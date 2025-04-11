# Import the System.Windows.Forms namespace for creating forms and controls
Add-Type -AssemblyName System.Windows.Forms

# Import the System.Drawing namespace for colors, sizes, and fonts
Add-Type -AssemblyName System.Drawing
# Function to validate a DOI URL
function Test-DOI {
    param (
        [string]$doiUrl
    )

    try {
        $headers = @{ "Accept" = "text/bibliography; style=bibtex" }
        $response = Invoke-WebRequest -Uri $doiUrl -Method Get -Headers $headers -ErrorAction Stop
        write-host "URL: $doiUrl"
        Write-Host "Success: Status Code $($response.StatusCode)"
        $authorstart = $response.Content.IndexOf("author={",[System.StringComparison]::OrdinalIgnoreCase)+8
        $authorend = $response.Content.IndexOf("}",$authorstart )
        $Authors =$response.Content.Substring($authorstart,($authorend-$authorstart)).Trim()
        $titlestart = $response.Content.IndexOf("title={",[System.StringComparison]::OrdinalIgnoreCase)+7
        $titleend = $response.Content.IndexOf("}",$titlestart ) 
        $Title =$response.Content.Substring($titlestart,($titleend-$titlestart)).Trim()
    } catch {
        Write-Host "Error: $($_.Exception.Message)"
    }
    return @{
        "StatusCode" = $response.StatusCode
        "Authors" = $Authors
        "Title" = $Title
    }   
}

#function to split the text into lines and extract authors and DOI URLs
function split-text {
    param (
        [string]$text
    )

    $authors=@()
    $doiUrl = @()
    $validURL = @()
    $DOIAuthors = @()
    $DOITitle = @()

    $reflist = $text.split("`n") 
    $counter =0
    foreach ($line in $reflist) {
        if ($line.length -eq 0) {
            continue
        }   
        $line = $line.Trim()
        write-host $line
            if ($line.Contains(").")) {
                $authorend = $line.indexof(").")+2

                $authors+= $line.Substring(0, $authorend).Trim()
            } else {
                Write-Host "Warning: Unable to extract author from the line." -ForegroundColor Yellow
                $authors+= "Author not valid"
            }

            $DOIStart = $line.indexOf("https://doi.org/", [System.StringComparison]::OrdinalIgnoreCase)

            if ($DOIStart -ne -1) {
                $doiUrl+= $line.Substring($DOIStart).Trim()
                $urldata = test-DOI -doiUrl $doiUrl[$counter]
               
            if ($urldata -and ($urldata.StatusCode -eq 200 -or $urldata.StatusCode -eq 302)) {
                $validURL += $true
                $DOIAuthors += $urldata.Authors
                $DOITitle += $urldata.Title
                Write-Host "Valid DOI URL: $($doiUrl[$counter])" -ForegroundColor Green
            } else {
                Write-Host "Invalid DOI URL: $($doiUrl[$counter])" -ForegroundColor Red
                $validURL+= $false
                $DOIAuthors += "No Lookup"
                $DOITitle += "No Lookup"
            }



            } else {
                $doiUrl+= "NO DOI URL"
                Write-Host "DOI URL not found in the line." -ForegroundColor Yellow
                $validURL+= $false
                $DOIAuthors += "No Lookup"
                $DOITitle += "No Lookup"
            }

    $counter++
    $reflist=$reflist.TrimEnd("`r","`n")
        }
        for ($j=0; $j -lt $reflist.Count; $j++) {
           write-host "Reference: " + $reflist[$j] + "`r`n"
           write-host "Authors: " + $authors[$j] + "`r`n"
              write-host "DOI Authors: " + $DOIAuthors[$j] + "`r`n"

              write-host "DOI Title: " + $DOITitle[$j] + "`r`n"
              write-host "Valid DOI URL: " + $doiUrl[$j] + "`r`n"

            


        }
        Showresultsform -authors $authors -doiUrl $doiUrl -validURL $validURL -DOIAuthors $DOIAuthors -DOITitle $DOITitle -reflist $reflist

    }
    # Function to display the results in a Windows Form
    function Showresultsform {
        param (
            [Array]$authors,
            [Array]$doiUrl,
            [Array]$validURL,
            [Array]$DOIAuthors,
            [Array]$DOITitle,
            [Array]$reflist
        )
        $screen = [System.Windows.Forms.Screen]::PrimaryScreen
        $resultsform = New-Object System.Windows.Forms.Form
        
        $resultsform.Width = $screen.WorkingArea.Width * 0.8  # 80% of screen width
        $resultsform.Height = $screen.WorkingArea.Height * 0.8  # 80% of screen height
        $resultsform.Text = "DOI Validation Results"
        $resultsform.Font = New-Object System.Drawing.Font("Arial", 12)
        $resultsform.StartPosition = "CenterScreen"


        $textBox = New-Object System.Windows.Forms.RichTextBox
        $textBox.Multiline = $true
        $textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $textBox.Font = New-Object System.Drawing.Font("Arial", 12)

        $textBox.ReadOnly = $true
        $resultsform.Controls.Add($textBox)

        $Backcolorcount=0
        $backcolor=[system.drawing.color]::gray
        for ($i=0; $i -lt $reflist.Count; $i++) {
            if($backColorcount %2 -eq 0){
                $backcolor=[system.drawing.color]::Black
            }else{
                $backcolor=[system.drawing.color]::Blue
            }
            $backColorcount++

                $textBox.SelectionColor = $backcolor
                $textBox.AppendText("Reference: " + $reflist[$i] + "`r`n")
                $textBox.SelectionColor = $backcolor
                $textBox.AppendText("Authors: " + $authors[$i] + "`r`n")
                $textBox.SelectionColor = $backcolor
                $textBox.AppendText("DOI Authors: " + $DOIAuthors[$i] + "`r`n")
                $textBox.SelectionColor = $backcolor
                $textBox.AppendText("DOI Title: " + $DOITitle[$i] + "`r`n")
            if ($validURL[$i] -eq $true) {
                $textBox.SelectionColor =[System.Drawing.Color]::Green
                $textBox.AppendText("Valid DOI URL: " + $doiUrl[$i] + "`r`n`r`n")
            } else {
                $textBox.SelectionColor = [System.Drawing.Color]::Red
                $textBox.AppendText("Invalid DOI URL: " + $doiUrl[$i] + "`r`n`r`n")
            }
        }

        
        $resultsform.Controls.Add($dataGridView)
        #>
        
        # Show the form
        [void]$resultsform.ShowDialog()
    }

# Add a Windows Form for DOI input
Add-Type -AssemblyName System.Windows.Forms


$form = New-Object System.Windows.Forms.Form
$form.Text = "DOI Validator"
$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$form.Width = $screen.WorkingArea.Width * 0.8  # 80% of screen width
$form.Height = $screen.WorkingArea.Height * 0.8  # 80% of screen height
$form.Font = New-Object System.Drawing.Font("Arial", 12)
$form.StartPosition = "CenterScreen"
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.Size = New-Object System.Drawing.Size($form.Width - 40, $form.Height - 350) 
$textBox.Multiline = $true
$textBox.Location = New-Object System.Drawing.Point(20, 20)
$textBox.ScrollBars = "Vertical"    
$textBox.Font = New-Object System.Drawing.Font("Arial", 12)
$textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($textBox)

$button = New-Object System.Windows.Forms.Button
$button.Text = "Validate"
$button.Dock = [System.Windows.Forms.DockStyle]::Bottom
$button.Size = New-Object System.Drawing.Size(150, 40)  
$button.Location = New-Object System.Drawing.Point(140, 200)
$button.Font = New-Object System.Drawing.Font("Arial", 12)
$form.Controls.Add($button)


<#This is for a potential featuer later
$checkBox = New-Object System.Windows.Forms.CheckBox
$checkBox.Text = "Check All URLs"
$checkBox.Size = New-Object System.Drawing.Size(150, 30)  
$checkBox.Dock = [System.Windows.Forms.DockStyle]::Bottom
$checkBox.Font = New-Object System.Drawing.Font("Arial", 12)
$form.Controls.Add($checkBox)
#>
# Create a Button for validation


$button.Add_Click({
    $incomingText = $textBox.Text.Trim()
    split-text -text $incomingText
    })
[void]$form.ShowDialog()




