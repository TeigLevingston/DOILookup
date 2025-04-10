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
        #Write-Host "Response Content: ($response.Content)" -ForegroundColor Green
        $authorstart = $response.Content.IndexOf("author={",[System.StringComparison]::OrdinalIgnoreCase)+8
        $authorend = $response.Content.IndexOf("}",$authorstart )
        $Authors =$response.Content.Substring($authorstart,($authorend-$authorstart)).Trim()
        $titlestart = $response.Content.IndexOf("title={",[System.StringComparison]::OrdinalIgnoreCase)+7
        $titleend = $response.Content.IndexOf("}",$titlestart ) 
        $Title =$response.Content.Substring($titlestart,($titleend-$titlestart)).Trim()
        #Write-Host $Authors
    } catch {
        Write-Host "Error: $($_.Exception.Message)"
    }
    return @{
        "StatusCode" = $response.StatusCode
        "Authors" = $Authors
        "Title" = $Title
    }   
}


function split-text {
    param (
        [string]$text
    )
    #write-host $text
    $authors=@()
    $doiUrl = @()
    $validURL = @()
    $DOIAuthors = @()
    $DOITitle = @()

    $reflist = $text.split("`n") # | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    $counter =0
    foreach ($line in $reflist) {
        if ($line.length -eq 0) {
            continue
        }   
        $line = $line.Trim()
        write-host $line
            if ($line.Contains(").")) {
                $authorend = $line.indexof(").")+2
                #write-host $authorend
                $authors+= $line.Substring(0, $authorend).Trim()
            } else {
                Write-Host "Warning: Unable to extract author from the line." -ForegroundColor Yellow
                $authors+= "Author not valid"
            }
            #write-host $authors[$counter]
            $DOIStart = $line.indexOf("https://doi.org/", [System.StringComparison]::OrdinalIgnoreCase)
            #Write-Host $DOIStart
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




                write-host $validURL[$counter]
                #write-host $doiUrl[$counter]
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
        #$resultsform.Size = New-Object System.Drawing.Size(800, 600)
        $resultsform.StartPosition = "CenterScreen"

        # Create a TextBox to display the results
        $textBox = New-Object System.Windows.Forms.RichTextBox
        $textBox.Multiline = $true
        $textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
        #$textBox.Size = New-Object System.Drawing.Size(800, 600)
        #$textBox.Location = New-Object System.Drawing.Point(20, 20)
        $textBox.ReadOnly = $true
        $resultsform.Controls.Add($textBox)
        # Populate the TextBox with the results
        $Backcolorcount=0
        $backcolor=[system.drawing.color]::gray
        for ($i=0; $i -lt $reflist.Count; $i++) {
            if($backColorcount %2 -eq 0){
                $backcolor=[system.drawing.color]::Black
            }else{
                $backcolor=[system.drawing.color]::Blue
            }
            $backColorcount++
        
            #.TrimEnd("`r","`n")
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
<##>
        # Create a DataGridView for displaying tabular data
        <#
        $dataGridView = New-Object System.Windows.Forms.DataGridView
        $dataGridView.Size = New-Object System.Drawing.Size(350, 100)
        $dataGridView.Location = New-Object System.Drawing.Point(20, 240)
        $dataGridView.ColumnCount = 5
        $dataGridView.Columns[0].Name = "APA Authors"
        $dataGridView.Columns[1].Name = "DOI URL"
        $dataGridView.Columns[2].Name = "DOI Authors"
        $dataGridView.Columns[3].Name = "DOI Title"
        $dataGridView.Columns[4].Name = "Valid DOI URL"
        $dataGridView.AutoSizeColumnsMode = "Fill"
        $dataGridView.AllowUserToAddRows = $false
        $dataGridView.AllowUserToDeleteRows = $false
        for ($i=0; $i -lt $authors.Count; $i++) {

            $dataGridView.Rows.Add($authors[$i],$doiUrl[$i],$DOIAuthors[$i],$DOITitle[$i],$doiUrl[$i]) 
            <#
            $textBox.AppendText("Authors: " + + "`r`n")
            $textBox.AppendText("Valid DOI URL: " +  + "`r`n")
            $textBox.AppendText("DOI Authors: " +  + "`r`n")
            $textBox.AppendText("DOI Title: " +  + "`r`n`r`n")
        if ($validURL[$i]) {
            $textBox.AppendText("Valid DOI URL: " +  + "`r`n")
        } else {$textBox.AppendText("Invalid DOI URL: " + $doiUrl[$i] + "`r`n")
            $textBox.AppendText("Invalid DOI URL: " + $doiUrl[$i] + "`r`n")
        }
    }

        $dataGridView.Rows.Add("Row 2, Col 1", "Row 2, Col 2")
        #>
        
        $resultsform.Controls.Add($dataGridView)
        #>
        
        # Show the form
        [void]$resultsform.ShowDialog()
    }

# Add a Windows Form for DOI input and validation
Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "DOI Validator"
$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$form.Width = $screen.WorkingArea.Width * 0.8  # 80% of screen width
$form.Height = $screen.WorkingArea.Height * 0.8  # 80% of screen height
#form.Size = New-Object System.Drawing.Size(400, 300)
$form.StartPosition = "CenterScreen"

# Create a multiline TextBox for input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$boxheight = $form.ClientSize.Height - 300
$textBox.Size = New-Object System.Drawing.Size($form.ClientSize.Width, $boxheight)
$textBox.Multiline = $true
$textBox.Location = New-Object System.Drawing.Point(20, 20)
$textBox.ScrollBars = "Vertical"    
# The size will be automatically adjusted by the Dock property
#$textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($textBox)

$button = New-Object System.Windows.Forms.Button
$button.Text = "Validate"
$button.Dock = [System.Windows.Forms.DockStyle]::Bottom
#$button.Dock = [System.Windows.Forms.DockStyle]::Right

$button.Size = New-Object System.Drawing.Size(100, 30)
$button.Location = New-Object System.Drawing.Point(140, 200)
$form.Controls.Add($button)

# Create a CheckBox for additional options
$checkBox = New-Object System.Windows.Forms.CheckBox
$checkBox.Text = "Check All URLs"
$checkBox.Size = New-Object System.Drawing.Size(100, 20)
$checkBox.Dock = [System.Windows.Forms.DockStyle]::Bottom
#$checkBox.Dock = [System.Windows.Forms.DockStyle]::Left
#$checkBox.Location = New-Object System.Drawing.Point(20, 180)
$form.Controls.Add($checkBox)

# Create a Button for validation


# Add event handler for the button click
$button.Add_Click({
    $incomingText = $textBox.Text.Trim()
    split-text -text $incomingText
    })
[void]$form.ShowDialog()



# Show the form
