

#----------------INFO----------------
#
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Just a scratchpad
#


Write-Output "Scratchpad"
Write-Output "GPL-3.0 Stella Ménier, Project manager Skrivanek BELGIUM - <stella.menier@gmx.de>"
Write-Output "Git: https://github.com/teamcons/Skrivanek-CountingSheeps"
Write-Output ""
Write-Output ""

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

# Define main
$MainWindow                   = New-Object System.Windows.Forms.Form
$MainWindow.Text              = "Scratchpad"
$MainWindow.StartPosition     = 'CenterScreen'
$MainWindow.Width     			= 300
$MainWindow.Height     			= 200
$MainWindow.StartPosition     = 'CenterScreen'
$MainWindow.Topmost           = $True
$MainWindow.Icon              = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\notepad.exe")
$MainWindow.AllowTransparency = $true
$MainWindow.AllowDrop         = $true
#$MainWindow.AutoSize          = $true

# Define textbox
$textBox 			        = New-Object System.Windows.Forms.RichTextBox
$textBox.text			    = -join("~ Write stuff here ~","`r`n","`r`n",(Get-Clipboard),"`r`n")
$textBox.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Regular)
$textBox.Dock 			    = "Fill"
$textBox.Multiline 		    = $true
#$textbox.AcceptsReturn 		= $true
$textBox.BorderStyle 		    = "None"
$textBox.AllowDrop         = $true


## Configure the Gridview
$datagridview                           			= New-Object System.Windows.Forms.DataGridView
$datagridview.Height                  				= 60
$datagridview.Anchor                    			= "Left,Bottom,Top,Right"
$datagridview.ColumnCount               			= 2
$datagridview.ColumnHeadersVisible      			= $false
$datagridview.RowHeadersVisible         			= $false
$datagridview.ReadOnly                  			= $true
$datagridview.AutoSizeColumnsMode       			= "Fill"
$datagridview.Dock 			    					= "Bottom"
$datagridview.ReadOnly                  			= $false

$datagridview.Columns[0].Width = 70
$datagridview.Columns[1].Width = 30

[void]$datagridview.Rows.Add()
[string]$datagridview.Rows[0].Cells[0].Value 		= "1+1"
$calculated 										= [Data.DataTable]::New().Compute($datagridview.Rows[0].Cells[0].Value, $null)
$datagridview.Rows[0].Cells[1].Value 				= -join("= ",$calculated)
$datagridview.Rows[0].Cells[1].ReadOnly 			= $true
$datagridview.Rows[0].Cells[1].Style.BackColor 		= 'LightGray'

[void]$datagridview.Rows.Add()
[string]$datagridview.Rows[1].Cells[0].Value 		= "1+1"
$calculated 										= [Data.DataTable]::New().Compute($datagridview.Rows[1].Cells[0].Value, $null)
$datagridview.Rows[1].Cells[1].Value 				= -join("= ",$calculated)
$datagridview.Rows[1].Cells[1].ReadOnly 			= $true
$datagridview.Rows[1].Cells[1].Style.BackColor 		= 'LightGray'


#======================================
#                Events               =
#======================================

# Add path to file when a file is dropped on it
$DragOver = [System.Windows.Forms.DragEventHandler]{
	if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop))
	{
	    $_.Effect = 'Copy'
	}
	else
	{
	    $_.Effect = 'None'
	}
}
$DragDrop = [System.Windows.Forms.DragEventHandler]{
	foreach ($filepath in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
    {
        $textBox.text = -join($textBox.text,"`r`n",$filepath,"`r`n")
	}
    
}

# Add those events
$textBox.Add_DragOver($DragOver)
$textBox.Add_DragDrop($DragDrop)

# Change opacity depending on focus
$textBox.Add_GotFocus({ $MainWindow.Opacity  = 1 })
$textBox.Add_LostFocus({ $MainWindow.Opacity = .55 })

$datagridview.Add_GotFocus({ $MainWindow.Opacity  = 1 })
$datagridview.Add_LostFocus({ $MainWindow.Opacity = .55 })


#
$datagridview.Rows[0].Cells[0].CellVa



# Add and done
$MainWindow.Controls.Add($datagridview)
$MainWindow.Controls.Add($textBox)
$MainWindow.ShowDialog()
