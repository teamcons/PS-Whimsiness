

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
$MainWindow.Topmost           = $True
$MainWindow.Icon              = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\notepad.exe")
$MainWindow.AllowTransparency = $true
$MainWindow.AllowDrop         = $true
$MainWindow.AutoSize          = $true

# Define textbox
$textBox 			        = New-Object System.Windows.Forms.RichTextBox
$textBox.text			    = -join("~ Write stuff here ~","`r`n","`r`n",(Get-Clipboard),"`r`n")
$textBox.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Regular)
$textBox.Dock 			    = "Fill"
$textBox.Multiline 		    = $true
#$textbox.AcceptsReturn 		= $true
$textBox.BorderStyle 		    = "None"
$textBox.AllowDrop         = $true


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

# Add and done
$MainWindow.Controls.Add($textBox)
$MainWindow.ShowDialog()
