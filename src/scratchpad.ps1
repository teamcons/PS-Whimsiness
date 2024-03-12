

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

$MainWindow                   = New-Object System.Windows.Forms.Form
$MainWindow.Text              = "Scratchpad"
$MainWindow.StartPosition     = 'CenterScreen'
$MainWindow.Topmost           = $True
$MainWindow.Icon              = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\notepad.exe")
$MainWindow.AllowTransparency = $true
$MainWindow.Opacity 		  = .85

$textBox 			        = New-Object System.Windows.Forms.TextBox
$textBox.text			    = Get-Clipboard
$textBox.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Regular)
$textBox.Dock 			    = "Fill"
$textBox.Multiline 		    = $true
$textbox.AcceptsReturn 		= $true

$MainWindow.Controls.Add($textBox)
$MainWindow.ShowDialog()



