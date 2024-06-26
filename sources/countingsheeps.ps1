﻿#================================================================================================================================

#----------------INFO----------------
#
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Project creator for Skrivanek GmbH
#
# Usage: powershell.exe -executionpolicy bypass -file ".\Rocketlaunch.ps1"
# Usage: Compiled form, just double-click.
#
#
#----------------STEPS----------------
#
# Initialization
# GUI
# Processing Input
# Build the project
# Bonus
#
#-------------------------------------


#===============================================
#                Initialization                =
#===============================================

# Arg
param([String]$arg)


# Fancy !
Write-Output "================================"
Write-Output "=        CountingSheeps        ="
Write-Output "================================"

Write-Output ""
Write-Output "For Skrivanek GmbH - Count number of words really, really quick!"
Write-Output "GPL-3.0 Stella Ménier, Project manager Skrivanek BELGIUM - <stella.menier@gmx.de>"
Write-Output "Git: https://github.com/teamcons/Skrivanek-CountingSheeps"
Write-Output ""
Write-Output ""


#========================================
# Get all important variables in place 

Write-Output "[STARTUP] Getting all variables in place"
[string]$APPNAME                        = "Counting Sheeps !"
[string]$SEP                            = ";"
[int]$WORDS_PER_HOUR                    = 1800
[int]$DECIMALS                          = 2
[string]$Form_Theme                     = '241,241,241' #"White" #"LightGreen"


#========================================
# Localization

[string]$text_column_type               = "Typ"
[string]$text_column_file               = "Datei"
[string]$text_column_words              = "Wortzahl"
[string]$text_column_proofreadtime      = "Std"
[string]$text_totalsum                  = "TOTAL"

[string]$text_about = "CountingSheeps V0.9
Für die Skrivanek GmbH - Zähle die Wörter sehr, sehr schnell!
GPL-3.0 Stella Ménier - stella.menier@gmx.de

Projektseite auf Github öffnen?"

[string]$GITHUB_LINK                    = "https://github.com/teamcons/Skrivanek-CountingSheeps"
[string]$text_label_how                 = "Einfach Dateien per Drag & Drop auf der Wiese ablegen!"
[string]$text_button_close              = "Schließen"
[string]$text_button_save               = "Speichern"
[string]$text_keepontop                 = "Über alle Fenster"



#===============================================
#                Some more init                =
#===============================================


if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }





Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 




#===========================================
#                BIG STRING                =
#===========================================



#================================
# Project icon in Base 64
# [Convert]::ToBase64String((Get-Content "..\assets\Skrivanek-Rocketlaunch-Icon.ico" -Encoding Byte))
[string]$iconBase64 = 'AAABAAEAQEAAAAEAIAAoQgAAFgAAACgAAABAAAAAgAAAAAEAIAAAAAAAAEAAAMMOAADDDgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALQAAAOwAAAD/T0FB/+O8vP//1tb//9TU///T0///09P//9PT///T0///4eL//+zu//3r7f//7/H//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7/H/x7i6/1lTVP9IQ0T/XFVW/1dQUf81MTL/CwsL/wAAAP8rFwD/iEgA/7FeAP+wXgD/XzIA/wEBAP8AAACyAAAADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEsAAAD3LiYm/923t///1tb//9PT///T0///09P//9PT///T0///1tb//+jq///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///u8P/86ev/+ebo//7r7f/96uz/793f/7yur/9PSUr/AQID/xcMAP+GRwD/jUsA/xMKAP8AAADpAAAAOwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAACcAgEB/5d9ff//1tb//9PT///T0///09P//9PT///T0///09P//9vc///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///t7///8PL/8+Hj/3pxcv8EBQb/HxAA/ysXAP8AAAD9AAAAfAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVAAAA0h0YGP/ctrb//9XV///T0///09P//9PT///T0///09P//9PT///g4f//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///t7///7/H//+/x//7r7f/l1Nb/S0VG/wAAAP8AAAD/AAAA2wAAAEUAAAAgAAAAEAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKwAAAOo6MDD/88nJ///U1P//09P//9PT///T0///09P//9PT///T0///4uP//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+3v///v8f/35Ob/zLy+/46Ehf9XUFH/MS4v/xEREv8AAAD/AAAA/wAAAPkAAADtAAAA4gAAAMwAAACmAAAAbgAAADEAAAAHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALgAAAKoAAAD9YFBQ//3S0v//09P//9PT///T0///09P//9PT///T0///09P//+Hi///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+/x//Ti5P+sn6H/TUdI/xAQEP8AAAD/AgAA/xUKBf8qFw3/OyET/0MlFv9CJRb/OiAT/yoYDv8XDQf/BgMC/wAAAP8AAADrAAAArQAAAE0AAAAJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAgIAAAAATAAAANwAAAD/Niws/8uoqP//1dX//9PT///T0///09P//9PT///T0///09P//9PT///e3v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+7w/8y9vv9RS0v/BQUG/wEAAP8qFw3/Zjki/5hVMv+3Zz3/x3BC/850RP/RdUX/0XVF/850RP/Hb0L/tmY8/5ZUMv9jOCH/KBcN/wMCAf8AAAD3AAAAkwAAAAcAAAAA03dGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAALsKCAj/dWFh/+a+vv//1NT//9PT///T0///09P//9PT///T0///09P//9PT///T0///2Nn//+rs///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu/6icnf8fHR3/AAAA/zEbEP+KTi7/xG5B/9Z4R//XeUf/1nhH/9V3Rv/Ud0b/1HdG/9R3Rv/Ud0b/1XdG/9Z4R//XeUf/1nhH/8JtQP+FSyz/MBsQ/zAbEN+QUTAmmlYzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADqx8cA58HBAPHKyhGliYnFrY+P//rPz///1dX//9PT///T0///09P//9PT///T0///09P//9PT///T0///09L//9na///r7f//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+3v/6OXmP8QDxD/BwMB/2Y6Iv/DbUD/13lH/9V4Rv/Ud0b/1HdG/9R3Rv/Ud0b/1HdG/9R3Rv/Ud0b/1HdG/9R3Rv/Ud0b/1HdG/9R3Rv/VeEb/13lH/8ZvQf/Gb0H+1XdGvMpxQysAAAAIAwIBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFhMTAHhjYwD/2ttj/9bW+//W1v//09P//9PT///T0///09P//9PT///T0///09P//9PT///T0///09P//9vb///o6f//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+/x/8GztP8YFhf/CQQC/39HKv/SdkX/1XhG/9R3Rv/Ud0b/1HdG/9R3Rv/Ud0b/1HdG/9R3Rv/Ud0b/1HdG/9R3Rv/Ud0b/1HdG/9R3Rv/Ud0b/1HdG/9R3Rv/VeEb/1XhG/9Z4R/99RinaAAAAtwAAADQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAABZnIKC1v3R0f//09P//9PT///T0///09P//9PT///T0///09P//9PT///T0///09P//9vb///q7P//7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu///s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AP9LKAD/AAAA/zMdEf+qXzj/Zzoi/wwGA/8AAAD/AQID/xIPEP91YWH/78XF///X1///3d7//+zu///w8v/r2tv/bmZm/xAPEP8BAgP/AAAA/w8IBP9uPiT/p143/y0ZD/8AAAD/UiwA/69dAP+mWAD/ajgA/xgNAP8AAAD/AAAAsgAAACEAAAAAAAAASAAAAOoTDQ3/hl9c/9GUj//enZj/05WQ/25OTP8AAAD/CwkJ/451df/5zs7//9fX///l5v//7O///+zu///s7v//7vD///Dy///w8v//7vD//+zu///s7v//7vD/59bY/25lZv8KCgv/AQAA/zIaAP94PwD/lU8A/zwgAP8BAAD/FQwH/wIBAP8hGxz/hm9v/7OUlP/Pq6v/+8/P///U1P//1tb//+jp///s7v//7O7//+7w//nn6f/Nvb//sqWm/4F4ef8cGxv/AwEA/xQLB/8BAAD/QiMA/5ZQAP90PgD/LxkA/wMCAP8AAADxAAAAigAAABQAAAAAAAAAAAAAAAEAAABoAAAA7wUDA/8pHRz/Piwr/ykdHP8GBAT/AAAA8QAAAPAIBgb/YVBQ/8mmpv/10tP//+vt///t7//55uj/2MjJ/5WKi/+MgYL/1sbI///v8f//7e///+zu///u8P/76Or/saSl/0E8Pf8GBgf/BAIA/yUTAP8+IQD/DwgA/wAAAP8aFRX/wJ+f///W1v//19f//9XV///T0///09P//9/g///s7v//7O7//+zu///s7v//7O7//+/x///w8v//7/H/uKqs/xQTE/8AAAD/EgoA/z8hAP8iEgD/BAIA/wAAAPkAAAC7AAAASQAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAEgAAACwAAAA4gAAAO4AAADjAAAAsAAAAEgAAABOAAAA0QAAAP8TDw//QDU1/2FYWP9lXl//SkVF/xwaG/8AAAD/AAAA/yYjI/+bj5H/6tja///s7v//7/H///Dy///t7//Iubv/NDAx/wAAAP8AAAD9AAAA/gEAAP8AAAD/ZVNT///T0///09P//9PT///T0///09P//9TU///m5///7O7//+zu///s7v//7O7//+zu///s7v//7O7//+zu//3q7P9ZU1P/AAAA/wEAAP8AAAD+AAAA6QAAALAAAABVAAAADgAAAAAAAAAAAAAAAAAA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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////////////////////////////////////////////////////////////////////////////////n/////8Hh/4D/////gAB/AH////8AAD4Af////gAAPgAB///+AAA8AAB///4AABwAAD///gAAHAAAP//+AAAcAAA///8AAAAAAB///wAAAAAAH///gAAAAAA///+AAAAAAD///4AAAAAAf///gAAAAAB///+AAAAAAP///wAAAAAB////AAAAAAA///8AAAAAAAf//gAAAAAAAf/8AAAAAAAA//gAAAAAAAD/+AAAAAAAAD/4AAAAAAAAH+AAAAAAAAAP4AAAAAAAAA/gAAAAAAAAB+AAAAAAAAAH4AAAAAAAAAHgAAAAAAAAAHAAAAAAAAAAMAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAwAAAAAAAAAHgAAAAAAAAB/gwAAAAAAAf//wIAMAAAP////4BwAAA///////gAAH///////AAA///////+AAH////////gH/////////z/////////////////////////////////////////////////////////////////////////////8='

# phew

# Rebuild an image from base 64
$iconBytes = [Convert]::FromBase64String($iconBase64)

# initialize a Memory stream holding the bytes
$stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

# This way we can draw icons without having any external file
$icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))




#==========================================
#                Functions                =
#==========================================


#================================
# Save to CSV
function Save-DataGridView
{

    [int]$totalwords                            =  $datagridview.Rows[ ($datagridview.Rows.Count - 1) ].Cells[2].Value
    [string]$defaultname                        = -join("CountingsSheeps-",$totalwords,"w.csv")

    $OpenFileDialog                             = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory            = $ScriptPath
    $OpenFileDialog.filename                    = $defaultname
    $OpenFileDialog.AddExtension                = $true
    $OpenFileDialog.filter                      = "Comma Separated Values (*.csv)|*.csv|All files (*.*)|*.*"

    $result = $OpenFileDialog.ShowDialog()
    $analysisfile = $OpenFileDialog.filename
    
    # Cancel culture : Drop everything if cancel
    if ( ($result -ne [System.Windows.Forms.DialogResult]::Cancel) -and ($analysisfile -ne $none) -and ($analysisfile -ne ""))
    {


        # Create the CSV, specify separator to avoid issues opening the csv in your fav office software
        Write-Output "sep=$SEP" | Out-File -FilePath "$analysisfile"

        # Add column headers
        $top = -join($datagridview.Columns[1].Name,$SEP,$datagridview.Columns[2].Name,$SEP,$datagridview.Columns[3].Name,$SEP)
        Write-Output $top | Out-File -FilePath "$analysisfile" -Append 


        # Rebuild and append each line
        foreach ($row in $datagridview.Rows )
        {
            $line = -join($row[0].Cells[1].Value,$SEP,$row[0].Cells[2].Value,$SEP,$row[0].Cells[3].Value,$SEP)
            Write-Output $line | Out-File -FilePath "$analysisfile" -Append
        }
    }
}


#==============================================================
#                GUI - Ask the Right Questions                =
#==============================================================


#================================
#= INITIAL WORK =

[int]$MainWindow_leftalign = 10
[int]$MainWindow_verticalalign = 180


$MainWindow                   = New-Object System.Windows.Forms.Form
$MainWindow.Text              = $APPNAME
$MainWindow.Size              = New-Object System.Drawing.Size(335,($MainWindow_verticalalign + 25))
$MainWindow.MinimumSize       = New-Object System.Drawing.Size(335,($MainWindow_verticalalign + 25))
$MainWindow.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Regular)
$MainWindow.StartPosition     = 'CenterScreen'
$MainWindow.MaximizeBox       = $True
$MainWindow.Topmost           = $True
$MainWindow.BackColor         = $Form_Theme
$MainWindow.Icon              = $icon
$MainWindow.AllowDrop         = $True


#================================
# FANCY ICON
$pictureBox             = new-object Windows.Forms.PictureBox
$pictureBox.Location    = New-Object System.Drawing.Point($MainWindow_leftalign,0)
$img = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
$pictureBox.Width       = 64 #$img.Size.Width
$pictureBox.Height      = 64 #$img.Size.Height
$pictureBox.Image       = $img;
$pictureBox.Add_Click({
                    $Result = [System.Windows.Forms.MessageBox]::Show($text_about,$APPNAME,4,[System.Windows.Forms.MessageBoxIcon]::Information)
                    If ($Result -eq "Yes") { Start-Process $GITHUB_LINK } })

#================================
# LABEL AND TEXT
$label                  = New-Object System.Windows.Forms.Label
$label.Text             = $APPNAME
$label.Location         = New-Object System.Drawing.Point(($MainWindow_leftalign + 75),10) # leftalign+80 if icon
$label.Size             = New-Object System.Drawing.Size(180,20)
$label.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Bold)

$MainWindow.controls.add($pictureBox)
$MainWindow.Controls.Add($label)


#================================
# Else the label doesnt wrap around neatly
$wraparound_panel                       = New-Object System.Windows.Forms.Panel
$wraparound_panel.Location              = New-Object System.Drawing.Point(($MainWindow_leftalign + 75),30)
$wraparound_panel.Height                = 30
$wraparound_panel.Width                 = 235
$wraparound_panel.AutoSize              = $true
$wraparound_panel.BackColor             = $Form_Theme
$wraparound_panel.Anchor                      = "Right,Left,Top"

#================================
# Label and button
$labelgrid                  = New-Object System.Windows.Forms.Label
$labelgrid.Text             = $text_label_how
$labelgrid.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Italic)
$labelgrid.Dock             = "Fill"


$wraparound_panel.Controls.Add($labelgrid)
$MainWindow.Controls.Add($wraparound_panel)


#===============================================
#                GUI - THE GRID                =
#===============================================



## Configure the Gridview
$datagridview                           = New-Object System.Windows.Forms.DataGridView
$datagridview.Location                  = New-Object System.Drawing.Size($MainWindow_leftalign,65)
$datagridview.Size                      = New-Object System.Drawing.Size(300,55)
$datagridview.AutoSize                  = $true
$datagridview.BackgroundColor           = $Form_Theme
$datagridview.Anchor                    = "Left,Bottom,Top,Right"
$datagridview.AllowDrop                 = $True
$datagridview.ColumnCount               = 3
$datagridview.ColumnHeadersVisible      = $true
$datagridview.RowHeadersVisible         = $false
$datagridview.ReadOnly                  = $true
$datagridview.AutoSizeRowsMode          = "AllCells"
$datagridview.AutoSizeColumnsMode       = "Fill"
$datagridview.AllowUserToAddRows        = "False"
$datagridview.AllowUserToResizeRows     = "False"
$datagridview.AllowUserToDeleteRows     = $true
$datagridview.BorderStyle               = "None"
$datagridview.BorderStyle                      = "FixedSingle"
$datagridview.CellBorderStyle                  = "SingleHorizontal"
#$datagridview.AlternatingRowsDefaultCellStyle.BackColor = "White"

$datagridview.Columns[0].Name = $text_column_file
$datagridview.Columns[0].Width = 100

$datagridview.Columns[1].Name = $text_column_words
$datagridview.Columns[1].Width = 40
$datagridview.Columns[1].DefaultCellStyle.Alignment = "MiddleRight" 
$datagridview.Columns[1].HeaderCell.Style.Alignment = "MiddleRight" 

$datagridview.Columns[2].Name = $text_column_proofreadtime
$datagridview.Columns[2].Width = 40
$datagridview.Columns[2].DefaultCellStyle.Alignment = "MiddleRight" 
$datagridview.Columns[2].HeaderCell.Style.Alignment = "MiddleRight" 


#================================
# Add an image column. Has to be inserted afterward. Idk why
$ImageColumn                                    = New-Object System.Windows.Forms.DataGridViewImageColumn
$ImageColumn.HeaderText                         = $text_column_type
$ImageColumn.Resizable                         = "false"
$ImageColumn.AutoSizeMode                      = "none"
$datagridview.Columns.Insert(0, $ImageColumn);
$datagridview.Columns[0].Width                  = 32



# Adding an image column adds a weird unremovable line. Use it for sum.
$miniico =  ([System.Drawing.Icon]::ExtractAssociatedIcon(([Environment]::GetCommandLineArgs()[0])) ).ToBitmap()
$datagridview.Rows[0].Cells[0].Value            = $miniico


#================================
# Add a button column. Has to be inserted afterward. Idk why
$ButtonColumn                                   = New-Object System.Windows.Forms.DataGridViewButtonColumn
$ButtonColumn.HeaderText                         = "X"
#$ButtonColumn.FlatStyle                         = "Flat"
$ButtonColumn.Resizable                         = "false"
$ButtonColumn.AutoSizeMode                      = "none"
$datagridview.Columns.Insert(4, $ButtonColumn);
$datagridview.Columns[4].Width = 24




[string]$datagridview.Rows[0].Cells[1].Value = $text_totalsum
[int]$datagridview.Rows[0].Cells[2].Value = 0
[int]$datagridview.Rows[0].Cells[3].Value = 0
[string]$datagridview.Rows[0].Cells[4].Value = ""

<# $font                                   = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$datagridview.Rows[0].Cells[1].DefaultCellStyle.Font     = $font
$datagridview.Rows[0].Cells[2].Font     = $font
$datagridview.Rows[0].Cells[3].Font     = $font #>



$MainWindow.Controls.Add($datagridview)

#===================================================
#                GUI - Down Buttons                =
#===================================================


#================================
$gui_panel                                  = New-Object System.Windows.Forms.Panel
$gui_panel.Left                             = 0
$gui_panel.Top                              = ($MainWindow_verticalalign)
$gui_panel.Width                            = 440
$gui_panel.Height                           = 35
$gui_panel.BackColor                        = '241,241,241'
$gui_panel.Dock                             = "Bottom"

#================================
$gui_keepontop                              = New-Object System.Windows.Forms.Checkbox
$gui_keepontop.Location                     = New-Object System.Drawing.Point(($MainWindow_leftalign),7)
$gui_keepontop.Size                         = New-Object System.Drawing.Size(120,20)
$gui_keepontop.Text                         = $text_keepontop
$gui_keepontop.UseVisualStyleBackColor      = $True
$gui_keepontop.Anchor                       = "Bottom,Left"
$gui_keepontop.Checked                      = $MainWindow.Topmost


#================================
$gui_saveButton                               = New-Object System.Windows.Forms.Button
$gui_saveButton.Location                      = New-Object System.Drawing.Point(($MainWindow_leftalign + 255),3)
$gui_saveButton.Size                          = New-Object System.Drawing.Size(80,25)
$gui_saveButton.Text                          = $text_button_save
$gui_saveButton.UseVisualStyleBackColor       = $True
$gui_saveButton.Anchor                        = "Bottom,Right"
#$gui_okButton.BackColor                     = ”Green”
$gui_saveButton.Add_Click({Save-DataGridView})
#$gui_saveButton.DialogResult              = [System.Windows.Forms.DialogResult]::OK
#$MainWindow.AcceptButton                          = $gui_saveButton

#================================
$gui_cancelButton                           = New-Object System.Windows.Forms.Button
$gui_cancelButton.Location                  = New-Object System.Drawing.Point(($MainWindow_leftalign + 340),3)
$gui_cancelButton.Size                      = New-Object System.Drawing.Size(80,25)
$gui_cancelButton.Text                      = $text_button_close
$gui_cancelButton.UseVisualStyleBackColor   = $True
$gui_cancelButton.Anchor                    = "Bottom,Right"
#$gui_cancelButton.BackColor                  = ”Red”
$gui_cancelButton.DialogResult              = [System.Windows.Forms.DialogResult]::Cancel
$MainWindow.CancelButton                    = $gui_cancelButton
#$gui_cancelButton.Add_Click({$MainWindow_FormClosed})

$gui_panel.Controls.Add($gui_saveButton)
$gui_panel.Controls.Add($gui_keepontop)
$gui_panel.Controls.Add($gui_cancelButton)
$gui_panel.Show()

[void]$MainWindow.Controls.Add($gui_panel)



#=========================================
#                Handlers                =
#=========================================



### Write event handlers ###


#================================
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


#================================
$DragDrop = [System.Windows.Forms.DragEventHandler]{
	foreach ($filepath in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
    {
        $file = Get-Item $filepath

        # Use different backend depending on what needed
        # Each time, check the extension to know what we deal with
        if ($file.Extension -match ".doc[|x]" )
        {
            # OPEN IN WORD, PROCESS COUNT
            $filecontent = $word.Documents.Open($file.FullName)
            [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticWords)
            #CLOSE FILE
            $filecontent.Close()
            
        }
<#         elseif ($file.Extension -match ".xls[|x]" )
        {

            #foreach ($cell in $b.ActiveSheet.Rows[3].Cells) { if ($cell.Text -ne "") {$cell.Text} }

            # OPEN IN EXCEL, PROCESS COUNT
            $filecontent = $excel.Workbooks.Open($file.FullName)
            [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Excel.WdStatistic]::wdStatisticWords)
            #CLOSE FILE
            $filecontent.Close()
        }
        elseif ($file.Extension -match ".ppt[|x]" )
        {
            # OPEN IN POWRPOINT, PROCESS COUNT
            $filecontent = $powerpoint.Documents.Open($file.FullName)
            [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Powerpoint.WdStatistic]::wdStatisticWords)
            #CLOSE FILE
            $filecontent.Close()
        } #>
        elseif ($file.Extension -match ".pdf" )
        {
            # COUNT WORDS IN PDF FILE
            [int]$wordcount = (Get-Content $file.FullName | Measure-Object –Word).Words
        }
    
        elseif ($file.Extension -match ".[txt|csv|md|log]" )
        {
            # COUNT WORDS IN TXT FILE
            [int]$wordcount = (Get-Content $file.FullName | Measure-Object –Word).Words
        }
        else
        {
            # IDK
            [int]$wordcount = 0
        }
            
        # Update totalcount
        $ico =  ([System.Drawing.Icon]::ExtractAssociatedIcon($filepath) ).ToBitmap()
        $proofreadtime = [math]::round(($wordcount / $WORDS_PER_HOUR),$DECIMALS)

        # Add
        $datagridview.Rows[ ($datagridview.Rows.Count - 1) ].Cells[2].Value += $wordcount
        $datagridview.Rows[ ($datagridview.Rows.Count - 1) ].Cells[3].Value += $proofreadtime
        $datagridview.Rows.Add($ico,$file.Name,$wordcount,$proofreadtime,"X");

        # HIGHER
        $MainWindow.Height = ($MainWindow.Height + 33)



	} # End of processing list
    
}


#================================
$DeleteRow = {
    # Check if we are not trying to delete the last
    # And if its actually the column with a X
	if (($datagridview.CurrentCell.RowIndex -ne ($datagridview.Rows.Count - 1) ) -and ($datagridview.CurrentCell.ColumnIndex -eq 4 ))
	{

        # Delete offending line
        $datagridview.Rows.RemoveAt($datagridview.CurrentCell.RowIndex)

        # Init count
        $wordcount          = 0
        $proofreadtime      = 0

        # Get old values out
        $datagridview.Rows[ ($datagridview.Rows.Count - 1) ].Cells[2].Value = $wordcount
        $datagridview.Rows[ ($datagridview.Rows.Count - 1) ].Cells[3].Value = $proofreadtime


        # Recount
        foreach ($row in $datagridview.Rows)
        {
            # Do not count the last
            if ($row.RowIndex -ne $datagridview.Rows.Count - 1 )
            {
                $wordcount      += $row.Cells[2].Value 
                $proofreadtime  += $row.Cells[3].Value
            }
        }

        # Wake up babe new count just dropped
        $datagridview.Rows[ ($datagridview.Rows.Count - 1) ].Cells[2].Value += $wordcount
        $datagridview.Rows[ ($datagridview.Rows.Count - 1) ].Cells[3].Value += $proofreadtime


        # LOWER
        $MainWindow.Height = ($MainWindow.Height - 33)

	}
}
 


#================================
$MainWindow_FormClosed = {
	try
    {
        $datagridview.remove_Click($button_Click)
		$datagridview.remove_DragOver($DragOver)
		$datagridview.remove_DragDrop($DragDrop)
        $datagridview.remove_DragDrop($DragDrop)
		$MainWindow.remove_FormClosed($MainWindow_Cleanup_FormClosed)
        $MainWindow.Close()

        # Wont need Word anymore
        $word.Quit()
        $excel.Quit()
        $powerpoint.Quit()

	}
	catch [Exception]
    { }
}
 
 


#================================
### Wire up events ###
 
#$button.Add_Click($button_Click)

$gui_keepontop.Add_Click({$MainWindow.Topmost = $gui_keepontop.Checked})

<# $label.Add_Click({
    echo "click"
    switch ($MainWindow.BackColor) {
        "LightGray"     { $MainWindow.BackColor = "White" }
        "White"         { $MainWindow.BackColor = "LightGreen" }
        "LightGreen"    { $MainWindow.BackColor = "LightGray" }
    }
    $wraparound_panel.BackColor     = $MainWindow.BackColor
    $datagridview.BackColor         = $MainWindow.BackColor
}) #>


# All the attach
$datagridview.Add_DragOver($DragOver)
$datagridview.Add_DragDrop($DragDrop)
$datagridview.Add_CellMouseClick($DeleteRow)
$MainWindow.Add_DragOver($DragOver)
$MainWindow.Add_DragDrop($DragDrop)
$MainWindow.Add_FormClosed($MainWindow_FormClosed)


$script:word = New-Object -ComObject Word.Application



# Go
$MainWindow.ShowDialog()

# Slow shit in the background, no one will notice
#$script:word = New-Object -ComObject Word.Application
#$script:excel = New-Object -ComObject Excel.Application
#$script:powerpoint = New-Object -ComObject Powerpoint.Application

# This makes it pop up
#$MainWindow.Activate()
 
# Create an application context for it to all run within. 
# This helps with responsiveness and threading.
#$appContext = New-Object System.Windows.Forms.ApplicationContext 
#[void][System.Windows.Forms.Application]::Run($appContext)
