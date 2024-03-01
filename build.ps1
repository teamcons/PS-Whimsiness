
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }

ps2exe `
-inputFile $ScriptPath\src\countingsheeps.ps1 `
-iconFile $ScriptPath\assets\bluesheep.ico `
-noConsole `
-noOutput `
-exitOnCancel `
-title "CountingSheeps" `
-description "Count words quick !" `
-company "Skrivanek GmbH" `
-copyright "CC0 1.0 Universal Stella - stella.menier@gmx.de" `
-version 0.9 `
-Verbose `
-outputFile $ScriptPath\countingsheeps.exe
