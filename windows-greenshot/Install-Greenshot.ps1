#region input

$currentVersion = "https://github.com/greenshot/greenshot/releases/download/Greenshot-RELEASE-1.2.10.6/Greenshot-NO-INSTALLER-1.2.10.6-RELEASE.zip"

#endregion

#region computed variables

$tempPath = Join-Path -Path $env:TEMP -ChildPath "Greenshot.zip"

$installBasePath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles)
$installPath = Join-Path -Path $installBasePath -ChildPath "Greenshot"

$startMenuPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonStartMenu)
$shortcutPath  = Join-Path -Path $startMenuPath -ChildPath "Greenshot.lnk"

#endregion

#region installation

Invoke-WebRequest -Uri $currentVersion -OutFile $tempPath -UseBasicParsing

New-Item -Path $installPath -ItemType Directory | Out-Null

Expand-Archive -Path $tempPath -DestinationPath $installPath

Remove-Item -Path $tempPath

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = Join-Path -Path $installPath -ChildPath "Greenshot.exe"
$Shortcut.Save()

#endregion
