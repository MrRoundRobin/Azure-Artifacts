#region input

$currentVersion = "https://www.xmind.net/xmind/downloads/xmind-8-update8-windows.zip"

#endregion

#region computed variables

$tempPath = Join-Path -Path $env:TEMP -ChildPath "XMind.zip"

$installBasePath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles)
$installPath = Join-Path -Path $installBasePath -ChildPath "XMind 8 Update 8"

$startMenuPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonStartMenu)
$shortcutPath  = Join-Path -Path $startMenuPath -ChildPath "XMind.lnk"

#endregion

#region installation

Invoke-WebRequest -Uri $currentVersion -OutFile $tempPath -UseBasicParsing

Expand-Archive -Path $tempPath -DestinationPath $installBasePath

$configPath = Join-Path -Path $installPath -ChildPath "configuration"
$acl = Get-Acl -Path $configPath

$identity = New-Object System.Security.Principal.SecurityIdentifier("BU") # Users
$ace = New-Object System.Security.AccessControl.FileSystemAccessRule($identity, 'Modify', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
$acl.AddAccessRule($ace)

Set-Acl -Path $configPath -AclObject $acl

Remove-Item -Path $tempPath

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = Join-Path -Path $installPath -ChildPath "XMind.exe"
$Shortcut.WorkingDirectory = $installPath
$Shortcut.Save()

#endregion
