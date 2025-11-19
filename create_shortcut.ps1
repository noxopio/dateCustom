# create_shortcut.ps1
# Crea un acceso directo en el Escritorio que abre CMD en el folder
# y ejecuta `setDate_quick.bat`.

$WshShell = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'setDate'

# Detectar la carpeta donde está este script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = 'C:\Windows\System32\cmd.exe'
$args = "/k cd /d `"$scriptPath`" && setDate_quick.bat"

$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $target
$shortcut.Arguments = $args
$shortcut.WorkingDirectory = $scriptPath
$shortcut.IconLocation = 'C:\Windows\System32\cmd.exe,0'
$shortcut.Save()

Write-Output "Acceso directo creado: $shortcutPath"
