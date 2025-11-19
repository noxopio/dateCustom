# create_shortcut.ps1
# Crea un acceso directo en el Escritorio que abre CMD en el folder
# y ejecuta `setDate_quick.bat`.

$WshShell = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
# El nombre completo del acceso directo debe terminar en .lnk o .url
$shortcutPath = Join-Path $desktop 'setDate.lnk'

# Detectar la carpeta donde está este script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = 'C:\Windows\System32\cmd.exe'
$shortcutArgs = "/k cd /d `"$scriptPath`" && setDate_quick.bat"

try {
	$shortcut = $WshShell.CreateShortcut($shortcutPath)
	$shortcut.TargetPath = $target
	$shortcut.Arguments = $shortcutArgs
	$shortcut.WorkingDirectory = $scriptPath
	$shortcut.IconLocation = 'C:\Windows\System32\cmd.exe,0'
	$shortcut.Save()
	Write-Output "Acceso directo creado: $shortcutPath"
}
catch {
	Write-Error "No se pudo crear el acceso directo: $_"
}
