param(
    [Parameter(Mandatory = $true)]
    [string]$Target,
 
    [Parameter(Mandatory = $true)]
    [string]$Timestamp,  # Formato: AAAAMMDDhhmm[ss]
 
    [switch]$n  # dry-run
)
 
# Validar formato de fecha
if ($Timestamp -notmatch '^\d{12}(\d{2})?$') {
    Write-Host "Error: Formato inválido. Use AAAAMMDDhhmm o AAAAMMDDhhmmss"
    exit 1
}
 
# Convertir a DateTime
try {
    if ($Timestamp.Length -eq 12) {
        $Date = [DateTime]::ParseExact($Timestamp, "yyyyMMddHHmm", $null)
    } else {
        $Date = [DateTime]::ParseExact($Timestamp, "yyyyMMddHHmmss", $null)
    }
} catch {
    Write-Host "Error: No se pudo convertir la fecha."
    exit 1
}
 
# Resolver ruta completa
$FullPath = Resolve-Path $Target -ErrorAction SilentlyContinue
if (-not $FullPath) {
    Write-Host "Error: Archivo o directorio no existe."
    exit 1
}
 
$FullPath = $FullPath.Path
 
# --- Función para procesar un archivo ---
function Process-File {
    param([string]$FilePath)
 
    Write-Host "Procesando: $FilePath"
 
    if ($n) {
        Write-Host "  [DRY RUN] → Se simularía recrear archivo para cambiar ctime"
        Write-Host "  [DRY RUN] → Se aplicaría fecha: $Date"
        return
    }
 
    # Crear archivo temporal
    $tmp = New-TemporaryFile
 
    # Copiar bytes sin tocar metadata
    [System.IO.File]::Copy($FilePath, $tmp.FullName, $true)
 
    # Eliminar archivo original
    Remove-Item -Path $FilePath -Force
 
    # Recrearlo copiando contenido
    [System.IO.File]::Copy($tmp.FullName, $FilePath, $true)
 
    # Cambiar mtime / atime / ctime usando .NET
    $fileInfo = Get-Item $FilePath
    $fileInfo.CreationTime = $Date
    $fileInfo.LastWriteTime = $Date
    $fileInfo.LastAccessTime = $Date
 
    # Cambiar propiedades internas OOXML (para Office: xlsx, docx, pptx)
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    if ($extension -in @('.xlsx', '.xlsm', '.docx', '.docm', '.pptx', '.pptm')) {
        try {
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace([System.IO.Path]::GetDirectoryName($FilePath))
            $item = $folder.ParseName([System.IO.Path]::GetFileName($FilePath))
            
            # Índices de propiedades en Office:
            # 208 = Contenido creado
            # 209 = Guardado el (última modificación)
            $dateString = $Date.ToString("dd/MM/yyyy HH:mm")
            
            # Nota: Shell.Application no permite SET de propiedades extendidas directamente
            # Usamos DSOFile o modificación directa del ZIP/OOXML
            
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zip = [System.IO.Compression.ZipFile]::Open($FilePath, 'Update')
            $coreXml = $zip.Entries | Where-Object { $_.FullName -eq 'docProps/core.xml' }
            
            if ($coreXml) {
                $stream = $coreXml.Open()
                $reader = New-Object System.IO.StreamReader($stream)
                $content = $reader.ReadToEnd()
                $reader.Close()
                
                # Formato ISO 8601 para OOXML
                $isoDate = $Date.ToString("yyyy-MM-ddTHH:mm:ssZ")
                
                # Reemplazar fechas en XML
                $content = $content -replace '<dcterms:created[^>]*>.*?</dcterms:created>', "<dcterms:created xsi:type=`"dcterms:W3CDTF`">$isoDate</dcterms:created>"
                $content = $content -replace '<dcterms:modified[^>]*>.*?</dcterms:modified>', "<dcterms:modified xsi:type=`"dcterms:W3CDTF`">$isoDate</dcterms:modified>"
                
                # Escribir contenido actualizado
                $coreXml.Delete()
                $newEntry = $zip.CreateEntry('docProps/core.xml')
                $writer = New-Object System.IO.StreamWriter($newEntry.Open())
                $writer.Write($content)
                $writer.Close()
            }
            
            $zip.Dispose()
            
            # Reestablecer fechas del sistema (se modifican al editar el ZIP)
            Start-Sleep -Milliseconds 100
            $fileInfo = Get-Item $FilePath
            $fileInfo.CreationTime = $Date
            $fileInfo.LastWriteTime = $Date
            $fileInfo.LastAccessTime = $Date
            
        } catch {
            Write-Host "  Advertencia: No se pudieron modificar las propiedades internas OOXML: $_"
        }
    }
 
    Write-Host "  Fechas aplicadas correctamente."
}
 
# --- Detectar si es archivo o carpeta ---
if (Test-Path $FullPath -PathType Leaf) {
    Process-File -FilePath $FullPath
}
elseif (Test-Path $FullPath -PathType Container) {
    Write-Host "Escaneando carpeta: $FullPath"
    Get-ChildItem -Path $FullPath -File -Recurse | ForEach-Object {
        Process-File $_.FullName
    }
}
else {
    Write-Host "Error: Ruta no válida."
    exit 1
}
 
Write-Host "Completado."