param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments,
 
    [switch]$n  # dry-run
)

# Si no hay argumentos, salir
if ($Arguments.Count -eq 0) {
    Write-Host "Error: No se proporcionaron argumentos."
    Write-Host "Uso: setDate.ps1 archivo1 fecha1 [archivo2 fecha2 ...]"
    Write-Host "Formato de fecha: AAAAMMDDhhmm o AAAAMMDDhhmmss"
    exit 1
}

# Verificar que hay un número par de argumentos
if ($Arguments.Count % 2 -ne 0) {
    Write-Host "Error: Debe proporcionar pares de archivo y fecha."
    Write-Host "Uso: setDate.ps1 archivo1 fecha1 [archivo2 fecha2 ...]"
    exit 1
}

# --- Función para procesar un archivo con su fecha ---
function Process-FileWithDate {
    param(
        [string]$Target,
        [string]$Timestamp
    )

    # Validar formato de fecha
    if ($Timestamp -notmatch '^\d{12}(\d{2})?$') {
        Write-Host "Error: Formato inválido para '$Target'. Use AAAAMMDDhhmm o AAAAMMDDhhmmss"
        return $false
    }

    # Convertir a DateTime
    try {
        if ($Timestamp.Length -eq 12) {
            $Date = [DateTime]::ParseExact($Timestamp, "yyyyMMddHHmm", $null)
        } else {
            $Date = [DateTime]::ParseExact($Timestamp, "yyyyMMddHHmmss", $null)
        }
    } catch {
        Write-Host "Error: No se pudo convertir la fecha para '$Target'."
        return $false
    }

    # Resolver ruta completa
    $FullPath = Resolve-Path $Target -ErrorAction SilentlyContinue
    if (-not $FullPath) {
        Write-Host "Error: Archivo o directorio no existe: '$Target'"
        return $false
    }

    $FullPath = $FullPath.Path

    Write-Host ""
    Write-Host "================================"
    Write-Host "Target: $Target"
    Write-Host "Fecha: $($Date.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Host "================================"

    # --- Detectar si es archivo o carpeta ---
    if (Test-Path $FullPath -PathType Leaf) {
        Process-File -FilePath $FullPath -Date $Date
    }
    elseif (Test-Path $FullPath -PathType Container) {
        Write-Host "Escaneando carpeta: $FullPath"
        Get-ChildItem -Path $FullPath -File -Recurse | ForEach-Object {
            Process-File -FilePath $_.FullName -Date $Date
        }
    }
    else {
        Write-Host "Error: Ruta no válida."
        return $false
    }

    return $true
}

 
# --- Función para procesar un archivo ---
function Process-File {
    param(
        [string]$FilePath,
        [DateTime]$Date
    )
 
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

# --- Procesar todos los pares de archivo/fecha ---
$totalPairs = $Arguments.Count / 2
$successCount = 0
$failCount = 0

Write-Host ""
Write-Host "========================================"
Write-Host "Procesando $totalPairs archivo(s)..."
Write-Host "========================================"

for ($i = 0; $i -lt $Arguments.Count; $i += 2) {
    $target = $Arguments[$i]
    $timestamp = $Arguments[$i + 1]
    
    Write-Host ""
    Write-Host "[$($i/2 + 1)/$totalPairs] Procesando: $target"
    
    if (Process-FileWithDate -Target $target -Timestamp $timestamp) {
        $successCount++
    } else {
        $failCount++
    }
}

Write-Host ""
Write-Host "========================================"
Write-Host "Resumen:"
Write-Host "  Total: $totalPairs"
Write-Host "  Exitosos: $successCount"
Write-Host "  Fallidos: $failCount"
Write-Host "========================================"
Write-Host "Completado."