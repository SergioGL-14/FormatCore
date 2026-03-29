<#
.SYNOPSIS
Reduce imagenes y rehace PDFs usando componentes nativos de Windows.

.DESCRIPTION
Procesa imagenes y PDFs desde PowerShell, por arrastrar y soltar sobre un EXE
empaquetado o abriendo un selector cuando no se pasan rutas.

La ruta PDF actual no usa PDF24 ni QPDF. Usa Windows.Data.Pdf para renderizar
cada pagina del PDF a JPEG y despues reconstruye un PDF nuevo.

Consecuencias de este enfoque:
- el PDF no se comprime de forma estructural; se rasteriza y se reconstruye
- se pierde texto seleccionable y busqueda
- se pierden vectoriales, capas y contenido nativo del PDF
- algunos PDFs de texto o vector pueden crecer en tamano

Si el PDF rasterizado sigue superando el umbral, el script intenta dividirlo en
varias partes.

.PARAMETER Files
Lista de archivos o carpetas a procesar. Si se pasan carpetas, por defecto solo
se toman los archivos del primer nivel. Usa -Recurse para incluir subcarpetas.

Si no se pasan archivos, se abre un cuadro de seleccion.

.PARAMETER Threshold
Tamano objetivo en bytes para el resultado o para cada parte de un PDF dividido.

.PARAMETER Quality
Calidad JPEG usada para imagenes y para las paginas rasterizadas de PDF.

.PARAMETER MaxWidth
Ancho maximo para el redimensionado de imagenes.

.PARAMETER MaxHeight
Alto maximo para el redimensionado de imagenes.

.PARAMETER PDFRenderDPI
Resolucion usada al rasterizar cada pagina PDF antes de reconstruir el archivo.

.PARAMETER OutputFormat
Formato de salida para imagenes: Auto, Original, Jpg o Png.

.PARAMETER KeepLargerOutput
Conserva resultados que no mejoran el tamano original.

.PARAMETER Overwrite
Sobrescribe resultados existentes en lugar de crear un sufijo incremental.

.PARAMETER OutputDirectory
Carpeta de salida opcional. Si no se indica, se usa la carpeta del archivo de entrada.

.PARAMETER PDFSplitMarginPercent
Margen aplicado al objetivo efectivo por parte al dividir PDFs.

.PARAMETER PdfMode
Controla el tratamiento de PDFs: Raster, Skip o Auto.

.PARAMETER Recurse
Procesa tambien subcarpetas cuando se pasan directorios de entrada.

.PARAMETER LogPath
Guarda un transcript de la ejecucion en un archivo de log.

.PARAMETER KeepTransparency
Si la imagen tiene transparencia, evita acabar en JPEG y fuerza una salida compatible.

.NOTES
Archivo actual: FormatCore.ps1
PDFs: requieren Windows.Data.Pdf disponible en Windows.
HEIC: depende del codec instalado en el sistema. Si el codec no esta presente,
el archivo puede no abrirse con System.Drawing.
Si ya existe un resultado con ese nombre, por defecto se crea un sufijo incremental.
Si una imagen procesada no mejora el tamano original, por defecto se descarta.
Si `PdfMode=Auto`, el script intenta estimar si un PDF va a crecer antes de rasterizarlo.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\Docs\informe.pdf" "C:\Docs\foto.png"

.EXAMPLE
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Position = 0)]
    [string[]]$Files,
    [ValidateRange(1024, 2147483647)][int]$Threshold = 1048576,
    [ValidateRange(1, 100)][int]$Quality = 80,
    [ValidateRange(1, 12000)][int]$MaxWidth = 1920,
    [ValidateRange(1, 12000)][int]$MaxHeight = 1080,
    [ValidateRange(72, 600)][int]$PDFRenderDPI = 150,
    [ValidateSet('Auto','Original','Jpg','Png')][string]$OutputFormat = 'Auto',
    [switch]$KeepLargerOutput,
    [switch]$KeepTransparency,
    [switch]$Overwrite,
    [ValidateSet('Raster','Skip','Auto')][string]$PdfMode = 'Auto',
    [switch]$Recurse,
    [string]$OutputDirectory,
    [string]$LogPath,
    [ValidateRange(50, 100)][int]$PDFSplitMarginPercent = 92
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$script:WinRTAvailable = $false
$script:SupportedExtensions = @('.pdf','.jpg','.jpeg','.png','.bmp','.gif','.tif','.tiff','.heic','.jfif')
$script:Results = New-Object System.Collections.Generic.List[object]
$script:TranscriptStarted = $false
$script:TranscriptPath = $null
$script:ScriptCmdlet = $PSCmdlet
if ((-not $Files -or $Files.Count -eq 0) -and $args.Count -gt 0) { $Files = @($args) }

function Get-JpegCodec {
    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
    if (-not $codec) { throw 'No se encontro el codec JPEG de System.Drawing.' }
    $codec
}

function Format-Size {
    param([Nullable[long]]$Bytes)
    if ($null -eq $Bytes) { return '' }
    if ($Bytes -lt 1024) { return "$Bytes B" }
    if ($Bytes -lt 1MB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    return ('{0:N2} MB' -f ($Bytes / 1MB))
}

function Ensure-Directory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Add-ProcessingResult {
    param(
        [string]$InputPath,
        [string]$Type,
        [string]$Status,
        [string]$Output = '',
        [string]$Format = '',
        [Nullable[long]]$OriginalBytes = $null,
        [Nullable[long]]$ResultBytes = $null,
        [string]$Note = ''
    )
    $script:Results.Add([PSCustomObject]@{
        Archivo = Split-Path $InputPath -Leaf
        Tipo = $Type
        Estado = $Status
        Salida = $Output
        Formato = $Format
        Original = Format-Size -Bytes $OriginalBytes
        Resultado = Format-Size -Bytes $ResultBytes
        Nota = $Note
    }) | Out-Null
}

function Show-ProcessingSummary {
    if ($script:Results.Count -eq 0) { return }
    Write-Host ''
    Write-Host 'Resumen final:' -ForegroundColor Cyan
    $table = $script:Results | Format-Table Archivo,Tipo,Estado,Salida,Formato,Original,Resultado,Nota -AutoSize -Wrap | Out-String
    Write-Host $table.TrimEnd()
}

function Start-ProcessingLog {
    if ([string]::IsNullOrWhiteSpace($LogPath)) { return }
    $resolvedLogPath = [System.IO.Path]::GetFullPath($LogPath)
    $resolvedLogPath = Get-UniqueOutputPath -DesiredPath $resolvedLogPath -AllowOverwrite:$Overwrite
    if ($WhatIfPreference) {
        Write-Host "WhatIf: el log se guardaria en $resolvedLogPath" -ForegroundColor DarkYellow
        $script:TranscriptPath = $resolvedLogPath
        return
    }
    Ensure-Directory -Path ([System.IO.Path]::GetDirectoryName($resolvedLogPath))
    Start-Transcript -Path $resolvedLogPath -Force | Out-Null
    $script:TranscriptStarted = $true
    $script:TranscriptPath = $resolvedLogPath
    Write-Host "Log: $resolvedLogPath" -ForegroundColor DarkGray
}

function Stop-ProcessingLog {
    if ($script:TranscriptStarted) {
        Stop-Transcript | Out-Null
        $script:TranscriptStarted = $false
    }
}

function Resolve-InputFiles {
    param([string[]]$Paths)
    $resolved = New-Object System.Collections.Generic.List[string]
    foreach ($rawPath in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($rawPath)) { continue }
        $candidate = $rawPath.Trim().Trim('"')
        if (-not (Test-Path -LiteralPath $candidate)) { continue }
        $item = Get-Item -LiteralPath $candidate
        if ($item.PSIsContainer) {
            Get-ChildItem -LiteralPath $item.FullName -File -Recurse:$Recurse | Where-Object { $script:SupportedExtensions -contains $_.Extension.ToLower() } | Sort-Object FullName | ForEach-Object { $resolved.Add($_.FullName) }
            continue
        }
        if ($script:SupportedExtensions -contains $item.Extension.ToLower()) { $resolved.Add($item.FullName) }
    }
    @($resolved.ToArray() | Select-Object -Unique)
}

function Initialize-WindowsForms {
    Add-Type -AssemblyName System.Windows.Forms
}

function Get-CanonicalExtension {
    param([string]$Extension)
    $safeExtension = if ($null -eq $Extension) { '' } else { $Extension.ToLower() }
    switch ($safeExtension) {
        '.jpeg' { '.jpg' }
        '.jfif' { '.jpg' }
        '.tif' { '.tiff' }
        default { $safeExtension }
    }
}

function Resolve-OutputDirectoryPath {
    param([string]$SourcePath)
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        return [System.IO.Path]::GetDirectoryName($SourcePath)
    }
    [System.IO.Path]::GetFullPath($OutputDirectory)
}

function Get-UniqueOutputPath {
    param(
        [string]$DesiredPath,
        [switch]$AllowOverwrite
    )
    $fullPath = [System.IO.Path]::GetFullPath($DesiredPath)
    if ($AllowOverwrite -or -not (Test-Path -LiteralPath $fullPath)) { return $fullPath }
    $dir = [System.IO.Path]::GetDirectoryName($fullPath)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
    $ext = [System.IO.Path]::GetExtension($fullPath)
    $index = 1
    do {
        $candidate = Join-Path $dir ("{0}_{1:D3}{2}" -f $base, $index, $ext)
        $index++
    } while (Test-Path -LiteralPath $candidate)
    $candidate
}

function Get-PreferredImageExtension {
    param(
        [string]$SourceExtension,
        [string]$Mode = 'Auto',
        [bool]$HasTransparency = $false
    )
    if ($HasTransparency -and ($KeepTransparency -or $Mode -eq 'Auto')) {
        switch ($Mode) {
            'Original' {
                switch ($SourceExtension.ToLower()) {
                    '.png' { return '.png' }
                    '.gif' { return '.gif' }
                    '.tif' { return '.tif' }
                    '.tiff' { return '.tiff' }
                    default { return '.png' }
                }
            }
            'Png' { return '.png' }
            default { return '.png' }
        }
    }
    switch ($Mode) {
        'Jpg' { return '.jpg' }
        'Png' { return '.png' }
        'Original' {
            switch ($SourceExtension.ToLower()) {
                '.heic' { return '.jpg' }
                '.jfif' { return '.jpg' }
                default { return $SourceExtension.ToLower() }
            }
        }
        default {
            switch ($SourceExtension.ToLower()) {
                '.jpg' { return '.jpg' }
                '.jpeg' { return '.jpeg' }
                '.png' { return '.png' }
                default { return '.jpg' }
            }
        }
    }
}

function Test-ImageHasTransparency {
    param([System.Drawing.Image]$Image)
    if ($null -eq $Image) { return $false }
    $bitmap = $null
    $ownsBitmap = $false
    try {
        $pixelFormatValue = [int]$Image.PixelFormat
        $alphaFlags = [int][System.Drawing.Imaging.PixelFormat]::Alpha -bor [int][System.Drawing.Imaging.PixelFormat]::PAlpha
        if ($Image -is [System.Drawing.Bitmap]) {
            $bitmap = $Image
        } else {
            $bitmap = New-Object System.Drawing.Bitmap($Image)
            $ownsBitmap = $true
        }
        if ($bitmap -and $bitmap.Palette) {
            foreach ($entry in $bitmap.Palette.Entries) {
                if ($entry.A -lt 255) { return $true }
            }
        }
        if (($pixelFormatValue -band $alphaFlags) -ne 0) {
            $stepX = [Math]::Max(1, [int][Math]::Floor($bitmap.Width / 128))
            $stepY = [Math]::Max(1, [int][Math]::Floor($bitmap.Height / 128))
            for ($y = 0; $y -lt $bitmap.Height; $y += $stepY) {
                for ($x = 0; $x -lt $bitmap.Width; $x += $stepX) {
                    if ($bitmap.GetPixel($x, $y).A -lt 255) { return $true }
                }
            }
        }
    } catch {
        return $false
    } finally {
        if ($ownsBitmap -and $null -ne $bitmap -and $bitmap -is [System.IDisposable]) { $bitmap.Dispose() }
    }
    $false
}

function Load-ImageUnlocked {
    param([string]$Path)
    $fileStream = $null; $memoryStream = $null; $sourceImage = $null; $clone = $null
    try {
        $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $memoryStream = New-Object System.IO.MemoryStream
        $fileStream.CopyTo($memoryStream)
        $memoryStream.Position = 0
        $sourceImage = [System.Drawing.Image]::FromStream($memoryStream, $true, $true)
        $clone = New-Object System.Drawing.Bitmap($sourceImage)
        $clone
    } finally {
        if ($null -ne $sourceImage -and $sourceImage -is [System.IDisposable]) { $sourceImage.Dispose() }
        if ($null -ne $memoryStream -and $memoryStream -is [System.IDisposable]) { $memoryStream.Dispose() }
        if ($null -ne $fileStream -and $fileStream -is [System.IDisposable]) { $fileStream.Dispose() }
    }
}

function New-FlatBitmap {
    param([System.Drawing.Image]$Image)
    $flat = New-Object System.Drawing.Bitmap($Image.Width, $Image.Height)
    $graphics = $null
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($flat)
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.DrawImage($Image, 0, 0)
        $flat
    } finally {
        if ($null -ne $graphics -and $graphics -is [System.IDisposable]) { $graphics.Dispose() }
    }
}

function Save-BitmapToPath {
    param(
        [System.Drawing.Image]$Image,
        [string]$OutputPath,
        [int]$Quality = 80
    )
    $extension = [System.IO.Path]::GetExtension($OutputPath).ToLower()
    Ensure-Directory -Path ([System.IO.Path]::GetDirectoryName($OutputPath))
    switch ($extension) {
        '.jpg' { }
        '.jpeg' { }
        '.png' {
            $Image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
            return
        }
        '.bmp' {
            $Image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Bmp)
            return
        }
        '.gif' {
            $Image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Gif)
            return
        }
        '.tif' { 
            $Image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Tiff)
            return
        }
        '.tiff' {
            $Image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Tiff)
            return
        }
        default {
            throw "Formato de salida no soportado: $extension"
        }
    }
    $jpegCodec = Get-JpegCodec
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality,[long]$Quality)
    $Image.Save($OutputPath, $jpegCodec, $encoderParams)
}

function Initialize-WinRT {
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Write-Warning 'PowerShell 7 detectado. La ruta PDF con Windows.Data.Pdf esta orientada a Windows PowerShell 5.1 y puede fallar.'
    }
    if ([Environment]::OSVersion.Version.Major -lt 10) {
        Write-Warning 'Windows.Data.Pdf requiere una version moderna de Windows. Windows 10 o superior es lo esperado.'
    }
    try {
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        $script:WinRTAvailable = $true
        Write-Host '  Windows.Data.Pdf disponible (Windows 10+)' -ForegroundColor Green
    } catch {
        $script:WinRTAvailable = $false
        Write-Warning 'Windows.Data.Pdf no disponible. Los PDFs no podran procesarse en este entorno.'
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            Write-Warning 'Prueba con Windows PowerShell 5.1 si necesitas procesar PDFs.'
        }
    }
}

function Invoke-WinRTOp {
    param($AsyncOp,[Type]$ResultType)
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 } | Select-Object -First 1
    if (-not $method) { throw 'No se encontro AsTask para IAsyncOperation.' }
    $task = $method.MakeGenericMethod($ResultType).Invoke($null, @($AsyncOp))
    $task.Wait()
    $task.Result
}

function Invoke-WinRTAction {
    param($AsyncAction)
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and -not $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 } | Select-Object -First 1
    if (-not $method) { throw 'No se encontro AsTask para IAsyncAction.' }
    $task = $method.Invoke($null, @($AsyncAction))
    $task.Wait()
}

function ConvertTo-NetStream {
    param($WinRTStream)
    [System.IO.WindowsRuntimeStreamExtensions]::AsStream([Windows.Storage.Streams.IRandomAccessStream]$WinRTStream)
}

function Test-PDFReadable {
    param([string]$PdfPath)
    if (-not $script:WinRTAvailable) { return $true }
    $pdfDoc = $null
    try {
        $storageFile = Invoke-WinRTOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync([System.IO.Path]::GetFullPath($PdfPath))) ([Windows.Storage.StorageFile])
        $pdfDoc = Invoke-WinRTOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])
        return ($pdfDoc.PageCount -gt 0)
    } catch {
        return $false
    } finally {
        if ($null -ne $pdfDoc -and $pdfDoc -is [System.IDisposable]) { $pdfDoc.Dispose() }
    }
}

function Get-PDFPageCountNative {
    param([string]$PdfPath)
    $pdfDoc = $null
    try {
        $storageFile = Invoke-WinRTOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync([System.IO.Path]::GetFullPath($PdfPath))) ([Windows.Storage.StorageFile])
        $pdfDoc = Invoke-WinRTOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])
        [int]$pdfDoc.PageCount
    } finally {
        if ($null -ne $pdfDoc -and $pdfDoc -is [System.IDisposable]) { $pdfDoc.Dispose() }
    }
}

function Get-PDFSampleIndices {
    param([int]$PageCount,[int]$MaxSamples = 3)
    if ($PageCount -le 0) { return @() }
    if ($PageCount -le $MaxSamples) { return @(0..($PageCount - 1)) }
    $indices = New-Object System.Collections.Generic.List[int]
    $indices.Add(0)
    $indices.Add([int][Math]::Floor(($PageCount - 1) / 2))
    $indices.Add($PageCount - 1)
    @($indices | Select-Object -Unique | Sort-Object)
}

function Estimate-PDFRasterSize {
    param([string]$PdfPath,[int]$RenderDPI = 150,[int]$JpegQuality = 80)
    if (-not $script:WinRTAvailable) { return $null }
    $pageCount = Get-PDFPageCountNative -PdfPath $PdfPath
    if ($pageCount -le 0) { return $null }
    $sampleIndices = Get-PDFSampleIndices -PageCount $pageCount
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("pdfestimate_{0}" -f ([Guid]::NewGuid().ToString('N')))
    Ensure-Directory -Path $tempDir
    try {
        $samplePages = @(Get-PDFPageImages -PdfPath $PdfPath -TempDir $tempDir -RenderDPI $RenderDPI -JpegQuality $JpegQuality -PageIndices $sampleIndices -Quiet)
        if (-not $samplePages -or $samplePages.Count -eq 0) { return $null }
        $sizes = @($samplePages | ForEach-Object { (Get-Item -LiteralPath $_).Length })
        $averageBytes = [long][Math]::Round((($sizes | Measure-Object -Sum).Sum) / [double]$sizes.Count)
        $projectedBytes = [long][Math]::Round(($averageBytes * $pageCount) + [Math]::Max(4096, $pageCount * 256))
        [PSCustomObject]@{
            PageCount = $pageCount
            SampleCount = $samplePages.Count
            AveragePageBytes = $averageBytes
            ProjectedBytes = $projectedBytes
        }
    } finally {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-OutputImagePath {
    param(
        [string]$FilePath,
        [bool]$HasTransparency = $false
    )
    $dir = Resolve-OutputDirectoryPath -SourcePath $FilePath
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $sourceExtension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    $targetExtension = Get-PreferredImageExtension -SourceExtension $sourceExtension -Mode $OutputFormat -HasTransparency:$HasTransparency
    $sourceTag = $sourceExtension.TrimStart('.')
    $useSourceTag = (Get-CanonicalExtension -Extension $sourceExtension) -ne (Get-CanonicalExtension -Extension $targetExtension)
    $leaf = if ($useSourceTag) {
        "${base}_REDUCIDO_${sourceTag}${targetExtension}"
    } else {
        "${base}_REDUCIDO${targetExtension}"
    }
    Get-UniqueOutputPath -DesiredPath (Join-Path $dir $leaf) -AllowOverwrite:$Overwrite
}

function Compress-ImageNative {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [int]$Quality = 80,
        [int]$MaxWidth = 1920,
        [int]$MaxHeight = 1080,
        [int]$Threshold = 1048576,
        [switch]$AllowJpegFallback,
        [bool]$HasTransparency = $false
    )
    if (-not (Test-Path -LiteralPath $InputPath)) { throw "Archivo no encontrado: $InputPath" }
    $src = $null; $resized = $null; $flattened = $null; $graphics = $null; $resultPath = $null
    $originalSize = (Get-Item -LiteralPath $InputPath).Length
    try {
        $src = Load-ImageUnlocked -Path $InputPath
        if (-not $HasTransparency) { $HasTransparency = Test-ImageHasTransparency -Image $src }
        $ratio = [Math]::Min([Math]::Min($MaxWidth / [double]$src.Width, $MaxHeight / [double]$src.Height), 1.0)
        $newWidth = [int][Math]::Max(1, [Math]::Round($src.Width * $ratio))
        $newHeight = [int][Math]::Max(1, [Math]::Round($src.Height * $ratio))
        $resized = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
        $graphics = [System.Drawing.Graphics]::FromImage($resized)
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.DrawImage($src, 0, 0, $newWidth, $newHeight)
        $graphics.Dispose(); $graphics = $null
        $resultPath = $OutputPath
        $saveExtension = [System.IO.Path]::GetExtension($resultPath).ToLower()
        if ($HasTransparency -and $saveExtension -in @('.jpg','.jpeg')) {
            $resultPath = Get-UniqueOutputPath -DesiredPath ([System.IO.Path]::ChangeExtension($resultPath, '.png')) -AllowOverwrite:$Overwrite
            $saveExtension = '.png'
            Write-Host '  Transparencia detectada. Se fuerza una salida PNG para no perder alpha.' -ForegroundColor DarkYellow
        }
        $saveBitmap = $resized
        $currentQuality = $Quality
        $minQuality = 30
        $currentSize = 0
        $note = ''

        if ($saveExtension -in @('.jpg','.jpeg')) {
            $flattened = New-FlatBitmap -Image $resized
            $saveBitmap = $flattened
            while ($currentQuality -ge $minQuality) {
                if (Test-Path -LiteralPath $resultPath) { Remove-Item -LiteralPath $resultPath -Force }
                Save-BitmapToPath -Image $saveBitmap -OutputPath $resultPath -Quality $currentQuality
                $currentSize = (Get-Item -LiteralPath $resultPath).Length
                Write-Host "  Calidad $currentQuality%: $([Math]::Round($currentSize / 1024, 1)) KB" -ForegroundColor Cyan
                if ($currentSize -le $Threshold) { break }
                $currentQuality -= 10
            }
        } else {
            if (Test-Path -LiteralPath $resultPath) { Remove-Item -LiteralPath $resultPath -Force }
            Save-BitmapToPath -Image $saveBitmap -OutputPath $resultPath -Quality $Quality
            $currentSize = (Get-Item -LiteralPath $resultPath).Length
            Write-Host "  Guardado $saveExtension`: $([Math]::Round($currentSize / 1024, 1)) KB" -ForegroundColor Cyan
            if ($currentSize -gt $Threshold -and $AllowJpegFallback -and -not $HasTransparency) {
                Write-Host '  El PNG sigue superando el umbral. Convirtiendo a JPG...' -ForegroundColor Yellow
                Remove-Item -LiteralPath $resultPath -Force
                $resultPath = Get-UniqueOutputPath -DesiredPath ([System.IO.Path]::ChangeExtension($resultPath, '.jpg')) -AllowOverwrite:$Overwrite
                if ($null -eq $flattened) { $flattened = New-FlatBitmap -Image $resized }
                $saveBitmap = $flattened
                $saveExtension = '.jpg'
                $currentQuality = $Quality
                while ($currentQuality -ge $minQuality) {
                    if (Test-Path -LiteralPath $resultPath) { Remove-Item -LiteralPath $resultPath -Force }
                    Save-BitmapToPath -Image $saveBitmap -OutputPath $resultPath -Quality $currentQuality
                    $currentSize = (Get-Item -LiteralPath $resultPath).Length
                    Write-Host "  Calidad $currentQuality%: $([Math]::Round($currentSize / 1024, 1)) KB" -ForegroundColor Cyan
                    if ($currentSize -le $Threshold) { break }
                    $currentQuality -= 10
                }
            }
        }

        if ($currentSize -gt $Threshold) {
            Write-Warning '  No se pudo bajar del umbral con el formato/configuracion actual.'
            $note = 'Supera el umbral objetivo.'
        } else {
            Write-Host "  OK: $(Split-Path $resultPath -Leaf)" -ForegroundColor Green
        }

        if (-not $KeepLargerOutput -and $currentSize -ge $originalSize) {
            if (Test-Path -LiteralPath $resultPath) { Remove-Item -LiteralPath $resultPath -Force }
            Write-Warning '  Se descarta la salida porque no mejora el tamano original. Usa -KeepLargerOutput para conservarla.'
            return [PSCustomObject]@{
                Status = 'Descartado'
                OutputPaths = @()
                OutputSizeBytes = $currentSize
                FinalFormat = $saveExtension
                Note = 'No mejora el tamano original.'
            }
        }

        [PSCustomObject]@{
            Status = 'Reducido'
            OutputPaths = @($resultPath)
            OutputSizeBytes = $currentSize
            FinalFormat = $saveExtension
            Note = $note
        }
    } catch {
        if ([System.IO.Path]::GetExtension($InputPath).ToLower() -eq '.heic') {
            throw "Error comprimiendo imagen HEIC: $($_.Exception.Message). Verifica que el codec HEIC/HEIF este instalado en Windows."
        } else {
            throw "Error comprimiendo imagen: $($_.Exception.Message)"
        }
    } finally {
        foreach ($resource in @($graphics, $flattened, $resized, $src)) { if ($null -ne $resource -and $resource -is [System.IDisposable]) { $resource.Dispose() } }
    }
}

function Get-PDFPageImages {
    param(
        [string]$PdfPath,
        [string]$TempDir,
        [int]$RenderDPI = 150,
        [int]$JpegQuality = 80,
        [int[]]$PageIndices,
        [switch]$Quiet
    )
    $absPath = [System.IO.Path]::GetFullPath($PdfPath)
    $jpegCodec = Get-JpegCodec
    $pdfDoc = $null
    try {
        $storageFile = Invoke-WinRTOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync($absPath)) ([Windows.Storage.StorageFile])
        $pdfDoc = Invoke-WinRTOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])
    } catch {
        Write-Warning "No se pudo abrir el PDF: $($_.Exception.Message)"
        return @()
    }
    try {
        $pageCount = $pdfDoc.PageCount
        if ($pageCount -le 0) { return @() }
        $selectedIndices = if ($PageIndices -and $PageIndices.Count -gt 0) {
            @($PageIndices | Where-Object { $_ -ge 0 -and $_ -lt $pageCount } | Select-Object -Unique | Sort-Object)
        } else {
            @(0..($pageCount - 1))
        }
        if (-not $selectedIndices -or $selectedIndices.Count -eq 0) { return @() }
        if (-not $Quiet) {
            $scopeLabel = if ($selectedIndices.Count -eq $pageCount) { "$pageCount pagina(s)" } else { "$($selectedIndices.Count)/$pageCount pagina(s) (muestra)" }
            Write-Host "  PDF cargado: $scopeLabel | ${RenderDPI} DPI | Calidad $JpegQuality%" -ForegroundColor Cyan
        }
        $jpegPaths = @()
        foreach ($pageIndex in $selectedIndices) {
            $page = $null; $ras = $null; $netStream = $null; $bitmap = $null; $rgb = $null; $graphics = $null
            try {
                $page = $pdfDoc.GetPage([uint32]$pageIndex)
                $targetWidth = [uint32]([Math]::Max(1, [Math]::Round($page.Size.Width * $RenderDPI / 96)))
                $targetHeight = [uint32]([Math]::Max(1, [Math]::Round($page.Size.Height * $RenderDPI / 96)))
                $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
                $options.DestinationWidth = $targetWidth
                $options.DestinationHeight = $targetHeight
                $ras = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
                Invoke-WinRTAction ($page.RenderToStreamAsync($ras, $options))
                $netStream = ConvertTo-NetStream $ras
                if ($netStream.CanSeek) { $netStream.Position = 0 } else {
                    $seekable = New-Object System.IO.MemoryStream
                    $netStream.CopyTo($seekable)
                    $seekable.Position = 0
                    $netStream.Dispose()
                    $netStream = $seekable
                }
                $bitmap = [System.Drawing.Bitmap]::FromStream($netStream)
                $rgb = New-Object System.Drawing.Bitmap($bitmap.Width, $bitmap.Height)
                $graphics = [System.Drawing.Graphics]::FromImage($rgb)
                $graphics.Clear([System.Drawing.Color]::White)
                $graphics.DrawImage($bitmap, 0, 0)
                $graphics.Dispose(); $graphics = $null
                $jpegPath = Join-Path $TempDir ('page_{0:D4}.jpg' -f $pageIndex)
                $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
                $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality,[long]$JpegQuality)
                $rgb.Save($jpegPath, $jpegCodec, $encoderParams)
                if (-not $Quiet) {
                    $sizeKB = [Math]::Round((Get-Item -LiteralPath $jpegPath).Length / 1024, 1)
                    Write-Host "  Pag. $($pageIndex + 1)/$pageCount -> $sizeKB KB" -ForegroundColor Gray
                }
                $jpegPaths += $jpegPath
            } catch {
                Write-Warning "Error en pagina $($pageIndex + 1): $($_.Exception.Message)"
            } finally {
                foreach ($resource in @($graphics, $rgb, $bitmap, $netStream, $ras, $page)) { if ($null -ne $resource -and $resource -is [System.IDisposable]) { $resource.Dispose() } }
            }
        }
        @($jpegPaths)
    } finally {
        if ($null -ne $pdfDoc -and $pdfDoc -is [System.IDisposable]) { $pdfDoc.Dispose() }
    }
}

function New-PDFFromJpegs {
    param([string[]]$JpegPaths,[string]$OutputPath,[int]$RenderDPI = 150)
    if (-not $JpegPaths -or $JpegPaths.Count -eq 0) { return $false }
    $encoding = [System.Text.Encoding]::GetEncoding('ISO-8859-1')
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $buffer = New-Object System.IO.MemoryStream
    try {
        Ensure-Directory -Path ([System.IO.Path]::GetDirectoryName($OutputPath))
        function Write-PdfText([string]$Text) { $bytes = $encoding.GetBytes($Text); $buffer.Write($bytes, 0, $bytes.Length) }
        function Write-PdfBytes([byte[]]$Bytes) { $buffer.Write($Bytes, 0, $Bytes.Length) }
        $pages = @(foreach ($jpegPath in $JpegPaths) {
            $img = $null
            try {
                $img = [System.Drawing.Image]::FromFile($jpegPath)
                [PSCustomObject]@{
                    Bytes = [System.IO.File]::ReadAllBytes($jpegPath)
                    PxWidth = $img.Width
                    PxHeight = $img.Height
                    PtWidth = ([Math]::Round($img.Width * 72.0 / $RenderDPI, 3)).ToString($culture)
                    PtHeight = ([Math]::Round($img.Height * 72.0 / $RenderDPI, 3)).ToString($culture)
                }
            } finally {
                if ($null -ne $img) { $img.Dispose() }
            }
        })
        $pageCount = @($pages).Count
        $offsets = @{}
        Write-PdfText "%PDF-1.4`r`n"
        $buffer.WriteByte(0x25); @(0xE2,0xE3,0xCF,0xD3) | ForEach-Object { $buffer.WriteByte($_) }; Write-PdfText "`r`n"
        $offsets[1] = $buffer.Position
        Write-PdfText "1 0 obj`r`n<< /Type /Catalog /Pages 2 0 R >>`r`nendobj`r`n"
        $offsets[2] = $buffer.Position
        $kids = (0..($pageCount - 1) | ForEach-Object { "$(3 + $_) 0 R" }) -join ' '
        Write-PdfText "2 0 obj`r`n<< /Type /Pages /Kids [$kids] /Count $pageCount >>`r`nendobj`r`n"
        for ($i = 0; $i -lt $pageCount; $i++) {
            $pageObject = 3 + $i; $imageObject = 3 + $pageCount + $i; $contentObject = 3 + ($pageCount * 2) + $i; $pageData = $pages[$i]
            $offsets[$pageObject] = $buffer.Position
            Write-PdfText "$pageObject 0 obj`r`n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 $($pageData.PtWidth) $($pageData.PtHeight)] /Resources << /ProcSet [/PDF /ImageC] /XObject << /Im$i $imageObject 0 R >> >> /Contents $contentObject 0 R >>`r`nendobj`r`n"
            $offsets[$imageObject] = $buffer.Position
            Write-PdfText "$imageObject 0 obj`r`n<< /Type /XObject /Subtype /Image /Width $($pageData.PxWidth) /Height $($pageData.PxHeight) /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length $($pageData.Bytes.Length) >>`r`nstream`r`n"
            Write-PdfBytes $pageData.Bytes
            Write-PdfText "`r`nendstream`r`nendobj`r`n"
            $content = "q $($pageData.PtWidth) 0 0 $($pageData.PtHeight) 0 0 cm /Im$i Do Q"
            $contentBytes = $encoding.GetBytes($content)
            $offsets[$contentObject] = $buffer.Position
            Write-PdfText "$contentObject 0 obj`r`n<< /Length $($contentBytes.Length) >>`r`nstream`r`n"
            Write-PdfBytes $contentBytes
            Write-PdfText "`r`nendstream`r`nendobj`r`n"
        }
        $xrefPosition = $buffer.Position
        $totalObjects = (3 * $pageCount) + 2
        Write-PdfText "xref`r`n0 $($totalObjects + 1)`r`n"
        Write-PdfText "0000000000 65535 f`r`n"
        for ($objectNumber = 1; $objectNumber -le $totalObjects; $objectNumber++) { Write-PdfText "$($offsets[$objectNumber].ToString('D10')) 00000 n`r`n" }
        Write-PdfText "trailer`r`n<< /Size $($totalObjects + 1) /Root 1 0 R >>`r`nstartxref`r`n$xrefPosition`r`n%%EOF`r`n"
        [System.IO.File]::WriteAllBytes($OutputPath, $buffer.ToArray())
        $true
    } catch {
        Write-Warning "No se pudo ensamblar el PDF: $($_.Exception.Message)"
        $false
    } finally {
        $buffer.Dispose()
    }
}

function Compress-PDFNative {
    param([string]$InputPDF,[string]$OutputPDF,[int]$RenderDPI = 150,[int]$JpegQuality = 80)
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("pdfcomp_{0}" -f ([Guid]::NewGuid().ToString('N')))
    Ensure-Directory -Path $tempDir
    try {
        $pages = @(Get-PDFPageImages -PdfPath $InputPDF -TempDir $tempDir -RenderDPI $RenderDPI -JpegQuality $JpegQuality)
        if (-not $pages -or $pages.Count -eq 0) { return $false }
        $ok = New-PDFFromJpegs -JpegPaths $pages -OutputPath $OutputPDF -RenderDPI $RenderDPI
        if (-not $ok -or -not (Test-Path -LiteralPath $OutputPDF)) { return $false }
        if (-not (Test-PDFReadable -PdfPath $OutputPDF)) {
            Write-Warning 'El PDF generado no se pudo volver a abrir con Windows.Data.Pdf.'
            Remove-Item -LiteralPath $OutputPDF -Force -ErrorAction SilentlyContinue
            return $false
        }
        $true
    } finally {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Measure-PDFChunkActual {
    param(
        [string[]]$JpegPaths,
        [string]$TempDir,
        [int]$RenderDPI = 150
    )
    $measurePath = Join-Path $TempDir ("chunk_{0}.pdf" -f ([Guid]::NewGuid().ToString('N')))
    if (-not (New-PDFFromJpegs -JpegPaths $JpegPaths -OutputPath $measurePath -RenderDPI $RenderDPI)) { return $null }
    if (-not (Test-PDFReadable -PdfPath $measurePath)) {
        Remove-Item -LiteralPath $measurePath -Force -ErrorAction SilentlyContinue
        return $null
    }
    [PSCustomObject]@{
        Path = $measurePath
        Size = (Get-Item -LiteralPath $measurePath).Length
    }
}

function Split-PDFByThreshold {
    param([string]$InputPDF,[string]$BaseOutputName,[string]$OutputDir,[int]$Threshold = 1048576,[int]$RenderDPI = 150,[int]$JpegQuality = 80)
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("pdfsplit_{0}" -f ([Guid]::NewGuid().ToString('N')))
    Ensure-Directory -Path $tempDir
    try {
        $allPages = @(Get-PDFPageImages -PdfPath $InputPDF -TempDir $tempDir -RenderDPI $RenderDPI -JpegQuality $JpegQuality)
        if (-not $allPages -or $allPages.Count -eq 0) { return @() }
        $pageInfos = @($allPages | ForEach-Object {
            [PSCustomObject]@{
                Path = $_
                Size = (Get-Item -LiteralPath $_).Length
            }
        })
        $targetBudget = [Math]::Max(1, [int][Math]::Floor($Threshold * ($PDFSplitMarginPercent / 100.0)))
        Write-Host "  $($pageInfos.Count) paginas | margen de division: $PDFSplitMarginPercent% | objetivo efectivo: $([Math]::Round($targetBudget / 1024, 1)) KB | umbral maximo: $([Math]::Round($Threshold / 1024, 1)) KB" -ForegroundColor Gray
        $outputFiles = @(); $part = 1; $start = 0
        while ($start -lt $pageInfos.Count) {
            $remainingCount = $pageInfos.Count - $start
            $bestMeasurement = $null
            $bestChunkSize = 0
            $low = 1
            $high = $remainingCount
            while ($low -le $high) {
                $mid = [int][Math]::Floor(($low + $high) / 2)
                $end = $start + $mid - 1
                $chunkInfo = @($pageInfos[$start..$end])
                $chunk = @($chunkInfo | ForEach-Object { $_.Path })
                $estimatedKB = [Math]::Round((($chunkInfo | Measure-Object -Property Size -Sum).Sum) / 1024, 1)
                $limitKB = [Math]::Round($targetBudget / 1024, 1)
                Write-Host "  Parte ${part}: prueba $mid pagina(s) | estimado $estimatedKB KB | limite efectivo $limitKB KB..." -ForegroundColor Cyan
                $measurement = Measure-PDFChunkActual -JpegPaths $chunk -TempDir $tempDir -RenderDPI $RenderDPI
                if (-not $measurement) { Write-Warning "No se pudo medir una parte PDF."; return $outputFiles }
                if ($measurement.Size -le $targetBudget) {
                    if ($null -ne $bestMeasurement -and (Test-Path -LiteralPath $bestMeasurement.Path)) {
                        Remove-Item -LiteralPath $bestMeasurement.Path -Force -ErrorAction SilentlyContinue
                    }
                    $bestMeasurement = $measurement
                    $bestChunkSize = $mid
                    $low = $mid + 1
                } else {
                    Remove-Item -LiteralPath $measurement.Path -Force -ErrorAction SilentlyContinue
                    $high = $mid - 1
                }
            }

            if ($bestChunkSize -eq 0) {
                $measurement = Measure-PDFChunkActual -JpegPaths @($pageInfos[$start].Path) -TempDir $tempDir -RenderDPI $RenderDPI
                if (-not $measurement) { Write-Warning "No se pudo crear la pagina individual para la parte $part."; return $outputFiles }
                $bestMeasurement = $measurement
                $bestChunkSize = 1
            }

            $finalEnd = $start + $bestChunkSize - 1
            $finalPath = Get-UniqueOutputPath -DesiredPath (Join-Path $OutputDir ("{0}_part{1}.pdf" -f $BaseOutputName, $part)) -AllowOverwrite:$Overwrite
            Ensure-Directory -Path ([System.IO.Path]::GetDirectoryName($finalPath))
            Move-Item -LiteralPath $bestMeasurement.Path -Destination $finalPath -Force
            $sizeKB = [Math]::Round($bestMeasurement.Size / 1024, 1)
            if ($bestMeasurement.Size -le $targetBudget) {
                Write-Host "  OK Parte ${part}: paginas $($start + 1)-$($finalEnd + 1) | $sizeKB KB" -ForegroundColor Green
            } elseif ($bestChunkSize -eq 1 -and $bestMeasurement.Size -le $Threshold) {
                Write-Warning "  La pagina $($start + 1) no cabe en el objetivo con margen ($([Math]::Round($targetBudget / 1024, 1)) KB), pero se conserva como parte unica dentro del umbral maximo."
            } else {
                Write-Warning "  La parte $part tiene una sola pagina y aun supera el umbral ($sizeKB KB)."
            }
            $outputFiles += [PSCustomObject]@{ Path = $finalPath; Size = $bestMeasurement.Size }
            $start = $finalEnd + 1
            $part++
        }
        $outputFiles
    } finally {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Test-PDFRasterizationDecision {
    param([string]$FilePath,[int]$OriginalSize,[int]$RenderDPI = 150,[int]$JpegQuality = 80)
    switch ($PdfMode) {
        'Skip' {
            return [PSCustomObject]@{
                ShouldRasterize = $false
                Note = 'Omitido por PdfMode=Skip.'
                Estimate = $null
            }
        }
        'Raster' {
            return [PSCustomObject]@{
                ShouldRasterize = $true
                Note = 'Rasterizacion forzada por PdfMode=Raster.'
                Estimate = $null
            }
        }
        default {
            $estimate = Estimate-PDFRasterSize -PdfPath $FilePath -RenderDPI $RenderDPI -JpegQuality $JpegQuality
            if ($null -eq $estimate) {
                return [PSCustomObject]@{
                    ShouldRasterize = $false
                    Note = 'No se pudo estimar de forma fiable el resultado rasterizado.'
                    Estimate = $null
                }
            }
            Write-Host "  Estimacion previa PDF: ~$([Math]::Round($estimate.ProjectedBytes / 1024, 1)) KB usando $($estimate.SampleCount) pagina(s) de muestra." -ForegroundColor DarkGray
            if ($estimate.ProjectedBytes -gt [long][Math]::Round($OriginalSize * 1.05)) {
                return [PSCustomObject]@{
                    ShouldRasterize = $false
                    Note = "La estimacion apunta a crecimiento respecto al original ($([Math]::Round((($estimate.ProjectedBytes - $OriginalSize) / [double]$OriginalSize) * 100, 1))%)."
                    Estimate = $estimate
                }
            }
            return [PSCustomObject]@{
                ShouldRasterize = $true
                Note = 'Estimacion favorable para rasterizar.'
                Estimate = $estimate
            }
        }
    }
}

function Handle-PDF {
    param([string]$FilePath,[int]$Threshold = 1048576,[int]$RenderDPI = 150,[int]$JpegQuality = 80)
    $originalSize = (Get-Item -LiteralPath $FilePath).Length
    if ($originalSize -le $Threshold) {
        Write-Host '  El PDF ya esta dentro del umbral. No se genera copia.' -ForegroundColor Green
        return [PSCustomObject]@{
            Status = 'Sin cambios'
            OutputPaths = @()
            OutputSizeBytes = $originalSize
            Note = 'Ya estaba dentro del umbral.'
        }
    }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $directory = Resolve-OutputDirectoryPath -SourcePath $FilePath
    Ensure-Directory -Path $directory
    $outputPDF = Get-UniqueOutputPath -DesiredPath (Join-Path $directory "${baseName}_REDUCIDO.pdf") -AllowOverwrite:$Overwrite
    $splitBaseName = [System.IO.Path]::GetFileNameWithoutExtension($outputPDF)
    Write-Warning '  Este flujo rasteriza cada pagina del PDF a JPEG y reconstruye un PDF nuevo.'
    Write-Warning '  Se pierden texto seleccionable, busqueda y vectoriales.'
    Write-Host "  Rasterizando PDF (DPI: $RenderDPI | Calidad JPEG: $JpegQuality%)..." -ForegroundColor Cyan
    $ok = Compress-PDFNative -InputPDF $FilePath -OutputPDF $outputPDF -RenderDPI $RenderDPI -JpegQuality $JpegQuality
    if (-not $ok -or -not (Test-Path -LiteralPath $outputPDF)) {
        Write-Warning 'Fallo la conversion nativa del PDF.'
        return [PSCustomObject]@{
            Status = 'Error'
            OutputPaths = @()
            OutputSizeBytes = $null
            Note = 'Fallo la conversion nativa del PDF.'
        }
    }
    $compressedSize = (Get-Item -LiteralPath $outputPDF).Length
    $note = ''
    if ($compressedSize -lt $originalSize) {
        $reduction = [Math]::Round((($originalSize - $compressedSize) / $originalSize) * 100, 1)
        Write-Host "  PDF generado: $([Math]::Round($compressedSize / 1024, 1)) KB | reduccion: $reduction%" -ForegroundColor Gray
    } elseif ($compressedSize -gt $originalSize) {
        $increase = [Math]::Round((($compressedSize - $originalSize) / $originalSize) * 100, 1)
        Write-Warning "  El PDF rasterizado es mas grande que el original (+$increase%)."
        $note = "El PDF rasterizado crece +$increase%."
    } else {
        Write-Host "  PDF generado: $([Math]::Round($compressedSize / 1024, 1)) KB" -ForegroundColor Gray
    }
    if ($compressedSize -le $Threshold) {
        Write-Host '  OK: PDF dentro del umbral.' -ForegroundColor Green
        return [PSCustomObject]@{
            Status = 'Reducido'
            OutputPaths = @($outputPDF)
            OutputSizeBytes = $compressedSize
            Note = $note
        }
    }
    Write-Host "  El PDF supera el umbral ($([Math]::Round($Threshold / 1024, 1)) KB). Dividiendo..." -ForegroundColor Yellow
    $parts = Split-PDFByThreshold -InputPDF $FilePath -BaseOutputName $splitBaseName -OutputDir $directory -Threshold $Threshold -RenderDPI $RenderDPI -JpegQuality $JpegQuality
    if ($parts.Count -gt 0 -and (Test-Path -LiteralPath $outputPDF)) { Remove-Item -LiteralPath $outputPDF -Force; Write-Host "  Eliminado intermedio: $(Split-Path $outputPDF -Leaf)" -ForegroundColor Gray }
    if ($parts.Count -gt 0) {
        return [PSCustomObject]@{
            Status = 'Dividido'
            OutputPaths = @($parts | ForEach-Object { $_.Path })
            OutputSizeBytes = [long](($parts | Measure-Object -Property Size -Sum).Sum)
            Note = "Se generaron $($parts.Count) parte(s)."
        }
    }
    [PSCustomObject]@{
        Status = 'Reducido'
        OutputPaths = @($outputPDF)
        OutputSizeBytes = $compressedSize
        Note = if ([string]::IsNullOrWhiteSpace($note)) { 'Supera el umbral y no se pudo dividir.' } else { "$note No se pudo dividir mejor." }
    }
}

function Process-File {
    param([string]$FilePath)
    if (-not (Test-Path -LiteralPath $FilePath)) { throw "No encontrado: $FilePath" }
    $resolvedPath = (Resolve-Path -LiteralPath $FilePath).Path
    $extension = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
    $originalSize = (Get-Item -LiteralPath $resolvedPath).Length
    Write-Host "  Original: $([Math]::Round($originalSize / 1024, 1)) KB" -ForegroundColor Gray
    switch ($extension) {
        '.pdf' {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
            $plannedPdf = Get-UniqueOutputPath -DesiredPath (Join-Path (Resolve-OutputDirectoryPath -SourcePath $resolvedPath) "${baseName}_REDUCIDO.pdf") -AllowOverwrite:$Overwrite
            if (-not $script:WinRTAvailable) {
                Write-Warning 'Windows.Data.Pdf no esta disponible. Este PDF no puede procesarse aqui.'
                Add-ProcessingResult -InputPath $resolvedPath -Type 'PDF' -Status 'Omitido' -OriginalBytes $originalSize -Note 'Windows.Data.Pdf no disponible.'
                return
            }
            if ($originalSize -le $Threshold) {
                Write-Host '  El PDF ya esta dentro del umbral. No se genera copia.' -ForegroundColor Green
                Add-ProcessingResult -InputPath $resolvedPath -Type 'PDF' -Status 'Sin cambios' -OriginalBytes $originalSize -ResultBytes $originalSize -Note 'Ya estaba dentro del umbral.'
                return
            }
            if (-not $script:ScriptCmdlet.ShouldProcess($resolvedPath, "Procesar PDF en modo $PdfMode")) {
                Add-ProcessingResult -InputPath $resolvedPath -Type 'PDF' -Status 'Simulado' -Output (Split-Path $plannedPdf -Leaf) -Format '.pdf' -OriginalBytes $originalSize -Note "Modo $PdfMode. Simulacion sin escritura."
                return
            }
            $decision = Test-PDFRasterizationDecision -FilePath $resolvedPath -OriginalSize $originalSize -RenderDPI $PDFRenderDPI -JpegQuality $Quality
            if (-not $decision.ShouldRasterize) {
                Write-Warning "  PDF omitido: $($decision.Note)"
                Add-ProcessingResult -InputPath $resolvedPath -Type 'PDF' -Status 'Omitido' -OriginalBytes $originalSize -Note $decision.Note
                return
            }
            $pdfResult = Handle-PDF -FilePath $resolvedPath -Threshold $Threshold -RenderDPI $PDFRenderDPI -JpegQuality $Quality
            Add-ProcessingResult -InputPath $resolvedPath -Type 'PDF' -Status $pdfResult.Status -Output (($pdfResult.OutputPaths | ForEach-Object { Split-Path $_ -Leaf }) -join ', ') -Format '.pdf' -OriginalBytes $originalSize -ResultBytes $pdfResult.OutputSizeBytes -Note $pdfResult.Note
            return
        }
        { $_ -in @('.jpg','.jpeg','.png','.bmp','.gif','.tif','.tiff','.heic','.jfif') } {
            $probeImage = $null
            $hasTransparency = $extension -in @('.png','.gif','.bmp','.tif','.tiff','.heic')
            if ($hasTransparency) {
                try {
                    $probeImage = Load-ImageUnlocked -Path $resolvedPath
                    $hasTransparency = Test-ImageHasTransparency -Image $probeImage
                } catch {
                    $hasTransparency = $false
                } finally {
                    if ($null -ne $probeImage -and $probeImage -is [System.IDisposable]) { $probeImage.Dispose() }
                }
            }
            if ($extension -eq '.heic') {
                Write-Host '  HEIC depende del codec del sistema. Si falla la carga, instala el codec HEIF/HEIC correspondiente.' -ForegroundColor DarkYellow
            }
            $outputPath = Get-OutputImagePath -FilePath $resolvedPath -HasTransparency:$hasTransparency
            $plannedFormat = [System.IO.Path]::GetExtension($outputPath).ToLower()
            if (-not $script:ScriptCmdlet.ShouldProcess($resolvedPath, "Procesar imagen a $plannedFormat")) {
                Add-ProcessingResult -InputPath $resolvedPath -Type 'Imagen' -Status 'Simulado' -Output (Split-Path $outputPath -Leaf) -Format $plannedFormat -OriginalBytes $originalSize -Note 'Simulacion sin escritura.'
                return
            }
            $allowJpegFallback = (($OutputFormat -eq 'Auto') -and ($plannedFormat -eq '.png') -and -not $hasTransparency)
            $imageResult = Compress-ImageNative -InputPath $resolvedPath -OutputPath $outputPath -Quality $Quality -MaxWidth $MaxWidth -MaxHeight $MaxHeight -Threshold $Threshold -AllowJpegFallback:$allowJpegFallback -HasTransparency:$hasTransparency
            if ($imageResult.OutputPaths.Count -gt 0) {
                $resultPath = $imageResult.OutputPaths[0]
                $newSize = $imageResult.OutputSizeBytes
                if ($newSize -lt $originalSize) {
                    $reduction = [Math]::Round((($originalSize - $newSize) / $originalSize) * 100, 1)
                    Write-Host "  Resultado: $(Split-Path $resultPath -Leaf) | $([Math]::Round($newSize / 1024, 1)) KB | reduccion: $reduction%" -ForegroundColor Gray
                } else {
                    Write-Warning "  Resultado: $(Split-Path $resultPath -Leaf) | $([Math]::Round($newSize / 1024, 1)) KB"
                }
                $newExtension = [System.IO.Path]::GetExtension($resultPath).ToLower()
                if ($newExtension -ne $extension) { Write-Host "  Formato de salida: $newExtension" -ForegroundColor Gray }
            }
            Add-ProcessingResult -InputPath $resolvedPath -Type 'Imagen' -Status $imageResult.Status -Output (($imageResult.OutputPaths | ForEach-Object { Split-Path $_ -Leaf }) -join ', ') -Format $imageResult.FinalFormat -OriginalBytes $originalSize -ResultBytes $imageResult.OutputSizeBytes -Note $imageResult.Note
            return
        }
        default {
            Write-Warning "Formato no soportado: $extension"
            Add-ProcessingResult -InputPath $resolvedPath -Type 'Otro' -Status 'Omitido' -OriginalBytes $originalSize -Note "Formato no soportado: $extension"
            return
        }
    }
}

function Main {
    try {
        Start-ProcessingLog
        Write-Host ''
        Write-Host '--------------------------------------------------------' -ForegroundColor DarkCyan
        Write-Host '  FormatCore' -ForegroundColor Cyan
        Write-Host '  Reduccion y conversion nativa para Windows' -ForegroundColor Gray
        Write-Host '  Arrastra archivos al EXE o abre el EXE/PS1 para elegirlos' -ForegroundColor Gray
        Write-Host '  PDF: rasterizacion por pagina a JPEG; no conserva texto ni vectoriales' -ForegroundColor DarkYellow
        Write-Host "  Carpetas: $(if ($Recurse) { 'recursivo' } else { 'solo primer nivel' }) | PdfMode: $PdfMode | Salida imagen: $OutputFormat" -ForegroundColor Gray
        Write-Host "  Umbral: $([Math]::Round($Threshold / 1024, 1)) KB | Calidad: $Quality% | PDF DPI: $PDFRenderDPI | Margen PDF: $PDFSplitMarginPercent%" -ForegroundColor Gray
        Write-Host "  Overwrite: $($Overwrite.IsPresent) | KeepTransparency: $($KeepTransparency.IsPresent) | WhatIf: $WhatIfPreference" -ForegroundColor Gray
        if (-not $KeepLargerOutput) { Write-Host '  Salidas peores de imagen se descartan automaticamente' -ForegroundColor Gray }
        Write-Host '--------------------------------------------------------' -ForegroundColor DarkCyan
        Write-Host ''
        Initialize-WinRT
        $filesToProcess = @()
        if ($Files -and $Files.Count -gt 0) {
            $missingCount = 0
            foreach ($rawPath in @($Files)) {
                if ([string]::IsNullOrWhiteSpace($rawPath)) { continue }
                $candidate = $rawPath.Trim().Trim('"')
                if (-not (Test-Path -LiteralPath $candidate)) { $missingCount++ }
            }
            $filesToProcess = Resolve-InputFiles -Paths $Files
            if ($missingCount -gt 0) { Write-Warning "$missingCount ruta(s) no encontrada(s)." }
        } else {
            Initialize-WindowsForms
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = 'Selecciona archivos para comprimir'
            $dialog.Filter = 'Archivos soportados|*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff;*.heic;*.jfif|Todos|*.*'
            $dialog.Multiselect = $true
            $dialog.CheckFileExists = $true
            $dialog.RestoreDirectory = $true
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $filesToProcess = Resolve-InputFiles -Paths $dialog.FileNames
            } else {
                Write-Host 'Cancelado por el usuario.' -ForegroundColor Yellow
                return
            }
        }
        if (-not $filesToProcess -or $filesToProcess.Count -eq 0) {
            Write-Host 'No hay archivos validos para procesar.' -ForegroundColor Yellow
            return
        }
        Write-Host "Procesando $($filesToProcess.Count) archivo(s)...`n" -ForegroundColor White
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($file in $filesToProcess) {
            Write-Host "--- $(Split-Path $file -Leaf) --------------------------------" -ForegroundColor White
            try {
                Process-File -FilePath $file
            } catch {
                Write-Warning "Error procesando $file`: $($_.Exception.Message)"
                $resolved = if (Test-Path -LiteralPath $file) { (Resolve-Path -LiteralPath $file).Path } else { $file }
                $size = if (Test-Path -LiteralPath $file) { (Get-Item -LiteralPath $file).Length } else { $null }
                $type = if ([System.IO.Path]::GetExtension($file).ToLower() -eq '.pdf') { 'PDF' } else { 'Imagen' }
                Add-ProcessingResult -InputPath $resolved -Type $type -Status 'Error' -OriginalBytes $size -Note $_.Exception.Message
            }
            Write-Host ''
        }
        $stopwatch.Stop()
        Show-ProcessingSummary
        Write-Host '--------------------------------------------------------' -ForegroundColor DarkCyan
        Write-Host "  Completado en $([Math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) segundos." -ForegroundColor Green
        if ($script:TranscriptPath) { Write-Host "  Log: $script:TranscriptPath" -ForegroundColor DarkGray }
    } finally {
        Stop-ProcessingLog
    }
}

Main
