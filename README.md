# FormatCore

Script PowerShell para reducir imagenes y rehacer PDFs usando componentes nativos
de Windows.

El archivo principal actual es `FormatCore.ps1`.

Este README sustituye documentacion antigua con nombres previos y referencias a
PDF24 y QPDF. Esa descripcion ya no corresponde al script actual del proyecto.

## Estado actual del proyecto

- El flujo real de PDF no usa PDF24 ni QPDF.
- El flujo real de PDF usa `Windows.Data.Pdf` para rasterizar cada pagina.
- El flujo de imagen usa `System.Drawing`.
- El script puede ejecutarse con rutas en linea de comandos o sin parametros.
- Si no se pasan archivos, se abre un selector.
- Si el script se empaqueta como EXE, se le pueden arrastrar archivos encima.

## Advertencias clave

- El PDF no se comprime de forma estructural; se rasteriza y se reconstruye.
- Puedes perder texto seleccionable, busqueda, vectoriales y otros elementos del PDF.
- Algunos PDFs pueden quedar peor o incluso mas grandes.
- `PdfMode=Auto` intenta evitar esos casos, pero no puede garantizarlo al 100%.
- La ruta PDF sigue siendo adecuada sobre todo para escaneos o documentos ya basados en imagen.

## Como funciona de verdad

### Imagenes

El script abre la imagen, la copia a memoria, la redimensiona si hace falta y la
guarda de nuevo con compresion JPEG o PNG segun el caso.

Comportamiento actual:

- Por defecto `OutputFormat=Auto`.
- En `Auto`, `jpg` y `jpeg` siguen en JPEG.
- En `Auto`, `png` intenta mantenerse en PNG y solo pasa a JPG si sigue fuera del umbral y no hay transparencia.
- En `Auto`, `bmp`, `gif`, `tif`, `tiff`, `heic` y `jfif` salen como JPG salvo que se detecte transparencia.
- Si una salida JPEG proviene de una imagen con transparencia, se rellena con blanco.
- El script ya permite elegir `-OutputFormat Auto|Original|Jpg|Png`.
- El script ya permite `-KeepTransparency` para forzar una salida compatible con alpha.
- Por defecto, si la salida final no mejora el tamano original, se descarta.
- Si se quiere conservar una salida peor o igual, hay que usar `-KeepLargerOutput`.

### PDF

Este punto es critico: el script no hace "compresion PDF" clasica.

El flujo actual es:

1. Abrir el PDF con `Windows.Data.Pdf`.
2. Si `PdfMode=Auto`, estimar con una muestra si el PDF probablemente va a crecer.
3. Renderizar cada pagina a imagen JPEG.
4. Reconstruir un PDF nuevo a partir de esas imagenes.
5. Si el resultado sigue superando el umbral, dividirlo en varias partes.

Eso significa:

- se pierde texto seleccionable
- se pierde busqueda dentro del PDF
- se pierden vectoriales, capas y estructura nativa del PDF
- formularios, anotaciones y elementos avanzados pueden perderse
- algunos PDFs, sobre todo los de texto o vector, pueden aumentar de tamano

Esta limitacion no es un detalle secundario. Es la caracteristica central del
comportamiento PDF actual.

## Requisitos reales

- Windows
- Windows PowerShell 5.1 recomendado
- `System.Drawing` disponible para imagenes
- `Windows.Data.Pdf` disponible para PDFs

Notas importantes:

- En PowerShell 7 la ruta WinRT para PDF puede fallar segun el entorno.
- Si `Windows.Data.Pdf` no esta disponible, el script no procesara PDFs.
- HEIC depende del codec instalado en Windows. Si falta el codec HEIF/HEIC,
  esos archivos pueden no abrirse.

## Parametros actuales

El script expone estos parametros:

- `Files`
- `Threshold`
- `Quality`
- `MaxWidth`
- `MaxHeight`
- `PDFRenderDPI`
- `OutputFormat`
- `KeepLargerOutput`
- `KeepTransparency`
- `Overwrite`
- `PdfMode`
- `Recurse`
- `OutputDirectory`
- `LogPath`
- `PDFSplitMarginPercent`

Resumen:

- `Files`: archivos o carpetas a procesar.
- `Threshold`: tamano objetivo en bytes.
- `Quality`: calidad JPEG.
- `MaxWidth` y `MaxHeight`: limites de redimensionado para imagenes.
- `PDFRenderDPI`: DPI usados al rasterizar PDFs.
- `OutputFormat`: comportamiento de conversion para imagenes.
- `KeepLargerOutput`: conserva resultados que no mejoran el original.
- `KeepTransparency`: evita acabar en JPEG cuando hay alpha.
- `Overwrite`: reutiliza el nombre de salida exacto y sobrescribe si existe.
- `PdfMode`: `Raster`, `Skip` o `Auto` para decidir la ruta PDF.
- `Recurse`: entra tambien en subcarpetas.
- `OutputDirectory`: guarda resultados en otra carpeta.
- `LogPath`: guarda un transcript de la ejecucion.
- `PDFSplitMarginPercent`: ajusta el objetivo efectivo real del algoritmo de division PDF.

## Carpetas y recursion

Si se pasa una carpeta, el script solo procesa los archivos soportados del primer
nivel de esa carpeta.

Si se usa `-Recurse`, tambien entra en subcarpetas.

## Nombres de salida y colisiones

Las salidas usan sufijos como estos:

- `archivo_REDUCIDO.jpg`
- `archivo_REDUCIDO_001.jpg`
- `archivo_REDUCIDO_bmp.png`
- `archivo_REDUCIDO.pdf`
- `archivo_REDUCIDO_part1.pdf`

Comportamiento actual:

- por defecto no se sobrescribe
- si el nombre ya existe, se crea un sufijo incremental como `_001`
- `-Overwrite` activa la sobrescritura explicita
- `-OutputDirectory` permite sacar los resultados a otra carpeta

## Lo que el usuario puede y no puede decidir hoy

La herramienta ya permite decidir el formato de salida de imagen con:

- `-OutputFormat Auto`
- `-OutputFormat Original`
- `-OutputFormat Jpg`
- `-OutputFormat Png`

Tambien permite decidir si se conservan o no resultados peores con:

- `-KeepLargerOutput`

Tambien permite decidir como tratar PDFs con:

- `-PdfMode Auto`
- `-PdfMode Raster`
- `-PdfMode Skip`

Tambien permite controlar transparencia con:

- `-KeepTransparency`

No existen aun:

- `-ForceJpegForPng`
- un backend de compresion PDF estructural

## Salida por consola

La consola ahora debe dejar claro que:

- el PDF se rasteriza
- puede perder informacion propia del PDF
- puede crecer de tamano
- `PdfMode=Auto` puede omitir PDFs si estima que creceran
- las carpetas se exploran solo en el primer nivel salvo `-Recurse`
- las imagenes peores se descartan salvo que se use `-KeepLargerOutput`
- los resultados existentes no se pisan por defecto
- al final se muestra un resumen por archivo
- `-LogPath` puede guardar el transcript completo

## Uso

Procesar archivos concretos:

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\ruta\documento.pdf" "C:\ruta\imagen.png"
```

Forzar PNG como salida de imagen y usar otra carpeta:

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\ruta\imagen.bmp" -OutputFormat Png -OutputDirectory "C:\ruta\salida"
```

Conservar una imagen aunque no reduzca tamano:

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\ruta\foto.jpg" -KeepLargerOutput
```

Procesar una carpeta completa con subcarpetas y guardar log:

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\ruta\carpeta" -Recurse -LogPath "C:\ruta\ejecucion.log"
```

Simular la ejecucion sin escribir archivos:

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\ruta\carpeta" -Recurse -WhatIf
```

Abrir el selector de archivos:

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1
```

Si esta empaquetado como EXE:

- arrastra archivos al EXE para procesarlos directamente
- si abres el EXE sin argumentos, aparecera el selector de archivos

## Limitaciones conocidas

- La reduccion de PDF depende de rasterizacion, no de compresion estructural.
- Algunos PDFs quedaran peor o mas grandes que el original.
- `PdfMode=Auto` usa una estimacion por muestra; es util, pero no infalible.
- La division PDF usa busqueda adaptativa con tamano real generado y respeta el objetivo efectivo del margen; si una sola pagina no cabe ahi, se conserva como parte unica.
- `LogPath` guarda transcript, no un sistema de logging estructurado por niveles.
- La transparencia puede preservarse evitando JPEG, pero no existe un parametro para recomponer alpha dentro de un JPEG porque el formato no lo soporta.

## Siguiente mejora recomendada

Si se quiere mejorar la reduccion de PDFs sin perder texto ni vectoriales, hace
falta una ruta distinta de la actual. Mientras esa ruta no exista, la documentacion
debe hablar siempre de "rasterizacion y reconstruccion" y no de "compresion PDF"
en sentido clasico.
