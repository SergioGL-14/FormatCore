# FormatCore

FormatCore es un script para Windows pensado para un caso muy concreto: bajar
el peso de imagenes y rehacer PDFs sin depender de herramientas externas como
PDF24 o QPDF.

Con imagenes funciona como cabria esperar: redimensiona, recomprime y, si hace
falta, cambia de formato. Con PDF hay una pega importante que conviene dejar
clara desde el principio: no hace compresion estructural del PDF. Lo que hace
es renderizar cada pagina como imagen y montar un PDF nuevo a partir de esas
imagenes.

Eso tiene consecuencias. Se puede perder texto seleccionable, busqueda,
vectoriales, capas y otros elementos propios del PDF original. En escaneos suele
tener sentido. En manuales, facturas, documentacion tecnica o PDFs con mucho
texto, no siempre. Algunos pueden quedar peor o incluso crecer.

## Lo importante antes de usarlo con PDFs

- Si quieres conservar texto seleccionable o busqueda, esta no es la mejor opcion.
- Si el PDF ya esta bien optimizado, rasterizarlo puede empeorarlo.
- `-PdfMode Auto` intenta evitar esos casos haciendo una estimacion previa.
- `-PdfMode Skip` ignora los PDFs por completo.
- `-PdfMode Raster` fuerza el rehacer del PDF aunque la estimacion diga que no compensa.

## Como se usa

Hay tres formas normales de usarlo:

- Arrastrar archivos sobre el `.exe` si lo tienes empaquetado.
- Abrir el `.exe` o el `.ps1` sin argumentos para que salga el selector.
- Ejecutarlo por consola pasando rutas.

Ejemplos:

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\Docs\informe.pdf" "C:\Docs\foto.png"
```

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\Entrada" -Recurse
```

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\Docs\scan.pdf" -PdfMode Auto -LogPath "C:\Logs\formatcore.log"
```

## Opciones utiles de verdad

- `-PdfMode Auto|Skip|Raster`: decide que hacer con los PDFs.
- `-Recurse`: entra tambien en subcarpetas cuando pasas directorios.
- `-WhatIf`: muestra lo que haria sin escribir nada.
- `-LogPath`: guarda un transcript de la ejecucion.
- `-OutputDirectory`: manda todos los resultados a otra carpeta.
- `-Overwrite`: reutiliza el mismo nombre de salida. Si no se usa, FormatCore crea sufijos `_001`, `_002`, etc.
- `-KeepLargerOutput`: conserva resultados que no mejoran el tamano original.
- `-OutputFormat Auto|Original|Jpg|Png`: formato de salida para imagenes.
- `-KeepTransparency`: si la imagen tiene transparencia, evita acabar en JPEG.
- `-PDFSplitMarginPercent`: hace el troceado PDF mas conservador o mas agresivo.

## Como decide los nombres de salida

Por defecto no sobrescribe resultados anteriores. Si ya existe un archivo como
`foto_REDUCIDO.jpg`, el siguiente sera `foto_REDUCIDO_001.jpg`, luego
`foto_REDUCIDO_002.jpg`, y asi sucesivamente.

Si prefieres que reemplace siempre, usa `-Overwrite`.

## Sobre formatos de imagen

FormatCore puede trabajar con `jpg`, `jpeg`, `png`, `bmp`, `gif`, `tif`,
`tiff`, `jfif` y `heic`.

Hay dos matices practicos:

- `HEIC` depende del codec instalado en Windows. Si el sistema no lo tiene, el archivo no abrira bien.
- Si una imagen tiene transparencia y la salida acabaria en JPEG, `-KeepTransparency` fuerza una salida compatible, normalmente `PNG`.

## Sobre PDF

La ruta PDF depende de `Windows.Data.Pdf`, asi que esta pensada para Windows
PowerShell 5.1 en un Windows compatible.

El proceso real es este:

1. Se abre el PDF con `Windows.Data.Pdf`.
2. Cada pagina se rasteriza a imagen JPEG.
3. Con esas imagenes se reconstruye un PDF nuevo.
4. Si aun supera el umbral, se intenta dividir en varias partes.

`PDFSplitMarginPercent` si influye en el corte. Un valor mas bajo hace que el
script meta menos paginas por parte para ir mas sobrado con el tamano final.

## Lo que hace por defecto

- Si no le pasas archivos, abre un selector.
- Si le pasas carpetas, procesa archivos directos del primer nivel.
- Si usas `-Recurse`, baja tambien a subcarpetas.
- Si el resultado no mejora y no has usado `-KeepLargerOutput`, lo descarta.
- Si el nombre de salida ya existe y no has usado `-Overwrite`, crea uno nuevo.

## Limitaciones que siguen ahi

- No hace compresion estructural de PDF.
- No es la mejor herramienta para PDF documental con texto o vectoriales.
- `HEIC` no esta garantizado en todos los equipos.
- La calidad final del PDF depende mucho del DPI y de la calidad JPEG que elijas.

## Requisitos

- Windows
- Windows PowerShell 5.1
- `System.Drawing`
- `Windows.Data.Pdf` disponible en el sistema

Si lo empaquetas como `.exe`, el comportamiento esperado sigue siendo el mismo:
arrastrar archivos encima o abrirlo sin argumentos para elegirlos.
