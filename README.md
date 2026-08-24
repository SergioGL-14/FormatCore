# FormatCore

FormatCore is a Windows script for a very specific case: bringing down the
weight of images and remaking PDFs without depending on external tools like
PDF24 or QPDF.

With images it works as you would expect: it resizes, recompresses and, if
needed, changes the format. With PDFs there is one important catch that
should be clear from the start: it does no structural compression of the
PDF. What it does is render every page as an image and assemble a new PDF
from those images.

That has consequences. You can lose selectable text, search, vectors, layers
and other elements of the original PDF. For scans it usually makes sense.
For manuals, invoices, technical documentation or text-heavy PDFs, not
always. Some may end up worse, or even larger.

## The important part before using it on PDFs

- If you want to keep selectable or searchable text, this is not the best option.
- If the PDF is already well optimized, rasterizing it can make it worse.
- `-PdfMode Auto` tries to avoid those cases with a prior estimate.
- `-PdfMode Skip` ignores PDFs entirely.
- `-PdfMode Raster` forces the remake even when the estimate says it is not worth it.

## How to use it

There are three normal ways:

- Drag files onto the `.exe` if you have it packaged.
- Open the `.exe` or the `.ps1` with no arguments to get the file picker.
- Run it from a console passing paths.

Examples:

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\Docs\report.pdf" "C:\Docs\photo.png"
```

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\Inbox" -Recurse
```

```powershell
powershell -ExecutionPolicy Bypass -File .\FormatCore.ps1 "C:\Docs\scan.pdf" -PdfMode Auto -LogPath "C:\Logs\formatcore.log"
```

## Options that are actually useful

- `-PdfMode Auto|Skip|Raster`: decides what to do with PDFs.
- `-Recurse`: also goes into subfolders when you pass directories.
- `-WhatIf`: shows what it would do without writing anything.
- `-LogPath`: saves a transcript of the run.
- `-OutputDirectory`: sends all results to another folder.
- `-Overwrite`: reuses the same output name. Without it, FormatCore creates `_001`, `_002`, etc. suffixes.
- `-KeepLargerOutput`: keeps results that do not improve on the original size.
- `-OutputFormat Auto|Original|Jpg|Png`: output format for images.
- `-KeepTransparency`: if the image has transparency, avoids ending up in JPEG.
- `-PDFSplitMarginPercent`: makes PDF splitting more conservative or more aggressive.

## How output names are decided

By default it never overwrites previous results. If a file like
`photo_REDUCED.jpg` already exists, the next one will be
`photo_REDUCED_001.jpg`, then `photo_REDUCED_002.jpg`, and so on.

If you prefer that it always replaces, use `-Overwrite`.

## About image formats

FormatCore can work with `jpg`, `jpeg`, `png`, `bmp`, `gif`, `tif`,
`tiff`, `jfif` and `heic`.

Two practical nuances:

- `HEIC` depends on the codec installed in Windows. If the system does not
  have it, the file will not open properly.
- If an image has transparency and the output would end up as JPEG,
  `-KeepTransparency` forces a compatible output, normally `PNG`.

## About PDF

The PDF path relies on `Windows.Data.Pdf`, so it is meant for Windows
PowerShell 5.1 on a compatible Windows.

The actual process:

1. The PDF is opened with `Windows.Data.Pdf`.
2. Every page is rasterized to a JPEG image.
3. A new PDF is rebuilt from those images.
4. If it still exceeds the threshold, it is split into several parts.

Rasterization has no fixed megabyte limit for the input PDF: it depends on
available memory, the number of pages and the size of each rasterized page.
In this implementation the rebuild keeps the PDF and the JPEG bytes in
memory before writing it out, so the practical limit for a generated PDF is
around 2 GB on Windows PowerShell 5.1. A PDF or a single page getting close
to that size can fail on memory even with plenty of disk space; it is not
considered a supported case. `Threshold` uses Int64 so sizes and split
decisions do not overflow near 2 GB, but that does not raise the
rasterization limits of `System.Drawing` or `Windows.Data.Pdf`.

`PDFSplitMarginPercent` does influence the cut. A lower value makes the
script put fewer pages per part, leaving more headroom on the final size.

Keep in mind that final PDF quality depends heavily on the DPI and JPEG
quality chosen for the rasterization.

The repository also carries a small regression check at
`tests/Verify-FileSizeTypes.ps1`. It inspects the source and fails if the
file-size handling ever goes back to Int32 — exactly the overflow this
script hit once near the 2 GB mark. Run it after touching any size-related
code:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\Verify-FileSizeTypes.ps1
```

CI runs the same check on every push.

## Default behavior

- No files passed: opens a file picker.
- Folders passed: processes direct files from the first level only.
- `-Recurse`: also goes into subfolders.
- Result is not smaller and `-KeepLargerOutput` unused: discarded.
- Output name exists and `-Overwrite` unused: creates a new name.

## Requirements

- Windows
- Windows PowerShell 5.1
- `System.Drawing`
- `Windows.Data.Pdf` available on the system

If you package it as an `.exe`, the expected behavior stays the same: drag
files onto it or open it with no arguments to pick them.
