# DWG to PDF Tool

Batch-publish AutoCAD DWG layouts to PDF using PowerShell and `AcCoreConsole.exe`. No AutoLISP or additional PowerShell modules required.

## Files

- `dwg-to-pdf.ps1` — publishing engine.
- `dwg-to-pdf.config.psd1` — user configuration. Edit this file, not the engine.
- `samples/` — sample DWG files for testing.

## Requirements

- Windows and AutoCAD for Windows, including `AcCoreConsole.exe` and PDF publishing support.
- Windows PowerShell 5.1 — the supported and tested environment for this release.
- PowerShell 7 has not been fully tested. Windows PowerShell 5.1 is recommended for reproducible results.

## Setup and run

1. Keep the script and configuration file in the same folder.
2. Edit `Paths.InputFolder`, `Paths.OutputFolder`, and `Paths.AutoCADCoreConsole` in `dwg-to-pdf.config.psd1`.
3. Start with `Publishing.DryRun = $true` to preview the planned operations.
4. Run the script in PowerShell:

   ```powershell
   .\dwg-to-pdf.ps1
   ```

5. Review the planned output paths. To publish PDFs, set `Publishing.DryRun = $false` and run the script again.

If Windows blocks script execution, follow your organization's PowerShell policy. Do not change the machine-wide execution policy just to run this tool.

## Options

- `Publishing.PdfMode`: `SEPARATE` (one PDF per layout) or `COMBINED` (one multi-page PDF per DWG).
- `Publishing.ExistingFileAction`: `OVERWRITE`, `SKIP`, or `ERROR`.
- `Revisions.Source`: `NONE`, `AUTO`, or `FIXED`.
- Naming tokens include `{file}`, `{filename}`, `{layout}`, `{date:yyyyMMdd}`, `{dwgdate:yyyyMMdd}`, `{rev}`, and named groups from `Naming.FileNameRegex`.

This basic version does not read paper format, orientation, or revision values from title-block attributes.

## Temporary files

Working files are created automatically under `%TEMP%\DWG_TO_PDF_WORK`. No manual setup of temporary folders, SCR files, or DSD files is required.

Successful jobs and dry runs remove their per-drawing work folders. Failed jobs may retain temporary files for diagnosis when `Publishing.KeepTempFilesOnError = $true`.