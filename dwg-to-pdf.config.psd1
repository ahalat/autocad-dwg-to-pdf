@{
    # =====================================================================
    # DWG TO PDF - USER CONFIGURATION
    # =====================================================================
    # This file configures DWG_TO_PDF.ps1. Edit this file instead of the
    # main script. No extra PowerShell modules are required.
    #
    # AI / LLM INSTRUCTIONS
    # ---------------------
    # This file may be supplied to an AI model with a plain-language request
    # describing a company's drawing naming and publishing standard.
    # The model should:
    # 1. Read all comments and settings before making changes.
    # 2. Change configuration values only; do not rewrite the engine.
    # 3. Preserve all required keys and the hashtable structure.
    # 4. Use only the supported tokens and values listed below.
    # 5. Keep valid PowerShell Data File (.psd1) syntax.
    # 6. Never add executable PowerShell code, functions, or script blocks.
    # 7. If a requested feature is not supported, explain this explicitly.
    # 8. Return the complete updated configuration file, not just a diff.
    # 9. Do not assume that paper format or revision attributes can be read
    #    in this basic version; these features are not implemented here.
    #
    # SUPPORTED NAMING TOKENS
    # -----------------------
    # {file}                  DWG filename without extension
    # {filename}              DWG filename including extension
    # {layout}                Layout name; SEPARATE mode only
    # {date:yyyyMMdd}         Publishing date, using .NET date format syntax
    # {dwgdate:yyyyMMdd}      DWG file modification date
    # {rev}                   Revision, formatted by Revisions.Style
    # {rev:00}                Numeric revision formatting, e.g. 01, 02
    # {rev:alpha}             Alphabetic revision: 1=A, 2=B, 27=AA
    # {rev:roman}             Roman numeral revision: 1=I, 2=II
    #
    # Named groups from Naming.FileNameRegex are also available as tokens.
    # Example filename: PROJECT42_STR_1234.dwg
    # Regex: ^(?<project>[^_]+)_(?<discipline>[^_]+)_(?<number>[^_]+)$
    # Tokens: {project}, {discipline}, {number}
    #
    # EXAMPLE PATTERNS
    # ----------------
    # Separate PDFs:
    # {project}-{discipline}-{number}-{layout}-{date:yyyyMMdd}-REV{rev:00}
    # Combined PDF:
    # {file}-{date:yyyy-MM-dd}-REV{rev:00}
    # Alphabetic revision:
    # {project}-{number}-{layout}-REV{rev:alpha}
    # Revision folder:
    # REV{rev:00}
    # Nested revision folder:
    # {project}\{file}\REV{rev:00}
    #
    # REVISION OPTIONS
    # ----------------
    # Revisions.Source:
    #   NONE  - no revision is generated; remove {rev} from active patterns.
    #   AUTO  - find the next unused revision based on expected output paths.
    #   FIXED - use Revisions.FixedValue.
    #
    # Revisions.Scope applies to AUTO:
    #   PER_DRAWING - each DWG has its own revision shared by its layouts.
    #   PER_PACKAGE - all DWGs in this run share one revision.
    #   PER_LAYOUT  - each layout gets its own revision; SEPARATE mode only.
    #
    # IMPORTANT LIMITATIONS OF THIS BASIC VERSION
    # -------------------------------------------
    # - The tool discovers layout names through AcCoreConsole command output.
    # - It does not read paper size or orientation from layout metadata.
    # - It does not read a revision value from title-block attributes/text.
    # - Available revision sources are NONE, AUTO and FIXED only.
    # - It does not implement arbitrary conditional rules.
    #
    # SAFETY
    # ------
    # DryRun is TRUE by default. Inspect the planned paths before publishing.
    # AUTO mode requires {rev} in the active file pattern or FolderPattern.
    # =====================================================================

    ConfigVersion = 1

    Paths = @{
        # Folder containing input DWG files.
        InputFolder = "C:\DWG_TO_PDF\IN"

        # Root folder for generated PDFs.
        OutputFolder = "C:\DWG_TO_PDF\OUT"

        # Path to AutoCAD Core Console.
        AutoCADCoreConsole = "C:\Program Files\Autodesk\AutoCAD 2027\accoreconsole.exe"

        # Include DWG files in child folders.
        SearchSubfolders = $false
    }

    Publishing = @{
        # SEPARATE: one PDF per layout.
        # COMBINED: one multi-page PDF per DWG.
        PdfMode = "SEPARATE"

        # OVERWRITE: replace existing target PDFs.
        # SKIP: leave existing targets in place and skip them.
        # ERROR: stop processing a DWG if a target path exists.
        ExistingFileAction = "OVERWRITE"

        # TRUE previews planned output paths without publishing.
        # Change to FALSE only after checking the preview.
        DryRun = $true

        # Maximum number of candidate revisions checked in AUTO mode.
        MaxAutoRevision = 999

        # Keep temporary SCR/DSD files if a scan or publish fails.
        KeepTempFilesOnError = $true
    }

    Naming = @{
        # Regex is applied to the DWG filename WITHOUT the extension.
        # Empty means only {file} and {filename} are available by default.
        FileNameRegex = ''

        # Used when PdfMode = SEPARATE. Must uniquely identify each layout.
        SeparatePattern = "{file}-{layout}-REV{rev:00}"

        # Used when PdfMode = COMBINED. Do not use {layout} here.
        CombinedPattern = "{file}-REV{rev:00}"

        # Default .NET date format used by {date} and {dwgdate}.
        DefaultDateFormat = "yyyyMMdd"
    }

    Revisions = @{
        # NONE, AUTO, or FIXED.
        Source = "AUTO"

        # Used by AUTO: PER_DRAWING, PER_PACKAGE, or PER_LAYOUT.
        Scope = "PER_DRAWING"

        # Used by AUTO: NUMERIC, ALPHA, or ROMAN.
        Style = "NUMERIC"
        Start = 1

        # Numeric formatting examples: "0", "00", "000".
        NumericFormat = "00"

        # Used when Source = FIXED.
        FixedValue = "A"

        # Empty means no revision subfolder.
        # Example: "REV{rev:00}"
        # Example: "{file}\REV{rev:00}"
        # FolderPattern must be relative to OutputFolder.
        FolderPattern = ""
    }
}
