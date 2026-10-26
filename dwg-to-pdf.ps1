[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$NoPause
)

# Resolve the default config path after parameter binding.
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptDirectory)) {
    $scriptDirectory = (Get-Location).Path
}
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $scriptDirectory "DWG_TO_PDF.config.psd1"
}

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Configuration file was not found: $ConfigPath"
}

$script:Config = Import-PowerShellDataFile -Path $ConfigPath
$script:PublishDate = Get-Date
$script:Culture = [System.Globalization.CultureInfo]::InvariantCulture
$script:ACCORE = [string]$script:Config.Paths.AutoCADCoreConsole
$script:OutputRoot = [System.IO.Path]::GetFullPath([string]$script:Config.Paths.OutputFolder)
if ($script:OutputRoot.Length -gt 3) { $script:OutputRoot = $script:OutputRoot.TrimEnd('\') }

function Assert-Config {
    $cfg = $script:Config

    foreach ($key in @('InputFolder', 'OutputFolder', 'AutoCADCoreConsole')) {
        if ([string]::IsNullOrWhiteSpace([string]$cfg.Paths[$key])) {
            throw "Configuration value Paths.$key is required."
        }
    }

    if (!(Test-Path -LiteralPath $cfg.Paths.InputFolder -PathType Container)) {
        throw "Input folder was not found: $($cfg.Paths.InputFolder)"
    }
    if (!(Test-Path -LiteralPath $cfg.Paths.AutoCADCoreConsole -PathType Leaf)) {
        throw "AcCoreConsole.exe was not found: $($cfg.Paths.AutoCADCoreConsole)"
    }

    if ($cfg.Publishing.PdfMode -notin @('SEPARATE', 'COMBINED')) {
        throw "Publishing.PdfMode must be SEPARATE or COMBINED."
    }
    if ($cfg.Publishing.ExistingFileAction -notin @('OVERWRITE', 'SKIP', 'ERROR')) {
        throw "Publishing.ExistingFileAction must be OVERWRITE, SKIP, or ERROR."
    }
    if ($cfg.Revisions.Source -notin @('NONE', 'AUTO', 'FIXED')) {
        throw "Revisions.Source must be NONE, AUTO, or FIXED."
    }
    if ($cfg.Revisions.Scope -notin @('PER_DRAWING', 'PER_PACKAGE', 'PER_LAYOUT')) {
        throw "Revisions.Scope must be PER_DRAWING, PER_PACKAGE, or PER_LAYOUT."
    }
    if ($cfg.Revisions.Style -notin @('NUMERIC', 'ALPHA', 'ROMAN')) {
        throw "Revisions.Style must be NUMERIC, ALPHA, or ROMAN."
    }
    if ($cfg.Publishing.PdfMode -eq 'COMBINED' -and $cfg.Revisions.Source -eq 'AUTO' -and $cfg.Revisions.Scope -eq 'PER_LAYOUT') {
        throw "PER_LAYOUT automatic revisions require PdfMode = SEPARATE."
    }

    if (![string]::IsNullOrWhiteSpace([string]$cfg.Naming.FileNameRegex)) {
        try { $rx = [regex]::new([string]$cfg.Naming.FileNameRegex) }
        catch { throw "Naming.FileNameRegex is invalid: $($_.Exception.Message)" }

        $reserved = @('file', 'filename', 'layout', 'date', 'dwgdate', 'rev')
        foreach ($groupName in $rx.GetGroupNames()) {
            if ($groupName -notmatch '^\d+$' -and $reserved -contains $groupName.ToLowerInvariant()) {
                throw "Regex group '$groupName' conflicts with a reserved token name."
            }
        }
    }

    $activePattern = if ($cfg.Publishing.PdfMode -eq 'SEPARATE') {
        [string]$cfg.Naming.SeparatePattern
    } else {
        [string]$cfg.Naming.CombinedPattern
    }

    if ($cfg.Revisions.Source -eq 'NONE') {
        if (($activePattern -match '(?i)\{rev(?::[^{}]+)?\}') -or
            ([string]$cfg.Revisions.FolderPattern -match '(?i)\{rev(?::[^{}]+)?\}')) {
            throw "Revisions.Source is NONE, but an active pattern contains a {rev} token. Remove the token."
        }
    }

    if ($cfg.Revisions.Source -eq 'AUTO') {
        $revInName = $activePattern -match '(?i)\{rev(?::[^{}]+)?\}'
        $revInFolder = ([string]$cfg.Revisions.FolderPattern) -match '(?i)\{rev(?::[^{}]+)?\}'
        if (!$revInName -and !$revInFolder) {
            throw "AUTO revision mode requires a {rev} token in the active naming pattern or Revisions.FolderPattern."
        }
    }

    if ($cfg.Publishing.PdfMode -eq 'COMBINED') {
        if ($activePattern -match '(?i)\{layout(?::[^{}]+)?\}') {
            throw "CombinedPattern cannot use {layout}; one PDF contains multiple layouts."
        }
        if ([string]$cfg.Revisions.FolderPattern -match '(?i)\{layout(?::[^{}]+)?\}') {
            throw "Revisions.FolderPattern cannot use {layout} in COMBINED mode."
        }
    }

    if ([int]$cfg.Publishing.MaxAutoRevision -lt 1) {
        throw "Publishing.MaxAutoRevision must be greater than zero."
    }
    if ([int]$cfg.Revisions.Start -lt 1) {
        throw "Revisions.Start must be greater than zero."
    }
}

function Invoke-LayoutScan {
    param(
        [System.IO.FileInfo]$Dwg,
        [string]$JobDir
    )

    $scrPath = Join-Path $JobDir 'SCAN_LAYOUTS.scr'
    @(
        '_.-LAYOUT'
        '?'
    ) | Set-Content -Path $scrPath -Encoding ASCII

    Push-Location $JobDir
    try {
        $scanOutput = @(
            & $script:ACCORE /i $Dwg.FullName /s $scrPath /l en-US 2>&1 |
                ForEach-Object { [string]$_ }
        )
        $exitCode = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    if ($exitCode -ne 0) {
        Write-Host 'Layout scan output:' -ForegroundColor Yellow
        $scanOutput | Select-Object -Last 25 | ForEach-Object { Write-Host $_ }
        throw "Core Console layout scan failed for $($Dwg.Name) with exit code $exitCode."
    }

    $layouts = @()
    foreach ($line in $scanOutput) {
        $clean = [string]$line -replace "`0", ''
        $clean = [regex]::Replace($clean, '[\x01-\x08\x0B\x0C\x0E-\x1F]', '')

        if ($clean -notmatch 'Paper_Space') { continue }

        # Parse the label before the block metadata without relying on UI language.
        $match = [regex]::Match($clean, '^\s*[^:]+:\s*(?<name>.*?)\s{2,}.*Paper_Space')

        if ($match.Success) {
            $name = $match.Groups['name'].Value.Trim()
            if ($name -and $name -ine 'Model') {
                $layouts += [pscustomobject]@{ Name = $name }
            }
        }
    }

    $layouts = @($layouts | Group-Object -Property Name | ForEach-Object { $_.Group[0] })

    if ($layouts.Count -eq 0) {
        Write-Host 'Layout scan output:' -ForegroundColor Yellow
        $scanOutput | Where-Object { $_ -match 'Paper_Space' } |
            ForEach-Object { Write-Host $_ }
        throw "Could not read layout names from $($Dwg.Name)."
    }

    return $layouts
}

function Get-NameTokens {
    param([System.IO.FileInfo]$Dwg)

    $tokens = @{
        file = $Dwg.BaseName
        filename = $Dwg.Name
    }

    $pattern = [string]$script:Config.Naming.FileNameRegex
    if (![string]::IsNullOrWhiteSpace($pattern)) {
        $rx = [regex]::new($pattern)
        $match = $rx.Match($Dwg.BaseName)
        if (!$match.Success) {
            throw "Filename does not match Naming.FileNameRegex: $($Dwg.Name)"
        }

        foreach ($groupName in $rx.GetGroupNames()) {
            if ($groupName -notmatch '^\d+$') {
                $tokens[$groupName] = $match.Groups[$groupName].Value
            }
        }
    }

    return $tokens
}

function ConvertTo-SafeToken {
    param([object]$Value)
    $text = [string]$Value
    return ($text -replace '[<>:"/\\|?*\x00-\x1F]', '_').Trim().TrimEnd(' ', '.')
}

function ConvertTo-SafeStem {
    param([string]$Value)
    $text = $Value -replace '[<>:"/\\|?*\x00-\x1F]', '_'
    $text = $text.Trim().TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw "The naming pattern produced an empty filename."
    }
    if ($text -match '(?i)\.pdf$') {
        throw "Naming patterns must not include the .pdf extension: $Value"
    }
    return $text
}

function ConvertTo-AlphaRevision {
    param([long]$Number)
    if ($Number -lt 1) { return [string]$Number }
    $result = ''
    $n = $Number
    while ($n -gt 0) {
        $n--
        $result = [string][char](65 + ($n % 26)) + $result
        $n = [math]::Floor($n / 26)
    }
    return $result
}

function ConvertTo-RomanRevision {
    param([long]$Number)
    if ($Number -lt 1 -or $Number -gt 3999) { return [string]$Number }
    $map = @(
        @{ Value = 1000; Text = 'M' }, @{ Value = 900; Text = 'CM' },
        @{ Value = 500; Text = 'D' }, @{ Value = 400; Text = 'CD' },
        @{ Value = 100; Text = 'C' }, @{ Value = 90; Text = 'XC' },
        @{ Value = 50; Text = 'L' }, @{ Value = 40; Text = 'XL' },
        @{ Value = 10; Text = 'X' }, @{ Value = 9; Text = 'IX' },
        @{ Value = 5; Text = 'V' }, @{ Value = 4; Text = 'IV' },
        @{ Value = 1; Text = 'I' }
    )
    $n = $Number
    $result = ''
    foreach ($entry in $map) {
        while ($n -ge $entry.Value) {
            $result += $entry.Text
            $n -= $entry.Value
        }
    }
    return $result
}

function Get-RevisionDisplay {
    param([object]$RawRevision)

    if ($null -eq $RawRevision -or [string]::IsNullOrWhiteSpace([string]$RawRevision)) {
        return ''
    }

    $text = [string]$RawRevision
    if ($script:Config.Revisions.Source -ne 'AUTO') { return $text }

    $number = 0L
    $isNumber = [long]::TryParse($text, [ref]$number)
    if (!$isNumber) { return $text }

    switch ([string]$script:Config.Revisions.Style) {
        'NUMERIC' {
            try { return $number.ToString([string]$script:Config.Revisions.NumericFormat, $script:Culture) }
            catch { throw "Invalid Revisions.NumericFormat." }
        }
        'ALPHA' { return ConvertTo-AlphaRevision $number }
        'ROMAN' { return ConvertTo-RomanRevision $number }
    }
}

function Get-TokenValue {
    param(
        [string]$Name,
        [string]$Format,
        [hashtable]$Values,
        [object]$RawRevision,
        [string]$RevisionDisplay,
        [datetime]$DwgDate
    )

    $key = $Name.ToLowerInvariant()

    switch ($key) {
        'date' {
            $dateFormat = if ($Format) { $Format } else { [string]$script:Config.Naming.DefaultDateFormat }
            try { return $script:PublishDate.ToString($dateFormat, $script:Culture) }
            catch { throw "Invalid date format in token {date:$dateFormat}." }
        }
        'dwgdate' {
            $dateFormat = if ($Format) { $Format } else { [string]$script:Config.Naming.DefaultDateFormat }
            try { return $DwgDate.ToString($dateFormat, $script:Culture) }
            catch { throw "Invalid date format in token {dwgdate:$dateFormat}." }
        }
        'rev' {
            if ([string]::IsNullOrWhiteSpace([string]$RawRevision)) { return '' }
            if ([string]::IsNullOrWhiteSpace($Format)) { return $RevisionDisplay }

            $rawText = [string]$RawRevision
            $number = 0L
            $isNumber = [long]::TryParse($rawText, [ref]$number)

            switch ($Format.ToLowerInvariant()) {
                'alpha' {
                    if ($isNumber) { return ConvertTo-AlphaRevision $number }
                    return $rawText
                }
                'roman' {
                    if ($isNumber) { return ConvertTo-RomanRevision $number }
                    return $rawText
                }
                default {
                    if ($isNumber) {
                        try { return $number.ToString($Format, $script:Culture) }
                        catch { throw "Invalid revision format '$Format'." }
                    }
                    return $rawText
                }
            }
        }
        default {
            if ($Values.ContainsKey($key)) {
                return (ConvertTo-SafeToken $Values[$key])
            }
            $tokenDescription = $Name
            if ($Format) { $tokenDescription = "$Name`:$Format" }
            throw "Unknown or unavailable naming token: {$tokenDescription}"
        }
    }
}

function Expand-Pattern {
    param(
        [string]$Pattern,
        [hashtable]$Values,
        [object]$RawRevision,
        [string]$RevisionDisplay,
        [datetime]$DwgDate
    )

    $tokenRegex = '\{(?<name>[A-Za-z][A-Za-z0-9_]*)(?::(?<format>[^{}]+))?\}'
    $found = [regex]::Matches($Pattern, $tokenRegex)
    $result = $Pattern

    for ($i = $found.Count - 1; $i -ge 0; $i--) {
        $item = $found[$i]
        $name = $item.Groups['name'].Value
        $format = $item.Groups['format'].Value
        $replacement = Get-TokenValue -Name $name -Format $format -Values $Values -RawRevision $RawRevision -RevisionDisplay $RevisionDisplay -DwgDate $DwgDate
        $result = $result.Remove($item.Index, $item.Length).Insert($item.Index, [string]$replacement)
    }

    if ($result -match '[{}]') {
        throw "Malformed or unsupported token in pattern: $Pattern"
    }

    return $result
}

function Get-OutputFolderForContext {
    param(
        [hashtable]$Values,
        [object]$RawRevision,
        [string]$RevisionDisplay,
        [datetime]$DwgDate
    )

    $folderPattern = [string]$script:Config.Revisions.FolderPattern
    if ([string]::IsNullOrWhiteSpace($folderPattern)) {
        return $script:OutputRoot
    }

    $relative = Expand-Pattern -Pattern $folderPattern -Values $Values -RawRevision $RawRevision -RevisionDisplay $RevisionDisplay -DwgDate $DwgDate
    if ([System.IO.Path]::IsPathRooted($relative) -or $relative -match '^[A-Za-z]:') {
        throw "Revisions.FolderPattern must be relative to Paths.OutputFolder."
    }

    $segments = $relative -split '[\\/]'
    if ($segments -contains '..') {
        throw "Revisions.FolderPattern cannot contain '..'."
    }

    $full = [System.IO.Path]::GetFullPath((Join-Path $script:OutputRoot $relative))
    $rootPrefix = $script:OutputRoot.TrimEnd('\') + '\'
    if (!$full.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase) -and $full -ne $script:OutputRoot) {
        throw "Resolved output folder escapes Paths.OutputFolder."
    }

    return $full
}

function New-LayoutPlan {
    param(
        [object]$Record,
        [object]$Layout,
        [object]$RawRevision
    )

    $values = @{}
    foreach ($key in $Record.NameTokens.Keys) { $values[$key.ToLowerInvariant()] = $Record.NameTokens[$key] }
    $values['layout'] = $Layout.Name

    $revisionDisplay = Get-RevisionDisplay $RawRevision
    $outFolder = Get-OutputFolderForContext -Values $values -RawRevision $RawRevision -RevisionDisplay $revisionDisplay -DwgDate $Record.Dwg.LastWriteTime
    $stem = Expand-Pattern -Pattern ([string]$script:Config.Naming.SeparatePattern) -Values $values -RawRevision $RawRevision -RevisionDisplay $revisionDisplay -DwgDate $Record.Dwg.LastWriteTime
    $stem = ConvertTo-SafeStem $stem
    $fullPath = Join-Path $outFolder ($stem + '.pdf')

    return [pscustomobject]@{
        DwgPath = $Record.Dwg.FullName
        BaseName = $Record.Dwg.BaseName
        Layout = $Layout
        Sheets = @($Layout)
        RawRevision = $RawRevision
        RevisionDisplay = $revisionDisplay
        Stem = $stem
        OutputFolder = $outFolder
        Path = $fullPath
        Mode = 'SEPARATE'
    }
}

function New-CombinedPlan {
    param(
        [object]$Record,
        [object]$RawRevision
    )

    $values = @{}
    foreach ($key in $Record.NameTokens.Keys) { $values[$key.ToLowerInvariant()] = $Record.NameTokens[$key] }

    $revisionDisplay = Get-RevisionDisplay $RawRevision
    $outFolder = Get-OutputFolderForContext -Values $values -RawRevision $RawRevision -RevisionDisplay $revisionDisplay -DwgDate $Record.Dwg.LastWriteTime
    $stem = Expand-Pattern -Pattern ([string]$script:Config.Naming.CombinedPattern) -Values $values -RawRevision $RawRevision -RevisionDisplay $revisionDisplay -DwgDate $Record.Dwg.LastWriteTime
    $stem = ConvertTo-SafeStem $stem
    $fullPath = Join-Path $outFolder ($stem + '.pdf')

    return [pscustomobject]@{
        DwgPath = $Record.Dwg.FullName
        BaseName = $Record.Dwg.BaseName
        Layout = $null
        Sheets = @($Record.Layouts)
        RawRevision = $RawRevision
        RevisionDisplay = $revisionDisplay
        Stem = $stem
        OutputFolder = $outFolder
        Path = $fullPath
        Mode = 'COMBINED'
    }
}

function New-PlansForRecord {
    param(
        [object]$Record,
        [object]$AutoRevision
    )

    $source = [string]$script:Config.Revisions.Source
    $mode = [string]$script:Config.Publishing.PdfMode

    if ($mode -eq 'SEPARATE') {
        $plans = @()
        foreach ($layout in $Record.Layouts) {
            switch ($source) {
                'AUTO'  { $rawRevision = $AutoRevision }
                'FIXED' { $rawRevision = [string]$script:Config.Revisions.FixedValue }
                'NONE'  { $rawRevision = '' }
            }
            $plans += New-LayoutPlan -Record $Record -Layout $layout -RawRevision $rawRevision
        }
        return $plans
    }

    switch ($source) {
        'AUTO'  { $rawRevision = $AutoRevision }
        'FIXED' { $rawRevision = [string]$script:Config.Revisions.FixedValue }
        'NONE'  { $rawRevision = '' }
    }

    return @(New-CombinedPlan -Record $Record -RawRevision $rawRevision)
}

function Assert-UniquePlanPaths {
    param([object[]]$Plans)
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($plan in $Plans) {
        if (!$seen.Add([string]$plan.Path)) {
            throw "Naming collision: more than one sheet maps to '$($plan.Path)'. Add {layout} or another unique token to the naming pattern."
        }
    }
}

function Test-PlansAvailable {
    param([object[]]$Plans)
    foreach ($plan in $Plans) {
        if (Test-Path -LiteralPath $plan.Path) { return $false }
    }
    return $true
}

function Get-AutoRevisionPlans {
    param([object[]]$Records)

    $start = [int]$script:Config.Revisions.Start
    $max = [int]$script:Config.Publishing.MaxAutoRevision
    $scope = [string]$script:Config.Revisions.Scope
    $result = @{}

    if ($scope -eq 'PER_PACKAGE') {
        for ($revision = $start; $revision -lt ($start + $max); $revision++) {
            $allPlans = @()
            foreach ($record in $Records) {
                $allPlans += New-PlansForRecord -Record $record -AutoRevision $revision
            }
            Assert-UniquePlanPaths -Plans $allPlans
            if (Test-PlansAvailable -Plans $allPlans) {
                foreach ($record in $Records) { $result[$record.Dwg.FullName] = @(New-PlansForRecord -Record $record -AutoRevision $revision) }
                Write-Host "Package revision selected: $(Get-RevisionDisplay $revision)" -ForegroundColor Cyan
                return $result
            }
        }
        throw "No unused package revision was found within the configured search range."
    }

    foreach ($record in $Records) {
        if ($scope -eq 'PER_DRAWING') {
            $selected = $null
            for ($revision = $start; $revision -lt ($start + $max); $revision++) {
                $candidatePlans = @(New-PlansForRecord -Record $record -AutoRevision $revision)
                Assert-UniquePlanPaths -Plans $candidatePlans
                if (Test-PlansAvailable -Plans $candidatePlans) {
                    $selected = $candidatePlans
                    Write-Host "Revision for $($record.Dwg.Name): $(Get-RevisionDisplay $revision)" -ForegroundColor Cyan
                    break
                }
            }
            if ($null -eq $selected) { throw "No unused revision found for $($record.Dwg.Name)." }
            $result[$record.Dwg.FullName] = @($selected)
        }
        elseif ($scope -eq 'PER_LAYOUT') {
            $selectedPlans = @()
            $usedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

            foreach ($layout in $record.Layouts) {
                $chosen = $null
                for ($revision = $start; $revision -lt ($start + $max); $revision++) {
                    $candidate = New-LayoutPlan -Record $record -Layout $layout -RawRevision $revision
                    if (!$usedPaths.Contains([string]$candidate.Path) -and !(Test-Path -LiteralPath $candidate.Path)) {
                        $chosen = $candidate
                        break
                    }
                }
                if ($null -eq $chosen) { throw "No unused revision found for layout '$($layout.Name)' in '$($record.Dwg.Name)'." }
                [void]$usedPaths.Add([string]$chosen.Path)
                $selectedPlans += $chosen
                Write-Host "Revision for $($record.Dwg.Name) / $($layout.Name): $($chosen.RevisionDisplay)" -ForegroundColor Cyan
            }

            $result[$record.Dwg.FullName] = @($selectedPlans)
        }
    }

    return $result
}

function Apply-ExistingFilePolicy {
    param([object[]]$Plans)

    $action = [string]$script:Config.Publishing.ExistingFileAction
    $exists = @($Plans | Where-Object { Test-Path -LiteralPath $_.Path })

    if ($exists.Count -eq 0) { return @($Plans) }

    switch ($action) {
        'ERROR' {
            throw "Target PDF already exists: $($exists[0].Path)"
        }
        'SKIP' {
            $remaining = @($Plans | Where-Object { !(Test-Path -LiteralPath $_.Path) })
            if ($remaining.Count -eq 0) {
                Write-Host "Skipped: all target files already exist." -ForegroundColor Yellow
            } else {
                Write-Host "Skipping $($exists.Count) existing target file(s)." -ForegroundColor Yellow
            }
            return $remaining
        }
        'OVERWRITE' {
            foreach ($plan in $exists) {
                Remove-Item -LiteralPath $plan.Path -Force
            }
            return @($Plans)
        }
    }
}

function New-DsdFile {
    param(
        [object[]]$Plans,
        [string]$DsdPath,
        [string]$TargetPath,
        [string]$OutputFolder,
        [string]$Mode
    )

    $lines = New-Object 'System.Collections.Generic.List[string]'
    [void]$lines.Add('[DWF6Version]')
    [void]$lines.Add('Ver=1')
    [void]$lines.Add('[DWF6MinorVersion]')
    [void]$lines.Add('MinorVer=1')

    foreach ($plan in $Plans) {
        foreach ($sheet in $plan.Sheets) {
            if ($Mode -eq 'SEPARATE') {
                $sheetTitle = $plan.Stem
            } else {
                $safeLayout = ConvertTo-SafeStem ([string]$sheet.Name)
                $safeBase = ConvertTo-SafeStem ([string]$plan.BaseName)
                $sheetTitle = "$safeBase-$safeLayout"
            }

            [void]$lines.Add('')
            [void]$lines.Add("[DWF6Sheet:$sheetTitle]")
            [void]$lines.Add('DWG="' + $plan.DwgPath + '"')
            [void]$lines.Add('Layout=' + [string]$sheet.Name)
            [void]$lines.Add('Setup=')
            [void]$lines.Add('OriginalSheetPath="' + $plan.DwgPath + '"')
            [void]$lines.Add('Has Plot Port=0')
            [void]$lines.Add('Has3DDWF=0')
        }
    }

    $targetType = if ($Mode -eq 'COMBINED') { 6 } else { 5 }
    [void]$lines.Add('')
    [void]$lines.Add('[Target]')
    [void]$lines.Add("Type=$targetType")
    [void]$lines.Add('DWF="' + $TargetPath + '"')
    [void]$lines.Add('OUT="' + $OutputFolder + '"')
    [void]$lines.Add('PWD=')

    [void]$lines.Add('')
    [void]$lines.Add('[PdfOptions]')
    [void]$lines.Add('IncludeHyperlinks=TRUE')
    [void]$lines.Add('CreateBookmarks=TRUE')
    [void]$lines.Add('CaptureFontsInDrawing=TRUE')
    [void]$lines.Add('ConvertTextToGeometry=FALSE')
    [void]$lines.Add('VectorResolution=1200')
    [void]$lines.Add('RasterResolution=400')

    [void]$lines.Add('')
    [void]$lines.Add('[AutoCAD Block Data]')
    [void]$lines.Add('IncludeBlockInfo=0')
    [void]$lines.Add('BlockTmplFilePath=')

    [void]$lines.Add('')
    [void]$lines.Add('[SheetSet Properties]')
    [void]$lines.Add('IsSheetSet=FALSE')
    [void]$lines.Add('IsHomogeneous=FALSE')
    [void]$lines.Add('SheetSet Name=')
    [void]$lines.Add('NoOfCopies=1')
    [void]$lines.Add('PlotStampOn=FALSE')
    [void]$lines.Add('ViewFile=FALSE')
    [void]$lines.Add('JobID=0')
    [void]$lines.Add('SelectionSetName=')
    [void]$lines.Add('AcadProfile=')
    [void]$lines.Add('CategoryName=')
    [void]$lines.Add('LogFilePath=')
    [void]$lines.Add('IncludeLayer=TRUE')
    [void]$lines.Add('LineMerge=FALSE')
    [void]$lines.Add('CurrentPrecision=')
    [void]$lines.Add('PromptForDwfName=FALSE')
    [void]$lines.Add('PwdProtectPublishedDWF=FALSE')
    [void]$lines.Add('PromptForPwd=FALSE')
    [void]$lines.Add('RepublishingMarkups=FALSE')
    [void]$lines.Add('PublishSheetSetMetadata=FALSE')
    [void]$lines.Add('PublishSheetMetadata=FALSE')
    [void]$lines.Add('3DDWFOptions=0 1')

    [System.IO.File]::WriteAllLines($DsdPath, $lines.ToArray(), [System.Text.Encoding]::Default)
}

function Invoke-PublishGroup {
    param(
        [object[]]$Plans,
        [string]$JobDir,
        [string]$Mode
    )

    if ($Plans.Count -eq 0) { return }

    $outputFolder = [string]$Plans[0].OutputFolder
    foreach ($plan in $Plans) {
        if ($plan.OutputFolder -ne $outputFolder) {
            throw "Internal error: one DSD group contains multiple output folders."
        }
    }

    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
    $dsdPath = Join-Path $JobDir ('PUBLISH_' + [guid]::NewGuid().ToString('N') + '.dsd')
    $scrPath = Join-Path $JobDir ('PUBLISH_' + [guid]::NewGuid().ToString('N') + '.scr')
    $targetPath = [string]$Plans[0].Path

    New-DsdFile -Plans $Plans -DsdPath $dsdPath -TargetPath $targetPath -OutputFolder $outputFolder -Mode $Mode

    $scrLines = @(
        'FILEDIA'
        '0'
        'BACKGROUNDPLOT'
        '0'
        '_.-PUBLISH'
        $dsdPath.Replace('\', '/')
    )
    [System.IO.File]::WriteAllLines($scrPath, $scrLines, [System.Text.Encoding]::ASCII)

    Push-Location $JobDir
    try {
        $publishOutput = @(& $script:ACCORE /i $Plans[0].DwgPath /s $scrPath /l en-US 2>&1 | ForEach-Object { [string]$_ })
        $exitCode = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    if ($exitCode -ne 0) {
        Write-Host "Core Console output:" -ForegroundColor Yellow
        $publishOutput | Select-Object -Last 30 | ForEach-Object { Write-Host $_ }
        throw "AcCoreConsole publishing process failed with exit code $exitCode."
    }

    $missing = @($Plans | Where-Object { !(Test-Path -LiteralPath $_.Path -PathType Leaf) })
    if ($missing.Count -gt 0) {
        Write-Host "Core Console output:" -ForegroundColor Yellow
        $publishOutput | Select-Object -Last 30 | ForEach-Object { Write-Host $_ }
        Write-Host 'Missing target files:' -ForegroundColor Yellow
        $missing | ForEach-Object { Write-Host "  $($_.Path)" }
        throw "Publish returned without creating all expected PDF files."
    }
}

function Get-RecordPlansForFixedSource {
    param([object]$Record)
    return @(New-PlansForRecord -Record $Record -AutoRevision $null)
}

Assert-Config

$inputFolder = [System.IO.Path]::GetFullPath([string]$script:Config.Paths.InputFolder)
$outputFolder = [System.IO.Path]::GetFullPath([string]$script:Config.Paths.OutputFolder)
New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null

$searchSubfolders = [bool]$script:Config.Paths.SearchSubfolders
$dwgFiles = @(Get-ChildItem -LiteralPath $inputFolder -Filter '*.dwg' -File -Recurse:$searchSubfolders | Sort-Object FullName)
if ($dwgFiles.Count -eq 0) {
    Write-Host "No DWG files found in input folder." -ForegroundColor Yellow
    exit 0
}

$workRoot = Join-Path $env:TEMP 'DWG_TO_PDF_WORK'
New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
$records = New-Object 'System.Collections.Generic.List[object]'
$jobDirectories = @{}
$failed = 0
$success = 0

Write-Host "DWG files found: $($dwgFiles.Count)" -ForegroundColor Cyan
Write-Host "PDF mode: $($script:Config.Publishing.PdfMode)"
Write-Host "Revision source: $($script:Config.Revisions.Source)"
Write-Host "Dry run: $($script:Config.Publishing.DryRun)"

foreach ($dwg in $dwgFiles) {
    $jobDir = Join-Path $workRoot ([guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $jobDir -Force | Out-Null
    $jobDirectories[$dwg.FullName] = $jobDir

    Write-Host "Scanning: $($dwg.Name)" -ForegroundColor Cyan
    try {
        $layouts = @(Invoke-LayoutScan -Dwg $dwg -JobDir $jobDir)
        $tokens = Get-NameTokens -Dwg $dwg

        $record = [pscustomobject]@{
            Dwg = $dwg
            NameTokens = $tokens
            Layouts = $layouts
        }
        $records.Add($record)

        Write-Host "  Layouts detected: $($layouts.Count)"
        foreach ($layout in $layouts) { Write-Host "    $($layout.Name)" }
    }
    catch {
        $failed++
        Write-Host "SCAN FAILED: $($dwg.Name): $($_.Exception.Message)" -ForegroundColor Red
        if (!$script:Config.Publishing.KeepTempFilesOnError) {
            Remove-Item $jobDir -Recurse -Force -ErrorAction SilentlyContinue
            $jobDirectories.Remove($dwg.FullName)
        } else {
            Write-Host "Temporary files: $jobDir"
        }
    }
}

$recordsArray = @($records.ToArray())
if ($recordsArray.Count -eq 0) {
    Write-Host "No drawings could be scanned." -ForegroundColor Red
    if (!$NoPause) { Read-Host "Press Enter to exit" }
    exit 1
}

$plansByDwg = @{}
try {
    if ($script:Config.Revisions.Source -eq 'AUTO') {
        $plansByDwg = Get-AutoRevisionPlans -Records $recordsArray
    } else {
        foreach ($record in $recordsArray) {
            $plansByDwg[$record.Dwg.FullName] = @(Get-RecordPlansForFixedSource -Record $record)
        }
    }
}
catch {
    Write-Host "Planning failed: $($_.Exception.Message)" -ForegroundColor Red
    if (!$script:Config.Publishing.KeepTempFilesOnError) {
        foreach ($jobDir in $jobDirectories.Values) {
            Remove-Item $jobDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    if (!$NoPause) { Read-Host "Press Enter to exit" }
    exit 1
}

$allPlans = @()
foreach ($record in $recordsArray) {
    if ($plansByDwg.ContainsKey($record.Dwg.FullName)) {
        $allPlans += @($plansByDwg[$record.Dwg.FullName])
    }
}
try {
    Assert-UniquePlanPaths -Plans $allPlans
}
catch {
    Write-Host "Output naming validation failed: $($_.Exception.Message)" -ForegroundColor Red
    if (!$script:Config.Publishing.KeepTempFilesOnError) {
        foreach ($jobDir in $jobDirectories.Values) {
            Remove-Item $jobDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    if (!$NoPause) { Read-Host "Press Enter to exit" }
    exit 1
}

Write-Host ""
Write-Host "Planned output:" -ForegroundColor Cyan
foreach ($plan in $allPlans) { Write-Host "  $($plan.Path)" }

if ($script:Config.Publishing.DryRun) {
    Write-Host ""
    Write-Host "DRY RUN: no PDFs were published. Set Publishing.DryRun = `$false in the config to publish." -ForegroundColor Yellow
    foreach ($record in $recordsArray) {
        $jobDir = $jobDirectories[$record.Dwg.FullName]
        Remove-Item $jobDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Host "Finished preview."
    if (!$NoPause) { Read-Host "Press Enter to exit" }
    if ($failed -gt 0) { exit 1 } else { exit 0 }
}

foreach ($record in $recordsArray) {
    $jobDir = $jobDirectories[$record.Dwg.FullName]
    $plans = @($plansByDwg[$record.Dwg.FullName])

    if ($plans.Count -eq 0) {
        Write-Host "Skipping $($record.Dwg.Name): no output plans remain." -ForegroundColor Yellow
        continue
    }

    try {
        $plans = @(Apply-ExistingFilePolicy -Plans $plans)
        if ($plans.Count -eq 0) { continue }
        Assert-UniquePlanPaths -Plans $plans

        if ($script:Config.Publishing.PdfMode -eq 'COMBINED') {
            Invoke-PublishGroup -Plans $plans -JobDir $jobDir -Mode 'COMBINED'
        } else {
            $groups = @($plans | Group-Object -Property OutputFolder)
            foreach ($group in $groups) {
                Invoke-PublishGroup -Plans @($group.Group) -JobDir $jobDir -Mode 'SEPARATE'
            }
        }

        $success++
        Write-Host "PUBLISHED: $($record.Dwg.Name)" -ForegroundColor Green
        Remove-Item $jobDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        $failed++
        Write-Host "PUBLISH FAILED: $($record.Dwg.Name): $($_.Exception.Message)" -ForegroundColor Red
        if (!$script:Config.Publishing.KeepTempFilesOnError) {
            Remove-Item $jobDir -Recurse -Force -ErrorAction SilentlyContinue
        } else {
            Write-Host "Temporary files: $jobDir"
        }
    }

    if (!$script:Config.Publishing.KeepTempFilesOnError) {
        Remove-Item $jobDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "Batch finished." -ForegroundColor Green
Write-Host "Drawings published: $success"
Write-Host "Drawings failed or skipped during scan: $failed"
Write-Host "Output folder: $outputFolder"
if (!$NoPause) { Read-Host "Press Enter to exit" }
if ($failed -gt 0) { exit 1 } else { exit 0 }
