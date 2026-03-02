function Resolve-OpenXmlDllPath {
    param(
        [Parameter(Mandatory=$true)][string]$ModulesRoot
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    # Also allow LibRoot (Modules/Lib split)
    try {
        if ($global:LibRoot -and ($global:LibRoot -ne $ModulesRoot)) {
            $lib = ($global:LibRoot + '')
            if ($lib) {
                $candidates.Add((Join-Path $lib 'DocumentFormat.OpenXml.dll')) | Out-Null
            }
        }
    } catch {}

    # 1) Directly in Modules
    $candidates.Add((Join-Path $ModulesRoot 'DocumentFormat.OpenXml.dll'))

    # 2) NuGet extracted layout (preferred)
    $candidates.Add((Join-Path $ModulesRoot 'OpenXMLSDK\DocumentFormat.OpenXml.2.20.0\lib\net46\DocumentFormat.OpenXml.dll'))
    $candidates.Add((Join-Path $ModulesRoot 'OpenXMLSDK\DocumentFormat.OpenXml.2.20.0\lib\net40\DocumentFormat.OpenXml.dll'))

    # 3) Alternative layout
    $candidates.Add((Join-Path $ModulesRoot 'OpenXMLSDK\lib\net46\DocumentFormat.OpenXml.dll'))
    $candidates.Add((Join-Path $ModulesRoot 'OpenXMLSDK\lib\net40\DocumentFormat.OpenXml.dll'))

    foreach ($p in $candidates) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Import-OpenXmlSdk {
    param(
        [Parameter(Mandatory=$true)][string]$ModulesRoot
    )

    # Already loaded?
    if ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'DocumentFormat.OpenXml' }) {
        return $true
    }

    # Try .dll on disk first
    $dllPath = Resolve-OpenXmlDllPath -ModulesRoot $ModulesRoot
    if ($dllPath -and (Test-Path -LiteralPath $dllPath)) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($dllPath)   # Bypasses MOTW
            [void][System.Reflection.Assembly]::Load($bytes)
            return $true
        } catch {
            throw "Kunde inte ladda OpenXML DLL: $dllPath`n$($_.Exception.Message)"
        }
    }

    # Optional fallback: Base64-encoded DLL as text file
    $b64Path = $null
    try { if ($global:LibRoot) { $b64Path = Join-Path $global:LibRoot 'OpenXMLAssembly.txt' } } catch {}
    if (-not $b64Path) { $b64Path = Join-Path $ModulesRoot 'OpenXMLAssembly.txt' }
    if (Test-Path -LiteralPath $b64Path) {
        try {
            $base64 = Get-Content -LiteralPath $b64Path -Raw
            $bytes  = [Convert]::FromBase64String($base64)
            [void][System.Reflection.Assembly]::Load($bytes)
            return $true
        } catch {
            throw "Kunde inte ladda OpenXML DLL från Base64 ($b64Path):`n$($_.Exception.Message)"
        }
    }

    throw "DocumentFormat.OpenXml.dll hittades inte i Lib/Modules (eller Base64-fallback). Lägg den i Modules eller i Modules\OpenXMLSDK\... (net46 föredras, net40 fallback) eller placera OpenXMLAssembly.txt."
}

function Normalize-OpenXmlText {
    param([string]$Text)
    $t = ($Text + '')
    $t = $t -replace [char]0x00A0,' '
    $t = $t.Trim()
    $t = [regex]::Replace($t,'\s+',' ')
    return $t
}

function Get-OpenXmlChildrenOfType {
    param(
        [Parameter(Mandatory=$true)]$Parent,
        [Parameter(Mandatory=$true)][Type]$Type
    )
    if (-not $Parent) { return @() }
    $out = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($ch in $Parent.ChildElements) {
            if ($ch -is $Type) { $out.Add($ch) }
        }
    } catch {}
    return $out
}

function Get-OpenXmlDescendantsOfType {
    param(
        [Parameter(Mandatory=$true)]$Parent,
        [Parameter(Mandatory=$true)][Type]$Type
    )
    if (-not $Parent) { return @() }
    $out = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($d in $Parent.Descendants()) {
            if ($d -is $Type) { $out.Add($d) }
        }
    } catch {}
    return $out
}

function Convert-ColLetterToIndex {
    param([string]$Col)
    $c = ($Col + '').Trim().ToUpperInvariant()
    if (-not $c) { return 0 }
    $sum = 0
    foreach ($ch in $c.ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { return 0 }
        $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $sum
}

function Convert-ColIndexToLetter {
    param([int]$Index)
    if ($Index -le 0) { return '' }
    $i = $Index
    $letters = New-Object System.Text.StringBuilder
    while ($i -gt 0) {
        $i--  # 1-based
        $rem = ($i % 26)
        [void]$letters.Insert(0, [char]([int][char]'A' + $rem))
        $i = [int][math]::Floor($i / 26)
    }
    return $letters.ToString()
}

function Get-MergeCellMap_OpenXml {

    param(
        [Parameter(Mandatory=$true)]$WorksheetPart
    )

    $map = @{}
    $ws = $WorksheetPart.Worksheet
    if (-not $ws) { return $map }

    $mergeCells = (Get-OpenXmlChildrenOfType -Parent $ws -Type ([DocumentFormat.OpenXml.Spreadsheet.MergeCells]))[0]
    if (-not $mergeCells) { return $map }

    $mergeCellItems = Get-OpenXmlChildrenOfType -Parent $mergeCells -Type ([DocumentFormat.OpenXml.Spreadsheet.MergeCell])
    foreach ($mc in $mergeCellItems) {
        if (-not $mc.Reference) { continue }
        $ref = ($mc.Reference.Value + '').Trim()
        if (-not $ref) { continue }

        # Expect "A1:B2" or "C73:D73". If single ref, treat as owner only.
        $parts = $ref -split ':'
        $a = ($parts[0] + '').Trim().ToUpperInvariant()
        $b = if ($parts.Count -ge 2) { ($parts[1] + '').Trim().ToUpperInvariant() } else { $a }

        if ($a -notmatch '^([A-Z]+)(\d+)$') { continue }
        $colA = $matches[1]; $rowA = [int]$matches[2]
        if ($b -notmatch '^([A-Z]+)(\d+)$') { continue }
        $colB = $matches[1]; $rowB = [int]$matches[2]

        $c1 = Convert-ColLetterToIndex -Col $colA
        $c2 = Convert-ColLetterToIndex -Col $colB
        if ($c1 -le 0 -or $c2 -le 0 -or $rowA -le 0 -or $rowB -le 0) { continue }

        $cMin = [math]::Min($c1, $c2)
        $cMax = [math]::Max($c1, $c2)
        $rMin = [math]::Min($rowA, $rowB)
        $rMax = [math]::Max($rowA, $rowB)

        $owner = ("{0}{1}" -f (Convert-ColIndexToLetter -Index $cMin), $rMin)

        for ($r = $rMin; $r -le $rMax; $r++) {
            for ($c = $cMin; $c -le $cMax; $c++) {
                $cellRef = ("{0}{1}" -f (Convert-ColIndexToLetter -Index $c), $r)
                if (-not $map.ContainsKey($cellRef)) { $map[$cellRef] = $owner }
            }
        }
    }
    return $map
 }

function Get-MergeRanges_OpenXml {
    <#
    .SYNOPSIS
        Parses worksheet MergeCells and returns a list of merge range objects.
        Each object contains: Ref, OwnerRef, StartRow, EndRow, StartColIndex, EndColIndex, OwnerRow, OwnerColIndex
    #>
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart
    )

    $ranges = New-Object System.Collections.Generic.List[object]
    $ws = $WorksheetPart.Worksheet
    if (-not $ws) { return $ranges }

    $mergeCells = (Get-OpenXmlChildrenOfType -Parent $ws -Type ([DocumentFormat.OpenXml.Spreadsheet.MergeCells]))[0]
    if (-not $mergeCells) { return $ranges }

    foreach ($mc in (Get-OpenXmlChildrenOfType -Parent $mergeCells -Type ([DocumentFormat.OpenXml.Spreadsheet.MergeCell]))) {
        if (-not $mc.Reference) { continue }
        $ref = ($mc.Reference.Value + '').Trim().ToUpperInvariant()
        if (-not $ref) { continue }

        $parts = $ref -split ':'
        $a = ($parts[0] + '').Trim()
        $b = if ($parts.Count -ge 2) { ($parts[1] + '').Trim() } else { $a }

        if ($a -notmatch '^([A-Z]+)(\d+)$') { continue }
        $colA = $matches[1]; $rowA = [int]$matches[2]
        if ($b -notmatch '^([A-Z]+)(\d+)$') { continue }
        $colB = $matches[1]; $rowB = [int]$matches[2]

        $c1 = Convert-ColLetterToIndex -Col $colA
        $c2 = Convert-ColLetterToIndex -Col $colB
        if ($c1 -le 0 -or $c2 -le 0 -or $rowA -le 0 -or $rowB -le 0) { continue }

        $cMin = [math]::Min($c1,$c2); $cMax = [math]::Max($c1,$c2)
        $rMin = [math]::Min($rowA,$rowB); $rMax = [math]::Max($rowA,$rowB)

        $ownerCol = Convert-ColIndexToLetter -Index $cMin
        $ownerRef = ("{0}{1}" -f $ownerCol, $rMin)

        $ranges.Add([pscustomobject]@{
            Ref          = $ref
            OwnerRef     = $ownerRef
            StartRow     = $rMin
            EndRow       = $rMax
            StartColIndex= $cMin
            EndColIndex  = $cMax
            OwnerRow     = $rMin
            OwnerColIndex= $cMin
        }) | Out-Null
    }
    return $ranges
}

function Get-MergeRangeForCellRef_OpenXml {
    <#
    .SYNOPSIS
        Returns the merge range object that contains the given cellRef (e.g. "E54"), or $null.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$CellRef,
        [Parameter(Mandatory=$true)][object]$MergeRanges
    )
    $cr = ($CellRef + '').Trim().ToUpperInvariant()
    if (-not $cr -or -not $MergeRanges) { return $null }
    if ($cr -notmatch '^([A-Z]+)(\d+)$') { return $null }
    $col = $matches[1]; $row = [int]$matches[2]
    $cIdx = Convert-ColLetterToIndex -Col $col
    if ($cIdx -le 0 -or $row -le 0) { return $null }

    foreach ($mr in $MergeRanges) {
        if ($row -ge $mr.StartRow -and $row -le $mr.EndRow -and $cIdx -ge $mr.StartColIndex -and $cIdx -le $mr.EndColIndex) {
            return $mr
        }
    }
    return $null
}

function Get-CellRefsInMergeRange_OpenXml {
    <#
    .SYNOPSIS
        Expands a merge range object to all cellRefs in the range.
    #>
    param(
        [Parameter(Mandatory=$true)]$MergeRange
    )
    $refs = New-Object System.Collections.Generic.List[string]
    if (-not $MergeRange) { return $refs }
    for ($r = [int]$MergeRange.StartRow; $r -le [int]$MergeRange.EndRow; $r++) {
        for ($c = [int]$MergeRange.StartColIndex; $c -le [int]$MergeRange.EndColIndex; $c++) {
            $refs.Add(("{0}{1}" -f (Convert-ColIndexToLetter -Index $c), $r)) | Out-Null
        }
    }
    return $refs
}

function Clear-OpenXmlCellContent {
    param([Parameter(Mandatory=$true)]$Cell)
    if (-not $Cell) { return }
    try { $Cell.CellFormula = $null } catch {}
    try { $Cell.CellValue   = $null } catch {}
    try { $Cell.InlineString= $null } catch {}
    try { $Cell.DataType    = $null } catch {}
}


function Get-OpenXmlCellText {
    param(
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [Parameter(Mandatory=$true)]$Cell
    )
    if (-not $Cell) { return '' }

    $val = $Cell.CellValue
    if (-not $val) {
        if ($Cell.InlineString -and $Cell.InlineString.Text) { return ($Cell.InlineString.Text.Text + '') }
        return ''
    }

    $raw = ($val.Text + '')
    if (-not $raw) { return '' }

    if ($Cell.DataType -and $Cell.DataType.Value -eq [DocumentFormat.OpenXml.Spreadsheet.CellValues]::SharedString) {
        $sst = $WorkbookPart.SharedStringTablePart
        if (-not $sst -or -not $sst.SharedStringTable) { return '' }
        $idx = 0
        if (-not [int]::TryParse($raw, [ref]$idx)) { return '' }

        $items = Get-OpenXmlChildrenOfType -Parent $sst.SharedStringTable -Type ([DocumentFormat.OpenXml.Spreadsheet.SharedStringItem])
        if ($idx -lt 0 -or $idx -ge $items.Count) { return '' }
        $item = $items[$idx]

        if ($item.Text) { return ($item.Text.Text + '') }

        $sb = New-Object System.Text.StringBuilder
        foreach ($t in (Get-OpenXmlDescendantsOfType -Parent $item -Type ([DocumentFormat.OpenXml.Spreadsheet.Text]))) {
            [void]$sb.Append(($t.Text + ''))
        }
        return $sb.ToString()
    }

    return $raw
}

function Ensure-OpenXmlCell {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)][int]$RowIndex,
        [Parameter(Mandatory=$true)][string]$ColLetter
    )
    $ws = $WorksheetPart.Worksheet

    $sheetData = (Get-OpenXmlChildrenOfType -Parent $ws -Type ([DocumentFormat.OpenXml.Spreadsheet.SheetData]))[0]
    if (-not $sheetData) {
        $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
        $ws.AppendChild($sheetData) | Out-Null
    }

    $rows = Get-OpenXmlChildrenOfType -Parent $sheetData -Type ([DocumentFormat.OpenXml.Spreadsheet.Row])
    $row = @($rows | Where-Object { $_.RowIndex -and $_.RowIndex.Value -eq $RowIndex })[0]
    if (-not $row) {
        $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
        $row.RowIndex = [uint32]$RowIndex

        $ref = @($rows | Where-Object { $_.RowIndex -and $_.RowIndex.Value -gt $RowIndex })[0]
        if ($ref) { $sheetData.InsertBefore($row, $ref) | Out-Null } else { $sheetData.AppendChild($row) | Out-Null }
    }

    $cellRef = "$ColLetter$RowIndex"
    $cells = Get-OpenXmlChildrenOfType -Parent $row -Type ([DocumentFormat.OpenXml.Spreadsheet.Cell])
    $cell  = @($cells | Where-Object { $_.CellReference -and $_.CellReference.Value -eq $cellRef })[0]
    if ($cell) { return $cell }

    $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
    $cell.CellReference = $cellRef

    $targetIdx = Convert-ColLetterToIndex -Col $ColLetter
    $refCell = @($cells | Where-Object {
        $_.CellReference -and (Convert-ColLetterToIndex -Col (($_.CellReference.Value -replace '\d+$',''))) -gt $targetIdx
    })[0]

    if ($refCell) { $row.InsertBefore($cell, $refCell) | Out-Null } else { $row.AppendChild($cell) | Out-Null }
    return $cell
}

function Set-OpenXmlCellText {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [Parameter(Mandatory=$true)][int]$RowIndex,
        [Parameter(Mandatory=$true)][string]$ColLetter,
        [Parameter(Mandatory=$true)][string]$Value,
        [bool]$Overwrite = $false,
        [hashtable]$MergeMap,
        [object]$MergeRanges,
        [string]$TailFillValue = 'N/A'
    )
    if ($RowIndex -lt 1) { return $false }

    $origTargetRef = ("{0}{1}" -f $ColLetter.ToUpperInvariant(), $RowIndex)

    # Determine merge range (if any). We prefer MergeRanges (exact range) but can fall back to MergeMap for owner resolution.
    $mr = $null
    try { if ($MergeRanges) { $mr = Get-MergeRangeForCellRef_OpenXml -CellRef $origTargetRef -MergeRanges $MergeRanges } } catch {}

    if (-not $mr -and $MergeMap) {
        try {
            if ($MergeMap.ContainsKey($origTargetRef)) {
                $ownerRef = ($MergeMap[$origTargetRef] + '')
                if ($MergeRanges) { $mr = Get-MergeRangeForCellRef_OpenXml -CellRef $ownerRef -MergeRanges $MergeRanges }
            }
        } catch {}
    }

    # If in a merge, write to owner. Also prepare list of all refs in the merge group.
    $mergeRefs = $null
    $ownerRef  = $origTargetRef
    if ($mr) {
        $ownerRef = ($mr.OwnerRef + '')
        $mergeRefs = Get-CellRefsInMergeRange_OpenXml -MergeRange $mr
        if ($ownerRef -match '^([A-Z]+)(\d+)$') {
            $ColLetter = $matches[1]
            $RowIndex  = [int]$matches[2]
        }
    } elseif ($MergeMap) {
        # Fallback: owner-only resolution
        try {
            if ($MergeMap.ContainsKey($origTargetRef)) {
                $ownerRef = ($MergeMap[$origTargetRef] + '').ToUpperInvariant().Trim()
                if ($ownerRef -match '^([A-Z]+)(\d+)$') {
                    $ColLetter = $matches[1]
                    $RowIndex  = [int]$matches[2]
                }
            }
        } catch {}
    }


    $cell = Ensure-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex $RowIndex -ColLetter $ColLetter
    if (-not $cell) { return $false }

    # --- Typ-normalisering / guard ---
    if ($cell -is [DocumentFormat.OpenXml.Spreadsheet.CellValue]) { $cell = $cell.Parent }
    if ($cell -is [DocumentFormat.OpenXml.Spreadsheet.InlineString]) { $cell = $cell.Parent }
    if (-not ($cell -is [DocumentFormat.OpenXml.Spreadsheet.Cell])) {
        throw "Ensure-OpenXmlCell returned unexpected type (after normalization): $($cell.GetType().FullName)"
    }
    # --- /Typ-normalisering ---
    $existing = Normalize-OpenXmlText (Get-OpenXmlCellText -WorkbookPart $WorkbookPart -Cell $cell)
    # Treat cells that still contain a label (e.g. "Recorded By:", "Performed By:", "PQC Reviewed By:", "Date:") as blank.
    if ($existing -match '^(?i)(Recorded By:|Performed By:|PQC Reviewed By:|Date:)$') {
        $existing = ''
    }

    # Merge-aware existing check:
    # - If NOT overwriting: if any cell in the merge group contains a non-empty value, skip.
    # - If overwriting: clear the whole merge group first (removes "spökceller").
    if (-not $Overwrite) {
        if ($mergeRefs -and $mergeRefs.Count -gt 0) {
            foreach ($rref in $mergeRefs) {
                if ($rref -eq $ownerRef) { continue }
                if ($rref -notmatch '^([A-Z]+)(\d+)$') { continue }
                $c = Find-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex ([int]$matches[2]) -ColLetter $matches[1]
                if (-not $c) { continue }
                $t = Normalize-OpenXmlText (Get-OpenXmlCellText -WorkbookPart $WorkbookPart -Cell $c)
                if ($t -match '^(?i)(Recorded By:|Performed By:|PQC Reviewed By:|Date:)$') { $t = '' }
                if ($t) { return $false }
            }
        }
        if ($existing) { return $false }
    } else {
        # Overwrite: clear all cells in merge group (or just the resolved cell if not merged)
        if ($mergeRefs -and $mergeRefs.Count -gt 0) {
            foreach ($rref in $mergeRefs) {
                if ($rref -notmatch '^([A-Z]+)(\d+)$') { continue }
                $ccol = $matches[1]; $rrow = [int]$matches[2]
                $c = Ensure-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex $rrow -ColLetter $ccol
                Clear-OpenXmlCellContent -Cell $c
            }
        } else {
            Clear-OpenXmlCellContent -Cell $cell
        }
    }


    # If overwriting in a horizontal merge: fill tail cells with TailFillValue to eliminate "osynliga" tail artifacts.
    if ($Overwrite -and $mr -and $TailFillValue) {
        try {
            foreach ($rref in $mergeRefs) {
                if ($rref -eq $ownerRef) { continue }
                if ($rref -notmatch '^([A-Z]+)(\d+)$') { continue }
                $tcol = $matches[1]; $trow = [int]$matches[2]
                # Only fill same row as owner (horizontal tail). For vertical merges, leave tail empty.
                if ($trow -ne [int]$mr.OwnerRow) { continue }
                $tc = Ensure-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex $trow -ColLetter $tcol
                $tc.DataType    = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
                $tc.CellValue   = $null
                $tc.InlineString      = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
                $tc.InlineString.Text = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
                $tc.InlineString.Text.Text = $TailFillValue
            }
        } catch {}
    }



function Find-FirstRowByContains_OpenXml {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [Parameter(Mandatory=$true)][string]$ColLetter,
        [Parameter(Mandatory=$true)][string]$Needle
    )
    $needleNorm = (Normalize-OpenXmlText $Needle).ToLowerInvariant()
    $sheetData = (Get-OpenXmlChildrenOfType -Parent $WorksheetPart.Worksheet -Type ([DocumentFormat.OpenXml.Spreadsheet.SheetData]))[0]
    if (-not $sheetData) { return $null }

    foreach ($row in (Get-OpenXmlChildrenOfType -Parent $sheetData -Type ([DocumentFormat.OpenXml.Spreadsheet.Row]))) {
        $r = 0
        if ($row.RowIndex) { $r = [int]$row.RowIndex.Value }
        if ($r -le 0) { continue }
        $cellRef = "$ColLetter$r"

        $cell = @(
            Get-OpenXmlChildrenOfType -Parent $row -Type ([DocumentFormat.OpenXml.Spreadsheet.Cell]) |
            Where-Object { $_.CellReference -and $_.CellReference.Value -eq $cellRef }
        )[0]

        if (-not $cell) { continue }
        $txt = Normalize-OpenXmlText (Get-OpenXmlCellText -WorkbookPart $WorkbookPart -Cell $cell)
        if (-not $txt) { continue }
        if ($txt.ToLowerInvariant().Contains($needleNorm)) { return $r }
    }
    return $null
}

function Test-OpenXmlDataSummaryHasData {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [string]$DataColLetter = 'C',
        [hashtable]$MergeMap
    )
    $sheetData = (Get-OpenXmlChildrenOfType -Parent $WorksheetPart.Worksheet -Type ([DocumentFormat.OpenXml.Spreadsheet.SheetData]))[0]
    if (-not $sheetData) { return $false }

    foreach ($row in (Get-OpenXmlChildrenOfType -Parent $sheetData -Type ([DocumentFormat.OpenXml.Spreadsheet.Row]))) {
        $r = 0
        if ($row.RowIndex) { $r = [int]$row.RowIndex.Value }
        if ($r -le 0) { continue }
        $cellRef = "$DataColLetter$r"
        if ($MergeMap -and $MergeMap.ContainsKey($cellRef)) {
            $cellRef = ($MergeMap[$cellRef] + '')
        }

        $cell = @(
            Get-OpenXmlChildrenOfType -Parent $row -Type ([DocumentFormat.OpenXml.Spreadsheet.Cell]) |
            Where-Object { $_.CellReference -and $_.CellReference.Value -eq $cellRef }
        )[0]

        if (-not $cell) { continue }
        $txt = Normalize-OpenXmlText (Get-OpenXmlCellText -WorkbookPart $WorkbookPart -Cell $cell)
        if (-not $txt) { continue }
        if ($txt -match '^(?i)(Performed By:|Recorded By:|PQC Reviewed By:|Date:)$') { continue }
        return $true
    }
    return $false
}

function Invoke-WorksheetSignature_OpenXml {
    <#
    Mode:
      - Sammanstallning
      - Granskning

    Returns: object with Written/Skipped.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$FullName,
        [Parameter(Mandatory=$true)][string]$SignDateYmd,
        [Parameter(Mandatory=$true)][ValidateSet('Sammanstallning','Granskning')][string]$Mode,
        [switch]$HasResample,
        [switch]$Overwrite,
        [Parameter(Mandatory=$true)][string]$ModulesRoot
    )

        Set-StrictMode -Version 2.0

    Import-OpenXmlSdk -ModulesRoot $ModulesRoot | Out-Null

    $res = [pscustomobject]@{
        Mode         = $Mode
        Written      = New-Object System.Collections.Generic.List[string]
        Skipped      = New-Object System.Collections.Generic.List[string]
        WrittenCells = New-Object System.Collections.Generic.List[pscustomobject]
    }

    $doc = $null
    try {
        $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($Path, $true)
        $wbp = $doc.WorkbookPart
        if (-not $wbp -or -not $wbp.Workbook) { throw 'WorkbookPart saknas.' }
        $sheets = $wbp.Workbook.Sheets
        if (-not $sheets) { throw 'Sheets saknas.' }

        $targets = @()
        if ($Mode -eq 'Sammanstallning') {
            $targets = @(
                @{ Name='Test Summary';                 NameLabelCol='B'; NameNeedle='Recorded By:';       NameWriteCol='C'; DateLabelCols=@('I');     DateNeedle='Date:'; DateWriteCol='J'; Offset=0;  DataSummaryGuard=$false },
                @{ Name='Data Summary';                 NameLabelCol='A'; NameNeedle='Recorded By:';       NameWriteCol='B'; DateLabelCols=@('C');     DateNeedle='Date:'; DateWriteCol='D'; Offset=0;  DataSummaryGuard=$true  },
                @{ Name='Extra Data Summary';           NameLabelCol='A'; NameNeedle='Recorded By:';       NameWriteCol='B'; DateLabelCols=@('C');     DateNeedle='Date:'; DateWriteCol='D'; Offset=0;  DataSummaryGuard=$true  },
                # Resample: signera endast om fliken finns OCH innehåller data
                @{ Name='Resample Data Summary';         NameLabelCol='A'; NameNeedle='Recorded By:';       NameWriteCol='B'; DateLabelCols=@('C');     DateNeedle='Date:'; DateWriteCol='D'; Offset=0;  DataSummaryGuard=$true; ResampleOnly=$true },
                @{ Name='Resample Date Summary';        NameLabelCol='A'; NameNeedle='Recorded By:';       NameWriteCol='B'; DateLabelCols=@('C');     DateNeedle='Date:'; DateWriteCol='D'; Offset=0;  DataSummaryGuard=$true; ResampleOnly=$true },
                @{ Name='Seal Test Failure Count';      NameLabelCol='K'; NameNeedle='Performed By:';      NameWriteCol='L'; DateLabelCols=@('K');     DateNeedle='Date:'; DateWriteCol='L'; Offset=-1; DataSummaryGuard=$false },
                @{ Name='Statistical Process Control';  NameLabelCol='K'; NameNeedle='Performed By:';      NameWriteCol='L'; DateLabelCols=@('K');     DateNeedle='Date:'; DateWriteCol='L'; Offset=-1; DataSummaryGuard=$false },
                @{ Name='Vacuum Seal Data';             NameLabelCol='B'; NameNeedle='Recorded By:';       NameWriteCol='C'; DateLabelCols=@('D','E'); DateNeedle='Date:'; DateWriteCol='F'; Offset=0;  DataSummaryGuard=$false }
            )
        } else {
            $targets = @(
                @{ Name='Test Summary'; NameLabelCol='B'; NameNeedle='PQC Reviewed By:'; NameWriteCol='C'; DateLabelCols=@('I'); DateNeedle='Date:'; DateWriteCol='J'; Offset=0; DataSummaryGuard=$false }
            )
        }

        foreach ($t in $targets) {
            $sheetName = $t.Name

            if ($t.ContainsKey('ResampleOnly') -and $t.ResampleOnly -and (-not $HasResample)) {
                [void]$res.Skipped.Add("$sheetName (ej resample)")
                continue
            }

            $sheet = @(
                Get-OpenXmlChildrenOfType -Parent $sheets -Type ([DocumentFormat.OpenXml.Spreadsheet.Sheet]) |
                Where-Object { $_.Name -and $_.Name.Value -eq $sheetName }
            )[0]
            if (-not $sheet) {
                [void]$res.Skipped.Add($sheetName)
                continue
            }
            $wsp = $wbp.GetPartById($sheet.Id)
            if (-not $wsp) {
                [void]$res.Skipped.Add($sheetName)
                continue
            }

            # Build merge map once per sheet (used for merge-aware DataSummary guards)
            $mergeMap = $null
            try { $mergeMap = Get-MergeCellMap_OpenXml -WorksheetPart $wsp } catch { $mergeMap = $null }
            $mergeRanges = $null
            try { $mergeRanges = Get-MergeRanges_OpenXml -WorksheetPart $wsp } catch { $mergeRanges = $null }

            if ($t.DataSummaryGuard) {
                $hasData = Test-OpenXmlDataSummaryHasData -WorksheetPart $wsp -WorkbookPart $wbp -DataColLetter 'C' -MergeMap $mergeMap
                if (-not $hasData) {
                    [void]$res.Skipped.Add("$sheetName (ingen data)")
                    continue
                }
            }

            $nameRow = Find-FirstRowByContains_OpenXml -WorksheetPart $wsp -WorkbookPart $wbp -ColLetter $t.NameLabelCol -Needle $t.NameNeedle
            $dateRow = $null
            foreach ($dc in $t.DateLabelCols) {
                $dateRow = Find-FirstRowByContains_OpenXml -WorksheetPart $wsp -WorkbookPart $wbp -ColLetter $dc -Needle $t.DateNeedle
                if ($dateRow) { break }
            }

            $wroteAny = $false
            if ($nameRow) {
                $nameWritten = (Set-OpenXmlCellText -WorksheetPart $wsp -WorkbookPart $wbp -RowIndex ($nameRow + $t.Offset) -ColLetter $t.NameWriteCol -Value $FullName -Overwrite:$Overwrite -MergeMap $mergeMap -MergeRanges $mergeRanges -TailFillValue 'N/A')
                if ($nameWritten) {
                    [void]$res.WrittenCells.Add([pscustomobject]@{ Sheet=$sheetName; Row=($nameRow + $t.Offset); Col=$t.NameWriteCol; Value=$FullName })
                }
                $wroteAny = $nameWritten -or $wroteAny
            }
            if ($dateRow) {
                $dateWritten = (Set-OpenXmlCellText -WorksheetPart $wsp -WorkbookPart $wbp -RowIndex ($dateRow + $t.Offset) -ColLetter $t.DateWriteCol -Value $SignDateYmd -Overwrite:$Overwrite -MergeMap $mergeMap -MergeRanges $mergeRanges -TailFillValue 'N/A')
                if ($dateWritten) {
                    [void]$res.WrittenCells.Add([pscustomobject]@{ Sheet=$sheetName; Row=($dateRow + $t.Offset); Col=$t.DateWriteCol; Value=$SignDateYmd })
                }
                $wroteAny = $dateWritten -or $wroteAny
            }

            if ($wroteAny) {
                [void]$res.Written.Add($sheetName)
                try { $wsp.Worksheet.Save() } catch {
                    throw "Kunde inte spara WorksheetPart '$sheetName': $($_.Exception.Message)"
                }
            } else {
                $skipReason = ''
                if (-not $nameRow -and -not $dateRow) {
                    $skipReason = 'labels saknas'
                } else {
                    $skipReason = 'redan ifyllt'
                }
                [void]$res.Skipped.Add("$sheetName ($skipReason)")
            }
        }

        $wbp.Workbook.Save()
    } catch {
        try { if ($doc) { $doc.Dispose() } } catch {}
        throw
    }
    try { if ($doc) { $doc.Close(); $doc.Dispose() } } catch {
        try { if ($doc) { $doc.Dispose() } } catch {}
    }

    return $res
}

function Find-OpenXmlCell {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)][int]$RowIndex,
        [Parameter(Mandatory=$true)][string]$ColLetter
    )
    $ws = $WorksheetPart.Worksheet
    $sheetData = (Get-OpenXmlChildrenOfType -Parent $ws -Type ([DocumentFormat.OpenXml.Spreadsheet.SheetData]))[0]
    if (-not $sheetData) { return $null }

    $rows = Get-OpenXmlChildrenOfType -Parent $sheetData -Type ([DocumentFormat.OpenXml.Spreadsheet.Row])
    $row = @($rows | Where-Object { $_.RowIndex -and $_.RowIndex.Value -eq $RowIndex })[0]
    if (-not $row) { return $null }

    $cellRef = "$ColLetter$RowIndex"
    $cells = Get-OpenXmlChildrenOfType -Parent $row -Type ([DocumentFormat.OpenXml.Spreadsheet.Cell])
    $cell = @($cells | Where-Object { $_.CellReference -and $_.CellReference.Value -eq $cellRef })[0]
    return $cell
}

function Verify-WorksheetSignatures_OpenXml {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][object[]]$WrittenCells,
        [Parameter(Mandatory=$true)][string]$ModulesRoot
    )

    $result = [pscustomobject]@{
        OK             = $false
        CellsChecked   = 0
        CellsVerified  = 0
        Mismatches     = New-Object System.Collections.Generic.List[string]
        Error          = $null
    }

    if (-not $WrittenCells -or $WrittenCells.Count -eq 0) {
        $result.OK = $true
        return $result
    }

    Import-OpenXmlSdk -ModulesRoot $ModulesRoot | Out-Null

    $doc = $null
    try {
        $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($Path, $false)  # read-only
        $wbp = $doc.WorkbookPart
        if (-not $wbp -or -not $wbp.Workbook) { throw 'WorkbookPart saknas vid verifiering.' }
        $sheets = $wbp.Workbook.Sheets

        foreach ($wc in $WrittenCells) {
            $result.CellsChecked++
            $sheet = @(
                Get-OpenXmlChildrenOfType -Parent $sheets -Type ([DocumentFormat.OpenXml.Spreadsheet.Sheet]) |
                Where-Object { $_.Name -and $_.Name.Value -eq $wc.Sheet }
            )[0]
            if (-not $sheet) {
                [void]$result.Mismatches.Add("$($wc.Sheet) $($wc.Col)$($wc.Row): Fliken hittades inte vid verifiering")
                continue
            }
            $wsp = $wbp.GetPartById($sheet.Id)
            if (-not $wsp) {
                [void]$result.Mismatches.Add("$($wc.Sheet) $($wc.Col)$($wc.Row): WorksheetPart saknas")
                continue
            }
            $cell = Find-OpenXmlCell -WorksheetPart $wsp -RowIndex $wc.Row -ColLetter $wc.Col
            if (-not $cell) {
                $actual = ''
            } else {
                $actual = Normalize-OpenXmlText (Get-OpenXmlCellText -WorkbookPart $wbp -Cell $cell)
            }
            $actual = ($actual + '').Trim()
            $expected = ($wc.Value + '').Trim()
            if ($actual -eq $expected) {
                $result.CellsVerified++
            } else {
                [void]$result.Mismatches.Add("$($wc.Sheet) $($wc.Col)$($wc.Row): Förväntade='$expected' Faktisk='$actual'")
            }
        }

        $result.OK = ($result.Mismatches.Count -eq 0 -and $result.CellsVerified -eq $WrittenCells.Count)
    } catch {
        $result.Error = $_.Exception.Message
    } finally {
        try { if ($doc) { $doc.Dispose() } } catch {}
    }
    return $result
}