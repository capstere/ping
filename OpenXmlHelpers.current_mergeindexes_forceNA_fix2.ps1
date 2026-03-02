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


function Get-MergeIndexes_OpenXml {
    <#
    .SYNOPSIS
        Returns forward and reverse merge indexes:
          - Forward: cellRef -> ownerRef
          - Reverse: ownerRef -> List[cellRef] within the same merge range
        Keeps Get-MergeCellMap_OpenXml intact for backwards compatibility.
    #>
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart
    )

    $forward = @{}
    $reverse = @{}

    $ws = $WorksheetPart.Worksheet
    if (-not $ws) { return @{ Forward = $forward; Reverse = $reverse } }

    $mergeCells = (Get-OpenXmlChildrenOfType -Parent $ws -Type ([DocumentFormat.OpenXml.Spreadsheet.MergeCells]))[0]
    if (-not $mergeCells) { return @{ Forward = $forward; Reverse = $reverse } }

    $mergeCellItems = Get-OpenXmlChildrenOfType -Parent $mergeCells -Type ([DocumentFormat.OpenXml.Spreadsheet.MergeCell])
    foreach ($mc in $mergeCellItems) {
        if (-not $mc.Reference) { continue }
        $ref = (($mc.Reference.Value + '')).Trim().ToUpperInvariant()
        if (-not $ref) { continue }

        $parts = $ref -split ':'
        $a = (($parts[0] + '')).Trim().ToUpperInvariant()
        $b = if ($parts.Count -ge 2) { (($parts[1] + '')).Trim().ToUpperInvariant() } else { $a }

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

        $cellsInRange = New-Object System.Collections.Generic.List[string]
        for ($r = $rMin; $r -le $rMax; $r++) {
            for ($c = $cMin; $c -le $cMax; $c++) {
                $cellRef = ("{0}{1}" -f (Convert-ColIndexToLetter -Index $c), $r)
                $forward[$cellRef] = $owner
                [void]$cellsInRange.Add($cellRef)
            }
        }

        if (-not $reverse.ContainsKey($owner)) {
            $reverse[$owner] = $cellsInRange
        } else {
            foreach ($x in $cellsInRange) { [void]$reverse[$owner].Add($x) }
        }
    }

    return @{ Forward = $forward; Reverse = $reverse }
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


function Normalize-OpenXmlCellObject {
    param([Parameter(Mandatory=$true)]$CellObj)
    $c = $CellObj
    try {
        if ($c -is [DocumentFormat.OpenXml.Spreadsheet.CellValue]) { $c = $c.Parent }
        if ($c -is [DocumentFormat.OpenXml.Spreadsheet.InlineString]) { $c = $c.Parent }
    } catch {}
    return $c
}

function Set-OpenXmlCellText {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [Parameter(Mandatory=$true)][int]$RowIndex,
        [Parameter(Mandatory=$true)][string]$ColLetter,
        [Parameter(Mandatory=$true)][string]$Value,
        [bool]$Overwrite = $false,
        [hashtable]$MergeMap
    )
    if ($RowIndex -lt 1) { return $false }

    # Merge-aware: if the target cell is part of a merged range, write to the owner (top-left) cell.
    # Excel displays merged cell values from the owner only; writing to a non-owner cell may appear as no-op.
    if ($MergeMap) {
        try {
            $origRef = (("{0}{1}" -f $ColLetter.ToUpperInvariant(), $RowIndex) + '')
            if ($MergeMap.ContainsKey($origRef)) {
                $ownerRef = ($MergeMap[$origRef] + '').ToUpperInvariant().Trim()
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

    if (-not $Overwrite) {
        if ($existing) { return $false }
    }

$cell.DataType    = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
    $cell.CellValue   = $null
    $cell.InlineString      = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
    $cell.InlineString.Text = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
    $cell.InlineString.Text.Text = $Value
    return $true
}


function Write-OpenXmlCellText_Force {
    <#
    .SYNOPSIS
        Force-write text to the merge-owner and make the merge range deterministic:
          - Clears owner + horizontal tail cells (same row as owner)
          - Writes owner Value
          - Writes TailValue ("N/A") into horizontal tail cells
        This mirrors the practical effect of AutoM's EPPlus approach without reading styles/colors.
    .RETURNS
        PSCustomObject { Written; Row; Col; OwnerRef }
    #>
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [Parameter(Mandatory=$true)][int]$RowIndex,
        [Parameter(Mandatory=$true)][string]$ColLetter,
        [Parameter(Mandatory=$true)][string]$Value,
        [hashtable]$ForwardMap,
        [hashtable]$ReverseMap,
        [string]$TailValue = 'N/A'
    )

    if ($RowIndex -lt 1) { return [pscustomobject]@{ Written=$false; Row=$RowIndex; Col=$ColLetter; OwnerRef=("{0}{1}" -f $ColLetter,$RowIndex) } }

    $origRef  = ("{0}{1}" -f $ColLetter.ToUpperInvariant(), $RowIndex)
    $ownerRef = $origRef
    $ownerRow = $RowIndex
    $ownerCol = $ColLetter.ToUpperInvariant()

    if ($ForwardMap -and $ForwardMap.ContainsKey($origRef)) {
        $ownerRef = (($ForwardMap[$origRef] + '')).ToUpperInvariant().Trim()
        if ($ownerRef -match '^([A-Z]+)(\d+)$') {
            $ownerCol = $matches[1]
            $ownerRow = [int]$matches[2]
        }
    }

    $ownerRef = ("{0}{1}" -f $ownerCol, $ownerRow)
    $ownerRef = ($ownerRef + '').ToUpperInvariant().Trim()

    # Collect refs for this merge group (if any), else just owner
    $refs = @($ownerRef)
    if ($ReverseMap -and $ReverseMap.ContainsKey($ownerRef)) {
        $refs = @($ReverseMap[$ownerRef])
    }

    # Phase 1: clear owner + horizontal tails (same row as owner)
    foreach ($ref in $refs) {
        if ($ref -notmatch '^([A-Z]+)(\d+)$') { continue }
        $c = $matches[1]; $r = [int]$matches[2]
        if ($r -ne $ownerRow) { continue }  # only horizontal row

        $cell = Ensure-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex $r -ColLetter $c
        if (-not $cell) { continue }
                $cell = Normalize-OpenXmlCellObject $cell
        if (-not ($cell -is [DocumentFormat.OpenXml.Spreadsheet.Cell])) { throw "Ensure-OpenXmlCell returned unexpected type: $($cell.GetType().FullName)" }
try { $cell.CellFormula = $null } catch {}
        try { $cell.CellValue   = $null } catch {}
        try { $cell.InlineString= $null } catch {}
        try { $cell.DataType    = $null } catch {}
    }

    # Phase 2: write owner
    $ocell = Ensure-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex $ownerRow -ColLetter $ownerCol
    if (-not $ocell) { return [pscustomobject]@{ Written=$false; Row=$ownerRow; Col=$ownerCol; OwnerRef=$ownerRef } }

        $ocell = Normalize-OpenXmlCellObject $ocell
    if (-not ($ocell -is [DocumentFormat.OpenXml.Spreadsheet.Cell])) { throw "Ensure-OpenXmlCell returned unexpected type: $($ocell.GetType().FullName)" }
$ocell.DataType    = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
    $ocell.CellValue   = $null
    $ocell.InlineString      = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
    $ocell.InlineString.Text = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
    $ocell.InlineString.Text.Text = $Value

    # Phase 3: write TailValue into horizontal tails (excluding owner)
    if ($TailValue) {
        foreach ($ref in $refs) {
            if ($ref -eq $ownerRef) { continue }
            if ($ref -notmatch '^([A-Z]+)(\d+)$') { continue }
            $c = $matches[1]; $r = [int]$matches[2]
            if ($r -ne $ownerRow) { continue }

            $tcell = Ensure-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex $r -ColLetter $c
            if (-not $tcell) { continue }

                        $tcell = Normalize-OpenXmlCellObject $tcell
            if (-not ($tcell -is [DocumentFormat.OpenXml.Spreadsheet.Cell])) { throw "Ensure-OpenXmlCell returned unexpected type: $($tcell.GetType().FullName)" }
$tcell.DataType    = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
            $tcell.CellValue   = $null
            $tcell.InlineString      = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
            $tcell.InlineString.Text = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
            $tcell.InlineString.Text.Text = $TailValue
        }
    }

    return [pscustomobject]@{ Written=$true; Row=$ownerRow; Col=$ownerCol; OwnerRef=$ownerRef }
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
        [bool]$Overwrite = $false,
        [Parameter(Mandatory=$true)][string]$ModulesRoot
    )

        Set-StrictMode -Version 2.0

    # Deterministic bool for overwrite
    $overwriteBool = $false
    try { $overwriteBool = [bool]$Overwrite.IsPresent } catch { $overwriteBool = [bool]$Overwrite }

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

            # Build merge indexes once per sheet (Forward + Reverse). Backwards compatible.
            $mergeIdx = $null
            $forwardMap = $null
            $reverseMap = $null
            try {
                $mergeIdx = Get-MergeIndexes_OpenXml -WorksheetPart $wsp
                $forwardMap = $mergeIdx.Forward
                $reverseMap = $mergeIdx.Reverse
            } catch {
                $forwardMap = $null
                $reverseMap = $null
            }

            # Legacy forward map for existing helper calls
            $mergeMap = $forwardMap


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
                $nameWriteRow = ($nameRow + $t.Offset)
                $nameWriteCol = $t.NameWriteCol
                $nameLogRow   = $nameWriteRow
                $nameLogCol   = $nameWriteCol
                try {
                    if ($mergeMap) {
                        $ref0 = ("{0}{1}" -f ($nameWriteCol.ToUpperInvariant()), $nameWriteRow)
                        if ($mergeMap.ContainsKey($ref0)) {
                            $ownerRef = ($mergeMap[$ref0] + '').ToUpperInvariant().Trim()
                            if ($ownerRef -match '^([A-Z]+)(\d+)$') {
                                $nameLogCol = $matches[1]
                                $nameLogRow = [int]$matches[2]
                            }
                        }
                    }
                } catch {}
                if ($overwriteBool) {
                $nameWrittenObj = Write-OpenXmlCellText_Force -WorksheetPart $wsp -WorkbookPart $wbp -RowIndex $nameWriteRow -ColLetter $nameWriteCol -Value $FullName -ForwardMap $forwardMap -ReverseMap $reverseMap -TailValue 'N/A'
                $nameWritten = [bool]$nameWrittenObj.Written
                if ($nameWritten) {
                    $nameLogRow = [int]$nameWrittenObj.Row
                    $nameLogCol = ($nameWrittenObj.Col + '')
                }
            } else {
                $nameWritten = (Set-OpenXmlCellText -WorksheetPart $wsp -WorkbookPart $wbp -RowIndex $nameWriteRow -ColLetter $nameWriteCol -Value $FullName -Overwrite:$false -MergeMap $mergeMap)
            }
                if ($nameWritten) {
                    [void]$res.WrittenCells.Add([pscustomobject]@{ Sheet=$sheetName; Row=$nameLogRow; Col=$nameLogCol; Value=$FullName })
                }
                $wroteAny = $nameWritten -or $wroteAny
            }
            if ($dateRow) {
                $dateWriteRow = ($dateRow + $t.Offset)
                $dateWriteCol = $t.DateWriteCol
                $dateLogRow   = $dateWriteRow
                $dateLogCol   = $dateWriteCol
                try {
                    if ($mergeMap) {
                        $ref0 = ("{0}{1}" -f ($dateWriteCol.ToUpperInvariant()), $dateWriteRow)
                        if ($mergeMap.ContainsKey($ref0)) {
                            $ownerRef = ($mergeMap[$ref0] + '').ToUpperInvariant().Trim()
                            if ($ownerRef -match '^([A-Z]+)(\d+)$') {
                                $dateLogCol = $matches[1]
                                $dateLogRow = [int]$matches[2]
                            }
                        }
                    }
                } catch {}
                if ($overwriteBool) {
                $dateWrittenObj = Write-OpenXmlCellText_Force -WorksheetPart $wsp -WorkbookPart $wbp -RowIndex $dateWriteRow -ColLetter $dateWriteCol -Value $SignDateYmd -ForwardMap $forwardMap -ReverseMap $reverseMap -TailValue 'N/A'
                $dateWritten = [bool]$dateWrittenObj.Written
                if ($dateWritten) {
                    $dateLogRow = [int]$dateWrittenObj.Row
                    $dateLogCol = ($dateWrittenObj.Col + '')
                }
            } else {
                $dateWritten = (Set-OpenXmlCellText -WorksheetPart $wsp -WorkbookPart $wbp -RowIndex $dateWriteRow -ColLetter $dateWriteCol -Value $SignDateYmd -Overwrite:$false -MergeMap $mergeMap)
            }
                if ($dateWritten) {
                    [void]$res.WrittenCells.Add([pscustomobject]@{ Sheet=$sheetName; Row=$dateLogRow; Col=$dateLogCol; Value=$SignDateYmd })
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