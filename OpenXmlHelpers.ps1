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

function Split-OpenXmlCellRef {
    param([string]$CellRef)
    $ref = ($CellRef + '').Trim().ToUpperInvariant()
    if ($ref -notmatch '^([A-Z]+)(\d+)$') { return $null }
    return [pscustomobject]@{
        Ref = $ref
        Col = $matches[1]
        Row = [int]$matches[2]
    }
}

function Get-MergeIndexes_OpenXml {
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
        $ref = ($mc.Reference.Value + '').Trim()
        if (-not $ref) { continue }

        $parts = $ref -split ':'
        $a = ($parts[0] + '').Trim().ToUpperInvariant()
        $b = if ($parts.Count -ge 2) { ($parts[1] + '').Trim().ToUpperInvariant() } else { $a }

        $pA = Split-OpenXmlCellRef -CellRef $a
        $pB = Split-OpenXmlCellRef -CellRef $b
        if (-not $pA -or -not $pB) { continue }

        $c1 = Convert-ColLetterToIndex -Col $pA.Col
        $c2 = Convert-ColLetterToIndex -Col $pB.Col
        $rowA = [int]$pA.Row
        $rowB = [int]$pB.Row
        if ($c1 -le 0 -or $c2 -le 0 -or $rowA -le 0 -or $rowB -le 0) { continue }

        $cMin = [math]::Min($c1, $c2)
        $cMax = [math]::Max($c1, $c2)
        $rMin = [math]::Min($rowA, $rowB)
        $rMax = [math]::Max($rowA, $rowB)

        $owner = ("{0}{1}" -f (Convert-ColIndexToLetter -Index $cMin), $rMin)
        if (-not $reverse.ContainsKey($owner)) {
            $reverse[$owner] = New-Object System.Collections.Generic.List[string]
        }

        for ($r = $rMin; $r -le $rMax; $r++) {
            for ($c = $cMin; $c -le $cMax; $c++) {
                $cellRef = ("{0}{1}" -f (Convert-ColIndexToLetter -Index $c), $r)
                if (-not $forward.ContainsKey($cellRef)) { $forward[$cellRef] = $owner }
                if (-not $reverse[$owner].Contains($cellRef)) { [void]$reverse[$owner].Add($cellRef) }
            }
        }
    }

    # Convert lists to arrays for read-only style access outside this function.
    $reverseFrozen = @{}
    foreach ($k in $reverse.Keys) {
        $reverseFrozen[$k] = @($reverse[$k].ToArray())
    }
    return @{ Forward = $forward; Reverse = $reverseFrozen }
}

function Get-MergeCellMap_OpenXml {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart
    )
    $idx = Get-MergeIndexes_OpenXml -WorksheetPart $WorksheetPart
    if (-not $idx) { return @{} }
    if (-not $idx.ContainsKey('Forward')) { return @{} }
    return $idx.Forward
}

function Get-OpenXmlCellText {
    param(
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [Parameter(Mandatory=$true)]$Cell
    )
    if ($null -eq $Cell) { return '' }

    $val = $Cell.CellValue
    if ($null -eq $val) {
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
    if ($null -ne $cell) { return ,$cell }

    $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
    $cell.CellReference = $cellRef

    $targetIdx = Convert-ColLetterToIndex -Col $ColLetter
    $refCell = @($cells | Where-Object {
        $_.CellReference -and (Convert-ColLetterToIndex -Col (($_.CellReference.Value -replace '\d+$',''))) -gt $targetIdx
    })[0]

    if ($null -ne $refCell) { $row.InsertBefore($cell, $refCell) | Out-Null } else { $row.AppendChild($cell) | Out-Null }
    return ,$cell
}

function Test-OpenXmlTreatAsBlankText {
    param([string]$Text)
    $t = Normalize-OpenXmlText $Text
    if (-not $t) { return $true }
    if ($t -match '^(?i)(Recorded By:|Performed By:|PQC Reviewed By:|Date:)$') { return $true }
    if ($t -match '^(?i)(N\/?A|NA)$') { return $true }
    return $false
}

function Get-OpenXmlCellSafe {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)][int]$RowIndex,
        [Parameter(Mandatory=$true)][string]$ColLetter
    )

    if ($RowIndex -lt 1) { return $null }
    $col = ($ColLetter + '').Trim().ToUpperInvariant()
    if (-not $col) { return $null }

    $cell = $null
    try {
        $cell = Ensure-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex $RowIndex -ColLetter $col
    } catch {
        $cell = $null
    }

    if ($cell -is [DocumentFormat.OpenXml.Spreadsheet.CellValue]) { $cell = $cell.Parent }
    if ($cell -is [DocumentFormat.OpenXml.Spreadsheet.InlineString]) { $cell = $cell.Parent }
    if ($cell -is [DocumentFormat.OpenXml.Spreadsheet.Cell]) { return ,$cell }

    $cell = Find-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex $RowIndex -ColLetter $col
    if ($cell -is [DocumentFormat.OpenXml.Spreadsheet.CellValue]) { $cell = $cell.Parent }
    if ($cell -is [DocumentFormat.OpenXml.Spreadsheet.InlineString]) { $cell = $cell.Parent }
    if ($cell -is [DocumentFormat.OpenXml.Spreadsheet.Cell]) { return ,$cell }

    return $null
}

function Clear-OpenXmlCellContent {
    param([Parameter(Mandatory=$true)]$Cell)
    if (-not ($Cell -is [DocumentFormat.OpenXml.Spreadsheet.Cell])) { return $false }
    $Cell.CellFormula  = $null
    $Cell.CellValue    = $null
    $Cell.InlineString = $null
    $Cell.DataType     = $null
    return $true
}

function Set-OpenXmlCellInlineText {
    param(
        [Parameter(Mandatory=$true)]$Cell,
        [Parameter(Mandatory=$true)][string]$Value
    )
    if (-not ($Cell -is [DocumentFormat.OpenXml.Spreadsheet.Cell])) { return $false }
    [void](Clear-OpenXmlCellContent -Cell $Cell)
    $Cell.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
    $Cell.InlineString = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
    $Cell.InlineString.Text = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
    $Cell.InlineString.Text.Text = ($Value + '')
    return $true
}

function Resolve-OpenXmlMergeOwner {
    param(
        [Parameter(Mandatory=$true)][int]$RowIndex,
        [Parameter(Mandatory=$true)][string]$ColLetter,
        [hashtable]$ForwardMap
    )
    $col = ($ColLetter + '').Trim().ToUpperInvariant()
    $targetRef = ("{0}{1}" -f $col, $RowIndex)
    $ownerRef = $targetRef
    if ($ForwardMap -and $ForwardMap.ContainsKey($targetRef)) {
        $ownerRef = (($ForwardMap[$targetRef] + '').Trim().ToUpperInvariant())
    }

    $ownerInfo = Split-OpenXmlCellRef -CellRef $ownerRef
    if (-not $ownerInfo) { $ownerInfo = Split-OpenXmlCellRef -CellRef $targetRef }
    if (-not $ownerInfo) { return $null }

    return [pscustomobject]@{
        TargetRef = $targetRef
        OwnerRef  = $ownerInfo.Ref
        OwnerCol  = $ownerInfo.Col
        OwnerRow  = [int]$ownerInfo.Row
    }
}

function Get-OpenXmlMergeGroupOwnerRowRefs {
    param(
        [Parameter(Mandatory=$true)][string]$OwnerRef,
        [Parameter(Mandatory=$true)][int]$OwnerRow,
        [hashtable]$ReverseMap
    )

    $refs = New-Object System.Collections.Generic.List[string]
    if ($ReverseMap -and $ReverseMap.ContainsKey($OwnerRef)) {
        foreach ($r in @($ReverseMap[$OwnerRef])) {
            $info = Split-OpenXmlCellRef -CellRef $r
            if ($info -and ([int]$info.Row -eq [int]$OwnerRow)) {
                if (-not $refs.Contains($info.Ref)) { [void]$refs.Add($info.Ref) }
            }
        }
    }
    if (-not $refs.Contains($OwnerRef)) { [void]$refs.Add($OwnerRef) }
    return @($refs.ToArray())
}

function Write-OpenXmlCellText_DeterministicMerge {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [Parameter(Mandatory=$true)][int]$RowIndex,
        [Parameter(Mandatory=$true)][string]$ColLetter,
        [Parameter(Mandatory=$true)][string]$Value,
        [bool]$Overwrite = $false,
        [hashtable]$ForwardMap,
        [hashtable]$ReverseMap,
        [string]$TailValue
    )

    $result = [pscustomobject]@{
        Written     = $false
        TargetRef   = $null
        OwnerRef    = $null
        OwnerRow    = $null
        OwnerCol    = $null
        GroupRefs   = @()
        Reason      = $null
        BlockedBy   = $null
    }

    if ($RowIndex -lt 1) {
        $result.Reason = 'invalid row'
        return $result
    }

    $owner = Resolve-OpenXmlMergeOwner -RowIndex $RowIndex -ColLetter $ColLetter -ForwardMap $ForwardMap
    if (-not $owner) {
        $result.Reason = 'could not resolve owner'
        return $result
    }

    $result.TargetRef = $owner.TargetRef
    $result.OwnerRef = $owner.OwnerRef
    $result.OwnerRow = [int]$owner.OwnerRow
    $result.OwnerCol = $owner.OwnerCol

    $groupRefs = Get-OpenXmlMergeGroupOwnerRowRefs -OwnerRef $owner.OwnerRef -OwnerRow $owner.OwnerRow -ReverseMap $ReverseMap
    $result.GroupRefs = @($groupRefs)

    if (-not $Overwrite) {
        foreach ($ref in $groupRefs) {
            $ri = Split-OpenXmlCellRef -CellRef $ref
            if (-not $ri) { continue }
            $existingCell = Find-OpenXmlCell -WorksheetPart $WorksheetPart -RowIndex ([int]$ri.Row) -ColLetter $ri.Col
            if ($null -eq $existingCell) { continue }
            $existingRaw = Get-OpenXmlCellText -WorkbookPart $WorkbookPart -Cell $existingCell
            if (-not (Test-OpenXmlTreatAsBlankText -Text $existingRaw)) {
                $result.Reason = 'already filled'
                $result.BlockedBy = $ri.Ref
                return $result
            }
        }
    } else {
        foreach ($ref in $groupRefs) {
            $ri = Split-OpenXmlCellRef -CellRef $ref
            if (-not $ri) { continue }
            $clearCell = Get-OpenXmlCellSafe -WorksheetPart $WorksheetPart -RowIndex ([int]$ri.Row) -ColLetter $ri.Col
            if ($clearCell) { [void](Clear-OpenXmlCellContent -Cell $clearCell) }
        }
    }

    $ownerCell = Get-OpenXmlCellSafe -WorksheetPart $WorksheetPart -RowIndex $owner.OwnerRow -ColLetter $owner.OwnerCol
    if ($null -eq $ownerCell) {
        $result.Reason = 'owner cell missing'
        return $result
    }
    if (-not (Set-OpenXmlCellInlineText -Cell $ownerCell -Value $Value)) {
        $result.Reason = 'owner write failed'
        return $result
    }

    $tail = Normalize-OpenXmlText $TailValue
    if ($Overwrite -and $tail) {
        foreach ($ref in $groupRefs) {
            if ($ref -eq $owner.OwnerRef) { continue }
            $ri = Split-OpenXmlCellRef -CellRef $ref
            if (-not $ri) { continue }
            $tailCell = Get-OpenXmlCellSafe -WorksheetPart $WorksheetPart -RowIndex ([int]$ri.Row) -ColLetter $ri.Col
            if ($null -eq $tailCell) { continue }
            [void](Set-OpenXmlCellInlineText -Cell $tailCell -Value $tail)
        }
    }

    $result.Written = $true
    $result.Reason = 'written'
    return $result
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
    $w = Write-OpenXmlCellText_DeterministicMerge `
        -WorksheetPart $WorksheetPart `
        -WorkbookPart $WorkbookPart `
        -RowIndex $RowIndex `
        -ColLetter $ColLetter `
        -Value $Value `
        -Overwrite $Overwrite `
        -ForwardMap $MergeMap `
        -ReverseMap $null `
        -TailValue $null
    return [bool]$w.Written
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

        if ($null -eq $cell) { continue }
        $txt = Normalize-OpenXmlText (Get-OpenXmlCellText -WorkbookPart $WorkbookPart -Cell $cell)
        if (-not $txt) { continue }
        if ($txt.ToLowerInvariant().Contains($needleNorm)) { return $r }
    }
    return $null
}

function Find-FirstRowByContains_FromRow_OpenXml {
    param(
        [Parameter(Mandatory=$true)]$WorksheetPart,
        [Parameter(Mandatory=$true)]$WorkbookPart,
        [Parameter(Mandatory=$true)][string]$ColLetter,
        [Parameter(Mandatory=$true)][string]$Needle,
        [Parameter(Mandatory=$true)][int]$StartRow
    )
    $needleNorm = (Normalize-OpenXmlText $Needle).ToLowerInvariant()
    $sheetData = (Get-OpenXmlChildrenOfType -Parent $WorksheetPart.Worksheet -Type ([DocumentFormat.OpenXml.Spreadsheet.SheetData]))[0]
    if (-not $sheetData) { return $null }

    $fromRow = $StartRow
    if ($fromRow -lt 1) { $fromRow = 1 }

    foreach ($row in (Get-OpenXmlChildrenOfType -Parent $sheetData -Type ([DocumentFormat.OpenXml.Spreadsheet.Row]))) {
        $r = 0
        if ($row.RowIndex) { $r = [int]$row.RowIndex.Value }
        if ($r -lt $fromRow) { continue }
        if ($r -le 0) { continue }
        $cellRef = "$ColLetter$r"

        $cell = @(
            Get-OpenXmlChildrenOfType -Parent $row -Type ([DocumentFormat.OpenXml.Spreadsheet.Cell]) |
            Where-Object { $_.CellReference -and $_.CellReference.Value -eq $cellRef }
        )[0]

        if ($null -eq $cell) { continue }
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

        if ($null -eq $cell) { continue }
        $txt = Normalize-OpenXmlText (Get-OpenXmlCellText -WorkbookPart $WorkbookPart -Cell $cell)
        if (-not $txt) { continue }
        if ($txt -match '^(?i)(Performed By:|Recorded By:|PQC Reviewed By:|Date:)$') { continue }
        return $true
    }
    return $false
}

function Write-OpenXmlSignDebug {
    param(
        [bool]$Enabled,
        [System.Collections.Generic.List[string]]$Events,
        [string]$Message
    )
    $msg = ($Message + '')
    if ($null -ne $Events) { [void]$Events.Add($msg) }
    if ($Enabled) { Write-Verbose ("[OpenXML] $msg") -Verbose }
}

function Invoke-WorksheetSignature_OpenXml {
    <#
    Mode:
      - Sammanstallning
      - Granskning

    Debug example:
      Invoke-WorksheetSignature_OpenXml ... -Overwrite $true -DebugLog
      [OpenXML] Invoke start: mode=Sammanstallning, overwrite=True, file=...
      [OpenXML] Sheet 'Test Summary': nameRow=54, dateRow=54
      [OpenXML] Sheet 'Test Summary' Name: target=C54, owner=C54, written=True, reason=written

    Returns: object with Written/Skipped.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$FullName,
        [Parameter(Mandatory=$true)][string]$SignDateYmd,
        [Parameter(Mandatory=$true)][ValidateSet('Sammanstallning','Granskning')][string]$Mode,
        [switch]$HasResample,
        [bool]$Overwrite = $false,
        [switch]$DebugLog,
        [Parameter(Mandatory=$true)][string]$ModulesRoot
    )

    Set-StrictMode -Version 2.0

    Import-OpenXmlSdk -ModulesRoot $ModulesRoot | Out-Null

    $res = [pscustomobject]@{
        Mode              = $Mode
        OverwriteRequested= [bool]$Overwrite
        Written           = New-Object System.Collections.Generic.List[string]
        Skipped           = New-Object System.Collections.Generic.List[string]
        WrittenCells      = New-Object System.Collections.Generic.List[pscustomobject]
        DebugEvents       = New-Object System.Collections.Generic.List[string]
    }

    Write-OpenXmlSignDebug -Enabled ([bool]$DebugLog) -Events $res.DebugEvents -Message ("Invoke start: mode={0}, overwrite={1}, file={2}" -f $Mode, [bool]$Overwrite, (Split-Path $Path -Leaf))

    $pendingSheetVerifications = New-Object System.Collections.Generic.List[object]
    $tailFillOnOverwrite = 'N/A'
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

            $mergeIndexes = $null
            try { $mergeIndexes = Get-MergeIndexes_OpenXml -WorksheetPart $wsp } catch { $mergeIndexes = $null }
            $forwardMap = if ($mergeIndexes -and $mergeIndexes.ContainsKey('Forward')) { $mergeIndexes.Forward } else { $null }
            $reverseMap = if ($mergeIndexes -and $mergeIndexes.ContainsKey('Reverse')) { $mergeIndexes.Reverse } else { $null }
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
            $isReviewTestSummary = (
                $Mode -eq 'Granskning' -and
                $sheetName -eq 'Test Summary' -and
                $t.NameNeedle -eq 'PQC Reviewed By:'
            )
            if ($isReviewTestSummary -and $nameRow) {
                $startAt = [math]::Max(1, ([int]$nameRow - 1))
                $dateRow = Find-FirstRowByContains_FromRow_OpenXml `
                    -WorksheetPart $wsp `
                    -WorkbookPart $wbp `
                    -ColLetter 'I' `
                    -Needle $t.DateNeedle `
                    -StartRow $startAt
                Write-OpenXmlSignDebug -Enabled ([bool]$DebugLog) -Events $res.DebugEvents -Message ("Sheet '{0}' review date search: startRow={1}, found={2}" -f $sheetName, $startAt, $(if($dateRow){$dateRow}else{'-' }))
            } else {
                foreach ($dc in $t.DateLabelCols) {
                    $dateRow = Find-FirstRowByContains_OpenXml -WorksheetPart $wsp -WorkbookPart $wbp -ColLetter $dc -Needle $t.DateNeedle
                    if ($dateRow) { break }
                }
            }

            Write-OpenXmlSignDebug -Enabled ([bool]$DebugLog) -Events $res.DebugEvents -Message ("Sheet '{0}': nameRow={1}, dateRow={2}" -f $sheetName, $(if($nameRow){$nameRow}else{'-' }), $(if($dateRow){$dateRow}else{'-' }))

            $wroteAny = $false
            $sheetWrittenCells = New-Object System.Collections.Generic.List[pscustomobject]
            if ($nameRow) {
                $nameWriteRow = ($nameRow + $t.Offset)
                $nameWriteCol = $t.NameWriteCol
                $nameWrite = Write-OpenXmlCellText_DeterministicMerge `
                    -WorksheetPart $wsp `
                    -WorkbookPart $wbp `
                    -RowIndex $nameWriteRow `
                    -ColLetter $nameWriteCol `
                    -Value $FullName `
                    -Overwrite ([bool]$Overwrite) `
                    -ForwardMap $forwardMap `
                    -ReverseMap $reverseMap `
                    -TailValue $tailFillOnOverwrite

                Write-OpenXmlSignDebug -Enabled ([bool]$DebugLog) -Events $res.DebugEvents -Message ("Sheet '{0}' Name: target={1}, owner={2}, written={3}, reason={4}" -f $sheetName, $nameWrite.TargetRef, $nameWrite.OwnerRef, $nameWrite.Written, $nameWrite.Reason)

                if ($nameWrite.Written) {
                    [void]$sheetWrittenCells.Add([pscustomobject]@{ Sheet=$sheetName; Row=[int]$nameWrite.OwnerRow; Col=($nameWrite.OwnerCol + ''); Value=$FullName })
                }
                $wroteAny = [bool]$nameWrite.Written -or $wroteAny
            }
            if ($dateRow) {
                $dateWriteRow = ($dateRow + $t.Offset)
                $dateWriteCol = $t.DateWriteCol
                $dateWrite = Write-OpenXmlCellText_DeterministicMerge `
                    -WorksheetPart $wsp `
                    -WorkbookPart $wbp `
                    -RowIndex $dateWriteRow `
                    -ColLetter $dateWriteCol `
                    -Value $SignDateYmd `
                    -Overwrite ([bool]$Overwrite) `
                    -ForwardMap $forwardMap `
                    -ReverseMap $reverseMap `
                    -TailValue $tailFillOnOverwrite

                Write-OpenXmlSignDebug -Enabled ([bool]$DebugLog) -Events $res.DebugEvents -Message ("Sheet '{0}' Date: target={1}, owner={2}, written={3}, reason={4}" -f $sheetName, $dateWrite.TargetRef, $dateWrite.OwnerRef, $dateWrite.Written, $dateWrite.Reason)

                if ($dateWrite.Written) {
                    [void]$sheetWrittenCells.Add([pscustomobject]@{ Sheet=$sheetName; Row=[int]$dateWrite.OwnerRow; Col=($dateWrite.OwnerCol + ''); Value=$SignDateYmd })
                }
                $wroteAny = [bool]$dateWrite.Written -or $wroteAny
            }

            if ($wroteAny) {
                try { $wsp.Worksheet.Save() } catch {
                    throw "Kunde inte spara WorksheetPart '$sheetName': $($_.Exception.Message)"
                }
                if ([bool]$Overwrite) {
                    [void]$pendingSheetVerifications.Add([pscustomobject]@{ Sheet = $sheetName; Cells = @($sheetWrittenCells.ToArray()) })
                } else {
                    [void]$res.Written.Add($sheetName)
                    foreach ($wc in $sheetWrittenCells) { [void]$res.WrittenCells.Add($wc) }
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
        throw
    } finally {
        try { if ($doc) { $doc.Close(); $doc.Dispose() } } catch {
            try { if ($doc) { $doc.Dispose() } } catch {}
        }
    }

    if ([bool]$Overwrite -and $pendingSheetVerifications.Count -gt 0) {
        foreach ($sv in $pendingSheetVerifications) {
            $cells = @($sv.Cells)
            if (-not $cells -or $cells.Count -eq 0) { continue }
            $sheetVerify = Verify-WorksheetSignatures_OpenXml -Path $Path -WrittenCells $cells -ModulesRoot $ModulesRoot
            if ($sheetVerify.OK) {
                [void]$res.Written.Add(($sv.Sheet + ''))
                foreach ($wc in $cells) { [void]$res.WrittenCells.Add($wc) }
                Write-OpenXmlSignDebug -Enabled ([bool]$DebugLog) -Events $res.DebugEvents -Message ("Sheet '{0}' overwrite verify: OK ({1}/{2})" -f $sv.Sheet, $sheetVerify.CellsVerified, $sheetVerify.CellsChecked)
            } else {
                $reason = if ($sheetVerify.Error) { $sheetVerify.Error } else { ($sheetVerify.Mismatches -join '; ') }
                [void]$res.Skipped.Add(("{0} (verify mismatch)" -f $sv.Sheet))
                Write-OpenXmlSignDebug -Enabled ([bool]$DebugLog) -Events $res.DebugEvents -Message ("Sheet '{0}' overwrite verify failed: {1}" -f $sv.Sheet, $reason)
                throw ("OpenXML overwrite verification failed for sheet '{0}': {1}" -f $sv.Sheet, $reason)
            }
        }
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
    return ,$cell
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
            if ($null -eq $cell) {
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
