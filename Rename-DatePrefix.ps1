# RENAME - GUI, datum-prefix-byte för LSP-mappar (PS 5.1, WinForms)
# ────────────────────────────────────────────────────────────────────
# Flöde:
#  1) Skriv LSP-nummer → [Sök] → välj LSP-mapp (listan fylls).
#  2) Välj/ändra datum-prefix → [Skanna] (listar ALLA mål - mappar + filer).
#  3) Granska två kolumner (Nuvarande / Nytt) → [Byt datum-prefix].
#  4) Logg och summering visas nederst. Dubbelklick öppnar filens mapp.
# ────────────────────────────────────────────────────────────────────

#region STA Relaunch
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $self = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    $exe  = (Get-Command powershell.exe).Source
    $relaunchArgs = @('-NoLogo','-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$self`"")
    if ($PSBoundParameters.ContainsKey('Lsp'))           { $relaunchArgs += @('-Lsp', $Lsp) }
    if ($PSBoundParameters.ContainsKey('HeadFolder'))    { $relaunchArgs += @('-HeadFolder', $HeadFolder) }
    if ($PSBoundParameters.ContainsKey('NewDatePrefix')) { $relaunchArgs += @('-NewDatePrefix', $NewDatePrefix) }
    Start-Process -FilePath $exe -ArgumentList ($relaunchArgs -join ' ')
    return
}
#endregion

[CmdletBinding()]
param(
    [string]$Lsp,
    [string]$HeadFolder,
    [string]$NewDatePrefix
)

# =====================[ KONFIG ]=====================
$RootFolders = @(
    '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Tests',
    '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING',
    '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING'
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique

if (-not $RootFolders -or $RootFolders.Count -eq 0) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Inga giltiga rotmappar hittades i konfigurationen.",
        "RENAME", 'OK', 'Warning') | Out-Null
    return
}

# =====================[ WinForms-init ]=====================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==== Script-state ====
$script:Items       = @()      # objekt från skanningen (mappar + filer)
$script:Head        = $null    # vald LSP-huvudmapp (fullpath)
$script:LspHits     = @()      # alla hittade LSP-mappar
$script:LogTb       = $null    # logg-rutan (TextBox)
$script:BusyLock    = $false   # förhindrar dubbelklick under körning

# =====================[ Hjälpfunktioner ]=====================

function Parse-DatePrefixName {
    <#
      Parsar ett fil-/mappnamn som börjar med "yyyy MM dd" (eller "yyyyMMdd", "yyyy-MM-dd").
      Returnerar $null om inget datum-prefix hittas.
      OBS: Extension hanteras INTE här -- anroparen ansvarar för det.
    #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }

    # Matcha ETT datum-block i början (inte flera)
    $m = [regex]::Match($Name, '^\s*(\d{4})[-\s]?(\d{2})[-\s]?(\d{2})\s*(.*)')
    if (-not $m.Success) { return $null }

    $yyyy = $m.Groups[1].Value
    $MM   = $m.Groups[2].Value
    $dd   = $m.Groups[3].Value
    $rest = $m.Groups[4].Value

    # Validera att det ser ut som ett rimligt datum
    try {
        [void][datetime]::ParseExact("$yyyy-$MM-$dd", 'yyyy-MM-dd', $null)
    } catch { return $null }

    # Trimma separatorer men INTE punkt (bevarar .xlsx-extension om direkt efter datum)
    $rest = $rest.TrimStart(" -_".ToCharArray())

    [pscustomobject]@{
        Date = "$yyyy $MM $dd"
        Rest = $rest
    }
}

function Build-NewName {
    <#
      Bygger nytt namn givet parsad info, nytt prefix och om det är en fil (har extension).
    #>
    param(
        [pscustomobject]$Parsed,
        [string]$NewPrefix,
        [bool]$IsFile,
        [string]$OriginalName
    )
    $rest = $Parsed.Rest

    if ($IsFile) {
        # Separera extension från rest
        # rest kan vara "Report.xlsx" eller ".xlsx" eller "Report" (utan ext) eller tom
        $ext  = [System.IO.Path]::GetExtension($OriginalName)
        $stem = $rest
        if ($ext -and $stem.EndsWith($ext, [System.StringComparison]::OrdinalIgnoreCase)) {
            $stem = $stem.Substring(0, $stem.Length - $ext.Length)
        }
        $stem = $stem.TrimEnd(" -_".ToCharArray())

        if ([string]::IsNullOrWhiteSpace($stem)) {
            return "$NewPrefix$ext"
        } else {
            return "$NewPrefix $stem$ext"
        }
    } else {
        # Mapp -- ingen extension att oroa sig för
        if ([string]::IsNullOrWhiteSpace($rest)) {
            return $NewPrefix
        } else {
            return "$NewPrefix $rest"
        }
    }
}

function Get-RelativePath {
    param([string]$Base, [string]$Full)
    if ([string]::IsNullOrWhiteSpace($Base) -or [string]::IsNullOrWhiteSpace($Full)) { return $Full }
    $b = (Resolve-Path -LiteralPath $Base).ProviderPath.TrimEnd('\')
    if ($Full.StartsWith($b, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $Full.Substring($b.Length).TrimStart('\')
    }
    return $Full
}

function Scan-DatePrefixItems {
    <#
      Skannar en mapp rekursivt och hittar alla mappar/filer med datum-prefix.
      Använder per-katalog try-catch för att hantera otillgängliga undermappar.
    #>
    param([string]$Base, [string]$Target)
    $out = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $Base -PathType Container)) { return @() }
    $baseResolved = (Resolve-Path -LiteralPath $Base).ProviderPath

    # Rekursiv mappsamling med per-mapp felhantering
    $allDirs  = New-Object System.Collections.ArrayList
    $allFiles = New-Object System.Collections.ArrayList
    $queue    = New-Object System.Collections.Queue
    $queue.Enqueue($baseResolved)

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        try {
            foreach ($d in [System.IO.Directory]::GetDirectories($current)) {
                [void]$allDirs.Add($d)
                $queue.Enqueue($d)
            }
        } catch { <# otillgänglig mapp -- hoppa över #> }
        try {
            foreach ($f in [System.IO.Directory]::GetFiles($current)) {
                [void]$allFiles.Add($f)
            }
        } catch { }
        if (($allDirs.Count + $allFiles.Count) % 500 -eq 0) {
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    # Analysera mappar
    foreach ($d in $allDirs) {
        $nm = [System.IO.Path]::GetFileName($d)
        $p  = Parse-DatePrefixName $nm
        if ($null -eq $p) { continue }
        $newName = Build-NewName -Parsed $p -NewPrefix $Target -IsFile $false -OriginalName $nm
        [void]$out.Add([pscustomobject]@{
            Type     = 'Mapp'
            FullName = $d
            RelPath  = (Get-RelativePath $baseResolved $d)
            OldName  = $nm
            NewName  = $newName
            IsDir    = $true
        })
    }

    # Analysera filer
    foreach ($f in $allFiles) {
        $nm = [System.IO.Path]::GetFileName($f)
        $p  = Parse-DatePrefixName $nm
        if ($null -eq $p) { continue }
        $newName = Build-NewName -Parsed $p -NewPrefix $Target -IsFile $true -OriginalName $nm
        [void]$out.Add([pscustomobject]@{
            Type     = 'Fil'
            FullName = $f
            RelPath  = (Get-RelativePath $baseResolved $f)
            OldName  = $nm
            NewName  = $newName
            IsDir    = $false
        })
    }

    return @($out)
}

function Test-FileLocked {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    try {
        $fs = New-Object System.IO.FileStream(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None)
        $fs.Close()
        return $false
    } catch {
        return $true
    }
}

function Get-UniqueLeafName {
    <#
      Returnerar ett unikt namn genom att lägga till (1), (2), ... om kollision finns.
      Returnerar $null om inget unikt namn kan skapas inom 99 försök.
    #>
    param([string]$ParentDir, [string]$DesiredName, [bool]$IsFile)
    $target = Join-Path $ParentDir $DesiredName
    if (-not (Test-Path -LiteralPath $target)) { return $DesiredName }

    if ($IsFile) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($DesiredName)
        $ext  = [System.IO.Path]::GetExtension($DesiredName)
    } else {
        $stem = $DesiredName
        $ext  = ''
    }

    for ($i = 1; $i -le 99; $i++) {
        $candidate = "$stem ($i)$ext"
        $target = Join-Path $ParentDir $candidate
        if (-not (Test-Path -LiteralPath $target)) { return $candidate }
    }
    return $null
}

function Apply-Rename {
    param([string]$Base, [object[]]$Items, [string]$Target)
    $DirOK = 0; $DirErr = 0; $FileOK = 0; $FileErr = 0; $Locked = @()

    # ── 1. Mappar först (djupaste först så att barn byts innan förälder) ──
    $dItems = @($Items | Where-Object { $_.IsDir }) |
        Sort-Object { ($_.FullName.Split('\')).Length } -Descending

    $i = 0
    foreach ($it in $dItems) {
        $i++
        if (($i % 50) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
        if (-not (Test-Path -LiteralPath $it.FullName -PathType Container)) { $DirErr++; continue }

        $parent = [System.IO.Path]::GetDirectoryName($it.FullName)
        $leaf   = Get-UniqueLeafName -ParentDir $parent -DesiredName $it.NewName -IsFile $false

        if (-not $leaf) { $DirErr++; continue }
        if ($it.FullName.TrimEnd('\') -ieq (Join-Path $parent $leaf).TrimEnd('\')) { continue }

        try {
            Rename-Item -LiteralPath $it.FullName -NewName $leaf -ErrorAction Stop
            $DirOK++
        } catch { $DirErr++ }
    }

    # ── 2. Skanna om filer (sökvägar kan ha ändrats av mapp-rename) ──
    $fItems = Scan-DatePrefixItems -Base $Base -Target $Target | Where-Object { -not $_.IsDir }

    $j = 0
    foreach ($it in $fItems) {
        $j++
        if (($j % 100) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
        if (-not (Test-Path -LiteralPath $it.FullName -PathType Leaf)) { $FileErr++; continue }
        if (Test-FileLocked $it.FullName) { $FileErr++; $Locked += $it.FullName; continue }

        $parent = [System.IO.Path]::GetDirectoryName($it.FullName)
        $leaf   = Get-UniqueLeafName -ParentDir $parent -DesiredName $it.NewName -IsFile $true

        if (-not $leaf) { $FileErr++; continue }
        if ($it.FullName -ieq (Join-Path $parent $leaf)) { continue }

        try {
            Rename-Item -LiteralPath $it.FullName -NewName $leaf -ErrorAction Stop
            $FileOK++
        } catch { $FileErr++ }
    }

    [pscustomobject]@{
        DirOK  = $DirOK;  DirErr  = $DirErr
        FileOK = $FileOK; FileErr = $FileErr
        LockedFiles = $Locked
    }
}

function Find-LspHeadFolders {
    param([string]$LspNumber)
    $hits = @()
    if ([string]::IsNullOrWhiteSpace($LspNumber)) { return $hits }
    $rx = ('(?i)#{0}(?!\d)' -f [regex]::Escape($LspNumber))

    foreach ($root in $RootFolders) {
        try {
            $children = Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue
        } catch { $children = @() }

        foreach ($c in $children) {
            if ($c.Name -match $rx) {
                $hits += [pscustomobject]@{ Name = $c.Name; FullPath = $c.FullName }
            }
        }
    }
    $hits
}

function Log {
    param([string]$t)
    if ($script:LogTb -and -not $script:LogTb.IsDisposed) {
        $ts = (Get-Date).ToString('HH:mm:ss')
        $entry = "[$ts]  $t"
        if ($script:LogTb.TextLength -gt 0) {
            $script:LogTb.AppendText("`r`n$entry")
        } else {
            $script:LogTb.AppendText($entry)
        }
        $script:LogTb.SelectionStart = $script:LogTb.TextLength
        $script:LogTb.ScrollToCaret()
    }
}

function Set-Busy {
    param([bool]$Busy)
    $script:BusyLock = $Busy
    $btnFind.Enabled   = -not $Busy
    $btnScan.Enabled   = -not $Busy
    $btnRename.Enabled = (-not $Busy -and $script:Items.Count -gt 0)
    $txtLsp.Enabled    = -not $Busy
    $txtDate.Enabled   = -not $Busy
    $listHits.Enabled  = -not $Busy
    if ($Busy) {
        $prg.Style = 'Marquee'
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    } else {
        $prg.Style = 'Blocks'; $prg.Value = 0
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# =====================[ GUI ]=====================

# ── Färgtema ──
$BgMain     = [System.Drawing.Color]::FromArgb(245, 247, 250)
$BgGroup    = [System.Drawing.Color]::White
$AccentBlue = [System.Drawing.Color]::FromArgb(0, 102, 178)
$AccentGreen = [System.Drawing.Color]::FromArgb(40, 140, 60)
$AccentRed  = [System.Drawing.Color]::FromArgb(200, 50, 50)
$GridHeader = [System.Drawing.Color]::FromArgb(68, 84, 106)
$GridHeaderFg = [System.Drawing.Color]::White
$GridAltRow = [System.Drawing.Color]::FromArgb(235, 240, 248)

$fontMain  = New-Object System.Drawing.Font("Segoe UI", 9)
$fontMono  = New-Object System.Drawing.Font("Consolas", 9)
$fontTitle = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# ── Form ──
$form = New-Object System.Windows.Forms.Form
$form.Text            = "RENAME - LSP datum-prefix"
$form.StartPosition   = "CenterScreen"
$form.Size            = New-Object System.Drawing.Size(1100, 780)
$form.MinimumSize     = New-Object System.Drawing.Size(900, 600)
$form.Font            = $fontMain
$form.BackColor       = $BgMain
$form.KeyPreview      = $true

# ── Layout: top panel ──
$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Dock   = 'Top'
$panelTop.Height = 160
$panelTop.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 4)
$form.Controls.Add($panelTop)

# ── GroupBox: LSP ──
$gpLsp = New-Object System.Windows.Forms.GroupBox
$gpLsp.Text      = " 1. Välj LSP-mapp "
$gpLsp.Font      = $fontTitle
$gpLsp.Location  = New-Object System.Drawing.Point(10, 4)
$gpLsp.Size      = New-Object System.Drawing.Size(1060, 85)
$gpLsp.Anchor    = 'Top,Left,Right'
$gpLsp.BackColor = $BgGroup
$panelTop.Controls.Add($gpLsp)

$lblLsp = New-Object System.Windows.Forms.Label
$lblLsp.Text = "LSP-nummer:"; $lblLsp.Font = $fontMain
$lblLsp.Location = New-Object System.Drawing.Point(12, 28); $lblLsp.AutoSize = $true
$gpLsp.Controls.Add($lblLsp)

$txtLsp = New-Object System.Windows.Forms.TextBox
$txtLsp.Font     = $fontMain
$txtLsp.Location = New-Object System.Drawing.Point(100, 25); $txtLsp.Width = 110
if ($Lsp) { $txtLsp.Text = $Lsp }
$gpLsp.Controls.Add($txtLsp)

$btnFind = New-Object System.Windows.Forms.Button
$btnFind.Text      = "🔍 Sök"
$btnFind.Font      = $fontMain
$btnFind.FlatStyle = 'Flat'
$btnFind.BackColor = $AccentBlue
$btnFind.ForeColor = [System.Drawing.Color]::White
$btnFind.Location  = New-Object System.Drawing.Point(225, 23)
$btnFind.Size      = New-Object System.Drawing.Size(90, 28)
$gpLsp.Controls.Add($btnFind)

$lblHits = New-Object System.Windows.Forms.Label
$lblHits.Text = "Mappar:"; $lblHits.Font = $fontMain
$lblHits.Location = New-Object System.Drawing.Point(12, 58); $lblHits.AutoSize = $true
$gpLsp.Controls.Add($lblHits)

$listHits = New-Object System.Windows.Forms.ComboBox
$listHits.Font          = $fontMain
$listHits.DropDownStyle  = 'DropDownList'
$listHits.Location      = New-Object System.Drawing.Point(100, 55)
$listHits.Width         = 700
$listHits.Anchor        = 'Top,Left,Right'
$gpLsp.Controls.Add($listHits)

# ── GroupBox: Datum ──
$gpDate = New-Object System.Windows.Forms.GroupBox
$gpDate.Text      = " 2. Datum-prefix "
$gpDate.Font      = $fontTitle
$gpDate.Location  = New-Object System.Drawing.Point(10, 92)
$gpDate.Size      = New-Object System.Drawing.Size(1060, 60)
$gpDate.Anchor    = 'Top,Left,Right'
$gpDate.BackColor = $BgGroup
$panelTop.Controls.Add($gpDate)

$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "Nytt prefix (yyyy MM dd):"; $lblDate.Font = $fontMain
$lblDate.Location = New-Object System.Drawing.Point(12, 26); $lblDate.AutoSize = $true
$gpDate.Controls.Add($lblDate)

$txtDate = New-Object System.Windows.Forms.TextBox
$txtDate.Font     = $fontMain
$txtDate.Location = New-Object System.Drawing.Point(180, 23); $txtDate.Width = 110
$txtDate.Text = if ($NewDatePrefix) { $NewDatePrefix } else { (Get-Date -Format 'yyyy MM dd') }
$gpDate.Controls.Add($txtDate)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text      = "📋 Skanna"
$btnScan.Font      = $fontMain
$btnScan.FlatStyle = 'Flat'
$btnScan.BackColor = $AccentBlue
$btnScan.ForeColor = [System.Drawing.Color]::White
$btnScan.Location  = New-Object System.Drawing.Point(310, 21)
$btnScan.Size      = New-Object System.Drawing.Size(110, 28)
$gpDate.Controls.Add($btnScan)

$lblDateHint = New-Object System.Windows.Forms.Label
$lblDateHint.Text      = "Ange nytt datum och klicka Skanna för att förhandsgranska."
$lblDateHint.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblDateHint.ForeColor = [System.Drawing.Color]::Gray
$lblDateHint.Location  = New-Object System.Drawing.Point(435, 28)
$lblDateHint.AutoSize  = $true
$gpDate.Controls.Add($lblDateHint)

# ── Separator ──
$sep = New-Object System.Windows.Forms.Label
$sep.Dock      = 'Top'
$sep.Height    = 1
$sep.BackColor = [System.Drawing.Color]::FromArgb(200, 210, 220)
$form.Controls.Add($sep)

# ── Middle: DataGridView (preview) ──
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock                  = 'Fill'
$grid.ReadOnly              = $true
$grid.AllowUserToAddRows    = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToOrderColumns = $false
$grid.RowHeadersVisible     = $false
$grid.SelectionMode         = 'FullRowSelect'
$grid.MultiSelect           = $false
$grid.AutoSizeRowsMode      = 'DisplayedCells'
$grid.Font                  = $fontMono
$grid.BorderStyle           = 'None'
$grid.BackgroundColor       = [System.Drawing.Color]::White
$grid.GridColor             = [System.Drawing.Color]::FromArgb(220, 225, 230)
$grid.CellBorderStyle       = 'SingleHorizontal'
$grid.AlternatingRowsDefaultCellStyle.BackColor = $GridAltRow
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor   = $GridHeader
$grid.ColumnHeadersDefaultCellStyle.ForeColor   = $GridHeaderFg
$grid.ColumnHeadersDefaultCellStyle.Font        = $fontTitle
$grid.ColumnHeadersDefaultCellStyle.WrapMode    = 'False'
$grid.ColumnHeadersHeight   = 30
$grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)

# Tre kolumner: Typ, Nuvarande, Nytt
$colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colType.Name = "Type"; $colType.HeaderText = "Typ"; $colType.Width = 55
$colType.AutoSizeMode = 'None'
[void]$grid.Columns.Add($colType)

$colOld = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colOld.Name = "OldName"; $colOld.HeaderText = "Nuvarande namn"
$colOld.AutoSizeMode = 'Fill'; $colOld.FillWeight = 50
[void]$grid.Columns.Add($colOld)

$colNew = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNew.Name = "NewName"; $colNew.HeaderText = "Nytt namn"
$colNew.AutoSizeMode = 'Fill'; $colNew.FillWeight = 50
[void]$grid.Columns.Add($colNew)

$form.Controls.Add($grid)

# ── Bottom panel ──
$panelBottom = New-Object System.Windows.Forms.Panel
$panelBottom.Dock   = 'Bottom'
$panelBottom.Height = 175
$panelBottom.BackColor = $BgGroup
$panelBottom.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
$form.Controls.Add($panelBottom)

$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text     = "Ingen skanning utförd ännu."
$lblSummary.Font     = $fontTitle
$lblSummary.Location = New-Object System.Drawing.Point(10, 6)
$lblSummary.AutoSize = $true
$panelBottom.Controls.Add($lblSummary)

$prg = New-Object System.Windows.Forms.ProgressBar
$prg.Location = New-Object System.Drawing.Point(10, 28)
$prg.Width    = 860; $prg.Height = 18
$prg.Style    = 'Blocks'; $prg.Minimum = 0; $prg.Maximum = 100; $prg.Value = 0
$prg.Anchor   = 'Top,Left,Right'
$panelBottom.Controls.Add($prg)

$tbLog = New-Object System.Windows.Forms.TextBox
$tbLog.Location   = New-Object System.Drawing.Point(10, 52)
$tbLog.Width      = 860; $tbLog.Height = 80
$tbLog.Multiline  = $true
$tbLog.ScrollBars = 'Vertical'
$tbLog.ReadOnly   = $true
$tbLog.Font       = $fontMono
$tbLog.BackColor  = [System.Drawing.Color]::FromArgb(250, 250, 250)
$tbLog.Anchor     = 'Top,Left,Right'
$panelBottom.Controls.Add($tbLog)
$script:LogTb = $tbLog

$btnRename = New-Object System.Windows.Forms.Button
$btnRename.Text      = "✅ Byt datum-prefix"
$btnRename.Font      = $fontTitle
$btnRename.FlatStyle = 'Flat'
$btnRename.BackColor = $AccentGreen
$btnRename.ForeColor = [System.Drawing.Color]::White
$btnRename.Enabled   = $false
$btnRename.Location  = New-Object System.Drawing.Point(890, 52)
$btnRename.Size      = New-Object System.Drawing.Size(170, 36)
$btnRename.Anchor    = 'Top,Right'
$panelBottom.Controls.Add($btnRename)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text      = "Stäng"
$btnClose.Font      = $fontMain
$btnClose.FlatStyle = 'Flat'
$btnClose.BackColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$btnClose.ForeColor = [System.Drawing.Color]::White
$btnClose.Location  = New-Object System.Drawing.Point(890, 96)
$btnClose.Size      = New-Object System.Drawing.Size(170, 36)
$btnClose.Anchor    = 'Top,Right'
$btnClose.Add_Click({ $form.Close() })
$panelBottom.Controls.Add($btnClose)

# ── Statusfält (nederst) ──
$statusBar = New-Object System.Windows.Forms.StatusStrip
$slMain = New-Object System.Windows.Forms.ToolStripStatusLabel
$slMain.Text = "Redo."
$slMain.Spring = $true
$slMain.TextAlign = 'MiddleLeft'
$statusBar.Items.Add($slMain) | Out-Null
$form.Controls.Add($statusBar)

# ── Snabbtangenter ──
$txtLsp.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; $btnFind.PerformClick() }
})
$txtDate.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; $btnScan.PerformClick() }
})

# ── Dubbelklick i grid → öppna mapp i Explorer ──
$grid.Add_CellDoubleClick({
    param($s, $e)
    if ($e.RowIndex -lt 0 -or $e.RowIndex -ge $script:Items.Count) { return }
    $item = $script:Items[$e.RowIndex]
    if (-not $item) { return }
    $folder = if ($item.IsDir) { $item.FullName } else { [System.IO.Path]::GetDirectoryName($item.FullName) }
    if (Test-Path -LiteralPath $folder) {
        Start-Process explorer.exe "/select,`"$($item.FullName)`""
    }
})

# =====================[ Händelser ]=====================

# ── Sök LSP ──
$btnFind.Add_Click({
    if ($script:BusyLock) { return }
    $grid.Rows.Clear(); $script:Items = @(); $btnRename.Enabled = $false
    $listHits.Items.Clear(); $script:Head = $null; $script:LspHits = @()

    $l = ($txtLsp.Text -replace '\D', '')
    if ([string]::IsNullOrWhiteSpace($l)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Ange ett LSP-nummer (siffror).", "RENAME", 'OK', 'Information') | Out-Null
        return
    }
    $txtLsp.Text = $l

    Set-Busy $true
    Log "Söker LSP-mappar för #$l ..."
    $slMain.Text = "Söker..."

    $hits = Find-LspHeadFolders $l
    $script:LspHits = $hits

    if (-not $hits -or $hits.Count -eq 0) {
        Log "Inga LSP-huvudmappar hittades."
        $slMain.Text = "Inga mappar hittades."
        Set-Busy $false
        return
    }

    foreach ($h in $hits) { [void]$listHits.Items.Add($h.Name) }
    $listHits.SelectedIndex = 0
    $script:Head = $hits[0].FullPath
    Log ("Hittade {0} LSP-mapp(ar)." -f $hits.Count)
    $slMain.Text = ("Vald: {0}" -f $hits[0].Name)
    Set-Busy $false
})

# ── Byte av LSP i dropdown ──
$listHits.Add_SelectedIndexChanged({
    if ($script:BusyLock) { return }
    $grid.Rows.Clear(); $script:Items = @(); $btnRename.Enabled = $false
    if ($listHits.SelectedIndex -ge 0 -and $listHits.SelectedIndex -lt $script:LspHits.Count) {
        $script:Head = $script:LspHits[$listHits.SelectedIndex].FullPath
        $slMain.Text = ("Vald: {0}" -f $script:LspHits[$listHits.SelectedIndex].Name)
    }
})

# ── Skanna ──
$btnScan.Add_Click({
    if ($script:BusyLock) { return }
    $grid.Rows.Clear(); $script:Items = @(); $btnRename.Enabled = $false

    if (-not $script:Head) {
        [System.Windows.Forms.MessageBox]::Show(
            "Sök och välj en LSP-mapp först.", "RENAME", 'OK', 'Information') | Out-Null
        return
    }
    $prefix = $txtDate.Text.Trim()
    if ($prefix -notmatch '^\d{4}\s\d{2}\s\d{2}$') {
        [System.Windows.Forms.MessageBox]::Show(
            "Ogiltigt datum-prefix.`nAnvänd formatet 'yyyy MM dd' (t.ex. 2025 02 08).",
            "RENAME", 'OK', 'Warning') | Out-Null
        return
    }

    # Validera datumet
    try {
        $checkDate = $prefix -replace '\s', '-'
        [void][datetime]::ParseExact($checkDate, 'yyyy-MM-dd', $null)
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "'$prefix' är inte ett giltigt datum.",
            "RENAME", 'OK', 'Warning') | Out-Null
        return
    }

    Set-Busy $true
    Log ("Skannar '{0}' ..." -f (Split-Path -Leaf $script:Head))
    $slMain.Text = "Skannar..."

    $items = Scan-DatePrefixItems -Base $script:Head -Target $prefix
    $script:Items = $items

    if (-not $items -or $items.Count -eq 0) {
        $lblSummary.Text = "Inga filer/mappar med datum-prefix hittades."
        Log "Inga mål hittades."
        $slMain.Text = "Inga mål."
        Set-Busy $false
        return
    }

    # Fyll grid
    $grid.SuspendLayout()
    foreach ($it in $items) {
        $rowIdx = $grid.Rows.Add($it.Type, $it.OldName, $it.NewName)

        # Markera rader som INTE ändras med grå text
        if ($it.OldName -ieq $it.NewName) {
            $grid.Rows[$rowIdx].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Silver
        }

        # Färgkoda Typ-kolumnen
        if ($it.IsDir) {
            $grid.Rows[$rowIdx].Cells[0].Style.ForeColor = $AccentBlue
        }
    }
    $grid.ResumeLayout()

    $dirs  = @($items | Where-Object { $_.IsDir }).Count
    $files = @($items | Where-Object { -not $_.IsDir }).Count
    $unchanged = @($items | Where-Object { $_.OldName -ieq $_.NewName }).Count
    $willChange = $items.Count - $unchanged

    $lblSummary.Text = "$dirs mappar, $files filer ($willChange att ändra, $unchanged redan rätt)."
    Log ("Listat {0} poster ({1} att ändra)." -f $items.Count, $willChange)
    $slMain.Text = ("{0} poster listade." -f $items.Count)

    $btnRename.Enabled = ($willChange -gt 0)
    Set-Busy $false
})

# ── Byt namn ──
$btnRename.Add_Click({
    if ($script:BusyLock) { return }
    if (-not $script:Items -or $script:Items.Count -eq 0) { return }

    $prefix = $txtDate.Text.Trim()
    $dirs   = @($script:Items | Where-Object { $_.IsDir }).Count
    $files  = @($script:Items | Where-Object { -not $_.IsDir }).Count

    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Byta datum-prefix till '$prefix' för $dirs mappar och $files filer?`n`nDenna åtgärd kan inte ångras.",
        "Bekräfta", 'YesNo', 'Warning')
    if ($ans -ne 'Yes') { Log "Avbrutet av användaren."; return }

    Set-Busy $true
    Log "Byter datum-prefix... (fönstret kan frysa tillfälligt)"
    $slMain.Text = "Byter namn..."

    $res = Apply-Rename -Base $script:Head -Items $script:Items -Target $prefix

    Log ("Mappar: {0} OK, {1} fel." -f $res.DirOK, $res.DirErr)
    Log ("Filer:  {0} OK, {1} fel." -f $res.FileOK, $res.FileErr)

    if ($res.LockedFiles -and $res.LockedFiles.Count -gt 0) {
        Log ("[!] {0} fil(er) var låsta:" -f $res.LockedFiles.Count)
        ($res.LockedFiles | Select-Object -First 5) | ForEach-Object { Log "  [LAST] $_" }
        if ($res.LockedFiles.Count -gt 5) {
            Log ("  ...och {0} till." -f ($res.LockedFiles.Count - 5))
        }
    }

    $total = $res.DirOK + $res.FileOK
    $errors = $res.DirErr + $res.FileErr
    $lblSummary.Text = "Klart! $total bytta, $errors fel."
    $slMain.Text = "Klart."
    Log "Datum-prefix byte slutfört."

    # Skanna om automatiskt för att visa kvarvarande poster
    $script:Items = @()
    $grid.Rows.Clear()
    $btnRename.Enabled = $false
    Set-Busy $false

    if ($errors -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "$total poster bytta.`n$errors fel -- se loggen.",
            "Resultat", 'OK', 'Warning') | Out-Null
    }

    # Kör ny skanning så användaren ser aktuellt läge
    $btnScan.PerformClick()
})

# =====================[ Startup ]=====================

# Auto-fylla om -HeadFolder angavs
if ($HeadFolder -and (Test-Path -LiteralPath $HeadFolder -PathType Container)) {
    $script:Head = (Resolve-Path -LiteralPath $HeadFolder).ProviderPath
    $listHits.Items.Clear()
    [void]$listHits.Items.Add( (Split-Path -Leaf $script:Head) )
    $listHits.SelectedIndex = 0
    $slMain.Text = ("Vald: {0}" -f (Split-Path -Leaf $script:Head))
}

if ($Lsp) { $txtLsp.Text = $Lsp }
if ($NewDatePrefix) { $txtDate.Text = $NewDatePrefix }

[void]$form.ShowDialog()
