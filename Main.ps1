#region Importering och config
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe = Join-Path $PSHome 'powershell.exe'
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}

Add-Type -AssemblyName System.Windows.Forms
try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.ComponentModel
try {
    Add-Type -AssemblyName 'Microsoft.VisualBasic' -ErrorAction SilentlyContinue
} catch {}

$scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
$ScriptRootPath = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($scriptPath))
$PSScriptRoot = $ScriptRootPath
try {
    $cwd = (Get-Location).Path
    Write-Host ("[EXEC] Script={0} | ScriptRoot={1} | CWD={2}" -f $scriptPath, $ScriptRootPath, $cwd)
} catch {}
$modulesRoot = Join-Path $ScriptRootPath 'Modules'


# samlade layoutkonstanter, lite rörigt
$Layout = @{ SignatureCell = 'B47' }

. (Join-Path $modulesRoot 'Config.ps1') -ScriptRoot $ScriptRootPath
. (Join-Path $modulesRoot 'Splash.ps1')
. (Join-Path $modulesRoot 'UiStyling.ps1')
. (Join-Path $modulesRoot 'Logging.ps1')
. (Join-Path $modulesRoot 'SharePointClient.ps1')

try {
    $netRoot = ($env:IPT_NETWORK_ROOT + '').Trim()
    $iptRoot = ($global:IPT_ROOT_EFFECTIVE + '').Trim()
    $iptSrc  = ($global:IPT_ROOT_SOURCE + '').Trim()
    $logPath = ($global:LogPath + '').Trim()
    $jsonl   = ($global:StructuredLogPath + '').Trim()

    if (-not $netRoot) { $netRoot = '<empty>' }
    if (-not $iptRoot) { $iptRoot = '<empty>' }
    if (-not $iptSrc)  { $iptSrc  = '<empty>' }
    if (-not $logPath) { $logPath = '<empty>' }
    if (-not $jsonl)   { $jsonl   = '<empty>' }

    $msg = "Sanity: IPT_NETWORK_ROOT=$netRoot | IPT_ROOT_EFFECTIVE=$iptRoot | IPT_ROOT_SOURCE=$iptSrc | LogPath=$logPath | StructuredLogPath=$jsonl"
    try { Gui-Log -Text $msg -Severity 'Info' -Category 'SANITY' } catch { Write-Host $msg }
} catch { }

. (Join-Path $modulesRoot 'DataHelpers.ps1')
. (Join-Path $modulesRoot 'SignatureHelpers.ps1')
. (Join-Path $modulesRoot 'RuleEngine.ps1')

$global:SpEnabled = Get-ConfigFlag -Name 'EnableSharePoint' -Default $true -ConfigOverride $Config
$global:SpAutoConnect = if ($global:SpEnabled) {
    Get-ConfigFlag -Name 'EnableSharePointAutoConnect' -Default $true -ConfigOverride $Config
} else { $false }

$global:StartupReady = $true
$configStatus = $null

try {

    $configStatus = Test-Config
    if ($configStatus) {
        foreach ($err in $configStatus.Errors) { Gui-Log "❌ Konfig-fel: $err" 'Error' }
        foreach ($warn in $configStatus.Warnings) { Gui-Log "⚠️ Konfig-varning: $warn" 'Warn' }
        if (-not $configStatus.Ok) {
            $global:StartupReady = $false
            try { [System.Windows.Forms.MessageBox]::Show("Startkontroll misslyckades. Se logg för detaljer.","Startkontroll") | Out-Null } catch {}
        }
    }
} catch { Gui-Log "❌ Test-Config misslyckades: $($_.Exception.Message)" 'Error'; $global:StartupReady = $false }

$Host.UI.RawUI.WindowTitle = "Startar…"
Show-Splash "Laddar PowerShell…"
$env:PNPPOWERSHELL_UPDATECHECK = 'Off'  # stäng av update-check tidigt

$global:SpConnected = $false
$global:SpError     = $null

# Skydd mot dubbelklick / parallella anslutningsförsök
$script:SpConnectInProgress = $false

 # Capture UI runspace so BackgroundWorker threads can execute PowerShell scriptblocks
 try {
     $script:UiRunspace = [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace
     if (-not $script:UiRunspace) { $script:UiRunspace = $Host.Runspace }
 } catch {
     $script:UiRunspace = $null
 }

# Initiera dedikerad SharePoint-runspace (alla PnP-anrop ska gå via SharePointClient.ps1)
if ($global:SpEnabled) {
    try { Start-SPClient } catch { $global:SpError = "SP-runspace init misslyckades: $($_.Exception.Message)" }
}

if ($global:SpAutoConnect) {
    try {
        $null = Get-PackageProvider -Name "NuGet" -ForceBootstrap -ErrorAction SilentlyContinue
    } catch {}
}

if (-not $global:SpEnabled) {
    $global:SpError = 'SharePoint avstängt i Config'
    try { Gui-Log "ℹ️ SharePoint är avstängt i konfigurationen." 'Info' } catch {}
} else {
    # SpEnabled = true men AutoConnect = false → snabbstart utan PnP
    # Detta är INTE ett fel – bara att användaren ansluter manuellt vid behov.
    $global:SpError = $null
    try { Gui-Log "ℹ️ SharePoint: manuell anslutning (klicka '🔌 Anslut SP' vid behov)." 'Info' } catch {}
}
try { $null = Ensure-EPPlus -Version '4.5.3.3' } catch { Gui-Log "⚠️ EPPlus-förkontroll misslyckades: $($_.Exception.Message)" 'Warn' }

if ($global:SpAutoConnect -and -not $global:SpError) {
    try {
        Update-Splash "Ansluter till SharePoint"
        $r = Connect-SPClient -Url $global:SP_SiteUrl -Tenant $global:SP_Tenant -ClientId $global:SP_ClientId -CertificateBase64Encoded $global:SP_CertBase64
        if ($r -and $r.Ok) {
            $global:SpConnected = $true
            $global:SpError = $null
        } else {
            $global:SpConnected = $false
            $global:SpError = if ($r -and $r.Err) { $r.Err } else { 'Okänt fel' }
            Update-Splash ("Connect-SPClient misslyckades: " + $global:SpError)
        }
    } catch {
        $global:SpConnected = $false
        $global:SpError = "Connect-SPClient misslyckades: $($_.Exception.ToString())"
        Update-Splash $global:SpError
    }
}
#endregion Importering och config

#region GUI 

Update-Splash "Startar…"
Close-Splash
$form = New-Object System.Windows.Forms.Form
$form.Text = "$ScriptVersion"
$form.AutoScaleMode = 'Dpi'
$form.Size = New-Object System.Drawing.Size(860,870)
$form.MinimumSize = New-Object System.Drawing.Size(860,870)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.AutoScroll  = $false
$form.MaximizeBox = $false
$form.Padding     = New-Object System.Windows.Forms.Padding(8)
$form.Font        = New-Object System.Drawing.Font('Segoe UI',10)
$form.KeyPreview = $true
$form.add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $form.Close() } })

# Städa upp SharePoint-runspace när appen stängs
try {
    $form.Add_FormClosed({
        try { Stop-SPClient } catch {}
    })
} catch {}

# ---------- Menyrad ----------
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock='Top'; $menuStrip.GripStyle='Hidden'
$menuStrip.ImageScalingSize = New-Object System.Drawing.Size(20,20)
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(8,6,0,6)
$menuStrip.Font = New-Object System.Drawing.Font('Segoe UI',10)
$miArkiv   = New-Object System.Windows.Forms.ToolStripMenuItem('🗂️ Arkiv')
$miVerktyg = New-Object System.Windows.Forms.ToolStripMenuItem('🛠️ Verktyg')
$miSettings= New-Object System.Windows.Forms.ToolStripMenuItem('⚙️ Inställningar')
$miHelp    = New-Object System.Windows.Forms.ToolStripMenuItem('📖 Instruktioner')
$miAbout   = New-Object System.Windows.Forms.ToolStripMenuItem('ℹ️ Om')
$miScan  = New-Object System.Windows.Forms.ToolStripMenuItem('🔍 Sök filer')
$miBuild = New-Object System.Windows.Forms.ToolStripMenuItem('✅ Skapa rapport')
$miExit  = New-Object System.Windows.Forms.ToolStripMenuItem('❌ Avsluta')

# Rensa ev. gamla undermenyer
$miArkiv.DropDownItems.Clear()
$miVerktyg.DropDownItems.Clear()
$miSettings.DropDownItems.Clear()
$miHelp.DropDownItems.Clear()

# ----- Arkiv -----
$miNew         = New-Object System.Windows.Forms.ToolStripMenuItem('🆕 Nytt')
$miOpenRecent  = New-Object System.Windows.Forms.ToolStripMenuItem('📂 Öppna senaste rapport')
$miArkiv.DropDownItems.AddRange(@(
    $miNew,
    $miOpenRecent,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miExit
))

# ----- Verktyg -----
$miScript1   = New-Object System.Windows.Forms.ToolStripMenuItem('📜 Kontrollprovsfil-skript')
$miScript2   = New-Object System.Windows.Forms.ToolStripMenuItem('📜 Aktivera Kontrollprovsfil')
$miScript3   = New-Object System.Windows.Forms.ToolStripMenuItem('📅 Ändra datum på filer')
$miScript4   = New-Object System.Windows.Forms.ToolStripMenuItem('🗂️ AutoMappscript Control')
$miScript5   = New-Object System.Windows.Forms.ToolStripMenuItem('📄 AutoMappscript Dashboard')
$miToggleSign = New-Object System.Windows.Forms.ToolStripMenuItem('✅ Aktivera Seal Test-signatur')
$miVerktyg.DropDownItems.AddRange(@(
    $miScript1,
    $miScript2,
    $miScript3,
    $miScript4,
    $miScript5,
    $miToggleSign
))

# ----- Inställningar -----
$miTheme = New-Object System.Windows.Forms.ToolStripMenuItem('🎨 Tema')
$miLightTheme = New-Object System.Windows.Forms.ToolStripMenuItem('☀️ Ljust (default)')
$miDarkTheme  = New-Object System.Windows.Forms.ToolStripMenuItem('🌙 Mörkt')
$miTheme.DropDownItems.AddRange(@($miLightTheme,$miDarkTheme))
$miSettings.DropDownItems.Add($miTheme)

# ----- Instruktioner -----
$miShowInstr   = New-Object System.Windows.Forms.ToolStripMenuItem('📖 Visa instruktioner')
$miFAQ         = New-Object System.Windows.Forms.ToolStripMenuItem('❓ FAQ')
$miHelpDlg     = New-Object System.Windows.Forms.ToolStripMenuItem('🆘 Hjälp')
$miHelp.DropDownItems.AddRange(@($miShowInstr,$miFAQ,$miHelpDlg))

$miGenvagar = New-Object System.Windows.Forms.ToolStripMenuItem('🔗 Genvägar')
$ShortcutGroups = Get-ConfigValue -Name 'ShortcutGroups' -Default $null -ConfigOverride $Config
if (-not $ShortcutGroups) {
    # Fallback om config saknar genvägar
    $ShortcutGroups = @{
        '🗂️ IPT-mappar' = @(
            @{ Text='📂 IPT - PÅGÅENDE KÖRNINGAR';        Target='N:\QC\QC-1\IPT\2. IPT - PÅGÅENDE KÖRNINGAR' },
            @{ Text='📂 IPT - KLART FÖR SAMMANSTÄLLNING'; Target='N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING' },
            @{ Text='📂 IPT - KLART FÖR GRANSKNING';      Target='N:\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING' },
            @{ Text='📂 SPT Macro Assay';                 Target='N:\QC\QC-0\SPT\SPT macros\Assay' }
        )
        '📄 Dokument' = @(
            @{ Text='🧰 Utrustningslista';    Target=$UtrustningListPath },
            @{ Text='🧪 Kontrollprovsfil';    Target=$RawDataPath }
        )
        '🌐 Länkar' = @(
            @{ Text='⚡ IPT App';              Target='https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83' },
            @{ Text='🌐 MES';                  Target='http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM' },
            @{ Text='🌐 CSV Uploader';         Target='http://auw2wgxtpap01.cepaws.com/Welcome.aspx' },
            @{ Text='🌐 BMRAM';                Target='https://cepheid62468.coolbluecloud.com/' },
            @{ Text='🌐 Agile';                Target='https://agileprod.cepheid.com/Agile/default/login-cms.jsp' }
        )
    }
}

foreach ($grp in $ShortcutGroups.GetEnumerator()) {

    $grpMenu = New-Object System.Windows.Forms.ToolStripMenuItem($grp.Key)
    foreach ($entry in $grp.Value) { Add-ShortcutItem -Parent $grpMenu -Text $entry.Text -Target $entry.Target }
    [void]$miGenvagar.DropDownItems.Add($grpMenu)

}

$miOm = New-Object System.Windows.Forms.ToolStripMenuItem('ℹ️ Om det här verktyget'); $miAbout.DropDownItems.Add($miOm)
$menuStrip.Items.AddRange(@($miArkiv,$miVerktyg,$miGenvagar,$miSettings,$miHelp,$miAbout))
$form.MainMenuStrip=$menuStrip

# ---------- Header ----------
$panelHeader = New-Object System.Windows.Forms.Panel
$panelHeader.Dock='Top'; $panelHeader.Height=82
$panelHeader.BackColor=[System.Drawing.Color]::SteelBlue
$panelHeader.Padding = New-Object System.Windows.Forms.Padding(10,8,10,8)

$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Dock = 'Left'
$picLogo.Width = 50
$picLogo.BorderStyle = 'FixedSingle'
$picLogo.SizeMode = 'Zoom'
if (Test-Path $ikonSokvag) { $picLogo.Image = [System.Drawing.Image]::FromFile($ikonSokvag) }

$panelText = New-Object System.Windows.Forms.Panel
$panelText.Dock = 'Fill'
$panelText.BackColor = $panelHeader.BackColor

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "$ScriptVersion - IPTCompile"
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',13)
$lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTitle.Dock = 'Top'
$lblTitle.Height = 36
$lblTitle.Padding = New-Object System.Windows.Forms.Padding(8,0,0,0)

$lblUpdate = New-Object System.Windows.Forms.Label
$lblUpdate.Text      = 'DONE: Additional QC Data Check.'
$lblUpdate.ForeColor = [System.Drawing.Color]::FromArgb(200,210,230)
$lblUpdate.Font      = New-Object System.Drawing.Font('Segoe UI',8)
$lblUpdate.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblUpdate.Dock      = 'Top'
$lblUpdate.Height    = 16
$lblUpdate.Padding   = New-Object System.Windows.Forms.Padding(8,0,0,0)

$lblExtra = New-Object System.Windows.Forms.Label
$lblExtra.Text = 'TODO1: Fixa utrustningslistan. TODO2: Fixa Resp Panel.'
$lblExtra.ForeColor = [System.Drawing.Color]::FromArgb(200,210,230)
$lblExtra.Font = New-Object System.Drawing.Font('Segoe UI',8)
$lblExtra.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblExtra.Dock = 'Top'
$lblExtra.Height = 14
$lblExtra.Padding = New-Object System.Windows.Forms.Padding(8,1,0,0)

$panelText.Controls.Add($lblExtra)
$panelText.Controls.Add($lblUpdate)
$panelText.Controls.Add($lblTitle)

$panelHeader.Controls.Add($panelText)
$panelHeader.Controls.Add($picLogo)

$panelHeader.Add_Resize({
    $innerH = $panelHeader.Height - $panelHeader.Padding.Top - $panelHeader.Padding.Bottom
    $picLogo.Height = $innerH
    $picLogo.Width  = $innerH 
})

$form.add_FormClosed({ try { if ($picLogo.Image) { $picLogo.Image.Dispose() } } catch {} })

# ---------- Sök-rad ----------

$tlSearch = New-Object System.Windows.Forms.TableLayoutPanel
$tlSearch.Dock='Top'; $tlSearch.AutoSize=$true; $tlSearch.AutoSizeMode='GrowAndShrink'
$tlSearch.Padding = New-Object System.Windows.Forms.Padding(0,10,0,8)
$tlSearch.ColumnCount=3
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,130)))

$lblLSP = New-Object System.Windows.Forms.Label
$lblLSP.Text='LSP:'; $lblLSP.Anchor='Left'; $lblLSP.AutoSize=$true
$lblLSP.Margin = New-Object System.Windows.Forms.Padding(0,6,8,0)
$txtLSP = New-Object System.Windows.Forms.TextBox
$txtLSP.Dock='Fill'
$txtLSP.Margin = New-Object System.Windows.Forms.Padding(0,2,10,2)
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text='Sök filer'; $btnScan.Dock='Fill'; Set-AccentButton $btnScan -Primary
$btnScan.Margin= New-Object System.Windows.Forms.Padding(0,2,0,2)

$tlSearch.Controls.Add($lblLSP,0,0)
$tlSearch.Controls.Add($txtLSP,1,0)
$tlSearch.Controls.Add($btnScan,2,0)

$pLog = New-Object System.Windows.Forms.Panel
$pLog.Dock='Top'; $pLog.Height=220; $pLog.Padding=New-Object System.Windows.Forms.Padding(0,0,0,8)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline=$true; $outputBox.ScrollBars='Vertical'; $outputBox.ReadOnly=$true
$outputBox.BackColor='White'; $outputBox.Dock='Fill'
$outputBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
$pLog.Controls.Add($outputBox)
try { Set-LogOutputControl -Control $outputBox } catch {}

$grpPick = New-Object System.Windows.Forms.GroupBox
$grpPick.Text='Välj filer för rapport'
$grpPick.Dock='Top'
$grpPick.Padding = New-Object System.Windows.Forms.Padding(10,12,10,14)
$grpPick.AutoSize=$false
$grpPick.Height = (78*3) + $grpPick.Padding.Top + $grpPick.Padding.Bottom +15

$tlPick = New-Object System.Windows.Forms.TableLayoutPanel
$tlPick.Dock='Fill'; $tlPick.ColumnCount=3; $tlPick.RowCount=3
$tlPick.GrowStyle=[System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,100)))
for($i=0;$i -lt 3;$i++){ [void]$tlPick.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,78))) }

function New-ListRow {
    param([string]$labelText,[ref]$lbl,[ref]$clb,[ref]$btn)
    $lbl.Value = New-Object System.Windows.Forms.Label
    $lbl.Value.Text=$labelText
    $lbl.Value.Anchor='Left'
    $lbl.Value.AutoSize=$true
    $lbl.Value.Margin=New-Object System.Windows.Forms.Padding(0,12,6,0)
    $clb.Value = New-Object System.Windows.Forms.CheckedListBox
    $clb.Value.Dock='Fill'
    $clb.Value.Margin=New-Object System.Windows.Forms.Padding(0,6,8,6)
    $clb.Value.Height=70
    $clb.Value.IntegralHeight=$false
    $clb.Value.CheckOnClick = $true
    $clb.Value.DisplayMember = 'Name'

    $btn.Value = New-Object System.Windows.Forms.Button
    $btn.Value.Text='Bläddra…'
    $btn.Value.Dock='Fill'
    $btn.Value.Margin=New-Object System.Windows.Forms.Padding(0,6,0,6)
    Set-AccentButton $btn.Value
}

$lblCsv=$null;$clbCsv=$null;$btnCsvBrowse=$null
New-ListRow -labelText 'CSV-fil:' -lbl ([ref]$lblCsv) -clb ([ref]$clbCsv) -btn ([ref]$btnCsvBrowse)
$lblNeg=$null;$clbNeg=$null;$btnNegBrowse=$null
New-ListRow -labelText 'Seal Test Neg:' -lbl ([ref]$lblNeg) -clb ([ref]$clbNeg) -btn ([ref]$btnNegBrowse)
$lblPos=$null;$clbPos=$null;$btnPosBrowse=$null
New-ListRow -labelText 'Seal Test Pos:' -lbl ([ref]$lblPos) -clb ([ref]$clbPos) -btn ([ref]$btnPosBrowse)

try {
    if ($tlPick.RowCount -lt 4) {
        $tlPick.RowCount = 4
        for ($i=$tlPick.RowStyles.Count; $i -lt 4; $i++) {
            $null = $tlPick.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 78)) )
        }
        $grpPick.Height = (78*4) + $grpPick.Padding.Top + $grpPick.Padding.Bottom + 15
    }
} catch {}

$lblLsp = $null; $clbLsp = $null; $btnLspBrowse = $null
New-ListRow -labelText 'Worksheet:' -lbl ([ref]$lblLsp) -clb ([ref]$clbLsp) -btn ([ref]$btnLspBrowse)

$tlPick.Controls.Add($lblLsp,  0, 3)
$tlPick.Controls.Add($clbLsp,  1, 3)
$tlPick.Controls.Add($btnLspBrowse, 2, 3)

$clbLsp.add_ItemCheck({
    param($s,$e)
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        for ($i=0; $i -lt $s.Items.Count; $i++) {
            if ($i -ne $e.Index) { $s.SetItemChecked($i, $false) }
        }
    }
})

$btnLspBrowse.Add_Click({
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
        $dlg.Title  = "Välj LSP Worksheet"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $f = Get-Item -LiteralPath $dlg.FileName
            Add-CLBItems -clb $clbLsp -files @($f) -AutoCheckFirst
            if (Get-Command Update-StatusBar -ErrorAction SilentlyContinue) { Update-StatusBar }
        }
    } catch {
        Gui-Log ("⚠️ LSP-browse fel: " + $_.Exception.Message) 'Warn'
    }
})

# Lägg in i tabellen
$tlPick.Controls.Add($lblCsv,0,0); $tlPick.Controls.Add($clbCsv,1,0); $tlPick.Controls.Add($btnCsvBrowse,2,0)
$tlPick.Controls.Add($lblNeg,0,1); $tlPick.Controls.Add($clbNeg,1,1); $tlPick.Controls.Add($btnNegBrowse,2,1)
$tlPick.Controls.Add($lblPos,0,2); $tlPick.Controls.Add($clbPos,1,2); $tlPick.Controls.Add($btnPosBrowse,2,2)
$grpPick.Controls.Add($tlPick)

# ---------- Signatur ----------
$grpSign = New-Object System.Windows.Forms.GroupBox
$grpSign.Text = "Lägg till signatur i Seal Test-filerna"
$grpSign.Dock='Top'
$grpSign.Padding = New-Object System.Windows.Forms.Padding(10,8,10,10)
$grpSign.AutoSize = $false
$grpSign.Height = 88

$tlSign = New-Object System.Windows.Forms.TableLayoutPanel
$tlSign.Dock='Fill'; $tlSign.ColumnCount=2; $tlSign.RowCount=2
[void]$tlSign.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlSign.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlSign.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,28)))
[void]$tlSign.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,28)))

$lblSigner = New-Object System.Windows.Forms.Label
$lblSigner.Text = 'Fullständigt namn, signatur och datum:'
$lblSigner.Anchor='Left'; $lblSigner.AutoSize=$true

$txtSigner = New-Object System.Windows.Forms.TextBox
$txtSigner.Dock='Fill'; $txtSigner.Margin = New-Object System.Windows.Forms.Padding(6,2,0,2)
$chkWriteSign = New-Object System.Windows.Forms.CheckBox
$chkWriteSign.Text = 'Signera Seal Test-Filerna'
$chkWriteSign.Anchor='Left'
$chkWriteSign.AutoSize = $true

$chkOverwriteSign = New-Object System.Windows.Forms.CheckBox
$chkOverwriteSign.Text = 'Aktivera'

$chkOverwriteSign.Anchor='Left'
$chkOverwriteSign.AutoSize = $true
$chkOverwriteSign.Enabled = $false
$chkWriteSign.add_CheckedChanged({ $chkOverwriteSign.Enabled = $chkWriteSign.Checked })

$tlSign.Controls.Add($lblSigner,0,0); $tlSign.Controls.Add($txtSigner,1,0)
$tlSign.Controls.Add($chkWriteSign,0,1); $tlSign.Controls.Add($chkOverwriteSign,1,1)
$grpSign.Controls.Add($tlSign)

$grpSign.Visible = $false
$baseHeight = $form.Height

# ---------- Rapport-utdata ----------
# (GUI-val borttaget) Rapporten sparas alltid temporärt och SharePoint Info inkluderas alltid.
$grpSave = $null
$rbSaveInLsp = $null
$rbTempOnly = $null
$chkSharePointInfo = $null

# ---------- Primärknapp ----------

$btnBuild = New-Object System.Windows.Forms.Button
$btnBuild.Text='Skapa rapport'; $btnBuild.Dock='Top'; $btnBuild.Height=40
$btnBuild.Margin = New-Object System.Windows.Forms.Padding(0,16,0,8)
$btnBuild.Enabled=$false; Set-AccentButton $btnBuild -Primary

# ---------- Statusrad ----------
$status = New-Object System.Windows.Forms.StatusStrip
$status.SizingGrip=$false; $status.Dock='Bottom'; $status.Font=New-Object System.Drawing.Font('Segoe UI',9)
$status.ShowItemToolTips = $true
$slCount = New-Object System.Windows.Forms.ToolStripStatusLabel; $slCount.Text='0 filer valda'; $slCount.Spring=$false
$slWork  = New-Object System.Windows.Forms.ToolStripStatusLabel
$slWork.Text   = ''
$slWork.Spring = $true

$pbWork = New-Object System.Windows.Forms.ToolStripProgressBar
$pbWork.Visible = $false
$pbWork.Style   = 'Marquee'
$pbWork.MarqueeAnimationSpeed = 30
$pbWork.AutoSize = $false
$pbWork.Width    = 140

$slSpacer= New-Object System.Windows.Forms.ToolStripStatusLabel; $slSpacer.Spring=$true

# --- Klickbar SharePoint-länk ---
$slBatchLink = New-Object System.Windows.Forms.ToolStripStatusLabel
$slBatchLink.IsLink   = $true
$slBatchLink.Text     = 'SharePoint: —'
$slBatchLink.Enabled  = $false
$slBatchLink.Tag      = $null
$slBatchLink.ToolTipText = 'Direktlänk aktiveras när Batch# hittas i filer.'
$slBatchLink.add_Click({
    if ($this.Enabled -and $this.Tag) {
        try { Start-Process $this.Tag } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna:`n$($this.Tag)`n$($_.Exception.Message)","Länk") | Out-Null
        }
    }
})

# --- SharePoint On-Demand Connect-knapp ---
$script:btnSpConnect = New-Object System.Windows.Forms.ToolStripButton
$script:btnSpConnect.DisplayStyle = 'Text'
$script:btnSpConnect.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)

if ($global:SpConnected) {
    $script:btnSpConnect.Text = '✅ SP'
    $script:btnSpConnect.ToolTipText = 'SharePoint är anslutet.'
    $script:btnSpConnect.Enabled = $false
} elseif ($global:SpEnabled) {
    $script:btnSpConnect.Text = '🔌 Anslut SP'
    $script:btnSpConnect.ToolTipText = 'Klicka för att ansluta till SharePoint (tar ~10–15 sek).'
    $script:btnSpConnect.Enabled = $true
} else {
    $script:btnSpConnect.Text = '⛔ SP av'
    $script:btnSpConnect.ToolTipText = 'SharePoint är avstängt i konfigurationen.'
    $script:btnSpConnect.Enabled = $false
}

$script:btnSpConnect.add_Click({
    if (-not $global:SpEnabled) { return }
    if ($global:SpConnected) { return }
    if ($script:SpConnectInProgress) { return }

    $script:SpConnectInProgress = $true

    # Uppdatera knappen och tvinga SP
    $this.Text        = '⏳ Ansluter…'
    $this.Enabled     = $false
    $this.ToolTipText = 'Ansluter till SharePoint – vänta…'

    # Splash visas på UI-tråden (GUI ska INTE frysa under själva anslutningen)
    Show-Splash 'Ansluter till SharePoint…'

    # Kör anslutningen i bakgrunden för att undvika UI-freeze
    $bw = New-Object System.ComponentModel.BackgroundWorker
    $bw.WorkerReportsProgress = $true

    $bw.add_DoWork({
        param($sender,$e)
     if ($script:UiRunspace) {
         [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace = $script:UiRunspace
     }

        $ok  = $false
        $err = $null

        try {
            $sender.ReportProgress(0, 'Loggar in…')
            $r = Connect-SPClient -Url $global:SP_SiteUrl -Tenant $global:SP_Tenant -ClientId $global:SP_ClientId -CertificateBase64Encoded $global:SP_CertBase64
            if ($r -and $r.Ok) {
                $ok = $true
            } else {
                $ok = $false
                $err = if ($r -and $r.Err) { $r.Err } else { 'Okänt fel' }
            }
        } catch {
            # Spara full feltext (mer användbar än bara Message)
            try { $err = $_.Exception.ToString() } catch { $err = $_ | Out-String }
        }

        $e.Result = [pscustomobject]@{ Ok=$ok; Err=$err }
    })

    $bw.add_ProgressChanged({
        param($sender,$e)
        if ($e.UserState) {
            try { Update-Splash ([string]$e.UserState) } catch {}
        }
    })

    $bw.add_RunWorkerCompleted({
        param($sender,$e)

        $script:SpConnectInProgress = $false
        Close-Splash

        # Först: worker-exception (ska inte hända, men hantera robust)
        # OBS: $e.Error ÄR redan Exception-objektet (System.Exception).
        #      .Exception finns inte (det heter .InnerException) → null → krasch.
        if ($e.Error) {
            $global:SpConnected = $false
            $global:SpError     = $e.Error.ToString()

            $script:btnSpConnect.Text        = '❌ SP – retry?'
            $script:btnSpConnect.ToolTipText = "Misslyckades: $($global:SpError)`nKlicka för att försöka igen."
            $script:btnSpConnect.Enabled     = $true
            Gui-Log ("⚠️ SharePoint-anslutning misslyckades (worker exception): " + $global:SpError) 'Warn'
            return
        }

        $r = $e.Result
        if ($r -and $r.Ok) {
            $global:SpConnected = $true
            $global:SpError     = $null

            $script:btnSpConnect.Text        = '✅ SP'
            $script:btnSpConnect.ToolTipText = 'SharePoint anslutet.'
            $script:btnSpConnect.Enabled     = $false
            Gui-Log '✅ SharePoint anslutet (on-demand).' 'Info'
        } else {
            $global:SpConnected = $false
            $global:SpError     = if ($r -and $r.Err) { $r.Err } else { 'Okänt fel' }

            $script:btnSpConnect.Text        = '❌ SP – retry?'
            $script:btnSpConnect.ToolTipText = "Misslyckades: $($global:SpError)`nKlicka för att försöka igen."
            $script:btnSpConnect.Enabled     = $true
            Gui-Log ("⚠️ SharePoint-anslutning misslyckades: " + $global:SpError) 'Warn'
        }
    })

    $bw.RunWorkerAsync()
})

$status.Items.AddRange(@($slCount,$slWork,$pbWork,$script:btnSpConnect,$slBatchLink))
$tsc = New-Object System.Windows.Forms.ToolStripContainer
$tsc.Dock = 'Fill'
$tsc.LeftToolStripPanelVisible  = $false
$tsc.RightToolStripPanelVisible = $false

$form.SuspendLayout()
$form.Controls.Clear()
$form.Controls.Add($tsc)

# Meny högst upp
$tsc.TopToolStripPanel.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip

# Status längst ner
$tsc.BottomToolStripPanel.Controls.Add($status)

# Content i mitten
$content = New-Object System.Windows.Forms.Panel
$content.Dock='Fill'
$content.BackColor = $form.BackColor
$tsc.ContentPanel.Controls.Add($content)

# Dock=Top: nedersta först
$content.SuspendLayout()
$content.Controls.Add($btnBuild)
$content.Controls.Add($grpSign)
$content.Controls.Add($grpPick)
$content.Controls.Add($pLog)
$content.Controls.Add($tlSearch)
$content.Controls.Add($panelHeader)
$content.ResumeLayout()
$form.ResumeLayout()
$form.PerformLayout()
$form.AcceptButton = $btnScan

#endregion GUI Construction

function Add-CLBItems {
    param([System.Windows.Forms.CheckedListBox]$clb,[System.IO.FileInfo[]]$files,[switch]$AutoCheckFirst)
    $clb.BeginUpdate()
    $clb.Items.Clear()
    foreach($f in $files){
        if ($f -isnot [System.IO.FileInfo]) { try { $f = Get-Item -LiteralPath $f } catch { continue } }
        [void]$clb.Items.Add($f, $false)
    }
    $clb.EndUpdate()
    if ($AutoCheckFirst -and $clb.Items.Count -gt 0) { $clb.SetItemChecked(0,$true) }
    Update-StatusBar
}

function Get-CheckedFilePath { param([System.Windows.Forms.CheckedListBox]$clb)
    for($i=0;$i -lt $clb.Items.Count;$i++){
        if ($clb.GetItemChecked($i)) {
            $fi = [System.IO.FileInfo]$clb.Items[$i]
            return $fi.FullName
        }
    }
    return $null
}

function Clear-GUI {
    $txtLSP.Text = ''
    $txtSigner.Text = ''
    $chkWriteSign.Checked = $false
    $chkOverwriteSign.Checked = $false
    Add-CLBItems -clb $clbCsv -files @()
    Add-CLBItems -clb $clbNeg -files @()
    Add-CLBItems -clb $clbPos -files @()
    Add-CLBItems -clb $clbLsp -files @()
    $outputBox.Clear()
    Update-BuildEnabled
    Gui-Log "🧹 GUI rensat." 'Info'
    Update-BatchLink
}

$onExclusive = {
    $clb = $this
    if ($_.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            if ($i -ne $_.Index -and $clb.GetItemChecked($i)) { $clb.SetItemChecked($i, $false) }
        }
    }
    $clb.BeginInvoke([Action]{ Update-BuildEnabled }) | Out-Null
}
$clbCsv.add_ItemCheck($onExclusive)
$clbNeg.add_ItemCheck($onExclusive)
$clbPos.add_ItemCheck($onExclusive)

function Get-SelectedFileCount {
    $n=0
    if (Get-CheckedFilePath $clbCsv) { $n++ }
    if (Get-CheckedFilePath $clbNeg) { $n++ }
    if (Get-CheckedFilePath $clbPos) { $n++ }
    if (Get-CheckedFilePath $clbLsp) { $n++ }
    return $n
}

function Update-StatusBar { $slCount.Text = "$(Get-SelectedFileCount) filer valda" }

function Invoke-UiPump {
    try { [System.Windows.Forms.Application]::DoEvents() } catch {}
}

function Set-UiBusy {
    param(
        [Parameter(Mandatory)][bool]$Busy,
        [string]$Message = ''
    )
    try {
        if ($Busy) {
            $slWork.Text = $Message
            $pbWork.Visible = $true
            $pbWork.Style   = 'Marquee'
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnScan.Enabled  = $false
            $btnBuild.Enabled = $false
        } else {
            $pbWork.Visible = $false
            $pbWork.Style   = 'Blocks'
            $pbWork.Value   = 0
            $slWork.Text = ''
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnScan.Enabled = $true
            Update-BuildEnabled
        }
        $status.Refresh()
        $form.Refresh()
        Invoke-UiPump
    } catch {}
}

function Set-UiStep {
    param(
        [Parameter(Mandatory)][int]$Percent,
        [string]$Message = ''
    )
    try {
        if ($Percent -lt 0) { $Percent = 0 }
        if ($Percent -gt 100) { $Percent = 100 }
        $slWork.Text = $Message
        $pbWork.Visible = $true
        $pbWork.Style   = 'Blocks'
        $pbWork.Minimum = 0
        $pbWork.Maximum = 100
        if ($pbWork.Value -ne $Percent) { $pbWork.Value = $Percent }
        $status.Refresh()
        $form.Refresh()
        Invoke-UiPump
    } catch {}
}


function Update-BuildEnabled {
    $btnBuild.Enabled = ((Get-CheckedFilePath $clbNeg) -and (Get-CheckedFilePath $clbPos))
    Update-StatusBar
}

$script:LastScanResult = $null

# Reentrancy guards (DoEvents can otherwise allow nested clicks)
$script:ScanInProgress  = $false
$script:BuildInProgress = $false

function Get-BatchLinkInfo {
    param(
        [string]$SealPosPath,
        [string]$SealNegPath,
        [string]$Lsp
    )

    $batch = $null
    try { if ($SealPosPath) { $batch = Get-BatchNumberFromSealFile $SealPosPath } } catch {}
    if (-not $batch) {
        try { if ($SealNegPath) { $batch = Get-BatchNumberFromSealFile $SealNegPath } } catch {}
    }

    $batchEsc = if ($batch) { [uri]::EscapeDataString($batch) } else { '' }
    $lspEsc   = if ($Lsp)   { [uri]::EscapeDataString($Lsp) }   else { '' }

    $url = if ($SharePointBatchLinkTemplate) {
        ($SharePointBatchLinkTemplate -replace '\{BatchNumber\}', $batchEsc) -replace '\{LSP\}', $lspEsc
    } else {
        "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/AllItems.aspx?view=7&q=$batchEsc"
    }
    $linkText = if ($batch) { "Öppna $batch" } else { 'Ingen batch funnen' }

    return [pscustomobject]@{
        Batch    = $batch
        Url      = $url
        LinkText = $linkText
    }
}

function Assert-StartupReady {
    if ($global:StartupReady) { return $true }
    Gui-Log "❌ Startkontroll misslyckades. Åtgärda konfigurationsfel innan du fortsätter." 'Error'
    return $false
}

function Find-LspFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Lsp,
        [Parameter(Mandatory)] [string[]]$Roots
    )

    $lspRaw = ($Lsp + '').Trim()
    if (-not $lspRaw) { return $null }

    # Tillåt att användaren skriver t.ex. "#38401" eller "38401" eller "LSP 38401".
    $lspDigits = ($lspRaw -replace '\D', '')
    if (-not $lspDigits) { return $null }

    # Matcha som "eget tal": undvik att 3840 matchar 38401 osv.
    $rxToken = "(?<!\d)#?\s*$lspDigits(?!\d)"

    # Snabb väg: använd -Filter först (fil-systemet kan optimera) + matcha med rxToken för att undvika falska träffar.
    $filters = @("*$lspDigits*", "*#$lspDigits*")

    foreach ($root in $Roots) {
        if (-not $root) { continue }
        if (-not (Test-Path -LiteralPath $root)) { continue }

        # 1) Försök hitta i översta nivån (vanligaste fallet)
        foreach ($f in $filters) {
            try {
                $hit = Get-ChildItem -LiteralPath $root -Directory -Filter $f -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -match $rxToken } |
                       Select-Object -First 1
                if ($hit) { return $hit }
            } catch {}
        }

        # 2) Försök med -Recurse + -Filter
        foreach ($f in $filters) {
            try {
                $hit = Get-ChildItem -LiteralPath $root -Directory -Recurse -Filter $f -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -match $rxToken } |
                       Select-Object -First 1
                if ($hit) { return $hit }
            } catch {}
        }

        # 3) Fallback
        try {
            $hit = Get-ChildItem -LiteralPath $root -Directory -Recurse -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -match $rxToken } |
                   Select-Object -First 1
            if ($hit) { return $hit }
        } catch {}
    }

    return $null
}

#region Event Handlers
$miScan.add_Click({ $btnScan.PerformClick() })
$miBuild.add_Click({ if ($btnBuild.Enabled) { $btnBuild.PerformClick() } })
$miExit.add_Click({ $form.Close() })
$miNew.add_Click({ Clear-GUI })

$miOpenRecent.add_Click({
    if ($global:LastReportPath -and (Test-Path -LiteralPath $global:LastReportPath)) {
        try { Start-Process -FilePath $global:LastReportPath } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna rapporten:\n$($_.Exception.Message)","Öppna senaste rapport") | Out-Null
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ingen rapport har genererats i denna session.","Öppna senaste rapport") | Out-Null
    }
})

# Skript1..3
$miScript1.add_Click({
    $p = $Script1Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("Ange sökvägen till Skript1 i variabeln `$Script1Path.","Skript1") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("Filen hittades inte:\n$Script1Path","Skript1") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript1") | Out-Null } }
    }
})

$miScript2.add_Click({
    $p = $Script2Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("Ange sökvägen till Skript2 i variabeln `$Script2Path.","Skript2") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("Filen hittades inte:\n$Script2Path","Skript2") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript2") | Out-Null } }
    }
})

$miScript3.add_Click({
    $p = $Script3Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript3") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript3") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript3") | Out-Null } }
    }
})

$miScript4.add_Click({
    $p = $Script4Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript4") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript4") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript4") | Out-Null } }
    }
})

$miScript5.add_Click({
    $p = $Script5Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript5") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript5") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript5") | Out-Null } }
    }
})

$miToggleSign.add_Click({
    $lsp = $txtLSP.Text.Trim()
    if (-not $lsp) {
        Gui-Log "⚠️ Ange och sök ett LSP först innan du aktiverar Seal Test-signatur." 'Warn'
        return
    }
    $selNeg = Get-CheckedFilePath $clbNeg
    $selPos = Get-CheckedFilePath $clbPos
    if (-not $selNeg -or -not $selPos) {
        Gui-Log "⚠️ Du måste först välja både Seal Test NEG och POS innan Seal Test-signatur kan aktiveras." 'Warn'
        return
    }
    $grpSign.Visible = -not $grpSign.Visible
    if ($grpSign.Visible) {
        $form.Height = $baseHeight + $grpSign.Height + 40
        $miToggleSign.Text  = '❌ Dölj Seal Test-signatur'
    }
    else {
        $form.Height = $baseHeight
        $miToggleSign.Text  = '✅ Aktivera Seal Test-signatur'
    }
})

function Set-Theme {
    param([string]$Theme)
    if ($Theme -eq 'dark') {
        $global:CurrentTheme = 'dark'
        $form.BackColor        = [System.Drawing.Color]::FromArgb(35,35,35)
        $content.BackColor     = $form.BackColor
        $panelHeader.BackColor = [System.Drawing.Color]::DarkSlateBlue
        $pLog.BackColor        = [System.Drawing.Color]::FromArgb(45,45,45)
        $grpPick.BackColor     = $form.BackColor  
        $grpSign.BackColor     = $form.BackColor
        if ($grpSave) { $grpSave.BackColor = $form.BackColor }
        $tlSearch.BackColor    = $form.BackColor
        $outputBox.BackColor   = [System.Drawing.Color]::FromArgb(55,55,55)
        $outputBox.ForeColor   = [System.Drawing.Color]::White
        $lblLSP.ForeColor      = [System.Drawing.Color]::White
        $lblCsv.ForeColor      = [System.Drawing.Color]::White
        $lblNeg.ForeColor      = [System.Drawing.Color]::White
        $lblPos.ForeColor      = [System.Drawing.Color]::White
        if ($lblLsp) { $lblLsp.ForeColor = [System.Drawing.Color]::White }
        $grpPick.ForeColor     = [System.Drawing.Color]::White
        $grpSign.ForeColor     = [System.Drawing.Color]::White
        if ($grpSave) { $grpSave.ForeColor = [System.Drawing.Color]::White }
        $pLog.ForeColor        = [System.Drawing.Color]::White
        $tlSearch.ForeColor    = [System.Drawing.Color]::White
    } else {
        $global:CurrentTheme = 'light'
        $form.BackColor        = [System.Drawing.Color]::WhiteSmoke
        $content.BackColor     = $form.BackColor
        $panelHeader.BackColor = [System.Drawing.Color]::SteelBlue
        $pLog.BackColor        = [System.Drawing.Color]::White
        $grpPick.BackColor     = $form.BackColor
        $grpSign.BackColor     = $form.BackColor
        if ($grpSave) { $grpSave.BackColor = $form.BackColor }
        $tlSearch.BackColor    = $form.BackColor
        $outputBox.BackColor   = [System.Drawing.Color]::White
        $outputBox.ForeColor   = [System.Drawing.Color]::Black
        $lblLSP.ForeColor      = [System.Drawing.Color]::Black
        $lblCsv.ForeColor      = [System.Drawing.Color]::Black
        $lblNeg.ForeColor      = [System.Drawing.Color]::Black
        $lblPos.ForeColor      = [System.Drawing.Color]::Black
        if ($lblLsp) { $lblLsp.ForeColor = [System.Drawing.Color]::Black }
        $grpPick.ForeColor     = [System.Drawing.Color]::Black
        $grpSign.ForeColor     = [System.Drawing.Color]::Black
        if ($grpSave) { $grpSave.ForeColor = [System.Drawing.Color]::Black }
        $pLog.ForeColor        = [System.Drawing.Color]::Black
        $tlSearch.ForeColor    = [System.Drawing.Color]::Black
    }
}

$miLightTheme.add_Click({ Set-Theme 'light' })
$miDarkTheme.add_Click({ Set-Theme 'dark' })

# Instruktioner

$miShowInstr.add_Click({
    $msg = @"
Snabbguide

1. Skriv in ditt LSP och klicka "Sök Filer eller använd Bläddra..."

2. Klicka på "Skapa rapport"


"@
    [System.Windows.Forms.MessageBox]::Show($msg,"Instruktioner") | Out-Null
})

$miFAQ.add_Click({
    $faq = @"
Vad gör skriptet?

Det skapar en excel-rapport för sökt LSP.

Excelrapport öppnas med följande flikar:
   • Information + SharePoint-Info
   • Seal Test Info
   • STF Sum (och minusvärden -3 mg gräns)
   • Utrustninglista
   • Kontrollmaterial
   • CSV-Summering

"@
    [System.Windows.Forms.MessageBox]::Show($faq,"Vanliga frågor") | Out-Null
})

$miHelpDlg.add_Click({
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.Text = 'Skicka meddelande'
    $helpForm.Size = New-Object System.Drawing.Size(400,300)
    $helpForm.StartPosition = 'CenterParent'
    $helpForm.Font = $form.Font
    $helpBox = New-Object System.Windows.Forms.TextBox
    $helpBox.Multiline = $true
    $helpBox.ScrollBars = 'Vertical'
    $helpBox.Dock = 'Fill'
    $helpBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
    $helpBox.Margin = New-Object System.Windows.Forms.Padding(10)
    $panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelButtons.Dock = 'Bottom'
    $panelButtons.FlowDirection = 'RightToLeft'
    $panelButtons.Padding = New-Object System.Windows.Forms.Padding(10)
    $btnSend = New-Object System.Windows.Forms.Button
    $btnSend.Text = 'Skicka'
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Avbryt'
    $panelButtons.Controls.Add($btnSend)
    $panelButtons.Controls.Add($btnCancel)
    $helpForm.Controls.Add($helpBox)
    $helpForm.Controls.Add($panelButtons)
    $btnSend.Add_Click({
        $msg = $helpBox.Text.Trim()
        if (-not $msg) { [System.Windows.Forms.MessageBox]::Show('Ange ett meddelande innan du skickar.','Hjälp') | Out-Null; return }
        try {
            $helpDir = (Get-ConfigValue -Name 'HelpFeedbackDir' -Default (Join-Path $PSScriptRoot 'help') -ConfigOverride $null)
            if (-not (Test-Path -LiteralPath $helpDir)) { New-Item -ItemType Directory -Path $helpDir -Force | Out-Null }
            $ts   = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $user = ($env:USERNAME + '').Trim()
            if (-not $user) { $user = 'unknown' }
            $file = Join-Path $helpDir ("help_{0}_{1}.txt" -f $user, $ts)

            $body = @(
                "User: $user"
                "Computer: $($env:COMPUTERNAME)"
                "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                ''
                $msg
            ) -join "`r`n"
             Set-Content -Path $file -Value $body -Encoding UTF8
             [System.Windows.Forms.MessageBox]::Show('Meddelandet sparades. Tack!','Hjälp') | Out-Null
             $helpForm.Close()
         } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte spara meddelandet:\n$($_.Exception.Message)",'Hjälp') | Out-Null
        }
    })
    $btnCancel.Add_Click({ $helpForm.Close() })
    $helpForm.ShowDialog() | Out-Null
})

# Om
$miOm.add_Click({ [System.Windows.Forms.MessageBox]::Show("OBS! Detta är endast ett hjälpmedel och ersätter inte någon IPT-process.`n $ScriptVersion`nav Jesper","Om") | Out-Null })

$btnScan.Add_Click({
    if (-not (Assert-StartupReady)) { return }

    $lspInput = ($txtLSP.Text + '').Trim()
    if (-not $lspInput) { Gui-Log "⚠️ Ange ett LSP-nummer" 'Warn'; return }

    # Normalisera: använd endast siffrorna som nyckel (tillåt '#38401', 'LSP 38401' osv.)
    $lsp = ($lspInput -replace '\D', '')
    if (-not $lsp) { Gui-Log "⚠️ Ange ett giltigt LSP-nummer (siffror)" 'Warn'; return }

    if ($script:BuildInProgress) { Gui-Log "⚠️ Rapportgenerering pågår – vänta tills den är klar." 'Warn'; return }
    if ($script:ScanInProgress)  { Gui-Log "⚠️ Sökning pågår redan – vänta." 'Warn'; return }
    $script:ScanInProgress = $true


    Gui-Log -Text ("🔎 Söker filer för {0}…" -f $lsp) -Severity Info -Category USER -Immediate

    try {
        # Återanvänd cache om LSP + mapp fortfarande finns
        if ($script:LastScanResult -and $script:LastScanResult.Lsp -eq $lsp -and
            $script:LastScanResult.Folder -and (Test-Path -LiteralPath $script:LastScanResult.Folder)) {

            Gui-Log -Text ("♻️ Återanvänder senaste sökresultatet för {0}." -f $lsp) -Severity Info -Category USER
            Add-CLBItems -clb $clbCsv -files $script:LastScanResult.Csv -AutoCheckFirst
            Add-CLBItems -clb $clbNeg -files $script:LastScanResult.Neg -AutoCheckFirst
            Add-CLBItems -clb $clbPos -files $script:LastScanResult.Pos -AutoCheckFirst
            Add-CLBItems -clb $clbLsp -files $script:LastScanResult.LspFiles -AutoCheckFirst
            Update-BuildEnabled
            Update-BatchLink
            return
        }
       
$folder = Find-LspFolder -Lsp $lsp -Roots $RootPaths
if (-not $folder) {
    Gui-Log "❌ Ingen LSP-mapp hittad för $lsp" 'Warn'
    if ($env:IPT_ROOT) { Gui-Log "ℹ️ IPT_ROOT=$($env:IPT_ROOT)" 'Info' 'DEBUG' }
    $rootInfo = $RootPaths | ForEach-Object { "{0} ({1})" -f $_, $(if (Test-Path -LiteralPath $_) { "OK" } else { "MISSING" }) }
    Gui-Log -Text ("Sökvägar som provats: " + ($rootInfo -join " | ")) -Severity Info -Category USER
    return
}

if (-not (Test-Path -LiteralPath $folder.FullName)) {
    Gui-Log "❌ LSP-mappen hittades men finns inte längre: $($folder.FullName)" 'Warn'
    $rootInfo = $RootPaths | ForEach-Object { "{0} ({1})" -f $_, $(if (Test-Path -LiteralPath $_) { "OK" } else { "MISSING" }) }
    Gui-Log -Text ("Sökvägar som provats: " + ($rootInfo -join " | ")) -Severity Info -Category USER
    return
}
        Gui-Log -Text ("📂 Hittad mapp: {0}" -f $folder.FullName) -Severity Info -Category USER

        # Plocka filer EN gång
        $files = Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue

        $candCsv = $files | Where-Object { $_.Extension -ieq '.csv' -and ( $_.Name -match [regex]::Escape($lsp) -or $_.Length -gt 100kb ) } | Sort-Object LastWriteTime -Descending
        $candNeg = $files | Where-Object { $_.Name -match '(?i)Neg.*\.xls[xm]$' -and $_.Name -match [regex]::Escape($lsp) } | Sort-Object LastWriteTime -Descending
        $candPos = $files | Where-Object { $_.Name -match '(?i)Pos.*\.xls[xm]$' -and $_.Name -match [regex]::Escape($lsp) } | Sort-Object LastWriteTime -Descending
        $candLsp = $files | Where-Object {
            ($_.Name -match '(?i)worksheet') -and ($_.Name -match [regex]::Escape($lsp)) -and ($_.Extension -match '^(\.xlsx|\.xlsm|\.xls)$')
        } | Sort-Object LastWriteTime -Descending

        Add-CLBItems -clb $clbCsv -files $candCsv -AutoCheckFirst
        Add-CLBItems -clb $clbNeg -files $candNeg -AutoCheckFirst
        Add-CLBItems -clb $clbPos -files $candPos -AutoCheckFirst
        Add-CLBItems -clb $clbLsp -files $candLsp -AutoCheckFirst

        if ($candCsv.Count -eq 0) { Gui-Log "ℹ️ Ingen CSV hittad (endast .csv visas)." 'Info' }
        if ($candNeg.Count -eq 0) { Gui-Log "⚠️ Ingen Seal NEG hittad." 'Warn' }
        if ($candPos.Count -eq 0) { Gui-Log "⚠️ Ingen Seal POS hittad." 'Warn' }
        if ($candLsp.Count -eq 0) { Gui-Log "ℹ️ Ingen LSP Worksheet hittad." 'Info' }

        Update-BuildEnabled
        Update-BatchLink

        # Cachea FileInfo-objekt
        $script:LastScanResult = [pscustomobject]@{
            Lsp      = $lsp
            Folder   = $folder.FullName
            Csv      = @($candCsv)
            Neg      = @($candNeg)
            Pos      = @($candPos)
            LspFiles = @($candLsp)
        }

        Gui-Log -Text "✅ Filer laddade." -Severity Info -Category USER
    }
    catch {
        Gui-Log ("❌ Filsökning misslyckades: " + $_.Exception.Message) 'Error'
    }
    finally {
        $script:ScanInProgress = $false
    }
})

$btnCsvBrowse.Add_Click({
        $dlg = $null
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "CSV|*.csv|Alla filer|*.*"
        if ($dlg.ShowDialog() -eq 'OK') {
            $f = Get-Item -LiteralPath $dlg.FileName
            Add-CLBItems -clb $clbCsv -files @($f) -AutoCheckFirst
            Update-BuildEnabled
            Update-BatchLink
        }
    } finally { if ($dlg) { try { $dlg.Dispose() } catch {} } }
})

$btnNegBrowse.Add_Click({
        $dlg = $null
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
        if ($dlg.ShowDialog() -eq 'OK') {
            $f = Get-Item -LiteralPath $dlg.FileName
            Add-CLBItems -clb $clbNeg -files @($f) -AutoCheckFirst
            Update-BuildEnabled
            Update-BatchLink
        }
    } finally { if ($dlg) { try { $dlg.Dispose() } catch {} } }
})

$btnPosBrowse.Add_Click({
        $dlg = $null
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
        if ($dlg.ShowDialog() -eq 'OK') {
            $f = Get-Item -LiteralPath $dlg.FileName
            Add-CLBItems -clb $clbPos -files @($f) -AutoCheckFirst
            Update-BuildEnabled
            Update-BatchLink
        }
    } finally { if ($dlg) { try { $dlg.Dispose() } catch {} } }
})

# --- Helper: Convert A1 address (e.g. 'E1') to row/col integers (EPPlus safe) ---
if (-not (Get-Command Convert-A1ToRowCol -ErrorAction SilentlyContinue)) {
    function Convert-A1ToRowCol {
        param(
            [string]$A1,
            [int]$DefaultRow = 1,
            [int]$DefaultCol = 5
        )

        if ([string]::IsNullOrWhiteSpace($A1)) {
            return [pscustomobject]@{ Row = $DefaultRow; Col = $DefaultCol }
        }

        $s = ($A1 + '').Trim().ToUpper()
        # Allow optional sheet prefix like 'Information!E1' (we only use the address part)
        if ($s -match '^(?:[^!]+!)?([A-Z]+)(\d+)$') {
            $letters = $matches[1]
            $row = [int]$matches[2]
            $col = 0
            foreach ($ch in $letters.ToCharArray()) {
                $col = ($col * 26) + ([int][char]$ch - [int][char]'A' + 1)
            }
            if ($row -lt 1) { $row = $DefaultRow }
            if ($col -lt 1) { $col = $DefaultCol }
            return [pscustomobject]@{ Row = $row; Col = $col }
        }

        return [pscustomobject]@{ Row = $DefaultRow; Col = $DefaultCol }
    }
}

if (-not (Get-Command Write-SPBlockIntoInformation -ErrorAction SilentlyContinue)) {
    function Write-SPBlockIntoInformation {
        param(
            [Parameter(Mandatory)][OfficeOpenXml.ExcelPackage]$Pkg,
            [Parameter()][object[]]$Rows,
            [Parameter()][string]$Batch,
            [Parameter()][string]$TargetSheetName = 'Information',
            [Parameter()][int]$StartRow = 1,
            [Parameter()][int]$StartCol = 5 # E = 5
        )

        if (-not $Pkg) { return $false }
        $Rows = @($Rows)

        $ws = $Pkg.Workbook.Worksheets[$TargetSheetName]
        if (-not $ws) { return $false }

        $labelCol = $StartCol
        $valueCol = $StartCol + 1

        # Harmoniserade färger (matchar CSV Sammanfattning)
        $HeaderBg   = [System.Drawing.Color]::FromArgb(68, 84, 106)    # Mörkblå
        $HeaderFg   = [System.Drawing.Color]::White
        $SectionBg  = [System.Drawing.Color]::FromArgb(217, 225, 242)  # Ljusblå
        $SectionFg  = [System.Drawing.Color]::FromArgb(0, 32, 96)      # Mörkblå text
        $BorderColor = [System.Drawing.Color]::FromArgb(68, 84, 106)

        # Rensa tidigare block (bara i block-ytan, påverkar inte A-D)
        try {
            $clearRows = 120
            $rngClear = $ws.Cells[$StartRow, $labelCol, ($StartRow + $clearRows - 1), $valueCol]
            $rngClear.Clear()
        } catch {}

        # Huvudrubrik
        $ws.Cells[$StartRow, $labelCol].Value = "SharePoint Info"
        $ws.Cells[$StartRow, $valueCol].Value = ""
        $ws.Cells[$StartRow, $labelCol, $StartRow, $valueCol].Merge = $true
        $ws.Cells[$StartRow, $labelCol].Style.Font.Bold = $true
        $ws.Cells[$StartRow, $labelCol].Style.Font.Size = 14
        $ws.Cells[$StartRow, $labelCol].Style.Font.Name = "Calibri"
        $ws.Cells[$StartRow, $labelCol].Style.Font.Color.SetColor($HeaderFg)
        $ws.Cells[$StartRow, $labelCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $ws.Cells[$StartRow, $labelCol].Style.Fill.BackgroundColor.SetColor($HeaderBg)
        $ws.Cells[$StartRow, $labelCol].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        $ws.Cells[$StartRow, $labelCol].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $ws.Row($StartRow).Height = 22

        $r = $StartRow + 1

        if (-not $Rows -or $Rows.Count -eq 0 -or $Rows[0] -eq $null) {
            # Tom data
            $ws.Cells[$r, $labelCol].Value = "Batch"
            $ws.Cells[$r, $valueCol].Value = $Batch
            $r++
            $ws.Cells[$r, $labelCol].Value = "Status"
            $ws.Cells[$r, $valueCol].Value = "SharePoint avstängt för optimering."
            $lastRow = $r
        } else {
            foreach ($row in $Rows) {
                $ws.Cells[$r, $labelCol].Value = $row.Rubrik
                # Normalisera värdet (tar bort NBSP/tabbar/extra whitespace som kan få kolumnen att se "bred" ut)
                $val = $row.'Värde'
                $valN = Normalize-HeaderText (($val + ''))
                $valN = ($valN -replace '\s+', ' ').Trim()
                $ws.Cells[$r, $valueCol].Value = $valN
                $r++
            }
            $lastRow = $r - 1
        }

        # Styling rubrik-kolumn
        try {
            $labelRange = $ws.Cells[($StartRow + 1), $labelCol, $lastRow, $labelCol]
            $labelRange.Style.Font.Bold = $true
            $labelRange.Style.Font.Name = "Calibri"
            $labelRange.Style.Font.Size = 10
            $labelRange.Style.Font.Color.SetColor($SectionFg)
            $labelRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $labelRange.Style.Fill.BackgroundColor.SetColor($SectionBg)

            # Styling värde-kolumn
            $valueRange = $ws.Cells[($StartRow + 1), $valueCol, $lastRow, $valueCol]
            $valueRange.Style.Font.Name = "Calibri"
            $valueRange.Style.Font.Size = 10
            $valueRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $valueRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::White)
            $valueRange.Style.WrapText = $true
            $valueRange.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
            $valueRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left

            # Borders kring block
            $rng = $ws.Cells[$StartRow, $labelCol, $lastRow, $valueCol]
            $rng.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $rng.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $rng.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $rng.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $rng.Style.Border.Top.Color.SetColor($BorderColor)
            $rng.Style.Border.Bottom.Color.SetColor($BorderColor)
            $rng.Style.Border.Left.Color.SetColor($BorderColor)
            $rng.Style.Border.Right.Color.SetColor($BorderColor)
        } catch {}

        # Kolumnbredd (endast block-kolumnerna)
        try { $ws.Column($labelCol).Width = 30 } catch {}
        try { $ws.Column($valueCol).Width = 35 } catch {}
        try { $ws.Column($valueCol).Style.ShrinkToFit = $false } catch {}

        return $true
    }
}

# ============================
# ===== RAPPORTLOGIK =========
# ============================
        
$btnBuild.Add_Click({
    if (-not (Assert-StartupReady)) { return }

    if ($script:ScanInProgress) { Gui-Log "⚠️ Sökning pågår – vänta innan du skapar rapport." 'Warn'; return }
    if ($script:BuildInProgress) { Gui-Log "⚠️ Rapportgenerering kör redan – vänta." 'Warn'; return }
    $script:BuildInProgress = $true

    Gui-Log -Text '🔃 Skapar rapport…' -Severity Info -Category USER -Immediate
    Set-UiBusy -Busy $true -Message 'Skapar rapport…'
    Set-UiStep 5 'Initierar…'

    $pkgNeg = $null
    $pkgPos = $null
    $pkgOut = $null

    # RuleEngine caches (per-körning)
    $script:RuleEngineShadow  = $null
    $script:RuleEngineCsvObjs = $null
    $script:RuleBankCache     = $null

    try {
        if (-not (Load-EPPlus)) { Gui-Log "❌ EPPlus kunde inte laddas – avbryter." 'Error'; return }

        Set-UiStep 10 '🔃 Läser in Seal Test-filer…'

        $selCsv = Get-CheckedFilePath $clbCsv
        $selNeg = Get-CheckedFilePath $clbNeg
        $selPos = Get-CheckedFilePath $clbPos

        if (-not $selNeg -or -not $selPos) { Gui-Log "❌ Du måste välja en Seal NEG och en Seal POS." 'Error'; return }

        $lspRaw    = ($txtLSP.Text + '').Trim()
        $lspDigits = ($lspRaw -replace '\D','')
        $hasLsp    = -not [string]::IsNullOrWhiteSpace($lspDigits)

        if ($hasLsp) {
            $lsp = $lspDigits
        } else {
            $lsp = 'MANUELL'
            Gui-Log "ℹ️ Ingen LSP angiven – kör manuellt läge (filval via Bläddra). Rapporten märks som '$lsp' och SharePoint/LSP-koppling hoppas över." 'Warn'
        }

        $lspForLinks = if ($hasLsp) { $lsp } else { '' }

        # ---- Local staging for network files (N:\ or UNC) ----
        $stageDir = $null
        $enableStaging = Get-ConfigFlag -Name 'EnableLocalStaging' -Default $true -ConfigOverride $Config
        if ($enableStaging) {
            try {
                $stageDir = Join-Path ([IO.Path]::GetTempPath()) ("QC_Stage_" + $lsp + "_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
                New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

                $selNeg = Stage-InputFile -Path $selNeg -StageDir $stageDir
                $selPos = Stage-InputFile -Path $selPos -StageDir $stageDir
                if ($selCsv) { $selCsv = Stage-InputFile -Path $selCsv -StageDir $stageDir }
            } catch {
            }
        }

        Gui-Log "📄 Neg-fil: $(Split-Path $selNeg -Leaf)" 'Info'
        Gui-Log "📄 Pos-fil: $(Split-Path $selPos -Leaf)" 'Info'
        if ($selCsv) { Gui-Log "📄 CSV: $(Split-Path $selCsv -Leaf)" 'Info' } else { Gui-Log "ℹ️ Ingen CSV vald." 'Info' }

        $negWritable = $true; $posWritable = $true
        if ($chkWriteSign.Checked) {
            $negWritable = -not (Test-FileLocked $selNeg); if (-not $negWritable) { Gui-Log "🔒 NEG är låst (öppen i Excel?)." 'Warn' }
            $posWritable = -not (Test-FileLocked $selPos); if (-not $posWritable) { Gui-Log "🔒 POS är låst (öppen i Excel?)." 'Warn' }
        }

        # ----------------------------
        # Open packages
        # ----------------------------
        try {
            $pkgNeg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selNeg))
            $pkgPos = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selPos))
        } catch {
            Gui-Log "❌ Kunde inte öppna NEG/POS: $($_.Exception.Message)" 'Error'
            return
        }

        $templatePath = Join-Path $PSScriptRoot "output_template-v4.xlsx"
        if (-not (Test-Path -LiteralPath $templatePath)) { Gui-Log "❌ Mallfilen 'output_template-v4.xlsx' saknas!" 'Error'; return }
        try {
            $pkgOut = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($templatePath))
        } catch {
            Gui-Log "❌ Kunde inte läsa mall: $($_.Exception.Message)" 'Error'
            return
        }

        # ============================
        # === SIGNATUR I NEG/POS  ====
        # ============================

        $signToWrite = ($txtSigner.Text + '').Trim()
        if ($chkWriteSign.Checked) {
            if (-not $signToWrite) { Gui-Log "❌ Ingen signatur angiven. Avbryter."; return }
            if (-not (Confirm-SignatureInput -Text $signToWrite)) { Gui-Log "🛑 Signatur ej bekräftad. Avbryter."; return }

            $negWritten = 0; $posWritten = 0; $negSkipped = 0; $posSkipped = 0

            foreach ($ws in $pkgNeg.Workbook.Worksheets) {
                if ($ws.Name -eq 'Worksheet Instructions') { continue }
                $h3 = ($ws.Cells['H3'].Text + '').Trim()

                if ($h3 -match '^[0-9]') {
                    $existing = ($ws.Cells[$Layout.SignatureCell].Text + '').Trim()
                    if ($existing -and -not $chkOverwriteSign.Checked) { $negSkipped++; continue }
                    $ws.Cells[$Layout.SignatureCell].Style.Numberformat.Format = '@'
                    $ws.Cells[$Layout.SignatureCell].Value = $signToWrite
                    $negWritten++
                }
                elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
                    break
                }
            }

            foreach ($ws in $pkgPos.Workbook.Worksheets) {
                if ($ws.Name -eq 'Worksheet Instructions') { continue }
                $h3 = ($ws.Cells['H3'].Text + '').Trim()

                if ($h3 -match '^[0-9]') {
                    $existing = ($ws.Cells[$Layout.SignatureCell].Text + '').Trim()
                    if ($existing -and -not $chkOverwriteSign.Checked) { $posSkipped++; continue }
                    $ws.Cells[$Layout.SignatureCell].Style.Numberformat.Format = '@'
                    $ws.Cells[$Layout.SignatureCell].Value = $signToWrite
                    $posWritten++
                }
                elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
                    break
                }
            }

            try {
                if ($negWritten -eq 0 -and $negSkipped -eq 0 -and $posWritten -eq 0 -and $posSkipped -eq 0) {
                    Gui-Log "ℹ️ Inga databladsflikar efter flik 1 att sätta signatur i (ingen åtgärd)."
                } else {
                    if ($negWritten -gt 0 -and $negWritable) { $pkgNeg.Save() } elseif ($negWritten -gt 0) { Gui-Log "🔒 Kunde inte spara NEG (låst)." 'Warn' }
                    if ($posWritten -gt 0 -and $posWritable) { $pkgPos.Save() } elseif ($posWritten -gt 0) { Gui-Log "🔒 Kunde inte spara POS (låst)." 'Warn' }
                    Gui-Log "🖊️ Signatur satt: NEG $negWritten blad (överhoppade $negSkipped), POS $posWritten blad (överhoppade $posSkipped)."
                }
            } catch {
                Gui-Log "⚠️ Kunde inte spara signatur i NEG/POS: $($_.Exception.Message)" 'Warn'
            }
        }

        # ============================
        # === CSV (Info/Control)  ====
        # ============================

        $csvRows = @()
        $runAssay = $null

        if ($selCsv) {
            try {
                $csvInfo = Get-Item -LiteralPath $selCsv -ErrorAction Stop
                $thresholdMb = 25
                try {
                    if ($Config -and ($Config -is [System.Collections.IDictionary]) -and $Config.Contains('CsvStreamingThresholdMB')) {
                        $thresholdMb = [int]$Config['CsvStreamingThresholdMB']
                    }
                } catch {}

                $useStreaming = ($csvInfo.Length -ge ($thresholdMb * 1MB))
                if ($useStreaming) {
                    Gui-Log ("⏳ CSV är stor ({0:N1} MB) – använder streaming-import…" -f ($csvInfo.Length / 1MB)) 'Info'
                    Set-UiStep 35 "Läser CSV (streaming)…"
                    $list = New-Object System.Collections.Generic.List[object]
                    Import-CsvRowsStreaming -Path $selCsv -StartRow 10 -ProcessRow {
                        param($Fields,$RowIndex)
                        [void]$list.Add($Fields)
                        if (($RowIndex % 2000) -eq 0) { Invoke-UiPump }
                    }
                    $csvRows = @($list.ToArray())
                } else {
                    Set-UiStep 35 "Läser CSV…"
                    $csvRows = Import-CsvRows -Path $selCsv -StartRow 10
                }

                # --- Robusthet: Sortera på kolumn C (Sample ID) innan vidare bearbetning.
                # Många downstream-steg (gruppering/skrivning) förutsätter att raderna är konsekvent sorterade.
                try {
                    if ($csvRows -and $csvRows.Count -gt 1) {
                        if ($csvRows[0] -is [object[]]) {
                            # Import-CsvRows* returnerar fält-arrayer: kolumn C = index 2
                            $csvRows = @($csvRows | Sort-Object { [string]($_[2]) })
                        } else {
                            # Om vi i framtiden får PSCustomObject-rader
                            $csvRows = @($csvRows | Sort-Object { [string]($_.'Sample ID') })
                        }
                        Gui-Log ("🔃 CSV sorterad på kolumn C (Sample ID). Rader: {0}" -f $csvRows.Count) 'Info'
                    }
                } catch {
                    Gui-Log "⚠️ Kunde inte sortera CSV på kolumn C (Sample ID)." 'Warn'
                }
            } catch {
                Gui-Log "⚠️ CSV-import misslyckades: $($_.Exception.Message)" 'Warn'
                $csvRows = @()
            }
            try { $runAssay = Get-AssayFromCsv -Path $selCsv -StartRow 10 } catch {}
            if ($runAssay) { Gui-Log "🔎 Assay från CSV: $runAssay" }
        }

        # ============================
        # === RuleEngine (shadow)  ===
        # ============================
        try {
            if ((Get-ConfigFlag -Name 'EnableRuleEngine' -Default $false -ConfigOverride $Config) -and
                $selCsv -and (Test-Path -LiteralPath $selCsv)) {

                # Load rulebank once
                $rb = Load-RuleBank -RuleBankDir $Config.RuleBankDir
                try { $rb = Compile-RuleBank -RuleBank $rb } catch {}
                $script:RuleBankCache = $rb

                # Build csvObjs once (prefer Import-CsvRows, else raw fallback)
                $csvObjs = @()
                if ($csvRows -and $csvRows.Count -gt 0) {
                    $csvObjs = @($csvRows)
                    Gui-Log ("🧠 Regelmotor: CSV-källa: Import-CsvRows ({0})" -f $csvObjs.Count) 'Info'
                } else {
                    try {
                        $all = Get-Content -LiteralPath $selCsv
                        if ($all -and $all.Count -gt 9) {
                            $del = Get-CsvDelimiter -Path $selCsv
                            $hdr = ConvertTo-CsvFields $all[7]
                            $dl  = $all[9..($all.Count-1)] | Where-Object { $_ -and $_.Trim() }
                            $csvObjs = @(ConvertFrom-Csv -InputObject ($dl -join "`n") -Delimiter $del -Header $hdr)
                        }
                    } catch {
                        $csvObjs = @()
                    }
                    Gui-Log ("🧠 Regelmotor: CSV-källa: Fallback-raw ({0})" -f ($csvObjs.Count)) 'Info'
                }

                $script:RuleEngineCsvObjs = $csvObjs

                if (-not $csvObjs -or $csvObjs.Count -eq 0) {
                    Gui-Log "⚠️ Regelmotor: CSV-objekt saknas (0 rader) – hoppar över." 'Warn'
                    $script:RuleEngineShadow = $null
                } else {
                    $re = Invoke-RuleEngine -CsvObjects $csvObjs -RuleBank $rb -CsvPath $selCsv
                    $script:RuleEngineShadow = $re

                    if ($re -and $re.Summary) {
                        $pairs = @()
                        foreach ($k in $re.Summary.ObservedCounts.Keys) { $pairs += ("$k=$($re.Summary.ObservedCounts[$k])") }
                        if ($pairs.Count -gt 0) { Gui-Log -Text ("🧠 Regelmotor: ObservedCall counts: " + ($pairs -join ', ')) -Severity Info -Category RuleEngineStats }

                        $dpairs = @()
                        foreach ($k2 in $re.Summary.DeviationCounts.Keys) { $dpairs += ("$k2=$($re.Summary.DeviationCounts[$k2])") }
                        if ($dpairs.Count -gt 0) { Gui-Log -Text ("🧠 Regelmotor: Deviations: " + ($dpairs -join ', ')) -Severity Info -Category RuleEngineStats }

                        if ($re.Summary.RetestYes -gt 0) { Gui-Log -Text ("🧠 Regelmotor: Retest=YES count: " + $re.Summary.RetestYes) -Severity Info -Category RuleEngineStats }

                        # Single user-facing summary line
                        if (Get-ConfigFlag -Name 'EnableRuleEngineSummaryLog' -Default $false -ConfigOverride $Config) {
                            try {
                                $pos = 0; $neg = 0; $err = 0
                                if ($re.Summary.ObservedCounts) {
                                    if ($re.Summary.ObservedCounts.ContainsKey('POS'))   { $pos = [int]$re.Summary.ObservedCounts['POS'] }
                                    if ($re.Summary.ObservedCounts.ContainsKey('NEG'))   { $neg = [int]$re.Summary.ObservedCounts['NEG'] }
                                    if ($re.Summary.ObservedCounts.ContainsKey('ERROR')) { $err = [int]$re.Summary.ObservedCounts['ERROR'] }
                                }

                                # Deviations: OK / FP / ERROR
                                $ok = 0; $fp = 0; $fn = 0; $derr = 0
                                if ($re.Summary.DeviationCounts) {
                                    if ($re.Summary.DeviationCounts.ContainsKey('OK'))    { $ok   = [int]$re.Summary.DeviationCounts['OK'] }
                                    if ($re.Summary.DeviationCounts.ContainsKey('FP'))    { $fp   = [int]$re.Summary.DeviationCounts['FP'] }
                                    if ($re.Summary.DeviationCounts.ContainsKey('FN'))    { $fn   = [int]$re.Summary.DeviationCounts['FN'] }
                                    if ($re.Summary.DeviationCounts.ContainsKey('ERROR')) { $derr = [int]$re.Summary.DeviationCounts['ERROR'] }
                                }

                                $rt = [int]$re.Summary.RetestYes
                                $sum = "🧠 Regelkontroll: POS=$pos, NEG=$neg, Error=$err | OK=$ok, FP=$fp, FN=$fn, Minor Func=$derr | Instrument Error=$rt"
                                Gui-Log -Text $sum -Severity Info -Category SUMMARY
                            } catch { }
                        }

                        if ((Get-ConfigFlag -Name 'EnableShadowCompare' -Default $false -ConfigOverride $Config) -and $re.TopDeviations) {
                            $n = 0
                            foreach ($d in $re.TopDeviations) {
                                $n++; if ($n -gt 20) { break }
                                $msg = "⚠️ RuleEngine dev: " + ($d.Deviation + '') + " | " + ($d.SampleId + '') + " | Exp=" + ($d.ExpectedCall + '') + " | Obs=" + ($d.ObservedCall + '')
                                if (($d.ErrorCode + '').Trim()) { $msg += " | Err=" + ($d.ErrorCode + '') }
                                Gui-Log -Text $msg -Severity Info -Category RuleEngineDev
                            }
                        }
                    }
                }
            }
        } catch {
            Gui-Log ("⚠️ RuleEngine (shadow) fel: " + $_.Exception.Message) 'Warn'
        }

        $controlTab = $null
        if ($runAssay) { $controlTab = Get-ControlTabName -AssayName $runAssay }
        if ($controlTab) { Gui-Log "🧪 Kontrollmaterial-flik: $controlTab" } else { Gui-Log "ℹ️ Ingen Kontrollmaterialsflik (fortsätter utan)." }

        # ============================
        # === Läs avvikelser       ===
        # ============================

        $violationsNeg = @(); $violationsPos = @(); $failNegCount = 0; $failPosCount = 0

        foreach ($ws in $pkgNeg.Workbook.Worksheets) {
            if ($ws.Name -eq "Worksheet Instructions") { continue }
            if (-not $ws.Dimension) { continue }
            $obsC = Find-ObservationCol $ws

            for ($r = 3; $r -le 45; $r++) {
                $valK = $ws.Cells["K$r"].Value
                $textL = $ws.Cells["L$r"].Text

                if ($valK -ne $null -and $valK -is [double]) {
                    if ($textL -eq "FAIL" -or $valK -le -3.0) {
                        $obsTxt = $ws.Cells[$r, $obsC].Text
                        $violationsNeg += [PSCustomObject]@{
                            Sheet      = $ws.Name
                            Cartridge  = $ws.Cells["H$r"].Text
                            InitialW   = $ws.Cells["I$r"].Value
                            FinalW     = $ws.Cells["J$r"].Value
                            WeightLoss = $valK
                            Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                            Obs        = $obsTxt
                        }
                        if ($textL -eq "FAIL") { $failNegCount++ }
                    }
                }
            }
        }

        foreach ($ws in $pkgPos.Workbook.Worksheets) {
            if ($ws.Name -eq "Worksheet Instructions") { continue }
            if (-not $ws.Dimension) { continue }
            $obsC = Find-ObservationCol $ws

            for ($r = 3; $r -le 45; $r++) {
                $valK = $ws.Cells["K$r"].Value
                $textL = $ws.Cells["L$r"].Text

                if ($valK -ne $null -and $valK -is [double]) {
                    if ($textL -eq "FAIL" -or $valK -le -3.0) {
                        $obsTxt = $ws.Cells[$r, $obsC].Text
                        $violationsPos += [PSCustomObject]@{
                            Sheet      = $ws.Name
                            Cartridge  = $ws.Cells["H$r"].Text
                            InitialW   = $ws.Cells["I$r"].Value
                            FinalW     = $ws.Cells["J$r"].Value
                            WeightLoss = $valK
                            Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                            Obs        = $obsTxt
                        }
                        if ($textL -eq "FAIL") { $failPosCount++ }
                    }
                }
            }
        }

        # ============================
        # === Seal Test Info (blad) ==
        # ============================

        $wsOut1 = $pkgOut.Workbook.Worksheets["Seal Test Info"]
        if (-not $wsOut1) { Gui-Log "❌ Fliken 'Seal Test Info' saknas i mallen"; return }

        for ($row = 3; $row -le 15; $row++) {
            $wsOut1.Cells["D$row"].Value = $null
            try { $wsOut1.Cells["D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None } catch {}
        }

        $fields = @(
            @{ Label = "ROBAL";                         Cell = "F2"  }
            @{ Label = "Part Number";                   Cell = "B2"  }
            @{ Label = "Batch Number";                  Cell = "D2"  }
            @{ Label = "Cartridge Number (LSP)";        Cell = "B6"  }
            @{ Label = "PO Number";                     Cell = "B10" }
            @{ Label = "Assay Family";                  Cell = "D10" }
            @{ Label = "Weight Loss Spec";              Cell = "F10" }
            @{ Label = "Balance ID Number";             Cell = "B14" }
            @{ Label = "Balance Cal Due Date";          Cell = "D14" }
            @{ Label = "Vacuum Oven ID Number";         Cell = "B20" }
            @{ Label = "Vacuum Oven Cal Due Date";      Cell = "D20" }
            @{ Label = "Timer ID Number";               Cell = "B25" }
            @{ Label = "Timer Cal Due Date";            Cell = "D25" }
        )

        $forceText = @("ROBAL","Part Number","Batch Number","Cartridge Number (LSP)","PO Number","Assay Family","Balance ID Number","Vacuum Oven ID Number","Timer ID Number")
        $mismatchFields = $fields[0..6] | ForEach-Object { $_.Label }


        # --- Equipment expected list (optional, från equipment.xml) ---
        $equip = $null
        $equipPath = $null
        try {
            $equipPath = Get-ConfigValue -Name 'EquipmentXmlPath' -Default (Join-Path $PSScriptRoot 'equipment.xml') -ConfigOverride $Config
            if ($equipPath -and -not (Test-Path -LiteralPath $equipPath)) {
                Gui-Log ("⚠️ Equipment.xml hittades inte: " + $equipPath) 'Warn'
            }
            if ($equipPath -and (Test-Path -LiteralPath $equipPath)) {
                $equip = Import-Clixml -LiteralPath $equipPath
            }
        } catch {
            $equip = $null
        }

        function _EquipTokens([string]$s) {
            if ([string]::IsNullOrWhiteSpace($s)) { return @() }
            $parts = $s -split '[,;]' | ForEach-Object { ($_ -replace '\s+', ' ').Trim() } | Where-Object { $_ }
            $tok = @()
            foreach ($p in $parts) {
                $v = (Normalize-HeaderText $p).Trim().ToUpper()
                $v = $v -replace '^(NR\.?|NO\.?)\s*', ''
                $v = $v -replace '[^A-Z0-9_-]', ''
                if ($v) { $tok += $v }
            }
            return ($tok | Sort-Object -Unique)
        }

        function _EquipPretty([string[]]$tokens, [string]$label) {
    if (-not $tokens -or $tokens.Count -eq 0) { return '' }
    $t = @($tokens | Sort-Object -Unique)

    if ($label -ieq 'Timer ID Number') {
        # Visa som "Nr. 18, Nr. 19"
        return (($t | ForEach-Object { "Nr. $_" }) -join ', ')
    }

    return ($t -join ', ')
}

        function _FixMonthText([string]$s) {
            if ([string]::IsNullOrWhiteSpace($s)) { return '' }
            $t = (Normalize-HeaderText $s).Trim()
            # Byt ut kyrillisk A (А/а) mot latin A – vanligt i "Аpr-26"
            $t = $t -replace '[Аа]', 'A'
            $t = ($t -replace '\s+', ' ').Trim()
            return $t
        }

        function _EquipEvalList([string]$actual, [string]$expected) {
            # Returnerar: Match | Delvis | Saknas | Mismatch | NoRef
            if ([string]::IsNullOrWhiteSpace($expected)) { return 'NoRef' }
            $aT = _EquipTokens $actual
            $eT = _EquipTokens $expected
            if ($aT.Count -eq 0) { return 'Saknas' }
            foreach ($a in $aT) { if (-not ($eT -contains $a)) { return 'Mismatch' } }
            foreach ($e in $eT) { if (-not ($aT -contains $e)) { return 'Delvis' } }
            return 'Match'
        }

        function _EquipEvalMonth([string]$actual, [string]$expected) {
            if ([string]::IsNullOrWhiteSpace($expected)) { return 'NoRef' }
            if ([string]::IsNullOrWhiteSpace($actual))   { return 'Saknas' }
            $a = (_FixMonthText $actual).ToUpper()
            $e = (_FixMonthText $expected).ToUpper()
            if ($a -eq $e) { return 'Match' }
            return 'Mismatch'
        }

function _IsActiveSealSheet($ws) {
    try {
        if (-not $ws) { return $false }
        if ($ws.Name -eq "Worksheet Instructions") { return $false }

        $h3 = (($ws.Cells["H3"].Text + '')).Trim()
        if (-not $h3) { return $false }

        $h3n = (Normalize-HeaderText $h3).Trim().ToUpper()
        if ($h3n -eq 'NA' -or $h3n -eq 'N/A') { return $false }

        return $true
    } catch {
        return $false
    }
}

function _GetPerSheetValues($pkg, [string]$addr) {
    # Returnerar en lista med värden (en per "aktiv" flik)
    $vals = @()

    foreach ($ws in $pkg.Workbook.Worksheets) {
        if (-not (_IsActiveSealSheet $ws)) { continue }

        try {
            $cell = $ws.Cells[$addr]
            $v = ''
            if ($cell -and $cell.Value -ne $null) {
                if ($cell.Value -is [datetime]) { $v = $cell.Value.ToString('MMM-yy') } else { $v = $cell.Text }
            }
            $vals += ($v + '')
        } catch {
            $vals += ''
        }
    }

    return $vals
}


function _AggregateStatus([string[]]$statuses) {
    # Summerar: Mismatch > Saknas > Delvis > Match
    if (-not $statuses -or $statuses.Count -eq 0) { return 'NoRef' }
    if ($statuses -contains 'Mismatch') { return 'Mismatch' }
    if ($statuses -contains 'Saknas')   { return 'Saknas' }
    if ($statuses -contains 'Delvis')   { return 'Delvis' }
    if ($statuses -contains 'Match')    { return 'Match' }
    return 'NoRef'
}

        $equipMap = @{
            'Balance ID Number'        = @{ NegKey='SCALESNEG';       PosKey='SCALESPOS';       Mode='LIST'  }
            'Balance Cal Due Date'     = @{ NegKey='CAL_D_SCALES';    PosKey='CAL_D_SCALES';    Mode='MONTH' }
            'Vacuum Oven ID Number'    = @{ NegKey='OVENSNEG';        PosKey='OVENSPOS';        Mode='LIST'  }
            'Vacuum Oven Cal Due Date' = @{ NegKey='CAL_D_OVENS';     PosKey='CAL_D_OVENS';     Mode='MONTH' }
            'Timer ID Number'          = @{ NegKey='TIMERSNEG';       PosKey='TIMERSPOS';       Mode='LIST'  }
            'Timer Cal Due Date'       = @{ NegKey='CAL_D_TIMERSNEG'; PosKey='CAL_D_TIMERSPOS'; Mode='MONTH' }
        }

        $row = 3
        foreach ($f in $fields) {
$valNeg=''; $valPos=''

# För equipment-rader: samla per flik (inte break på första!)
$perNeg = $null
$perPos = $null
$isEquip = ($equip -and $equipMap -and $equipMap.ContainsKey($f.Label))

if ($isEquip) {
    $perNeg = _GetPerSheetValues $pkgNeg $f.Cell
    $perPos = _GetPerSheetValues $pkgPos $f.Cell

    # Visa i B/C: union (för LIST) eller första icke-tom (för MONTH), bara för läsbarhet
    $map = $equipMap[$f.Label]
    if ($map.Mode -eq 'MONTH') {
        $valNeg = ($perNeg | Where-Object { $_ -and $_.Trim() } | Select-Object -First 1)
        $valPos = ($perPos | Where-Object { $_ -and $_.Trim() } | Select-Object -First 1)
    } else {
        $tokN = @()
        foreach ($x in $perNeg) { $tokN += (_EquipTokens $x) }
        $tokP = @()
        foreach ($x in $perPos) { $tokP += (_EquipTokens $x) }
$valNeg = _EquipPretty ($tokN | Sort-Object -Unique) $f.Label
$valPos = _EquipPretty ($tokP | Sort-Object -Unique) $f.Label
    }
}
else {
    # Befintligt beteende för "vanliga" fält: första flik med värde
    foreach ($wsN in $pkgNeg.Workbook.Worksheets) {
        if ($wsN.Name -eq "Worksheet Instructions") { continue }
        $cell = $wsN.Cells[$f.Cell]
        if ($cell.Value -ne $null) {
            if ($cell.Value -is [datetime]) { $valNeg = $cell.Value.ToString('MMM-yy') } else { $valNeg = $cell.Text }
            break
        }
    }

    foreach ($wsP in $pkgPos.Workbook.Worksheets) {
        if ($wsP.Name -eq "Worksheet Instructions") { continue }
        $cell = $wsP.Cells[$f.Cell]
        if ($cell.Value -ne $null) {
            if ($cell.Value -is [datetime]) { $valPos = $cell.Value.ToString('MMM-yy') } else { $valPos = $cell.Text }
            break
        }
    }
}


            if ($forceText -contains $f.Label) {
                $wsOut1.Cells["B$row"].Style.Numberformat.Format = '@'
                $wsOut1.Cells["C$row"].Style.Numberformat.Format = '@'
            }

            $wsOut1.Cells["B$row"].Value = $valNeg
            $wsOut1.Cells["C$row"].Value = $valPos
            $wsOut1.Cells["B$row"].Style.Border.Right.Style = "Medium"
            $wsOut1.Cells["C$row"].Style.Border.Left.Style  = "Medium"

        if ($mismatchFields -contains $f.Label) {
                # D3:D9: visa tydlig Match/Mismatch med symboler
                if ($valNeg -and $valPos) {
                    if ($valNeg -ne $valPos) {
                        $wsOut1.Cells["D$row"].Value = "⚠ Mismatch"
                        Style-Cell $wsOut1.Cells["D$row"] $true "FF0000" "Medium" "FFFFFF"
                        Gui-Log "⚠️ Avvikelse: $($f.Label) (NEG='$valNeg' vs POS='$valPos')"
                    } else {
                        $wsOut1.Cells["D$row"].Value = "✓ Match"
                        Style-Cell $wsOut1.Cells["D$row"] $true "C6EFCE" "Medium" "006100"
                    }
                } elseif ($valNeg -or $valPos) {
                    # Bara en av filerna har värde - markera som varning
                    $wsOut1.Cells["D$row"].Value = "⚠ Saknas"
                    Style-Cell $wsOut1.Cells["D$row"] $true "FFE699" "Medium" "806000"
                }
            }


            # --- Equipment check (POS/NEG har egna tillåtna värden i equipment.xml) ---
            if ($equip -and $equipMap -and $equipMap.ContainsKey($f.Label)) {
                $map = $equipMap[$f.Label]
                $expNeg = $null; $expPos = $null
                try { if ($equip.Contains($map.NegKey)) { $expNeg = (""+$equip[$map.NegKey]).Trim() } } catch {}
                try { if ($equip.Contains($map.PosKey)) { $expPos = (""+$equip[$map.PosKey]).Trim() } } catch {}

# Utvärdera PER FLIK och summera
$stNegAll = @()
$stPosAll = @()

if ($map.Mode -eq 'MONTH') {
    foreach ($v in $perNeg) { $stNegAll += (_EquipEvalMonth $v $expNeg) }
    foreach ($v in $perPos) { $stPosAll += (_EquipEvalMonth $v $expPos) }
} else {
    foreach ($v in $perNeg) { $stNegAll += (_EquipEvalList $v $expNeg) }
    foreach ($v in $perPos) { $stPosAll += (_EquipEvalList $v $expPos) }
}

$sNeg = _AggregateStatus $stNegAll
$sPos = _AggregateStatus $stPosAll

                function _Txt($s) {
                    switch ($s) {
                        'Match'   { '✓ Match' }
                        'Delvis'  { '✓ Delvis' }
                        'Saknas'  { '⚠ Saknas' }
                        'Mismatch'{ '⚠ Mismatch' }
                        default   { '?' }
                    }
                }

                $wsOut1.Cells["D$row"].Value = ('NEG ' + (_Txt $sNeg) + ' / POS ' + (_Txt $sPos))

                # Style D utifrån värsta status: Mismatch > Saknas/Delvis > Match
                $rank = @{ 'NoRef'=0; 'Match'=1; 'Delvis'=2; 'Saknas'=2; 'Mismatch'=3 }
                $sev = [Math]::Max($rank[$sNeg], $rank[$sPos])

                if ($sev -ge 3) {
                    Style-Cell $wsOut1.Cells["D$row"] $true "FF0000" "Medium" "FFFFFF"
                } elseif ($sev -ge 2) {
                    Style-Cell $wsOut1.Cells["D$row"] $true "FFE699" "Medium" "806000"
                } elseif ($sev -ge 1) {
                    Style-Cell $wsOut1.Cells["D$row"] $true "C6EFCE" "Medium" "006100"
                }
            }
            $row++
        }

        $wsOut1.Cells["D:D"].Style.WrapText = $false
        $wsOut1.Column(4).AutoFit()
        $wsOut1.Column(4).Width += 1.5
        $wsOut1.Column(4).BestFit = $true

        # ============================
        # === Testare (B43)        ===
        # ============================

        $testersNeg = @(); $testersPos = @()
        foreach ($s in $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) {
            $t=$s.Cells["B43"].Text
            if ($t) { $testersNeg += ($t -split ",") }
        }
        foreach ($s in $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) {
            $t=$s.Cells["B43"].Text
            if ($t) { $testersPos += ($t -split ",") }
        }
        $testersNeg = $testersNeg | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
        $testersPos = $testersPos | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

        $wsOut1.Cells["B16"].Value = "Name of Tester"
        $wsOut1.Cells["B16:C16"].Merge = $true
        $wsOut1.Cells["B16"].Style.HorizontalAlignment = "Center"

        $maxTesters = [Math]::Max($testersNeg.Count, $testersPos.Count)
        $initialRows = 11
        if ($maxTesters -lt $initialRows) { $wsOut1.DeleteRow(17 + $maxTesters, $initialRows - $maxTesters) }
        if ($maxTesters -gt $initialRows) {
            $rowsToAdd = $maxTesters - $initialRows
            $lastRow = 16 + $initialRows
            for ($i = 1; $i -le $rowsToAdd; $i++) { $wsOut1.InsertRow($lastRow + 1, 1, $lastRow) }
        }

        for ($i = 0; $i -lt $maxTesters; $i++) {
            $rowIndex = 17 + $i
            $wsOut1.Cells["A$rowIndex"].Value = $null
            $wsOut1.Cells["B$rowIndex"].Value = if ($i -lt $testersNeg.Count) { $testersNeg[$i] } else { "N/A" }
            $wsOut1.Cells["C$rowIndex"].Value = if ($i -lt $testersPos.Count) { $testersPos[$i] } else { "N/A" }

            $topStyle    = if ($i -eq 0) { "Medium" } else { "Thin" }
            $bottomStyle = if ($i -eq $maxTesters - 1) { "Medium" } else { "Thin" }

            foreach ($col in @("B","C")) {
                $cell = $wsOut1.Cells["$col$rowIndex"]
                $cell.Style.Border.Top.Style    = $topStyle
                $cell.Style.Border.Bottom.Style = $bottomStyle
                $cell.Style.Border.Left.Style   = "Medium"
                $cell.Style.Border.Right.Style  = "Medium"
                $cell.Style.Fill.PatternType = "Solid"
                $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
            }
        }

        # ============================
        # === Signatur-jämförelse  ===
        # ============================

        $negSigSet = Get-SignatureSetForDataSheets -Pkg $pkgNeg
        $posSigSet = Get-SignatureSetForDataSheets -Pkg $pkgPos

        $negSet = New-Object 'System.Collections.Generic.HashSet[string]'
        $posSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($n in $negSigSet.NormSet) { [void]$negSet.Add($n) }
        foreach ($p in $posSigSet.NormSet) { [void]$posSet.Add($p) }

        $hasNeg = ($negSet.Count -gt 0)
        $hasPos = ($posSet.Count -gt 0)

        $onlyNeg = @(); $onlyPos = @(); $sigMismatch = $false
        if ($hasNeg -and $hasPos) {
            foreach ($n in $negSet) { if (-not $posSet.Contains($n)) { $onlyNeg += $n } }
            foreach ($p in $posSet) { if (-not $negSet.Contains($p)) { $onlyPos += $p } }
            $sigMismatch = ($onlyNeg.Count -gt 0 -or $onlyPos.Count -gt 0)
        } else {
            $sigMismatch = $false
        }

        $mismatchSheets = @()
        if ($sigMismatch) {
            foreach ($k in $onlyNeg) {
                $raw = if ($negSigSet.RawByNorm.ContainsKey($k)) { $negSigSet.RawByNorm[$k] } else { $k }
                $where = if ($negSigSet.Occ.ContainsKey($k)) { ($negSigSet.Occ[$k] -join ', ') } else { '—' }
                $mismatchSheets += ("NEG: " + $raw + "  [Blad: " + $where + "]")
            }
            foreach ($k in $onlyPos) {
                $raw = if ($posSigSet.RawByNorm.ContainsKey($k)) { $posSigSet.RawByNorm[$k] } else { $k }
                $where = if ($posSigSet.Occ.ContainsKey($k)) { ($posSigSet.Occ[$k] -join ', ') } else { '—' }
                $mismatchSheets += ("POS: " + $raw + "  [Blad: " + $where + "]")
            }
            Gui-Log "⚠️ Avvikelse: Print Full Name, Sign, and Date (NEG vs POS)"
        }

        if (-not (Get-Command Set-MergedWrapAutoHeight -ErrorAction SilentlyContinue)) {
            function Set-MergedWrapAutoHeight {
                param([OfficeOpenXml.ExcelWorksheet]$Sheet,[int]$RowIndex,[int]$ColStart=2,[int]$ColEnd=3,[string]$Text)
                $rng = $Sheet.Cells[$RowIndex, $ColStart, $RowIndex, $ColEnd]
                $rng.Style.WrapText = $true
                $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None
                $Sheet.Row($RowIndex).CustomHeight = $false
                try {
                    $wChars = [Math]::Floor(($Sheet.Column($ColStart).Width + $Sheet.Column($ColEnd).Width) - 2); if ($wChars -lt 1) { $wChars = 1 }
                    $segments = $Text -split "(\r\n|\n|\r)"; $lineCount = 0
                    foreach ($seg in $segments) { if (-not $seg) { $lineCount++ } else { $lineCount += [Math]::Ceiling($seg.Length / $wChars) } }
                    if ($lineCount -lt 1) { $lineCount = 1 }
                    $targetHeight = [Math]::Max(15, [Math]::Ceiling(15 * $lineCount * 2.15))
                    if ($Sheet.Row($RowIndex).Height -lt $targetHeight) {
                        $Sheet.Row($RowIndex).Height = $targetHeight
                        $Sheet.Row($RowIndex).CustomHeight = $true
                    }
                } catch { $Sheet.Row($RowIndex).CustomHeight = $false }
            }
        }

        $signRow = 17 + $maxTesters + 3
        $displaySignNeg = $null; $displaySignPos = $null

        if ($signToWrite) {
            $displaySignNeg = $signToWrite
            $displaySignPos = $signToWrite
        } else {
            $displaySignNeg = if ($negSigSet.RawFirst) { $negSigSet.RawFirst } else { '—' }
            $displaySignPos = if ($posSigSet.RawFirst) { $posSigSet.RawFirst } else { '—' }
        }

        $wsOut1.Cells["B$signRow"].Style.Numberformat.Format = '@'
        $wsOut1.Cells["C$signRow"].Style.Numberformat.Format = '@'
        $wsOut1.Cells["B$signRow"].Value = $displaySignNeg
        $wsOut1.Cells["C$signRow"].Value = $displaySignPos

        foreach ($col in @('B','C')) {
            $cell = $wsOut1.Cells["${col}$signRow"]
            Style-Cell $cell $false 'CCFFFF' 'Medium' $null
            $cell.Style.HorizontalAlignment = 'Center'
        }

        try { $wsOut1.Column(2).Width = 40; $wsOut1.Column(3).Width = 40 } catch {}

        if ($sigMismatch) {
            # === MISMATCH: Röd markering och detaljerad tabell med fet border ===
            $mismatchCell = $wsOut1.Cells["D$signRow"]
            $mismatchCell.Value = '⚠ Mismatch'
            Style-Cell $mismatchCell $true 'FF0000' 'Medium' 'FFFFFF'

            if ($mismatchSheets.Count -gt 0) {
                # Rubrikrad för mismatch-detaljer
                $headerRow = $signRow + 1
                $wsOut1.Cells["A$headerRow"].Value = 'Fil'
                $wsOut1.Cells["B$headerRow"].Value = 'Signatur'
                $wsOut1.Cells["C$headerRow"].Value = 'Blad'
                foreach ($col in @('A','B','C')) {
                    $hdrCell = $wsOut1.Cells["$col$headerRow"]
                    $hdrCell.Style.Font.Bold = $true
                    $hdrCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $hdrCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml('#FFE699'))
                    $hdrCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                }

                $lastDataRow = $headerRow
                for ($j = 0; $j -lt $mismatchSheets.Count; $j++) {
                    $rowIdx = $headerRow + 1 + $j
                    $lastDataRow = $rowIdx
                    $text = $mismatchSheets[$j]

                    # Parsa "NEG: signatur  [Blad: namn]" eller "POS: signatur  [Blad: namn]"
                    $filType = ''; $sigVal = ''; $bladVal = ''
                    if ($text -match '^(NEG|POS):\s*(.+?)\s*\[Blad:\s*(.+?)\]$') {
                        $filType = $matches[1]
                        $sigVal  = $matches[2].Trim()
                        $bladVal = $matches[3].Trim()
                    } else {
                        $sigVal = $text
                    }

                    $wsOut1.Cells["A$rowIdx"].Value = $filType
                    $wsOut1.Cells["B$rowIdx"].Value = $sigVal
                    $wsOut1.Cells["C$rowIdx"].Value = $bladVal
                    # WrapText för "Blad" vid signatur-mismatch (bättre läsbarhet)
                    try {
                        $cBlad = $wsOut1.Cells["C$rowIdx"]
                        $cBlad.Style.WrapText = $true
                        $cBlad.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
                        $wsOut1.Row($rowIdx).CustomHeight = $true
                    } catch {}

                    # Färgkoda per fil
                    $bgColor = if ($filType -eq 'NEG') { '#B5E6A2' } else { '#FFB3B3' }
                    foreach ($col in @('A','B','C')) {
                        $dataCell = $wsOut1.Cells["$col$rowIdx"]
                        $dataCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $dataCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($bgColor))
                    }
                }
                
                # === FET YTTRE BORDER runt hela mismatch-tabellen ===
                try {
                    $borderColor = [System.Drawing.Color]::FromArgb(0, 0, 0) # SVART
                    
                    # Topp-border på header-raden
                    for ($c = 1; $c -le 3; $c++) {
                        $wsOut1.Cells[$headerRow, $c].Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thick
                        $wsOut1.Cells[$headerRow, $c].Style.Border.Top.Color.SetColor($borderColor)
                    }
                    
                    # Botten-border på sista data-raden
                    for ($c = 1; $c -le 3; $c++) {
                        $wsOut1.Cells[$lastDataRow, $c].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thick
                        $wsOut1.Cells[$lastDataRow, $c].Style.Border.Bottom.Color.SetColor($borderColor)
                    }
                    
                    # Vänster-border på kolumn A (alla rader)
                    for ($r = $headerRow; $r -le $lastDataRow; $r++) {
                        $wsOut1.Cells[$r, 1].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thick
                        $wsOut1.Cells[$r, 1].Style.Border.Left.Color.SetColor($borderColor)
                    }
                    
                    # Höger-border på kolumn C (alla rader)
                    for ($r = $headerRow; $r -le $lastDataRow; $r++) {
                        $wsOut1.Cells[$r, 3].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thick
                        $wsOut1.Cells[$r, 3].Style.Border.Right.Color.SetColor($borderColor)
                    }
                    
                    # Inre tunna borders
                    for ($r = $headerRow; $r -le $lastDataRow; $r++) {
                        for ($c = 1; $c -le 3; $c++) {
                            if ($r -lt $lastDataRow) {
                                $wsOut1.Cells[$r, $c].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                            }
                            if ($c -lt 3) {
                                $wsOut1.Cells[$r, $c].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                            }
                        }
                    }
                } catch {}
            }
        } else {
            # === MATCH: Grön markering när signaturer stämmer ===
            if ($hasNeg -and $hasPos) {
                $matchCell = $wsOut1.Cells["D$signRow"]
                $matchCell.Value = '✓ Match'
                Style-Cell $matchCell $true 'C6EFCE' 'Medium' '006100'
            }
        }

        # ============================
        # === STF Sum              ===
        # ============================

        $wsOut2 = $pkgOut.Workbook.Worksheets["STF Sum"]
        if (-not $wsOut2) { Gui-Log "❌ Fliken 'STF Sum' saknas i mallen!"; return }

        $totalRows = $violationsNeg.Count + $violationsPos.Count
        $currentRow = 2

        if ($totalRows -eq 0) {
            Gui-Log "✅ Seal Test hittades"
            $wsOut2.Cells["B1:H1"].Value = $null
            $wsOut2.Cells["A1"].Value = "Inga STF hittades!"
            Style-Cell $wsOut2.Cells["A1"] $true "D9EAD3" "Medium" "006100"
            $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
            if ($wsOut2.Dimension -and $wsOut2.Dimension.End.Row -gt 1) { $wsOut2.DeleteRow(2, $wsOut2.Dimension.End.Row - 1) }
        }
        else {
            Gui-Log "❗ $failNegCount avvikelser i NEG, $failPosCount i POS"

            $oldDataRows = 0
            if ($wsOut2.Dimension) {
                $oldDataRows = $wsOut2.Dimension.End.Row - 1
                if ($oldDataRows -lt 0) { $oldDataRows = 0 }
            }

            if ($totalRows -lt $oldDataRows) {
                $wsOut2.DeleteRow(2 + $totalRows, $oldDataRows - $totalRows)
            }
            elseif ($totalRows -gt $oldDataRows) {
                $wsOut2.InsertRow(2 + $oldDataRows, $totalRows - $oldDataRows, 1 + $oldDataRows)
            }

            $currentRow = 2
            foreach ($v in $violationsNeg) {
                $wsOut2.Cells["A$currentRow"].Value = "NEG"
                $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
                $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
                $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
                $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
                $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
                $wsOut2.Cells["G$currentRow"].Value = $v.Status
                $wsOut2.Cells["H$currentRow"].Value = if ([string]::IsNullOrWhiteSpace($v.Obs)) { 'NA' } else { $v.Obs }

                Style-Cell $wsOut2.Cells["A$currentRow"] $true "B5E6A2" "Medium" $null
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
                $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))

                if ($v.Status -in @("FAIL","Minusvärde")) {
                    $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                    $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                }

                Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
                $currentRow++
            }

            foreach ($v in $violationsPos) {
                $wsOut2.Cells["A$currentRow"].Value = "POS"
                $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
                $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
                $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
                $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
                $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
                $wsOut2.Cells["G$currentRow"].Value = $v.Status
                $wsOut2.Cells["H$currentRow"].Value = if ($v.Obs) { $v.Obs } else { 'NA' }

                Style-Cell $wsOut2.Cells["A$currentRow"] $true "FFB3B3" "Medium" $null
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
                $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))

                if ($v.Status -in @("FAIL","Minusvärde")) {
                    $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                    $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                }

                Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
                $currentRow++
            }

            $wsOut2.Cells.Style.WrapText = $false
            $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
            try { $wsOut2.Cells[2,6,([Math]::Max($currentRow-1,2)),6].Style.Numberformat.Format = '0.0' } catch {}
            if ($wsOut2.Dimension) {
                try {
                    if (Get-Command Safe-AutoFitColumns -ErrorAction SilentlyContinue) {
                        Safe-AutoFitColumns -Ws $wsOut2 -Context 'OutputSheet'
                    } else {
                        $wsOut2.Cells[$wsOut2.Dimension.Address].AutoFitColumns() | Out-Null
                    }
                } catch {}
            }
        }


# ============================
# === Information-blad     ===
# ============================

try {
    if (-not (Get-Command Add-Hyperlink -ErrorAction SilentlyContinue)) {
        function Add-Hyperlink {
            param(
                [OfficeOpenXml.ExcelRange]$Cell,
                [string]$Text,
                [string]$Url
            )
            try {
                $Cell.Value = $Text
                $Cell.Hyperlink = [Uri]$Url
                $Cell.Style.Font.UnderLine = $true
                $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0,102,204))
            } catch {}
        }
    }

    if (-not (Get-Command Find-RegexCell -ErrorAction SilentlyContinue)) {
        function Find-RegexCell {
            param(
                [OfficeOpenXml.ExcelWorksheet]$Ws,
                [regex]$Rx,
                [int]$MaxRows = 200,
                [int]$MaxCols = 40
            )
            if (-not $Ws -or -not $Ws.Dimension) { return $null }

            $rMax = [Math]::Min($Ws.Dimension.End.Row, $MaxRows)
            $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)

            for ($r = 1; $r -le $rMax; $r++) {
                for ($c = 1; $c -le $cMax; $c++) {
                    $t = Normalize-HeaderText ($Ws.Cells[$r,$c].Text + '')
                    if ($t -and $Rx.IsMatch($t)) {
                        return @{ Row = $r; Col = $c; Text = $t }
                    }
                }
            }
            return $null
        }
    }

    if (-not (Get-Command Get-SealHeaderDocInfo -ErrorAction SilentlyContinue)) {
        function Get-SealHeaderDocInfo {
            param([OfficeOpenXml.ExcelPackage]$Pkg)

            $result = [pscustomobject]@{ Raw=''; DocNo=''; Rev='' }
            if (-not $Pkg) { return $result }

            $ws = $Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
            if (-not $ws) { return $result }

            try {
                $lt = ($ws.HeaderFooter.OddHeader.LeftAlignedText + '').Trim()
                if (-not $lt) { $lt = ($ws.HeaderFooter.EvenHeader.LeftAlignedText + '').Trim() }

                $result.Raw = $lt

                $rx = [regex]'(?i)(?:document\s*(?:no|nr|#|number)\s*[:#]?\s*([A-Z0-9\-_\.\/]+))?.*?(?:rev(?:ision)?\.?\s*[:#]?\s*([A-Z0-9\-_\.]+))?'
                $m = $rx.Match($lt)
                if ($m.Success) {
                    if ($m.Groups[1].Value) { $result.DocNo = $m.Groups[1].Value.Trim() }
                    if ($m.Groups[2].Value) { $result.Rev   = $m.Groups[2].Value.Trim() }
                }
            } catch {}

            return $result
        }
    }

    $wsInfo = $pkgOut.Workbook.Worksheets['Information']
    if (-not $wsInfo) { $wsInfo = $pkgOut.Workbook.Worksheets.Add('Information') }

    try {
        $csvLines = $null
        $csvStats = $null

        # ✅ Viktigt: initiera alltid, även om ingen CSV är vald
        $csvInstrumentSerials = @()

        if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
            try { $csvLines = Get-Content -LiteralPath $selCsv } catch { Gui-Log ("⚠️ Kunde inte läsa CSV: " + $_.Exception.Message) 'Warn' }
            try { $csvStats = Get-CsvStats -Path $selCsv -Lines $csvLines } catch { Gui-Log ("⚠️ Get-CsvStats: " + $_.Exception.Message) 'Warn' }

            # Extract exact Instrument S/N list from CSV (for Equipment sheet when WS scan is disabled)
            try {
                if ($csvLines -and $csvLines.Count -gt 8) {
                    $hdr = ConvertTo-CsvFields $csvLines[7]

                    $idx = -1
                    for ($ii=0; $ii -lt $hdr.Count; $ii++) {
                        $h = (($hdr[$ii] + '').Trim('\"').ToLower())
                        if ($h -eq 'instrument s/n' -or $h -match 'instrument') { $idx = $ii; break }
                    }

                    if ($idx -ge 0) {
                        $set = New-Object System.Collections.Generic.HashSet[string]
                        for ($rr=9; $rr -lt $csvLines.Count; $rr++) {
                            $ln = $csvLines[$rr]
                            if (-not $ln -or -not $ln.Trim()) { continue }
                            $f = ConvertTo-CsvFields $ln
                            if ($f.Count -gt $idx) {
                                $v = ($f[$idx] + '').Trim().Trim('\"')
                                if ($v) { $null = $set.Add($v) }
                            }
                        }

                        # HashSet[T] enumerate (PS 5.1 safe)
                        $csvInstrumentSerials = @($set | Sort-Object)
                    }
                }
            } catch {
                Gui-Log ("⚠️ Kunde inte extrahera Instrument S/N från CSV: " + $_.Exception.Message) 'Warn'
                $csvInstrumentSerials = @()
            }
        }

        if (-not $csvStats) {
            $csvStats = [pscustomobject]@{
                TestCount    = 0
                DupCount     = 0
                Duplicates   = @()
                LspValues    = @()
                LspOK        = $null
                InstrumentByType = [ordered]@{}
            }
        }

        $infSN = @()
        if ($script:GXINF_Map) {
            foreach ($k in $script:GXINF_Map.Keys) {
                if ($k -like 'Infinity-*') {
                    $infSN += ($script:GXINF_Map[$k].Split(',') | ForEach-Object { ($_ + '').Trim() } | Where-Object { $_ })
                }
            }
        }
        $infSN = $infSN | Select-Object -Unique

        $infSummary = '—'
        try {
            if ($selCsv -and (Test-Path -LiteralPath $selCsv) -and $infSN.Count -gt 0) {
                $infSummary = Get-InfinitySpFromCsvStrict -Path $selCsv -InfinitySerials $infSN -Lines $csvLines
            }
        } catch {
            Gui-Log ("Infinity SP fel: " + $_.Exception.Message) 'Warn'
        }

        # --- Dubbletter Sample ID ---
        $dupSampleCount = 0
        $dupSampleList  = @()
        if ($csvLines -and $csvLines.Count -gt 8) {
            try {
                $headerFields = ConvertTo-CsvFields $csvLines[7]
                $sampleIdx = -1
                for ($i=0; $i -lt $headerFields.Count; $i++) {
                    $hf = ($headerFields[$i] + '').Trim().ToLower()
                    if ($hf -match 'sample') { $sampleIdx = $i; break }
                }
                if ($sampleIdx -ge 0) {
                    $samples = @()
                    for ($r=9; $r -lt $csvLines.Count; $r++) {
                        $line = $csvLines[$r]
                        if (-not $line -or -not $line.Trim()) { continue }
                        $fields = ConvertTo-CsvFields $line
                        if ($fields.Count -gt $sampleIdx) {
                            $val = ($fields[$sampleIdx] + '').Trim()
                            if ($val) { $samples += $val }
                        }
                    }
                    if ($samples.Count -gt 0) {
                        $counts = @{}
                        foreach ($s in $samples) { if (-not $counts.ContainsKey($s)) { $counts[$s] = 0 }; $counts[$s]++ }
                        foreach ($entry in $counts.GetEnumerator()) {
                            if ($entry.Value -gt 1) { $dupSampleList += ("$($entry.Key) x$($entry.Value)") }
                        }
                        $dupSampleCount = $dupSampleList.Count
                    }
                }
            } catch {
                Gui-Log ("⚠️ Fel vid analys av Sample ID: " + $_.Exception.Message) 'Warn'
            }
        }

        $dupSampleText = if ($dupSampleCount -gt 0) {
            $show = ($dupSampleList | Select-Object -First 8) -join ', '
            "$dupSampleCount ($show)"
        } else { 'N/A' }

        $dupCartText = if ($csvStats.DupCount -gt 0) {
            $show = ($csvStats.Duplicates | Select-Object -First 8) -join ', '
            "$($csvStats.DupCount) ($show)"
        } else { 'N/A' }

        # --- LSP summary ---
        $lspSummary = ''
        try {
            if ($csvLines -and $csvLines.Count -gt 8) {
                $counts = @{}
                for ($rr = 9; $rr -lt $csvLines.Count; $rr++) {
                    $ln = $csvLines[$rr]
                    if (-not $ln -or -not $ln.Trim()) { continue }
                    $fs = ConvertTo-CsvFields $ln
                    if ($fs.Count -gt 4) {
                        $raw = ($fs[4] + '').Trim()
                        if ($raw) {
                            $mLsp = [regex]::Match($raw,'(\d{5})')
                            $code = if ($mLsp.Success) { $mLsp.Groups[1].Value } else { $raw }
                            if (-not $counts.ContainsKey($code)) { $counts[$code] = 0 }
                            $counts[$code]++
                        }
                    }
                }

                if ($counts.Count -gt 0) {
                    $sorted = $counts.GetEnumerator() | Sort-Object Key
                    $parts = @()
                    foreach ($kvp in $sorted) {
                        $parts += $(if ($kvp.Value -gt 1) { "$($kvp.Key) x$($kvp.Value)" } else { $kvp.Key })
                    }
                    if ($sorted.Count -eq 1) { $lspSummary = $sorted[0].Key }
                    else { $lspSummary = "$($sorted.Count) (" + ($parts -join ', ') + ")" }
                }
            }
        } catch {
            Gui-Log ("⚠️ Fel vid extraktion av LSP från CSV: " + $_.Exception.Message) 'Warn'
            $lspSummary = ''
        }

        $instText = if ($csvStats.InstrumentByType.Keys.Count -gt 0) {
            ($csvStats.InstrumentByType.GetEnumerator() | ForEach-Object { "$($_.Key)" } | Sort-Object) -join '; '
        } else { '' }

        function Find-InfoRow {
            param([OfficeOpenXml.ExcelWorksheet]$Ws, [string]$Label)
            if (-not $Ws -or -not $Ws.Dimension) { return $null }
            $maxRow = [Math]::Min($Ws.Dimension.End.Row, 300)
            for ($ri=1; $ri -le $maxRow; $ri++) {
                $txt = (($Ws.Cells[$ri,1].Text) + '').Trim()
                if (-not $txt) { continue }
                if ($txt.ToLowerInvariant() -eq $Label.ToLowerInvariant()) { return $ri }
            }
            return $null
        }

        # ✅ Default: anta INTE new layout om vi inte hittar ankaret
        $isNewLayout = $false
        try {
            $tmpRow = Find-InfoRow -Ws $wsInfo -Label 'CSV-Info'
            if ($tmpRow) { $isNewLayout = $true }
        } catch {}

        $rowCsvFile    = Find-InfoRow -Ws $wsInfo -Label 'CSV'
        $rowLsp        = Find-InfoRow -Ws $wsInfo -Label 'LSP'
        $rowAntal      = Find-InfoRow -Ws $wsInfo -Label 'Antal tester'
        $rowDupSample  = Find-InfoRow -Ws $wsInfo -Label 'Dubblett Sample ID'
        if (-not $rowDupSample) { $rowDupSample = Find-InfoRow -Ws $wsInfo -Label 'Dublett Sample ID' }

        $rowDupCart    = Find-InfoRow -Ws $wsInfo -Label 'Dubblett Cartridge S/N'
        if (-not $rowDupCart) { $rowDupCart = Find-InfoRow -Ws $wsInfo -Label 'Dublett Cartridge S/N' }

        $rowInst = Find-InfoRow -Ws $wsInfo -Label 'Använda INF/GX'

        $rowBag = Find-InfoRow -Ws $wsInfo -Label 'Bag Numbers Tested Using Infinity'
        if (-not $rowBag) { $rowBag = Find-InfoRow -Ws $wsInfo -Label 'Bag Numbers Tested Using Infinity:' }
        if (-not $rowBag) { $rowBag = 14 }

        $wsInfo.Cells["B$rowBag"].Style.Numberformat.Format = '@'
        $wsInfo.Cells["B$rowBag"].Value = $infSummary

        if ($isNewLayout) {
            if (-not $rowCsvFile)   { $rowCsvFile   = 8 }
            if (-not $rowLsp)       { $rowLsp       = 9 }
            if (-not $rowAntal)     { $rowAntal     = 10 }
            if (-not $rowDupSample) { $rowDupSample = 11 }
            if (-not $rowDupCart)   { $rowDupCart   = 12 }
            if (-not $rowInst)      { $rowInst      = 13 }
        }

        if ($selCsv) {
            $wsInfo.Cells["B$rowCsvFile"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowCsvFile"].Value = (Split-Path $selCsv -Leaf)
        } else {
            $wsInfo.Cells["B$rowCsvFile"].Value = ''
        }

        $wsInfo.Cells["B$rowLsp"].Style.Numberformat.Format = '@'
        $wsInfo.Cells["B$rowLsp"].Value = $(if ($lspSummary) { $lspSummary } else { $lsp })

        # ✅ Skriv Antal tester EN gång (som text för att inte bli 1.00E+3 etc)
        $wsInfo.Cells["B$rowAntal"].Style.Numberformat.Format = '@'
        $wsInfo.Cells["B$rowAntal"].Value = "$($csvStats.TestCount)"

        if ($rowDupSample) { $wsInfo.Cells["B$rowDupSample"].Value = $dupSampleText }
        if ($rowDupCart)   { $wsInfo.Cells["B$rowDupCart"].Value   = $dupCartText }

        if ($rowInst) { $wsInfo.Cells["B$rowInst"].Value = $instText }

    } catch {
        Gui-Log ("⚠️ CSV data-fel: " + $_.Exception.Message) 'Warn'
    }

    # --- Macro / docinfo ---
    $assayForMacro = ''
    if ($runAssay) { $assayForMacro = $runAssay }
    elseif ($wsOut1) { $assayForMacro = ($wsOut1.Cells['D10'].Text + '').Trim() }

    $miniVal = ''
    if (Get-Command Get-MinitabMacro -ErrorAction SilentlyContinue) {
        $miniVal = Get-MinitabMacro -AssayName $assayForMacro
    }
    if (-not $miniVal) { $miniVal = 'N/A' }

    $hdNeg = $null; $hdPos = $null
    try { $hdNeg = Get-SealHeaderDocInfo -Pkg $pkgNeg } catch {}
    try { $hdPos = Get-SealHeaderDocInfo -Pkg $pkgPos } catch {}
    if (-not $hdNeg) { $hdNeg = [pscustomobject]@{ Raw=''; DocNo=''; Rev='' } }
    if (-not $hdPos) { $hdPos = [pscustomobject]@{ Raw=''; DocNo=''; Rev='' } }

    $wsInfo.Cells['B2'].Value = $ScriptVersion
    $wsInfo.Cells['B3'].Value = $env:USERNAME
    $wsInfo.Cells['B4'].Value = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    $wsInfo.Cells['B5'].Value = $miniVal

    # --- Batch + länkar ---
    $selLsp = $null
    try {
        if (Get-Variable -Name clbLsp -ErrorAction SilentlyContinue) {
            $selLsp = Get-CheckedFilePath $clbLsp
        }
    } catch {}

    $batchInfo = Get-BatchLinkInfo -SealPosPath $selPos -SealNegPath $selNeg -Lsp $lspForLinks
    $batch     = $batchInfo.Batch

    $wsInfo.Cells['A34'].Value = 'SharePoint Batch'
    $wsInfo.Cells['A34'].Style.Font.Bold = $true
    Add-Hyperlink -Cell $wsInfo.Cells['B34'] -Text $batchInfo.LinkText -Url $batchInfo.Url

    $linkMap = [ordered]@{
        'IPT App'      = 'https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83'
        'MES Login'    = 'http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM'
        'CSV Uploader' = 'http://auw2wgxtpap01.cepaws.com/Welcome.aspx'
        'BMRAM'        = 'https://cepheid62468.coolbluecloud.com/'
        'Agile'        = 'https://agileprod.cepheid.com/Agile/default/login-cms.jsp'
    }

    $rowLink = 35
    foreach ($key in $linkMap.Keys) {
        $wsInfo.Cells["A$rowLink"].Value = $key
        Add-Hyperlink -Cell $wsInfo.Cells["B$rowLink"] -Text 'LÄNK' -Url $linkMap[$key]
        $rowLink++
    }

} catch {
    Gui-Log ("⚠️ Information-blad fel: " + $_.Exception.Message) 'Warn'
}
                        
# ----------------------------------------------------------------
# WS (LSP Worksheet): hitta fil och skriv in i Information-bladet
# ----------------------------------------------------------------
try {
    if (-not $selLsp) {
        $probeDir = $null
        if ($selPos) { $probeDir = Split-Path -Parent $selPos }
        if (-not $probeDir -and $selNeg) { $probeDir = Split-Path -Parent $selNeg }

        if ($probeDir -and (Test-Path -LiteralPath $probeDir)) {
            $cand = Get-ChildItem -LiteralPath $probeDir -File -ErrorAction SilentlyContinue |
                    Where-Object {
                        ($_.Name -match '(?i)worksheet') -and
                        ($_.Name -match [regex]::Escape($lsp)) -and
                        ($_.Extension -match '^\.(xlsx|xlsm|xls)$')
                    } |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1

            if ($cand) { $selLsp = $cand.FullName }
        }
    }

    if (-not (Get-Command Find-LabelValueRightward -ErrorAction SilentlyContinue)) {
        function Find-LabelValueRightward {
            param(
                [OfficeOpenXml.ExcelWorksheet]$Ws,
                [string]$Label,
                [int]$MaxRows = 200,
                [int]$MaxCols = 40
            )
            if (-not $Ws -or -not $Ws.Dimension) { return $null }

            $normLbl = Normalize-HeaderText $Label
            $pat = '^(?i)\s*' + [regex]::Escape($normLbl).Replace('\ ', '\s*') + '\s*[:\.]*\s*$'
            $rx  = [regex]::new($pat, [Text.RegularExpressions.RegexOptions]::IgnoreCase)

            $hit = Find-RegexCell -Ws $Ws -Rx $rx -MaxRows $MaxRows -MaxCols $MaxCols
            if (-not $hit) { return $null }

            $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
            for ($c = $hit.Col + 1; $c -le $cMax; $c++) {
                $t = Normalize-HeaderText ($Ws.Cells[$hit.Row,$c].Text + '')
                if ($t) { return $t }
            }
            return $null
        }
    }

    if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
        Gui-Log ("🔎 WS hittad: " + (Split-Path $selLsp -Leaf)) 'Info'
    } else {
        Gui-Log "ℹ️ Ingen WS-fil vald/hittad (LSP Worksheet). Hoppar över WS-extraktion." 'Info'
    }
} catch {
    Gui-Log ("⚠️ WS-block fel: " + $_.Exception.Message) 'Warn'
}

try {
    # Se till att dessa alltid finns (de används senare)
    $headerWs  = $null
    $headerNeg = $null
    $headerPos = $null
    # OBS: $eqInfo används senare vid Equipment-blad → init här
    if (-not (Get-Variable -Name eqInfo -Scope 1 -ErrorAction SilentlyContinue)) {
        $eqInfo = $null
    }

    # --- WS öppnas EN gång och dispose:as alltid ---
    $tmpPkg = $null
    try {
        if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
            try {
                $tmpPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selLsp))

                # ==============================
                # Equipment / TestSummary
                # ==============================
                try {
                    # "EquipmentSheet" ska endast generera själva BLADET/mallen.
                    # Det ska *inte* automatiskt hämta pipetter/instrument från Test Summary/CSV här,
                    # annars blir det en blandning (auto + manuellt) som förvirrar användaren.
                    if ($Config -and $Config.Contains('EnableEquipmentSheet') -and -not $Config.EnableEquipmentSheet) {
                        $eqInfo = $null
                        Gui-Log 'ℹ️ Utrustningslista avstängt. Hoppar över.' 'Info'
                    } else {
                        $eqInfo = $null
                        Gui-Log '✅ Utrustningslista hämtad.' 'Info'
                    }
                } catch {
                    # Utrustning ska aldrig få stoppa rapporten
                    try { Gui-Log ("⚠️ Utrustningslista: kunde inte hämtas: " + $_.Exception.Message) 'Warn' } catch {}
                    $eqInfo = $null
                }

                # ==============================
                # Worksheet header (Extract + Compare)
                # ==============================
                try {
                    $headerWs = Extract-WorksheetHeader -Pkg $tmpPkg
                } catch {
                    $headerWs = $null
                }

                # Fallback om Extract gav null: skapa tomt objekt så .PartNo osv aldrig kraschar
                if (-not $headerWs) {
                    $headerWs = [pscustomobject]@{
                        WorksheetName   = ''
                        PartNo          = ''
                        BatchNo         = ''
                        CartridgeNo     = ''
                        DocumentNumber  = ''
                        Rev             = ''
                        Effective       = ''
                        Attachment      = ''
                    }
                }

                try {
                    $wsHeaderRows  = Get-WorksheetHeaderPerSheet -Pkg $tmpPkg
                    $wsHeaderCheck = Compare-WorksheetHeaderSet   -Rows $wsHeaderRows
                    try {
                        if ($wsHeaderCheck.Issues -gt 0 -and $wsHeaderCheck.Summary) {
                            Gui-Log ("⚠️ Worksheet header-avvikelser: {0} – se Information!" -f $wsHeaderCheck.Summary) 'Warn'
                        } else {
                            Gui-Log "✅ Worksheet header korrekt" 'Info'
                        }
                    } catch {}
                } catch {
                    # Behåll tyst/robust
                }

                # ==============================
                # Förstärk headerWs via labels (om Extract missade)
                # ==============================
                try {
                    $wsLsp = $tmpPkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsLsp) {

                        if (-not $headerWs.PartNo) {
                            $val = $null
                            $labels = @('Part No.','Part No.:','Part No','Part Number','Part Number:','Part Number.','Part Number.:')
                            foreach ($lbl in $labels) { $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl; if ($val) { break } }
                            if ($val) { $headerWs.PartNo = $val }
                        }

                        if (-not $headerWs.BatchNo) {
                            $val = $null
                            $labels = @(
                                'Batch No(s)','Batch No(s).','Batch No(s):','Batch No(s).:',
                                'Batch No','Batch No.','Batch No:','Batch No.:',
                                'Batch Number','Batch Number.','Batch Number:','Batch Number.:'
                            )
                            foreach ($lbl in $labels) { $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl; if ($val) { break } }
                            if ($val) { $headerWs.BatchNo = $val }
                        }

                        if (-not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.') {
                            $val = $null
                            $labels = @(
                                'Cartridge No. (LSP)','Cartridge No. (LSP):','Cartridge No. (LSP) :',
                                'Cartridge No (LSP)','Cartridge No (LSP):','Cartridge No (LSP) :',
                                'Cartridge Number (LSP)','Cartridge Number (LSP):','Cartridge Number (LSP) :',
                                'Cartridge No.','Cartridge No.:','Cartridge No. :','Cartridge No :',
                                'Cartridge Number','Cartridge Number:','Cartridge Number :',
                                'Cartridge No','Cartridge No:','Cartridge No :'
                            )
                            foreach ($lbl in $labels) { $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl; if ($val) { break } }

                            if (-not $val) {
                                $rxCart = [regex]::new('(?i)Cartridge.*\(LSP\)')
                                $maxCols = [Math]::Min($wsLsp.Dimension.End.Column, 100)
                                $hitCart = Find-RegexCell -Ws $wsLsp -Rx $rxCart -MaxRows 200 -MaxCols $maxCols
                                if ($hitCart) {
                                    for ($c = $hitCart.Col + 1; $c -le $wsLsp.Dimension.End.Column; $c++) {
                                        $cellVal = ($wsLsp.Cells[$hitCart.Row, $c].Text + '').Trim()
                                        if ($cellVal) { $val = $cellVal; break }
                                    }
                                }
                            }

                            if ($val) { $headerWs.CartridgeNo = $val }
                        }

                        if (-not $headerWs.Effective) {
                            $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective'
                            if (-not $val) { $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective Date' }
                            if ($val) { $headerWs.Effective = $val }
                        }
                    }
                } catch {}

                # Filename fallback för CartridgeNo om fortfarande tomt
                try {
                    if ($selLsp -and (-not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.' -or $headerWs.CartridgeNo -eq '')) {
                        $fn = Split-Path $selLsp -Leaf
                        $m = [regex]::Matches($fn, '(?<!\d)(\d{5,7})(?!\d)')
                        if ($m.Count -gt 0) { $headerWs.CartridgeNo = $m[0].Groups[1].Value }
                    }
                } catch {}

                # ==============================
                # QC Reminder – HIV/HBV/HCV assays (läser Test Summary B3 medan $tmpPkg är öppen)
                # ==============================
                $script:QcReminderB3 = $null
                try {
                    $qcTriggerAssays = @(
                        'Xpert_HIV-1 Viral Load',
                        'Xpert HIV-1 Viral Load XC',
                        'Xpert_HCV Viral Load',
                        'Xpert HCV VL Fingerstick',
                        'Xpert HBV Viral Load',
                        'Xpert_HIV-1 Qual',
                        'HIV-1_Qual RUO',
                        'HIV-1 Qual XC DBS RUO',
                        'HIV-1 Qual XC DBS IUO',
                        'Xpert HIV-1 Qual XC PQC',
                        'Xpert HIV-1 Qual XC PQC RUO'


                    )
                    $qcValidPartNos = @(
                        'GXHIV-VL-CE-10','GXHIV-VL-IN-10','GXHIV-VL-CN-10',
                        'GXHIV-VL-XC-CE-10','GXHCV-VL-CE-10','GXHCV-VL-IN-10',
                        'GXHCV-FS-CE-10','GXHBV-VL-CE-10',
                        'GXHIV-QA-CE-10','RHIVQ-10'
                        '700-6098/HIV-1 QUAL XC, RUO','700-6137/HIV-1 QUAL XC, IUO',
                        '700-6793/ HIV-1 QUAL XC, CE-IVD','700-6911/ HIV-1 QUAL XC, RUO'
                    )

                    $qcAssayMatch = $false
                    if ($runAssay) {
                        foreach ($_qa in $qcTriggerAssays) {
                            if ($runAssay -ilike "$_qa*") { $qcAssayMatch = $true; break }
                        }
                    }

                    if ($qcAssayMatch) {
                        $tsWs = $tmpPkg.Workbook.Worksheets | Where-Object { $_.Name -ieq 'Test Summary' } | Select-Object -First 1
                        if ($tsWs) {
                            $b3Val = ($tsWs.Cells['B3'].Text + '').Trim()
                            if ($b3Val -and ($qcValidPartNos -contains $b3Val)) {
                                $script:QcReminderB3 = $b3Val
                                Gui-Log ("🔬 QC Reminder: Assay={0}, Test Summary B3={1} → aktiverad" -f $runAssay, $b3Val) 'Info'
                            } else {
                                Gui-Log ("ℹ️ QC Reminder: Assay matchar men B3='{0}' ej i listan → hoppar över." -f $b3Val) 'Info'
                            }
                        } else {
                            Gui-Log "ℹ️ QC Reminder: Assay matchar men 'Test Summary'-blad saknas i Worksheet." 'Info'
                        }
                    }
                } catch {
                    Gui-Log ("⚠️ QC Reminder (läsning): " + $_.Exception.Message) 'Warn'
                }

            } catch {
                # Om WS öppnas men något inne fallerar: låt det inte döda hela rapporten
                Gui-Log ("⚠️ WS-parse fel: " + $_.Exception.Message) 'Warn'
            }
        }
    } finally {
        if ($tmpPkg) { try { $tmpPkg.Dispose() } catch {} }
    }

    # SealTest headers
    try { $headerNeg = Extract-SealTestHeader -Pkg $pkgNeg } catch {}
    try { $headerPos = Extract-SealTestHeader -Pkg $pkgPos } catch {}

    # Effective fallback för Seal POS/NEG om saknas
    try {
        if ($pkgPos -and $headerPos -and -not $headerPos.Effective) {
            $wsPos = $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
            if ($wsPos) {
                $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective'
                if (-not $val) { $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective Date' }
                if ($val) { $headerPos.Effective = $val }
            }
        }
    } catch {}

    try {
        if ($pkgNeg -and $headerNeg -and -not $headerNeg.Effective) {
            $wsNeg = $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
            if ($wsNeg) {
                $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective'
                if (-not $val) { $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective Date' }
                if ($val) { $headerNeg.Effective = $val }
            }
        }
    } catch {}

    # --------------------------
    # Header summary → Information
    # --------------------------
    try {
        $wsBatch   = if ($headerWs -and $headerWs.BatchNo) { $headerWs.BatchNo } else { $null }
        $sealBatch = $batch
        if (-not $sealBatch) {
            try { if ($selPos) { $sealBatch = Get-BatchNumberFromSealFile $selPos } } catch {}
            if (-not $sealBatch) { try { if ($selNeg) { $sealBatch = Get-BatchNumberFromSealFile $selNeg } } catch {} }
        }

        $rowWsFile = Find-InfoRow -Ws $wsInfo -Label 'Worksheet'
        if (-not $rowWsFile) { $rowWsFile = 17 }
        $rowPart  = $rowWsFile + 1
        $rowBatch = $rowWsFile + 2
        $rowCart  = $rowWsFile + 3
        $rowDoc   = $rowWsFile + 4
        $rowRev   = $rowWsFile + 5
        $rowEff   = $rowWsFile + 6

        $rowPosFile = Find-InfoRow -Ws $wsInfo -Label 'Seal Test POS'
        if (-not $rowPosFile) { $rowPosFile = $rowWsFile + 7 }
        $rowPosDoc = $rowPosFile + 1
        $rowPosRev = $rowPosFile + 2
        $rowPosEff = $rowPosFile + 3

        $rowNegFile = Find-InfoRow -Ws $wsInfo -Label 'Seal Test NEG'
        if (-not $rowNegFile) { $rowNegFile = $rowPosFile + 4 }
        $rowNegDoc = $rowNegFile + 1
        $rowNegRev = $rowNegFile + 2
        $rowNegEff = $rowNegFile + 3

        if ($selLsp) {
            $wsInfo.Cells["B$rowWsFile"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowWsFile"].Value = (Split-Path $selLsp -Leaf)
        } else {
            $wsInfo.Cells["B$rowWsFile"].Value = ''
        }

        $consPart  = Get-ConsensusValue -Type 'Part'      -Ws $headerWs.PartNo      -Pos $headerPos.PartNumber   -Neg $headerNeg.PartNumber
        $consBatch = Get-ConsensusValue -Type 'Batch'     -Ws $headerWs.BatchNo     -Pos $headerPos.BatchNumber  -Neg $headerNeg.BatchNumber
        $consCart  = Get-ConsensusValue -Type 'Cartridge' -Ws $headerWs.CartridgeNo -Pos $headerPos.CartridgeNo  -Neg $headerNeg.CartridgeNo

        if (-not $consCart.Value -and $selLsp) {
            $fnCart = Split-Path $selLsp -Leaf
            $mCart  = [regex]::Match($fnCart,'(?<!\d)(\d{5,7})(?!\d)')
            if ($mCart.Success) {
                $consCart = @{ Value=$mCart.Groups[1].Value; Source='FILENAME'; Note='Filename fallback' }
            }
        }

        $wsInfo.Cells["B$rowPart"].Value  = if ($consPart.Value)  { $consPart.Value }  else { '' }
        $wsInfo.Cells["B$rowBatch"].Value = if ($consBatch.Value) { $consBatch.Value } else { '' }
        $wsInfo.Cells["B$rowCart"].Value  = if ($consCart.Value)  { $consCart.Value }  else { '' }

        # wsHeaderCheck → skriv Match/Mismatch i C-kolumn för Part/Batch/Cartridge
        try {
            $StyleMismatchCell = {
                param($Cell, $Text)
                $Cell.Style.Numberformat.Format = '@'
                $Cell.Value = '⚠ ' + $Text
                $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml('#FFCCCC'))
                $Cell.Style.Font.Bold = $true
                $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
            }

            $StyleMatchCell = {
                param($Cell)
                $Cell.Style.Numberformat.Format = '@'
                $Cell.Value = '✓ Alla flikar matchar'
                $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml('#C6EFCE'))
                $Cell.Style.Font.Bold = $true
                $Cell.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#006100'))
            }

            $devPart = $null; $devBatch = $null; $devCart = $null

            if ($wsHeaderCheck -and $wsHeaderCheck.Details) {
                $linesDev = ($wsHeaderCheck.Details -split "`r?`n")
                foreach ($ln in $linesDev) {
                    if ($ln -match '^-\s*PartNo[^:]*:\s*(.+)$')      { $devPart  = $matches[1].Trim() }
                    elseif ($ln -match '^-\s*BatchNo[^:]*:\s*(.+)$') { $devBatch = $matches[1].Trim() }
                    elseif ($ln -match '^-\s*CartridgeNo[^:]*:\s*(.+)$') { $devCart = $matches[1].Trim() }
                }
            }

            if ($wsHeaderCheck) {
                if ($devPart)  { & $StyleMismatchCell $wsInfo.Cells["C$rowPart"]  ('Avvikande: ' + $devPart) }
                else           { & $StyleMatchCell $wsInfo.Cells["C$rowPart"] }
                
                if ($devBatch) { & $StyleMismatchCell $wsInfo.Cells["C$rowBatch"] ('Avvikande: ' + $devBatch) }
                else           { & $StyleMatchCell $wsInfo.Cells["C$rowBatch"] }
                
                if ($devCart)  { & $StyleMismatchCell $wsInfo.Cells["C$rowCart"]  ('Avvikande: ' + $devCart) }
                else           { & $StyleMatchCell $wsInfo.Cells["C$rowCart"] }
            }
        } catch {}

        if ($headerWs) {
            $doc = $headerWs.DocumentNumber
            if ($doc) { $doc = ($doc -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$', '').Trim() }
            if ($headerWs.Attachment -and ($doc -notmatch '(?i)\bAttachment\s+\w+\b')) {
                $doc = "$doc Attachment $($headerWs.Attachment)"
            }
            $wsInfo.Cells["B$rowDoc"].Value = $doc
            $wsInfo.Cells["B$rowRev"].Value = $headerWs.Rev
            $wsInfo.Cells["B$rowEff"].Value = $headerWs.Effective
            
            if ($wsHeaderCheck -and $doc) { & $StyleMatchCell $wsInfo.Cells["C$rowDoc"] }
            if ($wsHeaderCheck -and $headerWs.Rev) { & $StyleMatchCell $wsInfo.Cells["C$rowRev"] }
            if ($wsHeaderCheck -and $headerWs.Effective) { & $StyleMatchCell $wsInfo.Cells["C$rowEff"] }
        } else {
            $wsInfo.Cells["B$rowDoc"].Value = ''
            $wsInfo.Cells["B$rowRev"].Value = ''
            $wsInfo.Cells["B$rowEff"].Value = ''
        }

        if ($selPos) {
            $wsInfo.Cells["B$rowPosFile"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowPosFile"].Value = (Split-Path $selPos -Leaf)
        } else { $wsInfo.Cells["B$rowPosFile"].Value = '' }

        if ($headerPos) {
            $docPos = $headerPos.DocumentNumber
            if ($docPos) { $docPos = ($docPos -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
            $wsInfo.Cells["B$rowPosDoc"].Value = $docPos
            $wsInfo.Cells["B$rowPosRev"].Value = $headerPos.Rev
            $wsInfo.Cells["B$rowPosEff"].Value = $headerPos.Effective

            $posHeaderCheck = $null
            try { $posHeaderCheck = $headerPos.HeaderCheck } catch {}
            $devPosDoc = $null; $devPosRev = $null; $devPosEff = $null

            if ($posHeaderCheck -and $posHeaderCheck.Details) {
                $linesDev = ($posHeaderCheck.Details -split "`r?`n")
                foreach ($ln in $linesDev) {
                    if ($ln -match '^-\s*DocumentNumber[^|]*\|\s*avvikande flikar:\s*(.+)$') { $devPosDoc = $matches[1].Trim() }
                    elseif ($ln -match '^-\s*Rev[^|]*\|\s*avvikande flikar:\s*(.+)$') { $devPosRev = $matches[1].Trim() }
                    elseif ($ln -match '^-\s*Effective[^|]*\|\s*avvikande flikar:\s*(.+)$') { $devPosEff = $matches[1].Trim() }
                }
            }

            if ($posHeaderCheck) {
                if ($docPos) {
                    if ($devPosDoc) { & $StyleMismatchCell $wsInfo.Cells["C$rowPosDoc"] ('Avvikande flikar: ' + $devPosDoc) }
                    else { & $StyleMatchCell $wsInfo.Cells["C$rowPosDoc"] }
                }
                if ($headerPos.Rev) {
                    if ($devPosRev) { & $StyleMismatchCell $wsInfo.Cells["C$rowPosRev"] ('Avvikande flikar: ' + $devPosRev) }
                    else { & $StyleMatchCell $wsInfo.Cells["C$rowPosRev"] }
                }
                if ($headerPos.Effective) {
                    if ($devPosEff) { & $StyleMismatchCell $wsInfo.Cells["C$rowPosEff"] ('Avvikande flikar: ' + $devPosEff) }
                    else { & $StyleMatchCell $wsInfo.Cells["C$rowPosEff"] }
                }
            } else {
            }
        } else {
            $wsInfo.Cells["B$rowPosDoc"].Value = ''
            $wsInfo.Cells["B$rowPosRev"].Value = ''
            $wsInfo.Cells["B$rowPosEff"].Value = ''
        }

        if ($selNeg) {
            $wsInfo.Cells["B$rowNegFile"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowNegFile"].Value = (Split-Path $selNeg -Leaf)
        } else { $wsInfo.Cells["B$rowNegFile"].Value = '' }

        if ($headerNeg) {
            $docNeg = $headerNeg.DocumentNumber
            if ($docNeg) { $docNeg = ($docNeg -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
            $wsInfo.Cells["B$rowNegDoc"].Value = $docNeg
            $wsInfo.Cells["B$rowNegRev"].Value = $headerNeg.Rev
            $wsInfo.Cells["B$rowNegEff"].Value = $headerNeg.Effective
            
            $negHeaderCheck = $null
            try { $negHeaderCheck = $headerNeg.HeaderCheck } catch {}
            $devNegDoc = $null; $devNegRev = $null; $devNegEff = $null

            if ($negHeaderCheck -and $negHeaderCheck.Details) {
                $linesDev = ($negHeaderCheck.Details -split "`r?`n")
                foreach ($ln in $linesDev) {
                    if ($ln -match '^-\s*DocumentNumber[^|]*\|\s*avvikande flikar:\s*(.+)$') { $devNegDoc = $matches[1].Trim() }
                    elseif ($ln -match '^-\s*Rev[^|]*\|\s*avvikande flikar:\s*(.+)$') { $devNegRev = $matches[1].Trim() }
                    elseif ($ln -match '^-\s*Effective[^|]*\|\s*avvikande flikar:\s*(.+)$') { $devNegEff = $matches[1].Trim() }
                }
            }

            if ($negHeaderCheck) {
                if ($docNeg) {
                    if ($devNegDoc) { & $StyleMismatchCell $wsInfo.Cells["C$rowNegDoc"] ('Avvikande flikar: ' + $devNegDoc) }
                    else { & $StyleMatchCell $wsInfo.Cells["C$rowNegDoc"] }
                }
                if ($headerNeg.Rev) {
                    if ($devNegRev) { & $StyleMismatchCell $wsInfo.Cells["C$rowNegRev"] ('Avvikande flikar: ' + $devNegRev) }
                    else { & $StyleMatchCell $wsInfo.Cells["C$rowNegRev"] }
                }
                if ($headerNeg.Effective) {
                    if ($devNegEff) { & $StyleMismatchCell $wsInfo.Cells["C$rowNegEff"] ('Avvikande flikar: ' + $devNegEff) }
                    else { & $StyleMatchCell $wsInfo.Cells["C$rowNegEff"] }
                }
            } else {
            }
        } else {
            $wsInfo.Cells["B$rowNegDoc"].Value = ''
            $wsInfo.Cells["B$rowNegRev"].Value = ''
            $wsInfo.Cells["B$rowNegEff"].Value = ''
        }

    } catch {
        Gui-Log ("⚠️ Header summary fel: " + $_.Exception.Message) 'Warn'
    }

} catch {
    Gui-Log "⚠️ Information-blad fel: $($_.Exception.Message)" 'Warn'
}

# ============================
# === Equipment-blad       ===
# ============================
# OBS: Endast mall kopieras - ingen automatisk ifyllning av pipetter/instrument.
# Användaren fyller i manuellt i output-filen.
if ($Config -and $Config.Contains('EnableEquipmentSheet') -and -not $Config.EnableEquipmentSheet) {
    Gui-Log 'ℹ️ Utrustningslista avstängt. Hoppar över.' 'Info'
} else {
    try {
        if (Test-Path -LiteralPath $UtrustningListPath) {
            $srcPkg = $null
            try {
                $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($UtrustningListPath))

                $srcWs = $srcPkg.Workbook.Worksheets['Sheet1']
                if (-not $srcWs) { $srcWs = $srcPkg.Workbook.Worksheets[1] }

                if ($srcWs) {
                    # Ta bort befintlig flik om den finns
                    $wsEq = $pkgOut.Workbook.Worksheets['Utrustningslista']
                    if ($wsEq) { $pkgOut.Workbook.Worksheets.Delete($wsEq) }

                    # Kopiera mallen som en ny flik
                    $wsEq = $pkgOut.Workbook.Worksheets.Add('Utrustningslista', $srcWs)

                    # Ta bort formler och behåll bara värden
                    if ($wsEq.Dimension) {
                        foreach ($cell in $wsEq.Cells[$wsEq.Dimension.Address]) {
                            if ($cell.Formula -or $cell.FormulaR1C1) {
                                $val = $cell.Value
                                $cell.Formula     = $null
                                $cell.FormulaR1C1 = $null
                                $cell.Value       = $val
                            }
                        }
                        # Kopiera kolumnbredder
                        $colCount = $srcWs.Dimension.End.Column
                        for ($c = 1; $c -le $colCount; $c++) {
                            try { $wsEq.Column($c).Width = $srcWs.Column($c).Width } catch {}
                        }
                    }
                    Gui-Log "✅ Utrustningslista kopierad." 'Info'
                }
            } finally {
                if ($srcPkg) { try { $srcPkg.Dispose() } catch {} }
            }
        } else {
            Gui-Log ("ℹ️ Utrustningslista saknas: $UtrustningListPath") 'Info'
        }
    } catch {
        Gui-Log "⚠️ Kunde inte skapa Utrustningslista-flik: $($_.Exception.Message)" 'Warn'
    }
}

# ============================
# === Control Material     ===
# ============================
try {
    if ($controlTab -and (Test-Path -LiteralPath $RawDataPath)) {
        $srcPkg = $null
        try {
            $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($RawDataPath))
            try { $srcPkg.Workbook.Calculate() } catch {}

            $candidates = if ($controlTab -match '\|') {
                $controlTab -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            } else { @($controlTab) }

            $srcWs = $null
            foreach ($cand in $candidates) {
                $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -eq $cand } | Select-Object -First 1
                if ($srcWs) { break }
                $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -like "*$cand*" } | Select-Object -First 1
                if ($srcWs) { break }
            }

            if ($srcWs) {
                $safeName = if ($srcWs.Name.Length -gt 31) { $srcWs.Name.Substring(0,31) } else { $srcWs.Name }
                $destName = $safeName; $n = 1
                while ($pkgOut.Workbook.Worksheets[$destName]) {
                    $base = if ($safeName.Length -gt 27) { $safeName.Substring(0,27) } else { $safeName }
                    $destName = "$base($n)"; $n++
                }

                $wsCM = $pkgOut.Workbook.Worksheets.Add($destName, $srcWs)
                if ($wsCM.Dimension) {
                    foreach ($cell in $wsCM.Cells[$wsCM.Dimension.Address]) {
                        if ($cell.Formula -or $cell.FormulaR1C1) {
                            $v = $cell.Value
                            $cell.Formula = $null
                            $cell.FormulaR1C1 = $null
                            $cell.Value = $v
                        }
                    }
                    try {
                        if (Get-Command Safe-AutoFitColumns -ErrorAction SilentlyContinue) {
                            Safe-AutoFitColumns -Ws $wsCM -Context 'ControlMaterial'
                        } else {
                            $wsCM.Cells[$wsCM.Dimension.Address].AutoFitColumns() | Out-Null
                        }
                    } catch {}
                }

                Gui-Log "✅ Kontrollmaterial: '$($srcWs.Name)' → '$destName'" 'Info'
            } else {
                Gui-Log "ℹ️ Hittade inget blad i kontrollfilen som matchar '$controlTab'." 'Info'
            }
        } finally {
            if ($srcPkg) { try { $srcPkg.Dispose() } catch {} }
        }
    } else {
        Gui-Log "ℹ️ Ingen Control-flik skapad (saknar mappning eller kontrollfil)." 'Info'
    }
} catch {
    Gui-Log "⚠️ Control Material-fel: $($_.Exception.Message)" 'Warn'
}

# ============================
# === SharePoint Info      ===
# ============================
try {
    $skipSpInfo = -not $global:SpEnabled

    if ($skipSpInfo) {
        Gui-Log "ℹ️ SharePoint Info avstängt i konfigurationen – hoppar över." 'Info'
        try {
            $old = $pkgOut.Workbook.Worksheets["SharePoint Info"]
            if ($old) { $pkgOut.Workbook.Worksheets.Delete($old) }
        } catch {}
    } else {

        $spOk = $false
        if ($global:SpConnected) { $spOk = $true }
        elseif (Get-Command Test-SPClientConnection -ErrorAction SilentlyContinue) {
            try { $spOk = [bool](Test-SPClientConnection) } catch { $spOk = $false }
        }

        if (-not $spOk) {
            # Om manuell anslutning är valt: detta är normalt tills användaren klickar "🔌 Anslut SP".
            if ($global:SpEnabled -and (-not $global:SpAutoConnect)) {
                Gui-Log "ℹ️ SharePoint ej anslutet (manuellt läge). Klicka '🔌 Anslut SP' för att hämta SharePoint-data." 'Info'
            }
            else {
                $errMsg = if ($global:SpError) { $global:SpError } else { 'Okänt fel' }
                Gui-Log ("⚠️ SharePoint ej tillgängligt: $errMsg") 'Warn'
            }
        }

        $batchInfo = Get-BatchLinkInfo -SealPosPath $selPos -SealNegPath $selNeg -Lsp $lspForLinks
        $batch = $batchInfo.Batch

        # Startcell för SharePoint Info-blocket i bladet 'Information' (konfigurerbart)
        $spStartCell = $null
        try { $spStartCell = Get-ConfigValue -Name 'SharePointInfoStartCell' -Default 'E1' } catch { $spStartCell = 'E1' }
        $spRc = Convert-A1ToRowCol -A1 $spStartCell -DefaultRow 1 -DefaultCol 5
        $spStartRow = $spRc.Row
        $spStartCol = $spRc.Col


        if (-not $batch) {
            Gui-Log "ℹ️ Inget Batch # i POS/NEG – skriver SharePoint Info i 'Information'." 'Info'

            # Skriv tomt block (1:1 layout) in i Information (till höger)
            [void](Write-SPBlockIntoInformation -Pkg $pkgOut -Rows @() -Batch '—' -TargetSheetName 'Information' -StartRow $spStartRow -StartCol $spStartCol)

            # Säkerställ att separat flik inte finns kvar (banta rapporten)
            try {
                $old = $pkgOut.Workbook.Worksheets["SharePoint Info"]
                if ($old) { $pkgOut.Workbook.Worksheets.Delete($old) }
            } catch {}
        } else {
            Gui-Log "🔎 Batch hittad: $batch" 'Info'

            $fields = @(
                'Work_x0020_Center','Title','Batch_x0023_','SAP_x0020_Batch_x0023__x0020_2',
                'LSP','Material','BBD_x002f_SLED','Actual_x0020_startdate_x002f__x0',
                'PAL_x0020__x002d__x0020_Sample_x','Sample_x0020_Reagent_x0020_P_x00',
                'Order_x0020_quantity','Total_x0020_good','ITP_x0020_Test_x0020_results',
                'IPT_x0020__x002d__x0020_Testing_0','MES_x0020__x002d__x0020_Order_x0'
            )
            $renameMap = @{
                'Work Center'            = 'Work Center'
                'Title'                  = 'Order#'
                'Batch#'                 = 'SAP Batch#'
                'SAP Batch# 2'           = 'SAP Batch# 2'
                'LSP'                    = 'LSP'
                'Material'               = 'Material'
                'BBD/SLED'               = 'BBD/SLED'
                'Actual startdate/_x0'   = 'ROBAL - Actual start date/time'
                'PAL - Sample_x'         = 'Sample Reagent use'
                'Sample Reagent P'       = 'Sample Reagent P/N'
                'Order quantity'         = 'Order quantity'
                'Total good'             = 'ROBAL - Till Packning'
                'IPT Test results'       = 'IPT Test results'
                'IPT - Testing_0'        = 'IPT - Testing Finalized'
                'MES - Order_x0'         = 'MES Order'
            }

            $desiredOrder = @(
                'Work Center','Order#','SAP Batch#','SAP Batch# 2','LSP','Material','BBD/SLED',
                'ROBAL - Actual start date/time','Sample Reagent use','Sample Reagent P/N',
                'Order quantity','ROBAL - Till Packning','IPT Test results',
                'IPT - Testing Finalized','MES Order'
            )

            $dateFields      = @('BBD/SLED','ROBAL - Actual start date/time','IPT - Testing Finalized')
            $shortDateFields = @('BBD/SLED')

            $rows = @()
            if ($spOk) {
                try {
                    $items = Invoke-SPClient -ScriptBlock {
                        param($flds)
                        Get-PnPListItem -List "Cepheid | Production orders" -Fields $flds -PageSize 2000 -ErrorAction Stop
                    } -Arguments @($fields)
                    $match = $items | Where-Object {
                        $v1 = $_['Batch_x0023_']; $v2 = $_['SAP_x0020_Batch_x0023__x0020_2']
                        $s1 = if ($null -ne $v1) { ([string]$v1).Trim() } else { '' }
                        $s2 = if ($null -ne $v2) { ([string]$v2).Trim() } else { '' }
                        $s1 -eq $batch -or $s2 -eq $batch
                    } | Select-Object -First 1

                    if ($match) {
                        foreach ($f in $fields) {
                            $val = $match[$f]
                            $label = $f -replace '_x0020_', ' ' `
                                         -replace '_x002d_', '-' `
                                         -replace '_x0023_', '#' `
                                         -replace '_x002f_', '/' `
                                         -replace '_x2013_', '–' `
                                         -replace '_x00',''
                            $label = $label.Trim()
                            if ($renameMap.ContainsKey($label)) { $label = $renameMap[$label] }

                            if ($null -ne $val -and $val -ne '') {
                                if ($val -eq $true) { $val = 'JA' }
                                elseif ($val -eq $false) { $val = 'NEJ' }

                                $dt = $null
                                if ($val -is [datetime]) { $dt = [datetime]$val }
                                else { try { $dt = [datetime]::Parse($val) } catch { $dt = $null } }

                                if ($dt -ne $null -and ($dateFields -contains $label)) {
                                    $fmt = if ($shortDateFields -contains $label) { 'yyyy-MM-dd' } else { 'yyyy-MM-dd HH:mm' }
                                    $val = $dt.ToString($fmt)
                                }

                                $rows += [pscustomobject]@{ Rubrik = $label; 'Värde' = $val }
                            }
                        }

                        if ($rows.Count -gt 0) {
                            $ordered = @()
                            foreach ($label in $desiredOrder) {
                                $hit = $rows | Where-Object { $_.Rubrik -eq $label } | Select-Object -First 1
                                if ($hit) { $ordered += $hit }
                            }
                            if ($ordered.Count -gt 0) { $rows = $ordered }
                        }

                        Gui-Log "📄 SharePoint-post hittad – skriver blad." 'Info'
                    } else {
                        Gui-Log "ℹ️ Ingen post i SharePoint för Batch=$batch." 'Info'
                    }
                } catch {
                    Gui-Log "⚠️ SP: Get-PnPListItem misslyckades: $($_.Exception.Message)" 'Warn'
                }
            }

            [void](Write-SPBlockIntoInformation -Pkg $pkgOut -Rows $rows -Batch $batch -TargetSheetName 'Information' -StartRow $spStartRow -StartCol $spStartCol)

            try {
                $old = $pkgOut.Workbook.Worksheets["SharePoint Info"]
                if ($old) { $pkgOut.Workbook.Worksheets.Delete($old) }
            } catch {}

            try {
                if ($slBatchLink -and $batch) {
                    $slBatchLink.Text = "SharePoint: $batch"
                    $slBatchLink.Tag  = $batchInfo.Url
                    $slBatchLink.Enabled = $true
                }
            } catch {}

            try {
                $wsSP = $pkgOut.Workbook.Worksheets['Information']
                if ($wsSP) {
                    $labelCol = $spStartCol
                $valueCol = $spStartCol + 1
                for ($r = ($spStartRow + 1); $r -le ($spStartRow + 120); $r++) {
                        if ((Normalize-HeaderText (($wsSP.Cells[$r,$labelCol].Text + '')).Trim()) -ieq 'Sample Reagent use') {
                            $wsSP.Cells[$r,$valueCol].Style.WrapText = $true
                            $wsSP.Cells[$r,$valueCol].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
                            $wsSP.Row($r).CustomHeight = $true
                            break
                        }
                    }
                }
            } catch {
                Gui-Log "⚠️ WrapText på 'Sample Reagent use' misslyckades: $($_.Exception.Message)" 'Warn'
            }

            # ============================================================
            # === Sample Reagent Checklist-highlight (additiv)         ===
            # ============================================================
            try {
                $_srhPNs = @()
                try { $_srhPNs = @(Get-ConfigValue -Name 'SampleReagentChecklistPNs' -Default @()) } catch {}

                if ($_srhPNs.Count -gt 0) {
                    $_srhWs = $pkgOut.Workbook.Worksheets['Information']
                    if ($_srhWs -and $_srhWs.Dimension) {
                        $_srhLabelCol = $spStartCol
                        $_srhValueCol = $spStartCol + 1
                        $_srhPNRow  = -1; $_srhUseRow = -1
                        $_srhPNVal  = ''; $_srhUseVal = ''

                        for ($_sr = ($spStartRow + 1); $_sr -le ($spStartRow + 120); $_sr++) {
                            $_srhLabel = (Normalize-HeaderText ($_srhWs.Cells[$_sr, $_srhLabelCol].Text + '')).Trim()
                            if ($_srhLabel -ieq 'Sample Reagent P/N') {
                                $_srhPNRow = $_sr
                                $_srhPNVal = ($_srhWs.Cells[$_sr, $_srhValueCol].Text + '').Trim()
                            }
                            if ($_srhLabel -ieq 'Sample Reagent use') {
                                $_srhUseRow = $_sr
                                $_srhUseVal = ($_srhWs.Cells[$_sr, $_srhValueCol].Text + '').Trim()
                            }
                        }

                        $_srhMatch = $false
                        if ($_srhPNVal) {
                            foreach ($_pn in $_srhPNs) {
                                if ($_srhPNVal -eq ($_pn + '').Trim()) { $_srhMatch = $true; break }
                            }
                        }

                        if ($_srhMatch) {
                            $_srhHighlightBg = [System.Drawing.Color]::FromArgb(255, 255, 200)
                            $_srhWarningBg   = [System.Drawing.Color]::FromArgb(255, 235, 156)
                            $_srhWarningFg   = [System.Drawing.Color]::FromArgb(156, 101, 0)
                            $_srhOkBg        = [System.Drawing.Color]::FromArgb(198, 239, 206)
                            $_srhOkFg        = [System.Drawing.Color]::FromArgb(0, 97, 0)

                            if ($_srhPNRow -gt 0) {
                                $_srhWs.Cells[$_srhPNRow, $_srhValueCol].Style.Fill.PatternType = 'Solid'
                                $_srhWs.Cells[$_srhPNRow, $_srhValueCol].Style.Fill.BackgroundColor.SetColor($_srhHighlightBg)
                                $_srhWs.Cells[$_srhPNRow, $_srhValueCol].Style.Font.Bold = $true
                            }

                            if ($_srhUseRow -gt 0) {
                                if ($_srhUseVal) {
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Style.Fill.PatternType = 'Solid'
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Style.Fill.BackgroundColor.SetColor($_srhOkBg)
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Style.Font.Color.SetColor($_srhOkFg)
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Style.Font.Bold = $true
                                } else {
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Value = '⚠ SAKNAS'
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Style.Fill.PatternType = 'Solid'
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Style.Fill.BackgroundColor.SetColor($_srhWarningBg)
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Style.Font.Color.SetColor($_srhWarningFg)
                                    $_srhWs.Cells[$_srhUseRow, $_srhValueCol].Style.Font.Bold = $true
                                    Gui-Log '⚠️ Sample Reagent Saknas' 'Warn'
                                }
                            }
                        }
                    }
                }
            } catch {
                Gui-Log "⚠️ Sample Reagent highlight: $($_.Exception.Message)" 'Warn'
            }

        }
    }
} catch {
    Gui-Log "⚠️ SP-blad: $($_.Exception.Message)" 'Warn'
}

# ============================
# === QC Reminder (HIV/HBV/HCV) → Information E16:F18 ===
# ============================
try {
    if ($script:QcReminderB3) {
        $wsInfo = $pkgOut.Workbook.Worksheets['Information']
        if ($wsInfo) {
            $qcRow = 16
            $eCol  = 5   # E
            $fCol  = 6   # F

            # Färger (samma som SharePoint Info Header)
            $HeaderBg  = [System.Drawing.Color]::FromArgb(68, 84, 106)
            $HeaderFg  = [System.Drawing.Color]::White
            $SectionBg = [System.Drawing.Color]::FromArgb(217, 225, 242)
            $SectionFg = [System.Drawing.Color]::FromArgb(0, 32, 96)
            $BorderClr = [System.Drawing.Color]::FromArgb(68, 84, 106)

            # --- Rad 16: Header (merged E16:F16) ---
            $wsInfo.Cells[$qcRow, $eCol, $qcRow, $fCol].Merge = $true
            $hdrCell = $wsInfo.Cells[$qcRow, $eCol]
            $hdrCell.Value = ('QC Reminder – ' + $script:QcReminderB3)
            $hdrCell.Style.Font.Bold = $true
            $hdrCell.Style.Font.Size = 14
            $hdrCell.Style.Font.Name = 'Calibri'
            $hdrCell.Style.Font.Color.SetColor($HeaderFg)
            $hdrCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $hdrCell.Style.Fill.BackgroundColor.SetColor($HeaderBg)
            $hdrCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
            $hdrCell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
            $wsInfo.Row($qcRow).Height = 22

            # --- Rad 17: "Additional QC-Data?" / "JA" ---
            $wsInfo.Cells[($qcRow+1), $eCol].Value = 'Additional QC-Data?'
            $wsInfo.Cells[($qcRow+1), $eCol].Style.Font.Bold = $true
            $wsInfo.Cells[($qcRow+1), $eCol].Style.Font.Name = 'Calibri'
            $wsInfo.Cells[($qcRow+1), $eCol].Style.Font.Size = 10
            $wsInfo.Cells[($qcRow+1), $eCol].Style.Font.Color.SetColor($SectionFg)
            $wsInfo.Cells[($qcRow+1), $eCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $wsInfo.Cells[($qcRow+1), $eCol].Style.Fill.BackgroundColor.SetColor($SectionBg)

            $wsInfo.Cells[($qcRow+1), $fCol].Value = 'JA'
            $wsInfo.Cells[($qcRow+1), $fCol].Style.Font.Bold = $true
            $wsInfo.Cells[($qcRow+1), $fCol].Style.Font.Name = 'Calibri'
            $wsInfo.Cells[($qcRow+1), $fCol].Style.Font.Size = 10
            $wsInfo.Cells[($qcRow+1), $fCol].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))
            $wsInfo.Cells[($qcRow+1), $fCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $wsInfo.Cells[($qcRow+1), $fCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::White)
            $wsInfo.Cells[($qcRow+1), $fCol].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left

            # --- Rad 18: "Lathund QC data" med länk (merged E18:F18) ---
            $wsInfo.Cells[($qcRow+2), $eCol, ($qcRow+2), $fCol].Merge = $true
            $linkCell = $wsInfo.Cells[($qcRow+2), $eCol]

            $lathundPath = ''
            if ($Config) {
                # Primär nyckel (finns i Config.ps1 i LIVE-zippen)
                if ($Config.Contains('QcDataCheatSheetLink')) {
                    $lathundPath = $Config.QcDataCheatSheetLink
                }
                # Fallback om du någon gång vill använda ett annat namn
                elseif ($Config.Contains('QcReminderLathundPath')) {
                    $lathundPath = $Config.QcReminderLathundPath
                }
            }

            $linkCell.Value = 'Lathund QC data'
            $linkCell.Style.Font.Name = 'Calibri'
            $linkCell.Style.Font.Size = 10
            $linkCell.Style.Font.UnderLine = $true
            $linkCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(5, 99, 193))
            $linkCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $linkCell.Style.Fill.BackgroundColor.SetColor($SectionBg)
            # Centera merged E18:F18 (både horisontellt & vertikalt)
            $linkCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
            $linkCell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

            if ($lathundPath) {
                try {
                    # Bygg stabil URI:
                    $uriText = $lathundPath
                    if ($uriText -match '^(?i)https?://') {
                        $linkCell.Hyperlink = New-Object System.Uri($uriText)
                   }
                    elseif ($uriText -match '^[A-Za-z]:\\') {
                        # ex: N:\Folder\File.docx -> file:///N:/Folder/File.docx
                        $fileUri = 'file:///' + ($uriText -replace '\\','/')
                        $fileUri = [System.Uri]::EscapeUriString($fileUri)
                        $linkCell.Hyperlink = New-Object System.Uri($fileUri)
                    }
                    elseif ($uriText -match '^\\\\') {
                        # ex: \\server\share\dir\file.docx -> file://server/share/dir/file.docx
                        $fileUri = 'file://' + (($uriText -replace '^\\\\','') -replace '\\','/')
                        $fileUri = [System.Uri]::EscapeUriString($fileUri)
                        $linkCell.Hyperlink = New-Object System.Uri($fileUri)
                    }
                    else {
                        # Sista försök (om du ger en redan korrekt URI)
                        $linkCell.Hyperlink = New-Object System.Uri($uriText)
                    }
                 } catch {
                     Gui-Log ("⚠️ QC Reminder: Kunde inte skapa hyperlänk till '{0}': {1}" -f $lathundPath, $_.Exception.Message) 'Warn'
                 }
             }

            # Borders kring hela QC Reminder-blocket (E16:F18)
            $rng = $wsInfo.Cells[$qcRow, $eCol, ($qcRow+2), $fCol]
            $rng.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $rng.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $rng.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $rng.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $rng.Style.Border.Top.Color.SetColor($BorderClr)
            $rng.Style.Border.Bottom.Color.SetColor($BorderClr)
            $rng.Style.Border.Left.Color.SetColor($BorderClr)
            $rng.Style.Border.Right.Color.SetColor($BorderClr)

            Gui-Log ("✅ QC Reminder skrivet i Information E{0}:F{1} (B3={2})" -f $qcRow, ($qcRow+2), $script:QcReminderB3) 'Info'
        }
    }
} catch {
    Gui-Log ("⚠️ QC Reminder (skrivning): " + $_.Exception.Message) 'Warn'
}

# ============================
# === Header watermark     ===
# ============================
try {
    foreach ($ws in $pkgOut.Workbook.Worksheets) {
        try {
            $ws.HeaderFooter.OddHeader.CenteredText   = '&"Arial,Bold"&14 UNCONTROLLED'
            $ws.HeaderFooter.EvenHeader.CenteredText  = '&"Arial,Bold"&14 UNCONTROLLED'
            $ws.HeaderFooter.FirstHeader.CenteredText = '&"Arial,Bold"&14 UNCONTROLLED'
        } catch {
            Gui-Log ("⚠️ Kunde inte sätta header på blad: " + $ws.Name) 'Warn'
        }
    }
} catch {
    Gui-Log "⚠️ Fel vid vattenstämpling av rapporten." 'Warn'
}

        # ============================
        # === Tab-färger (innan Save)
        # ============================
        try {
            $wsT = $pkgOut.Workbook.Worksheets['Information'];     if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 52, 152, 219) }
            $wsT = $pkgOut.Workbook.Worksheets['Utrustningslista'];     if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 33, 115, 70) }
        } catch {
        Gui-Log "⚠️ Kunde inte sätta tab-färg: $($_.Exception.Message)" 'Warn'
    }

        # ============================
        # === Spara & Audit        ===
        # ============================

        $nowTs    = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseName = "$($env:USERNAME)_output_${lsp}_$nowTs.xlsx"# (GUI-val borttaget) Spara alltid i temporär katalog
        $saveDir  = $env:TEMP
        $SavePath = Join-Path $saveDir $baseName
        Gui-Log "💾 Sparläge: Temporärt → $SavePath"
        try {
            $pkgOut.Workbook.View.ActiveTab = 0
            $wsInitial = $pkgOut.Workbook.Worksheets["Information"]
            if ($wsInitial) { $wsInitial.View.TabSelected = $true }

            # ============================
            # === RuleEngine Debug Sheet ==
            # ============================
            try {
                if ((Get-ConfigFlag -Name 'EnableRuleEngine' -Default $false -ConfigOverride $Config) -and
                    (Get-ConfigFlag -Name 'EnableRuleEngineDebugSheet' -Default $false -ConfigOverride $Config) -and
                    $selCsv -and (Test-Path -LiteralPath $selCsv)) {

                    if (-not $script:RuleEngineShadow -or -not $script:RuleEngineShadow.Rows -or $script:RuleEngineShadow.Rows.Count -eq 0) {

                        Gui-Log "🧠 Regelmotor: saknas vid Save → bygger nu..." 'Warn'

                        $rb2 = $script:RuleBankCache
                        if (-not $rb2) {
                            $rb2 = Load-RuleBank -RuleBankDir $Config.RuleBankDir
                            try { $rb2 = Compile-RuleBank -RuleBank $rb2 } catch {}
                        }

                        $csvObjs2 = $script:RuleEngineCsvObjs
                        if (-not $csvObjs2 -or $csvObjs2.Count -eq 0) {
                            $csvObjs2 = @()
                            if ($csvRows -and $csvRows.Count -gt 0) {
                                $csvObjs2 = @($csvRows)
                            } else {
                                try {
                                    $all = Get-Content -LiteralPath $selCsv
                                    if ($all -and $all.Count -gt 9) {
                                        $hdr = ConvertTo-CsvFields $all[7]
                                        $dl  = $all[9..($all.Count-1)] | Where-Object { $_ -and $_.Trim() }
                                        $del = Get-CsvDelimiter -Path $selCsv
                                        $csvObjs2 = @(ConvertFrom-Csv -InputObject ($dl -join "`n") -Delimiter $del -Header $hdr)
                                    }
                                } catch { $csvObjs2 = @() }
                            }
                        }

                        if ($csvObjs2 -and $csvObjs2.Count -gt 0) {
                            $script:RuleEngineShadow = Invoke-RuleEngine -CsvObjects $csvObjs2 -RuleBank $rb2 -CsvPath $selCsv
                        } else {
                            Gui-Log "⚠️ Regelmotor: kunde inte bygga vid Save (0 rader)." 'Warn'
                        }
                    }

                    if ($script:RuleEngineShadow -and $script:RuleEngineShadow.Rows -and $script:RuleEngineShadow.Rows.Count -gt 0) {
                        Gui-Log "🧠 Skriver CSV-Sammanfattning..." 'Info'
                        $includeAll = Get-ConfigFlag -Name 'RuleEngineDebugIncludeAllRows' -Default $false -ConfigOverride $Config
                        [void](Write-RuleEngineDebugSheet -Pkg $pkgOut -RuleEngineResult $script:RuleEngineShadow -IncludeAllRows $includeAll)
                    } else {
                        Gui-Log "⚠️ Kunde inte skriva CSV-Sammanfattning." 'Warn'
                    }
                }
            } catch {
                Gui-Log ("⚠️ Kunde inte skriva CSV-Sammanfattning: " + $_.Exception.Message) 'Warn'
            }

            Set-UiStep 90 'Sparar rapport…'
            $pkgOut.SaveAs($SavePath)
            Set-UiStep 100 'Klar ✅'
            Gui-Log -Text ("✅ Rapport sparad: {0}" -f $SavePath) -Severity Info -Category RESULT
            $global:LastReportPath = $SavePath

            try {
                $auditDir = Join-Path $PSScriptRoot 'audit'
                if (-not (Test-Path $auditDir)) { New-Item -ItemType Directory -Path $auditDir -Force | Out-Null }

                $auditObj = [pscustomobject]@{
                    DatumTid        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    Användare       = $env:USERNAME
                    LSP             = $lsp
                    ValdCSV         = if ($selCsv) { Split-Path $selCsv -Leaf } else { '' }
                    ValdSealNEG     = Split-Path $selNeg -Leaf
                    ValdSealPOS     = Split-Path $selPos -Leaf
                    SignaturSkriven = if ($chkWriteSign.Checked) { 'Ja' } else { 'Nej' }
                    OverwroteSign   = if ($chkOverwriteSign.Checked) { 'Ja' } else { 'Nej' }
                    SigMismatch     = if ($sigMismatch) { 'Ja' } else { 'Nej' }
                    MismatchSheets  = if ($mismatchSheets -and $mismatchSheets.Count -gt 0) { ($mismatchSheets -join ';') } else { '' }
                    ViolationsNEG   = $violationsNeg.Count
                    ViolationsPOS   = $violationsPos.Count
                    Violations      = ($violationsNeg.Count + $violationsPos.Count)
                    Sparläge        = 'Temporärt'
                    OutputFile      = $SavePath
                    Kommentar       = 'UNCONTROLLED rapport, ingen källfil ändrades automatiskt.'
                    ScriptVersion   = $ScriptVersion
                }

                $auditFile = Join-Path $auditDir ("$($env:USERNAME)_audit_${nowTs}.csv")
                $auditObj | Export-Csv -Path $auditFile -NoTypeInformation -Encoding UTF8

                try {
                    $statusText = 'OK'
                    if (($violationsNeg.Count + $violationsPos.Count) -gt 0 -or $sigMismatch -or ($mismatchSheets -and $mismatchSheets.Count -gt 0)) {
                        $statusText = 'Warnings'
                    }
                    $auditTests = $null
                    try { if ($csvStats) { $auditTests = $csvStats.TestCount } } catch {}
                    Add-AuditEntry -Lsp $lsp -Assay $runAssay -BatchNumber $batch -TestCount $auditTests -Status $statusText -ReportPath $SavePath
                } catch {
                    Gui-Log "⚠️ Kunde inte skriva audit-CSV: $($_.Exception.Message)" 'Warn'
                }
            } catch {
                Gui-Log "⚠️ Kunde inte skriva revisionsfil: $($_.Exception.Message)" 'Warn'
            }

            try { Start-Process -FilePath "excel.exe" -ArgumentList "`"$SavePath`"" } catch {}
        }
        catch {
            Gui-Log "⚠️ Kunde inte spara/öppna: $($_.Exception.Message)" 'Warn'
        }

    } finally {
        try { if ($pkgNeg) { $pkgNeg.Dispose() } } catch {}
        try { if ($pkgPos) { $pkgPos.Dispose() } } catch {}
        try { if ($pkgOut) { $pkgOut.Dispose() } } catch {}
        Set-UiBusy -Busy $false
        $script:BuildInProgress = $false
    }
})

#endregion Event Handlers

# === Tooltip-inställningar ===
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 8000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay  = 500
$toolTip.ShowAlways   = $true
$toolTip.SetToolTip($txtLSP, 'Ange LSP-numret utan "#" och klicka på Sök filer.')
$toolTip.SetToolTip($btnScan, 'Sök efter LSP och lista tillgängliga filer.')
$toolTip.SetToolTip($clbCsv,  'Välj CSV-fil.')
$toolTip.SetToolTip($clbNeg,  'Välj Seal Test Neg-fil.')
$toolTip.SetToolTip($clbPos,  'Välj Seal Test Pos-fil.')
$toolTip.SetToolTip($btnCsvBrowse, 'Bläddra efter en CSV-fil manuellt.')
$toolTip.SetToolTip($btnNegBrowse, 'Bläddra efter Seal Test Neg-fil manuellt.')
$toolTip.SetToolTip($btnPosBrowse, 'Bläddra efter Seal Test Pos-fil manuellt.')
$toolTip.SetToolTip($txtSigner, 'Skriv fullständigt namn, signatur och datum (separerat med kommatecken).')
$toolTip.SetToolTip($chkWriteSign, 'Signatur appliceras på flikar.')
$toolTip.SetToolTip($chkOverwriteSign, 'Dubbelkontroll för att aktivera signering')
$miToggleSign.ToolTipText = 'Visa eller dölj panelen för att lägga till signatur.'
if ($rbSaveInLsp) { $toolTip.SetToolTip($rbSaveInLsp, 'Spara rapporten i mappen för ditt LSP.') }
if ($rbTempOnly) { $toolTip.SetToolTip($rbTempOnly, 'Skapa rapporten temporär utan att spara.') }
$toolTip.SetToolTip($btnBuild, 'Skapa och öppna rapporten baserat på de valda filerna.')
if ($chkSharePointInfo) { $toolTip.SetToolTip($chkSharePointInfo, 'Exportera med SharePoint Info.') }
$txtLSP.add_TextChanged({ Update-BatchLink })

#region Main Run / Orchestration
# =============== SLUT ===============
function Enable-DoubleBuffer {
    $pi = [Windows.Forms.Control].GetProperty('DoubleBuffered',[Reflection.BindingFlags]'NonPublic,Instance')
    foreach($c in @($content,$pLog,$grpPick,$grpSign,$grpSave)) { if ($c) { $pi.SetValue($c,$true,$null) } }
}
try { Set-Theme 'light' } catch {}
Enable-DoubleBuffer
Update-BatchLink
[System.Windows.Forms.Application]::Run($form)

try{ Stop-Transcript | Out-Null }catch{}
#endregion Main Run / Orchestration