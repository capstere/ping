param(
    [string]$ScriptRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

$script:ScriptRoot = $ScriptRoot
$script:DefaultIptRoot = '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT'.TrimEnd('\')

function Get-EnvNonEmpty {
    param([Parameter(Mandatory=$true)][string]$Name)
    try {
        $v = [Environment]::GetEnvironmentVariable($Name)
    } catch {
        $v = $null
    }
    if ([string]::IsNullOrWhiteSpace($v)) { return '' }
    return $v.Trim()
}

$envIptRaw = Get-EnvNonEmpty -Name 'IPT_ROOT'
$candidate = if ($envIptRaw) { $envIptRaw.TrimEnd('\') } else { '' }

$script:CandidateExists = $false
if ($candidate) {
    try { $script:CandidateExists = (Test-Path -LiteralPath $candidate) } catch { $script:CandidateExists = $false }
}

$script:IPTRoot = if ($candidate -and $script:CandidateExists) { $candidate } else { $script:DefaultIptRoot }

$global:IPT_ROOT_EFFECTIVE = $script:IPTRoot
$global:IPT_ROOT_SOURCE = if ($envIptRaw) {
    if ($candidate -and $script:CandidateExists) { 'ENV' } else { 'ENV_INVALID_FALLBACK' }
} else {
    'DEFAULT'
}

function Resolve-IptPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }

    if ($Path.StartsWith($script:DefaultIptRoot + '\', [System.StringComparison]::OrdinalIgnoreCase) -or
    $Path.Equals($script:DefaultIptRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ($script:IPTRoot + $Path.Substring($script:DefaultIptRoot.Length))
    }
    return $Path
}

function Resolve-IptPathList {
    param([object[]]$Paths)
    if ($null -eq $Paths) { return @() }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Paths) {
        if ($null -eq $p) { continue }
        $s = [string]$p
        if ([string]::IsNullOrWhiteSpace($s)) { continue }
        $out.Add((Resolve-IptPath $s))
    }
    return $out.ToArray()
}

function Test-IsNetworkPathSimple {
    param([Parameter(Mandatory=$true)][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if ($Path -like '\\*') { return $true }
    try {
        $root = [System.IO.Path]::GetPathRoot($Path)
        if (-not $root) { return $false }
        $driveName = $root.TrimEnd('\')
        $di = New-Object System.IO.DriveInfo($driveName)
        return ($di.DriveType -eq [System.IO.DriveType]::Network)
    } catch {
        return $false
    }
}

function Resolve-LocalFirstFile {
    param(
        [Parameter(Mandatory=$true)][string]$LocalRelativePath,
        [Parameter(Mandatory=$true)][string]$NetworkPath
    )
    try {
        $local = Join-Path $script:ScriptRoot $LocalRelativePath
        if (Test-Path -LiteralPath $local) { return $local }
    } catch { }
    return (Resolve-IptPath $NetworkPath)
}

 # ============================
 # === Script Version       ===
 # ============================
# Default/fallback (used if Version.txt is missing)
$ScriptVersion = "v95.5"

$RootPaths = @(
    (Join-Path $script:DefaultIptRoot 'Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Tests'),
    (Join-Path $script:DefaultIptRoot '3. IPT - KLART FÖR SAMMANSTÄLLNING'),
    (Join-Path $script:DefaultIptRoot '4. IPT - KLART FÖR GRANSKNING')
)
$RootPaths = Resolve-IptPathList $RootPaths

$ikonSokvag = Join-Path $ScriptRoot "icon.png"

# Redigeringsfil (för genvägar / användare)
$equipXlsCandidate = Resolve-IptPath (Join-Path $script:DefaultIptRoot 'Utrustning, kontrollprover & förbrukningsmaterial IPT\Utrustningslista In Process Testing.xls')
$UtrustningListXlsPath = if (Test-IsNetworkPathSimple -Path $equipXlsCandidate) { $equipXlsCandidate } else { (Join-Path $script:DefaultIptRoot 'Utrustning, kontrollprover & förbrukningsmaterial IPT\Utrustningslista In Process Testing.xls') }

# Läsfil för IPTCompile (synkad kopia)
$equipXlsxCandidate = Resolve-IptPath (Join-Path $script:DefaultIptRoot 'Skiftspecifika dokument\PQC analyst\JESPER\Scripts\zz_IPTCompile_Shortcut\Utrustningslista In Process Testing kopia.xlsx')
$UtrustningListPath = if (Test-IsNetworkPathSimple -Path $equipXlsxCandidate) { $equipXlsxCandidate } else { (Join-Path $script:DefaultIptRoot 'Skiftspecifika dokument\PQC analyst\JESPER\Scripts\zz_IPTCompile_Shortcut\Utrustningslista In Process Testing kopia.xlsx') }

$rawCandidate = Resolve-IptPath (Join-Path $script:DefaultIptRoot 'KONTROLLPROVSFIL - Version 2.5.xlsm')
$RawDataPath  = if (Test-IsNetworkPathSimple -Path $rawCandidate) { $rawCandidate } else { (Join-Path $script:DefaultIptRoot 'KONTROLLPROVSFIL - Version 2.5.xlsm') }
$OtherScriptPath = 'TBD'

$Script1Path = (Join-Path $script:DefaultIptRoot '8. IPT - WR + Rework\1. PQC - Kontrollprovsfil - RÖR EJ -\Script Raw Data\Kontrollprovsfil_EPPlus_2025_ver3.ps1')
$Script2Path = Resolve-IptPath (Join-Path $script:DefaultIptRoot '8. IPT - WR + Rework\1. PQC - Kontrollprovsfil - RÖR EJ -\Script Raw Data\aktivera_makro.ps1')
$Script3Path = Resolve-IptPath (Join-Path $script:DefaultIptRoot 'Skiftspecifika dokument\PQC analyst\JESPER\Scripts\zz_IPTCompile_Shortcut\rename-GUI.bat')
$Script4Path = Resolve-IptPath (Join-Path $script:DefaultIptRoot 'Skiftspecifika dokument\PQC analyst\JESPER\Scripts\zz_IPTCompile_Shortcut\Run_ScriptControl.bat')
$Script5Path = Resolve-IptPath (Join-Path $script:DefaultIptRoot 'Skiftspecifika dokument\PQC analyst\JESPER\Scripts\zz_IPTCompile_Shortcut\AutoMappscript Dashboard.ps1')

$env:PNPPOWERSHELL_UPDATECHECK = "Off"
$global:SP_ClientId   = "O"
$global:SP_Tenant     = "danaher.onmicrosoft.com"
$global:SP_CertBase64 = "O"
$global:SP_SiteUrl    = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management"

$Config = [ordered]@{
    AppName               = 'IPTCompile'
    EntryPointFolderName  = 'zz_IPTCompile'
    HelpFeedbackDir = (Resolve-IptPath (Join-Path $script:DefaultIptRoot 'Skiftspecifika dokument\PQC analyst\JESPER\zz_IPTCompile\help'))
    QcDataCheatSheetLink = (Resolve-IptPath (Join-Path $script:DefaultIptRoot 'IPT - LATHUND\2. Lathund - sammanställning och granskning\Lathund för additional QC data.docx'))

    EnableEquipmentSheet          = $true
    EnableEquipmentScan           = $true
    EnableRuleEngine              = $true
    EnableShadowCompare           = $true
    EnableRuleEngineSummaryLog    = $true
    EnableRuleEngineDebugSheet    = $true
    RuleEngineDebugIncludeAllRows = $false
    EnableRuleEngineRowSkipTrace  = $false
    EnableLocalStaging            = $true
    TempSnapshotRoot              = 'C:\IPTCompile_TEMP'

    GuiLogVerbosity                  = 'QUIET'
    GuiLogInfoCategoriesQuiet        = @('SUMMARY','RESULT','USER','UI')
    GuiLogInfoHiddenCategoriesNormal = @('DEBUG','SANITY','PROGRESS','RuleEngineStats','RuleEngineDev')

    EpplusAutoFitMode    = 'SMART'
    EpplusAutoFitMaxRows = 500

    RuleBankDir = (Join-Path $ScriptRoot 'RuleBank')
    RuleBankRequireCompiled = $true

    CsvStreamingThresholdMB = 25
    CsvPath        = ''
    SealNegPath    = ''
    SealPosPath    = ''
    WorksheetPath  = ''

    # --- Post-download organizer (Camstar -> Downloads -> N:) ---
    # Om tomt används standard för användarens "Downloads".
    DownloadsDir = ''

    # Rotmappar där LSP-mapparna ligger (per status). Används av "Flytta nedladdade filer"-knapparna.
    Stage3RootPath = (Resolve-IptPath (Join-Path $script:DefaultIptRoot 'Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Tests'))
    Stage4RootPath = (Resolve-IptPath (Join-Path $script:DefaultIptRoot '4. IPT - KLART FÖR GRANSKNING'))

    EnableSharePoint = $true
    # Styr om PnP-modulen laddas och Connect-PnPOnline körs vid uppstart.
    # Sätt $false för att hoppa över inloggning (SP Info skrivs ej, men rapporten genereras snabbare).
    # EnableSharePoint måste vara $true för att denna ska ha effekt.
    EnableSharePointAutoConnect = $true
    SiteUrl      = $global:SP_SiteUrl
    Tenant       = $global:SP_Tenant
    ClientId     = $global:SP_ClientId
    Certificate  = $global:SP_CertBase64

    EpplusDllPath = (Join-Path $ScriptRoot 'Modules\EPPlus.dll')
    EpplusVersion = '4.5.3.3'
    AllowNuGetDownload = $false
    
    NetworkOnlyFileNames = @('KONTROLLPROVSFIL - Version 2.5.xlsm')

    EquipmentXmlPath = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript\equipment.xml'

    # P/N som gul-highlightas i SharePoint Info-blocket (Sample Reagent P/N + Sample Reagent use)
    SampleReagentChecklistPNs = @(
        '700-6052','700-6609','700-8870','700-6822',
        '700-5280','700-5197','700-6787','700-5375',
        '700-4521','700-4383','700-5194','700-5666',
        '700-5667','700-5196','700-5662','700-6379'
    )

    # Startcell (A1) för SharePoint Info-blocket som skrivs in i bladet 'Information'
    SharePointInfoStartCell = 'E1'

    ShortcutGroups = [ordered]@{
    '🗂️ IPT-mappar' = @(
        @{ Text='📂 IPT - PÅGÅENDE KÖRNINGAR';        Target=(Resolve-IptPath (Join-Path $script:DefaultIptRoot '2. IPT - PÅGÅENDE KÖRNINGAR')) },
        @{ Text='📂 IPT - KLART FÖR SAMMANSTÄLLNING'; Target=(Resolve-IptPath (Join-Path $script:DefaultIptRoot '3. IPT - KLART FÖR SAMMANSTÄLLNING')) },
        @{ Text='📂 IPT - KLART FÖR GRANSKNING';      Target=(Resolve-IptPath (Join-Path $script:DefaultIptRoot '4. IPT - KLART FÖR GRANSKNING')) },
        @{ Text='📂 SPT Macro Assay';                 Target='N:\QC\QC-0\SPT\SPT macros\Assay' }
    )
    '📄 Dokument' = @(
        # Öppna redigeringsfilen (.xls) för användare
        @{ Text='🧰 Utrustningslista';    Target=$UtrustningListXlsPath },
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

try { $global:Config = $Config } catch {}

function Get-ConfigValue {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        $Default = $null,
        [object]$ConfigOverride
    )

    $cfg = $ConfigOverride
    if (-not $cfg) {
        try { if ($global:Config) { $cfg = $global:Config } } catch {}
    }
    if (-not $cfg) {
        try { $cfg = $Config } catch {}
    }

    if (-not $cfg) { return $Default }

    try {
        if ($cfg -is [hashtable]) {
            if ($cfg.ContainsKey($Name)) { return $cfg[$Name] }
            return $Default
        }
        if ($cfg -is [System.Collections.IDictionary]) {
            if ($cfg.Contains($Name)) { return $cfg[$Name] }
            foreach ($k in $cfg.Keys) {
                if (("$k") -eq $Name) { return $cfg[$k] }
            }
            return $Default
        }
    } catch { return $Default }

    try {
        $prop = $cfg.PSObject.Properties[$Name]
        if ($prop) { return $prop.Value }
    } catch {}

    return $Default
}

function Get-ConfigFlag {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [bool]$Default = $false,
        [object]$ConfigOverride
    )
    try {
        return [bool](Get-ConfigValue -Name $Name -Default $Default -ConfigOverride $ConfigOverride)
    } catch {
        return $Default
    }
}

$script:GXINF_Map = @{
    'Infinity-VI'   = '847922'
    'Infinity-VIII' = '803094'
    'GX5'           = '750210,750211,750212,750213'
    'GX6'           = '750246,750247,750248,750249'
    'GX1'           = '709863,709864,709865,709866'
    'GX2'           = '709951,709952,709953,709954'
    'GX3'           = '710084,710085,710086,710087'
    'GX7'           = '750170,750171,750172,750213'
    'Infinity-I'    = '802069'
    'Infinity-III'  = '807363'
    'Infinity-V'    = '839032'
}

$script:CalDueMap = @{
    'Infinity-VI'   = 'May-26'
    'Infinity-VIII' = 'Jul-26'
    'GX5'           = 'Jul-26'
    'GX6'           = 'Apr-26'
    'GX1'           = 'Mar-26'
    'GX2'           = 'Mar-26'
    'GX3'           = 'Mar-26'
    'GX7'           = 'Mar-26'
    'Infinity-I'    = 'Jun-26'
    'Infinity-III'  = 'Apr-26'
    'Infinity-V'    = 'Jul-26'
}

$SharePointBatchLinkTemplate = (Get-ConfigValue -Name 'SharePointBatchLinkTemplate' -Default 'https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/ROBAL.aspx?viewid=6c9e53c9-a377-40c1-a154-13a13866b52b&view=7&q={BatchNumber}')

$logRootOverride = (Get-EnvNonEmpty -Name 'IPT_LOG_ROOT')
$netRootForLogs  = (Get-EnvNonEmpty -Name 'IPT_NETWORK_ROOT')

$DevLogDir = $null

if ($logRootOverride) {
    try {
        if (-not (Test-Path -LiteralPath $logRootOverride)) {
            New-Item -ItemType Directory -Path $logRootOverride -Force | Out-Null
        }
        $DevLogDir = $logRootOverride
        $global:IPT_LOG_ROOT_EFFECTIVE = $DevLogDir
        $global:IPT_LOG_ROOT_SOURCE = 'ENV_IPT_LOG_ROOT'
    } catch {
        $DevLogDir = $null
    }
}

if (-not $DevLogDir) {
    if ($netRootForLogs -and (Test-Path -LiteralPath $netRootForLogs)) {
        $DevLogDir = Join-Path $netRootForLogs 'Loggar'
        $global:IPT_LOG_ROOT_EFFECTIVE = $DevLogDir
        $global:IPT_LOG_ROOT_SOURCE = 'IPT_NETWORK_ROOT'
    } else {
        $DevLogDir = Join-Path $ScriptRoot 'Loggar'
        $global:IPT_LOG_ROOT_EFFECTIVE = $DevLogDir
        $global:IPT_LOG_ROOT_SOURCE = 'LOCAL_FALLBACK'
    }
}

if (-not (Test-Path -LiteralPath $DevLogDir)) { New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null }

$global:IPT_LOG_MIRROR_DIR = $null
if ($netRootForLogs -and (Test-Path -LiteralPath $netRootForLogs)) {
    try { $global:IPT_LOG_MIRROR_DIR = (Join-Path $netRootForLogs 'Loggar') } catch {}
}

$global:LogPath = Join-Path $DevLogDir ("$($env:USERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt")
$global:StructuredLogPath = [System.IO.Path]::ChangeExtension($global:LogPath, '.jsonl')

function Test-Config {
    $result = [pscustomobject]@{
        Ok       = $true
        Errors   = New-Object System.Collections.Generic.List[object]
        Warnings = New-Object System.Collections.Generic.List[object]
    }

    try {
        $templatePath = Join-Path $ScriptRoot 'output_template-v4.xlsx'
        if (-not (Test-Path -LiteralPath $templatePath)) {
            $null = $result.Errors.Add("Mallfil saknas: $templatePath")
        }
    } catch {
        $null = $result.Errors.Add("Test-Config (template): $($_.Exception.Message)")
    }

    try {
        if (-not (Test-Path -LiteralPath $UtrustningListPath)) {
            $null = $result.Warnings.Add("Utrustningslista saknas: $UtrustningListPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (utrustning): $($_.Exception.Message)")
    }

    try {
        if (-not (Test-Path -LiteralPath $RawDataPath)) {
            $null = $result.Warnings.Add("Kontrollprovsfil saknas: $RawDataPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (rawdata): $($_.Exception.Message)")
    }

    try {
        if (-not (Test-Path -LiteralPath $DevLogDir)) {
            New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null
        }
        $probe = Join-Path $DevLogDir "write_probe.txt"
        Set-Content -Path $probe -Value 'probe' -Encoding UTF8 -Force
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    } catch {
        $null = $result.Warnings.Add("Kunde inte verifiera skrivning till loggmapp: $($_.Exception.Message)")
    }

    if ($result.Errors.Count -gt 0) { $result.Ok = $false }
    return $result
}

# ============================================================================
# CENTRALISERADE KONSTANTER
# ============================================================================
# Dessa värden används av Main.ps1, RuleEngine.ps1 och DataHelpers.ps1.
# Ändra här för att slippa leta efter hårdkodade värden i koden.

$global:IPTConstants = @{

    # --- VL-analyser (referens, används ej aktivt sedan MISQ blev universell) ---
    VlAssays = @(
        'Xpert_HIV-1 Viral Load'
        'HIV-1 Viral Load RUO'
        'HIV-1 Viral Load XC IUO'
        'HIV-1 Viral Load XC RUO'
        'Xpert HIV-1 Viral Load XC'
        'Xpert HCV VL Fingerstick'
        'HCV VL WB RUO'
        'Xpert_HCV Viral Load'
        'HCV Viral Load RUO'
        'Xpert HBV Viral Load'
    )

    # --- HPV-analyser som kräver analysspecifik skanning ---
    HpvAssays = @(
        'Xpert HPV HR'
        'Xpert HPV v2 HR'
        'HPV HR AND GENOTYPE RUO ASSAY'
    )

    # =========================================================================
    # DATA SUMMARY-REGLER  –  Kolumn C → allvar + etikett
    # =========================================================================
    # Pattern   = -ilike-wildcard mot kolumn C (skiftlägesokänsligt)
    # Severity  = 'Major' eller 'Minor'
    # Label     = Visningstext i rapporten
    # Scope     = 'All' (alla analyser) eller 'Hpv' (bara HpvAssays ovan)
    #
    # Lägg till fler rader efter samma mönster för framtida statusar!
    # =========================================================================
    DataSummaryRules = @(
        # === SPECIFIKA MÖNSTER (matchas först → ger detaljerad etikett) ===
        @{ Pattern = '*MAJOR FUNCTIONAL (HPV 16)*';     Severity = 'Major'; Label = 'Major Functional (HPV 16)';       Scope = 'All' }
        @{ Pattern = '*MAJOR FUNCTIONAL (HPV 18_45)*';  Severity = 'Major'; Label = 'Major Functional (HPV 18_45)';    Scope = 'All' }
        @{ Pattern = '*MAJOR FUNCTIONAL (P3)*';         Severity = 'Major'; Label = 'Major Functional (P3)';           Scope = 'All' }
        @{ Pattern = '*MAJOR FUNCTIONAL (P4)*';         Severity = 'Major'; Label = 'Major Functional (P4)';           Scope = 'All' }
        @{ Pattern = '*MAJOR FUNCTIONAL (P5)*';         Severity = 'Major'; Label = 'Major Functional (P5)';           Scope = 'All' }
        @{ Pattern = '*SAC CT FAIL*';                   Severity = 'Minor'; Label = 'Minor Functional (SAC CT FAIL)';  Scope = 'All' }
        @{ Pattern = '*MINOR FUNCTIONAL (SAC)*';        Severity = 'Minor'; Label = 'Minor Functional (SAC)';          Scope = 'All' }
        @{ Pattern = '*MINOR VISUAL*';                  Severity = 'Minor'; Label = 'Minor Functional (Visual)';       Scope = 'All' }
        @{ Pattern = '*MAJOR VISUAL*';                  Severity = 'Major'; Label = 'Major Functional (Visual)';       Scope = 'All' }
        @{ Pattern = '*BARCODE FAIL*';                  Severity = 'Major'; Label = 'Major Functional (Barcode Scan)'; Scope = 'All' }
        @{ Pattern = '*DELAMINATION*';                  Severity = 'Minor'; Label = 'Minor Functional (Delamination)'; Scope = 'All' }

        # === UNIVERSELLA CATCH-ALL (sist → fångar allt som inte träffade ovan) ===
        @{ Pattern = '*MISQUANTITATION*';               Severity = 'Major'; Label = 'Misquantitation';                 Scope = 'All' }
        <#
        @{ Pattern = '*MAJOR FUNCTIONAL*';              Severity = 'Major'; Label = 'Major Functional';                Scope = 'All' }
        @{ Pattern = '*MINOR FUNCTIONAL*';              Severity = 'Minor'; Label = 'Minor Functional';                Scope = 'All' }
        #>
    )

    # --- Tröskelvärden ---
    Thresholds = @{
        PressurePsiMax     = 90       # Maxtryck (PSI) ≥ detta → flaggas
        DataSummaryRowMin  = 10       # Startrad i Data Summary
        DataSummaryRowMax  = 340      # Slutrad i Data Summary
    }

    # --- Fliknamn i output-mallen ---
    OutSheetNames = @{
        SealTestInfo          = 'Seal Test Info'
        StfSum                = 'STF Summary'
        Information           = 'Run Information'
        CsvSammanfattning     = 'QC Summary'
        DataSummary           = 'Data Summary'
        ResampleDataSummary   = 'Resample Data Summary'
    }

    # --- Funktionsflaggor (standardvärden, kan överskridas i $Config) ---

    FeatureFlags = @{
        EnableRuleEngine           = $true
        EnableEquipmentSheet       = $true
        EnableRuleEngineDebugSheet = $true
        EnableRuleEngineRowSkipTrace = $false
        EnableSharePoint           = $true
    }
}
