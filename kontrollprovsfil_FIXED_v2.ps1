<# 
Kontrollprovsfil - Slim (CSV log only)
- EN rad per lyckad uppdatering i CSV (append-only)
- Ingen XLSX-logg, ingen backend, inga viewer-filer
- Behåller befintliga funktioner som ni faktiskt använder: uppdatering + rapporter
- "Less is more" = bortstädad logg-infrastruktur, färger, self-heal, osv.

Krav:
- PowerShell 5.1
- EPPlus 4.5.3.3 (EPPlus.dll måste finnas i känd sökväg eller bredvid skriptet)
#>

param(
    [string]$RawDataPath  = "N:\QC\QC-1\IPT\8. IPT - WR + Rework\1. PQC - Kontrollprovsfil - RÖR EJ -\Script Raw Data\raw_data.xlsx",
    [string]$OutputDir    = "N:\QC\QC-1\IPT\8. IPT - WR + Rework\1. PQC - Kontrollprovsfil - RÖR EJ -\Inventeringsrapport",
    [string]$UpdateLogCsvPath = ""
)

# =====================[ Kolumnschema för rådata ]=====================
$InventoryColumns = @{
    PN              = 1
    Lot             = 2
    Exp             = 3
    Qty             = 4
    LastUpdate      = 5
    Signature       = 6
    ProductStartCol = 7
    ProductEndCol   = 13
    Description     = 14
    LabbDescription = 15
    LabbCode        = 16
}

# =====================[ Minimal update-logg (CSV) ]=====================
$script:RawDataBackedUp = $false

if ([string]::IsNullOrWhiteSpace($UpdateLogCsvPath)) {
    $rawDir = Split-Path $RawDataPath -Parent
    $UpdateLogCsvPath = Join-Path $rawDir "UpdateLog_Kontrollprovsfil.csv"
}
$script:UpdateLogCsvPath = $UpdateLogCsvPath

# =====================[ EPPlus bootstrap (minimal) ]=====================

function Ensure-EPPlus {
    param(
        [string] $SourceDllPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\lib\net35\EPPlus.dll"
    )

    $candidatePaths = @()
    if (-not [string]::IsNullOrWhiteSpace($SourceDllPath)) { $candidatePaths += $SourceDllPath }
    $candidatePaths += (Join-Path $PSScriptRoot 'EPPlus.dll')

    foreach ($cand in $candidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -LiteralPath $cand)) {
            return $cand
        }
    }

    Write-Warning "❌ EPPlus.dll hittades inte. Lägg EPPlus.dll bredvid skriptet eller uppdatera SourceDllPath i Ensure-EPPlus."
    return $null
}

function Load-EPPlus {
    if ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' }) {
        return $true
    }

    $dllPath = Ensure-EPPlus
    if (-not $dllPath) { return $false }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($dllPath)
        [System.Reflection.Assembly]::Load($bytes) | Out-Null
        return $true
    }
    catch {
        Write-Warning "❌ EPPlus-fel vid inläsning från '$dllPath': $($_.Exception.Message)"
        return $false
    }
}

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue

if (-not (Load-EPPlus)) {
    Write-Host "Kritisk komponent (EPPlus) kunde inte laddas. Skriptet avbryts." -ForegroundColor Red
    exit 1
}

# =====================[ Utilities ]=====================

function Backup-RawDataFile {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [string]$BackupDir = "$PSScriptRoot\Backups"
    )
    try {
        if (-not (Test-Path -LiteralPath $BackupDir)) {
            New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
        }
        $timestamp  = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $filename   = Split-Path $FilePath -Leaf
        $backupPath = Join-Path $BackupDir "$($filename)_$timestamp.bak"
        Copy-Item -LiteralPath $FilePath -Destination $backupPath -Force
        Write-Host "Säkerhetskopia skapad: $backupPath" -ForegroundColor Green
    } catch {
        Write-Host "⚠ Kunde inte skapa backup: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Pause-AnyKey {
    param([string]$Prompt = "Tryck Enter för att återgå till huvudmenyn...")

    try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") }
    catch { $null = Read-Host $Prompt }
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory=$true)][scriptblock]$Action,
        [int]$MaxAttempts = 6,
        [int]$BaseDelayMs = 200
    )
    for ($i=1; $i -le $MaxAttempts; $i++) {
        try { return & $Action }
        catch {
            if ($i -eq $MaxAttempts) { throw }
            $delay = [int]($BaseDelayMs * [Math]::Pow(1.7, ($i-1)) + (Get-Random -Minimum 0 -Maximum 120))
            Start-Sleep -Milliseconds $delay
        }
    }
}

# =====================[ CSV log (append-only) ]=====================
# Målet: 1 rad per faktisk ändring (source of truth), robust på nätverks-share.
# Vi undviker separata ".lock"-filer (kan fallera pga rättigheter) och låser istället CSV-filen kort vid append.

function Ensure-UpdateLogCsv {
    param([Parameter(Mandatory=$true)][string]$Path)

    $dir = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    if (Test-Path -LiteralPath $Path) { return }

    # V2-header: inkluderar Action + Row (hjälper spårbarhet vid dubbletter)
    $header = 'Timestamp;User;Signature;Action;PN;Row;OldLot;OldExp;OldQty;NewLot;NewExp;NewQty;Machine'
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)

    # Atomisk skapning (undviker race om två startar samtidigt)
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
        try {
            $bytes = $utf8Bom.GetBytes($header + [Environment]::NewLine)
            $fs.Write($bytes, 0, $bytes.Length)
            $fs.Flush()
        } finally {
            $fs.Close()
        }
    }
    catch [System.IO.IOException] {
        # Troligen skapad av någon annan i samma ögonblick
        if (-not (Test-Path -LiteralPath $Path)) {
            throw "Kunde inte skapa CSV-logg '$Path': $($_.Exception.Message)"
        }
    }
    catch {
        throw "Kunde inte skapa CSV-logg '$Path': $($_.Exception.Message)"
    }
}

function CsvCell {
    param([string]$Text)
    if ($null -eq $Text) { $Text = "" }
    $t = $Text -replace "(\r\n|\n|\r|\t)", " "
    $t = $t.Trim()
    $t = $t -replace '"','""'
    return '"' + $t + '"'
}

function Append-CsvLineWithRetry {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Line
    )

    Invoke-WithRetry -Action {
        try {
            # Öppna kort exklusivt för skriv (läsare kan läsa när vi stängt filen).
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
            try {
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($Line + [Environment]::NewLine)
                $fs.Write($bytes, 0, $bytes.Length)
                $fs.Flush()
            } finally {
                $fs.Close()
            }
        }
        catch [System.IO.IOException] {
            # Dvs. typiskt "The process cannot access the file" (Excel har CSV öppen, AV-scan, nätverks-lås osv.)
            throw
        }
    } | Out-Null
}

# =====================[ CSV spool / fail-safe ]=====================
# Problem vi löser:
# - Om någon har CSV-loggen öppen i Excel (vanligt), kan append ge "sharing violation".
# - Då förloras loggrad om vi bara ger upp.
# Lösning:
# - Vid append-problem: skriv rad(er) till lokal spool-fil i %TEMP%.
# - Vid nästa lyckade loggning: försök först flush:a spool -> main CSV.

function Get-UpdateLogSpoolPath {
    param([Parameter(Mandatory=$true)][string]$MainCsvPath)

    # Spoola i *samma mapp* som main CSV => central spårbarhet (inte %TEMP%).
    # Namn per dator för att minimera krockar.
    $dir = Split-Path $MainCsvPath -Parent
    $name = [System.IO.Path]::GetFileNameWithoutExtension($MainCsvPath)
    $spoolName = "{0}.SPOOL.{1}.csv" -f $name, $env:COMPUTERNAME
    return (Join-Path $dir $spoolName)
}

function Write-UpdateLogSpoolLine {
    param(
        [Parameter(Mandatory=$true)][string]$MainCsvPath,
        [Parameter(Mandatory=$true)][string]$Line
    )

    $spool = Get-UpdateLogSpoolPath -MainCsvPath $MainCsvPath

    if (-not (Test-Path -LiteralPath $spool)) {
        # Skriv samma header som main (v2) för enkel merge.
        $header = 'Timestamp;User;Signature;Action;PN;Row;OldLot;OldExp;OldQty;NewLot;NewExp;NewQty;Machine'
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)

        try {
            $fs = [System.IO.File]::Open($spool, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
            try {
                $bytes = $utf8Bom.GetBytes($header + [Environment]::NewLine)
                $fs.Write($bytes, 0, $bytes.Length)
                $fs.Flush()
            } finally {
                $fs.Close()
            }
        }
        catch [System.IO.IOException] {
            # Om den skapades i samma ögonblick av annan instans, ok.
        }
    }

    # Append med samma robusta metod som main (spool kan också bli låst om någon öppnar den).
    Append-CsvLineWithRetry -Path $spool -Line $Line
}

function Flush-UpdateLogSpool {
    param([Parameter(Mandatory=$true)][string]$MainCsvPath)

    $dir  = Split-Path $MainCsvPath -Parent
    $name = [System.IO.Path]::GetFileNameWithoutExtension($MainCsvPath)

    # Flush:a *alla* spools i mappen (så att vem som helst kan "städa upp" kvarlämnade rader).
    $spools = Get-ChildItem -LiteralPath $dir -Filter ("{0}.SPOOL.*.csv" -f $name) -ErrorAction SilentlyContinue
    if (-not $spools) { return }

    foreach ($spoolFile in $spools) {
        try {
            $lines = Get-Content -LiteralPath $spoolFile.FullName -ErrorAction Stop
            if ($lines.Count -le 1) {
                Remove-Item -LiteralPath $spoolFile.FullName -Force -ErrorAction SilentlyContinue
                continue
            }
            for ($i=1; $i -lt $lines.Count; $i++) {
                Append-CsvLineWithRetry -Path $MainCsvPath -Line $lines[$i]
            }
            Remove-Item -LiteralPath $spoolFile.FullName -Force -ErrorAction SilentlyContinue
        }
        catch {
            # Om flush misslyckas låter vi spool ligga kvar (ingen dataförlust).
            continue
        }
    }
}

function Write-UpdateLogRow {
    param(
        [Parameter(Mandatory=$true)][string]$PN,
        [Parameter(Mandatory=$true)][string]$Signature,
        [Parameter(Mandatory=$true)][ValidateSet("ADD","REMOVE","QTY","OVERWRITE")][string]$Action,
        [int]$RowNumber = 0,
        [string]$OldLot, [string]$OldExp, [string]$OldQty,
        [string]$NewLot, [string]$NewExp, [string]$NewQty
    )

    $path = $script:UpdateLogCsvPath
    Ensure-UpdateLogCsv -Path $path

    # Om vi har en lokal spool från tidigare (t.ex. CSV var låst i Excel), försök flush:a den först.
    Flush-UpdateLogSpool -MainCsvPath $path

    $ts   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $user = [Environment]::UserName
    $pc   = $env:COMPUTERNAME

        $line = (@(
        CsvCell $ts,
        CsvCell $user,
        CsvCell $Signature,
        CsvCell $Action,
        CsvCell $PN,
        CsvCell ([string]$RowNumber),
        CsvCell $OldLot, CsvCell $OldExp, CsvCell $OldQty,
        CsvCell $NewLot, CsvCell $NewExp, CsvCell $NewQty,
        CsvCell $pc
    ) -join ';')

    try {
        Append-CsvLineWithRetry -Path $path -Line $line
    }
    catch {
        # Fail-safe: skriv till lokal spool om main CSV är låst/otillgänglig.
        Write-UpdateLogSpoolLine -MainCsvPath $path -Line $line
        throw "CSV-loggen var låst/otillgänglig. Loggrad spools i samma mapp som huvud-CSV och flushas automatiskt nästa gång (eller av valfri körning när CSV:n går att skriva till). Orsak: $($_.Exception.Message)"
    }
}


# =====================[ Auth + Signature ]=====================

function Request-Password {
    $autoUsers = @("vivian.dao", "elin.sidstedt", "jesper.fredriksson", "afnan.vijitraphongs")
    $currentUser = [Environment]::UserName.ToLower()
    if ($autoUsers -contains $currentUser) {
         Write-Host "Automatiskt godkänd för $currentUser utan lösenord." -ForegroundColor Green
         return $true
    }
    do {
         $securePwd = Read-Host "Ange lösenord eller tryck Enter för PQC-konto" -AsSecureString
         $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePwd)
         $unsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
         [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
         if ($unsecurePassword -eq 'labbkontroll') { return $true }
         Write-Host "Fel lösenord. Tryck (1) för att återgå till huvudmenyn." -ForegroundColor Red -BackgroundColor DarkYellow
         $choice = Read-Host "Ange ditt val"
         if ($choice -eq '1') { return $false }
    } while ($true)
}

function Get-UserSignature {
    param([string]$Prompt = "Ange din signatur för att bekräfta ändringarna")

    $userSignature = @{
        "jesper.fredriksson"   = "JESP"
        "elin.sidstedt"        = "ELS"
        "vivian.dao"           = "vdao"
        "afnan.vijitraphongs"  = "AFVI"
    }

    $currentUser = [Environment]::UserName.ToLower()

    if ($userSignature.ContainsKey($currentUser)) {
        Write-Host "Automatisk signatur för $currentUser $($userSignature[$currentUser])" -ForegroundColor Green
        return $userSignature[$currentUser]
    }

    do {
        $signature = Read-Host $Prompt
        if ($signature.Length -ge 3 -and $signature.Length -le 4) { return $signature }
        Write-Host "Ogiltig signatur – måste vara exakt 3-4 tecken." -ForegroundColor Red
    } while ($true)
}

# =====================[ Date parsing ]=====================

function Try-ParseInventoryDate {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text) -or $Text -eq 'N/A') { return $null }

    try {
        if ($Text -match '^\d{4}-\d{2}-\d{2}$') {
            return [datetime]$Text
        }
        elseif ($Text -match '^\d{4}-\d{2}$') {
            $year  = [int]$Text.Substring(0,4)
            $month = [int]$Text.Substring(5,2)
            $first = New-Object datetime($year, $month, 1)
            return $first.AddMonths(1).AddDays(-1)
        }
        return $null
    } catch { return $null }
}

function Read-ValidExpiryDate {
    param([string]$Prompt = "Ange utgångsdatum (YYYY-MM-DD eller YYYY-MM)")

    while ($true) {
        $input = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($input)) {
            Write-Host "Utgångsdatum får inte vara tomt." -ForegroundColor Red
            continue
        }
        if ($input -notmatch '^\d{4}-\d{2}-\d{2}$' -and $input -notmatch '^\d{4}-\d{2}$') {
            Write-Host "Ogiltigt format. Använd 'YYYY-MM-DD' eller 'YYYY-MM'." -ForegroundColor Red
            continue
        }
        $parsed = Try-ParseInventoryDate -Text $input
        if (-not $parsed) {
            Write-Host "Ogiltigt datum (t.ex. fel månad/dag). Försök igen." -ForegroundColor Red
            continue
        }
        return $input
    }
}

# =====================[ EPPlus helpers ]=====================

function Open-ExcelPackageWithRetry {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$MaxAttempts = 5,
        [int]$DelaySeconds = 2
    )

    for ($attempt=1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $file = New-Object System.IO.FileInfo($Path)
            return (New-Object OfficeOpenXml.ExcelPackage($file))
        }
        catch {
            Write-Host "Kunde inte öppna '$Path' (försök $attempt/$MaxAttempts). Fel: $($_.Exception.Message)" -ForegroundColor Yellow
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    Write-Host "❌ Gav upp efter $MaxAttempts försök att öppna '$Path'." -ForegroundColor Red
    return $null
}

# =====================[ Core: Update inventory ]=====================

function Update-ExcelInventory {
    param([Parameter(Mandatory=$true)][string]$FilePath)

    $package = $null
    try {
        $package = Open-ExcelPackageWithRetry -Path $FilePath
        if (-not $package) { return }

        $sheet = $package.Workbook.Worksheets[1]
        if (-not $sheet) {
            Write-Host "Blad 1 saknas i filen. Avbryter..." -ForegroundColor Red
            return
        }

        $rowCount = $sheet.Dimension.End.Row
        while ($true) {
            Clear-Host
            Write-Host "================ Uppdatera Kontrollprovsfil =================" -ForegroundColor Cyan
            Write-Host "CSV-logg: $($script:UpdateLogCsvPath)" -ForegroundColor DarkGray
            Write-Host "`nAnge P/N eller skriv [Q] för att avbryta"
            $pnToUpdate = Read-Host
            if ($pnToUpdate -eq 'Q') { return }
            if ([string]::IsNullOrWhiteSpace($pnToUpdate)) {
                Write-Host "P/N får inte vara tomt. Försök igen!" -ForegroundColor DarkYellow
                Pause-AnyKey
                continue
            }

            $foundRows = @()
            for ($i = 2; $i -le $rowCount; $i++) {
                if ($sheet.Cells[$i,$InventoryColumns.PN].Text -eq $pnToUpdate) {
                    $foundRows += $i
                    if ($foundRows.Count -eq 4) { break }
                }
            }
            if ($foundRows.Count -eq 0) {
                Write-Host "Inget P/N hittades som matchar '$pnToUpdate'." -ForegroundColor Red
                Pause-AnyKey
                continue
            }

            $index = 1
            foreach ($row in $foundRows) {
                $lotNum     = $sheet.Cells[$row,$InventoryColumns.Lot].Text
                $expiryDate = $sheet.Cells[$row,$InventoryColumns.Exp].Text
                $quantity   = $sheet.Cells[$row,$InventoryColumns.Qty].Text
                Write-Host "`n${index}: Lot: ${lotNum}, Exp: ${expiryDate}, Qty: ${quantity}" -ForegroundColor Green
                $index++
            }

            while ($true) {
                Write-Host "`nVälj en rad att uppdatera (1-$($foundRows.Count)) eller tryck 'B' för att backa"
                $choice = Read-Host
                if ($choice -eq 'Q') { return }
                if ($choice -eq 'B') { break }

                if (($choice -as [int]) -and ($choice -gt 0) -and ($choice -le $foundRows.Count)) {
                    $selectedRow = $foundRows[$choice - 1]
                }
                else {
                    Write-Host "Ogiltigt val, försök igen." -ForegroundColor Red
                    continue
                }

                $oldLot = $sheet.Cells[$selectedRow,$InventoryColumns.Lot].Text
                $oldExp = $sheet.Cells[$selectedRow,$InventoryColumns.Exp].Text
                $oldQty = $sheet.Cells[$selectedRow,$InventoryColumns.Qty].Text

                Write-Host "=============================================" -ForegroundColor Cyan
                Write-Host "Vad vill du uppdatera?"
                Write-Host "1: Nytt Lotnummer, Exp och Qty" -ForegroundColor Magenta
                Write-Host "2: Endast kvantitet"           -ForegroundColor Magenta
                Write-Host "3: Ta bort post (markeras som N/A)" -ForegroundColor Magenta
                Write-Host "B: Back"
                $updateChoice = Read-Host "Välj"
                if ($updateChoice -eq 'B') { continue }

                $newLot = $oldLot; $newExp = $oldExp; $newQty = $oldQty

                switch ($updateChoice) {
                    '1' {
                        $newLot = Read-Host "Ange Lotnummer"
                        $newExp = Read-ValidExpiryDate
                        $newQty = Read-Host "Ange antal"
                    }
                    '2' { $newQty = Read-Host "Ange nytt antal" }
                    '3' { $newLot = "N/A"; $newExp = "N/A"; $newQty = "N/A" }
                    default {
                        Write-Host "Ogiltigt val." -ForegroundColor Red
                        continue
                    }
                }

                # Inga faktiska ändringar? -> spara/logga inte.
                if ($oldLot -eq $newLot -and $oldExp -eq $newExp -and $oldQty -eq $newQty) {
                    Write-Host "Inga ändringar jämfört med befintligt värde. Ingen uppdatering/loggning görs." -ForegroundColor Yellow
                    continue
                }

                if ($updateChoice -ne '3') {
                    Write-Host "`nBekräfta uppdateringen:" -ForegroundColor Cyan
                    Write-Host "Lotnummer: $newLot" -ForegroundColor DarkYellow
                    Write-Host "Utgångsdatum: $newExp" -ForegroundColor DarkYellow
                    Write-Host "Antal: $newQty" -ForegroundColor DarkYellow
                    $confirm = Read-Host "`nÄr detta korrekt? (1) Ja / (2) Nej"
                    if ($confirm -ne '1') {
                        Write-Host "Inga ändringar sparades."
                        continue
                    }
                }

                $userSignature = Get-UserSignature -Prompt ($(if ($updateChoice -eq '3') { "Ange signatur för borttagning" } else { "Ange din signatur" }))
                $today = (Get-Date).ToString("yyyy-MM-dd")

                # Skriv ändringar till ark
                $sheet.Cells[$selectedRow,$InventoryColumns.Lot].Value        = $newLot
                $sheet.Cells[$selectedRow,$InventoryColumns.Exp].Value        = $newExp
                $sheet.Cells[$selectedRow,$InventoryColumns.Qty].Value        = $newQty
                $sheet.Cells[$selectedRow,$InventoryColumns.LastUpdate].Value = ($(if ($updateChoice -eq '3') { "N/A" } else { $today }))
                $sheet.Cells[$selectedRow,$InventoryColumns.Signature].Value  = ($(if ($updateChoice -eq '3') { "N/A" } else { $userSignature }))

                # Backup bara första gången vi faktiskt sparar under denna körning
                if (-not $script:RawDataBackedUp) {
                    Backup-RawDataFile -FilePath $FilePath
                    $script:RawDataBackedUp = $true
                }

                # Klassificera ändringen för CSV (det du faktiskt vill se)
                $action = "OVERWRITE"
                if ($newLot -eq "N/A" -and $newExp -eq "N/A" -and $newQty -eq "N/A") {
                    $action = "REMOVE"
                }
                elseif ($oldLot -eq "N/A" -and $oldExp -eq "N/A" -and $oldQty -eq "N/A" -and
                        ($newLot -ne "N/A" -or $newExp -ne "N/A" -or $newQty -ne "N/A")) {
                    $action = "ADD"
                }
                elseif ($oldLot -eq $newLot -and $oldExp -eq $newExp -and $oldQty -ne $newQty) {
                    $action = "QTY"
                }

                $saved = $false
                try {
                    $package.Save()
                    $saved = $true
                }
                catch {
                    Write-Host "❌ Kunde inte spara ändringen i raw_data.xlsx. CSV-logg skrivs inte. Fel: $($_.Exception.Message)" -ForegroundColor Red
                }

                if ($saved) {
                    try {
                        Write-UpdateLogRow -PN $pnToUpdate -Signature $userSignature -Action $action -RowNumber $selectedRow `
                            -OldLot $oldLot -OldExp $oldExp -OldQty $oldQty `
                            -NewLot $newLot -NewExp $newExp -NewQty $newQty

                        Write-Host "`n✅ Uppdatering sparad + loggad i CSV ($action)." -ForegroundColor Green
                    }
                    catch {
                        Write-Host "⚠ Uppdatering sparad, men kunde inte skriva till CSV-loggen. Fel: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

$nextAction = Read-Host "`nUppdatera (1) nytt P/N, (2) avsluta, (Enter) fortsätt med samma P/N"
                if ($nextAction -eq '1') { break }
                if ($nextAction -eq '2') { return }

                # Visa rader igen för samma P/N
                Clear-Host
                Write-Host "================ Uppdatera Kontrollprovsfil =================" -ForegroundColor Cyan
                Write-Host "`nP/N: $pnToUpdate"
                $index = 1
                foreach ($row in $foundRows) {
                    $lotNum     = $sheet.Cells[$row,$InventoryColumns.Lot].Text
                    $expiryDate = $sheet.Cells[$row,$InventoryColumns.Exp].Text
                    $quantity   = $sheet.Cells[$row,$InventoryColumns.Qty].Text
                    Write-Host "`n${index}: Lot: ${lotNum}, Exp: ${expiryDate}, Qty: ${quantity}" -ForegroundColor Green
                    $index++
                }
            }
        }
    }
    finally {
        if ($package) { $package.Dispose() }
    }
}

# =====================[ Reports (kept, no logging) ]=====================

function Generate-InventoryReport {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter(Mandatory=$true)][string]$OutputDir
    )

    $package = $null
    try {
        $package = Open-ExcelPackageWithRetry -Path $FilePath
        if (-not $package) { return }

        $sheet = $package.Workbook.Worksheets[1]
        if (-not $sheet) { Write-Host "Blad 1 saknas, avbryter..." -ForegroundColor Red; return }

        $rowCount = $sheet.Dimension.End.Row
        $data = @()
        for ($i=2; $i -le $rowCount; $i++) {
            $data += [pscustomobject]@{
                PN              = $sheet.Cells[$i,$InventoryColumns.PN].Text
                LotNr           = $sheet.Cells[$i,$InventoryColumns.Lot].Text
                Exp             = $sheet.Cells[$i,$InventoryColumns.Exp].Text
                Qty             = $sheet.Cells[$i,$InventoryColumns.Qty].Text
                LabbCode        = $sheet.Cells[$i,$InventoryColumns.LabbCode].Text
                LabbDescription = $sheet.Cells[$i,$InventoryColumns.LabbDescription].Text
            }
        }

        if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

        $dateStamp = (Get-Date -Format "yyyyMMdd")
        $xlsxPath  = Join-Path $OutputDir ("InventoryReport_{0}.xlsx" -f $dateStamp)
        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }

        $invPkg = New-Object OfficeOpenXml.ExcelPackage
        try {
            $ws = $invPkg.Workbook.Worksheets.Add("Inventory")

            $ws.Cells["A1"].Value = "PN"
            $ws.Cells["B1"].Value = "LotNr"
            $ws.Cells["C1"].Value = "Exp"
            $ws.Cells["D1"].Value = "Qty"
            $ws.Cells["E1"].Value = "LabbCode"
            $ws.Cells["F1"].Value = "LabbDescription"

            $r=2
            foreach ($it in $data) {
                $ws.Cells["A$r"].Value = $it.PN
                $ws.Cells["B$r"].Value = $it.LotNr
                $ws.Cells["C$r"].Value = $it.Exp
                $ws.Cells["D$r"].Value = $it.Qty
                $ws.Cells["E$r"].Value = $it.LabbCode
                $ws.Cells["F$r"].Value = $it.LabbDescription
                $r++
            }

            if ($r -gt 2) {
                $lastRow = $r-1
                $used = $ws.Cells["A1:F$lastRow"]
                $hdr = $ws.Cells["A1:F1"]
                $hdr.Style.Font.Bold = $true
                $ws.View.FreezePanes(2,1)
                $used.AutoFitColumns()
            }

            $invPkg.SaveAs((New-Object System.IO.FileInfo($xlsxPath)))
            Write-Host "✅ XLSX-rapport genererad: $xlsxPath" -ForegroundColor Green
        } finally { if ($invPkg) { $invPkg.Dispose() } }
    }
    catch {
        Write-Host "Fel vid inventeringsrapport: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally { if ($package) { $package.Dispose() } }
}

function Generate-ExpiringPNsReport {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter(Mandatory=$true)][string]$OutputDir
    )

    $package = $null
    try {
        Clear-Host
        Write-Host "==== Material med kort utgångsdatum ====" -ForegroundColor Cyan

        if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

        $package = Open-ExcelPackageWithRetry -Path $FilePath
        if (-not $package) { return }

        $sheet = $package.Workbook.Worksheets[1]
        if (-not $sheet) { Write-Host "Blad saknas, avbryter..." -ForegroundColor Red; return }

        $rowCount  = $sheet.Dimension.End.Row
        $now       = Get-Date
        $lowerDate = $now.AddDays(-15)
        $upperDate = $now.AddDays(30)

        $dateStamp = (Get-Date -Format "yyyyMMdd")
        $xlsxPath  = Join-Path $OutputDir ("ExpiringPNsReport_{0}.xlsx" -f $dateStamp)
        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }

        $repPkg = New-Object OfficeOpenXml.ExcelPackage
        try {
            $ws = $repPkg.Workbook.Worksheets.Add("Expiring")
            $ws.Cells["A1"].Value = "PN"
            $ws.Cells["B1"].Value = "Lotnummer"
            $ws.Cells["C1"].Value = "Exp. Date"
            $ws.Cells["D1"].Value = "Beskrivning"

            $outRow=2
            for ($i=2; $i -le $rowCount; $i++) {
                $expText = $sheet.Cells[$i,$InventoryColumns.Exp].Text
                $expDate = Try-ParseInventoryDate -Text $expText
                if ($expDate -and $expDate -ge $lowerDate -and $expDate -le $upperDate) {
                    $ws.Cells["A$outRow"].Value = $sheet.Cells[$i,$InventoryColumns.PN].Text
                    $ws.Cells["B$outRow"].Value = $sheet.Cells[$i,$InventoryColumns.Lot].Text
                    $ws.Cells["C$outRow"].Value = $expText
                    $ws.Cells["D$outRow"].Value = $sheet.Cells[$i,$InventoryColumns.Description].Text
                    $outRow++
                }
            }

            if ($outRow -gt 2) {
                $lastRow = $outRow-1
                $ws.View.FreezePanes(2,1)
                $ws.Cells["A1:D$lastRow"].AutoFitColumns()
                $repPkg.SaveAs((New-Object System.IO.FileInfo($xlsxPath)))
                Write-Host "✅ Rapport genererad: $xlsxPath" -ForegroundColor Green
                Invoke-Item $xlsxPath
            } else {
                Write-Host "Inga poster hittades inom datumintervallet." -ForegroundColor Yellow
            }
        } finally { if ($repPkg) { $repPkg.Dispose() } }
    }
    catch {
        Write-Host "Fel vid generering: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally { if ($package) { $package.Dispose() } }
}

function ShowProductInfo {
    param([Parameter(Mandatory=$true)][string]$FilePath)

    Clear-Host
    Write-Host "================ Sök på Produkt för material =================" -ForegroundColor Cyan
    $productName = Read-Host "`nAnge produktnamn"
    if ([string]::IsNullOrWhiteSpace($productName)) { return }

    $package = $null
    try {
        $package = Open-ExcelPackageWithRetry -Path $FilePath
        if (-not $package) { return }

        $sheet = $package.Workbook.Worksheets[1]
        if (-not $sheet) { Write-Host "Inget blad hittades." -ForegroundColor Red; return }

        $rowCount = $sheet.Dimension.End.Row
        $productInfo = @()

        for ($i=2; $i -le $rowCount; $i++) {
            for ($j=$InventoryColumns.ProductStartCol; $j -le $InventoryColumns.ProductEndCol; $j++) {
                if ($sheet.Cells[$i,$j].Text -eq $productName) {
                    $lot = $sheet.Cells[$i,$InventoryColumns.Lot].Text
                    $exp = $sheet.Cells[$i,$InventoryColumns.Exp].Text
                    if ($lot -ne "N/A" -and $exp -ne "N/A") {
                        $productInfo += [pscustomobject]@{
                            PN          = $sheet.Cells[$i,$InventoryColumns.PN].Text
                            LotNr       = $lot
                            Exp         = $exp
                            Description = $sheet.Cells[$i,$InventoryColumns.Description].Text
                        }
                    }
                }
            }
        }

        if ($productInfo.Count -gt 0) {
            Write-Host "`nProduktinformation hittades:" -ForegroundColor Green
            $productInfo | Format-Table -AutoSize
        } else {
            Write-Host "`nIngen information hittades för '$productName'." -ForegroundColor Yellow
        }
    }
    finally { if ($package) { $package.Dispose() } }
}

function Collect-Feedback {
    Clear-Host
    Write-Host "========== Feedback ==========" -ForegroundColor Cyan
    Write-Host "OBS: Feedback loggas inte i denna slim-version." -ForegroundColor DarkGray
    $null = Read-Host "`nAnge din feedback (tryck Enter för att avsluta)"
}

# =====================[ Main menu ]=====================

function Main-Menu {
    do {
        Clear-Host
        Write-Host "Rådatafil : $RawDataPath" -ForegroundColor DarkGray
        Write-Host "CSV-logg  : $($script:UpdateLogCsvPath)" -ForegroundColor DarkGray

        Write-Host "`nKontrollprovsfil - Slim (CSV log only)" -ForegroundColor Cyan
        Write-Host "==============================================================" -ForegroundColor Cyan
        Write-Host "1: Uppdatera Kontrollprovsfil (lösenord krävs)" -ForegroundColor Magenta
        Write-Host "2: Generera inventeringsrapport (lösenord krävs)" -ForegroundColor Magenta
        Write-Host "3: Visa material med kort utgångsdatum"
        Write-Host "4: Visa materialinformation för produkt"
        Write-Host "5: Feedback (ingen loggning)"
        Write-Host "6: Avsluta"
        Write-Host "==============================================================" -ForegroundColor Cyan

        $choice = Read-Host "Välj"
        switch ($choice) {
            '1' { if (Request-Password) { Update-ExcelInventory -FilePath $RawDataPath } }
            '2' { if (Request-Password) { Generate-InventoryReport -FilePath $RawDataPath -OutputDir $OutputDir } }
            '3' { Generate-ExpiringPNsReport -FilePath $RawDataPath -OutputDir $OutputDir }
            '4' { ShowProductInfo -FilePath $RawDataPath }
            '5' { Collect-Feedback }
            '6' { return }
            default { Write-Host "Ogiltigt val." -ForegroundColor Red }
        }

        Pause-AnyKey
    } while ($true)
}

# =====================[ Startup ]=====================

if (-not (Test-Path -LiteralPath $RawDataPath)) {
    Write-Host "❌ Hittar inte rådatafilen:" -ForegroundColor Red
    Write-Host "   $RawDataPath" -ForegroundColor Yellow
    exit 1
}

Ensure-UpdateLogCsv -Path $script:UpdateLogCsvPath
Main-Menu