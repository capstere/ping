Förbättrad spårbarhet i IPTCompile 16.3-rapporten
Bakgrund och mål
Du vill ENDAST förbättra rapportens spårbarhet (traceability) utan att ändra funktionalitet eller regel-/tolkningslogik. Det konkreta målet är att rapporten tydligt visar exakta källor (filnamn) för:

Primary CSV och Resample CSV, Worksheet-filen som Data Summary/Resample Data Summary kommer från, samt gärna Seal POS/NEG när de redan finns i flödet.

Din tekniska ram är Windows PowerShell 5.1 och EPPlus 4.5.3.3 (som är en “legacy/deprecated” version på NuGet, vilket stärker argumentet för små, lågrisk ändringar istället för större refaktor eller biblioteksbyte).

Nuvarande rapportarkitektur i IPTCompile 16.3
Rapporten byggs i praktiken av två delar:

Template-filen output_template-v4.xlsx ger grundflikar som Run Information (där “CSV”, “Worksheet”, “Seal Test POS/NEG” typiskt skrivs in). Sedan genererar du en extra flik CSV Sammanfattning vid rapportskapandet via regelmotorn (Write-RuleEngineDebugSheet i Modules/RuleEngine.ps1).

I EPPlus skapar/manipulerar man rapporten via ExcelPackage/ExcelWorksheet och skriver/stylar celler via Cells[...] och Style (t.ex. Numberformat.Format, WrapText, Merge, fill/font etc.). 

Det som är viktigt ur spårbarhetssynpunkt i din kodbas är att CSV Sammanfattning skapas från scratch i Write-RuleEngineDebugSheet och att den i din nuvarande layout lämnar en “luft-rad” (rad 2) mellan titel och första sektion. Det är därför en ideal plats att lägga en spårbarhetsbanner utan att flytta eller skriva över existerande tabell/sektioner.

Robusthets- och riskkriterier för ändringen
För att uppfylla “minimal risk” och “ingen funktionspåverkan” är det här de mest relevanta designvalen:

Ändra endast presentation i rapporten: inga nya regler, inga nya filtreringar, inga nya beräkningar.

Lägg spårbarhet på en plats som inte påverkar befintliga rader/sektioner: i CSV Sammanfattning används en rad som redan är tom.

Använd EPPlus-styling på range-nivå (merge, numberformat, wrap, font) på samma sätt som resten av rapporten gör: EPPlus rekommenderar att cellformat sätts via Style på Cells[]/range. 

Gör ändringen bakåtkompatibel internt: genom att lägga till optionella parametrar med defaultvärden i Write-RuleEngineDebugSheet, så går funktionen att anropa även utan att skicka spårbarhetsdata. (I PowerShell är defaultvärden i param-blocket ett standardmönster för att göra funktioner valfria/robusta.) 

Rekommenderad lösning för spårbarhet i rapporten
Lösningen nedan gör exakt det du efterfrågar och inget mer:

I Modules/RuleEngine.ps1 utökas Write-RuleEngineDebugSheet med fem optionella textparametrar: PrimaryCsvFile, ResampleCsvFile, WorksheetFile, SealPosFile, SealNegFile.

Dessa skrivs som en banner på rad 2 (merge A2:H2) i CSV Sammanfattning med texten “Källor: …” och en tydlig, men diskret stil: textformat ('@'), kursiv, liten font och wrap.

I Main.ps1 beräknas filnamnen (redan “cleanade” från eventuella TEMP-prefix genom din befintliga Get-CleanLeaf) och skickas in till Write-RuleEngineDebugSheet. Samtidigt tas din tidigare “post-write” hack för enbart CSV bort, eftersom banner-raden nu ägs av regelmotorfliken (mer sammanhållet, mindre dubbellogik).

Som extra tydlighet i Run Information skrivs en kort notis i kolumn C på “Worksheet”-raden: att Worksheet är källa för Data Summary/Resample Data Summary. Detta är fortfarande en ren presentationsdetalj.

Unified diff patch
Patchen nedan antar din struktur:

Main.ps1
Modules/RuleEngine.ps1
Patch för Modules/RuleEngine.ps1
diff
Kopiera
--- a/Modules/RuleEngine.ps1
+++ b/Modules/RuleEngine.ps1
@@ -2103,7 +2103,13 @@
     [Parameter(Mandatory)][pscustomobject]$RuleEngineResult,
     [Parameter(Mandatory=$false)][bool]$IncludeAllRows = $false,
     [Parameter(Mandatory=$false)][object[]]$DataSummaryFindings = @(),
-    [Parameter(Mandatory=$false)][int]$StfCount = 0
+    [Parameter(Mandatory=$false)][int]$StfCount = 0,
+    # Spårbarhet (filnamn) – endast visning i rapport, ingen logik
+    [Parameter(Mandatory=$false)][string]$PrimaryCsvFile  = '',
+    [Parameter(Mandatory=$false)][string]$ResampleCsvFile = '',
+    [Parameter(Mandatory=$false)][string]$WorksheetFile   = '',
+    [Parameter(Mandatory=$false)][string]$SealPosFile     = '',
+    [Parameter(Mandatory=$false)][string]$SealNegFile     = ''
 )
 
 # ============================================================================
@@ -2316,6 +2322,29 @@
 $titleRng.Style.Fill.BackgroundColor.SetColor($Colors.HeaderBg)
 $titleRng.Style.Font.Color.SetColor($Colors.HeaderFg)
 $ws.Row($row).Height = 25
+
+# Rad 2 är tom i grundlayouten → nyttjas för spårbarhet (filnamn)
+try {
+    $metaParts = @()
+    if ($PrimaryCsvFile)  { $metaParts += ("CSV: {0}" -f $PrimaryCsvFile) }
+    if ($ResampleCsvFile) { $metaParts += ("RES CSV: {0}" -f $ResampleCsvFile) }
+    if ($WorksheetFile)   { $metaParts += ("Worksheet (Data Summary/Resample): {0}" -f $WorksheetFile) }
+    if ($SealPosFile)     { $metaParts += ("Seal POS: {0}" -f $SealPosFile) }
+    if ($SealNegFile)     { $metaParts += ("Seal NEG: {0}" -f $SealNegFile) }
+
+    if ($metaParts -and $metaParts.Count -gt 0) {
+        $metaText = "Källor: " + ($metaParts -join ' | ')
+        $metaRng = $ws.Cells[2, 1, 2, 8]
+        $metaRng.Merge = $true
+        $metaRng.Style.Numberformat.Format = '@'
+        $metaRng.Style.Font.Italic = $true
+        $metaRng.Style.Font.Size = 9
+        $metaRng.Style.WrapText = $true
+        $ws.Cells[2, 1].Value = $metaText
+        try { $ws.Row(2).Height = 18 } catch {}
+    }
+} catch {}
+
 $row += 2
 
 # ============================================================================
Patch för Main.ps1
diff
Kopiera
--- a/Main.ps1
+++ b/Main.ps1
@@ -3877,8 +3877,17 @@
         if ($selLsp) {
             $wsInfo.Cells["B$rowWsFile"].Style.Numberformat.Format = '@'
             $wsInfo.Cells["B$rowWsFile"].Value = (Get-CleanLeaf $selLsp)
+            # Spårbarhet: tydliggör att Data Summary / Resample Data Summary kommer från Worksheet
+            try {
+                $wsInfo.Cells["C$rowWsFile"].Style.Numberformat.Format = '@'
+                if (-not (($wsInfo.Cells["C$rowWsFile"].Text + '').Trim())) {
+                    $wsInfo.Cells["C$rowWsFile"].Style.Font.Italic = $true
+                    $wsInfo.Cells["C$rowWsFile"].Value = 'Källa för Data Summary / Resample Data Summary'
+                }
+            } catch {}
         } else {
             $wsInfo.Cells["B$rowWsFile"].Value = ''
+            try { $wsInfo.Cells["C$rowWsFile"].Value = '' } catch {}
         }
 
         $consPart  = Get-ConsensusValue -Type 'Part'      -Ws $headerWs.PartNo      -Pos $headerPos.PartNumber   -Neg $headerNeg.PartNumber
@@ -4725,32 +4734,28 @@
                         if ($script:DataSummaryFindings) { $dsf = @($script:DataSummaryFindings) }
                         $stf = 0
                         if ($script:StfCount) { $stf = [int]$script:StfCount }
-                        [void](Write-RuleEngineDebugSheet -Pkg $pkgOut -RuleEngineResult $script:RuleEngineShadow -IncludeAllRows $includeAll -DataSummaryFindings $dsf -StfCount $stf)
-
-                        # Visa vilken CSV som klassades som primär/resample direkt i CSV Sammanfattning (rad 2)
+                        # Spårbarhet (filnamn) i rapporten: CSV / Resample / Worksheet (Data Summary) / Seal POS/NEG
+                        # OBS: Endast visning i output. Ändrar inte regel-logik.
+                        $tracePrimary  = ''
+                        $traceResample = ''
+                        $traceWs       = ''
+                        $tracePos      = ''
+                        $traceNeg      = ''
                         try {
-                            $wsDbg = $pkgOut.Workbook.Worksheets['CSV Sammanfattning']
-                            if ($wsDbg) {
-                                $pPath = if ($selCsvOrig) { $selCsvOrig } elseif ($selCsv) { $selCsv } else { '' }
-                                $rPath = if ($selCsvResOrig) { $selCsvResOrig } elseif ($selCsvRes) { $selCsvRes } else { '' }
-                                $pLeaf = Get-CleanLeaf $pPath
-                                $rLeaf = Get-CleanLeaf $rPath
-                                $lbl = if ($pLeaf -and $rLeaf) {
-                                    "CSV: {0} | RES CSV: {1}" -f $pLeaf, $rLeaf
-                                } elseif ($pLeaf) {
-                                    "CSV: {0}" -f $pLeaf
-                                } elseif ($rLeaf) {
-                                    "Resample CSV: {0}" -f $rLeaf
-                                } else { '' }
-                                if ($lbl) {
-                                    $rng = $wsDbg.Cells[2, 1, 2, 8]
-                                    $rng.Merge = $true
-                                    $rng.Style.Numberformat.Format = '@'
-                                    $rng.Style.Font.Italic = $true
-                                    $wsDbg.Cells[2, 1].Value = $lbl
-                                }
-                            }
+                            $pPath = if ($selCsvOrig)    { $selCsvOrig }    elseif ($selCsv)    { $selCsv }    else { $null }
+                            $rPath = if ($selCsvResOrig) { $selCsvResOrig } elseif ($selCsvRes) { $selCsvRes } else { $null }
+                            $wPath = if ($selLspOrig)    { $selLspOrig }    elseif ($selLsp)    { $selLsp }    else { $null }
+                            $posP  = if ($selPosOrig)    { $selPosOrig }    elseif ($selPos)    { $selPos }    else { $null }
+                            $negP  = if ($selNegOrig)    { $selNegOrig }    elseif ($selNeg)    { $selNeg }    else { $null }
+
+                            $tracePrimary  = Get-CleanLeaf $pPath
+                            $traceResample = Get-CleanLeaf $rPath
+                            $traceWs       = Get-CleanLeaf $wPath
+                            $tracePos      = Get-CleanLeaf $posP
+                            $traceNeg      = Get-CleanLeaf $negP
                         } catch {}
+
+                        [void](Write-RuleEngineDebugSheet -Pkg $pkgOut -RuleEngineResult $script:RuleEngineShadow -IncludeAllRows $includeAll -DataSummaryFindings $dsf -StfCount $stf -PrimaryCsvFile $tracePrimary -ResampleCsvFile $traceResample -WorksheetFile $traceWs -SealPosFile $tracePos -SealNegFile $traceNeg)
                     } else {
                         Gui-Log "⚠️ Kunde inte skriva CSV-Sammanfattning." 'Warn'
                     }
Vad rapporten får efter patchen
På fliken CSV Sammanfattning (högst upp) får du nu en tydlig “Källor: …”-rad som kan innehålla:

CSV: <Primary.csv>
RES CSV: <Resample.csv>
Worksheet (Data Summary/Resample): <Worksheet.xlsx>
Seal POS: <SealPOS.xlsx>
Seal NEG: <SealNEG.xlsx>

Allt är filnamn, inte full path, vilket matchar ditt krav.

På fliken Run Information blir det också mer explicit att Worksheet-raden är källan för Data Summary/Resample Data Summary.

Verifiering utan att riskera funktionaliteten
För att verifiera att det här verkligen bara är output och inte påverkar logiken, räcker det att testköra fyra realistiska scenarier:

Körning med endast Primary CSV valt. CSV Sammanfattning ska visa “CSV: …” men inga RES/Worksheet/Seal om de inte finns.

Körning med Primary + Resample CSV. Banner-raden ska visa båda.

Körning där Worksheet saknas. Banner-raden ska då sakna Worksheet-delen (och Run Information notisen ska rensas/inte skrivas).

Körning med Seal POS/NEG valda. Banner-raden ska inkludera Seal POS/NEG (om de faktiskt är valda i flödet), samtidigt som övriga rapporttabeller ser likadana ut som före patchen.

EPPlus-stylingen i patchen bygger på standardmönster för cellformat (merge + style på range, wrap text via Style.WrapText, och textformat via Style.Numberformat.Format). 