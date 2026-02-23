Ja — nu är jag säker på att det är IPTCompile 16.3.

Jag har packat upp /mnt/data/IPTCompile 16.3.zip och verifierat att filerna ligger här:
	•	IPTCompile_16/Modules/RuleEngine.ps1 (har Write-RuleEngineDebugSheet på rad ~2100)
	•	IPTCompile_16/Main.ps1 (har “Seal Test Info (blad)” runt rad ~2510 och debug-sheet skrivning runt ~4728)

Nedan får du hela diffarna (2 st), korrigerade så de fungerar ihop och följer ditt krav:

✅ Endast spårbarhet i rapporten
✅ Ingen förändring av regel-/beslutslogik
✅ Seal Test Info: source-sheet trace BARA vid mismatch, en rad, snyggt i kolumn D
✅ CSV Sammanfattning: “Källor: …” skrivs av RuleEngine (rad 2), filnamn
✅ Run Information: gör det explicit att Worksheet är källa för Data Summary/Resample (i kolumn C om tom)

⸻

DIFF 1/2 — IPTCompile_16/Modules/RuleEngine.ps1

--- a/IPTCompile_16/Modules/RuleEngine.ps1
+++ b/IPTCompile_16/Modules/RuleEngine.ps1
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
@@ -2316,7 +2322,35 @@
 $titleRng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
 $titleRng.Style.Fill.BackgroundColor.SetColor($Colors.HeaderBg)
 $titleRng.Style.Font.Color.SetColor($Colors.HeaderFg)
 $ws.Row($row).Height = 25
+
+# Rad 2 är tom i grundlayouten → nyttjas för spårbarhet (filnamn)
+try {
+    if (-not (($ws.Cells[2,1].Text + '').Trim())) {
+        $metaParts = @()
+        if ($PrimaryCsvFile)  { $metaParts += ("CSV: {0}" -f $PrimaryCsvFile) }
+        if ($ResampleCsvFile) { $metaParts += ("RES CSV: {0}" -f $ResampleCsvFile) }
+        if ($WorksheetFile)   { $metaParts += ("Worksheet (Data Summary/Resample): {0}" -f $WorksheetFile) }
+        if ($SealPosFile)     { $metaParts += ("Seal POS: {0}" -f $SealPosFile) }
+        if ($SealNegFile)     { $metaParts += ("Seal NEG: {0}" -f $SealNegFile) }
+
+        if ($metaParts -and $metaParts.Count -gt 0) {
+            $metaText = "Källor: " + ($metaParts -join ' | ')
+            $metaRng = $ws.Cells[2, 1, 2, 8]
+            $metaRng.Merge = $true
+            $metaRng.Style.Numberformat.Format = '@'
+            $metaRng.Style.Font.Italic = $true
+            $metaRng.Style.Font.Size = 9
+            $metaRng.Style.WrapText = $true
+            $ws.Cells[2, 1].Value = $metaText
+            try { $ws.Row(2).Height = 30 } catch {}
+        }
+    }
+} catch {}
+
 $row += 2
 
 # ============================================================================


⸻

DIFF 2/2 — IPTCompile_16/Main.ps1

--- a/IPTCompile_16/Main.ps1
+++ b/IPTCompile_16/Main.ps1
@@ -2669,6 +2669,7 @@
         $row = 3
         foreach ($f in $fields) {
 $valNeg=''; $valPos=''
+$srcNeg=''; $srcPos=''
 
 # För utrustningsrader: samla per flik (inte avbrott på första!)
 $perNeg = $null
 $perPos = $null
@@ -2694,21 +2695,25 @@
 else {
     # Befintligt beteende för "vanliga" fält: första flik med värde
     foreach ($wsN in $pkgNeg.Workbook.Worksheets) {
         if ($wsN.Name -eq "Worksheet Instructions") { continue }
         $cell = $wsN.Cells[$f.Cell]
         if ($cell.Value -ne $null) {
             if ($cell.Value -is [datetime]) { $valNeg = $cell.Value.ToString('MMM-yy') } else { $valNeg = $cell.Text }
+            $srcNeg = $wsN.Name
             break
         }
     }
 
     foreach ($wsP in $pkgPos.Workbook.Worksheets) {
         if ($wsP.Name -eq "Worksheet Instructions") { continue }
         $cell = $wsP.Cells[$f.Cell]
         if ($cell.Value -ne $null) {
             if ($cell.Value -is [datetime]) { $valPos = $cell.Value.ToString('MMM-yy') } else { $valPos = $cell.Text }
+            $srcPos = $wsP.Name
             break
         }
     }
 }
@@ -2726,22 +2731,26 @@
         if ($mismatchFields -contains $f.Label) {
                 # D3:D9: visa tydlig Match/Mismatch med symboler
                 if ($valNeg -and $valPos) {
                     if ($valNeg -ne $valPos) {
-                        $wsOut1.Cells["D$row"].Value = "⚠ Mismatch"
+                        $traceNeg = if ($srcNeg) { "NEG:$srcNeg@$($f.Cell)" } else { "NEG@$($f.Cell)" }
+                        $tracePos = if ($srcPos) { "POS:$srcPos@$($f.Cell)" } else { "POS@$($f.Cell)" }
+                        $wsOut1.Cells["D$row"].Value = ("⚠ Mismatch ({0} | {1})" -f $traceNeg, $tracePos)
+                        $wsOut1.Cells["D$row"].Style.WrapText = $false
+                        try { $wsOut1.Cells["D$row"].Style.Font.Size = 9 } catch {}
                         Style-Cell $wsOut1.Cells["D$row"] $true "FF0000" "Medium" "FFFFFF"
                         Gui-Log "⚠️ Avvikelse: $($f.Label) (NEG='$valNeg' vs POS='$valPos')"
                     } else {
                         $wsOut1.Cells["D$row"].Value = "✓ Match"
+                        $wsOut1.Cells["D$row"].Style.WrapText = $false
                         Style-Cell $wsOut1.Cells["D$row"] $true "C6EFCE" "Medium" "006100"
                     }
                 } elseif ($valNeg -or $valPos) {
                     # Bara en av filerna har värde - markera som varning
                     $wsOut1.Cells["D$row"].Value = "⚠ Saknas"
+                    $wsOut1.Cells["D$row"].Style.WrapText = $false
                     Style-Cell $wsOut1.Cells["D$row"] $true "FFE699" "Medium" "806000"
                 }
             }
@@ -3877,8 +3886,20 @@
         if ($selLsp) {
             $wsInfo.Cells["B$rowWsFile"].Style.Numberformat.Format = '@'
             $wsInfo.Cells["B$rowWsFile"].Value = (Get-CleanLeaf $selLsp)
+            # Spårbarhet: gör explicit att Data Summary / Resample Data Summary kommer från Worksheet
+            try {
+                $cellC = $wsInfo.Cells["C$rowWsFile"]
+                $cellC.Style.Numberformat.Format = '@'
+                if (-not (($cellC.Formula + '').Trim()) -and -not (($cellC.Text + '').Trim())) {
+                    $cellC.Style.Font.Italic = $true
+                    $cellC.Value = 'Källa för Data Summary / Resample Data Summary'
+                }
+            } catch {}
         } else {
             $wsInfo.Cells["B$rowWsFile"].Value = ''
+            try { $wsInfo.Cells["C$rowWsFile"].Value = '' } catch {}
         }
@@ -4728,32 +4750,36 @@
                         $dsf = @()
                         if ($script:DataSummaryFindings) { $dsf = @($script:DataSummaryFindings) }
                         $stf = 0
                         if ($script:StfCount) { $stf = [int]$script:StfCount }
-                        [void](Write-RuleEngineDebugSheet -Pkg $pkgOut -RuleEngineResult $script:RuleEngineShadow -IncludeAllRows $includeAll -DataSummaryFindings $dsf -StfCount $stf)
-
-                        # Visa vilken CSV som klassades som primär/resample direkt i CSV Sammanfattning (rad 2)
-                        try {
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
-                        } catch {}
+                        # Spårbarhet (filnamn) i rapporten: CSV / Resample / Worksheet (Data Summary) / Seal POS/NEG
+                        # OBS: Endast visning i output. Ändrar inte regel-logik.
+                        $tracePrimary  = ''
+                        $traceResample = ''
+                        $traceWs       = ''
+                        $tracePos      = ''
+                        $traceNeg      = ''
+                        try {
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
+                        } catch {}
+
+                        [void](Write-RuleEngineDebugSheet -Pkg $pkgOut -RuleEngineResult $script:RuleEngineShadow -IncludeAllRows $includeAll -DataSummaryFindings $dsf -StfCount $stf -PrimaryCsvFile $tracePrimary -ResampleCsvFile $traceResample -WorksheetFile $traceWs -SealPosFile $tracePos -SealNegFile $traceNeg)
                     } else {
                         Gui-Log "⚠️ Kunde inte skriva CSV-Sammanfattning." 'Warn'
                     }


⸻

Hur layouten blir (det du vill ha)

Seal Test Info (kolumn D)
	•	Match: ✓ Match
	•	Mismatch: ⚠ Mismatch (NEG:Sheet@D14 | POS:Other@D14)  ✅ (bara då)

CSV Sammanfattning (rad 2)

Källor: CSV: X.csv | RES CSV: Y.csv | Worksheet (Data Summary/Resample): Z.xlsx | Seal POS: ... | Seal NEG: ...

Run Information

Worksheet-raden får text i C-kolumnen:
Källa för Data Summary / Resample Data Summary (bara om cellen är tom och ej formel)

⸻

Om du vill kan vi även göra mismatch-strängen ännu “renare” (t.ex. utan ordet “Mismatch” i parentesen), men ovan är redan max tydlighet + minimal visuellt brus.