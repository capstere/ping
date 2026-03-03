# OpenXML deterministic merge overwrite (PowerShell 5.1)

This project now signs worksheet name/date cells with deterministic merge-aware logic.

## What changed

- `Invoke-WorksheetSignature_OpenXml` now takes `-Overwrite` as a strict `[bool]`.
- Main call site passes overwrite as boolean: `-Overwrite ([bool]$m.Overwrite)`.
- Signing writes now resolve merged target -> merge owner (top-left) using:
  - `Get-MergeIndexes_OpenXml` (`Forward` + `Reverse` maps)
  - `Write-OpenXmlCellText_DeterministicMerge`
- `Overwrite=$false` checks the full merge owner-row group (owner + tails) before writing.
- `Overwrite=$true` clears owner-row merge group first, writes owner value, then fills tails with `N/A`.
- Writes are recorded using resolved owner coordinates in `WrittenCells`.
- Overwrite mode performs per-sheet read-back verification (`Verify-WorksheetSignatures_OpenXml`) before reporting the sheet as written.

## Debug logging

Enable per-sheet debug logs with `-DebugLog`:

```powershell
Invoke-WorksheetSignature_OpenXml -Path $path -FullName 'Jane Doe' -SignDateYmd '2026-03-03' -Mode Sammanstallning -Overwrite $true -DebugLog -ModulesRoot $modulesRoot
```

Example log lines:

```text
[OpenXML] Invoke start: mode=Sammanstallning, overwrite=True, file=Worksheet.xlsx
[OpenXML] Sheet 'Test Summary': nameRow=54, dateRow=54
[OpenXML] Sheet 'Test Summary' Name: target=C54, owner=C54, written=True, reason=written
```
