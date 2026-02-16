#Requires -Version 5.1
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$BasePath      = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\Skift 1\Mahdi\powerpoint\AutoMappscript'
$EquipmentPath = Join-Path $BasePath 'equipment.xml'
$StatusPath    = Join-Path $BasePath 'status.txt'

function ConvertTo-DataTable {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [PSObject[]] $InputObject,

        [Parameter()]
        [string] $DefaultType = 'System.String'
    )

    begin {
        $dataTable = New-Object -TypeName 'System.Data.DataTable'
        $first = $true
        $types = @(
            'System.String','System.Boolean','System.Byte[]','System.Byte','System.Char','System.DateTime',
            'System.Decimal','System.Double','System.Guid','System.Int16','System.Int32','System.Int64',
            'System.Single','System.UInt16','System.UInt32','System.UInt64'
        )
    }

    process {
        foreach ($object in $InputObject) {
            $dataRow = $dataTable.NewRow()

            foreach ($property in $object.PSObject.get_properties()) {
                if ($first) {
                    if ($types -contains $property.TypeNameOfValue) { $dataType = $property.TypeNameOfValue }
                    else { $dataType = $DefaultType }

                    $dataColumn = New-Object 'System.Data.DataColumn' $property.Name, $dataType
                    $null = $dataTable.Columns.Add($dataColumn)
                }

                if ($property.Value -ne $null) {
                    if (($property.Value.GetType().IsArray) -or ($property.TypeNameOfValue -like '*collection*')) {
                        $dataRow.Item($property.Name) = $property.Value | ConvertTo-Xml -As 'String' -NoTypeInformation -Depth 1
                    } else {
                        $dataRow.Item($property.Name) = $property.Value -as $dataType
                    }
                }
            }

            $null = $dataTable.Rows.Add($dataRow)
            $first = $false
        }
    }

    end { Write-Output (,($dataTable)) }
}

try {
    Import-Module PnP.PowerShell -ErrorAction Stop
} catch {
    Install-Module PnP.PowerShell -MaximumVersion 1.12.0 -Scope CurrentUser -Force
    Import-Module PnP.PowerShell
}
$env:PNPPOWERSHELL_UPDATECHECK = "Off"

$equipment = Import-Clixml -LiteralPath $EquipmentPath

function Get-StatusText {
    if (-not (Test-Path -LiteralPath $StatusPath)) { return 'Disabled' }
    ($null = $null)
    try { return ((Get-Content -LiteralPath $StatusPath -ErrorAction Stop) + '').Trim() } catch { return 'Disabled' }
}

function Set-StatusText([string]$NewStatus) {
    Set-Content -LiteralPath $StatusPath -Value $NewStatus
}

function Show-EquipmentDialog {
    param([hashtable]$Hashtable)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Instrumentlist (AutoMappscript)"
    $form.StartPosition = 'CenterParent'
    $form.Size = New-Object System.Drawing.Size(700, 520)
    $form.MinimumSize = New-Object System.Drawing.Size(650, 450)

    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'
    $panelTop.Height = 56
    $panelTop.BackColor = [System.Drawing.Color]::FromArgb(31,78,121)
$form.Controls.Add($panelTop)

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "Ändringar sparas automatiskt. Stäng fönstret när du är klar."
$lbl.ForeColor = [System.Drawing.Color]::White
$lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lbl.AutoSize = $true
$lbl.Location = New-Object System.Drawing.Point(16, 18)
$panelTop.Controls.Add($lbl)

$panelBody = New-Object System.Windows.Forms.Panel
$panelBody.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelBody.Padding = New-Object System.Windows.Forms.Padding(0)
$form.Controls.Add($panelBody)
$panelBody.BringToFront()

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode = 'CellSelect'
$grid.MultiSelect = $false
$grid.RowHeadersVisible = $false
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = 'FixedSingle'
$grid.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
$grid.EnableHeadersVisualStyles = $false
$panelBody.Controls.Add($grid)

    $dataTable = $Hashtable.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{ Key = $_.Key; Value = $_.Value }
    } | ConvertTo-DataTable

    $grid.DataSource = $dataTable

    try { $grid.Columns["Key"].ReadOnly = $true } catch {}

    $grid.Add_CellValueChanged({
        try {
            $key = $grid.Rows[$_.RowIndex].Cells["Key"].Value
            $value = $grid.Rows[$_.RowIndex].Cells["Value"].Value
            $Hashtable[$key] = $value

            Export-Clixml -LiteralPath $EquipmentPath -InputObject $Hashtable
        } catch {
        }
    })

    $grid.Add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
    })

    $null = $form.ShowDialog()
    return
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "AutoMappscript – Status & Instruments"
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(520, 320)
$form.MinimumSize = New-Object System.Drawing.Size(520, 320)
$form.MaximizeBox = $false
$form.FormBorderStyle = 'FixedDialog'

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 72
$header.BackColor = [System.Drawing.Color]::FromArgb(31,78,121)
$form.Controls.Add($header)

$title = New-Object System.Windows.Forms.Label
$title.Text = "AutoMappscript Status Control"
$title.ForeColor = [System.Drawing.Color]::White
$title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(16, 18)
$header.Controls.Add($title)

$statusLbl = New-Object System.Windows.Forms.Label
$statusLbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$statusLbl.AutoSize = $true
$statusLbl.Location = New-Object System.Drawing.Point(20, 92)
$form.Controls.Add($statusLbl)

function Refresh-StatusUi {
    $s = Get-StatusText
    if ($s -eq 'Enabled') {
        $statusLbl.Text = "Status: ENABLED"
        $statusLbl.ForeColor = [System.Drawing.Color]::FromArgb(0,128,0)
    } else {
        $statusLbl.Text = "Status: DISABLED"
        $statusLbl.ForeColor = [System.Drawing.Color]::FromArgb(178,34,34)
    }
}

$btnToggle = New-Object System.Windows.Forms.Button
$btnToggle.Text = "Enable / Disable"
$btnToggle.Size = New-Object System.Drawing.Size(200, 42)
$btnToggle.Location = New-Object System.Drawing.Point(20, 140)
$btnToggle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($btnToggle)

$btnEdit = New-Object System.Windows.Forms.Button
$btnEdit.Text = "View / Change Instrumentlist"
$btnEdit.Size = New-Object System.Drawing.Size(260, 42)
$btnEdit.Location = New-Object System.Drawing.Point(240, 140)
$btnEdit.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($btnEdit)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Size = New-Object System.Drawing.Size(120, 36)
$btnExit.Location = New-Object System.Drawing.Point(380, 232)
$btnExit.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($btnExit)

$btnToggle.Add_Click({
    $current = Get-StatusText
    $new = if ($current -eq 'Enabled') { 'Disabled' } else { 'Enabled' }
    try {
        Set-StatusText -NewStatus $new
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Kunde inte uppdatera statusfilen.`r`n$StatusPath`r`n`r`n$($_.Exception.Message)", "Fel", 'OK', 'Error') | Out-Null
        return
    }
    Refresh-StatusUi
})

$btnEdit.Add_Click({
    try {
        $script:equipment = Import-Clixml -LiteralPath $EquipmentPath
        Show-EquipmentDialog -Hashtable $script:equipment
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna instrumentlistan.`r`n$EquipmentPath`r`n`r`n$($_.Exception.Message)", "Fel", 'OK', 'Error') | Out-Null
    }
})

$btnExit.Add_Click({ $form.Close() })

$form.Add_Shown({ Refresh-StatusUi })
[void]$form.ShowDialog()