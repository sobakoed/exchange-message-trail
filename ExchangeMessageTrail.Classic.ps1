#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Message Trail (EMT): the classic window (Windows Forms). A fallback for one version.

.DESCRIPTION
    The usual window is ExchangeMessageTrail.ps1 (WPF, EMT.cmd). This classic window has the same functions
    and the same logic (src\EMT.Core.ps1). Start it with EMT-Classic.cmd only if the usual window does not operate.
    The classic window will be removed in a later version.
    Tab "Message search": find a message in the message tracking log, see the delivery status of each recipient,
    see why a message did not arrive, examine the mailbox of a recipient. Read-only. Seconds.
    Tab "Delete": find the copies of a message in the mailboxes and delete them.
    All Exchange work runs in a background runspace. The window does not freeze.
    The user interface languages are the files in the folder "Lang" (*.lng).
    See README.md for requirements, roles and the procedure.

.PARAMETER Server
    The FQDN of the Exchange server. If you do not set it, the tool uses the last value.

.PARAMETER Language
    The language code of a file in "Lang" (for example ru or en). If you do not set it,
    the tool uses the last value, or the Windows language.

.EXAMPLE
    .\ExchangeMessageTrail.ps1 -Server exch01.contoso.local -Language en
#>
[CmdletBinding()]
param([string] $Server, [string] $Language)

$ErrorActionPreference = 'Stop'

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Start the classic window in STA mode: powershell.exe -STA -File ExchangeMessageTrail.Classic.ps1 (or use EMT-Classic.cmd).'
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

# The console is hidden: show a start error in a window. The language files are not loaded yet: English text.
trap {
    [void][System.Windows.Forms.MessageBox]::Show(
        "Exchange Message Trail could not start.`r`n`r`n$($_.Exception.Message)`r`n`r`n$($_.InvocationInfo.PositionMessage)",
        'Exchange Message Trail', 'OK', 'Error')
    break
}

# Hide the PowerShell console window, but only if it belongs to EMT alone.
# One process in the console: the console was made for EMT (EMT.cmd, a shortcut, "Run").
# More processes: the user started EMT from an open console. That console stays visible.
if (-not ('EmtConsole' -as [type])) {
    Add-Type -Namespace '' -Name 'EmtConsole' -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
[DllImport("kernel32.dll")] public static extern uint GetConsoleProcessList(uint[] list, uint count);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int cmd);
'@
}
$consoleWindow = [EmtConsole]::GetConsoleWindow()
if ($consoleWindow -ne [IntPtr]::Zero -and [EmtConsole]::GetConsoleProcessList((New-Object 'System.UInt32[]' 4), 4) -eq 1 -and
    [Environment]::CommandLine -match '\s-File\s') {
    [void][EmtConsole]::ShowWindow($consoleWindow, 0)    # 0 = SW_HIDE
}

$CorePath = Join-Path $PSScriptRoot 'src\EMT.Core.ps1'
. $CorePath

#region ---------- Paths, settings, shared state ----------

$AppDir       = Join-Path $env:USERPROFILE 'ExchangeMessageTrail'
$SettingsPath = Join-Path $env:APPDATA 'ExchangeMessageTrail\settings.json'
$AuditPath    = Join-Path $AppDir 'audit.csv'
$LogPath      = Join-Path $AppDir ('logs\EMT_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
foreach ($d in (Split-Path $SettingsPath), (Split-Path $LogPath)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# The older version (MailCleaner) kept its files in other folders. Copy them once: the audit history stays.
$script:Migrated = @()
$oldFiles = @(
    @{ From = (Join-Path $env:APPDATA 'MailCleaner\settings.json'); To = $SettingsPath }
    @{ From = (Join-Path $env:USERPROFILE 'MailCleaner\audit.csv');  To = $AuditPath }
)
foreach ($f in $oldFiles) {
    if ((Test-Path $f.From) -and -not (Test-Path $f.To)) {
        try { Copy-Item -Path $f.From -Destination $f.To; $script:Migrated += $f.From } catch { }
    }
}

$Sync = [hashtable]::Synchronized(@{
    Log      = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    Progress = 0
    Status   = ''
    Cancel   = $false
    Lang     = 'en'      # the window and the background runspace read the language here
})
$global:EmtSync = $Sync

$script:Caps    = $null      # capabilities after Connect
$script:Trace   = $null      # result of the last trace
$script:MailboxCheck = $null # result of the last mailbox check (MessageId, Mailbox, Findings)
$script:Result  = $null      # result of the last search on the tab "Delete"
$script:Check   = $null      # result of the last verification
$script:Job     = $null      # running background operation
$script:Loading = $false     # true while the UI fills controls (no invalidation)
$script:ConnectedServer = ''
$script:DelFromTrace    = ''       # MessageId when the tab "Delete" has the data of a traced message
$script:CheckForMessage = ''       # MessageId of the running mailbox check
$script:ReportByMsg     = @{}      # MessageId -> the report rows of the message (Show-EmtTraceViews)
$script:ResolveNext     = $null    # the action after the address book search
$script:PendingCleanup  = $null    # names confirmed for cleanup; the timer starts the removal
$script:CleanupNames    = $null
$script:GridNeedsSize   = @{}      # grid -> true: set the initial column widths after the next binding
$script:GridHeaders     = @{}      # grid -> ordered map "property" -> "text key" of the column header
$script:GridHidden      = @{}      # grid -> names of the columns that are data for the row colors only
$script:Deleted         = ''       # last delete on the current result: '', 'Soft' or 'Hard'
$script:Localized       = New-Object System.Collections.Generic.List[object]   # control + text key
$script:LocalizedTips   = New-Object System.Collections.Generic.List[object]   # control + tooltip key
$script:LanguageCodes   = @()      # language codes in the order of the language list

function Read-EmtSettings {
    $s = @{}
    if (Test-Path $SettingsPath) {
        try {
            $j = Get-Content $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $j.PSObject.Properties) { $s[$p.Name] = $p.Value }
        } catch { }
    }
    $s
}

function Save-EmtSettings {
    # Senders and subject are not kept: they can be confidential.
    # Other keys of the file stay (for example Theme and WpfWindow of the usual window).
    $s = Read-EmtSettings
    $s.Server          = $tbServer.Text.Trim()
    $s.Language        = [string]$Sync.Lang
    $s.IncludeForwards = $chkForwards.Checked
    $s.AllMailboxes    = $chkAll.Checked
    $s.TimeoutMinutes  = [int]$nudTimeout.Value
    # The window: the normal size and position, also when the window is maximized or minimized.
    $b = if ($form.WindowState -eq 'Normal') { $form.Bounds } else { $form.RestoreBounds }
    if ($b.Width -gt 0 -and $b.Height -gt 0) {
        $s.Window = [ordered]@{ X = $b.X; Y = $b.Y; Width = $b.Width; Height = $b.Height
                                Maximized = ($form.WindowState -eq 'Maximized') }
    }
    try { $s | ConvertTo-Json | Set-Content -Path $SettingsPath -Encoding UTF8 } catch { }
}

#endregion

#region ---------- Background runspace ----------

$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$iss.ExecutionPolicy = 'Bypass'
$Runspace = [runspacefactory]::CreateRunspace($iss)
$Runspace.ThreadOptions = 'ReuseThread'
$Runspace.Open()
$Runspace.SessionStateProxy.SetVariable('EmtSync', $Sync)

$OpScript = @'
param($CorePath, $Name, $Params)
. $CorePath
Invoke-EmtOperation -Name $Name -Params $Params
'@

function Start-EmtOperation {
    param(
        [Parameter(Mandatory)] [string] $Name,
        [hashtable] $Params = @{},
        [Parameter(Mandatory)] [scriptblock] $OnSuccess,
        [string] $StatusText = (T 'status.running')
    )
    if ($script:Job) { return }
    $Sync.Cancel   = $false
    $Sync.Progress = 0
    $Sync.Status   = $StatusText
    $ps = [powershell]::Create()
    $ps.Runspace = $Runspace
    [void]$ps.AddScript($OpScript).AddArgument($CorePath).AddArgument($Name).AddArgument($Params)
    $script:Job = [pscustomobject]@{
        PS = $ps; Handle = $ps.BeginInvoke(); Name = $Name; Params = $Params; OnSuccess = $OnSuccess
    }
    Set-UiState
}

function Complete-EmtOperation {
    $job = $script:Job
    $script:Job = $null
    $result = $null
    $err = $null
    try {
        $out = $job.PS.EndInvoke($job.Handle)
        $result = if ($out.Count -eq 1) { $out[0] } else { @($out) }
    } catch {
        $err = $_.Exception
        while ($err -is [System.Management.Automation.MethodInvocationException] -and $err.InnerException) {
            $err = $err.InnerException
        }
        # "-ErrorAction Stop" wraps the real error: show the original message.
        if ($err -is [System.Management.Automation.ActionPreferenceStopException] -and $err.ErrorRecord) {
            $err = $err.ErrorRecord.Exception
        }
    }
    foreach ($e in $job.PS.Streams.Error) { Write-EmtLog ([string]$e) 'Warn' }
    $job.PS.Dispose()
    $progress.Value = 0
    $lblStatus.Text = ''
    $tips.SetToolTip($lblStatus, '')

    if ($err) {
        if ($Sync.Cancel) {
            Write-EmtLog (T 'log.opCancelled' $job.Name) 'Warn'
        } else {
            Write-EmtLog (T 'log.opFailed' $job.Name $err.Message) 'Error'
            [void][System.Windows.Forms.MessageBox]::Show($form, $err.Message, (T 'dlg.error'),
                'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        # A delete that stops can be partial: the audit must show it.
        if ($job.Name -in 'SoftDelete', 'HardDelete') {
            try {
                Write-EmtAudit -Path $AuditPath -Record @{
                    Server = $script:ConnectedServer; Action = $job.Name
                    SearchName = [string]$job.Params.SearchName; Query = [string]$job.Params.Query
                    Mailboxes = @($job.Params.Mailboxes).Count
                    Result = $(if ($Sync.Cancel) { 'Cancelled' } else { 'Failed' }); Details = $err.Message
                }
            } catch { Write-EmtLog (T 'log.auditFailed' $_.Exception.Message) 'Error' }
        }
    } else {
        try { & $job.OnSuccess $result }
        catch { Write-EmtLog (T 'log.handlerFailed' $_.Exception.Message) 'Error' }
    }
    Set-UiState
}

#endregion

#region ---------- Localization of controls ----------

function Set-EmtLoc {
    <# Registers a control for the language change and sets its text now. Returns the control. #>
    param($Control, [string] $Key)
    $script:Localized.Add([pscustomobject]@{ Control = $Control; Key = $Key })
    $Control.Text = T $Key
    $Control
}

function Set-EmtTip {
    param($Control, [string] $Key)
    $script:LocalizedTips.Add([pscustomobject]@{ Control = $Control; Key = $Key })
    $tips.SetToolTip($Control, (T $Key))
}

function Update-EmtLanguage {
    <# Applies the current language to all controls, headers and result views. #>
    $form.Text = T 'app.title'
    foreach ($x in $script:Localized)     { $x.Control.Text = T $x.Key }
    foreach ($x in $script:LocalizedTips) { $tips.SetToolTip($x.Control, (T $x.Key)) }
    Update-EmtCriteriaLabels
    Update-EmtConnIndicator
    Update-EmtDeleteSource
    if ($script:Trace) { Show-EmtTraceViews $script:Trace } else { Clear-EmtTraceViews }
    if ($script:Result) { Show-EmtResultViews $script:Result } else { Clear-EmtResultViews }
    if ($script:Check)  { Show-EmtCheckView $script:Check }
    Update-EmtTabTitles
    Set-UiState
}

#endregion

#region ---------- UI helpers ----------

function New-EmtLabel([string] $Key) {
    $l = New-Object System.Windows.Forms.Label
    $l.AutoSize = $true; $l.Anchor = 'Left'; $l.Margin = '3,7,3,3'
    if ($Key) { [void](Set-EmtLoc $l $Key) }
    $l
}

function New-EmtTextBox {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Dock = 'Fill'; $t.Margin = '3,4,12,3'
    $t
}

function New-EmtButton([string] $Key, [int] $Width = 150) {
    # AutoSize with a minimum width: a longer translation makes the button wider, not cut.
    $b = New-Object System.Windows.Forms.Button
    $b.AutoSize = $true; $b.AutoSizeMode = 'GrowOnly'
    $b.MinimumSize = New-Object System.Drawing.Size($Width, 28)
    $b.Height = 28; $b.Margin = '3,3,6,3'
    if ($Key) { [void](Set-EmtLoc $b $Key) }
    $b
}

function Set-EmtAccentButton([System.Windows.Forms.Button] $Button) {
    # The main action of a tab: bold text on a light green background.
    # A disabled button gets the usual grey background: green must not show an unavailable action.
    $Button.Font = New-Object System.Drawing.Font($Button.Font, [System.Drawing.FontStyle]::Bold)
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(110, 170, 110)
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(200, 235, 200)
    $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(175, 220, 175)
    $paint = {
        param($s, $e)
        if ($s.Enabled) {
            $s.BackColor = [System.Drawing.Color]::FromArgb(222, 245, 222)
            $s.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(110, 170, 110)
        } else {
            $s.BackColor = [System.Drawing.SystemColors]::Control
            $s.FlatAppearance.BorderColor = [System.Drawing.SystemColors]::ControlDark
        }
    }
    $Button.Add_EnabledChanged($paint)
    & $paint $Button $null
}

function New-EmtIndicator([string] $Key) {
    <# A round color mark and a text. Set-EmtIndicator sets the color. #>
    $dot = New-Object System.Windows.Forms.Label
    $dot.AutoSize = $false; $dot.Size = '12,12'; $dot.Margin = '6,5,2,3'; $dot.Anchor = 'Left'
    $dot.Tag = [System.Drawing.Color]::Gray
    $dot.Add_Paint({
        param($s, $e)
        $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]$s.Tag)
        $e.Graphics.FillEllipse($brush, 1, 1, $s.Width - 3, $s.Height - 3)
        $brush.Dispose()
    })
    $text = New-EmtLabel $Key
    $text.Margin = '0,3,4,3'
    [pscustomobject]@{ Dot = $dot; Text = $text }
}

function Set-EmtIndicator($Indicator, [System.Drawing.Color] $Color) {
    $Indicator.Dot.Tag = $Color
    $Indicator.Dot.Invalidate()
}

function New-EmtCheckBox([string] $Key) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.AutoSize = $true; $c.Margin = '3,7,12,3'
    [void](Set-EmtLoc $c $Key)
    $c
}

function New-EmtDatePicker {
    <#
        The date and time in the format of the Windows regional settings, not of the interface language.
        Thus a user in the USA sees 11/12/2026 12:00 AM, and a user in Europe sees 12.11.2026 00:00.
        The width comes from the longest date of the format (the AM/PM text and the separators can be longer).
    #>
    $f = [Globalization.CultureInfo]::CurrentCulture.DateTimeFormat
    $d = New-Object System.Windows.Forms.DateTimePicker
    $d.Format = 'Custom'; $d.CustomFormat = '{0} {1}' -f $f.ShortDatePattern, $f.ShortTimePattern
    $sample = ([datetime]'2026-12-28 23:58').ToString($d.CustomFormat)
    $d.Width = [math]::Max(135, [System.Windows.Forms.TextRenderer]::MeasureText($sample, $form.Font).Width + 44)
    $d.Margin = '3,4,6,3'
    $d
}

function New-EmtSummaryLabel {
    $l = New-Object System.Windows.Forms.Label
    $l.Dock = 'Top'; $l.Height = 28; $l.Padding = '10,6,6,0'; $l.AutoEllipsis = $true
    $l.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
    $l
}

function New-EmtPeriodPanel($StartPicker, $EndPicker) {
    <# The period in one line: start - end and the quick period buttons. Returns the panel and the buttons. #>
    $today     = New-EmtButton 'crit.today' 72
    $yesterday = New-EmtButton 'crit.yesterday' 64
    $week      = New-EmtButton 'crit.week' 64
    foreach ($b in $today, $yesterday, $week) {
        $b.MinimumSize = New-Object System.Drawing.Size($b.MinimumSize.Width, 24); $b.Height = 24; $b.Margin = '0,3,4,3'
    }
    $today.Margin = '8,3,4,3'
    $dash = New-EmtLabel ''
    $dash.Text = [string][char]0x2014; $dash.Margin = '0,7,3,3'
    $p = New-Object System.Windows.Forms.FlowLayoutPanel
    $p.AutoSize = $true; $p.WrapContents = $false; $p.Margin = '0,0,0,0'
    $p.Controls.AddRange(@($StartPicker, $dash, $EndPicker, $today, $yesterday, $week))
    [pscustomobject]@{ Panel = $p; Today = $today; Yesterday = $yesterday; Week = $week }
}

function New-EmtCriteriaTable([int] $Rows) {
    <#
        The criteria: two pairs "label + field" in each row.
        The two field columns have the same width. Thus all fields have the same size.
    #>
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.Dock = 'Top'; $t.AutoSize = $true; $t.ColumnCount = 4; $t.RowCount = $Rows
    foreach ($i in 1, 2) {
        [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
        [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
    }
    $t
}

function New-EmtActionRow($Left, $Button) {
    <# The last row of the criteria: controls on the left, the main button on the right edge of the fields. #>
    $r = New-Object System.Windows.Forms.TableLayoutPanel
    $r.Dock = 'Fill'; $r.AutoSize = $true; $r.ColumnCount = 3; $r.RowCount = 1; $r.Margin = '0,0,0,0'
    [void]$r.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$r.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$r.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    $Left.Anchor = 'Left'
    $Button.Anchor = 'Right'; $Button.Margin = '12,3,12,3'    # the right margin of a text field
    $r.Controls.Add($Left, 0, 0)
    $r.Controls.Add($Button, 2, 0)
    $r
}

function New-EmtCriteriaGroup([string] $Key, $Table) {
    $g = New-Object System.Windows.Forms.GroupBox
    [void](Set-EmtLoc $g $Key)
    $g.Dock = 'Top'; $g.AutoSize = $true; $g.AutoSizeMode = 'GrowAndShrink'; $g.Padding = '8,4,8,6'
    $g.Controls.Add($Table)
    $g
}

function New-EmtGrid([scriptblock] $RowColor) {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.MultiSelect = $false
    # Auto-size modes lock the column width. Use None: the user can drag the column borders.
    $g.AutoSizeColumnsMode = 'None'
    $g.AllowUserToResizeColumns = $true
    $g.AllowUserToOrderColumns = $true
    # The row height and the header height are fixed. Only the column width can change.
    $g.AllowUserToResizeRows = $false
    $g.ColumnHeadersHeightSizeMode = 'DisableResizing'
    $g.ClipboardCopyMode = 'EnableWithAutoHeaderText'    # Ctrl+C copies the selected rows with headers
    $g.BackgroundColor = [System.Drawing.SystemColors]::Window
    $g.Tag = $RowColor
    $g.Add_DataBindingComplete({
        param($s, $e)
        foreach ($c in $s.Columns) {
            if ($c.ValueType -eq [datetime]) { $c.DefaultCellStyle.Format = T 'fmt.datetime' }
            $c.AutoSizeMode = 'None'
            $c.Resizable = 'True'
            if ($script:GridHidden.ContainsKey($s) -and $c.Name -in $script:GridHidden[$s]) { $c.Visible = $false }
        }
        Update-EmtGridHeaders $s
        Resize-EmtGridColumns $s
    })
    # A grid on a hidden tab cannot measure its cells. Size it when the tab opens.
    $g.Add_VisibleChanged({ param($s, $e) Resize-EmtGridColumns $s })
    $g.Add_CellFormatting({
        param($s, $e)
        if (-not $s.Tag -or $e.RowIndex -lt 0) { return }
        $view = $s.Rows[$e.RowIndex].DataBoundItem
        if ($view) {
            $color = & $s.Tag $view.Row
            if ($color) { $e.CellStyle.BackColor = $color }
        }
    })
    $g
}

function Update-EmtGridHeaders($Grid) {
    <# The DataTable columns have the property names. The headers come from the current language. #>
    if (-not $script:GridHeaders.ContainsKey($Grid)) { return }
    $map = $script:GridHeaders[$Grid]
    foreach ($c in $Grid.Columns) {
        if ($map.Contains($c.Name) -and $map[$c.Name]) { $c.HeaderText = T $map[$c.Name] }
    }
}

function Resize-EmtGridColumns($Grid) {
    <#
        Initial column width from the content (header and all rows), 60–450 px.
        The last column gets the free space. It runs once after new data, and only
        when the grid is visible: a sort must keep the widths that the user set.
    #>
    if (-not ($script:GridNeedsSize.ContainsKey($Grid) -and $script:GridNeedsSize[$Grid])) { return }
    if (-not $Grid.Visible -or -not $Grid.Columns.Count) { return }
    $script:GridNeedsSize[$Grid] = $false
    $Grid.AutoResizeColumns('AllCells')
    $used = 0
    $last = $null
    foreach ($c in $Grid.Columns) {
        if (-not $c.Visible) { continue }
        if ($c.Width -gt 450) { $c.Width = 450 }
        if ($c.Width -lt 60)  { $c.Width = 60 }
        $used += $c.Width
        $last = $c
    }
    $free = $Grid.ClientSize.Width - $used - 4
    if ($last -and $free -gt 0) { $last.Width += $free }
}

function Set-EmtGridData {
    <#
        Binds objects to a grid. $Columns: ordered map "Property" -> "header text key".
        $Hidden: columns that only give data for the row colors (not shown).
    #>
    param($Grid, [object[]] $Items = @(), [System.Collections.Specialized.OrderedDictionary] $Columns, [string[]] $Hidden = @())
    $items = @($Items | Where-Object { $_ -ne $null })
    $dt = New-Object System.Data.DataTable
    foreach ($prop in $Columns.Keys) {
        $sample = $items | ForEach-Object { Get-EmtProp $_ $prop } | Where-Object { $_ -ne $null } | Select-Object -First 1
        $type = [string]
        if ($sample -is [int] -or $sample -is [long]) { $type = [long] }
        elseif ($sample -is [double] -or $sample -is [decimal]) { $type = [double] }
        elseif ($sample -is [datetime]) { $type = [datetime] }
        [void]$dt.Columns.Add($prop, $type)
    }
    foreach ($it in $items) {
        $row = $dt.NewRow()
        foreach ($prop in $Columns.Keys) {
            $v = Get-EmtProp $it $prop
            if ($null -eq $v) { $row[$prop] = [DBNull]::Value }
            elseif ($dt.Columns[$prop].DataType -eq [string]) {
                $row[$prop] = if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { (@($v) -join '; ') } else { [string]$v }
            }
            else { $row[$prop] = $v }
        }
        $dt.Rows.Add($row)
    }
    $script:GridHeaders[$Grid] = $Columns
    $script:GridHidden[$Grid] = $Hidden
    $script:GridNeedsSize[$Grid] = $true
    $Grid.DataSource = $dt
    Update-EmtGridHeaders $Grid
}

function Get-EmtSelectedRow($Grid) {
    <# The DataRow of the selected row of a grid, or $null. #>
    if ($Grid.SelectedRows.Count -ne 1) { return $null }
    $view = $Grid.SelectedRows[0].DataBoundItem
    if ($view) { $view.Row } else { $null }
}

function Show-EmtConfirm {
    <# Confirmation dialog. With -RequireWord the OK button is enabled only after the user types the word. #>
    param([string] $Title, [string] $Text, [string] $OkText, [string] $RequireWord)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title; $dlg.Size = '700,560'; $dlg.StartPosition = 'CenterParent'
    $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false; $dlg.Font = $form.Font; $dlg.ShowIcon = $false

    $body = New-Object System.Windows.Forms.TextBox
    $body.Multiline = $true; $body.ReadOnly = $true; $body.ScrollBars = 'Vertical'; $body.Dock = 'Fill'
    $body.Text = $Text -replace "`r?`n", "`r`n"
    $body.BackColor = [System.Drawing.SystemColors]::Window

    $bottom = New-Object System.Windows.Forms.FlowLayoutPanel
    $bottom.Dock = 'Bottom'; $bottom.AutoSize = $true; $bottom.Padding = '6,6,6,6'; $bottom.FlowDirection = 'LeftToRight'

    $ok = New-EmtButton '' 160
    $ok.Text = $OkText
    $ok.DialogResult = 'OK'
    $cancel = New-EmtButton '' 120
    $cancel.Text = T 'dlg.cancel'
    $cancel.DialogResult = 'Cancel'
    $dlg.CancelButton = $cancel

    if ($RequireWord) {
        $prompt = New-EmtLabel ''
        $prompt.Text = T 'dlg.typeWord' $RequireWord
        $bottom.Controls.Add($prompt)
        $word = New-Object System.Windows.Forms.TextBox
        $word.Width = 140; $word.Margin = '3,4,12,3'
        $bottom.Controls.Add($word)
        $ok.Enabled = $false
        $word.Add_TextChanged({ $ok.Enabled = ($word.Text -ceq $RequireWord) }.GetNewClosure())
    }
    $bottom.Controls.Add($ok)
    $bottom.Controls.Add($cancel)

    $dlg.Controls.Add($body)
    $dlg.Controls.Add($bottom)
    $dlg.Add_Shown({ $cancel.Focus() }.GetNewClosure())
    $answer = $dlg.ShowDialog($form)
    $dlg.Dispose()
    $answer -eq [System.Windows.Forms.DialogResult]::OK
}

function Add-EmtLogLine([string] $Line) {
    $log.SelectionStart = $log.TextLength
    $log.SelectionLength = 0
    $log.SelectionColor = switch -Regex ($Line) {
        '\[ERROR\]' { [System.Drawing.Color]::Firebrick; break }
        '\[WARN'    { [System.Drawing.Color]::DarkOrange; break }
        '\[OK'      { [System.Drawing.Color]::ForestGreen; break }
        default     { [System.Drawing.SystemColors]::WindowText }
    }
    $log.AppendText($Line + [Environment]::NewLine)
    $log.ScrollToCaret()
    try { Add-Content -Path $LogPath -Value $Line -Encoding UTF8 } catch { }
}

function Export-EmtSets([string] $Stem, [System.Collections.Specialized.OrderedDictionary] $Sets) {
    <# Saves each set as <folder>\<stem>-<name>.csv. The user selects the folder. #>
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = T 'dlg.exportFolder'
    $dlg.SelectedPath = $AppDir
    if ($dlg.ShowDialog($form) -ne 'OK') { return }
    $base = Join-Path $dlg.SelectedPath $Stem
    foreach ($k in $Sets.Keys) {
        if (@($Sets[$k]).Count) {
            $Sets[$k] | Export-Csv -Path "$base-$k.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
        }
    }
    Write-EmtLog (T 'log.exported' $base) 'Ok'
}

#endregion

#region ---------- Language (before the layout: the controls read the texts) ----------

$settings = Read-EmtSettings
$available = @(Get-EmtLanguages)
$script:LanguageCodes = @($available | ForEach-Object Code)
$wanted = if ($Language) { $Language.ToLowerInvariant() }
          elseif ($settings.ContainsKey('Language') -and $settings.Language) { ([string]$settings.Language).ToLowerInvariant() }
          else { Get-EmtDefaultLanguage }
$Sync.Lang = if ($wanted -in $script:LanguageCodes) { $wanted } elseif ('en' -in $script:LanguageCodes) { 'en' } else { @($script:LanguageCodes + 'en')[0] }

#endregion

#region ---------- Form layout: frame and connection ----------

$form = New-Object System.Windows.Forms.Form
$form.Text = T 'app.title'
$form.Size = '1280,880'
$form.MinimumSize = '1100,720'
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$tips = New-Object System.Windows.Forms.ToolTip

$colorRed    = [System.Drawing.Color]::MistyRose
$colorYellow = [System.Drawing.Color]::LightYellow
$colorGreen  = [System.Drawing.Color]::Honeydew
$colorBlue   = [System.Drawing.Color]::AliceBlue
$colorGrey   = [System.Drawing.Color]::WhiteSmoke
# Colors of the indicators in the bottom line.
$dotGreen = [System.Drawing.Color]::FromArgb(46, 160, 67)
$dotGrey  = [System.Drawing.Color]::FromArgb(160, 160, 160)

# --- Connection bar (common for both tabs) ---
$pnlConn = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlConn.Dock = 'Top'; $pnlConn.AutoSize = $true; $pnlConn.Padding = '6,6,6,0'; $pnlConn.WrapContents = $true
$tbServer = New-Object System.Windows.Forms.TextBox
$tbServer.Width = 240; $tbServer.Margin = '3,4,6,3'
$btnConnect = New-EmtButton 'conn.connect' 140
$nudTimeout = New-Object System.Windows.Forms.NumericUpDown
$nudTimeout.Minimum = 5; $nudTimeout.Maximum = 240; $nudTimeout.Value = 60; $nudTimeout.Width = 60; $nudTimeout.Margin = '3,4,12,3'
$cmbLanguage = New-Object System.Windows.Forms.ComboBox
$cmbLanguage.DropDownStyle = 'DropDownList'; $cmbLanguage.Width = 120; $cmbLanguage.Margin = '3,4,12,3'
foreach ($l in $available) { [void]$cmbLanguage.Items.Add($l.Name) }
$pnlConn.Controls.AddRange(@(
    (New-EmtLabel 'conn.server'), $tbServer, $btnConnect,
    (New-EmtLabel 'conn.timeout'), $nudTimeout,
    (New-EmtLabel 'conn.language'), $cmbLanguage))

# --- Main tabs ---
$mainTabs = New-Object System.Windows.Forms.TabControl
$mainTabs.Dock = 'Fill'
$pgTrace  = New-Object System.Windows.Forms.TabPage
$pgDelete = New-Object System.Windows.Forms.TabPage
[void](Set-EmtLoc $pgTrace 'main.trace')
[void](Set-EmtLoc $pgDelete 'main.delete')
$pgTrace.Padding = '4,4,4,4'; $pgDelete.Padding = '4,4,4,4'
$mainTabs.TabPages.AddRange(@($pgTrace, $pgDelete))

#endregion

#region ---------- Form layout: tab "Trace" ----------

$tbTrSenders   = New-EmtTextBox
$tbTrRecipient = New-EmtTextBox
$tbTrSubject   = New-EmtTextBox
$tbTrMessageId = New-EmtTextBox
$dtTrStart     = New-EmtDatePicker
$dtTrEnd       = New-EmtDatePicker
$trPeriod      = New-EmtPeriodPanel $dtTrStart $dtTrEnd
$btnTrFind     = New-EmtButton 'crit.find' 130
Set-EmtAccentButton $btnTrFind

# Row 0: who.  Row 1: what.  Row 2: when and Find.  The tab "Delete" uses the same order.
$tlpTr = New-EmtCriteriaTable 3
$tlpTr.Controls.Add((New-EmtLabel 'crit.senders'), 0, 0);   $tlpTr.Controls.Add($tbTrSenders, 1, 0)
$tlpTr.Controls.Add((New-EmtLabel 'crit.recipient'), 2, 0); $tlpTr.Controls.Add($tbTrRecipient, 3, 0)
$tlpTr.Controls.Add((New-EmtLabel 'crit.subject'), 0, 1);   $tlpTr.Controls.Add($tbTrSubject, 1, 1)
$tlpTr.Controls.Add((New-EmtLabel 'crit.messageId'), 2, 1); $tlpTr.Controls.Add($tbTrMessageId, 3, 1)
$tlpTr.Controls.Add((New-EmtLabel 'crit.period'), 0, 2)
$trRowFind = New-EmtActionRow $trPeriod.Panel $btnTrFind
$tlpTr.Controls.Add($trRowFind, 1, 2)
$tlpTr.SetColumnSpan($trRowFind, 3)
$grpTr = New-EmtCriteriaGroup 'trcrit.group' $tlpTr

$lblTrSummary = New-EmtSummaryLabel
$lblTrSummary.Text = T 'trsum.none'

# Messages: red = a recipient did not get the message, yellow = a delivery is not complete.
$gridTrMessages = New-EmtGrid {
    param($r)
    if ([long]$r['Problems'] -gt 0) { $colorRed } elseif ([long]$r['Pending'] -gt 0) { $colorYellow }
}
# Recipients: the color of the delivery status.
$gridTrRecipients = New-EmtGrid {
    param($r)
    switch ([string]$r['StatusCode']) {
        'Delivered' { $colorGreen }  'External' { $colorBlue }
        'Failed'    { $colorRed }    'Dropped'  { $colorRed }    'GroupBlocked' { $colorRed }
        'Deferred'  { $colorYellow } 'Forwarded' { $colorYellow } default { $colorGrey }
    }
}
$gridTrEvents = New-EmtGrid {
    param($r)
    if ([string]$r['EventId'] -in 'FAIL', 'DROP', 'POISONMESSAGE') { $colorRed } elseif ([string]$r['EventId'] -eq 'DEFER') { $colorYellow }
}
$gridTrMailbox = New-EmtGrid {
    param($r)
    switch ([string]$r['SeverityCode']) { 'Alert' { $colorRed } 'Warn' { $colorYellow } default { $null } }
}

$trDetail = New-Object System.Windows.Forms.TabControl
$trDetail.Dock = 'Fill'
$trPages = [ordered]@{ 'recipients' = $gridTrRecipients; 'events' = $gridTrEvents; 'mailbox' = $gridTrMailbox }
$trPage = @{}
foreach ($k in $trPages.Keys) {
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text = T "trtab.$k"
    $p.Controls.Add($trPages[$k])
    $trDetail.TabPages.Add($p)
    $trPage[$k] = $p
}
# A note under the mailbox findings: the tracking log ends at the mailbox.
$lblTrMailboxNote = New-Object System.Windows.Forms.Label
$lblTrMailboxNote.Dock = 'Bottom'; $lblTrMailboxNote.Height = 22; $lblTrMailboxNote.Padding = '6,4,6,0'
$lblTrMailboxNote.ForeColor = [System.Drawing.SystemColors]::GrayText
[void](Set-EmtLoc $lblTrMailboxNote 'trmail.note')
$trPage['mailbox'].Controls.Add($lblTrMailboxNote)

$trSplit = New-Object System.Windows.Forms.SplitContainer
$trSplit.Dock = 'Fill'; $trSplit.Orientation = 'Horizontal'
$trSplit.Panel1.Controls.Add($gridTrMessages)
$trSplit.Panel2.Controls.Add($trDetail)

$pnlTrActions = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlTrActions.Dock = 'Bottom'; $pnlTrActions.AutoSize = $true; $pnlTrActions.WrapContents = $false; $pnlTrActions.Padding = '0,4,0,0'
$btnTrCheck  = New-EmtButton 'act.checkMailbox' 190
$btnTrDelete = New-EmtButton 'act.deleteThis' 170
$btnTrDelete.ForeColor = [System.Drawing.Color]::Firebrick
$btnTrExport = New-EmtButton 'act.export' 115
$btnTrExport.Margin = '18,3,6,3'
$pnlTrActions.Controls.AddRange(@($btnTrCheck, $btnTrDelete, $btnTrExport))

$pgTrace.Controls.Add($trSplit)
$pgTrace.Controls.Add($pnlTrActions)
$pgTrace.Controls.Add($lblTrSummary)
$pgTrace.Controls.Add($grpTr)

#endregion

#region ---------- Form layout: tab "Delete" ----------

$tbSenders    = New-EmtTextBox
$tbRecipient  = New-EmtTextBox
$tbSubject    = New-EmtTextBox
$tbMessageId  = New-EmtTextBox
$tbAttachment = New-EmtTextBox
$tbKeywords   = New-EmtTextBox
$dtStart      = New-EmtDatePicker
$dtEnd        = New-EmtDatePicker
$delPeriod    = New-EmtPeriodPanel $dtStart $dtEnd

$chkForwards = New-EmtCheckBox 'crit.forwards'
$chkForwards.Checked = $true
$chkOnly = New-EmtCheckBox 'crit.only'
$chkOnly.Enabled = $false
$chkAll = New-EmtCheckBox 'crit.all'
$pnlOptions = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlOptions.AutoSize = $true; $pnlOptions.WrapContents = $false; $pnlOptions.Margin = '0,0,0,0'
$pnlOptions.Controls.AddRange(@($chkForwards, $chkOnly, $chkAll))
$btnSearch = New-EmtButton 'crit.findCopies' 130
Set-EmtAccentButton $btnSearch

# The same order as on the tab "Trace": who, what, content, when. The last row: options and Find copies.
$tlp = New-EmtCriteriaTable 5
$tlp.Controls.Add((New-EmtLabel 'crit.senders'), 0, 0);    $tlp.Controls.Add($tbSenders, 1, 0)
$tlp.Controls.Add((New-EmtLabel 'crit.recipient'), 2, 0);  $tlp.Controls.Add($tbRecipient, 3, 0)
$tlp.Controls.Add((New-EmtLabel 'crit.subject'), 0, 1);    $tlp.Controls.Add($tbSubject, 1, 1)
$tlp.Controls.Add((New-EmtLabel 'crit.messageId'), 2, 1);  $tlp.Controls.Add($tbMessageId, 3, 1)
$tlp.Controls.Add((New-EmtLabel 'crit.attachment'), 0, 2); $tlp.Controls.Add($tbAttachment, 1, 2)
$tlp.Controls.Add((New-EmtLabel 'crit.keywords'), 2, 2);   $tlp.Controls.Add($tbKeywords, 3, 2)
$tlp.Controls.Add((New-EmtLabel 'crit.period'), 0, 3);     $tlp.Controls.Add($delPeriod.Panel, 1, 3)
$tlp.SetColumnSpan($delPeriod.Panel, 3)
$delRowFind = New-EmtActionRow $pnlOptions $btnSearch
$tlp.Controls.Add($delRowFind, 1, 4)
$tlp.SetColumnSpan($delRowFind, 3)
$grpCrit = New-EmtCriteriaGroup 'crit.group' $tlp

function Update-EmtCriteriaLabels {
    <#
        The label columns get the same width on both tabs. Thus the fields have the same size and position
        on "Trace" and on "Delete". Call it again after a language change: the texts have a new length.
    #>
    foreach ($col in 0, 2) {
        $labels = foreach ($t in $tlpTr, $tlp) {
            foreach ($row in 0..($t.RowCount - 1)) {
                $c = $t.GetControlFromPosition($col, $row)
                if ($c -is [System.Windows.Forms.Label]) { $c }
            }
        }
        $w = 0
        foreach ($l in $labels) { $l.MinimumSize = [System.Drawing.Size]::Empty; $w = [math]::Max($w, $l.PreferredWidth) }
        foreach ($l in $labels) { $l.MinimumSize = New-Object System.Drawing.Size($w, 0) }
    }
}

# Where the criteria come from: a traced message (confirmed by the tracking log) or the user.
$lblDelSource = New-Object System.Windows.Forms.Label
$lblDelSource.Dock = 'Top'; $lblDelSource.Height = 22; $lblDelSource.Padding = '10,4,6,0'; $lblDelSource.AutoEllipsis = $true

$lblSummary = New-EmtSummaryLabel
$lblSummary.Text = T 'sum.none'

# Row colors use the DataTable column names (the property names), not the header texts.
$gridMessages   = New-EmtGrid { param($r) if ([string]$r['Note']) { $colorYellow } }
$gridTracking   = New-EmtGrid $null
$gridRecipients = New-EmtGrid { param($r) if ([string]$r['KindCode'] -eq 'External') { $colorRed } }
$gridFound      = New-EmtGrid $null
$gridMissing    = New-EmtGrid $null
$gridCheck      = New-EmtGrid {
    param($r)
    if ([long]$r['Visible'] -lt 0) { $colorYellow }
    elseif ([long]$r['Visible'] -gt 0) { $colorRed }
    else { $colorGreen }
}

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$pages = [ordered]@{
    'messages'   = $gridMessages
    'tracking'   = $gridTracking
    'recipients' = $gridRecipients
    'found'      = $gridFound
    'missing'    = $gridMissing
    'check'      = $gridCheck
}
$tabPage = @{}
foreach ($k in $pages.Keys) {
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text = T "tab.$k"
    $p.Controls.Add($pages[$k])
    $tabs.TabPages.Add($p)
    $tabPage[$k] = $p
}

$pnlActions = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlActions.Dock = 'Bottom'; $pnlActions.AutoSize = $true; $pnlActions.WrapContents = $false; $pnlActions.Padding = '0,4,0,0'
$btnSoft    = New-EmtButton 'act.soft' 140
$btnHard    = New-EmtButton 'act.hard' 165
$btnHard.ForeColor = [System.Drawing.Color]::Firebrick
$btnVerify  = New-EmtButton 'act.verify' 100
$btnExport  = New-EmtButton 'act.export' 115
$btnExport.Margin = '18,3,6,3'
$btnCleanup = New-EmtButton 'conn.cleanup' 180
$btnCleanup.Margin = '18,3,6,3'
$pnlActions.Controls.AddRange(@($btnSoft, $btnHard, $btnVerify, $btnExport, $btnCleanup))

$pgDelete.Controls.Add($tabs)
$pgDelete.Controls.Add($pnlActions)
$pgDelete.Controls.Add($lblSummary)
$pgDelete.Controls.Add($lblDelSource)
$pgDelete.Controls.Add($grpCrit)

#endregion

#region ---------- Form layout: log, run area, tooltips ----------

$log = New-Object System.Windows.Forms.RichTextBox
$log.Dock = 'Fill'; $log.ReadOnly = $true; $log.Font = New-Object System.Drawing.Font('Consolas', 9)
$log.BackColor = [System.Drawing.SystemColors]::Window; $log.WordWrap = $false

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'; $split.Orientation = 'Horizontal'
$split.Panel1.Controls.Add($mainTabs)
$split.Panel2.Controls.Add($log)

# The bottom line (common for both tabs):
# left - the connection; right - the running operation: status, progress, cancel.
# The missing roles are in the log (after the connection). Disabled buttons show them too.
$pnlConnState = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlConnState.AutoSize = $true; $pnlConnState.WrapContents = $false; $pnlConnState.Margin = '0,0,0,0'; $pnlConnState.Anchor = 'Left'
$indServer = New-EmtIndicator ''
$pnlConnState.Controls.AddRange(@($indServer.Dot, $indServer.Text))

$btnCancel = New-EmtButton 'act.cancel' 90
$btnCancel.Anchor = 'Right'
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Width = 160; $progress.Height = 20; $progress.Margin = '6,3,6,3'; $progress.Anchor = 'Right'
$lblStatus = New-EmtLabel ''
$lblStatus.AutoSize = $false; $lblStatus.Dock = 'Fill'; $lblStatus.AutoEllipsis = $true
$lblStatus.TextAlign = 'MiddleRight'; $lblStatus.Margin = '8,3,3,3'

$pnlRun = New-Object System.Windows.Forms.TableLayoutPanel
$pnlRun.Dock = 'Bottom'; $pnlRun.AutoSize = $true; $pnlRun.Padding = '6,4,6,4'
$pnlRun.ColumnCount = 4; $pnlRun.RowCount = 1
[void]$pnlRun.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$pnlRun.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$pnlRun.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$pnlRun.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$pnlRun.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$pnlRun.Controls.Add($pnlConnState, 0, 0)
$pnlRun.Controls.Add($lblStatus, 1, 0)
$pnlRun.Controls.Add($progress, 2, 0)
$pnlRun.Controls.Add($btnCancel, 3, 0)

Set-EmtTip $tbTrSenders     'tip.senders'
Set-EmtTip $tbTrRecipient   'tip.trRecipient'
Set-EmtTip $tbTrMessageId   'tip.messageId'
Set-EmtTip $trPeriod.Today     'tip.today'
Set-EmtTip $trPeriod.Yesterday 'tip.yesterday'
Set-EmtTip $trPeriod.Week      'tip.week'
Set-EmtTip $btnTrCheck      'tip.checkMailbox'
Set-EmtTip $btnTrDelete     'tip.deleteThis'
Set-EmtTip $btnTrExport     'tip.export'
Set-EmtTip $tbSenders       'tip.senders'
Set-EmtTip $chkAll          'tip.all'
Set-EmtTip $delPeriod.Today     'tip.today'
Set-EmtTip $delPeriod.Yesterday 'tip.yesterday'
Set-EmtTip $delPeriod.Week      'tip.week'
Set-EmtTip $nudTimeout      'tip.timeout'
Set-EmtTip $btnCleanup      'tip.cleanup'
Set-EmtTip $btnExport       'tip.export'
Set-EmtTip $tbRecipient     'tip.recipient'
Set-EmtTip $chkOnly         'tip.only'
Set-EmtTip $tbMessageId     'tip.messageId'
Set-EmtTip $chkForwards     'tip.forwards'
Set-EmtTip $btnSoft         'tip.soft'
Set-EmtTip $btnHard         'tip.hard'
Set-EmtTip $btnVerify       'tip.verify'
Set-EmtTip $cmbLanguage     'tip.language'

# Dock order: the control added last docks first (top-most).
$form.Controls.Add($split)
$form.Controls.Add($pnlRun)
$form.Controls.Add($pnlConn)

#endregion

#region ---------- State machine ----------

function Get-EmtSelectedMessageId {
    $row = Get-EmtSelectedRow $gridTrMessages
    if ($row) { [string]$row['MessageId'] } else { '' }
}

function Get-EmtSelectedRecipient {
    <# The selected row of the recipients grid on the tab "Trace": Address and KindCode, or $null. #>
    $row = Get-EmtSelectedRow $gridTrRecipients
    if (-not $row) { return $null }
    [pscustomobject]@{ Address = [string]$row['Recipient']; KindCode = [string]$row['KindCode'] }
}

function Set-UiState {
    $busy      = [bool]$script:Job
    $caps      = $script:Caps
    $connected = [bool]$caps
    $r         = $script:Result
    $hasFound  = [bool]($r -and @($r.Found).Count -gt 0)
    $clean     = [bool]($hasFound -and -not $r.Errors)
    $deleted   = $script:Deleted     # '', 'Soft' or 'Hard'

    $tbServer.Enabled    = -not $busy
    $btnConnect.Enabled  = -not $busy
    $btnConnect.Text     = if ($connected) { T 'conn.reconnect' } else { T 'conn.connect' }
    $cmbLanguage.Enabled = -not $busy

    # ---- Trace ----
    $grpTr.Enabled       = -not $busy
    $btnTrFind.Enabled   = $connected -and -not $busy -and $caps.Tracking
    $msgId = Get-EmtSelectedMessageId
    $rcpt  = Get-EmtSelectedRecipient
    $btnTrCheck.Enabled  = $connected -and -not $busy -and $caps.MailboxCheck -and [bool]$msgId -and
                           [bool]$rcpt -and $rcpt.KindCode -eq 'Internal'
    $btnTrDelete.Enabled = $connected -and -not $busy -and [bool]$msgId -and ($caps.Search -or $caps.Verify)
    $btnTrExport.Enabled = (-not $busy) -and [bool]$script:Trace

    # ---- Delete ----
    # The criteria can be typed before the connection. Only Find needs it.
    $grpCrit.Enabled     = -not $busy
    # The quick search needs only Search-Mailbox. A soft delete always needs a Compliance Search.
    $btnSearch.Enabled   = $connected -and -not $busy -and ($caps.Search -or $caps.Verify)
    # After a soft delete, a hard delete is the next step. After a hard delete, only Verify is useful.
    $btnSoft.Enabled     = (-not $busy) -and $clean -and $connected -and (-not $deleted) -and
                           $caps.SoftDelete -and $caps.Search -and $caps.Verify
    $btnHard.Enabled     = (-not $busy) -and $clean -and $connected -and ($deleted -ne 'Hard') -and $caps.HardDelete
    $btnVerify.Enabled   = (-not $busy) -and $hasFound -and $connected -and $caps.Verify
    $btnExport.Enabled   = (-not $busy) -and [bool]$r
    $btnCleanup.Enabled  = (-not $busy) -and $connected -and $caps.Search
    $btnCancel.Enabled   = $busy
}

function Update-EmtTabTitles {
    <# Tab title = name and row count, e.g. "Found in mailboxes (42)". #>
    foreach ($k in $pages.Keys) {
        $dt = $pages[$k].DataSource
        $name = T "tab.$k"
        $tabPage[$k].Text = if ($dt -is [System.Data.DataTable]) { "$name ($($dt.Rows.Count))" } else { $name }
    }
    foreach ($k in $trPages.Keys) {
        $dt = $trPages[$k].DataSource
        $name = T "trtab.$k"
        # The mailbox tab shows the checked mailbox: "Recipient mailbox: x@contoso.com (5)".
        if ($k -eq 'mailbox' -and $dt -is [System.Data.DataTable] -and $script:MailboxCheck) { $name = T 'trtab.mailboxOf' $script:MailboxCheck.Mailbox }
        $trPage[$k].Text = if ($dt -is [System.Data.DataTable]) { "$name ($($dt.Rows.Count))" } else { $name }
    }
}

function Update-EmtConnIndicator {
    <# The bottom line: green = connected to the server, grey = not connected. #>
    if ($script:Caps) {
        $indServer.Text.Text = T 'conn.connectedTo' $script:ConnectedServer
        Set-EmtIndicator $indServer $dotGreen
    } else {
        $indServer.Text.Text = T 'conn.notConnected'
        Set-EmtIndicator $indServer $dotGrey
    }
}

function Update-EmtDeleteSource {
    <# The line over the delete result: data from the tracking log (green) or from the user (orange). #>
    if ($script:DelFromTrace) {
        $lblDelSource.Text = T 'del.fromTrace' $script:DelFromTrace
        $lblDelSource.ForeColor = [System.Drawing.Color]::ForestGreen
    } else {
        $lblDelSource.Text = T 'del.manual'
        $lblDelSource.ForeColor = [System.Drawing.Color]::DarkOrange
    }
}

function Clear-EmtResultViews {
    foreach ($g in $pages.Values) { $g.DataSource = $null }
    Update-EmtTabTitles
    $lblSummary.Text = T 'sum.none'
    $lblSummary.ForeColor = [System.Drawing.SystemColors]::ControlText
    $tips.SetToolTip($lblSummary, '')
}

function Clear-EmtResults {
    $script:Result  = $null
    $script:Check   = $null
    $script:Deleted = ''
    Clear-EmtResultViews
}

function Invoke-EmtInvalidate {
    <# A change of a delete criterion: the result is old, and the data is not from the tracking log any more. #>
    if ($script:Loading) { return }
    if ($script:DelFromTrace) { $script:DelFromTrace = ''; Update-EmtDeleteSource }
    if (-not $script:Result) { return }
    Clear-EmtResults
    Write-EmtLog (T 'log.invalidated') 'Warn'
    Set-UiState
}

function Get-EmtCriteria {
    @{
        Senders         = @(Split-EmtAddressList $tbSenders.Text)
        Subject         = $tbSubject.Text.Trim()
        Recipient       = $tbRecipient.Text.Trim().ToLowerInvariant()
        MessageId       = $tbMessageId.Text.Trim()
        Attachment      = $tbAttachment.Text.Trim()
        Keywords        = $tbKeywords.Text.Trim()
        Start           = $dtStart.Value
        End             = Get-EmtPeriodEnd $dtEnd.Value      # the end includes its minute
        IncludeForwards = $chkForwards.Checked
        AllMailboxes    = $chkAll.Checked
        OnlyRecipient   = $chkOnly.Checked -and $chkOnly.Enabled
    }
}

function Get-EmtTraceCriteria {
    @{
        Senders   = @(Split-EmtAddressList $tbTrSenders.Text)
        Recipient = $tbTrRecipient.Text.Trim().ToLowerInvariant()
        Subject   = $tbTrSubject.Text.Trim()
        MessageId = $tbTrMessageId.Text.Trim()
        Start     = $dtTrStart.Value
        End       = Get-EmtPeriodEnd $dtTrEnd.Value
    }
}

#endregion

#region ---------- Trace views ----------

function Get-EmtReportOf([string] $MessageId) {
    <# The report rows (recipients) of one message. Show-EmtTraceViews makes the map. #>
    if ($script:ReportByMsg -and $script:ReportByMsg.ContainsKey($MessageId)) { $script:ReportByMsg[$MessageId].ToArray() }
}

function Clear-EmtTraceViews {
    foreach ($g in $trPages.Values) { $g.DataSource = $null }
    $gridTrMessages.DataSource = $null
    Update-EmtTabTitles
    $lblTrSummary.Text = T 'trsum.none'
    $lblTrSummary.ForeColor = [System.Drawing.SystemColors]::ControlText
    $tips.SetToolTip($lblTrSummary, '')
}

function Show-EmtTraceViews($t) {
    <# Fills the message grid and the summary line of the trace. No audit: it also runs after a language change. #>
    $keep = Get-EmtSelectedMessageId
    # The recipients of each message in one pass. A search in the full report for each message is slow:
    # 2000 messages and 6000 recipients take 30 seconds.
    $script:ReportByMsg = @{}
    foreach ($r in @($t.Report)) {
        if (-not $script:ReportByMsg.ContainsKey($r.MessageId)) { $script:ReportByMsg[$r.MessageId] = New-Object System.Collections.ArrayList }
        [void]$script:ReportByMsg[$r.MessageId].Add($r)
    }
    $msgs = @($t.Messages | ForEach-Object {
        [pscustomobject]@{
            Time = $_.Time; Sender = $_.Sender; Subject = $_.Subject; Recipients = $_.Recipients
            StatusLine = (Get-EmtMessageStatusLine @(Get-EmtReportOf $_.MessageId))
            Note = $(if ($_.ServerGenerated) { T 'note.serverGenerated' } else { '' })
            MessageId = $_.MessageId; Problems = $_.Problems; Pending = $_.Pending
        }
    })
    Set-EmtGridData $gridTrMessages $msgs ([ordered]@{
        Time = 'col.time'; Sender = 'col.sender'; Subject = 'col.subject'; Recipients = 'col.recipientCount'
        StatusLine = 'col.statusLine'; Note = 'col.note'; MessageId = 'col.messageId'; Problems = ''; Pending = ''
    }) -Hidden 'Problems', 'Pending'

    $report = @($t.Report)
    $count = { param($codes) @($report | Where-Object { $_.Status -in $codes }).Count }
    $failed = & $count @(Get-EmtErrorStatuses)
    if ($msgs.Count) {
        $lblTrSummary.Text = T 'trsum.text' $msgs.Count $report.Count (& $count @('Delivered')) (& $count @('External')) `
                                            $failed (& $count @('Deferred', 'InTransit', 'Unknown')) $t.Seconds
        $lblTrSummary.ForeColor = $(if ($failed) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.SystemColors]::ControlText })
        $tips.SetToolTip($lblTrSummary, $lblTrSummary.Text)
    } else {
        $lblTrSummary.Text = T 'trsum.empty'
        $lblTrSummary.ForeColor = [System.Drawing.Color]::DarkOrange
        $tips.SetToolTip($lblTrSummary, (@($t.NoResultHints) -join "`r`n"))
    }
    if (@($t.TrackingErrors).Count) {
        $lblTrSummary.Text = (T 'sum.errors') + '   ' + $lblTrSummary.Text
        $lblTrSummary.ForeColor = [System.Drawing.Color]::Firebrick
    }
    # Keep the selected message after a language change.
    if ($keep) {
        foreach ($row in $gridTrMessages.Rows) {
            if ([string]$row.DataBoundItem.Row['MessageId'] -eq $keep) { $row.Selected = $true; break }
        }
    }
    Show-EmtTraceDetail
}

function Show-EmtTraceDetail {
    <# The recipients, the events and the mailbox findings of the selected message. #>
    $t = $script:Trace
    $msgId = Get-EmtSelectedMessageId
    if (-not $t -or -not $msgId) {
        foreach ($g in $trPages.Values) { $g.DataSource = $null }
        Update-EmtTabTitles
        Set-UiState
        return
    }
    $rows = @(Get-EmtReportOf $msgId | ForEach-Object {
        [pscustomobject]@{
            Recipient = $_.Recipient; Status = (T "st.$($_.Status)")
            Kind = $(if ($_.Kind) { T "kind.$($_.Kind)" } else { '' }); Time = $_.Time; Server = $_.Server
            Folder = (Get-EmtProp $_ 'Folder'); Detail = $_.Detail; Hint = (Get-EmtStatusHint $_.Detail)
            StatusCode = $_.Status; KindCode = $_.Kind
        }
    } | Sort-Object { -(Get-EmtStatusRank $_.StatusCode) }, Recipient)
    # Problems first: failed and dropped, then the others.
    $first = @(Get-EmtErrorStatuses) + 'Deferred', 'Forwarded'
    $rows = @(@($rows | Where-Object { $_.StatusCode -in $first }) + @($rows | Where-Object { $_.StatusCode -notin $first }))
    Set-EmtGridData $gridTrRecipients $rows ([ordered]@{
        Recipient = 'col.recipient'; Status = 'col.status'; Kind = 'col.kind'; Time = 'col.time'; Server = 'col.server'
        Hint = 'col.hint'; Folder = 'col.folder'; Detail = 'col.detail'; StatusCode = ''; KindCode = ''
    }) -Hidden 'StatusCode', 'KindCode'
    $events = @($t.Events | Where-Object { $_.MessageId -eq $msgId })
    Set-EmtGridData $gridTrEvents $events ([ordered]@{
        Timestamp = 'col.time'; EventId = 'col.event'; Source = 'col.source'; ServerHostname = 'col.server'
        Recipients = 'col.recipients'; RecipientStatus = 'col.detail'; Sender = 'col.sender'
    })
    $mc = $script:MailboxCheck
    if ($mc -and $mc.MessageId -eq $msgId) { Show-EmtMailboxView $mc } else { $gridTrMailbox.DataSource = $null }
    Update-EmtTabTitles
    Set-UiState
}

function Show-EmtMailboxView($mc) {
    # The findings are texts: make them again from the data, in the current language.
    $findings = if (Get-EmtProp $mc 'Data') { @(Get-EmtMailboxFindings -Data $mc.Data -Sender $mc.Sender -Subject $mc.Subject) } else { @($mc.Findings) }
    $rows = @($findings | ForEach-Object {
        [pscustomobject]@{
            Severity = (T "sev.$($_.Severity)"); Area = (T "area.$($_.Area)"); Text = $_.Text
            SeverityCode = $_.Severity
        }
    })
    Set-EmtGridData $gridTrMailbox $rows ([ordered]@{
        Severity = 'col.severity'; Area = 'col.area'; Text = 'col.finding'; SeverityCode = ''
    }) -Hidden 'SeverityCode'
}

function Show-EmtTraceResult($t) {
    $script:Trace = $t
    $script:MailboxCheck = $null
    Show-EmtTraceViews $t
    if (-not @($t.Messages).Count) {
        Write-EmtLog (T 'log.traceEmpty') 'Warn'
        foreach ($h in @($t.NoResultHints)) { Write-EmtLog ('  - ' + $h) 'Warn' }
    }
    if ($gridTrMessages.Rows.Count) { $gridTrMessages.Rows[0].Selected = $true; Show-EmtTraceDetail }
    $c = $t.Criteria
    Write-EmtAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = 'Trace'
        Query = "Senders=$(@($c.Senders) -join ' '); Recipient=$($c.Recipient); Subject=$($c.Subject); MessageId=$($c.MessageId); Start=$($c.Start); End=$($c.End)"
        Items = @($t.Messages).Count; Mailboxes = @($t.Report).Count
        Result = $(if (@($t.TrackingErrors).Count) { 'Errors' } else { 'OK' })
        Details = "Failed=$(@($t.Report | Where-Object { $_.Status -in (Get-EmtErrorStatuses) }).Count); Seconds=$($t.Seconds)"
    }
}

#endregion

#region ---------- Delete views ----------

function Show-EmtResultViews($r) {
    <# Fills the grids and the summary line. No audit: it also runs after a language change. #>
    $msgs = @($r.Messages | ForEach-Object {
        [pscustomobject]@{
            Time = $_.Time; Sender = $_.Sender; Subject = $_.Subject; Recipients = $_.Recipients
            Note = $(if ($_.ServerGenerated) { T 'note.serverGenerated' } else { '' }); MessageId = $_.MessageId
        }
    })
    Set-EmtGridData $gridMessages $msgs ([ordered]@{
        Time = 'col.time'; Sender = 'col.sender'; Subject = 'col.subject'; Recipients = 'col.recipientCount'
        Note = 'col.note'; MessageId = 'col.messageId' })
    Set-EmtGridData $gridTracking @($r.Events) ([ordered]@{
        Timestamp = 'col.time'; ServerHostname = 'col.server'; EventId = 'col.event'; Source = 'col.source'
        Sender = 'col.sender'; Recipients = 'col.recipients'; MessageSubject = 'col.subject'; MessageId = 'col.messageId' })
    $rcpt = @($r.Recipients | ForEach-Object {
        [pscustomobject]@{ Address = $_.Address; Kind = (T "kind.$($_.Kind)"); Mailbox = $_.Mailbox; Type = $_.Type; KindCode = $_.Kind }
    } | Sort-Object Kind, Address)
    Set-EmtGridData $gridRecipients $rcpt ([ordered]@{
        Address = 'col.address'; Kind = 'col.kind'; Mailbox = 'col.mailbox'; Type = 'col.exchangeType'; KindCode = '' }) -Hidden 'KindCode'
    Set-EmtGridData $gridFound @($r.Found) ([ordered]@{ Mailbox = 'col.mailbox'; Items = 'col.items'; SizeKB = 'col.sizeKB' })
    $miss = @($r.Missing | ForEach-Object { [pscustomobject]@{ Mailbox = $_; Reason = (T 'missing.reason') } })
    Set-EmtGridData $gridMissing $miss ([ordered]@{ Mailbox = 'col.mailbox'; Reason = 'col.reason' })

    $unindexedText = if ($r.Unindexed -lt 0) { T 'sum.na' } else { [string]$r.Unindexed }
    $methodText    = if ($r.Method -eq 'SearchMailbox') { T 'sum.quick' } else { 'Compliance Search' }
    $scopeText     = if (@($r.Locations) -contains 'All') { T 'common.all' } else { [string]@($r.Locations).Count }
    $lblSummary.Text = T 'sum.text' $r.Items @($r.Found).Count $scopeText $r.SizeKB @($r.External).Count `
                                    $unindexedText $methodText $r.Seconds
    $lblSummary.ForeColor = [System.Drawing.SystemColors]::ControlText
    if (-not @($r.Found).Count -and -not $r.Errors) { $lblSummary.Text = (T 'sum.nothing') + '   ' + $lblSummary.Text }
    # More items than messages in a mailbox: the query can also find other messages (orange warning).
    $extra = @(Get-EmtExtraItemMailboxes -Found @($r.Found) -Messages @($r.Messages).Count)
    if ($extra.Count) {
        $lblSummary.Text = (T 'sum.extraItems' $extra.Count @($r.Messages).Count) + '   ' + $lblSummary.Text
        $lblSummary.ForeColor = [System.Drawing.Color]::DarkOrange
    }
    $unconfirmed = [bool]($r.PSObject.Properties['Unconfirmed'] -and $r.Unconfirmed)
    if ($unconfirmed) {
        $lblSummary.Text = (T 'sum.unconfirmed') + '   ' + $lblSummary.Text
        $lblSummary.ForeColor = [System.Drawing.Color]::Firebrick
    }
    if ($r.Errors) {
        $lblSummary.Text = (T 'sum.errors') + '   ' + $lblSummary.Text
        $lblSummary.ForeColor = [System.Drawing.Color]::Firebrick
    }
    $tips.SetToolTip($lblSummary, $(if ($r.Errors) { "$($lblSummary.Text)`r`n`r`n$($r.Errors)" } else { $lblSummary.Text }))
    Update-EmtTabTitles
}

function Show-EmtSearchResult($r) {
    $script:Result = $r
    $script:Check  = $null
    $gridCheck.DataSource = $null
    Show-EmtResultViews $r
    $tabs.SelectedTab = $tabPage['found']

    Write-EmtAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = 'Search'; SearchName = $r.SearchName; Query = $r.Query
        Mailboxes = @($r.Found).Count; Items = $r.Items
        Result = $(if ($r.Errors) { 'Errors' } else { 'OK' })
        Details = "Method=$($r.Method); Seconds=$($r.Seconds); FromTrace=$($script:DelFromTrace); External=$(@($r.External | ForEach-Object Address) -join ' '); Unindexed=$($r.Unindexed)"
    }
}

function Show-EmtCheckView([object[]] $Check) {
    Set-EmtGridData $gridCheck @($Check) ([ordered]@{
        Mailbox = 'col.mailbox'; Visible = 'col.visible'; Total = 'col.total'; Note = 'col.note' })
    Update-EmtTabTitles
}

function Show-EmtCheckResult([object[]] $Check, [string] $Action, $DeleteInfo) {
    $script:Check = @($Check)
    Show-EmtCheckView $script:Check
    $tabs.SelectedTab = $tabPage['check']
    $visible = @($script:Check | Where-Object { $_.Visible -gt 0 })
    $hidden  = @($script:Check | Where-Object { $_.Visible -eq 0 -and $_.Total -gt 0 })
    $failed  = @($script:Check | Where-Object { $_.Visible -lt 0 })
    if ($visible.Count) { Write-EmtLog (T 'log.checkVisible' $visible.Count) 'Warn' }
    else { Write-EmtLog (T 'log.checkClean') 'Ok' }
    if ($hidden.Count) { Write-EmtLog (T 'log.checkHidden' $hidden.Count) 'Warn' }
    if ($failed.Count) { Write-EmtLog (T 'log.checkFailed' $failed.Count) 'Error' }

    $resultText = if ($failed.Count) { 'CheckFailed' } elseif ($visible.Count) { 'Visible' } elseif ($hidden.Count) { 'RecoverableOnly' } else { 'Clean' }
    Write-EmtAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = $Action; SearchName = $script:Result.SearchName
        Query = $script:Result.Query; Mailboxes = $script:Check.Count; Items = $script:Result.Items
        Result = $resultText; Details = [string]$DeleteInfo
    }
}

function Get-EmtDeleteWarning($r) {
    <#
        All warnings for the delete confirmation: not confirmed by the tracking log, large delete,
        more items than messages. A warning makes the typed word necessary also for a soft delete.
    #>
    $w = @()
    if ($r.PSObject.Properties['Unconfirmed'] -and $r.Unconfirmed) { $w += (T 'confirm.unconfirmed') }
    $large = Get-EmtLargeDeleteWarning -Items $r.Items -Messages @($r.Messages).Count
    if ($large) { $w += $large }
    $extra = @(Get-EmtExtraItemMailboxes -Found @($r.Found) -Messages @($r.Messages).Count)
    if ($extra.Count) { $w += (T 'confirm.extraItems' @($r.Messages).Count ($extra -join ', ')) }
    $w -join "`r`n`r`n"
}

function Get-EmtDeleteText {
    <# Text of the delete confirmation: warnings, counts, the messages and the mailboxes. #>
    param($r, [ValidateSet('Soft', 'Hard')] [string] $Kind, [string] $Warning)
    $mbx = @($r.Found | ForEach-Object Mailbox)
    $msgs = @($r.Messages)
    $fmt = T 'fmt.date'
    $msgLines = @($msgs | Select-Object -First 30 | ForEach-Object {
        '{0}  {1}  «{2}»' -f ([datetime]$_.Time).ToString($fmt), $_.Sender, $_.Subject
    })
    if ($msgs.Count -gt 30) { $msgLines += (T 'confirm.more' ($msgs.Count - 30)) }
    if (-not $msgs.Count)   { $msgLines = @(T 'confirm.noMessages') }
    $unindexed = if ($r.Unindexed -lt 0) { T 'confirm.unindexedNa' } else { [string]$r.Unindexed }
    $external  = @($r.External | ForEach-Object Address) -join ', '
    if (-not $external) { $external = T 'confirm.none' }

    $head = if ($Kind -eq 'Soft') {
        @((T 'confirm.softHead'), '', (T 'confirm.soft1'), (T 'confirm.soft2'), (T 'confirm.soft3'))
    } else {
        @((T 'confirm.hardHead'), '', (T 'confirm.hard1'), (T 'confirm.hard2'), (T 'confirm.hard3'), (T 'confirm.hard4'))
    }
    $search = if ($r.SearchName) { $r.SearchName } elseif ($Kind -eq 'Soft') { T 'confirm.searchNew' } else { '—' }

    $lines = @()
    if ($Warning) { $lines += $Warning; $lines += '' }
    $lines += $head
    $lines += ''
    $lines += (T 'confirm.items' $r.Items)
    $lines += (T 'confirm.mailboxes' $mbx.Count)
    $lines += (T 'confirm.search' $search)
    $lines += (T 'confirm.kql' $r.Query)
    $lines += (T 'confirm.external' $external)
    $lines += (T 'confirm.unindexed' $unindexed)
    $lines += ''
    $lines += (T 'confirm.messages' $msgs.Count)
    $lines += $msgLines
    $lines += ''
    $lines += (T 'confirm.mailboxList')
    $lines += $mbx
    $lines -join "`r`n"
}

#endregion

#region ---------- Address book search (partial addresses) ----------

function Show-EmtPick {
    <# Selection of one or more addresses from the candidates. Returns the selected addresses, or $null on cancel. #>
    param([string] $Text, [object[]] $Candidates, [switch] $Multi, [int] $Max = 25)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = T 'pick.title'; $dlg.Size = '640,420'; $dlg.StartPosition = 'CenterParent'
    $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false; $dlg.Font = $form.Font; $dlg.ShowIcon = $false

    $top = New-EmtLabel ''
    $top.Dock = 'Top'; $top.AutoSize = $false; $top.Height = 44; $top.Padding = '6,6,6,0'
    $top.Text = $(if ($Multi) { T 'pick.promptMulti' $Text } else { T 'pick.prompt' $Text })
    if ($Candidates.Count -ge $Max) { $top.Text += "`r`n" + (T 'pick.tooMany' $Max) }

    $list = New-Object System.Windows.Forms.ListBox
    $list.Dock = 'Fill'; $list.IntegralHeight = $false
    $list.SelectionMode = $(if ($Multi) { 'MultiExtended' } else { 'One' })
    foreach ($c in $Candidates) { [void]$list.Items.Add(('{0}  <{1}>  {2}' -f $c.Name, $c.Address, $c.Type)) }
    if ($list.Items.Count) { $list.SelectedIndex = 0 }

    $bottom = New-Object System.Windows.Forms.FlowLayoutPanel
    $bottom.Dock = 'Bottom'; $bottom.AutoSize = $true; $bottom.Padding = '6,6,6,6'
    $ok = New-EmtButton '' 120; $ok.Text = T 'pick.ok'; $ok.DialogResult = 'OK'
    $cancel = New-EmtButton '' 120; $cancel.Text = T 'dlg.cancel'; $cancel.DialogResult = 'Cancel'
    $bottom.Controls.AddRange(@($ok, $cancel))
    $dlg.AcceptButton = $ok; $dlg.CancelButton = $cancel
    $list.Add_DoubleClick({ $dlg.DialogResult = 'OK' }.GetNewClosure())

    $dlg.Controls.Add($list); $dlg.Controls.Add($bottom); $dlg.Controls.Add($top)
    $answer = $dlg.ShowDialog($form)
    $picked = @($list.SelectedIndices | ForEach-Object { $Candidates[$_].Address })
    $dlg.Dispose()
    if ($answer -ne [System.Windows.Forms.DialogResult]::OK -or -not $picked.Count) { return $null }
    , $picked
}

function Complete-EmtAddressResolve {
    <#
        Puts the full addresses in the fields: one candidate = automatic, more = the user selects,
        none = error. Returns $true if the search can start.
    #>
    param([object[]] $Resolved, $SendersBox, $RecipientBox)
    $senders = @(Split-EmtAddressList $SendersBox.Text)
    $recipient = $RecipientBox.Text.Trim().ToLowerInvariant()
    $map = @{}
    $notFound = @()
    foreach ($x in @($Resolved | Where-Object { $_ })) {
        $cands = @($x.Candidates)
        if (-not $cands.Count) { $notFound += $x.Input; continue }
        if ($cands.Count -eq 1) {
            $map[$x.Input] = @($cands[0].Address)
            Write-EmtLog (T 'log.resolvedOne' $x.Input $cands[0].Address) 'Ok'
            continue
        }
        # Several senders are possible. The recipient is one mailbox.
        $multi = ($x.Input -in $senders) -and ($x.Input -ne $recipient)
        $picked = Show-EmtPick -Text $x.Input -Candidates $cands -Multi:$multi -Max $x.Max
        if (-not $picked) { Write-EmtLog (T 'log.resolveCancelled') 'Warn'; return $false }
        $map[$x.Input] = @($picked)
        Write-EmtLog (T 'log.resolvedPick' $x.Input ($picked -join ', ')) 'Ok'
    }
    if ($notFound.Count) {
        [void][System.Windows.Forms.MessageBox]::Show($form, (T 'resolve.notFound' ($notFound -join ', ')),
            (T 'resolve.title'), 'OK', [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    $script:Loading = $true     # the fields change, but this is not a change of the criteria by the user
    try {
        $newSenders = @($senders | ForEach-Object { if ($map.ContainsKey($_)) { $map[$_] } else { $_ } } | Sort-Object -Unique)
        $SendersBox.Text = $newSenders -join '; '
        if ($recipient -and $map.ContainsKey($recipient)) { $RecipientBox.Text = @($map[$recipient])[0] }
    } finally { $script:Loading = $false }
    $true
}

function Invoke-EmtWithResolve {
    <#
        Starts an action after the address book search of the partial addresses (if any).
        $Then runs when all addresses are full.
    #>
    param([hashtable] $Criteria, $SendersBox, $RecipientBox, [scriptblock] $Then)
    $partial = @(@($Criteria.Senders) + @($Criteria.Recipient) | Where-Object { $_ -and -not (Test-EmtFullAddress $_) } | Sort-Object -Unique)
    if (-not $partial.Count) { & $Then; return }
    $script:ResolveNext = @{ Senders = $SendersBox; Recipient = $RecipientBox; Then = $Then }
    Start-EmtOperation -Name 'Resolve' -Params @{ Inputs = $partial } -StatusText (T 'status.resolve') -OnSuccess {
        param($x)
        $n = $script:ResolveNext
        $script:ResolveNext = $null
        if (Complete-EmtAddressResolve -Resolved @($x) -SendersBox $n.Senders -RecipientBox $n.Recipient) { & $n.Then }
    }
}

#endregion

#region ---------- Actions ----------

function Start-EmtTrace {
    $crit = Get-EmtTraceCriteria
    $script:Trace = $null
    $script:MailboxCheck = $null
    Clear-EmtTraceViews
    $fmt = T 'fmt.date'
    Write-EmtLog (T 'log.traceStart' ($crit.Senders -join ', ') $crit.Recipient $crit.Subject $crit.Start.ToString($fmt) $crit.End.ToString($fmt))
    Start-EmtOperation -Name 'Trace' -Params @{ Criteria = $crit } -StatusText (T 'status.trace') -OnSuccess {
        param($t) Show-EmtTraceResult $t
    }
}

function Start-EmtSearch {
    $crit = Get-EmtCriteria
    Save-EmtSettings
    Clear-EmtResults
    $fmt = T 'fmt.date'
    Write-EmtLog (T 'log.search' ($crit.Senders -join ', ') $crit.Subject $crit.Start.ToString($fmt) $crit.End.ToString($fmt))
    Start-EmtOperation -Name 'Search' -Params @{ Criteria = $crit; TimeoutMinutes = [int]$nudTimeout.Value } `
        -StatusText (T 'status.search') -OnSuccess { param($r) Show-EmtSearchResult $r }
}

function Open-EmtDeleteForMessage {
    <#
        Fills the tab "Delete" with the data of the selected traced message and opens the tab.
        The data comes from the tracking log: MessageId, sender, exact subject, the day of the message.
    #>
    $msgId = Get-EmtSelectedMessageId
    $m = @($script:Trace.Messages | Where-Object MessageId -eq $msgId)[0]
    if (-not $m) { return }
    $script:Loading = $true
    try {
        $tbSenders.Text = $m.Sender
        $tbSubject.Text = [string]$m.Subject
        $tbMessageId.Text = $m.MessageId
        $tbRecipient.Text = ''; $tbAttachment.Text = ''; $tbKeywords.Text = ''
        $chkOnly.Checked = $false; $chkAll.Checked = $false
        # The period: from some minutes before the message to 23:59 of its day. The tracking log uses the exact time.
        # The query (KQL) uses whole days: this period gives one day, not the full period of the message search.
        $sent = [datetime]$m.Time
        $from = $sent.AddMinutes(-5)
        $dtStart.Value = $from.AddSeconds(-$from.Second).AddMilliseconds(-$from.Millisecond)
        $dtEnd.Value = $sent.Date.AddDays(1).AddMinutes(-1)    # 23:59 of the same day
    } finally { $script:Loading = $false }
    Clear-EmtResults
    $script:DelFromTrace = $m.MessageId
    Update-EmtDeleteSource
    $mainTabs.SelectedTab = $pgDelete
    Write-EmtLog (T 'log.delPrefilled' $m.MessageId)
    Set-UiState
}

#endregion

#region ---------- Event handlers ----------

$btnConnect.Add_Click({
    $srv = $tbServer.Text.Trim()
    if (-not $srv) { [void][System.Windows.Forms.MessageBox]::Show($form, (T 'conn.enterServer'), (T 'app.short')); return }
    $script:Caps = $null
    Clear-EmtResults
    Update-EmtConnIndicator
    Start-EmtOperation -Name 'Connect' -Params @{ Server = $srv } -StatusText (T 'status.connecting' $srv) -OnSuccess {
        param($caps)
        $script:Caps = $caps
        $script:ConnectedServer = $tbServer.Text.Trim()
        Update-EmtConnIndicator
        Save-EmtSettings
        # The missing roles: one warning for each function that is not available.
        if (-not $caps.Tracking)     { Write-EmtLog (T 'log.noTrackRole') 'Warn' }
        if (-not $caps.MailboxCheck) { Write-EmtLog (T 'log.noMailboxRole') 'Warn' }
        if (-not $caps.Search)       { Write-EmtLog (T 'log.noSearchRole') 'Warn' }
        if (-not $caps.Verify)       { Write-EmtLog (T 'log.noVerifyRole') 'Warn' }
        if (-not $caps.SoftDelete)   { Write-EmtLog (T 'log.noSoftRole') 'Warn' }
        if (-not $caps.HardDelete)   { Write-EmtLog (T 'log.noHardRole') 'Warn' }
    }
})

$cmbLanguage.Add_SelectedIndexChanged({
    if ($script:Loading -or $cmbLanguage.SelectedIndex -lt 0) { return }
    $code = $script:LanguageCodes[$cmbLanguage.SelectedIndex]
    if ($code -eq $Sync.Lang) { return }
    $Sync.Lang = $code
    Update-EmtLanguage
    Save-EmtSettings
    Write-EmtLog (T 'log.language')
})

# Enter starts the search of the open tab.
$mainTabs.Add_SelectedIndexChanged({
    $form.AcceptButton = if ($mainTabs.SelectedTab -eq $pgTrace) { $btnTrFind } else { $btnSearch }
})

# ---- Trace ----

# The quick periods: from 0:00 to 23:59 of the last day (the end includes its minute).
$trPeriod.Today.Add_Click({     $dtTrStart.Value = (Get-Date).Date;            $dtTrEnd.Value = (Get-Date).Date.AddDays(1).AddMinutes(-1) })
$trPeriod.Yesterday.Add_Click({ $dtTrStart.Value = (Get-Date).Date.AddDays(-1); $dtTrEnd.Value = (Get-Date).Date.AddMinutes(-1) })
$trPeriod.Week.Add_Click({      $dtTrStart.Value = (Get-Date).Date.AddDays(-6); $dtTrEnd.Value = (Get-Date).Date.AddDays(1).AddMinutes(-1) })

$btnTrFind.Add_Click({
    $crit = Get-EmtTraceCriteria
    $problems = @(Test-EmtCriteria $crit -ForTrace)
    if ($problems.Count) {
        [void][System.Windows.Forms.MessageBox]::Show($form, ($problems -join "`r`n"), (T 'crit.checkTitle'),
            'OK', [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Invoke-EmtWithResolve -Criteria $crit -SendersBox $tbTrSenders -RecipientBox $tbTrRecipient -Then { Start-EmtTrace }
})

$gridTrMessages.Add_SelectionChanged({ if (-not $script:Loading) { Show-EmtTraceDetail } })
$gridTrMessages.Add_CellDoubleClick({ param($s, $e) if ($e.RowIndex -ge 0) { $trDetail.SelectedTab = $trPage['events'] } })
$gridTrRecipients.Add_SelectionChanged({ Set-UiState })

$btnTrCheck.Add_Click({
    $rcpt = Get-EmtSelectedRecipient
    $msgId = Get-EmtSelectedMessageId
    $m = @($script:Trace.Messages | Where-Object MessageId -eq $msgId)[0]
    if (-not $rcpt -or -not $m) { return }
    # The user can select another message during the check: keep the message of this check.
    $script:CheckForMessage = $msgId
    Start-EmtOperation -Name 'CheckMailbox' -StatusText (T 'progress.checkMailbox' $rcpt.Address) -Params @{
        Mailbox = $rcpt.Address; Sender = $m.Sender; Subject = [string]$m.Subject
    } -OnSuccess {
        param($x)
        $script:MailboxCheck = [pscustomobject]@{
            MessageId = $script:CheckForMessage; Mailbox = $x.Mailbox; Findings = @($x.Findings)
            Data = $x.Data; Sender = $x.Sender; Subject = $x.Subject
        }
        Write-EmtAudit -Path $AuditPath -Record @{
            Server = $script:ConnectedServer; Action = 'MailboxCheck'; Mailboxes = 1
            Items = @($x.Findings | Where-Object Severity -eq 'Alert').Count; Result = 'OK'
            Details = "Mailbox=$($x.Mailbox); MessageId=$($script:MailboxCheck.MessageId)"
        }
        # Show the findings only with their message. For another message they come back with it.
        if ((Get-EmtSelectedMessageId) -ne $script:CheckForMessage) { Write-EmtLog (T 'log.checkOtherMessage'); return }
        Show-EmtMailboxView $script:MailboxCheck
        Update-EmtTabTitles
        $trDetail.SelectedTab = $trPage['mailbox']
    }
})

$btnTrDelete.Add_Click({ Open-EmtDeleteForMessage })

$btnTrExport.Add_Click({
    $t = $script:Trace
    Export-EmtSets -Stem ('EMT_trace_{0:yyyyMMdd_HHmmss}' -f (Get-Date)) -Sets ([ordered]@{
        'messages'   = @($t.Messages | Select-Object Time, Sender, Subject, Recipients, StatusLine, MessageId)
        'recipients' = @($t.Report | Select-Object MessageId, Recipient, Kind, Status, Time, Server, Folder, Detail, Hint)
        'events'     = @($t.Events | Select-Object Timestamp, ServerHostname, EventId, Source, Sender,
                            @{ n = 'Recipients'; e = { $_.Recipients -join '; ' } },
                            @{ n = 'RecipientStatus'; e = { $_.RecipientStatus -join '; ' } }, MessageSubject, MessageId)
        'mailbox'    = @($(if ($script:MailboxCheck) { $script:MailboxCheck.Findings }))
    })
})

# ---- Delete ----

$delPeriod.Today.Add_Click({     $dtStart.Value = (Get-Date).Date;            $dtEnd.Value = (Get-Date).Date.AddDays(1).AddMinutes(-1) })
$delPeriod.Yesterday.Add_Click({ $dtStart.Value = (Get-Date).Date.AddDays(-1); $dtEnd.Value = (Get-Date).Date.AddMinutes(-1) })
$delPeriod.Week.Add_Click({      $dtStart.Value = (Get-Date).Date.AddDays(-6); $dtEnd.Value = (Get-Date).Date.AddDays(1).AddMinutes(-1) })

$btnSearch.Add_Click({
    $crit = Get-EmtCriteria
    $problems = @(Test-EmtCriteria $crit)
    if ($problems.Count) {
        [void][System.Windows.Forms.MessageBox]::Show($form, ($problems -join "`r`n"), (T 'crit.checkTitle'),
            'OK', [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Invoke-EmtWithResolve -Criteria $crit -SendersBox $tbSenders -RecipientBox $tbRecipient -Then { Start-EmtSearch }
})

$btnSoft.Add_Click({
    $r = $script:Result
    $mbx = @($r.Found | ForEach-Object Mailbox)
    $warning = Get-EmtDeleteWarning $r
    $text = Get-EmtDeleteText -r $r -Kind 'Soft' -Warning $warning
    # A large or not confirmed delete needs the typed word also for a soft delete.
    $word = if ($warning) { T 'confirm.word' } else { '' }
    if (-not (Show-EmtConfirm -Title (T 'confirm.softTitle') -Text $text -OkText (T 'act.soft') -RequireWord $word)) { return }
    Start-EmtOperation -Name 'SoftDelete' -StatusText (T 'status.soft') -Params @{
        SearchName = $r.SearchName; Mailboxes = $mbx; Query = $r.Query; TimeoutMinutes = [int]$nudTimeout.Value
    } -OnSuccess {
        param($x)
        $script:Deleted = 'Soft'
        $info = "Status=$($x.Delete.Status); Search=$($x.SearchName)"
        if ($x.Delete.Errors) { $info += "; Errors=$($x.Delete.Errors)" }
        Show-EmtCheckResult -Check @($x.Check) -Action 'SoftDelete' -DeleteInfo $info
    }
})

$btnHard.Add_Click({
    $r = $script:Result
    $mbx = @($r.Found | ForEach-Object Mailbox)
    $expected = @{}
    foreach ($f in @($r.Found)) { $expected[$f.Mailbox] = [int]$f.Items }
    $warning = Get-EmtDeleteWarning $r
    $text = Get-EmtDeleteText -r $r -Kind 'Hard' -Warning $warning
    if (-not (Show-EmtConfirm -Title (T 'confirm.hardTitle') -Text $text -OkText (T 'act.hard') -RequireWord (T 'confirm.word'))) { return }
    Start-EmtOperation -Name 'HardDelete' -StatusText (T 'status.hard') -Params @{
        Mailboxes = $mbx; Query = $r.Query; Expected = $expected; SearchName = $r.SearchName
    } -OnSuccess {
        param($x)
        $script:Deleted = 'Hard'
        $deleted = [int](Get-EmtSum -Items @($x.Delete) -Property Deleted)
        $errors  = @($x.Delete | Where-Object Error).Count
        $info    = "Deleted=$deleted; MailboxErrors=$errors; Processed=$(@($x.Delete).Count)"
        if ($x.Cancelled) {
            Write-EmtLog (T 'log.hardCancelled' $deleted) 'Warn'
            Write-EmtAudit -Path $AuditPath -Record @{
                Server = $script:ConnectedServer; Action = 'HardDelete'; SearchName = $script:Result.SearchName
                Query = $script:Result.Query; Mailboxes = @($x.Delete).Count; Items = $deleted
                Result = 'Cancelled'; Details = $info
            }
        } else {
            Show-EmtCheckResult -Check @($x.Check) -Action 'HardDelete' -DeleteInfo $info
        }
    }
})

$btnVerify.Add_Click({
    $r = $script:Result
    Start-EmtOperation -Name 'Verify' -StatusText (T 'status.verify') -Params @{
        Mailboxes = @($r.Found | ForEach-Object Mailbox); Query = $r.Query
    } -OnSuccess { param($x) Show-EmtCheckResult -Check @($x.Check) -Action 'Verify' -DeleteInfo '' }
})

$btnExport.Add_Click({
    $r = $script:Result
    # The quick search has no search name: use a time stamp.
    $stem = if ($r.SearchName) { $r.SearchName } else { 'EMT_quick_{0:yyyyMMdd_HHmmss}' -f (Get-Date) }
    Export-EmtSets -Stem $stem -Sets ([ordered]@{
        'messages'   = @($r.Messages)
        'tracking'   = @($r.Events | Select-Object Timestamp, ServerHostname, EventId, Source, Sender,
                            @{ n = 'Recipients'; e = { $_.Recipients -join '; ' } }, MessageSubject, MessageId)
        'recipients' = @($r.Recipients)
        'found'      = @($r.Found)
        'missing'    = @($r.Missing | ForEach-Object { [pscustomobject]@{ Mailbox = $_ } })
        'check'      = @($script:Check)
        'errors'     = @(@($r.TrackingErrors | ForEach-Object { [pscustomobject]@{ Source = 'Tracking'; Error = $_ } }) +
                         @($(if ($r.Errors) { [pscustomobject]@{ Source = 'Search'; Error = $r.Errors } })))
    })
})

$btnCleanup.Add_Click({
    $days = [Microsoft.VisualBasic.Interaction]::InputBox((T 'dlg.cleanupPrompt'), (T 'dlg.cleanupTitle'), '7')
    if (-not $days) { return }
    if ($days -notmatch '^\d+$') { Write-EmtLog (T 'log.badDays') 'Warn'; return }
    Start-EmtOperation -Name 'ListOld' -Params @{ Days = [int]$days } -StatusText (T 'status.listOld') -OnSuccess {
        param($list)
        $list = @($list | Where-Object { $_ })
        if (-not $list.Count) { Write-EmtLog (T 'log.noOldSearches') 'Ok'; return }
        $names = @($list | ForEach-Object Name)
        $text = (T 'dlg.cleanupList' $names.Count) + "`r`n`r`n" + ($names -join "`r`n")
        if (-not (Show-EmtConfirm -Title (T 'dlg.cleanupTitle') -Text $text -OkText (T 'dlg.cleanupOk'))) { return }
        $script:PendingCleanup = $names
    }
})

$btnCancel.Add_Click({
    if ($script:Job) {
        $Sync.Cancel = $true
        Write-EmtLog (T 'log.cancelRequested') 'Warn'
    }
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 400
$timer.Add_Tick({
    $line = $null
    while ($Sync.Log.TryDequeue([ref]$line)) { Add-EmtLogLine $line }
    if ($script:Job) {
        $progress.Value = [int]$Sync.Progress
        $st = [string]$Sync.Status
        # A long status text can be cut: the tooltip shows the full text.
        if ($lblStatus.Text -ne $st) { $lblStatus.Text = $st; $tips.SetToolTip($lblStatus, $st) }
        if ($script:Job.Handle.IsCompleted) {
            # Stop the timer: a result handler can open a modal dialog (no re-entry).
            $timer.Stop()
            try {
                Complete-EmtOperation
                while ($Sync.Log.TryDequeue([ref]$line)) { Add-EmtLogLine $line }
            } finally { $timer.Start() }
        }
    }
    elseif ($script:PendingCleanup) {
        $script:CleanupNames = $script:PendingCleanup
        $script:PendingCleanup = $null
        Start-EmtOperation -Name 'RemoveOld' -Params @{ Names = $script:CleanupNames } -StatusText (T 'status.removeOld') -OnSuccess {
            Write-EmtAudit -Path $AuditPath -Record @{
                Server = $script:ConnectedServer; Action = 'CleanupSearches'; Result = 'OK'
                Details = ($script:CleanupNames -join ' ')
            }
        }
    }
})

foreach ($c in $tbSenders, $tbSubject, $tbRecipient, $tbMessageId, $tbAttachment, $tbKeywords) {
    $c.Add_TextChanged({ Invoke-EmtInvalidate })
}
$dtStart.Add_ValueChanged({ Invoke-EmtInvalidate })
$dtEnd.Add_ValueChanged({ Invoke-EmtInvalidate })
$chkForwards.Add_CheckedChanged({ Invoke-EmtInvalidate })
# "All mailboxes" and "Only the recipient mailbox" exclude each other.
$chkAll.Add_CheckedChanged({ if ($chkAll.Checked) { $chkOnly.Checked = $false }; Invoke-EmtInvalidate })
$chkOnly.Add_CheckedChanged({ if ($chkOnly.Checked) { $chkAll.Checked = $false }; Invoke-EmtInvalidate })
$tbRecipient.Add_TextChanged({ $chkOnly.Enabled = [bool]$tbRecipient.Text.Trim() })

$form.Add_Shown({
    # Show the window in front of the other windows once. TopMost stays off: dialogs of other programs are not hidden.
    $form.TopMost = $true
    $form.Activate()
    $form.TopMost = $false
    $split.SplitterDistance = [math]::Max(200, $split.Height - 140)
    $trSplit.SplitterDistance = [int]($trSplit.Height * 0.45)
    Write-EmtLog (T 'log.started' $LogPath)
    Write-EmtLog (T 'log.auditFile' $AuditPath)
    foreach ($m in $script:Migrated) { Write-EmtLog (T 'log.migrated' $m) 'Ok' }
    if ($tbServer.Text) { $btnConnect.PerformClick() }
})

$form.Add_FormClosing({
    param($s, $e)
    if ($script:Job) {
        $a = [System.Windows.Forms.MessageBox]::Show($form, (T 'dlg.closeRunning'),
            (T 'app.short'), 'YesNo', [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($a -ne 'Yes') { $e.Cancel = $true; return }
        $Sync.Cancel = $true
        # BeginStop does not block the window. Stop() can wait for a long remote call.
        try { [void]$script:Job.PS.BeginStop($null, $null) } catch { }
    }
    $timer.Stop()
    Save-EmtSettings
})

#endregion

#region ---------- Start ----------

$script:Loading = $true
$defaultServer = if ($Server) { $Server }
                 elseif ($settings.ContainsKey('Server') -and $settings.Server) { [string]$settings.Server }
                 elseif ($env:ExchangeInstallPath) { [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName }
                 else { '' }
$tbServer.Text = $defaultServer
if ($settings.ContainsKey('IncludeForwards')) { $chkForwards.Checked = [bool]$settings.IncludeForwards }
if ($settings.ContainsKey('AllMailboxes'))    { $chkAll.Checked = [bool]$settings.AllMailboxes }
if ($settings.ContainsKey('TimeoutMinutes'))  { $nudTimeout.Value = [math]::Max(5, [math]::Min(240, [int]$settings.TimeoutMinutes)) }
$cmbLanguage.SelectedIndex = [array]::IndexOf($script:LanguageCodes, [string]$Sync.Lang)
# The last window size and position. Only if the window is on a screen now (a monitor can be removed).
if ($settings.ContainsKey('Window') -and $settings.Window) {
    try {
        $wb = $settings.Window
        $rect = New-Object System.Drawing.Rectangle([int]$wb.X, [int]$wb.Y,
            [math]::Max([int]$wb.Width, $form.MinimumSize.Width), [math]::Max([int]$wb.Height, $form.MinimumSize.Height))
        $onScreen = $false
        foreach ($scr in [System.Windows.Forms.Screen]::AllScreens) {
            $part = [System.Drawing.Rectangle]::Intersect($scr.WorkingArea, $rect)
            if ($part.Width -ge 200 -and $part.Height -ge 100) { $onScreen = $true }
        }
        if ($onScreen) {
            $form.StartPosition = 'Manual'
            $form.Bounds = $rect
            if ([bool]$wb.Maximized) { $form.WindowState = 'Maximized' }
        }
    } catch { }
}
foreach ($pair in @(@($dtTrStart, $dtTrEnd), @($dtStart, $dtEnd))) {
    $pair[0].Value = (Get-Date).Date
    $pair[1].Value = (Get-Date).Date.AddDays(1).AddMinutes(-1)
}
$form.AcceptButton = $btnTrFind
Update-EmtCriteriaLabels    # after the labels are on the form: they have the font of the form now
Update-EmtConnIndicator
Update-EmtDeleteSource
$script:Loading = $false

Set-UiState
$timer.Start()
try {
    [void]$form.ShowDialog()
} finally {
    $timer.Stop()
    if ($script:Job) {
        # An operation still stops: do not wait for it. Exchange closes the idle session itself.
        try { $Runspace.CloseAsync() } catch { }
    } else {
        try {
            $ps = [powershell]::Create()
            $ps.Runspace = $Runspace
            [void]$ps.AddScript('if ($global:EmtExSession) { Remove-PSSession $global:EmtExSession -ErrorAction SilentlyContinue }')
            [void]$ps.Invoke()
            $ps.Dispose()
        } catch { }
        $Runspace.Close()
        $Runspace.Dispose()
    }
    $form.Dispose()
    Remove-Variable EmtSync -Scope Global -ErrorAction SilentlyContinue
}

#endregion
