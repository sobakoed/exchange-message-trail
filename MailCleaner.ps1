#Requires -Version 5.1
<#
.SYNOPSIS
    MailCleaner: a GUI to find and remove misdirected e-mail in Exchange Server 2016/2019 (on-premises).

.DESCRIPTION
    Use this tool to find a message, see its recipients, and remove it from mailboxes.
    All Exchange work runs in a background runspace. The window does not freeze.
    The user interface languages are the files in the folder "Lang" (*.lng).
    See README.md for requirements, roles and the procedure.

.PARAMETER Server
    The FQDN of the Exchange server. If you do not set it, the tool uses the last value.

.PARAMETER Language
    The language code of a file in "Lang" (for example ru or en). If you do not set it,
    the tool uses the last value, or the Windows language.

.EXAMPLE
    .\MailCleaner.ps1 -Server exch01.contoso.local -Language en
#>
[CmdletBinding()]
param([string] $Server, [string] $Language)

$ErrorActionPreference = 'Stop'

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Start MailCleaner in STA mode: powershell.exe -STA -File MailCleaner.ps1 (or use MailCleaner.cmd).'
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

$CorePath = Join-Path $PSScriptRoot 'src\MailCleaner.Core.ps1'
. $CorePath

#region ---------- Paths, settings, shared state ----------

$AppDir       = Join-Path $env:USERPROFILE 'MailCleaner'
$SettingsPath = Join-Path $env:APPDATA 'MailCleaner\settings.json'
$AuditPath    = Join-Path $AppDir 'audit.csv'
$LogPath      = Join-Path $AppDir ('logs\MailCleaner_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
foreach ($d in (Split-Path $SettingsPath), (Split-Path $LogPath)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

$Sync = [hashtable]::Synchronized(@{
    Log      = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    Progress = 0
    Status   = ''
    Cancel   = $false
    Lang     = 'en'      # the window and the background runspace read the language here
})
$global:McSync = $Sync

$script:Caps    = $null      # capabilities after Connect
$script:Result  = $null      # result of the last search
$script:Check   = $null      # result of the last verification
$script:Job     = $null      # running background operation
$script:Loading = $false     # true while the UI fills controls (no invalidation)
$script:ConnectedServer = ''
$script:PendingCleanup  = $null    # names confirmed for cleanup; the timer starts the removal
$script:CleanupNames    = $null
$script:GridNeedsSize   = @{}      # grid -> true: set the initial column widths after the next binding
$script:GridHeaders     = @{}      # grid -> ordered map "property" -> "text key" of the column header
$script:Deleted         = ''       # last delete on the current result: '', 'Soft' or 'Hard'
$script:Localized       = New-Object System.Collections.Generic.List[object]   # control + text key
$script:LocalizedTips   = New-Object System.Collections.Generic.List[object]   # control + tooltip key
$script:LanguageCodes   = @()      # language codes in the order of the language list

function Read-McSettings {
    $s = @{}
    if (Test-Path $SettingsPath) {
        try {
            $j = Get-Content $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $j.PSObject.Properties) { $s[$p.Name] = $p.Value }
        } catch { }
    }
    $s
}

function Save-McSettings {
    # Senders and subject are not kept: they can be confidential.
    $s = [ordered]@{
        Server          = $tbServer.Text.Trim()
        Language        = [string]$Sync.Lang
        IncludeForwards = $chkForwards.Checked
        AllMailboxes    = $chkAll.Checked
        TimeoutMinutes  = [int]$nudTimeout.Value
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
$Runspace.SessionStateProxy.SetVariable('McSync', $Sync)

$OpScript = @'
param($CorePath, $Name, $Params)
. $CorePath
Invoke-McOperation -Name $Name -Params $Params
'@

function Start-McOperation {
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

function Complete-McOperation {
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
    foreach ($e in $job.PS.Streams.Error) { Write-McLog ([string]$e) 'Warn' }
    $job.PS.Dispose()
    $progress.Value = 0
    $lblStatus.Text = ''

    if ($err) {
        if ($Sync.Cancel) {
            Write-McLog (T 'log.opCancelled' $job.Name) 'Warn'
        } else {
            Write-McLog (T 'log.opFailed' $job.Name $err.Message) 'Error'
            [void][System.Windows.Forms.MessageBox]::Show($form, $err.Message, (T 'dlg.error'),
                'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        # A delete that stops can be partial: the audit must show it.
        if ($job.Name -in 'SoftDelete', 'HardDelete') {
            try {
                Write-McAudit -Path $AuditPath -Record @{
                    Server = $script:ConnectedServer; Action = $job.Name
                    SearchName = [string]$job.Params.SearchName; Query = [string]$job.Params.Query
                    Mailboxes = @($job.Params.Mailboxes).Count
                    Result = $(if ($Sync.Cancel) { 'Cancelled' } else { 'Failed' }); Details = $err.Message
                }
            } catch { Write-McLog (T 'log.auditFailed' $_.Exception.Message) 'Error' }
        }
    } else {
        try { & $job.OnSuccess $result }
        catch { Write-McLog (T 'log.handlerFailed' $_.Exception.Message) 'Error' }
    }
    Set-UiState
}

#endregion

#region ---------- Localization of controls ----------

function Set-McLoc {
    <# Registers a control for the language change and sets its text now. Returns the control. #>
    param($Control, [string] $Key)
    $script:Localized.Add([pscustomobject]@{ Control = $Control; Key = $Key })
    $Control.Text = T $Key
    $Control
}

function Set-McTip {
    param($Control, [string] $Key)
    $script:LocalizedTips.Add([pscustomobject]@{ Control = $Control; Key = $Key })
    $tips.SetToolTip($Control, (T $Key))
}

function Update-McLanguage {
    <# Applies the current language to all controls, headers and result views. #>
    $form.Text = T 'app.title'
    foreach ($x in $script:Localized)     { $x.Control.Text = T $x.Key }
    foreach ($x in $script:LocalizedTips) { $tips.SetToolTip($x.Control, (T $x.Key)) }
    foreach ($d in $dtStart, $dtEnd) { $d.CustomFormat = T 'fmt.picker' }
    Update-McCapsLabel
    if ($script:Result) { Show-McResultViews $script:Result } else { Clear-McResultViews }
    if ($script:Check)  { Show-McCheckView $script:Check }
    Update-McTabTitles
    Set-UiState
}

#endregion

#region ---------- UI helpers ----------

function New-McLabel([string] $Key) {
    $l = New-Object System.Windows.Forms.Label
    $l.AutoSize = $true; $l.Anchor = 'Left'; $l.Margin = '3,7,3,3'
    if ($Key) { [void](Set-McLoc $l $Key) }
    $l
}

function New-McTextBox {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Dock = 'Fill'; $t.Margin = '3,4,12,3'
    $t
}

function New-McButton([string] $Key, [int] $Width = 150) {
    # AutoSize with a minimum width: a longer translation makes the button wider, not cut.
    $b = New-Object System.Windows.Forms.Button
    $b.AutoSize = $true; $b.AutoSizeMode = 'GrowOnly'
    $b.MinimumSize = New-Object System.Drawing.Size($Width, 28)
    $b.Height = 28; $b.Margin = '3,3,6,3'
    if ($Key) { [void](Set-McLoc $b $Key) }
    $b
}

function New-McCheckBox([string] $Key) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.AutoSize = $true; $c.Margin = '3,7,12,3'
    [void](Set-McLoc $c $Key)
    $c
}

function New-McDatePicker {
    $d = New-Object System.Windows.Forms.DateTimePicker
    $d.Format = 'Custom'; $d.CustomFormat = T 'fmt.picker'; $d.Width = 135; $d.Margin = '3,4,6,3'
    $d
}

function New-McGrid([scriptblock] $RowColor) {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'
    # Auto-size modes lock the column width. Use None: the user can drag the column borders.
    $g.AutoSizeColumnsMode = 'None'
    $g.AllowUserToResizeColumns = $true
    $g.AllowUserToOrderColumns = $true
    $g.ClipboardCopyMode = 'EnableWithAutoHeaderText'    # Ctrl+C copies the selected rows with headers
    $g.BackgroundColor = [System.Drawing.SystemColors]::Window
    $g.Tag = $RowColor
    $g.Add_DataBindingComplete({
        param($s, $e)
        foreach ($c in $s.Columns) {
            if ($c.ValueType -eq [datetime]) { $c.DefaultCellStyle.Format = T 'fmt.datetime' }
            $c.AutoSizeMode = 'None'
            $c.Resizable = 'True'
        }
        Update-McGridHeaders $s
        Resize-McGridColumns $s
    })
    # A grid on a hidden tab cannot measure its cells. Size it when the tab opens.
    $g.Add_VisibleChanged({ param($s, $e) Resize-McGridColumns $s })
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

function Update-McGridHeaders($Grid) {
    <# The DataTable columns have the property names. The headers come from the current language. #>
    if (-not $script:GridHeaders.ContainsKey($Grid)) { return }
    $map = $script:GridHeaders[$Grid]
    foreach ($c in $Grid.Columns) {
        if ($map.Contains($c.Name)) { $c.HeaderText = T $map[$c.Name] }
    }
}

function Resize-McGridColumns($Grid) {
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
    foreach ($c in $Grid.Columns) {
        if ($c.Width -gt 450) { $c.Width = 450 }
        if ($c.Width -lt 60)  { $c.Width = 60 }
        $used += $c.Width
    }
    $free = $Grid.ClientSize.Width - $used - 4
    if ($free -gt 0) { $Grid.Columns[$Grid.Columns.Count - 1].Width += $free }
}

function Set-McGridData {
    <# Binds objects to a grid. $Columns: ordered map "Property" -> "header text key". #>
    param($Grid, [object[]] $Items = @(), [System.Collections.Specialized.OrderedDictionary] $Columns)
    $items = @($Items | Where-Object { $_ -ne $null })
    $dt = New-Object System.Data.DataTable
    foreach ($prop in $Columns.Keys) {
        $sample = $items | ForEach-Object { $_.$prop } | Where-Object { $_ -ne $null } | Select-Object -First 1
        $type = [string]
        if ($sample -is [int] -or $sample -is [long]) { $type = [long] }
        elseif ($sample -is [double] -or $sample -is [decimal]) { $type = [double] }
        elseif ($sample -is [datetime]) { $type = [datetime] }
        [void]$dt.Columns.Add($prop, $type)
    }
    foreach ($it in $items) {
        $row = $dt.NewRow()
        foreach ($prop in $Columns.Keys) {
            $v = $it.$prop
            if ($null -eq $v) { $row[$prop] = [DBNull]::Value }
            elseif ($dt.Columns[$prop].DataType -eq [string]) {
                $row[$prop] = if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { (@($v) -join '; ') } else { [string]$v }
            }
            else { $row[$prop] = $v }
        }
        $dt.Rows.Add($row)
    }
    $script:GridHeaders[$Grid] = $Columns
    $script:GridNeedsSize[$Grid] = $true
    $Grid.DataSource = $dt
    Update-McGridHeaders $Grid
}

function Show-McConfirm {
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

    $ok = New-McButton '' 160
    $ok.Text = $OkText
    $ok.DialogResult = 'OK'
    $cancel = New-McButton '' 120
    $cancel.Text = T 'dlg.cancel'
    $cancel.DialogResult = 'Cancel'
    $dlg.CancelButton = $cancel

    if ($RequireWord) {
        $prompt = New-McLabel ''
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

function Add-McLogLine([string] $Line) {
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

#endregion

#region ---------- Language (before the layout: the controls read the texts) ----------

$settings = Read-McSettings
$available = @(Get-McLanguages)
$script:LanguageCodes = @($available | ForEach-Object Code)
$wanted = if ($Language) { $Language.ToLowerInvariant() }
          elseif ($settings.ContainsKey('Language') -and $settings.Language) { ([string]$settings.Language).ToLowerInvariant() }
          else { Get-McDefaultLanguage }
$Sync.Lang = if ($wanted -in $script:LanguageCodes) { $wanted } elseif ('en' -in $script:LanguageCodes) { 'en' } else { @($script:LanguageCodes + 'en')[0] }

#endregion

#region ---------- Form layout ----------

$form = New-Object System.Windows.Forms.Form
$form.Text = T 'app.title'
$form.Size = '1280,860'
$form.MinimumSize = '1100,700'
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$tips = New-Object System.Windows.Forms.ToolTip

# --- Connection bar ---
$pnlConn = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlConn.Dock = 'Top'; $pnlConn.AutoSize = $true; $pnlConn.Padding = '6,6,6,0'; $pnlConn.WrapContents = $true
$tbServer = New-Object System.Windows.Forms.TextBox
$tbServer.Width = 240; $tbServer.Margin = '3,4,6,3'
$btnConnect = New-McButton 'conn.connect' 140
$nudTimeout = New-Object System.Windows.Forms.NumericUpDown
$nudTimeout.Minimum = 5; $nudTimeout.Maximum = 240; $nudTimeout.Value = 60; $nudTimeout.Width = 60; $nudTimeout.Margin = '3,4,12,3'
$btnCleanup = New-McButton 'conn.cleanup' 180
$cmbLanguage = New-Object System.Windows.Forms.ComboBox
$cmbLanguage.DropDownStyle = 'DropDownList'; $cmbLanguage.Width = 120; $cmbLanguage.Margin = '3,4,12,3'
foreach ($l in $available) { [void]$cmbLanguage.Items.Add($l.Name) }
$lblCaps = New-McLabel ''
$pnlConn.Controls.AddRange(@(
    (New-McLabel 'conn.server'), $tbServer, $btnConnect,
    (New-McLabel 'conn.timeout'), $nudTimeout, $btnCleanup,
    (New-McLabel 'conn.language'), $cmbLanguage, $lblCaps))

# --- Criteria ---
$grpCrit = New-Object System.Windows.Forms.GroupBox
[void](Set-McLoc $grpCrit 'crit.group')
$grpCrit.Dock = 'Top'; $grpCrit.Height = 178; $grpCrit.Padding = '8,4,8,4'

$tlp = New-Object System.Windows.Forms.TableLayoutPanel
$tlp.Dock = 'Fill'; $tlp.ColumnCount = 6; $tlp.RowCount = 4
# The middle column holds the end date and the quick period buttons: it gets more space.
foreach ($pct in 29, 42, 29) {
    [void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', $pct)))
}

$tbSenders    = New-McTextBox
$tbSubject    = New-McTextBox
$tbRecipient  = New-McTextBox
$dtStart      = New-McDatePicker
$dtEnd        = New-McDatePicker
$tbMessageId  = New-McTextBox
$tbAttachment = New-McTextBox
$tbKeywords   = New-McTextBox

# Period: end date and the quick period buttons in one cell.
$btnToday     = New-McButton 'crit.today' 72
$btnYesterday = New-McButton 'crit.yesterday' 64
foreach ($b in $btnToday, $btnYesterday) { $b.MinimumSize = New-Object System.Drawing.Size($b.MinimumSize.Width, 24); $b.Height = 24; $b.Margin = '0,3,4,3' }
$pnlEnd = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlEnd.AutoSize = $true; $pnlEnd.WrapContents = $false; $pnlEnd.Margin = '0,0,0,0'
$pnlEnd.Controls.AddRange(@($dtEnd, $btnToday, $btnYesterday))

# Row 0: who and what.  Row 1: when and message ID.  Row 2: content.  Row 3: options and Find.
$tlp.Controls.Add((New-McLabel 'crit.senders'), 0, 0);    $tlp.Controls.Add($tbSenders, 1, 0)
$tlp.Controls.Add((New-McLabel 'crit.subject'), 2, 0);    $tlp.Controls.Add($tbSubject, 3, 0)
$tlp.Controls.Add((New-McLabel 'crit.recipient'), 4, 0);  $tlp.Controls.Add($tbRecipient, 5, 0)
$tlp.Controls.Add((New-McLabel 'crit.start'), 0, 1);      $tlp.Controls.Add($dtStart, 1, 1)
$tlp.Controls.Add((New-McLabel 'crit.end'), 2, 1);        $tlp.Controls.Add($pnlEnd, 3, 1)
$tlp.Controls.Add((New-McLabel 'crit.messageId'), 4, 1);  $tlp.Controls.Add($tbMessageId, 5, 1)
$tlp.Controls.Add((New-McLabel 'crit.attachment'), 0, 2); $tlp.Controls.Add($tbAttachment, 1, 2)
$tlp.Controls.Add((New-McLabel 'crit.keywords'), 2, 2);   $tlp.Controls.Add($tbKeywords, 3, 2)

$pnlSearch = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlSearch.Dock = 'Fill'; $pnlSearch.AutoSize = $true; $pnlSearch.WrapContents = $false
$chkForwards = New-McCheckBox 'crit.forwards'
$chkForwards.Checked = $true
$chkOnly = New-McCheckBox 'crit.only'
$chkOnly.Enabled = $false
$chkAll = New-McCheckBox 'crit.all'
$chkAll.Margin = '3,7,24,3'
$btnSearch = New-McButton 'crit.find' 130
$btnSearch.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
$pnlSearch.Controls.AddRange(@($chkForwards, $chkOnly, $chkAll, $btnSearch))
$tlp.Controls.Add($pnlSearch, 0, 3)
$tlp.SetColumnSpan($pnlSearch, 6)
$grpCrit.Controls.Add($tlp)
$form.AcceptButton = $btnSearch

# --- Summary line ---
$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Dock = 'Top'; $lblSummary.Height = 28; $lblSummary.Padding = '10,6,6,0'; $lblSummary.AutoEllipsis = $true
$lblSummary.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
$lblSummary.Text = T 'sum.none'

# --- Tabs with grids + log ---
$colorExternal = [System.Drawing.Color]::MistyRose
$colorNote     = [System.Drawing.Color]::LightYellow
$colorClean    = [System.Drawing.Color]::Honeydew

# Row colors use the DataTable column names (the property names), not the header texts.
$gridMessages   = New-McGrid { param($r) if ([string]$r['Note']) { $colorNote } }
$gridTracking   = New-McGrid $null
$gridRecipients = New-McGrid { param($r) if ([string]$r['Kind'] -eq (T 'kind.External')) { $colorExternal } }
$gridFound      = New-McGrid $null
$gridMissing    = New-McGrid $null
$gridCheck      = New-McGrid {
    param($r)
    if ([long]$r['Visible'] -lt 0) { $colorNote }
    elseif ([long]$r['Visible'] -gt 0) { $colorExternal }
    else { $colorClean }
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

$log = New-Object System.Windows.Forms.RichTextBox
$log.Dock = 'Fill'; $log.ReadOnly = $true; $log.Font = New-Object System.Drawing.Font('Consolas', 9)
$log.BackColor = [System.Drawing.SystemColors]::Window; $log.WordWrap = $false

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'; $split.Orientation = 'Horizontal'
$split.Panel1.Controls.Add($tabs)
$split.Panel2.Controls.Add($log)

# --- Actions bar ---
# Left: actions on the result. Right: the running operation (status, progress, cancel).
$pnlBottom = New-Object System.Windows.Forms.TableLayoutPanel
$pnlBottom.Dock = 'Bottom'; $pnlBottom.AutoSize = $true; $pnlBottom.ColumnCount = 2; $pnlBottom.RowCount = 1
$pnlBottom.Padding = '6,4,6,4'
[void]$pnlBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
[void]$pnlBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))

$pnlActions = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlActions.AutoSize = $true; $pnlActions.WrapContents = $false; $pnlActions.Margin = '0,0,0,0'
$btnSoft   = New-McButton 'act.soft' 140
$btnHard   = New-McButton 'act.hard' 165
$btnHard.ForeColor = [System.Drawing.Color]::Firebrick
$btnVerify = New-McButton 'act.verify' 100
$btnExport = New-McButton 'act.export' 115
$btnExport.Margin = '18,3,6,3'
$pnlActions.Controls.AddRange(@($btnSoft, $btnHard, $btnVerify, $btnExport))

$pnlRun = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlRun.Dock = 'Fill'; $pnlRun.AutoSize = $true; $pnlRun.WrapContents = $false
$pnlRun.FlowDirection = 'RightToLeft'; $pnlRun.Margin = '0,0,0,0'
$btnCancel = New-McButton 'act.cancel' 90
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Width = 140; $progress.Height = 20; $progress.Margin = '6,7,6,3'
$lblStatus = New-McLabel ''
$lblStatus.AutoSize = $false; $lblStatus.Width = 260; $lblStatus.Height = 20; $lblStatus.AutoEllipsis = $true
$lblStatus.TextAlign = 'MiddleRight'
$pnlRun.Controls.AddRange(@($btnCancel, $progress, $lblStatus))

$pnlBottom.Controls.Add($pnlActions, 0, 0)
$pnlBottom.Controls.Add($pnlRun, 1, 0)

Set-McTip $tbSenders    'tip.senders'
Set-McTip $chkAll       'tip.all'
Set-McTip $btnToday     'tip.today'
Set-McTip $btnYesterday 'tip.yesterday'
Set-McTip $nudTimeout   'tip.timeout'
Set-McTip $btnCleanup   'tip.cleanup'
Set-McTip $btnExport    'tip.export'
Set-McTip $tbRecipient  'tip.recipient'
Set-McTip $chkOnly      'tip.only'
Set-McTip $tbMessageId  'tip.messageId'
Set-McTip $chkForwards  'tip.forwards'
Set-McTip $btnSoft      'tip.soft'
Set-McTip $btnHard      'tip.hard'
Set-McTip $btnVerify    'tip.verify'
Set-McTip $cmbLanguage  'tip.language'

# Dock order: the control added last docks first (top-most).
$form.Controls.Add($split)
$form.Controls.Add($pnlBottom)
$form.Controls.Add($lblSummary)
$form.Controls.Add($grpCrit)
$form.Controls.Add($pnlConn)

#endregion

#region ---------- State machine ----------

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

function Update-McTabTitles {
    <# Tab title = name and row count, e.g. "Found in mailboxes (42)". #>
    foreach ($k in $pages.Keys) {
        $dt = $pages[$k].DataSource
        $name = T "tab.$k"
        $tabPage[$k].Text = if ($dt -is [System.Data.DataTable]) { "$name ($($dt.Rows.Count))" } else { $name }
    }
}

function Update-McCapsLabel {
    $c = $script:Caps
    if (-not $c) { $lblCaps.Text = T 'conn.notConnected'; $tips.SetToolTip($lblCaps, ''); return }
    $mark = { param($v) if ($v) { '✓' } else { '✗' } }
    $lblCaps.Text = T 'conn.caps' $script:ConnectedServer (& $mark $c.Tracking) (& $mark $c.Search) (& $mark $c.Verify) `
                                  (& $mark $c.SoftDelete) (& $mark $c.HardDelete)
    $tips.SetToolTip($lblCaps, $(if (-not $c.HardDelete) { T 'conn.noHardTip' } else { '' }))
}

function Clear-McResultViews {
    foreach ($g in $pages.Values) { $g.DataSource = $null }
    Update-McTabTitles
    $lblSummary.Text = T 'sum.none'
    $lblSummary.ForeColor = [System.Drawing.SystemColors]::ControlText
    $tips.SetToolTip($lblSummary, '')
}

function Clear-McResults {
    $script:Result  = $null
    $script:Check   = $null
    $script:Deleted = ''
    Clear-McResultViews
}

function Invoke-McInvalidate {
    if ($script:Loading -or -not $script:Result) { return }
    Clear-McResults
    Write-McLog (T 'log.invalidated') 'Warn'
    Set-UiState
}

function Get-McCriteria {
    @{
        Senders         = @(Split-McAddressList $tbSenders.Text)
        Subject         = $tbSubject.Text.Trim()
        Recipient       = $tbRecipient.Text.Trim().ToLowerInvariant()
        MessageId       = $tbMessageId.Text.Trim()
        Attachment      = $tbAttachment.Text.Trim()
        Keywords        = $tbKeywords.Text.Trim()
        Start           = $dtStart.Value
        End             = $dtEnd.Value
        IncludeForwards = $chkForwards.Checked
        AllMailboxes    = $chkAll.Checked
        OnlyRecipient   = $chkOnly.Checked -and $chkOnly.Enabled
    }
}

#endregion

#region ---------- Result views ----------

function Show-McResultViews($r) {
    <# Fills the grids and the summary line. No audit: it also runs after a language change. #>
    $msgs = @($r.Messages | ForEach-Object {
        [pscustomobject]@{
            Time = $_.Time; Sender = $_.Sender; Subject = $_.Subject; Recipients = $_.Recipients
            Note = $(if ($_.ServerGenerated) { T 'note.serverGenerated' } else { '' }); MessageId = $_.MessageId
        }
    })
    Set-McGridData $gridMessages $msgs ([ordered]@{
        Time = 'col.time'; Sender = 'col.sender'; Subject = 'col.subject'; Recipients = 'col.recipientCount'
        Note = 'col.note'; MessageId = 'col.messageId' })
    Set-McGridData $gridTracking @($r.Events) ([ordered]@{
        Timestamp = 'col.time'; ServerHostname = 'col.server'; EventId = 'col.event'; Source = 'col.source'
        Sender = 'col.sender'; Recipients = 'col.recipients'; MessageSubject = 'col.subject'; MessageId = 'col.messageId' })
    $rcpt = @($r.Recipients | ForEach-Object {
        [pscustomobject]@{ Address = $_.Address; Kind = (T "kind.$($_.Kind)"); Mailbox = $_.Mailbox; Type = $_.Type }
    } | Sort-Object Kind, Address)
    Set-McGridData $gridRecipients $rcpt ([ordered]@{
        Address = 'col.address'; Kind = 'col.kind'; Mailbox = 'col.mailbox'; Type = 'col.exchangeType' })
    Set-McGridData $gridFound @($r.Found) ([ordered]@{ Mailbox = 'col.mailbox'; Items = 'col.items'; SizeKB = 'col.sizeKB' })
    $miss = @($r.Missing | ForEach-Object { [pscustomobject]@{ Mailbox = $_; Reason = (T 'missing.reason') } })
    Set-McGridData $gridMissing $miss ([ordered]@{ Mailbox = 'col.mailbox'; Reason = 'col.reason' })

    $unindexedText = if ($r.Unindexed -lt 0) { T 'sum.na' } else { [string]$r.Unindexed }
    $methodText    = if ($r.Method -eq 'SearchMailbox') { T 'sum.quick' } else { 'Compliance Search' }
    $scopeText     = if (@($r.Locations) -contains 'All') { T 'common.all' } else { [string]@($r.Locations).Count }
    $lblSummary.Text = T 'sum.text' $r.Items @($r.Found).Count $scopeText $r.SizeKB @($r.External).Count `
                                    $unindexedText $methodText $r.Seconds
    $lblSummary.ForeColor = [System.Drawing.SystemColors]::ControlText
    if ($r.Errors) {
        $lblSummary.Text = (T 'sum.errors') + $lblSummary.Text
        $lblSummary.ForeColor = [System.Drawing.Color]::Firebrick
    }
    $tips.SetToolTip($lblSummary, $(if ($r.Errors) { "$($lblSummary.Text)`r`n`r`n$($r.Errors)" } else { $lblSummary.Text }))
    Update-McTabTitles
}

function Show-McSearchResult($r) {
    $script:Result = $r
    $script:Check  = $null
    $gridCheck.DataSource = $null
    Show-McResultViews $r
    $tabs.SelectedTab = $tabPage['found']

    Write-McAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = 'Search'; SearchName = $r.SearchName; Query = $r.Query
        Mailboxes = @($r.Found).Count; Items = $r.Items
        Result = $(if ($r.Errors) { 'Errors' } else { 'OK' })
        Details = "Method=$($r.Method); Seconds=$($r.Seconds); External=$(@($r.External | ForEach-Object Address) -join ' '); Unindexed=$($r.Unindexed)"
    }
}

function Show-McCheckView([object[]] $Check) {
    Set-McGridData $gridCheck @($Check) ([ordered]@{
        Mailbox = 'col.mailbox'; Visible = 'col.visible'; Total = 'col.total'; Note = 'col.note' })
    Update-McTabTitles
}

function Show-McCheckResult([object[]] $Check, [string] $Action, $DeleteInfo) {
    $script:Check = @($Check)
    Show-McCheckView $script:Check
    $tabs.SelectedTab = $tabPage['check']
    $visible = @($script:Check | Where-Object { $_.Visible -gt 0 })
    $hidden  = @($script:Check | Where-Object { $_.Visible -eq 0 -and $_.Total -gt 0 })
    $failed  = @($script:Check | Where-Object { $_.Visible -lt 0 })
    if ($visible.Count) { Write-McLog (T 'log.checkVisible' $visible.Count) 'Warn' }
    else { Write-McLog (T 'log.checkClean') 'Ok' }
    if ($hidden.Count) { Write-McLog (T 'log.checkHidden' $hidden.Count) 'Warn' }
    if ($failed.Count) { Write-McLog (T 'log.checkFailed' $failed.Count) 'Error' }

    $resultText = if ($failed.Count) { 'CheckFailed' } elseif ($visible.Count) { 'Visible' } elseif ($hidden.Count) { 'RecoverableOnly' } else { 'Clean' }
    Write-McAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = $Action; SearchName = $script:Result.SearchName
        Query = $script:Result.Query; Mailboxes = $script:Check.Count; Items = $script:Result.Items
        Result = $resultText; Details = [string]$DeleteInfo
    }
}

function Get-McDeleteText {
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

#region ---------- Event handlers ----------

$btnConnect.Add_Click({
    $srv = $tbServer.Text.Trim()
    if (-not $srv) { [void][System.Windows.Forms.MessageBox]::Show($form, (T 'conn.enterServer'), 'MailCleaner'); return }
    $script:Caps = $null
    Clear-McResults
    Update-McCapsLabel
    Start-McOperation -Name 'Connect' -Params @{ Server = $srv } -StatusText (T 'status.connecting' $srv) -OnSuccess {
        param($caps)
        $script:Caps = $caps
        $script:ConnectedServer = $tbServer.Text.Trim()
        Update-McCapsLabel
        Save-McSettings
        if (-not $caps.HardDelete) { Write-McLog (T 'log.noHardRole') 'Warn' }
        if (-not $caps.Tracking)   { Write-McLog (T 'log.noTrackRole') 'Warn' }
    }
})

$cmbLanguage.Add_SelectedIndexChanged({
    if ($script:Loading -or $cmbLanguage.SelectedIndex -lt 0) { return }
    $code = $script:LanguageCodes[$cmbLanguage.SelectedIndex]
    if ($code -eq $Sync.Lang) { return }
    $Sync.Lang = $code
    Update-McLanguage
    Save-McSettings
    Write-McLog (T 'log.language')
})

$btnToday.Add_Click({
    $dtStart.Value = (Get-Date).Date
    $dtEnd.Value   = (Get-Date).Date.AddDays(1)
})

$btnYesterday.Add_Click({
    $dtStart.Value = (Get-Date).Date.AddDays(-1)
    $dtEnd.Value   = (Get-Date).Date
})

function Show-McPick {
    <# Selection of one or more addresses from the candidates. Returns the selected addresses, or $null on cancel. #>
    param([string] $Text, [object[]] $Candidates, [switch] $Multi, [int] $Max = 25)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = T 'pick.title'; $dlg.Size = '640,420'; $dlg.StartPosition = 'CenterParent'
    $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false; $dlg.Font = $form.Font; $dlg.ShowIcon = $false

    $top = New-McLabel ''
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
    $ok = New-McButton '' 120; $ok.Text = T 'pick.ok'; $ok.DialogResult = 'OK'
    $cancel = New-McButton '' 120; $cancel.Text = T 'dlg.cancel'; $cancel.DialogResult = 'Cancel'
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

function Complete-McAddressResolve {
    <#
        Puts the full addresses in the fields: one candidate = automatic, more = the user selects,
        none = error. Returns $true if the search can start.
    #>
    param([object[]] $Resolved)
    $senders = @(Split-McAddressList $tbSenders.Text)
    $recipient = $tbRecipient.Text.Trim().ToLowerInvariant()
    $map = @{}
    $notFound = @()
    foreach ($x in @($Resolved | Where-Object { $_ })) {
        $cands = @($x.Candidates)
        if (-not $cands.Count) { $notFound += $x.Input; continue }
        if ($cands.Count -eq 1) {
            $map[$x.Input] = @($cands[0].Address)
            Write-McLog (T 'log.resolvedOne' $x.Input $cands[0].Address) 'Ok'
            continue
        }
        # Several senders are possible. The recipient is one mailbox.
        $multi = ($x.Input -in $senders) -and ($x.Input -ne $recipient)
        $picked = Show-McPick -Text $x.Input -Candidates $cands -Multi:$multi -Max $x.Max
        if (-not $picked) { Write-McLog (T 'log.resolveCancelled') 'Warn'; return $false }
        $map[$x.Input] = @($picked)
        Write-McLog (T 'log.resolvedPick' $x.Input ($picked -join ', ')) 'Ok'
    }
    if ($notFound.Count) {
        [void][System.Windows.Forms.MessageBox]::Show($form, (T 'resolve.notFound' ($notFound -join ', ')),
            (T 'resolve.title'), 'OK', [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    $script:Loading = $true     # the fields change, but this is not a change of the criteria by the user
    try {
        $newSenders = @($senders | ForEach-Object { if ($map.ContainsKey($_)) { $map[$_] } else { $_ } } | Sort-Object -Unique)
        $tbSenders.Text = $newSenders -join '; '
        if ($recipient -and $map.ContainsKey($recipient)) { $tbRecipient.Text = @($map[$recipient])[0] }
    } finally { $script:Loading = $false }
    $true
}

function Start-McSearch {
    $crit = Get-McCriteria
    Save-McSettings
    Clear-McResults
    $fmt = T 'fmt.date'
    Write-McLog (T 'log.search' ($crit.Senders -join ', ') $crit.Subject $crit.Start.ToString($fmt) $crit.End.ToString($fmt))
    Start-McOperation -Name 'Search' -Params @{ Criteria = $crit; TimeoutMinutes = [int]$nudTimeout.Value } `
        -StatusText (T 'status.search') -OnSuccess { param($r) Show-McSearchResult $r }
}

$btnSearch.Add_Click({
    $crit = Get-McCriteria
    $problems = @(Test-McCriteria $crit)
    if ($problems.Count) {
        [void][System.Windows.Forms.MessageBox]::Show($form, ($problems -join "`r`n"), (T 'crit.checkTitle'),
            'OK', [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    # A part of an address or a name: find the full address in the address book first.
    $partial = @(@($crit.Senders) + @($crit.Recipient) | Where-Object { $_ -and -not (Test-McFullAddress $_) } | Sort-Object -Unique)
    if ($partial.Count) {
        Start-McOperation -Name 'Resolve' -Params @{ Inputs = $partial } -StatusText (T 'status.resolve') -OnSuccess {
            param($x)
            if (Complete-McAddressResolve -Resolved @($x)) { Start-McSearch }
        }
        return
    }
    Start-McSearch
})

$btnSoft.Add_Click({
    $r = $script:Result
    $mbx = @($r.Found | ForEach-Object Mailbox)
    $warning = Get-McLargeDeleteWarning -Items $r.Items -Messages @($r.Messages).Count
    $text = Get-McDeleteText -r $r -Kind 'Soft' -Warning $warning
    # A large delete needs the typed word also for a soft delete.
    $word = if ($warning) { T 'confirm.word' } else { '' }
    if (-not (Show-McConfirm -Title (T 'confirm.softTitle') -Text $text -OkText (T 'act.soft') -RequireWord $word)) { return }
    Start-McOperation -Name 'SoftDelete' -StatusText (T 'status.soft') -Params @{
        SearchName = $r.SearchName; Mailboxes = $mbx; Query = $r.Query; TimeoutMinutes = [int]$nudTimeout.Value
    } -OnSuccess {
        param($x)
        $script:Deleted = 'Soft'
        $info = "Status=$($x.Delete.Status); Search=$($x.SearchName)"
        if ($x.Delete.Errors) { $info += "; Errors=$($x.Delete.Errors)" }
        Show-McCheckResult -Check @($x.Check) -Action 'SoftDelete' -DeleteInfo $info
    }
})

$btnHard.Add_Click({
    $r = $script:Result
    $mbx = @($r.Found | ForEach-Object Mailbox)
    $expected = @{}
    foreach ($f in @($r.Found)) { $expected[$f.Mailbox] = [int]$f.Items }
    $warning = Get-McLargeDeleteWarning -Items $r.Items -Messages @($r.Messages).Count
    $text = Get-McDeleteText -r $r -Kind 'Hard' -Warning $warning
    if (-not (Show-McConfirm -Title (T 'confirm.hardTitle') -Text $text -OkText (T 'act.hard') -RequireWord (T 'confirm.word'))) { return }
    Start-McOperation -Name 'HardDelete' -StatusText (T 'status.hard') -Params @{
        Mailboxes = $mbx; Query = $r.Query; Expected = $expected; SearchName = $r.SearchName
    } -OnSuccess {
        param($x)
        $script:Deleted = 'Hard'
        $deleted = [int](@($x.Delete) | Measure-Object Deleted -Sum).Sum
        $errors  = @($x.Delete | Where-Object Error).Count
        $info    = "Deleted=$deleted; MailboxErrors=$errors; Processed=$(@($x.Delete).Count)"
        if ($x.Cancelled) {
            Write-McLog (T 'log.hardCancelled' $deleted) 'Warn'
            Write-McAudit -Path $AuditPath -Record @{
                Server = $script:ConnectedServer; Action = 'HardDelete'; SearchName = $script:Result.SearchName
                Query = $script:Result.Query; Mailboxes = @($x.Delete).Count; Items = $deleted
                Result = 'Cancelled'; Details = $info
            }
        } else {
            Show-McCheckResult -Check @($x.Check) -Action 'HardDelete' -DeleteInfo $info
        }
    }
})

$btnVerify.Add_Click({
    $r = $script:Result
    Start-McOperation -Name 'Verify' -StatusText (T 'status.verify') -Params @{
        Mailboxes = @($r.Found | ForEach-Object Mailbox); Query = $r.Query
    } -OnSuccess { param($x) Show-McCheckResult -Check @($x.Check) -Action 'Verify' -DeleteInfo '' }
})

$btnExport.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = T 'dlg.exportFolder'
    $dlg.SelectedPath = $AppDir
    if ($dlg.ShowDialog($form) -ne 'OK') { return }
    $r = $script:Result
    # The quick search has no search name: use a time stamp.
    $stem = if ($r.SearchName) { $r.SearchName } else { 'MC_quick_{0:yyyyMMdd_HHmmss}' -f (Get-Date) }
    $base = Join-Path $dlg.SelectedPath $stem
    $sets = [ordered]@{
        'messages'   = @($r.Messages)
        'tracking'   = @($r.Events | Select-Object Timestamp, ServerHostname, EventId, Source, Sender,
                            @{ n = 'Recipients'; e = { $_.Recipients -join '; ' } }, MessageSubject, MessageId)
        'recipients' = @($r.Recipients)
        'found'      = @($r.Found)
        'missing'    = @($r.Missing | ForEach-Object { [pscustomobject]@{ Mailbox = $_ } })
        'check'      = @($script:Check)
        'errors'     = @(@($r.TrackingErrors | ForEach-Object { [pscustomobject]@{ Source = 'Tracking'; Error = $_ } }) +
                         @($(if ($r.Errors) { [pscustomobject]@{ Source = 'Search'; Error = $r.Errors } })))
    }
    foreach ($k in $sets.Keys) {
        if (@($sets[$k]).Count) {
            $sets[$k] | Export-Csv -Path "$base-$k.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
        }
    }
    Write-McLog (T 'log.exported' $base) 'Ok'
})

$btnCleanup.Add_Click({
    $days = [Microsoft.VisualBasic.Interaction]::InputBox((T 'dlg.cleanupPrompt'), (T 'dlg.cleanupTitle'), '7')
    if (-not $days) { return }
    if ($days -notmatch '^\d+$') { Write-McLog (T 'log.badDays') 'Warn'; return }
    Start-McOperation -Name 'ListOld' -Params @{ Days = [int]$days } -StatusText (T 'status.listOld') -OnSuccess {
        param($list)
        $list = @($list | Where-Object { $_ })
        if (-not $list.Count) { Write-McLog (T 'log.noOldSearches') 'Ok'; return }
        $names = @($list | ForEach-Object Name)
        $text = (T 'dlg.cleanupList' $names.Count) + "`r`n`r`n" + ($names -join "`r`n")
        if (-not (Show-McConfirm -Title (T 'dlg.cleanupTitle') -Text $text -OkText (T 'dlg.cleanupOk'))) { return }
        $script:PendingCleanup = $names
    }
})

$btnCancel.Add_Click({
    if ($script:Job) {
        $Sync.Cancel = $true
        Write-McLog (T 'log.cancelRequested') 'Warn'
    }
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 400
$timer.Add_Tick({
    $line = $null
    while ($Sync.Log.TryDequeue([ref]$line)) { Add-McLogLine $line }
    if ($script:Job) {
        $progress.Value = [int]$Sync.Progress
        $lblStatus.Text = [string]$Sync.Status
        if ($script:Job.Handle.IsCompleted) {
            # Stop the timer: a result handler can open a modal dialog (no re-entry).
            $timer.Stop()
            try {
                Complete-McOperation
                while ($Sync.Log.TryDequeue([ref]$line)) { Add-McLogLine $line }
            } finally { $timer.Start() }
        }
    }
    elseif ($script:PendingCleanup) {
        $script:CleanupNames = $script:PendingCleanup
        $script:PendingCleanup = $null
        Start-McOperation -Name 'RemoveOld' -Params @{ Names = $script:CleanupNames } -StatusText (T 'status.removeOld') -OnSuccess {
            Write-McAudit -Path $AuditPath -Record @{
                Server = $script:ConnectedServer; Action = 'CleanupSearches'; Result = 'OK'
                Details = ($script:CleanupNames -join ' ')
            }
        }
    }
})

foreach ($c in $tbSenders, $tbSubject, $tbRecipient, $tbMessageId, $tbAttachment, $tbKeywords) {
    $c.Add_TextChanged({ Invoke-McInvalidate })
}
$dtStart.Add_ValueChanged({ Invoke-McInvalidate })
$dtEnd.Add_ValueChanged({ Invoke-McInvalidate })
$chkForwards.Add_CheckedChanged({ Invoke-McInvalidate })
# "All mailboxes" and "Only the recipient mailbox" exclude each other.
$chkAll.Add_CheckedChanged({ if ($chkAll.Checked) { $chkOnly.Checked = $false }; Invoke-McInvalidate })
$chkOnly.Add_CheckedChanged({ if ($chkOnly.Checked) { $chkAll.Checked = $false }; Invoke-McInvalidate })
$tbRecipient.Add_TextChanged({ $chkOnly.Enabled = [bool]$tbRecipient.Text.Trim() })

$form.Add_Shown({
    $split.SplitterDistance = [int]($split.Height * 0.66)
    Write-McLog (T 'log.started' $LogPath)
    Write-McLog (T 'log.auditFile' $AuditPath)
    if ($tbServer.Text) { $btnConnect.PerformClick() }
})

$form.Add_FormClosing({
    param($s, $e)
    if ($script:Job) {
        $a = [System.Windows.Forms.MessageBox]::Show($form, (T 'dlg.closeRunning'),
            'MailCleaner', 'YesNo', [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($a -ne 'Yes') { $e.Cancel = $true; return }
        $Sync.Cancel = $true
        # BeginStop does not block the window. Stop() can wait for a long remote call.
        try { [void]$script:Job.PS.BeginStop($null, $null) } catch { }
    }
    $timer.Stop()
    Save-McSettings
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
$dtStart.Value = (Get-Date).Date
$dtEnd.Value   = (Get-Date).Date.AddDays(1)
Update-McCapsLabel
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
            [void]$ps.AddScript('if ($global:McExSession) { Remove-PSSession $global:McExSession -ErrorAction SilentlyContinue }')
            [void]$ps.Invoke()
            $ps.Dispose()
        } catch { }
        $Runspace.Close()
        $Runspace.Dispose()
    }
    $form.Dispose()
    Remove-Variable McSync -Scope Global -ErrorAction SilentlyContinue
}

#endregion
