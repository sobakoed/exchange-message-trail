#Requires -Version 5.1
<#
.SYNOPSIS
    MailCleaner: a GUI to find and remove misdirected e-mail in Exchange Server 2016/2019 (on-premises).

.DESCRIPTION
    Use this tool to find a message, see its recipients, and remove it from mailboxes.
    All Exchange work runs in a background runspace. The window does not freeze.
    See README.md for requirements, roles and the procedure.

.PARAMETER Server
    The FQDN of the Exchange server. If you do not set it, the tool uses the last value.

.EXAMPLE
    .\MailCleaner.ps1 -Server exch01.contoso.local
#>
[CmdletBinding()]
param([string] $Server)

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
    $s = [ordered]@{
        Server          = $tbServer.Text.Trim()
        Senders         = $tbSenders.Text.Trim()
        Subject         = $tbSubject.Text.Trim()
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
        [string] $StatusText = 'Выполняется...'
    )
    if ($script:Job) { return }
    $Sync.Cancel   = $false
    $Sync.Progress = 0
    $Sync.Status   = $StatusText
    $ps = [powershell]::Create()
    $ps.Runspace = $Runspace
    [void]$ps.AddScript($OpScript).AddArgument($CorePath).AddArgument($Name).AddArgument($Params)
    $script:Job = [pscustomobject]@{
        PS = $ps; Handle = $ps.BeginInvoke(); Name = $Name; OnSuccess = $OnSuccess; Started = Get-Date
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
            Write-McLog "Операция «$($job.Name)» отменена." 'Warn'
        } else {
            Write-McLog "Операция «$($job.Name)» завершилась с ошибкой: $($err.Message)" 'Error'
            [void][System.Windows.Forms.MessageBox]::Show($form, $err.Message, 'Ошибка',
                'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        try { & $job.OnSuccess $result }
        catch { Write-McLog "Ошибка обработки результата: $($_.Exception.Message)" 'Error' }
    }
    Set-UiState
}

#endregion

#region ---------- UI helpers ----------

function New-McLabel([string] $Text) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.AutoSize = $true; $l.Anchor = 'Left'; $l.Margin = '3,7,3,3'
    $l
}

function New-McTextBox {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Dock = 'Fill'; $t.Margin = '3,4,12,3'
    $t
}

function New-McButton([string] $Text, [int] $Width = 150) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Width = $Width; $b.Height = 28; $b.Margin = '3,3,6,3'
    $b
}

function New-McDatePicker {
    $d = New-Object System.Windows.Forms.DateTimePicker
    $d.Format = 'Custom'; $d.CustomFormat = 'dd.MM.yyyy HH:mm'; $d.Width = 150; $d.Margin = '3,4,12,3'
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
    $g.AutoSizeColumnsMode = 'DisplayedCells'
    $g.BackgroundColor = [System.Drawing.SystemColors]::Window
    $g.Tag = $RowColor
    $g.Add_DataBindingComplete({
        param($s, $e)
        foreach ($c in $s.Columns) {
            if ($c.ValueType -eq [datetime]) { $c.DefaultCellStyle.Format = 'dd.MM.yyyy HH:mm:ss' }
        }
        if ($s.Columns.Count) { $s.Columns[$s.Columns.Count - 1].AutoSizeMode = 'Fill' }
    })
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

function Set-McGridData {
    <# Binds objects to a grid. $Columns: ordered map "Header" -> "Property". #>
    param($Grid, [object[]] $Items = @(), [System.Collections.Specialized.OrderedDictionary] $Columns)
    $items = @($Items | Where-Object { $_ -ne $null })
    $dt = New-Object System.Data.DataTable
    foreach ($h in $Columns.Keys) {
        $prop = $Columns[$h]
        $sample = $items | ForEach-Object { $_.$prop } | Where-Object { $_ -ne $null } | Select-Object -First 1
        $type = [string]
        if ($sample -is [int] -or $sample -is [long]) { $type = [long] }
        elseif ($sample -is [double] -or $sample -is [decimal]) { $type = [double] }
        elseif ($sample -is [datetime]) { $type = [datetime] }
        [void]$dt.Columns.Add($h, $type)
    }
    foreach ($it in $items) {
        $row = $dt.NewRow()
        foreach ($h in $Columns.Keys) {
            $v = $it.($Columns[$h])
            if ($null -eq $v) { $row[$h] = [DBNull]::Value }
            elseif ($dt.Columns[$h].DataType -eq [string]) {
                $row[$h] = if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { (@($v) -join '; ') } else { [string]$v }
            }
            else { $row[$h] = $v }
        }
        $dt.Rows.Add($row)
    }
    $Grid.DataSource = $dt
}

function Show-McConfirm {
    <# Confirmation dialog. With -RequireWord the OK button is enabled only after the user types the word. #>
    param([string] $Title, [string] $Text, [string] $OkText = 'Удалить', [string] $RequireWord)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title; $dlg.Size = '700,560'; $dlg.StartPosition = 'CenterParent'
    $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false; $dlg.Font = $form.Font; $dlg.ShowIcon = $false

    $body = New-Object System.Windows.Forms.TextBox
    $body.Multiline = $true; $body.ReadOnly = $true; $body.ScrollBars = 'Vertical'; $body.Dock = 'Fill'
    $body.Text = $Text -replace "`r?`n", "`r`n"
    $body.BackColor = [System.Drawing.SystemColors]::Window

    $bottom = New-Object System.Windows.Forms.FlowLayoutPanel
    $bottom.Dock = 'Bottom'; $bottom.AutoSize = $true; $bottom.Padding = '6,6,6,6'; $bottom.FlowDirection = 'LeftToRight'

    $ok = New-McButton $OkText 160
    $ok.DialogResult = 'OK'
    $cancel = New-McButton 'Отмена' 120
    $cancel.DialogResult = 'Cancel'
    $dlg.CancelButton = $cancel

    if ($RequireWord) {
        $bottom.Controls.Add((New-McLabel "Для подтверждения введите $($RequireWord):"))
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

#region ---------- Form layout ----------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'MailCleaner — поиск и удаление писем в Exchange'
$form.Size = '1280,860'
$form.MinimumSize = '1050,680'
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# --- Connection bar ---
$pnlConn = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlConn.Dock = 'Top'; $pnlConn.AutoSize = $true; $pnlConn.Padding = '6,6,6,0'; $pnlConn.WrapContents = $false
$tbServer = New-Object System.Windows.Forms.TextBox
$tbServer.Width = 260; $tbServer.Margin = '3,4,6,3'
$btnConnect = New-McButton 'Подключиться' 130
$lblCaps = New-McLabel 'Не подключено.'
$pnlConn.Controls.AddRange(@((New-McLabel 'Сервер Exchange (FQDN):'), $tbServer, $btnConnect, $lblCaps))

# --- Criteria ---
$grpCrit = New-Object System.Windows.Forms.GroupBox
$grpCrit.Text = 'Критерии поиска'; $grpCrit.Dock = 'Top'; $grpCrit.Height = 178; $grpCrit.Padding = '8,4,8,4'

$tlp = New-Object System.Windows.Forms.TableLayoutPanel
$tlp.Dock = 'Fill'; $tlp.ColumnCount = 6; $tlp.RowCount = 4
foreach ($i in 0..2) {
    [void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 33.3)))
}

$tbSenders    = New-McTextBox
$tbSubject    = New-McTextBox
$tbRecipient  = New-McTextBox
$dtStart      = New-McDatePicker
$dtEnd        = New-McDatePicker
$tbMessageId  = New-McTextBox
$tbAttachment = New-McTextBox
$tbKeywords   = New-McTextBox
$nudTimeout   = New-Object System.Windows.Forms.NumericUpDown
$nudTimeout.Minimum = 5; $nudTimeout.Maximum = 240; $nudTimeout.Value = 60; $nudTimeout.Width = 70; $nudTimeout.Margin = '3,4,12,3'

$tlp.Controls.Add((New-McLabel 'Отправители:'), 0, 0);    $tlp.Controls.Add($tbSenders, 1, 0)
$tlp.Controls.Add((New-McLabel 'Тема содержит:'), 2, 0);  $tlp.Controls.Add($tbSubject, 3, 0)
$tlp.Controls.Add((New-McLabel 'Получатель:'), 4, 0);     $tlp.Controls.Add($tbRecipient, 5, 0)
$tlp.Controls.Add((New-McLabel 'Начало периода:'), 0, 1); $tlp.Controls.Add($dtStart, 1, 1)
$tlp.Controls.Add((New-McLabel 'Конец периода:'), 2, 1);  $tlp.Controls.Add($dtEnd, 3, 1)
$tlp.Controls.Add((New-McLabel 'MessageId:'), 4, 1);      $tlp.Controls.Add($tbMessageId, 5, 1)
$tlp.Controls.Add((New-McLabel 'Имя вложения:'), 0, 2);   $tlp.Controls.Add($tbAttachment, 1, 2)
$tlp.Controls.Add((New-McLabel 'Слова в тексте:'), 2, 2); $tlp.Controls.Add($tbKeywords, 3, 2)
$tlp.Controls.Add((New-McLabel 'Таймаут, мин:'), 4, 2);   $tlp.Controls.Add($nudTimeout, 5, 2)

$pnlSearch = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlSearch.Dock = 'Fill'; $pnlSearch.AutoSize = $true; $pnlSearch.WrapContents = $false
$chkForwards = New-Object System.Windows.Forms.CheckBox
$chkForwards.Text = 'Включать пересылки (FW:) от других отправителей'; $chkForwards.AutoSize = $true
$chkForwards.Checked = $true; $chkForwards.Margin = '3,7,12,3'
$chkAll = New-Object System.Windows.Forms.CheckBox
$chkAll.Text = 'Искать во всех ящиках'; $chkAll.AutoSize = $true; $chkAll.Margin = '3,7,12,3'
$btnToday     = New-McButton 'Сегодня' 90
$btnYesterday = New-McButton 'Вчера' 90
$btnSearch    = New-McButton 'Найти' 130
$btnSearch.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
$pnlSearch.Controls.AddRange(@($chkForwards, $chkAll, $btnToday, $btnYesterday, $btnSearch))
$tlp.Controls.Add($pnlSearch, 0, 3)
$tlp.SetColumnSpan($pnlSearch, 6)
$grpCrit.Controls.Add($tlp)

# --- Summary line ---
$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Dock = 'Top'; $lblSummary.Height = 28; $lblSummary.Padding = '10,6,6,0'
$lblSummary.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
$lblSummary.Text = 'Результата нет.'

# --- Tabs with grids + log ---
$colorExternal = [System.Drawing.Color]::MistyRose
$colorNote     = [System.Drawing.Color]::LightYellow
$colorClean    = [System.Drawing.Color]::Honeydew

$gridMessages   = New-McGrid { param($r) if ([string]$r['Пометка']) { $colorNote } }
$gridTracking   = New-McGrid $null
$gridRecipients = New-McGrid { param($r) if ($r['Тип'] -eq 'Внешний') { $colorExternal } }
$gridFound      = New-McGrid $null
$gridMissing    = New-McGrid $null
$gridCheck      = New-McGrid {
    param($r)
    if ([long]$r['В видимых папках'] -lt 0) { $colorNote }
    elseif ([long]$r['В видимых папках'] -gt 0) { $colorExternal }
    else { $colorClean }
}

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$pages = [ordered]@{
    'Письма'           = $gridMessages
    'Журнал доставки'  = $gridTracking
    'Получатели'       = $gridRecipients
    'Найдено в ящиках' = $gridFound
    'Не найдено'       = $gridMissing
    'Проверка'         = $gridCheck
}
$tabPage = @{}
foreach ($k in $pages.Keys) {
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text = $k
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
$pnlActions = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlActions.Dock = 'Bottom'; $pnlActions.AutoSize = $true; $pnlActions.Padding = '6,6,6,6'; $pnlActions.WrapContents = $false
$btnSoft    = New-McButton 'Удалить (мягко)' 150
$btnHard    = New-McButton 'Удалить безвозвратно' 170
$btnHard.ForeColor = [System.Drawing.Color]::Firebrick
$btnVerify  = New-McButton 'Проверить' 110
$btnExport  = New-McButton 'Экспорт в CSV' 120
$btnCleanup = New-McButton 'Очистить старые поиски' 180
$btnCancel  = New-McButton 'Отмена' 100
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Width = 200; $progress.Height = 20; $progress.Margin = '12,7,6,3'
$lblStatus = New-McLabel ''
$pnlActions.Controls.AddRange(@($btnSoft, $btnHard, $btnVerify, $btnExport, $btnCleanup, $btnCancel, $progress, $lblStatus))

$tips = New-Object System.Windows.Forms.ToolTip
$tips.SetToolTip($tbSenders,   'Один или несколько адресов через запятую, точку с запятой или пробел.')
$tips.SetToolTip($tbRecipient, 'Только для журнала доставки: письма, которые получил этот адрес.')
$tips.SetToolTip($tbMessageId, 'Internet MessageId из журнала доставки или заголовков письма.')
$tips.SetToolTip($chkForwards, 'Добавить в поиск письма с той же темой от тех, кто переслал исходное письмо.')
$tips.SetToolTip($btnSoft,     'SoftDelete: письма переходят в «Элементы для восстановления». Пользователь может их восстановить.')
$tips.SetToolTip($btnHard,     'Search-Mailbox -DeleteContent: письма удаляются и из «Элементов для восстановления». Нужна роль Mailbox Import Export.')
$tips.SetToolTip($btnVerify,   'Проверить каждый ящик через Search-Mailbox: в видимых папках и всего.')

# Dock order: the control added last docks first (top-most).
$form.Controls.Add($split)
$form.Controls.Add($pnlActions)
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

    $tbServer.Enabled   = -not $busy
    $btnConnect.Enabled = -not $busy
    $grpCrit.Enabled    = $connected -and -not $busy
    $btnSearch.Enabled  = $connected -and -not $busy -and $caps.Search
    $btnSoft.Enabled    = (-not $busy) -and $clean -and $connected -and $caps.SoftDelete -and $caps.Verify
    $btnHard.Enabled    = (-not $busy) -and $clean -and $connected -and $caps.HardDelete
    $btnVerify.Enabled  = (-not $busy) -and $hasFound -and $connected -and $caps.Verify
    $btnExport.Enabled  = (-not $busy) -and [bool]$r
    $btnCleanup.Enabled = (-not $busy) -and $connected -and $caps.Search
    $btnCancel.Enabled  = $busy
    $form.UseWaitCursor = $false
}

function Update-McCapsLabel {
    $c = $script:Caps
    if (-not $c) { $lblCaps.Text = 'Не подключено.'; return }
    $mark = { param($v) if ($v) { '✓' } else { '✗' } }
    $lblCaps.Text = ('Подключено: {0}   Журнал {1}   Поиск {2}   Проверка {3}   Мягкое удаление {4}   Безвозвратное {5}' -f
        $script:ConnectedServer, (& $mark $c.Tracking), (& $mark $c.Search), (& $mark $c.Verify),
        (& $mark $c.SoftDelete), (& $mark $c.HardDelete))
    if (-not $c.HardDelete) { $tips.SetToolTip($lblCaps, 'Безвозвратное удаление недоступно: нужна роль Mailbox Import Export.') }
}

function Clear-McResults {
    $script:Result = $null
    $script:Check  = $null
    foreach ($g in $pages.Values) { $g.DataSource = $null }
    $lblSummary.Text = 'Результата нет.'
}

function Invoke-McInvalidate {
    if ($script:Loading -or -not $script:Result) { return }
    Clear-McResults
    Write-McLog 'Критерии изменены. Результат сброшен. Выполните поиск снова.' 'Warn'
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
    }
}

#endregion

#region ---------- Result views ----------

function Show-McSearchResult($r) {
    $script:Result = $r
    $script:Check  = $null
    Set-McGridData $gridMessages @($r.Messages) ([ordered]@{
        'Время' = 'Time'; 'Отправитель' = 'Sender'; 'Тема' = 'Subject'; 'Получателей' = 'Recipients'
        'Пометка' = 'Note'; 'MessageId' = 'MessageId' })
    Set-McGridData $gridTracking @($r.Events) ([ordered]@{
        'Время' = 'Timestamp'; 'Сервер' = 'ServerHostname'; 'Событие' = 'EventId'; 'Источник' = 'Source'
        'Отправитель' = 'Sender'; 'Получатели' = 'Recipients'; 'Тема' = 'MessageSubject'; 'MessageId' = 'MessageId' })
    $kindText = @{ Internal = 'Внутренний ящик'; External = 'Внешний'; Group = 'Группа'; Other = 'Другое' }
    $rcpt = @($r.Recipients | ForEach-Object {
        [pscustomobject]@{ Address = $_.Address; Kind = $kindText[$_.Kind]; Mailbox = $_.Mailbox; Type = $_.Type }
    } | Sort-Object Kind, Address)
    Set-McGridData $gridRecipients $rcpt ([ordered]@{
        'Адрес' = 'Address'; 'Тип' = 'Kind'; 'Ящик' = 'Mailbox'; 'Тип в Exchange' = 'Type' })
    Set-McGridData $gridFound @($r.Found) ([ordered]@{ 'Ящик' = 'Mailbox'; 'Писем' = 'Items'; 'Размер, КБ' = 'SizeKB' })
    $miss = @($r.Missing | ForEach-Object {
        [pscustomobject]@{ Mailbox = $_; Reason = 'Письмо удалено пользователем, перемещено или не проиндексировано.' }
    })
    Set-McGridData $gridMissing $miss ([ordered]@{ 'Ящик' = 'Mailbox'; 'Возможная причина' = 'Reason' })
    $gridCheck.DataSource = $null

    $lblSummary.Text = ('Найдено писем: {0} в {1} ящиках ({2:N1} КБ).   Не найдено в ящиках: {3}.   Внешних получателей: {4}.   Неиндексированных: {5}.' -f
        $r.Items, @($r.Found).Count, $r.SizeKB, @($r.Missing).Count, @($r.External).Count, $r.Unindexed)
    $tabs.SelectedTab = $tabPage['Найдено в ящиках']

    Write-McAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = 'Search'; SearchName = $r.SearchName; Query = $r.Query
        Mailboxes = @($r.Found).Count; Items = $r.Items
        Result = $(if ($r.Errors) { 'Errors' } else { 'OK' })
        Details = "External=$(@($r.External | ForEach-Object Address) -join ' '); Unindexed=$($r.Unindexed)"
    }
}

function Show-McCheckResult([object[]] $Check, [string] $Action, $DeleteInfo) {
    $script:Check = @($Check)
    Set-McGridData $gridCheck $script:Check ([ordered]@{
        'Ящик' = 'Mailbox'; 'В видимых папках' = 'Visible'; 'Всего (с «Элементами для восстановления»)' = 'Total'; 'Примечание' = 'Note' })
    $tabs.SelectedTab = $tabPage['Проверка']
    $visible = @($script:Check | Where-Object { $_.Visible -gt 0 })
    $hidden  = @($script:Check | Where-Object { $_.Visible -eq 0 -and $_.Total -gt 0 })
    $failed  = @($script:Check | Where-Object { $_.Visible -lt 0 })
    if ($visible.Count) { Write-McLog "Письмо видно пользователю в $($visible.Count) ящиках." 'Warn' }
    else { Write-McLog 'В видимых папках письма нет ни в одном ящике.' 'Ok' }
    if ($hidden.Count) {
        Write-McLog "Письмо есть только в «Элементах для восстановления» в $($hidden.Count) ящиках. Пользователь может его восстановить." 'Warn'
    }
    if ($failed.Count) { Write-McLog "Проверка не выполнена для $($failed.Count) ящиков. См. вкладку «Проверка»." 'Error' }

    $resultText = if ($failed.Count) { 'CheckFailed' } elseif ($visible.Count) { 'Visible' } elseif ($hidden.Count) { 'RecoverableOnly' } else { 'Clean' }
    Write-McAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = $Action; SearchName = $script:Result.SearchName
        Query = $script:Result.Query; Mailboxes = $script:Check.Count; Items = $script:Result.Items
        Result = $resultText; Details = [string]$DeleteInfo
    }
}

#endregion

#region ---------- Event handlers ----------

$btnConnect.Add_Click({
    $srv = $tbServer.Text.Trim()
    if (-not $srv) { [void][System.Windows.Forms.MessageBox]::Show($form, 'Введите FQDN сервера Exchange.', 'MailCleaner'); return }
    $script:Caps = $null
    Clear-McResults
    Update-McCapsLabel
    Start-McOperation -Name 'Connect' -Params @{ Server = $srv } -StatusText "Подключение к $srv..." -OnSuccess {
        param($caps)
        $script:Caps = $caps
        $script:ConnectedServer = $tbServer.Text.Trim()
        Update-McCapsLabel
        Save-McSettings
        if (-not $caps.HardDelete) { Write-McLog 'Безвозвратное удаление недоступно: нет роли Mailbox Import Export.' 'Warn' }
        if (-not $caps.Tracking)   { Write-McLog 'Журнал доставки недоступен: нет роли Message Tracking.' 'Warn' }
    }
})

$btnToday.Add_Click({
    $dtStart.Value = (Get-Date).Date
    $dtEnd.Value   = (Get-Date).Date.AddDays(1)
})

$btnYesterday.Add_Click({
    $dtStart.Value = (Get-Date).Date.AddDays(-1)
    $dtEnd.Value   = (Get-Date).Date
})

$btnSearch.Add_Click({
    $crit = Get-McCriteria
    $problems = @(Test-McCriteria $crit)
    if ($problems.Count) {
        [void][System.Windows.Forms.MessageBox]::Show($form, ($problems -join "`r`n"), 'Проверьте критерии',
            'OK', [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Save-McSettings
    Clear-McResults
    Write-McLog ('Поиск: отправители [{0}], тема [{1}], период {2:dd.MM.yyyy HH:mm} — {3:dd.MM.yyyy HH:mm}.' -f
        ($crit.Senders -join ', '), $crit.Subject, $crit.Start, $crit.End)
    Start-McOperation -Name 'Search' -Params @{ Criteria = $crit; TimeoutMinutes = [int]$nudTimeout.Value } `
        -StatusText 'Поиск...' -OnSuccess { param($r) Show-McSearchResult $r }
})

$btnSoft.Add_Click({
    $r = $script:Result
    $mbx = @($r.Found | ForEach-Object Mailbox)
    $text = @"
МЯГКОЕ УДАЛЕНИЕ (SoftDelete)

Писем: $($r.Items)
Ящиков: $($mbx.Count)
Поиск: $($r.SearchName)
KQL: $($r.Query)

Письма перейдут в «Элементы для восстановления».
Пользователь не увидит их в папках.
Пользователь может восстановить их в Outlook.

Внешние получатели (удалить нельзя): $(@($r.External | ForEach-Object Address) -join ', ')
Неиндексированные элементы не удаляются: $($r.Unindexed)

Ящики:
$($mbx -join "`r`n")
"@
    if (-not (Show-McConfirm -Title 'Мягкое удаление' -Text $text -OkText 'Удалить (мягко)')) { return }
    Start-McOperation -Name 'SoftDelete' -StatusText 'Мягкое удаление...' -Params @{
        SearchName = $r.SearchName; Mailboxes = $mbx; Query = $r.Query; TimeoutMinutes = [int]$nudTimeout.Value
    } -OnSuccess {
        param($x)
        Show-McCheckResult -Check @($x.Check) -Action 'SoftDelete' -DeleteInfo "Status=$($x.Delete.Status)"
    }
})

$btnHard.Add_Click({
    $r = $script:Result
    $mbx = @($r.Found | ForEach-Object Mailbox)
    $text = @"
БЕЗВОЗВРАТНОЕ УДАЛЕНИЕ (Search-Mailbox -DeleteContent)

Писем: $($r.Items)
Ящиков: $($mbx.Count)
KQL: $($r.Query)

Письма будут удалены из всех папок.
Письма будут удалены из «Элементов для восстановления».
Пользователь НЕ может восстановить эти письма.

Ящик на удержании (Hold) сохраняет копию в скрытой папке.
Пользователь не видит эту копию.

Внешние получатели (удалить нельзя): $(@($r.External | ForEach-Object Address) -join ', ')
Неиндексированные элементы не удаляются: $($r.Unindexed)

Ящики:
$($mbx -join "`r`n")
"@
    if (-not (Show-McConfirm -Title 'Безвозвратное удаление' -Text $text -OkText 'Удалить безвозвратно' -RequireWord 'УДАЛИТЬ')) { return }
    Start-McOperation -Name 'HardDelete' -StatusText 'Безвозвратное удаление...' -Params @{
        Mailboxes = $mbx; Query = $r.Query
    } -OnSuccess {
        param($x)
        $deleted = (@($x.Delete) | Measure-Object Deleted -Sum).Sum
        $errors  = @($x.Delete | Where-Object Error).Count
        Show-McCheckResult -Check @($x.Check) -Action 'HardDelete' -DeleteInfo "Deleted=$deleted; MailboxErrors=$errors"
    }
})

$btnVerify.Add_Click({
    $r = $script:Result
    Start-McOperation -Name 'Verify' -StatusText 'Проверка...' -Params @{
        Mailboxes = @($r.Found | ForEach-Object Mailbox); Query = $r.Query
    } -OnSuccess { param($x) Show-McCheckResult -Check @($x.Check) -Action 'Verify' -DeleteInfo '' }
})

$btnExport.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Папка для CSV-файлов'
    $dlg.SelectedPath = $AppDir
    if ($dlg.ShowDialog($form) -ne 'OK') { return }
    $r = $script:Result
    $base = Join-Path $dlg.SelectedPath $r.SearchName
    $sets = [ordered]@{
        'messages'   = @($r.Messages)
        'tracking'   = @($r.Events | Select-Object Timestamp, ServerHostname, EventId, Source, Sender,
                            @{ n = 'Recipients'; e = { $_.Recipients -join '; ' } }, MessageSubject, MessageId)
        'recipients' = @($r.Recipients)
        'found'      = @($r.Found)
        'missing'    = @($r.Missing | ForEach-Object { [pscustomobject]@{ Mailbox = $_ } })
        'check'      = @($script:Check)
    }
    foreach ($k in $sets.Keys) {
        if (@($sets[$k]).Count) {
            $sets[$k] | Export-Csv -Path "$base-$k.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
        }
    }
    Write-McLog "Экспорт выполнен: $base-*.csv" 'Ok'
})

$btnCleanup.Add_Click({
    $days = [Microsoft.VisualBasic.Interaction]::InputBox(
        'Удалить поиски MailCleaner (MC_*) старше, дней:', 'Очистка старых поисков', '7')
    if (-not $days) { return }
    if ($days -notmatch '^\d+$') { Write-McLog 'Введите целое число дней.' 'Warn'; return }
    Start-McOperation -Name 'ListOld' -Params @{ Days = [int]$days } -StatusText 'Список старых поисков...' -OnSuccess {
        param($list)
        $list = @($list | Where-Object { $_ })
        if (-not $list.Count) { Write-McLog 'Старых поисков MailCleaner нет.' 'Ok'; return }
        $names = @($list | ForEach-Object Name)
        $text = "Будут удалены поиски и их действия ($($names.Count)):`r`n`r`n" + ($names -join "`r`n")
        if (-not (Show-McConfirm -Title 'Очистка старых поисков' -Text $text -OkText 'Удалить поиски')) { return }
        $script:PendingCleanup = $names
    }
})

$btnCancel.Add_Click({
    if ($script:Job) {
        $Sync.Cancel = $true
        Write-McLog 'Запрошена отмена. Дождитесь завершения текущего шага.' 'Warn'
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
        Start-McOperation -Name 'RemoveOld' -Params @{ Names = $script:CleanupNames } -StatusText 'Удаление старых поисков...' -OnSuccess {
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
$chkAll.Add_CheckedChanged({ Invoke-McInvalidate })

$form.Add_Shown({
    $split.SplitterDistance = [int]($split.Height * 0.66)
    Write-McLog "MailCleaner запущен. Лог: $LogPath"
    Write-McLog "Журнал действий: $AuditPath"
    if ($tbServer.Text) { $btnConnect.PerformClick() }
})

$form.Add_FormClosing({
    param($s, $e)
    if ($script:Job) {
        $a = [System.Windows.Forms.MessageBox]::Show($form, 'Операция ещё выполняется. Закрыть окно и прервать её?',
            'MailCleaner', 'YesNo', [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($a -ne 'Yes') { $e.Cancel = $true; return }
        $Sync.Cancel = $true
        try { $script:Job.PS.Stop() } catch { }
    }
    $timer.Stop()
    Save-McSettings
})

#endregion

#region ---------- Start ----------

$script:Loading = $true
$settings = Read-McSettings
$defaultServer = if ($Server) { $Server }
                 elseif ($settings.ContainsKey('Server') -and $settings.Server) { [string]$settings.Server }
                 elseif ($env:ExchangeInstallPath) { [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName }
                 else { '' }
$tbServer.Text = $defaultServer
if ($settings.ContainsKey('Senders'))         { $tbSenders.Text = [string]$settings.Senders }
if ($settings.ContainsKey('Subject'))         { $tbSubject.Text = [string]$settings.Subject }
if ($settings.ContainsKey('IncludeForwards')) { $chkForwards.Checked = [bool]$settings.IncludeForwards }
if ($settings.ContainsKey('AllMailboxes'))    { $chkAll.Checked = [bool]$settings.AllMailboxes }
if ($settings.ContainsKey('TimeoutMinutes'))  { $nudTimeout.Value = [math]::Max(5, [math]::Min(240, [int]$settings.TimeoutMinutes)) }
$dtStart.Value = (Get-Date).Date
$dtEnd.Value   = (Get-Date).Date.AddDays(1)
$script:Loading = $false

Set-UiState
$timer.Start()
try {
    [void]$form.ShowDialog()
} finally {
    $timer.Stop()
    try {
        $ps = [powershell]::Create()
        $ps.Runspace = $Runspace
        [void]$ps.AddScript('if ($global:McExSession) { Remove-PSSession $global:McExSession -ErrorAction SilentlyContinue }')
        [void]$ps.Invoke()
        $ps.Dispose()
    } catch { }
    $Runspace.Close()
    $Runspace.Dispose()
    $form.Dispose()
    Remove-Variable McSync -Scope Global -ErrorAction SilentlyContinue
}

#endregion
