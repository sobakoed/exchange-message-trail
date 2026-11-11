#Requires -Version 5.1
<#
.SYNOPSIS
    Makes the screenshots of the user guide (help\img\<language>\*.png) again.

.DESCRIPTION
    The script starts the real window of EMT outside the screen, with the dark theme.
    It shows example data (contoso), not real data. The core functions make the delivery statuses.
    No modal window opens. No connection to Exchange occurs.
    The settings, the log and the audit file go to a temporary folder. The script removes this folder at the end.
    Run the script after each change of the window. Then examine the screenshots.

.PARAMETER Language
    The languages of the screenshots. Default: en and ru.

.EXAMPLE
    powershell.exe -STA -ExecutionPolicy Bypass -File .\tools\Make-HelpScreenshots.ps1
#>
[CmdletBinding()]
param([string[]] $Language = @('en', 'ru'), [switch] $Child)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot

if (-not $Child) {
    # Each language in its own process: the window, the culture and the texts start clean.
    foreach ($l in $Language) {
        Write-Host "Screenshots: $l"
        & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $PSCommandPath -Language $l -Child
        if ($LASTEXITCODE) { throw "The screenshots of the language '$l' failed." }
    }
    return
}

# ---------- Child process: one language ----------
$lang = $Language[0]
$culture = if ($lang -eq 'ru') { 'ru-RU' } else { 'en-US' }
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($culture)
$out = Join-Path $repo "help\img\$lang"
if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }
$tmp = Join-Path $env:TEMP ('EMT-help-' + [guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $tmp | Out-Null
Set-Content -Path (Join-Path $tmp 'settings.json') -Encoding UTF8 -Value (@{ Theme = 'Dark'; Language = $lang; Server = '' } | ConvertTo-Json)

$text = [IO.File]::ReadAllText((Join-Path $repo 'ExchangeMessageTrail.ps1'), [Text.Encoding]::UTF8)
$text = $text.Replace('$PSScriptRoot', "'$repo'")
$text = $text.Replace("Join-Path `$env:USERPROFILE 'ExchangeMessageTrail'", "'$tmp'")
$text = $text.Replace("Join-Path `$env:APPDATA 'ExchangeMessageTrail\settings.json'", "'$tmp\settings.json'")
$text = $text.Replace("Join-Path `$env:APPDATA 'MailCleaner\settings.json'", "'$tmp\none.json'")
$text = $text.Replace("Join-Path `$env:USERPROFILE 'MailCleaner\audit.csv'", "'$tmp\none.csv'")
$text = $text -replace '(?s)trap \{.*?break\r?\n\}', 'trap { Write-Host "ERROR: $$($$_.Exception.Message) $$($$_.InvocationInfo.PositionMessage)"; exit 1 }'

$scenes = @'
$window.WindowState = 'Normal'      # a small screen opens the window maximized: the screenshots need the same size always
$window.WindowStartupLocation = 'Manual'; $window.Left = -20000; $window.Top = 0
$window.Width = 1280; $window.Height = 1000; $window.ShowActivated = $false
$window.Show()
function Pump { for ($i = 0; $i -lt 3; $i++) { $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::ContextIdle); Start-Sleep -Milliseconds 60 } }
function Click($b) { $b.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))); Pump }
function Snap($Element, [string] $Name, [int] $Pad = 12) {
    Pump
    $w = [int]$Element.ActualWidth + 2 * $Pad; $h = [int]$Element.ActualHeight + 2 * $Pad
    $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($w, $h, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32))
    $dv = New-Object System.Windows.Media.DrawingVisual; $dc = $dv.RenderOpen()
    $dc.DrawRectangle($window.Background, $null, (New-Object System.Windows.Rect(0, 0, $w, $h)))
    $vb = New-Object System.Windows.Media.VisualBrush($Element); $vb.Stretch = 'None'; $vb.AlignmentX = 'Left'; $vb.AlignmentY = 'Top'
    $dc.DrawRectangle($vb, $null, (New-Object System.Windows.Rect($Pad, $Pad, $Element.ActualWidth, $Element.ActualHeight)))
    $dc.Close(); $bmp.Render($dv)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder; $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bmp))
    $fs = [IO.File]::Create((Join-Path '__OUT__' "$Name.png")); $enc.Save($fs); $fs.Close()
    Write-Host "  $Name.png ${w}x${h}"
}
# A screenshot of a full window: the title bar, the icon and the frame too. Windows draws them, so a WPF render cannot.
# Two parts: PrintWindow gives the title bar and the frame (it operates also outside the screen).
# A WPF render of the window gives the client area (PrintWindow cannot see it outside the screen).
# WM_NCACTIVATE draws the title bar as active, but the window does not get the focus.
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition '
using System; using System.Drawing; using System.IO; using System.Runtime.InteropServices;
public static class EmtShot {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h, ref POINT p);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
    [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr h, int attr, out RECT r, int size);
    public static int[] ClientSize(IntPtr h) { RECT c; GetClientRect(h, out c); return new int[] { c.R - c.L, c.B - c.T }; }
    public static Bitmap Capture(IntPtr h, byte[] clientPng) {
        SendMessage(h, 0x86, (IntPtr)1, IntPtr.Zero);
        RECT w; GetWindowRect(h, out w);
        RECT f; if (DwmGetWindowAttribute(h, 9, out f, Marshal.SizeOf(typeof(RECT))) != 0) { f = w; }
        Bitmap full = new Bitmap(w.R - w.L, w.B - w.T, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
        using (Graphics g = Graphics.FromImage(full)) { IntPtr dc = g.GetHdc(); PrintWindow(h, dc, 2); g.ReleaseHdc(dc); }
        Bitmap res = full.Clone(new Rectangle(f.L - w.L, f.T - w.T, f.R - f.L, f.B - f.T), full.PixelFormat);
        full.Dispose();
        POINT p = new POINT(); ClientToScreen(h, ref p);
        using (MemoryStream ms = new MemoryStream(clientPng)) using (Bitmap client = new Bitmap(ms))
        using (Graphics g = Graphics.FromImage(res)) { g.DrawImageUnscaled(client, p.X - f.L, p.Y - f.T); }
        return res;
    }
}'
function SnapWindow($Win, [string] $Name) {
    Pump
    $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($Win)).Handle
    $size = [EmtShot]::ClientSize($hwnd)
    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($size[0], $size[1], 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32))
    $rtb.Render($Win)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder; $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $ms = New-Object IO.MemoryStream; $enc.Save($ms)
    $bmp = [EmtShot]::Capture($hwnd, $ms.ToArray())
    $bmp.Save((Join-Path '__OUT__' "$Name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Host "  $Name.png $($bmp.Width)x$($bmp.Height)"
    $bmp.Dispose()
}
# The example day is 2026-11-10. "Today" is always 2026-11-11 for the mark of the quick period:
# the screenshots show the same mark ("Yesterday") on each day.
$script:RealQuickPeriod = ${function:Get-EmtQuickPeriod}
function Get-EmtQuickPeriod { param($Start, $End, $Now) & $script:RealQuickPeriod -Start $Start -End $End -Now ([datetime]'2026-11-11 09:31') }
foreach ($set in $script:PeriodSets) { $set.Mark = $null }
Update-EmtPeriodMarks
function Show-EmtMessage { param($Text, $Title, $Kind) $true }
function Show-EmtConfirm { param($Title, $Text, $OkText, $RequireWord) $true }
function Export-EmtSets { param($Stem, $Sets) }
function Set-Log([string[]] $Lines) {
    $script:LogPara.Inlines.Clear()
    foreach ($l in $Lines) {
        $run = New-Object System.Windows.Documents.Run($l)
        $key = switch -Regex ($l) { '\[ERROR\]' { 'Text.Red'; break } '\[WARN' { 'Text.Orange'; break } '\[OK' { 'Text.Green'; break } default { 'Fg' } }
        Set-EmtBrush $run $key ([System.Windows.Documents.TextElement]::ForegroundProperty)
        $script:LogPara.Inlines.Add($run); $script:LogPara.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
    }
    # The same page width as Add-EmtLogLine gives: the longest line.
    $script:LogTextWidth = [double](@($Lines | ForEach-Object { Measure-EmtText $_ $log }) | Measure-Object -Maximum).Maximum
    Update-EmtLogWidth
}

# ---------- Example data ----------
$ru = $Sync.Lang -eq 'ru'
$day = [datetime]'2026-11-10'
$subj = if ($ru) { @('Квартальный отчёт', 'План проекта', 'Заметки со встречи', 'Счёт 2026-114', 'Еженедельная сводка', 'Обновление сервера') }
        else      { @('Quarterly report', 'Project plan', 'Meeting notes', 'Invoice 2026-114', 'Weekly update', 'Server update') }
function Ev($id, $time, $event, [string[]] $to, [string[]] $st, $server = 'EXCH01', $sender = 'jsmith@contoso.com', $source = 'SMTP') {
    [pscustomobject]@{ MessageId = $id; Timestamp = $time; EventId = $event; Recipients = $to; RecipientStatus = $st
                       ServerHostname = $server; Source = $source; Sender = $sender; MessageSubject = $script:SubjectOf[$id] }
}
$script:SubjectOf = @{}
$ids = @('<a1b2c3@contoso.com>', '<d4e5f6@contoso.com>', '<g7h8i9@contoso.com>', '<j1k2l3@contoso.com>', '<m4n5o6@contoso.com>', '<p7q8r9@contoso.com>')
for ($i = 0; $i -lt $ids.Count; $i++) { $script:SubjectOf[$ids[$i]] = $subj[$i] }
$led = { param($x) "[{LED=250 2.1.5 $x};{MSG=};{FQDN=};{IP=};{LRT=}]" }
$t1 = $day.AddHours(9).AddMinutes(12); $t2 = $day.AddHours(10).AddMinutes(41); $t3 = $day.AddHours(11).AddMinutes(5)
$t4 = $day.AddHours(13).AddMinutes(20); $t5 = $day.AddHours(15).AddMinutes(2); $t6 = $day.AddHours(16).AddMinutes(48)
$events = @(
    Ev $ids[0] $t1 'RECEIVE' @('akim@contoso.com', 'mlee@contoso.com', 'tnovak@contoso.com', 'nobody@contoso.com', 'partner@fabrikam.com') @('', '', '', '', '')
    Ev $ids[0] $t1.AddSeconds(2) 'DELIVER' @('akim@contoso.com', 'mlee@contoso.com', 'tnovak@contoso.com') @('', '', '') 'EXCH01' 'jsmith@contoso.com' 'STOREDRIVER'
    Ev $ids[0] $t1.AddSeconds(2) 'FAIL' @('nobody@contoso.com') @('550 5.1.1 RESOLVER.ADR.RecipNotFound; not found')
    Ev $ids[0] $t1.AddSeconds(4) 'SEND' @('partner@fabrikam.com') @('250 2.0.0 OK')
    Ev $ids[1] $t2 'RECEIVE' @('akim@contoso.com', 'mlee@contoso.com', 'team@contoso.com') @('', '', '')
    Ev $ids[1] $t2.AddSeconds(1) 'DROP' @('team@contoso.com') @((& $led 'RESOLVER.GRP.Expanded; distribution list expanded'))
    Ev $ids[1] $t2.AddSeconds(2) 'DELIVER' @('akim@contoso.com') @('Projects') 'EXCH02' 'jsmith@contoso.com' 'STOREDRIVER'
    Ev $ids[1] $t2.AddSeconds(2) 'DELIVER' @('mlee@contoso.com') @('') 'EXCH01' 'jsmith@contoso.com' 'STOREDRIVER'
    Ev $ids[2] $t3 'RECEIVE' @('akim@contoso.com', 'partner@northwind.com') @('', '')
    Ev $ids[2] $t3.AddSeconds(2) 'DELIVER' @('akim@contoso.com') @('') 'EXCH01' 'jsmith@contoso.com' 'STOREDRIVER'
    Ev $ids[2] $t3.AddSeconds(30) 'DEFER' @('partner@northwind.com') @('451 4.4.0 Primary target IP address responded with: 421 Service not available')
    Ev $ids[3] $t4 'RECEIVE' @('billing@fabrikam.com', 'monitor@security.contoso.net') @('', '')
    Ev $ids[3] $t4.AddSeconds(3) 'SEND' @('billing@fabrikam.com') @('250 2.0.0 OK')
    Ev $ids[3] $t4.AddSeconds(3) 'SENDEXTERNAL' @('monitor@security.contoso.net') @('250 2.1.5 Ok')
    Ev $ids[4] $t5 'RECEIVE' @('akim@contoso.com', 'mlee@contoso.com', 'tnovak@contoso.com', 'opetrov@contoso.com') @('', '', '', '')
    Ev $ids[4] $t5.AddSeconds(2) 'DELIVER' @('akim@contoso.com', 'mlee@contoso.com', 'tnovak@contoso.com') @('', '', '') 'EXCH01' 'jsmith@contoso.com' 'STOREDRIVER'
    Ev $ids[4] $t5.AddSeconds(2) 'DROP' @('opetrov@contoso.com') @((& $led 'RESOLVER.FWD.Forwarded; recipient forwarded'))
    Ev $ids[5] $t6 'RECEIVE' @('akim@contoso.com', 'mlee@contoso.com') @('', '')
    Ev $ids[5] $t6.AddSeconds(2) 'DELIVER' @('akim@contoso.com', 'mlee@contoso.com') @('', '') 'EXCH02' 'jsmith@contoso.com' 'STOREDRIVER'
)
$kinds = @{}
foreach ($a in 'akim', 'mlee', 'tnovak', 'opetrov', 'nobody') { $kinds["$a@contoso.com"] = 'Internal' }
$kinds['team@contoso.com'] = 'Group'
foreach ($a in 'partner@fabrikam.com', 'partner@northwind.com', 'billing@fabrikam.com', 'monitor@security.contoso.net') { $kinds[$a] = 'External' }
$report = @(Get-EmtDeliveryReport -Events $events -Kinds $kinds)
$messages = @(foreach ($id in $ids) {
    $rows = @($report | Where-Object MessageId -eq $id)
    $first = @($events | Where-Object MessageId -eq $id)[0]
    [pscustomobject]@{ Time = $first.Timestamp; Sender = 'jsmith@contoso.com'; Subject = $script:SubjectOf[$id]; Recipients = $rows.Count
                       MessageId = $id; Problems = 0; Pending = 0; StatusLine = ''; ServerGenerated = $false }
})
$crit = @{ Senders = @('jsmith@contoso.com'); Recipient = ''; Subject = ''; MessageId = ''; Start = $day; End = $day.AddDays(1).AddTicks(-1) }
$trace = [pscustomobject]@{ Criteria = $crit; Messages = $messages; Report = $report; Events = $events; Recipients = @()
                           TrackingErrors = @(); NoResultHints = @(); Seconds = 2; Partial = $false }

$script:NextResult = $null
$script:DetailResult = $null
function Start-EmtOperation {
    param($Name, $Params, $OnSuccess, $StatusText)
    $r = switch ($Name) {
        'Connect'     { [pscustomobject]@{ Tracking = $true; MailboxCheck = $true; Search = $true; Verify = $true; SoftDelete = $true; HardDelete = $true } }
        'TraceDetail' { $script:DetailResult }
        default       { $script:NextResult }
    }
    & $OnSuccess $r
    Set-UiState
}
function Select-Message([int] $Index) {
    # The grid sets the column widths after the rows are on the screen: wait for it.
    $gridTrMessages.SelectedIndex = $Index; Pump; Start-Sleep -Milliseconds 300; Pump
}
function Select-Recipient([string] $Address) {
    foreach ($v in $gridTrRecipients.Items) { if ([string]$v.Row['Recipient'] -eq $Address) { $gridTrRecipients.SelectedItem = $v; break } }
    Pump
}

Pump
$tbServer.Text = 'exch01.contoso.com'; Click $btnConnect
foreach ($p in @(@($dtTrStart, $tbTrStartTime, $day), @($dtTrEnd, $tbTrEndTime, $day.AddDays(1).AddMinutes(-1)))) { Set-EmtPeriodValue $p[0] $p[1] $p[2] }
$tbTrSenders.Text = 'jsmith@contoso.com'
$script:NextResult = $trace
Click $btnTrFind
$logSearch = if ($ru) {
    @('09:30:02 [INFO ] Exchange Message Trail запущен.', '09:30:04 [OK   ] Подключено к exch01.contoso.com.',
      '09:31:10 [INFO ] Поиск писем: отправители [jsmith@contoso.com], получатель [], тема [], период 10.11.2026 0:00 — 10.11.2026 23:59.',
      '09:31:12 [OK   ] Поиск писем завершён: найдено писем 6, получателей 18, ошибок доставки 1 (2 с).')
} else {
    @('09:30:02 [INFO ] Exchange Message Trail started.', '09:30:04 [OK   ] Connected to exch01.contoso.com.',
      '09:31:10 [INFO ] Message search: senders [jsmith@contoso.com], recipient [], subject [], period 11/10/2026 12:00 AM to 11/10/2026 11:59 PM.',
      '09:31:12 [OK   ] Message search complete: 6 messages, 18 recipients, 1 delivery errors (2 s).')
}
Set-Log $logSearch

# 1. The window and its parts
Select-Message 0
SnapWindow $window 'window'
Snap $tbServer.Parent.Parent 'top-line' 8
Snap $lblConn.Parent.Parent 'bottom-line' 8
Snap $cardTr 'search-criteria' 10
Snap $gridTrMessages.Parent 'message-list' 6

# 2. The details of a message
Snap $trDetail 'recipients-failed' 6
Select-Message 1
Snap $trDetail 'recipients-folder' 6
Select-Message 0
$trDetail.SelectedItem = $tabTrEvents; Pump
Snap $trDetail 'events' 6
$trDetail.SelectedItem = $tabTrRecipients; Pump

# 3. The mailbox check (an inbox rule moves the message to a folder)
Select-Message 1
Select-Recipient 'akim@contoso.com'
$script:NextResult = [pscustomobject]@{ Mailbox = 'akim@contoso.com'; Sender = 'jsmith@contoso.com'; Subject = $subj[1]; Data = $null
    Findings = @(
        [pscustomobject]@{ Severity = 'Alert'; Area = 'Rules'; Text = $(if ($ru) { 'Правило «Проекты»: письма от jsmith@contoso.com перемещаются в папку «Projects». Правило применяется к этому письму.' } else { 'Rule "Projects": messages from jsmith@contoso.com move to the folder "Projects". The rule applies to this message.' }) }
        [pscustomobject]@{ Severity = 'Info'; Area = 'Junk'; Text = $(if ($ru) { 'Фильтр нежелательной почты включён. Отправитель не в списке заблокированных.' } else { 'The junk e-mail filter is on. The sender is not in the blocked list.' }) }
        [pscustomobject]@{ Severity = 'Info'; Area = 'Quota'; Text = $(if ($ru) { 'Ящик заполнен на 42 %.' } else { 'The mailbox is 42% full.' }) }
    ) }
Click $btnTrCheck
Snap $trDetail 'mailbox-check' 6
$trDetail.SelectedItem = $tabTrRecipients; Pump

# 4. Other statuses: forwarding of a mailbox, a copy for the security team, a delay
Select-Message 4
Snap $trDetail 'recipients-forwarded' 6
Select-Message 3
Snap $trDetail 'recipients-sendexternal' 6
Select-Message 2
Snap $trDetail 'recipients-deferred' 6

# 5. The calendar
$dtTrStart.IsDropDownOpen = $true; Pump
Snap $dtTrStart.Template.FindName('PART_Popup', $dtTrStart).Child 'calendar' 8
$dtTrStart.IsDropDownOpen = $false; Pump

# 6. A search by a recipient: the other recipients load when the user selects a message
$part = @($report | Where-Object { $_.Recipient -eq 'akim@contoso.com' })
$partIds = @($part | ForEach-Object MessageId)
$pmsgs = @($messages | Where-Object { $_.MessageId -in $partIds })
$script:NextResult = [pscustomobject]@{ Criteria = @{ Senders = @(); Recipient = 'akim@contoso.com'; Subject = ''; MessageId = ''; Start = $day; End = $day.AddDays(1).AddTicks(-1) }
    Messages = $pmsgs; Report = $part; Events = @($events | Where-Object { @($_.Recipients) -contains 'akim@contoso.com' }); Recipients = @()
    TrackingErrors = @(); NoResultHints = @(); Seconds = 1; Partial = $true }
# The first message loads its details at once: the detail result is the full data of this message.
$firstId = $pmsgs[0].MessageId
$script:DetailResult = [pscustomobject]@{ MessageIds = @($firstId); Events = @($events | Where-Object MessageId -eq $firstId)
    Report = @($report | Where-Object MessageId -eq $firstId); Recipients = @(); Errors = @() }
$tbTrSenders.Text = ''; $tbTrRecipient.Text = 'akim@contoso.com'
Click $btnTrFind
Set-Log $(if ($ru) { @('09:40:01 [INFO ] Поиск по получателю akim@contoso.com: в списке показана доставка этому получателю. Остальные получатели письма загружаются при выборе письма.') }
          else { @('09:40:01 [INFO ] Search by the recipient akim@contoso.com: the list shows the delivery to this recipient. The other recipients of a message load when you select the message.') })
[void]$gridTrMessages.Focus()     # no text selection in a date field on the screenshot
SnapWindow $window 'recipient-search'

# 7. The tab "Delete" after "Delete this message..."
$tbTrRecipient.Text = ''; $tbTrSenders.Text = 'jsmith@contoso.com'
$script:NextResult = $trace
Click $btnTrFind
Select-Message 0
Click $btnTrDelete
$q = New-EmtSearchQuery -Senders @('jsmith@contoso.com') -Subject $subj[0] -Start $t1.AddMinutes(-5) -End $day.AddDays(1).AddTicks(-1)
$found = @(
    [pscustomobject]@{ Mailbox = 'jsmith@contoso.com'; Items = 2; SizeKB = 168 }
    [pscustomobject]@{ Mailbox = 'akim@contoso.com'; Items = 1; SizeKB = 84 }
    [pscustomobject]@{ Mailbox = 'mlee@contoso.com'; Items = 1; SizeKB = 84 }
    [pscustomobject]@{ Mailbox = 'tnovak@contoso.com'; Items = 1; SizeKB = 84 }
)
$script:NextResult = [pscustomobject]@{
    Messages = @([pscustomobject]@{ Time = $t1; Sender = 'jsmith@contoso.com'; Subject = $subj[0]; Recipients = 5; MessageId = $ids[0]; ServerGenerated = $false })
    Events = @($events | Where-Object MessageId -eq $ids[0])
    Recipients = @(
        [pscustomobject]@{ Address = 'akim@contoso.com'; Kind = 'Internal'; Mailbox = 'akim@contoso.com'; Type = 'UserMailbox' }
        [pscustomobject]@{ Address = 'mlee@contoso.com'; Kind = 'Internal'; Mailbox = 'mlee@contoso.com'; Type = 'UserMailbox' }
        [pscustomobject]@{ Address = 'tnovak@contoso.com'; Kind = 'Internal'; Mailbox = 'tnovak@contoso.com'; Type = 'UserMailbox' }
        [pscustomobject]@{ Address = 'partner@fabrikam.com'; Kind = 'External'; Mailbox = ''; Type = '' })
    Found = $found; Missing = @(); External = @([pscustomobject]@{ Address = 'partner@fabrikam.com' })
    Items = 5; SizeKB = 420; Unindexed = 0; Method = 'SearchMailbox'; Locations = @('a', 'b', 'c', 'd'); Seconds = 3
    Errors = ''; SearchName = ''; Query = $q; Unconfirmed = $false; TrackingErrors = @() }
Click $btnSearch
Set-Log $(if ($ru) { @("09:45:20 [INFO ] Вкладка «Удаление» заполнена данными письма $($ids[0]). Нажмите «Найти копии».", '09:45:24 [WARN ] Внешних получателей: 1. У них письмо удалить нельзя.', '09:45:27 [OK   ] Найдено писем: 5 в 4 ящиках, 420,0 КБ.') }
          else { @("09:45:20 [INFO ] The tab ""Delete"" has the data of the message $($ids[0]). Click ""Find copies"".", '09:45:24 [WARN ] External recipients: 1. The message cannot be deleted for them.', '09:45:27 [OK   ] Found 5 messages in 4 mailboxes, 420.0 KB.') })
SnapWindow $window 'delete'
Snap $cardDel 'delete-criteria' 10

# 8. The confirmation of a hard delete (made, not shown as a modal window)
$r = $script:Result
$d = New-EmtConfirmWindow -Title (T 'confirm.hardTitle') -Text (Get-EmtDeleteText -r $r -Kind 'Hard' -Warning (Get-EmtDeleteWarning $r)) -OkText (T 'act.hard') -RequireWord (T 'confirm.word')
$d.Owner = $null; $d.WindowStartupLocation = 'Manual'; $d.Left = -20000; $d.Top = 0; $d.ShowActivated = $false
$d.Show(); Pump
SnapWindow $d 'confirm-hard'
$d.Close()

# 9. The verification after a soft delete
$script:NextResult = [pscustomobject]@{ Delete = [pscustomobject]@{ Status = 'Completed'; Errors = '' }; SearchName = 'EMT_20261110_094610'
    Check = @(
        [pscustomobject]@{ Mailbox = 'jsmith@contoso.com'; Visible = 0; Total = 2; Note = '' }
        [pscustomobject]@{ Mailbox = 'akim@contoso.com'; Visible = 0; Total = 1; Note = '' }
        [pscustomobject]@{ Mailbox = 'mlee@contoso.com'; Visible = 1; Total = 1; Note = '' }
        [pscustomobject]@{ Mailbox = 'tnovak@contoso.com'; Visible = -1; Total = -1; Note = 'Search-Mailbox: timeout' }) }
Click $btnSoft
Snap $tabs.Parent 'verification' 6

# 10. The selection of an address (a part of a name)
$cands = @(
    [pscustomobject]@{ Name = 'Anna Kim'; Address = 'akim@contoso.com'; Type = 'UserMailbox' }
    [pscustomobject]@{ Name = 'Andrew Kimball'; Address = 'akimball@contoso.com'; Type = 'UserMailbox' }
    [pscustomobject]@{ Name = 'Kim Support'; Address = 'kim.support@contoso.com'; Type = 'SharedMailbox' })
$pw = New-EmtPickWindow -Text 'kim' -Candidates $cands
$pw.Window.Owner = $null; $pw.Window.WindowStartupLocation = 'Manual'; $pw.Window.Left = -20000; $pw.Window.Top = 0; $pw.Window.ShowActivated = $false
$pw.Window.Show(); Pump
SnapWindow $pw.Window 'pick-address'
$pw.Window.Close()

# 11. The top line at the minimum width: all controls must be visible
$window.Width = $window.MinWidth; Pump
Snap $tbServer.Parent.Parent 'top-line-min' 8

$window.Close()
'@
$text = $text.Replace('[void]$window.ShowDialog()', $scenes.Replace('__OUT__', $out))
try {
    & ([scriptblock]::Create($text))
} finally {
    Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
