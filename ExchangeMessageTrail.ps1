#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Message Trail (EMT): message search and, when necessary, delete for Exchange Server 2016/2019 (on-premises).

.DESCRIPTION
    Tab "Message search": find a message in the message tracking log, see the delivery status of each recipient,
    see why a message did not arrive, examine the mailbox of a recipient. Read-only. Seconds.
    Tab "Delete": find the copies of a message in the mailboxes and delete them.
    All Exchange work runs in a background runspace. The window does not freeze.
    The window is WPF: gui\MainWindow.xaml, the styles gui\Styles.xaml, the themes gui\Theme.*.xaml.
    The user interface languages are the files in the folder "Lang" (*.lng).
    See README.md for requirements, roles and the procedure.

.PARAMETER Server
    The FQDN of the Exchange server. If you do not set it, the tool uses the last value.

.PARAMETER Language
    The language code of a file in "Lang" (for example ru or en). If you do not set it,
    the tool uses the last value, or the Windows language.

.PARAMETER Theme
    Dark or Light. If you do not set it, the tool uses the last value, or Dark.

.EXAMPLE
    .\ExchangeMessageTrail.ps1 -Server exch01.contoso.local -Language en -Theme Dark
#>
[CmdletBinding()]
param([string] $Server, [string] $Language, [ValidateSet('', 'Dark', 'Light')] [string] $Theme = '')

$ErrorActionPreference = 'Stop'

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Start Exchange Message Trail in STA mode: powershell.exe -STA -File ExchangeMessageTrail.ps1 (or use EMT.cmd).'
}

# Windows Forms: only for the folder selection of the export (WPF has no folder dialog).
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms

# The console is hidden: show a start error in a window. The language files are not loaded yet: English text.
trap {
    [void][System.Windows.MessageBox]::Show(
        "Exchange Message Trail could not start.`r`n`r`n$($_.Exception.Message)`r`n`r`n$($_.InvocationInfo.PositionMessage)",
        'Exchange Message Trail', 'OK', 'Error')
    break
}

# One Add-Type for all Win32 functions: each Add-Type starts the C# compiler (about 0.1 s).
# Only the classes that are not loaded yet (the other EMT window can have loaded some in this session).
$win32 = [ordered]@{
    EmtConsole = @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("kernel32.dll")] public static extern uint GetConsoleProcessList(uint[] list, uint count);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int cmd);
'@
    # The color of the title bar (Windows 10 1809 and later, Windows Server 2019 and later).
    EmtDwm = @'
[DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
'@
    EmtShell = @'
[DllImport("shell32.dll")] public static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string id);
'@
}
$missing = @($win32.Keys | Where-Object { -not ($_ -as [type]) })
if ($missing.Count) {
    Add-Type -TypeDefinition ("using System;`nusing System.Runtime.InteropServices;`n" +
        (($missing | ForEach-Object { "public static class $_ {`n$($win32[$_])}" }) -join "`n"))
}
# An own application ID: the taskbar shows EMT as a separate button with the EMT icon, not as PowerShell.
# Set it before the first window opens.
try { [void][EmtShell]::SetCurrentProcessExplicitAppUserModelID('ExchangeMessageTrail') } catch { }

# Hide the PowerShell console window, but only if it belongs to EMT alone.
# One process in the console: the console was made for EMT (EMT.cmd, a shortcut, "Run").
# More processes: the user started EMT from an open console. That console stays visible.
$consoleWindow = [EmtConsole]::GetConsoleWindow()
if ($consoleWindow -ne [IntPtr]::Zero -and [EmtConsole]::GetConsoleProcessList((New-Object 'System.UInt32[]' 4), 4) -eq 1 -and
    [Environment]::CommandLine -match '\s-File\s') {
    [void][EmtConsole]::ShowWindow($consoleWindow, 0)    # 0 = SW_HIDE
}

$CorePath = Join-Path $PSScriptRoot 'src\EMT.Core.ps1'
. $CorePath
$GuiDir = Join-Path $PSScriptRoot 'gui'

# The EMT icon: the title bar of all windows and the taskbar. A missing file gives the PowerShell icon.
$script:EmtIcon = $null
# OnLoad: the file is read at once and stays free (an update of the folder can replace it while EMT runs).
try { $script:EmtIcon = [System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object Uri (Join-Path $GuiDir 'emt.ico')),
          [System.Windows.Media.Imaging.BitmapCreateOptions]::None, [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad) } catch { }

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

$script:Caps            = $null      # capabilities after Connect
$script:Trace           = $null      # result of the last message search
$script:MailboxCheck    = $null      # result of the last mailbox check (MessageId, Mailbox, Findings)
$script:Result          = $null      # result of the last search on the tab "Delete"
$script:Compact         = $null      # the criteria cards of a small window (see Update-EmtCompact)
$script:Check           = $null      # result of the last verification
$script:Job             = $null      # running background operation
$script:Loading         = $false     # true while the UI fills controls (no invalidation)
$script:ConnectedServer = ''
$script:DelFromTrace    = ''         # MessageId when the tab "Delete" has the data of a found message
$script:CheckForMessage = ''         # MessageId of the running mailbox check
$script:ResolveNext     = $null      # the action after the address book search
$script:PendingCleanup  = $null      # names confirmed for cleanup; the timer starts the removal
$script:CleanupNames    = $null
$script:Deleted         = ''         # last delete on the current result: '', 'Soft' or 'Hard'
$script:ReportByMsg     = @{}        # MessageId -> the report rows of the message
$script:EventsByMsg     = $null      # MessageId -> the events of the message; $null = make it again from $script:Trace
$script:DetailLoaded    = @{}        # MessageId -> $true: the message has all recipients (partial message search)
$script:PendingDetail   = ''         # MessageId: load its recipients when no operation runs
$script:PcCulture       = [Globalization.CultureInfo]::CurrentCulture   # the regional settings of Windows (Set-EmtDateCulture)
$script:Theme           = 'Dark'
$script:LanguageCodes   = @()
$script:ThemeCodes      = @('Dark', 'Light')
$script:GridHeaders     = @{}        # grid -> ordered map "property" -> "text key" of the column header
$script:GridNeedsSize   = @{}        # grid -> true: set the column widths when the grid is visible
$script:GridLastWidth   = @{}        # grid -> the width of the last column from its content
$script:GridResizeHooked = @{}       # grid -> true: the grid has the SizeChanged handler
$script:Tips            = New-Object System.Collections.Generic.List[object]   # control + tooltip key
$script:XamlText        = @{}        # file name -> text (the dialogs load the theme again)
$script:LastTime        = @{}        # time field -> the last good time text

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
    <#
        Senders and subject are not kept: they can be confidential.
        The window size is in WPF units (key WpfWindow). The classic window keeps its size in pixels (key Window).
        Other keys of the file stay: the classic window can use the same file.
    #>
    $s = Read-EmtSettings
    $s.Server          = $tbServer.Text.Trim()
    $s.Language        = [string]$Sync.Lang
    $s.Theme           = $script:Theme
    $s.IncludeForwards = [bool]$chkForwards.IsChecked
    $s.AllMailboxes    = [bool]$chkAll.IsChecked
    $s.TimeoutMinutes  = Get-EmtTimeout
    # The window: the normal size and position, also when the window is maximized.
    $b = if ($window.WindowState -eq 'Normal') { New-Object System.Windows.Rect($window.Left, $window.Top, $window.ActualWidth, $window.ActualHeight) }
         else { $window.RestoreBounds }
    if ($b.Width -gt 0 -and $b.Height -gt 0) {
        $s.WpfWindow = [ordered]@{ X = [int]$b.X; Y = [int]$b.Y; Width = [int]$b.Width; Height = [int]$b.Height
                                   Maximized = ($window.WindowState -eq 'Maximized') }
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

# The core loads once: the functions stay in the runspace for the next operations (about 0.1 s for each operation).
$OpScript = @'
param($CorePath, $Name, $Params)
if (-not (Get-Command Invoke-EmtOperation -ErrorAction SilentlyContinue)) { . $CorePath }
Invoke-EmtOperation -Name $Name -Params $Params
'@

# An operation that runs longer than this number of seconds writes its time to the log.
$script:SlowOpSeconds = 5

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
        Timer = [System.Diagnostics.Stopwatch]::StartNew()
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
    $seconds = $job.Timer.Elapsed.TotalSeconds
    if ($seconds -gt $script:SlowOpSeconds) { Write-EmtLog (T 'log.opSlow' $job.Name $seconds) }
    foreach ($e in $job.PS.Streams.Error) { Write-EmtLog ([string]$e) 'Warn' }
    $job.PS.Dispose()
    $progress.Value = 0
    Set-EmtStatusText ''

    if ($err) {
        if ($Sync.Cancel) {
            Write-EmtLog (T 'log.opCancelled' $job.Name) 'Warn'
        } else {
            Write-EmtLog (T 'log.opFailed' $job.Name $err.Message) 'Error'
            [void](Show-EmtMessage -Text $err.Message -Title (T 'dlg.error') -Kind Error)
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

#region ---------- XAML, theme, texts ----------

function Read-EmtXaml([string] $Name) {
    <# Makes the object of a XAML file in the folder "gui". #>
    if (-not $script:XamlText.ContainsKey($Name)) {
        $script:XamlText[$Name] = [IO.File]::ReadAllText((Join-Path $GuiDir $Name), [Text.Encoding]::UTF8)
    }
    [System.Windows.Markup.XamlReader]::Parse($script:XamlText[$Name])
}

function Get-EmtThemeDictionary([string] $Name) {
    <#
        The colors of a theme. The standard templates use the system colors (for example the corner of the scroll bars).
        The function adds these keys with the colors of the theme. XAML cannot have more than one system key in one file.
    #>
    $d = Read-EmtXaml "Theme.$Name.xaml"
    $map = [ordered]@{
        ControlBrushKey                    = 'Bg.Window'
        WindowBrushKey                     = 'Bg.Input'
        ControlTextBrushKey                = 'Fg'
        WindowTextBrushKey                 = 'Fg'
        GrayTextBrushKey                   = 'Fg.Disabled'
        HighlightBrushKey                  = 'Sel.Bg'
        HighlightTextBrushKey              = 'Sel.Fg'
        InactiveSelectionHighlightBrushKey = 'Sel.Bg'
    }
    foreach ($k in $map.Keys) { $d[[System.Windows.SystemColors]::$k] = $d[$map[$k]] }
    $d
}

function Add-EmtTheme($Target) {
    <# Adds the theme colors and the styles to a window. The theme is always the first dictionary. #>
    $Target.Resources.MergedDictionaries.Clear()
    $Target.Resources.MergedDictionaries.Add((Get-EmtThemeDictionary $script:Theme))
    $Target.Resources.MergedDictionaries.Add((Read-EmtXaml 'Styles.xaml'))
}

function Set-EmtTitleBar($Target) {
    <# A dark title bar for the dark theme. Older Windows versions ignore the call. #>
    try {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($Target)).EnsureHandle()
        $v = [int]($script:Theme -eq 'Dark')
        foreach ($attr in 20, 19) {     # 20 = DWMWA_USE_IMMERSIVE_DARK_MODE; 19 = the same before Windows 10 20H1
            if ([EmtDwm]::DwmSetWindowAttribute($hwnd, $attr, [ref]$v, 4) -eq 0) { break }
        }
    } catch { }
}

function Set-EmtTheme([string] $Name) {
    $script:Theme = $Name
    $window.Resources.MergedDictionaries[0] = Get-EmtThemeDictionary $Name
    Set-EmtTitleBar $window
    # The title bar changes only after the next activation. A quick size change shows it now.
    if ($window.IsLoaded -and $window.WindowState -eq 'Normal') { $window.Width += 1; $window.Width -= 1 }
}

function Set-EmtBrush($Element, [string] $Key, [System.Windows.DependencyProperty] $Property = [System.Windows.Controls.TextBlock]::ForegroundProperty) {
    <# Connects a property to a theme color. The color follows a theme change. #>
    $Element.SetResourceReference($Property, $Key)
}

function Get-EmtTextKey($Element) {
    if ($Element -is [System.Windows.FrameworkElement] -and $Element.Tag -is [string] -and $Element.Tag) { [string]$Element.Tag }
}

function Update-EmtTexts($Node) {
    <# Sets the texts of all elements with a text key in Tag (TextBlock, button, check box, tab). #>
    $key = Get-EmtTextKey $Node
    if ($key) {
        $text = T $key
        if ($Node -is [System.Windows.Controls.TextBlock]) { $Node.Text = $text }
        elseif ($Node -is [System.Windows.Controls.HeaderedContentControl]) { $Node.Header = $text }
        elseif ($Node -is [System.Windows.Controls.ContentControl]) { $Node.Content = $text }
    }
    if ($Node -is [System.Windows.DependencyObject]) {
        foreach ($c in [System.Windows.LogicalTreeHelper]::GetChildren($Node)) { Update-EmtTexts $c }
    }
}

function Set-EmtTip($Control, [string] $Key) {
    $script:Tips.Add([pscustomobject]@{ Control = $Control; Key = $Key })
    $Control.ToolTip = T $Key
    [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($Control, $true)
}

function Set-EmtStatusText([string] $Text) {
    $lblStatus.Text = $Text
    $lblStatus.ToolTip = $(if ($Text) { $Text } else { $null })
}

function Set-EmtDateCulture {
    <#
        The date pickers use the culture of the window thread. This culture is a copy of the regional settings of Windows:
        the date format and the first day of the week are the same as in the other programs on this computer.
        The names of the months and the week days come from the interface language.
        The window must not set its property Language: then the date pickers use the culture of the thread.
    #>
    $c = $script:PcCulture.Clone()
    try {
        $code = if ($Sync.Lang -eq 'en') { 'en-US' } else { [string]$Sync.Lang }
        $ui = [Globalization.CultureInfo]::CreateSpecificCulture($code).DateTimeFormat
        $f = $c.DateTimeFormat
        $f.MonthNames                    = $ui.MonthNames
        $f.MonthGenitiveNames            = $ui.MonthGenitiveNames
        $f.AbbreviatedMonthNames         = $ui.AbbreviatedMonthNames
        $f.AbbreviatedMonthGenitiveNames = $ui.AbbreviatedMonthGenitiveNames
        $f.DayNames                      = $ui.DayNames
        $f.AbbreviatedDayNames           = $ui.AbbreviatedDayNames
        $f.ShortestDayNames              = $ui.ShortestDayNames
        $f.YearMonthPattern              = $ui.YearMonthPattern    # the header of the calendar: "November 2026"
    } catch { }    # an unknown language code: the names of Windows stay
    [Threading.Thread]::CurrentThread.CurrentCulture = $c
    # The text of the date pickers: show the dates again in the new culture.
    # The same dates again: this is not a change of the criteria (no invalidation of the delete result).
    $was = $script:Loading
    $script:Loading = $true
    try {
        foreach ($d in $dtTrStart, $dtTrEnd, $dtStart, $dtEnd) {
            $v = $d.SelectedDate
            $d.SelectedDate = $null
            $d.SelectedDate = $v
        }
    } finally { $script:Loading = $was }
}

function Update-EmtLanguage {
    <# Applies the current language to all controls, headers and result views. #>
    $window.Title = T 'app.title'
    Set-EmtDateCulture
    Update-EmtTexts $window
    foreach ($x in $script:Tips) { $x.Control.ToolTip = T $x.Key }
    if ($script:Compact) { foreach ($c in $script:Compact) { if ($c.On) { $c.Text.Text = Get-EmtCompactText $c; $c.Text.ToolTip = $c.Text.Text } } }
    $was = $script:Loading
    $script:Loading = $true
    try {
        $cmbTheme.Items.Clear()
        foreach ($t in $script:ThemeCodes) { [void]$cmbTheme.Items.Add((T "theme.$t")) }
        $cmbTheme.SelectedIndex = [array]::IndexOf($script:ThemeCodes, $script:Theme)
    } finally { $script:Loading = $was }
    Update-EmtConnIndicator
    Update-EmtDeleteSource
    if ($script:Trace) { Show-EmtTraceViews $script:Trace } else { Clear-EmtTraceViews }
    if ($script:Result) { Show-EmtResultViews $script:Result } else { Clear-EmtResultViews }
    if ($script:Check) { Show-EmtCheckView $script:Check }
    Update-EmtTabTitles
    Set-UiState
}

#endregion

#region ---------- Dialogs ----------

function New-EmtDialog([string] $Title, [double] $Width) {
    $w = New-Object System.Windows.Window
    $w.Title = $Title
    $w.WindowStartupLocation = 'CenterOwner'
    $w.ShowInTaskbar = $false
    $w.ResizeMode = 'NoResize'
    $w.FontFamily = 'Segoe UI'; $w.FontSize = 12
    $w.UseLayoutRounding = $true
    if ($script:EmtIcon) { $w.Icon = $script:EmtIcon }
    if ($Width) { $w.Width = $Width; $w.SizeToContent = 'Height' } else { $w.SizeToContent = 'WidthAndHeight' }
    if ($window -and $window.IsLoaded) { $w.Owner = $window }
    Add-EmtTheme $w
    Set-EmtBrush $w 'Bg.Window' ([System.Windows.Controls.Control]::BackgroundProperty)
    Set-EmtBrush $w 'Fg' ([System.Windows.Controls.Control]::ForegroundProperty)
    Set-EmtTitleBar $w
    $w
}

function Show-EmtMessage {
    <# A message in the colors of the theme. Kind Question has the buttons Yes and No. Returns $true for OK or Yes. #>
    param([string] $Text, [string] $Title = (T 'app.short'), [ValidateSet('Info', 'Warn', 'Error', 'Question')] [string] $Kind = 'Info')
    $w = New-EmtDialog -Title $Title -Width 520
    $root = New-Object System.Windows.Controls.DockPanel
    $root.Margin = '20,18,20,16'

    $buttons = New-Object System.Windows.Controls.StackPanel
    $buttons.Orientation = 'Horizontal'; $buttons.HorizontalAlignment = 'Right'; $buttons.Margin = '0,18,0,0'
    [System.Windows.Controls.DockPanel]::SetDock($buttons, 'Bottom')
    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = $(if ($Kind -eq 'Question') { T 'dlg.yes' } else { T 'dlg.ok' })
    $ok.IsDefault = $true; $ok.MinWidth = 96
    $ok.Add_Click({ $w.DialogResult = $true }.GetNewClosure())
    [void]$buttons.Children.Add($ok)
    if ($Kind -eq 'Question') {
        $no = New-Object System.Windows.Controls.Button
        $no.Content = T 'dlg.no'; $no.IsCancel = $true; $no.MinWidth = 96; $no.Margin = '8,0,0,0'
        [void]$buttons.Children.Add($no)
    }
    [void]$root.Children.Add($buttons)

    # A color bar on the left: red = error, orange = warning or question, green = information.
    $bar = New-Object System.Windows.Controls.Border
    $bar.Width = 4; $bar.CornerRadius = 2; $bar.Margin = '0,0,14,0'
    $barKey = switch ($Kind) { 'Error' { 'Text.Red' } 'Info' { 'Accent' } default { 'Text.Orange' } }
    Set-EmtBrush $bar $barKey ([System.Windows.Controls.Border]::BackgroundProperty)
    [System.Windows.Controls.DockPanel]::SetDock($bar, 'Left')
    [void]$root.Children.Add($bar)

    $body = New-Object System.Windows.Controls.TextBlock
    $body.Text = $Text -replace "`r?`n", "`r`n"
    $body.TextWrapping = 'Wrap'
    $scroll = New-Object System.Windows.Controls.ScrollViewer
    $scroll.MaxHeight = 360; $scroll.VerticalScrollBarVisibility = 'Auto'
    $scroll.Content = $body
    [void]$root.Children.Add($scroll)

    $w.Content = $root
    $answer = $w.ShowDialog()
    [bool]$answer
}

function Show-EmtPick {
    <# Selection of one or more addresses from the candidates. Returns the selected addresses, or $null on cancel. #>
    param([string] $Text, [object[]] $Candidates, [switch] $Multi, [int] $Max = 25)
    $d = New-EmtPickWindow -Text $Text -Candidates $Candidates -Multi:$Multi -Max $Max
    $answer = $d.Window.ShowDialog()
    $picked = @(foreach ($item in $d.List.SelectedItems) { $Candidates[$d.List.Items.IndexOf($item)].Address })
    if (-not $answer -or -not $picked.Count) { return $null }
    , $picked
}

function New-EmtPickWindow {
    <# Makes the selection window of the addresses. It does not show it. Returns the window and the list. #>
    param([string] $Text, [object[]] $Candidates, [switch] $Multi, [int] $Max = 25)
    $w = New-EmtDialog -Title (T 'pick.title') -Width 640
    $root = New-Object System.Windows.Controls.DockPanel
    $root.Margin = '18,16,18,14'

    $top = New-Object System.Windows.Controls.TextBlock
    $top.TextWrapping = 'Wrap'; $top.Margin = '0,0,0,10'
    $top.Text = $(if ($Multi) { T 'pick.promptMulti' $Text } else { T 'pick.prompt' $Text })
    if ($Candidates.Count -ge $Max) { $top.Text += "`r`n" + (T 'pick.tooMany' $Max) }
    [System.Windows.Controls.DockPanel]::SetDock($top, 'Top')
    [void]$root.Children.Add($top)

    $buttons = New-Object System.Windows.Controls.StackPanel
    $buttons.Orientation = 'Horizontal'; $buttons.HorizontalAlignment = 'Right'; $buttons.Margin = '0,12,0,0'
    [System.Windows.Controls.DockPanel]::SetDock($buttons, 'Bottom')
    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = T 'pick.ok'; $ok.IsDefault = $true; $ok.MinWidth = 110
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = T 'dlg.cancel'; $cancel.IsCancel = $true; $cancel.MinWidth = 110; $cancel.Margin = '8,0,0,0'
    [void]$buttons.Children.Add($ok); [void]$buttons.Children.Add($cancel)
    [void]$root.Children.Add($buttons)

    $list = New-Object System.Windows.Controls.ListBox
    $list.Height = 260
    $list.SelectionMode = $(if ($Multi) { 'Extended' } else { 'Single' })
    foreach ($c in $Candidates) { [void]$list.Items.Add(('{0}  <{1}>  {2}' -f $c.Name, $c.Address, $c.Type)) }
    if ($list.Items.Count) { $list.SelectedIndex = 0 }
    [void]$root.Children.Add($list)

    $ok.Add_Click({ $w.DialogResult = $true }.GetNewClosure())
    $list.Add_MouseDoubleClick({ if ($list.SelectedIndex -ge 0) { $w.DialogResult = $true } }.GetNewClosure())

    $w.Content = $root
    [pscustomobject]@{ Window = $w; List = $list }
}

function New-EmtDialogButtons([string] $OkText, [string] $OkStyle = '') {
    <# The buttons at the bottom of a dialog: OK (or $OkText) and Cancel. Returns the panel, OK and Cancel. #>
    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Orientation = 'Horizontal'; $panel.HorizontalAlignment = 'Right'; $panel.Margin = '0,14,0,0'
    [System.Windows.Controls.DockPanel]::SetDock($panel, 'Bottom')
    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = $OkText; $ok.MinWidth = 130
    if ($OkStyle) { $ok.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, $OkStyle) }
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = T 'dlg.cancel'; $cancel.IsCancel = $true; $cancel.MinWidth = 110; $cancel.Margin = '8,0,0,0'
    [void]$panel.Children.Add($ok); [void]$panel.Children.Add($cancel)
    [pscustomobject]@{ Panel = $panel; Ok = $ok; Cancel = $cancel }
}

function Show-EmtConfirm {
    <# Shows the confirmation dialog (New-EmtConfirmWindow). Returns $true if the user confirms. #>
    param([string] $Title, [string] $Text, [string] $OkText, [string] $RequireWord)
    [bool](New-EmtConfirmWindow -Title $Title -Text $Text -OkText $OkText -RequireWord $RequireWord).ShowDialog()
}

function New-EmtConfirmWindow {
    <#
        Makes the confirmation dialog of a delete or a cleanup. It does not show it.
        The text is in a read-only box with a scroll bar.
        With -RequireWord the OK button is enabled only after the user types the word.
        The focus is on Cancel: Enter does not confirm a delete by mistake.
    #>
    param([string] $Title, [string] $Text, [string] $OkText, [string] $RequireWord)
    $w = New-EmtDialog -Title $Title -Width 700
    $root = New-Object System.Windows.Controls.DockPanel
    $root.Margin = '18,16,18,14'

    $b = New-EmtDialogButtons -OkText $OkText -OkStyle 'DangerButton'
    $ok = $b.Ok; $cancel = $b.Cancel
    [void]$root.Children.Add($b.Panel)
    if ($RequireWord) {
        # The word has its own line over the buttons: a long prompt must not push the buttons out of the window.
        $line = New-Object System.Windows.Controls.StackPanel
        $line.Orientation = 'Horizontal'; $line.Margin = '0,12,0,0'
        [System.Windows.Controls.DockPanel]::SetDock($line, 'Bottom')
        $prompt = New-Object System.Windows.Controls.TextBlock
        $prompt.Text = T 'dlg.typeWord' $RequireWord; $prompt.Margin = '0,0,10,0'
        $word = New-Object System.Windows.Controls.TextBox
        $word.Width = 160
        [void]$line.Children.Add($prompt); [void]$line.Children.Add($word)
        [void]$root.Children.Add($line)
        $ok.IsEnabled = $false
        $word.Add_TextChanged({ $ok.IsEnabled = ($word.Text -ceq $RequireWord) }.GetNewClosure())
    }

    $body = New-Object System.Windows.Controls.TextBox
    $body.Text = $Text -replace "`r?`n", "`r`n"
    $body.IsReadOnly = $true; $body.TextWrapping = 'Wrap'; $body.AcceptsReturn = $true
    $body.Height = 420; $body.VerticalContentAlignment = 'Top'; $body.Padding = '8,6'
    $body.VerticalScrollBarVisibility = 'Auto'
    # The text box style hides the scroll bars (a one-line field). A long text needs them.
    $body.Template = [System.Windows.Markup.XamlReader]::Parse(@'
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="TextBox">
  <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="4">
    <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"
                  VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"/>
  </Border>
</ControlTemplate>
'@)
    [void]$root.Children.Add($body)

    $ok.Add_Click({ $w.DialogResult = $true }.GetNewClosure())
    $w.Add_ContentRendered({ [void]$cancel.Focus() }.GetNewClosure())
    $w.Content = $root
    $w
}

function Show-EmtInput {
    <# A question with a one-line answer. Returns the text, or $null on cancel. #>
    param([string] $Title, [string] $Prompt, [string] $Value)
    $w = New-EmtDialog -Title $Title -Width 520
    $root = New-Object System.Windows.Controls.DockPanel
    $root.Margin = '18,16,18,14'
    $b = New-EmtDialogButtons -OkText (T 'dlg.ok')
    $b.Ok.IsDefault = $true
    [void]$root.Children.Add($b.Panel)
    $stack = New-Object System.Windows.Controls.StackPanel
    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = $Prompt; $label.TextWrapping = 'Wrap'; $label.Margin = '0,0,0,10'
    $box = New-Object System.Windows.Controls.TextBox
    $box.Text = $Value; $box.Width = 120; $box.HorizontalAlignment = 'Left'
    [void]$stack.Children.Add($label); [void]$stack.Children.Add($box)
    [void]$root.Children.Add($stack)
    $b.Ok.Add_Click({ $w.DialogResult = $true }.GetNewClosure())
    $w.Add_ContentRendered({ [void]$box.Focus(); $box.SelectAll() }.GetNewClosure())
    $w.Content = $root
    if ($w.ShowDialog()) { $box.Text } else { $null }
}

function Open-EmtHelp {
    <# Opens the user guide in the interface language. If the guide has no page for this language, the English page opens. #>
    $dir = Join-Path $PSScriptRoot 'help'
    $page = Join-Path $dir "$($Sync.Lang).html"
    if (-not (Test-Path $page)) { $page = Join-Path $dir 'en.html' }
    if (-not (Test-Path $page)) { [void](Show-EmtMessage -Text (T 'help.missing' $page) -Kind Warn); return }
    Start-Process -FilePath $page
}

function Export-EmtSets([string] $Stem, [System.Collections.Specialized.OrderedDictionary] $Sets) {
    <# Saves each set as <folder>\<stem>-<name>.csv. The user selects the folder (the standard Windows dialog). #>
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = T 'dlg.exportFolder'
    $dlg.SelectedPath = $AppDir
    $owner = New-Object System.Windows.Forms.NativeWindow
    $owner.AssignHandle((New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle)
    try {
        if ($dlg.ShowDialog($owner) -ne 'OK') { return }
        $folder = $dlg.SelectedPath
    } finally { $owner.ReleaseHandle(); $dlg.Dispose() }
    $base = Join-Path $folder $Stem
    foreach ($k in $Sets.Keys) {
        if (@($Sets[$k]).Count) {
            $Sets[$k] | Export-Csv -Path "$base-$k.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
        }
    }
    Write-EmtLog (T 'log.exported' $base) 'Ok'
}

#endregion

#region ---------- Tables ----------

function Set-EmtGridData {
    <#
        Binds objects to a grid. $Columns: ordered map "Property" -> "header text key".
        A property with an empty key is data only (not shown).
        $Mark gives the mark on the left of a row: Red (an error), Yellow (a delay, a warning) or nothing.
        $PartsColumn shows the statuses with colors: $Parts gives the parts of an object (hashtables Text, Tone).
        The sort of this column uses the text of the property $PartsColumn.
    #>
    param($Grid, [object[]] $Items = @(), [System.Collections.Specialized.OrderedDictionary] $Columns,
          [scriptblock] $Mark, [string] $PartsColumn, [scriptblock] $Parts)
    $items = @($Items | Where-Object { $_ -ne $null })
    $dt = New-Object System.Data.DataTable
    foreach ($prop in $Columns.Keys) {
        $sample = $null
        foreach ($it in $items) { $p = $it.PSObject.Properties[$prop]; if ($p -and $null -ne $p.Value) { $sample = $p.Value; break } }
        $type = [string]
        if ($sample -is [int] -or $sample -is [long]) { $type = [long] }
        elseif ($sample -is [double] -or $sample -is [decimal]) { $type = [double] }
        elseif ($sample -is [datetime]) { $type = [datetime] }
        [void]$dt.Columns.Add($prop, $type)
    }
    [void]$dt.Columns.Add('Mark', [string])
    [void]$dt.Columns.Add('Parts', [object])
    foreach ($it in $items) {
        $row = $dt.NewRow()
        # The property directly, not Get-EmtProp: a function call for each cell is slow for thousands of rows.
        $props = $it.PSObject.Properties
        foreach ($prop in $Columns.Keys) {
            $p = $props[$prop]
            $v = if ($p) { $p.Value } else { $null }
            if ($null -eq $v) { $row[$prop] = [DBNull]::Value }
            elseif ($dt.Columns[$prop].DataType -eq [string]) {
                $row[$prop] = if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { (@($v) -join '; ') } else { [string]$v }
            }
            else { $row[$prop] = $v }
        }
        $row['Mark'] = $(if ($Mark) { [string](& $Mark $it) } else { '' })
        if ($Parts) { $row['Parts'] = [object[]]@(& $Parts $it) }
        $dt.Rows.Add($row)
    }

    $Grid.ItemsSource = $null
    $Grid.Columns.Clear()
    if ($Mark) {
        $col = New-Object System.Windows.Controls.DataGridTemplateColumn
        $col.CellTemplate = $window.FindResource('EmtMarkCell')
        $col.CellStyle = $window.FindResource('EmtMarkCellStyle')
        $col.Width = 3; $col.MinWidth = 3; $col.CanUserResize = $false; $col.CanUserSort = $false; $col.CanUserReorder = $false
        $Grid.Columns.Add($col)
    }
    foreach ($prop in $Columns.Keys) {
        if (-not $Columns[$prop]) { continue }
        if ($Parts -and $prop -eq $PartsColumn) {
            $col = New-Object System.Windows.Controls.DataGridTemplateColumn
            $col.CellTemplate = $window.FindResource('EmtPartsCell')
        } else {
            $col = New-Object System.Windows.Controls.DataGridTextColumn
            $b = New-Object System.Windows.Data.Binding($prop)
            if ($dt.Columns[$prop].DataType -eq [datetime]) { $b.StringFormat = T 'fmt.datetime' }
            $col.Binding = $b
        }
        $col.SortMemberPath = $prop
        $col.Header = T $Columns[$prop]
        $col.Width = [System.Windows.Controls.DataGridLength]::Auto
        $Grid.Columns.Add($col)
    }
    $script:GridHeaders[$Grid] = $Columns
    $script:GridNeedsSize[$Grid] = $true
    $script:GridLastWidth.Remove($Grid)
    if (-not $script:GridResizeHooked.ContainsKey($Grid)) {
        $script:GridResizeHooked[$Grid] = $true
        $Grid.Add_SizeChanged({ param($s, $e) if ($e.WidthChanged) { Update-EmtLastColumn $s } })
    }
    $Grid.ItemsSource = $dt.DefaultView
    # The widths come from the content after the layout. The dispatcher calls the function after the layout.
    [void]$window.Dispatcher.BeginInvoke([Action]{ Resize-EmtGrids }, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
}

function Clear-EmtGrid($Grid) {
    $Grid.ItemsSource = $null
    $Grid.Columns.Clear()
    $script:GridNeedsSize[$Grid] = $false
}

$script:Typefaces = @{}

function Measure-EmtText([string] $Text, $Grid, [switch] $Bold) {
    <# The width of a text in the font of a grid, in pixels. #>
    if (-not $Text) { return 0 }
    $weight = if ($Bold) { [System.Windows.FontWeights]::SemiBold } else { $Grid.FontWeight }
    # One typeface for each font: the column widths measure thousands of texts. ::new() is faster than New-Object.
    $key = '{0}|{1}|{2}|{3}' -f $Grid.FontFamily, $Grid.FontStyle, $weight, $Grid.FontStretch
    $face = $script:Typefaces[$key]
    if (-not $face) {
        $face = [System.Windows.Media.Typeface]::new($Grid.FontFamily, $Grid.FontStyle, $weight, $Grid.FontStretch)
        $script:Typefaces[$key] = $face
    }
    $ft = [System.Windows.Media.FormattedText]::new($Text, [Globalization.CultureInfo]::CurrentCulture,
              [System.Windows.FlowDirection]::LeftToRight, $face, $Grid.FontSize, [System.Windows.Media.Brushes]::Black)
    $ft.WidthIncludingTrailingWhitespace
}

function Get-EmtColumnWidth($Grid, $Column) {
    <#
        The width of a column from its header and its values (up to 300 rows): the widest text and the paddings.
        The automatic width of WPF can come before the rows are on the grid: then it fits the header only.
    #>
    $w = (Measure-EmtText ([string]$Column.Header) $Grid -Bold) + 30      # header padding 8 + 14, the sort arrow
    $prop = $Column.SortMemberPath
    $format = if ($Column -is [System.Windows.Controls.DataGridTextColumn] -and $Column.Binding.StringFormat) { $Column.Binding.StringFormat } else { '' }
    $n = 0
    foreach ($v in @($Grid.ItemsSource)) {
        if (++$n -gt 300) { break }
        $cell = 0
        if ($Column -is [System.Windows.Controls.DataGridTemplateColumn] -and $v.Row.Table.Columns.Contains('Parts') -and $v.Row['Parts'] -is [array]) {
            # The status parts: a dot (14 px) and a text for each part, 14 px between the parts.
            foreach ($p in $v.Row['Parts']) { $cell += 14 + (Measure-EmtText ([string]$p['Text']) $Grid) + 14 }
        } elseif ($prop -and $v.Row.Table.Columns.Contains($prop)) {
            $x = $v.Row[$prop]
            $s = if ($x -is [DBNull]) { '' } elseif ($format -and $x -is [datetime]) { $x.ToString($format) } else { [string]$x }
            $cell = Measure-EmtText $s $Grid
        }
        $w = [math]::Max($w, $cell + 32)                                  # cell padding 8 + 8, the margin of the text, the grid line
    }
    [math]::Ceiling($w)
}

function Resize-EmtGrids {
    <#
        Initial column width from the content: 60-450 px. Then the width is fixed:
        a sort and a scroll keep the widths that the user set. The last column gets the free space.
        A grid on a hidden tab gets the widths when the tab opens.
    #>
    foreach ($g in @($script:GridNeedsSize.Keys)) {
        if (-not $script:GridNeedsSize[$g] -or -not $g.IsVisible -or -not $g.Columns.Count) { continue }
        $script:GridNeedsSize[$g] = $false
        # The mark column has a fixed width: it cannot change.
        $cols = @($g.Columns | Where-Object { $_.Visibility -eq 'Visible' -and $_.CanUserResize })
        $sum = 0
        foreach ($c in $cols) {
            $w = [math]::Max(60, [math]::Min(450, (Get-EmtColumnWidth $g $c)))
            $c.Width = New-Object System.Windows.Controls.DataGridLength($w)
            $sum += $w
        }
        # The last column gets the free space only if all columns fit. A star column with too little space
        # makes the grid narrow the other columns: then the texts are cut. Without free space, the grid scrolls.
        $script:GridLastWidth[$g] = $cols[-1].Width.Value
        Update-EmtLastColumn $g
    }
}

function Update-EmtLastColumn($Grid) {
    <#
        The last column gets the free space only if all columns fit. A star column with too little space
        makes the grid narrow the other columns: then the texts are cut. Without free space, the grid scrolls.
        The grid calls this function also when its width changes (the window size).
    #>
    if (-not $script:GridLastWidth.ContainsKey($Grid) -or $script:GridNeedsSize[$Grid]) { return }
    $cols = @($Grid.Columns | Where-Object { $_.Visibility -eq 'Visible' -and $_.CanUserResize })
    if (-not $cols.Count) { return }
    $last = $cols[-1]; $natural = $script:GridLastWidth[$Grid]
    $sum = $natural
    for ($i = 0; $i -lt $cols.Count - 1; $i++) { $sum += $cols[$i].ActualWidth }
    $fits = ($Grid.ActualWidth - 30 - $sum) -gt 0
    if ($fits -and -not $last.Width.IsStar) {
        $last.MinWidth = $natural
        $last.Width = New-Object System.Windows.Controls.DataGridLength(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
    } elseif (-not $fits -and $last.Width.IsStar) {
        $last.MinWidth = 20
        $last.Width = New-Object System.Windows.Controls.DataGridLength($natural)
    }
}

function Get-EmtSelectedRow($Grid) {
    <# The DataRow of the selected row of a grid, or $null. #>
    $item = $Grid.SelectedItem
    if ($item -is [System.Data.DataRowView]) { $item.Row } else { $null }
}

$script:LogTextWidth = 0

function Update-EmtLogWidth {
    <#
        No word wrap in the log: one log line is one line. The page has the width of the visible area,
        or the width of the longest line. A horizontal scroll bar shows only for a line that is longer than the area.
    #>
    if (-not $log -or -not $log.Document) { return }
    $view = $log.ViewportWidth
    if (-not ($view -gt 0)) { $view = 600 }
    $log.Document.PageWidth = [math]::Max($view, $script:LogTextWidth + 8)
}

# The log window keeps the last lines only: a long log in a RichTextBox makes the window slow.
# The log file keeps all lines.
$script:LogMaxLines = 2000

function Add-EmtLogLines([string[]] $Lines) {
    <# Adds lines to the log window and to the log file. One file write and one scroll for all lines. #>
    if (-not $Lines -or -not $Lines.Count) { return }
    $inlines = $script:LogPara.Inlines
    foreach ($line in $Lines) {
        $run = New-Object System.Windows.Documents.Run($line)
        $key = switch -Regex ($line) {
            '\[ERROR\]' { 'Text.Red'; break }
            '\[WARN'    { 'Text.Orange'; break }
            '\[OK'      { 'Text.Green'; break }
            default     { 'Fg' }
        }
        Set-EmtBrush $run $key ([System.Windows.Documents.TextElement]::ForegroundProperty)
        $inlines.Add($run)
        $inlines.Add((New-Object System.Windows.Documents.LineBreak))
        $w = Measure-EmtText $line $log
        if ($w -gt $script:LogTextWidth) { $script:LogTextWidth = $w; Update-EmtLogWidth }
    }
    # Two inlines for each line: the text and the line break.
    while ($inlines.Count -gt 2 * $script:LogMaxLines) { [void]$inlines.Remove($inlines.FirstInline) }
    $log.ScrollToEnd()
    try { Add-Content -Path $LogPath -Value $Lines -Encoding UTF8 } catch { }
}

#endregion

#region ---------- Language and settings (before the window) ----------

$settings = Read-EmtSettings
$available = @(Get-EmtLanguages)
$script:LanguageCodes = @($available | ForEach-Object Code)
$wanted = if ($Language) { $Language.ToLowerInvariant() }
          elseif ($settings.ContainsKey('Language') -and $settings.Language) { ([string]$settings.Language).ToLowerInvariant() }
          else { Get-EmtDefaultLanguage }
$Sync.Lang = if ($wanted -in $script:LanguageCodes) { $wanted } elseif ('en' -in $script:LanguageCodes) { 'en' } else { @($script:LanguageCodes + 'en')[0] }
$script:Theme = if ($Theme) { $Theme }
                elseif ($settings.ContainsKey('Theme') -and [string]$settings.Theme -in $script:ThemeCodes) { [string]$settings.Theme }
                else { 'Dark' }

#endregion

#region ---------- Window ----------

$xamlPath = Join-Path $GuiDir 'MainWindow.xaml'
[xml]$xamlDoc = [IO.File]::ReadAllText($xamlPath, [Text.Encoding]::UTF8)
$window = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xamlDoc))
if ($script:EmtIcon) { $window.Icon = $script:EmtIcon }
Add-EmtTheme $window

# A variable for each element with x:Name: $tbServer, $btnTrFind, ...
$xNs = 'http://schemas.microsoft.com/winfx/2006/xaml'
foreach ($n in $xamlDoc.SelectNodes('//*[@*[local-name()="Name"]]')) {
    $name = $n.GetAttribute('Name', $xNs)
    if ($name) { Set-Variable -Name $name -Value $window.FindName($name) -Scope Script }
}

$doc = New-Object System.Windows.Documents.FlowDocument
$doc.PagePadding = '0'
$script:LogPara = New-Object System.Windows.Documents.Paragraph
$script:LogPara.Margin = '0'
$doc.Blocks.Add($script:LogPara)
$log.Document = $doc
Update-EmtLogWidth
# The visible area changes with the window size and with the vertical scroll bar: the page width follows it.
$log.AddHandler([System.Windows.Controls.ScrollViewer]::ScrollChangedEvent, [System.Windows.Controls.ScrollChangedEventHandler]{
    param($s, $e)
    if ($e.ViewportWidthChange -ne 0) { Update-EmtLogWidth }
})

foreach ($l in $available) { [void]$cmbLanguage.Items.Add($l.Name) }

# The tabs and the grids of the results: key -> tab and grid. The tab title shows the row count.
$trPages = [ordered]@{ 'recipients' = @($tabTrRecipients, $gridTrRecipients); 'events' = @($tabTrEvents, $gridTrEvents)
                       'mailbox' = @($tabTrMailbox, $gridTrMailbox) }
$delPages = [ordered]@{
    'messages' = @($tabMessages, $gridMessages); 'tracking' = @($tabTracking, $gridTracking)
    'recipients' = @($tabRecipients, $gridRecipients); 'found' = @($tabFound, $gridFound)
    'missing' = @($tabMissing, $gridMissing); 'check' = @($tabCheck, $gridCheck)
}
$script:AllGrids = @($gridTrMessages, $gridTrRecipients, $gridTrEvents, $gridTrMailbox) + @($delPages.Values | ForEach-Object { $_[1] })

Set-EmtTip $tbTrSenders     'tip.senders'
Set-EmtTip $tbTrRecipient   'tip.trRecipient'
Set-EmtTip $tbTrMessageId   'tip.messageId'
Set-EmtTip $btnTrToday      'tip.today'
Set-EmtTip $btnTrYesterday  'tip.yesterday'
Set-EmtTip $btnTrWeek       'tip.week'
Set-EmtTip $btnTrDays14     'tip.days14'
Set-EmtTip $btnTrShowCriteria 'tip.show'
Set-EmtTip $btnTrCheck      'tip.checkMailbox'
Set-EmtTip $btnTrDelete     'tip.deleteThis'
Set-EmtTip $btnTrExport     'tip.export'
Set-EmtTip $tbSenders       'tip.senders'
Set-EmtTip $tbRecipient     'tip.recipient'
Set-EmtTip $tbMessageId     'tip.messageId'
Set-EmtTip $btnToday        'tip.today'
Set-EmtTip $btnYesterday    'tip.yesterday'
Set-EmtTip $btnWeek         'tip.week'
Set-EmtTip $btnDays14       'tip.days14'
Set-EmtTip $btnDelShowCriteria 'tip.show'
Set-EmtTip $chkForwards     'tip.forwards'
Set-EmtTip $chkOnly         'tip.only'
Set-EmtTip $chkAll          'tip.all'
Set-EmtTip $btnSoft         'tip.soft'
Set-EmtTip $btnHard         'tip.hard'
Set-EmtTip $btnVerify       'tip.verify'
Set-EmtTip $btnExport       'tip.export'
Set-EmtTip $btnCleanup      'tip.cleanup'
Set-EmtTip $tbTimeout       'tip.timeout'
foreach ($tb in $tbTrStartTime, $tbTrEndTime, $tbStartTime, $tbEndTime) { Set-EmtTip $tb 'tip.time' }
Set-EmtTip $cmbLanguage     'tip.language'
Set-EmtTip $cmbTheme        'tip.theme'
Set-EmtTip $btnHelp         'tip.help'

#endregion

#region ---------- State ----------

function Get-EmtSelectedMessageId {
    $row = Get-EmtSelectedRow $gridTrMessages
    if ($row) { [string]$row['MessageId'] } else { '' }
}

function Get-EmtSelectedRecipient {
    <# The selected row of the recipients grid on the tab "Message search": Address and KindCode, or $null. #>
    $row = Get-EmtSelectedRow $gridTrRecipients
    if (-not $row) { return $null }
    [pscustomobject]@{ Address = [string]$row['Recipient']; KindCode = [string]$row['KindCode'] }
}

function Get-EmtTimeout {
    <# The timeout of the operations in minutes: 5-240. A wrong value gives 60. #>
    $n = 0
    if ([int]::TryParse($tbTimeout.Text.Trim(), [ref]$n) -and $n -ge 5 -and $n -le 240) { $n } else { 60 }
}

function Set-UiState {
    $busy      = [bool]$script:Job
    $caps      = $script:Caps
    $connected = [bool]$caps
    $r         = $script:Result
    $hasFound  = [bool]($r -and @($r.Found).Count -gt 0)
    $clean     = [bool]($hasFound -and -not $r.Errors)
    $deleted   = $script:Deleted     # '', 'Soft' or 'Hard'

    $tbServer.IsEnabled    = -not $busy
    $btnConnect.IsEnabled  = -not $busy
    $btnConnect.Content    = if ($connected) { T 'conn.reconnect' } else { T 'conn.connect' }
    $cmbLanguage.IsEnabled = -not $busy

    # ---- Message search ----
    $cardTr.IsEnabled      = -not $busy
    $btnTrFind.IsEnabled   = $connected -and -not $busy -and $caps.Tracking
    $msgId = Get-EmtSelectedMessageId
    $rcpt  = Get-EmtSelectedRecipient
    $btnTrCheck.IsEnabled  = $connected -and -not $busy -and $caps.MailboxCheck -and [bool]$msgId -and
                             [bool]$rcpt -and $rcpt.KindCode -eq 'Internal'
    $btnTrDelete.IsEnabled = $connected -and -not $busy -and [bool]$msgId -and ($caps.Search -or $caps.Verify)
    $btnTrExport.IsEnabled = (-not $busy) -and [bool]$script:Trace

    # ---- Delete ----
    # The criteria can be typed before the connection. Only Find copies needs it.
    $cardDel.IsEnabled     = -not $busy
    # The quick search needs only Search-Mailbox. A soft delete always needs a Compliance Search.
    $btnSearch.IsEnabled   = $connected -and -not $busy -and ($caps.Search -or $caps.Verify)
    # After a soft delete, a hard delete is the next step. After a hard delete, only Verify is useful.
    $btnSoft.IsEnabled     = (-not $busy) -and $clean -and $connected -and (-not $deleted) -and
                             $caps.SoftDelete -and $caps.Search -and $caps.Verify
    $btnHard.IsEnabled     = (-not $busy) -and $clean -and $connected -and ($deleted -ne 'Hard') -and $caps.HardDelete
    $btnVerify.IsEnabled   = (-not $busy) -and $hasFound -and $connected -and $caps.Verify
    $btnExport.IsEnabled   = (-not $busy) -and [bool]$r
    $btnCleanup.IsEnabled  = (-not $busy) -and $connected -and $caps.Search
    $btnCancel.IsEnabled   = $busy
}

function Update-EmtTabTitles {
    <# Tab title = name and row count, e.g. "Found in mailboxes (42)". #>
    foreach ($k in $trPages.Keys) {
        $tab, $grid = $trPages[$k]
        $name = T "trtab.$k"
        $view = $grid.ItemsSource
        # The mailbox tab shows the checked mailbox: "Recipient mailbox: x@contoso.com (5)".
        if ($k -eq 'mailbox' -and $view -and $script:MailboxCheck) { $name = T 'trtab.mailboxOf' $script:MailboxCheck.Mailbox }
        $tab.Header = if ($view -is [System.Data.DataView]) { "$name ($($view.Count))" } else { $name }
    }
    foreach ($k in $delPages.Keys) {
        $tab, $grid = $delPages[$k]
        $name = T "tab.$k"
        $view = $grid.ItemsSource
        $tab.Header = if ($view -is [System.Data.DataView]) { "$name ($($view.Count))" } else { $name }
    }
}

function Update-EmtConnIndicator {
    <# The bottom line: green = connected to the server, grey = not connected. #>
    if ($script:Caps) {
        $lblConn.Text = T 'conn.connectedTo' $script:ConnectedServer
        Set-EmtBrush $dotConn 'Dot.On' ([System.Windows.Shapes.Shape]::FillProperty)
    } else {
        $lblConn.Text = T 'conn.notConnected'
        Set-EmtBrush $dotConn 'Dot.Off' ([System.Windows.Shapes.Shape]::FillProperty)
    }
}

function Update-EmtDeleteSource {
    <# The line over the delete result: data from the tracking log (green) or from the user (orange). #>
    if ($script:DelFromTrace) {
        $lblDelSource.Text = T 'del.fromTrace' $script:DelFromTrace
        Set-EmtBrush $lblDelSource 'Text.Green'
    } else {
        $lblDelSource.Text = T 'del.manual'
        Set-EmtBrush $lblDelSource 'Text.Orange'
    }
}

function Clear-EmtResultViews {
    foreach ($p in $delPages.Values) { Clear-EmtGrid $p[1] }
    Update-EmtTabTitles
    Set-EmtLabel $lblSummary (T 'sum.none') 'Fg' ''
    if ($script:Compact) { Request-EmtCompact }
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

function Get-EmtTimeFormat {
    <# The short time format of the Windows regional settings, for example H:mm or h:mm tt. #>
    $script:PcCulture.DateTimeFormat.ShortTimePattern
}

function ConvertTo-EmtTime([string] $Text) {
    <# The time of a time field, or $null for a wrong text. The format of Windows and the 24-hour format are correct. #>
    $Text = $Text.Trim()
    if (-not $Text) { return [timespan]::Zero }
    $d = [datetime]::MinValue
    foreach ($c in $script:PcCulture, [Globalization.CultureInfo]::InvariantCulture) {
        if ([datetime]::TryParse($Text, $c, [Globalization.DateTimeStyles]::NoCurrentDateDefault, [ref]$d)) { return $d.TimeOfDay }
    }
    $null
}

function Get-EmtPeriodValue($Picker, $TimeBox) {
    <# The date of a date picker and the time of its time field. #>
    # PowerShell gives a DateTime for "DateTime?": the property "Value" does not exist.
    $d = $Picker.SelectedDate
    $date = if ($null -ne $d) { ([datetime]$d).Date } else { (Get-Date).Date }
    $time = ConvertTo-EmtTime $TimeBox.Text
    if ($null -eq $time) { $time = [timespan]::Zero }
    $date + $time
}

function Set-EmtPeriodValue($Picker, $TimeBox, [datetime] $Value) {
    $Picker.SelectedDate = $Value.Date
    $TimeBox.Text = $Value.ToString((Get-EmtTimeFormat), $script:PcCulture)
    $script:LastTime[$TimeBox] = $TimeBox.Text     # a wrong time later goes back to this time
}

function Get-EmtCriteria {
    @{
        Senders         = @(Split-EmtAddressList $tbSenders.Text)
        Subject         = $tbSubject.Text.Trim()
        Recipient       = $tbRecipient.Text.Trim().ToLowerInvariant()
        MessageId       = $tbMessageId.Text.Trim()
        Attachment      = $tbAttachment.Text.Trim()
        Keywords        = $tbKeywords.Text.Trim()
        Start           = Get-EmtPeriodValue $dtStart $tbStartTime
        End             = Get-EmtPeriodEnd (Get-EmtPeriodValue $dtEnd $tbEndTime)      # the end includes its minute
        IncludeForwards = [bool]$chkForwards.IsChecked
        AllMailboxes    = [bool]$chkAll.IsChecked
        OnlyRecipient   = [bool]$chkOnly.IsChecked -and $chkOnly.IsEnabled
    }
}

function Get-EmtTraceCriteria {
    @{
        Senders   = @(Split-EmtAddressList $tbTrSenders.Text)
        Recipient = $tbTrRecipient.Text.Trim().ToLowerInvariant()
        Subject   = $tbTrSubject.Text.Trim()
        MessageId = $tbTrMessageId.Text.Trim()
        Start     = Get-EmtPeriodValue $dtTrStart $tbTrStartTime
        End       = Get-EmtPeriodEnd (Get-EmtPeriodValue $dtTrEnd $tbTrEndTime)
    }
}

#endregion

#region ---------- Trace views ----------

function Set-EmtLabel($Label, [string] $Text, [string] $ColorKey, [string] $Tip) {
    <# One text in one color, and a tooltip (empty = no tooltip). #>
    $Label.Text = $Text
    Set-EmtBrush $Label $ColorKey
    $Label.ToolTip = $(if ($Tip) { $Tip } else { $null })
}

function Set-EmtSummaryParts {
    <#
        A summary line in parts. A part is a label and its value from the template, for example "Errors: 2.".
        The part with the index $Alert is red if its value is more than 0.
        $Prefix: red texts before the parts (errors, warnings). $Lead: a usual text before the parts.
        The language files do not change: the parts come from the places {0}, {1}, ... of the template.
    #>
    param($Label, [string] $Template, [object[]] $Values, [int] $Alert = -1,
          [string[]] $Prefix = @(), [string[]] $Warn = @(), [string] $Lead = '', [string] $Tip = '')
    $Label.Inlines.Clear()
    Set-EmtBrush $Label 'Fg'
    $add = {
        param([string] $Text, [string] $ColorKey)
        if (-not $Text) { return }
        $run = New-Object System.Windows.Documents.Run($Text)
        if ($ColorKey) { Set-EmtBrush $run $ColorKey ([System.Windows.Documents.TextElement]::ForegroundProperty) }
        $Label.Inlines.Add($run)
    }
    foreach ($p in @($Prefix | Where-Object { $_ })) { & $add ($p + '   ') 'Text.Red' }
    foreach ($p in @($Warn | Where-Object { $_ })) { & $add ($p + '   ') 'Text.Orange' }
    if ($Lead) { & $add ($Lead + '   ') '' }
    $pos = 0
    # A place is {n} or {n:format}, for example {3:N1}. The text after the place up to a space belongs to the part.
    foreach ($m in [regex]::Matches($Template, '\{(\d+)(:[^}]*)?\}([^\s{]*)')) {
        $i = [int]$m.Groups[1].Value
        $before = $Template.Substring($pos, $m.Index - $pos)
        $text = $before.TrimStart()
        & $add $before.Substring(0, $before.Length - $text.Length) ''          # the spaces between the parts
        $value = if ($i -lt $Values.Count) { try { ('{0' + $m.Groups[2].Value + '}') -f $Values[$i] } catch { [string]$Values[$i] } } else { '' }
        $tail = $m.Groups[3].Value                                             # for example the dot after the value
        $red = $i -eq $Alert -and $i -lt $Values.Count -and [double]$Values[$i] -gt 0
        & $add ($text + $value + $tail) $(if ($red) { 'Text.Red' } else { '' })
        $pos = $m.Index + $m.Length
    }
    & $add $Template.Substring($pos) ''
    $full = try { $Template -f $Values } catch { $Template }
    $line = (@($Prefix | Where-Object { $_ }) + @($Warn | Where-Object { $_ }) + @($Lead | Where-Object { $_ }) + $full) -join '   '
    $Label.ToolTip = $(if ($Tip) { "$line`r`n`r`n$Tip" } else { $line })
}

function Get-EmtStatusTone([string] $Status) {
    <#
        The color of a delivery status: green = delivered, blue = sent out, red = a non-delivery,
        yellow = a delay or a forward of the mailbox, grey = other (in transit, group, not sent by a setting).
    #>
    if ($Status -in (Get-EmtErrorStatuses)) { return 'Red' }
    switch ($Status) {
        'Delivered' { 'Green' } 'External' { 'Blue' } 'Deferred' { 'Yellow' } 'Forwarded' { 'Yellow' } default { 'Grey' }
    }
}

function Get-EmtStatusParts([object[]] $Rows = @(), [switch] $Shadow) {
    <# The parts of the status of one message: one part for each status, in the order of Get-EmtMessageStatusLine. #>
    if ($Shadow) { return @(@{ Text = (T 'line.shadow'); Tone = 'Grey' }) }
    $count = Get-EmtStatusCount $Rows
    $parts = foreach ($s in Get-EmtStatusOrder) {
        $n = [int]$count[$s]
        if ($n) { @{ Text = (T "line.$s" $n); Tone = (Get-EmtStatusTone $s) } }
    }
    if ($parts) { @($parts) } else { @(@{ Text = (T 'line.none'); Tone = 'Grey' }) }
}

function Get-EmtStatusMark([object[]] $Rows = @()) {
    <#
        The mark on the left edge of a row: the worst status of the recipients.
        Red = a recipient did not get the message. Yellow = a delivery has a delay, or a mailbox forwards the message.
        Grey = a delivery is not complete.
        Blue = all recipients have the message, some recipients are outside. Green = all recipients have the message.
        A group has no own delivery: the members have their own rows.
    #>
    $count = Get-EmtStatusCount $Rows
    if (-not $count.Count) { return 'Grey' }
    $has = { param([string[]] $codes) foreach ($c in $codes) { if ($count.ContainsKey($c)) { return $true } }; $false }
    if (& $has @(Get-EmtErrorStatuses))     { return 'Red' }
    if (& $has 'Deferred', 'Forwarded')     { return 'Yellow' }
    if (& $has 'InTransit', 'Unknown')      { return 'Grey' }
    if ($count.ContainsKey('External'))     { return 'Blue' }
    'Green'
}

function Get-EmtReportOf([string] $MessageId) {
    <# The report rows (recipients) of one message. Show-EmtTraceViews makes the map. #>
    if ($script:ReportByMsg -and $script:ReportByMsg.ContainsKey($MessageId)) { $script:ReportByMsg[$MessageId].ToArray() }
}

function Test-EmtDetailLoaded([string] $MessageId) {
    <# $true if the message has all its recipients: a full message search, or the details are loaded. #>
    -not ($script:Trace -and (Get-EmtProp $script:Trace 'Partial')) -or $script:DetailLoaded.ContainsKey($MessageId)
}

function Get-EmtMessageNote($Message) {
    <# The note of a message: a shadow copy, made by Exchange, or the other recipients are not loaded yet. #>
    if (Get-EmtProp $Message 'Shadow') { T 'note.shadow' }
    elseif ($Message.ServerGenerated) { T 'note.serverGenerated' }
    elseif (-not (Test-EmtDetailLoaded $Message.MessageId)) {
        # A message to a group: the user sees the group size before the load (a large group takes time).
        $n = [int](Get-EmtProp $Message 'Expanded')
        if ($n -gt 0) { T 'note.partialGroup' $n }
        elseif (Get-EmtProp $Message 'ToGroup') { T 'note.partialToGroup' }
        else { T 'note.partial' }
    }
    else { '' }
}

function Request-EmtDetail([string] $MessageId) {
    <#
        Loads all recipients of one message (operation TraceDetail). If an operation runs, the timer starts the load later.
        The user can select other messages quickly: the timer loads only the message that is selected at that time.
    #>
    if ($script:Job) { $script:PendingDetail = $MessageId; return }
    $script:PendingDetail = ''
    $c = $script:Trace.Criteria
    Start-EmtOperation -Name 'TraceDetail' -StatusText (T 'status.detail') -Params @{
        MessageIds = @($MessageId); Start = $c.Start; End = $c.End
    } -OnSuccess {
        param($d)
        Merge-EmtDetail $d
        Show-EmtTraceDetail
    }
}

function Merge-EmtDetail($Detail) {
    <# Puts the full data of messages (Get-EmtMessageDetail) into the result of the message search, and updates their rows. #>
    $t = $script:Trace
    if (-not $t -or -not $Detail) { return }
    $ids = @($Detail.MessageIds)
    # A set and a foreach loop: Where-Object with -notin is slow for 20 000 events.
    $idSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($id in $ids) { if ($id) { [void]$idSet.Add([string]$id) } }
    $t.Events = @(@(foreach ($e in @($t.Events)) { if ($e -and -not $idSet.Contains([string]$e.MessageId)) { $e } }) + @($Detail.Events))
    $t.Report = @(@(foreach ($r in @($t.Report)) { if ($r -and -not $idSet.Contains([string]$r.MessageId)) { $r } }) + @($Detail.Report))
    $script:EventsByMsg = $null
    $byMsg = @{}
    foreach ($r in @($Detail.Report)) {
        if (-not $byMsg.ContainsKey($r.MessageId)) { $byMsg[$r.MessageId] = New-Object System.Collections.ArrayList }
        [void]$byMsg[$r.MessageId].Add($r)
    }
    $messages = @{}
    foreach ($m in @($t.Messages)) { $messages[$m.MessageId] = $m }
    $views = @{}
    foreach ($v in @($gridTrMessages.ItemsSource)) { $views[[string]$v.Row['MessageId']] = $v }
    foreach ($id in $ids) {
        # Assignment in each branch: `$x = if (...) { $list }` sends the list through the pipeline and unrolls it.
        if ($byMsg.ContainsKey($id)) { $list = $byMsg[$id] } else { $list = New-Object System.Collections.ArrayList }
        $script:ReportByMsg[$id] = $list
        $script:DetailLoaded[$id] = $true
        $rows = @($list.ToArray())
        $m = $messages[$id]
        $shadow = [bool]($m -and (Get-EmtProp $m 'Shadow'))
        if ($m) {
            $m.Recipients = $rows.Count
            $m.Problems   = @($rows | Where-Object { $_.Status -in (Get-EmtErrorStatuses) }).Count
            $m.StatusLine = Get-EmtMessageStatusLine $rows -Shadow:$shadow
        }
        # The row changes in place: the grid keeps its scroll position and the selection.
        $v = $views[$id]
        if ($v) {
            $v.Row['Recipients'] = $rows.Count
            $v.Row['StatusLine'] = Get-EmtMessageStatusLine $rows -Shadow:$shadow
            $v.Row['Parts']      = [object[]]@(Get-EmtStatusParts $rows -Shadow:$shadow)
            $v.Row['Mark']       = [string](Get-EmtStatusMark $rows)
            $v.Row['Note']       = $(if ($m) { Get-EmtMessageNote $m } else { '' })
        }
    }
}

function Export-EmtTrace {
    <# Saves the result of the message search as CSV files. The column Complete tells if a message has all recipients. #>
    $t = $script:Trace
    Export-EmtSets -Stem ('EMT_trace_{0:yyyyMMdd_HHmmss}' -f (Get-Date)) -Sets ([ordered]@{
        'messages'   = @($t.Messages | Select-Object Time, Sender, Subject, Recipients, MessageId,
                            @{ n = 'Complete'; e = { Test-EmtDetailLoaded $_.MessageId } })
        'recipients' = @($t.Report | Select-Object MessageId, Recipient, Kind, Status, Time, Server, Folder, Detail, Hint)
        'events'     = @($t.Events | Select-Object Timestamp, ServerHostname, EventId, Source, Sender,
                            @{ n = 'Recipients'; e = { $_.Recipients -join '; ' } },
                            @{ n = 'RecipientStatus'; e = { $_.RecipientStatus -join '; ' } }, MessageSubject, MessageId)
        'mailbox'    = @($(if ($script:MailboxCheck) { $script:MailboxCheck.Findings }))
    })
}

function Clear-EmtTraceViews {
    foreach ($g in $gridTrMessages, $gridTrRecipients, $gridTrEvents, $gridTrMailbox) { Clear-EmtGrid $g }
    Update-EmtTabTitles
    Set-EmtLabel $lblTrSummary (T 'trsum.none') 'Fg' ''
    if ($script:Compact) { Request-EmtCompact }
}

function Show-EmtTraceViews($t) {
    <# Fills the message grid and the summary line. No audit: it also runs after a language change. #>
    $keep = Get-EmtSelectedMessageId
    # The recipients of each message in one pass. A search in the full report for each message is slow:
    # 2000 messages and 6000 recipients take 30 seconds.
    $script:ReportByMsg = @{}
    $script:EventsByMsg = $null
    foreach ($r in @($t.Report)) {
        if (-not $script:ReportByMsg.ContainsKey($r.MessageId)) { $script:ReportByMsg[$r.MessageId] = New-Object System.Collections.ArrayList }
        [void]$script:ReportByMsg[$r.MessageId].Add($r)
    }
    $msgs = @($t.Messages | ForEach-Object {
        $rcpt = @(Get-EmtReportOf $_.MessageId)
        $shadow = [bool](Get-EmtProp $_ 'Shadow')
        [pscustomobject]@{
            Time = $_.Time; Sender = $_.Sender; Subject = $_.Subject; Recipients = $_.Recipients
            StatusLine = (Get-EmtMessageStatusLine $rcpt -Shadow:$shadow)
            Note = (Get-EmtMessageNote $_)
            MessageId = $_.MessageId; Rows = $rcpt; Shadow = $shadow
        }
    })
    # A message can have more statuses: each status has its color. The mark: red = an error, yellow = a delay.
    $script:Loading = $true
    try {
        Set-EmtGridData $gridTrMessages $msgs ([ordered]@{
            Time = 'col.time'; Sender = 'col.sender'; Subject = 'col.subject'; Recipients = 'col.recipientCount'
            StatusLine = 'col.statusLine'; Note = 'col.note'; MessageId = 'col.messageId'
        }) -PartsColumn 'StatusLine' -Parts { param($m) Get-EmtStatusParts $m.Rows -Shadow:$m.Shadow } -Mark { param($m) Get-EmtStatusMark $m.Rows }
    } finally { $script:Loading = $false }

    $report = @($t.Report)
    # One count of the statuses, then sums: not one pass over the report for each number.
    $byStatus = Get-EmtStatusCount $report
    $count = { param($codes) $n = 0; foreach ($c in $codes) { $n += [int]$byStatus[$c] }; $n }
    $failed = & $count @(Get-EmtErrorStatuses)
    $prefix = if (@($t.TrackingErrors).Count) { T 'sum.errors' } else { '' }
    if ($msgs.Count) {
        # Only the number of errors is red. The other parts have the usual text color.
        $values = @((Get-EmtMessageCountText $t.Messages), $report.Count, (& $count @('Delivered')), (& $count @('External')),
                    $failed, (& $count @('Deferred', 'InTransit', 'Unknown')), $t.Seconds)
        # A partial search counts the searched recipient only: the line starts with "Recipient x:".
        $lead = if (Get-EmtProp $t 'Partial') { T 'trsum.partial' $t.Criteria.Recipient } else { '' }
        Set-EmtSummaryParts -Label $lblTrSummary -Template (T 'trsum.text') -Values $values -Alert 4 -Prefix $prefix -Lead $lead
    } else {
        Set-EmtLabel $lblTrSummary (T 'trsum.empty') 'Text.Orange' (@($t.NoResultHints) -join "`r`n")
        if ($prefix) { Set-EmtLabel $lblTrSummary ($prefix + '   ' + $lblTrSummary.Text) 'Text.Red' $lblTrSummary.ToolTip }
    }
    # Keep the selected message after a language change.
    $script:Loading = $true
    try {
        foreach ($v in $gridTrMessages.ItemsSource) {
            if ($keep -and [string]$v.Row['MessageId'] -eq $keep) { $gridTrMessages.SelectedItem = $v; break }
        }
    } finally { $script:Loading = $false }
    Show-EmtTraceDetail
}

function Show-EmtTraceDetail {
    <# The recipients, the events and the mailbox findings of the selected message. #>
    $t = $script:Trace
    $msgId = Get-EmtSelectedMessageId
    if (-not $t -or -not $msgId) {
        foreach ($g in $gridTrRecipients, $gridTrEvents, $gridTrMailbox) { Clear-EmtGrid $g }
        Update-EmtTabTitles
        Set-UiState
        return
    }
    # A partial message search: the other recipients of this message are not loaded yet.
    if (-not (Test-EmtDetailLoaded $msgId)) {
        foreach ($g in $gridTrRecipients, $gridTrEvents, $gridTrMailbox) { Clear-EmtGrid $g }
        Update-EmtTabTitles
        Request-EmtDetail $msgId
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
    }) -PartsColumn 'Status' -Parts { param($r) @{ Text = $r.Status; Tone = (Get-EmtStatusTone $r.StatusCode) } } `
       -Mark { param($r) Get-EmtStatusTone $r.StatusCode }
    # The events of each message in one pass, once for each result: a click does not read all events.
    if (-not $script:EventsByMsg) {
        $script:EventsByMsg = @{}
        foreach ($e in @($t.Events)) {
            if (-not $e -or -not $e.MessageId) { continue }
            $k = [string]$e.MessageId
            if (-not $script:EventsByMsg.ContainsKey($k)) { $script:EventsByMsg[$k] = New-Object System.Collections.ArrayList }
            [void]$script:EventsByMsg[$k].Add($e)
        }
    }
    $events = @(if ($script:EventsByMsg.ContainsKey($msgId)) { $script:EventsByMsg[$msgId].ToArray() })
    Set-EmtGridData $gridTrEvents $events ([ordered]@{
        Timestamp = 'col.time'; EventId = 'col.event'; Source = 'col.source'; ServerHostname = 'col.server'
        Recipients = 'col.recipients'; RecipientStatus = 'col.detail'; Sender = 'col.sender'
    }) -Mark {
        param($e)
        if ([string]$e.EventId -in 'FAIL', 'DROP', 'POISONMESSAGE') { 'Red' } elseif ([string]$e.EventId -eq 'DEFER') { 'Yellow' }
    }
    $mc = $script:MailboxCheck
    if ($mc -and $mc.MessageId -eq $msgId) { Show-EmtMailboxView $mc } else { Clear-EmtGrid $gridTrMailbox }
    Update-EmtTabTitles
    Set-UiState
}

function Show-EmtMailboxView($mc) {
    $findings = if (Get-EmtProp $mc 'Data') { @(Get-EmtMailboxFindings -Data $mc.Data -Sender $mc.Sender -Subject $mc.Subject) } else { @($mc.Findings) }
    $rows = @($findings | ForEach-Object {
        [pscustomobject]@{ Severity = (T "sev.$($_.Severity)"); Area = (T "area.$($_.Area)"); Text = $_.Text; SeverityCode = $_.Severity }
    })
    Set-EmtGridData $gridTrMailbox $rows ([ordered]@{
        Severity = 'col.severity'; Area = 'col.area'; Text = 'col.finding'; SeverityCode = ''
    }) -Mark { param($f) switch ([string]$f.SeverityCode) { 'Alert' { 'Red' } 'Warn' { 'Yellow' } } }
}

function Show-EmtTraceResult($t) {
    $script:Trace = $t
    $script:EventsByMsg = $null
    $script:MailboxCheck = $null
    $script:DetailLoaded = @{}
    $script:PendingDetail = ''
    Show-EmtTraceViews $t
    if (-not @($t.Messages).Count) {
        Write-EmtLog (T 'log.traceEmpty') 'Warn'
        foreach ($h in @($t.NoResultHints)) { Write-EmtLog ('  - ' + $h) 'Warn' }
    }
    $first = @($gridTrMessages.ItemsSource)[0]
    if ($first) { $gridTrMessages.SelectedItem = $first }
    Request-EmtCompact -AfterSearch
    $c = $t.Criteria
    Write-EmtAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = 'Trace'
        Query = "Senders=$(@($c.Senders) -join ' '); Recipient=$($c.Recipient); Subject=$($c.Subject); MessageId=$($c.MessageId); Start=$($c.Start); End=$($c.End)"
        Items = @($t.Messages).Count; Mailboxes = @($t.Report).Count
        Result = $(if (@($t.TrackingErrors).Count) { 'Partial' } else { 'OK' })
        Details = "Seconds=$($t.Seconds)"
    }
}

#endregion

#region ---------- Delete views ----------

function Show-EmtResultViews($r) {
    <# Fills the grids and the summary line of the tab "Delete". No audit: it also runs after a language change. #>
    $msgs = @($r.Messages | ForEach-Object {
        [pscustomobject]@{
            Time = $_.Time; Sender = $_.Sender; Subject = $_.Subject; Recipients = $_.Recipients
            Note = $(if ($_.ServerGenerated) { T 'note.serverGenerated' } else { '' }); MessageId = $_.MessageId
        }
    })
    # A message that Exchange made itself (a notification) has a yellow mark.
    Set-EmtGridData $gridMessages $msgs ([ordered]@{
        Time = 'col.time'; Sender = 'col.sender'; Subject = 'col.subject'; Recipients = 'col.recipientCount'
        Note = 'col.note'; MessageId = 'col.messageId' }) -Mark { param($m) if ($m.Note) { 'Yellow' } }
    Set-EmtGridData $gridTracking @($r.Events) ([ordered]@{
        Timestamp = 'col.time'; ServerHostname = 'col.server'; EventId = 'col.event'; Source = 'col.source'
        Sender = 'col.sender'; Recipients = 'col.recipients'; MessageSubject = 'col.subject'; MessageId = 'col.messageId' })
    # An external recipient has a red mark: EMT cannot delete the message outside the organization.
    $rcpt = @($r.Recipients | ForEach-Object {
        [pscustomobject]@{ Address = $_.Address; Kind = (T "kind.$($_.Kind)"); Mailbox = $_.Mailbox; Type = $_.Type; KindCode = $_.Kind }
    } | Sort-Object Kind, Address)
    Set-EmtGridData $gridRecipients $rcpt ([ordered]@{
        Address = 'col.address'; Kind = 'col.kind'; Mailbox = 'col.mailbox'; Type = 'col.exchangeType'; KindCode = ''
    }) -Mark { param($x) if ([string]$x.KindCode -eq 'External') { 'Red' } }
    Set-EmtGridData $gridFound @($r.Found) ([ordered]@{ Mailbox = 'col.mailbox'; Items = 'col.items'; SizeKB = 'col.sizeKB' })
    $miss = @($r.Missing | ForEach-Object { [pscustomobject]@{ Mailbox = $_; Reason = (T 'missing.reason') } })
    Set-EmtGridData $gridMissing $miss ([ordered]@{ Mailbox = 'col.mailbox'; Reason = 'col.reason' })

    $unindexedText = if ($r.Unindexed -lt 0) { T 'sum.na' } else { [string]$r.Unindexed }
    $methodText    = if ($r.Method -eq 'SearchMailbox') { T 'sum.quick' } else { 'Compliance Search' }
    $scopeText     = if (@($r.Locations) -contains 'All') { T 'common.all' } else { [string]@($r.Locations).Count }
    $values = @($r.Items, @($r.Found).Count, $scopeText, $r.SizeKB, @($r.External).Count, $unindexedText, $methodText, $r.Seconds)
    $prefix = @()
    if ($r.Errors) { $prefix += (T 'sum.errors') }
    if ($r.PSObject.Properties['Unconfirmed'] -and $r.Unconfirmed) { $prefix += (T 'sum.unconfirmed') }
    $lead = if (-not @($r.Found).Count -and -not $r.Errors) { T 'sum.nothing' } else { '' }
    # More items than messages in a mailbox: the query can also find other messages (orange warning).
    $extra = @(Get-EmtExtraItemMailboxes -Found @($r.Found) -Messages @($r.Messages).Count)
    $warn = if ($extra.Count) { T 'sum.extraItems' $extra.Count @($r.Messages).Count } else { '' }
    Set-EmtSummaryParts -Label $lblSummary -Template (T 'sum.text') -Values $values -Prefix $prefix -Warn $warn -Lead $lead -Tip ([string]$r.Errors)
    Update-EmtTabTitles
}

function Show-EmtSearchResult($r) {
    $script:Result = $r
    $script:Check  = $null
    Clear-EmtGrid $gridCheck
    Show-EmtResultViews $r
    $tabs.SelectedItem = $tabFound
    Request-EmtCompact -AfterSearch

    Write-EmtAudit -Path $AuditPath -Record @{
        Server = $script:ConnectedServer; Action = 'Search'; SearchName = $r.SearchName; Query = $r.Query
        Mailboxes = @($r.Found).Count; Items = $r.Items
        Result = $(if ($r.Errors) { 'Errors' } else { 'OK' })
        Details = "Method=$($r.Method); Seconds=$($r.Seconds); FromTrace=$($script:DelFromTrace); External=$(@($r.External | ForEach-Object Address) -join ' '); Unindexed=$($r.Unindexed)"
    }
}

function Show-EmtCheckView([object[]] $Check) {
    # Red = the message is still in the visible folders. Yellow = the check failed. Green = no copy in the visible folders.
    Set-EmtGridData $gridCheck @($Check) ([ordered]@{
        Mailbox = 'col.mailbox'; Visible = 'col.visible'; Total = 'col.total'; Note = 'col.note'
    }) -Mark { param($c) if ([long]$c.Visible -lt 0) { 'Yellow' } elseif ([long]$c.Visible -gt 0) { 'Red' } else { 'Green' } }
    Update-EmtTabTitles
}

function Show-EmtCheckResult([object[]] $Check, [string] $Action, $DeleteInfo) {
    $script:Check = @($Check)
    Show-EmtCheckView $script:Check
    $tabs.SelectedItem = $tabCheck
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
        [void](Show-EmtMessage -Text (T 'resolve.notFound' ($notFound -join ', ')) -Title (T 'resolve.title') -Kind Warn)
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
    # A search by a recipient is partial: the other recipients of a message load when the user selects it.
    Start-EmtOperation -Name 'Trace' -Params @{ Criteria = $crit; Partial = [bool]$crit.Recipient } -StatusText (T 'status.trace') -OnSuccess {
        param($t) Show-EmtTraceResult $t
    }
}

function Start-EmtSearch {
    $crit = Get-EmtCriteria
    Save-EmtSettings
    Clear-EmtResults
    $fmt = T 'fmt.date'
    Write-EmtLog (T 'log.search' ($crit.Senders -join ', ') $crit.Subject $crit.Start.ToString($fmt) $crit.End.ToString($fmt))
    Start-EmtOperation -Name 'Search' -Params @{ Criteria = $crit; TimeoutMinutes = (Get-EmtTimeout) } `
        -StatusText (T 'status.search') -OnSuccess { param($r) Show-EmtSearchResult $r }
}

function Open-EmtDeleteForMessage {
    <#
        Fills the tab "Delete" with the data of the selected found message and opens the tab.
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
        $chkOnly.IsChecked = $false; $chkAll.IsChecked = $false
        # The period: from some minutes before the message to 23:59 of its day. The tracking log uses the exact time.
        # The query (KQL) uses whole days: this period gives one day, not the full period of the message search.
        $sent = [datetime]$m.Time
        $from = $sent.AddMinutes(-5); $from = $from.AddSeconds(-$from.Second).AddMilliseconds(-$from.Millisecond)
        Set-EmtPeriodValue $dtStart $tbStartTime $from
        Set-EmtPeriodValue $dtEnd $tbEndTime $sent.Date.AddDays(1).AddMinutes(-1)    # 23:59 of the same day
    } finally { $script:Loading = $false }
    Clear-EmtResults
    $script:DelFromTrace = $m.MessageId
    Update-EmtDeleteSource
    $mainTabs.SelectedItem = $pgDelete
    Write-EmtLog (T 'log.delPrefilled' $m.MessageId)
    Set-UiState
}

function Set-EmtPeriod($Start, $StartTime, $End, $EndTime, [int] $StartDays, [int] $EndDays) {
    <#
        A quick period: today, yesterday, the last 7 days or the last 14 days. The start is 0:00, the end is 23:59 of the last day
        (the end includes its minute). $EndDays is the day after the last day: 1 = today is the last day.
    #>
    Set-EmtPeriodValue $Start $StartTime (Get-Date).Date.AddDays($StartDays)
    Set-EmtPeriodValue $End $EndTime (Get-Date).Date.AddDays($EndDays).AddMinutes(-1)
}

#endregion

#region ---------- Event handlers ----------

# An error in an event handler must not close the window. The dispatcher gives the error here.
$window.Dispatcher.Add_UnhandledException({
    param($s, $e)
    $e.Handled = $true
    $msg = $e.Exception.Message
    try { Write-EmtLog (T 'log.handlerFailed' $msg) 'Error' } catch { }
    try { [void](Show-EmtMessage -Text $msg -Title (T 'dlg.error') -Kind Error) } catch { }
})

$btnConnect.Add_Click({
    $srv = $tbServer.Text.Trim()
    if (-not $srv) { [void](Show-EmtMessage -Text (T 'conn.enterServer') -Kind Warn); return }
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

# The timeout: only digits. A wrong value goes back to the last good value.
$tbTimeout.Add_PreviewTextInput({ param($s, $e) if ($e.Text -notmatch '^\d+$') { $e.Handled = $true } })
$tbTimeout.Add_LostFocus({
    $n = 0
    if (-not ([int]::TryParse($tbTimeout.Text.Trim(), [ref]$n) -and $n -ge 5 -and $n -le 240)) {
        Write-EmtLog (T 'conn.timeoutBad') 'Warn'
        $tbTimeout.Text = [string]$script:LastTimeout
    } else {
        $script:LastTimeout = $n
    }
})

# A time field: a wrong time goes back to the last good time.
foreach ($tb in $tbTrStartTime, $tbTrEndTime, $tbStartTime, $tbEndTime) {
    $tb.Add_LostFocus({
        param($s, $e)
        $time = ConvertTo-EmtTime $s.Text
        if ($null -eq $time) {
            Write-EmtLog (T 'crit.timeBad') 'Warn'
            $s.Text = $(if ($script:LastTime.ContainsKey($s)) { $script:LastTime[$s] } else { ([datetime]::Today).ToString((Get-EmtTimeFormat), $script:PcCulture) })
        } else {
            $s.Text = ([datetime]::Today + $time).ToString((Get-EmtTimeFormat), $script:PcCulture)
            $script:LastTime[$s] = $s.Text
        }
    })
}

$cmbLanguage.Add_SelectionChanged({
    if ($script:Loading -or $cmbLanguage.SelectedIndex -lt 0) { return }
    $code = $script:LanguageCodes[$cmbLanguage.SelectedIndex]
    if ($code -eq $Sync.Lang) { return }
    $Sync.Lang = $code
    Update-EmtLanguage
    Save-EmtSettings
    Write-EmtLog (T 'log.language')
})

# Help: the button "?" and the key F1 open the user guide (help\<language>.html) in the browser.
$btnHelp.Add_Click({ Open-EmtHelp })
$window.Add_PreviewKeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::F1) { $e.Handled = $true; Open-EmtHelp } })

$cmbTheme.Add_SelectionChanged({
    if ($script:Loading -or $cmbTheme.SelectedIndex -lt 0) { return }
    $name = $script:ThemeCodes[$cmbTheme.SelectedIndex]
    if ($name -eq $script:Theme) { return }
    Set-EmtTheme $name
    Save-EmtSettings
})

# Enter starts the search of the open tab. The event also comes from the tables and lists inside the tab: examine the source.
$mainTabs.Add_SelectionChanged({
    param($s, $e)
    if ($e.OriginalSource -ne $mainTabs) { return }
    $onTrace = $mainTabs.SelectedItem -eq $pgTrace
    $btnTrFind.IsDefault = $onTrace
    $btnSearch.IsDefault = -not $onTrace
})

# ---- Message search ----

$btnTrToday.Add_Click({     Set-EmtPeriod $dtTrStart $tbTrStartTime $dtTrEnd $tbTrEndTime 0 1 })
$btnTrYesterday.Add_Click({ Set-EmtPeriod $dtTrStart $tbTrStartTime $dtTrEnd $tbTrEndTime -1 0 })
$btnTrWeek.Add_Click({      Set-EmtPeriod $dtTrStart $tbTrStartTime $dtTrEnd $tbTrEndTime -6 1 })
$btnTrDays14.Add_Click({    Set-EmtPeriod $dtTrStart $tbTrStartTime $dtTrEnd $tbTrEndTime -13 1 })

$btnTrFind.Add_Click({
    $crit = Get-EmtTraceCriteria
    $problems = @(Test-EmtCriteria $crit -ForTrace)
    if ($problems.Count) {
        [void](Show-EmtMessage -Text ($problems -join "`r`n") -Title (T 'crit.checkTitle') -Kind Warn)
        return
    }
    Invoke-EmtWithResolve -Criteria $crit -SendersBox $tbTrSenders -RecipientBox $tbTrRecipient -Then { Start-EmtTrace }
})

$gridTrMessages.Add_SelectionChanged({ if (-not $script:Loading) { Show-EmtTraceDetail } })
$gridTrMessages.Add_MouseDoubleClick({ if ($gridTrMessages.SelectedItem) { $trDetail.SelectedItem = $tabTrEvents } })
$gridTrRecipients.Add_SelectionChanged({ Set-UiState })
# A grid on a hidden tab gets the column widths when it becomes visible.
foreach ($g in $script:AllGrids) {
    $g.Add_IsVisibleChanged({
        [void]$window.Dispatcher.BeginInvoke([Action]{ Resize-EmtGrids }, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
    })
}

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
        $trDetail.SelectedItem = $tabTrMailbox
    }
})

$btnTrDelete.Add_Click({ Open-EmtDeleteForMessage })

$btnTrExport.Add_Click({
    # A partial message search: the user can load all recipients of all messages before the export.
    $t = $script:Trace
    $missing = @($t.Messages | Where-Object { -not (Test-EmtDetailLoaded $_.MessageId) } | ForEach-Object { [string]$_.MessageId })
    $total = @($t.Messages).Count
    if ($missing.Count -and (Show-EmtMessage -Text (T 'dlg.exportPartial' ($total - $missing.Count) $total) -Kind Question)) {
        Start-EmtOperation -Name 'TraceDetail' -StatusText (T 'status.detail') -Params @{
            MessageIds = $missing; Start = $t.Criteria.Start; End = $t.Criteria.End
        } -OnSuccess { param($d) Merge-EmtDetail $d; Show-EmtTraceDetail; Export-EmtTrace }
        return
    }
    Export-EmtTrace
})

# ---- Delete ----

$btnToday.Add_Click({     Set-EmtPeriod $dtStart $tbStartTime $dtEnd $tbEndTime 0 1 })
$btnYesterday.Add_Click({ Set-EmtPeriod $dtStart $tbStartTime $dtEnd $tbEndTime -1 0 })
$btnWeek.Add_Click({      Set-EmtPeriod $dtStart $tbStartTime $dtEnd $tbEndTime -6 1 })
$btnDays14.Add_Click({    Set-EmtPeriod $dtStart $tbStartTime $dtEnd $tbEndTime -13 1 })

$btnSearch.Add_Click({
    $crit = Get-EmtCriteria
    $problems = @(Test-EmtCriteria $crit)
    if ($problems.Count) {
        [void](Show-EmtMessage -Text ($problems -join "`r`n") -Title (T 'crit.checkTitle') -Kind Warn)
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
        SearchName = $r.SearchName; Mailboxes = $mbx; Query = $r.Query; TimeoutMinutes = (Get-EmtTimeout)
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
    $days = Show-EmtInput -Title (T 'dlg.cleanupTitle') -Prompt (T 'dlg.cleanupPrompt') -Value '7'
    if (-not $days) { return }
    if ($days -notmatch '^\d+$') { Write-EmtLog (T 'log.badDays') 'Warn'; return }
    Start-EmtOperation -Name 'ListOld' -Params @{ Days = [int]$days } -StatusText (T 'status.listOld') -OnSuccess {
        param($list)
        $list = @($list | Where-Object { $_ })
        if (-not $list.Count) { Write-EmtLog (T 'log.noOldSearches') 'Ok'; return }
        $names = @($list | ForEach-Object Name)
        $text = (T 'dlg.cleanupList' $names.Count) + "`r`n`r`n" + ($names -join "`r`n")
        if (-not (Show-EmtConfirm -Title (T 'dlg.cleanupTitle') -Text $text -OkText (T 'dlg.cleanupOk'))) { return }
        # The removal starts in the timer: an operation cannot start inside the result handler of another operation.
        $script:PendingCleanup = $names
    }
})

# A change of a delete criterion makes the result old.
foreach ($c in $tbSenders, $tbSubject, $tbRecipient, $tbMessageId, $tbAttachment, $tbKeywords) {
    $c.Add_TextChanged({ Invoke-EmtInvalidate })
}
foreach ($d in $dtStart, $dtEnd) { $d.Add_SelectedDateChanged({ Invoke-EmtInvalidate }) }
foreach ($tb in $tbStartTime, $tbEndTime) { $tb.Add_TextChanged({ Invoke-EmtInvalidate }) }
$chkForwards.Add_Checked({ Invoke-EmtInvalidate })
$chkForwards.Add_Unchecked({ Invoke-EmtInvalidate })
# "All mailboxes" and "Only the recipient mailbox" exclude each other.
$chkAll.Add_Checked({ $chkOnly.IsChecked = $false; Invoke-EmtInvalidate })
$chkAll.Add_Unchecked({ Invoke-EmtInvalidate })
$chkOnly.Add_Checked({ $chkAll.IsChecked = $false; Invoke-EmtInvalidate })
$chkOnly.Add_Unchecked({ Invoke-EmtInvalidate })
$tbRecipient.Add_TextChanged({ $chkOnly.IsEnabled = [bool]$tbRecipient.Text.Trim() })

# ---- The mark of the quick period ----
# The button of the current period has a green border and a light green background, also after a search.
# The mark comes from the period fields: a manual change removes it. Midnight changes "today": the timer examines it.
$script:PeriodSets = @(
    @{ Start = $dtTrStart; StartTime = $tbTrStartTime; End = $dtTrEnd; EndTime = $tbTrEndTime; Mark = $null
       Buttons = @{ Today = $btnTrToday; Yesterday = $btnTrYesterday; Week = $btnTrWeek; Days14 = $btnTrDays14 } }
    @{ Start = $dtStart; StartTime = $tbStartTime; End = $dtEnd; EndTime = $tbEndTime; Mark = $null
       Buttons = @{ Today = $btnToday; Yesterday = $btnYesterday; Week = $btnWeek; Days14 = $btnDays14 } }
)
$script:PeriodMarkDay = (Get-Date).Date

function Update-EmtPeriodMarks {
    foreach ($set in $script:PeriodSets) {
        $k = Get-EmtQuickPeriod -Start (Get-EmtPeriodValue $set.Start $set.StartTime) -End (Get-EmtPeriodValue $set.End $set.EndTime)
        if ($set.Mark -eq $k) { continue }
        $set.Mark = $k
        foreach ($name in $set.Buttons.Keys) {
            $b = $set.Buttons[$name]
            if ($name -eq $k) {
                Set-EmtBrush $b 'Accent.Border' ([System.Windows.Controls.Control]::BorderBrushProperty)
                Set-EmtBrush $b 'Accent.Soft' ([System.Windows.Controls.Control]::BackgroundProperty)
            } else {
                $b.ClearValue([System.Windows.Controls.Control]::BorderBrushProperty)
                $b.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
            }
        }
    }
}
foreach ($set in $script:PeriodSets) {
    foreach ($d in $set.Start, $set.End) { $d.Add_SelectedDateChanged({ Update-EmtPeriodMarks }) }
    foreach ($tb in $set.StartTime, $set.EndTime) { $tb.Add_TextChanged({ Update-EmtPeriodMarks }) }
}

# ---- A small window: the criteria card becomes one line ----
# Only if the result has too little space (a small screen). After a search, one line with the criteria
# and "Show criteria" replace the card. A larger window shows the card again.
$script:Compact = @(
    @{ Card = $cardTr; Title = $lblTrTitle; Criteria = $trCriteria; Bar = $trCompact; Text = $lblTrCompact
       Area = $gridTrMessages.Parent; Need = 300; On = $false; Saved = 0; UserOpened = $false
       HasResult = { [bool]$script:Trace -and [bool]@($script:Trace.Messages).Count }
       Fields = { @(@('crit.senders', $tbTrSenders.Text), @('crit.recipient', $tbTrRecipient.Text), @('crit.subject', $tbTrSubject.Text),
                    @('crit.messageId', $tbTrMessageId.Text)) }
       Options = { @() }
       Period = @($dtTrStart, $tbTrStartTime, $dtTrEnd, $tbTrEndTime) }
    @{ Card = $cardDel; Title = $lblDelTitle; Criteria = $delCriteria; Bar = $delCompact; Text = $lblDelCompact
       Area = $tabs; Need = 220; On = $false; Saved = 0; UserOpened = $false
       HasResult = { [bool]$script:Result }
       Fields = { @(@('crit.senders', $tbSenders.Text), @('crit.recipient', $tbRecipient.Text), @('crit.subject', $tbSubject.Text),
                    @('crit.messageId', $tbMessageId.Text), @('crit.attachment', $tbAttachment.Text), @('crit.keywords', $tbKeywords.Text)) }
       Options = { @(foreach ($x in $chkForwards, $chkOnly, $chkAll) { if ($x.IsChecked -and $x.IsEnabled) { [string]$x.Content } }) }
       Period = @($dtStart, $tbStartTime, $dtEnd, $tbEndTime) }
)

function Get-EmtCompactText($c) {
    <# The criteria in one line: only the fields with a value, then the period. #>
    $parts = @(foreach ($f in (& $c.Fields)) { if ([string]$f[1].Trim()) { "$(T $f[0]) $($f[1].Trim())" } })
    $p = $c.Period
    $fmt = $script:PcCulture.DateTimeFormat.ShortDatePattern
    $s = (Get-EmtPeriodValue $p[0] $p[1]).ToString($fmt, $script:PcCulture) + ' ' + $p[1].Text
    $e = (Get-EmtPeriodValue $p[2] $p[3]).ToString($fmt, $script:PcCulture) + ' ' + $p[3].Text
    $parts += "$(T 'crit.period') $s $([char]0x2014) $e"
    $parts += @(& $c.Options)
    $parts -join '    '
}

function Set-EmtCompact($c, [bool] $On) {
    if ($c.On -eq $On) { return }
    if ($On) { $c.Saved = $c.Card.ActualHeight }
    $c.On = $On
    $full = if ($On) { 'Collapsed' } else { 'Visible' }
    $c.Title.Visibility = $full; $c.Criteria.Visibility = $full
    $c.Bar.Visibility = if ($On) { 'Visible' } else { 'Collapsed' }
    if ($On) { $c.Text.Text = Get-EmtCompactText $c; $c.Text.ToolTip = $c.Text.Text }
}

function Update-EmtCompact([switch] $AfterSearch) {
    <#
        The card becomes one line if the result area is lower than Need (with the full card).
        It becomes full again if the window is large enough, or if the result is gone.
        After "Show criteria", the card stays full until the next search.
    #>
    if (-not $script:Compact) { return }
    foreach ($c in $script:Compact) {
        if (-not (& $c.HasResult)) { Set-EmtCompact $c $false; $c.UserOpened = $false; continue }
        if ($AfterSearch) { $c.UserOpened = $false }
        $h = $c.Area.ActualHeight
        if (-not ($h -gt 0)) { continue }              # the tab is not visible: no size
        if ($c.On) {
            $full = $h - ($c.Saved - $c.Card.ActualHeight)
            if ($full -ge $c.Need) { Set-EmtCompact $c $false }
        } elseif (-not $c.UserOpened -and $h -lt $c.Need) {
            Set-EmtCompact $c $true
        }
    }
}

$script:CompactAfterSearch = $false
function Request-EmtCompact([switch] $AfterSearch) {
    <# After the layout: the heights are then correct. #>
    if ($AfterSearch) { $script:CompactAfterSearch = $true }
    [void]$window.Dispatcher.BeginInvoke([Action]{
        $a = $script:CompactAfterSearch; $script:CompactAfterSearch = $false
        Update-EmtCompact -AfterSearch:$a
    }, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
}

$btnTrShowCriteria.Add_Click({ $c = $script:Compact[0]; $c.UserOpened = $true; Set-EmtCompact $c $false })
$btnDelShowCriteria.Add_Click({ $c = $script:Compact[1]; $c.UserOpened = $true; Set-EmtCompact $c $false })
$window.Add_SizeChanged({ Request-EmtCompact })
$mainTabs.Add_SelectionChanged({ param($s, $e) if ($e.OriginalSource -eq $mainTabs) { Request-EmtCompact } })

# ---- Common ----

$btnCancel.Add_Click({
    if ($script:Job) {
        $Sync.Cancel = $true
        Write-EmtLog (T 'log.cancelRequested') 'Warn'
    }
})

function Receive-EmtLog {
    <# Moves the waiting log lines from the background runspace to the log, all in one step. #>
    $lines = New-Object System.Collections.Generic.List[string]
    $line = $null
    while ($Sync.Log.TryDequeue([ref]$line)) { $lines.Add($line) }
    if ($lines.Count) { Add-EmtLogLines $lines.ToArray() }
}

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(400)
$timer.Add_Tick({
    if ((Get-Date).Date -ne $script:PeriodMarkDay) { $script:PeriodMarkDay = (Get-Date).Date; Update-EmtPeriodMarks }
    Receive-EmtLog
    if ($script:Job) {
        $progress.Value = [int]$Sync.Progress
        $st = [string]$Sync.Status
        # A long status text can be cut: the tooltip shows the full text.
        if ($lblStatus.Text -ne $st) { Set-EmtStatusText $st }
        if ($script:Job.Handle.IsCompleted) {
            # Stop the timer: a result handler can open a modal dialog (no re-entry).
            $timer.Stop()
            try {
                Complete-EmtOperation
                Receive-EmtLog
            } finally { $timer.Start() }
        }
    }
    elseif ($script:PendingDetail) {
        # Load the message that is selected now. The user can have selected another message during the operation.
        $id = Get-EmtSelectedMessageId
        $script:PendingDetail = ''
        if ($id -and -not (Test-EmtDetailLoaded $id)) { Request-EmtDetail $id }
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

$window.Add_SourceInitialized({ Set-EmtTitleBar $window })

function Update-EmtShortcut {
    <#
        The shortcut "Exchange Message Trail.lnk" with the EMT icon is in the EMT folder. It opens EMT.cmd.
        EMT makes it at the first start and after a move of the folder (the shortcut points to an earlier folder).
        The key ShortcutFolder of the settings keeps the folder of the last try. A shortcut that is removed later
        (by the user or by an antivirus) stays removed: no new try, no new antivirus alert at each start.
        src\New-EmtShortcut.ps1 makes the shortcut. The user can run it to make the shortcut again.
    #>
    $file = Join-Path $PSScriptRoot 'Exchange Message Trail.lnk'
    $cmd = Join-Path $PSScriptRoot 'EMT.cmd'
    $isGood = {
        (Test-Path -LiteralPath $file) -and
        [string]::Equals((New-Object -ComObject WScript.Shell).CreateShortcut($file).TargetPath, $cmd, [StringComparison]::OrdinalIgnoreCase)
    }
    try {
        if (& $isGood) { return }
        $s = Read-EmtSettings
        $tried = $s.ContainsKey('ShortcutFolder') -and [string]::Equals([string]$s.ShortcutFolder, $PSScriptRoot, [StringComparison]::OrdinalIgnoreCase)
        if ($tried -and -not (Test-Path -LiteralPath $file)) { return }
        $s.ShortcutFolder = $PSScriptRoot
        try { $s | ConvertTo-Json | Set-Content -Path $SettingsPath -Encoding UTF8 } catch { }
        # Some errors (for example, a script that cannot run) do not stop the call: collect them.
        $errs = @(& (Join-Path $PSScriptRoot 'src\New-EmtShortcut.ps1') 2>&1 6>$null |
                  Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
        # The result counts, not the call: examine the shortcut again.
        if (& $isGood) { Write-EmtLog (T 'log.shortcutMade' $file); return }
        $why = if ($errs.Count) { $errs[0].Exception.Message } else { $file }
        Write-EmtLog (T 'log.shortcutFailed' $why) 'Warn'
    } catch {
        Write-EmtLog (T 'log.shortcutFailed' $_.Exception.Message) 'Warn'
    }
}

$window.Add_ContentRendered({
    # Show the window in front of the other windows once. Topmost stays off: dialogs of other programs are not hidden.
    $window.Topmost = $true
    [void]$window.Activate()
    $window.Topmost = $false
    Write-EmtLog (T 'log.started' $LogPath)
    Write-EmtLog (T 'log.auditFile' $AuditPath)
    foreach ($m in $script:Migrated) { Write-EmtLog (T 'log.migrated' $m) 'Ok' }
    Update-EmtShortcut
    if ($tbServer.Text) { $btnConnect.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) }
})

$window.Add_Closing({
    param($s, $e)
    if ($script:Job) {
        if (-not (Show-EmtMessage -Text (T 'dlg.closeRunning') -Kind Question)) { $e.Cancel = $true; return }
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
if ($settings.ContainsKey('IncludeForwards')) { $chkForwards.IsChecked = [bool]$settings.IncludeForwards }
if ($settings.ContainsKey('AllMailboxes'))    { $chkAll.IsChecked = [bool]$settings.AllMailboxes }
$script:LastTimeout = 60
if ($settings.ContainsKey('TimeoutMinutes')) {
    $script:LastTimeout = [math]::Max(5, [math]::Min(240, [int]$settings.TimeoutMinutes))
}
$tbTimeout.Text = [string]$script:LastTimeout
$cmbLanguage.SelectedIndex = [array]::IndexOf($script:LanguageCodes, [string]$Sync.Lang)
# Small screens (HD 1280 x 720, or Full HD with the scale 150 %): the window must fit in the work area
# (the screen without the taskbar). The minimum size cannot be larger than the work area.
# If the default size does not fit, the window opens maximized. The log is smaller: the grids get the space.
$wa = [System.Windows.SystemParameters]::WorkArea
if ($wa.Width -gt 0 -and $wa.Height -gt 0) {
    $window.MinWidth  = [math]::Min($window.MinWidth, $wa.Width)
    $window.MinHeight = [math]::Min($window.MinHeight, $wa.Height)
    if ($window.Width -gt $wa.Width -or $window.Height -gt $wa.Height) {
        $window.Width = [math]::Min($window.Width, $wa.Width); $window.Height = [math]::Min($window.Height, $wa.Height)
        $window.WindowState = 'Maximized'
    }
    if ($wa.Height -lt 800) { $log.Parent.RowDefinitions[3].Height = New-Object System.Windows.GridLength(72) }
}
# The last window size and position. Only if the window is on the screen now (a monitor can be removed).
# The size cannot be larger than all screens together.
if ($settings.ContainsKey('WpfWindow') -and $settings.WpfWindow) {
    try {
        $wb = $settings.WpfWindow
        $vs = New-Object System.Windows.Rect([System.Windows.SystemParameters]::VirtualScreenLeft, [System.Windows.SystemParameters]::VirtualScreenTop,
                                             [System.Windows.SystemParameters]::VirtualScreenWidth, [System.Windows.SystemParameters]::VirtualScreenHeight)
        $rect = New-Object System.Windows.Rect([int]$wb.X, [int]$wb.Y,
            [math]::Min($vs.Width, [math]::Max([int]$wb.Width, $window.MinWidth)),
            [math]::Min($vs.Height, [math]::Max([int]$wb.Height, $window.MinHeight)))
        $part = [System.Windows.Rect]::Intersect($vs, $rect)
        if (-not $part.IsEmpty -and $part.Width -ge 200 -and $part.Height -ge 100) {
            $window.WindowStartupLocation = 'Manual'
            $window.Left = $rect.X; $window.Top = $rect.Y; $window.Width = $rect.Width; $window.Height = $rect.Height
            $window.WindowState = if ([bool]$wb.Maximized) { 'Maximized' } else { 'Normal' }
        }
    } catch { }
}
Set-EmtPeriod $dtTrStart $tbTrStartTime $dtTrEnd $tbTrEndTime 0 1
Set-EmtPeriod $dtStart $tbStartTime $dtEnd $tbEndTime 0 1
Update-EmtPeriodMarks
# The first day of the week comes from the regional settings of Windows, not from the interface language.
# For example: Monday in Europe, Sunday in the USA. The user can change it in the Windows settings.
foreach ($d in $dtTrStart, $dtTrEnd, $dtStart, $dtEnd) { $d.FirstDayOfWeek = $script:PcCulture.DateTimeFormat.FirstDayOfWeek }
$btnTrFind.IsDefault = $true
$script:Loading = $false

Update-EmtLanguage
$timer.Start()
try {
    [void]$window.ShowDialog()
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
    Remove-Variable EmtSync -Scope Global -ErrorAction SilentlyContinue
}

#endregion
