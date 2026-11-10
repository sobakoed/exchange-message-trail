#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Message Trail (EMT): the WPF prototype of the window. Only the tab "Message search".

.DESCRIPTION
    The prototype shows the new look: a dark or light theme, the styles of Exchange Mailbox Cleanup.
    The logic is the same as in ExchangeMessageTrail.ps1: src\EMT.Core.ps1 in a background runspace.
    The prototype has the connection, the message search, the recipients, the events and the mailbox check.
    The tab "Delete" is not in the prototype. Use ExchangeMessageTrail.ps1 to delete messages.
    The prototype keeps its settings in a separate file: settings.wpf.json.

.PARAMETER Server
    The FQDN of the Exchange server. If you do not set it, the prototype uses the last value.

.PARAMETER Language
    The language code of a file in "Lang" (for example ru or en).

.PARAMETER Theme
    Dark or Light. If you do not set it, the prototype uses the last value, or Dark.

.EXAMPLE
    .\ExchangeMessageTrail.Wpf.ps1 -Server exch01.contoso.local -Language en -Theme Dark
#>
[CmdletBinding()]
param([string] $Server, [string] $Language, [ValidateSet('', 'Dark', 'Light')] [string] $Theme = '')

$ErrorActionPreference = 'Stop'

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Start the prototype in STA mode: powershell.exe -STA -File ExchangeMessageTrail.Wpf.ps1 (or use EMT-WPF.cmd).'
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

# The console is hidden: show a start error in a window. The language files are not loaded yet: English text.
trap {
    [void][System.Windows.MessageBox]::Show(
        "Exchange Message Trail could not start.`r`n`r`n$($_.Exception.Message)`r`n`r`n$($_.InvocationInfo.PositionMessage)",
        'Exchange Message Trail', 'OK', 'Error')
    break
}

if (-not ('EmtConsole' -as [type])) {
    Add-Type -Namespace '' -Name 'EmtConsole' -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
[DllImport("kernel32.dll")] public static extern uint GetConsoleProcessList(uint[] list, uint count);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int cmd);
'@
}
if (-not ('EmtDwm' -as [type])) {
    # The color of the title bar (Windows 10 1809 and later, Windows Server 2019 and later).
    Add-Type -Namespace '' -Name 'EmtDwm' -MemberDefinition @'
[DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(System.IntPtr hwnd, int attr, ref int value, int size);
'@
}

# Hide the PowerShell console window, but only if it belongs to EMT alone (see ExchangeMessageTrail.ps1).
$consoleWindow = [EmtConsole]::GetConsoleWindow()
if ($consoleWindow -ne [IntPtr]::Zero -and [EmtConsole]::GetConsoleProcessList((New-Object 'System.UInt32[]' 4), 4) -eq 1 -and
    [Environment]::CommandLine -match '\s-File\s') {
    [void][EmtConsole]::ShowWindow($consoleWindow, 0)    # 0 = SW_HIDE
}

$CorePath = Join-Path $PSScriptRoot 'src\EMT.Core.ps1'
. $CorePath
$GuiDir = Join-Path $PSScriptRoot 'gui'

#region ---------- Paths, settings, shared state ----------

$AppDir          = Join-Path $env:USERPROFILE 'ExchangeMessageTrail'
$SettingsPath    = Join-Path $env:APPDATA 'ExchangeMessageTrail\settings.json'       # read only: the settings of EMT
$WpfSettingsPath = Join-Path $env:APPDATA 'ExchangeMessageTrail\settings.wpf.json'   # the settings of the prototype
$AuditPath       = Join-Path $AppDir 'audit.csv'
$LogPath         = Join-Path $AppDir ('logs\EMT_WPF_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
foreach ($d in (Split-Path $WpfSettingsPath), (Split-Path $LogPath)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

$Sync = [hashtable]::Synchronized(@{
    Log      = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    Progress = 0
    Status   = ''
    Cancel   = $false
    Lang     = 'en'
})
$global:EmtSync = $Sync

$script:Caps            = $null
$script:Trace           = $null
$script:MailboxCheck    = $null
$script:Job             = $null
$script:Loading         = $false
$script:ConnectedServer = ''
$script:CheckForMessage = ''
$script:ReportByMsg     = @{}     # MessageId -> the report rows of the message
$script:PcCulture       = [Globalization.CultureInfo]::CurrentCulture   # the regional settings of Windows (Set-EmtDateCulture)
$script:ResolveNext     = $null
$script:Theme           = 'Dark'
$script:LanguageCodes   = @()
$script:ThemeCodes      = @('Dark', 'Light')
$script:GridHeaders     = @{}     # grid -> ordered map "property" -> "text key" of the column header
$script:GridNeedsSize   = @{}     # grid -> true: set the column widths when the grid is visible
$script:Tips            = New-Object System.Collections.Generic.List[object]   # control + tooltip key
$script:XamlText        = @{}     # file name -> text (the dialogs load the theme again)

# The texts of the prototype. They are not in the language files: the prototype can change.
$ProtoTexts = @{
    ru = @{
        'theme' = 'Тема:'; 'theme.Dark' = 'Тёмная'; 'theme.Light' = 'Светлая'
        'title' = 'прототип WPF'
        'noDelete' = 'Вкладки «Удаление» нет в прототипе. Чтобы удалить письма, запустите EMT.cmd.'
        'ok' = 'OK'; 'yes' = 'Да'; 'no' = 'Нет'
    }
    en = @{
        'theme' = 'Theme:'; 'theme.Dark' = 'Dark'; 'theme.Light' = 'Light'
        'title' = 'WPF prototype'
        'noDelete' = 'The tab "Delete" is not in the prototype. To delete messages, start EMT.cmd.'
        'ok' = 'OK'; 'yes' = 'Yes'; 'no' = 'No'
    }
}
function PT([string] $Key) {
    $lang = if ($ProtoTexts.ContainsKey([string]$Sync.Lang)) { [string]$Sync.Lang } else { 'en' }
    $ProtoTexts[$lang][$Key]
}

function Read-EmtJson([string] $Path) {
    $s = @{}
    if (Test-Path $Path) {
        try {
            $j = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $j.PSObject.Properties) { $s[$p.Name] = $p.Value }
        } catch { }
    }
    $s
}

function Save-EmtSettings {
    # Senders and subject are not kept: they can be confidential.
    $s = [ordered]@{
        Server   = $tbServer.Text.Trim()
        Language = [string]$Sync.Lang
        Theme    = $script:Theme
    }
    # The window: the normal size and position, also when the window is maximized.
    $b = if ($window.WindowState -eq 'Normal') { New-Object System.Windows.Rect($window.Left, $window.Top, $window.ActualWidth, $window.ActualHeight) }
         else { $window.RestoreBounds }
    if ($b.Width -gt 0 -and $b.Height -gt 0) {
        $s.Window = [ordered]@{ X = [int]$b.X; Y = [int]$b.Y; Width = [int]$b.Width; Height = [int]$b.Height
                                Maximized = ($window.WindowState -eq 'Maximized') }
    }
    try { $s | ConvertTo-Json | Set-Content -Path $WpfSettingsPath -Encoding UTF8 } catch { }
}

#endregion

#region ---------- Background runspace (the same as in ExchangeMessageTrail.ps1) ----------

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
        if ($err -is [System.Management.Automation.ActionPreferenceStopException] -and $err.ErrorRecord) {
            $err = $err.ErrorRecord.Exception
        }
    }
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
    <# Sets the texts of all elements with a text key in Tag. "P:" = a text of the prototype. #>
    $key = Get-EmtTextKey $Node
    if ($key) {
        $text = if ($key.StartsWith('P:')) { PT $key.Substring(2) } else { T $key }
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
    foreach ($d in $dtTrStart, $dtTrEnd) {
        $v = $d.SelectedDate
        $d.SelectedDate = $null
        $d.SelectedDate = $v
    }
}

function Update-EmtLanguage {
    <# Applies the current language to all controls, headers and result views. #>
    $window.Title = '{0} ({1})' -f (T 'app.title'), (PT 'title')
    Set-EmtDateCulture
    Update-EmtTexts $window
    foreach ($x in $script:Tips) { $x.Control.ToolTip = T $x.Key }
    $script:Loading = $true
    try {
        $cmbTheme.Items.Clear()
        foreach ($t in $script:ThemeCodes) { [void]$cmbTheme.Items.Add((PT "theme.$t")) }
        $cmbTheme.SelectedIndex = [array]::IndexOf($script:ThemeCodes, $script:Theme)
    } finally { $script:Loading = $false }
    Update-EmtConnIndicator
    if ($script:Trace) { Show-EmtTraceViews $script:Trace } else { Clear-EmtTraceViews }
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
    $ok.Content = $(if ($Kind -eq 'Question') { PT 'yes' } else { PT 'ok' })
    $ok.IsDefault = $true; $ok.MinWidth = 96
    $ok.Add_Click({ $w.DialogResult = $true }.GetNewClosure())
    [void]$buttons.Children.Add($ok)
    if ($Kind -eq 'Question') {
        $no = New-Object System.Windows.Controls.Button
        $no.Content = PT 'no'; $no.IsCancel = $true; $no.MinWidth = 96; $no.Margin = '8,0,0,0'
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
    $answer = $w.ShowDialog()
    $picked = @(foreach ($item in $list.SelectedItems) { $Candidates[$list.Items.IndexOf($item)].Address })
    if (-not $answer -or -not $picked.Count) { return $null }
    , $picked
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
        $sample = $items | ForEach-Object { Get-EmtProp $_ $prop } | Where-Object { $_ -ne $null } | Select-Object -First 1
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
        foreach ($prop in $Columns.Keys) {
            $v = Get-EmtProp $it $prop
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
    $Grid.ItemsSource = $dt.DefaultView
    # The widths come from the content after the layout. The dispatcher calls the function after the layout.
    [void]$window.Dispatcher.BeginInvoke([Action]{ Resize-EmtGrids }, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
}

function Clear-EmtGrid($Grid) {
    $Grid.ItemsSource = $null
    $Grid.Columns.Clear()
    $script:GridNeedsSize[$Grid] = $false
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
        foreach ($c in $cols) {
            $w = [math]::Max(60, [math]::Min(450, $c.ActualWidth))
            $c.Width = New-Object System.Windows.Controls.DataGridLength($w)
        }
        $last = $cols[-1]
        $last.MinWidth = $last.ActualWidth
        $last.Width = New-Object System.Windows.Controls.DataGridLength(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
    }
}

function Get-EmtSelectedRow($Grid) {
    <# The DataRow of the selected row of a grid, or $null. #>
    $item = $Grid.SelectedItem
    if ($item -is [System.Data.DataRowView]) { $item.Row } else { $null }
}

function Add-EmtLogLine([string] $Line) {
    $run = New-Object System.Windows.Documents.Run($Line)
    $key = switch -Regex ($Line) {
        '\[ERROR\]' { 'Text.Red'; break }
        '\[WARN'    { 'Text.Orange'; break }
        '\[OK'      { 'Text.Green'; break }
        default     { 'Fg' }
    }
    Set-EmtBrush $run $key ([System.Windows.Documents.TextElement]::ForegroundProperty)
    $script:LogPara.Inlines.Add($run)
    $script:LogPara.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
    $log.ScrollToEnd()
    try { Add-Content -Path $LogPath -Value $Line -Encoding UTF8 } catch { }
}

#endregion

#region ---------- Language and settings (before the window) ----------

$settings = Read-EmtJson $SettingsPath
foreach ($kv in (Read-EmtJson $WpfSettingsPath).GetEnumerator()) { $settings[$kv.Key] = $kv.Value }
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
Add-EmtTheme $window

# A variable for each element with x:Name: $tbServer, $btnTrFind, ...
$xNs = 'http://schemas.microsoft.com/winfx/2006/xaml'
foreach ($n in $xamlDoc.SelectNodes('//*[@*[local-name()="Name"]]')) {
    $name = $n.GetAttribute('Name', $xNs)
    if ($name) { Set-Variable -Name $name -Value $window.FindName($name) -Scope Script }
}

$doc = New-Object System.Windows.Documents.FlowDocument
$doc.PagePadding = '0'
$doc.PageWidth = 4000      # no word wrap: one log line is one line
$script:LogPara = New-Object System.Windows.Documents.Paragraph
$script:LogPara.Margin = '0'
$doc.Blocks.Add($script:LogPara)
$log.Document = $doc

foreach ($l in $available) { [void]$cmbLanguage.Items.Add($l.Name) }

Set-EmtTip $tbTrSenders     'tip.senders'
Set-EmtTip $tbTrRecipient   'tip.trRecipient'
Set-EmtTip $tbTrMessageId   'tip.messageId'
Set-EmtTip $btnTrToday      'tip.today'
Set-EmtTip $btnTrYesterday  'tip.yesterday'
Set-EmtTip $btnTrWeek       'tip.week'
Set-EmtTip $btnTrCheck      'tip.checkMailbox'
Set-EmtTip $cmbLanguage     'tip.language'

#endregion

#region ---------- State ----------

function Get-EmtSelectedMessageId {
    $row = Get-EmtSelectedRow $gridTrMessages
    if ($row) { [string]$row['MessageId'] } else { '' }
}

function Get-EmtSelectedRecipient {
    $row = Get-EmtSelectedRow $gridTrRecipients
    if (-not $row) { return $null }
    [pscustomobject]@{ Address = [string]$row['Recipient']; KindCode = [string]$row['KindCode'] }
}

function Set-UiState {
    $busy      = [bool]$script:Job
    $caps      = $script:Caps
    $connected = [bool]$caps

    $tbServer.IsEnabled    = -not $busy
    $btnConnect.IsEnabled  = -not $busy
    $btnConnect.Content    = if ($connected) { T 'conn.reconnect' } else { T 'conn.connect' }
    $cmbLanguage.IsEnabled = -not $busy

    $cardTr.IsEnabled      = -not $busy
    $btnTrFind.IsEnabled   = $connected -and -not $busy -and $caps.Tracking
    $msgId = Get-EmtSelectedMessageId
    $rcpt  = Get-EmtSelectedRecipient
    $btnTrCheck.IsEnabled  = $connected -and -not $busy -and $caps.MailboxCheck -and [bool]$msgId -and
                             [bool]$rcpt -and $rcpt.KindCode -eq 'Internal'
    $btnCancel.IsEnabled   = $busy
}

function Update-EmtTabTitles {
    $pages = [ordered]@{ 'recipients' = @($tabTrRecipients, $gridTrRecipients); 'events' = @($tabTrEvents, $gridTrEvents)
                         'mailbox' = @($tabTrMailbox, $gridTrMailbox) }
    foreach ($k in $pages.Keys) {
        $tab, $grid = $pages[$k]
        $name = T "trtab.$k"
        $view = $grid.ItemsSource
        if ($k -eq 'mailbox' -and $view -and $script:MailboxCheck) { $name = T 'trtab.mailboxOf' $script:MailboxCheck.Mailbox }
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

function Get-EmtPickerDate($Picker) {
    # PowerShell gives a DateTime for "DateTime?": the property "Value" does not exist.
    $d = $Picker.SelectedDate
    if ($null -ne $d) { ([datetime]$d).Date } else { (Get-Date).Date }
}

function Get-EmtTraceCriteria {
    @{
        Senders   = @(Split-EmtAddressList $tbTrSenders.Text)
        Recipient = $tbTrRecipient.Text.Trim().ToLowerInvariant()
        Subject   = $tbTrSubject.Text.Trim()
        MessageId = $tbTrMessageId.Text.Trim()
        Start     = Get-EmtPickerDate $dtTrStart
        End       = Get-EmtPickerDate $dtTrEnd
    }
}

#endregion

#region ---------- Trace views ----------

function Set-EmtSummary([string] $Text, [string] $ColorKey, [string] $Tip) {
    $lblTrSummary.Text = $Text
    Set-EmtBrush $lblTrSummary $ColorKey
    $lblTrSummary.ToolTip = $(if ($Tip) { $Tip } else { $null })
}

function Set-EmtSummaryParts {
    <#
        The summary line in parts. A part is a label and its value from the template, for example "Errors: 2.".
        The part with the index $Alert is red if its value is more than 0. $Prefix is a red text before the parts.
        The language files do not change: the parts come from the places {0}, {1}, ... of the template.
    #>
    param([string] $Template, [object[]] $Values, [int] $Alert = -1, [string] $Prefix = '')
    $lblTrSummary.Inlines.Clear()
    Set-EmtBrush $lblTrSummary 'Fg'
    $add = {
        param([string] $Text, [string] $ColorKey)
        if (-not $Text) { return }
        $run = New-Object System.Windows.Documents.Run($Text)
        if ($ColorKey) { Set-EmtBrush $run $ColorKey ([System.Windows.Documents.TextElement]::ForegroundProperty) }
        $lblTrSummary.Inlines.Add($run)
    }
    if ($Prefix) { & $add ($Prefix + '   ') 'Text.Red' }
    $pos = 0
    foreach ($m in [regex]::Matches($Template, '\{(\d+)\}[^\s{]*')) {
        $i = [int]$m.Groups[1].Value
        $before = $Template.Substring($pos, $m.Index - $pos)
        $label = $before.TrimStart()
        & $add $before.Substring(0, $before.Length - $label.Length) ''          # the spaces between the parts
        $value = if ($i -lt $Values.Count) { '{0}' -f $Values[$i] } else { '' }
        $tail = $m.Value.Substring($m.Groups[1].Length + 2)                     # for example the dot after the value
        $red = $i -eq $Alert -and $i -lt $Values.Count -and [double]$Values[$i] -gt 0
        & $add ($label + $value + $tail) $(if ($red) { 'Text.Red' } else { '' })
        $pos = $m.Index + $m.Length
    }
    & $add $Template.Substring($pos) ''
    $full = try { $Template -f $Values } catch { $Template }
    $lblTrSummary.ToolTip = $(if ($Prefix) { "$Prefix   $full" } else { $full })
}

function Get-EmtStatusTone([string] $Status) {
    <# The color of a delivery status: green = delivered, blue = sent out, red = error, yellow = delay, grey = other. #>
    switch ($Status) {
        'Delivered' { 'Green' } 'External' { 'Blue' } 'Failed' { 'Red' } 'Dropped' { 'Red' } 'Deferred' { 'Yellow' } default { 'Grey' }
    }
}

function Get-EmtStatusParts([object[]] $Rows = @()) {
    <# The parts of the status of one message: one part for each status, in the order of Get-EmtMessageStatusLine. #>
    $order = 'Delivered', 'External', 'Failed', 'Dropped', 'Deferred', 'InTransit', 'Group', 'Unknown'
    $parts = foreach ($s in $order) {
        $n = @($Rows | Where-Object { $_.Status -eq $s }).Count
        if ($n) { @{ Text = (T "line.$s" $n); Tone = (Get-EmtStatusTone $s) } }
    }
    if ($parts) { @($parts) } else { @(@{ Text = (T 'line.none'); Tone = 'Grey' }) }
}

function Get-EmtStatusMark([object[]] $Rows = @()) {
    <#
        The mark on the left edge of a row: the worst status of the recipients.
        Red = a recipient did not get the message. Yellow = a delivery has a delay. Grey = a delivery is not complete.
        Blue = all recipients have the message, some recipients are outside. Green = all recipients have the message.
        A group has no own delivery: the members have their own rows.
    #>
    $codes = @($Rows | ForEach-Object { [string]$_.Status })
    if (-not $codes.Count) { return 'Grey' }
    if ($codes | Where-Object { $_ -in 'Failed', 'Dropped' }) { return 'Red' }
    if ($codes -contains 'Deferred') { return 'Yellow' }
    if ($codes | Where-Object { $_ -in 'InTransit', 'Unknown' }) { return 'Grey' }
    if ($codes -contains 'External') { return 'Blue' }
    'Green'
}

function Get-EmtReportOf([string] $MessageId) {
    <# The report rows (recipients) of one message. Show-EmtTraceViews makes the map. #>
    if ($script:ReportByMsg -and $script:ReportByMsg.ContainsKey($MessageId)) { $script:ReportByMsg[$MessageId].ToArray() }
}

function Clear-EmtTraceViews {
    foreach ($g in $gridTrMessages, $gridTrRecipients, $gridTrEvents, $gridTrMailbox) { Clear-EmtGrid $g }
    Update-EmtTabTitles
    Set-EmtSummary (T 'trsum.none') 'Fg' ''
}

function Show-EmtTraceViews($t) {
    <# Fills the message grid and the summary line. No audit: it also runs after a language change. #>
    $keep = Get-EmtSelectedMessageId
    # The recipients of each message in one pass. A search in the full report for each message is slow:
    # 2000 messages and 6000 recipients take 30 seconds.
    $script:ReportByMsg = @{}
    foreach ($r in @($t.Report)) {
        if (-not $script:ReportByMsg.ContainsKey($r.MessageId)) { $script:ReportByMsg[$r.MessageId] = New-Object System.Collections.ArrayList }
        [void]$script:ReportByMsg[$r.MessageId].Add($r)
    }
    $msgs = @($t.Messages | ForEach-Object {
        $rcpt = @(Get-EmtReportOf $_.MessageId)
        [pscustomobject]@{
            Time = $_.Time; Sender = $_.Sender; Subject = $_.Subject; Recipients = $_.Recipients
            StatusLine = (Get-EmtMessageStatusLine $rcpt)
            Note = $(if ($_.ServerGenerated) { T 'note.serverGenerated' } else { '' })
            MessageId = $_.MessageId; Rows = $rcpt
        }
    })
    # A message can have more statuses: each status has its color. The mark: red = an error, yellow = a delay.
    $script:Loading = $true
    try {
        Set-EmtGridData $gridTrMessages $msgs ([ordered]@{
            Time = 'col.time'; Sender = 'col.sender'; Subject = 'col.subject'; Recipients = 'col.recipientCount'
            StatusLine = 'col.statusLine'; Note = 'col.note'; MessageId = 'col.messageId'
        }) -PartsColumn 'StatusLine' -Parts { param($m) Get-EmtStatusParts $m.Rows } -Mark { param($m) Get-EmtStatusMark $m.Rows }
    } finally { $script:Loading = $false }

    $report = @($t.Report)
    $count = { param($codes) @($report | Where-Object { $_.Status -in $codes }).Count }
    $failed = & $count @('Failed', 'Dropped')
    $prefix = if (@($t.TrackingErrors).Count) { T 'sum.errors' } else { '' }
    if ($msgs.Count) {
        # Only the number of errors is red. The other parts have the usual text color.
        $values = @($msgs.Count, $report.Count, (& $count @('Delivered')), (& $count @('External')),
                    $failed, (& $count @('Deferred', 'InTransit', 'Unknown')), $t.Seconds)
        Set-EmtSummaryParts -Template (T 'trsum.text') -Values $values -Alert 4 -Prefix $prefix
    } else {
        Set-EmtSummary (T 'trsum.empty') 'Text.Orange' (@($t.NoResultHints) -join "`r`n")
        if ($prefix) { Set-EmtSummary ($prefix + '   ' + $lblTrSummary.Text) 'Text.Red' $lblTrSummary.ToolTip }
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
    $rows = @(Get-EmtReportOf $msgId | ForEach-Object {
        [pscustomobject]@{
            Recipient = $_.Recipient; Status = (T "st.$($_.Status)")
            Kind = $(if ($_.Kind) { T "kind.$($_.Kind)" } else { '' }); Time = $_.Time; Server = $_.Server
            Detail = $_.Detail; Hint = (Get-EmtStatusHint $_.Detail); StatusCode = $_.Status; KindCode = $_.Kind
        }
    } | Sort-Object { -(Get-EmtStatusRank $_.StatusCode) }, Recipient)
    # Problems first: failed and dropped, then the others.
    $rows = @(@($rows | Where-Object { $_.StatusCode -in 'Failed', 'Dropped', 'Deferred' }) +
              @($rows | Where-Object { $_.StatusCode -notin 'Failed', 'Dropped', 'Deferred' }))
    Set-EmtGridData $gridTrRecipients $rows ([ordered]@{
        Recipient = 'col.recipient'; Status = 'col.status'; Kind = 'col.kind'; Time = 'col.time'; Server = 'col.server'
        Hint = 'col.hint'; Detail = 'col.detail'; StatusCode = ''; KindCode = ''
    }) -PartsColumn 'Status' -Parts { param($r) @{ Text = $r.Status; Tone = (Get-EmtStatusTone $r.StatusCode) } } `
       -Mark { param($r) Get-EmtStatusTone $r.StatusCode }
    $events = @($t.Events | Where-Object { $_.MessageId -eq $msgId })
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
    $script:MailboxCheck = $null
    Show-EmtTraceViews $t
    if (-not @($t.Messages).Count) {
        Write-EmtLog (T 'log.traceEmpty') 'Warn'
        foreach ($h in @($t.NoResultHints)) { Write-EmtLog ('  - ' + $h) 'Warn' }
    }
    $first = @($gridTrMessages.ItemsSource)[0]
    if ($first) { $gridTrMessages.SelectedItem = $first }
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

#region ---------- Address book search (the same logic as in ExchangeMessageTrail.ps1) ----------

function Complete-EmtAddressResolve {
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
    $newSenders = @($senders | ForEach-Object { if ($map.ContainsKey($_)) { $map[$_] } else { $_ } } | Sort-Object -Unique)
    $SendersBox.Text = $newSenders -join '; '
    if ($recipient -and $map.ContainsKey($recipient)) { $RecipientBox.Text = @($map[$recipient])[0] }
    $true
}

function Invoke-EmtWithResolve {
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

function Set-EmtPeriod([int] $StartDays, [int] $EndDays) {
    $dtTrStart.SelectedDate = (Get-Date).Date.AddDays($StartDays)
    $dtTrEnd.SelectedDate   = (Get-Date).Date.AddDays($EndDays)
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
    Update-EmtConnIndicator
    Start-EmtOperation -Name 'Connect' -Params @{ Server = $srv } -StatusText (T 'status.connecting' $srv) -OnSuccess {
        param($caps)
        $script:Caps = $caps
        $script:ConnectedServer = $tbServer.Text.Trim()
        Update-EmtConnIndicator
        Save-EmtSettings
        if (-not $caps.Tracking)     { Write-EmtLog (T 'log.noTrackRole') 'Warn' }
        if (-not $caps.MailboxCheck) { Write-EmtLog (T 'log.noMailboxRole') 'Warn' }
    }
})

$cmbLanguage.Add_SelectionChanged({
    if ($script:Loading -or $cmbLanguage.SelectedIndex -lt 0) { return }
    $code = $script:LanguageCodes[$cmbLanguage.SelectedIndex]
    if ($code -eq $Sync.Lang) { return }
    $Sync.Lang = $code
    Update-EmtLanguage
    Save-EmtSettings
    Write-EmtLog (T 'log.language')
})

$cmbTheme.Add_SelectionChanged({
    if ($script:Loading -or $cmbTheme.SelectedIndex -lt 0) { return }
    $name = $script:ThemeCodes[$cmbTheme.SelectedIndex]
    if ($name -eq $script:Theme) { return }
    Set-EmtTheme $name
    Save-EmtSettings
})

$btnTrToday.Add_Click({     Set-EmtPeriod 0 1 })
$btnTrYesterday.Add_Click({ Set-EmtPeriod -1 0 })
$btnTrWeek.Add_Click({      Set-EmtPeriod -6 1 })

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
foreach ($g in $gridTrMessages, $gridTrRecipients, $gridTrEvents, $gridTrMailbox) {
    $g.Add_IsVisibleChanged({
        [void]$window.Dispatcher.BeginInvoke([Action]{ Resize-EmtGrids }, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
    })
}

$btnTrCheck.Add_Click({
    $rcpt = Get-EmtSelectedRecipient
    $msgId = Get-EmtSelectedMessageId
    $m = @($script:Trace.Messages | Where-Object MessageId -eq $msgId)[0]
    if (-not $rcpt -or -not $m) { return }
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
        if ((Get-EmtSelectedMessageId) -ne $script:CheckForMessage) { Write-EmtLog (T 'log.checkOtherMessage'); return }
        Show-EmtMailboxView $script:MailboxCheck
        Update-EmtTabTitles
        $trDetail.SelectedItem = $tabTrMailbox
    }
})

$btnCancel.Add_Click({
    if ($script:Job) {
        $Sync.Cancel = $true
        Write-EmtLog (T 'log.cancelRequested') 'Warn'
    }
})

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(400)
$timer.Add_Tick({
    $line = $null
    while ($Sync.Log.TryDequeue([ref]$line)) { Add-EmtLogLine $line }
    if ($script:Job) {
        $progress.Value = [int]$Sync.Progress
        $st = [string]$Sync.Status
        if ($lblStatus.Text -ne $st) { Set-EmtStatusText $st }
        if ($script:Job.Handle.IsCompleted) {
            # Stop the timer: a result handler can open a modal dialog (no re-entry).
            $timer.Stop()
            try {
                Complete-EmtOperation
                while ($Sync.Log.TryDequeue([ref]$line)) { Add-EmtLogLine $line }
            } finally { $timer.Start() }
        }
    }
})

$window.Add_SourceInitialized({ Set-EmtTitleBar $window })

$window.Add_ContentRendered({
    # Show the window in front of the other windows once. Topmost stays off.
    $window.Topmost = $true
    [void]$window.Activate()
    $window.Topmost = $false
    Write-EmtLog (T 'log.started' $LogPath)
    Write-EmtLog (T 'log.auditFile' $AuditPath)
    if ($tbServer.Text) { $btnConnect.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) }
})

$window.Add_Closing({
    param($s, $e)
    if ($script:Job) {
        if (-not (Show-EmtMessage -Text (T 'dlg.closeRunning') -Kind Question)) { $e.Cancel = $true; return }
        $Sync.Cancel = $true
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
$cmbLanguage.SelectedIndex = [array]::IndexOf($script:LanguageCodes, [string]$Sync.Lang)
# The last window size and position. Only if the window is on the screen now (a monitor can be removed).
if ($settings.ContainsKey('Window') -and $settings.Window) {
    try {
        $wb = $settings.Window
        $vs = New-Object System.Windows.Rect([System.Windows.SystemParameters]::VirtualScreenLeft, [System.Windows.SystemParameters]::VirtualScreenTop,
                                             [System.Windows.SystemParameters]::VirtualScreenWidth, [System.Windows.SystemParameters]::VirtualScreenHeight)
        $rect = New-Object System.Windows.Rect([int]$wb.X, [int]$wb.Y,
            [math]::Max([int]$wb.Width, $window.MinWidth), [math]::Max([int]$wb.Height, $window.MinHeight))
        $part = [System.Windows.Rect]::Intersect($vs, $rect)
        if (-not $part.IsEmpty -and $part.Width -ge 200 -and $part.Height -ge 100) {
            $window.WindowStartupLocation = 'Manual'
            $window.Left = $rect.X; $window.Top = $rect.Y; $window.Width = $rect.Width; $window.Height = $rect.Height
            if ([bool]$wb.Maximized) { $window.WindowState = 'Maximized' }
        }
    } catch { }
}
Set-EmtPeriod 0 1
# The first day of the week comes from the regional settings of Windows, not from the interface language.
# For example: Monday in Europe, Sunday in the USA. The user can change it in the Windows settings.
foreach ($d in $dtTrStart, $dtTrEnd) { $d.FirstDayOfWeek = $script:PcCulture.DateTimeFormat.FirstDayOfWeek }
$script:Loading = $false

Update-EmtLanguage
$timer.Start()
try {
    [void]$window.ShowDialog()
} finally {
    $timer.Stop()
    if ($script:Job) {
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
