#Requires -Version 5.1
<#
.SYNOPSIS
  Floating status window for the tem-ord-cards CI/CD monitor.
  Auto-refreshes every 30 seconds. Reads monitor.log and state.json.

.USAGE
  # Open the window:
  powershell -ExecutionPolicy Bypass -File status-window.ps1

  # Create a Desktop shortcut (run once):
  powershell -ExecutionPolicy Bypass -File status-window.ps1 -CreateShortcut
#>

param([switch]$CreateShortcut)

$MONITOR_DIR = 'C:\Users\josehoon\scripts\tem-ord-cards-monitor'
$LOG_FILE    = "$MONITOR_DIR\monitor.log"
$STATE_FILE  = "$MONITOR_DIR\state.json"
$TASK_NAME   = 'tem-ord-cards PDF Monitor'
$REFRESH_SEC = 30

# ── Desktop shortcut ────────────────────────────────────────────────────────
if ($CreateShortcut) {
    $wsh = New-Object -ComObject WScript.Shell
    $sc  = $wsh.CreateShortcut("$([Environment]::GetFolderPath('Desktop'))\tem-ord-cards Monitor.lnk")
    $sc.TargetPath       = 'powershell.exe'
    $sc.Arguments        = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$MONITOR_DIR\status-window.ps1`""
    $sc.WorkingDirectory = $MONITOR_DIR
    $sc.IconLocation     = 'shell32.dll,167'
    $sc.Save()
    Write-Host "Shortcut created on Desktop."
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ── Color helpers ────────────────────────────────────────────────────────────
function mc([string]$hex) {
    $h = $hex.TrimStart('#')
    [System.Windows.Media.Color]::FromRgb(
        [Convert]::ToByte($h.Substring(0,2),16),
        [Convert]::ToByte($h.Substring(2,2),16),
        [Convert]::ToByte($h.Substring(4,2),16))
}
function br([string]$hex) { [System.Windows.Media.SolidColorBrush]::new((mc $hex)) }

# ── Status palette: Bg / text-Fg / accent-Bar / sub-Fg ──────────────────────
$STATES = @{
    Idle    = @('#0D2640','#4A82AA','#3A6A9A','#1E3D55')
    Running = @('#1A1008','#C9924A','#B07A30','#4A3810')
    Pending = @('#1A1008','#C9924A','#B07A30','#4A3810')
    Pass    = @('#081E14','#3DAA7D','#2A9A6A','#1A4A30')
    Fail    = @('#1E0808','#C44B4B','#A43A3A','#4A1818')
}

# ── XAML ─────────────────────────────────────────────────────────────────────
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="tem-ord-cards Monitor"
    Width="430" Height="590"
    WindowStyle="SingleBorderWindow"
    ResizeMode="CanResizeWithGrip"
    MinWidth="360" MinHeight="480">

  <Window.Background><SolidColorBrush Color="#071B2E"/></Window.Background>

  <Window.Resources>
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Name="bd" Background="{TemplateBinding Background}"
                    CornerRadius="3" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.75"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="34"/>
      <RowDefinition Height="108"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="48"/>
    </Grid.RowDefinitions>

    <!-- Top bar -->
    <Grid Grid.Row="0" Background="#030F1E">
      <TextBlock Text="TEM-ORD-CARDS  ·  DEPLOY MONITOR"
                 Foreground="#163040" FontFamily="Consolas" FontSize="10" FontWeight="Bold"
                 VerticalAlignment="Center" Margin="14,0,0,0"/>
      <TextBlock Name="RefreshLabel" Text="—"
                 Foreground="#163040" FontFamily="Consolas" FontSize="10"
                 HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,14,0"/>
    </Grid>

    <!-- Status banner -->
    <Grid Grid.Row="1">
      <Border Name="StatusBanner" Background="#0D2640" Padding="24,0,14,0">
        <StackPanel VerticalAlignment="Center">
          <TextBlock Name="StatusWord" Text="IDLE"
                     FontFamily="Segoe UI" FontSize="46" FontWeight="Black"
                     Foreground="#4A82AA" LineHeight="50"/>
          <TextBlock Name="StatusSub" Text="Waiting for next deployment"
                     FontFamily="Consolas" FontSize="11" Foreground="#1E3D55" Margin="2,5,0,0"/>
        </StackPanel>
      </Border>
      <Rectangle Name="AccentBar" Width="4" HorizontalAlignment="Left" Fill="#3A6A9A"/>
    </Grid>

    <!-- Info grid -->
    <Grid Grid.Row="2" Margin="14,14,14,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="112"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Grid.RowDefinitions>
        <RowDefinition Height="24"/>
        <RowDefinition Height="24"/>
        <RowDefinition Height="24"/>
        <RowDefinition Height="24"/>
        <RowDefinition Height="24"/>
      </Grid.RowDefinitions>

      <TextBlock Grid.Row="0" Grid.Column="0" Text="latest deploy"  Foreground="#163040" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>
      <TextBlock Grid.Row="1" Grid.Column="0" Text="last tested"    Foreground="#163040" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>
      <TextBlock Grid.Row="2" Grid.Column="0" Text="checked at"     Foreground="#163040" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>
      <TextBlock Grid.Row="3" Grid.Column="0" Text="healthcheck"    Foreground="#163040" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>
      <TextBlock Grid.Row="4" Grid.Column="0" Text="scheduler"      Foreground="#163040" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>

      <TextBlock Name="ValLatest"    Grid.Row="0" Grid.Column="1" Text="—" Foreground="#7AAAC8" FontFamily="Consolas" FontSize="12" FontWeight="Bold" VerticalAlignment="Center"/>
      <TextBlock Name="ValTested"    Grid.Row="1" Grid.Column="1" Text="—" Foreground="#7AAAC8" FontFamily="Consolas" FontSize="12" VerticalAlignment="Center"/>
      <TextBlock Name="ValCheck"     Grid.Row="2" Grid.Column="1" Text="—" Foreground="#2A5070" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>
      <TextBlock Name="ValHealth"    Grid.Row="3" Grid.Column="1" Text="—" Foreground="#2A5070" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>
      <TextBlock Name="ValScheduler" Grid.Row="4" Grid.Column="1" Text="—" Foreground="#2A5070" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>
    </Grid>

    <!-- Log panel -->
    <Border Grid.Row="3" Background="#030F1E" Margin="14,10,14,0"
            CornerRadius="3" Padding="10,8">
      <ScrollViewer Name="LogScroll" VerticalScrollBarVisibility="Auto"
                    HorizontalScrollBarVisibility="Disabled">
        <TextBlock Name="LogBlock" FontFamily="Consolas" FontSize="11"
                   TextWrapping="Wrap" LineHeight="17"/>
      </ScrollViewer>
    </Border>

    <!-- Action bar -->
    <Grid Grid.Row="4" Margin="14,0,14,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock Name="CountdownLabel" Grid.Column="0"
                 Text="" Foreground="#162A3A" FontFamily="Consolas" FontSize="10"
                 VerticalAlignment="Center"/>
      <Button Name="BtnRunNow" Grid.Column="1"
              Style="{StaticResource Btn}"
              Content="▶  Run Now"
              Padding="16,8" FontWeight="SemiBold"
              Background="#4A2808" Foreground="#C9924A"/>
      <Button Name="BtnOpenLog" Grid.Column="3"
              Style="{StaticResource Btn}"
              Content="Open Log"
              Padding="14,8"
              Background="#0D2640" Foreground="#2A5070"/>
    </Grid>
  </Grid>
</Window>
"@

# ── Load ─────────────────────────────────────────────────────────────────────
$reader = [System.Xml.XmlNodeReader]::new($XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# XamlReader.Load() doesn't always register all named elements in the window
# namescope; fall back to LogicalTreeHelper which walks the tree directly.
function Find-WpfControl([string]$name) {
    $ctrl = $window.FindName($name)
    if (-not $ctrl) { $ctrl = [System.Windows.LogicalTreeHelper]::FindLogicalNode($window, $name) }
    $ctrl
}

$statusBanner    = Find-WpfControl 'StatusBanner'
$statusWord      = Find-WpfControl 'StatusWord'
$statusSub       = Find-WpfControl 'StatusSub'
$accentBar       = Find-WpfControl 'AccentBar'
$refreshLabel    = Find-WpfControl 'RefreshLabel'
$valLatest       = Find-WpfControl 'ValLatest'
$valTested       = Find-WpfControl 'ValTested'
$valCheck        = Find-WpfControl 'ValCheck'
$valHealth       = Find-WpfControl 'ValHealth'
$valScheduler    = Find-WpfControl 'ValScheduler'
$logBlock        = Find-WpfControl 'LogBlock'
$logScroll       = Find-WpfControl 'LogScroll'
$btnRunNow       = Find-WpfControl 'BtnRunNow'
$btnOpenLog      = Find-WpfControl 'BtnOpenLog'
$countdownLabel  = Find-WpfControl 'CountdownLabel'

# ── Apply status palette ──────────────────────────────────────────────────────
function Set-Status([string]$key, [string]$word, [string]$sub) {
    $p = $STATES[$key]
    $statusBanner.Background = br $p[0]
    $statusWord.Foreground   = br $p[1]
    $accentBar.Fill          = br $p[2]
    $statusSub.Foreground    = br $p[3]
    $statusWord.Text         = $word
    $statusSub.Text          = $sub
}

# ── Log line helper (per-line coloring via Inlines) ───────────────────────────
function Add-LogLine([string]$line) {
    $run = [System.Windows.Documents.Run]::new("$line`n")
    $run.Foreground = switch -Regex ($line) {
        '\[PASS\]'  { br '#1A5A38' }
        '\[FAIL\]'  { br '#5A2020' }
        '\[WARN\]'  { br '#4A3818' }
        '\[ERROR\]' { br '#6A2020' }
        'Running PDF tests|Starting PDF' { br '#2A4A68' }
        default     { br '#1E3D55' }
    }
    $logBlock.Inlines.Add($run)
}

# ── Main refresh ──────────────────────────────────────────────────────────────
$script:nextRefresh = (Get-Date).AddSeconds($REFRESH_SEC)

function Update-UI {
    # --- state.json ---
    $knownSha = '—'; $testedSha = '—'
    if (Test-Path $STATE_FILE) {
        try {
            $s = Get-Content $STATE_FILE -Raw | ConvertFrom-Json
            if ($s.lastKnownSha  -and $s.lastKnownSha.Length  -ge 8) { $knownSha  = $s.lastKnownSha.Substring(0,8)  }
            if ($s.lastTestedSha -and $s.lastTestedSha.Length -ge 8) { $testedSha = $s.lastTestedSha.Substring(0,8) }
        } catch {}
    }
    $valLatest.Text = $knownSha
    $valTested.Text = $testedSha

    # --- log ---
    $lines = @()
    if (Test-Path $LOG_FILE) {
        try { $lines = Get-Content $LOG_FILE -Tail 100 } catch {}
    }

    # last check time
    $chk = $lines | Where-Object { $_ -match '\[INFO\]\s+Monitor check started' } | Select-Object -Last 1
    $valCheck.Text = if ($chk -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})') { $matches[1] } else { '—' }

    # healthcheck
    $hcLine = $lines | Where-Object { $_ -match '(Healthcheck|healthcheck)' } | Select-Object -Last 1
    if ($hcLine -match 'OK') {
        $valHealth.Text       = '✓  reachable (VPN up)'
        $valHealth.Foreground = br '#2A6A4A'
    } elseif ($hcLine -match 'skipped') {
        $valHealth.Text       = '⊘  skipped — VPN not up'
        $valHealth.Foreground = br '#2A5070'
    } else {
        $valHealth.Text       = '—'
        $valHealth.Foreground = br '#2A5070'
    }

    # scheduler
    try {
        $ts = Get-ScheduledTask     -TaskName $TASK_NAME -ErrorAction SilentlyContinue
        $ti = Get-ScheduledTaskInfo -TaskName $TASK_NAME -ErrorAction SilentlyContinue
        if ($ts) {
            $nxt = if ($ti -and $ti.NextRunTime) { '  next ' + $ti.NextRunTime.ToString('HH:mm') } else { '' }
            $valScheduler.Text = "$($ts.State.ToString().ToLower())$nxt"
        } else { $valScheduler.Text = 'not registered' }
    } catch { $valScheduler.Text = 'unknown' }

    # --- determine status ---
    $lastPass = $lines | Where-Object { $_ -match '\[PASS\]' } | Select-Object -Last 1
    $lastFail = $lines | Where-Object { $_ -match '\[FAIL\]' } | Select-Object -Last 1

    function LineDate([string]$l) {
        if ($l -and $l -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})') { return [datetime]$matches[1] }
        return [datetime]::MinValue
    }

    $lastTestStart = $lines | Where-Object { $_ -match 'Running PDF tests' } | Select-Object -Last 1
    $lastComplete  = $lines | Where-Object { $_ -match 'Monitor check complete' } | Select-Object -Last 1
    $testsRunning  = $lastTestStart -and ((LineDate $lastTestStart) -gt (LineDate $lastComplete))

    $taskRunning = $false
    try {
        $tsObj = Get-ScheduledTask -TaskName $TASK_NAME -ErrorAction SilentlyContinue
        if ($tsObj -and $tsObj.State -eq 'Running') { $taskRunning = $true }
    } catch {}

    $isRunning = $testsRunning -or $taskRunning

    if ($isRunning) {
        Set-Status 'Running' 'RUNNING' "testing sha=$knownSha  ·  please wait"
    } elseif ($knownSha -ne '—' -and $testedSha -ne '—' -and $knownSha -ne $testedSha) {
        Set-Status 'Pending' 'PENDING' "new deploy $knownSha not yet tested"
    } elseif ($lastPass -or $lastFail) {
        $pd = LineDate $lastPass
        $fd = LineDate $lastFail
        if ($pd -ge $fd) {
            $mins = if ($lastPass -match '\((\d+)m\)') { "  ·  $($matches[1]) min" } else { '' }
            Set-Status 'Pass' 'PASSED' "sha=$testedSha$mins"
        } else {
            Set-Status 'Fail' 'FAILED' "sha=$testedSha  ·  open log for details"
        }
    } else {
        Set-Status 'Idle' 'IDLE' 'Waiting for next deployment'
    }

    # --- log block (last 15 lines, per-line color) ---
    $logBlock.Inlines.Clear()
    $recent = if ($lines.Count -gt 0) { $lines | Select-Object -Last 15 } else { @('— no log yet —') }
    foreach ($ln in $recent) { Add-LogLine $ln }
    $logScroll.ScrollToBottom()

    $refreshLabel.Text         = "refreshed $(Get-Date -Format 'HH:mm:ss')"
    $script:nextRefresh        = (Get-Date).AddSeconds($REFRESH_SEC)
}

# ── Countdown (every second) ──────────────────────────────────────────────────
$ticker = [System.Windows.Threading.DispatcherTimer]::new()
$ticker.Interval = [TimeSpan]::FromSeconds(1)
$ticker.Add_Tick({
    $rem = [int][Math]::Max(0, ($script:nextRefresh - (Get-Date)).TotalSeconds)
    $countdownLabel.Text = "refresh in ${rem}s"
})
$ticker.Start()

# ── Refresh timer ─────────────────────────────────────────────────────────────
$refreshTimer = [System.Windows.Threading.DispatcherTimer]::new()
$refreshTimer.Interval = [TimeSpan]::FromSeconds($REFRESH_SEC)
$refreshTimer.Add_Tick({ Update-UI })
$refreshTimer.Start()

# ── Button handlers ───────────────────────────────────────────────────────────
$btnRunNow.Add_Click({
    try {
        Start-ScheduledTask -TaskName $TASK_NAME -ErrorAction Stop
        Set-Status 'Running' 'RUNNING' 'manually triggered  ·  please wait'
    } catch {
        [System.Windows.MessageBox]::Show("Could not start task:`n$_", 'Error', 'OK', 'Error') | Out-Null
    }
})

$btnOpenLog.Add_Click({
    if (Test-Path $LOG_FILE) { Start-Process notepad.exe $LOG_FILE }
    else { [System.Windows.MessageBox]::Show("Log not found: $LOG_FILE") | Out-Null }
})

# ── Show ──────────────────────────────────────────────────────────────────────
Update-UI
$window.ShowDialog() | Out-Null
$ticker.Stop()
$refreshTimer.Stop()
