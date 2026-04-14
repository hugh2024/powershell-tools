#Requires -Version 5.0
<#
.SYNOPSIS    NetProbe - Network Scanner GUI
.DESCRIPTION Portable network scanning tool for personal lab and security testing.
             Uses portable nmap bundled in .\nmap\nmap.exe
             No installation or admin required for TCP scans.
.NOTES       Version: 1.0
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$Script:AppName = 'NetProbe'
$Script:AppVersion = '1.0'

# ── NMAP PATH ─────────────────────────────────────────────────────────────────
$ScriptDir = if ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($PSCommandPath) {
    Split-Path -Parent $PSCommandPath
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}

$NmapCandidates = @(
    (Join-Path $ScriptDir "nmap\nmap.exe"),
    (Join-Path $ScriptDir "nmap.exe"),
    "$env:ProgramFiles\Nmap\nmap.exe",
    "${env:ProgramFiles(x86)}\Nmap\nmap.exe"
) | Where-Object { $_ }

$NmapPath = $NmapCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $NmapPath) {
    $sys = Get-Command nmap -ErrorAction SilentlyContinue
    if ($sys -and $sys.Source) {
        $NmapPath = $sys.Source
    }
}

# ── XAML ──────────────────────────────────────────────────────────────────────
[xml]$XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="NetProbe - Network Scanner"
    Height="820" Width="1260" MinHeight="580" MinWidth="860"
    WindowStartupLocation="CenterScreen"
    Background="#0a0e14" FontFamily="Consolas" FontSize="12">

  <Window.Resources>

    <Style x:Key="SBtn" TargetType="Button">
      <Setter Property="Background" Value="#131c2e"/>
      <Setter Property="Foreground" Value="#cdd8e8"/>
      <Setter Property="BorderBrush" Value="#1e2d45"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Name="Bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#00c8ff"/>
                <Setter Property="Foreground" Value="#00c8ff"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="RunBtn" TargetType="Button">
      <Setter Property="Background" Value="#0084aa"/>
      <Setter Property="Foreground" Value="#000000"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="5" Padding="14,7">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#00c8ff"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Background" Value="#1e2d45"/>
                <Setter Property="Foreground" Value="#5a7294"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="DkChk" TargetType="CheckBox">
      <Setter Property="Foreground" Value="#cdd8e8"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Margin" Value="0,3"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>

    <Style x:Key="DkTxt" TargetType="TextBox">
      <Setter Property="Background" Value="#131c2e"/>
      <Setter Property="Foreground" Value="#cdd8e8"/>
      <Setter Property="BorderBrush" Value="#1e2d45"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="7,5"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="CaretBrush" Value="#00c8ff"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border Name="Bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsFocused" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#0084aa"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="DkCbo" TargetType="ComboBox">
      <Setter Property="Background" Value="#131c2e"/>
      <Setter Property="Foreground" Value="#cdd8e8"/>
      <Setter Property="BorderBrush" Value="#1e2d45"/>
      <Setter Property="Padding" Value="7,5"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="11"/>
    </Style>

    <Style x:Key="DkTab" TargetType="TabItem">
      <Setter Property="Background" Value="#0e1420"/>
      <Setter Property="Foreground" Value="#5a7294"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Padding" Value="14,7"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border Name="Bd" Background="#0e1420"
                    BorderBrush="Transparent" BorderThickness="0,0,0,2"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter Name="Cp"
                TextElement.Foreground="{TemplateBinding Foreground}"
                HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#00c8ff"/>
                <Setter TargetName="Cp" Property="TextElement.Foreground" Value="#00c8ff"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="54"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="22"/>
    </Grid.RowDefinitions>

    <!-- HEADER -->
    <Border Grid.Row="0" Background="#0e1420" BorderBrush="#1e2d45" BorderThickness="0,0,0,1">
      <Grid Margin="14,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
          <Border Background="#0084aa" CornerRadius="7" Width="30" Height="30" Margin="0,0,10,0">
            <TextBlock Text="N" Foreground="#00c8ff" FontWeight="Bold" FontSize="16"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel VerticalAlignment="Center">
            <TextBlock FontFamily="Segoe UI" FontWeight="ExtraBold" FontSize="15" Foreground="White">
              <Run Text="Net"/><Run Text="Probe" Foreground="#00c8ff"/>
            </TextBlock>
            <TextBlock Text="INTERNAL NETWORK SCANNER" FontSize="9" Foreground="#5a7294"/>
          </StackPanel>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0,0,0">
          <Button x:Name="BtnTabNmap"  Style="{StaticResource SBtn}" Content="NMAP SCANNER"
                  Background="#0e1420" BorderThickness="0,0,0,2" BorderBrush="#00c8ff"
                  Foreground="#00c8ff" Margin="0,0,4,0" Padding="14,8"/>
          <Button x:Name="BtnTabTrace" Style="{StaticResource SBtn}" Content="TRACEROUTE"
                  Background="#0e1420" BorderThickness="0,0,0,2" BorderBrush="Transparent"
                  Foreground="#5a7294" Margin="0,0,4,0" Padding="14,8"/>
          <Button x:Name="BtnTabHist"  Style="{StaticResource SBtn}" Content="HISTORY"
                  Background="#0e1420" BorderThickness="0,0,0,2" BorderBrush="Transparent"
                  Foreground="#5a7294" Padding="14,8"/>
        </StackPanel>
        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse Width="7" Height="7" Fill="#00e676" Margin="0,0,6,0"/>
          <TextBlock x:Name="TxtNmapStatus" Text="CHECKING..." Foreground="#5a7294" FontSize="10" VerticalAlignment="Center" Margin="0,0,14,0"/>
          <Border BorderBrush="#1e2d45" BorderThickness="1" CornerRadius="8" Padding="7,2">
            <TextBlock Text="PERSONAL REPO BUILD" Foreground="#5a7294" FontSize="9"/>
          </Border>
        </StackPanel>
      </Grid>
    </Border>

    <!-- TARGET BAR -->
    <Border Grid.Row="1" Background="#0e1420" BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,8,14,10">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="2*"/>
          <ColumnDefinition Width="152"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Margin="0,0,8,0">
          <TextBlock Text="HOST / IP / CIDR" Foreground="#5a7294" FontSize="9" Margin="0,0,0,3"/>
          <TextBox x:Name="TxtTarget" Style="{StaticResource DkTxt}" MinHeight="32"
                   ToolTip="IP, hostname, or CIDR. e.g. 192.168.1.1 or 10.0.0.0/24"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Margin="0,0,8,0">
          <TextBlock Text="PORTS" Foreground="#5a7294" FontSize="9" Margin="0,0,0,3"/>
          <TextBox x:Name="TxtPorts" Style="{StaticResource DkTxt}" MinHeight="32"
                   ToolTip="22,80,443 or 1-1024 or top100. Blank = top 1000"/>
        </StackPanel>
        <StackPanel Grid.Column="2" VerticalAlignment="Bottom" HorizontalAlignment="Right" Orientation="Horizontal" Margin="4,0,0,0">
          <Button x:Name="BtnRun"  Content="Run Scan" Style="{StaticResource RunBtn}" Height="32" Width="100" Margin="0,0,8,0"/>
          <Button x:Name="BtnStop" Content="Stop"     Style="{StaticResource SBtn}"   Height="32" Width="48"  IsEnabled="False"
                  Foreground="#ff4444" BorderBrush="#ff4444"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- MAIN CONTENT -->
    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="270" MinWidth="180" MaxWidth="420"/>
        <ColumnDefinition Width="5"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- SIDEBAR -->
      <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Background="#0e1420">
        <StackPanel>

          <!-- NMAP OPTIONS -->
          <StackPanel x:Name="PanelNmap">

            <!-- Timing -->
            <Border BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,10">
              <StackPanel>
                <TextBlock Text="TIMING" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="80"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0" Margin="0,0,6,0">
                    <TextBlock Text="Speed" Foreground="#5a7294" FontSize="10" Margin="0,0,0,2"/>
                    <ComboBox x:Name="CboTiming" Style="{StaticResource DkCbo}"
                              ToolTip="T2=Polite, T3=Normal, T4=Aggressive (LAN), T5=Insane">
                      <ComboBoxItem Content="T2 - Polite"/>
                      <ComboBoxItem Content="T3 - Normal" IsSelected="True"/>
                      <ComboBoxItem Content="T4 - Aggressive"/>
                      <ComboBoxItem Content="T5 - Insane"/>
                    </ComboBox>
                  </StackPanel>
                  <StackPanel Grid.Column="1">
                    <TextBlock Text="Timeout(s)" Foreground="#5a7294" FontSize="10" Margin="0,0,0,2"/>
                    <TextBox x:Name="TxtTimeout" Text="120" Style="{StaticResource DkTxt}"/>
                  </StackPanel>
                </Grid>
                <TextBlock Text="T4 for LAN. T2 for remote targets. T5 may miss ports."
                           Foreground="#3a4f6a" FontSize="10" Margin="0,4,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>

            <!-- Scan Type -->
            <Border BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,10">
              <StackPanel>
                <TextBlock Text="SCAN TYPE" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <CheckBox x:Name="ChkSynS"  Style="{StaticResource DkChk}" Content="-sS  SYN Stealth (fast, needs admin)"
                          ToolTip="Sends SYN only. Fast and stealthy. Requires admin privileges."/>
                <CheckBox x:Name="ChkTcpC"  Style="{StaticResource DkChk}" Content="-sT  TCP Connect (no admin needed)" IsChecked="True"
                          ToolTip="Full TCP handshake. Works without admin. Safe default."/>
                <CheckBox x:Name="ChkUdp"   Style="{StaticResource DkChk}" Content="-sU  UDP Scan (DNS, SNMP, DHCP)"
                          ToolTip="Scans UDP ports. Finds DNS(53), SNMP(161), DHCP(67). Slow."/>
                <CheckBox x:Name="ChkAckS"  Style="{StaticResource DkChk}" Content="-sA  ACK Scan (map firewall rules)"
                          ToolTip="Maps firewall rules. Shows filtered vs unfiltered ports."/>
              </StackPanel>
            </Border>

            <!-- Detection -->
            <Border BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,10">
              <StackPanel>
                <TextBlock Text="DETECTION" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <CheckBox x:Name="ChkSV"  Style="{StaticResource DkChk}" Content="-sV  Version Detection"
                          ToolTip="Probes open ports to detect service name and version."/>
                <CheckBox x:Name="ChkO"   Style="{StaticResource DkChk}" Content="-O   OS Detection (needs admin)"
                          ToolTip="Tries to guess remote OS (Windows/Linux). Needs admin."/>
                <CheckBox x:Name="ChkSC"  Style="{StaticResource DkChk}" Content="-sC  Default NSE Scripts"
                          ToolTip="Runs nmap default scripts. Checks banners and vulnerabilities."/>
                <CheckBox x:Name="ChkA"   Style="{StaticResource DkChk}" Content="-A   Aggressive (-sV -O -sC)"
                          ToolTip="Combines version, OS, scripts and traceroute. Thorough but slow."/>
              </StackPanel>
            </Border>

            <!-- Host Discovery -->
            <Border BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,10">
              <StackPanel>
                <TextBlock Text="HOST DISCOVERY" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <CheckBox x:Name="ChkPn"  Style="{StaticResource DkChk}" Content="-Pn  Skip Ping  [Recommended]" IsChecked="True"
                          ToolTip="Skips ping. Scans even if host does not respond to ICMP. Best for firewalled hosts."/>
                <CheckBox x:Name="ChkPS"  Style="{StaticResource DkChk}" Content="-PS  TCP SYN Ping (port 80)"
                          ToolTip="Sends TCP SYN to port 80 to check if host is up."/>
                <CheckBox x:Name="ChkPA"  Style="{StaticResource DkChk}" Content="-PA  TCP ACK Ping (port 80)"
                          ToolTip="Sends TCP ACK to port 80. Useful when SYN is blocked."/>
                <CheckBox x:Name="ChkPE"  Style="{StaticResource DkChk}" Content="-PE  ICMP Echo (traditional ping)"
                          ToolTip="Classic ICMP ping. Fast but often blocked by firewalls."/>
              </StackPanel>
            </Border>

            <!-- Output Options -->
            <Border BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,10">
              <StackPanel>
                <TextBlock Text="OUTPUT OPTIONS" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <CheckBox x:Name="ChkV"    Style="{StaticResource DkChk}" Content="-v   Verbose output" IsChecked="True"
                          ToolTip="Shows open ports as found and more scan detail."/>
                <CheckBox x:Name="ChkOpen" Style="{StaticResource DkChk}" Content="--open  Show open ports only"
                          ToolTip="Hides closed and filtered ports. Cleaner output."/>
                <CheckBox x:Name="ChkRsn"  Style="{StaticResource DkChk}" Content="--reason  Show port state reason"
                          ToolTip="Shows why each port is open/closed. e.g. syn-ack = open."/>
                <CheckBox x:Name="ChkN"    Style="{StaticResource DkChk}" Content="-n   No DNS (faster)"
                          ToolTip="Skips DNS lookup for each IP. Speeds up scans."/>
              </StackPanel>
            </Border>

            <!-- Custom Flags -->
            <Border BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,10">
              <StackPanel>
                <TextBlock Text="CUSTOM FLAGS" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <TextBox x:Name="TxtCustom" Style="{StaticResource DkTxt}"
                         ToolTip="Raw nmap flags. e.g. --script=banner --max-retries 1"/>
                <TextBlock Text="Raw nmap flags. Use with care." Foreground="#3a4f6a" FontSize="10" Margin="0,3,0,0"/>
              </StackPanel>
            </Border>

            <!-- Presets -->
            <Border Padding="14,10">
              <StackPanel>
                <TextBlock Text="QUICK PRESETS" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <Button x:Name="BtnPQ" Style="{StaticResource SBtn}" Margin="0,2" HorizontalContentAlignment="Left"
                        Content="[FAST]  Quick Scan - top 100 ports"
                        ToolTip="Top 100 ports, no ping, TCP Connect. Fast."/>
                <Button x:Name="BtnPF" Style="{StaticResource SBtn}" Margin="0,2" HorizontalContentAlignment="Left"
                        Content="[SAFE]  No-Ping Full - all 65535 ports"
                        ToolTip="All 65535 ports with -Pn. Best for firewalled hosts."/>
                <Button x:Name="BtnPS" Style="{StaticResource SBtn}" Margin="0,2" HorizontalContentAlignment="Left"
                        Content="[DEEP]  Service Fingerprint - -sV -O -sC"
                        ToolTip="Version, OS, scripts. Most thorough. Slow."/>
                <Button x:Name="BtnPSt" Style="{StaticResource SBtn}" Margin="0,2" HorizontalContentAlignment="Left"
                        Content="[QUIET] Stealth - SYN T2 low footprint"
                        ToolTip="SYN scan at T2. Low footprint, less IDS alerts."/>
                <Button x:Name="BtnPW" Style="{StaticResource SBtn}" Margin="0,2" HorizontalContentAlignment="Left"
                        Content="[WEB]   Web Ports - 80,443,8080,8443"
                        ToolTip="HTTP/HTTPS ports only."/>
                <Button x:Name="BtnPN" Style="{StaticResource SBtn}" Margin="0,2" HorizontalContentAlignment="Left"
                        Content="[DISC]  Subnet Discovery - host only"
                        ToolTip="Discovers alive hosts on subnet. No port scan."/>
              </StackPanel>
            </Border>

          </StackPanel>

          <!-- TRACEROUTE OPTIONS -->
          <StackPanel x:Name="PanelTrace" Visibility="Collapsed">
            <Border BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,10">
              <StackPanel>
                <TextBlock Text="METHOD" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <TextBlock Text="Protocol" Foreground="#5a7294" FontSize="10" Margin="0,0,0,2"/>
                <ComboBox x:Name="CboTrMethod" Style="{StaticResource DkCbo}">
                  <ComboBoxItem Content="UDP - standard, most compatible" IsSelected="True"/>
                  <ComboBoxItem Content="TCP - bypasses UDP/ICMP blocks"/>
                  <ComboBoxItem Content="ICMP - like Windows tracert"/>
                </ComboBox>
                <CheckBox x:Name="ChkTrNmap" Style="{StaticResource DkChk}" Margin="0,8,0,0"
                          Content="Use nmap --traceroute (more detail)"
                          ToolTip="Uses nmap traceroute which shows port and OS info per hop."/>
              </StackPanel>
            </Border>
            <Border BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="14,10">
              <StackPanel>
                <TextBlock Text="OPTIONS" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0" Margin="0,0,4,0">
                    <TextBlock Text="Max Hops" Foreground="#5a7294" FontSize="10" Margin="0,0,0,2"/>
                    <TextBox x:Name="TxtTrHops" Text="30" Style="{StaticResource DkTxt}"/>
                  </StackPanel>
                  <StackPanel Grid.Column="1">
                    <TextBlock Text="Hop Timeout(s)" Foreground="#5a7294" FontSize="10" Margin="0,0,0,2"/>
                    <TextBox x:Name="TxtTrHopTo" Text="3" Style="{StaticResource DkTxt}"/>
                  </StackPanel>
                </Grid>
                <CheckBox x:Name="ChkTrDns" Style="{StaticResource DkChk}" Margin="0,7,0,0"
                          Content="Resolve hostnames per hop (slower)"
                          ToolTip="Adds DNS names to each hop. Slower but readable."/>
              </StackPanel>
            </Border>
            <Border Padding="14,10">
              <StackPanel>
                <TextBlock Text="QUICK TARGETS" Foreground="#5a7294" FontSize="9" Margin="0,0,0,7"/>
                <Button x:Name="BtnTr8" Style="{StaticResource SBtn}" Margin="0,2"
                        Content="Google DNS - 8.8.8.8"/>
                <Button x:Name="BtnTr1" Style="{StaticResource SBtn}" Margin="0,2"
                        Content="Cloudflare - 1.1.1.1"/>
                <Button x:Name="BtnTrGW" Style="{StaticResource SBtn}" Margin="0,2"
                        Content="Default Gateway"/>
                <Button x:Name="BtnRunTr" Style="{StaticResource RunBtn}" Margin="0,10,0,0"
                        Content="Run Traceroute"/>
              </StackPanel>
            </Border>
          </StackPanel>

        </StackPanel>
      </ScrollViewer>

      <!-- SPLITTER -->
      <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch"
                    Background="#1e2d45" ShowsPreview="False"/>

      <!-- OUTPUT -->
      <Grid Grid.Column="2" Background="#0a0e14">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Command Preview -->
        <Border Grid.Row="0" Background="#0e1420" BorderBrush="#1e2d45" BorderThickness="0,0,0,1" Padding="12,7">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="$ " Foreground="#00c8ff" FontWeight="Bold" VerticalAlignment="Center"/>
            <TextBlock Grid.Column="1" x:Name="TxtCmd" Foreground="#00e676" FontSize="11"
                       VerticalAlignment="Center" TextWrapping="Wrap"
                       Text="nmap -sT -Pn -v -T3 &lt;target&gt;"/>
            <Button Grid.Column="2" x:Name="BtnCopyCmd" Content="COPY"
                    Style="{StaticResource SBtn}" Padding="7,3" FontSize="10"/>
          </Grid>
        </Border>

        <!-- Tabs -->
        <TabControl Grid.Row="1" x:Name="TabOut" Background="#0a0e14" BorderThickness="0">
          <TabControl.Resources>
            <Style TargetType="TabItem" BasedOn="{StaticResource DkTab}"/>
          </TabControl.Resources>

          <!-- Raw Output -->
          <TabItem Header="RAW OUTPUT">
            <Grid Background="#0e1420">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="#131c2e" BorderBrush="#1e2d45"
                      BorderThickness="0,0,0,1" Padding="10,5">
                <Grid>
                  <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Scan Output  " Foreground="#cdd8e8"
                               FontFamily="Segoe UI" FontWeight="SemiBold" FontSize="12"/>
                    <Border x:Name="ChipStat" Background="#0a0e14" BorderBrush="#1e2d45"
                            BorderThickness="1" CornerRadius="3" Padding="5,1" Margin="4,0">
                      <TextBlock x:Name="TxtStat" Text="" Foreground="#5a7294" FontSize="10"/>
                    </Border>
                    <Border x:Name="ChipTime" Background="#0a0e14" BorderBrush="#1e2d45"
                            BorderThickness="1" CornerRadius="3" Padding="5,1" Margin="4,0">
                      <TextBlock x:Name="TxtTime" Text="" Foreground="#5a7294" FontSize="10"/>
                    </Border>
                  </StackPanel>
                  <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="BtnCopyOut"  Content="COPY"  Style="{StaticResource SBtn}" Padding="7,3" FontSize="10" Margin="0,0,4,0"/>
                    <Button x:Name="BtnSaveOut"  Content="SAVE"  Style="{StaticResource SBtn}" Padding="7,3" FontSize="10" Margin="0,0,4,0"/>
                    <Button x:Name="BtnClearOut" Content="CLEAR" Style="{StaticResource SBtn}" Padding="7,3" FontSize="10"/>
                  </StackPanel>
                </Grid>
              </Border>
              <RichTextBox Grid.Row="1" x:Name="Rtb" Background="#0a0e14" Foreground="#cdd8e8"
                           FontFamily="Consolas" FontSize="12" IsReadOnly="True"
                           BorderThickness="0" Padding="12" VerticalScrollBarVisibility="Auto"/>
            </Grid>
          </TabItem>

          <!-- Port Table -->
          <TabItem Header="PORT TABLE">
            <DataGrid x:Name="DgPorts" Background="#0a0e14" Foreground="#cdd8e8"
                      FontFamily="Consolas" FontSize="11" BorderThickness="0"
                      GridLinesVisibility="Horizontal" RowBackground="#0a0e14"
                      AlternatingRowBackground="#0e1420" HeadersVisibility="Column"
                      IsReadOnly="True" AutoGenerateColumns="False" CanUserResizeRows="False">
              <DataGrid.ColumnHeaderStyle>
                <Style TargetType="DataGridColumnHeader">
                  <Setter Property="Background" Value="#131c2e"/>
                  <Setter Property="Foreground" Value="#5a7294"/>
                  <Setter Property="FontSize"   Value="10"/>
                  <Setter Property="Padding"    Value="10,5"/>
                  <Setter Property="BorderThickness" Value="0,0,0,1"/>
                  <Setter Property="BorderBrush" Value="#1e2d45"/>
                </Style>
              </DataGrid.ColumnHeaderStyle>
              <DataGrid.Columns>
                <DataGridTextColumn Header="PORT"    Binding="{Binding Port}"    Width="70"/>
                <DataGridTextColumn Header="PROTO"   Binding="{Binding Proto}"   Width="65"/>
                <DataGridTextColumn Header="STATE"   Binding="{Binding State}"   Width="100"/>
                <DataGridTextColumn Header="SERVICE" Binding="{Binding Service}" Width="120"/>
                <DataGridTextColumn Header="INFO"    Binding="{Binding Info}"    Width="*"/>
              </DataGrid.Columns>
            </DataGrid>
          </TabItem>

          <!-- Traceroute -->
          <TabItem Header="TRACEROUTE">
            <Grid Background="#0e1420">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="180"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="#131c2e" BorderBrush="#1e2d45"
                      BorderThickness="0,0,0,1" Padding="10,5">
                <Grid>
                  <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Traceroute Output  " Foreground="#cdd8e8"
                               FontFamily="Segoe UI" FontWeight="SemiBold" FontSize="12"/>
                    <Border Background="#0a0e14" BorderBrush="#1e2d45"
                            BorderThickness="1" CornerRadius="3" Padding="5,1" Margin="4,0">
                      <TextBlock x:Name="TxtTrStat" Text="" Foreground="#5a7294" FontSize="10"/>
                    </Border>
                  </StackPanel>
                  <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="BtnCopyTr"  Content="COPY"  Style="{StaticResource SBtn}" Padding="7,3" FontSize="10" Margin="0,0,4,0"/>
                    <Button x:Name="BtnClearTr" Content="CLEAR" Style="{StaticResource SBtn}" Padding="7,3" FontSize="10"/>
                  </StackPanel>
                </Grid>
              </Border>
              <RichTextBox Grid.Row="1" x:Name="RtbTr" Background="#0a0e14" Foreground="#cdd8e8"
                           FontFamily="Consolas" FontSize="12" IsReadOnly="True"
                           BorderThickness="0" Padding="12" VerticalScrollBarVisibility="Auto"/>
              <DataGrid Grid.Row="2" x:Name="DgHops" Background="#0a0e14" Foreground="#cdd8e8"
                        FontFamily="Consolas" FontSize="11" BorderThickness="0,1,0,0"
                        BorderBrush="#1e2d45" GridLinesVisibility="Horizontal"
                        RowBackground="#0a0e14" AlternatingRowBackground="#0e1420"
                        HeadersVisibility="Column" IsReadOnly="True"
                        AutoGenerateColumns="False" CanUserResizeRows="False">
                <DataGrid.ColumnHeaderStyle>
                  <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="#131c2e"/>
                    <Setter Property="Foreground" Value="#5a7294"/>
                    <Setter Property="FontSize"   Value="10"/>
                    <Setter Property="Padding"    Value="10,4"/>
                    <Setter Property="BorderThickness" Value="0,0,0,1"/>
                    <Setter Property="BorderBrush" Value="#1e2d45"/>
                  </Style>
                </DataGrid.ColumnHeaderStyle>
                <DataGrid.Columns>
                  <DataGridTextColumn Header="#"   Binding="{Binding Hop}" Width="40"/>
                  <DataGridTextColumn Header="IP"  Binding="{Binding IP}"  Width="140"/>
                  <DataGridTextColumn Header="RTT" Binding="{Binding RTT}" Width="110"/>
                  <DataGridTextColumn Header="RAW" Binding="{Binding Raw}" Width="*"/>
                </DataGrid.Columns>
              </DataGrid>
            </Grid>
          </TabItem>

          <!-- History -->
          <TabItem Header="HISTORY">
            <Grid Background="#0e1420">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="#131c2e" BorderBrush="#1e2d45"
                      BorderThickness="0,0,0,1" Padding="10,5">
                <Grid>
                  <TextBlock Text="Recent scans - double-click to reload" Foreground="#5a7294"
                             FontSize="10" VerticalAlignment="Center"/>
                  <Button x:Name="BtnClrHist" Content="CLEAR" Style="{StaticResource SBtn}"
                          HorizontalAlignment="Right" Padding="7,3" FontSize="10"/>
                </Grid>
              </Border>
              <ListBox x:Name="LstHist" Grid.Row="1" Background="#0a0e14"
                       BorderThickness="0" Foreground="#cdd8e8" FontFamily="Consolas" FontSize="11"/>
            </Grid>
          </TabItem>

        </TabControl>
      </Grid>
    </Grid>

    <!-- STATUS BAR -->
    <Border Grid.Row="3" Background="#0e1420" BorderBrush="#1e2d45" BorderThickness="0,1,0,0" Padding="12,3">
      <Grid>
        <TextBlock x:Name="TxtSBar" Foreground="#5a7294" FontSize="10" VerticalAlignment="Center"
                   Text="Ready - enter a target and click Run Scan"/>
        <TextBlock x:Name="TxtNmapPath" Foreground="#3a4f6a" FontSize="10"
                   HorizontalAlignment="Right" VerticalAlignment="Center"/>
      </Grid>
    </Border>

  </Grid>
</Window>
'@

# ── LOAD WINDOW ───────────────────────────────────────────────────────────────
$Reader = [System.Xml.XmlNodeReader]::new($XAML)
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# Fit window inside the usable desktop area (helps on scaling / smaller screens)
$wa = [System.Windows.SystemParameters]::WorkArea
$Window.MaxWidth  = [math]::Max(860, [int]$wa.Width)
$Window.MaxHeight = [math]::Max(580, [int]$wa.Height)
if ($Window.Width  -gt ($wa.Width  - 20)) { $Window.Width  = [math]::Max(860, [int]($wa.Width  - 20)) }
if ($Window.Height -gt ($wa.Height - 20)) { $Window.Height = [math]::Max(580, [int]($wa.Height - 20)) }
$Window.Add_SourceInitialized({
    $wa2 = [System.Windows.SystemParameters]::WorkArea
    if ($this.Left -lt $wa2.Left) { $this.Left = $wa2.Left + 8 }
    if ($this.Top  -lt $wa2.Top)  { $this.Top  = $wa2.Top  + 8 }
    if (($this.Left + $this.Width)  -gt ($wa2.Right - 8))  { $this.Left = [math]::Max($wa2.Left + 8, $wa2.Right  - $this.Width  - 8) }
    if (($this.Top  + $this.Height) -gt ($wa2.Bottom - 8)) { $this.Top  = [math]::Max($wa2.Top  + 8, $wa2.Bottom - $this.Height - 8) }
})

function G($n) { $Window.FindName($n) }

$BtnTabNmap  = G 'BtnTabNmap';  $BtnTabTrace = G 'BtnTabTrace'; $BtnTabHist = G 'BtnTabHist'
$TxtTarget   = G 'TxtTarget';   $TxtPorts    = G 'TxtPorts';    $TxtTimeout = G 'TxtTimeout'
$TxtCmd      = G 'TxtCmd';      $TxtSBar     = G 'TxtSBar';     $TxtNmapPath= G 'TxtNmapPath'
$TxtNmapStatus=G 'TxtNmapStatus'; $TxtStat   = G 'TxtStat';     $TxtTime    = G 'TxtTime'
$TxtTrStat   = G 'TxtTrStat';   $CboTiming   = G 'CboTiming'
$ChkSynS     = G 'ChkSynS';     $ChkTcpC     = G 'ChkTcpC';     $ChkUdp     = G 'ChkUdp'
$ChkAckS     = G 'ChkAckS';     $ChkSV       = G 'ChkSV';       $ChkO       = G 'ChkO'
$ChkSC       = G 'ChkSC';       $ChkA        = G 'ChkA'
$ChkPn       = G 'ChkPn';       $ChkPS       = G 'ChkPS';       $ChkPA      = G 'ChkPA'; $ChkPE = G 'ChkPE'
$ChkV        = G 'ChkV';        $ChkOpen     = G 'ChkOpen';     $ChkRsn     = G 'ChkRsn'; $ChkN = G 'ChkN'
$TxtCustom   = G 'TxtCustom'
$BtnRun      = G 'BtnRun';      $BtnPQ       = G 'BtnPQ';       $BtnPF      = G 'BtnPF'
$BtnPS       = G 'BtnPS';       $BtnPSt      = G 'BtnPSt';      $BtnPW      = G 'BtnPW'; $BtnPN = G 'BtnPN'
$PanelNmap   = G 'PanelNmap';   $PanelTrace  = G 'PanelTrace'
$CboTrMethod = G 'CboTrMethod'; $ChkTrNmap   = G 'ChkTrNmap';   $TxtTrHops  = G 'TxtTrHops'
$TxtTrHopTo  = G 'TxtTrHopTo'; $ChkTrDns    = G 'ChkTrDns'
$BtnTr8      = G 'BtnTr8';      $BtnTr1      = G 'BtnTr1';      $BtnTrGW    = G 'BtnTrGW'
$BtnRunTr    = G 'BtnRunTr';    $TabOut      = G 'TabOut'
$Rtb         = G 'Rtb';         $DgPorts     = G 'DgPorts';     $RtbTr      = G 'RtbTr'
$DgHops      = G 'DgHops';      $LstHist     = G 'LstHist'
$BtnCopyCmd  = G 'BtnCopyCmd';  $BtnCopyOut  = G 'BtnCopyOut';  $BtnSaveOut = G 'BtnSaveOut'
$BtnClearOut = G 'BtnClearOut'; $BtnCopyTr   = G 'BtnCopyTr';   $BtnClearTr = G 'BtnClearTr'
$BtnClrHist  = G 'BtnClrHist';  $BtnStop     = G 'BtnStop'

# ── STATE ─────────────────────────────────────────────────────────────────────
$Script:Scanning  = $false
$Script:LastOut   = ''
$Script:LastTr    = ''
$Script:History   = [System.Collections.Generic.List[string]]::new()

# ── NMAP STATUS ───────────────────────────────────────────────────────────────
if ($NmapPath -and (Test-Path $NmapPath)) {
    $TxtNmapStatus.Text       = 'NMAP READY'
    $TxtNmapStatus.Foreground = [Windows.Media.Brushes]::LightGreen
    $TxtNmapPath.Text         = $NmapPath
} else {
    $TxtNmapStatus.Text = 'NMAP NOT FOUND'
    $TxtNmapStatus.Foreground = [Windows.Media.SolidColorBrush][Windows.Media.ColorConverter]::ConvertFromString('#ff4444')
    $BtnRun.IsEnabled   = $false
    $TxtSBar.Text       = 'nmap.exe not found. Place it in .\nmap\nmap.exe alongside this script.'
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
function Build-Args {
    $a = [System.Collections.Generic.List[string]]::new()
    if ($ChkSynS.IsChecked)  { $a.Add('-sS') }
    if ($ChkTcpC.IsChecked)  { $a.Add('-sT') }
    if ($ChkUdp.IsChecked)   { $a.Add('-sU') }
    if ($ChkAckS.IsChecked)  { $a.Add('-sA') }
    if ($ChkSV.IsChecked)    { $a.Add('-sV') }
    if ($ChkO.IsChecked)     { $a.Add('-O')  }
    if ($ChkSC.IsChecked)    { $a.Add('-sC') }
    if ($ChkA.IsChecked)     { $a.Add('-A')  }
    if ($ChkPn.IsChecked)    { $a.Add('-Pn') }
    if ($ChkPS.IsChecked)    { $a.Add('-PS') }
    if ($ChkPA.IsChecked)    { $a.Add('-PA') }
    if ($ChkPE.IsChecked)    { $a.Add('-PE') }
    if ($ChkV.IsChecked)     { $a.Add('-v')       }
    if ($ChkOpen.IsChecked)  { $a.Add('--open')   }
    if ($ChkRsn.IsChecked)   { $a.Add('--reason') }
    if ($ChkN.IsChecked)     { $a.Add('-n')        }
    $t = @('-T2','-T3','-T4','-T5')[$CboTiming.SelectedIndex]
    if ($t) { $a.Add($t) }
    $p = $TxtPorts.Text.Trim()
    if ($p -eq 'top100')      { $a.Add('--top-ports'); $a.Add('100')  }
    elseif ($p -eq 'top1000') { $a.Add('--top-ports'); $a.Add('1000') }
    elseif ($p)               { $a.Add('-p'); $a.Add($p) }
    $TxtCustom.Text.Trim() -split '\s+' | Where-Object {$_} | ForEach-Object { $a.Add($_) }
    return $a
}

function Update-Preview {
    $tgt = if ($TxtTarget.Text.Trim()) { $TxtTarget.Text.Trim() } else { '<target>' }
    $TxtCmd.Text = "nmap $((Build-Args) -join ' ') $tgt"
}

function Append-Rtb {
    param($rtb, [string]$text, [string]$defaultColor = '#cdd8e8')
    $Window.Dispatcher.Invoke([Action]{
        $para = [Windows.Documents.Paragraph]::new()
        foreach ($line in ($text -split "`n")) {
            $c = $defaultColor
            if ($line -match '\bopen\b')              { $c = '#00e676' }
            elseif ($line -match '\bclosed\b')        { $c = '#ff4444' }
            elseif ($line -match '\bfiltered\b')      { $c = '#ff9800' }
            elseif ($line -match 'Nmap scan report')  { $c = '#80deea' }
            elseif ($line -match 'Host is up')        { $c = '#00e676' }
            elseif ($line -match 'WARNING')           { $c = '#ff9800' }
            $run = [Windows.Documents.Run]::new($line + "`n")
            $run.Foreground = [Windows.Media.SolidColorBrush][Windows.Media.ColorConverter]::ConvertFromString($c)
            $para.Inlines.Add($run)
        }
        $rtb.Document.Blocks.Add($para)
        $rtb.ScrollToEnd()
    })
}

function Parse-Ports { param([string]$out)
    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $out -split "`n" | ForEach-Object {
        if ($_ -match '^(\d+)/(tcp|udp)\s+(open\S*)\s+(\S+)(.*)') {
            $rows.Add([PSCustomObject]@{ Port=$Matches[1]; Proto=$Matches[2]; State=$Matches[3]; Service=$Matches[4]; Info=$Matches[5].Trim() })
        }
    }
    $Window.Dispatcher.Invoke([Action]{ $DgPorts.ItemsSource = $rows })
}

function Parse-Hops { param([string]$out)
    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $out -split "`n" | ForEach-Object {
        if ($_ -match '^\s*(\d+)\s+(.+)$') {
            $rest = $Matches[2].Trim()
            $ips  = [regex]::Matches($rest,'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b') | ForEach-Object {$_.Value}
            $rtts = [regex]::Matches($rest,'[\d.]+\s*ms') | ForEach-Object {$_.Value}
            $stars= ($rest.ToCharArray() | Where-Object {$_ -eq '*'}).Count
            $rows.Add([PSCustomObject]@{
                Hop = $Matches[1]
                IP  = if ($stars -ge 3) {'* * *'} elseif ($ips) {$ips[0]} else {'?'}
                RTT = if ($stars -ge 3) {'—'} else {$rtts -join ' / '}
                Raw = $rest
            })
        }
    }
    $Window.Dispatcher.Invoke([Action]{ $DgHops.ItemsSource = $rows })
}

function Add-Hist { param([string]$e)
    $Script:History.Insert(0,$e)
    if ($Script:History.Count -gt 50) { $Script:History.RemoveAt(50) }
    $Window.Dispatcher.Invoke([Action]{
        $LstHist.Items.Clear()
        $Script:History | ForEach-Object { $LstHist.Items.Add($_) }
    })
}

function Set-Stat { param([string]$txt,[string]$col='#5a7294')
    $Window.Dispatcher.Invoke([Action]{
        $TxtStat.Text = $txt
        $TxtStat.Foreground=[Windows.Media.SolidColorBrush][Windows.Media.ColorConverter]::ConvertFromString($col)
    })
}

function Clear-Checks {
    $ChkSynS.IsChecked=$false; $ChkTcpC.IsChecked=$false; $ChkUdp.IsChecked=$false; $ChkAckS.IsChecked=$false
    $ChkSV.IsChecked=$false;   $ChkO.IsChecked=$false;    $ChkSC.IsChecked=$false;  $ChkA.IsChecked=$false
    $ChkPn.IsChecked=$false;   $ChkPS.IsChecked=$false;   $ChkPA.IsChecked=$false;  $ChkPE.IsChecked=$false
    $ChkV.IsChecked=$false;    $ChkOpen.IsChecked=$false; $ChkRsn.IsChecked=$false; $ChkN.IsChecked=$false
    $TxtPorts.Text=''; $TxtCustom.Text=''
}

function Switch-Tab { param([string]$tab)
    $acc = [Windows.Media.SolidColorBrush][Windows.Media.ColorConverter]::ConvertFromString('#00c8ff')
    $dim = [Windows.Media.SolidColorBrush][Windows.Media.ColorConverter]::ConvertFromString('#5a7294')
    foreach ($b in @($BtnTabNmap,$BtnTabTrace,$BtnTabHist)) {
        $b.BorderBrush=$([Windows.Media.Brushes]::Transparent); $b.Foreground=$dim
    }
    switch ($tab) {
        'nmap'  { $BtnTabNmap.BorderBrush=$acc;  $BtnTabNmap.Foreground=$acc;  $PanelNmap.Visibility='Visible'; $PanelTrace.Visibility='Collapsed'; $TabOut.SelectedIndex=0 }
        'trace' { $BtnTabTrace.BorderBrush=$acc; $BtnTabTrace.Foreground=$acc; $PanelNmap.Visibility='Collapsed'; $PanelTrace.Visibility='Visible'; $TabOut.SelectedIndex=2 }
        'hist'  { $BtnTabHist.BorderBrush=$acc;  $BtnTabHist.Foreground=$acc;  $PanelNmap.Visibility='Visible'; $PanelTrace.Visibility='Collapsed'; $TabOut.SelectedIndex=3 }
    }
}

# ── PRESETS ───────────────────────────────────────────────────────────────────
$BtnPQ.Add_Click({  Clear-Checks; $ChkTcpC.IsChecked=$true; $ChkPn.IsChecked=$true; $ChkV.IsChecked=$true; $CboTiming.SelectedIndex=2; $TxtPorts.Text='top100'; Update-Preview })
$BtnPF.Add_Click({  Clear-Checks; $ChkTcpC.IsChecked=$true; $ChkPn.IsChecked=$true; $ChkV.IsChecked=$true; $ChkOpen.IsChecked=$true; $CboTiming.SelectedIndex=2; $TxtPorts.Text='1-65535'; Update-Preview })
$BtnPS.Add_Click({  Clear-Checks; $ChkTcpC.IsChecked=$true; $ChkSV.IsChecked=$true; $ChkO.IsChecked=$true; $ChkSC.IsChecked=$true; $ChkPn.IsChecked=$true; $ChkV.IsChecked=$true; $CboTiming.SelectedIndex=2; Update-Preview })
$BtnPSt.Add_Click({ Clear-Checks; $ChkSynS.IsChecked=$true; $ChkPn.IsChecked=$true; $CboTiming.SelectedIndex=0; Update-Preview })
$BtnPW.Add_Click({  Clear-Checks; $ChkTcpC.IsChecked=$true; $ChkPn.IsChecked=$true; $ChkV.IsChecked=$true; $CboTiming.SelectedIndex=2; $TxtPorts.Text='80,443,8080,8443,8000,8888'; Update-Preview })
$BtnPN.Add_Click({  Clear-Checks; $ChkPE.IsChecked=$true; $ChkPS.IsChecked=$true; $ChkV.IsChecked=$true; $CboTiming.SelectedIndex=2; $TxtCustom.Text='-sn'; Update-Preview })

# ── LIVE PREVIEW ──────────────────────────────────────────────────────────────
@($ChkSynS,$ChkTcpC,$ChkUdp,$ChkAckS,$ChkSV,$ChkO,$ChkSC,$ChkA,
  $ChkPn,$ChkPS,$ChkPA,$ChkPE,$ChkV,$ChkOpen,$ChkRsn,$ChkN) | ForEach-Object {
    $_.Add_Checked({Update-Preview}); $_.Add_Unchecked({Update-Preview})
}
$TxtTarget.Add_TextChanged({Update-Preview})
$TxtPorts.Add_TextChanged({Update-Preview})
$TxtCustom.Add_TextChanged({Update-Preview})
$CboTiming.Add_SelectionChanged({Update-Preview})

# ── RUN SCAN ──────────────────────────────────────────────────────────────────
$Script:ScanJob     = $null
$Script:ScanTimer   = $null
$Script:ScanStart   = $null
$Script:ScanTarget  = $null
$Script:ScanTimeout = 120

$Script:ResetScanUi = {
    $Script:Scanning   = $false
    $BtnRun.IsEnabled  = $true
    $BtnRun.Content    = 'Run Scan'
    $BtnStop.IsEnabled = $false
}

$BtnStop.Add_Click({
    if ($Script:ScanTimer) {
        $Script:ScanTimer.Stop()
        $Script:ScanTimer = $null
    }
    if ($Script:ScanJob) {
        try { Stop-Job -Job $Script:ScanJob -ErrorAction SilentlyContinue | Out-Null } catch {}
        try { Remove-Job -Job $Script:ScanJob -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        $Script:ScanJob = $null
    }
    & $Script:ResetScanUi
    Set-Stat 'STOPPED' '#ff9800'
    $TxtSBar.Text = 'Scan stopped by user.'
    Append-Rtb $Rtb "`n[Scan stopped by user]" '#ff9800'
})

$BtnRun.Add_Click({
    if ($Script:Scanning) { return }

    $Script:ScanTarget = $TxtTarget.Text.Trim()
    if (-not $Script:ScanTarget) {
        [System.Windows.MessageBox]::Show('Enter a target host or IP.','NetProbe','OK','Warning')
        return
    }

    $Script:Scanning  = $true
    $Script:ScanStart = [datetime]::Now
    $BtnRun.IsEnabled = $false
    $BtnRun.Content   = 'Scanning...'
    $BtnStop.IsEnabled = $true
    $Rtb.Document.Blocks.Clear()
    $DgPorts.ItemsSource = $null
    Set-Stat 'SCANNING...' '#00c8ff'
    $TxtTime.Text = ''
    $TxtSBar.Text = "Scanning $($Script:ScanTarget) ..."

    $nmapArgs = Build-Args
    $Script:ScanTimeout = [int]($TxtTimeout.Text -replace '\D','')
    if ($Script:ScanTimeout -lt 5) { $Script:ScanTimeout = 120 }

    $Script:ScanJob = Start-Job -ScriptBlock {
        param($nm,$a,$t)
        & $nm @a $t 2>&1
    } -ArgumentList $NmapPath,$nmapArgs,$Script:ScanTarget

    $Script:ScanTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:ScanTimer.Interval = [TimeSpan]::FromMilliseconds(800)
    $Script:ScanTimer.Add_Tick({
        if (-not $Script:Scanning) {
            if ($Script:ScanTimer) {
                $Script:ScanTimer.Stop()
                $Script:ScanTimer = $null
            }
            return
        }

        if ($null -eq $Script:ScanStart) {
            $Script:ScanStart = [datetime]::Now
        }

        $elapsed = [int](([datetime]::Now - $Script:ScanStart).TotalSeconds)
        $TxtSBar.Text = "Scanning $($Script:ScanTarget) ... ${elapsed}s"

        if ($null -eq $Script:ScanJob) {
            if ($Script:ScanTimer) {
                $Script:ScanTimer.Stop()
                $Script:ScanTimer = $null
            }
            & $Script:ResetScanUi
            return
        }

        $timedOut = $elapsed -ge $Script:ScanTimeout
        $finished = $Script:ScanJob.State -in @('Completed','Failed','Stopped')

        if (-not $timedOut -and -not $finished) {
            return
        }

        if ($Script:ScanTimer) {
            $Script:ScanTimer.Stop()
            $Script:ScanTimer = $null
        }

        $out = ''
        if ($timedOut -and $Script:ScanJob.State -notin @('Completed','Failed','Stopped')) {
            try { Stop-Job -Job $Script:ScanJob -ErrorAction SilentlyContinue | Out-Null } catch {}
            $out = "Scan timed out after $($Script:ScanTimeout) seconds."
            try { Remove-Job -Job $Script:ScanJob -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        } else {
            try {
                $out = Receive-Job -Job $Script:ScanJob -Keep -ErrorAction Stop 2>&1 | Out-String
            } catch {
                $out = "Scan error: $($_.Exception.Message)"
            } finally {
                try { Remove-Job -Job $Script:ScanJob -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
            }
        }

        $Script:ScanJob = $null
        $Script:LastOut = $out
        Append-Rtb $Rtb $out
        Parse-Ports $out

        $elapsed2 = [int](([datetime]::Now - $Script:ScanStart).TotalSeconds)
        $TxtTime.Text = "${elapsed2}s"

        if ($timedOut) {
            Set-Stat 'TIMEOUT' '#ff4444'
            Add-Hist "ERR $(Get-Date -f 'HH:mm:ss')  $($Script:ScanTarget)  [timeout]"
            $TxtSBar.Text = "Scan timed out - $($Script:ScanTarget) (${elapsed2}s)"
        } else {
            Set-Stat 'COMPLETE' '#00e676'
            Add-Hist "OK  $(Get-Date -f 'HH:mm:ss')  $($Script:ScanTarget)  [${elapsed2}s]"
            $TxtSBar.Text = "Scan complete - $($Script:ScanTarget) (${elapsed2}s)"
        }

        & $Script:ResetScanUi
    })
    $Script:ScanTimer.Start()
})

# ── TRACEROUTE ────────────────────────────────────────────────────────────────
$Script:TraceJob    = $null
$Script:TraceTimer  = $null
$Script:TraceStart  = $null
$Script:TraceTarget = $null
$Script:TraceMaxSec = 120

$BtnRunTr.Add_Click({
    $Script:TraceTarget = $TxtTarget.Text.Trim()
    if (-not $Script:TraceTarget) {
        [System.Windows.MessageBox]::Show('Enter a target in the HOST field.','NetProbe','OK','Warning')
        return
    }

    $RtbTr.Document.Blocks.Clear()
    $DgHops.ItemsSource = $null
    $TxtTrStat.Text = 'TRACING...'
    $BtnRunTr.IsEnabled = $false

    $hops  = [int]($TxtTrHops.Text  -replace '\D',''); if ($hops -lt 1) { $hops = 30 }
    $hopTo = [int]($TxtTrHopTo.Text -replace '\D',''); if ($hopTo -lt 1) { $hopTo = 3 }
    $useNm = $ChkTrNmap.IsChecked

    if ($useNm) {
        $cmd = $NmapPath
        $trArgs = @('--traceroute','-Pn','-n','--max-retries','1','-T4',$Script:TraceTarget)
    } else {
        $cmd = 'tracert'
        $trArgs = @('-h',$hops,'-w',($hopTo * 1000),$Script:TraceTarget)
    }

    $Script:TraceStart = [datetime]::Now
    $Script:TraceJob = Start-Job -ScriptBlock {
        param($c,$a)
        & $c @a 2>&1
    } -ArgumentList $cmd,$trArgs

    $Script:TraceTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:TraceTimer.Interval = [TimeSpan]::FromMilliseconds(600)
    $Script:TraceTimer.Add_Tick({
        if ($null -eq $Script:TraceStart) {
            $Script:TraceStart = [datetime]::Now
        }

        $elapsed = [int](([datetime]::Now - $Script:TraceStart).TotalSeconds)
        $jobState = if ($Script:TraceJob) { $Script:TraceJob.State } else { 'Stopped' }
        $timedOut = $elapsed -ge $Script:TraceMaxSec
        $finished = $jobState -in @('Completed','Failed','Stopped')

        if (-not $timedOut -and -not $finished) {
            return
        }

        if ($Script:TraceTimer) {
            $Script:TraceTimer.Stop()
            $Script:TraceTimer = $null
        }

        $out = ''
        if ($Script:TraceJob) {
            if ($timedOut -and $Script:TraceJob.State -notin @('Completed','Failed','Stopped')) {
                try { Stop-Job -Job $Script:TraceJob -ErrorAction SilentlyContinue | Out-Null } catch {}
                $out = "Traceroute timed out after $($Script:TraceMaxSec) seconds."
                try { Remove-Job -Job $Script:TraceJob -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
            } else {
                try {
                    $out = Receive-Job -Job $Script:TraceJob -Keep -ErrorAction Stop 2>&1 | Out-String
                } catch {
                    $out = "Traceroute error: $($_.Exception.Message)"
                } finally {
                    try { Remove-Job -Job $Script:TraceJob -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
                }
            }
        } else {
            $out = 'Traceroute job was not available.'
        }

        $Script:TraceJob = $null
        $Script:LastTr = $out
        Append-Rtb $RtbTr $out
        Parse-Hops $out

        $elapsed2 = [int](([datetime]::Now - $Script:TraceStart).TotalSeconds)
        if ($timedOut) {
            $TxtTrStat.Text = "TIMEOUT (${elapsed2}s)"
        } else {
            $TxtTrStat.Text = "COMPLETE (${elapsed2}s)"
        }
        $BtnRunTr.IsEnabled = $true
        Add-Hist "TR  $(Get-Date -f 'HH:mm:ss')  $($Script:TraceTarget)  [traceroute]"
    })
    $Script:TraceTimer.Start()
})

# ── QUICK TRACE ───────────────────────────────────────────────────────────────
$BtnTr8.Add_Click({  $TxtTarget.Text='8.8.8.8'; Switch-Tab 'trace' })
$BtnTr1.Add_Click({  $TxtTarget.Text='1.1.1.1'; Switch-Tab 'trace' })
$BtnTrGW.Add_Click({
    $gw = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Sort-Object RouteMetric | Select-Object -First 1).NextHop
    $TxtTarget.Text = if ($gw) { $gw } else { '192.168.1.1' }
    Switch-Tab 'trace'
})

# ── COPY / SAVE / CLEAR ───────────────────────────────────────────────────────
$BtnCopyCmd.Add_Click({ [System.Windows.Clipboard]::SetText($TxtCmd.Text) })
$BtnCopyOut.Add_Click({ [System.Windows.Clipboard]::SetText($Script:LastOut) })
$BtnCopyTr.Add_Click({  [System.Windows.Clipboard]::SetText($Script:LastTr)  })
$BtnSaveOut.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'Text (*.txt)|*.txt|All (*.*)|*.*'
    $dlg.FileName = "nmap_$(Get-Date -f 'yyyyMMdd_HHmm').txt"
    if ($dlg.ShowDialog()) { $Script:LastOut | Set-Content $dlg.FileName -Encoding UTF8; $TxtSBar.Text="Saved: $($dlg.FileName)" }
})
$BtnClearOut.Add_Click({
    $Rtb.Document.Blocks.Clear(); $DgPorts.ItemsSource=$null
    $TxtStat.Text=''; $TxtTime.Text=''; $Script:LastOut=''
})
$BtnClearTr.Add_Click({
    $RtbTr.Document.Blocks.Clear(); $DgHops.ItemsSource=$null
    $TxtTrStat.Text=''; $Script:LastTr=''
})
$BtnClrHist.Add_Click({ $Script:History.Clear(); $LstHist.Items.Clear() })
$LstHist.Add_MouseDoubleClick({
    $sel = $LstHist.SelectedItem
    if ($sel) {
        $ip = [regex]::Match($sel,'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b').Value
        if ($ip) { $TxtTarget.Text=$ip }
    }
})

# ── NAV ───────────────────────────────────────────────────────────────────────
$BtnTabNmap.Add_Click({  Switch-Tab 'nmap'  })
$BtnTabTrace.Add_Click({ Switch-Tab 'trace' })
$BtnTabHist.Add_Click({  Switch-Tab 'hist'  })
$Window.Add_KeyDown({
    if ($_.Key -eq 'Return' -and $_.KeyboardDevice.Modifiers -eq 'Control') {
        $BtnRun.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
})

# ── LAUNCH ────────────────────────────────────────────────────────────────────
Update-Preview
Switch-Tab 'nmap'
$TxtTarget.Focus()
$Window.ShowDialog() | Out-Null
