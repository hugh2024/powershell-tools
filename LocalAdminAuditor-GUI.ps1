#region Prerequisite Checks
Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Install-RSAT {
    Write-Host "  Installing RSAT ActiveDirectory module..." -ForegroundColor Cyan
    # Windows 10/11 - Windows Capability
    $cap = Get-WindowsCapability -Online -Name "Rsat.ActiveDirectory*" -ErrorAction SilentlyContinue |
           Where-Object State -ne "Installed" | Select-Object -First 1
    if ($cap) {
        Add-WindowsCapability -Online -Name $cap.Name -ErrorAction Stop | Out-Null
        return
    }
    # Windows Server - Install-WindowsFeature
    if (Get-Command Install-WindowsFeature -ErrorAction SilentlyContinue) {
        Install-WindowsFeature -Name "RSAT-AD-PowerShell" -ErrorAction Stop | Out-Null
        return
    }
    throw "Could not find a supported install method for RSAT on this OS."
}

function Enable-WinRM {
    Write-Host "  Enabling WinRM / PowerShell Remoting..." -ForegroundColor Cyan
    Enable-PSRemoting -Force -ErrorAction Stop | Out-Null
}

# --- Collect what's missing ---
$missing = [System.Collections.Generic.List[string]]::new()
$psOk    = $PSVersionTable.PSVersion -ge [Version]"5.1"
$rsatOk  = [bool](Get-Module -ListAvailable -Name ActiveDirectory)
$winrmOk = (Get-Service WinRM -ErrorAction SilentlyContinue).Status -eq 'Running'

if (-not $psOk)   { $missing.Add("PowerShell 5.1 or later") }
if (-not $rsatOk) { $missing.Add("RSAT - ActiveDirectory Module") }
if (-not $winrmOk){ $missing.Add("WinRM / PowerShell Remoting") }

# --- If nothing missing, move on ---
if ($missing.Count -eq 0) {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
}
else {
    # Build the prompt message
    $itemList = ($missing | ForEach-Object { "  - $_" }) -join "`n"
    $prompt = "The following prerequisites are missing:`n`n$itemList`n`n"

    if (-not $psOk) {
        $prompt += "PowerShell 5.1 cannot be auto-installed.`n"
        $prompt += "Download WMF 5.1 from: https://aka.ms/wmf51download`n`n"
    }

    if ($rsatOk -eq $false -or $winrmOk -eq $false) {
        $prompt += "Click YES to install/enable the missing items automatically.`n"
        $prompt += "Click NO to exit (you can install them manually and relaunch)."
    }

    $title  = "Missing Prerequisites ($($missing.Count) item$(if($missing.Count -gt 1){'s'}))"
    $buttons = if ($psOk) { "YesNo" } else { "OK" }
    $result = [System.Windows.MessageBox]::Show($prompt, $title, $buttons, "Warning")

    # Hard exit if PS version is too old - can't fix at runtime
    if (-not $psOk) { exit 1 }

    if ($result -ne "Yes") {
        [System.Windows.MessageBox]::Show(
            "Manual install steps:`n`n" +
            "RSAT:`n  Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'`n`n" +
            "WinRM (run as Administrator):`n  Enable-PSRemoting -Force`n`n" +
            "Then relaunch this script.",
            "Manual Setup", "OK", "Information") | Out-Null
        exit 0
    }

    # --- User said Yes - check admin rights before attempting ---
    if (-not (Test-IsAdmin)) {
        [System.Windows.MessageBox]::Show(
            "Administrator rights are required to install prerequisites.`n`n" +
            "Please relaunch PowerShell as Administrator:`n" +
            "  Right-click PowerShell > Run as Administrator`n`n" +
            "Then run this script again.",
            "Administrator Required", "OK", "Warning") | Out-Null
        exit 1
    }

    # --- Install / enable each missing item ---
    $errors = [System.Collections.Generic.List[string]]::new()

    if (-not $rsatOk) {
        try   { Install-RSAT }
        catch { $errors.Add("RSAT install failed: $($_.Exception.Message)") }
    }

    if (-not $winrmOk) {
        try   { Enable-WinRM }
        catch { $errors.Add("WinRM enable failed: $($_.Exception.Message)") }
    }

    # --- Report results ---
    if ($errors.Count -gt 0) {
        $errText = $errors -join "`n`n"
        [System.Windows.MessageBox]::Show(
            "Some items could not be configured automatically:`n`n$errText`n`n" +
            "Please resolve these manually and relaunch.",
            "Setup Incomplete", "OK", "Warning") | Out-Null
        exit 1
    }

    # Verify RSAT actually landed
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.MessageBox]::Show(
            "RSAT was installed but the ActiveDirectory module is not yet visible.`n`n" +
            "This is normal - PowerShell caches the module list at startup.`n`n" +
            "Please close this window and relaunch the script.`n" +
            "The tool will start normally on the next run.",
            "Relaunch Required", "OK", "Information") | Out-Null
        exit 0
    }

    Import-Module ActiveDirectory -ErrorAction SilentlyContinue

    [System.Windows.MessageBox]::Show(
        "All prerequisites are now configured.`n`nThe tool will launch now.",
        "Setup Complete", "OK", "Information") | Out-Null
}

# --- Domain membership soft-warning (non-blocking) ---
$domainJoined = (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue).PartOfDomain
if (-not $domainJoined) {
    $r = [System.Windows.MessageBox]::Show(
        "This machine is not joined to an Active Directory domain.`n`n" +
        "This tool is designed for domain environments. Without a domain you won't`n" +
        "be able to enumerate AD computers, but you can still use 'Specific Computers'`n" +
        "mode to scan individual machines you have access to.`n`n" +
        "Continue anyway?",
        "Not Domain-Joined", "YesNo", "Warning")
    if ($r -ne "Yes") { exit 0 }
}
#endregion Prerequisite Checks

<#
.SYNOPSIS
    GUI-based Local Administrator Group Auditor for Active Directory environments.

.DESCRIPTION
    A comprehensive WPF-based GUI application that audits local Administrators group
    membership across Windows Servers and Workstations in your Active Directory domain.

    Features:
    - Unified scanning of Servers, Workstations, Both, or Specific Computers
    - Real-time progress tracking with cancellation and resume support
    - Configurable exclusion patterns for approved accounts
    - Computer exclusions for honeypots, test systems, and dev machines
    - Sortable and filterable results grid with risk scoring
    - Unified Scanner tab with live statistics and risk badges
    - Export to CSV, HTML, PDF, or clipboard
    - Compliance reports (PCI-DSS, HIPAA, SOX, CIS)
    - Trend dashboard with historical analysis
    - Parallel scanning with throttling control and runspace pools
    - Remote remediation with multi-select support and double-confirmation
    - Scan comparison/diff mode with NEW/EXISTING/REMOVED badges

.NOTES
    GitHub  : https://Hugh Gozlou  |  github.com/hugh2024/powershell-tools
    Version : 9.0
    License : MIT
    Requires: ActiveDirectory module, PowerShell remoting enabled on targets

    MIT License
    Copyright (c) 2025 hugh2024
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.

    Version History:
    - 2.0: Initial GUI version
    - 2.1: Added Risk Severity scoring with color-coded badges
    - 2.2: Added JSON persistence for settings and scan history
    - 2.3: Added History tab with NEW/EXISTING badges
    - 2.4: Added context menus, keyboard shortcuts, sound notifications, status bar
    - 2.5: Added Group-By views and Approved Service Accounts for risk reduction
    - 2.6: Fixed UI freeze after scan completion (async Dispatcher, background priority)
    - 2.7: Fixed 5% scan freeze (pre-cached UI settings), added attribution footer
    - 2.8: Performance optimizations (HashSet, hashtable caching, single-pass stats)
    - 2.9: Fixed results not showing (completion handler on main thread for function access)
    - 3.0: Added Computer Exclusions feature for honeypots and test systems
    - 3.1: Added comprehensive tooltips, Pro-Tips bar, removed Status column
    - 3.2: Scan Resume, Remote Remediation (multi-select), Trend Dashboard, PDF Export,
           Compliance Reports, Specific Computers mode, Compare/Diff, Quick Rescan,
           Parallel AD queries, Window size bounds checking, Double-click detail pane
    - 3.3: Credential Authentication - Authenticate button for alternate domain admin credentials,
           credential validation via Get-ADDomain, splatted injection at all remote operations,
           credential error detection, button disabled during active scans
    - 9.0: Redesigned blue/cyan security theme, MIT license, public release cleanup,
           improved remediation confirmation with activity logging, security hardening
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#region XAML Definition
$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Local Administrator Auditor v9.0"
        Height="800" Width="1200"
        MinHeight="600" MinWidth="900"
        WindowStartupLocation="CenterScreen"
        Background="#0D1117">

    <Window.Resources>
        <!-- Color Definitions - Dark Blue/Cyan Security Theme -->
        <Color x:Key="BackgroundColor">#0D1117</Color>
        <Color x:Key="CardColor">#161B22</Color>
        <Color x:Key="CardHoverColor">#1C2333</Color>
        <Color x:Key="AccentColor">#0EA5E9</Color>
        <Color x:Key="AccentHoverColor">#38BDF8</Color>
        <Color x:Key="SuccessColor">#10B981</Color>
        <Color x:Key="WarningColor">#F59E0B</Color>
        <Color x:Key="DangerColor">#EF4444</Color>
        <Color x:Key="TextPrimaryColor">#E6EDF3</Color>
        <Color x:Key="TextSecondaryColor">#8B949E</Color>
        <Color x:Key="BorderColor">#30363D</Color>

        <!-- Risk Level Colors (8:1+ contrast) -->
        <Color x:Key="RiskHighColor">#FF6B6B</Color>
        <Color x:Key="RiskMediumColor">#FFD93D</Color>
        <Color x:Key="RiskLowColor">#4ECDC4</Color>
        <Color x:Key="RiskInfoColor">#8B949E</Color>

        <!-- Risk Level Background Colors -->
        <Color x:Key="RiskHighBgColor">#3D1212</Color>
        <Color x:Key="RiskMediumBgColor">#3D2E05</Color>
        <Color x:Key="RiskLowBgColor">#0D2F2D</Color>

        <!-- Brushes -->
        <SolidColorBrush x:Key="BackgroundBrush" Color="{StaticResource BackgroundColor}"/>
        <SolidColorBrush x:Key="CardBrush" Color="{StaticResource CardColor}"/>
        <SolidColorBrush x:Key="CardBrushHover" Color="{StaticResource CardHoverColor}"/>
        <SolidColorBrush x:Key="AccentBrush" Color="{StaticResource AccentColor}"/>
        <SolidColorBrush x:Key="AccentBrushHover" Color="{StaticResource AccentHoverColor}"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="{StaticResource SuccessColor}"/>
        <SolidColorBrush x:Key="WarningBrush" Color="{StaticResource WarningColor}"/>
        <SolidColorBrush x:Key="DangerBrush" Color="{StaticResource DangerColor}"/>
        <SolidColorBrush x:Key="TextPrimaryBrush" Color="{StaticResource TextPrimaryColor}"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="{StaticResource TextSecondaryColor}"/>
        <SolidColorBrush x:Key="TextMutedBrush" Color="#484F58"/>
        <SolidColorBrush x:Key="BorderBrush" Color="{StaticResource BorderColor}"/>

        <!-- Risk Level Brushes -->
        <SolidColorBrush x:Key="RiskHighBrush" Color="{StaticResource RiskHighColor}"/>
        <SolidColorBrush x:Key="RiskMediumBrush" Color="{StaticResource RiskMediumColor}"/>
        <SolidColorBrush x:Key="RiskLowBrush" Color="{StaticResource RiskLowColor}"/>
        <SolidColorBrush x:Key="RiskInfoBrush" Color="{StaticResource RiskInfoColor}"/>
        <SolidColorBrush x:Key="RiskHighBgBrush" Color="{StaticResource RiskHighBgColor}"/>
        <SolidColorBrush x:Key="RiskMediumBgBrush" Color="{StaticResource RiskMediumBgColor}"/>
        <SolidColorBrush x:Key="RiskLowBgBrush" Color="{StaticResource RiskLowBgColor}"/>

        <!-- Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="20,12"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}"
                                BorderBrush="#1D3A4A"
                                BorderThickness="1">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource AccentBrushHover}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="{StaticResource CardBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource CardBrushHover}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#7F1D1D"/>
            <Setter Property="Foreground" Value="#FCA5A5"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource DangerBrush}"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Card Style -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="{StaticResource CardBrush}"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="20"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>

        <!-- TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#0D1117"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,10"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CaretBrush" Value="{StaticResource AccentBrush}"/>
            <Setter Property="SelectionBrush" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ComboBox Toggle Button -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="30"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2"
                        Background="#0D1117"
                        BorderBrush="{StaticResource BorderBrush}"
                        BorderThickness="1"
                        CornerRadius="5"/>
                <Border Grid.Column="0" Margin="1"
                        Background="#0D1117"
                        CornerRadius="4,0,0,4"/>
                <Path x:Name="Arrow" Grid.Column="1"
                      Fill="{StaticResource TextPrimaryBrush}"
                      HorizontalAlignment="Center"
                      VerticalAlignment="Center"
                      Data="M 0 0 L 4 4 L 8 0 Z"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="{StaticResource CardBrushHover}"/>
                    <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <!-- ComboBox TextBox -->
        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="TextBox">
            <Border x:Name="PART_ContentHost" Focusable="False" Background="{TemplateBinding Background}"/>
        </ControlTemplate>

        <!-- ComboBox Style -->
        <Style TargetType="ComboBox">
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="MinHeight" Value="38"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton"
                                          Template="{StaticResource ComboBoxToggleButton}"
                                          Grid.Column="2"
                                          Focusable="False"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter Name="ContentSite"
                                              IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="12,3,30,3"
                                              VerticalAlignment="Center"
                                              HorizontalAlignment="Left"/>
                            <TextBox x:Name="PART_EditableTextBox"
                                     Style="{x:Null}"
                                     Template="{StaticResource ComboBoxTextBox}"
                                     HorizontalAlignment="Left"
                                     VerticalAlignment="Center"
                                     Margin="3,3,23,3"
                                     Focusable="True"
                                     Background="Transparent"
                                     Foreground="{StaticResource TextPrimaryBrush}"
                                     Visibility="Hidden"
                                     IsReadOnly="{TemplateBinding IsReadOnly}"/>
                            <Popup Name="Popup"
                                   Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True"
                                   Focusable="False"
                                   PopupAnimation="Slide">
                                <Grid Name="DropDown"
                                      SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}"
                                      MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder"
                                            Background="{StaticResource CardBrush}"
                                            BorderThickness="1"
                                            BorderBrush="{StaticResource AccentBrush}"
                                            CornerRadius="5"
                                            Margin="0,2,0,0"/>
                                    <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ComboBoxItem Style -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border Name="Border"
                                Padding="{TemplateBinding Padding}"
                                Background="{TemplateBinding Background}"
                                CornerRadius="4">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource CardBrushHover}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Margin" Value="0,5"/>
        </Style>

        <!-- DataGrid Styles -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="RowBackground" Value="Transparent"/>
            <Setter Property="AlternatingRowBackground" Value="#12181F"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
            <Setter Property="CanUserResizeRows" Value="False"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CellStyle">
                <Setter.Value>
                    <Style TargetType="DataGridCell">
                        <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
                        <Setter Property="Padding" Value="12,8"/>
                        <Setter Property="BorderThickness" Value="0"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="DataGridCell">
                                    <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                                        <ContentPresenter VerticalAlignment="Center"/>
                                    </Border>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                        <Style.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
                                <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#0C1A24"/>
            <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="12,10"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="BorderBrush" Value="#0EA5E9"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Margin" Value="0,1"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource CardBrushHover}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- TabControl Style -->
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextSecondaryBrush}"/>
            <Setter Property="Padding" Value="18,11"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Border" Background="Transparent" Padding="{TemplateBinding Padding}"
                                BorderThickness="0,0,0,2" BorderBrush="Transparent" Margin="0,0,2,0">
                            <ContentPresenter ContentSource="Header"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource CardBrush}"/>
                                <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                                <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
                                <Setter TargetName="Border" Property="Background" Value="#0C1A24"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ProgressBar Style -->
        <Style TargetType="ProgressBar">
            <Setter Property="Height" Value="8"/>
            <Setter Property="Background" Value="{StaticResource CardBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ProgressBar">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <Border x:Name="PART_Track">
                                <Border x:Name="PART_Indicator"
                                        Background="{TemplateBinding Foreground}"
                                        CornerRadius="4"
                                        HorizontalAlignment="Left"/>
                            </Border>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ScrollBar Style for dark theme -->
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="Transparent"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Margin="0,0,0,20" Background="{StaticResource CardBrush}"
                CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1">
            <Grid>
                <!-- Cyan accent bar on left edge -->
                <Border Width="4" HorizontalAlignment="Left" CornerRadius="8,0,0,8">
                    <Border.Background>
                        <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                            <GradientStop Color="#0EA5E9" Offset="0"/>
                            <GradientStop Color="#0284C7" Offset="1"/>
                        </LinearGradientBrush>
                    </Border.Background>
                </Border>
                <Grid Margin="20,14,20,14">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- Shield icon -->
                    <Border Grid.Column="0" Width="44" Height="44" CornerRadius="8" Margin="0,0,16,0">
                        <Border.Background>
                            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                                <GradientStop Color="#0C2D3F" Offset="0"/>
                                <GradientStop Color="#0EA5E9" Offset="1"/>
                            </LinearGradientBrush>
                        </Border.Background>
                        <TextBlock Text="LA" FontSize="15" FontWeight="Bold"
                                   HorizontalAlignment="Center" VerticalAlignment="Center"
                                   Foreground="White" FontFamily="Consolas"/>
                    </Border>

                    <StackPanel Grid.Column="1" VerticalAlignment="Center">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Local Administrator Auditor"
                                       FontSize="22" FontWeight="Bold"
                                       Foreground="{StaticResource TextPrimaryBrush}"/>
                            <Border Background="#0C2D3F" CornerRadius="4" Padding="8,3" Margin="12,3,0,0"
                                    BorderBrush="{StaticResource AccentBrush}" BorderThickness="1">
                                <TextBlock Text="v9.0" FontSize="11" FontWeight="Bold" Foreground="{StaticResource AccentBrush}"/>
                            </Border>
                        </StackPanel>
                        <TextBlock Text="Audit local Administrators group membership across your Active Directory domain"
                                   FontSize="13"
                                   Foreground="{StaticResource TextSecondaryBrush}"
                                   Margin="0,4,0,0"/>
                    </StackPanel>

                    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                        <Button x:Name="btnHelp" Content="?" Style="{StaticResource SecondaryButton}" Width="36" Padding="0,10"
                                ToolTip="Show help and keyboard shortcuts" FontSize="14" FontWeight="Bold"/>
                    </StackPanel>
                </Grid>
            </Grid>
        </Border>

        <!-- Main Content -->
        <TabControl Grid.Row="1" x:Name="MainTabControl">

            <!-- Scanner Tab (unified scan controls, stats, and results) -->
            <TabItem Header="Scanner" ToolTip="Scan computers and view results (Ctrl+1)">
                <Grid Margin="0,20,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>  <!-- Row 0: Resume Banner (collapsed) -->
                        <RowDefinition Height="Auto"/>  <!-- Row 1: Scan Configuration -->
                        <RowDefinition Height="Auto"/>  <!-- Row 2: Specific Computers Panel (collapsed) -->
                        <RowDefinition Height="Auto"/>  <!-- Row 3: Progress Section (collapsed) -->
                        <RowDefinition Height="Auto"/>  <!-- Row 4: Statistics -->
                        <RowDefinition Height="Auto"/>  <!-- Row 5: Filter Bar -->
                        <RowDefinition Height="*"/>     <!-- Row 6: Results Grid -->
                    </Grid.RowDefinitions>

                    <!-- Row 0: Resume Interrupted Scan Banner (Collapsed by default) -->
                    <Border Grid.Row="0" x:Name="resumeScanBanner" Visibility="Collapsed"
                            Background="#0C1A24" BorderBrush="#0EA5E9" BorderThickness="1" CornerRadius="6"
                            Margin="0,0,0,15" Padding="15,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <TextBlock Grid.Column="0" Text="||" FontSize="24" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,15,0"
                                       Foreground="#0EA5E9"/>

                            <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                <TextBlock x:Name="txtResumeTitle" Text="Resume Interrupted Scan?"
                                           FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimaryBrush}"/>
                                <TextBlock x:Name="txtResumeDetails" Text="Previous scan was 45% complete (67 of 150 computers)"
                                           FontSize="12" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,2,0,0"/>
                                <TextBlock x:Name="txtResumeSavedAt" Text="Saved: Today at 2:30 PM"
                                           FontSize="11" Foreground="{StaticResource TextMutedBrush}" Margin="0,2,0,0"/>
                            </StackPanel>

                            <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button x:Name="btnResumeScan" Content="Resume Scan" Style="{StaticResource ModernButton}"
                                        Margin="0,0,10,0" ToolTip="Continue scanning from where it left off"/>
                                <Button x:Name="btnDiscardResume" Content="Discard" Style="{StaticResource SecondaryButton}"
                                        ToolTip="Discard saved state and start fresh"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Row 1: Scan Configuration (Always Visible) -->
                    <Border Grid.Row="1" Style="{StaticResource Card}" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" Orientation="Horizontal">
                                <TextBlock Text="Scan Target:" FontSize="14" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}"
                                           VerticalAlignment="Center" Margin="0,0,15,0"/>
                                <ComboBox x:Name="cmbScanTarget" Width="220" VerticalAlignment="Center"
                                          ToolTip="Select which types of computers to scan for unexpected local administrators">
                                    <ComboBoxItem Content="Servers Only" IsSelected="True"/>
                                    <ComboBoxItem Content="Workstations Only"/>
                                    <ComboBoxItem Content="Both (Servers + Workstations)"/>
                                    <ComboBoxItem Content="Specific Computers"/>
                                </ComboBox>

                                <TextBlock Text="Throttle:" FontSize="14" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}"
                                           VerticalAlignment="Center" Margin="30,0,15,0"/>
                                <ComboBox x:Name="cmbThrottle" Width="100" VerticalAlignment="Center"
                                          ToolTip="Maximum concurrent connections. Lower values reduce network load but increase scan time">
                                    <ComboBoxItem Content="32" IsSelected="True"/>
                                    <ComboBoxItem Content="16"/>
                                    <ComboBoxItem Content="64"/>
                                    <ComboBoxItem Content="128"/>
                                </ComboBox>
                            </StackPanel>

                            <StackPanel Grid.Column="2" Orientation="Horizontal">
                                <Button x:Name="btnStartScan" Content="Start Scan" Style="{StaticResource ModernButton}" Margin="0,0,10,0"
                                        ToolTip="Start scanning for unexpected local administrators (F5)"/>
                                <Button x:Name="btnRescan" Content="Rescan" Style="{StaticResource SecondaryButton}" Margin="0,0,10,0"
                                        IsEnabled="False" ToolTip="Repeat the last scan with same settings (Ctrl+R)"/>
                                <Button x:Name="btnCancelScan" Content="Cancel" Style="{StaticResource DangerButton}" IsEnabled="False"
                                        ToolTip="Cancel the running scan (Escape)"/>
                                <Border Width="1" Background="{StaticResource BorderBrush}" Margin="15,4,15,4"/>
                                <Button x:Name="btnElevate" Content="Authenticate" Style="{StaticResource ModernButton}" Margin="0,0,10,0"
                                        ToolTip="Authenticate with your admin account (e.g. DOMAIN\adminusername) for scan operations"/>
                                <TextBlock x:Name="txtCredentialStatus" Text="" FontSize="11"
                                           Foreground="{StaticResource TextSecondaryBrush}" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Specific Computers Panel (Collapsed by default) -->
                    <Border Grid.Row="2" Style="{StaticResource Card}" Margin="0,0,0,15" x:Name="specificComputersCard" Visibility="Collapsed">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Text="Enter computer names (one per line) or import from a file:" FontSize="12"
                                           Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,8"/>
                                <StackPanel Grid.Column="1" Orientation="Horizontal">
                                    <Button x:Name="btnImportComputers" Content="Import from File" Style="{StaticResource SecondaryButton}"
                                            Padding="10,5" Margin="0,0,8,0" ToolTip="Import computer names from a text file (one per line)"/>
                                    <Button x:Name="btnClearComputers" Content="Clear" Style="{StaticResource SecondaryButton}"
                                            Padding="10,5" ToolTip="Clear all entered computer names"/>
                                </StackPanel>
                            </Grid>
                            <TextBox x:Name="txtSpecificComputers" Height="100" AcceptsReturn="True" TextWrapping="Wrap"
                                     VerticalScrollBarVisibility="Auto" Background="#0D1117" Foreground="{StaticResource TextPrimaryBrush}"
                                     BorderBrush="{StaticResource BorderBrush}" Padding="8"
                                     ToolTip="Enter computer names, one per line. Example:&#x0a;SERVER01&#x0a;SERVER02&#x0a;WORKSTATION01"/>
                            <TextBlock x:Name="txtSpecificComputersCount" Text="0 computers entered" FontSize="11"
                                       Foreground="{StaticResource TextSecondaryBrush}" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Border>

                    <!-- Row 3: Progress Section (Collapsed by default) -->
                    <Border Grid.Row="3" Style="{StaticResource Card}" Margin="0,0,0,15" x:Name="progressCard" Visibility="Collapsed">
                        <StackPanel>
                            <Grid Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name="txtProgressStatus" Text="Initializing..." FontSize="14"
                                           Foreground="{StaticResource TextPrimaryBrush}"/>
                                <TextBlock x:Name="txtProgressPercent" Grid.Column="1" Text="0%" FontSize="14"
                                           Foreground="{StaticResource TextSecondaryBrush}"/>
                            </Grid>
                            <ProgressBar x:Name="progressBar" Value="0" Maximum="100"/>
                            <TextBlock x:Name="txtProgressDetail" Text="" FontSize="12"
                                       Foreground="{StaticResource TextSecondaryBrush}" Margin="0,10,0,0"/>
                        </StackPanel>
                    </Border>

                    <!-- Row 4: Stats Cards (Collapsible Expander) -->
                    <Expander Grid.Row="4" Header="Statistics" IsExpanded="True" Margin="0,0,0,15"
                              Foreground="{StaticResource TextPrimaryBrush}">
                        <Grid Margin="0,15,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="15"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="15"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="15"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <!-- Computers Found -->
                            <Border Grid.Column="0" Style="{StaticResource Card}"
                                    ToolTip="Total enabled Windows computers found in Active Directory matching scan target">
                                <StackPanel>
                                    <TextBlock Text="COMPUTERS IN AD" FontSize="11" FontWeight="SemiBold"
                                               Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,8"/>
                                    <TextBlock x:Name="txtTotalComputers" Text="--" FontSize="32" FontWeight="Bold"
                                               Foreground="{StaticResource TextPrimaryBrush}"/>
                                    <TextBlock x:Name="txtComputerBreakdown" Text="Servers: -- | Workstations: --"
                                               FontSize="11" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>

                            <!-- Reached -->
                            <Border Grid.Column="2" Style="{StaticResource Card}"
                                    ToolTip="Computers successfully contacted via WinRM (PowerShell remoting)">
                                <StackPanel>
                                    <TextBlock Text="REACHED" FontSize="11" FontWeight="SemiBold"
                                               Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,8"/>
                                    <TextBlock x:Name="txtReached" Text="--" FontSize="32" FontWeight="Bold"
                                               Foreground="{StaticResource SuccessBrush}"/>
                                    <TextBlock x:Name="txtReachedPercent" Text="--% success rate"
                                               FontSize="11" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>

                            <!-- Unreachable -->
                            <Border Grid.Column="4" Style="{StaticResource Card}"
                                    ToolTip="Computers that could not be contacted - may be offline, firewalled, or WinRM not enabled">
                                <StackPanel>
                                    <TextBlock Text="UNREACHABLE" FontSize="11" FontWeight="SemiBold"
                                               Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,8"/>
                                    <TextBlock x:Name="txtUnreachable" Text="--" FontSize="32" FontWeight="Bold"
                                               Foreground="{StaticResource WarningBrush}"/>
                                    <TextBlock x:Name="txtUnreachableDetail" Text="Connection failed"
                                               FontSize="11" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>

                            <!-- Unexpected Admins -->
                            <Border Grid.Column="6" Style="{StaticResource Card}"
                                    ToolTip="Total unexpected accounts found in local Administrators groups (after exclusions applied)">
                                <StackPanel>
                                    <TextBlock Text="UNEXPECTED ADMINS" FontSize="11" FontWeight="SemiBold"
                                               Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,8"/>
                                    <TextBlock x:Name="txtUnexpectedAdmins" Text="--" FontSize="32" FontWeight="Bold"
                                               Foreground="{StaticResource DangerBrush}"/>
                                    <StackPanel Orientation="Horizontal" Margin="0,5,0,0">
                                        <Border Background="#3D1212" BorderBrush="#FF6B6B" BorderThickness="1" CornerRadius="3" Padding="6,2" Margin="0,0,5,0"
                                                ToolTip="HIGH: Domain users/groups on servers, or unknown accounts">
                                            <TextBlock x:Name="txtHighRisk" Text="0 HIGH" FontSize="10" FontWeight="Bold" Foreground="#FF6B6B"/>
                                        </Border>
                                        <Border Background="#3D2E05" BorderBrush="#FFD93D" BorderThickness="1" CornerRadius="3" Padding="6,2" Margin="0,0,5,0"
                                                ToolTip="MEDIUM: Local users on servers, or domain users on workstations">
                                            <TextBlock x:Name="txtMediumRisk" Text="0 MED" FontSize="10" FontWeight="Bold" Foreground="#FFD93D"/>
                                        </Border>
                                        <Border Background="#0D2F2D" BorderBrush="#4ECDC4" BorderThickness="1" CornerRadius="3" Padding="6,2"
                                                ToolTip="LOW: Local users on workstations, or approved service accounts">
                                            <TextBlock x:Name="txtLowRisk" Text="0 LOW" FontSize="10" FontWeight="Bold" Foreground="#4ECDC4"/>
                                        </Border>
                                    </StackPanel>
                                </StackPanel>
                            </Border>
                        </Grid>
                    </Expander>

                    <!-- Row 5: Filter Bar with Export Options -->
                    <Border Grid.Row="5" Style="{StaticResource Card}" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" Orientation="Horizontal">
                                <TextBlock Text="Filter:" FontSize="14" VerticalAlignment="Center"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,10,0"/>
                                <TextBox x:Name="txtResultsFilter" Width="200"
                                         ToolTip="Filter by computer name or admin account"/>

                                <TextBlock Text="Type:" FontSize="14" VerticalAlignment="Center"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="15,0,10,0"/>
                                <ComboBox x:Name="cmbResultsTypeFilter" Width="110"
                                          ToolTip="Filter results by computer type">
                                    <ComboBoxItem Content="All" IsSelected="True"/>
                                    <ComboBoxItem Content="Servers"/>
                                    <ComboBoxItem Content="Workstations"/>
                                </ComboBox>

                                <TextBlock Text="Group:" FontSize="14" VerticalAlignment="Center"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="15,0,10,0"/>
                                <ComboBox x:Name="cmbGroupBy" Width="130" ToolTip="Group results by selected field">
                                    <ComboBoxItem Content="None" IsSelected="True"/>
                                    <ComboBoxItem Content="Computer Name"/>
                                    <ComboBoxItem Content="Admin Account"/>
                                    <ComboBoxItem Content="Risk Level"/>
                                    <ComboBoxItem Content="Account Type"/>
                                </ComboBox>

                                <Button x:Name="btnClearFilter" Content="Clear" Style="{StaticResource SecondaryButton}"
                                        Margin="15,0,0,0" Padding="12,8"
                                        ToolTip="Reset all filters to show all results"/>
                                <Button x:Name="btnCompare" Content="Compare..." Style="{StaticResource SecondaryButton}"
                                        Margin="15,0,0,0" Padding="12,8"
                                        ToolTip="Compare current results with a previous scan CSV to see NEW and REMOVED admin accounts"/>
                                <Button x:Name="btnClearComparison" Content="Clear Diff" Style="{StaticResource SecondaryButton}"
                                        Margin="8,0,0,0" Padding="12,8" Visibility="Collapsed"
                                        ToolTip="Clear comparison and show current results only"/>
                            </StackPanel>

                            <StackPanel Grid.Column="1" Orientation="Horizontal">
                                <TextBlock x:Name="txtResultsCount" Text="0 results" FontSize="14"
                                           VerticalAlignment="Center" Foreground="{StaticResource TextSecondaryBrush}"
                                           Margin="0,0,15,0"/>
                                <Button x:Name="btnExecSummary" Content="Executive Summary" Style="{StaticResource ModernButton}" Margin="0,0,8,0"
                                        ToolTip="Generate a management-ready HTML summary report"/>
                                <Button x:Name="btnExportCSV" Content="CSV" Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"
                                        ToolTip="Export results to CSV file (Ctrl+S)"/>
                                <Button x:Name="btnExportHTML" Content="HTML" Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"
                                        ToolTip="Export results to detailed HTML report"/>
                                <Button x:Name="btnExportPDF" Content="PDF" Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"
                                        ToolTip="Export results to PDF report"/>
                                <Button x:Name="btnCopyClipboard" Content="Copy" Style="{StaticResource SecondaryButton}"
                                        ToolTip="Copy results to clipboard for pasting"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Row 6: Results Grid with Pro-Tips -->
                    <Border Grid.Row="6" Style="{StaticResource Card}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Pro-Tips Bar -->
                            <Border Grid.Row="0" Background="#0C1A24" CornerRadius="6" Padding="10,6" Margin="0,0,0,10">
                                <StackPanel Orientation="Horizontal">
                                    <TextBlock Text="Pro-Tip:" FontWeight="SemiBold" FontSize="11" Foreground="#3B82F6" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                    <TextBlock Text="Double-click a row to see all admin accounts for that computer. Right-click for more actions."
                                               FontSize="11" Foreground="#94A3B8" VerticalAlignment="Center"/>
                                </StackPanel>
                            </Border>

                        <DataGrid Grid.Row="1" x:Name="dgResults" IsReadOnly="True"
                                  CanUserSortColumns="True" CanUserReorderColumns="True">
                            <DataGrid.ContextMenu>
                                <ContextMenu>
                                    <MenuItem x:Name="ctxResultsCopyComputer" Header="Copy Computer Name"/>
                                    <MenuItem x:Name="ctxResultsCopyAdmin" Header="Copy Admin Account"/>
                                    <MenuItem x:Name="ctxResultsCopyRow" Header="Copy Row"/>
                                    <Separator/>
                                    <MenuItem x:Name="ctxResultsPing" Header="Ping Computer"/>
                                    <MenuItem x:Name="ctxResultsRDP" Header="Remote Desktop"/>
                                    <Separator/>
                                    <MenuItem x:Name="ctxResultsAddExclusion" Header="Add Account to Exclusions"/>
                                    <MenuItem x:Name="ctxResultsExcludeComputer" Header="Exclude Computer from Scans"/>
                                    <MenuItem x:Name="ctxResultsRescan" Header="Rescan This Computer"/>
                                    <Separator/>
                                    <MenuItem x:Name="ctxResultsRemoveAdmin" Header="Remove Selected Admin(s) (Remediate)"
                                              Foreground="#FF6B6B" ToolTip="CAUTION: Remotely remove this account from the local Administrators group"/>
                                </ContextMenu>
                            </DataGrid.ContextMenu>
                            <DataGrid.Columns>
                                <DataGridTemplateColumn Header="Risk" Width="75" SortMemberPath="RiskScore"
                                                        ToolTipService.ToolTip="Risk level based on account type and computer type. Click to sort by severity.">
                                    <DataGridTemplateColumn.CellTemplate>
                                        <DataTemplate>
                                            <Border CornerRadius="4" Padding="6,3" HorizontalAlignment="Center" BorderThickness="1">
                                                <Border.Style>
                                                    <Style TargetType="Border">
                                                        <Setter Property="Background" Value="#2D2D30"/>
                                                        <Setter Property="BorderBrush" Value="#A0AEC0"/>
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding RiskLevel}" Value="HIGH">
                                                                <Setter Property="Background" Value="#3D1212"/>
                                                                <Setter Property="BorderBrush" Value="#FF6B6B"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding RiskLevel}" Value="MEDIUM">
                                                                <Setter Property="Background" Value="#3D2E05"/>
                                                                <Setter Property="BorderBrush" Value="#FFD93D"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding RiskLevel}" Value="LOW">
                                                                <Setter Property="Background" Value="#0D2F2D"/>
                                                                <Setter Property="BorderBrush" Value="#4ECDC4"/>
                                                            </DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </Border.Style>
                                                <TextBlock Text="{Binding RiskLevel}" FontWeight="Bold" FontSize="10" HorizontalAlignment="Center">
                                                    <TextBlock.Style>
                                                        <Style TargetType="TextBlock">
                                                            <Setter Property="Foreground" Value="#A0AEC0"/>
                                                            <Style.Triggers>
                                                                <DataTrigger Binding="{Binding RiskLevel}" Value="HIGH">
                                                                    <Setter Property="Foreground" Value="#FF6B6B"/>
                                                                </DataTrigger>
                                                                <DataTrigger Binding="{Binding RiskLevel}" Value="MEDIUM">
                                                                    <Setter Property="Foreground" Value="#FFD93D"/>
                                                                </DataTrigger>
                                                                <DataTrigger Binding="{Binding RiskLevel}" Value="LOW">
                                                                    <Setter Property="Foreground" Value="#4ECDC4"/>
                                                                </DataTrigger>
                                                            </Style.Triggers>
                                                        </Style>
                                                    </TextBlock.Style>
                                                </TextBlock>
                                            </Border>
                                        </DataTemplate>
                                    </DataGridTemplateColumn.CellTemplate>
                                </DataGridTemplateColumn>
                                <DataGridTemplateColumn Header="Change" Width="70" SortMemberPath="ChangeStatus"
                                                        ToolTipService.ToolTip="Comparison with previous scan: NEW (added), REMOVED (no longer present), or blank (unchanged)">
                                    <DataGridTemplateColumn.CellTemplate>
                                        <DataTemplate>
                                            <Border CornerRadius="4" Padding="6,3" HorizontalAlignment="Center" BorderThickness="1">
                                                <Border.Style>
                                                    <Style TargetType="Border">
                                                        <Setter Property="Background" Value="Transparent"/>
                                                        <Setter Property="BorderBrush" Value="Transparent"/>
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding ChangeStatus}" Value="NEW">
                                                                <Setter Property="Background" Value="#0D3D1D"/>
                                                                <Setter Property="BorderBrush" Value="#4ADE80"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding ChangeStatus}" Value="REMOVED">
                                                                <Setter Property="Background" Value="#3D1212"/>
                                                                <Setter Property="BorderBrush" Value="#FF6B6B"/>
                                                            </DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </Border.Style>
                                                <TextBlock Text="{Binding ChangeStatus}" FontWeight="Bold" FontSize="10" HorizontalAlignment="Center">
                                                    <TextBlock.Style>
                                                        <Style TargetType="TextBlock">
                                                            <Setter Property="Foreground" Value="Transparent"/>
                                                            <Style.Triggers>
                                                                <DataTrigger Binding="{Binding ChangeStatus}" Value="NEW">
                                                                    <Setter Property="Foreground" Value="#4ADE80"/>
                                                                </DataTrigger>
                                                                <DataTrigger Binding="{Binding ChangeStatus}" Value="REMOVED">
                                                                    <Setter Property="Foreground" Value="#FF6B6B"/>
                                                                </DataTrigger>
                                                            </Style.Triggers>
                                                        </Style>
                                                    </TextBlock.Style>
                                                </TextBlock>
                                            </Border>
                                        </DataTemplate>
                                    </DataGridTemplateColumn.CellTemplate>
                                </DataGridTemplateColumn>
                                <DataGridTextColumn Header="Computer Name" Binding="{Binding ComputerName}" Width="150" SortMemberPath="ComputerName"/>
                                <DataGridTextColumn Header="Type" Binding="{Binding ComputerType}" Width="80" SortMemberPath="ComputerType"/>
                                <DataGridTextColumn Header="Operating System" Binding="{Binding OperatingSystem}" Width="160" SortMemberPath="OperatingSystem"/>
                                <DataGridTextColumn Header="Unexpected Admin Account" Binding="{Binding AdminName}" Width="*" SortMemberPath="AdminName"/>
                                <DataGridTextColumn Header="Account Type" Binding="{Binding ObjectClass}" Width="90" SortMemberPath="ObjectClass"/>
                            </DataGrid.Columns>
                        </DataGrid>

                        <!-- Computer Detail Pane (shown on double-click) -->
                        <Border Grid.Row="2" x:Name="computerDetailPane" Visibility="Collapsed"
                                Background="#0C1A24" CornerRadius="6" Padding="15" Margin="0,10,0,0">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <!-- Header with computer info -->
                                <StackPanel Grid.Row="0" Grid.Column="0" Orientation="Horizontal" Margin="0,0,0,10">
                                    <TextBlock Text="Computer:" FontWeight="SemiBold" FontSize="14"
                                               Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,8,0"/>
                                    <TextBlock x:Name="txtDetailComputerName" Text="" FontWeight="Bold" FontSize="14"
                                               Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,20,0"/>
                                    <TextBlock x:Name="txtDetailOS" Text="" FontSize="12"
                                               Foreground="{StaticResource TextSecondaryBrush}"/>
                                </StackPanel>

                                <!-- Close button -->
                                <Button Grid.Row="0" Grid.Column="1" x:Name="btnCloseDetail" Content="X"
                                        Style="{StaticResource SecondaryButton}" Padding="8,4" FontSize="12"
                                        ToolTip="Close detail pane"/>

                                <!-- Admin accounts list -->
                                <Border Grid.Row="1" Grid.ColumnSpan="2" Background="#161B22" CornerRadius="4" Padding="10">
                                    <StackPanel>
                                        <TextBlock Text="All Unexpected Admin Accounts on this Computer:" FontSize="12"
                                                   Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,8"/>
                                        <ItemsControl x:Name="lstDetailAdmins">
                                            <ItemsControl.ItemTemplate>
                                                <DataTemplate>
                                                    <Border Background="#0D1117" CornerRadius="4" Padding="10,8" Margin="0,0,0,5"
                                                            BorderBrush="{StaticResource BorderBrush}" BorderThickness="1">
                                                        <Grid>
                                                            <Grid.ColumnDefinitions>
                                                                <ColumnDefinition Width="60"/>
                                                                <ColumnDefinition Width="*"/>
                                                                <ColumnDefinition Width="80"/>
                                                            </Grid.ColumnDefinitions>
                                                            <!-- Risk Badge -->
                                                            <Border Grid.Column="0" CornerRadius="4" Padding="6,3" HorizontalAlignment="Left" BorderThickness="1">
                                                                <Border.Style>
                                                                    <Style TargetType="Border">
                                                                        <Setter Property="Background" Value="#2D2D30"/>
                                                                        <Setter Property="BorderBrush" Value="#A0AEC0"/>
                                                                        <Style.Triggers>
                                                                            <DataTrigger Binding="{Binding RiskLevel}" Value="HIGH">
                                                                                <Setter Property="Background" Value="#3D1212"/>
                                                                                <Setter Property="BorderBrush" Value="#FF6B6B"/>
                                                                            </DataTrigger>
                                                                            <DataTrigger Binding="{Binding RiskLevel}" Value="MEDIUM">
                                                                                <Setter Property="Background" Value="#3D2E05"/>
                                                                                <Setter Property="BorderBrush" Value="#FFD93D"/>
                                                                            </DataTrigger>
                                                                            <DataTrigger Binding="{Binding RiskLevel}" Value="LOW">
                                                                                <Setter Property="Background" Value="#0D2F2D"/>
                                                                                <Setter Property="BorderBrush" Value="#4ECDC4"/>
                                                                            </DataTrigger>
                                                                        </Style.Triggers>
                                                                    </Style>
                                                                </Border.Style>
                                                                <TextBlock Text="{Binding RiskLevel}" FontWeight="Bold" FontSize="10" HorizontalAlignment="Center">
                                                                    <TextBlock.Style>
                                                                        <Style TargetType="TextBlock">
                                                                            <Setter Property="Foreground" Value="#A0AEC0"/>
                                                                            <Style.Triggers>
                                                                                <DataTrigger Binding="{Binding RiskLevel}" Value="HIGH">
                                                                                    <Setter Property="Foreground" Value="#FF6B6B"/>
                                                                                </DataTrigger>
                                                                                <DataTrigger Binding="{Binding RiskLevel}" Value="MEDIUM">
                                                                                    <Setter Property="Foreground" Value="#FFD93D"/>
                                                                                </DataTrigger>
                                                                                <DataTrigger Binding="{Binding RiskLevel}" Value="LOW">
                                                                                    <Setter Property="Foreground" Value="#4ECDC4"/>
                                                                                </DataTrigger>
                                                                            </Style.Triggers>
                                                                        </Style>
                                                                    </TextBlock.Style>
                                                                </TextBlock>
                                                            </Border>
                                                            <!-- Admin Name -->
                                                            <TextBlock Grid.Column="1" Text="{Binding AdminName}" FontSize="13"
                                                                       Foreground="{StaticResource TextPrimaryBrush}" VerticalAlignment="Center" Margin="10,0"/>
                                                            <!-- Account Type -->
                                                            <TextBlock Grid.Column="2" Text="{Binding ObjectClass}" FontSize="11"
                                                                       Foreground="{StaticResource TextSecondaryBrush}" VerticalAlignment="Center"/>
                                                        </Grid>
                                                    </Border>
                                                </DataTemplate>
                                            </ItemsControl.ItemTemplate>
                                        </ItemsControl>
                                    </StackPanel>
                                </Border>
                            </Grid>
                        </Border>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Unreachable Tab -->
            <TabItem Header="Unreachable" ToolTip="View computers that could not be contacted (Ctrl+2)">
                <Grid Margin="0,20,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Style="{StaticResource Card}" Margin="0,0,0,20">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" Orientation="Horizontal">
                                <TextBlock Text="Search:" FontSize="14" VerticalAlignment="Center"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,10,0"/>
                                <TextBox x:Name="txtUnreachableFilter" Width="300"
                                         ToolTip="Filter unreachable computers by name or error message"/>
                            </StackPanel>

                            <StackPanel Grid.Column="1" Orientation="Horizontal">
                                <TextBlock x:Name="txtUnreachableCount" Text="0 unreachable" FontSize="14"
                                           VerticalAlignment="Center" Foreground="{StaticResource TextSecondaryBrush}"
                                           Margin="0,0,20,0"/>
                                <Button x:Name="btnExportUnreachable" Content="Export List" Style="{StaticResource SecondaryButton}" Margin="0,0,10,0"
                                        ToolTip="Export unreachable computers list to CSV"/>
                                <Button x:Name="btnRetryUnreachable" Content="Retry Selected" Style="{StaticResource ModernButton}"
                                        ToolTip="Retry scanning selected unreachable computers (Ctrl+R)"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <Border Grid.Row="1" Style="{StaticResource Card}">
                        <DataGrid x:Name="dgUnreachable" IsReadOnly="True"
                                  CanUserSortColumns="True" CanUserReorderColumns="True">
                            <DataGrid.ContextMenu>
                                <ContextMenu>
                                    <MenuItem x:Name="ctxUnreachableCopyName" Header="Copy Computer Name"/>
                                    <MenuItem x:Name="ctxUnreachableCopyError" Header="Copy Error Message"/>
                                    <Separator/>
                                    <MenuItem x:Name="ctxUnreachablePing" Header="Ping Computer"/>
                                    <MenuItem x:Name="ctxUnreachableRetryOne" Header="Retry This Computer"/>
                                    <MenuItem x:Name="ctxUnreachableRetrySelected" Header="Retry Selected"/>
                                    <Separator/>
                                    <MenuItem x:Name="ctxUnreachableExcludeComputer" Header="Exclude Computer from Scans"/>
                                </ContextMenu>
                            </DataGrid.ContextMenu>
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Computer Name" Binding="{Binding ComputerName}" Width="200" SortMemberPath="ComputerName"/>
                                <DataGridTextColumn Header="Type" Binding="{Binding ComputerType}" Width="100" SortMemberPath="ComputerType"/>
                                <DataGridTextColumn Header="Operating System" Binding="{Binding OperatingSystem}" Width="250" SortMemberPath="OperatingSystem"/>
                                <DataGridTextColumn Header="Error" Binding="{Binding ErrorMessage}" Width="*" SortMemberPath="ErrorMessage"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Settings Tab -->
            <TabItem Header="Settings" ToolTip="Configure exclusions, output, and preferences (Ctrl+3)">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,20,0,0">
                    <StackPanel>
                        <!-- Exclusion Patterns -->
                        <Border Style="{StaticResource Card}" Margin="0,0,0,20">
                            <StackPanel>
                                <TextBlock Text="Exclusion Patterns" FontSize="18" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock Text="Accounts matching these patterns will be excluded from results (regex supported)"
                                           FontSize="13" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,20"/>

                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>

                                    <StackPanel Grid.Column="0">

                                        <!-- Secondary Admin Account (replaces hardcoded Fadmin) -->
                                        <Border Background="#0C1A24" CornerRadius="6" Padding="14,12" Margin="0,0,0,16"
                                                BorderBrush="{StaticResource AccentBrush}" BorderThickness="1">
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="Auto"/>
                                                </Grid.ColumnDefinitions>
                                                <StackPanel Grid.Column="0">
                                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                                                        <CheckBox x:Name="chkExcludeFadmin" IsChecked="False" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                                        <TextBlock Text="Secondary Admin Account Prefix" FontSize="13" FontWeight="SemiBold"
                                                                   Foreground="{StaticResource TextPrimaryBrush}" VerticalAlignment="Center"/>
                                                        <Border Background="#0EA5E9" CornerRadius="3" Padding="6,2" Margin="10,0,0,0">
                                                            <TextBlock Text="EDITABLE" FontSize="9" FontWeight="Bold" Foreground="White"/>
                                                        </Border>
                                                    </StackPanel>
                                                    <TextBlock TextWrapping="Wrap" FontSize="12" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,8">
                                                        Your organization's standard secondary admin account name or prefix (e.g. Fadmin, ladmin, secadmin).
                                                        All accounts matching this prefix will be excluded from results.
                                                    </TextBlock>
                                                    <Grid>
                                                        <Grid.ColumnDefinitions>
                                                            <ColumnDefinition Width="200"/>
                                                            <ColumnDefinition Width="*"/>
                                                        </Grid.ColumnDefinitions>
                                                        <TextBox x:Name="txtFadminPattern" Text="Fadmin"
                                                                 ToolTip="Enter your secondary admin account name or prefix. Matches any account containing this text (case-insensitive)."
                                                                 Padding="10,7" FontSize="13"/>
                                                        <TextBlock Grid.Column="1" FontSize="11" Foreground="{StaticResource TextSecondaryBrush}"
                                                                   VerticalAlignment="Center" Margin="12,0,0,0"
                                                                   Text="Matches accounts containing this text (e.g. 'Fadmin' matches CORP\FadminJohn, FadminSvc, etc.)"/>
                                                    </Grid>
                                                </StackPanel>
                                            </Grid>
                                        </Border>

                                        <CheckBox x:Name="chkExcludeDomainAdmins" Content="Domain Admins group" IsChecked="True"
                                                  ToolTip="Exclude Domain Admins group members from results"/>
                                        <CheckBox x:Name="chkExcludeBuiltinAdmin" Content="Built-in Administrator account" IsChecked="True"
                                                  ToolTip="Exclude the built-in local Administrator account from results"/>

                                        <TextBlock Text="Custom Exclusion Patterns (one per line):"
                                                   FontSize="13" Foreground="{StaticResource TextSecondaryBrush}"
                                                   Margin="0,20,0,10"/>
                                        <TextBox x:Name="txtCustomExclusions" Height="120"
                                                 AcceptsReturn="True" TextWrapping="Wrap"
                                                 VerticalScrollBarVisibility="Auto"
                                                 ToolTip="Enter regex patterns, one per line. Example: \\ServiceAccount$"/>
                                    </StackPanel>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- Computer Exclusions (Honeypots, Test Systems, etc.) -->
                        <Border Style="{StaticResource Card}" Margin="0,0,0,20">
                            <StackPanel>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock Text="Computer Exclusions" FontSize="18" FontWeight="SemiBold"
                                                   Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,5"/>
                                        <TextBlock Text="Computers that will never be scanned (honeypots, isolated test systems, etc.)"
                                                   FontSize="13" Foreground="{StaticResource TextSecondaryBrush}"/>
                                    </StackPanel>
                                    <Border Grid.Column="1" Background="#3D1F1F" CornerRadius="6" Padding="10,6" VerticalAlignment="Center">
                                        <StackPanel Orientation="Horizontal">
                                            <TextBlock Text="!" FontWeight="Bold" Foreground="#EF4444" Margin="0,0,6,0"/>
                                            <TextBlock Text="Security Feature" FontSize="11" Foreground="#EF4444"/>
                                        </StackPanel>
                                    </Border>
                                </Grid>

                                <Grid Margin="0,20,0,0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>

                                    <!-- Exclusion List -->
                                    <Border Grid.Column="0" Background="#161B22" CornerRadius="8" Padding="0" Margin="0,0,15,0">
                                        <ListBox x:Name="lstComputerExclusions" MinHeight="180" MaxHeight="250"
                                                 Background="Transparent" BorderThickness="0"
                                                 ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                                                 SelectionMode="Extended">
                                            <ListBox.ItemTemplate>
                                                <DataTemplate>
                                                    <Grid Margin="8,6">
                                                        <Grid.ColumnDefinitions>
                                                            <ColumnDefinition Width="Auto"/>
                                                            <ColumnDefinition Width="*"/>
                                                            <ColumnDefinition Width="Auto"/>
                                                        </Grid.ColumnDefinitions>
                                                        <!-- Icon based on type -->
                                                        <Border Grid.Column="0" Width="28" Height="28" CornerRadius="6" Margin="0,0,12,0"
                                                                Background="{Binding IconBackground}">
                                                            <TextBlock Text="{Binding Icon}" FontSize="14"
                                                                       HorizontalAlignment="Center" VerticalAlignment="Center"
                                                                       Foreground="{Binding IconForeground}"/>
                                                        </Border>
                                                        <!-- Computer name and reason -->
                                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                                            <TextBlock Text="{Binding ComputerName}" FontSize="14" FontWeight="SemiBold"
                                                                       Foreground="#F8FAFC"/>
                                                            <TextBlock Text="{Binding Reason}" FontSize="11"
                                                                       Foreground="#94A3B8" Margin="0,2,0,0"/>
                                                        </StackPanel>
                                                        <!-- Type badge -->
                                                        <Border Grid.Column="2" Background="{Binding BadgeBackground}"
                                                                CornerRadius="4" Padding="8,3" VerticalAlignment="Center">
                                                            <TextBlock Text="{Binding ExclusionType}" FontSize="10" FontWeight="SemiBold"
                                                                       Foreground="{Binding BadgeForeground}"/>
                                                        </Border>
                                                    </Grid>
                                                </DataTemplate>
                                            </ListBox.ItemTemplate>
                                            <ListBox.ItemContainerStyle>
                                                <Style TargetType="ListBoxItem">
                                                    <Setter Property="Background" Value="Transparent"/>
                                                    <Setter Property="Foreground" Value="#F8FAFC"/>
                                                    <Setter Property="Padding" Value="4"/>
                                                    <Setter Property="Margin" Value="4,2"/>
                                                    <Setter Property="Template">
                                                        <Setter.Value>
                                                            <ControlTemplate TargetType="ListBoxItem">
                                                                <Border x:Name="Border" Background="{TemplateBinding Background}"
                                                                        CornerRadius="6" Padding="{TemplateBinding Padding}">
                                                                    <ContentPresenter/>
                                                                </Border>
                                                                <ControlTemplate.Triggers>
                                                                    <Trigger Property="IsMouseOver" Value="True">
                                                                        <Setter TargetName="Border" Property="Background" Value="#1C2333"/>
                                                                    </Trigger>
                                                                    <Trigger Property="IsSelected" Value="True">
                                                                        <Setter TargetName="Border" Property="Background" Value="#0EA5E9"/>
                                                                    </Trigger>
                                                                </ControlTemplate.Triggers>
                                                            </ControlTemplate>
                                                        </Setter.Value>
                                                    </Setter>
                                                </Style>
                                            </ListBox.ItemContainerStyle>
                                        </ListBox>
                                    </Border>

                                    <!-- Action buttons -->
                                    <StackPanel Grid.Column="1" Width="160">
                                        <Button x:Name="btnAddComputerExclusion" Content="+ Add Computer"
                                                Style="{StaticResource ModernButton}" Margin="0,0,0,8"
                                                ToolTip="Add a computer to exclude from scans (test/dev systems)"/>
                                        <Button x:Name="btnAddHoneypot" Content="+ Add Honeypot"
                                                Style="{StaticResource SecondaryButton}" Margin="0,0,0,8"
                                                ToolTip="Add a honeypot computer that should never be scanned">
                                            <Button.Background>
                                                <SolidColorBrush Color="#3D1F1F"/>
                                            </Button.Background>
                                        </Button>
                                        <Button x:Name="btnRemoveComputerExclusion" Content="Remove Selected"
                                                Style="{StaticResource SecondaryButton}" Margin="0,0,0,20"
                                                ToolTip="Remove selected computers from the exclusion list"/>

                                        <TextBlock Text="Quick Add:" FontSize="11" Foreground="#94A3B8" Margin="0,0,0,8"/>
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>
                                            <TextBox x:Name="txtQuickAddComputer" Grid.Column="0"
                                                     ToolTip="Enter computer name and press Enter or click +"
                                                     Height="36"/>
                                            <Button x:Name="btnQuickAdd" Grid.Column="1" Content="+"
                                                    Style="{StaticResource SecondaryButton}"
                                                    Width="36" Height="36" Padding="0" Margin="4,0,0,0"
                                                    ToolTip="Add computer to exclusion list"/>
                                        </Grid>
                                    </StackPanel>
                                </Grid>

                                <!-- Info panel -->
                                <Border Background="#0C1A24" CornerRadius="6" Padding="12,10" Margin="0,15,0,0">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Grid.Column="0" Text="i" FontWeight="Bold" FontSize="12"
                                                   Foreground="#3B82F6" Margin="0,0,10,0" VerticalAlignment="Top"/>
                                        <TextBlock Grid.Column="1" TextWrapping="Wrap" FontSize="12" Foreground="#94A3B8">
                                            <Run FontWeight="SemiBold" Foreground="#F8FAFC">Honeypots</Run>
                                            <Run>are decoy systems designed to detect attacks. Scanning them could trigger alerts or pollute audit logs.</Run>
                                            <LineBreak/>
                                            <Run FontWeight="SemiBold" Foreground="#F8FAFC">Test systems</Run>
                                            <Run>may have intentionally misconfigured permissions that shouldn't appear in compliance reports.</Run>
                                        </TextBlock>
                                    </Grid>
                                </Border>
                            </StackPanel>
                        </Border>

                        <!-- Approved Service Accounts -->
                        <Border Style="{StaticResource Card}" Margin="0,0,0,20">
                            <StackPanel>
                                <TextBlock Text="Approved Service Accounts" FontSize="18" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock Text="Service accounts listed here are considered approved and will receive a reduced risk score"
                                           FontSize="13" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,20"/>

                                <TextBlock Text="Enter account names or patterns (one per line):"
                                           FontSize="13" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,10"/>
                                <TextBox x:Name="txtApprovedAccounts" Height="100"
                                         AcceptsReturn="True" TextWrapping="Wrap"
                                         VerticalScrollBarVisibility="Auto"
                                         ToolTip="Enter account names or patterns. Example: svc_backup, YOURDOM\\svc_*"/>

                                <StackPanel Orientation="Horizontal" Margin="0,15,0,0">
                                    <Border Background="#0C1A24" CornerRadius="4" Padding="8,4" Margin="0,0,10,0">
                                        <TextBlock Text="Matched accounts get -50 risk score" FontSize="11"
                                                   Foreground="#3B82F6"/>
                                    </Border>
                                    <TextBlock Text="Tip: Use partial names to match patterns (e.g., 'svc_' matches all service accounts starting with svc_)"
                                               FontSize="11" Foreground="{StaticResource TextSecondaryBrush}"
                                               VerticalAlignment="Center"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <!-- Output Settings -->
                        <Border Style="{StaticResource Card}" Margin="0,0,0,20">
                            <StackPanel>
                                <TextBlock Text="Output Settings" FontSize="18" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock Text="Configure where scan results are saved"
                                           FontSize="13" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,20"/>

                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>

                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Output Directory:"
                                               FontSize="14" Foreground="{StaticResource TextPrimaryBrush}"
                                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                                    <StackPanel Grid.Row="1" Grid.Column="0" Orientation="Horizontal">
                                        <TextBox x:Name="txtOutputDir" Width="500" IsReadOnly="True"
                                                 ToolTip="Directory where CSV and HTML exports will be saved"/>
                                        <Button x:Name="btnBrowseOutput" Content="Browse..."
                                                Style="{StaticResource SecondaryButton}" Margin="10,0,0,0"
                                                ToolTip="Select a different output directory"/>
                                    </StackPanel>
                                </Grid>

                                <StackPanel Margin="0,20,0,0">
                                    <CheckBox x:Name="chkAutoExportCSV" Content="Automatically export CSV after each scan" IsChecked="True"
                                              ToolTip="Save results to CSV file automatically when scan completes"/>
                                    <CheckBox x:Name="chkAutoExportUnreachable" Content="Automatically export unreachable machines log" IsChecked="True"
                                              ToolTip="Save list of unreachable computers to CSV when scan completes"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <!-- Connection Settings -->
                        <Border Style="{StaticResource Card}" Margin="0,0,0,20">
                            <StackPanel>
                                <TextBlock Text="Connection Settings" FontSize="18" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock Text="Configure PowerShell remoting behavior"
                                           FontSize="13" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,20"/>

                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>

                                    <TextBlock Grid.Column="0" Text="Connection Timeout (seconds):"
                                               FontSize="14" Foreground="{StaticResource TextPrimaryBrush}"
                                               VerticalAlignment="Center" Margin="0,0,15,0"/>
                                    <TextBox x:Name="txtTimeout" Grid.Column="1" Width="80" Text="30"
                                             ToolTip="How long to wait for each computer to respond before marking as unreachable"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- Notifications -->
                        <Border Style="{StaticResource Card}" Margin="0,0,0,20">
                            <StackPanel>
                                <TextBlock Text="Notifications" FontSize="18" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock Text="Configure sound and notification preferences"
                                           FontSize="13" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,20"/>

                                <CheckBox x:Name="chkPlaySound" Content="Play sound when scan completes" IsChecked="True"
                                          ToolTip="Play a notification sound when the scan finishes"/>
                            </StackPanel>
                        </Border>

                        <!-- Keyboard Shortcuts Reference -->
                        <Border Style="{StaticResource Card}">
                            <StackPanel>
                                <TextBlock Text="Keyboard Shortcuts" FontSize="18" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock Text="Quick access to common functions"
                                           FontSize="13" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,20"/>

                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="120"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>

                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="F5" FontWeight="SemiBold" Foreground="{StaticResource AccentBrush}" Margin="0,4"/>
                                    <TextBlock Grid.Row="0" Grid.Column="1" Text="Start Scan" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,4"/>

                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Escape" FontWeight="SemiBold" Foreground="{StaticResource AccentBrush}" Margin="0,4"/>
                                    <TextBlock Grid.Row="1" Grid.Column="1" Text="Cancel Scan" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,4"/>

                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Ctrl+S" FontWeight="SemiBold" Foreground="{StaticResource AccentBrush}" Margin="0,4"/>
                                    <TextBlock Grid.Row="2" Grid.Column="1" Text="Export Results to CSV" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,4"/>

                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Ctrl+F" FontWeight="SemiBold" Foreground="{StaticResource AccentBrush}" Margin="0,4"/>
                                    <TextBlock Grid.Row="3" Grid.Column="1" Text="Focus Filter Field" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,4"/>

                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Ctrl+R" FontWeight="SemiBold" Foreground="{StaticResource AccentBrush}" Margin="0,4"/>
                                    <TextBlock Grid.Row="4" Grid.Column="1" Text="Retry Selected Unreachable" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,4"/>

                                    <TextBlock Grid.Row="5" Grid.Column="0" Text="Ctrl+1 to 6" FontWeight="SemiBold" Foreground="{StaticResource AccentBrush}" Margin="0,4"/>
                                    <TextBlock Grid.Row="5" Grid.Column="1" Text="Switch Tabs" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,4"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- History Tab -->
            <TabItem Header="History" ToolTip="View and load previous scan exports (Ctrl+4)">
                <Grid Margin="0,20,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- History Header -->
                    <Border Grid.Row="0" Style="{StaticResource Card}" Margin="0,0,0,20">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0">
                                <TextBlock Text="Scan History" FontSize="18" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimaryBrush}"/>
                                <TextBlock x:Name="txtHistoryInfo" Text="Loading history..." FontSize="13"
                                           Foreground="{StaticResource TextSecondaryBrush}" Margin="0,5,0,0"/>
                            </StackPanel>

                            <StackPanel Grid.Column="1" Orientation="Horizontal">
                                <Button x:Name="btnRefreshHistory" Content="Refresh" Style="{StaticResource SecondaryButton}" Margin="0,0,10,0"
                                        ToolTip="Rescan output folders for export files"/>
                                <Button x:Name="btnLoadHistory" Content="Load Selected" Style="{StaticResource ModernButton}" Margin="0,0,10,0"
                                        ToolTip="Load selected CSV into Results grid, or open HTML in browser"/>
                                <Button x:Name="btnDeleteHistory" Content="Delete Selected" Style="{StaticResource DangerButton}" Margin="0,0,10,0"
                                        ToolTip="Permanently delete selected export file"/>
                                <Button x:Name="btnClearHistory" Content="Clear All" Style="{StaticResource SecondaryButton}"
                                        ToolTip="Delete all export files (irreversible)"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- History Grid -->
                    <Border Grid.Row="1" Style="{StaticResource Card}">
                        <DataGrid x:Name="dgHistory" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Date" Binding="{Binding DateFormatted}" Width="160"/>
                                <DataGridTextColumn Header="File Name" Binding="{Binding FileName}" Width="220"/>
                                <DataGridTextColumn Header="Type" Binding="{Binding FileType}" Width="120"/>
                                <DataGridTextColumn Header="Rows" Binding="{Binding RowCount}" Width="60"/>
                                <DataGridTextColumn Header="Size" Binding="{Binding FileSize}" Width="80"/>
                                <DataGridTextColumn Header="Folder" Binding="{Binding FolderDate}" Width="100"/>
                                <DataGridTextColumn Header="Full Path" Binding="{Binding FullPath}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>

                    <!-- Loaded Scan Details -->
                    <Border Grid.Row="2" Style="{StaticResource Card}" Margin="0,20,0,0" x:Name="borderHistoryDetails" Visibility="Collapsed">
                        <StackPanel>
                            <TextBlock x:Name="txtLoadedScanInfo" Text="" FontSize="14" FontWeight="SemiBold"
                                       Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,10"/>
                            <TextBlock Text="Use 'Load Selected' to view details from a past scan in the Results tab."
                                       FontSize="13" Foreground="{StaticResource TextSecondaryBrush}"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Trends Tab -->
            <TabItem Header="Trends" ToolTip="View historical trends and charts (Ctrl+5)">
                <Grid Margin="0,20,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Controls -->
                    <Border Grid.Row="0" Style="{StaticResource Card}" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Orientation="Horizontal">
                                <TextBlock Text="Historical Trend Analysis" FontSize="18" FontWeight="Bold"
                                           Foreground="{StaticResource TextPrimaryBrush}" VerticalAlignment="Center" Margin="0,0,20,0"/>
                                <TextBlock x:Name="txtTrendDateRange" Text="No data loaded" FontSize="13"
                                           Foreground="{StaticResource TextSecondaryBrush}" VerticalAlignment="Center"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal">
                                <Button x:Name="btnRefreshTrends" Content="Refresh Data" Style="{StaticResource SecondaryButton}" Margin="0,0,10,0"
                                        ToolTip="Scan historical CSV files and refresh trend data"/>
                                <Button x:Name="btnExportTrendChart" Content="Export Chart" Style="{StaticResource ModernButton}"
                                        ToolTip="Export trend chart as HTML"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Summary Stats -->
                    <Border Grid.Row="1" Style="{StaticResource Card}" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" HorizontalAlignment="Center">
                                <TextBlock Text="SCANS ANALYZED" FontSize="11" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock x:Name="txtTrendScanCount" Text="--" FontSize="28" FontWeight="Bold"
                                           Foreground="{StaticResource TextPrimaryBrush}"/>
                            </StackPanel>

                            <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                                <TextBlock Text="AVG FINDINGS/SCAN" FontSize="11" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock x:Name="txtTrendAvgFindings" Text="--" FontSize="28" FontWeight="Bold"
                                           Foreground="{StaticResource WarningBrush}"/>
                            </StackPanel>

                            <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                                <TextBlock Text="TREND DIRECTION" FontSize="11" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock x:Name="txtTrendDirection" Text="--" FontSize="28" FontWeight="Bold"
                                           Foreground="{StaticResource TextPrimaryBrush}"/>
                            </StackPanel>

                            <StackPanel Grid.Column="3" HorizontalAlignment="Center">
                                <TextBlock Text="HIGH RISK TREND" FontSize="11" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextSecondaryBrush}" Margin="0,0,0,5"/>
                                <TextBlock x:Name="txtTrendHighRisk" Text="--" FontSize="28" FontWeight="Bold"
                                           Foreground="{StaticResource DangerBrush}"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Trend Data Grid -->
                    <Border Grid.Row="2" Style="{StaticResource Card}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>

                            <TextBlock Grid.Row="0" Text="Historical Scan Data" FontSize="14" FontWeight="SemiBold"
                                       Foreground="{StaticResource TextPrimaryBrush}" Margin="0,0,0,10"/>

                            <DataGrid Grid.Row="1" x:Name="dgTrends" IsReadOnly="True"
                                      CanUserSortColumns="True" CanUserReorderColumns="True">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Date" Binding="{Binding Date}" Width="100" SortMemberPath="SortDate"/>
                                    <DataGridTextColumn Header="Computers Scanned" Binding="{Binding ComputersScanned}" Width="130"/>
                                    <DataGridTextColumn Header="Total Findings" Binding="{Binding TotalFindings}" Width="110"/>
                                    <DataGridTextColumn Header="High Risk" Binding="{Binding HighRisk}" Width="90"/>
                                    <DataGridTextColumn Header="Medium Risk" Binding="{Binding MediumRisk}" Width="100"/>
                                    <DataGridTextColumn Header="Low Risk" Binding="{Binding LowRisk}" Width="90"/>
                                    <DataGridTextColumn Header="Change" Binding="{Binding Change}" Width="80"/>
                                    <DataGridTextColumn Header="Source File" Binding="{Binding SourceFile}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Log Tab -->
            <TabItem Header="Activity Log" ToolTip="View detailed operation logs (Ctrl+6)">
                <Grid Margin="0,20,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Style="{StaticResource Card}" Margin="0,0,0,20">
                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="btnClearLog" Content="Clear Log" Style="{StaticResource SecondaryButton}" Margin="0,0,10,0"
                                    ToolTip="Clear all log entries from the display"/>
                            <Button x:Name="btnSaveLog" Content="Save Log" Style="{StaticResource SecondaryButton}"
                                    ToolTip="Save the activity log to a text file"/>
                        </StackPanel>
                    </Border>

                    <Border Grid.Row="1" Style="{StaticResource Card}">
                        <TextBox x:Name="txtLog" IsReadOnly="True"
                                 ToolTip="Detailed activity log of all scan operations and events"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"
                                 TextWrapping="NoWrap"
                                 FontFamily="Consolas"
                                 FontSize="12"/>
                    </Border>
                </Grid>
            </TabItem>
        </TabControl>

        <!-- Status Bar with Attribution -->
        <Border Grid.Row="2" Background="{StaticResource CardBrush}" CornerRadius="8" Padding="15,10" Margin="0,15,0,0"
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Status Bar Row -->
                <TextBlock x:Name="txtStatusBar" Grid.Row="0" Grid.Column="0" Text="Ready" FontSize="12"
                           Foreground="{StaticResource TextSecondaryBrush}" VerticalAlignment="Center"/>

                <TextBlock x:Name="txtLastScanTime" Grid.Row="0" Grid.Column="1" Text="" FontSize="12"
                           Foreground="{StaticResource TextSecondaryBrush}" VerticalAlignment="Center" Margin="0,0,30,0"/>

                <TextBlock Grid.Row="0" Grid.Column="2" Text="v9.0" FontSize="12"
                           Foreground="{StaticResource AccentBrush}" FontWeight="SemiBold" VerticalAlignment="Center"
                           ToolTip="Local Administrator Auditor v9.0 - Hover over controls for tips"/>

                <!-- Attribution Row -->
                <StackPanel Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="3"
                            Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,8,0,0">
                    <TextBlock FontSize="11" Foreground="#484F58" VerticalAlignment="Center">
                        <TextBlock.Inlines>
                            <Run Text="Hugh Gozlou" Foreground="#8B949E"/>
                            <Run Text="  |  " Foreground="#30363D"/>
                        </TextBlock.Inlines>
                    </TextBlock>
                    <TextBlock FontSize="11" VerticalAlignment="Center">
                        <TextBlock.Foreground>
                            <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                                <GradientStop Color="#0EA5E9" Offset="0"/>
                                <GradientStop Color="#38BDF8" Offset="1"/>
                            </LinearGradientBrush>
                        </TextBlock.Foreground>
                        <TextBlock.Text>github.com/hugh2024/powershell-tools</TextBlock.Text>
                    </TextBlock>
                    <TextBlock Text="  |  MIT License" FontSize="11" Foreground="#484F58" VerticalAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@
#endregion

#region Initialize Window
$Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$XAML)
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# Get all named controls (Script-scoped for Dispatcher callback access)
$Script:Controls = @{}
$XAML -split "`n" | ForEach-Object {
    if ($_ -match 'x:Name="([^"]+)"') {
        $name = $Matches[1]
        $Script:Controls[$name] = $Window.FindName($name)
    }
}
#endregion

#region Global Variables
$Script:ScanResults = [System.Collections.ArrayList]::new()
$Script:UnreachableComputers = [System.Collections.ArrayList]::new()
$Script:AllComputers = @()
$Script:CancelRequested = $false
$Script:OutputDirectory = if ([string]::IsNullOrEmpty($PSScriptRoot)) { Get-Location } else { $PSScriptRoot }
$Script:LastScanSettings = $null  # Stores settings for Quick Rescan feature
$Script:ElevatedCredential = $null  # Alternate credentials for scan operations (memory-only, never persisted)

# JSON Persistence Paths
$Script:AppDataPath = Join-Path $env:APPDATA "LocalAdminAuditor"
$Script:HistoryFile = Join-Path $Script:AppDataPath "ScanHistory.json"
$Script:SettingsFile = Join-Path $Script:AppDataPath "UserSettings.json"
$Script:ScanStateFile = Join-Path $Script:AppDataPath "ScanState.json"
$Script:ScanHistory = @{ version = "1.0"; scans = @(); knownFindings = @{} }
$Script:UserSettings = $null
$Script:ComputerExclusions = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()

# Set initial output directory
$Controls['txtOutputDir'].Text = $Script:OutputDirectory
#endregion

#region Helper Functions
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message`r`n"

    $Window.Dispatcher.Invoke([Action]{
        # Use $Script:Controls explicitly for Dispatcher callback access
        $Script:Controls['txtLog'].AppendText($logEntry)
        $Script:Controls['txtLog'].ScrollToEnd()
    })
}

function Update-Progress {
    param(
        [string]$Status,
        [int]$Percent,
        [string]$Detail = ""
    )

    $Window.Dispatcher.Invoke([Action]{
        # Use $Script:Controls explicitly for Dispatcher callback access
        $Script:Controls['txtProgressStatus'].Text = $Status
        $Script:Controls['txtProgressPercent'].Text = "$Percent%"
        $Script:Controls['progressBar'].Value = $Percent
        $Script:Controls['txtProgressDetail'].Text = $Detail
    })
}

function Update-Statistics {
    $Window.Dispatcher.Invoke([Action]{
        $total = $Script:AllComputers.Count
        $reached = ($Script:ScanResults | Select-Object -ExpandProperty ComputerName -Unique).Count
        $unreachable = $Script:UnreachableComputers.Count
        $unexpected = $Script:ScanResults.Count

        $servers = ($Script:AllComputers | Where-Object { $_.OperatingSystem -like "*Server*" }).Count
        $workstations = $total - $servers

        # Use $Script:Controls explicitly for Dispatcher callback access
        $Script:Controls['txtTotalComputers'].Text = $total.ToString()
        $Script:Controls['txtComputerBreakdown'].Text = "Servers: $servers | Workstations: $workstations"

        $Script:Controls['txtReached'].Text = $reached.ToString()
        $reachedPercent = if ($total -gt 0) { [math]::Round(($reached / $total) * 100, 1) } else { 0 }
        $Script:Controls['txtReachedPercent'].Text = "$reachedPercent% success rate"

        $Script:Controls['txtUnreachable'].Text = $unreachable.ToString()
        $Script:Controls['txtUnreachableDetail'].Text = "Connection failed"

        $Script:Controls['txtUnexpectedAdmins'].Text = $unexpected.ToString()

        # Update risk badges (these exist instead of txtUnexpectedDetail)
        $highCount = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'HIGH' }).Count
        $medCount = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'MEDIUM' }).Count
        $lowCount = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'LOW' }).Count
        $Script:Controls['txtHighRisk'].Text = "$highCount HIGH"
        $Script:Controls['txtMediumRisk'].Text = "$medCount MED"
        $Script:Controls['txtLowRisk'].Text = "$lowCount LOW"

        # Update counts on other tabs
        $Script:Controls['txtResultsCount'].Text = "$unexpected results"
        $Script:Controls['txtUnreachableCount'].Text = "$unreachable unreachable"
    })
}

function Get-ExclusionPatterns {
    $patterns = @()

    if ($Controls['chkExcludeFadmin'].IsChecked) {
        $fadminPattern = $Controls['txtFadminPattern'].Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($fadminPattern)) {
            $patterns += $fadminPattern
        }
    }
    if ($Controls['chkExcludeDomainAdmins'].IsChecked) { $patterns += '\\Domain Admins$' }
    if ($Controls['chkExcludeBuiltinAdmin'].IsChecked) { $patterns += '\\Administrator$' }

    # Add custom patterns
    $customPatterns = $Controls['txtCustomExclusions'].Text -split "`r`n" | Where-Object { $_.Trim() -ne '' }
    $patterns += $customPatterns

    return $patterns
}

#region Scan State Management (Resume Interrupted Scans)
function Save-ScanState {
    <#
    .SYNOPSIS
    Saves the current scan state to allow resuming after interruption.
    #>
    param(
        [array]$AllComputers,
        [array]$ProcessedComputers,
        [array]$Results,
        [array]$Unreachable,
        [hashtable]$Settings
    )

    try {
        $state = @{
            Version = "1.0"
            SavedAt = (Get-Date).ToString('o')
            ScanSettings = @{
                ScanTarget = $Settings.ScanTarget
                ThrottleLimit = $Settings.ThrottleLimit
                SpecificComputers = $Settings.SpecificComputers
                ExclusionPatterns = $Settings.ExclusionPatterns
            }
            AllComputers = @($AllComputers | ForEach-Object {
                @{
                    Name = $_.Name
                    OperatingSystem = $_.OperatingSystem
                    DistinguishedName = $_.DistinguishedName
                }
            })
            ProcessedComputers = @($ProcessedComputers)
            Results = @($Results | ForEach-Object {
                @{
                    ComputerName = $_.ComputerName
                    ComputerType = $_.ComputerType
                    OperatingSystem = $_.OperatingSystem
                    AccountName = $_.AccountName
                    AccountType = $_.AccountType
                    RiskLevel = $_.RiskLevel
                    RiskScore = $_.RiskScore
                    Status = $_.Status
                }
            })
            Unreachable = @($Unreachable | ForEach-Object {
                @{
                    ComputerName = $_.ComputerName
                    OperatingSystem = $_.OperatingSystem
                    ErrorMessage = $_.ErrorMessage
                }
            })
            TotalComputers = $AllComputers.Count
            CompletedCount = $ProcessedComputers.Count
        }

        # Ensure directory exists
        if (-not (Test-Path $Script:AppDataPath)) {
            New-Item -ItemType Directory -Path $Script:AppDataPath -Force | Out-Null
        }

        $state | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:ScanStateFile -Encoding UTF8
        return $true
    }
    catch {
        Write-Log "Failed to save scan state: $_" -Level "ERROR"
        return $false
    }
}

function Load-ScanState {
    <#
    .SYNOPSIS
    Loads a saved scan state for resuming an interrupted scan.
    #>

    if (-not (Test-Path $Script:ScanStateFile)) {
        return $null
    }

    try {
        $json = Get-Content -Path $Script:ScanStateFile -Raw | ConvertFrom-Json

        # Validate state structure
        if (-not $json.Version -or -not $json.AllComputers -or -not $json.ScanSettings) {
            Write-Log "Invalid scan state file structure" -Level "WARN"
            return $null
        }

        # Check if state is too old (older than 24 hours)
        $savedTime = [DateTime]::Parse($json.SavedAt)
        $hoursOld = (Get-Date).Subtract($savedTime).TotalHours
        if ($hoursOld -gt 24) {
            Write-Log "Scan state is $([int]$hoursOld) hours old, discarding" -Level "WARN"
            Clear-ScanState
            return $null
        }

        return $json
    }
    catch {
        Write-Log "Failed to load scan state: $_" -Level "ERROR"
        return $null
    }
}

function Clear-ScanState {
    <#
    .SYNOPSIS
    Clears the saved scan state after successful completion or manual dismissal.
    #>

    if (Test-Path $Script:ScanStateFile) {
        try {
            Remove-Item -Path $Script:ScanStateFile -Force
            Write-Log "Cleared saved scan state" -Level "INFO"
        }
        catch {
            Write-Log "Failed to clear scan state: $_" -Level "WARN"
        }
    }
}

function Get-ScanStateInfo {
    <#
    .SYNOPSIS
    Gets summary information about a saved scan state without loading full data.
    #>

    $state = Load-ScanState
    if ($null -eq $state) {
        return $null
    }

    $savedTime = [DateTime]::Parse($state.SavedAt)
    $remaining = $state.TotalComputers - $state.CompletedCount
    $percentComplete = if ($state.TotalComputers -gt 0) { [math]::Round(($state.CompletedCount / $state.TotalComputers) * 100, 1) } else { 0 }

    return @{
        SavedAt = $savedTime
        ScanTarget = $state.ScanSettings.ScanTarget
        TotalComputers = $state.TotalComputers
        CompletedCount = $state.CompletedCount
        RemainingCount = $remaining
        PercentComplete = $percentComplete
        ResultsFound = $state.Results.Count
        UnreachableCount = $state.Unreachable.Count
    }
}
#endregion

function New-ComputerExclusionItem {
    <#
    .SYNOPSIS
    Creates a formatted computer exclusion item for the ListBox with visual styling.
    #>
    param(
        [string]$ComputerName,
        [string]$ExclusionType = "EXCLUDED",
        [string]$Reason = "User excluded"
    )

    # Define visual styling based on exclusion type
    $styling = switch ($ExclusionType.ToUpper()) {
        "HONEYPOT" {
            @{
                Icon = "!"
                IconBackground = "#3D1F1F"
                IconForeground = "#EF4444"
                BadgeBackground = "#3D1F1F"
                BadgeForeground = "#EF4444"
            }
        }
        "TEST" {
            @{
                Icon = "T"
                IconBackground = "#0C1A24"
                IconForeground = "#3B82F6"
                BadgeBackground = "#0C1A24"
                BadgeForeground = "#3B82F6"
            }
        }
        "DEV" {
            @{
                Icon = "D"
                IconBackground = "#1F3D2E"
                IconForeground = "#10B981"
                BadgeBackground = "#1F3D2E"
                BadgeForeground = "#10B981"
            }
        }
        default {
            @{
                Icon = "X"
                IconBackground = "#161B22"
                IconForeground = "#8B949E"
                BadgeBackground = "#161B22"
                BadgeForeground = "#8B949E"
            }
        }
    }

    return [PSCustomObject]@{
        ComputerName = $ComputerName.ToUpper()
        ExclusionType = $ExclusionType.ToUpper()
        Reason = $Reason
        Icon = $styling.Icon
        IconBackground = $styling.IconBackground
        IconForeground = $styling.IconForeground
        BadgeBackground = $styling.BadgeBackground
        BadgeForeground = $styling.BadgeForeground
    }
}

function Get-ExcludedComputerNames {
    <#
    .SYNOPSIS
    Returns a HashSet of excluded computer names for fast O(1) lookups during scanning.
    #>
    $excludedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $Script:ComputerExclusions) {
        [void]$excludedSet.Add($item.ComputerName)
    }
    return $excludedSet
}

function Add-ComputerExclusion {
    <#
    .SYNOPSIS
    Adds a computer to the exclusion list.
    #>
    param(
        [string]$ComputerName,
        [string]$ExclusionType = "EXCLUDED",
        [string]$Reason = "User excluded"
    )

    $computerName = $ComputerName.Trim().ToUpper()
    if ([string]::IsNullOrWhiteSpace($computerName)) { return $false }

    # Check if already exists
    $existing = $Script:ComputerExclusions | Where-Object { $_.ComputerName -eq $computerName }
    if ($existing) {
        Write-Log "Computer '$computerName' is already in exclusion list" "WARNING"
        return $false
    }

    $item = New-ComputerExclusionItem -ComputerName $computerName -ExclusionType $ExclusionType -Reason $Reason
    $Script:ComputerExclusions.Add($item)
    Write-Log "Added '$computerName' to computer exclusions as $ExclusionType"
    return $true
}

function Remove-ComputerExclusion {
    <#
    .SYNOPSIS
    Removes selected computers from the exclusion list.
    #>
    param([array]$SelectedItems)

    $removed = 0
    foreach ($item in $SelectedItems) {
        $toRemove = $Script:ComputerExclusions | Where-Object { $_.ComputerName -eq $item.ComputerName }
        if ($toRemove) {
            $Script:ComputerExclusions.Remove($toRemove)
            $removed++
            Write-Log "Removed '$($item.ComputerName)' from computer exclusions"
        }
    }
    return $removed
}

function Test-ExcludeAccount {
    param([string]$AccountName)

    $patterns = Get-ExclusionPatterns
    foreach ($pattern in $patterns) {
        if ($AccountName -match $pattern) { return $true }
    }
    return $false
}

function Get-OutputPath {
    <#
    .SYNOPSIS
    Returns a file path within a date-based subfolder, creating the folder if needed.
    #>
    param([string]$FileName)

    $dateFolder = Join-Path $Script:OutputDirectory (Get-Date -Format 'yyyy-MM-dd')
    if (-not (Test-Path $dateFolder)) {
        New-Item -Path $dateFolder -ItemType Directory -Force | Out-Null
        Write-Log "Created output folder: $dateFolder"
    }
    return Join-Path $dateFolder $FileName
}

function Invoke-OutputCleanup {
    <#
    .SYNOPSIS
    Removes date-based output folders older than the specified retention period.
    #>
    param([int]$RetentionDays = 30)

    if (-not (Test-Path $Script:OutputDirectory)) { return }

    $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
    $dateFolders = Get-ChildItem -Path $Script:OutputDirectory -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' }

    $cleanedCount = 0
    foreach ($folder in $dateFolders) {
        try {
            $folderDate = [DateTime]::ParseExact($folder.Name, 'yyyy-MM-dd', $null)
            if ($folderDate -lt $cutoffDate) {
                Remove-Item -Path $folder.FullName -Recurse -Force
                $cleanedCount++
                Write-Log "Cleaned up old scan folder: $($folder.Name)"
            }
        } catch {
            # Skip folders that don't parse as dates
        }
    }

    if ($cleanedCount -gt 0) {
        Write-Log "Cleanup complete: removed $cleanedCount folder(s) older than $RetentionDays days"
    }
}

function Export-ResultsToCSV {
    param([string]$FilePath)

    $Script:ScanResults | Sort-Object -Property RiskScore -Descending |
        Select-Object RiskLevel, RiskScore, ComputerName, ComputerType, OperatingSystem, AdminName, ObjectClass |
        Export-Csv -Path $FilePath -NoTypeInformation

    Write-Log "Results exported to: $FilePath"
}

function Export-ResultsToHTML {
    param([string]$FilePath)

    # Calculate risk distribution
    $highRisk = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'HIGH' }).Count
    $mediumRisk = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'MEDIUM' }).Count
    $lowRisk = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'LOW' }).Count

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Local Administrator Audit Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #1e293b; margin-bottom: 10px; }
        .subtitle { color: #64748b; margin-bottom: 30px; }
        .stats { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
        .stat-card { background: #f8fafc; padding: 20px; border-radius: 8px; flex: 1; min-width: 150px; }
        .stat-value { font-size: 32px; font-weight: bold; color: #1e293b; }
        .stat-label { color: #64748b; font-size: 12px; text-transform: uppercase; }
        .risk-stats { display: flex; gap: 15px; margin-bottom: 30px; }
        .risk-badge { padding: 8px 16px; border-radius: 6px; font-weight: bold; font-size: 14px; }
        .risk-high { background: #fee2e2; color: #ef4444; }
        .risk-medium { background: #fef3c7; color: #f59e0b; }
        .risk-low { background: #dbeafe; color: #3b82f6; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background: #0EA5E9; color: #0D1117; padding: 12px; text-align: left; }
        td { padding: 12px; border-bottom: 1px solid #e2e8f0; }
        tr:hover { background: #f8fafc; }
        tr.row-high { background: #fef2f2; }
        tr.row-medium { background: #fffbeb; }
        .badge { display: inline-block; padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: bold; }
        .badge-high { background: #ef4444; color: white; }
        .badge-medium { background: #f59e0b; color: white; }
        .badge-low { background: #3b82f6; color: white; }
        .badge-info { background: #6b7280; color: white; }
        .warning { color: #ef4444; font-weight: bold; }
        .timestamp { color: #94a3b8; font-size: 12px; margin-top: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Local Administrator Audit Report</h1>
        <p class="subtitle">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>

        <div class="stats">
            <div class="stat-card">
                <div class="stat-value">$($Script:AllComputers.Count)</div>
                <div class="stat-label">Total Computers</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">$(($Script:ScanResults | Select-Object -ExpandProperty ComputerName -Unique).Count)</div>
                <div class="stat-label">Computers Reached</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">$($Script:UnreachableComputers.Count)</div>
                <div class="stat-label">Unreachable</div>
            </div>
            <div class="stat-card">
                <div class="stat-value" $(if ($Script:ScanResults.Count -gt 0) { 'class="warning"' })>$($Script:ScanResults.Count)</div>
                <div class="stat-label">Unexpected Admins</div>
            </div>
        </div>

        <h3>Risk Distribution</h3>
        <div class="risk-stats">
            <span class="risk-badge risk-high">HIGH: $highRisk</span>
            <span class="risk-badge risk-medium">MEDIUM: $mediumRisk</span>
            <span class="risk-badge risk-low">LOW: $lowRisk</span>
        </div>

        <h2>Unexpected Local Administrators</h2>
        <table>
            <tr>
                <th>Risk</th>
                <th>Score</th>
                <th>Computer Name</th>
                <th>Type</th>
                <th>Operating System</th>
                <th>Admin Account</th>
                <th>Account Type</th>
            </tr>
            $($Script:ScanResults | Sort-Object -Property RiskScore -Descending | ForEach-Object {
                $rowClass = switch ($_.RiskLevel) { 'HIGH' { 'row-high' }; 'MEDIUM' { 'row-medium' }; default { '' } }
                $badgeClass = switch ($_.RiskLevel) { 'HIGH' { 'badge-high' }; 'MEDIUM' { 'badge-medium' }; 'LOW' { 'badge-low' }; default { 'badge-info' } }
                "<tr class='$rowClass'><td><span class='badge $badgeClass'>$($_.RiskLevel)</span></td><td>$($_.RiskScore)</td><td>$($_.ComputerName)</td><td>$($_.ComputerType)</td><td>$($_.OperatingSystem)</td><td>$($_.AdminName)</td><td>$($_.ObjectClass)</td></tr>"
            })
        </table>

        <p class="timestamp">Report generated by Local Administrator Auditor v9.0</p>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $FilePath -Encoding UTF8
    Write-Log "HTML report exported to: $FilePath"
}

function Export-ExecutiveSummary {
    param([string]$FilePath)

    # Calculate statistics
    $totalComputers = $Script:AllComputers.Count
    $reachedComputers = ($Script:ScanResults | Select-Object -ExpandProperty ComputerName -Unique).Count
    $unreachableCount = $Script:UnreachableComputers.Count
    $totalFindings = $Script:ScanResults.Count
    $coveragePercent = if ($totalComputers -gt 0) { [math]::Round(($reachedComputers / $totalComputers) * 100, 1) } else { 0 }

    # Risk distribution
    $highRisk = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'HIGH' }).Count
    $mediumRisk = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'MEDIUM' }).Count
    $lowRisk = ($Script:ScanResults | Where-Object { $_.RiskLevel -eq 'LOW' }).Count

    # Top 5 affected computers
    $topComputers = $Script:ScanResults | Group-Object -Property ComputerName |
        Sort-Object -Property Count -Descending |
        Select-Object -First 5 -Property Name, Count, @{N='MaxRisk';E={
            ($_.Group | Sort-Object -Property RiskScore -Descending | Select-Object -First 1).RiskLevel
        }}

    # Top 5 unexpected accounts
    $topAccounts = $Script:ScanResults | Group-Object -Property AdminName |
        Sort-Object -Property Count -Descending |
        Select-Object -First 5 -Property Name, Count, @{N='MaxRisk';E={
            ($_.Group | Sort-Object -Property RiskScore -Descending | Select-Object -First 1).RiskLevel
        }}

    # Generate recommendations based on findings
    $recommendations = @()
    if ($highRisk -gt 0) {
        $recommendations += "CRITICAL: Review and remediate $highRisk HIGH risk findings immediately - these include local user accounts with admin privileges on servers."
    }
    if ($mediumRisk -gt 0) {
        $recommendations += "Review $mediumRisk MEDIUM risk findings - investigate domain users with local admin access to verify business need."
    }
    if ($unreachableCount -gt 0) {
        $recommendations += "Investigate $unreachableCount unreachable computers - ensure WinRM is enabled and verify network connectivity."
    }
    if (($Script:ScanResults | Where-Object { $_.ObjectClass -eq 'Group' }).Count -gt 0) {
        $groupCount = ($Script:ScanResults | Where-Object { $_.ObjectClass -eq 'Group' }).Count
        $recommendations += "Review $groupCount group memberships in local admin - verify membership and consider using direct user assignments."
    }
    if ($recommendations.Count -eq 0) {
        $recommendations += "No critical issues identified. Continue regular monitoring and auditing."
    }

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Executive Summary - Local Administrator Audit</title>
    <style>
        @media print {
            body { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f8fafc; color: #1e293b; line-height: 1.5; }
        .page { max-width: 900px; margin: 0 auto; background: white; padding: 40px; }
        .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #7c3aed; padding-bottom: 20px; margin-bottom: 30px; }
        .header h1 { color: #1e293b; font-size: 28px; font-weight: 700; }
        .header .logo { font-size: 16px; color: #64748b; font-weight: 600; }
        .header .date { font-size: 14px; color: #94a3b8; }
        .metrics { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }
        .metric-card { background: #f8fafc; padding: 20px; border-radius: 10px; text-align: center; border: 1px solid #e2e8f0; }
        .metric-value { font-size: 36px; font-weight: 700; color: #1e293b; }
        .metric-value.warning { color: #ef4444; }
        .metric-value.success { color: #10b981; }
        .metric-label { font-size: 12px; color: #64748b; text-transform: uppercase; font-weight: 600; margin-top: 5px; }
        .section { margin-bottom: 25px; }
        .section-title { font-size: 16px; font-weight: 700; color: #1e293b; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e2e8f0; }
        .risk-bar { display: flex; gap: 10px; margin-bottom: 20px; }
        .risk-item { display: flex; align-items: center; gap: 8px; padding: 10px 15px; border-radius: 8px; }
        .risk-high { background: #fee2e2; color: #ef4444; }
        .risk-medium { background: #fef3c7; color: #f59e0b; }
        .risk-low { background: #dbeafe; color: #3b82f6; }
        .risk-count { font-size: 24px; font-weight: 700; }
        .risk-label { font-size: 12px; font-weight: 600; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 13px; }
        th { background: #f1f5f9; color: #475569; padding: 10px 12px; text-align: left; font-weight: 600; }
        td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; }
        tr:hover { background: #f8fafc; }
        .badge { display: inline-block; padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 700; }
        .badge-high { background: #ef4444; color: white; }
        .badge-medium { background: #f59e0b; color: white; }
        .badge-low { background: #3b82f6; color: white; }
        .recommendations { background: #fefce8; border: 1px solid #fef08a; padding: 20px; border-radius: 10px; }
        .recommendations ul { margin-left: 20px; margin-top: 10px; }
        .recommendations li { margin-bottom: 8px; color: #713f12; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #e2e8f0; text-align: center; color: #94a3b8; font-size: 12px; }
    </style>
</head>
<body>
    <div class="page">
        <div class="header">
            <div>
                <h1>Local Administrator Audit</h1>
                <div class="logo">Executive Summary</div>
            </div>
            <div class="date">Generated: $(Get-Date -Format 'MMMM d, yyyy h:mm tt')</div>
        </div>

        <div class="metrics">
            <div class="metric-card">
                <div class="metric-value">$totalComputers</div>
                <div class="metric-label">Total Computers</div>
            </div>
            <div class="metric-card">
                <div class="metric-value success">$coveragePercent%</div>
                <div class="metric-label">Coverage Rate</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">$unreachableCount</div>
                <div class="metric-label">Unreachable</div>
            </div>
            <div class="metric-card">
                <div class="metric-value $(if ($totalFindings -gt 0) { 'warning' })">$totalFindings</div>
                <div class="metric-label">Total Findings</div>
            </div>
        </div>

        <div class="section">
            <div class="section-title">Risk Distribution</div>
            <div class="risk-bar">
                <div class="risk-item risk-high">
                    <span class="risk-count">$highRisk</span>
                    <span class="risk-label">HIGH</span>
                </div>
                <div class="risk-item risk-medium">
                    <span class="risk-count">$mediumRisk</span>
                    <span class="risk-label">MEDIUM</span>
                </div>
                <div class="risk-item risk-low">
                    <span class="risk-count">$lowRisk</span>
                    <span class="risk-label">LOW</span>
                </div>
            </div>
        </div>

        <div class="section">
            <div class="section-title">Top 5 Affected Computers</div>
            <table>
                <tr><th>Computer Name</th><th>Findings</th><th>Highest Risk</th></tr>
                $($topComputers | ForEach-Object {
                    $badgeClass = switch ($_.MaxRisk) { 'HIGH' { 'badge-high' }; 'MEDIUM' { 'badge-medium' }; default { 'badge-low' } }
                    "<tr><td>$($_.Name)</td><td>$($_.Count)</td><td><span class='badge $badgeClass'>$($_.MaxRisk)</span></td></tr>"
                })
                $(if ($topComputers.Count -eq 0) { "<tr><td colspan='3' style='text-align:center;color:#94a3b8;'>No findings to display</td></tr>" })
            </table>
        </div>

        <div class="section">
            <div class="section-title">Top 5 Unexpected Accounts</div>
            <table>
                <tr><th>Account Name</th><th>Occurrences</th><th>Highest Risk</th></tr>
                $($topAccounts | ForEach-Object {
                    $badgeClass = switch ($_.MaxRisk) { 'HIGH' { 'badge-high' }; 'MEDIUM' { 'badge-medium' }; default { 'badge-low' } }
                    "<tr><td>$($_.Name)</td><td>$($_.Count)</td><td><span class='badge $badgeClass'>$($_.MaxRisk)</span></td></tr>"
                })
                $(if ($topAccounts.Count -eq 0) { "<tr><td colspan='3' style='text-align:center;color:#94a3b8;'>No findings to display</td></tr>" })
            </table>
        </div>

        <div class="section">
            <div class="section-title">Recommended Actions</div>
            <div class="recommendations">
                <ul>
                    $($recommendations | ForEach-Object { "<li>$_</li>" })
                </ul>
            </div>
        </div>

        <div class="footer">
            Report generated by Local Administrator Auditor v9.0 | github.com/hugh2024/powershell-tools
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $FilePath -Encoding UTF8
    Write-Log "Executive Summary exported to: $FilePath"
}

#region JSON Persistence Functions
function Initialize-AppDataFolder {
    <#
    .SYNOPSIS
    Creates the application data folder if it doesn't exist.
    #>
    if (-not (Test-Path $Script:AppDataPath)) {
        try {
            New-Item -Path $Script:AppDataPath -ItemType Directory -Force | Out-Null
            Write-Log "Created app data folder: $Script:AppDataPath"
        } catch {
            Write-Log "Failed to create app data folder: $_" "ERROR"
        }
    }
}

function Get-DefaultSettings {
    <#
    .SYNOPSIS
    Returns default settings object.
    #>
    return @{
        version = "1.1"
        scanSettings = @{
            defaultTarget = "Servers Only"
            defaultThrottle = 32
            timeout = 30
        }
        exclusionPatterns = @{
            excludeFadmin = $false
            excludeDomainAdmins = $true
            excludeBuiltinAdmin = $true
            customPatterns = @()
        }
        computerExclusions = @()
        outputSettings = @{
            outputDirectory = $Script:OutputDirectory
            autoExportCSV = $true
            autoExportUnreachable = $true
            retentionDays = 30
        }
        uiPreferences = @{
            playSound = $true
            windowWidth = 1200
            windowHeight = 800
        }
        riskScoring = @{
            approvedServiceAccounts = @()
        }
    }
}

function Load-UserSettings {
    <#
    .SYNOPSIS
    Loads user settings from JSON file.
    #>
    if (Test-Path $Script:SettingsFile) {
        try {
            $json = Get-Content $Script:SettingsFile -Raw | ConvertFrom-Json

            # Load computer exclusions (with migration support for older settings)
            $computerExclusions = @()
            if ($json.computerExclusions) {
                foreach ($exc in $json.computerExclusions) {
                    $computerExclusions += @{
                        computerName = $exc.computerName
                        exclusionType = $exc.exclusionType
                        reason = $exc.reason
                    }
                }
            } else {
                # Migration: ship with empty exclusion list for new installs
                $computerExclusions = @()
            }

            $Script:UserSettings = @{
                version = $json.version
                scanSettings = @{
                    defaultTarget = $json.scanSettings.defaultTarget
                    defaultThrottle = $json.scanSettings.defaultThrottle
                    timeout = $json.scanSettings.timeout
                }
                exclusionPatterns = @{
                    excludeFadmin = $json.exclusionPatterns.excludeFadmin
                    excludeDomainAdmins = $json.exclusionPatterns.excludeDomainAdmins
                    excludeBuiltinAdmin = $json.exclusionPatterns.excludeBuiltinAdmin
                    customPatterns = @($json.exclusionPatterns.customPatterns)
                }
                computerExclusions = $computerExclusions
                outputSettings = @{
                    outputDirectory = $json.outputSettings.outputDirectory
                    autoExportCSV = $json.outputSettings.autoExportCSV
                    autoExportUnreachable = $json.outputSettings.autoExportUnreachable
                    retentionDays = if ($json.outputSettings.retentionDays) { $json.outputSettings.retentionDays } else { 30 }
                }
                uiPreferences = @{
                    playSound = $json.uiPreferences.playSound
                    windowWidth = $json.uiPreferences.windowWidth
                    windowHeight = $json.uiPreferences.windowHeight
                }
                riskScoring = @{
                    approvedServiceAccounts = @($json.riskScoring.approvedServiceAccounts)
                }
            }
            Write-Log "Settings loaded from: $Script:SettingsFile"
            return $true
        } catch {
            Write-Log "Failed to load settings: $_" "WARNING"
            $Script:UserSettings = Get-DefaultSettings
            return $false
        }
    } else {
        $Script:UserSettings = Get-DefaultSettings
        return $false
    }
}

function Save-UserSettings {
    <#
    .SYNOPSIS
    Saves current user settings to JSON file.
    #>
    try {
        Initialize-AppDataFolder

        # Gather current settings from UI
        $targetIndex = $Controls['cmbScanTarget'].SelectedIndex
        $targetMap = @(0, 1, 2)
        $targetNames = @("Servers Only", "Workstations Only", "Both (Servers + Workstations)")

        $throttleIndex = $Controls['cmbThrottle'].SelectedIndex
        $throttleValues = @(32, 16, 64, 128)

        $customPatterns = @()
        if (-not [string]::IsNullOrWhiteSpace($Controls['txtCustomExclusions'].Text)) {
            $customPatterns = @($Controls['txtCustomExclusions'].Text -split "`r`n" | Where-Object { $_.Trim() -ne '' })
        }

        $settings = @{
            version = "1.0"
            scanSettings = @{
                defaultTarget = $targetNames[$targetIndex]
                defaultThrottle = $throttleValues[$throttleIndex]
                timeout = [int]$Controls['txtTimeout'].Text
            }
            exclusionPatterns = @{
                excludeFadmin = [bool]$Controls['chkExcludeFadmin'].IsChecked
                excludeDomainAdmins = [bool]$Controls['chkExcludeDomainAdmins'].IsChecked
                excludeBuiltinAdmin = [bool]$Controls['chkExcludeBuiltinAdmin'].IsChecked
                fadminPattern = $Controls['txtFadminPattern'].Text.Trim()
                customPatterns = $customPatterns
            }
            outputSettings = @{
                outputDirectory = $Script:OutputDirectory
                autoExportCSV = [bool]$Controls['chkAutoExportCSV'].IsChecked
                autoExportUnreachable = [bool]$Controls['chkAutoExportUnreachable'].IsChecked
                retentionDays = if ($Script:UserSettings.outputSettings.retentionDays) { $Script:UserSettings.outputSettings.retentionDays } else { 30 }
            }
            uiPreferences = @{
                playSound = [bool]$Controls['chkPlaySound'].IsChecked
                # Only save window size if not maximized (to prevent spanning multi-monitor sizes)
                windowWidth = if ($Window.WindowState -eq 'Normal') { [int]$Window.ActualWidth } else { 1200 }
                windowHeight = if ($Window.WindowState -eq 'Normal') { [int]$Window.ActualHeight } else { 800 }
            }
            riskScoring = @{
                approvedServiceAccounts = @(
                    if (-not [string]::IsNullOrWhiteSpace($Controls['txtApprovedAccounts'].Text)) {
                        $Controls['txtApprovedAccounts'].Text -split "`r`n" | Where-Object { $_.Trim() -ne '' }
                    }
                )
            }
            computerExclusions = @(
                foreach ($item in $Script:ComputerExclusions) {
                    @{
                        computerName = $item.ComputerName
                        exclusionType = $item.ExclusionType
                        reason = $item.Reason
                    }
                }
            )
        }

        $settings | ConvertTo-Json -Depth 5 | Out-File $Script:SettingsFile -Encoding UTF8
        Write-Log "Settings saved to: $Script:SettingsFile"
    } catch {
        Write-Log "Failed to save settings: $_" "ERROR"
    }
}

function Apply-UserSettings {
    <#
    .SYNOPSIS
    Applies loaded settings to UI controls.
    #>
    if ($null -eq $Script:UserSettings) { return }

    try {
        # Apply scan target
        $targetNames = @("Servers Only", "Workstations Only", "Both (Servers + Workstations)")
        $targetIndex = [array]::IndexOf($targetNames, $Script:UserSettings.scanSettings.defaultTarget)
        if ($targetIndex -ge 0) { $Controls['cmbScanTarget'].SelectedIndex = $targetIndex }

        # Apply throttle
        $throttleValues = @("32", "16", "64", "128")
        $throttleIndex = [array]::IndexOf($throttleValues, $Script:UserSettings.scanSettings.defaultThrottle.ToString())
        if ($throttleIndex -ge 0) { $Controls['cmbThrottle'].SelectedIndex = $throttleIndex }

        # Apply timeout
        $Controls['txtTimeout'].Text = $Script:UserSettings.scanSettings.timeout.ToString()

        # Apply exclusion patterns
        $null = $Controls['chkExcludeFadmin'].IsChecked = $Script:UserSettings.exclusionPatterns.excludeFadmin
        $null = $Controls['chkExcludeDomainAdmins'].IsChecked = $Script:UserSettings.exclusionPatterns.excludeDomainAdmins
        $null = $Controls['chkExcludeBuiltinAdmin'].IsChecked = $Script:UserSettings.exclusionPatterns.excludeBuiltinAdmin
        if ($Script:UserSettings.exclusionPatterns.fadminPattern) {
            $Controls['txtFadminPattern'].Text = $Script:UserSettings.exclusionPatterns.fadminPattern
        }

        if ($Script:UserSettings.exclusionPatterns.customPatterns.Count -gt 0) {
            $Controls['txtCustomExclusions'].Text = $Script:UserSettings.exclusionPatterns.customPatterns -join "`r`n"
        }

        # Apply output settings
        if (-not [string]::IsNullOrEmpty($Script:UserSettings.outputSettings.outputDirectory)) {
            if (Test-Path $Script:UserSettings.outputSettings.outputDirectory) {
                $Script:OutputDirectory = $Script:UserSettings.outputSettings.outputDirectory
                $Controls['txtOutputDir'].Text = $Script:OutputDirectory
            }
        }
        $null = $Controls['chkAutoExportCSV'].IsChecked = $Script:UserSettings.outputSettings.autoExportCSV
        $null = $Controls['chkAutoExportUnreachable'].IsChecked = $Script:UserSettings.outputSettings.autoExportUnreachable

        # Apply UI preferences
        $null = $Controls['chkPlaySound'].IsChecked = $Script:UserSettings.uiPreferences.playSound

        # Apply window size (with bounds checking to prevent multi-monitor spanning)
        $screenWidth = [System.Windows.SystemParameters]::PrimaryScreenWidth
        $screenHeight = [System.Windows.SystemParameters]::PrimaryScreenHeight
        $savedWidth = $Script:UserSettings.uiPreferences.windowWidth
        $savedHeight = $Script:UserSettings.uiPreferences.windowHeight

        if ($savedWidth -gt 0 -and $savedWidth -le $screenWidth) {
            $Window.Width = $savedWidth
        }
        if ($savedHeight -gt 0 -and $savedHeight -le $screenHeight) {
            $Window.Height = $savedHeight
        }

        # Apply approved service accounts
        if ($Script:UserSettings.riskScoring.approvedServiceAccounts.Count -gt 0) {
            $Controls['txtApprovedAccounts'].Text = $Script:UserSettings.riskScoring.approvedServiceAccounts -join "`r`n"
        }

        # Apply computer exclusions
        $Script:ComputerExclusions.Clear()
        if ($Script:UserSettings.computerExclusions.Count -gt 0) {
            foreach ($exc in $Script:UserSettings.computerExclusions) {
                $exclusionItem = New-ComputerExclusionItem -ComputerName $exc.computerName -ExclusionType $exc.exclusionType -Reason $exc.reason
                $Script:ComputerExclusions.Add($exclusionItem)
            }
        }
        $Controls['lstComputerExclusions'].ItemsSource = $Script:ComputerExclusions
        Write-Log "Loaded $($Script:ComputerExclusions.Count) computer exclusion(s)"

        Write-Log "User settings applied"
    } catch {
        Write-Log "Error applying settings: $_" "WARNING"
    }
}

function Load-ScanHistory {
    <#
    .SYNOPSIS
    Loads scan history from JSON file.
    #>
    if (Test-Path $Script:HistoryFile) {
        try {
            $json = Get-Content $Script:HistoryFile -Raw | ConvertFrom-Json

            # Convert JSON to hashtable structure
            $Script:ScanHistory = @{
                version = $json.version
                scans = @()
                knownFindings = @{}
            }

            # Convert scans array
            foreach ($scan in $json.scans) {
                $scanEntry = @{
                    scanId = $scan.scanId
                    timestamp = $scan.timestamp
                    scanTarget = $scan.scanTarget
                    statistics = @{
                        total = $scan.statistics.total
                        reached = $scan.statistics.reached
                        unreachable = $scan.statistics.unreachable
                        findings = $scan.statistics.findings
                        highRisk = $scan.statistics.highRisk
                        mediumRisk = $scan.statistics.mediumRisk
                        lowRisk = $scan.statistics.lowRisk
                    }
                    results = @($scan.results)
                    unreachable = @($scan.unreachable)
                }
                $Script:ScanHistory.scans += $scanEntry
            }

            # Convert known findings
            if ($json.knownFindings) {
                foreach ($prop in $json.knownFindings.PSObject.Properties) {
                    $Script:ScanHistory.knownFindings[$prop.Name] = @{
                        firstSeen = $prop.Value.firstSeen
                        lastSeen = $prop.Value.lastSeen
                        acknowledged = $prop.Value.acknowledged
                    }
                }
            }

            Write-Log "Scan history loaded: $($Script:ScanHistory.scans.Count) previous scans"
            return $true
        } catch {
            Write-Log "Failed to load scan history: $_" "WARNING"
            $Script:ScanHistory = @{ version = "1.0"; scans = @(); knownFindings = @{} }
            return $false
        }
    }
    return $false
}

function Save-ScanHistory {
    <#
    .SYNOPSIS
    Saves current scan to history JSON file.
    #>
    param(
        [string]$ScanTarget,
        [array]$Results,
        [array]$Unreachable,
        [hashtable]$Statistics
    )

    try {
        Initialize-AppDataFolder

        $scanId = [guid]::NewGuid().ToString()
        $timestamp = Get-Date -Format "o"  # ISO 8601 format

        # Create scan entry
        $scanEntry = @{
            scanId = $scanId
            timestamp = $timestamp
            scanTarget = $ScanTarget
            statistics = @{
                total = $Statistics.total
                reached = $Statistics.reached
                unreachable = $Statistics.unreachable
                findings = $Statistics.findings
                highRisk = $Statistics.highRisk
                mediumRisk = $Statistics.mediumRisk
                lowRisk = $Statistics.lowRisk
            }
            results = @($Results | ForEach-Object {
                @{
                    computerName = $_.ComputerName
                    computerType = $_.ComputerType
                    operatingSystem = $_.OperatingSystem
                    adminName = $_.AdminName
                    objectClass = $_.ObjectClass
                    riskScore = $_.RiskScore
                    riskLevel = $_.RiskLevel
                }
            })
            unreachable = @($Unreachable | ForEach-Object {
                @{
                    computerName = $_.ComputerName
                    computerType = $_.ComputerType
                    errorMessage = $_.ErrorMessage
                }
            })
        }

        # Update known findings
        foreach ($result in $Results) {
            $key = "$($result.ComputerName)|$($result.AdminName)"
            if (-not $Script:ScanHistory.knownFindings.ContainsKey($key)) {
                $Script:ScanHistory.knownFindings[$key] = @{
                    firstSeen = $timestamp
                    lastSeen = $timestamp
                    acknowledged = $false
                }
            } else {
                $Script:ScanHistory.knownFindings[$key].lastSeen = $timestamp
            }
        }

        # Add scan to history (keep last 50 scans)
        $Script:ScanHistory.scans += $scanEntry
        if ($Script:ScanHistory.scans.Count -gt 50) {
            $Script:ScanHistory.scans = $Script:ScanHistory.scans | Select-Object -Last 50
        }

        # Save to file
        $Script:ScanHistory | ConvertTo-Json -Depth 10 | Out-File $Script:HistoryFile -Encoding UTF8
        Write-Log "Scan saved to history (ID: $scanId)"

        return $scanId
    } catch {
        Write-Log "Failed to save scan history: $_" "ERROR"
        return $null
    }
}


function Refresh-HistoryGrid {
    <#
    .SYNOPSIS
    Refreshes the history DataGrid by scanning date folders for export files.
    #>
    $historyItems = @()
    $baseDir = $Script:OutputDirectory

    # Find all date folders (YYYY-MM-DD format)
    $dateFolders = Get-ChildItem -Path $baseDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' }

    foreach ($folder in $dateFolders) {
        # Find all CSV and HTML files in each date folder
        $files = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in '.csv', '.html' }

        foreach ($file in $files) {
            # Determine file type
            $fileType = switch -Regex ($file.Name) {
                'ExecutiveSummary' { 'Executive Summary' }
                '\.csv$' { 'CSV Export' }
                '\.html$' { 'HTML Export' }
                default { 'Export' }
            }

            # Parse time from filename (e.g., LocalAdminAudit_145122.csv -> 14:51:22)
            $timeMatch = [regex]::Match($file.BaseName, '_(\d{2})(\d{2})(\d{2})$')
            $timeStr = if ($timeMatch.Success) {
                "$($timeMatch.Groups[1].Value):$($timeMatch.Groups[2].Value):$($timeMatch.Groups[3].Value)"
            } else {
                $file.LastWriteTime.ToString("HH:mm:ss")
            }

            # Combine folder date with file time
            $dateTimeStr = "$($folder.Name) $timeStr"
            $dateTime = try { [DateTime]::Parse($dateTimeStr) } catch { $file.LastWriteTime }

            # Get row count for CSV files
            $rowCount = ""
            if ($file.Extension -eq '.csv') {
                try {
                    $lineCount = (Get-Content $file.FullName -ErrorAction SilentlyContinue | Measure-Object -Line).Lines
                    $rowCount = ($lineCount - 1).ToString()  # Subtract header row
                } catch { $rowCount = "?" }
            }

            $historyItems += [PSCustomObject]@{
                FileName = $file.Name
                DateFormatted = $dateTime.ToString("MMM d, yyyy h:mm tt")
                Timestamp = $dateTime
                FileType = $fileType
                RowCount = $rowCount
                FileSize = "{0:N0} KB" -f ($file.Length / 1KB)
                FolderDate = $folder.Name
                FullPath = $file.FullName
            }
        }
    }

    # Sort by date descending (most recent first)
    $historyItems = @($historyItems | Sort-Object -Property Timestamp -Descending)

    $Controls['dgHistory'].ItemsSource = $historyItems
    $Controls['txtHistoryInfo'].Text = "$($historyItems.Count) export file$(if ($historyItems.Count -ne 1) {'s'}) found"
}

function Load-HistoricalScan {
    <#
    .SYNOPSIS
    Loads a historical scan into the Results tab.
    #>
    param([string]$ScanId)

    $scan = $Script:ScanHistory.scans | Where-Object { $_.scanId -eq $ScanId }
    if (-not $scan) {
        Write-Log "Scan not found: $ScanId" "ERROR"
        return
    }

    $dateTime = [DateTime]::Parse($scan.timestamp)
    Write-Log "Loading historical scan from $($dateTime.ToString('MMM d, yyyy h:mm tt'))"

    # Clear current results
    $Script:ScanResults.Clear()
    $Script:UnreachableComputers.Clear()

    # Load results
    foreach ($result in $scan.results) {
        $null = $Script:ScanResults.Add([PSCustomObject]@{
            ComputerName = $result.computerName
            ComputerType = $result.computerType
            OperatingSystem = $result.operatingSystem
            AdminName = $result.adminName
            ObjectClass = $result.objectClass
            RiskScore = $result.riskScore
            RiskLevel = $result.riskLevel
            RiskColor = switch ($result.riskLevel) { 'HIGH' { '#FF6B6B' }; 'MEDIUM' { '#FFD93D' }; 'LOW' { '#4ECDC4' }; default { '#A0AEC0' } }
            RiskBgColor = switch ($result.riskLevel) { 'HIGH' { '#3D1212' }; 'MEDIUM' { '#3D2E05' }; 'LOW' { '#0D2F2D' }; default { '#2D3748' } }
        })
    }

    # Load unreachable
    foreach ($unreachable in $scan.unreachable) {
        $null = $Script:UnreachableComputers.Add([PSCustomObject]@{
            ComputerName = $unreachable.computerName
            ComputerType = $unreachable.computerType
            OperatingSystem = ""
            ErrorMessage = $unreachable.errorMessage
        })
    }

    # Update UI
    $resultsArray = @($Script:ScanResults | Sort-Object -Property RiskScore -Descending)
    $unreachableArray = @($Script:UnreachableComputers)

    $Controls['dgResults'].ItemsSource = $null
    $Controls['dgResults'].ItemsSource = $resultsArray
    $Controls['dgUnreachable'].ItemsSource = $null
    $Controls['dgUnreachable'].ItemsSource = $unreachableArray

    # Update statistics
    $Controls['txtTotalComputers'].Text = $scan.statistics.total.ToString()
    $Controls['txtReached'].Text = $scan.statistics.reached.ToString()
    $Controls['txtUnreachable'].Text = $scan.statistics.unreachable.ToString()
    $Controls['txtUnexpectedAdmins'].Text = $scan.statistics.findings.ToString()
    $Controls['txtResultsCount'].Text = "$($scan.statistics.findings) results"
    $Controls['txtUnreachableCount'].Text = "$($scan.statistics.unreachable) unreachable"

    $highRisk = if ($scan.statistics.highRisk) { $scan.statistics.highRisk } else { 0 }
    $mediumRisk = if ($scan.statistics.mediumRisk) { $scan.statistics.mediumRisk } else { 0 }
    $lowRisk = if ($scan.statistics.lowRisk) { $scan.statistics.lowRisk } else { 0 }
    $Controls['txtHighRisk'].Text = "$highRisk HIGH"
    $Controls['txtMediumRisk'].Text = "$mediumRisk MED"
    $Controls['txtLowRisk'].Text = "$lowRisk LOW"

    # Switch to Results tab
    $tabControl = $Window.FindName("tabControl")
    if ($tabControl) { $tabControl.SelectedIndex = 1 }

    Write-Log "Loaded historical scan: $($scan.statistics.findings) findings"
    [System.Windows.MessageBox]::Show(
        "Loaded scan from $($dateTime.ToString('MMM d, yyyy h:mm tt'))`n`nFindings: $($scan.statistics.findings)`nUnreachable: $($scan.statistics.unreachable)",
        "Historical Scan Loaded",
        "OK",
        "Information"
    )
}
#endregion

function Update-DataGrids {
    $Window.Dispatcher.Invoke([Action]{
        # Update main results grid
        $Controls['dgResults'].ItemsSource = $null
        $Controls['dgResults'].ItemsSource = $Script:ScanResults

        # Update unreachable grid
        $Controls['dgUnreachable'].ItemsSource = $null
        $Controls['dgUnreachable'].ItemsSource = $Script:UnreachableComputers
    })
}

function Apply-ResultsFilter {
    $filterText = $Controls['txtResultsFilter'].Text.ToLower()
    $typeFilter = $Controls['cmbResultsTypeFilter'].SelectedItem.Content
    $groupBy = if ($Controls['cmbGroupBy'].SelectedItem) { $Controls['cmbGroupBy'].SelectedItem.Content } else { "None" }

    $filtered = $Script:ScanResults | Where-Object {
        $matchesText = ($_.ComputerName.ToLower().Contains($filterText) -or
                       $_.AdminName.ToLower().Contains($filterText))

        $matchesType = switch ($typeFilter) {
            "Servers" { $_.ComputerType -eq "Server" }
            "Workstations" { $_.ComputerType -eq "Workstation" }
            default { $true }
        }

        $matchesText -and $matchesType
    }

    # Apply sorting based on group selection
    $sortedFiltered = switch ($groupBy) {
        "Computer Name" { $filtered | Sort-Object -Property ComputerName, @{Expression={$_.RiskScore}; Descending=$true} }
        "Admin Account" { $filtered | Sort-Object -Property AdminName, @{Expression={$_.RiskScore}; Descending=$true} }
        "Risk Level" { $filtered | Sort-Object -Property @{Expression={
            switch ($_.RiskLevel) { 'HIGH' { 0 }; 'MEDIUM' { 1 }; 'LOW' { 2 }; default { 3 } }
        }}, @{Expression={$_.RiskScore}; Descending=$true} }
        "Account Type" { $filtered | Sort-Object -Property ObjectClass, @{Expression={$_.RiskScore}; Descending=$true} }
        default { $filtered | Sort-Object -Property RiskScore -Descending }
    }

    # Convert to objects with risk properties for proper display
    $filteredArray = @($sortedFiltered | ForEach-Object {
        [PSCustomObject]@{
            ComputerName = [string]$_.ComputerName
            ComputerType = [string]$_.ComputerType
            OperatingSystem = [string]$_.OperatingSystem
            AdminName = [string]$_.AdminName
            ObjectClass = [string]$_.ObjectClass
            RiskScore = [int]$_.RiskScore
            RiskLevel = [string]$_.RiskLevel
            RiskColor = [string]$_.RiskColor
            RiskBgColor = [string]$_.RiskBgColor
        }
    })

    $Controls['dgResults'].ItemsSource = $null
    $Controls['dgResults'].ItemsSource = $filteredArray

    # Update count with grouping info
    $countText = "$($filteredArray.Count) results"
    if ($groupBy -ne "None") {
        $countText += " (grouped by $groupBy)"
    }
    $Controls['txtResultsCount'].Text = $countText
}

function Test-ApprovedAccount {
    <#
    .SYNOPSIS
    Checks if an account matches any approved service account pattern.
    #>
    param([string]$AccountName)

    $approvedPatterns = @()
    if ($Controls['txtApprovedAccounts'] -and -not [string]::IsNullOrWhiteSpace($Controls['txtApprovedAccounts'].Text)) {
        $approvedPatterns = @($Controls['txtApprovedAccounts'].Text -split "`r`n" | Where-Object { $_.Trim() -ne '' })
    }

    if ($approvedPatterns.Count -eq 0) { return $false }

    $accountPart = ($AccountName -split '\\')[-1]  # Get just the account name without domain

    foreach ($pattern in $approvedPatterns) {
        $pattern = $pattern.Trim()
        if ([string]::IsNullOrEmpty($pattern)) { continue }

        # Try exact match first (case-insensitive)
        if ($AccountName -ieq $pattern -or $accountPart -ieq $pattern) {
            return $true
        }

        # Try regex match only - no wildcard *contains* to avoid false suppression
        try {
            if ($AccountName -match $pattern -or $accountPart -match $pattern) {
                return $true
            }
        } catch {
            # Invalid regex pattern - skip silently
        }
    }

    return $false
}

function Get-TrendData {
    <#
    .SYNOPSIS
    Analyze historical CSV files to build trend data.
    #>
    $trendData = @()

    # Scan for date folders in output directory
    $dateFolders = Get-ChildItem -Path $Script:OutputDirectory -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } |
        Sort-Object Name

    foreach ($folder in $dateFolders) {
        $csvFiles = Get-ChildItem -Path $folder.FullName -Filter "LocalAdminAudit*.csv" -ErrorAction SilentlyContinue

        foreach ($csvFile in $csvFiles) {
            try {
                $data = Import-Csv $csvFile.FullName -ErrorAction Stop

                # Count findings by risk level (handle different column name formats)
                $highRisk = @($data | Where-Object {
                    ($_.RiskLevel -eq 'HIGH') -or ($_.'Risk Level' -eq 'HIGH')
                }).Count
                $mediumRisk = @($data | Where-Object {
                    ($_.RiskLevel -eq 'MEDIUM') -or ($_.'Risk Level' -eq 'MEDIUM')
                }).Count
                $lowRisk = @($data | Where-Object {
                    ($_.RiskLevel -eq 'LOW') -or ($_.'Risk Level' -eq 'LOW')
                }).Count

                # Get unique computer count
                $computers = @($data | ForEach-Object {
                    if ($_.ComputerName) { $_.ComputerName } elseif ($_.'Computer Name') { $_.'Computer Name' }
                } | Select-Object -Unique).Count

                $trendData += [PSCustomObject]@{
                    Date = $folder.Name
                    SortDate = [datetime]::ParseExact($folder.Name, 'yyyy-MM-dd', $null)
                    ComputersScanned = $computers
                    TotalFindings = $data.Count
                    HighRisk = $highRisk
                    MediumRisk = $mediumRisk
                    LowRisk = $lowRisk
                    Change = ""  # Will be calculated after sorting
                    SourceFile = $csvFile.Name
                }
            } catch {
                Write-Log "Error reading CSV for trends: $($csvFile.FullName) - $_" "WARNING"
            }
        }
    }

    # Sort by date and calculate changes
    $trendData = $trendData | Sort-Object SortDate
    for ($i = 1; $i -lt $trendData.Count; $i++) {
        $diff = $trendData[$i].TotalFindings - $trendData[$i-1].TotalFindings
        $trendData[$i].Change = if ($diff -gt 0) { "+$diff" } elseif ($diff -lt 0) { "$diff" } else { "0" }
    }
    if ($trendData.Count -gt 0) { $trendData[0].Change = "N/A" }

    return $trendData
}

function Refresh-TrendDashboard {
    Write-Log "Refreshing trend data..."
    $trendData = Get-TrendData

    if ($trendData.Count -eq 0) {
        $Controls['txtTrendDateRange'].Text = "No historical data found"
        $Controls['txtTrendScanCount'].Text = "0"
        $Controls['txtTrendAvgFindings'].Text = "--"
        $Controls['txtTrendDirection'].Text = "--"
        $Controls['txtTrendHighRisk'].Text = "--"
        $Controls['dgTrends'].ItemsSource = $null
        return
    }

    # Update summary stats
    $Controls['txtTrendScanCount'].Text = $trendData.Count.ToString()

    $avgFindings = [math]::Round(($trendData | Measure-Object -Property TotalFindings -Average).Average, 1)
    $Controls['txtTrendAvgFindings'].Text = $avgFindings.ToString()

    # Calculate trend direction (compare first half to second half)
    if ($trendData.Count -ge 2) {
        $firstHalf = ($trendData[0..([math]::Floor($trendData.Count/2)-1)] | Measure-Object -Property TotalFindings -Average).Average
        $secondHalf = ($trendData[[math]::Floor($trendData.Count/2)..($trendData.Count-1)] | Measure-Object -Property TotalFindings -Average).Average

        if ($secondHalf -lt $firstHalf * 0.9) {
            $Controls['txtTrendDirection'].Text = "DOWN"
            $Controls['txtTrendDirection'].Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(16, 185, 129)))
        } elseif ($secondHalf -gt $firstHalf * 1.1) {
            $Controls['txtTrendDirection'].Text = "UP"
            $Controls['txtTrendDirection'].Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(239, 68, 68)))
        } else {
            $Controls['txtTrendDirection'].Text = "STABLE"
            $Controls['txtTrendDirection'].Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(148, 163, 184)))
        }
    } else {
        $Controls['txtTrendDirection'].Text = "N/A"
    }

    # High risk trend
    $avgHighRisk = [math]::Round(($trendData | Measure-Object -Property HighRisk -Average).Average, 1)
    $Controls['txtTrendHighRisk'].Text = $avgHighRisk.ToString()

    # Date range
    $minDate = ($trendData | Sort-Object SortDate | Select-Object -First 1).Date
    $maxDate = ($trendData | Sort-Object SortDate | Select-Object -Last 1).Date
    $Controls['txtTrendDateRange'].Text = "Data from $minDate to $maxDate ($($trendData.Count) scans)"

    # Update grid
    $Controls['dgTrends'].ItemsSource = $null
    $Controls['dgTrends'].ItemsSource = @($trendData | Sort-Object SortDate -Descending)

    Write-Log "Trend data refreshed: $($trendData.Count) historical scans found"
}

function Export-TrendChart {
    param([string]$FilePath)

    $trendData = Get-TrendData | Sort-Object SortDate

    if ($trendData.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No historical data available to chart.", "No Data", "OK", "Warning")
        return
    }

    # Build chart data for Chart.js
    $labels = ($trendData | ForEach-Object { "`"$($_.Date)`"" }) -join ","
    $totalData = ($trendData | ForEach-Object { $_.TotalFindings }) -join ","
    $highData = ($trendData | ForEach-Object { $_.HighRisk }) -join ","
    $mediumData = ($trendData | ForEach-Object { $_.MediumRisk }) -join ","
    $lowData = ($trendData | ForEach-Object { $_.LowRisk }) -join ","

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Local Admin Audit - Trend Chart</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1e1e2e; color: #f8fafc; padding: 30px; }
        .container { max-width: 1000px; margin: 0 auto; }
        h1 { color: #f8fafc; margin-bottom: 10px; }
        .subtitle { color: #94a3b8; margin-bottom: 30px; }
        .chart-container { background: #2a2a3e; border-radius: 12px; padding: 30px; margin-bottom: 30px; }
        .stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }
        .stat-card { background: #2a2a3e; border-radius: 8px; padding: 20px; text-align: center; }
        .stat-value { font-size: 32px; font-weight: bold; color: #f8fafc; }
        .stat-value.danger { color: #ef4444; }
        .stat-value.warning { color: #f59e0b; }
        .stat-value.success { color: #10b981; }
        .stat-label { font-size: 12px; color: #94a3b8; text-transform: uppercase; margin-top: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Local Administrator Audit - Trend Analysis</h1>
        <p class="subtitle">Generated: $(Get-Date -Format 'MMMM d, yyyy h:mm tt') | Scans Analyzed: $($trendData.Count)</p>

        <div class="stats">
            <div class="stat-card">
                <div class="stat-value">$($trendData.Count)</div>
                <div class="stat-label">Scans Analyzed</div>
            </div>
            <div class="stat-card">
                <div class="stat-value danger">$([math]::Round(($trendData | Measure-Object -Property HighRisk -Average).Average, 1))</div>
                <div class="stat-label">Avg High Risk</div>
            </div>
            <div class="stat-card">
                <div class="stat-value warning">$([math]::Round(($trendData | Measure-Object -Property TotalFindings -Average).Average, 1))</div>
                <div class="stat-label">Avg Total Findings</div>
            </div>
            <div class="stat-card">
                <div class="stat-value success">$([math]::Round(($trendData | Measure-Object -Property ComputersScanned -Average).Average, 0))</div>
                <div class="stat-label">Avg Computers</div>
            </div>
        </div>

        <div class="chart-container">
            <canvas id="trendChart"></canvas>
        </div>

        <div class="chart-container">
            <canvas id="riskChart"></canvas>
        </div>
    </div>

    <script>
        new Chart(document.getElementById('trendChart'), {
            type: 'line',
            data: {
                labels: [$labels],
                datasets: [{
                    label: 'Total Findings',
                    data: [$totalData],
                    borderColor: '#7c3aed',
                    backgroundColor: 'rgba(124, 58, 237, 0.1)',
                    fill: true,
                    tension: 0.3
                }]
            },
            options: {
                responsive: true,
                plugins: { title: { display: true, text: 'Total Findings Over Time', color: '#f8fafc', font: { size: 16 } } },
                scales: { y: { beginAtZero: true, ticks: { color: '#94a3b8' }, grid: { color: '#3f3f5a' } }, x: { ticks: { color: '#94a3b8' }, grid: { color: '#3f3f5a' } } }
            }
        });

        new Chart(document.getElementById('riskChart'), {
            type: 'bar',
            data: {
                labels: [$labels],
                datasets: [
                    { label: 'High Risk', data: [$highData], backgroundColor: '#ef4444' },
                    { label: 'Medium Risk', data: [$mediumData], backgroundColor: '#f59e0b' },
                    { label: 'Low Risk', data: [$lowData], backgroundColor: '#10b981' }
                ]
            },
            options: {
                responsive: true,
                plugins: { title: { display: true, text: 'Risk Distribution Over Time', color: '#f8fafc', font: { size: 16 } } },
                scales: { y: { beginAtZero: true, stacked: true, ticks: { color: '#94a3b8' }, grid: { color: '#3f3f5a' } }, x: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: '#3f3f5a' } } }
            }
        });
    </script>
</body>
</html>
"@

    $html | Out-File -FilePath $FilePath -Encoding UTF8
    Write-Log "Trend chart exported to: $FilePath"
}

function Test-CredentialError {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $msg = $ErrorRecord.Exception.Message
    return ($msg -match 'Access is denied' -or $msg -match 'The user name or password is incorrect' -or $msg -match 'Logon failure')
}

function Get-RiskScore {
    <#
    .SYNOPSIS
    Calculate risk score for a local admin finding.

    .DESCRIPTION
    Scoring algorithm:
    - Local user account: +40 points
    - Domain user (non-service): +30 points
    - Unapproved service account: +20 points
    - Server (vs workstation): +10 points
    - Group account: +15 points (multiple users could have access)
    - Approved service account: -50 points (reduces risk)

    Risk Levels:
    - High: 60+ points
    - Medium: 30-59 points
    - Low: 1-29 points
    - Info: 0 or less points
    #>
    param(
        [string]$AccountName,
        [string]$ObjectClass,
        [string]$ComputerType
    )

    $score = 0

    # Determine account type and add base score
    $isLocalAccount = $AccountName -notmatch '\\'  # No backslash means local account
    $isDomainAccount = $AccountName -match '\\'

    if ($ObjectClass -eq 'User') {
        if ($isLocalAccount) {
            # Local user account - highest risk
            $score += 40
        } elseif ($isDomainAccount) {
            # Check if it's a service account pattern
            $accountPart = ($AccountName -split '\\')[-1]
            $isServiceAccount = $accountPart -match '^(svc|service|sql|app|sys|task|batch|iis|mssql|backup|scan)' -or
                               $accountPart -match '(svc|service)$'

            if ($isServiceAccount) {
                $score += 20  # Service account - medium concern
            } else {
                $score += 30  # Domain user - high concern
            }
        }
    } elseif ($ObjectClass -eq 'Group') {
        # Group membership - moderate risk (multiple users could have access)
        $score += 15
    }

    # Server bonus - findings on servers are more critical
    if ($ComputerType -eq 'Server') {
        $score += 10
    }

    # Check if account is on approved list - reduce score significantly
    if (Test-ApprovedAccount -AccountName $AccountName) {
        $score -= 50
    }

    # Ensure score doesn't go below 0
    if ($score -lt 0) { $score = 0 }

    # Determine risk level and colors
    $riskLevel = switch ($score) {
        { $_ -ge 60 } { 'HIGH' }
        { $_ -ge 30 } { 'MEDIUM' }
        { $_ -ge 1 }  { 'LOW' }
        default       { 'INFO' }
    }

    # Risk colors (Vibrant Dark Theme - 8:1+ contrast)
    $riskColor = switch ($riskLevel) {
        'HIGH'   { '#FF6B6B' }
        'MEDIUM' { '#FFD93D' }
        'LOW'    { '#4ECDC4' }
        default  { '#A0AEC0' }
    }

    $riskBgColor = switch ($riskLevel) {
        'HIGH'   { '#3D1212' }
        'MEDIUM' { '#3D2E05' }
        'LOW'    { '#0D2F2D' }
        default  { '#2D3748' }
    }

    return @{
        RiskScore = $score
        RiskLevel = $riskLevel
        RiskColor = $riskColor
        RiskBgColor = $riskBgColor
    }
}
#endregion

#region Scan Functions
function Start-AuditScan {
    param(
        [switch]$ResumeMode
    )

    # Check if already scanning (button should be disabled, but double-check)
    if (-not $Controls['btnStartScan'].IsEnabled) { return }

    $Script:CancelRequested = $false
    if ($Script:SyncHash) { $Script:SyncHash.CancelRequested = $false }

    # In resume mode, don't clear existing results - they were restored from saved state
    if (-not $ResumeMode) {
        $Script:ScanResults.Clear()
        $Script:UnreachableComputers.Clear()
    }

    # Update UI
    $Controls['btnStartScan'].IsEnabled = $false
    $Controls['btnCancelScan'].IsEnabled = $true
    $Controls['btnElevate'].IsEnabled = $false
    $Controls['progressCard'].Visibility = "Visible"

    $scanTarget = $Controls['cmbScanTarget'].SelectedItem.Content
    $throttleLimit = [int]$Controls['cmbThrottle'].SelectedItem.Content

    # Store scan target for completion handler (which runs on main thread)
    $Script:LastScanTarget = $scanTarget

    # Store settings for Quick Rescan feature
    $Script:LastScanSettings = @{
        ScanTarget = $scanTarget
        ThrottleLimit = $throttleLimit
        SpecificComputers = $Controls['txtSpecificComputers'].Text
    }

    # PRE-CACHE all UI settings BEFORE creating runspace (avoids blocking Dispatcher.Invoke calls)
    $excludeFadmin = [bool]$Controls['chkExcludeFadmin'].IsChecked
    $excludeDomainAdmins = [bool]$Controls['chkExcludeDomainAdmins'].IsChecked
    $excludeBuiltinAdmin = [bool]$Controls['chkExcludeBuiltinAdmin'].IsChecked
    $fadminPattern = $Controls['txtFadminPattern'].Text.Trim()
    $customExclusions = $Controls['txtCustomExclusions'].Text
    $autoExportCSV = [bool]$Controls['chkAutoExportCSV'].IsChecked
    $autoExportUnreachable = [bool]$Controls['chkAutoExportUnreachable'].IsChecked

    # Pre-cache excluded computers list (honeypots, test systems, etc.)
    $excludedComputers = Get-ExcludedComputerNames
    if ($excludedComputers.Count -gt 0) {
        Write-Log "Excluding $($excludedComputers.Count) computer(s) from scan: $($excludedComputers -join ', ')"
    }

    # Pre-cache specific computers list if selected
    $specificComputers = @()
    if ($scanTarget -eq "Specific Computers") {
        $specificComputers = @($Controls['txtSpecificComputers'].Text -split "`r`n" |
            Where-Object { $_.Trim() -ne '' } |
            ForEach-Object { $_.Trim() })
        if ($specificComputers.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Please enter at least one computer name.", "No Computers Specified", "OK", "Warning")
            return
        }
        Write-Log "Scanning specific computers: $($specificComputers.Count) specified"
    }

    Write-Log "Starting scan - Target: $scanTarget, Throttle: $throttleLimit"
    Update-Progress -Status "Querying Active Directory..." -Percent 5

    # Run scan in background
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()

    # Use synchronized hashtable for thread-safe UI communication
    # Store at script scope so event handler can access it
    $Script:SyncHash = [hashtable]::Synchronized(@{
        Window = $Window
        Controls = $Controls
        ScanTarget = $scanTarget
        CancelRequested = $false
        CredentialError = $false
    })
    $runspace.SessionStateProxy.SetVariable('SyncHash', $Script:SyncHash)
    $runspace.SessionStateProxy.SetVariable('ScanTarget', $scanTarget)
    $runspace.SessionStateProxy.SetVariable('ThrottleLimit', $throttleLimit)
    $runspace.SessionStateProxy.SetVariable('ScanResults', $Script:ScanResults)
    $runspace.SessionStateProxy.SetVariable('UnreachableComputers', $Script:UnreachableComputers)
    $runspace.SessionStateProxy.SetVariable('AllComputers', $Script:AllComputers)
    $runspace.SessionStateProxy.SetVariable('OutputDirectory', $Script:OutputDirectory)

    # Pass pre-cached UI settings to runspace (no blocking calls needed during scan)
    $runspace.SessionStateProxy.SetVariable('ExcludeFadmin', $excludeFadmin)
    $runspace.SessionStateProxy.SetVariable('FadminPattern', $fadminPattern)
    $runspace.SessionStateProxy.SetVariable('ExcludeDomainAdmins', $excludeDomainAdmins)
    $runspace.SessionStateProxy.SetVariable('ExcludeBuiltinAdmin', $excludeBuiltinAdmin)
    $runspace.SessionStateProxy.SetVariable('CustomExclusions', $customExclusions)
    $runspace.SessionStateProxy.SetVariable('AutoExportCSV', $autoExportCSV)
    $runspace.SessionStateProxy.SetVariable('AutoExportUnreachable', $autoExportUnreachable)
    $runspace.SessionStateProxy.SetVariable('ExcludedComputers', $excludedComputers)
    $runspace.SessionStateProxy.SetVariable('SpecificComputers', $specificComputers)

    # Resume mode variables
    $runspace.SessionStateProxy.SetVariable('IsResumeMode', $ResumeMode.IsPresent)
    $runspace.SessionStateProxy.SetVariable('ResumeRemainingComputers', $Script:ResumeRemainingComputers)
    $runspace.SessionStateProxy.SetVariable('ResumeProcessedComputers', $Script:ResumeProcessedComputers)
    $runspace.SessionStateProxy.SetVariable('ResumeStartPercent', $Script:ResumeStartPercent)
    $runspace.SessionStateProxy.SetVariable('ElevatedCredential', $Script:ElevatedCredential)

    $powershell = [powershell]::Create()
    $powershell.Runspace = $runspace

    [void]$powershell.AddScript({
        # Variables are passed via SetVariable - available directly
        # $SyncHash contains Window and Controls for thread-safe UI access

        function Write-Log {
            param([string]$Message, [string]$Level = "INFO")
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message`r`n"
            try {
                $SyncHash.Window.Dispatcher.Invoke([Action]{
                    $SyncHash.Controls['txtLog'].AppendText($logEntry)
                    $SyncHash.Controls['txtLog'].ScrollToEnd()
                }, [System.Windows.Threading.DispatcherPriority]::Background)
            } catch {
                Write-Host "Write-Log Error: $_" -ForegroundColor Red
            }
        }

        function Update-Progress {
            param([string]$Status, [int]$Percent, [string]$Detail = "")
            try {
                $SyncHash.Window.Dispatcher.Invoke([Action]{
                    $SyncHash.Controls['txtProgressStatus'].Text = $Status
                    $SyncHash.Controls['txtProgressPercent'].Text = "$Percent%"
                    $SyncHash.Controls['progressBar'].Value = $Percent
                    $SyncHash.Controls['txtProgressDetail'].Text = $Detail
                }, [System.Windows.Threading.DispatcherPriority]::Background)
            } catch {
                Write-Host "Update-Progress Error: $_" -ForegroundColor Red
            }
        }

        function Get-ExclusionPatterns {
            # Uses pre-cached values passed from main thread - NO blocking Dispatcher calls!
            $patterns = @()

            if ($ExcludeFadmin -and -not [string]::IsNullOrWhiteSpace($FadminPattern)) { $patterns += $FadminPattern }
            if ($ExcludeDomainAdmins) { $patterns += '\\Domain Admins$' }
            if ($ExcludeBuiltinAdmin) { $patterns += '\\Administrator$' }

            $custom = $CustomExclusions -split "`r`n" | Where-Object { $_.Trim() -ne '' }
            $patterns += $custom

            return $patterns
        }

        try {
            # Resume Mode: Skip AD query, use remaining computers from saved state
            if ($IsResumeMode -and $ResumeRemainingComputers -and $ResumeRemainingComputers.Count -gt 0) {
                Write-Log "Resume Mode: Continuing with $($ResumeRemainingComputers.Count) remaining computers"
                $computers = $ResumeRemainingComputers
                $totalFoundInAD = $AllComputers.Count  # Use original total for progress calculation
                $SyncHash.ProcessedComputers = $ResumeProcessedComputers
                Update-Progress -Status "Resuming scan..." -Percent $ResumeStartPercent -Detail "$($ResumeProcessedComputers.Count) already scanned, $($ResumeRemainingComputers.Count) remaining"
            }
            # Handle specific computers scan differently - use optimized batch LDAP query
            elseif ($ScanTarget -eq "Specific Computers") {
                Write-Log "Querying Active Directory for $($SpecificComputers.Count) specific computer(s) using parallel batch query"

                # Build optimized LDAP filter for batch query (much faster than sequential)
                # Split into batches of 50 to avoid LDAP filter size limits
                $batchSize = 50
                $computers = @()
                $notFound = @()

                for ($i = 0; $i -lt $SpecificComputers.Count; $i += $batchSize) {
                    $batch = $SpecificComputers[$i..([Math]::Min($i + $batchSize - 1, $SpecificComputers.Count - 1))]

                    # Build LDAP OR filter: (|(name=comp1)(name=comp2)...)
                    $nameFilters = ($batch | ForEach-Object { "(name=$_)" }) -join ''
                    $ldapFilter = "(&(objectClass=computer)(|$nameFilters))"

                    try {
                        $adParams = @{ LDAPFilter = $ldapFilter; Properties = @('Name','OperatingSystem'); ErrorAction = 'Stop' }
                        if ($ElevatedCredential) { $adParams['Credential'] = $ElevatedCredential }
                        $batchResults = Get-ADComputer @adParams | Select-Object Name, OperatingSystem
                        $computers += @($batchResults)

                        # Track which computers weren't found
                        $foundNames = @($batchResults | ForEach-Object { $_.Name.ToLower() })
                        foreach ($name in $batch) {
                            if ($name.ToLower() -notin $foundNames) {
                                $notFound += $name
                            }
                        }
                    } catch {
                        Write-Log "Error querying AD batch: $_" "WARNING"
                    }

                    # Progress update for large lists
                    if ($SpecificComputers.Count -gt $batchSize) {
                        $pct = [Math]::Min(10, [int](($i / $SpecificComputers.Count) * 10))
                        Update-Progress -Status "Querying Active Directory..." -Percent $pct -Detail "Processed $([Math]::Min($i + $batchSize, $SpecificComputers.Count)) of $($SpecificComputers.Count)"
                    }
                }

                # Log computers not found
                if ($notFound.Count -gt 0) {
                    Write-Log "Computers not found in AD: $($notFound -join ', ')" "WARNING"
                }

                Write-Log "Batch AD query complete: Found $($computers.Count) of $($SpecificComputers.Count) computers"
            } else {
                # Build AD filter based on scan target
                $filter = switch ($ScanTarget) {
                    "Servers Only" {
                        { OperatingSystem -like "*Windows*Server*" -and Enabled -eq $true }
                    }
                    "Workstations Only" {
                        { OperatingSystem -like "*Windows*" -and OperatingSystem -notlike "*Server*" -and Enabled -eq $true }
                    }
                    default {
                        { OperatingSystem -like "*Windows*" -and Enabled -eq $true }
                    }
                }

                Write-Log "Querying Active Directory with filter: $ScanTarget"
                $adParams = @{ Filter = $filter; Properties = @('Name','OperatingSystem') }
                if ($ElevatedCredential) { $adParams['Credential'] = $ElevatedCredential }
                $computers = Get-ADComputer @adParams | Select-Object Name, OperatingSystem
            }

            if (-not $computers) {
                Write-Log "No computers found matching criteria" "WARNING"
                return
            }

            $totalFoundInAD = @($computers).Count
            Write-Log "Found $totalFoundInAD computers in AD"

            # Filter out excluded computers (honeypots, test systems, etc.)
            # Note: Use -in operator (case-insensitive) instead of .Contains() because
            # HashSet loses its custom StringComparer when passed through SetVariable
            $excludedCount = 0
            $filteredComputers = @()
            foreach ($c in $computers) {
                if ($ExcludedComputers -and ($c.Name -in $ExcludedComputers)) {
                    $excludedCount++
                    Write-Log "EXCLUDED: '$($c.Name)' - skipping (honeypot/test/excluded system)"
                } else {
                    $filteredComputers += $c
                }
            }

            if ($excludedCount -gt 0) {
                Write-Log "Excluded $excludedCount computer(s) from scan"
            }

            $computers = $filteredComputers
            $totalCount = @($computers).Count

            if ($totalCount -eq 0) {
                Write-Log "No computers to scan after exclusions" "WARNING"
                return
            }

            Write-Log "Scanning $totalCount computers (after exclusions)"

            # Initialize ProcessedComputers tracker if not in resume mode
            if (-not $IsResumeMode) {
                $SyncHash.ProcessedComputers = [System.Collections.ArrayList]::new()
            }

            # Calculate progress offset for resume mode
            $progressOffset = if ($IsResumeMode -and $ResumeStartPercent) { $ResumeStartPercent } else { 0 }
            $progressRange = 100 - $progressOffset

            Update-Progress -Status "Found $totalCount computers" -Percent ([int]($progressOffset + ($progressRange * 0.10))) -Detail "Preparing to scan..."

            # BUILD LOOKUP HASHTABLE - O(1) lookups instead of O(n) Where-Object in loops
            $computerLookup = @{}
            $computerNames = [System.Collections.Generic.List[string]]::new($totalCount)
            foreach ($c in $computers) {
                $computerLookup[$c.Name] = $c
                $computerNames.Add($c.Name)
            }

            Update-Progress -Status "Scanning computers..." -Percent ([int]($progressOffset + ($progressRange * 0.15))) -Detail "Connecting via WinRM (Throttle: $ThrottleLimit)"

            # Check for cancellation before starting remote scan
            if ($SyncHash.CancelRequested) {
                Write-Log "Scan cancelled by user before remote connections" "WARNING"
                return
            }

            # Execute remote command with throttling
            Write-Log "Starting remote scan with throttle limit: $ThrottleLimit"

            $icParams = @{
                ComputerName  = $computerNames
                ScriptBlock   = { Get-LocalGroupMember -Name 'Administrators' -ErrorAction SilentlyContinue }
                ThrottleLimit = $ThrottleLimit
                ErrorAction   = 'SilentlyContinue'
                ErrorVariable = 'remoteErrors'
            }
            if ($ElevatedCredential) { $icParams['Credential'] = $ElevatedCredential }
            $localAdmins = Invoke-Command @icParams

            # Check for credential-related errors in remote scan
            if ($ElevatedCredential -and $remoteErrors) {
                foreach ($err in $remoteErrors) {
                    $errMsg = $err.Exception.Message
                    if ($errMsg -match 'Access is denied' -or $errMsg -match 'The user name or password is incorrect' -or $errMsg -match 'Logon failure') {
                        $SyncHash.CredentialError = $true
                        Write-Log "CREDENTIAL ERROR: Elevated credentials may be invalid or expired - $errMsg" "ERROR"
                        break
                    }
                }
            }

            # Track all computers as processed (for resume functionality)
            foreach ($name in $computerNames) {
                [void]$SyncHash.ProcessedComputers.Add($name)
            }

            Update-Progress -Status "Processing results..." -Percent ([int]($progressOffset + ($progressRange * 0.70)))

            # USE HASHSET for O(1) membership lookups (vs O(n) with -notin)
            $reachedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            if ($localAdmins) {
                foreach ($admin in $localAdmins) {
                    [void]$reachedSet.Add($admin.PSComputerName)
                }
            }

            $unreachableCount = $totalCount - $reachedSet.Count
            Write-Log "Reached: $($reachedSet.Count), Unreachable: $unreachableCount"

            # Add unreachable computers - O(n) with O(1) HashSet lookup
            foreach ($c in $computers) {
                if (-not $reachedSet.Contains($c.Name)) {
                    $isServer = $c.OperatingSystem -like "*Server*"
                    $null = $UnreachableComputers.Add([PSCustomObject]@{
                        ComputerName = $c.Name
                        ComputerType = if ($isServer) { "Server" } else { "Workstation" }
                        OperatingSystem = $c.OperatingSystem
                        ErrorMessage = "Connection failed or timed out"
                    })
                }
            }

            Update-Progress -Status "Filtering results..." -Percent ([int]($progressOffset + ($progressRange * 0.85)))

            # Get exclusion patterns
            $patterns = Get-ExclusionPatterns
            Write-Log "Applying exclusion patterns: $($patterns -join ', ')"

            # Filter and process results with risk scoring - O(n) with O(1) hashtable lookup
            if ($localAdmins) {
                foreach ($admin in $localAdmins) {
                    $accountName = $admin.Name
                    $excluded = $false

                    foreach ($pattern in $patterns) {
                        if ($accountName -match $pattern) {
                            $excluded = $true
                            break
                        }
                    }

                    if (-not $excluded) {
                        # O(1) HASHTABLE LOOKUP instead of O(n) Where-Object
                        $comp = $computerLookup[$admin.PSComputerName]
                        $isServer = $comp.OperatingSystem -like "*Server*"
                        $computerType = if ($isServer) { "Server" } else { "Workstation" }

                    # Calculate risk score
                    $riskScore = 0
                    $objectClass = $admin.ObjectClass

                    # Account type scoring
                    $isLocalAccount = $accountName -notmatch '\\'
                    $isDomainAccount = $accountName -match '\\'

                    if ($objectClass -eq 'User') {
                        if ($isLocalAccount) {
                            $riskScore += 40  # Local user - highest risk
                        } elseif ($isDomainAccount) {
                            $accountPart = ($accountName -split '\\')[-1]
                            $isServiceAccount = $accountPart -match '^(svc|service|sql|app|sys|task|batch|iis|mssql|backup|scan)' -or
                                               $accountPart -match '(svc|service)$'
                            if ($isServiceAccount) {
                                $riskScore += 20  # Service account
                            } else {
                                $riskScore += 30  # Domain user
                            }
                        }
                    } elseif ($objectClass -eq 'Group') {
                        $riskScore += 15  # Group
                    }

                    # Server bonus
                    if ($computerType -eq 'Server') {
                        $riskScore += 10
                    }

                    # Determine risk level and colors
                    $riskLevel = switch ($riskScore) {
                        { $_ -ge 60 } { 'HIGH' }
                        { $_ -ge 30 } { 'MEDIUM' }
                        { $_ -ge 1 }  { 'LOW' }
                        default       { 'INFO' }
                    }
                    $riskColor = switch ($riskLevel) {
                        'HIGH'   { '#FF6B6B' }
                        'MEDIUM' { '#FFD93D' }
                        'LOW'    { '#4ECDC4' }
                        default  { '#A0AEC0' }
                    }
                    $riskBgColor = switch ($riskLevel) {
                        'HIGH'   { '#3D1212' }
                        'MEDIUM' { '#3D2E05' }
                        'LOW'    { '#0D2F2D' }
                        default  { '#2D3748' }
                    }

                    $null = $ScanResults.Add([PSCustomObject]@{
                        ComputerName = $admin.PSComputerName
                        ComputerType = $computerType
                        OperatingSystem = $comp.OperatingSystem
                        AdminName = $accountName
                        ObjectClass = $objectClass
                        RiskScore = $riskScore
                        RiskLevel = $riskLevel
                        RiskColor = $riskColor
                        RiskBgColor = $riskBgColor
                    })
                    }
                }
            }

            Write-Log "Found $($ScanResults.Count) unexpected local administrators"

            Update-Progress -Status "Scan complete!" -Percent 100 -Detail "Found $($ScanResults.Count) unexpected admins"

            # Auto-export using pre-cached settings with date-based folder organization
            $dateFolder = Join-Path $OutputDirectory (Get-Date -Format 'yyyy-MM-dd')
            $timeStamp = Get-Date -Format 'HHmmss'

            # Create date folder if needed (only if we have something to export)
            if (($AutoExportCSV -and $ScanResults.Count -gt 0) -or ($AutoExportUnreachable -and $UnreachableComputers.Count -gt 0)) {
                if (-not (Test-Path $dateFolder)) {
                    New-Item -Path $dateFolder -ItemType Directory -Force | Out-Null
                    Write-Log "Created output folder: $dateFolder"
                }
            }

            if ($AutoExportCSV -and $ScanResults.Count -gt 0) {
                $csvPath = Join-Path $dateFolder "LocalAdminAudit_$timeStamp.csv"
                $ScanResults | Export-Csv -Path $csvPath -NoTypeInformation
                Write-Log "Auto-exported results to: $csvPath"
            }

            if ($AutoExportUnreachable -and $UnreachableComputers.Count -gt 0) {
                $logPath = Join-Path $dateFolder "Unreachable_$timeStamp.log"
                $UnreachableComputers | Select-Object -ExpandProperty ComputerName | Out-File $logPath
                Write-Log "Auto-exported unreachable list to: $logPath"
            }

        } catch {
            Write-Log "Error during scan: $_" "ERROR"
            Update-Progress -Status "Scan failed!" -Percent 0 -Detail $_.Exception.Message
        } finally {
            # OPTIMIZED: Single-pass statistics (avoids multiple pipeline iterations)
            $totalCount = if ($computers) { @($computers).Count } else { 0 }
            $unreachableCount = $UnreachableComputers.Count
            $unexpectedCount = $ScanResults.Count

            # Use HashSet for unique computer count (O(n) instead of pipeline)
            $uniqueComputers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $serverCount = 0
            $highRiskCount = 0
            $mediumRiskCount = 0
            $lowRiskCount = 0

            # SINGLE PASS over results for all stats
            foreach ($r in $ScanResults) {
                [void]$uniqueComputers.Add($r.ComputerName)
                switch ($r.RiskLevel) {
                    'HIGH'   { $highRiskCount++ }
                    'MEDIUM' { $mediumRiskCount++ }
                    'LOW'    { $lowRiskCount++ }
                }
            }

            # Count servers from computers list (if available)
            if ($computers) {
                foreach ($c in $computers) {
                    if ($c.OperatingSystem -like "*Server*") { $serverCount++ }
                }
            }

            $reachedCount = $uniqueComputers.Count
            $workstationCount = $totalCount - $serverCount
            $reachedPct = if ($totalCount -gt 0) { [math]::Round(($reachedCount / $totalCount) * 100, 1) } else { 0 }
            $unexpectedDetail = if ($unexpectedCount -gt 0) { "Requires review" } else { "All clear" }

            # Pre-build all strings and store in SyncHash for UI thread access
            # PowerShell Dispatcher.BeginInvoke doesn't capture outer variables like C# closures
            $SyncHash.strTotal = $totalCount.ToString()
            $SyncHash.strBreakdown = "Servers: $serverCount | Workstations: $workstationCount"
            $SyncHash.strReached = $reachedCount.ToString()
            $SyncHash.strReachedPct = "$reachedPct% success rate"
            $SyncHash.strUnreachable = $unreachableCount.ToString()
            $SyncHash.strUnexpected = $unexpectedCount.ToString()
            $SyncHash.strUnexpectedDetail = $unexpectedDetail
            $SyncHash.strResultsCount = "$unexpectedCount results"
            $SyncHash.strUnreachableCount = "$unreachableCount unreachable"
            $SyncHash.strHighRisk = "$highRiskCount HIGH"
            $SyncHash.strMediumRisk = "$mediumRiskCount MED"
            $SyncHash.strLowRisk = "$lowRiskCount LOW"

            # Convert results to PSCustomObject arrays for WPF DataGrid binding
            $resultsArray = @($ScanResults | Sort-Object -Property RiskScore -Descending | ForEach-Object {
                [PSCustomObject]@{
                    ComputerName = [string]$_.ComputerName
                    ComputerType = [string]$_.ComputerType
                    OperatingSystem = [string]$_.OperatingSystem
                    AdminName = [string]$_.AdminName
                    ObjectClass = [string]$_.ObjectClass
                    RiskScore = [int]$_.RiskScore
                    RiskLevel = [string]$_.RiskLevel
                    RiskColor = [string]$_.RiskColor
                    RiskBgColor = [string]$_.RiskBgColor
                }
            })

            $unreachableArray = @($UnreachableComputers | ForEach-Object {
                [PSCustomObject]@{
                    ComputerName = [string]$_.ComputerName
                    ComputerType = [string]$_.ComputerType
                    OperatingSystem = [string]$_.OperatingSystem
                    ErrorMessage = [string]$_.ErrorMessage
                }
            })

            Write-Log "Updating UI with results..."

            # Prepare statistics for history save
            $statsForHistory = @{
                total = $totalCount
                reached = $reachedCount
                unreachable = $unreachableCount
                findings = $unexpectedCount
                highRisk = $highRiskCount
                mediumRisk = $mediumRiskCount
                lowRisk = $lowRiskCount
            }
            $scanTargetForHistory = $ScanTarget

            # Store arrays in SyncHash so they're accessible in the UI thread action
            # PowerShell doesn't capture outer variables in Dispatcher actions like C# closures
            $SyncHash.ResultsArray = $resultsArray
            $SyncHash.UnreachableArray = $unreachableArray
            $SyncHash.StatsForHistory = $statsForHistory
            $SyncHash.ScanResultsRaw = $ScanResults
            $SyncHash.UnreachableRaw = $UnreachableComputers

            # PHASE 1: UI update - use Invoke (synchronous) like Update-Progress
            try {
                $SyncHash.Window.Dispatcher.Invoke([Action]{
                    # Re-enable buttons immediately so user knows scan is done
                    $SyncHash.Controls['btnStartScan'].IsEnabled = $true
                    $SyncHash.Controls['btnRescan'].IsEnabled = $true
                    $SyncHash.Controls['btnCancelScan'].IsEnabled = $false
                    $SyncHash.Controls['btnElevate'].IsEnabled = $true
                    $SyncHash.Controls['progressCard'].Visibility = 'Collapsed'

                    # Update status bar immediately
                    $SyncHash.Controls['txtStatusBar'].Text = "Scan complete - Processing results..."
                    $SyncHash.Controls['txtLastScanTime'].Text = "Last scan: $(Get-Date -Format 'MMM d, h:mm tt')"

                    # Update statistics using pre-built strings from SyncHash
                    $SyncHash.Controls['txtTotalComputers'].Text = $SyncHash.strTotal
                    $SyncHash.Controls['txtComputerBreakdown'].Text = $SyncHash.strBreakdown
                    $SyncHash.Controls['txtReached'].Text = $SyncHash.strReached
                    $SyncHash.Controls['txtReachedPercent'].Text = $SyncHash.strReachedPct
                    $SyncHash.Controls['txtUnreachable'].Text = $SyncHash.strUnreachable
                    $SyncHash.Controls['txtUnexpectedAdmins'].Text = $SyncHash.strUnexpected
                    $SyncHash.Controls['txtResultsCount'].Text = $SyncHash.strResultsCount
                    $SyncHash.Controls['txtUnreachableCount'].Text = $SyncHash.strUnreachableCount

                    # Update risk distribution
                    $SyncHash.Controls['txtHighRisk'].Text = $SyncHash.strHighRisk
                    $SyncHash.Controls['txtMediumRisk'].Text = $SyncHash.strMediumRisk
                    $SyncHash.Controls['txtLowRisk'].Text = $SyncHash.strLowRisk

                    # Update unreachable grid immediately (no processing needed)
                    $SyncHash.Controls['dgUnreachable'].ItemsSource = $SyncHash.UnreachableArray

                    # Bind results grid
                    $SyncHash.Controls['dgResults'].ItemsSource = $null
                    $SyncHash.Controls['dgResults'].ItemsSource = $SyncHash.ResultsArray

                    # Update final status
                    $SyncHash.Controls['txtStatusBar'].Text = "Ready - Found $($SyncHash.ResultsArray.Count) unexpected admin(s)"
                }, [System.Windows.Threading.DispatcherPriority]::Background)
            } catch {
                Write-Host "Phase 1 UI Error: $_" -ForegroundColor Red
            }

            Write-Log "Scan completed - results bound to UI"
        }
    })

    # Store for cleanup (before starting)
    $Script:CurrentRunspace = $runspace
    $Script:CurrentPowerShell = $powershell

    # Clean up any existing event subscription from previous runs
    Get-EventSubscriber -SourceIdentifier "ScanComplete" -ErrorAction SilentlyContinue | Unregister-Event -ErrorAction SilentlyContinue

    # Register completion handler BEFORE starting scan - critical to catch fast completions
    $null = Register-ObjectEvent -InputObject $powershell -EventName InvocationStateChanged -SourceIdentifier "ScanComplete" -Action {
        if ($Sender.InvocationStateInfo.State -eq 'Completed') {
            # Use Dispatcher to run on UI thread
            $Window.Dispatcher.BeginInvoke([Action]{
                try {
                    # Save scan to history (functions accessible in main script context)
                    if ($Script:SyncHash.ResultsArray.Count -gt 0 -or $Script:SyncHash.UnreachableArray.Count -gt 0) {
                        try {
                            Save-ScanHistory -ScanTarget $Script:SyncHash.ScanTarget `
                                -Results $Script:SyncHash.ScanResultsRaw `
                                -Unreachable $Script:SyncHash.UnreachableRaw `
                                -Statistics $Script:SyncHash.StatsForHistory
                            Refresh-HistoryGrid
                        } catch {
                            Write-Log "History save error: $_" "WARNING"
                        }
                    }

                    # Warn if credential errors were detected during scan
                    if ($Script:SyncHash.CredentialError) {
                        [System.Windows.MessageBox]::Show(
                            "Credential errors were detected during the scan. The elevated credentials may be invalid or expired.`n`nClick 'Elevate' to update credentials and try again.",
                            "Credential Warning",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                    }

                    # Clear any saved scan state since scan completed successfully
                    Clear-ScanState

                    # Hide resume banner if visible
                    $Controls['resumeScanBanner'].Visibility = 'Collapsed'

                    # Play sound notification
                    if ($Controls['chkPlaySound'].IsChecked) {
                        [System.Threading.ThreadPool]::QueueUserWorkItem({
                            try { [System.Media.SystemSounds]::Asterisk.Play() } catch { }
                        }) | Out-Null
                    }

                    Write-Log "Completion handler executed"
                } catch {
                    Write-Log "Error in completion handler: $_" "ERROR"
                    $Controls['txtStatusBar'].Text = "Error processing results: $_"
                }

                # Cleanup event registration
                Unregister-Event -SourceIdentifier "ScanComplete" -ErrorAction SilentlyContinue
                Remove-Job -Name "ScanComplete" -ErrorAction SilentlyContinue
            })
        }
    }

    # NOW start the scan (after event handler is registered)
    $Script:AsyncResult = $powershell.BeginInvoke()
}
#endregion

#region Event Handlers

# Scan Target Selection Change - show/hide specific computers panel
$Controls['cmbScanTarget'].Add_SelectionChanged({
    if ($Controls['cmbScanTarget'].SelectedItem.Content -eq "Specific Computers") {
        $Controls['specificComputersCard'].Visibility = 'Visible'
    } else {
        $Controls['specificComputersCard'].Visibility = 'Collapsed'
    }
})

# Specific Computers Text Change - update count
$Controls['txtSpecificComputers'].Add_TextChanged({
    $text = $Controls['txtSpecificComputers'].Text
    if ([string]::IsNullOrWhiteSpace($text)) {
        $Controls['txtSpecificComputersCount'].Text = "0 computers entered"
    } else {
        $computers = @($text -split "`r`n" | Where-Object { $_.Trim() -ne '' })
        $count = $computers.Count
        $Controls['txtSpecificComputersCount'].Text = "$count computer$(if($count -ne 1){'s'}) entered"
    }
})

# Import Computers from File Button
$Controls['btnImportComputers'].Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Title = "Import Computer Names"
    $dialog.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dialog.DefaultExt = ".txt"

    if ($dialog.ShowDialog()) {
        try {
            $content = Get-Content $dialog.FileName -Raw
            # Handle CSV - extract first column if comma-separated
            if ($dialog.FileName -match '\.csv$') {
                $lines = $content -split "`r`n" | Where-Object { $_.Trim() -ne '' }
                $content = ($lines | ForEach-Object { ($_ -split ',')[0].Trim().Trim('"') }) -join "`r`n"
            }
            $Controls['txtSpecificComputers'].Text = $content
            Write-Log "Imported computers from: $($dialog.FileName)"
        } catch {
            Write-Log "Error importing file: $_" "ERROR"
            [System.Windows.MessageBox]::Show("Error importing file:`n$_", "Import Error", "OK", "Error")
        }
    }
})

# Clear Computers Button
$Controls['btnClearComputers'].Add_Click({
    $Controls['txtSpecificComputers'].Text = ""
})

# Start Scan Button
$Controls['btnStartScan'].Add_Click({
    Start-AuditScan
})

# Rescan Button (Quick Rescan with last settings)
$Controls['btnRescan'].Add_Click({
    if ($null -eq $Script:LastScanSettings) {
        [System.Windows.MessageBox]::Show(
            "No previous scan to repeat. Run a scan first.",
            "No Previous Scan",
            "OK",
            "Warning"
        )
        return
    }

    # Restore last scan settings to UI
    $settings = $Script:LastScanSettings

    # Find and select the scan target
    foreach ($item in $Controls['cmbScanTarget'].Items) {
        if ($item.Content -eq $settings.ScanTarget) {
            $Controls['cmbScanTarget'].SelectedItem = $item
            break
        }
    }

    # Find and select the throttle limit
    foreach ($item in $Controls['cmbThrottle'].Items) {
        if ($item.Content -eq $settings.ThrottleLimit.ToString()) {
            $Controls['cmbThrottle'].SelectedItem = $item
            break
        }
    }

    # Restore specific computers if applicable
    if ($settings.ScanTarget -eq "Specific Computers") {
        $Controls['txtSpecificComputers'].Text = $settings.SpecificComputers
    }

    Write-Log "Quick Rescan: Repeating last scan ($($settings.ScanTarget), Throttle: $($settings.ThrottleLimit))"

    # Start the scan
    Start-AuditScan
})

# Resume Scan Button (Resume interrupted scan)
$Controls['btnResumeScan'].Add_Click({
    $state = Load-ScanState
    if ($null -eq $state) {
        [System.Windows.MessageBox]::Show(
            "No saved scan state found.",
            "No Saved State",
            "OK",
            "Warning"
        )
        $Controls['resumeScanBanner'].Visibility = 'Collapsed'
        return
    }

    Write-Log "Resuming interrupted scan..."

    # Restore previous results and unreachable computers
    $Script:ScanResults = [System.Collections.ArrayList]::new()
    foreach ($result in $state.Results) {
        [void]$Script:ScanResults.Add([PSCustomObject]@{
            ComputerName = $result.ComputerName
            ComputerType = $result.ComputerType
            OperatingSystem = $result.OperatingSystem
            AccountName = $result.AccountName
            AccountType = $result.AccountType
            RiskLevel = $result.RiskLevel
            RiskScore = $result.RiskScore
            Status = if ($result.Status) { $result.Status } else { "EXISTING" }
        })
    }

    $Script:UnreachableComputers = [System.Collections.ArrayList]::new()
    foreach ($unreachable in $state.Unreachable) {
        [void]$Script:UnreachableComputers.Add([PSCustomObject]@{
            ComputerName = $unreachable.ComputerName
            OperatingSystem = $unreachable.OperatingSystem
            ErrorMessage = $unreachable.ErrorMessage
        })
    }

    # Restore AllComputers list (computers to scan)
    $Script:AllComputers = @()
    foreach ($comp in $state.AllComputers) {
        $Script:AllComputers += [PSCustomObject]@{
            Name = $comp.Name
            OperatingSystem = $comp.OperatingSystem
            DistinguishedName = $comp.DistinguishedName
        }
    }

    # Filter to only remaining computers (not yet processed)
    $processedNames = @($state.ProcessedComputers)
    $remainingComputers = @($Script:AllComputers | Where-Object { $_.Name -notin $processedNames })

    Write-Log "Restored $($state.CompletedCount) previously scanned computers, $($remainingComputers.Count) remaining"

    # Hide the resume banner
    $Controls['resumeScanBanner'].Visibility = 'Collapsed'

    # Restore settings to UI
    foreach ($item in $Controls['cmbScanTarget'].Items) {
        if ($item.Content -eq $state.ScanSettings.ScanTarget) {
            $Controls['cmbScanTarget'].SelectedItem = $item
            break
        }
    }

    foreach ($item in $Controls['cmbThrottle'].Items) {
        if ($item.Content -eq $state.ScanSettings.ThrottleLimit.ToString()) {
            $Controls['cmbThrottle'].SelectedItem = $item
            break
        }
    }

    if ($state.ScanSettings.SpecificComputers) {
        $Controls['txtSpecificComputers'].Text = $state.ScanSettings.SpecificComputers
    }

    # Store these for the scan function
    $Script:ResumeMode = $true
    $Script:ResumeRemainingComputers = $remainingComputers
    $Script:ResumeProcessedComputers = [System.Collections.ArrayList]::new($processedNames)
    $Script:ResumeStartPercent = if ($state.TotalComputers -gt 0) { [math]::Round(($state.CompletedCount / $state.TotalComputers) * 100, 0) } else { 0 }

    # Update UI to show resumed state
    Update-Statistics
    $Controls['dgResults'].ItemsSource = $null
    $Controls['dgResults'].ItemsSource = @($Script:ScanResults)
    $Controls['dgUnreachable'].ItemsSource = $null
    $Controls['dgUnreachable'].ItemsSource = @($Script:UnreachableComputers)

    # Start the resumed scan
    Start-AuditScan -ResumeMode
})

# Discard Resume Button
$Controls['btnDiscardResume'].Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Are you sure you want to discard the saved scan state?`n`nThis cannot be undone.",
        "Discard Saved Scan",
        "YesNo",
        "Question"
    )

    if ($result -eq 'Yes') {
        Clear-ScanState
        $Controls['resumeScanBanner'].Visibility = 'Collapsed'
        Write-Log "Discarded saved scan state"
    }
})

# Cancel Scan Button
$Controls['btnCancelScan'].Add_Click({
    $Script:CancelRequested = $true
    if ($Script:SyncHash) { $Script:SyncHash.CancelRequested = $true }
    if ($Script:CurrentPowerShell) {
        $Script:CurrentPowerShell.Stop()
    }
    Write-Log "Scan cancelled by user" "WARNING"

    # Save scan state for potential resume
    if ($Script:AllComputers.Count -gt 0 -and $SyncHash.ProcessedComputers -and $SyncHash.ProcessedComputers.Count -gt 0) {
        $saveResult = Save-ScanState -AllComputers $Script:AllComputers `
            -ProcessedComputers @($SyncHash.ProcessedComputers) `
            -Results @($Script:ScanResults) `
            -Unreachable @($Script:UnreachableComputers) `
            -Settings @{
                ScanTarget = $Script:LastScanSettings.ScanTarget
                ThrottleLimit = $Script:LastScanSettings.ThrottleLimit
                SpecificComputers = $Script:LastScanSettings.SpecificComputers
                ExclusionPatterns = (Get-ExclusionPatterns)
            }

        if ($saveResult) {
            Write-Log "Scan state saved - can resume later" "INFO"
            Update-Progress -Status "Scan cancelled - state saved for resume" -Percent 0
        } else {
            Update-Progress -Status "Scan cancelled" -Percent 0
        }
    } else {
        Update-Progress -Status "Scan cancelled" -Percent 0
    }

    $Controls['btnStartScan'].IsEnabled = $true
    $Controls['btnRescan'].IsEnabled = ($null -ne $Script:LastScanSettings)
    $Controls['btnCancelScan'].IsEnabled = $false
    $Controls['btnElevate'].IsEnabled = $true
})

# Elevate Button - Alternate credential management
$Controls['btnElevate'].Add_Click({
    if ($Script:ElevatedCredential) {
        # Already elevated - offer to clear or change
        $result = [System.Windows.MessageBox]::Show(
            "Currently elevated as: $($Script:ElevatedCredential.UserName)`n`nYes = Clear credentials`nNo = Change credentials",
            "Credential Elevation",
            [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($result -eq 'Yes') {
            $Script:ElevatedCredential = $null
            $Script:Controls['btnElevate'].Content = "Authenticate"
            $Script:Controls['btnElevate'].Style = $Window.FindResource('ModernButton')
            $Script:Controls['txtCredentialStatus'].Text = ""
            $Script:Controls['txtStatusBar'].Text = "Credentials cleared - using current session identity"
            Write-Log "Credential elevation cleared"
        }
        elseif ($result -eq 'No') {
            $cred = Get-Credential -Message "Enter domain admin credentials for scan operations"
            if ($cred) {
                try {
                    $Script:Controls['txtStatusBar'].Text = "Validating credentials..."
                    $null = Get-ADDomain -Credential $cred -ErrorAction Stop
                    $Script:ElevatedCredential = $cred
                    $Script:Controls['btnElevate'].Content = $cred.UserName
                    $Script:Controls['btnElevate'].Style = $Window.FindResource('ModernButton')
                    $Script:Controls['txtCredentialStatus'].Text = ""
                    $Script:Controls['txtStatusBar'].Text = "Credentials updated: $($cred.UserName)"
                    Write-Log "Credential elevation changed to: $($cred.UserName)"
                }
                catch {
                    [System.Windows.MessageBox]::Show(
                        "Credential validation failed:`n$($_.Exception.Message)",
                        "Invalid Credentials",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error
                    )
                    Write-Log "Credential validation failed: $($_.Exception.Message)" "ERROR"
                }
            }
        }
    }
    else {
        # Not elevated - prompt for credentials
        $cred = Get-Credential -Message "Enter domain admin credentials for scan operations"
        if ($cred) {
            try {
                $Script:Controls['txtStatusBar'].Text = "Validating credentials..."
                $null = Get-ADDomain -Credential $cred -ErrorAction Stop
                $Script:ElevatedCredential = $cred
                $Script:Controls['btnElevate'].Content = $cred.UserName
                $Script:Controls['btnElevate'].Style = $Window.FindResource('ModernButton')
                $Script:Controls['txtCredentialStatus'].Text = ""
                $Script:Controls['txtStatusBar'].Text = "Elevated as: $($cred.UserName)"
                Write-Log "Credential elevation set: $($cred.UserName)"
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "Credential validation failed:`n$($_.Exception.Message)",
                    "Invalid Credentials",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
                $Script:Controls['txtStatusBar'].Text = "Credential validation failed"
                Write-Log "Credential validation failed: $($_.Exception.Message)" "ERROR"
            }
        }
    }
})

# Help Button
$Controls['btnHelp'].Add_Click({
    $helpText = @"
Local Administrator Auditor v2.6

This tool scans Windows computers in your Active Directory domain to identify unexpected local administrator accounts.

FEATURES:
- Scan servers, workstations, or both
- Risk severity scoring (HIGH/MEDIUM/LOW)
- NEW/EXISTING badges for findings
- Executive Summary reports
- History tab to view/load past scans
- Configurable exclusion patterns
- Real-time progress tracking
- Export to CSV, HTML, or clipboard
- Settings persistence (auto-saved)
- Scan history tracking (up to 50 scans)
- Context menus on data grids
- Sound notifications
- Group-By views (Computer, Admin, Risk, etc.)
- Approved service accounts (reduced risk score)

KEYBOARD SHORTCUTS:
F5         - Start Scan
Escape     - Cancel Scan
Ctrl+S     - Export CSV
Ctrl+F     - Focus Filter
Ctrl+R     - Retry Selected
Ctrl+1-6   - Switch Tabs

RISK SCORING:
- Local user: +40 pts
- Domain user: +30 pts
- Service account: +20 pts
- Group: +15 pts
- Server: +10 pts
- Approved account: -50 pts

DATA STORAGE:
Settings and history saved to:
$env:APPDATA\LocalAdminAuditor\

REQUIREMENTS:
- ActiveDirectory PowerShell module
- PowerShell remoting enabled on target computers
- Appropriate AD read permissions
- Remote admin access to targets

USAGE:
1. Select scan target (Servers/Workstations/Both)
2. Adjust throttle limit if needed
3. Click 'Start Scan' (or press F5)
4. Review results in the Results tab
5. Use Group-By to organize findings
6. Right-click for context menu options
7. Export findings as needed

For support, contact your IT administrator.
"@
    [System.Windows.MessageBox]::Show($helpText, "Help - Local Administrator Auditor", "OK", "Information")
})

# Results Filter
$Controls['txtResultsFilter'].Add_TextChanged({
    Apply-ResultsFilter
})

$Controls['cmbResultsTypeFilter'].Add_SelectionChanged({
    Apply-ResultsFilter
})

$Controls['cmbGroupBy'].Add_SelectionChanged({
    Apply-ResultsFilter
})

$Controls['btnClearFilter'].Add_Click({
    $Controls['txtResultsFilter'].Text = ""
    $Controls['cmbResultsTypeFilter'].SelectedIndex = 0
    $Controls['cmbGroupBy'].SelectedIndex = 0
    Apply-ResultsFilter
})

# Compare with Previous Scan Button
$Controls['btnCompare'].Add_Click({
    if ($Script:ScanResults.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No current scan results available. Please run a scan first.",
            "No Data",
            "OK",
            "Warning"
        )
        return
    }

    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Title = "Select Previous Scan CSV to Compare"
    $dialog.Filter = "CSV Files (*.csv)|*.csv"
    $dialog.InitialDirectory = $Script:OutputDirectory

    if ($dialog.ShowDialog()) {
        try {
            Write-Log "Loading previous scan for comparison: $($dialog.FileName)"

            # Load previous CSV
            $previousResults = Import-Csv $dialog.FileName

            # Build lookup sets for fast comparison (ComputerName + AdminName as key)
            $currentKeys = @{}
            foreach ($result in $Script:ScanResults) {
                $key = "$($result.ComputerName)|$($result.AdminName)"
                $currentKeys[$key] = $result
            }

            $previousKeys = @{}
            # Handle different possible column names in CSV
            foreach ($prev in $previousResults) {
                $compName = if ($prev.ComputerName) { $prev.ComputerName } elseif ($prev.'Computer Name') { $prev.'Computer Name' } else { $null }
                $adminName = if ($prev.AdminName) { $prev.AdminName } elseif ($prev.'Unexpected Admin Account') { $prev.'Unexpected Admin Account' } else { $null }
                if ($compName -and $adminName) {
                    $key = "$compName|$adminName"
                    $previousKeys[$key] = $prev
                }
            }

            # Build comparison results
            $comparisonResults = @()

            # Mark current results as NEW or unchanged
            foreach ($result in $Script:ScanResults) {
                $key = "$($result.ComputerName)|$($result.AdminName)"
                $newResult = $result.PSObject.Copy()
                if (-not $previousKeys.ContainsKey($key)) {
                    $newResult | Add-Member -NotePropertyName 'ChangeStatus' -NotePropertyValue 'NEW' -Force
                } else {
                    $newResult | Add-Member -NotePropertyName 'ChangeStatus' -NotePropertyValue '' -Force
                }
                $comparisonResults += $newResult
            }

            # Add REMOVED items (in previous but not in current)
            foreach ($key in $previousKeys.Keys) {
                if (-not $currentKeys.ContainsKey($key)) {
                    $prev = $previousKeys[$key]
                    # Create a result object for removed item
                    $compName = if ($prev.ComputerName) { $prev.ComputerName } elseif ($prev.'Computer Name') { $prev.'Computer Name' } else { 'Unknown' }
                    $adminName = if ($prev.AdminName) { $prev.AdminName } elseif ($prev.'Unexpected Admin Account') { $prev.'Unexpected Admin Account' } else { 'Unknown' }
                    $os = if ($prev.OperatingSystem) { $prev.OperatingSystem } elseif ($prev.'Operating System') { $prev.'Operating System' } else { '' }
                    $type = if ($prev.ComputerType) { $prev.ComputerType } elseif ($prev.Type) { $prev.Type } else { '' }
                    $objClass = if ($prev.ObjectClass) { $prev.ObjectClass } elseif ($prev.'Account Type') { $prev.'Account Type' } else { '' }

                    $removedItem = [PSCustomObject]@{
                        ComputerName = $compName
                        AdminName = $adminName
                        OperatingSystem = $os
                        ComputerType = $type
                        ObjectClass = $objClass
                        RiskLevel = 'REMOVED'
                        RiskScore = 0
                        ChangeStatus = 'REMOVED'
                    }
                    $comparisonResults += $removedItem
                }
            }

            # Count changes
            $newCount = @($comparisonResults | Where-Object { $_.ChangeStatus -eq 'NEW' }).Count
            $removedCount = @($comparisonResults | Where-Object { $_.ChangeStatus -eq 'REMOVED' }).Count

            # Update grid with comparison results
            $Controls['dgResults'].ItemsSource = $null
            $Controls['dgResults'].ItemsSource = @($comparisonResults)

            # Show Clear Diff button
            $Controls['btnClearComparison'].Visibility = 'Visible'

            # Update results count text
            $Controls['txtResultsCount'].Text = "$($comparisonResults.Count) results ($newCount NEW, $removedCount REMOVED)"

            Write-Log "Comparison complete: $newCount NEW, $removedCount REMOVED admin accounts"

            [System.Windows.MessageBox]::Show(
                "Comparison complete:`n`n$newCount NEW admin account(s) found`n$removedCount admin account(s) REMOVED`n`nResults updated in grid. Use 'Clear Diff' to return to current results only.",
                "Comparison Results",
                "OK",
                "Information"
            )
        } catch {
            Write-Log "Error during comparison: $_" "ERROR"
            [System.Windows.MessageBox]::Show(
                "Error comparing with previous scan:`n$_",
                "Comparison Error",
                "OK",
                "Error"
            )
        }
    }
})

# Clear Comparison Button
$Controls['btnClearComparison'].Add_Click({
    # Restore original results without ChangeStatus
    $Controls['dgResults'].ItemsSource = $null
    $Controls['dgResults'].ItemsSource = @($Script:ScanResults)
    $Controls['btnClearComparison'].Visibility = 'Collapsed'
    $Controls['txtResultsCount'].Text = "$($Script:ScanResults.Count) results"
    Write-Log "Comparison cleared, showing current results only"
})

# Executive Summary Button
$Controls['btnExecSummary'].Add_Click({
    if ($Script:ScanResults.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No scan results available. Please run a scan first.",
            "No Data",
            "OK",
            "Warning"
        )
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "HTML Files (*.html)|*.html"
    $dialog.FileName = "ExecutiveSummary_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $dialog.InitialDirectory = $Script:OutputDirectory
    $dialog.Title = "Save Executive Summary Report"

    if ($dialog.ShowDialog() -eq "OK") {
        Export-ExecutiveSummary -FilePath $dialog.FileName
        [System.Windows.MessageBox]::Show(
            "Executive Summary exported to:`n$($dialog.FileName)",
            "Export Complete",
            "OK",
            "Information"
        )

        # Offer to open the file
        $result = [System.Windows.MessageBox]::Show(
            "Would you like to open the Executive Summary?",
            "Open Report",
            "YesNo",
            "Question"
        )
        if ($result -eq "Yes") {
            Start-Process $dialog.FileName
        }
    }
})

# Export Buttons
$Controls['btnExportCSV'].Add_Click({
    # Use date-based folder for organization
    $dateFolder = Get-OutputPath -FileName ""
    $dateFolder = Split-Path $dateFolder -Parent  # Get just the folder path

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "CSV Files (*.csv)|*.csv"
    $dialog.FileName = "LocalAdminAudit_$(Get-Date -Format 'HHmmss').csv"
    $dialog.InitialDirectory = $dateFolder

    if ($dialog.ShowDialog() -eq "OK") {
        Export-ResultsToCSV -FilePath $dialog.FileName
        [System.Windows.MessageBox]::Show("Results exported to:`n$($dialog.FileName)", "Export Complete", "OK", "Information")
    }
})

$Controls['btnExportHTML'].Add_Click({
    # Use date-based folder for organization
    $dateFolder = Get-OutputPath -FileName ""
    $dateFolder = Split-Path $dateFolder -Parent  # Get just the folder path

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "HTML Files (*.html)|*.html"
    $dialog.FileName = "LocalAdminAudit_$(Get-Date -Format 'HHmmss').html"
    $dialog.InitialDirectory = $dateFolder

    if ($dialog.ShowDialog() -eq "OK") {
        Export-ResultsToHTML -FilePath $dialog.FileName
        [System.Windows.MessageBox]::Show("Report exported to:`n$($dialog.FileName)", "Export Complete", "OK", "Information")

        # Offer to open the file
        $result = [System.Windows.MessageBox]::Show("Would you like to open the report?", "Open Report", "YesNo", "Question")
        if ($result -eq "Yes") {
            Start-Process $dialog.FileName
        }
    }
})

# PDF Export
$Controls['btnExportPDF'].Add_Click({
    if ($Script:ScanResults.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No results to export.", "No Data", "OK", "Warning")
        return
    }

    # Use date-based folder for organization
    $dateFolder = Get-OutputPath -FileName ""
    $dateFolder = Split-Path $dateFolder -Parent

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "PDF Files (*.pdf)|*.pdf"
    $dialog.FileName = "LocalAdminAudit_$(Get-Date -Format 'HHmmss').pdf"
    $dialog.InitialDirectory = $dateFolder
    $dialog.Title = "Export PDF Report"

    if ($dialog.ShowDialog() -eq "OK") {
        $pdfPath = $dialog.FileName
        $tempHtmlPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.html'

        $Controls['txtStatusBar'].Text = "Generating PDF report..."
        Write-Log "Exporting to PDF: $pdfPath"

        try {
            # First generate HTML with print-optimized styles
            Export-ResultsToHTML -FilePath $tempHtmlPath

            # Try to find Edge or Chrome for headless PDF conversion
            $edgePath = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
            $edgePath2 = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
            $chromePath = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
            $chromePath2 = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"

            $browserPath = $null
            if (Test-Path $edgePath) { $browserPath = $edgePath }
            elseif (Test-Path $edgePath2) { $browserPath = $edgePath2 }
            elseif (Test-Path $chromePath) { $browserPath = $chromePath }
            elseif (Test-Path $chromePath2) { $browserPath = $chromePath2 }

            if ($browserPath) {
                # Use headless browser to convert HTML to PDF
                $htmlUri = "file:///$($tempHtmlPath -replace '\\', '/')"
                $args = @(
                    "--headless"
                    "--disable-gpu"
                    "--print-to-pdf=`"$pdfPath`""
                    "--no-pdf-header-footer"
                    $htmlUri
                )

                $process = Start-Process -FilePath $browserPath -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
                Start-Sleep -Milliseconds 500  # Give it time to write the file

                if (Test-Path $pdfPath) {
                    Write-Log "PDF exported successfully: $pdfPath"
                    [System.Windows.MessageBox]::Show(
                        "PDF report exported to:`n$pdfPath",
                        "Export Complete",
                        "OK",
                        "Information"
                    )

                    $result = [System.Windows.MessageBox]::Show("Would you like to open the PDF?", "Open Report", "YesNo", "Question")
                    if ($result -eq "Yes") {
                        Start-Process $pdfPath
                    }
                } else {
                    throw "PDF file was not created"
                }
            } else {
                # Fallback: Open HTML in browser for manual print-to-PDF
                Write-Log "No compatible browser found for headless PDF. Opening in browser." "WARNING"
                Start-Process $tempHtmlPath

                [System.Windows.MessageBox]::Show(
                    "PDF generation requires Microsoft Edge or Google Chrome.`n`n" +
                    "The report has been opened in your browser. To save as PDF:`n" +
                    "1. Press Ctrl+P to open Print dialog`n" +
                    "2. Select 'Save as PDF' or 'Microsoft Print to PDF'`n" +
                    "3. Click Save and choose your location",
                    "Manual PDF Export",
                    "OK",
                    "Information"
                )
            }
        }
        catch {
            Write-Log "PDF export failed: $_" "ERROR"
            [System.Windows.MessageBox]::Show(
                "Failed to export PDF:`n$_`n`nTry exporting to HTML and printing to PDF from your browser.",
                "Export Failed",
                "OK",
                "Error"
            )
        }
        finally {
            # Clean up temp file after a delay (browser might still be reading it)
            Start-Job -ScriptBlock {
                param($path)
                Start-Sleep -Seconds 5
                if (Test-Path $path) { Remove-Item $path -Force -ErrorAction SilentlyContinue }
            } -ArgumentList $tempHtmlPath | Out-Null
        }

        $Controls['txtStatusBar'].Text = "Ready"
    }
})

$Controls['btnCopyClipboard'].Add_Click({
    if ($Script:ScanResults.Count -gt 0) {
        $text = $Script:ScanResults | Format-Table -AutoSize | Out-String
        [System.Windows.Clipboard]::SetText($text)
        [System.Windows.MessageBox]::Show("Results copied to clipboard!", "Copy Complete", "OK", "Information")
    } else {
        [System.Windows.MessageBox]::Show("No results to copy.", "No Data", "OK", "Warning")
    }
})

# Unreachable Tab
$Controls['txtUnreachableFilter'].Add_TextChanged({
    $filterText = $Controls['txtUnreachableFilter'].Text.ToLower()
    $filtered = $Script:UnreachableComputers | Where-Object {
        $_.ComputerName.ToLower().Contains($filterText)
    }
    $Controls['dgUnreachable'].ItemsSource = $null
    $Controls['dgUnreachable'].ItemsSource = $filtered
    $Controls['txtUnreachableCount'].Text = "$($filtered.Count) unreachable"
})

$Controls['btnExportUnreachable'].Add_Click({
    # Use date-based folder for organization
    $dateFolder = Get-OutputPath -FileName ""
    $dateFolder = Split-Path $dateFolder -Parent  # Get just the folder path

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Log Files (*.log)|*.log|CSV Files (*.csv)|*.csv"
    $dialog.FileName = "Unreachable_$(Get-Date -Format 'HHmmss').log"
    $dialog.InitialDirectory = $dateFolder

    if ($dialog.ShowDialog() -eq "OK") {
        if ($dialog.FileName -like "*.csv") {
            $Script:UnreachableComputers | Export-Csv -Path $dialog.FileName -NoTypeInformation
        } else {
            $Script:UnreachableComputers | Select-Object -ExpandProperty ComputerName | Out-File $dialog.FileName
        }
        Write-Log "Unreachable list exported to: $($dialog.FileName)"
        [System.Windows.MessageBox]::Show("List exported to:`n$($dialog.FileName)", "Export Complete", "OK", "Information")
    }
})

# Retry Selected - Rescan unreachable computers
$Controls['btnRetryUnreachable'].Add_Click({
    $selectedItems = @($Controls['dgUnreachable'].SelectedItems)

    if ($selectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select one or more computers to retry.",
            "No Selection",
            "OK",
            "Warning"
        )
        return
    }

    $result = [System.Windows.MessageBox]::Show(
        "Retry scanning $($selectedItems.Count) unreachable computer(s)?",
        "Confirm Retry",
        "YesNo",
        "Question"
    )

    if ($result -ne "Yes") { return }

    Write-Log "Retrying $($selectedItems.Count) previously unreachable computers"

    # Disable button during retry
    $Controls['btnRetryUnreachable'].IsEnabled = $false
    $Controls['progressCard'].Visibility = "Visible"

    # Get computer info for retry
    $computersToRetry = @($selectedItems | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.ComputerName
            OperatingSystem = $_.OperatingSystem
            ComputerType = $_.ComputerType
        }
    })

    # Run retry in background
    $retryRunspace = [runspacefactory]::CreateRunspace()
    $retryRunspace.ApartmentState = "STA"
    $retryRunspace.ThreadOptions = "ReuseThread"
    $retryRunspace.Open()

    # Pre-cache exclusion settings for retry (same pattern as main scan)
    $retryExcludeFadmin = [bool]$Controls['chkExcludeFadmin'].IsChecked
    $retryFadminPattern = $Controls['txtFadminPattern'].Text.Trim()
    $retryExcludeDomainAdmins = [bool]$Controls['chkExcludeDomainAdmins'].IsChecked
    $retryExcludeBuiltinAdmin = [bool]$Controls['chkExcludeBuiltinAdmin'].IsChecked
    $retryCustomExclusions = $Controls['txtCustomExclusions'].Text

    # Use synchronized hashtable for thread-safe UI communication
    $retrySyncHash = [hashtable]::Synchronized(@{
        Window = $Window
        Controls = $Controls
    })
    $retryRunspace.SessionStateProxy.SetVariable('SyncHash', $retrySyncHash)
    $retryRunspace.SessionStateProxy.SetVariable('ComputersToRetry', $computersToRetry)
    $retryRunspace.SessionStateProxy.SetVariable('ScanResults', $Script:ScanResults)
    $retryRunspace.SessionStateProxy.SetVariable('UnreachableComputers', $Script:UnreachableComputers)
    $retryRunspace.SessionStateProxy.SetVariable('ThrottleLimit', [int]$Controls['cmbThrottle'].SelectedItem.Content)
    $retryRunspace.SessionStateProxy.SetVariable('ExcludeFadmin', $retryExcludeFadmin)
    $retryRunspace.SessionStateProxy.SetVariable('FadminPattern', $retryFadminPattern)
    $retryRunspace.SessionStateProxy.SetVariable('ExcludeDomainAdmins', $retryExcludeDomainAdmins)
    $retryRunspace.SessionStateProxy.SetVariable('ExcludeBuiltinAdmin', $retryExcludeBuiltinAdmin)
    $retryRunspace.SessionStateProxy.SetVariable('CustomExclusions', $retryCustomExclusions)
    $retryRunspace.SessionStateProxy.SetVariable('ElevatedCredential', $Script:ElevatedCredential)

    $retryPS = [powershell]::Create()
    $retryPS.Runspace = $retryRunspace

    [void]$retryPS.AddScript({
        # Variables are passed via SetVariable
        # $SyncHash contains Window and Controls for thread-safe UI access

        function Write-Log {
            param([string]$Message, [string]$Level = "INFO")
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message`r`n"
            try {
                $SyncHash.Window.Dispatcher.Invoke([Action]{
                    $SyncHash.Controls['txtLog'].AppendText($logEntry)
                    $SyncHash.Controls['txtLog'].ScrollToEnd()
                }, [System.Windows.Threading.DispatcherPriority]::Background)
            } catch {
                Write-Host "Retry Write-Log Error: $_" -ForegroundColor Red
            }
        }

        function Update-Progress {
            param([string]$Status, [int]$Percent, [string]$Detail = "")
            try {
                $SyncHash.Window.Dispatcher.Invoke([Action]{
                    $SyncHash.Controls['txtProgressStatus'].Text = $Status
                    $SyncHash.Controls['txtProgressPercent'].Text = "$Percent%"
                    $SyncHash.Controls['progressBar'].Value = $Percent
                    $SyncHash.Controls['txtProgressDetail'].Text = $Detail
                }, [System.Windows.Threading.DispatcherPriority]::Background)
            } catch {
                Write-Host "Retry Update-Progress Error: $_" -ForegroundColor Red
            }
        }

        function Get-ExclusionPatterns {
            # Uses pre-cached values - NO blocking Dispatcher calls!
            $patterns = @()
            if ($ExcludeFadmin -and -not [string]::IsNullOrWhiteSpace($FadminPattern)) { $patterns += $FadminPattern }
            if ($ExcludeDomainAdmins) { $patterns += '\\Domain Admins$' }
            if ($ExcludeBuiltinAdmin) { $patterns += '\\Administrator$' }
            $custom = $CustomExclusions -split "`r`n" | Where-Object { $_.Trim() -ne '' }
            $patterns += $custom
            return $patterns
        }

        try {
            $computerNames = @($ComputersToRetry | Select-Object -ExpandProperty Name)
            $totalCount = $computerNames.Count

            Write-Log "Starting retry scan for $totalCount computers"
            Update-Progress -Status "Retrying $totalCount computers..." -Percent 20

            $icParams = @{
                ComputerName  = $computerNames
                ScriptBlock   = { Get-LocalGroupMember -Name 'Administrators' -ErrorAction SilentlyContinue }
                ThrottleLimit = $ThrottleLimit
                ErrorAction   = 'SilentlyContinue'
            }
            if ($ElevatedCredential) { $icParams['Credential'] = $ElevatedCredential }
            $localAdmins = Invoke-Command @icParams

            Update-Progress -Status "Processing retry results..." -Percent 60

            $nowReached = @($localAdmins | Select-Object -ExpandProperty PSComputerName -Unique)
            $stillUnreachable = @($computerNames | Where-Object { $_ -notin $nowReached })

            Write-Log "Retry: $($nowReached.Count) now reachable, $($stillUnreachable.Count) still unreachable"

            $patterns = Get-ExclusionPatterns
            $newResults = @()

            foreach ($admin in $localAdmins) {
                $accountName = $admin.Name
                $excluded = $false
                foreach ($pattern in $patterns) {
                    if ($accountName -match $pattern) { $excluded = $true; break }
                }

                if (-not $excluded) {
                    $comp = $ComputersToRetry | Where-Object { $_.Name -eq $admin.PSComputerName }
                    $existing = $ScanResults | Where-Object { $_.ComputerName -eq $admin.PSComputerName -and $_.AdminName -eq $accountName }

                    if (-not $existing) {
                        $computerType = $comp.ComputerType
                        $objectClass = $admin.ObjectClass

                        # Calculate risk score
                        $riskScore = 0
                        $isLocalAccount = $accountName -notmatch '\\'
                        $isDomainAccount = $accountName -match '\\'

                        if ($objectClass -eq 'User') {
                            if ($isLocalAccount) {
                                $riskScore += 40
                            } elseif ($isDomainAccount) {
                                $accountPart = ($accountName -split '\\')[-1]
                                $isServiceAccount = $accountPart -match '^(svc|service|sql|app|sys|task|batch|iis|mssql|backup|scan)' -or
                                                   $accountPart -match '(svc|service)$'
                                if ($isServiceAccount) { $riskScore += 20 } else { $riskScore += 30 }
                            }
                        } elseif ($objectClass -eq 'Group') {
                            $riskScore += 15
                        }
                        if ($computerType -eq 'Server') { $riskScore += 10 }

                        $riskLevel = switch ($riskScore) {
                            { $_ -ge 60 } { 'HIGH' }
                            { $_ -ge 30 } { 'MEDIUM' }
                            { $_ -ge 1 }  { 'LOW' }
                            default       { 'INFO' }
                        }
                        $riskColor = switch ($riskLevel) {
                            'HIGH'   { '#FF6B6B' }; 'MEDIUM' { '#FFD93D' }; 'LOW' { '#4ECDC4' }; default { '#A0AEC0' }
                        }
                        $riskBgColor = switch ($riskLevel) {
                            'HIGH'   { '#3D1212' }; 'MEDIUM' { '#3D2E05' }; 'LOW' { '#0D2F2D' }; default { '#2D3748' }
                        }

                        $newResults += [PSCustomObject]@{
                            ComputerName = $admin.PSComputerName
                            ComputerType = $computerType
                            OperatingSystem = $comp.OperatingSystem
                            AdminName = $accountName
                            ObjectClass = $objectClass
                            RiskScore = $riskScore
                            RiskLevel = $riskLevel
                            RiskColor = $riskColor
                            RiskBgColor = $riskBgColor
                        }
                    }
                }
            }

            # Add new results
            foreach ($r in $newResults) { $null = $ScanResults.Add($r) }

            # Remove reached computers from unreachable list
            $toRemove = @($UnreachableComputers | Where-Object { $_.ComputerName -in $nowReached })
            foreach ($item in $toRemove) { $null = $UnreachableComputers.Remove($item) }

            Update-Progress -Status "Retry complete!" -Percent 100 -Detail "$($nowReached.Count) computers now reachable"

        } catch {
            Write-Log "Retry error: $_" "ERROR"
            Update-Progress -Status "Retry failed!" -Percent 0 -Detail $_.Exception.Message
        } finally {
            # Update UI
            $resultsArray = @($ScanResults | Sort-Object -Property RiskScore -Descending | ForEach-Object {
                [PSCustomObject]@{
                    ComputerName = [string]$_.ComputerName
                    ComputerType = [string]$_.ComputerType
                    OperatingSystem = [string]$_.OperatingSystem
                    AdminName = [string]$_.AdminName
                    ObjectClass = [string]$_.ObjectClass
                    RiskScore = [int]$_.RiskScore
                    RiskLevel = [string]$_.RiskLevel
                    RiskColor = [string]$_.RiskColor
                    RiskBgColor = [string]$_.RiskBgColor
                }
            })
            $unreachableArray = @($UnreachableComputers | ForEach-Object {
                [PSCustomObject]@{
                    ComputerName = [string]$_.ComputerName
                    ComputerType = [string]$_.ComputerType
                    OperatingSystem = [string]$_.OperatingSystem
                    ErrorMessage = [string]$_.ErrorMessage
                }
            })

            $resultsCount = $ScanResults.Count
            $unreachableCount = $UnreachableComputers.Count

            # Risk distribution counts
            $highRiskCount = @($ScanResults | Where-Object { $_.RiskLevel -eq 'HIGH' }).Count
            $mediumRiskCount = @($ScanResults | Where-Object { $_.RiskLevel -eq 'MEDIUM' }).Count
            $lowRiskCount = @($ScanResults | Where-Object { $_.RiskLevel -eq 'LOW' }).Count
            $strHighRisk = "$highRiskCount HIGH"
            $strMediumRisk = "$mediumRiskCount MED"
            $strLowRisk = "$lowRiskCount LOW"

            try {
                $SyncHash.Window.Dispatcher.Invoke([Action]{
                    $SyncHash.Controls['btnRetryUnreachable'].IsEnabled = $true

                    # Bind results
                    $SyncHash.Controls['dgResults'].ItemsSource = $null
                    $SyncHash.Controls['dgResults'].ItemsSource = $resultsArray
                    $SyncHash.Controls['dgUnreachable'].ItemsSource = $null
                    $SyncHash.Controls['dgUnreachable'].ItemsSource = $unreachableArray
                    $SyncHash.Controls['txtResultsCount'].Text = "$resultsCount results"
                    $SyncHash.Controls['txtUnreachableCount'].Text = "$unreachableCount unreachable"
                    $SyncHash.Controls['txtUnexpectedAdmins'].Text = $resultsCount.ToString()
                    $SyncHash.Controls['txtUnreachable'].Text = $unreachableCount.ToString()
                    $SyncHash.Controls['txtHighRisk'].Text = $strHighRisk
                    $SyncHash.Controls['txtMediumRisk'].Text = $strMediumRisk
                    $SyncHash.Controls['txtLowRisk'].Text = $strLowRisk
                }, [System.Windows.Threading.DispatcherPriority]::Background)
            } catch { }

            Write-Log "Retry operation completed"
        }
    })

    # Variables passed via SetVariable, not AddArgument
    $null = $retryPS.BeginInvoke()
})

# Output Directory Browse
$Controls['btnBrowseOutput'].Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $Script:OutputDirectory
    $dialog.Description = "Select output directory for scan results"

    if ($dialog.ShowDialog() -eq "OK") {
        $Script:OutputDirectory = $dialog.SelectedPath
        $Controls['txtOutputDir'].Text = $Script:OutputDirectory
        Write-Log "Output directory changed to: $Script:OutputDirectory"
    }
})

#region Computer Exclusion Buttons
$Controls['btnAddComputerExclusion'].Add_Click({
    $inputDialog = New-Object System.Windows.Window
    $inputDialog.Title = "Add Computer Exclusion"
    $inputDialog.Width = 420
    $inputDialog.Height = 320
    $inputDialog.WindowStartupLocation = "CenterOwner"
    $inputDialog.Owner = $Window
    $inputDialog.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#0D1117"))
    $inputDialog.ResizeMode = "NoResize"

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = "20"

    # Computer Name
    $lblName = New-Object System.Windows.Controls.TextBlock
    $lblName.Text = "Computer Name:"
    $lblName.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $lblName.FontSize = 14
    $lblName.Margin = "0,0,0,5"
    $lblName.VerticalAlignment = "Top"

    $txtName = New-Object System.Windows.Controls.TextBox
    $txtName.Margin = "0,25,0,0"
    $txtName.VerticalAlignment = "Top"
    $txtName.Height = 36
    $txtName.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#161B22"))
    $txtName.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $txtName.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#30363D"))
    $txtName.Padding = "10,8"

    # Exclusion Type
    $lblType = New-Object System.Windows.Controls.TextBlock
    $lblType.Text = "Exclusion Type:"
    $lblType.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $lblType.FontSize = 14
    $lblType.Margin = "0,75,0,5"
    $lblType.VerticalAlignment = "Top"

    $cmbType = New-Object System.Windows.Controls.ComboBox
    $cmbType.Margin = "0,100,0,0"
    $cmbType.VerticalAlignment = "Top"
    $cmbType.Height = 36
    $cmbType.Items.Add("EXCLUDED") | Out-Null
    $cmbType.Items.Add("TEST") | Out-Null
    $cmbType.Items.Add("DEV") | Out-Null
    $cmbType.SelectedIndex = 0

    # Reason
    $lblReason = New-Object System.Windows.Controls.TextBlock
    $lblReason.Text = "Reason (optional):"
    $lblReason.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $lblReason.FontSize = 14
    $lblReason.Margin = "0,150,0,5"
    $lblReason.VerticalAlignment = "Top"

    $txtReason = New-Object System.Windows.Controls.TextBox
    $txtReason.Margin = "0,175,0,0"
    $txtReason.VerticalAlignment = "Top"
    $txtReason.Height = 36
    $txtReason.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#161B22"))
    $txtReason.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $txtReason.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#30363D"))
    $txtReason.Padding = "10,8"
    $txtReason.Text = "User excluded"

    # Buttons
    $btnAdd = New-Object System.Windows.Controls.Button
    $btnAdd.Content = "Add Exclusion"
    $btnAdd.Margin = "0,225,110,0"
    $btnAdd.VerticalAlignment = "Top"
    $btnAdd.HorizontalAlignment = "Right"
    $btnAdd.Width = 120
    $btnAdd.Height = 36
    $btnAdd.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#0EA5E9"))
    $btnAdd.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $btnAdd.BorderThickness = "0"
    $btnAdd.Cursor = "Hand"

    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Cancel"
    $btnCancel.Margin = "0,225,0,0"
    $btnCancel.VerticalAlignment = "Top"
    $btnCancel.HorizontalAlignment = "Right"
    $btnCancel.Width = 100
    $btnCancel.Height = 36
    $btnCancel.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#161B22"))
    $btnCancel.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $btnCancel.BorderThickness = "0"
    $btnCancel.Cursor = "Hand"

    $btnCancel.Add_Click({ $inputDialog.Close() })
    $btnAdd.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
            [System.Windows.MessageBox]::Show("Please enter a computer name.", "Required", "OK", "Warning")
            return
        }
        $result = Add-ComputerExclusion -ComputerName $txtName.Text -ExclusionType $cmbType.SelectedItem -Reason $txtReason.Text
        if ($result) {
            $Controls['txtStatusBar'].Text = "Added '$($txtName.Text.ToUpper())' to computer exclusions"
        }
        $inputDialog.Close()
    })

    $grid.Children.Add($lblName)
    $grid.Children.Add($txtName)
    $grid.Children.Add($lblType)
    $grid.Children.Add($cmbType)
    $grid.Children.Add($lblReason)
    $grid.Children.Add($txtReason)
    $grid.Children.Add($btnAdd)
    $grid.Children.Add($btnCancel)

    $inputDialog.Content = $grid
    $inputDialog.ShowDialog()
})

$Controls['btnAddHoneypot'].Add_Click({
    $inputDialog = New-Object System.Windows.Window
    $inputDialog.Title = "Add Honeypot"
    $inputDialog.Width = 420
    $inputDialog.Height = 240
    $inputDialog.WindowStartupLocation = "CenterOwner"
    $inputDialog.Owner = $Window
    $inputDialog.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#0D1117"))
    $inputDialog.ResizeMode = "NoResize"

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = "20"

    # Warning banner
    $warningBorder = New-Object System.Windows.Controls.Border
    $warningBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#3D1F1F"))
    $warningBorder.CornerRadius = "6"
    $warningBorder.Padding = "12,8"
    $warningBorder.Margin = "0,0,0,15"
    $warningBorder.VerticalAlignment = "Top"

    $warningText = New-Object System.Windows.Controls.TextBlock
    $warningText.Text = "! Honeypots are security decoys. Never scan these systems."
    $warningText.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#EF4444"))
    $warningText.FontSize = 12
    $warningBorder.Child = $warningText

    # Computer Name
    $lblName = New-Object System.Windows.Controls.TextBlock
    $lblName.Text = "Honeypot Computer Name:"
    $lblName.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $lblName.FontSize = 14
    $lblName.Margin = "0,55,0,5"
    $lblName.VerticalAlignment = "Top"

    $txtName = New-Object System.Windows.Controls.TextBox
    $txtName.Margin = "0,80,0,0"
    $txtName.VerticalAlignment = "Top"
    $txtName.Height = 36
    $txtName.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#161B22"))
    $txtName.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $txtName.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#30363D"))
    $txtName.Padding = "10,8"

    # Buttons
    $btnAdd = New-Object System.Windows.Controls.Button
    $btnAdd.Content = "Add Honeypot"
    $btnAdd.Margin = "0,140,110,0"
    $btnAdd.VerticalAlignment = "Top"
    $btnAdd.HorizontalAlignment = "Right"
    $btnAdd.Width = 120
    $btnAdd.Height = 36
    $btnAdd.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#EF4444"))
    $btnAdd.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#FFFFFF"))
    $btnAdd.BorderThickness = "0"
    $btnAdd.Cursor = "Hand"

    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Cancel"
    $btnCancel.Margin = "0,140,0,0"
    $btnCancel.VerticalAlignment = "Top"
    $btnCancel.HorizontalAlignment = "Right"
    $btnCancel.Width = 100
    $btnCancel.Height = 36
    $btnCancel.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#161B22"))
    $btnCancel.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F8FAFC"))
    $btnCancel.BorderThickness = "0"
    $btnCancel.Cursor = "Hand"

    $btnCancel.Add_Click({ $inputDialog.Close() })
    $btnAdd.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
            [System.Windows.MessageBox]::Show("Please enter a honeypot computer name.", "Required", "OK", "Warning")
            return
        }
        $result = Add-ComputerExclusion -ComputerName $txtName.Text -ExclusionType "HONEYPOT" -Reason "Security honeypot - never scan"
        if ($result) {
            $Controls['txtStatusBar'].Text = "Added honeypot '$($txtName.Text.ToUpper())' to exclusions"
        }
        $inputDialog.Close()
    })

    $grid.Children.Add($warningBorder)
    $grid.Children.Add($lblName)
    $grid.Children.Add($txtName)
    $grid.Children.Add($btnAdd)
    $grid.Children.Add($btnCancel)

    $inputDialog.Content = $grid
    $inputDialog.ShowDialog()
})

$Controls['btnRemoveComputerExclusion'].Add_Click({
    $selected = @($Controls['lstComputerExclusions'].SelectedItems)
    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select one or more computers to remove from the exclusion list.",
            "No Selection",
            "OK",
            "Warning"
        )
        return
    }

    $names = ($selected | ForEach-Object { $_.ComputerName }) -join ", "
    $result = [System.Windows.MessageBox]::Show(
        "Remove the following computer(s) from exclusions?`n`n$names`n`nThese computers will be scanned in future audits.",
        "Confirm Removal",
        "YesNo",
        "Warning"
    )

    if ($result -eq "Yes") {
        $removed = Remove-ComputerExclusion -SelectedItems $selected
        $Controls['txtStatusBar'].Text = "Removed $removed computer(s) from exclusions"
    }
})

# Quick Add functionality
$Controls['btnQuickAdd'].Add_Click({
    $computerName = $Controls['txtQuickAddComputer'].Text.Trim()
    if ([string]::IsNullOrWhiteSpace($computerName)) { return }

    $result = Add-ComputerExclusion -ComputerName $computerName -ExclusionType "EXCLUDED" -Reason "Quick add exclusion"
    if ($result) {
        $Controls['txtQuickAddComputer'].Text = ""
        $Controls['txtStatusBar'].Text = "Added '$($computerName.ToUpper())' to exclusions"
    }
})

# Quick Add via Enter key
$Controls['txtQuickAddComputer'].Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq 'Return') {
        $computerName = $Controls['txtQuickAddComputer'].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($computerName)) { return }

        $result = Add-ComputerExclusion -ComputerName $computerName -ExclusionType "EXCLUDED" -Reason "Quick add exclusion"
        if ($result) {
            $Controls['txtQuickAddComputer'].Text = ""
            $Controls['txtStatusBar'].Text = "Added '$($computerName.ToUpper())' to exclusions"
        }
        $e.Handled = $true
    }
})
#endregion Computer Exclusion Buttons

# Trends Tab
$Controls['btnRefreshTrends'].Add_Click({
    Refresh-TrendDashboard
})

$Controls['btnExportTrendChart'].Add_Click({
    $dateFolder = Get-OutputPath -FileName ""
    $dateFolder = Split-Path $dateFolder -Parent

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "HTML Files (*.html)|*.html"
    $dialog.FileName = "TrendChart_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $dialog.InitialDirectory = $dateFolder
    $dialog.Title = "Save Trend Chart"

    if ($dialog.ShowDialog() -eq "OK") {
        Export-TrendChart -FilePath $dialog.FileName
        [System.Windows.MessageBox]::Show(
            "Trend chart exported to:`n$($dialog.FileName)",
            "Export Complete",
            "OK",
            "Information"
        )

        $result = [System.Windows.MessageBox]::Show("Would you like to open the chart?", "Open Chart", "YesNo", "Question")
        if ($result -eq "Yes") {
            Start-Process $dialog.FileName
        }
    }
})

# Log Tab
$Controls['btnClearLog'].Add_Click({
    $Controls['txtLog'].Clear()
})

$Controls['btnSaveLog'].Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt"
    $dialog.FileName = "AuditLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $dialog.InitialDirectory = $Script:OutputDirectory

    if ($dialog.ShowDialog() -eq "OK") {
        $Controls['txtLog'].Text | Out-File $dialog.FileName
        [System.Windows.MessageBox]::Show("Log saved to:`n$($dialog.FileName)", "Save Complete", "OK", "Information")
    }
})

# History Tab Buttons
$Controls['btnRefreshHistory'].Add_Click({
    Write-Log "Refreshing history - scanning date folders for export files..."
    Refresh-HistoryGrid
    Write-Log "History refreshed from: $Script:OutputDirectory"
})

$Controls['btnLoadHistory'].Add_Click({
    $selected = $Controls['dgHistory'].SelectedItem
    if (-not $selected) {
        [System.Windows.MessageBox]::Show(
            "Please select an export file from the list.",
            "No Selection",
            "OK",
            "Warning"
        )
        return
    }

    $filePath = $selected.FullPath

    if (-not (Test-Path $filePath)) {
        [System.Windows.MessageBox]::Show(
            "File not found:`n$filePath",
            "File Not Found",
            "OK",
            "Error"
        )
        return
    }

    # CSV files - load into results grid
    if ($filePath -match '\.csv$') {
        try {
            Write-Log "Loading CSV file: $filePath"
            $csvData = Import-Csv $filePath

            # Clear current results
            $Script:ScanResults.Clear()

            # Load CSV data into results
            foreach ($row in $csvData) {
                $riskInfo = Get-RiskScore -AdminName $row.AdminName -ComputerType $row.ComputerType -ObjectClass $row.ObjectClass
                $null = $Script:ScanResults.Add([PSCustomObject]@{
                    ComputerName = $row.ComputerName
                    ComputerType = $row.ComputerType
                    OperatingSystem = $row.OperatingSystem
                    AdminName = $row.AdminName
                    ObjectClass = $row.ObjectClass
                    RiskScore = $riskInfo.RiskScore
                    RiskLevel = $riskInfo.RiskLevel
                    RiskColor = $riskInfo.RiskColor
                    RiskBgColor = $riskInfo.RiskBgColor
                })
            }

            # Update grid
            $Controls['dgResults'].ItemsSource = $null
            $Controls['dgResults'].ItemsSource = $Script:ScanResults
            $Controls['txtResultsCount'].Text = "$($Script:ScanResults.Count) results (loaded from file)"

            # Switch to Scanner tab
            $Controls['MainTabControl'].SelectedIndex = 0

            Write-Log "Loaded $($Script:ScanResults.Count) results from CSV"
            [System.Windows.MessageBox]::Show(
                "Loaded $($Script:ScanResults.Count) results from:`n$($selected.FileName)",
                "File Loaded",
                "OK",
                "Information"
            )
        } catch {
            Write-Log "Error loading CSV: $_" "ERROR"
            [System.Windows.MessageBox]::Show("Error loading CSV file:`n$_", "Error", "OK", "Error")
        }
    }
    # HTML files - open in default browser
    else {
        Write-Log "Opening HTML file: $filePath"
        Start-Process $filePath
    }
})

$Controls['btnDeleteHistory'].Add_Click({
    $selected = $Controls['dgHistory'].SelectedItem
    if (-not $selected) {
        [System.Windows.MessageBox]::Show(
            "Please select a file to delete.",
            "No Selection",
            "OK",
            "Warning"
        )
        return
    }

    $result = [System.Windows.MessageBox]::Show(
        "Are you sure you want to delete this file?`n`nFile: $($selected.FileName)`nDate: $($selected.DateFormatted)`nPath: $($selected.FullPath)",
        "Confirm Delete",
        "YesNo",
        "Warning"
    )

    if ($result -eq "Yes") {
        try {
            Remove-Item $selected.FullPath -Force
            Write-Log "Deleted file: $($selected.FullPath)"
            Refresh-HistoryGrid
        } catch {
            Write-Log "Error deleting file: $_" "ERROR"
            [System.Windows.MessageBox]::Show("Error deleting file:`n$_", "Error", "OK", "Error")
        }
    }
})

$Controls['btnClearHistory'].Add_Click({
    # Count files in date folders
    $baseDir = $Script:OutputDirectory
    $allFiles = @()

    $dateFolders = Get-ChildItem -Path $baseDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' }

    foreach ($folder in $dateFolders) {
        $files = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in '.csv', '.html' }
        $allFiles += $files
    }

    if ($allFiles.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No export files to clear.", "Empty History", "OK", "Information")
        return
    }

    $result = [System.Windows.MessageBox]::Show(
        "Are you sure you want to delete ALL export files?`n`nThis will delete $($allFiles.Count) files from $($dateFolders.Count) date folders and cannot be undone.",
        "Confirm Clear All",
        "YesNo",
        "Warning"
    )

    if ($result -eq "Yes") {
        $deleted = 0
        $errors = 0

        foreach ($file in $allFiles) {
            try {
                Remove-Item $file.FullName -Force
                $deleted++
            } catch {
                $errors++
                Write-Log "Error deleting $($file.Name): $_" "ERROR"
            }
        }

        # Remove empty date folders
        foreach ($folder in $dateFolders) {
            $remaining = Get-ChildItem -Path $folder.FullName -ErrorAction SilentlyContinue
            if ($remaining.Count -eq 0) {
                Remove-Item $folder.FullName -Force -ErrorAction SilentlyContinue
            }
        }

        Refresh-HistoryGrid
        Write-Log "Cleared all history: deleted $deleted files, $errors errors"

        if ($errors -gt 0) {
            [System.Windows.MessageBox]::Show("Deleted $deleted files with $errors errors.`nCheck the Activity Log for details.", "History Cleared", "OK", "Warning")
        } else {
            [System.Windows.MessageBox]::Show("Deleted $deleted export files.", "History Cleared", "OK", "Information")
        }
    }
})

#region Results Grid Double-Click (Show Computer Detail Pane)
$Controls['dgResults'].Add_MouseDoubleClick({
    $selected = $Controls['dgResults'].SelectedItem
    if ($selected) {
        $computerName = $selected.ComputerName

        # Get all admin accounts for this computer
        $computerAdmins = @($Script:ScanResults | Where-Object { $_.ComputerName -eq $computerName })

        if ($computerAdmins.Count -gt 0) {
            # Update detail pane
            $Controls['txtDetailComputerName'].Text = $computerName
            $Controls['txtDetailOS'].Text = $selected.OperatingSystem
            $Controls['lstDetailAdmins'].ItemsSource = $computerAdmins

            # Show the pane
            $Controls['computerDetailPane'].Visibility = 'Visible'
            $Controls['txtStatusBar'].Text = "Showing $($computerAdmins.Count) admin account(s) for $computerName"
        }
    }
})

# Close Detail Pane Button
$Controls['btnCloseDetail'].Add_Click({
    $Controls['computerDetailPane'].Visibility = 'Collapsed'
})
#endregion

#region Context Menu Handlers - Results DataGrid
$Controls['ctxResultsCopyComputer'].Add_Click({
    $selected = $Controls['dgResults'].SelectedItem
    if ($selected) {
        [System.Windows.Clipboard]::SetText($selected.ComputerName)
        $Controls['txtStatusBar'].Text = "Copied computer name to clipboard: $($selected.ComputerName)"
    }
})

$Controls['ctxResultsCopyAdmin'].Add_Click({
    $selected = $Controls['dgResults'].SelectedItem
    if ($selected) {
        [System.Windows.Clipboard]::SetText($selected.AdminName)
        $Controls['txtStatusBar'].Text = "Copied admin account to clipboard: $($selected.AdminName)"
    }
})

$Controls['ctxResultsCopyRow'].Add_Click({
    $selected = $Controls['dgResults'].SelectedItem
    if ($selected) {
        $rowText = "$($selected.ComputerName)`t$($selected.AdminName)`t$($selected.ObjectClass)`t$($selected.RiskLevel)"
        [System.Windows.Clipboard]::SetText($rowText)
        $Controls['txtStatusBar'].Text = "Copied row to clipboard"
    }
})

$Controls['ctxResultsPing'].Add_Click({
    $selected = $Controls['dgResults'].SelectedItem
    if ($selected) {
        $Controls['txtStatusBar'].Text = "Pinging $($selected.ComputerName)..."
        Start-Process cmd -ArgumentList "/c ping $($selected.ComputerName) & pause"
    }
})

$Controls['ctxResultsRDP'].Add_Click({
    $selected = $Controls['dgResults'].SelectedItem
    if ($selected) {
        $Controls['txtStatusBar'].Text = "Opening Remote Desktop to $($selected.ComputerName)..."
        Start-Process mstsc -ArgumentList "/v:$($selected.ComputerName)"
    }
})

$Controls['ctxResultsAddExclusion'].Add_Click({
    $selected = $Controls['dgResults'].SelectedItem
    if ($selected) {
        # Extract account name without domain
        $accountName = $selected.AdminName -replace '^.*\\', ''
        $currentExclusions = $Controls['txtCustomExclusions'].Text
        if ($currentExclusions -and $currentExclusions.Trim()) {
            $Controls['txtCustomExclusions'].Text = "$currentExclusions`r`n$accountName"
        } else {
            $Controls['txtCustomExclusions'].Text = $accountName
        }
        $Controls['MainTabControl'].SelectedIndex = 3  # Switch to Settings tab
        $Controls['txtStatusBar'].Text = "Added '$accountName' to custom exclusions"
        Write-Log "Added '$accountName' to custom exclusions via context menu"
    }
})

$Controls['ctxResultsExcludeComputer'].Add_Click({
    $selected = $Controls['dgResults'].SelectedItem
    if ($selected) {
        $computerName = $selected.ComputerName
        $result = Add-ComputerExclusion -ComputerName $computerName -ExclusionType "EXCLUDED" -Reason "Excluded from Results grid"
        if ($result) {
            $Controls['MainTabControl'].SelectedIndex = 3  # Switch to Settings tab
            $Controls['txtStatusBar'].Text = "Added '$computerName' to computer exclusions - will not be scanned"
        }
    }
})

$Controls['ctxResultsRescan'].Add_Click({
    $selected = $Controls['dgResults'].SelectedItem
    if (-not $selected) { return }

    $computerName = $selected.ComputerName
    $Controls['txtStatusBar'].Text = "Rescanning $computerName..."
    Write-Log "Rescanning computer: $computerName"

    $Script:ScanResults = [System.Collections.ArrayList]@($Script:ScanResults | Where-Object { $_.ComputerName -ne $computerName })

    try {
        $icParams = @{
            ComputerName = $computerName
            ScriptBlock  = {
                Get-LocalGroupMember -Name 'Administrators' -ErrorAction SilentlyContinue |
                    Select-Object Name, ObjectClass
            }
            ErrorAction  = 'Stop'
        }
        if ($Script:ElevatedCredential) { $icParams['Credential'] = $Script:ElevatedCredential }
        $localAdmins = Invoke-Command @icParams

        foreach ($admin in $localAdmins) {
            if (-not (Test-ExcludeAccount -AccountName $admin.Name)) {
                $computerInfo = $Script:AllComputers | Where-Object { $_.Name -eq $computerName } | Select-Object -First 1
                $isServer = $computerInfo.OperatingSystem -like "*Server*"

                $riskScore = Get-RiskScore -ObjectClass $admin.ObjectClass -AdminName $admin.Name -IsServer $isServer
                $riskColors = Get-RiskColors -Score $riskScore

                [void]$Script:ScanResults.Add([PSCustomObject]@{
                    ComputerName = $computerName
                    ComputerType = if ($isServer) { "Server" } else { "Workstation" }
                    OperatingSystem = $computerInfo.OperatingSystem
                    AdminName = $admin.Name
                    ObjectClass = $admin.ObjectClass
                    RiskScore = $riskScore
                    RiskLevel = $riskColors.Level
                    RiskColor = $riskColors.Color
                    RiskBgColor = $riskColors.BgColor
                })
            }
        }

        Apply-ResultsFilter
        Update-Statistics
        $Controls['txtStatusBar'].Text = "Rescan complete for $computerName"
        Write-Log "Rescan complete for: $computerName"
    }
    catch {
        $Controls['txtStatusBar'].Text = "Failed to rescan $computerName"
        Write-Log "Failed to rescan $computerName`: $_" -Level "ERROR"
    }
})

# Remove Admin Account (Remediation) - Supports multiple selection
$Controls['ctxResultsRemoveAdmin'].Add_Click({
    $selectedItems = @($Controls['dgResults'].SelectedItems)
    if ($selectedItems.Count -eq 0) { return }

    # Build summary of accounts to remove
    $highRiskCount = ($selectedItems | Where-Object { $_.RiskLevel -eq 'HIGH' }).Count
    $totalCount = $selectedItems.Count

    $accountList = ($selectedItems | ForEach-Object { "  - $($_.AdminName) on $($_.ComputerName) [$($_.RiskLevel)]" }) -join "`n"

    # Show serious warning dialog
    $warningMessage = @"
*** REMOTE REMEDIATION WARNING ***

You are about to REMOVE $totalCount admin account(s):

$accountList

This action will:
- Remotely connect to each computer via WinRM
- Remove the listed accounts from local Administrators groups
- Changes take effect IMMEDIATELY

This action CANNOT be easily undone!

Are you absolutely sure you want to proceed?
"@

    $result = [System.Windows.MessageBox]::Show(
        $warningMessage,
        "Confirm Remote Remediation ($totalCount accounts)",
        "YesNo",
        "Warning"
    )

    if ($result -ne "Yes") {
        Write-Log "Bulk remediation cancelled by user ($totalCount accounts)"
        return
    }

    # Double confirmation if any HIGH risk accounts
    if ($highRiskCount -gt 0) {
        $result2 = [System.Windows.MessageBox]::Show(
            "WARNING: $highRiskCount of the selected accounts are HIGH RISK!`n`nAre you CERTAIN you want to remove all $totalCount accounts?",
            "Final Confirmation - HIGH RISK",
            "YesNo",
            "Exclamation"
        )
        if ($result2 -ne "Yes") {
            Write-Log "Bulk remediation cancelled at HIGH RISK confirmation ($totalCount accounts)"
            return
        }
    }

    # Process each account
    $successCount = 0
    $failCount = 0
    $failedItems = @()

    foreach ($item in $selectedItems) {
        $computerName = $item.ComputerName
        $adminName = $item.AdminName

        $Controls['txtStatusBar'].Text = "Removing $adminName from $computerName... ($($successCount + $failCount + 1) of $totalCount)"
        Write-Log "REMEDIATION: Attempting to remove '$adminName' from Administrators group on $computerName" "WARNING"

        try {
            $icParams = @{
                ComputerName = $computerName
                ScriptBlock  = {
                    param($Account)
                    try {
                        Remove-LocalGroupMember -Group 'Administrators' -Member $Account -ErrorAction Stop
                        return @{ Success = $true; Message = "Account removed successfully" }
                    } catch {
                        return @{ Success = $false; Message = $_.Exception.Message }
                    }
                }
                ArgumentList = $adminName
                ErrorAction  = 'Stop'
            }
            if ($Script:ElevatedCredential) { $icParams['Credential'] = $Script:ElevatedCredential }
            $removalResult = Invoke-Command @icParams

            if ($removalResult.Success) {
                $auditEntry = "[REMEDIATION] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | REMOVED '$adminName' from Administrators on $computerName | Operator: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
                Write-Log $auditEntry "WARNING"
                $successCount++

                # Remove from results
                $Script:ScanResults = [System.Collections.ArrayList]@($Script:ScanResults | Where-Object {
                    -not ($_.ComputerName -eq $computerName -and $_.AdminName -eq $adminName)
                })
            } else {
                throw $removalResult.Message
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Log "REMEDIATION FAILED: Could not remove '$adminName' from $computerName - $errorMsg" "ERROR"
            $failCount++
            $failedItems += "$adminName on $computerName`: $errorMsg"
        }
    }

    # Update UI
    Apply-ResultsFilter
    Update-Statistics

    # Show summary
    if ($failCount -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Successfully removed all $successCount account(s) from local Administrators groups.`n`nThe results have been updated.",
            "Remediation Complete",
            "OK",
            "Information"
        )
        $Controls['txtStatusBar'].Text = "Remediation complete: $successCount account(s) removed"
    } else {
        $failedList = $failedItems -join "`n"
        [System.Windows.MessageBox]::Show(
            "Remediation partially complete.`n`nSucceeded: $successCount`nFailed: $failCount`n`nFailed items:`n$failedList",
            "Remediation Results",
            "OK",
            "Warning"
        )
        $Controls['txtStatusBar'].Text = "Remediation: $successCount succeeded, $failCount failed"
    }
})
#endregion

#region Context Menu Handlers - Unreachable DataGrid
$Controls['ctxUnreachableCopyName'].Add_Click({
    $selected = $Controls['dgUnreachable'].SelectedItem
    if ($selected) {
        [System.Windows.Clipboard]::SetText($selected.ComputerName)
        $Controls['txtStatusBar'].Text = "Copied computer name to clipboard: $($selected.ComputerName)"
    }
})

$Controls['ctxUnreachableCopyError'].Add_Click({
    $selected = $Controls['dgUnreachable'].SelectedItem
    if ($selected) {
        [System.Windows.Clipboard]::SetText($selected.ErrorMessage)
        $Controls['txtStatusBar'].Text = "Copied error message to clipboard"
    }
})

$Controls['ctxUnreachablePing'].Add_Click({
    $selected = $Controls['dgUnreachable'].SelectedItem
    if ($selected) {
        $Controls['txtStatusBar'].Text = "Pinging $($selected.ComputerName)..."
        Start-Process cmd -ArgumentList "/c ping $($selected.ComputerName) & pause"
    }
})

$Controls['ctxUnreachableRetryOne'].Add_Click({
    $selected = $Controls['dgUnreachable'].SelectedItem
    if ($selected) {
        # Trigger retry via the existing retry mechanism
        $Controls['dgUnreachable'].SelectedItems.Clear()
        $Controls['dgUnreachable'].SelectedItem = $selected
        $Controls['btnRetryUnreachable'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
})

$Controls['ctxUnreachableRetrySelected'].Add_Click({
    $Controls['btnRetryUnreachable'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
})

$Controls['ctxUnreachableExcludeComputer'].Add_Click({
    $selected = $Controls['dgUnreachable'].SelectedItem
    if ($selected) {
        $computerName = $selected.ComputerName
        $result = Add-ComputerExclusion -ComputerName $computerName -ExclusionType "EXCLUDED" -Reason "Excluded from Unreachable grid - connection issues"
        if ($result) {
            $Controls['MainTabControl'].SelectedIndex = 3  # Switch to Settings tab
            $Controls['txtStatusBar'].Text = "Added '$computerName' to computer exclusions - will not be scanned"
        }
    }
})
#endregion

#region Keyboard Shortcuts
$Window.Add_KeyDown({
    param($sender, $e)

    # F5 - Start Scan
    if ($e.Key -eq [System.Windows.Input.Key]::F5) {
        if ($Controls['btnStartScan'].IsEnabled) {
            $Controls['btnStartScan'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
        $e.Handled = $true
    }

    # Escape - Cancel Scan
    if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
        if ($Controls['btnCancelScan'].IsEnabled) {
            $Controls['btnCancelScan'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
        $e.Handled = $true
    }

    # Ctrl+S - Export CSV
    if ($e.Key -eq [System.Windows.Input.Key]::S -and [System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control) {
        $Controls['btnExportCSV'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        $e.Handled = $true
    }

    # Ctrl+F - Focus Filter
    if ($e.Key -eq [System.Windows.Input.Key]::F -and [System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control) {
        $Controls['MainTabControl'].SelectedIndex = 1  # Switch to Results tab
        $Controls['txtResultsFilter'].Focus()
        $e.Handled = $true
    }

    # Ctrl+R - Retry Selected
    if ($e.Key -eq [System.Windows.Input.Key]::R -and [System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control) {
        $Controls['MainTabControl'].SelectedIndex = 2  # Switch to Unreachable tab
        $Controls['btnRetryUnreachable'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        $e.Handled = $true
    }

    # Ctrl+1 through Ctrl+6 - Switch tabs
    if ([System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control) {
        $tabIndex = switch ($e.Key) {
            D1 { 0 }
            D2 { 1 }
            D3 { 2 }
            D4 { 3 }
            D5 { 4 }
            D6 { 5 }
            default { -1 }
        }
        if ($tabIndex -ge 0 -and $tabIndex -lt $Controls['MainTabControl'].Items.Count) {
            $Controls['MainTabControl'].SelectedIndex = $tabIndex
            $e.Handled = $true
        }
    }
})
#endregion

#region Sound Notification Function
function Play-ScanCompleteSound {
    if ($Controls['chkPlaySound'].IsChecked) {
        try {
            [System.Media.SystemSounds]::Asterisk.Play()
        }
        catch {
            # Silently ignore if sound fails
        }
    }
}

function Update-StatusBar {
    param([string]$Message, [bool]$UpdateLastScan = $false)

    $Controls['txtStatusBar'].Text = $Message

    if ($UpdateLastScan) {
        $Controls['txtLastScanTime'].Text = "Last scan: $(Get-Date -Format 'MMM d, h:mm tt')"
    }
}
#endregion

# Window Closing
$Window.Add_Closing({
    # Save user settings before closing
    Save-UserSettings

    # Clean up event subscriptions
    Get-EventSubscriber -SourceIdentifier "ScanComplete" -ErrorAction SilentlyContinue | Unregister-Event -ErrorAction SilentlyContinue

    if ($Script:CurrentPowerShell) {
        $Script:CurrentPowerShell.Stop()
        $Script:CurrentPowerShell.Dispose()
    }
    if ($Script:CurrentRunspace) {
        $Script:CurrentRunspace.Close()
        $Script:CurrentRunspace.Dispose()
    }
})
#endregion

#region Initialization
# Initialize app data folder
Initialize-AppDataFolder

# Load user settings first (for retention days config)
if (Load-UserSettings) {
    Apply-UserSettings
}

# Auto-cleanup old scan folders (uses configured retention days)
$retentionDays = if ($Script:UserSettings.outputSettings.retentionDays) { $Script:UserSettings.outputSettings.retentionDays } else { 30 }
Invoke-OutputCleanup -RetentionDays $retentionDays

# Load scan history (for NEW/EXISTING detection)
Load-ScanHistory

# Populate history grid
Refresh-HistoryGrid

Write-Log "Local Administrator Auditor started"
Write-Log "App data folder: $Script:AppDataPath"
Write-Log "Output directory: $Script:OutputDirectory"

if ($Script:ScanHistory.scans.Count -gt 0) {
    $lastScan = $Script:ScanHistory.scans[-1]
    $lastScanTime = [DateTime]::Parse($lastScan.timestamp).ToString("MMM d, yyyy h:mm tt")
    Write-Log "Last scan: $lastScanTime ($($lastScan.statistics.findings) findings)"

    # Update status bar with last scan time
    $Controls['txtLastScanTime'].Text = "Last scan: $([DateTime]::Parse($lastScan.timestamp).ToString('MMM d, h:mm tt'))"
}

# Check for saved scan state and show resume banner if found
$savedStateInfo = Get-ScanStateInfo
if ($null -ne $savedStateInfo) {
    Write-Log "Found saved scan state from $($savedStateInfo.SavedAt.ToString('MMM d, h:mm tt'))"

    # Update resume banner with state info
    $Controls['txtResumeTitle'].Text = "Resume Interrupted Scan?"
    $Controls['txtResumeDetails'].Text = "Previous scan was $($savedStateInfo.PercentComplete)% complete ($($savedStateInfo.CompletedCount) of $($savedStateInfo.TotalComputers) computers)"
    $Controls['txtResumeSavedAt'].Text = "Saved: $($savedStateInfo.SavedAt.ToString('dddd, MMMM d, yyyy h:mm tt'))"

    # Show the resume banner
    $Controls['resumeScanBanner'].Visibility = 'Visible'
}
#endregion

# Show the window
$Window.ShowDialog() | Out-Null
