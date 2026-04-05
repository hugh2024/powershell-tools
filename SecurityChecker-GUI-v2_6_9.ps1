# ============================================================================
# SecurityChecker v2.6
# Enterprise Subnet Security Configuration Scanner
# WinPS 5.1 + PS7 compatible | Runspace parallelism | Responsive WinForms UI
# Checks: LLMNR, mDNS, NetBIOS, SMBv1, SMB Signing, WPAD, IPv6, RDP port
#
# Author: Hugh Gozlou
# GitHub: https://github.com/hughgozlou
# ============================================================================
#
# CHANGELOG:
#   v2.6 - NEW: Right-click "Enable WinRM" on error rows to remotely fix WinRM
#        - NEW: Full prerequisite check dialog with guided install for new users
#        - Prerequisites: WinRM, PSRemoting, Admin rights — prompts to fix all
#   v2.5 - FIXED: Grid uses fixed column widths + horizontal scroll (no truncation)
#        - FIXED: Prerequisites install silently (no PS prompts, no Y/N)
#        - Auto-configures WinRM/PSRemoting before GUI if needed
#   v2.4 - Prerequisite check dialog, shorter grid values
#   v2.3 - NetBIOS label fix, wider columns
#   v2.2 - Tooltips on all controls, RDP port 3389 check
#   v2.1 - Custom credential dialog, WinRM port pre-check, auth selector
#   v2.0 - Full rewrite: multi-check, professional UI, color-coded results
#
# ============================================================================

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

# =======================
# PREREQUISITE CHECK WITH GUI DIALOG
# =======================

# Check admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

function Show-PrerequisiteDialog {
    # -------------------------------------------------------
    # Gather ALL prerequisite statuses
    # -------------------------------------------------------
    $prereqs = @()

    # --- CATEGORY: Core Requirements ---

    # 1. Administrator
    $prereqs += [pscustomobject]@{
        Name        = "Run as Administrator"
        Category    = "Core"
        Status      = if ($isAdmin) { "OK" } else { "MISSING" }
        Description = "Required to configure services and install features. Right-click script > Run as Administrator."
        CanFix      = $false
        FixAction   = $null
        FixLabel    = ""
    }

    # 2. Execution Policy
    $execPolicy = Get-ExecutionPolicy -ErrorAction SilentlyContinue
    $execOk = $execPolicy -in @("RemoteSigned","Unrestricted","Bypass","AllSigned")
    $prereqs += [pscustomobject]@{
        Name        = "Execution Policy"
        Category    = "Core"
        Status      = if ($execOk) { "OK ($execPolicy)" } else { "RESTRICTED ($execPolicy)" }
        Description = "Must allow script execution. Will set to RemoteSigned."
        CanFix      = ($isAdmin -and -not $execOk)
        FixAction   = { Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force -ErrorAction Stop }
        FixLabel    = "Set-ExecutionPolicy RemoteSigned"
    }

    # --- CATEGORY: WinRM & Remoting ---

    # 3. WinRM Service
    $winrmSvc = Get-Service -Name "WinRM" -ErrorAction SilentlyContinue
    $winrmOk = ($winrmSvc -and $winrmSvc.Status -eq 'Running')
    $prereqs += [pscustomobject]@{
        Name        = "WinRM Service"
        Category    = "Remoting"
        Status      = if ($winrmOk) { "OK (Running)" } elseif (-not $winrmSvc) { "NOT FOUND" } else { "STOPPED" }
        Description = "Windows Remote Management service must be running."
        CanFix      = ($isAdmin -and $winrmSvc -and $winrmSvc.Status -ne 'Running')
        FixAction   = {
            Set-Service -Name "WinRM" -StartupType Automatic -ErrorAction Stop
            Start-Service -Name "WinRM" -ErrorAction Stop
        }
        FixLabel    = "Start WinRM service, set to Automatic"
    }

    # 4. PSRemoting / WinRM Listeners
    $hasListeners = $false
    try {
        $listeners = Get-ChildItem WSMan:\localhost\Listener -ErrorAction Stop 2>$null
        if ($listeners -and $listeners.Count -gt 0) { $hasListeners = $true }
    } catch { }
    $prereqs += [pscustomobject]@{
        Name        = "PS Remoting (WinRM Listeners)"
        Category    = "Remoting"
        Status      = if ($hasListeners) { "OK (Configured)" } else { "NOT CONFIGURED" }
        Description = "Enables Invoke-Command. Will run: winrm quickconfig + Enable-PSRemoting."
        CanFix      = ($isAdmin -and -not $hasListeners)
        FixAction   = {
            & winrm quickconfig /quiet /force 2>&1 | Out-Null
            Enable-PSRemoting -Force -SkipNetworkProfileCheck -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -Confirm:$false 2>$null | Out-Null
        }
        FixLabel    = "winrm quickconfig /quiet /force + Enable-PSRemoting -Force"
    }

    # 5. TrustedHosts (informational)
    $trustedHosts = ""
    try { $trustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value } catch {}
    $prereqs += [pscustomobject]@{
        Name        = "TrustedHosts (Optional)"
        Category    = "Remoting"
        Status      = if ($trustedHosts -and $trustedHosts -ne "") { "OK (SET: $trustedHosts)" } else { "EMPTY (OK for domain-joined PCs)" }
        Description = "Only needed for workgroup/cross-domain targets. Set to * for all or specific IPs."
        CanFix      = $false
        FixAction   = $null
        FixLabel    = ""
    }

    # --- CATEGORY: PowerShell Modules ---

    # 6. SmbShare module (Get-SmbServerConfiguration)
    $smbMod = Get-Module -ListAvailable -Name "SmbShare" -ErrorAction SilentlyContinue
    $prereqs += [pscustomobject]@{
        Name        = "SmbShare Module"
        Category    = "Modules"
        Status      = if ($smbMod) { "OK (Installed)" } else { "MISSING" }
        Description = "Provides Get-SmbServerConfiguration for SMBv1 and SMB Signing checks."
        CanFix      = ($isAdmin -and -not $smbMod)
        FixAction   = {
            # SmbShare is a Windows feature module — try to import or install via DISM
            try {
                Import-Module SmbShare -ErrorAction Stop
            } catch {
                # Try enabling the Windows feature
                $null = dism.exe /Online /Enable-Feature /FeatureName:SmbDirect /NoRestart 2>&1
                Import-Module SmbShare -ErrorAction SilentlyContinue
            }
        }
        FixLabel    = "Import-Module SmbShare (built-in Windows module)"
    }

    # 7. NetAdapter module (Get-NetAdapterBinding for IPv6)
    $netAdMod = Get-Module -ListAvailable -Name "NetAdapter" -ErrorAction SilentlyContinue
    $prereqs += [pscustomobject]@{
        Name        = "NetAdapter Module"
        Category    = "Modules"
        Status      = if ($netAdMod) { "OK (Installed)" } else { "MISSING" }
        Description = "Provides Get-NetAdapterBinding for IPv6 check."
        CanFix      = $false
        FixAction   = $null
        FixLabel    = "Built-in Windows module. Reinstall RSAT or repair Windows if missing."
    }

    # 8. CimCmdlets (Get-CimInstance for OS info)
    $cimMod = Get-Module -ListAvailable -Name "CimCmdlets" -ErrorAction SilentlyContinue
    $prereqs += [pscustomobject]@{
        Name        = "CimCmdlets Module"
        Category    = "Modules"
        Status      = if ($cimMod) { "OK (Installed)" } else { "MISSING" }
        Description = "Provides Get-CimInstance for OS detection on remote hosts."
        CanFix      = $false
        FixAction   = $null
        FixLabel    = "Built-in PowerShell module. Should be present on all Windows systems."
    }

    # 9. Microsoft.PowerShell.Management (Get-Service, Test-Connection, etc.)
    $mgmtMod = Get-Module -ListAvailable -Name "Microsoft.PowerShell.Management" -ErrorAction SilentlyContinue
    $prereqs += [pscustomobject]@{
        Name        = "PowerShell.Management Module"
        Category    = "Modules"
        Status      = if ($mgmtMod) { "OK (Installed)" } else { "MISSING" }
        Description = "Core module for Get-Service, Test-Connection, Get-ItemProperty."
        CanFix      = $false
        FixAction   = $null
        FixLabel    = "Core PowerShell module. Repair PowerShell installation if missing."
    }

    # --- CATEGORY: Windows Features ---

    # 10. Windows Remote Management (WinRM) Windows feature / firewall rule
    $winrmFwRule = $false
    try {
        $fwRules = Get-NetFirewallRule -DisplayName "Windows Remote Management*" -ErrorAction SilentlyContinue
        if ($fwRules) {
            $enabled = $fwRules | Where-Object { $_.Enabled -eq $true -and $_.Direction -eq "Inbound" }
            if ($enabled) { $winrmFwRule = $true }
        }
    } catch {
        # Get-NetFirewallRule may not exist on older systems, skip
        $winrmFwRule = $true  # Assume OK if we can't check
    }
    $prereqs += [pscustomobject]@{
        Name        = "WinRM Firewall Rule"
        Category    = "Windows Features"
        Status      = if ($winrmFwRule) { "OK (Enabled)" } else { "DISABLED or MISSING" }
        Description = "Inbound firewall rule for WinRM (TCP 5985/5986) must be enabled."
        CanFix      = ($isAdmin -and -not $winrmFwRule)
        FixAction   = {
            try {
                Enable-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)" -ErrorAction Stop
            } catch {
                # Fallback: use netsh
                & netsh advfirewall firewall set rule name="Windows Remote Management (HTTP-In)" dir=in new enable=yes 2>&1 | Out-Null
            }
        }
        FixLabel    = "Enable WinRM inbound firewall rule (TCP 5985)"
    }

    # 11. .NET Framework version (need 4.5+ for WinForms features used)
    $dotnetOk = $false
    try {
        $release = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release
        if ($release -ge 378389) { $dotnetOk = $true }  # 4.5+
    } catch {}
    $prereqs += [pscustomobject]@{
        Name        = ".NET Framework 4.5+"
        Category    = "Windows Features"
        Status      = if ($dotnetOk) { "OK (Installed)" } else { "MISSING or OLD" }
        Description = "Required for WinForms GUI. Windows 10/11 includes this by default."
        CanFix      = $false
        FixAction   = $null
        FixLabel    = "Install from microsoft.com/net/download if missing."
    }

    # 12. PowerShell version check
    $psVer = $PSVersionTable.PSVersion
    $psOk = ($psVer.Major -ge 5)
    $prereqs += [pscustomobject]@{
        Name        = "PowerShell 5.1+"
        Category    = "Windows Features"
        Status      = if ($psOk) { "OK (v$($psVer.Major).$($psVer.Minor))" } else { "OLD (v$($psVer.Major).$($psVer.Minor))" }
        Description = "PowerShell 5.1 or later is required. Windows 10/11 includes 5.1."
        CanFix      = $false
        FixAction   = $null
        FixLabel    = "Update via Windows Update or install WMF 5.1."
    }

    # -------------------------------------------------------
    # Determine if all OK — skip dialog if everything passes
    # -------------------------------------------------------
    $problemItems = @($prereqs | Where-Object {
        $_.Status -notmatch "^OK" -and
        $_.Name -ne "TrustedHosts (Optional)"
    })
    if ($problemItems.Count -eq 0) { return $true }

    # -------------------------------------------------------
    # BUILD THE GUI DIALOG
    # -------------------------------------------------------
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "SecurityChecker v2.6 - Prerequisite Check"
    $dlg.Size = New-Object System.Drawing.Size(720, 620)
    $dlg.StartPosition = "CenterScreen"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(236, 240, 241)
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

    # Title
    $lblT = New-Object System.Windows.Forms.Label
    $lblT.Text = "Prerequisite Check"
    $lblT.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblT.ForeColor = [System.Drawing.Color]::FromArgb(24, 42, 68)
    $lblT.Location = New-Object System.Drawing.Point(20, 12)
    $lblT.AutoSize = $true
    $dlg.Controls.Add($lblT)

    # Summary counts
    $fixable = @($prereqs | Where-Object { $_.CanFix -eq $true })
    $failCount = $problemItems.Count
    $fixCount = $fixable.Count
    $okCount = ($prereqs | Where-Object { $_.Status -match "^OK" }).Count

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = "$okCount passed, $failCount need attention ($fixCount auto-fixable). Review below:"
    $lblSub.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $lblSub.Location = New-Object System.Drawing.Point(20, 42)
    $lblSub.Size = New-Object System.Drawing.Size(660, 20)
    $dlg.Controls.Add($lblSub)

    # Prerequisite list panel with scrollbar
    $pnlList = New-Object System.Windows.Forms.Panel
    $pnlList.Location = New-Object System.Drawing.Point(20, 70)
    $pnlList.Size = New-Object System.Drawing.Size(665, 390)
    $pnlList.AutoScroll = $true
    $pnlList.BackColor = [System.Drawing.Color]::White
    $pnlList.BorderStyle = "FixedSingle"
    $dlg.Controls.Add($pnlList)

    $yPos = 6
    $lastCat = ""
    foreach ($p in $prereqs) {
        # Category header
        if ($p.Category -ne $lastCat) {
            $lastCat = $p.Category
            $lblCat = New-Object System.Windows.Forms.Label
            $catLabel = switch ($p.Category) {
                "Core"             { "CORE REQUIREMENTS" }
                "Remoting"         { "WINRM & REMOTING" }
                "Modules"          { "POWERSHELL MODULES" }
                "Windows Features" { "WINDOWS FEATURES" }
                default            { $p.Category.ToUpper() }
            }
            $lblCat.Text = $catLabel
            $lblCat.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
            $lblCat.ForeColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
            $lblCat.Location = New-Object System.Drawing.Point(10, $yPos)
            $lblCat.Size = New-Object System.Drawing.Size(630, 16)
            $pnlList.Controls.Add($lblCat)
            $yPos += 20
        }

        # Status icon/color
        $isGood = ($p.Status -match "^OK")
        $isInfo = ($p.Name -match "Optional|TrustedHosts")
        $statusColor = if ($isGood) { [System.Drawing.Color]::FromArgb(39, 174, 96) }
                       elseif ($isInfo) { [System.Drawing.Color]::FromArgb(52, 152, 219) }
                       else { [System.Drawing.Color]::FromArgb(192, 57, 43) }
        $icon = if ($isGood) { "[PASS]" } elseif ($isInfo) { "[INFO]" } else { "[FAIL]" }

        $lblIcon = New-Object System.Windows.Forms.Label
        $lblIcon.Text = $icon
        $lblIcon.Font = New-Object System.Drawing.Font("Consolas", 8.5, [System.Drawing.FontStyle]::Bold)
        $lblIcon.ForeColor = $statusColor
        $lblIcon.Location = New-Object System.Drawing.Point(12, $yPos)
        $lblIcon.Size = New-Object System.Drawing.Size(48, 18)
        $pnlList.Controls.Add($lblIcon)

        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text = "$($p.Name)  —  $($p.Status)"
        $lblName.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $lblName.ForeColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
        $lblName.Location = New-Object System.Drawing.Point(62, $yPos)
        $lblName.Size = New-Object System.Drawing.Size(580, 18)
        $pnlList.Controls.Add($lblName)

        $lblDesc = New-Object System.Windows.Forms.Label
        $lblDesc.Text = $p.Description
        $lblDesc.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
        $lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
        $lblDesc.Location = New-Object System.Drawing.Point(62, ($yPos + 19))
        $lblDesc.Size = New-Object System.Drawing.Size(580, 16)
        $pnlList.Controls.Add($lblDesc)

        # Show fix action label for fixable items
        if ($p.CanFix -and $p.FixLabel -ne "") {
            $lblFix = New-Object System.Windows.Forms.Label
            $lblFix.Text = "Fix: $($p.FixLabel)"
            $lblFix.ForeColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
            $lblFix.Font = New-Object System.Drawing.Font("Consolas", 7.5)
            $lblFix.Location = New-Object System.Drawing.Point(62, ($yPos + 36))
            $lblFix.Size = New-Object System.Drawing.Size(580, 14)
            $pnlList.Controls.Add($lblFix)
            $yPos += 56
        } elseif (-not $isGood -and -not $isInfo -and $p.FixLabel -ne "") {
            # Show manual fix hint for non-fixable failed items
            $lblManual = New-Object System.Windows.Forms.Label
            $lblManual.Text = "Manual: $($p.FixLabel)"
            $lblManual.ForeColor = [System.Drawing.Color]::FromArgb(211, 84, 0)
            $lblManual.Font = New-Object System.Drawing.Font("Consolas", 7.5)
            $lblManual.Location = New-Object System.Drawing.Point(62, ($yPos + 36))
            $lblManual.Size = New-Object System.Drawing.Size(580, 14)
            $pnlList.Controls.Add($lblManual)
            $yPos += 56
        } else {
            $yPos += 42
        }
    }

    # -------------------------------------------------------
    # Buttons
    # -------------------------------------------------------
    $hasFixable = $fixCount -gt 0

    $btnFix = New-Object System.Windows.Forms.Button
    $btnFix.Text = "Fix All ($fixCount items)"
    $btnFix.Size = New-Object System.Drawing.Size(160, 36)
    $btnFix.Location = New-Object System.Drawing.Point(210, 475)
    $btnFix.FlatStyle = "Flat"
    $btnFix.FlatAppearance.BorderSize = 0
    $btnFix.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $btnFix.ForeColor = [System.Drawing.Color]::White
    $btnFix.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnFix.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnFix.Enabled = $hasFixable
    $dlg.Controls.Add($btnFix)

    $btnContinue = New-Object System.Windows.Forms.Button
    $btnContinue.Text = "Continue Anyway"
    $btnContinue.Size = New-Object System.Drawing.Size(140, 36)
    $btnContinue.Location = New-Object System.Drawing.Point(380, 475)
    $btnContinue.FlatStyle = "Flat"
    $btnContinue.FlatAppearance.BorderSize = 1
    $btnContinue.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(189, 195, 199)
    $btnContinue.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $btnContinue.Cursor = [System.Windows.Forms.Cursors]::Hand
    $dlg.Controls.Add($btnContinue)

    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Exit"
    $btnExit.Size = New-Object System.Drawing.Size(80, 36)
    $btnExit.Location = New-Object System.Drawing.Point(530, 475)
    $btnExit.FlatStyle = "Flat"
    $btnExit.FlatAppearance.BorderSize = 0
    $btnExit.BackColor = [System.Drawing.Color]::FromArgb(231, 76, 60)
    $btnExit.ForeColor = [System.Drawing.Color]::White
    $btnExit.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $btnExit.Cursor = [System.Windows.Forms.Cursors]::Hand
    $dlg.Controls.Add($btnExit)

    # Result label
    $lblResult = New-Object System.Windows.Forms.Label
    $lblResult.Text = ""
    $lblResult.Location = New-Object System.Drawing.Point(20, 520)
    $lblResult.Size = New-Object System.Drawing.Size(660, 36)
    $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
    $lblResult.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dlg.Controls.Add($lblResult)

    # Admin tip
    if (-not $isAdmin) {
        $lblAdmin = New-Object System.Windows.Forms.Label
        $lblAdmin.Text = "Tip: Right-click this .ps1 file > Run with PowerShell as Administrator to auto-fix most issues."
        $lblAdmin.ForeColor = [System.Drawing.Color]::FromArgb(211, 84, 0)
        $lblAdmin.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
        $lblAdmin.Location = New-Object System.Drawing.Point(20, 558)
        $lblAdmin.Size = New-Object System.Drawing.Size(660, 18)
        $dlg.Controls.Add($lblAdmin)
    }

    $script:prereqResult = "exit"

    # -------------------------------------------------------
    # FIX ALL — with Yes/No confirmation listing what will happen
    # -------------------------------------------------------
    $btnFix.Add_Click({
        # Build confirmation message listing all actions
        $confirmMsg = "The following changes will be made to this computer:`n`n"
        $idx = 1
        foreach ($p in $fixable) {
            $confirmMsg += "  $idx. $($p.Name)`n"
            $confirmMsg += "     Action: $($p.FixLabel)`n`n"
            $idx++
        }
        $confirmMsg += "Do you want to proceed?"

        $answer = [System.Windows.Forms.MessageBox]::Show(
            $confirmMsg,
            "Confirm Prerequisite Installation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(211, 84, 0)
            $lblResult.Text = "Fix cancelled. Click 'Continue Anyway' to launch with current config, or 'Exit'."
            return
        }

        # User confirmed — proceed with fixes
        $errors = @()
        $fixed = 0
        foreach ($p in $fixable) {
            try {
                & $p.FixAction
                $fixed++
            } catch {
                $errors += "$($p.Name): $($_.Exception.Message)"
            }
        }

        if ($errors.Count -eq 0) {
            $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
            $lblResult.Text = "All $fixed items fixed successfully. Click 'Continue Anyway' to launch SecurityChecker."
            $btnFix.Enabled = $false
            $btnFix.Text = "Fixed!"
        } else {
            $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(192, 57, 43)
            $partialMsg = "$fixed fixed, $($errors.Count) failed: " + ($errors -join " | ")
            if ($partialMsg.Length -gt 200) { $partialMsg = $partialMsg.Substring(0, 197) + "..." }
            $lblResult.Text = $partialMsg
        }
    })

    $btnContinue.Add_Click({
        $script:prereqResult = "continue"
        $dlg.Close()
    })

    $btnExit.Add_Click({
        $script:prereqResult = "exit"
        $dlg.Close()
    })

    [void]$dlg.ShowDialog()
    $dlg.Dispose()

    return ($script:prereqResult -eq "continue")
}

# Run the prerequisite check
$proceed = Show-PrerequisiteDialog
if (-not $proceed) { return }

# =======================
# CONFIG
# =======================
$script:AppVersion   = "2.6"
$script:AppTitle     = "SecurityChecker v$($script:AppVersion)"
$script:DomainSuffix = ""

# =======================
# COLOR PALETTE
# =======================
$script:Colors = @{
    Primary     = [System.Drawing.Color]::FromArgb(24, 42, 68)
    Secondary   = [System.Drawing.Color]::FromArgb(44, 62, 88)
    Accent      = [System.Drawing.Color]::FromArgb(52, 152, 219)
    Success     = [System.Drawing.Color]::FromArgb(46, 204, 113)
    Warning     = [System.Drawing.Color]::FromArgb(241, 196, 15)
    Danger      = [System.Drawing.Color]::FromArgb(231, 76, 60)
    TextLight   = [System.Drawing.Color]::White
    TextDark    = [System.Drawing.Color]::FromArgb(44, 62, 80)
    GridBg      = [System.Drawing.Color]::FromArgb(248, 249, 250)
    GridAltRow  = [System.Drawing.Color]::FromArgb(236, 240, 245)
    GridHeader  = [System.Drawing.Color]::FromArgb(44, 62, 88)
    InputBg     = [System.Drawing.Color]::White
    InputBorder = [System.Drawing.Color]::FromArgb(189, 195, 199)
    PanelBg     = [System.Drawing.Color]::FromArgb(236, 240, 241)
}

# =======================
# TOOLTIP PROVIDER
# =======================
$script:ToolTip = New-Object System.Windows.Forms.ToolTip
$script:ToolTip.AutoPopDelay = 8000; $script:ToolTip.InitialDelay = 400; $script:ToolTip.ReshowDelay = 200
$script:ToolTip.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50); $script:ToolTip.ForeColor = [System.Drawing.Color]::White

function Set-Tip { param([System.Windows.Forms.Control]$Control, [string]$Text); $script:ToolTip.SetToolTip($Control, $Text) }

# =======================
# HELPERS
# =======================
function Convert-CidrToIPs {
    param([Parameter(Mandatory)][string]$Cidr)
    $Cidr = ($Cidr + "").Trim()
    if ($Cidr -notmatch '^(\d{1,3}\.){3}\d{1,3}\/\d{1,2}$') { throw ("Invalid CIDR: '{0}'" -f $Cidr) }
    $parts = $Cidr.Split('/'); $ip = $parts[0].Trim(); $prefix = [int]$parts[1]
    if ($prefix -lt 0 -or $prefix -gt 32) { throw "Prefix must be 0-32" }
    if ($prefix -lt 16) { throw "Prefix /$prefix would generate too many IPs (max /16 = 65534 hosts). Use a smaller range." }
    $addr = [System.Net.IPAddress]::Parse($ip); $baseBytes = $addr.GetAddressBytes()
    if ($prefix -eq 32) { $ips = New-Object System.Collections.Generic.List[string]; $ips.Add($ip); return $ips }
    $hostBits = 32 - $prefix; $hostCount = [math]::Pow(2, $hostBits)
    $baseInt = 0; foreach ($b in $baseBytes) { $baseInt = ($baseInt * 256) + [int]$b }
    $ips = New-Object System.Collections.Generic.List[string]
    if ($prefix -eq 31) { for ($i=0;$i -lt 2;$i++) { $a=$baseInt+$i; $ips.Add(((@(($a -shr 24)-band 0xFF;($a -shr 16)-band 0xFF;($a -shr 8)-band 0xFF;$a -band 0xFF))-join '.')) } }
    elseif ($hostCount -gt 2) { for ($i=1;$i -lt ($hostCount-1);$i++) { $a=$baseInt+$i; $ips.Add(((@(($a -shr 24)-band 0xFF;($a -shr 16)-band 0xFF;($a -shr 8)-band 0xFF;$a -band 0xFF))-join '.')) } }
    return $ips
}

function Parse-TargetInput {
    param([string]$InputText)
    $allIPs = New-Object System.Collections.Generic.List[string]
    $entries = $InputText -split '[,;\r\n]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    foreach ($entry in $entries) {
        if ($entry -match '/') { $cidrIps = Convert-CidrToIPs -Cidr $entry; foreach ($ip in $cidrIps) { $allIPs.Add($ip) } }
        elseif ($entry -match '^(\d{1,3}\.){3}\d{1,3}$') { $allIPs.Add($entry) }
        elseif ($entry -match '^(\d{1,3}\.){3}\d{1,3}\s*-\s*(\d{1,3}\.){3}\d{1,3}$') {
            $rp = $entry -split '\s*-\s*'
            $sb=[System.Net.IPAddress]::Parse($rp[0]).GetAddressBytes(); $eb=[System.Net.IPAddress]::Parse($rp[1]).GetAddressBytes()
            $si=0;foreach($b in $sb){$si=($si*256)+[int]$b};$ei=0;foreach($b in $eb){$ei=($ei*256)+[int]$b}
            for($i=$si;$i -le $ei;$i++){$allIPs.Add(((@(($i -shr 24)-band 0xFF;($i -shr 16)-band 0xFF;($i -shr 8)-band 0xFF;$i -band 0xFF))-join '.'))}
        }
        else { throw "Bad format: '$entry'" }
    }
    $unique=New-Object System.Collections.Generic.HashSet[string]; $result=New-Object System.Collections.Generic.List[string]
    foreach($ip in $allIPs){if($unique.Add($ip)){$result.Add($ip)}}; return $result
}

# =======================
# CREDENTIAL DIALOG
# =======================
function Show-CredentialDialog {
    param([string]$DefaultUser="")
    $f=New-Object System.Windows.Forms.Form; $f.Text="$($script:AppTitle) — Credentials"; $f.Size=New-Object System.Drawing.Size(440,300)
    $f.StartPosition="CenterParent"; $f.FormBorderStyle="FixedDialog"; $f.MaximizeBox=$false; $f.MinimizeBox=$false
    $f.BackColor=$script:Colors.PanelBg; $f.Font=New-Object System.Drawing.Font("Segoe UI",9.5); $f.TopMost=$true

    $hdr=New-Object System.Windows.Forms.Label; $hdr.Text="Enter Domain Credentials"
    $hdr.Font=New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
    $hdr.ForeColor=$script:Colors.Secondary; $hdr.Location=New-Object System.Drawing.Point(20,15); $hdr.AutoSize=$true; $f.Controls.Add($hdr)

    $desc=New-Object System.Windows.Forms.Label; $desc.Text="Used for WinRM. Click 'Use Current User' to use session identity."
    $desc.ForeColor=[System.Drawing.Color]::FromArgb(100,100,100); $desc.Location=New-Object System.Drawing.Point(20,45); $desc.Size=New-Object System.Drawing.Size(390,32); $f.Controls.Add($desc)

    $lU=New-Object System.Windows.Forms.Label; $lU.Text="Username (DOMAIN\user):"; $lU.Location=New-Object System.Drawing.Point(20,85); $lU.AutoSize=$true; $f.Controls.Add($lU)
    $tU=New-Object System.Windows.Forms.TextBox; $tU.Location=New-Object System.Drawing.Point(20,108); $tU.Size=New-Object System.Drawing.Size(385,28); $tU.BackColor=$script:Colors.InputBg; $tU.Text=$DefaultUser; $f.Controls.Add($tU)
    $lP=New-Object System.Windows.Forms.Label; $lP.Text="Password:"; $lP.Location=New-Object System.Drawing.Point(20,145); $lP.AutoSize=$true; $f.Controls.Add($lP)
    $tP=New-Object System.Windows.Forms.TextBox; $tP.Location=New-Object System.Drawing.Point(20,168); $tP.Size=New-Object System.Drawing.Size(385,28); $tP.UseSystemPasswordChar=$true; $tP.BackColor=$script:Colors.InputBg; $f.Controls.Add($tP)

    $bOK=New-Object System.Windows.Forms.Button; $bOK.Text="OK"; $bOK.Size=New-Object System.Drawing.Size(90,34); $bOK.Location=New-Object System.Drawing.Point(195,220)
    $bOK.BackColor=$script:Colors.Accent; $bOK.ForeColor=$script:Colors.TextLight; $bOK.FlatStyle="Flat"; $bOK.FlatAppearance.BorderSize=0
    $bOK.DialogResult=[System.Windows.Forms.DialogResult]::OK; $f.Controls.Add($bOK)

    $bCx=New-Object System.Windows.Forms.Button; $bCx.Text="Use Current User"; $bCx.Size=New-Object System.Drawing.Size(120,34); $bCx.Location=New-Object System.Drawing.Point(295,220)
    $bCx.FlatStyle="Flat"; $bCx.FlatAppearance.BorderSize=1; $bCx.FlatAppearance.BorderColor=$script:Colors.InputBorder
    $bCx.DialogResult=[System.Windows.Forms.DialogResult]::Cancel; $f.Controls.Add($bCx)
    $f.AcceptButton=$bOK; $f.CancelButton=$bCx

    $result=$f.ShowDialog()
    if($result -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($tU.Text) -and -not [string]::IsNullOrWhiteSpace($tP.Text)){
        $secPass=ConvertTo-SecureString $tP.Text -AsPlainText -Force
        $cred=New-Object System.Management.Automation.PSCredential($tU.Text,$secPass); $tP.Text=""; $f.Dispose(); return $cred
    }
    $f.Dispose(); return $null
}

# =======================
# BUILD GUI
# =======================
[System.Windows.Forms.Application]::EnableVisualStyles()

$form=New-Object System.Windows.Forms.Form; $form.Text=$script:AppTitle; $form.Size=New-Object System.Drawing.Size(1600,920)
$form.StartPosition="CenterScreen"; $form.Font=New-Object System.Drawing.Font("Segoe UI",9); $form.BackColor=$script:Colors.PanelBg
$form.MinimumSize=New-Object System.Drawing.Size(1200,700)

# --- Title bar ---
$pnlTitle=New-Object System.Windows.Forms.Panel; $pnlTitle.Dock="Top"; $pnlTitle.Height=50; $pnlTitle.BackColor=$script:Colors.Primary
$form.Controls.Add($pnlTitle)
$lblTitle=New-Object System.Windows.Forms.Label; $lblTitle.Text="SecurityChecker"
$lblTitle.Font=New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor=$script:Colors.TextLight; $lblTitle.Location=New-Object System.Drawing.Point(16,12); $lblTitle.AutoSize=$true; $pnlTitle.Controls.Add($lblTitle)
$lblVer=New-Object System.Windows.Forms.Label; $lblVer.Text="v$($script:AppVersion)"; $lblVer.Font=New-Object System.Drawing.Font("Segoe UI",9)
$lblVer.ForeColor=[System.Drawing.Color]::FromArgb(150,180,210); $lblVer.Location=New-Object System.Drawing.Point(185,18); $lblVer.AutoSize=$true; $pnlTitle.Controls.Add($lblVer)
$lblSub=New-Object System.Windows.Forms.Label; $lblSub.Text="Enterprise Subnet Security Configuration Scanner"
$lblSub.Font=New-Object System.Drawing.Font("Segoe UI",9); $lblSub.ForeColor=[System.Drawing.Color]::FromArgb(130,160,190)
$lblSub.Location=New-Object System.Drawing.Point(240,18); $lblSub.AutoSize=$true; $pnlTitle.Controls.Add($lblSub)

# --- Settings panel ---
$pnlS=New-Object System.Windows.Forms.Panel; $pnlS.Location=New-Object System.Drawing.Point(0,50); $pnlS.Size=New-Object System.Drawing.Size(1600,120)
$pnlS.BackColor=[System.Drawing.Color]::White; $pnlS.Anchor="Top,Left,Right"; $form.Controls.Add($pnlS)

# Row 1
$lblTgt=New-Object System.Windows.Forms.Label; $lblTgt.Text="Target(s):"; $lblTgt.Location=New-Object System.Drawing.Point(16,14); $lblTgt.AutoSize=$true
$lblTgt.ForeColor=$script:Colors.TextDark; $lblTgt.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold); $pnlS.Controls.Add($lblTgt)
$txtTarget=New-Object System.Windows.Forms.TextBox; $txtTarget.Location=New-Object System.Drawing.Point(90,11); $txtTarget.Size=New-Object System.Drawing.Size(340,26)
$txtTarget.BackColor=$script:Colors.InputBg; $txtTarget.Text="192.168.1.0/24"; $txtTarget.Font=New-Object System.Drawing.Font("Consolas",9.5); $pnlS.Controls.Add($txtTarget)
Set-Tip $txtTarget "Enter targets: CIDR (10.0.0.0/24), single IP, range (10.0.0.1-10.0.0.50), or comma-separated"

function New-StyledButton{param([string]$Text,[int]$X,[int]$Y,[int]$W=110,[int]$H=30,[System.Drawing.Color]$Bg,[System.Drawing.Color]$Fg=[System.Drawing.Color]::White)
    $b=New-Object System.Windows.Forms.Button; $b.Text=$Text; $b.Location=New-Object System.Drawing.Point($X,$Y); $b.Size=New-Object System.Drawing.Size($W,$H)
    $b.FlatStyle="Flat"; $b.FlatAppearance.BorderSize=0; $b.BackColor=$Bg; $b.ForeColor=$Fg
    $b.Font=New-Object System.Drawing.Font("Segoe UI",9); $b.Cursor=[System.Windows.Forms.Cursors]::Hand; return $b}

$btnCred=New-StyledButton "Set Credentials" 450 9 130 30 $script:Colors.Secondary
$btnScan=New-StyledButton "Scan" 590 9 100 30 $script:Colors.Accent
$btnCancel=New-StyledButton "Cancel" 700 9 90 30 $script:Colors.Danger; $btnCancel.Enabled=$false
$btnExport=New-StyledButton "Export CSV" 800 9 110 30 $script:Colors.Success
$pnlS.Controls.AddRange(@($btnCred,$btnScan,$btnCancel,$btnExport))
Set-Tip $btnCred   "Set explicit domain credentials for WinRM.`nSkip if running as domain admin."
Set-Tip $btnScan   "Start scanning all targets. Results appear in real-time."
Set-Tip $btnCancel "Stop the running scan. Completed results are kept."
Set-Tip $btnExport "Export all results to a timestamped CSV file."

$currentIdentity=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$lblCred=New-Object System.Windows.Forms.Label; $lblCred.Text="Auth: $currentIdentity (current session)"
$lblCred.Location=New-Object System.Drawing.Point(920,16); $lblCred.AutoSize=$true; $lblCred.ForeColor=$script:Colors.Success
$lblCred.Font=New-Object System.Drawing.Font("Segoe UI",9); $pnlS.Controls.Add($lblCred)
Set-Tip $lblCred "Green = current session | Blue = explicit credentials"

# Row 2: Checks
$lblChk=New-Object System.Windows.Forms.Label; $lblChk.Text="Checks:"; $lblChk.Location=New-Object System.Drawing.Point(16,52); $lblChk.AutoSize=$true
$lblChk.ForeColor=$script:Colors.TextDark; $lblChk.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold); $pnlS.Controls.Add($lblChk)
$lnkAll=New-Object System.Windows.Forms.LinkLabel; $lnkAll.Text="All"; $lnkAll.Location=New-Object System.Drawing.Point(75,53); $lnkAll.AutoSize=$true; $pnlS.Controls.Add($lnkAll)
Set-Tip $lnkAll "Select all security checks"
$lblS1=New-Object System.Windows.Forms.Label; $lblS1.Text="/"; $lblS1.Location=New-Object System.Drawing.Point(97,53); $lblS1.AutoSize=$true; $lblS1.ForeColor=[System.Drawing.Color]::Gray; $pnlS.Controls.Add($lblS1)
$lnkNone=New-Object System.Windows.Forms.LinkLabel; $lnkNone.Text="None"; $lnkNone.Location=New-Object System.Drawing.Point(106,53); $lnkNone.AutoSize=$true; $pnlS.Controls.Add($lnkNone)
Set-Tip $lnkNone "Deselect all security checks"

$chkPing=New-Object System.Windows.Forms.CheckBox; $chkPing.Text="Ping First"; $chkPing.Checked=$true; $chkPing.Location=New-Object System.Drawing.Point(155,50); $chkPing.AutoSize=$true; $chkPing.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkPing)
Set-Tip $chkPing "ICMP ping before scanning. Disable if ICMP is blocked."
$chkWinRM=New-Object System.Windows.Forms.CheckBox; $chkWinRM.Text="WinRM (5985)"; $chkWinRM.Checked=$true; $chkWinRM.Location=New-Object System.Drawing.Point(265,50); $chkWinRM.AutoSize=$true; $chkWinRM.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkWinRM)
Set-Tip $chkWinRM "TCP 5985 pre-check. Skips unreachable hosts fast."
$chkRDP=New-Object System.Windows.Forms.CheckBox; $chkRDP.Text="RDP (3389)"; $chkRDP.Checked=$true; $chkRDP.Location=New-Object System.Drawing.Point(400,50); $chkRDP.AutoSize=$true; $chkRDP.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkRDP)
Set-Tip $chkRDP "Check if RDP port 3389 is open. Risk if NLA not enforced."

$lblDiv=New-Object System.Windows.Forms.Label; $lblDiv.Text="|"; $lblDiv.Location=New-Object System.Drawing.Point(515,52); $lblDiv.AutoSize=$true; $lblDiv.ForeColor=[System.Drawing.Color]::FromArgb(180,180,180); $pnlS.Controls.Add($lblDiv)

$chkLLMNR=New-Object System.Windows.Forms.CheckBox; $chkLLMNR.Text="LLMNR"; $chkLLMNR.Checked=$true; $chkLLMNR.Location=New-Object System.Drawing.Point(530,50); $chkLLMNR.AutoSize=$true; $chkLLMNR.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkLLMNR)
Set-Tip $chkLLMNR "LLMNR: vulnerable to MITM/relay. Should be Disabled."
$chkMDNS=New-Object System.Windows.Forms.CheckBox; $chkMDNS.Text="mDNS"; $chkMDNS.Checked=$true; $chkMDNS.Location=New-Object System.Drawing.Point(615,50); $chkMDNS.AutoSize=$true; $chkMDNS.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkMDNS)
Set-Tip $chkMDNS "Multicast DNS: similar risk to LLMNR. Should be Disabled."
$chkNetBIOS=New-Object System.Windows.Forms.CheckBox; $chkNetBIOS.Text="NetBIOS"; $chkNetBIOS.Checked=$true; $chkNetBIOS.Location=New-Object System.Drawing.Point(695,50); $chkNetBIOS.AutoSize=$true; $chkNetBIOS.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkNetBIOS)
Set-Tip $chkNetBIOS "NetBIOS over TCP/IP: spoofing risk. Should be Disabled."
$chkSMBv1=New-Object System.Windows.Forms.CheckBox; $chkSMBv1.Text="SMBv1"; $chkSMBv1.Checked=$true; $chkSMBv1.Location=New-Object System.Drawing.Point(790,50); $chkSMBv1.AutoSize=$true; $chkSMBv1.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkSMBv1)
Set-Tip $chkSMBv1 "SMBv1: EternalBlue exploit target. Should be Disabled."
$chkSMBSign=New-Object System.Windows.Forms.CheckBox; $chkSMBSign.Text="SMB Signing"; $chkSMBSign.Checked=$true; $chkSMBSign.Location=New-Object System.Drawing.Point(878,50); $chkSMBSign.AutoSize=$true; $chkSMBSign.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkSMBSign)
Set-Tip $chkSMBSign "SMB signing: prevents relay attacks. Should be Required."
$chkWPAD=New-Object System.Windows.Forms.CheckBox; $chkWPAD.Text="WPAD"; $chkWPAD.Checked=$false; $chkWPAD.Location=New-Object System.Drawing.Point(1000,50); $chkWPAD.AutoSize=$true; $chkWPAD.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkWPAD)
Set-Tip $chkWPAD "WPAD: can redirect traffic to attacker proxy."
$chkIPv6=New-Object System.Windows.Forms.CheckBox; $chkIPv6.Text="IPv6"; $chkIPv6.Checked=$false; $chkIPv6.Location=New-Object System.Drawing.Point(1080,50); $chkIPv6.AutoSize=$true; $chkIPv6.ForeColor=$script:Colors.TextDark; $pnlS.Controls.Add($chkIPv6)
Set-Tip $chkIPv6 "IPv6: MITM risk if unmanaged. Disable if not needed."

$script:secCheckBoxes=@($chkLLMNR,$chkMDNS,$chkNetBIOS,$chkSMBv1,$chkSMBSign,$chkWPAD,$chkIPv6)
$lnkAll.Add_LinkClicked({ foreach($c in $script:secCheckBoxes){$c.Checked=$true} })
$lnkNone.Add_LinkClicked({ foreach($c in $script:secCheckBoxes){$c.Checked=$false} })

# Row 3
$lblTh=New-Object System.Windows.Forms.Label; $lblTh.Text="Threads:"; $lblTh.Location=New-Object System.Drawing.Point(16,88); $lblTh.AutoSize=$true
$lblTh.ForeColor=$script:Colors.TextDark; $lblTh.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold); $pnlS.Controls.Add($lblTh)
$numThreads=New-Object System.Windows.Forms.NumericUpDown; $numThreads.Location=New-Object System.Drawing.Point(80,85); $numThreads.Size=New-Object System.Drawing.Size(60,26)
$numThreads.Minimum=5; $numThreads.Maximum=100; $numThreads.Value=25; $numThreads.BackColor=$script:Colors.InputBg; $pnlS.Controls.Add($numThreads)
Set-Tip $numThreads "Parallel threads (5-100). 25 is good for /24."

$lblAu=New-Object System.Windows.Forms.Label; $lblAu.Text="Auth:"; $lblAu.Location=New-Object System.Drawing.Point(160,88); $lblAu.AutoSize=$true
$lblAu.ForeColor=$script:Colors.TextDark; $lblAu.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold); $pnlS.Controls.Add($lblAu)
$cmbAuth=New-Object System.Windows.Forms.ComboBox; $cmbAuth.Items.AddRange(@("Kerberos","Negotiate","Default"))
$cmbAuth.SelectedIndex=0; $cmbAuth.Location=New-Object System.Drawing.Point(200,85); $cmbAuth.Size=New-Object System.Drawing.Size(120,26)
$cmbAuth.DropDownStyle="DropDownList"; $cmbAuth.BackColor=$script:Colors.InputBg; $pnlS.Controls.Add($cmbAuth)
Set-Tip $cmbAuth "Kerberos = secure (needs FQDN)`nNegotiate = NTLM fallback`nDefault = auto"

$lblDm=New-Object System.Windows.Forms.Label; $lblDm.Text="Domain Suffix:"; $lblDm.Location=New-Object System.Drawing.Point(340,88); $lblDm.AutoSize=$true
$lblDm.ForeColor=$script:Colors.TextDark; $lblDm.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold); $pnlS.Controls.Add($lblDm)
$txtDomain=New-Object System.Windows.Forms.TextBox; $txtDomain.Location=New-Object System.Drawing.Point(440,85); $txtDomain.Size=New-Object System.Drawing.Size(160,26)
$txtDomain.Text=$script:DomainSuffix; $txtDomain.Font=New-Object System.Drawing.Font("Consolas",9.5); $txtDomain.BackColor=$script:Colors.InputBg; $pnlS.Controls.Add($txtDomain)
Set-Tip $txtDomain "Appended to short hostnames for Kerberos auth."

# ===================================================================
# DataGridView — FIXED COLUMN WIDTHS, HORIZONTAL SCROLL, NO FILL MODE
# ===================================================================
$grid=New-Object System.Windows.Forms.DataGridView
$grid.Location=New-Object System.Drawing.Point(0,173); $grid.Size=New-Object System.Drawing.Size(1584,610); $grid.Anchor="Top,Bottom,Left,Right"
$grid.ReadOnly=$true; $grid.AllowUserToAddRows=$false; $grid.AutoGenerateColumns=$false
$grid.SelectionMode="FullRowSelect"; $grid.MultiSelect=$true; $grid.RowHeadersVisible=$false
$grid.BorderStyle="None"; $grid.CellBorderStyle="SingleHorizontal"; $grid.BackgroundColor=$script:Colors.GridBg
$grid.DefaultCellStyle.BackColor=$script:Colors.GridBg; $grid.DefaultCellStyle.SelectionBackColor=[System.Drawing.Color]::FromArgb(200,220,240)
$grid.DefaultCellStyle.SelectionForeColor=$script:Colors.TextDark; $grid.DefaultCellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9)
$grid.DefaultCellStyle.Padding=New-Object System.Windows.Forms.Padding(4,2,4,2)
$grid.AlternatingRowsDefaultCellStyle.BackColor=$script:Colors.GridAltRow
$grid.ColumnHeadersDefaultCellStyle.BackColor=$script:Colors.GridHeader; $grid.ColumnHeadersDefaultCellStyle.ForeColor=$script:Colors.TextLight
$grid.ColumnHeadersDefaultCellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle.Padding=New-Object System.Windows.Forms.Padding(4,6,4,6)
$grid.ColumnHeadersHeightSizeMode="AutoSize"; $grid.EnableHeadersVisualStyles=$false; $grid.RowTemplate.Height=26; $grid.ShowCellToolTips=$true

# KEY FIX: Fixed widths with last column (Error) filling remaining space
$grid.AutoSizeColumnsMode = "None"
$grid.ScrollBars = "Both"

$form.Controls.Add($grid)

$dt=New-Object System.Data.DataTable
$columns=@("IP","Ping","WinRM","RDP","Target","Auth","Status","Computer","OS","LLMNR","mDNS","NetBIOS","SMBv1","SMBSign","WPAD","IPv6","Error")
foreach($c in $columns){[void]$dt.Columns.Add($c)}
$grid.DataSource=$dt

# Fixed-width column helper — each column gets EXACTLY the width specified
function Add-GridCol($name, $header, $width, $tip="") {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.DataPropertyName = $name
    $col.HeaderText       = $header
    $col.Width            = $width
    $col.MinimumWidth     = $width
    $col.Resizable        = [System.Windows.Forms.DataGridViewTriState]::True
    if ($tip -ne "") { $col.ToolTipText = $tip }
    $grid.Columns.Add($col) | Out-Null
}

# Fixed pixel widths — guaranteed no truncation
Add-GridCol "IP"       "IP"         115  "Target IP address"
Add-GridCol "Ping"     "Ping"        52  "ICMP ping: Up/Down"
Add-GridCol "WinRM"    "WinRM"       60  "TCP 5985 status"
Add-GridCol "RDP"      "RDP"         58  "TCP 3389 status"
Add-GridCol "Target"   "Target"     200  "Resolved FQDN or IP"
Add-GridCol "Auth"     "Auth"        72  "Auth method used"
Add-GridCol "Status"   "Status"      68  "OK, Error, or Skipped"
Add-GridCol "Computer" "Computer"   115  "Remote COMPUTERNAME"
Add-GridCol "OS"       "OS"         210  "Operating system"
Add-GridCol "LLMNR"    "LLMNR"       82  "Disabled = Secure | NotSet = Risk"
Add-GridCol "mDNS"     "mDNS"        82  "Disabled = Secure | NotSet = Risk"
Add-GridCol "NetBIOS"  "NetBIOS"    110  "Disabled = Secure | DHCP(Def) = Risk (uses DHCP setting) | Enabled = Risk"
Add-GridCol "SMBv1"    "SMBv1"       82  "Disabled = Secure | Enabled = Critical"
Add-GridCol "SMBSign"  "SMB Sign"    82  "Required = Secure | NotReq = Risk"
Add-GridCol "WPAD"     "WPAD"        72  "Stopped = Secure | Running = Risk"
Add-GridCol "IPv6"     "IPv6"        72  "Disabled = Secure (if not needed)"
Add-GridCol "Error"    "Error"      250  "Error details"

# Make the Error column (last) fill remaining space when window is wide
$errorCol = $grid.Columns[$grid.Columns.Count - 1]
$errorCol.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$errorCol.MinimumWidth = 250
$errorCol.FillWeight = 100

# Proportionally resize all columns when window expands beyond default
$script:baseFormWidth = 1600
$script:baseColWidths = @{}
foreach ($col in $grid.Columns) {
    if ($col.DataPropertyName -ne "Error") {
        $script:baseColWidths[$col.DataPropertyName] = $col.Width
    }
}

$form.Add_Resize({
    $extraWidth = $form.ClientSize.Width - $script:baseFormWidth
    if ($extraWidth -gt 50) {
        # Distribute extra width proportionally to wider columns
        $expandCols = @("Target","OS","Computer","Error")
        foreach ($col in $grid.Columns) {
            $name = $col.DataPropertyName
            if ($name -in $expandCols -and $script:baseColWidths.ContainsKey($name)) {
                $bonus = [int]($extraWidth * 0.15)
                $col.Width = $script:baseColWidths[$name] + $bonus
            }
        }
    } elseif ($extraWidth -le 0) {
        # Reset to base widths when shrunk
        foreach ($col in $grid.Columns) {
            $name = $col.DataPropertyName
            if ($script:baseColWidths.ContainsKey($name)) {
                $col.Width = $script:baseColWidths[$name]
            }
        }
    }
})

# Color-code cells
$grid.Add_CellFormatting({
    param($sender,$e)
    if($e.RowIndex -lt 0){return}
    $colName=$grid.Columns[$e.ColumnIndex].DataPropertyName; $val="$($e.Value)"
    if([string]::IsNullOrEmpty($val)){return}
    $green=[System.Drawing.Color]::FromArgb(39,174,96); $red=[System.Drawing.Color]::FromArgb(192,57,43)
    $orange=[System.Drawing.Color]::FromArgb(211,84,0); $bold=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)

    if($colName -in @("LLMNR","mDNS","NetBIOS","SMBv1","WPAD","IPv6")){
        if($val -match '^Disabled'){$e.CellStyle.ForeColor=$green; $e.CellStyle.Font=$bold}
        elseif($val -match 'Enabled|NotSet|DHCP|Running|Default'){$e.CellStyle.ForeColor=$red; $e.CellStyle.Font=$bold}
    }
    if($colName -eq "SMBSign"){
        if($val -match '^Required'){$e.CellStyle.ForeColor=$green; $e.CellStyle.Font=$bold}
        elseif($val -match 'Not|False'){$e.CellStyle.ForeColor=$red; $e.CellStyle.Font=$bold}
    }
    if($colName -in @("Ping","WinRM","RDP")){
        if($val -eq "Up" -or $val -eq "Open"){$e.CellStyle.ForeColor=$green; $e.CellStyle.Font=$bold}
        elseif($val -eq "Down" -or $val -eq "Closed"){$e.CellStyle.ForeColor=$red; $e.CellStyle.Font=$bold}
    }
    if($colName -eq "Status"){
        if($val -eq "OK"){$e.CellStyle.ForeColor=$green; $e.CellStyle.Font=$bold}
        elseif($val -eq "Error"){$e.CellStyle.ForeColor=$red; $e.CellStyle.Font=$bold}
        elseif($val -match 'Skip'){$e.CellStyle.ForeColor=$orange}
    }
})

# ===================================================================
# REMOTE SCRIPT BLOCK FOR RESCAN (must be defined before menu handlers)
# ===================================================================
$RemoteScriptForRescan = {
    param($Options)
    function RegVal($p,$n){try{(Get-ItemProperty $p -ErrorAction Stop).$n}catch{$null}}
    $r=[ordered]@{}
    $r.Computer=$env:COMPUTERNAME
    try{$r.OS=(Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption}catch{$r.OS="Unknown"}
    if($Options.LLMNR){
        $v=RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' 'EnableMulticast'
        if($v -eq 0){$r.LLMNR='Disabled'}elseif($v -eq 1){$r.LLMNR='Enabled'}else{$r.LLMNR='NotSet'}
    }
    if($Options.mDNS){
        $v=RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' 'EnableMDNS'
        if($v -eq 0){$r.mDNS='Disabled'}elseif($v -eq 1){$r.mDNS='Enabled'}else{$r.mDNS='NotSet'}
    }
    if($Options.NetBIOS){
        try{
            $adapters=Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" -ErrorAction Stop
            $st=@(); foreach($a in $adapters){$v=$a.TcpipNetbiosOptions; switch($v){0{$st+="DHCP(Def)"};1{$st+="Enabled"};2{$st+="Disabled"};default{$st+="Unknown($v)"}}}
            $u=@($st|Select-Object -Unique); if($u.Count -eq 0){$r.NetBIOS="N/A"}elseif($u.Count -eq 1){$r.NetBIOS=$u[0]}else{$r.NetBIOS=($u -join '/')}
        }catch{$r.NetBIOS="Error"}
    }
    if($Options.SMBv1){
        try{$smb=Get-SmbServerConfiguration -ErrorAction Stop
            if($smb.EnableSMB1Protocol -eq $true){$r.SMBv1='Enabled'}elseif($smb.EnableSMB1Protocol -eq $false){$r.SMBv1='Disabled'}else{$r.SMBv1='Unknown'}
        }catch{try{$v=RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'SMB1'
            if($v -eq 0){$r.SMBv1='Disabled'}elseif($v -eq 1){$r.SMBv1='Enabled'}else{$r.SMBv1='NotSet'}}catch{$r.SMBv1='Error'}}
    }
    if($Options.SMBSign){
        try{$smb=Get-SmbServerConfiguration -ErrorAction Stop
            if($smb.RequireSecuritySignature -eq $true){$r.SMBSign='Required'}else{$r.SMBSign='NotReq'}
        }catch{try{$v=RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'RequireSecuritySignature'
            if($v -eq 1){$r.SMBSign='Required'}elseif($v -eq 0){$r.SMBSign='NotReq'}else{$r.SMBSign='NotSet'}}catch{$r.SMBSign='Error'}}
    }
    if($Options.WPAD){
        try{$svc=Get-Service -Name "WinHttpAutoProxySvc" -ErrorAction SilentlyContinue
            if($svc -and $svc.Status -eq 'Running'){$r.WPAD='Running'}elseif($svc){$r.WPAD='Stopped'}else{$r.WPAD='N/A'}
        }catch{$r.WPAD='Error'}
    }
    if($Options.IPv6){
        try{$b=Get-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue; $en=$b|Where-Object{$_.Enabled -eq $true}
            if($en){$r.IPv6='Enabled'}else{$r.IPv6='Disabled'}
        }catch{try{$v=RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' 'DisabledComponents'
            if($v -eq 0xFF){$r.IPv6='Disabled'}elseif($null -eq $v){$r.IPv6='Enabled'}else{$r.IPv6='Partial'}}catch{$r.IPv6='Error'}}
    }
    [pscustomobject]$r
}

# ===================================================================
# RIGHT-CLICK CONTEXT MENU — Enable WinRM on failed hosts
# ===================================================================
$gridMenu = New-Object System.Windows.Forms.ContextMenuStrip
$gridMenu.BackColor = [System.Drawing.Color]::White
$gridMenu.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$menuEnableWinRM = New-Object System.Windows.Forms.ToolStripMenuItem
$menuEnableWinRM.Text = "Enable WinRM on this host (WMI/SMB)"
$menuEnableWinRM.ForeColor = [System.Drawing.Color]::FromArgb(44, 62, 80)

$menuRescan = New-Object System.Windows.Forms.ToolStripMenuItem
$menuRescan.Text = "Rescan this host"
$menuRescan.ForeColor = [System.Drawing.Color]::FromArgb(44, 62, 80)

$menuCopyIP = New-Object System.Windows.Forms.ToolStripMenuItem
$menuCopyIP.Text = "Copy IP address"
$menuCopyIP.ForeColor = [System.Drawing.Color]::FromArgb(44, 62, 80)

$menuCopyError = New-Object System.Windows.Forms.ToolStripMenuItem
$menuCopyError.Text = "Copy error message"
$menuCopyError.ForeColor = [System.Drawing.Color]::FromArgb(44, 62, 80)

[void]$gridMenu.Items.Add($menuEnableWinRM)
[void]$gridMenu.Items.Add($menuRescan)
$sep = New-Object System.Windows.Forms.ToolStripSeparator
[void]$gridMenu.Items.Add($sep)
[void]$gridMenu.Items.Add($menuCopyIP)
[void]$gridMenu.Items.Add($menuCopyError)

## Do NOT bind context menu to grid — we show it manually to control row selection
# $grid.ContextMenuStrip = $gridMenu

# Track which DataTable row index was right-clicked
$script:rightClickDtRowIndex = -1

# On right-click: find the row via HitTest, map to DataTable row, then show menu
$grid.Add_MouseDown({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $hit = $grid.HitTest($e.X, $e.Y)
        if ($hit.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::Cell -and $hit.RowIndex -ge 0) {
            # Map grid display row to underlying DataTable row
            $boundItem = $grid.Rows[$hit.RowIndex].DataBoundItem
            if ($boundItem -is [System.Data.DataRowView]) {
                $tableRow = $boundItem.Row
                $script:rightClickDtRowIndex = $dt.Rows.IndexOf($tableRow)
            } else {
                $script:rightClickDtRowIndex = $hit.RowIndex
            }
            $grid.ClearSelection()
            $grid.Rows[$hit.RowIndex].Selected = $true
            $gridMenu.Show($grid, (New-Object System.Drawing.Point($e.X, $e.Y)))
        }
    }
})

# Enable WinRM remotely using multiple methods — runs in BACKGROUND to keep GUI responsive
# Methods: 1) WMI/RPC  2) WMI quickconfig/RPC  3) sc.exe remote/SMB  4) Scheduled Task/SMB
$menuEnableWinRM.Add_Click({
    $ridx = $script:rightClickDtRowIndex
    if ($ridx -lt 0 -or $ridx -ge $dt.Rows.Count) {
        [System.Windows.Forms.MessageBox]::Show("No row selected. Right-click on a row first.", $script:AppTitle, "OK", "Warning") | Out-Null
        return
    }
    $ip = [string]$dt.Rows[$ridx]["IP"]
    $target = [string]$dt.Rows[$ridx]["Target"]
    if ([string]::IsNullOrWhiteSpace($ip)) {
        [System.Windows.Forms.MessageBox]::Show("Selected row has no IP address.", $script:AppTitle, "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($target)) { $target = $ip }

    $confirmMsg = "This will attempt to enable WinRM on:`n`n"
    $confirmMsg += "  IP: $ip`n  Target: $target`n`n"
    $confirmMsg += "Methods tried (in order):`n"
    $confirmMsg += "  1. WMI via RPC (TCP 135)`n"
    $confirmMsg += "  2. WMI quickconfig via RPC`n"
    $confirmMsg += "  3. sc.exe remote via SMB (TCP 445) — if RPC is blocked`n"
    $confirmMsg += "  4. Scheduled Task via SMB — last resort`n`n"
    $confirmMsg += "After sending the command:`n"
    $confirmMsg += "  - Waits 15 seconds for service to start`n"
    $confirmMsg += "  - Pings host to verify it's still up`n"
    $confirmMsg += "  - Checks if WinRM port 5985 opened`n`n"
    $confirmMsg += "This does NOT restart the host.`n"
    $confirmMsg += "Proceed?"

    $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "$($script:AppTitle) - Enable WinRM", "YesNo", "Question")
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $stLabel.Text = "Enabling WinRM on $ip — pinging first..."
    [System.Windows.Forms.Application]::DoEvents()

    # Run in background job so GUI stays responsive
    $cred = $script:Cred
    $script:winrmJob = Start-Job -ScriptBlock {
        param($ip, $cred)
        $result = @{ Success = $false; Message = ""; PingBefore = $false; PingAfter = $false; Method = "" }

        # Step 1: Ping before
        try {
            $result.PingBefore = [bool](Test-Connection -ComputerName $ip -Count 2 -Quiet -ErrorAction SilentlyContinue)
        } catch { $result.PingBefore = $false }

        if (-not $result.PingBefore) {
            $result.Message = "Host $ip is NOT responding to ping. Host may be offline."
            return $result
        }

        $commandSent = $false

        # --- Method 1: CIM via RPC ---
        try {
            $cimSessionOpt = New-CimSessionOption -Protocol Dcom
            $cimParams = @{ ErrorAction = "Stop" }
            if ($null -ne $cred) { $cimParams["Credential"] = $cred }
            $cimSession = New-CimSession -ComputerName $ip -SessionOption $cimSessionOpt @cimParams
            $wmiResult = Invoke-CimMethod -CimSession $cimSession -ClassName "Win32_Process" -MethodName "Create" -Arguments @{ CommandLine = "cmd.exe /c `"sc config WinRM start= auto & net start WinRM`"" } -ErrorAction Stop
            Remove-CimSession $cimSession -ErrorAction SilentlyContinue
            if ($wmiResult.ReturnValue -eq 0) {
                $result.Method = "CIM (RPC)"
                $result.Message = "[Method 1 - CIM/RPC] Command sent. PID: $($wmiResult.ProcessId)"
                $commandSent = $true
            } else {
                $result.Message = "[Method 1 - CIM/RPC] Return code: $($wmiResult.ReturnValue)"
            }
        } catch {
            $result.Message = "[Method 1 - CIM/RPC] $($_.Exception.Message)"
        }

        # --- Method 2: CIM quickconfig via RPC ---
        if (-not $commandSent) {
            try {
                $cimSessionOpt2 = New-CimSessionOption -Protocol Dcom
                $cimParams2 = @{ ErrorAction = "Stop" }
                if ($null -ne $cred) { $cimParams2["Credential"] = $cred }
                $cimSession2 = New-CimSession -ComputerName $ip -SessionOption $cimSessionOpt2 @cimParams2
                $wmiResult2 = Invoke-CimMethod -CimSession $cimSession2 -ClassName "Win32_Process" -MethodName "Create" -Arguments @{ CommandLine = "cmd.exe /c `"winrm quickconfig /quiet /force`"" } -ErrorAction Stop
                Remove-CimSession $cimSession2 -ErrorAction SilentlyContinue
                if ($wmiResult2.ReturnValue -eq 0) {
                    $result.Method = "CIM quickconfig (RPC)"
                    $result.Message += "`n[Method 2 - CIM quickconfig] Command sent. PID: $($wmiResult2.ProcessId)"
                    $commandSent = $true
                } else {
                    $result.Message += "`n[Method 2 - CIM quickconfig] Return code: $($wmiResult2.ReturnValue)"
                }
            } catch {
                $result.Message += "`n[Method 2 - CIM quickconfig] $($_.Exception.Message)"
            }
        }

        # --- Method 3: Native SMB remote commands (TCP 445) — built-in, no external tools ---
        if (-not $commandSent) {
            $result.Message += "`n`nRPC unavailable. Trying native SMB methods (TCP 445)..."

            # Check if SMB port 445 is reachable
            $smbOpen = $false
            try {
                $tcp = New-Object System.Net.Sockets.TcpClient
                $ar = $tcp.BeginConnect($ip, 445, $null, $null)
                $smbOpen = $ar.AsyncWaitHandle.WaitOne(3000, $false)
                if ($smbOpen -and $tcp.Connected) { $tcp.Close() } else { $tcp.Close(); $smbOpen = $false }
            } catch { $smbOpen = $false }

            if (-not $smbOpen) {
                $result.Message += "`n[Method 3 - SMB] Port 445 is CLOSED."
                $result.Message += "`n`nBoth RPC (135) and SMB (445) are blocked."
                $result.Message += "`nMust fix manually via RDP or physical access."
                return $result
            }

            # Method 3a: sc.exe \\remote — configure and start WinRM service directly over SMB
            try {
                # Configure WinRM to auto-start
                $scConfig = & sc.exe \\$ip config WinRM start= auto 2>&1
                $configOk = ($LASTEXITCODE -eq 0)
                
                # Start the WinRM service
                $scStart = & sc.exe \\$ip start WinRM 2>&1
                $startOk = ($LASTEXITCODE -eq 0) -or ("$scStart" -match "already been started")

                if ($configOk -or $startOk) {
                    $result.Method = "sc.exe remote (SMB)"
                    $result.Message += "`n[Method 3a - sc.exe \\\\$ip] "
                    if ($configOk) { $result.Message += "Service configured to auto-start. " }
                    if ($startOk) { $result.Message += "Service start command sent." }
                    $commandSent = $true
                } else {
                    $scErr = "$scConfig $scStart".Trim()
                    if ($scErr.Length -gt 200) { $scErr = $scErr.Substring(0, 200) }
                    $result.Message += "`n[Method 3a - sc.exe \\\\$ip] Failed: $scErr"
                }
            } catch {
                $result.Message += "`n[Method 3a - sc.exe \\\\$ip] Exception: $($_.Exception.Message)"
            }

            # Method 3b: Remote scheduled task — create a one-time task to enable WinRM+PSRemoting
            if (-not $commandSent) {
                try {
                    $taskName = "SecurityChecker_EnableWinRM_$(Get-Random -Minimum 1000 -Maximum 9999)"
                    $taskCmd = "cmd.exe /c `"sc config WinRM start= auto & net start WinRM & winrm quickconfig /quiet /force & schtasks /Delete /TN $taskName /F`""

                    $schtaskArgs = @("/Create", "/S", $ip, "/TN", $taskName, "/TR", $taskCmd, "/SC", "ONCE", "/ST", "00:00", "/RL", "HIGHEST", "/RU", "SYSTEM", "/F")
                    if ($null -ne $cred) {
                        $netCred = $cred.GetNetworkCredential()
                        # NOTE: schtasks.exe requires plaintext credentials via command-line args.
                        # This is a limitation of the schtasks.exe CLI — credentials are briefly
                        # visible in the process argument list. Only used as last-resort Method 3b.
                        $schtaskArgs += @("/U", "$($netCred.Domain)\$($netCred.UserName)", "/P", $netCred.Password)
                    }
                    
                    $createResult = & schtasks.exe @schtaskArgs 2>&1
                    $createOk = ($LASTEXITCODE -eq 0)

                    if ($createOk) {
                        # Run the task immediately
                        $runArgs = @("/Run", "/S", $ip, "/TN", $taskName)
                        if ($null -ne $cred) {
                            $netCred = $cred.GetNetworkCredential()
                            # See note above re: schtasks.exe plaintext credential limitation
                            $runArgs += @("/U", "$($netCred.Domain)\$($netCred.UserName)", "/P", $netCred.Password)
                        }
                        $runResult = & schtasks.exe @runArgs 2>&1
                        $runOk = ($LASTEXITCODE -eq 0)

                        if ($runOk) {
                            $result.Method = "Scheduled Task (SMB)"
                            $result.Message += "`n[Method 3b - schtasks] Task '$taskName' created and executed."
                            $result.Message += "`nThe task enables WinRM and deletes itself when done."
                            $commandSent = $true
                        } else {
                            $result.Message += "`n[Method 3b - schtasks] Task created but failed to run: $runResult"
                            # Clean up the task
                            try { & schtasks.exe /Delete /S $ip /TN $taskName /F 2>&1 | Out-Null } catch {}
                        }
                    } else {
                        $createErr = "$createResult".Trim()
                        if ($createErr.Length -gt 200) { $createErr = $createErr.Substring(0, 200) }
                        $result.Message += "`n[Method 3b - schtasks] Failed to create task: $createErr"
                    }
                } catch {
                    $result.Message += "`n[Method 3b - schtasks] Exception: $($_.Exception.Message)"
                }
            }
        }

        if (-not $commandSent) {
            return $result
        }

        # --- Post-command verification ---
        Start-Sleep -Seconds 15

        # Ping after
        try {
            $result.PingAfter = [bool](Test-Connection -ComputerName $ip -Count 2 -Quiet -ErrorAction SilentlyContinue)
        } catch { $result.PingAfter = $false }

        # Check WinRM port
        $winrmOpen = $false
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $ar = $tcp.BeginConnect($ip, 5985, $null, $null)
            $winrmOpen = $ar.AsyncWaitHandle.WaitOne(3000, $false)
            if ($winrmOpen -and $tcp.Connected) { $tcp.Close() } else { $tcp.Close(); $winrmOpen = $false }
        } catch { $winrmOpen = $false }

        if ($result.PingAfter -and $winrmOpen) {
            $result.Success = $true
            $result.Message += "`n`nHost is UP and WinRM port 5985 is now OPEN.`nRight-click > Rescan to verify."
        } elseif ($result.PingAfter -and -not $winrmOpen) {
            $result.Message += "`n`nHost is UP but WinRM port 5985 still CLOSED.`nService may need more time or firewall blocking 5985."
        } elseif (-not $result.PingAfter) {
            $result.Message += "`n`nWARNING: Host not responding to ping after command."
        }

        return $result
    } -ArgumentList $ip, $cred

    # Timer to poll the background job
    $script:winrmTimer = New-Object System.Windows.Forms.Timer
    $script:winrmTimer.Interval = 1000
    $script:winrmJobIP = $ip
    $script:winrmJobStart = [DateTime]::Now
    $script:winrmTimer.Add_Tick({
        $elapsed = [int]([DateTime]::Now - $script:winrmJobStart).TotalSeconds
        if ($script:winrmJob.State -eq "Completed") {
            $script:winrmTimer.Stop()
            $script:winrmTimer.Dispose()
            try {
                $jobResult = Receive-Job -Job $script:winrmJob
                Remove-Job -Job $script:winrmJob -Force -ErrorAction SilentlyContinue

                $methodUsed = if ($jobResult.Method) { " via $($jobResult.Method)" } else { "" }
                $pingInfo = "Ping before: $(if($jobResult.PingBefore){'UP'}else{'DOWN'}) | Ping after: $(if($jobResult.PingAfter){'UP'}else{'DOWN'})"

                if ($jobResult.Success) {
                    $stLabel.Text = "WinRM enabled on $($script:winrmJobIP)$methodUsed. Ready to rescan."
                    [System.Windows.Forms.MessageBox]::Show(
                        "Enable WinRM on $($script:winrmJobIP):`n`n$($jobResult.Message)`n`n$pingInfo",
                        "$($script:AppTitle) - Success", "OK", "Information") | Out-Null
                } else {
                    $stLabel.Text = "WinRM enable failed on $($script:winrmJobIP). $pingInfo"
                    $failMsg = "Enable WinRM on $($script:winrmJobIP):`n`n$($jobResult.Message)`n`n$pingInfo"
                    [System.Windows.Forms.MessageBox]::Show($failMsg, "$($script:AppTitle) - Result", "OK", "Warning") | Out-Null
                }
            } catch {
                $stLabel.Text = "WinRM job error: $($_.Exception.Message)"
            }
        } elseif ($script:winrmJob.State -eq "Failed") {
            $script:winrmTimer.Stop()
            $script:winrmTimer.Dispose()
            $stLabel.Text = "WinRM background job failed on $($script:winrmJobIP)."
            try { Remove-Job -Job $script:winrmJob -Force -ErrorAction SilentlyContinue } catch {}
            [System.Windows.Forms.MessageBox]::Show("Background job failed unexpectedly.",
                "$($script:AppTitle) - Error", "OK", "Warning") | Out-Null
        } else {
            # Still running — show progress with countdown
            if ($elapsed -lt 5) {
                $stLabel.Text = "Enabling WinRM on $($script:winrmJobIP) — trying WMI/RPC... ($elapsed`s)"
            } elseif ($elapsed -lt 15) {
                $stLabel.Text = "Enabling WinRM on $($script:winrmJobIP) — trying methods... ($elapsed`s)"
            } elseif ($elapsed -lt 35) {
                $remaining = 35 - $elapsed
                $stLabel.Text = "Enabling WinRM on $($script:winrmJobIP) — waiting for service... ($remaining`s)"
            } else {
                $stLabel.Text = "Enabling WinRM on $($script:winrmJobIP) — verifying... ($elapsed`s)"
            }
        }
    })
    $script:winrmTimer.Start()
})

# Rescan a single host
$menuRescan.Add_Click({
    $ridx = $script:rightClickDtRowIndex
    if ($ridx -lt 0 -or $ridx -ge $dt.Rows.Count) {
        [System.Windows.Forms.MessageBox]::Show("No row selected. Right-click on a row first.", $script:AppTitle, "OK", "Warning") | Out-Null
        return
    }
    $ip = [string]$dt.Rows[$ridx]["IP"]
    if ([string]::IsNullOrWhiteSpace($ip)) {
        [System.Windows.Forms.MessageBox]::Show("Selected row has no IP address.", $script:AppTitle, "OK", "Warning") | Out-Null
        return
    }

    $stLabel.Text = "Rescanning $ip..."
    [System.Windows.Forms.Application]::DoEvents()

    # Build check options from current UI state
    $checks = @{
        LLMNR   = [bool]$chkLLMNR.Checked
        mDNS    = [bool]$chkMDNS.Checked
        NetBIOS = [bool]$chkNetBIOS.Checked
        SMBv1   = [bool]$chkSMBv1.Checked
        SMBSign = [bool]$chkSMBSign.Checked
        WPAD    = [bool]$chkWPAD.Checked
        IPv6    = [bool]$chkIPv6.Checked
    }
    $authMethod = $cmbAuth.SelectedItem.ToString()
    $domSuffix = $txtDomain.Text.Trim()
    $cred = $script:Cred

    $sessOpt = New-PSSessionOption -OperationTimeout 8000 -OpenTimeout 5000

    try {
        # Ping
        $up = $false
        if ($chkPing.Checked) {
            try { $up = [bool](Test-Connection -ComputerName $ip -Count 1 -Quiet -ErrorAction SilentlyContinue) } catch { $up = $false }
        } else { $up = $true }

        if (-not $up) {
            $dt.Rows[$ridx]["Ping"] = "Down"
            $dt.Rows[$ridx]["Status"] = "Skipped"
            $stLabel.Text = "$ip is still down (ping failed)."
            $grid.Refresh(); return
        }
        $dt.Rows[$ridx]["Ping"] = "Up"

        # WinRM port check
        $winrmOpen = $false
        if ($chkWinRM.Checked) {
            try {
                $tcp = New-Object System.Net.Sockets.TcpClient
                $ar = $tcp.BeginConnect($ip, 5985, $null, $null)
                $winrmOpen = $ar.AsyncWaitHandle.WaitOne(2000, $false)
                if ($winrmOpen -and $tcp.Connected) { $tcp.Close() } else { $tcp.Close(); $winrmOpen = $false }
            } catch { $winrmOpen = $false }
        } else { $winrmOpen = $true }

        $dt.Rows[$ridx]["WinRM"] = if ($winrmOpen) { "Open" } else { "Closed" }
        if (-not $winrmOpen) {
            $dt.Rows[$ridx]["Status"] = "Skipped"
            $stLabel.Text = "$ip WinRM port 5985 still closed."
            $grid.Refresh(); return
        }

        # Resolve FQDN
        $target = $ip
        try {
            $h = [System.Net.Dns]::GetHostEntry($ip)
            if ($h -and $h.HostName) {
                $n = $h.HostName.Trim().TrimEnd('.')
                if ($n -match '\.') { $target = $n.ToLower() } else { $target = ("{0}.{1}" -f $n, $domSuffix).ToLower() }
            }
        } catch {}
        $dt.Rows[$ridx]["Target"] = $target

        # RDP check
        if ($chkRDP.Checked) {
            try {
                $tcp2 = New-Object System.Net.Sockets.TcpClient
                $ar2 = $tcp2.BeginConnect($ip, 3389, $null, $null)
                $rdpOpen = $ar2.AsyncWaitHandle.WaitOne(1500, $false)
                if ($rdpOpen -and $tcp2.Connected) { $tcp2.Close() } else { $tcp2.Close(); $rdpOpen = $false }
            } catch { $rdpOpen = $false }
            $dt.Rows[$ridx]["RDP"] = if ($rdpOpen) { "Open" } else { "Closed" }
        }

        # Invoke-Command
        $inv = @{
            ComputerName  = $target
            SessionOption = $sessOpt
            ScriptBlock   = $RemoteScriptForRescan
            ArgumentList  = @(,$checks)
            ErrorAction   = "Stop"
        }
        if ($authMethod -eq "Kerberos") { $inv["Authentication"] = "Kerberos" }
        elseif ($authMethod -eq "Negotiate") { $inv["Authentication"] = "Negotiate" }
        if ($null -ne $cred) { $inv["Credential"] = $cred }

        $r = Invoke-Command @inv
        $dt.Rows[$ridx]["Auth"] = if ($inv.ContainsKey("Authentication")) { $inv["Authentication"] } else { "Default" }
        $dt.Rows[$ridx]["Status"] = "OK"
        $dt.Rows[$ridx]["Computer"] = $r.Computer
        $dt.Rows[$ridx]["OS"] = $r.OS
        if ($checks.LLMNR)   { $dt.Rows[$ridx]["LLMNR"]   = "$($r.LLMNR)" }
        if ($checks.mDNS)    { $dt.Rows[$ridx]["mDNS"]    = "$($r.mDNS)" }
        if ($checks.NetBIOS) { $dt.Rows[$ridx]["NetBIOS"] = "$($r.NetBIOS)" }
        if ($checks.SMBv1)   { $dt.Rows[$ridx]["SMBv1"]   = "$($r.SMBv1)" }
        if ($checks.SMBSign) { $dt.Rows[$ridx]["SMBSign"]  = "$($r.SMBSign)" }
        if ($checks.WPAD)    { $dt.Rows[$ridx]["WPAD"]    = "$($r.WPAD)" }
        if ($checks.IPv6)    { $dt.Rows[$ridx]["IPv6"]    = "$($r.IPv6)" }
        $dt.Rows[$ridx]["Error"] = ""
        $stLabel.Text = "Rescan of $ip completed successfully."
    } catch {
        $dt.Rows[$ridx]["Status"] = "Error"
        $errMsg = $_.Exception.Message
        if ($errMsg.Length -gt 150) { $errMsg = $errMsg.Substring(0, 150) + "..." }
        $dt.Rows[$ridx]["Error"] = $errMsg
        $stLabel.Text = "Rescan of $ip failed: $errMsg"
    }
    $grid.Refresh()
})

$menuCopyIP.Add_Click({
    $ridx = $script:rightClickDtRowIndex
    if ($ridx -ge 0 -and $ridx -lt $dt.Rows.Count) {
        $val = [string]$dt.Rows[$ridx]["IP"]
        if ($val) { [System.Windows.Forms.Clipboard]::SetText($val) }
    }
})

$menuCopyError.Add_Click({
    $ridx = $script:rightClickDtRowIndex
    if ($ridx -ge 0 -and $ridx -lt $dt.Rows.Count) {
        $val = [string]$dt.Rows[$ridx]["Error"]
        if ($val) { [System.Windows.Forms.Clipboard]::SetText($val) }
    }
})

# --- Status bar ---
$statusStrip=New-Object System.Windows.Forms.StatusStrip; $statusStrip.BackColor=$script:Colors.Primary
$stLabel=New-Object System.Windows.Forms.ToolStripStatusLabel; $stLabel.Text="Ready"; $stLabel.ForeColor=$script:Colors.TextLight
$stLabel.Font=New-Object System.Drawing.Font("Segoe UI",9); [void]$statusStrip.Items.Add($stLabel)
$stCount=New-Object System.Windows.Forms.ToolStripStatusLabel; $stCount.Text=""; $stCount.ForeColor=[System.Drawing.Color]::FromArgb(150,180,210)
$stCount.Alignment="Right"; $stCount.Spring=$true; [void]$statusStrip.Items.Add($stCount); $form.Controls.Add($statusStrip)
$progress=New-Object System.Windows.Forms.ProgressBar; $progress.Dock="Bottom"; $progress.Height=5; $progress.Style="Continuous"; $progress.ForeColor=$script:Colors.Accent
$form.Controls.Add($progress)

# =======================
# RUNTIME STATE
# =======================
$script:Cred=$null; $script:queue=New-Object System.Collections.Concurrent.ConcurrentQueue[object]
$script:cancelCts=New-Object System.Threading.CancellationTokenSource
$script:pool=$null; $script:jobs=@(); $script:doneWorkers=0; $script:sharedCounter=[ref]([int]0); $script:totalIPs=0

# =======================
# UI TIMER
# =======================
$uiTimer=New-Object System.Windows.Forms.Timer; $uiTimer.Interval=100
$uiTimer.Add_Tick({
    $drained=0
    while($drained -lt 500){
        $item=$null; if(-not $script:queue.TryDequeue([ref]$item)){break}; $drained++
        if($item -is [hashtable] -and $item.ContainsKey("meta")){
            if($item.meta -eq "progress"){
                $progress.Value=[Math]::Min($progress.Maximum,[int]$item.value)
                $stLabel.Text="Scanning {0}/{1}" -f [int]$item.value,$progress.Maximum
                $stCount.Text="Rows: {0}" -f $dt.Rows.Count; continue
            }
            if($item.meta -eq "doneWorker"){
                $script:doneWorkers++
                if($script:doneWorkers -ge $script:jobs.Count){
                    $uiTimer.Stop()
                    try{if($script:pool){$script:pool.Close();$script:pool.Dispose()}}catch{}
                    $script:pool=$null; $script:jobs=@(); $btnScan.Enabled=$true; $btnCancel.Enabled=$false; $progress.Value=$progress.Maximum
                    $ok=($dt.Select("Status = 'OK'")).Count; $er=($dt.Select("Status = 'Error'")).Count
                    $sk=($dt.Select("Status LIKE 'Skip%'")).Count; $rdp=($dt.Select("RDP = 'Open'")).Count
                    $stLabel.Text="Done | OK: $ok | Errors: $er | Skipped: $sk | RDP Open: $rdp"
                    $stCount.Text="Total: {0}" -f $dt.Rows.Count
                }; continue
            }
        }
        if($item -is [hashtable] -and $item.ContainsKey("IP")){
            $row=$dt.NewRow()
            foreach($c in $columns){if($item.ContainsKey($c)){$row[$c]=[string]$item[$c]}else{$row[$c]=""}}
            [void]$dt.Rows.Add($row)
        }
    }
})

# =======================
# BUTTON EVENTS
# =======================
$btnCred.Add_Click({
    $c=Show-CredentialDialog
    if($null -ne $c){$script:Cred=$c; $lblCred.Text="Auth: $($c.UserName) (explicit)"; $lblCred.ForeColor=$script:Colors.Accent}
    else{$script:Cred=$null; $id=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name; $lblCred.Text="Auth: $id (current session)"; $lblCred.ForeColor=$script:Colors.Success}
})
$btnExport.Add_Click({
    if($dt.Rows.Count -eq 0){[System.Windows.Forms.MessageBox]::Show("No data.","Export","OK","Information")|Out-Null; return}
    $sfd=New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter="CSV (*.csv)|*.csv"; $sfd.FileName="SecurityChecker-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    if($sfd.ShowDialog() -eq "OK"){$dt|Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Saved: $($sfd.FileName)","Export","OK","Information")|Out-Null}
})
$btnCancel.Add_Click({$btnCancel.Enabled=$false; $stLabel.Text="Cancelling..."; try{$script:cancelCts.Cancel()}catch{}})
$form.Add_FormClosing({try{$script:cancelCts.Cancel()}catch{}; try{if($script:pool){$script:pool.Close();$script:pool.Dispose()}}catch{}})

# =======================
# WORKER SCRIPT
# =======================
$workerScript={
    param($chunk,[bool]$pingFirst,[bool]$winrmCheck,[bool]$rdpCheck,[hashtable]$checks,$cred,[string]$domainSuffix,[string]$authMethod,$ct,$queue,[int]$total,[ref]$sharedCounter)

    function Ensure-Fqdn{param([string]$N,[string]$S); if([string]::IsNullOrWhiteSpace($N)){return $null}; $n=$N.Trim().TrimEnd('.'); if($n -eq ''){return $null}; if($n -match '\.'){return $n.ToLower()}; if([string]::IsNullOrWhiteSpace($S)){return $n.ToLower()}; return ("{0}.{1}" -f $n,$S).ToLower()}
    function Resolve-TargetFqdnFromIp{param([string]$Ip,[string]$Suffix); try{$h=[System.Net.Dns]::GetHostEntry($Ip); if($h -and $h.HostName){$f=Ensure-Fqdn -N $h.HostName -S $Suffix; if($f){return $f}}}catch{}; return $Ip}
    function Test-TcpPort{param([string]$Ip,[int]$Port,[int]$T=2000); try{$tcp=New-Object System.Net.Sockets.TcpClient; $ar=$tcp.BeginConnect($Ip,$Port,$null,$null); $ok=$ar.AsyncWaitHandle.WaitOne($T,$false); if($ok -and $tcp.Connected){$tcp.Close();return $true}; $tcp.Close(); return $false}catch{return $false}}

    $sessOpt=New-PSSessionOption -OperationTimeout 8000 -OpenTimeout 5000

    $RemoteScript={
        param($Options)
        function RegVal($p,$n){try{(Get-ItemProperty $p -ErrorAction Stop).$n}catch{$null}}
        $r=[ordered]@{}
        $r.Computer=$env:COMPUTERNAME
        try{$r.OS=(Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption}catch{$r.OS="Unknown"}

        if($Options.LLMNR){
            $v=RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' 'EnableMulticast'
            if($v -eq 0){$r.LLMNR='Disabled'}elseif($v -eq 1){$r.LLMNR='Enabled'}else{$r.LLMNR='NotSet'}
        }
        if($Options.mDNS){
            $v=RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' 'EnableMDNS'
            if($v -eq 0){$r.mDNS='Disabled'}elseif($v -eq 1){$r.mDNS='Enabled'}else{$r.mDNS='NotSet'}
        }
        if($Options.NetBIOS){
            try{
                $adapters=Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" -ErrorAction Stop
                $st=@(); foreach($a in $adapters){$v=$a.TcpipNetbiosOptions; switch($v){0{$st+="DHCP(Def)"};1{$st+="Enabled"};2{$st+="Disabled"};default{$st+="Unknown($v)"}}}
                $u=@($st|Select-Object -Unique); if($u.Count -eq 0){$r.NetBIOS="N/A"}elseif($u.Count -eq 1){$r.NetBIOS=$u[0]}else{$r.NetBIOS=($u -join '/')}
            }catch{$r.NetBIOS="Error"}
        }
        if($Options.SMBv1){
            try{$smb=Get-SmbServerConfiguration -ErrorAction Stop
                if($smb.EnableSMB1Protocol -eq $true){$r.SMBv1='Enabled'}elseif($smb.EnableSMB1Protocol -eq $false){$r.SMBv1='Disabled'}else{$r.SMBv1='Unknown'}
            }catch{try{$v=RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'SMB1'
                if($v -eq 0){$r.SMBv1='Disabled'}elseif($v -eq 1){$r.SMBv1='Enabled'}else{$r.SMBv1='NotSet'}}catch{$r.SMBv1='Error'}}
        }
        if($Options.SMBSign){
            try{$smb=Get-SmbServerConfiguration -ErrorAction Stop
                if($smb.RequireSecuritySignature -eq $true){$r.SMBSign='Required'}else{$r.SMBSign='NotReq'}
            }catch{try{$v=RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'RequireSecuritySignature'
                if($v -eq 1){$r.SMBSign='Required'}elseif($v -eq 0){$r.SMBSign='NotReq'}else{$r.SMBSign='NotSet'}}catch{$r.SMBSign='Error'}}
        }
        if($Options.WPAD){
            try{$svc=Get-Service -Name "WinHttpAutoProxySvc" -ErrorAction SilentlyContinue
                if($svc -and $svc.Status -eq 'Running'){$r.WPAD='Running'}elseif($svc){$r.WPAD='Stopped'}else{$r.WPAD='N/A'}
            }catch{$r.WPAD='Error'}
        }
        if($Options.IPv6){
            try{$b=Get-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue; $en=$b|Where-Object{$_.Enabled -eq $true}
                if($en){$r.IPv6='Enabled'}else{$r.IPv6='Disabled'}
            }catch{try{$v=RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' 'DisabledComponents'
                if($v -eq 0xFF){$r.IPv6='Disabled'}elseif($null -eq $v){$r.IPv6='Enabled'}else{$r.IPv6='Partial'}}catch{$r.IPv6='Error'}}
        }
        [pscustomobject]$r
    }

    foreach($ip in $chunk){
        if($ct.IsCancellationRequested){break}
        $val=[System.Threading.Interlocked]::Increment($sharedCounter.Value)
        if(($val % 3) -eq 0 -or $val -eq $total){$queue.Enqueue(@{meta="progress";value=$val})}

        $row=@{IP=$ip;Ping="";WinRM="";RDP="";Target="";Auth="";Status="";Computer="";OS="";LLMNR="";mDNS="";NetBIOS="";SMBv1="";SMBSign="";WPAD="";IPv6="";Error=""}

        if($pingFirst){try{$up=[bool](Test-Connection -ComputerName $ip -Count 1 -Quiet -ErrorAction SilentlyContinue)}catch{$up=$false}; $row.Ping=if($up){"Up"}else{"Down"}; if(-not $up){$row.Status="Skipped"; $queue.Enqueue($row); continue}}
        if($rdpCheck){$ro=Test-TcpPort -Ip $ip -Port 3389 -T 1500; $row.RDP=if($ro){"Open"}else{"Closed"}}
        if($winrmCheck){$po=Test-TcpPort -Ip $ip -Port 5985 -T 2000; $row.WinRM=if($po){"Open"}else{"Closed"}; if(-not $po){$row.Status="Skipped"; $row.Target=Resolve-TargetFqdnFromIp -Ip $ip -Suffix $domainSuffix; $queue.Enqueue($row); continue}}

        $target=Resolve-TargetFqdnFromIp -Ip $ip -Suffix $domainSuffix; $row.Target=$target
        $inv=@{ComputerName=$target;SessionOption=$sessOpt;ScriptBlock=$RemoteScript;ArgumentList=@(,$checks);ErrorAction="Stop"}
        if($authMethod -eq "Kerberos"){$inv["Authentication"]="Kerberos"}elseif($authMethod -eq "Negotiate"){$inv["Authentication"]="Negotiate"}
        if($null -ne $cred){$inv["Credential"]=$cred}

        try{
            $r=Invoke-Command @inv
            $row.Auth=if($inv.ContainsKey("Authentication")){$inv["Authentication"]}else{"Default"}; $row.Status="OK"; $row.Computer=$r.Computer; $row.OS=$r.OS
            if($checks.LLMNR){$row.LLMNR="$($r.LLMNR)"}; if($checks.mDNS){$row.mDNS="$($r.mDNS)"}; if($checks.NetBIOS){$row.NetBIOS="$($r.NetBIOS)"}
            if($checks.SMBv1){$row.SMBv1="$($r.SMBv1)"}; if($checks.SMBSign){$row.SMBSign="$($r.SMBSign)"}; if($checks.WPAD){$row.WPAD="$($r.WPAD)"}; if($checks.IPv6){$row.IPv6="$($r.IPv6)"}
        }catch{
            $errMsg=$_.Exception.Message; $row.Status="Error"; $row.Auth=if($inv.ContainsKey("Authentication")){$inv["Authentication"]}else{"Default"}
            if($errMsg -match "Kerberos" -and $target -match '^\d+\.\d+\.\d+\.\d+$'){$row.Error="Kerberos needs FQDN. Try Negotiate."}
            elseif($errMsg -match "Access.*denied"){$row.Error="Access denied on $target"}
            elseif($errMsg -match "WinRM|cannot connect"){$row.Error="WinRM failed: $target"}
            else{if($errMsg.Length -gt 150){$errMsg=$errMsg.Substring(0,150)+"..."}; $row.Error=$errMsg}
        }
        $queue.Enqueue($row)
    }
    $queue.Enqueue(@{meta="doneWorker"})
}

# =======================
# SCAN BUTTON
# =======================
$btnScan.Add_Click({
    try{
        if($null -ne $script:pool){throw "Scan already running."}
        $dt.Rows.Clear(); $progress.Value=0; $script:doneWorkers=0; $script:sharedCounter.Value=0
        try{$script:cancelCts.Dispose()}catch{}; $script:cancelCts=New-Object System.Threading.CancellationTokenSource

        $t=$txtTarget.Text; if([string]::IsNullOrWhiteSpace($t)){[System.Windows.Forms.MessageBox]::Show("Enter targets.","Input","OK","Warning")|Out-Null; return}
        $ips=Parse-TargetInput -InputText $t
        if($null -eq $ips -or $ips.Count -eq 0){[System.Windows.Forms.MessageBox]::Show("No usable IPs.","Input","OK","Warning")|Out-Null; return}

        $threads=[int]$numThreads.Value; $pingFirst=[bool]$chkPing.Checked; $winrmCheck=[bool]$chkWinRM.Checked; $rdpCheck=[bool]$chkRDP.Checked
        $checks=@{LLMNR=[bool]$chkLLMNR.Checked;mDNS=[bool]$chkMDNS.Checked;NetBIOS=[bool]$chkNetBIOS.Checked;SMBv1=[bool]$chkSMBv1.Checked;SMBSign=[bool]$chkSMBSign.Checked;WPAD=[bool]$chkWPAD.Checked;IPv6=[bool]$chkIPv6.Checked}
        $cred=$script:Cred; $authMethod=$cmbAuth.SelectedItem.ToString(); $domSuffix=$txtDomain.Text.Trim(); $ct=$script:cancelCts.Token
        $script:totalIPs=$ips.Count; $progress.Minimum=0; $progress.Maximum=$script:totalIPs
        $btnScan.Enabled=$false; $btnCancel.Enabled=$true
        $ad=if($null -ne $cred){$cred.UserName}else{"current session"}
        $stLabel.Text="Scanning {0} hosts | {1} threads | {2} | {3}" -f $ips.Count,$threads,$authMethod,$ad; $stCount.Text=""; $uiTimer.Start()

        $script:pool=[runspacefactory]::CreateRunspacePool(1,$threads); $script:pool.ApartmentState="MTA"; $script:pool.Open()
        $chunks=@(); for($i=0;$i -lt $threads;$i++){$chunks+=,(New-Object System.Collections.Generic.List[string])}
        for($i=0;$i -lt $ips.Count;$i++){$chunks[$i % $threads].Add($ips[$i])}

        $script:jobs=@()
        for($w=0;$w -lt $threads;$w++){
            $chunk=$chunks[$w]; if($chunk.Count -eq 0){continue}
            $ps=[PowerShell]::Create(); $ps.RunspacePool=$script:pool
            $null=$ps.AddScript($workerScript).AddArgument($chunk).AddArgument($pingFirst).AddArgument($winrmCheck).AddArgument($rdpCheck).AddArgument($checks).AddArgument($cred).AddArgument($domSuffix).AddArgument($authMethod).AddArgument($ct).AddArgument($script:queue).AddArgument([int]$script:totalIPs).AddArgument($script:sharedCounter)
            $handle=$ps.BeginInvoke(); $script:jobs+=[pscustomobject]@{PS=$ps;Handle=$handle}
        }
    }catch{
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"$($script:AppTitle) — Error","OK","Error")|Out-Null
        $btnScan.Enabled=$true; $btnCancel.Enabled=$false; $uiTimer.Stop()
        try{if($script:pool){$script:pool.Close();$script:pool.Dispose()}}catch{}; $script:pool=$null; $script:jobs=@(); $stLabel.Text="Ready"
    }
})

[void]$form.ShowDialog()
try{$script:cancelCts.Dispose()}catch{}; try{if($script:pool){$script:pool.Close();$script:pool.Dispose()}}catch{}