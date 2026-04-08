#Requires -Version 5.1
<#
.SYNOPSIS
    AD Lockout Manager
    Real-time lockout detection, investigation, and response tool.
.NOTES
    Requires: RSAT ActiveDirectory module, PowerShell 5.1+, Domain Admin or Account Operators
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
#  CONFIG & STATE
# ============================================================
$script:Config = @{
    TeamsWebhook    = ""
    EmailFrom       = ""
    EmailTo         = ""
    SmtpServer      = ""
    SmtpPort        = 587
    RefreshInterval = 60
    TeamsEnabled    = $false
    EmailEnabled    = $false
    AuditLogPath    = "$env:USERPROFILE\Desktop\ADLockoutAudit.csv"
    MaxHistoryRows  = 500
}
$script:AuditLog     = [System.Collections.Generic.List[PSObject]]::new()
$script:MonitorTimer = $null
$script:LastAlerted  = @{}
$script:UnlockCount  = 0
$script:AlertCount   = 0

# ============================================================
#  TOOLTIP ENGINE
# ============================================================
$tt = New-Object System.Windows.Forms.ToolTip
$tt.AutoPopDelay = 6000
$tt.InitialDelay = 400
$tt.ReshowDelay  = 300
$tt.ShowAlways   = $true

function Set-Tip { param($ctrl,$text) $tt.SetToolTip($ctrl,$text) }

# ============================================================
#  COLORS
# ============================================================
$C = @{
    BG          = [System.Drawing.Color]::FromArgb(13,  17,  23)
    Panel       = [System.Drawing.Color]::FromArgb(22,  27,  34)
    Card        = [System.Drawing.Color]::FromArgb(30,  37,  46)
    Border      = [System.Drawing.Color]::FromArgb(48,  54,  61)
    Accent      = [System.Drawing.Color]::FromArgb(0,  168, 255)
    AccentHov   = [System.Drawing.Color]::FromArgb(0,  120, 212)
    Danger      = [System.Drawing.Color]::FromArgb(255,  85,  85)
    Warning     = [System.Drawing.Color]::FromArgb(255, 170,   0)
    Success     = [System.Drawing.Color]::FromArgb(35,  209, 139)
    TextPrimary = [System.Drawing.Color]::FromArgb(230, 237, 243)
    TextMuted   = [System.Drawing.Color]::FromArgb(125, 133, 144)
    TextDim     = [System.Drawing.Color]::FromArgb(70,  80,  90)
    RowAlt      = [System.Drawing.Color]::FromArgb(18,  22,  28)
    RowHover    = [System.Drawing.Color]::FromArgb(38,  45,  56)
}

# ============================================================
#  UI FACTORY FUNCTIONS
# ============================================================
function New-Btn {
    param($Text, $X, $Y, $W=160, $H=32, $BG=$null, $FS=9)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Location = [System.Drawing.Point]::new($X,$Y)
    $b.Size = [System.Drawing.Size]::new($W,$H)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderSize = 1
    $b.Font = New-Object System.Drawing.Font("Segoe UI Semibold",$FS)
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $col = if ($BG) { $BG } else { $C.Accent }
    $b.BackColor = $col; $b.ForeColor = $C.TextPrimary
    $b.FlatAppearance.BorderColor = $col
    $b.FlatAppearance.MouseOverBackColor = $C.AccentHov
    $b.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0,90,170)
    return $b
}

function New-Lbl {
    param($Text, $X, $Y, $W=200, $H=20, $FS=9, $Bold=$false, $FG=$null)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Location = [System.Drawing.Point]::new($X,$Y)
    $l.Size = [System.Drawing.Size]::new($W,$H)
    $st = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $l.Font = New-Object System.Drawing.Font("Segoe UI",$FS,$st)
    $l.ForeColor = if ($FG) { $FG } else { $C.TextPrimary }
    $l.BackColor = [System.Drawing.Color]::Transparent
    return $l
}

function New-TB {
    param($X, $Y, $W=200, $H=26, $PH="")
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = [System.Drawing.Point]::new($X,$Y)
    $t.Size = [System.Drawing.Size]::new($W,$H)
    $t.BackColor = $C.Card; $t.ForeColor = $C.TextPrimary
    $t.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $t.Font = New-Object System.Drawing.Font("Segoe UI",9)
    if ($PH) {
        $t.Text = $PH; $t.ForeColor = $C.TextMuted; $t.Tag = $PH
        $t.Add_Enter({ if ($this.Text -eq $this.Tag) { $this.Text=""; $this.ForeColor=$C.TextPrimary } })
        $t.Add_Leave({ if ($this.Text -eq "")        { $this.Text=$this.Tag; $this.ForeColor=$C.TextMuted } })
    }
    return $t
}

function New-Grid {
    param($X, $Y, $W, $H)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location = [System.Drawing.Point]::new($X,$Y)
    $g.Size = [System.Drawing.Size]::new($W,$H)
    $g.BackgroundColor = $C.BG; $g.GridColor = $C.Border
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $g.RowHeadersVisible = $false
    $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false
    $g.ReadOnly = $true
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $g.MultiSelect = $false
    $g.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $g.ColumnHeadersHeight = 34; $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor  = $C.Panel
    $g.ColumnHeadersDefaultCellStyle.ForeColor  = $C.TextMuted
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI Semibold",8)
    $g.ColumnHeadersDefaultCellStyle.Padding = [System.Windows.Forms.Padding]::new(8,0,0,0)
    $g.DefaultCellStyle.BackColor = $C.BG; $g.DefaultCellStyle.ForeColor = $C.TextPrimary
    $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI",9)
    $g.DefaultCellStyle.Padding = [System.Windows.Forms.Padding]::new(6,0,0,0)
    $g.DefaultCellStyle.SelectionBackColor = $C.RowHover
    $g.DefaultCellStyle.SelectionForeColor = $C.TextPrimary
    $g.AlternatingRowsDefaultCellStyle.BackColor = $C.RowAlt
    $g.RowTemplate.Height = 28
    return $g
}

function New-Toggle {
    param($X, $Y, $LabelText, $On=$false)
    $p = New-Object System.Windows.Forms.Panel
    $p.Location = [System.Drawing.Point]::new($X,$Y)
    $p.Size = [System.Drawing.Size]::new(210,24)
    $p.BackColor = [System.Drawing.Color]::Transparent
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Appearance = [System.Windows.Forms.Appearance]::Button
    $cb.Location = [System.Drawing.Point]::new(0,2)
    $cb.Size = [System.Drawing.Size]::new(44,20)
    $cb.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cb.FlatAppearance.BorderSize = 0
    $cb.Checked = $On
    $cb.Cursor = [System.Windows.Forms.Cursors]::Hand
    $cb.BackColor = if ($On) { $C.Accent } else { $C.Border }
    $cb.ForeColor = $C.TextPrimary
    $cb.Text = if ($On) { "ON" } else { "OFF" }
    $cb.Font = New-Object System.Drawing.Font("Segoe UI",7,[System.Drawing.FontStyle]::Bold)
    $cb.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $cb.Add_CheckedChanged({
        $this.BackColor = if ($this.Checked) { $C.Accent } else { $C.Border }
        $this.Text      = if ($this.Checked) { "ON" } else { "OFF" }
    })
    $lb = New-Object System.Windows.Forms.Label
    $lb.Text = $LabelText; $lb.Location = [System.Drawing.Point]::new(50,3)
    $lb.Size = [System.Drawing.Size]::new(160,18)
    $lb.Font = New-Object System.Drawing.Font("Segoe UI",8.5)
    $lb.ForeColor = $C.TextPrimary; $lb.BackColor = [System.Drawing.Color]::Transparent
    $p.Controls.AddRange(@($cb,$lb))
    return @{ Panel=$p; Toggle=$cb }
}

function New-Divider { param($Y) $d=New-Object System.Windows.Forms.Panel; $d.Location=[System.Drawing.Point]::new(12,$Y); $d.Size=[System.Drawing.Size]::new(210,1); $d.BackColor=$C.Border; return $d }

function New-SectionLabel { param($Text,$Y) return (New-Lbl "  $Text" 12 $Y 210 16 7.5 $true $C.TextMuted) }

# ============================================================
#  AD MODULE CHECK
# ============================================================
function Test-ADModule {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        return $true
    }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Missing Dependency - RSAT Required"
    $dlg.Size = [System.Drawing.Size]::new(500,260)
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $dlg.BackColor = $C.Panel; $dlg.ForeColor = $C.TextPrimary
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.TopMost = $true

    $dlg.Controls.Add((New-Lbl "!" 20 16 36 50 26 $true $C.Warning))
    $dlg.Controls.Add((New-Lbl "ActiveDirectory Module Not Found" 66 20 380 24 11 $true $C.Warning))

    $ml = New-Object System.Windows.Forms.Label
    $ml.Text = "The ActiveDirectory PowerShell module (RSAT) is required.`n`nInstall it now? Requires:`n   - Windows 10/11 or Server`n   - Internet / WSUS access`n   - Administrator rights"
    $ml.Location = [System.Drawing.Point]::new(66,50); $ml.Size = [System.Drawing.Size]::new(410,110)
    $ml.Font = New-Object System.Drawing.Font("Segoe UI",9)
    $ml.ForeColor = $C.TextPrimary; $ml.BackColor = [System.Drawing.Color]::Transparent
    $dlg.Controls.Add($ml)

    $bi = New-Btn "Install RSAT Now"      66  182 180 34 $C.Accent
    $bs = New-Btn "Continue Without AD"  256  182 180 34 $C.Card
    $bs.FlatAppearance.BorderColor = $C.Border
    $script:RsatChoice = "skip"
    $bi.Add_Click({ $script:RsatChoice = "install"; $dlg.Close() })
    $bs.Add_Click({ $script:RsatChoice = "skip";    $dlg.Close() })
    $dlg.Controls.AddRange(@($bi,$bs)); $dlg.ShowDialog() | Out-Null

    if ($script:RsatChoice -eq "install") {
        $pf = New-Object System.Windows.Forms.Form
        $pf.Text = "Installing RSAT..."; $pf.Size = [System.Drawing.Size]::new(440,130)
        $pf.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $pf.BackColor = $C.Panel; $pf.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $pf.MaximizeBox = $false; $pf.TopMost = $true
        $pl = New-Lbl "Installing, please wait..." 20 26 400 20 9 $false $C.TextPrimary
        $pb = New-Object System.Windows.Forms.ProgressBar
        $pb.Location = [System.Drawing.Point]::new(20,54); $pb.Size = [System.Drawing.Size]::new(390,18)
        $pb.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $pf.Controls.AddRange(@($pl,$pb)); $pf.Show()
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $os = (Get-WmiObject Win32_OperatingSystem).Caption
            if ($os -match "Server") { Install-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction Stop | Out-Null }
            else { Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop | Out-Null }
            $pf.Close()
            if (Get-Module -ListAvailable -Name ActiveDirectory) {
                Import-Module ActiveDirectory -ErrorAction SilentlyContinue
                [System.Windows.Forms.MessageBox]::Show("RSAT installed successfully! The AD module is now loaded.","Install Complete","OK","Information") | Out-Null
                return $true
            } else {
                [System.Windows.Forms.MessageBox]::Show("Install completed but module not detected. Please restart PowerShell as Administrator.","Verify Failed","OK","Warning") | Out-Null
            }
        } catch {
            $pf.Close()
            [System.Windows.Forms.MessageBox]::Show("Installation failed:`n$_`n`nManual: Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'","Install Failed","OK","Error") | Out-Null
        }
    }
    return $false
}

# ============================================================
#  AD FUNCTIONS
# ============================================================
function Get-AllDCs        { try { return @(Get-ADDomainController -Filter *).Name } catch { return @() } }
function Get-PDCEmulator   { try { return (Get-ADDomain).PDCEmulator } catch { return "Unknown" } }

function Search-LockedAccounts {
    try {
        return @(Search-ADAccount -LockedOut -UsersOnly |
            Get-ADUser -Properties LockedOut,BadLogonCount,LastBadPasswordAttempt,PasswordLastSet,PasswordNeverExpires,Department,Title |
            Select-Object SamAccountName,Name,Department,Title,LockedOut,BadLogonCount,LastBadPasswordAttempt,PasswordLastSet,PasswordNeverExpires,DistinguishedName)
    } catch { return @() }
}

function Get-LockoutSource {
    param($Username)
    try {
        $pdc = Get-PDCEmulator
        $evts = @(Get-WinEvent -ComputerName $pdc -FilterHashtable @{LogName='Security';Id=4740} -MaxEvents 200 -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -like "*$Username*" })
        return @($evts | ForEach-Object {
            $x = [xml]$_.ToXml()
            [PSCustomObject]@{
                Time          = $_.TimeCreated
                Username      = ($x.Event.EventData.Data | Where-Object Name -eq 'TargetUserName').'#text'
                SourceMachine = ($x.Event.EventData.Data | Where-Object Name -eq 'CallerComputerName').'#text'
                DC            = $_.MachineName
            }
        })
    } catch { return @() }
}

function Get-ProcessSource {
    param($Computer, $Username)
    try {
        $evts = @(Get-WinEvent -ComputerName $Computer -FilterHashtable @{LogName='Security';Id=4625} -MaxEvents 100 -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -like "*$Username*" })
        return @($evts | ForEach-Object {
            $x = [xml]$_.ToXml()
            [PSCustomObject]@{
                Time      = $_.TimeCreated
                Process   = ($x.Event.EventData.Data | Where-Object Name -eq 'ProcessName').'#text'
                LogonType = ($x.Event.EventData.Data | Where-Object Name -eq 'LogonType').'#text'
                SourceIP  = ($x.Event.EventData.Data | Where-Object Name -eq 'IpAddress').'#text'
            }
        } | Select-Object -First 20)
    } catch { return @() }
}

function Invoke-UnlockAccount { param($u); try { Unlock-ADAccount -Identity $u; return $true } catch { return $false } }

function Get-PasswordPolicy {
    try { return @{ Domain=(Get-ADDefaultDomainPasswordPolicy); FGPP=@(Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction SilentlyContinue) } }
    catch { return $null }
}

# ============================================================
#  NOTIFICATIONS & AUDIT
# ============================================================
function Send-TeamsAlert {
    param($Msg)
    if (-not $script:Config.TeamsEnabled -or -not $script:Config.TeamsWebhook) { return }
    try { Invoke-RestMethod -Uri $script:Config.TeamsWebhook -Method Post -Body (@{text=$Msg}|ConvertTo-Json) -ContentType "application/json" | Out-Null } catch {}
}
function Send-EmailAlert {
    param($Subj,$Body)
    if (-not $script:Config.EmailEnabled -or -not $script:Config.SmtpServer) { return }
    try { Send-MailMessage -From $script:Config.EmailFrom -To $script:Config.EmailTo -Subject $Subj -Body $Body -SmtpServer $script:Config.SmtpServer -Port $script:Config.SmtpPort -UseSsl -ErrorAction SilentlyContinue } catch {}
}
function Write-AuditEntry {
    param($Action,$Username,$Details,$Operator)
    $e = [PSCustomObject]@{ Timestamp=(Get-Date -Format "yyyy-MM-dd HH:mm:ss"); Operator=$Operator; Action=$Action; Username=$Username; Details=$Details }
    $script:AuditLog.Add($e)
    try { $e | Export-Csv $script:Config.AuditLogPath -Append -NoTypeInformation -ErrorAction SilentlyContinue } catch {}
}

# ============================================================
#  MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Lockout Manager"
$form.Size = [System.Drawing.Size]::new(1300,840)
$form.MinimumSize = [System.Drawing.Size]::new(1100,700)
$form.BackColor = $C.BG; $form.ForeColor = $C.TextPrimary
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.Font = New-Object System.Drawing.Font("Segoe UI",9)
try { $form.Icon = [System.Drawing.SystemIcons]::Shield } catch {}

# -- HEADER (Dock=Top) ---------------------------------------------------------
$hdr = New-Object System.Windows.Forms.Panel
$hdr.Dock = [System.Windows.Forms.DockStyle]::Top; $hdr.Height = 130; $hdr.BackColor = $C.Panel
$hdr.Controls.Add((New-Lbl "AD Lockout Manager" 18 8 420 28 15 $true $C.Accent))
$hdr.Controls.Add((New-Lbl "AD Lockout Manager  -  Real-time Investigation and Response" 22 34 500 18 8 $false $C.TextMuted))

$script:lblMonStatus = New-Lbl "[ IDLE ]" 0 14 140 22 8.5 $true $C.TextMuted
$script:lblMonStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:lblMonStatus.Anchor = "Top,Right"

$script:lblRefresh = New-Lbl "Last refresh: --" 0 34 230 18 7.5 $false $C.TextDim
$script:lblRefresh.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:lblRefresh.Anchor = "Top,Right"

$hdr.Controls.AddRange(@($script:lblMonStatus,$script:lblRefresh))
$hdr.Add_Resize({ $script:lblMonStatus.Left=$hdr.Width-152; $script:lblRefresh.Left=$hdr.Width-242 })
$form.Controls.Add($hdr)

# -- STATUS BAR (Dock=Bottom) --------------------------------------------------
$sbar = New-Object System.Windows.Forms.Panel
$sbar.Dock = [System.Windows.Forms.DockStyle]::Bottom; $sbar.Height = 26; $sbar.BackColor = $C.Panel

$script:lblStatus = New-Lbl "  Ready." 0 4 800 18 8 $false $C.TextMuted
$script:lblStatus.Anchor = "Left,Bottom"

$lblVer = New-Lbl "v1.0  -  AD Lockout Manager" 0 4 210 18 7.5 $false $C.TextDim
$lblVer.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $lblVer.Anchor = "Right,Bottom"
$sbar.Controls.AddRange(@($script:lblStatus,$lblVer))
$sbar.Add_Resize({ $lblVer.Left=$sbar.Width-220 })
$form.Controls.Add($sbar)

# -- CONTENT PANEL (Dock=Fill) -------------------------------------------------
# NOTE: Fill must be added BEFORE Left/Right docked panels
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$contentPanel.BackColor = $C.BG
$form.Controls.Add($contentPanel)

# -- SIDEBAR (Dock=Left, added AFTER Fill) -------------------------------------
$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.Dock = [System.Windows.Forms.DockStyle]::Left
$sidebar.Width = 244; $sidebar.BackColor = $C.Panel
$form.Controls.Add($sidebar)

function Set-Status { param($Msg,$Col=$null); $script:lblStatus.Text="  $Msg"; $script:lblStatus.ForeColor=if($Col){$Col}else{$C.TextMuted} }

# -- SIDEBAR CONTENTS ----------------------------------------------------------
$sidebar.Controls.Add((New-SectionLabel "ACCOUNT SEARCH" 14))
$tbSearch = New-TB 12 32 218 26 "Username or Display Name..."
$sidebar.Controls.Add($tbSearch)
$btnSearch       = New-Btn "Search"          12 66  218 32 $C.Accent
$btnSearchLocked = New-Btn "Show All Locked" 12 104 218 32 ([System.Drawing.Color]::FromArgb(160,45,45))
$sidebar.Controls.AddRange(@($btnSearch,$btnSearchLocked))
$sidebar.Controls.Add((New-Divider 148))

$sidebar.Controls.Add((New-SectionLabel "REAL-TIME MONITOR" 160))
$monToggle = New-Toggle 12 180 "Auto-Monitor"
$sidebar.Controls.Add($monToggle.Panel)
$sidebar.Controls.Add((New-Lbl "Refresh every (sec):" 12 212 140 16 8 $false $C.TextMuted))
$tbInterval = New-TB 12 230 76 26; $tbInterval.Text = "60"; $sidebar.Controls.Add($tbInterval)
$btnApply = New-Btn "Apply" 96 230 118 26 $C.Card 8; $btnApply.FlatAppearance.BorderColor=$C.Border; $sidebar.Controls.Add($btnApply)
$sidebar.Controls.Add((New-Divider 268))

$sidebar.Controls.Add((New-SectionLabel "NOTIFICATIONS" 280))
$teamsToggle = New-Toggle 12 300 "Teams Alerts"; $sidebar.Controls.Add($teamsToggle.Panel)
$emailToggle = New-Toggle 12 326 "Email Alerts";  $sidebar.Controls.Add($emailToggle.Panel)
$btnConfigNotif = New-Btn "Configure..." 12 358 218 28 $C.Card 8; $btnConfigNotif.FlatAppearance.BorderColor=$C.Border; $sidebar.Controls.Add($btnConfigNotif)
$sidebar.Controls.Add((New-Divider 398))

$sidebar.Controls.Add((New-SectionLabel "TOOLS" 410))
$btnViewAudit   = New-Btn "View Audit Log"   12 430 218 28 $C.Card 8
$btnExportAudit = New-Btn "Export Audit CSV" 12 464 218 28 $C.Card 8
$btnRefreshDCs  = New-Btn "Refresh DC List"  12 498 218 28 $C.Card 8
$btnViewAudit.FlatAppearance.BorderColor=$C.Border; $btnExportAudit.FlatAppearance.BorderColor=$C.Border; $btnRefreshDCs.FlatAppearance.BorderColor=$C.Border
$sidebar.Controls.AddRange(@($btnViewAudit,$btnExportAudit,$btnRefreshDCs))
$sidebar.Controls.Add((New-Divider 538))

$sidebar.Controls.Add((New-SectionLabel "DOMAIN CONTROLLERS" 550))
$lbDCs = New-Object System.Windows.Forms.ListBox
$lbDCs.Location=[System.Drawing.Point]::new(12,570); $lbDCs.Size=[System.Drawing.Size]::new(218,180)
$lbDCs.BackColor=$C.Card; $lbDCs.ForeColor=$C.TextPrimary
$lbDCs.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$lbDCs.Font=New-Object System.Drawing.Font("Segoe UI",8.5)
$sidebar.Controls.Add($lbDCs)

# -- STAT CARDS (inside header panel, Y=58) ------------------------------------
function New-StatCard {
    param($Label, $Val, $X, $Col)
    $card = New-Object System.Windows.Forms.Panel
    $card.Location=[System.Drawing.Point]::new($X,58); $card.Size=[System.Drawing.Size]::new(160,66); $card.BackColor=$C.Card
    $vl = New-Lbl $Val 10 4 140 36 20 $true $Col; $vl.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft
    $kl = New-Lbl $Label 10 40 140 20 8 $false $C.TextMuted
    $card.Controls.AddRange(@($vl,$kl))
    return @{Panel=$card; VL=$vl}
}
$sc1=New-StatCard "Locked Accounts" "0" 8   $C.Danger
$sc2=New-StatCard "DCs Queried"     "0" 178 $C.Accent
$sc3=New-StatCard "Unlocks Today"   "0" 348 $C.Success
$sc4=New-StatCard "Alerts Sent"     "0" 518 $C.Warning
$hdr.Controls.AddRange(@($sc1.Panel,$sc2.Panel,$sc3.Panel,$sc4.Panel))
$script:S_Locked=$sc1.VL; $script:S_DCs=$sc2.VL; $script:S_Unlocks=$sc3.VL; $script:S_Alerts=$sc4.VL

# -- TAB CONTROL (inside contentPanel, Dock=Fill) ------------------------------
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabs.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabs.ItemSize = [System.Drawing.Size]::new(172,32)
$tabs.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabs.BackColor = $C.BG; $tabs.Padding = [System.Drawing.Point]::new(14,6)
$tabs.Add_DrawItem({
    param($s,$e)
    $tab=$s.TabPages[$e.Index]; $sel=($e.Index -eq $s.SelectedIndex)
    $bg=if($sel){$C.Accent}else{$C.Card}; $fg=if($sel){$C.TextPrimary}else{$C.TextMuted}
    $br=New-Object System.Drawing.SolidBrush($bg); $e.Graphics.FillRectangle($br,$e.Bounds)
    $f=New-Object System.Drawing.Font("Segoe UI Semibold",8.5)
    $fb=New-Object System.Drawing.SolidBrush($fg); $sf=New-Object System.Drawing.StringFormat
    $sf.Alignment=[System.Drawing.StringAlignment]::Center; $sf.LineAlignment=[System.Drawing.StringAlignment]::Center
    $e.Graphics.DrawString($tab.Text,$f,$fb,[System.Drawing.RectangleF]$e.Bounds,$sf)
    $br.Dispose(); $fb.Dispose(); $f.Dispose()
})
$contentPanel.Controls.Add($tabs)

# ============================================================
#  TAB 1 - LOCKED ACCOUNTS
# ============================================================
$tabLocked = New-Object System.Windows.Forms.TabPage
$tabLocked.Text = "Locked Accounts"
$tabLocked.BackColor = $C.BG
$tabLocked.Padding = [System.Windows.Forms.Padding]::new(0)

# TableLayoutPanel: 3 rows - grid(Fill), buttons(Auto), detail(160px)
$tbl = New-Object System.Windows.Forms.TableLayoutPanel
$tbl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tbl.ColumnCount = 1
$tbl.RowCount = 3
$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))  | Out-Null
$tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 180))) | Out-Null
$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$tbl.BackColor = $C.BG
$tbl.Margin = [System.Windows.Forms.Padding]::new(0)

# Row 0: Grid
$gridLocked = New-Grid 0 0 100 100
$gridLocked.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridLocked.Margin = [System.Windows.Forms.Padding]::new(0)
$gridLocked.Columns.Add("SAM","Username")            | Out-Null
$gridLocked.Columns.Add("Name","Full Name")           | Out-Null
$gridLocked.Columns.Add("Dept","Department")          | Out-Null
$gridLocked.Columns.Add("Bad","Bad Logons")           | Out-Null
$gridLocked.Columns.Add("LastBad","Last Bad Attempt") | Out-Null
$gridLocked.Columns.Add("PwdSet","Pwd Last Set")      | Out-Null
$gridLocked.Columns.Add("Status","Status")            | Out-Null
$gridLocked.Columns["Bad"].FillWeight    = 9
$gridLocked.Columns["Status"].FillWeight = 11
$gridLocked.RowTemplate.Height = 28
$tbl.Controls.Add($gridLocked, 0, 0)

# Row 1: Action buttons
$actRow = New-Object System.Windows.Forms.Panel
$actRow.Dock = [System.Windows.Forms.DockStyle]::Fill
$actRow.Margin = [System.Windows.Forms.Padding]::new(0)
$actRow.BackColor = $C.Panel
$btnUnlock      = New-Btn "Unlock Selected"    4   5 170 32 $C.Success
$btnUnlockAll   = New-Btn "Unlock All"         182 5 130 32 ([System.Drawing.Color]::FromArgb(190,110,0))
$btnInvestigate = New-Btn "Investigate Source" 320 5 180 32 $C.Accent
$btnCopyUser    = New-Btn "Copy Username"      508 5 150 32 $C.Card
$btnCopyUser.FlatAppearance.BorderColor = $C.Border
$actRow.Controls.AddRange(@($btnUnlock,$btnUnlockAll,$btnInvestigate,$btnCopyUser))
$tbl.Controls.Add($actRow, 0, 1)

# Row 2: Detail pane
$detPane = New-Object System.Windows.Forms.Panel
$detPane.Dock = [System.Windows.Forms.DockStyle]::Fill
$detPane.Margin = [System.Windows.Forms.Padding]::new(0)
$detPane.BackColor = $C.Card
$lblDetTitle = New-Lbl "  Account Details" 8 6 300 18 9 $true $C.Accent
$script:rtDetail = New-Object System.Windows.Forms.RichTextBox
$script:rtDetail.Location = [System.Drawing.Point]::new(0, 28)
$script:rtDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:rtDetail.BackColor = $C.Card
$script:rtDetail.ForeColor = $C.TextPrimary
$script:rtDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:rtDetail.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$script:rtDetail.ReadOnly = $true
$detPane.Controls.Add($script:rtDetail)
$detPane.Controls.Add($lblDetTitle)
$tbl.Controls.Add($detPane, 0, 2)

$tabLocked.Controls.Add($tbl)
$tabs.TabPages.Add($tabLocked)

# ============================================================
#  TAB 2 - INVESTIGATION
# ============================================================
$tabInvest = New-Object System.Windows.Forms.TabPage
$tabInvest.Text = "Investigation"; $tabInvest.BackColor = $C.BG
$tabInvest.Padding = [System.Windows.Forms.Padding]::new(0)

# TableLayoutPanel for Investigation: toolbar, sources grid, proc toolbar, procs grid
$tblInv = New-Object System.Windows.Forms.TableLayoutPanel
$tblInv.Dock = [System.Windows.Forms.DockStyle]::Fill
$tblInv.ColumnCount = 1; $tblInv.RowCount = 6
$tblInv.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))  | Out-Null
$tblInv.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))  | Out-Null
$tblInv.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))   | Out-Null
$tblInv.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))  | Out-Null
$tblInv.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))  | Out-Null
$tblInv.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))   | Out-Null
$tblInv.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100))) | Out-Null
$tblInv.BackColor = $C.BG; $tblInv.Margin = [System.Windows.Forms.Padding]::new(0)

# Row 0: search bar
$invBar = New-Object System.Windows.Forms.Panel
$invBar.Dock = [System.Windows.Forms.DockStyle]::Fill; $invBar.BackColor = $C.BG
$invBar.Margin = [System.Windows.Forms.Padding]::new(0)
$invBar.Controls.Add((New-Lbl "Investigate Username:" 4 8 165 20 9 $false $C.TextMuted))
$tbInvUser = New-TB 172 5 200 26; $invBar.Controls.Add($tbInvUser)
$btnInvGo = New-Btn "Investigate" 380 4 150 28 $C.Accent 8.5; $invBar.Controls.Add($btnInvGo)
$tblInv.Controls.Add($invBar, 0, 0)

# Row 1: sources label
$lbSrc = New-Lbl "  LOCKOUT SOURCES  (Event 4740 from PDC Emulator)" 0 2 600 16 7.5 $true $C.TextMuted
$lbSrc.Dock = [System.Windows.Forms.DockStyle]::Fill; $lbSrc.Margin = [System.Windows.Forms.Padding]::new(0)
$tblInv.Controls.Add($lbSrc, 0, 1)

# Row 2: sources grid
$gridSources = New-Grid 0 0 100 100
$gridSources.Dock = [System.Windows.Forms.DockStyle]::Fill; $gridSources.Margin = [System.Windows.Forms.Padding]::new(0)
$gridSources.Columns.Add("Time","Time")          | Out-Null
$gridSources.Columns.Add("User","Username")      | Out-Null
$gridSources.Columns.Add("Src","Source Machine") | Out-Null
$gridSources.Columns.Add("DC","Locking DC")      | Out-Null
$tblInv.Controls.Add($gridSources, 0, 2)

# Row 3: process fetch bar
$procBar = New-Object System.Windows.Forms.Panel
$procBar.Dock = [System.Windows.Forms.DockStyle]::Fill; $procBar.BackColor = $C.BG
$procBar.Margin = [System.Windows.Forms.Padding]::new(0)
$tbSrcMachine = New-TB 4 3 220 26 "Source machine name..."; $procBar.Controls.Add($tbSrcMachine)
$btnFetchProcs = New-Btn "Fetch Processes" 232 3 160 26 $C.Card 8
$btnFetchProcs.FlatAppearance.BorderColor = $C.Border; $procBar.Controls.Add($btnFetchProcs)
$tblInv.Controls.Add($procBar, 0, 3)

# Row 4: procs label
$lbProc = New-Lbl "  PROCESS DETAILS  (Event 4625 from source machine)" 0 2 600 16 7.5 $true $C.TextMuted
$lbProc.Dock = [System.Windows.Forms.DockStyle]::Fill; $lbProc.Margin = [System.Windows.Forms.Padding]::new(0)
$tblInv.Controls.Add($lbProc, 0, 4)

# Row 5: procs grid
$gridProcs = New-Grid 0 0 100 100
$gridProcs.Dock = [System.Windows.Forms.DockStyle]::Fill; $gridProcs.Margin = [System.Windows.Forms.Padding]::new(0)
$gridProcs.Columns.Add("Time","Time")      | Out-Null
$gridProcs.Columns.Add("Proc","Process")   | Out-Null
$gridProcs.Columns.Add("LT","Logon Type")  | Out-Null
$gridProcs.Columns.Add("IP","Source IP")   | Out-Null
$tblInv.Controls.Add($gridProcs, 0, 5)

$tabInvest.Controls.Add($tblInv)
$tabs.TabPages.Add($tabInvest)

# ============================================================
#  TAB 3 - PASSWORD POLICY
# ============================================================
$tabPolicy = New-Object System.Windows.Forms.TabPage
$tabPolicy.Text = "Password Policy"; $tabPolicy.BackColor = $C.BG
$tabPolicy.Padding = [System.Windows.Forms.Padding]::new(0)

$tblPol = New-Object System.Windows.Forms.TableLayoutPanel
$tblPol.Dock = [System.Windows.Forms.DockStyle]::Fill; $tblPol.ColumnCount = 1; $tblPol.RowCount = 2
$tblPol.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 38))) | Out-Null
$tblPol.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$tblPol.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100))) | Out-Null
$tblPol.BackColor = $C.BG; $tblPol.Margin = [System.Windows.Forms.Padding]::new(0)

$polBar = New-Object System.Windows.Forms.Panel
$polBar.Dock = [System.Windows.Forms.DockStyle]::Fill; $polBar.BackColor = $C.BG
$polBar.Margin = [System.Windows.Forms.Padding]::new(0)
$btnLoadPolicy = New-Btn "Load Policy" 4 4 160 30 $C.Accent; $polBar.Controls.Add($btnLoadPolicy)
$tblPol.Controls.Add($polBar, 0, 0)

$script:rtPolicy = New-Object System.Windows.Forms.RichTextBox
$script:rtPolicy.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:rtPolicy.Margin = [System.Windows.Forms.Padding]::new(0)
$script:rtPolicy.BackColor = $C.Card; $script:rtPolicy.ForeColor = $C.TextPrimary
$script:rtPolicy.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:rtPolicy.Font = New-Object System.Drawing.Font("Consolas",9); $script:rtPolicy.ReadOnly = $true
$tblPol.Controls.Add($script:rtPolicy, 0, 1)

$tabPolicy.Controls.Add($tblPol)
$tabs.TabPages.Add($tabPolicy)

# ============================================================
#  TAB 4 - AUDIT LOG
# ============================================================
$tabAudit = New-Object System.Windows.Forms.TabPage; $tabAudit.Text="Audit Log"; $tabAudit.BackColor=$C.BG; $tabAudit.Padding=[System.Windows.Forms.Padding]::new(0)
$gridAudit = New-Grid 0 0 900 520; $gridAudit.Dock=[System.Windows.Forms.DockStyle]::Fill
$gridAudit.Columns.Add("TS","Timestamp")  | Out-Null
$gridAudit.Columns.Add("Op","Operator")   | Out-Null
$gridAudit.Columns.Add("Act","Action")    | Out-Null
$gridAudit.Columns.Add("User","Username") | Out-Null
$gridAudit.Columns.Add("Det","Details")   | Out-Null
$gridAudit.Columns["TS"].FillWeight=15; $gridAudit.Columns["Op"].FillWeight=13; $gridAudit.Columns["Act"].FillWeight=12
$gridAudit.Columns["User"].FillWeight=14; $gridAudit.Columns["Det"].FillWeight=46
$tabAudit.Controls.Add($gridAudit)
$tabs.TabPages.Add($tabAudit)

# ============================================================
#  TAB 5 - SETTINGS
# ============================================================
$tabSettings = New-Object System.Windows.Forms.TabPage; $tabSettings.Text="Settings"; $tabSettings.BackColor=$C.BG; $tabSettings.Padding=[System.Windows.Forms.Padding]::new(8,8,8,8)

function New-SettingRow { param($Label,$Val,$Y,$Pwd=$false)
    $tabSettings.Controls.Add((New-Lbl $Label 0 ($Y+3) 205 20 9 $false $C.TextMuted))
    $t=New-TB 210 $Y 400 26; $t.Text=$Val; if($Pwd){$t.UseSystemPasswordChar=$true}
    $tabSettings.Controls.Add($t); return $t }

$tabSettings.Controls.Add((New-Lbl "Notification Settings" 0 8 400 22 11 $true $C.Accent))
$tbTeamsWH  = New-SettingRow "Teams Webhook URL:"  $script:Config.TeamsWebhook 40
$tbEmailFr  = New-SettingRow "Email From:"         $script:Config.EmailFrom    74
$tbEmailTo  = New-SettingRow "Email To:"           $script:Config.EmailTo      108
$tbSmtp     = New-SettingRow "SMTP Server:"        $script:Config.SmtpServer   142
$tbSmtpPort = New-SettingRow "SMTP Port:"          "$($script:Config.SmtpPort)" 176
$tbAuditPath= New-SettingRow "Audit Log Path:"     $script:Config.AuditLogPath  210

$tabSettings.Controls.Add((New-Lbl "Application Settings" 0 260 400 22 11 $true $C.Accent))
$tabSettings.Controls.Add((New-Lbl "Max Audit Rows in Memory:" 0 294 220 20 9 $false $C.TextMuted))
$tbMaxRows = New-TB 224 290 80 26; $tbMaxRows.Text="$($script:Config.MaxHistoryRows)"; $tabSettings.Controls.Add($tbMaxRows)

$btnSaveSet  = New-Btn "Save Settings"  0   345 180 34 $C.Success
$btnTestTeam = New-Btn "Test Teams"     188 345 160 34 $C.Card; $btnTestTeam.FlatAppearance.BorderColor=$C.Border
$btnTestMail = New-Btn "Test Email"     356 345 160 34 $C.Card; $btnTestMail.FlatAppearance.BorderColor=$C.Border
$tabSettings.Controls.AddRange(@($btnSaveSet,$btnTestTeam,$btnTestMail))
$tabs.TabPages.Add($tabSettings)

# ============================================================
#  HELPER: Resize grids on tab resize
# ============================================================
function Resize-TabGrids {
    # All grids use Dock=Fill - WinForms handles resize automatically
}
$tabs.Add_SizeChanged({ Resize-TabGrids })

# ============================================================
#  HELPER: Update Locked Grid
# ============================================================
function Update-LockedGrid {
    param($accounts)
    $gridLocked.Rows.Clear()
    $arr = @($accounts)
    $lockedCount = @($arr | Where-Object { $_.LockedOut -eq $true }).Count
    $script:S_Locked.Text = $lockedCount.ToString()
    foreach ($acc in $arr) {
        $st  = if ($acc.LockedOut) { "LOCKED" } else { "OK" }
        $lb  = if ($acc.LastBadPasswordAttempt) { $acc.LastBadPasswordAttempt.ToString("MM/dd HH:mm:ss") } else { "--" }
        $ps  = if ($acc.PasswordLastSet)         { $acc.PasswordLastSet.ToString("MM/dd/yyyy") }          else { "--" }
        $idx = $gridLocked.Rows.Add($acc.SamAccountName, $acc.Name, $acc.Department, $acc.BadLogonCount, $lb, $ps, $st)
        $row = $gridLocked.Rows[$idx]
        if ($acc.LockedOut) {
            $row.DefaultCellStyle.ForeColor = $C.Danger
        } else {
            $row.DefaultCellStyle.ForeColor = $C.TextPrimary
            $row.Cells["Status"].Style.ForeColor = $C.Success
        }
    }
}


# ============================================================
#  TOOLTIPS - hover over any control for explanation
# ============================================================
Set-Tip $tbSearch         "Type a username, display name, or partial name to search Active Directory"
Set-Tip $btnSearch        "Search Active Directory for accounts matching the entered name"
Set-Tip $btnSearchLocked  "Query all Domain Controllers and list every currently locked account"
Set-Tip $monToggle.Toggle "Enable real-time monitoring - automatically refreshes locked accounts on the interval below"
Set-Tip $monToggle.Panel  "Enable real-time monitoring - automatically refreshes locked accounts on the interval below"
Set-Tip $tbInterval       "How often (in seconds) the tool auto-refreshes when monitoring is enabled. Minimum: 10 seconds"
Set-Tip $btnApply         "Apply the new refresh interval to the running monitor"
Set-Tip $teamsToggle.Toggle "Send a Microsoft Teams notification when a new lockout is detected (configure webhook in Settings)"
Set-Tip $teamsToggle.Panel  "Send a Microsoft Teams notification when a new lockout is detected (configure webhook in Settings)"
Set-Tip $emailToggle.Toggle "Send an email alert when a new lockout is detected (configure SMTP settings in Settings)"
Set-Tip $emailToggle.Panel  "Send an email alert when a new lockout is detected (configure SMTP settings in Settings)"
Set-Tip $btnConfigNotif   "Open the Settings tab to configure Teams webhook and email SMTP settings"
Set-Tip $btnViewAudit     "Open the Audit Log tab showing all actions taken in this session"
Set-Tip $btnExportAudit   "Export the full audit log to a CSV file for record-keeping or compliance"
Set-Tip $btnRefreshDCs    "Query Active Directory to refresh the list of Domain Controllers in your environment"
Set-Tip $lbDCs            "All Domain Controllers in your environment. [PDC] marks the PDC Emulator - lockout events are queried from this DC."
Set-Tip $btnUnlock        "Unlock the selected account - requires confirmation before proceeding"
Set-Tip $btnUnlockAll     "Unlock ALL currently locked accounts shown in the list - use with caution"
Set-Tip $btnInvestigate   "Investigate the selected account - queries Event ID 4740 from the PDC Emulator to find the lockout source machine"
Set-Tip $btnCopyUser      "Copy the selected username to the clipboard"
Set-Tip $gridLocked       "List of accounts. Click a row to see full account details below. Red = locked, green = unlocked."
Set-Tip $script:rtDetail  "Full account details from Active Directory for the selected user"
Set-Tip $tbInvUser        "Enter the username (SAMAccountName) to investigate lockout events for"
Set-Tip $btnInvGo         "Query Event ID 4740 from the PDC Emulator to find which machine caused the lockout"
Set-Tip $gridSources      "Lockout source events (4740) - shows which machine was sending bad credentials and when"
Set-Tip $tbSrcMachine     "Enter the source machine name from the table above to drill into which process caused the lockout"
Set-Tip $btnFetchProcs    "Query Event ID 4625 on the source machine to identify the exact process or service using stale credentials"
Set-Tip $gridProcs        "Failed logon events (4625) from the source machine - Logon Type 5 = Service, Type 4 = Scheduled Task, Type 3 = Network"
Set-Tip $btnLoadPolicy    "Load the default domain password policy and any Fine-Grained Password Policies (FGPP) from Active Directory"
Set-Tip $script:rtPolicy  "Password and lockout policy details. Lockout threshold = how many bad attempts before lockout. Observation window = the reset period."
Set-Tip $tbTeamsWH        "Full URL of your Microsoft Teams incoming webhook. Create one in Teams: channel settings > Connectors > Incoming Webhook"
Set-Tip $tbEmailFr        "The sender email address for alert notifications (e.g. it-alerts@yourcompany.com)"
Set-Tip $tbEmailTo        "The recipient email address for alert notifications"
Set-Tip $tbSmtp           "Your SMTP server hostname (e.g. smtp.office365.com or your internal mail relay)"
Set-Tip $tbSmtpPort       "SMTP port number - typically 587 for TLS/STARTTLS or 465 for SSL"
Set-Tip $tbAuditPath      "File path where the audit log CSV is saved automatically (default: your Desktop)"
Set-Tip $tbMaxRows        "Maximum number of audit entries to keep in memory during this session"
Set-Tip $btnSaveSet       "Save all settings for this session (note: settings do not persist after closing - add a config file path to make them permanent)"
Set-Tip $btnTestTeam      "Send a test message to your configured Teams webhook to verify the connection"
Set-Tip $btnTestMail      "Send a test email to verify your SMTP configuration is working"
Set-Tip $sc1.Panel        "Total number of accounts currently shown as locked in the results list"
Set-Tip $sc2.Panel        "Number of Domain Controllers discovered in your environment"
Set-Tip $sc3.Panel        "Number of accounts unlocked during this session"
Set-Tip $sc4.Panel        "Number of Teams/Email alerts sent during this session"

# ============================================================
#  EVENTS
# ============================================================

# Search
$btnSearch.Add_Click({
    $q = $tbSearch.Text.Trim()
    if (-not $q -or $q -eq $tbSearch.Tag) { [System.Windows.Forms.MessageBox]::Show("Enter a username or display name.","Search","OK","Information")|Out-Null; return }
    Set-Status "Searching for '$q'..." $C.Accent
    $form.Cursor=[System.Windows.Forms.Cursors]::WaitCursor
    try {
        $r = @(Get-ADUser -Filter "SamAccountName -like '*$q*' -or DisplayName -like '*$q*' -or Name -like '*$q*'" `
            -Properties LockedOut,BadLogonCount,LastBadPasswordAttempt,PasswordLastSet,Department,Title |
            Select-Object SamAccountName,Name,Department,Title,LockedOut,BadLogonCount,LastBadPasswordAttempt,PasswordLastSet)
        Update-LockedGrid $r
        Set-Status "Found $($r.Count) account(s) matching '$q'." $C.Success
        $tabs.SelectedIndex=0
    } catch { Set-Status "Search failed: $_" $C.Danger }
    $form.Cursor=[System.Windows.Forms.Cursors]::Default
})

# Show All Locked
$btnSearchLocked.Add_Click({
    Set-Status "Querying all DCs for locked accounts..." $C.Warning
    $form.Cursor=[System.Windows.Forms.Cursors]::WaitCursor
    try {
        $locked = Search-LockedAccounts
        Update-LockedGrid $locked
        $script:lblRefresh.Text = "Last refresh: $(Get-Date -Format 'HH:mm:ss')"
        $dcs = Get-AllDCs; $script:S_DCs.Text = $dcs.Count.ToString()
        if ($lbDCs.Items.Count -eq 0) {
            $pdc = Get-PDCEmulator; $pdcShort = ($pdc -split "\.")[0]
            foreach ($dc in $dcs) {
                $dcShort = ($dc -split "\.")[0]
                $label = if ($dcShort -eq $pdcShort -or $dc -eq $pdc) { "$dc  [PDC]" } else { $dc }
                $lbDCs.Items.Add($label) | Out-Null
            }
        }
        foreach ($acc in @($locked | Where-Object LockedOut)) {
            $key=$acc.SamAccountName
            if (-not $script:LastAlerted.ContainsKey($key)) {
                $script:LastAlerted[$key]=(Get-Date)
                $msg="LOCKOUT DETECTED: $($acc.Name) ($key) at $(Get-Date -Format 'HH:mm:ss') on $env:COMPUTERNAME"
                Send-TeamsAlert $msg; Send-EmailAlert "AD Lockout: $key" $msg
                $script:AlertCount++; $script:S_Alerts.Text=$script:AlertCount.ToString()
            }
        }
        Set-Status "Found $(@($locked).Count) locked account(s)." $(if (@($locked|Where-Object LockedOut).Count -gt 0){$C.Danger}else{$C.Success})
    } catch { Set-Status "Failed to query locked accounts: $_" $C.Danger }
    $form.Cursor=[System.Windows.Forms.Cursors]::Default
})

# Grid selection -> detail pane
$gridLocked.Add_SelectionChanged({
    if ($gridLocked.SelectedRows.Count -eq 0) { return }
    $sam=$gridLocked.SelectedRows[0].Cells["SAM"].Value
    try {
        $u=Get-ADUser $sam -Properties * -ErrorAction SilentlyContinue
        if ($u) {
            $script:rtDetail.Clear()
            $script:rtDetail.SelectionColor=$C.Accent
            $script:rtDetail.AppendText("  $($u.Name)  ($sam)`n")
            $script:rtDetail.SelectionColor=$C.TextPrimary
            @(
                "  Dept / Title   : $($u.Department) / $($u.Title)"
                "  Email          : $($u.EmailAddress)"
                "  Locked Out     : $($u.LockedOut)    Bad Logons: $($u.BadLogonCount)    Last Bad: $($u.LastBadPasswordAttempt)"
                "  Password Set   : $($u.PasswordLastSet)    Never Expires: $($u.PasswordNeverExpires)"
                "  Last Logon     : $($u.LastLogonDate)"
                "  Enabled        : $($u.Enabled)"
                "  OU             : $(($u.DistinguishedName -split ',',2)[1])"
            ) | ForEach-Object { $script:rtDetail.AppendText("$_`n") }
        }
    } catch {}
})

# Unlock selected
$btnUnlock.Add_Click({
    if ($gridLocked.SelectedRows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Select an account first.","Unlock","OK","Information")|Out-Null; return }
    $sam=$gridLocked.SelectedRows[0].Cells["SAM"].Value
    if (([System.Windows.Forms.MessageBox]::Show("Unlock account: $sam ?","Confirm","YesNo","Question")) -ne "Yes") { return }
    if (Invoke-UnlockAccount $sam) {
        $gridLocked.SelectedRows[0].Cells["Status"].Value="OK"
        $gridLocked.SelectedRows[0].DefaultCellStyle.ForeColor=$C.Success
        $script:UnlockCount++; $script:S_Unlocks.Text=$script:UnlockCount.ToString()
        Write-AuditEntry "UNLOCK" $sam "Manual unlock" $env:USERNAME
        if ($script:LastAlerted.ContainsKey($sam)) { $script:LastAlerted.Remove($sam) }
        Set-Status "Account '$sam' unlocked successfully." $C.Success
    } else { Set-Status "Failed to unlock '$sam'. Check permissions." $C.Danger }
})

# Unlock all
$btnUnlockAll.Add_Click({
    $rows=@($gridLocked.Rows | Where-Object { $_.Cells["Status"].Value -eq "LOCKED" })
    if ($rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No locked accounts in list.","Unlock All","OK","Information")|Out-Null; return }
    if (([System.Windows.Forms.MessageBox]::Show("Unlock ALL $($rows.Count) locked account(s)?`nThis will be logged.","Confirm Bulk Unlock","YesNo","Warning")) -ne "Yes") { return }
    $n=0
    foreach ($row in $rows) {
        $sam=$row.Cells["SAM"].Value
        if (Invoke-UnlockAccount $sam) {
            $row.Cells["Status"].Value="OK"; $row.DefaultCellStyle.ForeColor=$C.Success
            $n++; $script:UnlockCount++
            Write-AuditEntry "BULK UNLOCK" $sam "Bulk unlock" $env:USERNAME
            if ($script:LastAlerted.ContainsKey($sam)) { $script:LastAlerted.Remove($sam) }
        }
    }
    $script:S_Unlocks.Text=$script:UnlockCount.ToString()
    Set-Status "Bulk unlock done. $n account(s) unlocked." $C.Success
})

# Investigate from locked tab
$btnInvestigate.Add_Click({
    if ($gridLocked.SelectedRows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Select an account first.","Investigate","OK","Information")|Out-Null; return }
    $tbInvUser.Text=$gridLocked.SelectedRows[0].Cells["SAM"].Value
    $tabs.SelectedIndex=1; $btnInvGo.PerformClick()
})

# Copy username
$btnCopyUser.Add_Click({
    if ($gridLocked.SelectedRows.Count -eq 0) { return }
    $sam=$gridLocked.SelectedRows[0].Cells["SAM"].Value
    [System.Windows.Forms.Clipboard]::SetText($sam)
    Set-Status "Copied '$sam' to clipboard." $C.Success
})

# Investigation go
$btnInvGo.Add_Click({
    $sam=$tbInvUser.Text.Trim()
    if (-not $sam) { return }
    Set-Status "Investigating '$sam'..." $C.Accent
    $form.Cursor=[System.Windows.Forms.Cursors]::WaitCursor
    $gridSources.Rows.Clear()
    try {
        $src=Get-LockoutSource $sam
        foreach ($s in $src) { $gridSources.Rows.Add($s.Time.ToString("MM/dd HH:mm:ss"),$s.Username,$s.SourceMachine,$s.DC)|Out-Null }
        if ($src.Count -gt 0) { $tbSrcMachine.Text=$src[0].SourceMachine; $tbSrcMachine.ForeColor=$C.TextPrimary }
        Write-AuditEntry "INVESTIGATE" $sam "Lockout source query" $env:USERNAME
        Set-Status "Found $($src.Count) lockout event(s) for '$sam'." $(if($src.Count -gt 0){$C.Warning}else{$C.Success})
    } catch { Set-Status "Investigation failed: $_" $C.Danger }
    $form.Cursor=[System.Windows.Forms.Cursors]::Default
})

# Fetch processes from source machine
$btnFetchProcs.Add_Click({
    $sam=$tbInvUser.Text.Trim(); $machine=$tbSrcMachine.Text.Trim()
    if (-not $machine -or $machine -eq $tbSrcMachine.Tag) { [System.Windows.Forms.MessageBox]::Show("Enter source machine name first.","Fetch","OK","Information")|Out-Null; return }
    Set-Status "Fetching 4625 events from '$machine'..." $C.Accent
    $form.Cursor=[System.Windows.Forms.Cursors]::WaitCursor
    $gridProcs.Rows.Clear()
    try {
        $procs=Get-ProcessSource $machine $sam
        foreach ($p in $procs) { $gridProcs.Rows.Add($p.Time.ToString("MM/dd HH:mm:ss"),$p.Process,"Type $($p.LogonType)",$p.SourceIP)|Out-Null }
        Set-Status "Found $($procs.Count) failed logon event(s) on '$machine'." $C.Warning
    } catch { Set-Status "Could not reach '$machine': $_" $C.Danger }
    $form.Cursor=[System.Windows.Forms.Cursors]::Default
})

# Load password policy
$btnLoadPolicy.Add_Click({
    Set-Status "Loading password policy..." $C.Accent
    $form.Cursor=[System.Windows.Forms.Cursors]::WaitCursor
    $script:rtPolicy.Clear()
    try {
        $pol=Get-PasswordPolicy
        if ($pol) {
            $d=$pol.Domain
            $script:rtPolicy.SelectionColor=$C.Accent
            $script:rtPolicy.AppendText("=================================================`n  DEFAULT DOMAIN PASSWORD POLICY`n=================================================`n")
            @(
                "  Min Password Length          : $($d.MinPasswordLength) characters"
                "  Min Password Age             : $($d.MinPasswordAge.Days) days"
                "  Max Password Age             : $($d.MaxPasswordAge.Days) days"
                "  Password History Count       : $($d.PasswordHistoryCount)"
                "  Complexity Requirements      : $($d.ComplexityEnabled)"
                "  Reversible Encryption        : $($d.ReversibleEncryptionEnabled)"
                "  Lockout Threshold            : $($d.LockoutThreshold) bad attempts"
                "  Lockout Duration             : $($d.LockoutDuration.Minutes) minutes"
                "  Lockout Observation Window   : $($d.LockoutObservationWindow.Minutes) minutes"
            ) | ForEach-Object {
                $script:rtPolicy.SelectionColor=if($_-like"*Lockout*"){$C.Warning}else{$C.TextPrimary}
                $script:rtPolicy.AppendText("$_`n")
            }
            if ($pol.FGPP.Count -gt 0) {
                $script:rtPolicy.SelectionColor=$C.Accent
                $script:rtPolicy.AppendText("`n=================================================`n  FINE-GRAINED PASSWORD POLICIES`n=================================================`n")
                foreach ($fg in $pol.FGPP) {
                    $script:rtPolicy.SelectionColor=$C.Warning
                    $script:rtPolicy.AppendText("`n  Policy: $($fg.Name)  (Precedence: $($fg.Precedence))`n")
                    $script:rtPolicy.SelectionColor=$C.TextPrimary
                    @("  Applied To : $($fg.AppliesTo -join ', ')","  Min Pwd Len: $($fg.MinPasswordLength)","  Lockout Threshold: $($fg.LockoutThreshold)","  Lockout Duration : $($fg.LockoutDuration.Minutes) min","  Complexity: $($fg.ComplexityEnabled)") |
                        ForEach-Object { $script:rtPolicy.AppendText("$_`n") }
                }
            } else {
                $script:rtPolicy.SelectionColor=$C.TextMuted
                $script:rtPolicy.AppendText("`n  No Fine-Grained Password Policies found.`n")
            }
        }
        Set-Status "Password policy loaded." $C.Success
    } catch { Set-Status "Failed to load policy: $_" $C.Danger }
    $form.Cursor=[System.Windows.Forms.Cursors]::Default
})

# Monitor toggle
$monToggle.Toggle.Add_CheckedChanged({
    if ($monToggle.Toggle.Checked) {
        $script:MonitorTimer=New-Object System.Windows.Forms.Timer
        $script:MonitorTimer.Interval=$script:Config.RefreshInterval*1000
        $script:MonitorTimer.Add_Tick({ $btnSearchLocked.PerformClick() })
        $script:MonitorTimer.Start()
        $script:lblMonStatus.Text="[ MONITORING ]"; $script:lblMonStatus.ForeColor=$C.Success
        Set-Status "Monitor enabled. Refreshing every $($script:Config.RefreshInterval)s." $C.Success
        $btnSearchLocked.PerformClick()
    } else {
        if ($script:MonitorTimer) { $script:MonitorTimer.Stop(); $script:MonitorTimer.Dispose(); $script:MonitorTimer=$null }
        $script:lblMonStatus.Text="[ IDLE ]"; $script:lblMonStatus.ForeColor=$C.TextMuted
        Set-Status "Monitor disabled." $C.TextMuted
    }
})

$btnApply.Add_Click({
    $s=0; if ([int]::TryParse($tbInterval.Text,[ref]$s) -and $s -ge 10) {
        $script:Config.RefreshInterval=$s
        if ($script:MonitorTimer) { $script:MonitorTimer.Interval=$s*1000 }
        Set-Status "Refresh interval set to ${s}s." $C.Success
    } else { Set-Status "Invalid interval - minimum 10 seconds." $C.Danger }
})

$teamsToggle.Toggle.Add_CheckedChanged({ $script:Config.TeamsEnabled=$teamsToggle.Toggle.Checked })
$emailToggle.Toggle.Add_CheckedChanged({ $script:Config.EmailEnabled=$emailToggle.Toggle.Checked })
$btnConfigNotif.Add_Click({ $tabs.SelectedIndex=4 })

$btnViewAudit.Add_Click({
    $gridAudit.Rows.Clear()
    foreach ($e in $script:AuditLog) { $gridAudit.Rows.Add($e.Timestamp,$e.Operator,$e.Action,$e.Username,$e.Details)|Out-Null }
    $tabs.SelectedIndex=3
})

$btnExportAudit.Add_Click({
    $dlg=New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter="CSV Files (*.csv)|*.csv"; $dlg.FileName="ADLockoutAudit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if ($dlg.ShowDialog() -eq "OK") {
        try { $script:AuditLog|Export-Csv $dlg.FileName -NoTypeInformation; Set-Status "Exported to $($dlg.FileName)." $C.Success }
        catch { Set-Status "Export failed: $_" $C.Danger }
    }
})

$btnRefreshDCs.Add_Click({
    $lbDCs.Items.Clear()
    $dcs = Get-AllDCs
    $pdc = Get-PDCEmulator
    # Strip hostname from FQDN for comparison
    $pdcShort = ($pdc -split '\.')[0]
    foreach ($dc in $dcs) {
        $dcShort = ($dc -split '\.')[0]
        if ($dcShort -eq $pdcShort -or $dc -eq $pdc) {
            $lbDCs.Items.Add("$dc  [PDC]") | Out-Null
        } else {
            $lbDCs.Items.Add($dc) | Out-Null
        }
    }
    $script:S_DCs.Text = $dcs.Count.ToString()
    Set-Status "Found $($dcs.Count) DC(s). PDC Emulator: $pdc" $C.Success
})

$btnSaveSet.Add_Click({
    $script:Config.TeamsWebhook=$tbTeamsWH.Text; $script:Config.EmailFrom=$tbEmailFr.Text
    $script:Config.EmailTo=$tbEmailTo.Text; $script:Config.SmtpServer=$tbSmtp.Text
    $script:Config.AuditLogPath=$tbAuditPath.Text
    $p=587; [int]::TryParse($tbSmtpPort.Text,[ref]$p)|Out-Null; $script:Config.SmtpPort=$p
    $mr=500; [int]::TryParse($tbMaxRows.Text,[ref]$mr)|Out-Null; $script:Config.MaxHistoryRows=$mr
    Set-Status "Settings saved." $C.Success
})

$btnTestTeam.Add_Click({
    $script:Config.TeamsWebhook=$tbTeamsWH.Text
    Send-TeamsAlert "AD Lockout Manager - Teams test from $env:COMPUTERNAME by $env:USERNAME at $(Get-Date)"
    Set-Status "Teams test sent." $C.Success
})

$btnTestMail.Add_Click({
    $script:Config.SmtpServer=$tbSmtp.Text; $script:Config.EmailFrom=$tbEmailFr.Text; $script:Config.EmailTo=$tbEmailTo.Text
    Send-EmailAlert "AD Lockout Manager - Test" "Email test from $env:COMPUTERNAME by $env:USERNAME at $(Get-Date)"
    Set-Status "Test email sent to $($script:Config.EmailTo)." $C.Success
})

# ============================================================
#  STARTUP & SHUTDOWN
# ============================================================
$form.Add_Shown({
    $ok=Test-ADModule
    if ($ok) {
        $pdc=Get-PDCEmulator
        Set-Status "Ready.  PDC: $pdc  |  Click 'Refresh DC List' to discover Domain Controllers." $C.TextMuted
    } else {
        Set-Status "Limited mode - RSAT not installed. Install the ActiveDirectory module to enable all features." $C.Warning
    }
    Resize-TabGrids
    Write-AuditEntry "LAUNCH" "" "Tool started on $env:COMPUTERNAME" $env:USERNAME
})

$form.Add_Resize({ Resize-TabGrids })

$form.Add_FormClosing({
    if ($script:MonitorTimer) { $script:MonitorTimer.Stop(); $script:MonitorTimer.Dispose() }
    Write-AuditEntry "EXIT" "" "Tool closed" $env:USERNAME
})

[System.Windows.Forms.Application]::Run($form)
