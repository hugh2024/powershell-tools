# Network Speed Checker v4.4
# Best of v3.3 + v4.2 — Full feature merge
# ─────────────────────────────────────────────────────────────
# FROM v4.2:  Traceroute | Batch Parallel | TrustedHosts | CSV Export
#             Live output streaming | DNS average | Batch dialog
# FROM v3.3:  GUI Credential Dialog (no terminal prompts) | Session cred memory
#             Authentication=Negotiate (local-admin WinRM fix) | Input validation
#             CIM session disposal | HTTPS fallback servers | $winrmCmd defined once
# ─────────────────────────────────────────────────────────────

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

$script:IsBusy    = $false
$script:Job       = $null
$script:Timer     = $null
$script:SavedCred = $null   # PSCredential stored from GUI dialog — never plain text

# ═══════════════════════════════════════════════════════════════
#region === COLOR THEME ===
$C = @{
    Header     = [Drawing.Color]::FromArgb(22, 62, 117)
    Accent     = [Drawing.Color]::FromArgb(0, 120, 212)
    BtnLocal   = [Drawing.Color]::FromArgb(0, 120, 212)
    BtnRemote  = [Drawing.Color]::FromArgb(0, 137, 123)
    BtnBatch   = [Drawing.Color]::FromArgb(81, 45, 168)
    BtnTest    = [Drawing.Color]::FromArgb(56, 142, 60)
    BtnEnable  = [Drawing.Color]::FromArgb(245, 124, 0)
    BtnCancel  = [Drawing.Color]::FromArgb(198, 40, 40)
    BtnClear   = [Drawing.Color]::FromArgb(117, 117, 117)
    BtnExport  = [Drawing.Color]::FromArgb(69, 90, 100)
    BtnCSV     = [Drawing.Color]::FromArgb(46, 125, 50)
    BtnCred    = [Drawing.Color]::FromArgb(94, 53, 177)
    StatusOK   = @([Drawing.Color]::FromArgb(232,245,233), [Drawing.Color]::FromArgb(27,94,32))
    StatusRun  = @([Drawing.Color]::FromArgb(227,242,253), [Drawing.Color]::FromArgb(13,71,161))
    StatusWarn = @([Drawing.Color]::FromArgb(255,243,224), [Drawing.Color]::FromArgb(230,81,0))
    StatusFail = @([Drawing.Color]::FromArgb(255,235,238), [Drawing.Color]::FromArgb(183,28,28))
    CardBg     = [Drawing.Color]::White
    FormBg     = [Drawing.Color]::FromArgb(243, 246, 249)
    LabelDim   = [Drawing.Color]::FromArgb(100, 116, 139)
}
#endregion

# ═══════════════════════════════════════════════════════════════
#region === GUI CREDENTIAL DIALOG (v3.3) ===
# Full WinForms popup — password stored as SecureString, textbox wiped on OK.
# Replaces Get-Credential which fires in the terminal and breaks when running
# PowerShell as a local admin (no domain Kerberos context).

function Show-CredentialDialog {
    param(
        [string]$Title   = 'Enter Credentials',
        [string]$Message = 'Domain admin credentials for remote machine'
    )

    $dlg = New-Object Windows.Forms.Form -Property @{
        Text            = $Title
        Size            = New-Object Drawing.Size(400, 250)
        StartPosition   = 'CenterParent'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        BackColor       = $C.FormBg
        Font            = New-Object Drawing.Font('Segoe UI', 9.5)
    }

    $lblMsg = New-Object Windows.Forms.Label -Property @{
        Text      = $Message
        Location  = New-Object Drawing.Point(16, 14)
        Size      = New-Object Drawing.Size(360, 20)
        ForeColor = [Drawing.Color]::FromArgb(60,60,60)
    }
    $lblUser = New-Object Windows.Forms.Label -Property @{
        Text = 'Username:'; Location = New-Object Drawing.Point(16,46); AutoSize=$true; ForeColor=$C.LabelDim
    }
    $txtUser = New-Object Windows.Forms.TextBox -Property @{
        Location = New-Object Drawing.Point(100,43); Size = New-Object Drawing.Size(270,26)
        Font     = New-Object Drawing.Font('Segoe UI',10)
        Text     = if ($script:SavedCred) { $script:SavedCred.UserName } else { "$env:USERDOMAIN\$env:USERNAME" }
    }
    $lblPass = New-Object Windows.Forms.Label -Property @{
        Text = 'Password:'; Location = New-Object Drawing.Point(16,84); AutoSize=$true; ForeColor=$C.LabelDim
    }
    $txtPass = New-Object Windows.Forms.TextBox -Property @{
        Location = New-Object Drawing.Point(100,81); Size = New-Object Drawing.Size(240,26)
        Font     = New-Object Drawing.Font('Segoe UI',10)
        UseSystemPasswordChar = $true
    }
    # Show/Hide toggle
    $btnToggle = New-Object Windows.Forms.Button -Property @{
        Text='👁'; Location=New-Object Drawing.Point(344,80); Size=New-Object Drawing.Size(26,28)
        FlatStyle='Flat'; Cursor='Hand'; Font=New-Object Drawing.Font('Segoe UI',9)
    }
    $btnToggle.FlatAppearance.BorderSize = 1
    $btnToggle.Add_Click({
        $txtPass.UseSystemPasswordChar = -not $txtPass.UseSystemPasswordChar
        $btnToggle.Text = if ($txtPass.UseSystemPasswordChar) { '👁' } else { '🙈' }
    })

    $chkSave = New-Object Windows.Forms.CheckBox -Property @{
        Text='Remember for this session'; Location=New-Object Drawing.Point(16,118); AutoSize=$true
        Checked=$true; ForeColor=$C.LabelDim; Font=New-Object Drawing.Font('Segoe UI',8.5)
    }

    $btnOK = New-Object Windows.Forms.Button -Property @{
        Text='OK'; Location=New-Object Drawing.Point(210,165); Size=New-Object Drawing.Size(80,32)
        FlatStyle='Flat'; BackColor=$C.BtnLocal; ForeColor='White'; DialogResult='OK'
        Font=New-Object Drawing.Font('Segoe UI Semibold',9.5); Cursor='Hand'
    }
    $btnOK.FlatAppearance.BorderSize = 0

    $btnCancelCred = New-Object Windows.Forms.Button -Property @{
        Text='Cancel'; Location=New-Object Drawing.Point(300,165); Size=New-Object Drawing.Size(80,32)
        FlatStyle='Flat'; BackColor=$C.BtnClear; ForeColor='White'; DialogResult='Cancel'
        Font=New-Object Drawing.Font('Segoe UI Semibold',9.5); Cursor='Hand'
    }
    $btnCancelCred.FlatAppearance.BorderSize = 0

    $dlg.AcceptButton = $btnOK; $dlg.CancelButton = $btnCancelCred
    $dlg.Controls.AddRange(@($lblMsg,$lblUser,$txtUser,$lblPass,$txtPass,$btnToggle,$chkSave,$btnOK,$btnCancelCred))

    if ($dlg.ShowDialog() -ne 'OK') { $dlg.Dispose(); return $null }
    if (-not $txtUser.Text.Trim()) {
        [Windows.Forms.MessageBox]::Show('Username cannot be empty.','Credential Error','OK','Warning') | Out-Null
        $dlg.Dispose(); return $null
    }

    # SecureString conversion — plain text never stored in a variable
    $secPass = $txtPass.Text | ConvertTo-SecureString -AsPlainText -Force
    $txtPass.Text = ''   # wipe textbox immediately
    $cred = New-Object System.Management.Automation.PSCredential($txtUser.Text.Trim(), $secPass)

    if ($chkSave.Checked) { $script:SavedCred = $cred }
    $dlg.Dispose()
    return $cred
}
#endregion

# ═══════════════════════════════════════════════════════════════
#region === INPUT VALIDATION (v3.3) ===
# Only allow valid hostname/IP characters — blocks injection attempts
function Test-RemoteHostname([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return $false }
    return $name -match '^[A-Za-z0-9.\-_]+$'
}
#endregion

# ═══════════════════════════════════════════════════════════════
#region === BUILD GUI ===
$tt = New-Object Windows.Forms.ToolTip -Property @{ AutoPopDelay=8000; InitialDelay=300; ReshowDelay=200; ShowAlways=$true }

$form = New-Object Windows.Forms.Form -Property @{
    Text='Network Speed Checker v4.4'; Size=New-Object Drawing.Size(1100,980)
    MinimumSize=New-Object Drawing.Size(960,720); StartPosition='CenterScreen'
    BackColor=$C.FormBg; Font=New-Object Drawing.Font('Segoe UI',9.5)
}

# ── HEADER ──
$pnlHead  = New-Object Windows.Forms.Panel -Property @{ Dock='Top'; Height=50; BackColor=$C.Header }
$lblTitle = New-Object Windows.Forms.Label -Property @{
    Text='  Network Speed Checker v4.4'; ForeColor='White'
    Font=New-Object Drawing.Font('Segoe UI Semibold',15); AutoSize=$true; Location=New-Object Drawing.Point(10,12)
}
$lblSub = New-Object Windows.Forms.Label -Property @{
    Text='Cloudflare CDN | Traceroute | Batch | CSV Export | GUI Credentials'
    ForeColor=[Drawing.Color]::FromArgb(170,195,225)
    Font=New-Object Drawing.Font('Segoe UI',8.5); AutoSize=$true; Location=New-Object Drawing.Point(370,18)
}
$pnlHead.Controls.AddRange(@($lblTitle,$lblSub))

# ── STATUS BAR ──
$pnlStatus = New-Object Windows.Forms.Panel -Property @{
    Location=New-Object Drawing.Point(16,58); Size=New-Object Drawing.Size(1050,30)
    BackColor=$C.StatusOK[0]; Anchor='Top,Left,Right'
}
$lblStatus = New-Object Windows.Forms.Label -Property @{
    Text='  Ready'; Dock='Fill'; ForeColor=$C.StatusOK[1]
    Font=New-Object Drawing.Font('Segoe UI Semibold',10.5); TextAlign='MiddleLeft'
}
$pnlStatus.Controls.Add($lblStatus)

# ── REMOTE TARGET CARD ──
# Includes: Set Credentials button (GUI popup), cred status label, Auto TrustedHosts
$pnlRemote = New-Object Windows.Forms.Panel -Property @{
    BorderStyle='FixedSingle'; Location=New-Object Drawing.Point(16,96)
    Size=New-Object Drawing.Size(1050,52); BackColor=$C.CardBg; Anchor='Top,Left,Right'
}
$lblRmt    = New-Object Windows.Forms.Label   -Property @{ Text='Remote:'; Location=New-Object Drawing.Point(12,16); AutoSize=$true; ForeColor=$C.LabelDim }
$txtRemote = New-Object Windows.Forms.TextBox -Property @{ Location=New-Object Drawing.Point(75,12); Size=New-Object Drawing.Size(240,26); Font=New-Object Drawing.Font('Segoe UI',10) }
$tt.SetToolTip($txtRemote, "Single: PC001.domain.com`nBatch: PC001, PC002, PC003 (comma-separated)`nBatch button opens a multi-line entry dialog.")

$chkCred = New-Object Windows.Forms.CheckBox -Property @{ Text='Alt Creds'; Location=New-Object Drawing.Point(328,16); AutoSize=$true }
$tt.SetToolTip($chkCred, "Use alternate domain credentials for remote operations.`nClick 'Set Credentials' to enter them via GUI popup.`nLeave unchecked to use your current session identity.")

# Set Credentials — opens GUI dialog instead of terminal Get-Credential
$btnSetCred = New-Object Windows.Forms.Button -Property @{
    Text='Set Credentials'; Location=New-Object Drawing.Point(415,10); Size=New-Object Drawing.Size(118,30)
    FlatStyle='Flat'; BackColor=$C.BtnCred; ForeColor='White'
    Font=New-Object Drawing.Font('Segoe UI Semibold',8.5); Cursor='Hand'
}
$btnSetCred.FlatAppearance.BorderSize = 0
$tt.SetToolTip($btnSetCred, "Open popup to enter or update domain admin credentials.`nCredentials are held in memory only — never written to disk.")

$lblCredStatus = New-Object Windows.Forms.Label -Property @{
    Text='No credentials set'; Location=New-Object Drawing.Point(540,17); AutoSize=$true
    ForeColor=$C.LabelDim; Font=New-Object Drawing.Font('Segoe UI',8.5)
}

$chkTrusted = New-Object Windows.Forms.CheckBox -Property @{ Text='Auto TrustedHosts'; Location=New-Object Drawing.Point(730,16); AutoSize=$true }
$tt.SetToolTip($chkTrusted, "Auto-add remote to WinRM TrustedHosts before connecting.`nFixes auth errors for non-domain or cross-domain machines.`nRequires local admin rights.")

$lblWR        = New-Object Windows.Forms.Label    -Property @{ Text='WinRM:'; Location=New-Object Drawing.Point(880,16); AutoSize=$true; ForeColor=$C.LabelDim }
$cmbTransport = New-Object Windows.Forms.ComboBox -Property @{ Location=New-Object Drawing.Point(930,12); Size=New-Object Drawing.Size(72,26); DropDownStyle='DropDownList' }
[void]$cmbTransport.Items.AddRange(@('HTTP','HTTPS')); $cmbTransport.SelectedIndex=0
$tt.SetToolTip($cmbTransport, "HTTP = port 5985 (default, unencrypted)`nHTTPS = port 5986 (SSL/TLS encrypted)")

$pnlRemote.Controls.AddRange(@($lblRmt,$txtRemote,$chkCred,$btnSetCred,$lblCredStatus,$chkTrusted,$lblWR,$cmbTransport))

# Wire up Set Credentials button
$btnSetCred.Add_Click({
    $cred = Show-CredentialDialog -Title 'Set Credentials' -Message 'Enter domain admin credentials for remote operations:'
    if ($cred) {
        $script:SavedCred        = $cred
        $chkCred.Checked         = $true
        $lblCredStatus.Text      = "✔ Set: $($cred.UserName)"
        $lblCredStatus.ForeColor = [Drawing.Color]::FromArgb(27,94,32)
    }
})
# Unchecking Alt Creds clears saved credential from memory
$chkCred.Add_CheckedChanged({
    if (-not $chkCred.Checked) {
        $script:SavedCred        = $null
        $lblCredStatus.Text      = 'No credentials set'
        $lblCredStatus.ForeColor = $C.LabelDim
    }
})

# ── SETTINGS CARDS ──
$pnlHealth = New-Object Windows.Forms.Panel -Property @{ BorderStyle='FixedSingle'; Location=New-Object Drawing.Point(16,156); Size=New-Object Drawing.Size(510,115); BackColor=$C.CardBg }
$lblHT     = New-Object Windows.Forms.Label -Property @{ Text='HEALTH THRESHOLDS'; Location=New-Object Drawing.Point(12,8); AutoSize=$true; ForeColor=$C.Accent; Font=New-Object Drawing.Font('Segoe UI Semibold',8.5) }
function New-SettingRow($parent,$label,$tip,$y,$min,$max,$val,$dec=0) {
    $lbl = New-Object Windows.Forms.Label         -Property @{ Text=$label; Location=New-Object Drawing.Point(12,$y); AutoSize=$true; ForeColor=$C.LabelDim }
    $num = New-Object Windows.Forms.NumericUpDown -Property @{ Location=New-Object Drawing.Point(210,$($y-3)); Size=New-Object Drawing.Size(80,26); Minimum=$min; Maximum=$max; DecimalPlaces=$dec; Value=$val }
    $tt.SetToolTip($num,$tip); $tt.SetToolTip($lbl,$tip)
    $parent.Controls.AddRange(@($lbl,$num)); $num
}
$numMinDL = New-SettingRow $pnlHealth 'Min Download (Mbps)' "Flag if download < this.`n0 = disable."  32 0 10000 25 1
$numMinUL = New-SettingRow $pnlHealth 'Min Upload (Mbps)'   "Flag if upload < this.`n0 = disable."    62 0 10000 4  1
$numMaxPL = New-SettingRow $pnlHealth 'Max Packet Loss (%)' "Flag if loss > this.`n100 = disable."   92 0 100   2  1
$pnlHealth.Controls.Add($lblHT)

$pnlConfig = New-Object Windows.Forms.Panel -Property @{ BorderStyle='FixedSingle'; Location=New-Object Drawing.Point(544,156); Size=New-Object Drawing.Size(522,115); BackColor=$C.CardBg; Anchor='Top,Left,Right' }
$lblCF     = New-Object Windows.Forms.Label -Property @{ Text='TEST CONFIGURATION'; Location=New-Object Drawing.Point(12,8); AutoSize=$true; ForeColor=$C.Accent; Font=New-Object Drawing.Font('Segoe UI Semibold',8.5) }
$numPingCount = New-SettingRow $pnlConfig 'Pings per Target'  "ICMP pings per target. More = accurate."     32 1 50  5  0
$numDLSize    = New-SettingRow $pnlConfig 'DL Test Size (MB)' "Max DL payload MB. Bigger = better accuracy." 62 1 100 25 0
$numULSize    = New-SettingRow $pnlConfig 'UL Test Size (MB)' "Max UL payload MB."                           92 1 50  10 0
$chkTrace = New-Object Windows.Forms.CheckBox -Property @{ Text='Traceroute to:'; Location=New-Object Drawing.Point(320,32); AutoSize=$true }
$tt.SetToolTip($chkTrace, "Include traceroute. Shows each hop with latency + rDNS.`nAdds ~15-30 seconds to test time.")
$txtTraceTarget = New-Object Windows.Forms.TextBox -Property @{ Text='8.8.8.8'; Location=New-Object Drawing.Point(430,29); Size=New-Object Drawing.Size(80,24); Font=New-Object Drawing.Font('Segoe UI',9) }
$tt.SetToolTip($txtTraceTarget, "Traceroute destination. Default: 8.8.8.8`nExamples: 1.1.1.1, your-dc.domain.com")
$lblPar = New-Object Windows.Forms.Label -Property @{ Text='Parallel Jobs:'; Location=New-Object Drawing.Point(320,67); AutoSize=$true; ForeColor=$C.LabelDim }
$numParallel = New-Object Windows.Forms.NumericUpDown -Property @{ Location=New-Object Drawing.Point(430,64); Size=New-Object Drawing.Size(55,24); Minimum=1; Maximum=20; Value=5 }
$tt.SetToolTip($numParallel, "Max simultaneous machines in Batch mode.`nHigher = faster but more network/CPU load.")
$tt.SetToolTip($lblPar,      "Max simultaneous machines in Batch mode.")
$pnlConfig.Controls.AddRange(@($lblCF,$chkTrace,$txtTraceTarget,$lblPar,$numParallel))

# ── ACTION BUTTONS ──
$pnlBtn = New-Object Windows.Forms.Panel -Property @{ Location=New-Object Drawing.Point(16,280); Size=New-Object Drawing.Size(1050,44); Anchor='Top,Left,Right' }
function New-ActionBtn($text,$x,$w,$color,$tip) {
    $b = New-Object Windows.Forms.Button -Property @{
        Text=$text; Location=New-Object Drawing.Point($x,0); Size=New-Object Drawing.Size($w,40)
        FlatStyle='Flat'; BackColor=$color; ForeColor='White'
        Font=New-Object Drawing.Font('Segoe UI Semibold',9.5); Cursor='Hand'
    }
    $b.FlatAppearance.BorderSize=0
    $b.FlatAppearance.MouseOverBackColor=[Drawing.Color]::FromArgb(
        [Math]::Min($color.R+25,255),[Math]::Min($color.G+25,255),[Math]::Min($color.B+25,255))
    $tt.SetToolTip($b,$tip); $b
}
$btnLocal  = New-ActionBtn 'Run Local'    0   108 $C.BtnLocal  "Run all tests on THIS computer:`nAdapter info, public IP, ping, speed, traceroute, DNS."
$btnRemote = New-ActionBtn 'Run Remote'   116 116 $C.BtnRemote "Run all tests on the REMOTE computer via WinRM."
$btnBatch  = New-ActionBtn 'Batch Test'   240 108 $C.BtnBatch  "Test MULTIPLE machines in PARALLEL.`nOpens a multi-line machine entry dialog."
$btnTest   = New-ActionBtn 'Test WinRM'   356 108 $C.BtnTest   "Check WinRM on remote: DNS, TCP, WSMan, TrustedHosts."
$btnEnable = New-ActionBtn 'Enable WinRM' 472 120 $C.BtnEnable "Enable WinRM on remote via WMI/CIM or PsExec.`nRequires admin rights on the remote machine."
$btnCancel = New-ActionBtn 'Cancel'       600 80  $C.BtnCancel "Stop the current test."
$btnClear  = New-ActionBtn 'Clear'        688 80  $C.BtnClear  "Clear the output window."
$btnExport = New-ActionBtn 'Export TXT'   776 92  $C.BtnExport "Save output as a text file (CSV lines stripped)."
$btnCSV    = New-ActionBtn 'Export CSV'   876 92  $C.BtnCSV    "Export results as CSV for tracking over time."
$btnCancel.Enabled = $false
$pnlBtn.Controls.AddRange(@($btnLocal,$btnRemote,$btnBatch,$btnTest,$btnEnable,$btnCancel,$btnClear,$btnExport,$btnCSV))

# ── OUTPUT ──
$pnlOut = New-Object Windows.Forms.Panel -Property @{
    BorderStyle='FixedSingle'; Location=New-Object Drawing.Point(16,332)
    Size=New-Object Drawing.Size(1050,595); Anchor='Top,Left,Right,Bottom'; BackColor=$C.CardBg
}
$lblOut = New-Object Windows.Forms.Label -Property @{ Text='OUTPUT'; Location=New-Object Drawing.Point(12,8); AutoSize=$true; ForeColor=$C.Accent; Font=New-Object Drawing.Font('Segoe UI Semibold',8.5) }
$txtOut = New-Object Windows.Forms.TextBox -Property @{
    Multiline=$true; ScrollBars='Both'; ReadOnly=$true; WordWrap=$false
    Font=New-Object Drawing.Font('Consolas',10); BackColor=[Drawing.Color]::FromArgb(252,253,254)
    Location=New-Object Drawing.Point(10,28); Size=New-Object Drawing.Size(1028,556); Anchor='Top,Left,Right,Bottom'
}
$pnlOut.Controls.AddRange(@($lblOut,$txtOut))
$form.Controls.AddRange(@($pnlHead,$pnlStatus,$pnlRemote,$pnlHealth,$pnlConfig,$pnlBtn,$pnlOut))
#endregion

# ═══════════════════════════════════════════════════════════════
#region === STATUS HELPERS ===
function Set-UIBusy([bool]$b) {
    $script:IsBusy=$b
    foreach($btn in @($btnLocal,$btnRemote,$btnBatch,$btnTest,$btnEnable,$btnClear,$btnExport,$btnCSV)){$btn.Enabled=!$b}
    $btnCancel.Enabled=$b
    if($b){ $pnlStatus.BackColor=$C.StatusRun[0]; $lblStatus.ForeColor=$C.StatusRun[1] }
}
function Set-Done([string]$text,[string]$level) {
    $lblStatus.Text="  $text"
    $colors = switch($level){ 'OK'{$C.StatusOK} 'WARN'{$C.StatusWarn} 'FAIL'{$C.StatusFail} }
    $pnlStatus.BackColor=$colors[0]; $lblStatus.ForeColor=$colors[1]
}
#endregion

# ═══════════════════════════════════════════════════════════════
#region === ASYNC JOB RUNNER ===
function Start-Test([scriptblock]$Code,[hashtable]$Params,[string]$Label) {
    if ($script:IsBusy) { return }
    Set-UIBusy $true; $txtOut.Clear()
    $lblStatus.Text = "  $Label"

    $script:Job = Start-Job -ScriptBlock $Code -ArgumentList @(
        $Params.MinDL, $Params.MinUL, $Params.MaxPL,
        $Params.Remote, $Params.Transport, $Params.Cred, $Params.ScriptText,
        $Params.PingCount, $Params.DLSize, $Params.ULSize,
        $Params.DoTrace, $Params.AutoTrust, $Params.TraceTarget, $Params.MaxParallel
    )
    $startTick = [Environment]::TickCount

    if ($script:Timer) { $script:Timer.Stop(); $script:Timer.Dispose() }
    $script:Timer = New-Object Windows.Forms.Timer -Property @{ Interval = 500 }
    $script:Timer.Add_Tick({
        $sec = [int][Math]::Floor(([Environment]::TickCount - $startTick) / 1000)
        $m = [int][Math]::Floor($sec/60); $s = [int]($sec % 60)
        $lblStatus.Text = "  $Label ($m`:$($s.ToString('00')))"

        # Stream partial output while running (v4.2 feature)
        try {
            $partial = Receive-Job -Job $script:Job -Keep -ErrorAction SilentlyContinue
            if ($partial) {
                $current = ($partial | Out-String).TrimEnd()
                if ($current -ne $txtOut.Text) {
                    $txtOut.Text = $current
                    $txtOut.SelectionStart = $txtOut.TextLength
                    $txtOut.ScrollToCaret()
                }
            }
        } catch {}

        if ($script:Job.State -ne 'Running') {
            $script:Timer.Stop(); $script:Timer.Dispose(); $script:Timer = $null
            $output = $null; $errText = $null
            try { $output = Receive-Job -Job $script:Job -ErrorAction Stop }
            catch { $errText = $_.Exception.Message }
            if ($script:Job.State -eq 'Failed') {
                try { $errText = $script:Job.ChildJobs[0].JobStateInfo.Reason.Message } catch {}
            }
            Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue
            $script:Job = $null; Set-UIBusy $false

            if ($errText -and -not $output) { Set-Done 'Failed' 'FAIL'; $txtOut.Text = "ERROR:`r`n$errText"; return }
            if (-not $output) { Set-Done 'No result' 'FAIL'; $txtOut.Text = 'No output from job.'; return }

            $text = ($output | Out-String).Trim()
            $txtOut.Text = $text
            $txtOut.SelectionStart = $txtOut.TextLength; $txtOut.ScrollToCaret()

            if ($text -match 'ISSUE\(S\) FOUND|RESULT: FAILED|RESULT:\s+NOT READY') { Set-Done 'Issues Found' 'WARN' }
            elseif ($text -match 'ALL CHECKS PASSED|RESULT: SUCCESS|RESULT:\s+READY|BATCH COMPLETE') { Set-Done 'Completed' 'OK' }
            else { Set-Done 'Completed' 'OK' }
        }
    })
    $script:Timer.Start()
}
#endregion

# ═══════════════════════════════════════════════════════════════
#region === TEST SCRIPTS ===

$LocalTestCode = {
    param([double]$MinDL,[double]$MinUL,[double]$MaxPL,$u1,$u2,$u3,$u4,
          [int]$PingCount,[int]$DLSizeMB,[int]$ULSizeMB,[bool]$DoTrace,$u5,[string]$TraceTarget,$u6)
    if ($PingCount  -le 0) { $PingCount  = 5  }
    if ($DLSizeMB   -le 0) { $DLSizeMB   = 25 }
    if ($ULSizeMB   -le 0) { $ULSizeMB   = 10 }
    if ([string]::IsNullOrWhiteSpace($TraceTarget)) { $TraceTarget = '8.8.8.8' }
    $ProgressPreference = 'SilentlyContinue'
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.ServicePointManager]::SecurityProtocol } catch {}
    $issues = [System.Collections.ArrayList]::new()

    Write-Output "============================================"
    Write-Output "  NETWORK SPEED CHECK"
    Write-Output "  Computer: $env:COMPUTERNAME"
    Write-Output "  Time:     $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "============================================"

    # ── ADAPTER INFO ──
    Write-Output ""; Write-Output "--- NETWORK ADAPTER INFO ---"
    try {
        $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Sort-Object RouteMetric | Select-Object -First 1
        if ($route) {
            $ifIdx = $route.InterfaceIndex
            Write-Output "  Gateway:        $($route.NextHop)"
            $ip = Get-NetIPAddress -InterfaceIndex $ifIdx -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($ip) { Write-Output "  Private IP:     $($ip.IPAddress)/$($ip.PrefixLength)" }
            $nic = Get-NetAdapter -InterfaceIndex $ifIdx -ErrorAction SilentlyContinue
            if ($nic) { Write-Output "  Adapter:        $($nic.Name) ($($nic.LinkSpeed))" }
            $dns = Get-DnsClientServerAddress -InterfaceIndex $ifIdx -AddressFamily IPv4 -ErrorAction SilentlyContinue
            if ($dns.ServerAddresses) { Write-Output "  DNS Servers:    $($dns.ServerAddresses -join ', ')" }
        }
    } catch { Write-Output "  Error: $($_.Exception.Message)" }

    # ── PUBLIC IP ──
    Write-Output ""; Write-Output "--- PUBLIC IP & LOCATION ---"
    try {
        $ipData = Invoke-RestMethod 'http://ip-api.com/json/' -TimeoutSec 8 -ErrorAction Stop
        Write-Output "  Public IP:      $($ipData.query)"
        Write-Output "  Location:       $($ipData.city), $($ipData.regionName), $($ipData.country)"
        Write-Output "  ISP:            $($ipData.isp)"
        Write-Output "  Organization:   $($ipData.org)"
    } catch {
        try { $pub=(Invoke-RestMethod 'https://api.ipify.org' -TimeoutSec 5).Trim(); Write-Output "  Public IP:      $pub" }
        catch { Write-Output "  Public IP:      FAILED" }
    }

    # ── PING STABILITY ──
    Write-Output ""; Write-Output "--- PING STABILITY ($PingCount pings per target) ---"
    $pingTargets = @(
        @{Name='Google DNS'; Host='8.8.8.8'},
        @{Name='Cloudflare'; Host='1.1.1.1'},
        @{Name='Google.com'; Host='google.com'}
    )
    $totalSent=0; $totalLost=0; $allAvg=0; $avgCount=0
    foreach ($tgt in $pingTargets) {
        try {
            $pings = Test-Connection -ComputerName $tgt.Host -Count $PingCount -ErrorAction SilentlyContinue
            if ($pings) {
                $times = @($pings | ForEach-Object {
                    if ($null -ne $_.Latency) { $_.Latency } elseif ($null -ne $_.ResponseTime) { $_.ResponseTime } else { -1 }
                } | Where-Object { $_ -ge 0 })
                $recv=$times.Count; $lost=$PingCount-$recv
                if ($recv -gt 0) {
                    $filtered = @($times | Where-Object { $_ -lt 500 })
                    if ($filtered.Count -eq 0) { $filtered = $times }
                    $avg = [Math]::Round(($filtered | Measure-Object -Average).Average,1)
                    $mn  = [Math]::Round(($filtered | Measure-Object -Minimum).Minimum,1)
                    $mx  = [Math]::Round(($filtered | Measure-Object -Maximum).Maximum,1)
                    $jit = [Math]::Round($mx-$mn,1)
                    $lp  = [Math]::Round($lost/$PingCount*100,0)
                    $note= if ($filtered.Count -lt $recv) {" *"} else {""}
                    Write-Output ("  {0,-26} Avg:{1,6}ms  Min:{2,6}ms  Max:{3,6}ms  Jitter:{4,6}ms  Loss:{5}%{6}" -f "$($tgt.Name) ($($tgt.Host))",$avg,$mn,$mx,$jit,$lp,$note)
                    $allAvg+=$avg; $avgCount++
                } else {
                    Write-Output ("  {0,-26} ALL PACKETS LOST" -f "$($tgt.Name) ($($tgt.Host))")
                    $lost=$PingCount
                }
                $totalSent+=$PingCount; $totalLost+=$lost
            } else {
                Write-Output ("  {0,-26} NO REPLY" -f "$($tgt.Name) ($($tgt.Host))")
                $totalSent+=$PingCount; $totalLost+=$PingCount
            }
        } catch {
            Write-Output ("  {0,-26} ERROR: {1}" -f "$($tgt.Name) ($($tgt.Host))",$_.Exception.Message)
            $totalSent+=$PingCount; $totalLost+=$PingCount
        }
    }
    $overallLoss = if ($totalSent -gt 0) { [Math]::Round($totalLost/$totalSent*100,1) } else { 100 }
    $overallAvg  = if ($avgCount  -gt 0) { [Math]::Round($allAvg/$avgCount,1) } else { 0 }
    Write-Output ""
    Write-Output "  (* = cold-cache outlier >500ms excluded)"
    if ($overallLoss -le 2)     { Write-Output "  Overall:        EXCELLENT ($overallLoss% loss, avg ${overallAvg}ms)" }
    elseif ($overallLoss -le 5) { Write-Output "  Overall:        GOOD ($overallLoss% loss, avg ${overallAvg}ms)" }
    else { Write-Output "  Overall:        POOR ($overallLoss% loss, avg ${overallAvg}ms)"; [void]$issues.Add("Ping: $overallLoss% loss") }

    # ── TRACEROUTE (optional) ──
    if ($DoTrace) {
        Write-Output ""; Write-Output "--- TRACEROUTE to $TraceTarget ---"
        try {
            $tr = Test-NetConnection -ComputerName $TraceTarget -TraceRoute -WarningAction SilentlyContinue
            if ($tr.TraceRoute) {
                $hopNum = 0
                foreach ($hop in $tr.TraceRoute) {
                    $hopNum++
                    if ($hop -eq '0.0.0.0' -or [string]::IsNullOrEmpty($hop)) {
                        Write-Output ("  {0,3}.  *  Request timed out" -f $hopNum)
                    } else {
                        $ms = '?'
                        try {
                            $p = Test-Connection -ComputerName $hop -Count 1 -ErrorAction SilentlyContinue
                            if ($p) { $ms = if ($null -ne $p.Latency) { $p.Latency } elseif ($null -ne $p.ResponseTime) { $p.ResponseTime } else { '?' } }
                        } catch {}
                        $rDns = ''
                        try { $rDns = " [$([System.Net.Dns]::GetHostEntry($hop).HostName)]" } catch {}
                        Write-Output ("  {0,3}.  {1,-16} {2,6}ms{3}" -f $hopNum,$hop,$ms,$rDns)
                    }
                }
                Write-Output "  Total hops: $hopNum"
            } else { Write-Output "  No traceroute data returned." }
        } catch { Write-Output "  Traceroute error: $($_.Exception.Message)" }
    }

    # ── SPEED TEST ──
    Write-Output ""; Write-Output "--- SPEED TEST ---"
    $cfWorks=$true; $dlSpeeds=@(); $ulSpeeds=@()
    try {
        $wc=New-Object System.Net.WebClient; $wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); $wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultCredentials
        $sw=[System.Diagnostics.Stopwatch]::StartNew(); $data=$wc.DownloadData('https://speed.cloudflare.com/__down?bytes=1000000'); $sw.Stop(); $wc.Dispose()
        $warmMbps=[Math]::Round(($data.Length*8)/$sw.Elapsed.TotalSeconds/1000000,2)
        Write-Output "  Method: Cloudflare CDN (nearest edge server)"
        Write-Output ("  DL Warmup 1MB      {0,8:N2} Mbps  (1.0 MB in {1:N2}s)" -f $warmMbps,$sw.Elapsed.TotalSeconds)
    } catch { $cfWorks=$false; Write-Output "  Cloudflare unreachable. Using HTTPS fallback servers." }

    if ($cfWorks) {
        $dlBytes=@(10000000); if ($DLSizeMB -ge 25) { $dlBytes+=25000000 }; if ($DLSizeMB -ge 50) { $dlBytes+=50000000 }
        foreach ($size in $dlBytes) {
            $label="$([int]($size/1MB))MB"
            try {
                $wc=New-Object System.Net.WebClient; $wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); $wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultCredentials
                $sw=[System.Diagnostics.Stopwatch]::StartNew(); $data=$wc.DownloadData("https://speed.cloudflare.com/__down?bytes=$size"); $sw.Stop(); $wc.Dispose()
                $mbps=[Math]::Round(($data.Length*8)/$sw.Elapsed.TotalSeconds/1000000,2); $dlSpeeds+=$mbps
                Write-Output ("  DL Test {0,-10} {1,8:N2} Mbps  ({2:N1} MB in {3:N2}s)" -f $label,$mbps,($data.Length/1MB),$sw.Elapsed.TotalSeconds)
            } catch { Write-Output ("  DL Test {0,-10} FAILED: {1}" -f $label,$_.Exception.Message) }
        }
        $ulBytes=@(5000000); if ($ULSizeMB -ge 10) { $ulBytes+=10000000 }
        foreach ($size in $ulBytes) {
            $label="$([int]($size/1MB))MB"
            try {
                $wc=New-Object System.Net.WebClient; $wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); $wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultCredentials
                $wc.Headers.Add('Content-Type','application/octet-stream'); $up=[byte[]]::new($size)
                $sw=[System.Diagnostics.Stopwatch]::StartNew(); [void]$wc.UploadData('https://speed.cloudflare.com/__up','POST',$up); $sw.Stop(); $wc.Dispose()
                $mbps=[Math]::Round(($size*8)/$sw.Elapsed.TotalSeconds/1000000,2); $ulSpeeds+=$mbps
                Write-Output ("  UL Test {0,-10} {1,8:N2} Mbps  ({2:N1} MB in {3:N2}s)" -f $label,$mbps,($size/1MB),$sw.Elapsed.TotalSeconds)
            } catch { Write-Output ("  UL Test {0,-10} FAILED: {1}" -f $label,$_.Exception.Message) }
        }
    } else {
        # HTTPS fallback servers (v3.3 security improvement — no plain HTTP)
        Write-Output "  Method: HTTPS fallback servers"
        foreach ($dl in @(
            @{Url='https://speed.hetzner.de/10MB.bin'; Name='Hetzner 10MB'},
            @{Url='https://proof.ovh.net/files/10Mb.dat'; Name='OVH 10MB'}
        )) {
            try {
                $wc=New-Object System.Net.WebClient; $wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); $wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultCredentials
                $sw=[System.Diagnostics.Stopwatch]::StartNew(); $data=$wc.DownloadData($dl.Url); $sw.Stop(); $wc.Dispose()
                $mbps=[Math]::Round(($data.Length*8)/$sw.Elapsed.TotalSeconds/1000000,2); $dlSpeeds+=$mbps
                Write-Output ("  DL {0,-18} {1,8:N2} Mbps  ({2:N1} MB in {3:N1}s)" -f $dl.Name,$mbps,($data.Length/1MB),$sw.Elapsed.TotalSeconds)
            } catch { Write-Output ("  DL {0,-18} FAILED: {1}" -f $dl.Name,$_.Exception.Message) }
        }
        try {
            $wc=New-Object System.Net.WebClient; $wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); $wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultCredentials
            $wc.Headers.Add('Content-Type','application/octet-stream'); $up=[byte[]]::new(2MB)
            $sw=[System.Diagnostics.Stopwatch]::StartNew(); [void]$wc.UploadData('https://speed.cloudflare.com/__up','POST',$up); $sw.Stop(); $wc.Dispose()
            $mbps=[Math]::Round(($up.Length*8)/$sw.Elapsed.TotalSeconds/1000000,2); $ulSpeeds+=$mbps
            Write-Output ("  UL Cloudflare 2MB    {0,8:N2} Mbps" -f $mbps)
        } catch { Write-Output "  UL Fallback 2MB      FAILED: $($_.Exception.Message)" }
    }

    $bestDL=if($dlSpeeds){[Math]::Round(($dlSpeeds|Measure-Object -Maximum).Maximum,2)}else{0}
    $bestUL=if($ulSpeeds){[Math]::Round(($ulSpeeds|Measure-Object -Maximum).Maximum,2)}else{0}
    Write-Output ""; Write-Output "  Best Download:  $bestDL Mbps"; Write-Output "  Best Upload:    $bestUL Mbps"
    if ($bestDL -eq 0) { [void]$issues.Add("Download failed") } elseif ($bestDL -lt $MinDL) { [void]$issues.Add("DL $bestDL < $MinDL Mbps") }
    if ($bestUL -eq 0) { [void]$issues.Add("Upload failed")   } elseif ($bestUL -lt $MinUL) { [void]$issues.Add("UL $bestUL < $MinUL Mbps") }
    if ($overallLoss -gt $MaxPL) { [void]$issues.Add("Loss $overallLoss% > $MaxPL%") }

    # ── DNS ──
    Write-Output ""; Write-Output "--- DNS RESOLUTION ---"
    $dnsAvg=0; $dnsCnt=0
    foreach ($d in @('google.com','microsoft.com','amazon.com')) {
        try {
            $sw=[System.Diagnostics.Stopwatch]::StartNew(); Resolve-DnsName $d -ErrorAction Stop|Out-Null; $sw.Stop()
            $ms=$sw.Elapsed.TotalMilliseconds; $dnsAvg+=$ms; $dnsCnt++
            Write-Output ("  {0,-24} {1,6:N1}ms" -f $d,$ms)
        } catch { Write-Output ("  {0,-24} FAILED" -f $d) }
    }
    if ($dnsCnt -gt 0) { Write-Output ("  Average:               {0,6:N1}ms" -f ($dnsAvg/$dnsCnt)) }

    # ── SUMMARY ──
    Write-Output ""; Write-Output "============================================"
    if ($issues.Count -eq 0) { Write-Output "  RESULT: ALL CHECKS PASSED" }
    else { Write-Output "  RESULT: $($issues.Count) ISSUE(S) FOUND"; foreach ($iss in $issues) { Write-Output "  ! $iss" } }
    Write-Output "============================================"

    # Hidden CSV line — parsed by Export CSV button, stripped by Export TXT
    Write-Output "CSV_DATA:$env:COMPUTERNAME,$bestDL,$bestUL,$overallLoss,$overallAvg,$($issues.Count),$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

# ── REMOTE TEST ──
# Authentication=Negotiate is the KEY FIX for running PS as local admin:
# Lets Windows fall back to NTLM when no Kerberos context is available.
# UseSSL only added when HTTPS transport is selected.
$RemoteTestCode = {
    param([double]$MinDL,[double]$MinUL,[double]$MaxPL,[string]$Remote,[string]$Transport,
          [pscredential]$Cred,[string]$ScriptText,[int]$PingCount,[int]$DLSize,[int]$ULSize,
          [bool]$DoTrace,[bool]$AutoTrust,[string]$TraceTarget,$u1)
    $ErrorActionPreference = 'Stop'
    $port = if ($Transport -eq 'HTTPS') { 5986 } else { 5985 }

    if ($AutoTrust) {
        Write-Output "Adding $Remote to TrustedHosts..."
        try {
            $current=(Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
            if ($current -notmatch [regex]::Escape($Remote)) {
                $new = if ($current) { "$current,$Remote" } else { $Remote }
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $new -Force
                Write-Output "  OK: $new"
            } else { Write-Output "  Already trusted." }
        } catch { Write-Output "  WARNING (need Admin): $($_.Exception.Message)" }
    }

    Write-Output "Connecting to $Remote (${Transport}:${port})..."
    $tnc = Test-NetConnection $Remote -Port $port -WarningAction SilentlyContinue
    if (-not $tnc.TcpTestSucceeded) { throw "Cannot reach $Remote on port $port. WinRM not enabled? Try 'Enable WinRM'." }
    Write-Output "Connected. Running tests..."

    $sb = [scriptblock]::Create($ScriptText)
    $ic = @{
        ComputerName = $Remote
        ScriptBlock  = $sb
        ArgumentList = @($MinDL,$MinUL,$MaxPL,$null,$null,$null,$null,$PingCount,$DLSize,$ULSize,$DoTrace,$null,$TraceTarget,$null)
        ErrorAction  = 'Stop'
    }
    if ($Transport -eq 'HTTPS') { $ic.UseSSL = $true }
    if ($Cred) {
        # Explicit credential = local admin scenario — must use Negotiate (NTLM fallback)
        # Without it, Kerberos is attempted and fails with no domain context
        $ic.Credential     = $Cred
        $ic.Authentication = 'Negotiate'
    }
    # No credential = domain admin running as themselves — let Kerberos handle it natively,
    # do NOT set Authentication or it triggers TrustedHosts enforcement over HTTP
    Invoke-Command @ic
}

# ── BATCH TEST (PARALLEL) ──
$BatchTestCode = {
    param([double]$MinDL,[double]$MinUL,[double]$MaxPL,[string]$RemoteList,[string]$Transport,
          [pscredential]$Cred,[string]$ScriptText,[int]$PingCount,[int]$DLSize,[int]$ULSize,
          [bool]$DoTrace,[bool]$AutoTrust,[string]$TraceTarget,[int]$MaxParallel)
    $ErrorActionPreference = 'Continue'
    $machines = $RemoteList -split '[,;\s]+' | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() }
    $port     = if ($Transport -eq 'HTTPS') { 5986 } else { 5985 }
    $total    = $machines.Count
    if ($MaxParallel -le 0) { $MaxParallel = 5 }

    Write-Output "============================================"
    Write-Output "  BATCH NETWORK SPEED CHECK (PARALLEL)"
    Write-Output "  Machines:     $total"
    Write-Output "  Max Parallel: $MaxParallel"
    Write-Output "  Time:         $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "============================================"

    if ($AutoTrust -and $machines.Count -gt 0) {
        Write-Output ""; Write-Output "Adding machines to TrustedHosts..."
        try {
            $current=(Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
            $toAdd=@(); foreach ($m in $machines) { if ($current -notmatch [regex]::Escape($m)) { $toAdd+=$m } }
            if ($toAdd.Count -gt 0) {
                $new = if ($current) { "$current,$($toAdd -join ',')" } else { $toAdd -join ',' }
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $new -Force
                Write-Output "  Added: $($toAdd -join ', ')"
            } else { Write-Output "  All already trusted." }
        } catch { Write-Output "  WARNING: $($_.Exception.Message)" }
    }

    $sb       = [scriptblock]::Create($ScriptText)
    $csvLines = @()
    $jobs     = @{}
    $completed=0; $pass=0; $fail=0

    $queue = [System.Collections.Queue]::new()
    foreach ($m in $machines) { $queue.Enqueue($m) }
    Write-Output ""; Write-Output "Launching parallel tests..."

    while ($queue.Count -gt 0 -or $jobs.Count -gt 0) {
        while ($queue.Count -gt 0 -and $jobs.Count -lt $MaxParallel) {
            $machine = $queue.Dequeue()
            Write-Output "  Starting: $machine"
            try {
                $ic = @{
                    ComputerName = $machine
                    ScriptBlock  = $sb
                    AsJob        = $true
                    ArgumentList = @($MinDL,$MinUL,$MaxPL,$null,$null,$null,$null,$PingCount,$DLSize,$ULSize,$DoTrace,$null,$TraceTarget,$null)
                    ErrorAction  = 'Stop'
                }
                if ($Transport -eq 'HTTPS') { $ic.UseSSL = $true }
                if ($Cred) {
                    $ic.Credential     = $Cred
                    $ic.Authentication = 'Negotiate'
                }
                $j = Invoke-Command @ic
                $jobs[$machine] = $j
            } catch {
                Write-Output "  FAILED to start $machine`: $($_.Exception.Message)"
                $csvLines += "CSV_DATA:$machine,0,0,100,0,1,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),CONNECT_ERROR"
                $completed++; $fail++
            }
        }

        $done = @()
        foreach ($kv in $jobs.GetEnumerator()) { if ($kv.Value.State -ne 'Running') { $done += $kv.Key } }

        foreach ($machine in $done) {
            $completed++
            $j = $jobs[$machine]; $jobs.Remove($machine)
            Write-Output ""; Write-Output ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
            Write-Output "  [$completed/$total] $machine (COMPLETED)"
            Write-Output ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
            try {
                $result = Receive-Job -Job $j -ErrorAction Stop
                $result | ForEach-Object { Write-Output $_ }
                $csvLine = $result | Where-Object { $_ -match '^CSV_DATA:' } | Select-Object -Last 1
                if ($csvLine) { $csvLines += $csvLine }
                if ($result -match 'ALL CHECKS PASSED') { $pass++ } else { $fail++ }
            } catch {
                Write-Output "  ERROR: $($_.Exception.Message)"
                $csvLines += "CSV_DATA:$machine,0,0,100,0,1,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),ERROR"
                $fail++
            }
            Remove-Job -Job $j -Force -ErrorAction SilentlyContinue
        }

        if ($jobs.Count -gt 0) { Start-Sleep -Milliseconds 500 }
    }

    Write-Output ""; Write-Output "============================================"
    Write-Output "  BATCH COMPLETE: $total tested | $pass passed | $fail issues"
    Write-Output "============================================"
    foreach ($line in $csvLines) { Write-Output $line }
}

# ── WINRM TEST ──
$WinRMTestCode = {
    param([double]$u1,[double]$u2,[double]$u3,[string]$Remote,[string]$Transport,$u4,$u5,$u6,$u7,$u8,$u9,$u10,$u11,$u12)
    $port = if ($Transport -eq 'HTTPS') { 5986 } else { 5985 }
    Write-Output "WinRM Connectivity Test: $Remote (${Transport}:${port})"
    Write-Output "============================================"

    try {
        $dnsResult  = Resolve-DnsName $Remote -ErrorAction Stop
        $resolvedIP = ($dnsResult | Where-Object { $_.QueryType -eq 'A' } | Select-Object -First 1).IPAddress
        if ($resolvedIP) { Write-Output "DNS:       OK ($resolvedIP)" } else { Write-Output "DNS:       OK" }
    } catch { Write-Output "DNS:       FAIL - $($_.Exception.Message)"; Write-Output "RESULT:    NOT READY"; return }

    try {
        $tnc = Test-NetConnection $Remote -Port $port -WarningAction SilentlyContinue
        if ($tnc.TcpTestSucceeded) { Write-Output "TCP $port`:   OK" } else { Write-Output "TCP $port`:   FAIL" }
        if ($tnc.PingSucceeded)    { Write-Output "ICMP:      OK ($($tnc.PingReplyDetails.RoundtripTime)ms)" } else { Write-Output "ICMP:      Blocked/No reply" }
    } catch { Write-Output "TCP:       FAIL - $($_.Exception.Message)"; Write-Output "RESULT:    NOT READY"; return }

    $wsmanOk = $false
    if ($tnc.TcpTestSucceeded) {
        try { $ws=Test-WSMan $Remote -ErrorAction Stop; Write-Output "WSMan:     OK ($($ws.ProductVersion))"; $wsmanOk=$true }
        catch { Write-Output "WSMan:     FAIL - $($_.Exception.Message)" }
    } else { Write-Output "WSMan:     Skipped (port closed)" }

    # TrustedHosts status check (v4.2 feature)
    try {
        $th   = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
        $inTH = if ($th -eq '*' -or ($th -and $th -match [regex]::Escape($Remote))) { 'Yes' } else { 'No' }
        Write-Output "Trusted:   $inTH"
    } catch {}

    Write-Output "============================================"
    if ($wsmanOk) { Write-Output "RESULT:    READY" } else { Write-Output "RESULT:    NOT READY" }
}

# ── ENABLE WINRM ──
$EnableWinRMCode = {
    param([double]$u1,[double]$u2,[double]$u3,[string]$Remote,[string]$Transport,
          [pscredential]$Cred,$u4,$u5,$u6,$u7,$u8,[bool]$AutoTrust,$u9,$u10)
    $ErrorActionPreference = 'Stop'
    Write-Output "Enable WinRM: $Remote"
    Write-Output "============================================"

    Write-Output ""; Write-Output "[1/5] Checking current state..."
    try { Test-WSMan $Remote -ErrorAction Stop | Out-Null; Write-Output "  Already enabled!"; Write-Output "RESULT: SUCCESS"; return }
    catch { Write-Output "  Not responding. Proceeding." }

    Write-Output ""; Write-Output "[2/5] Connectivity check..."
    try {
        $tnc = Test-NetConnection $Remote -WarningAction SilentlyContinue
        if (-not $tnc.PingSucceeded) { Write-Output "  FAILED - machine may be offline."; Write-Output "RESULT: FAILED"; return }
        Write-Output "  Ping OK"
    } catch { Write-Output "  Error: $($_.Exception.Message)"; return }

    Write-Output ""; Write-Output "[3/5] Enabling via CIM/WMI..."
    $wmiOk = $false
    # Defined once — used for both CIM and legacy WMI paths
    $winrmCmd = 'cmd.exe /c winrm quickconfig -q & net stop winrm & net start winrm & netsh advfirewall firewall set rule group="Windows Remote Management" new enable=Yes'

    $cimSession = $null
    try {
        $cimP = @{ ComputerName=$Remote; ErrorAction='Stop' }
        if ($Cred) { $cimP.Credential=$Cred }
        $cimSession = New-CimSession @cimP
        $r = Invoke-CimMethod -CimSession $cimSession -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine=$winrmCmd}
        if ($r.ReturnValue -eq 0) { Write-Output "  CIM OK (PID:$($r.ProcessId)). Waiting 10s..."; Start-Sleep 10; $wmiOk=$true }
        else { Write-Output "  CIM code: $($r.ReturnValue)" }
    } catch {
        Write-Output "  CIM failed: $($_.Exception.Message). Trying legacy WMI..."
        try {
            $wmiP = @{ ComputerName=$Remote; ErrorAction='Stop' }
            if ($Cred) { $wmiP.Credential=$Cred }
            $r = Invoke-WmiMethod -Class Win32_Process -Name Create @wmiP -ArgumentList $winrmCmd
            if ($r.ReturnValue -eq 0) { Write-Output "  WMI OK (PID:$($r.ProcessId)). Waiting 10s..."; Start-Sleep 10; $wmiOk=$true }
            else { Write-Output "  WMI code: $($r.ReturnValue)" }
        } catch { Write-Output "  WMI also failed: $($_.Exception.Message)" }
    } finally {
        # Always dispose — prevents handle leaks across multiple Enable attempts
        if ($cimSession) { $cimSession.Dispose(); $cimSession=$null }
    }

    if (-not $wmiOk) {
        Write-Output "  Checking for PsExec..."
        $psexec = $null
        # Only well-known fixed paths — no user-controlled paths
        foreach ($p in @("$env:SystemRoot\System32","$env:SystemRoot",'C:\Tools','C:\SysInternals','C:\PSTools')) {
            $exe = Join-Path $p 'PsExec.exe'
            if (Test-Path $exe -PathType Leaf) { $psexec=$exe; break }
        }
        if (-not $psexec) { try { $psexec=(Get-Command psexec.exe -EA Stop).Source } catch {} }
        if ($psexec) {
            Write-Output "  Found: $psexec"
            try {
                $psi=New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName=$psexec
                $psi.Arguments="\\$Remote -accepteula -s cmd.exe /c `"winrm quickconfig -q`""
                $psi.UseShellExecute=$false; $psi.RedirectStandardOutput=$true; $psi.RedirectStandardError=$true; $psi.CreateNoWindow=$true
                $proc=[System.Diagnostics.Process]::Start($psi); $proc.WaitForExit(30000)
                if ($proc.ExitCode -eq 0) { Write-Output "  PsExec OK. Waiting 5s..."; Start-Sleep 5; $wmiOk=$true }
                else { Write-Output "  PsExec exit: $($proc.ExitCode)" }
            } catch { Write-Output "  PsExec error: $($_.Exception.Message)" }
        } else { Write-Output "  PsExec not found." }
    }

    if ($AutoTrust) {
        Write-Output ""; Write-Output "[4/5] TrustedHosts..."
        try {
            $cur=(Get-Item WSMan:\localhost\Client\TrustedHosts -EA SilentlyContinue).Value
            if ($cur -notmatch [regex]::Escape($Remote)) {
                $new = if ($cur) { "$cur,$Remote" } else { $Remote }
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $new -Force
                Write-Output "  Added: $new"
            } else { Write-Output "  Already trusted." }
        } catch { Write-Output "  Need Admin: $($_.Exception.Message)" }
    } else { Write-Output ""; Write-Output "[4/5] TrustedHosts: skipped (unchecked)" }

    Write-Output ""; Write-Output "[5/5] Verifying..."
    $verified=$false
    for ($i=1; $i -le 3; $i++) {
        try { Test-WSMan $Remote -ErrorAction Stop | Out-Null; $verified=$true; break }
        catch { if ($i -lt 3) { Write-Output "  Retry $i in 5s..."; Start-Sleep 5 } }
    }

    Write-Output ""; Write-Output "============================================"
    if ($verified) { Write-Output "  RESULT: SUCCESS - WinRM enabled on $Remote" }
    elseif ($wmiOk){ Write-Output "  RESULT: PARTIAL - Command sent. Try 'Test WinRM' in 30s." }
    else {
        Write-Output "  RESULT: FAILED"
        Write-Output "  Manual: RDP > Enable-PSRemoting -Force"
        Write-Output "  Or GPO: Computer Config > Windows Components > WinRM"
        Write-Output "  Or:     psexec \\$Remote -s winrm quickconfig -q"
    }
    Write-Output "============================================"
}
#endregion

# ═══════════════════════════════════════════════════════════════
#region === BUTTON HANDLERS ===
function Get-TestParams { @{
    MinDL      = [double]$numMinDL.Value
    MinUL      = [double]$numMinUL.Value
    MaxPL      = [double]$numMaxPL.Value
    Remote     = $txtRemote.Text.Trim()
    Transport  = $cmbTransport.SelectedItem
    Cred       = $null
    ScriptText = ''
    PingCount  = [int]$numPingCount.Value
    DLSize     = [int]$numDLSize.Value
    ULSize     = [int]$numULSize.Value
    DoTrace    = $chkTrace.Checked
    AutoTrust  = $chkTrusted.Checked
    TraceTarget= $txtTraceTarget.Text.Trim()
    MaxParallel= [int]$numParallel.Value
}}

# Resolves credentials for remote operations:
#   - If Alt Creds unchecked: returns null (use current session)
#   - If saved cred exists: reuses it silently
#   - Otherwise: opens GUI dialog (never touches the terminal)
function Get-RemoteCred {
    if (-not $chkCred.Checked) { return $null }
    if ($script:SavedCred)     { return $script:SavedCred }
    $cred = Show-CredentialDialog -Title 'Alt Credentials' -Message 'Enter domain admin credentials for remote machine:'
    if ($cred) {
        $script:SavedCred        = $cred
        $lblCredStatus.Text      = "✔ Set: $($cred.UserName)"
        $lblCredStatus.ForeColor = [Drawing.Color]::FromArgb(27,94,32)
    }
    return $cred
}

$btnCancel.Add_Click({
    if (-not $script:IsBusy) { return }
    if ($script:Timer) { $script:Timer.Stop(); $script:Timer.Dispose(); $script:Timer=$null }
    if ($script:Job)   { Stop-Job $script:Job -EA SilentlyContinue; Remove-Job $script:Job -Force -EA SilentlyContinue; $script:Job=$null }
    Set-UIBusy $false; Set-Done 'Cancelled' 'FAIL'; $txtOut.Text='Cancelled.'
})

$btnClear.Add_Click({
    if (-not $script:IsBusy) {
        $txtOut.Clear(); $lblStatus.Text='  Ready'
        $pnlStatus.BackColor=$C.StatusOK[0]; $lblStatus.ForeColor=$C.StatusOK[1]
    }
})

$btnExport.Add_Click({
    if (-not $txtOut.Text.Trim()) { return }
    $sfd=New-Object Windows.Forms.SaveFileDialog -Property @{ Filter='Text (*.txt)|*.txt'; FileName="SpeedCheck_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt" }
    if ($sfd.ShowDialog() -eq 'OK') {
        # Strip hidden CSV_DATA lines from text export
        ($txtOut.Text -split "`r`n" | Where-Object { $_ -notmatch '^CSV_DATA:' }) -join "`r`n" | Out-File $sfd.FileName -Encoding UTF8
        [Windows.Forms.MessageBox]::Show("Saved: $($sfd.FileName)",'Export','OK','Information') | Out-Null
    }
})

$btnCSV.Add_Click({
    if (-not $txtOut.Text.Trim()) { return }
    $lines = $txtOut.Text -split "`r`n|`n" | Where-Object { $_ -match '^CSV_DATA:' }
    if (-not $lines) { [Windows.Forms.MessageBox]::Show("No CSV data found. Run a test first.",'CSV','OK','Warning') | Out-Null; return }
    $sfd=New-Object Windows.Forms.SaveFileDialog -Property @{ Filter='CSV (*.csv)|*.csv'; FileName="SpeedCheck_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }
    if ($sfd.ShowDialog() -eq 'OK') {
        $csv=@('Computer,Download_Mbps,Upload_Mbps,PacketLoss_Pct,AvgPing_ms,Issues,Timestamp,Notes')
        foreach ($l in $lines) { $csv+=($l -replace '^CSV_DATA:','') }
        $csv -join "`r`n" | Out-File $sfd.FileName -Encoding UTF8
        [Windows.Forms.MessageBox]::Show("CSV saved: $($sfd.FileName)`n$($lines.Count) record(s)",'Export','OK','Information') | Out-Null
    }
})

$btnLocal.Add_Click({
    $p=Get-TestParams; Start-Test -Code $LocalTestCode -Params $p -Label 'Running local test...'
})

$btnRemote.Add_Click({
    if (-not $txtRemote.Text.Trim()) { [Windows.Forms.MessageBox]::Show('Enter a remote computer.','Input') | Out-Null; return }
    if (-not (Test-RemoteHostname $txtRemote.Text.Trim())) {
        [Windows.Forms.MessageBox]::Show("Invalid hostname: '$($txtRemote.Text.Trim())'`nOnly letters, numbers, hyphens, dots and underscores are allowed.",'Validation','OK','Warning') | Out-Null; return
    }
    $p=Get-TestParams
    $p.Cred=$( Get-RemoteCred )
    $p.ScriptText=$LocalTestCode.ToString()
    Start-Test -Code $RemoteTestCode -Params $p -Label "Remote: $($txtRemote.Text.Trim())..."
})

$btnBatch.Add_Click({
    # Multi-line machine entry dialog
    $dlg=New-Object Windows.Forms.Form -Property @{
        Text='Batch Test — Enter Machine Names'; Size=New-Object Drawing.Size(500,420)
        StartPosition='CenterParent'; FormBorderStyle='FixedDialog'; MaximizeBox=$false; MinimizeBox=$false
        BackColor=$C.FormBg; Font=New-Object Drawing.Font('Segoe UI',9.5)
    }
    $dlgLbl=New-Object Windows.Forms.Label -Property @{
        Text='Enter computer names (one per line or comma-separated):'; AutoSize=$true
        Location=New-Object Drawing.Point(16,14); ForeColor=$C.LabelDim
    }
    $dlgTxt=New-Object Windows.Forms.TextBox -Property @{
        Multiline=$true; ScrollBars='Vertical'; WordWrap=$true
        Location=New-Object Drawing.Point(16,40); Size=New-Object Drawing.Size(450,260)
        Font=New-Object Drawing.Font('Consolas',10)
    }
    if ($txtRemote.Text.Trim()) {
        $pre=($txtRemote.Text -split '[,;\s]+' | Where-Object {$_.Trim()} | ForEach-Object {$_.Trim()}) -join "`r`n"
        $dlgTxt.Text=$pre
    }
    $tt.SetToolTip($dlgTxt,"One hostname per line, or comma-separated.`nExample:`n  PC001.domain.com`n  PC002.domain.com")
    $dlgCount=New-Object Windows.Forms.Label -Property @{ Text='0 machines'; AutoSize=$true; Location=New-Object Drawing.Point(16,310); ForeColor=$C.LabelDim }
    $dlgTxt.Add_TextChanged({ $cnt=($dlgTxt.Text -split '[,;\r\n]+' | Where-Object {$_.Trim()}).Count; $dlgCount.Text="$cnt machine(s)" })
    $dlgOK=New-Object Windows.Forms.Button -Property @{
        Text='Run Batch'; Location=New-Object Drawing.Point(270,340); Size=New-Object Drawing.Size(100,36)
        FlatStyle='Flat'; BackColor=$C.BtnBatch; ForeColor='White'; DialogResult='OK'
        Font=New-Object Drawing.Font('Segoe UI Semibold',9.5)
    }; $dlgOK.FlatAppearance.BorderSize=0
    $dlgCancel=New-Object Windows.Forms.Button -Property @{
        Text='Cancel'; Location=New-Object Drawing.Point(380,340); Size=New-Object Drawing.Size(86,36)
        FlatStyle='Flat'; BackColor=$C.BtnClear; ForeColor='White'; DialogResult='Cancel'
        Font=New-Object Drawing.Font('Segoe UI Semibold',9.5)
    }; $dlgCancel.FlatAppearance.BorderSize=0
    $dlg.AcceptButton=$dlgOK; $dlg.CancelButton=$dlgCancel
    $dlg.Controls.AddRange(@($dlgLbl,$dlgTxt,$dlgCount,$dlgOK,$dlgCancel))
    $cnt=($dlgTxt.Text -split '[,;\r\n]+' | Where-Object {$_.Trim()}).Count; $dlgCount.Text="$cnt machine(s)"

    if ($dlg.ShowDialog() -eq 'OK') {
        $machineList=($dlgTxt.Text -split '[,;\r\n]+' | Where-Object {$_.Trim()} | ForEach-Object {$_.Trim()}) -join ','
        if (-not $machineList) { $dlg.Dispose(); return }
        # Validate each hostname before launching
        $invalid = $machineList -split ',' | Where-Object { -not ($_ -match '^[A-Za-z0-9.\-_]+$') }
        if ($invalid) {
            [Windows.Forms.MessageBox]::Show("Invalid hostname(s): $($invalid -join ', ')`nOnly letters, numbers, hyphens, dots and underscores allowed.",'Validation','OK','Warning') | Out-Null
            $dlg.Dispose(); return
        }
        $txtRemote.Text=$machineList
        $p=Get-TestParams; $p.Remote=$machineList
        $p.Cred=$( Get-RemoteCred )
        $p.ScriptText=$LocalTestCode.ToString()
        $cnt=($machineList -split ',' | Where-Object {$_.Trim()}).Count
        Start-Test -Code $BatchTestCode -Params $p -Label "Batch: $cnt machines (parallel)..."
    }
    $dlg.Dispose()
})

$btnTest.Add_Click({
    if (-not $txtRemote.Text.Trim()) { [Windows.Forms.MessageBox]::Show('Enter a remote computer.','Input') | Out-Null; return }
    if (-not (Test-RemoteHostname $txtRemote.Text.Trim())) {
        [Windows.Forms.MessageBox]::Show("Invalid hostname: '$($txtRemote.Text.Trim())'",'Validation','OK','Warning') | Out-Null; return
    }
    $p=Get-TestParams; Start-Test -Code $WinRMTestCode -Params $p -Label "WinRM test: $($p.Remote)..."
})

$btnEnable.Add_Click({
    if (-not $txtRemote.Text.Trim()) { [Windows.Forms.MessageBox]::Show('Enter a remote computer.','Input') | Out-Null; return }
    if (-not (Test-RemoteHostname $txtRemote.Text.Trim())) {
        [Windows.Forms.MessageBox]::Show("Invalid hostname: '$($txtRemote.Text.Trim())'",'Validation','OK','Warning') | Out-Null; return
    }
    $p=Get-TestParams
    $p.Cred=$( Get-RemoteCred )
    Start-Test -Code $EnableWinRMCode -Params $p -Label "Enable WinRM: $($p.Remote)..."
})
#endregion

[void]$form.ShowDialog()
