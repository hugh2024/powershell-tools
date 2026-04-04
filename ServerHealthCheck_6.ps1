#Requires -Version 5.1
<#
.SYNOPSIS  Server Health Diagnostic Tool v6.0
.NOTES     Run: powershell -ExecutionPolicy Bypass -File .\ServerHealthCheck.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#region ── COLORS & FONTS ────────────────────────────────────────────────────
$BG        = [System.Drawing.Color]::FromArgb(10,  13,  18)
$PANEL     = [System.Drawing.Color]::FromArgb(18,  22,  30)
$PANELMD   = [System.Drawing.Color]::FromArgb(22,  27,  36)
$BORDER    = [System.Drawing.Color]::FromArgb(38,  44,  54)
$ACCENT    = [System.Drawing.Color]::FromArgb(31,  111, 235)
$ACCENTHOV = [System.Drawing.Color]::FromArgb(66,  153, 255)
$ACCENTDIM = [System.Drawing.Color]::FromArgb(18,  65,  140)
$CRIT      = [System.Drawing.Color]::FromArgb(235, 68,  57)
$WARN      = [System.Drawing.Color]::FromArgb(208, 158, 28)
$SUCC      = [System.Drawing.Color]::FromArgb(48,  172, 68)
$INFO      = [System.Drawing.Color]::FromArgb(78,  152, 238)
$TXTPRI    = [System.Drawing.Color]::FromArgb(218, 228, 242)
$TXTSEC    = [System.Drawing.Color]::FromArgb(130, 142, 158)
$TXTDIM    = [System.Drawing.Color]::FromArgb(58,  66,  78)
$INPUTBG   = [System.Drawing.Color]::FromArgb(12,  16,  22)
$INPUTDIS  = [System.Drawing.Color]::FromArgb(18,  22,  28)
$WHITE     = [System.Drawing.Color]::White

$fTitle  = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$fH1     = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Bold)
$fBody   = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Regular)
$fSmall  = New-Object System.Drawing.Font("Segoe UI",  8, [System.Drawing.FontStyle]::Regular)
$fSmallB = New-Object System.Drawing.Font("Segoe UI",  8, [System.Drawing.FontStyle]::Bold)
$fTag    = New-Object System.Drawing.Font("Segoe UI",  7, [System.Drawing.FontStyle]::Bold)
$fMono   = New-Object System.Drawing.Font("Consolas",  9, [System.Drawing.FontStyle]::Regular)
$fMonoSm = New-Object System.Drawing.Font("Consolas",  8, [System.Drawing.FontStyle]::Regular)
#endregion

#region ── RICHTEXT HELPERS ──────────────────────────────────────────────────
function Write-RT {
    param($rtb, [string]$text, $color, $font = $null)
    if (-not $font) { $font = $fMono }
    $rtb.SelectionStart  = $rtb.TextLength
    $rtb.SelectionLength = 0
    $rtb.SelectionColor  = $color
    $rtb.SelectionFont   = $font
    $rtb.AppendText("$text`n")
}
function Write-RTSec { param($rtb,$num,$title)
    Write-RT $rtb "" $TXTDIM
    Write-RT $rtb "  ┌──────────────────────────────────────────────────────────────────" $BORDER $fMonoSm
    Write-RT $rtb "  │  $num  ·  $title" $ACCENT $fMono
    Write-RT $rtb "  └──────────────────────────────────────────────────────────────────" $BORDER $fMonoSm
}
function Write-RTItem { param($rtb,$sev,$msg,$detail="")
    $icon  = switch($sev){"CRITICAL"{"  ✖"};"WARNING"{"  ⚠"};"OK"{"  ✔"};"INFO"{"  ●"};default{"  ·"}}
    $color = switch($sev){"CRITICAL"{$CRIT};"WARNING"{$WARN};"OK"{$SUCC};default{$INFO}}
    Write-RT $rtb "$icon  $msg" $color $fMono
    if ($detail) { Write-RT $rtb "       $detail" $TXTSEC $fMonoSm }
}
#endregion

#region ── DIAGNOSTIC ENGINE ─────────────────────────────────────────────────
function Start-Diagnostic {
    param($Server,$Cred,$RTB,$StatusLbl,$ProgBar,$ProgLbl,[int]$Hours)

    $RTB.Clear()
    $issues  = [System.Collections.Generic.List[string]]::new()
    $sn      = 0

    function Upd { param($msg)
        $script:sn++
        $pct = [int](($script:sn/12)*100)
        $ProgBar.Value  = [Math]::Min($pct,100)
        $ProgLbl.Text   = "$pct%"
        $StatusLbl.Text = "  $msg"
        [System.Windows.Forms.Application]::DoEvents()
    }

    $wmi = @{ComputerName=$Server;ErrorAction="Stop"}
    $evt = @{ComputerName=$Server}
    if ($null -ne $Cred) { $wmi["Credential"]=$Cred; $evt["Credential"]=$Cred }

    $since    = (Get-Date).AddHours(-$Hours)
    $lastBoot = (Get-Date).AddHours(-1)
    $wl       = if ($Hours -ge 24 -and $Hours%24 -eq 0){"$($Hours/24) day(s)"}else{"$Hours hour(s)"}

    Write-RT $RTB "" $TXTDIM
    Write-RT $RTB "  ╔════════════════════════════════════════════════════════════════════╗" $BORDER $fMonoSm
    Write-RT $RTB "  ║  SERVER HEALTH DIAGNOSTIC REPORT                                  ║" $ACCENT $fMono
    Write-RT $RTB "  ╠════════════════════════════════════════════════════════════════════╣" $BORDER $fMonoSm
    Write-RT $RTB "  ║  Target  : $($Server.PadRight(56))║" $TXTPRI $fMono
    Write-RT $RTB "  ║  Scanned : $((Get-Date -f 'yyyy-MM-dd HH:mm:ss').PadRight(56))║" $TXTPRI $fMono
    Write-RT $RTB "  ║  Window  : Last $($wl.PadRight(52))║" $TXTPRI $fMono
    Write-RT $RTB "  ╚════════════════════════════════════════════════════════════════════╝" $BORDER $fMonoSm

    # 01
    Upd "01/12 · Checking connectivity..."
    Write-RTSec $RTB "01" "CONNECTIVITY"
    try {
        $ping=$null; $ping=Test-Connection -ComputerName $Server -Count 2 -ErrorAction Stop
        $ms=[math]::Round(($ping|Measure-Object ResponseTime -Average).Average,1)
        $s=if($ms -gt 100){"WARNING"}else{"OK"}
        Write-RTItem $RTB $s "Ping OK — ${ms}ms average response"
        if($ms -gt 100){$issues.Add("High latency: ${ms}ms")}
    } catch { Write-RTItem $RTB "CRITICAL" "Ping FAILED — server unreachable" "$_"; $issues.Add("Server unreachable via ICMP") }
    $w=Test-NetConnection -ComputerName $Server -Port 5985 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    if($w.TcpTestSucceeded){Write-RTItem $RTB "OK" "WinRM port 5985 reachable"}
    else{Write-RTItem $RTB "WARNING" "WinRM port 5985 not reachable";$issues.Add("WinRM 5985 blocked")}

    # 02
    Upd "02/12 · Gathering system information..."
    Write-RTSec $RTB "02" "SYSTEM INFORMATION"
    try {
        $os=Get-WmiObject Win32_OperatingSystem @wmi
        $cs=Get-WmiObject Win32_ComputerSystem  @wmi
        $bi=Get-WmiObject Win32_BIOS            @wmi
        $lastBoot=[Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
        $up=(Get-Date)-$lastBoot; $ups="$($up.Days)d $($up.Hours)h $($up.Minutes)m"
        Write-RTItem $RTB "INFO" "OS        :  $($os.Caption)  (Build $($os.BuildNumber))"
        Write-RTItem $RTB "INFO" "Hostname  :  $($cs.Name)   ·   Domain: $($cs.Domain)"
        Write-RTItem $RTB "INFO" "Last Boot :  $lastBoot"
        Write-RTItem $RTB "INFO" "Uptime    :  $ups"
        Write-RTItem $RTB "INFO" "BIOS      :  $($bi.SMBIOSBIOSVersion) — $($bi.Manufacturer)"
        if($up.TotalHours -lt 2){Write-RTItem $RTB "WARNING" "Server rebooted less than 2 hours ago";$issues.Add("Very recent reboot ($ups ago)")}
    } catch { Write-RTItem $RTB "CRITICAL" "WMI access failed" "$_"; $issues.Add("WMI failed — check creds/firewall") }

    # 03
    Upd "03/12 · Analyzing reboot history..."
    Write-RTSec $RTB "03" "REBOOT ANALYSIS  (last $wl)"
    try {
        $rb=Get-WinEvent @evt -FilterHashtable @{LogName='System';Id=41,1074,1076,6005,6006,6008;StartTime=$since} -EA SilentlyContinue
        if($rb){
            foreach($e in $rb|Sort-Object TimeCreated){
                $s=switch($e.Id){41{"CRITICAL"};6008{"CRITICAL"};1074{"WARNING"};6006{"OK"};6005{"OK"};default{"INFO"}}
                $d=switch($e.Id){41{"UNEXPECTED REBOOT  (Kernel Power / Crash)"};6008{"UNEXPECTED SHUTDOWN"};6006{"Clean shutdown"};6005{"System startup"};1074{"Planned restart/shutdown"};1076{"Admin recorded shutdown reason"};default{"Event $($e.Id)"}}
                Write-RTItem $RTB $s $d "$($e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))  ·  $($e.Message.Split("`n")[0].Trim())"
                if($s -eq "CRITICAL"){$issues.Add("Unexpected reboot/crash at $($e.TimeCreated)")}
            }
        } else { Write-RTItem $RTB "OK" "No reboot events in the last $wl" }
        $pre=Get-WinEvent @evt -FilterHashtable @{LogName='System';Level=1,2;StartTime=$lastBoot.AddMinutes(-30);EndTime=$lastBoot.AddMinutes(2)} -EA SilentlyContinue|Select-Object -First 15
        if($pre){ Write-RT $RTB "  ·  Errors before last reboot:" $TXTSEC $fMonoSm; foreach($e in $pre|Sort-Object TimeCreated){Write-RTItem $RTB "WARNING" "[ID:$($e.Id)]  $($e.ProviderName)" "$($e.TimeCreated.ToString('HH:mm:ss'))  ·  $($e.Message.Split("`n")[0].Trim())"} }
    } catch { Write-RTItem $RTB "WARNING" "Could not query reboot events" "$_" }

    # 04
    Upd "04/12 · Checking CPU and memory..."
    Write-RTSec $RTB "04" "CPU & MEMORY"
    try {
        $cpu=Get-WmiObject Win32_Processor @wmi
        $os2=Get-WmiObject Win32_OperatingSystem @wmi
        $pf=Get-WmiObject Win32_PerfFormattedData_PerfOS_Processor @wmi -Filter "Name='_Total'" -EA SilentlyContinue
        $ld=if($pf){$pf.PercentProcessorTime}else{"N/A"}
        $cs2=if($ld -ne "N/A" -and [int]$ld -gt 90){"CRITICAL"}elseif($ld -ne "N/A" -and [int]$ld -gt 70){"WARNING"}else{"OK"}
        Write-RTItem $RTB "INFO" "Processor :  $($cpu[0].Name.Trim())  [$(@($cpu).Count) socket · $($cpu[0].NumberOfCores) cores]"
        Write-RTItem $RTB $cs2  "CPU Load  :  ${ld}%$(if($ld -ne 'N/A'){if([int]$ld -gt 90){'  ◀ CRITICAL'}elseif([int]$ld -gt 70){'  ◀ High'}else{'  — Normal'}})"
        if($ld -ne "N/A" -and [int]$ld -gt 90){$issues.Add("CPU critically high: ${ld}%")}
        elseif($ld -ne "N/A" -and [int]$ld -gt 70){$issues.Add("CPU load elevated: ${ld}%")}
        $tot=[math]::Round($os2.TotalVisibleMemorySize/1MB,1); $free=[math]::Round($os2.FreePhysicalMemory/1MB,1)
        $used=[math]::Round($tot-$free,1); $pct=[math]::Round(($used/$tot)*100,1)
        $ms2=if($pct -gt 95){"CRITICAL"}elseif($pct -gt 85){"WARNING"}else{"OK"}
        Write-RTItem $RTB $ms2 "Memory    :  ${used} GB / ${tot} GB  (${pct}%)$(if($pct -gt 95){'  ◀ CRITICAL'}elseif($pct -gt 85){'  ◀ High'}else{'  — Normal'})"
        if($pct -gt 95){$issues.Add("Memory critically high: ${pct}%")}elseif($pct -gt 85){$issues.Add("Memory elevated: ${pct}%")}
    } catch { Write-RTItem $RTB "WARNING" "Could not retrieve CPU/memory" "$_" }

    # 05
    Upd "05/12 · Checking disk health..."
    Write-RTSec $RTB "05" "DISK HEALTH"
    try {
        foreach($d in (Get-WmiObject Win32_LogicalDisk @wmi -Filter "DriveType=3")){
            if($d.Size -eq 0){continue}
            $tot=[math]::Round($d.Size/1GB,1); $free=[math]::Round($d.FreeSpace/1GB,1); $used=[math]::Round($tot-$free,1); $pct=[math]::Round(($used/$tot)*100,1)
            $fill=[int]($pct/5); $bar="["+("█"*$fill)+("░"*(20-$fill))+"]"
            $s=if($pct -gt 95){"CRITICAL"}elseif($pct -gt 85){"WARNING"}else{"OK"}
            Write-RTItem $RTB $s "Drive $($d.DeviceID)  $bar  ${pct}%  ($used / $tot GB)" "Free: ${free} GB"
            if($pct -gt 95){$issues.Add("Disk $($d.DeviceID) critically full: ${pct}%")}elseif($pct -gt 85){$issues.Add("Disk $($d.DeviceID) low space: ${pct}%")}
        }
        $de=Get-WinEvent @evt -FilterHashtable @{LogName='System';Id=7,11,15,51,52,153;StartTime=$since} -EA SilentlyContinue|Select-Object -First 8
        if($de){foreach($e in $de|Sort-Object TimeCreated -Descending){Write-RTItem $RTB "CRITICAL" "Disk I/O Error  [ID:$($e.Id)]" "$($e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))  ·  $($e.Message.Split("`n")[0].Trim())";$issues.Add("Disk I/O error ID $($e.Id)")}}
        else{Write-RTItem $RTB "OK" "No disk I/O errors in event log"}
    } catch { Write-RTItem $RTB "WARNING" "Could not retrieve disk data" "$_" }

    # 06
    Upd "06/12 · Scanning system errors..."
    Write-RTSec $RTB "06" "SYSTEM EVENT ERRORS & WARNINGS  (last $wl)"
    try {
        $se=Get-WinEvent @evt -FilterHashtable @{LogName='System';Level=1,2;StartTime=$since} -EA SilentlyContinue|Select-Object -First 30
        if($se){foreach($g in($se|Group-Object Id|Sort-Object Count -Descending)){$e=$g.Group[0];$cnt=if($g.Count-gt 1){"  (×$($g.Count))"}else{""};Write-RTItem $RTB "CRITICAL" "[ID:$($e.Id)]  $($e.ProviderName)$cnt" "$($e.Message.Split("`n")[0].Trim())";$issues.Add("System Error $($e.Id) · $($e.ProviderName)$cnt")}}
        else{Write-RTItem $RTB "OK" "No critical system errors in the last $wl"}
        $sw=Get-WinEvent @evt -FilterHashtable @{LogName='System';Level=3;StartTime=$since} -EA SilentlyContinue|Select-Object -First 20
        if($sw){foreach($g in($sw|Group-Object Id|Sort-Object Count -Descending|Select-Object -First 10)){$e=$g.Group[0];$cnt=if($g.Count-gt 1){"  (×$($g.Count))"}else{""}; Write-RTItem $RTB "WARNING" "[ID:$($e.Id)]  $($e.ProviderName)$cnt" "$($e.Message.Split("`n")[0].Trim())"}}
    } catch { Write-RTItem $RTB "WARNING" "Could not query system events" "$_" }

    # 07
    Upd "07/12 · Scanning application errors..."
    Write-RTSec $RTB "07" "APPLICATION EVENT ERRORS  (last $wl)"
    try {
        $ae=Get-WinEvent @evt -FilterHashtable @{LogName='Application';Level=1,2;StartTime=$since} -EA SilentlyContinue|Select-Object -First 20
        if($ae){foreach($g in($ae|Group-Object Id|Sort-Object Count -Descending|Select-Object -First 10)){$e=$g.Group[0];$cnt=if($g.Count-gt 1){"  (×$($g.Count))"}else{""}; Write-RTItem $RTB "CRITICAL" "[ID:$($e.Id)]  $($e.ProviderName)$cnt" "$($e.Message.Split("`n")[0].Trim())"; $issues.Add("App Error $($e.Id) · $($e.ProviderName)$cnt")}}
        else{Write-RTItem $RTB "OK" "No critical application errors"}
    } catch { Write-RTItem $RTB "WARNING" "Could not query application events" "$_" }

    # 08
    Upd "08/12 · Checking critical services..."
    Write-RTSec $RTB "08" "CRITICAL SERVICES"
    $ws=@("wuauserv","WinDefend","MpsSvc","EventLog","Dnscache","LanmanServer","LanmanWorkstation","RpcSs","W32Time","BITS","CryptSvc","Spooler","winrm")
    try {
        $svcs=Get-WmiObject Win32_Service @wmi|Where-Object{$_.Name -in $ws}; $fail=0
        foreach($s in $svcs|Sort-Object Name){
            if($s.StartMode -eq "Auto" -and $s.State -ne "Running"){Write-RTItem $RTB "CRITICAL" "$($s.Name.PadRight(26)) STOPPED  (Auto-start)" "Expected: Running";$issues.Add("Stopped service: $($s.Name)");$fail++}
            else{Write-RTItem $RTB "OK" "$($s.Name.PadRight(26)) $($s.State)"}
        }
        if($fail -eq 0){Write-RT $RTB "  ·  All $(@($svcs).Count) monitored services running normally" $SUCC $fMonoSm}
        $sc=Get-WinEvent @evt -FilterHashtable @{LogName='System';Id=7031,7032,7034;StartTime=$since} -EA SilentlyContinue|Select-Object -First 10
        if($sc){Write-RT $RTB "  ·  Service crashes:" $WARN $fMonoSm; foreach($e in $sc|Sort-Object TimeCreated -Descending){Write-RTItem $RTB "WARNING" "[ID:$($e.Id)]  Service crash/restart" "$($e.TimeCreated.ToString('HH:mm:ss'))  ·  $($e.Message.Split("`n")[0].Trim())"; $issues.Add("Service crash: $($e.Message.Split("`n")[0].Trim())")}}
    } catch { Write-RTItem $RTB "WARNING" "Could not retrieve service data" "$_" }

    # 09
    Upd "09/12 · Checking Windows Update..."
    Write-RTSec $RTB "09" "WINDOWS UPDATE"
    try {
        $wu=Get-WinEvent @evt -FilterHashtable @{LogName='System';Id=19,20,21,22,43;StartTime=$since} -EA SilentlyContinue
        if($wu){foreach($e in $wu|Sort-Object TimeCreated -Descending|Select-Object -First 8){$s=if($e.Id -in 20,21){"WARNING"}else{"INFO"};Write-RTItem $RTB $s "[WU:$($e.Id)]  $($e.Message.Split("`n")[0].Trim())" "$($e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))";if($e.Id -in 20,21){$issues.Add("WU error ID $($e.Id)")}}}
        else{Write-RTItem $RTB "OK" "No Windows Update activity in this window"}
        try {
            $ia=@{ComputerName=$Server;ErrorAction="Stop";ScriptBlock={
                $p=$false
                if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){$p=$true}
                if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"){$p=$true}
                if(Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -EA SilentlyContinue){$p=$true}
                $p
            }}
            if($null -ne $Cred){$ia["Credential"]=$Cred}
            if((Invoke-Command @ia)){Write-RTItem $RTB "WARNING" "PENDING REBOOT — restart required";$issues.Add("Pending reboot required")}
            else{Write-RTItem $RTB "OK" "No pending reboot detected"}
        } catch { Write-RTItem $RTB "INFO" "Pending reboot check skipped (WinRM unavailable)" }
    } catch { Write-RTItem $RTB "WARNING" "Could not retrieve Windows Update data" "$_" }

    # 10
    Upd "10/12 · Scanning security events..."
    Write-RTSec $RTB "10" "SECURITY EVENTS  (last $wl)"
    try {
        $fl=Get-WinEvent @evt -FilterHashtable @{LogName='Security';Id=4625;StartTime=$since} -EA SilentlyContinue
        if($fl){$cnt=@($fl).Count;$s=if($cnt -gt 50){"CRITICAL"}elseif($cnt -gt 10){"WARNING"}else{"INFO"};Write-RTItem $RTB $s "Failed logins: $cnt$(if($cnt -gt 50){'  ◀ POSSIBLE BRUTE FORCE'}elseif($cnt -gt 10){'  ◀ Elevated'}else{'  — Normal'})" "Event 4625";if($cnt -gt 50){$issues.Add("HIGH failed logins: $cnt")}elseif($cnt -gt 10){$issues.Add("Elevated failed logins: $cnt")}}
        else{Write-RTItem $RTB "OK" "No failed login attempts"}
        $lo=Get-WinEvent @evt -FilterHashtable @{LogName='Security';Id=4740;StartTime=$since} -EA SilentlyContinue
        if($lo){Write-RTItem $RTB "WARNING" "Account lockouts: $(@($lo).Count)" "Event 4740";$issues.Add("$(@($lo).Count) lockout(s)")}else{Write-RTItem $RTB "OK" "No account lockouts"}
        $ap=Get-WinEvent @evt -FilterHashtable @{LogName='Security';Id=4719;StartTime=$since} -EA SilentlyContinue
        if($ap){Write-RTItem $RTB "WARNING" "Audit policy changed ($(@($ap).Count)×)" "Event 4719";$issues.Add("Audit policy changed")}else{Write-RTItem $RTB "OK" "No audit policy changes"}
        $cl=Get-WinEvent @evt -FilterHashtable @{LogName='Security';Id=1102;StartTime=$since} -EA SilentlyContinue
        if($cl){Write-RTItem $RTB "CRITICAL" "EVENT LOG CLEARED (ID 1102) — investigate immediately" "Possible tampering";$issues.Add("SECURITY: Event log cleared")}else{Write-RTItem $RTB "OK" "Event log intact"}
    } catch { Write-RTItem $RTB "INFO" "Security log limited — run as domain admin for full audit" }

    # 11
    Upd "11/12 · Checking network..."
    Write-RTSec $RTB "11" "NETWORK"
    try {
        $ne=Get-WinEvent @evt -FilterHashtable @{LogName='System';Id=4199,4201,4202,4227,4231;StartTime=$since} -EA SilentlyContinue
        if($ne){foreach($e in $ne|Sort-Object TimeCreated -Descending|Select-Object -First 8){Write-RTItem $RTB "WARNING" "[ID:$($e.Id)]  Network event" "$($e.TimeCreated.ToString('HH:mm:ss'))  ·  $($e.Message.Split("`n")[0].Trim())";$issues.Add("Network event $($e.Id)")}}
        else{Write-RTItem $RTB "OK" "No network errors in event log"}
        foreach($n in (Get-WmiObject Win32_NetworkAdapterConfiguration @wmi|Where-Object{$_.IPEnabled})){Write-RTItem $RTB "INFO" "NIC: $($n.Description)" "IP: $($n.IPAddress -join ', ')   GW: $($n.DefaultIPGateway -join ', ')   DNS: $($n.DNSServerSearchOrder -join ', ')"}
    } catch { Write-RTItem $RTB "WARNING" "Could not retrieve network data" "$_" }

    # 12
    Upd "12/12 · Checking hardware and drivers..."
    Write-RTSec $RTB "12" "HARDWARE & DRIVER ERRORS  (last $wl)"
    try {
        $hw=Get-WinEvent @evt -FilterHashtable @{LogName='System';Id=41,1001,6008,10,219,411,10016,9,11,15,117;StartTime=$since} -EA SilentlyContinue|Select-Object -First 15
        if($hw){foreach($e in $hw|Sort-Object TimeCreated -Descending){$s=if($e.Id -in 41,6008,9,11){"CRITICAL"}else{"WARNING"};Write-RTItem $RTB $s "[ID:$($e.Id)]  $($e.ProviderName)" "$($e.TimeCreated.ToString('HH:mm:ss'))  ·  $($e.Message.Split("`n")[0].Trim())";if($s -eq "CRITICAL"){$issues.Add("Hardware error ID $($e.Id)")}}}
        else{Write-RTItem $RTB "OK" "No hardware or driver errors detected"}
    } catch { Write-RTItem $RTB "WARNING" "Could not retrieve hardware events" "$_" }

    # Summary
    $ProgBar.Value=100; $ProgLbl.Text="Done"
    [System.Windows.Forms.Application]::DoEvents()
    Write-RT $RTB "" $TXTDIM
    Write-RT $RTB "  ╔════════════════════════════════════════════════════════════════════╗" $BORDER $fMonoSm
    Write-RT $RTB "  ║  DIAGNOSTIC SUMMARY                                               ║" $ACCENT $fMono
    Write-RT $RTB "  ╠════════════════════════════════════════════════════════════════════╣" $BORDER $fMonoSm
    if($issues.Count -eq 0){ Write-RT $RTB "  ║  ✔  No issues found — server appears healthy" $SUCC $fMono }
    else {
        $crit=@($issues|Where-Object{$_ -match "CRITICAL|critical|crash|Cleared|brute|SECURITY"}).Count
        Write-RT $RTB "  ║  Found $($issues.Count) issue(s):  $crit critical  ·  $($issues.Count-$crit) warning(s)" $WARN $fMono
        Write-RT $RTB "  ╠════════════════════════════════════════════════════════════════════╣" $BORDER $fMonoSm
        $i=1; foreach($iss in $issues){
            $col=if($iss -match "CRITICAL|critical|crash|Cleared|brute|SECURITY"){$CRIT}else{$WARN}
            $line="  $i.  $iss"; if($line.Length -gt 70){$line=$line.Substring(0,67)+"..."}
            Write-RT $RTB $line $col $fMono; $i++
        }
    }
    Write-RT $RTB "  ╠════════════════════════════════════════════════════════════════════╣" $BORDER $fMonoSm
    Write-RT $RTB "  ║  Completed: $(Get-Date -f 'yyyy-MM-dd HH:mm:ss')" $TXTSEC $fMonoSm
    Write-RT $RTB "  ╚════════════════════════════════════════════════════════════════════╝" $BORDER $fMonoSm
    $StatusLbl.Text="  ✔  Complete — $($issues.Count) issue(s) found on $Server  ·  $(Get-Date -f 'HH:mm:ss')"
    $StatusLbl.ForeColor=if($issues.Count -gt 0){$WARN}else{$SUCC}
}
#endregion

#region ── FORM ───────────────────────────────────────────────────────────────
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Server Health Diagnostic  v6.0"
$form.Size          = New-Object System.Drawing.Size(1160, 840)
$form.MinimumSize   = New-Object System.Drawing.Size(900,  640)
$form.BackColor     = $BG
$form.ForeColor     = $TXTPRI
$form.Font          = $fBody
$form.StartPosition = "CenterScreen"

# Use a TableLayoutPanel as the root container — 4 rows, fills form
# Row 0: title      58px fixed
# Row 1: toolbar    72px fixed
# Row 2: progress   20px fixed
# Row 3: output     fills remaining
# Row 4: statusbar  26px fixed
$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock        = "Fill"
$root.ColumnCount = 1
$root.RowCount    = 5
$root.BackColor   = $BG
$root.Padding     = New-Object System.Windows.Forms.Padding(0)
$root.Margin      = New-Object System.Windows.Forms.Padding(0)
[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))   # title
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 72)))   # toolbar
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))   # progress
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,  100)))  # output
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))   # statusbar
$form.Controls.Add($root)

# ── ROW 0: Title ──────────────────────────────────────────────────────────────
$titlePnl           = New-Object System.Windows.Forms.Panel
$titlePnl.Dock      = "Fill"
$titlePnl.BackColor = $PANEL
$titlePnl.Margin    = New-Object System.Windows.Forms.Padding(0)
$root.Controls.Add($titlePnl, 0, 0)

$lTitle             = New-Object System.Windows.Forms.Label
$lTitle.Text        = "SERVER HEALTH DIAGNOSTIC"
$lTitle.Font        = $fTitle
$lTitle.ForeColor   = $ACCENTHOV
$lTitle.AutoSize    = $true
$lTitle.Location    = New-Object System.Drawing.Point(16, 8)
$titlePnl.Controls.Add($lTitle)

$lSub               = New-Object System.Windows.Forms.Label
$lSub.Text          = "v6.0  ·  12-point check  ·  Connectivity  ·  Reboots  ·  CPU/Memory  ·  Disk  ·  Services  ·  Security  ·  Hardware"
$lSub.Font          = $fSmall
$lSub.ForeColor     = $TXTSEC
$lSub.AutoSize      = $true
$lSub.Location      = New-Object System.Drawing.Point(18, 36)
$titlePnl.Controls.Add($lSub)

# ── ROW 1: Toolbar ────────────────────────────────────────────────────────────
$toolPnl            = New-Object System.Windows.Forms.Panel
$toolPnl.Dock       = "Fill"
$toolPnl.BackColor  = $PANELMD
$toolPnl.Margin     = New-Object System.Windows.Forms.Padding(0)
$root.Controls.Add($toolPnl, 0, 1)

# Helper functions scoped to toolPnl
function TBLbl { param($t,$x)
    $l=New-Object System.Windows.Forms.Label; $l.Text=$t; $l.Font=$fTag; $l.ForeColor=$TXTDIM
    $l.Location=New-Object System.Drawing.Point($x,10); $l.AutoSize=$true; $toolPnl.Controls.Add($l) }
function TBTxt { param($x,$w,$ph="")
    $t=New-Object System.Windows.Forms.TextBox; $t.Location=New-Object System.Drawing.Point($x,28)
    $t.Size=New-Object System.Drawing.Size($w,26); $t.BackColor=$INPUTBG; $t.ForeColor=$TXTPRI
    $t.BorderStyle="FixedSingle"; $t.Font=$fBody
    if($ph){try{$t.PlaceholderText=$ph}catch{}}
    $toolPnl.Controls.Add($t); $t }

# Columns: SERVER[14] AUTH[222] USER[496] PASS[696] TIME[872] RUN[958] SAVE[1138]
TBLbl "SERVER / IP" 14
$script:txSrv = TBTxt 14 200

TBLbl "AUTHENTICATION" 222

$script:rbCur           = New-Object System.Windows.Forms.RadioButton
$script:rbCur.Text      = "Current User  ($env:USERDOMAIN\$env:USERNAME)"
$script:rbCur.Font      = $fSmall; $script:rbCur.ForeColor=$SUCC
$script:rbCur.Location  = New-Object System.Drawing.Point(222, 24)
$script:rbCur.AutoSize  = $true; $script:rbCur.Checked=$true
$toolPnl.Controls.Add($script:rbCur)

$script:rbCust          = New-Object System.Windows.Forms.RadioButton
$script:rbCust.Text     = "Custom Credentials"
$script:rbCust.Font     = $fSmall; $script:rbCust.ForeColor=$TXTSEC
$script:rbCust.Location = New-Object System.Drawing.Point(222, 46)
$script:rbCust.AutoSize = $true
$toolPnl.Controls.Add($script:rbCust)

TBLbl "USERNAME" 496
$script:txUsr           = TBTxt 496 188 "DOMAIN\username"
$script:txUsr.Enabled   = $false; $script:txUsr.BackColor=$INPUTDIS; $script:txUsr.ForeColor=$TXTDIM

TBLbl "PASSWORD" 694
$script:txPwd           = TBTxt 694 166
$script:txPwd.PasswordChar=[char]0x25CF; $script:txPwd.Enabled=$false
$script:txPwd.BackColor =$INPUTDIS; $script:txPwd.ForeColor=$TXTDIM

TBLbl "TIME WINDOW" 872

$script:numT            = New-Object System.Windows.Forms.NumericUpDown
$script:numT.Location   = New-Object System.Drawing.Point(872, 28)
$script:numT.Size       = New-Object System.Drawing.Size(58, 26)
$script:numT.BackColor  = $INPUTBG; $script:numT.ForeColor=$TXTPRI; $script:numT.Font=$fBody
$script:numT.Minimum=1; $script:numT.Maximum=999; $script:numT.Value=24
$toolPnl.Controls.Add($script:numT)

$script:cmbU            = New-Object System.Windows.Forms.ComboBox
$script:cmbU.Location   = New-Object System.Drawing.Point(936, 28)
$script:cmbU.Size       = New-Object System.Drawing.Size(74, 26)
$script:cmbU.BackColor  = $INPUTBG; $script:cmbU.ForeColor=$TXTPRI; $script:cmbU.Font=$fSmall
$script:cmbU.DropDownStyle="DropDownList"; $script:cmbU.FlatStyle="Flat"
[void]$script:cmbU.Items.Add("Hours"); [void]$script:cmbU.Items.Add("Days")
$script:cmbU.SelectedIndex=0
$toolPnl.Controls.Add($script:cmbU)
$script:cmbU.Add_SelectedIndexChanged({
    if($script:cmbU.SelectedItem -eq "Days"){$script:numT.Maximum=30;if($script:numT.Value -gt 30){$script:numT.Value=7}}
    else{$script:numT.Maximum=999}
})

$script:btnRun          = New-Object System.Windows.Forms.Button
$script:btnRun.Text     = "▶   RUN DIAGNOSTIC"
$script:btnRun.Location = New-Object System.Drawing.Point(1020, 14)
$script:btnRun.Size     = New-Object System.Drawing.Size(168, 44)
$script:btnRun.BackColor=$ACCENT; $script:btnRun.ForeColor=$WHITE
$script:btnRun.FlatStyle="Flat"; $script:btnRun.FlatAppearance.BorderSize=0
$script:btnRun.Font=$fH1; $script:btnRun.Cursor=[System.Windows.Forms.Cursors]::Hand
$toolPnl.Controls.Add($script:btnRun)

# Radio toggle
$script:rbCur.Add_CheckedChanged({
    $en=$script:rbCust.Checked
    $script:txUsr.Enabled=$en; $script:txPwd.Enabled=$en
    $script:txUsr.BackColor=if($en){$INPUTBG}else{$INPUTDIS}; $script:txPwd.BackColor=if($en){$INPUTBG}else{$INPUTDIS}
    $script:txUsr.ForeColor=if($en){$TXTPRI}else{$TXTDIM};   $script:txPwd.ForeColor=if($en){$TXTPRI}else{$TXTDIM}
    $script:rbCust.ForeColor=if($en){$SUCC}else{$TXTSEC};    $script:rbCur.ForeColor=if($en){$TXTSEC}else{$SUCC}
})

# ── ROW 2: Progress bar ───────────────────────────────────────────────────────
$progPnl            = New-Object System.Windows.Forms.Panel
$progPnl.Dock       = "Fill"
$progPnl.BackColor  = $BG
$progPnl.Margin     = New-Object System.Windows.Forms.Padding(0)
$root.Controls.Add($progPnl, 0, 2)

$script:progBar     = New-Object System.Windows.Forms.ProgressBar
$script:progBar.Location=New-Object System.Drawing.Point(0,0)
$script:progBar.Size=New-Object System.Drawing.Size(1060,20)
$script:progBar.Style="Continuous"; $script:progBar.ForeColor=$ACCENT; $script:progBar.BackColor=$BORDER
$script:progBar.Minimum=0; $script:progBar.Maximum=100; $script:progBar.Value=0
$progPnl.Controls.Add($script:progBar)

$script:progLbl     = New-Object System.Windows.Forms.Label
$script:progLbl.Text="Ready"; $script:progLbl.Font=$fSmallB; $script:progLbl.ForeColor=$TXTSEC
$script:progLbl.Location=New-Object System.Drawing.Point(1068,2); $script:progLbl.AutoSize=$true
$progPnl.Controls.Add($script:progLbl)

# ── ROW 3: Output RichTextBox ─────────────────────────────────────────────────
$script:rtb         = New-Object System.Windows.Forms.RichTextBox
$script:rtb.Dock    = "Fill"
$script:rtb.BackColor=$BG; $script:rtb.ForeColor=$TXTPRI; $script:rtb.Font=$fMono
$script:rtb.ReadOnly=$true; $script:rtb.BorderStyle="None"
$script:rtb.ScrollBars="Vertical"; $script:rtb.WordWrap=$false
$script:rtb.Margin  = New-Object System.Windows.Forms.Padding(0)
$root.Controls.Add($script:rtb, 0, 3)

# ── ROW 4: Status bar ─────────────────────────────────────────────────────────
$statPnl            = New-Object System.Windows.Forms.Panel
$statPnl.Dock       = "Fill"
$statPnl.BackColor  = $PANEL
$statPnl.Margin     = New-Object System.Windows.Forms.Padding(0)
$root.Controls.Add($statPnl, 0, 4)

$script:lblStat     = New-Object System.Windows.Forms.Label
$script:lblStat.Text="  Ready — enter server name or IP above and click  ▶ Run Diagnostic"
$script:lblStat.Font=$fSmall; $script:lblStat.ForeColor=$TXTSEC
$script:lblStat.Location=New-Object System.Drawing.Point(6,5); $script:lblStat.AutoSize=$true
$statPnl.Controls.Add($script:lblStat)

$lVer               = New-Object System.Windows.Forms.Label
$lVer.Text          = "IT Security · v6.0  "
$lVer.Font          = $fSmall; $lVer.ForeColor=$TXTDIM; $lVer.AutoSize=$true
$lVer.Location      = New-Object System.Drawing.Point(1060, 5)
$statPnl.Controls.Add($lVer)

# Welcome content
Write-RT $script:rtb "" $TXTDIM
Write-RT $script:rtb "  ╔════════════════════════════════════════════════════════════════════╗" $BORDER $fMonoSm
Write-RT $script:rtb "  ║  SERVER HEALTH DIAGNOSTIC TOOL  v6.0                              ║" $ACCENT $fMono
Write-RT $script:rtb "  ║  Comprehensive 12-point Windows Server health check                ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ╠════════════════════════════════════════════════════════════════════╣" $BORDER $fMonoSm
Write-RT $script:rtb "  ║  01 · Connectivity          ping latency + WinRM port              ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  02 · System Information    OS, hostname, uptime, BIOS             ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  03 · Reboot Analysis       planned vs unexpected + pre-boot err   ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  04 · CPU & Memory          utilization with thresholds            ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  05 · Disk Health           usage bar + I/O error events           ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  06 · System Event Errors   grouped by ID with count               ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  07 · Application Errors    critical + error events                ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  08 · Critical Services     auto-start + crash detection           ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  09 · Windows Update        status + pending reboot check          ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  10 · Security Events       failed logins, lockouts, log clear     ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  11 · Network               errors + NIC / IP / DNS info           ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ║  12 · Hardware & Drivers    kernel, storage, driver errors         ║" $TXTSEC $fMonoSm
Write-RT $script:rtb "  ╠════════════════════════════════════════════════════════════════════╣" $BORDER $fMonoSm
Write-RT $script:rtb "  ║  ▶  Enter a server name or IP above and click RUN DIAGNOSTIC       ║" $ACCENT $fMono
Write-RT $script:rtb "  ╚════════════════════════════════════════════════════════════════════╝" $BORDER $fMonoSm
#endregion

#region ── EVENTS ─────────────────────────────────────────────────────────────
$script:btnRun.Add_Click({
    $server=$script:txSrv.Text.Trim()
    if([string]::IsNullOrEmpty($server)){
        [System.Windows.Forms.MessageBox]::Show("Please enter a server name or IP.","Input Required",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)|Out-Null; return }
    $cred=$null
    if($script:rbCust.Checked){
        $user=$script:txUsr.Text.Trim()
        if([string]::IsNullOrEmpty($user)){
            [System.Windows.Forms.MessageBox]::Show("Please enter a username.","Input Required",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)|Out-Null; return }
        try{$sec=ConvertTo-SecureString $script:txPwd.Text -AsPlainText -Force; $cred=New-Object System.Management.Automation.PSCredential($user,$sec)}
        catch{[System.Windows.Forms.MessageBox]::Show("Credential error: $_","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)|Out-Null; return}
    }
    $hours=if($script:cmbU.SelectedItem -eq "Days"){[int]$script:numT.Value*24}else{[int]$script:numT.Value}
    $script:btnRun.Enabled=$false; $script:btnRun.Text="  Scanning..."
    $script:btnRun.BackColor=$ACCENTDIM
    $script:progBar.Value=0; $script:progLbl.Text="0%"
    $script:lblStat.Text="  Connecting to $server..."; $script:lblStat.ForeColor=$TXTSEC
    $form.Refresh()
    try{
        Start-Diagnostic -Server $server -Cred $cred -RTB $script:rtb `
            -StatusLbl $script:lblStat -ProgBar $script:progBar -ProgLbl $script:progLbl -Hours $hours
    } catch {
        Write-RT $script:rtb "  FATAL ERROR: $_" $CRIT $fMono
        $script:lblStat.Text="  Scan failed"; $script:lblStat.ForeColor=$CRIT
    } finally {
        $script:btnRun.Enabled=$true; $script:btnRun.Text="▶   RUN DIAGNOSTIC"; $script:btnRun.BackColor=$ACCENT
    }
})

# Save button — added to toolbar after Run (outside TableLayout for simplicity)
$btnSave            = New-Object System.Windows.Forms.Button
$btnSave.Text       = "💾  Save"
$btnSave.Location   = New-Object System.Drawing.Point(($script:btnRun.Right + 8), 14)
$btnSave.Size       = New-Object System.Drawing.Size(80, 44)
$btnSave.BackColor  = $PANELMD; $btnSave.ForeColor=$TXTSEC
$btnSave.FlatStyle  = "Flat"; $btnSave.FlatAppearance.BorderSize=1; $btnSave.FlatAppearance.BorderColor=$BORDER
$btnSave.Font=$fSmall; $btnSave.Cursor=[System.Windows.Forms.Cursors]::Hand
$toolPnl.Controls.Add($btnSave)

$btnSave.Add_Click({
    if($script:rtb.TextLength -eq 0){return}
    $dlg=New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter="Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $dlg.FileName="ServerHealth_$($script:txSrv.Text.Trim())_$(Get-Date -f 'yyyyMMdd_HHmmss').txt"
    if($dlg.ShowDialog() -eq "OK"){$script:rtb.Text|Out-File $dlg.FileName -Encoding UTF8;$script:lblStat.Text="  Saved: $($dlg.FileName)";$script:lblStat.ForeColor=$SUCC}
})

$form.KeyPreview=$true
$form.Add_KeyDown({if($_.KeyCode -eq "Return" -and $script:btnRun.Enabled){$script:btnRun.PerformClick()}})
$script:btnRun.Add_MouseEnter({if($script:btnRun.Enabled){$script:btnRun.BackColor=$ACCENTHOV}})
$script:btnRun.Add_MouseLeave({if($script:btnRun.Enabled){$script:btnRun.BackColor=$ACCENT}})
$btnSave.Add_MouseEnter({$btnSave.ForeColor=$TXTPRI}); $btnSave.Add_MouseLeave({$btnSave.ForeColor=$TXTSEC})

$form.Add_Resize({
    $script:progBar.Width=[Math]::Max(200,$form.ClientSize.Width-100)
    $script:progLbl.Left=$script:progBar.Right+8
    $lVer.Left=[Math]::Max(600,$form.ClientSize.Width-$lVer.Width-10)
})
#endregion

[System.Windows.Forms.Application]::Run($form)
