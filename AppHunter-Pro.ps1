#Requires -Version 5.1
# AppHunter Pro v5.0 - Enterprise Software & Service Discovery
# https://github.com/hugh2024

Set-StrictMode -Off
$ErrorActionPreference = 'SilentlyContinue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#region COLORS
$CB0  = [Drawing.Color]::FromArgb(10,12,18)
$CB1  = [Drawing.Color]::FromArgb(18,22,32)
$CB2  = [Drawing.Color]::FromArgb(28,34,50)
$CB3  = [Drawing.Color]::FromArgb(38,46,66)
$CAcc = [Drawing.Color]::FromArgb(0,168,255)
$CAccH= [Drawing.Color]::FromArgb(40,195,255)
$CAccD= [Drawing.Color]::FromArgb(0,100,160)
$CGrn = [Drawing.Color]::FromArgb(0,210,110)
$CGrnH= [Drawing.Color]::FromArgb(30,235,130)
$CYel = [Drawing.Color]::FromArgb(255,185,0)
$CYelH= [Drawing.Color]::FromArgb(255,210,50)
$CRed = [Drawing.Color]::FromArgb(215,50,60)
$CRedH= [Drawing.Color]::FromArgb(250,80,90)
$CTxt = [Drawing.Color]::FromArgb(215,225,242)
$CDim = [Drawing.Color]::FromArgb(95,112,145)
$CWht = [Drawing.Color]::FromArgb(255,255,255)
$CBdr = [Drawing.Color]::FromArgb(42,52,76)
$CAlt = [Drawing.Color]::FromArgb(22,27,42)
$CSel = [Drawing.Color]::FromArgb(0,75,125)
$COkBG= [Drawing.Color]::FromArgb(4,50,20)
$CWnBG= [Drawing.Color]::FromArgb(50,40,0)
$CErBG= [Drawing.Color]::FromArgb(60,10,10)
#endregion

#region FONTS
$FMono = New-Object Drawing.Font('Consolas',9,[Drawing.FontStyle]::Regular)
$FUI   = New-Object Drawing.Font('Segoe UI',9,[Drawing.FontStyle]::Regular)
$FUIB  = New-Object Drawing.Font('Segoe UI',9,[Drawing.FontStyle]::Bold)
$FSm   = New-Object Drawing.Font('Segoe UI',8,[Drawing.FontStyle]::Regular)
$FSmB  = New-Object Drawing.Font('Segoe UI',8,[Drawing.FontStyle]::Bold)
$FTtl  = New-Object Drawing.Font('Segoe UI',13,[Drawing.FontStyle]::Bold)
#endregion

#region SCRIPT-SCOPE STATE (all controls + data at script scope = no scope bugs)
$script:ActiveTab   = 'Software'
$script:UseAltCreds = $false
$script:AltCred     = $null
$script:PSExecPath  = ''
$script:AllResults  = [System.Collections.ArrayList]::new()

# UI controls - declared here so ALL event handlers can access them
$script:Grid       = $null
$script:TxtQ       = $null
$script:TxtTarget  = $null
$script:BtnSearch  = $null
$script:LblCount   = $null
$script:ProgBar    = $null
$script:StatPanel  = $null   # Panel used as status bar
$script:StatLabel  = $null   # Label INSIDE the panel (Label always has .Text)
$script:TrkThr     = $null
$script:NumTO      = $null
$script:ChkPing    = $null
$script:ScopeRBs   = @{}
$script:TabBtns    = @{}
$script:BtnCred    = $null
#endregion

#region HELPER: Create Label with explicit BG (fixes WinForms transparency bug)
function New-Lbl {
    param([string]$Text,[int]$X,[int]$Y,[int]$W,[int]$H,
          [Drawing.Color]$FG,[Drawing.Color]$BG,
          [Drawing.Font]$Font=$FUI,
          [Windows.Forms.ContentAlignment]$Align='MiddleLeft')
    $ctrl = New-Object Windows.Forms.Label
    $ctrl.Text=$Text; $ctrl.Location=[Drawing.Point]::new($X,$Y); $ctrl.Size=[Drawing.Size]::new($W,$H)
    $ctrl.ForeColor=$FG; $ctrl.BackColor=$BG; $ctrl.Font=$Font; $ctrl.TextAlign=$Align
    return $ctrl
}

function New-Btn {
    param([string]$Text,[int]$X,[int]$Y,[int]$W,[int]$H,
          [Drawing.Color]$BG,[Drawing.Color]$BGH,
          [Drawing.Color]$FG=$CB0,[Drawing.Font]$Font=$FUIB,
          [string]$Tip='',[object]$TC=$null)
    $ctrl = New-Object Windows.Forms.Button
    $ctrl.Text=$Text; $ctrl.Location=[Drawing.Point]::new($X,$Y); $ctrl.Size=[Drawing.Size]::new($W,$H)
    $ctrl.BackColor=$BG; $ctrl.ForeColor=$FG; $ctrl.Font=$Font
    $ctrl.FlatStyle='Flat'; $ctrl.FlatAppearance.BorderSize=0
    $ctrl.FlatAppearance.MouseOverBackColor=$BGH
    $ctrl.FlatAppearance.MouseDownBackColor=$CAccD
    $ctrl.Cursor=[Windows.Forms.Cursors]::Hand
    if ($Tip -and $TC) { $TC.SetToolTip($ctrl,$Tip) }
    return $ctrl
}

function New-HR { param([int]$X,[int]$Y,[int]$W,[Drawing.Color]$BG=$CBdr)
    $p=New-Object Windows.Forms.Panel; $p.Location=[Drawing.Point]::new($X,$Y)
    $p.Size=[Drawing.Size]::new($W,1); $p.BackColor=$BG; return $p
}
#endregion

#region STATUS BAR
function Set-Status {
    param([string]$Msg,[string]$Type='Info')
    if (-not $script:StatLabel) { return }
    $script:StatLabel.Text = "  $Msg"
    if ($Type -eq 'OK')   { $script:StatLabel.ForeColor = $CGrn; return }
    if ($Type -eq 'Warn') { $script:StatLabel.ForeColor = $CYel; return }
    if ($Type -eq 'Err')  { $script:StatLabel.ForeColor = $CRed; return }
    $script:StatLabel.ForeColor = $CDim
    [Windows.Forms.Application]::DoEvents()
}
#endregion

#region QUERY: auto-add wildcards
function Get-Filter { param([string]$q)
    $q = $q.Trim()
    if ($q -eq '') { return '*' }
    if ($q -notmatch '\*') { return "*$q*" }
    return $q
}
#endregion

#region LOCAL QUERIES (plain functions, no closures, no scope issues)
function Search-LocalSoftware { param([string]$q)
    $f = Get-Filter $q
    $keys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $out = [System.Collections.ArrayList]::new()
    $raw = Get-ItemProperty $keys -ErrorAction SilentlyContinue |
           Where-Object { $_.DisplayName -and ($_.DisplayName -like $f) }
    foreach ($item in @($raw)) {
        [void]$out.Add([PSCustomObject]@{
            Name        = "$($item.DisplayName)"
            Version     = "$($item.DisplayVersion)"
            Publisher   = "$($item.Publisher)"
            InstallDate = "$($item.InstallDate)"
            InstallPath = "$($item.InstallLocation)"
            Computer    = "$env:COMPUTERNAME"
            Via         = 'Local'
            _Uninstall  = "$($item.UninstallString)"
            _Quiet      = "$($item.QuietUninstallString)"
        })
    }
    return $out
}

function Search-LocalServices { param([string]$q)
    $f = Get-Filter $q
    $out = [System.Collections.ArrayList]::new()
    $raw = Get-WmiObject Win32_Service -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -like $f -or $_.DisplayName -like $f }
    foreach ($item in @($raw)) {
        [void]$out.Add([PSCustomObject]@{
            Name        = "$($item.Name)"
            DisplayName = "$($item.DisplayName)"
            State       = "$($item.State)"
            StartMode   = "$($item.StartMode)"
            RunAs       = "$($item.StartName)"
            Binary      = "$($item.PathName)"
            Computer    = "$env:COMPUTERNAME"
            Via         = 'Local'
        })
    }
    return $out
}

function Search-LocalProcesses { param([string]$q)
    $f = Get-Filter $q
    $out = [System.Collections.ArrayList]::new()
    $raw = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $f }
    foreach ($item in @($raw)) {
        $pth = ''
        try { $pth = "$($item.Path)" } catch {}
        [void]$out.Add([PSCustomObject]@{
            PID      = $item.Id
            Name     = "$($item.Name)"
            CPU_s    = [math]::Round($item.CPU,1)
            Mem_MB   = [math]::Round($item.WorkingSet64/1MB,1)
            Path     = $pth
            Since    = "$($item.StartTime)"
            Computer = "$env:COMPUTERNAME"
            Via      = 'Local'
        })
    }
    return $out
}

function Search-LocalTasks { param([string]$q)
    $f = Get-Filter $q
    $out = [System.Collections.ArrayList]::new()
    $raw = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like $f }
    foreach ($item in @($raw)) {
        $info = Get-ScheduledTaskInfo -TaskName $item.TaskName -TaskPath $item.TaskPath -ErrorAction SilentlyContinue
        [void]$out.Add([PSCustomObject]@{
            TaskName = "$($item.TaskName)"
            Path     = "$($item.TaskPath)"
            State    = "$($item.State)"
            RunAs    = "$($item.Principal.UserId)"
            LastRun  = "$($info.LastRunTime)"
            NextRun  = "$($info.NextRunTime)"
            Computer = "$env:COMPUTERNAME"
            Via      = 'Local'
        })
    }
    return $out
}
#endregion

#region REMOTE EXECUTION
function Invoke-Remote { param([string]$Computer,[scriptblock]$Block,[object[]]$ArgList=@())
    $cp = @{ComputerName=$Computer;ScriptBlock=$Block;ArgumentList=$ArgList;ErrorAction='Stop'}
    if ($script:UseAltCreds -and $script:AltCred) { $cp['Credential']=$script:AltCred }
    try {
        return @{OK=$true;Data=(Invoke-Command @cp);Via='WinRM'}
    } catch { $e="$_" }
    if ($script:PSExecPath) {
        try {
            $enc=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Block.ToString()))
            $r=& $script:PSExecPath "\\$Computer" -accepteula -nobanner -s powershell.exe -NonInteractive -EncodedCommand $enc 2>&1
            return @{OK=$true;Data=$r;Via='PSExec'}
        } catch { return @{OK=$false;Error="WinRM:$e PSExec:$_"} }
    }
    return @{OK=$false;Error=$e}
}

function Search-RemoteSoftware { param([string]$Computer,[string]$q)
    $f = Get-Filter $q
    $blk = { param($f)
        $keys=@('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
        Get-ItemProperty $keys -EA SilentlyContinue |
            Where-Object {$_.DisplayName -and $_.DisplayName -like $f} |
            Select-Object @{N='Name';E={"$($_.DisplayName)"}},@{N='Version';E={"$($_.DisplayVersion)"}},
                          @{N='Publisher';E={"$($_.Publisher)"}},@{N='InstallDate';E={"$($_.InstallDate)"}},
                          @{N='InstallPath';E={"$($_.InstallLocation)"}},
                          @{N='_Uninstall';E={"$($_.UninstallString)"}},@{N='_Quiet';E={"$($_.QuietUninstallString)"}}
    }
    $res=Invoke-Remote $Computer $blk @($f)
    if ($res.OK -and $res.Data) {
        return $res.Data | Select-Object *,@{N='Computer';E={$Computer}},@{N='Via';E={$res.Via}}
    }
    return @()
}

function Search-RemoteServices { param([string]$Computer,[string]$q)
    $f = Get-Filter $q
    $blk = { param($f)
        Get-WmiObject Win32_Service -EA SilentlyContinue |
            Where-Object {$_.Name -like $f -or $_.DisplayName -like $f} |
            Select-Object Name,DisplayName,State,StartMode,@{N='RunAs';E={"$($_.StartName)"}},@{N='Binary';E={"$($_.PathName)"}}
    }
    $res=Invoke-Remote $Computer $blk @($f)
    if ($res.OK -and $res.Data) {
        return $res.Data | Select-Object *,@{N='Computer';E={$Computer}},@{N='Via';E={$res.Via}}
    }
    return @()
}

function Search-RemoteProcesses { param([string]$Computer,[string]$q)
    $f = Get-Filter $q
    $blk = { param($f)
        Get-Process -EA SilentlyContinue | Where-Object {$_.Name -like $f} |
            Select-Object @{N='PID';E={$_.Id}},Name,@{N='CPU_s';E={[math]::Round($_.CPU,1)}},
                          @{N='Mem_MB';E={[math]::Round($_.WorkingSet64/1MB,1)}}
    }
    $res=Invoke-Remote $Computer $blk @($f)
    if ($res.OK -and $res.Data) {
        return $res.Data | Select-Object *,@{N='Computer';E={$Computer}},@{N='Via';E={$res.Via}}
    }
    return @()
}

function Search-RemoteTasks { param([string]$Computer,[string]$q)
    $f = Get-Filter $q
    $blk = { param($f)
        Get-ScheduledTask -EA SilentlyContinue | Where-Object {$_.TaskName -like $f} |
            Select-Object TaskName,@{N='Path';E={$_.TaskPath}},@{N='State';E={"$($_.State)"}},@{N='RunAs';E={$_.Principal.UserId}}
    }
    $res=Invoke-Remote $Computer $blk @($f)
    if ($res.OK -and $res.Data) {
        return $res.Data | Select-Object *,@{N='Computer';E={$Computer}},@{N='Via';E={$res.Via}}
    }
    return @()
}
#endregion

#region AD TARGETS
function Get-ScopeTargets { param([string]$Scope,[string]$Input)
    if ($Scope -eq 'Local')  { return @('LOCAL') }
    if ($Scope -eq 'Remote') { return @($Input.Trim()) }
    Import-Module ActiveDirectory -EA SilentlyContinue
    if ($Scope -eq 'Domain') { return @(Get-ADComputer -Filter * -EA SilentlyContinue | Select-Object -Expand Name) }
    if ($Scope -eq 'OU')     { return @(Get-ADComputer -SearchBase $Input -Filter * -EA SilentlyContinue | Select-Object -Expand Name) }
    if ($Scope -eq 'Group')  { return @(Get-ADGroupMember $Input -EA SilentlyContinue | Where-Object objectClass -eq 'computer' | Select-Object -Expand Name) }
    return @()
}
#endregion

#region PRE-FLIGHT
function Show-PreFlight {
    # Gather checks
    $psv   = $PSVersionTable.PSVersion
    $wm    = Get-Service WinRM -EA SilentlyContinue
    $ad    = Get-Module -ListAvailable ActiveDirectory -EA SilentlyContinue
    $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrator')
    $pxp   = $null
    foreach ($p in @("$PSScriptRoot\psexec.exe","$env:windir\System32\psexec.exe")) { if (Test-Path $p) { $pxp=$p; break } }
    $pxc   = Get-Command psexec.exe -EA SilentlyContinue
    if (-not $pxp -and $pxc) { $pxp = $pxc.Source }
    $script:PSExecPath = "$pxp"

    $chks = @(
        @{Lbl='PowerShell 5.1+';        OK=($psv.Major -ge 5);  Opt=$false; Det="Detected v$($psv.Major).$($psv.Minor)";   Fix=$null; FH=''}
        @{Lbl='WinRM Service';          OK=($wm -and $wm.Status-eq'Running'); Opt=$false
          Det=if($wm){"Status: $($wm.Status)"}else{'Not found'}
          Fix={Enable-PSRemoting -Force -SkipNetworkProfileCheck|Out-Null}; FH='Enable-PSRemoting -Force'}
        @{Lbl='RSAT: Active Directory'; OK=($null -ne $ad); Opt=$false
          Det=if($ad){"v$($ad.Version)"}else{'Not installed'}
          Fix={$os=(Get-WmiObject Win32_OperatingSystem -EA SilentlyContinue).Caption
               if($os-match'Server'){Add-WindowsFeature RSAT-AD-PowerShell|Out-Null}
               else{Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'|Out-Null}}
          FH='Add-WindowsCapability RSAT.AD'}
        @{Lbl='PSExec (optional)';      OK=($null -ne $pxp); Opt=$true
          Det=if($pxp){"$pxp"}else{'Not found - WinRM used (OK)'}; Fix=$null; FH='learn.microsoft.com/sysinternals/downloads/psexec'}
        @{Lbl='Running as Admin';       OK=$admin; Opt=$true
          Det=if($admin){'Elevated'}else{'Not elevated - some actions may fail'}; Fix=$null; FH='Right-click -> Run as Administrator'}
    )
    $critFail = @($chks | Where-Object { -not $_.OK -and -not $_.Opt }).Count

    $dlg = New-Object Windows.Forms.Form
    $dlg.Text='AppHunter Pro - Pre-Flight Check'; $dlg.Size=[Drawing.Size]::new(820,630)
    $dlg.StartPosition='CenterScreen'; $dlg.BackColor=$CB1
    $dlg.FormBorderStyle='FixedDialog'; $dlg.MaximizeBox=$false; $dlg.MinimizeBox=$false

    $tip = New-Object Windows.Forms.ToolTip; $tip.InitialDelay=300

    # Header
    $hp=New-Object Windows.Forms.Panel; $hp.Dock='Top'; $hp.Height=64; $hp.BackColor=$CB2
    $dlg.Controls.Add($hp)
    $hp.Controls.Add((New-Lbl 'AppHunter Pro' 14 8 300 28 $CAcc $CB2 $FTtl))
    $hp.Controls.Add((New-Lbl 'Pre-Flight Dependency Check' 16 38 400 20 $CDim $CB2 $FSm))
    $hp.Controls.Add((New-Lbl "PS $($psv.Major).$($psv.Minor)" 620 20 96 28 $CB0 $CAcc $FSmB 'MiddleCenter'))

    # Pre-flight rows: Labels added directly to the dialog form with absolute positions
    # Adding to form (not a panel) with explicit BackColor is the ONLY reliable method
    $ry = 72
    foreach ($chk in $chks) {
        $rowBG  = if ($chk.OK) { [Drawing.Color]::FromArgb(4,50,20) } elseif ($chk.Opt) { [Drawing.Color]::FromArgb(50,40,0) } else { [Drawing.Color]::FromArgb(62,10,10) }
        $pillFG = if ($chk.OK) { [Drawing.Color]::FromArgb(0,210,110) } elseif ($chk.Opt) { [Drawing.Color]::FromArgb(255,185,0) } else { [Drawing.Color]::FromArgb(215,50,60) }
        $nameFG = [Drawing.Color]::FromArgb(255,255,255)
        $detFG  = if ($chk.OK) { [Drawing.Color]::FromArgb(130,210,130) } elseif ($chk.Opt) { [Drawing.Color]::FromArgb(210,180,90) } else { [Drawing.Color]::FromArgb(210,130,130) }

        # Add text labels FIRST, then background - or skip bg label entirely
        # Each label gets the row BG color directly, covering the full area needed

        # Status pill
        $pillLbl = New-Object Windows.Forms.Label
        $pillLbl.Text      = if ($chk.OK) { ' PASS ' } elseif ($chk.Opt) { ' SKIP ' } else { ' FAIL ' }
        $pillLbl.Location  = [Drawing.Point]::new(10, $ry + 20)
        $pillLbl.Size      = [Drawing.Size]::new(54, 24)
        $pillLbl.BackColor = $pillFG
        $pillLbl.ForeColor = [Drawing.Color]::FromArgb(10,12,18)
        $pillLbl.Font      = New-Object Drawing.Font('Segoe UI', 8, [Drawing.FontStyle]::Bold)
        $pillLbl.TextAlign = 'MiddleCenter'
        $dlg.Controls.Add($pillLbl)
        $pillLbl.BringToFront()

        # Left edge strip (decorative bar)
        $barLbl = New-Object Windows.Forms.Label
        $barLbl.Location  = [Drawing.Point]::new(0, $ry)
        $barLbl.Size      = [Drawing.Size]::new(4, 64)
        $barLbl.BackColor = $pillFG
        $barLbl.Text      = ''
        $dlg.Controls.Add($barLbl)
        $barLbl.BringToFront()

        # Name label - spans from pill to right
        $nameLbl = New-Object Windows.Forms.Label
        $nameLbl.Text      = $chk.Lbl
        $nameLbl.Location  = [Drawing.Point]::new(72, $ry + 4)
        $nameLbl.Size      = [Drawing.Size]::new(460, 26)
        $nameLbl.BackColor = $rowBG
        $nameLbl.ForeColor = $nameFG
        $nameLbl.Font      = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
        $nameLbl.TextAlign = 'MiddleLeft'
        $dlg.Controls.Add($nameLbl)
        $nameLbl.BringToFront()

        # Detail label
        $detLbl = New-Object Windows.Forms.Label
        $detLbl.Text      = $chk.Det
        $detLbl.Location  = [Drawing.Point]::new(72, $ry + 34)
        $detLbl.Size      = [Drawing.Size]::new(460, 22)
        $detLbl.BackColor = $rowBG
        $detLbl.ForeColor = $detFG
        $detLbl.Font      = New-Object Drawing.Font('Segoe UI', 8, [Drawing.FontStyle]::Regular)
        $detLbl.TextAlign = 'MiddleLeft'
        $dlg.Controls.Add($detLbl)
        $detLbl.BringToFront()

        # Fill gaps on left (4-72) and right of pill row with row BG
        $leftFill = New-Object Windows.Forms.Label
        $leftFill.Location = [Drawing.Point]::new(4, $ry)
        $leftFill.Size     = [Drawing.Size]::new(736, 64)
        $leftFill.BackColor= $rowBG; $leftFill.Text=''
        $dlg.Controls.Add($leftFill)
        # Send background fill to back so text labels show on top
        $leftFill.SendToBack()

        # Auto-fix button for failed critical checks
        if (-not $chk.OK -and $chk.Fix) {
            $fa = $chk.Fix; $fh = $chk.FH
            $fb = New-Btn 'Auto-Fix' 554 ($ry + 16) 130 32 [Drawing.Color]::FromArgb(255,185,0) [Drawing.Color]::FromArgb(255,210,50) ([Drawing.Color]::FromArgb(10,12,18)) (New-Object Drawing.Font('Segoe UI',9,[Drawing.FontStyle]::Bold)) "Run: $fh" $tip
            $fb.Add_Click({
                $this.Enabled = $false; $this.Text = 'Running...'; $this.BackColor = [Drawing.Color]::FromArgb(95,112,145)
                [Windows.Forms.Application]::DoEvents()
                try   { & $fa; $this.Text = 'Done!';   $this.BackColor = [Drawing.Color]::FromArgb(0,210,110) }
                catch { $this.Text = 'Failed'; $this.BackColor = [Drawing.Color]::FromArgb(215,50,60)
                        [Windows.Forms.MessageBox]::Show("Error:`n$_",'Fix Failed','OK','Error') | Out-Null }
            }.GetNewClosure())
            $dlg.Controls.Add($fb)
        } elseif ($chk.FH -ne '') {
            $hintLbl = New-Object Windows.Forms.Label
            $hintLbl.Text      = $chk.FH
            $hintLbl.Location  = [Drawing.Point]::new(420, $ry + 22)
            $hintLbl.Size      = [Drawing.Size]::new(300, 20)
            $hintLbl.BackColor = $rowBG
            $hintLbl.ForeColor = [Drawing.Color]::FromArgb(150,148,70)
            $hintLbl.Font      = New-Object Drawing.Font('Segoe UI', 8, [Drawing.FontStyle]::Regular)
            $hintLbl.TextAlign = 'MiddleLeft'
            $dlg.Controls.Add($hintLbl)
        }

        $ry += 70
    }

    # Summary strip
    $sBG=if($critFail-eq 0){[Drawing.Color]::FromArgb(0,46,18)}else{[Drawing.Color]::FromArgb(50,36,0)}
    $sFG=if($critFail-eq 0){$CGrn}else{$CYel}
    $sTx=if($critFail-eq 0){'All critical checks passed - AppHunter Pro is ready!'}else{"$critFail critical check(s) failed. Some features limited."}
    $sp=New-Object Windows.Forms.Panel; $sp.Location=[Drawing.Point]::new(0,442)
    $sp.Size=[Drawing.Size]::new(740,34); $sp.BackColor=$sBG
    $sp.Controls.Add((New-Lbl $sTx 10 4 694 28 $sFG $sBG $FUIB)); $dlg.Controls.Add($sp)

    $bL=New-Btn 'Launch AppHunter' 430 490 206 44 $CAcc $CAccH $CB0 $FUIB 'Open the main window' $tip
    $bE=New-Btn 'Exit'              644 490  82 44 $CRed $CRedH $CWht $FUIB 'Exit' $tip
    $dlg.Controls.AddRange(@($bL,$bE))

    $script:pfResult=$false
    $bL.Add_Click({$script:pfResult=$true;$dlg.Close()})
    $bE.Add_Click({$script:pfResult=$false;$dlg.Close()})
    $dlg.Add_KeyDown({if($_.KeyCode-eq'Escape'){$script:pfResult=$false;$dlg.Close()}})
    [void]$dlg.ShowDialog()
    return $script:pfResult
}
#endregion

#region GRID OPERATIONS
function Update-Grid { param([object[]]$Data)
    $script:Grid.Columns.Clear(); $script:Grid.Rows.Clear(); $script:AllResults.Clear()

    if (-not $Data -or $Data.Count -eq 0) {
        $script:LblCount.Text='0 results'
        Set-Status 'No results found. Try a different search term.' 'Warn'
        return
    }

    foreach ($item in $Data) { [void]$script:AllResults.Add($item) }
    $script:LblCount.Text="$($Data.Count) result$(if($Data.Count-ne 1){'s'})"

    # Add checkbox column first - fixed width so it never collapses
    $chkCol=New-Object Windows.Forms.DataGridViewCheckBoxColumn
    $chkCol.HeaderText='[/]'; $chkCol.Name='_Select'
    $chkCol.Width=40; $chkCol.MinimumWidth=40
    $chkCol.AutoSizeMode='None'   # never auto-size this column
    $chkCol.Resizable='False'
    [void]$script:Grid.Columns.Add($chkCol)

    # Data columns: skip _prefixed (internal data)
    $props = $Data[0].PSObject.Properties |
             Where-Object { $_.MemberType -in 'NoteProperty','Property' -and $_.Name -notmatch '^_' }

    foreach ($p in $props) {
        $col=New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText=$p.Name; $col.Name=$p.Name; $col.SortMode='Automatic'
        $col.ReadOnly=$true
        [void]$script:Grid.Columns.Add($col)
    }
    foreach ($item in $Data) {
        $cells=@($false)  # checkbox starts unchecked
        foreach ($p in $props) { $cells+="$($item.($p.Name))" }
        [void]$script:Grid.Rows.Add($cells)
    }
    # Re-enforce checkbox column fixed width (Fill mode resets it)
    $script:Grid.Columns['_Select'].AutoSizeMode = 'None'
    $script:Grid.Columns['_Select'].Width = 40
    $script:Grid.Columns['_Select'].MinimumWidth = 40
    # Color-code service state column
    foreach ($row in $script:Grid.Rows) {
        $stCell=$row.Cells['State']
        if ($stCell) {
            if ($stCell.Value-eq'Running') { $row.DefaultCellStyle.ForeColor=$CGrn }
            if ($stCell.Value-eq'Stopped') { $row.DefaultCellStyle.ForeColor=[Drawing.Color]::FromArgb(200,100,100) }
        }
    }
    Set-Status "Found $($Data.Count) result(s)." 'OK'
}

function Get-SelectedRows {
    # Return checked rows first; fall back to selected (highlighted) rows if none checked
    $out=[System.Collections.ArrayList]::new()
    $checkedIdxs=@()
    foreach ($gridRow in $script:Grid.Rows) {
        $chkCell=$gridRow.Cells['_Select']
        if ($chkCell -and $chkCell.Value -eq $true) { $checkedIdxs+=$gridRow.Index }
    }
    if ($checkedIdxs.Count -gt 0) {
        foreach ($idx in $checkedIdxs) {
            if ($idx -ge 0 -and $idx -lt $script:AllResults.Count) { [void]$out.Add($script:AllResults[$idx]) }
        }
    } else {
        # Fall back to row selection (highlight)
        foreach ($gridRow in $script:Grid.SelectedRows) {
            $idx=$gridRow.Index
            if ($idx -ge 0 -and $idx -lt $script:AllResults.Count) { [void]$out.Add($script:AllResults[$idx]) }
        }
    }
    return @($out)
}
#endregion

#region MAIN FORM
function Show-Main {
    $tip=New-Object Windows.Forms.ToolTip; $tip.InitialDelay=350; $tip.AutoPopDelay=9000; $tip.ShowAlways=$true

    $frm=New-Object Windows.Forms.Form
    $frm.Text='AppHunter Pro  v5.0  |  Software & Service Discovery'
    $frm.Size=[Drawing.Size]::new(1240,840); $frm.MinimumSize=[Drawing.Size]::new(1050,740)
    $frm.StartPosition='CenterScreen'; $frm.BackColor=$CB0; $frm.ForeColor=$CTxt; $frm.Font=$FUI

    #-- Header
    $hdr=New-Object Windows.Forms.Panel; $hdr.Dock='Top'; $hdr.Height=56; $hdr.BackColor=$CB2
    $frm.Controls.Add($hdr)
    $hdr.Controls.Add((New-Lbl 'AppHunter Pro' 14 6 270 28 $CAcc $CB2 $FTtl))
    $hdr.Controls.Add((New-Lbl 'Enterprise Software & Service Discovery  v5.0' 16 34 440 18 $CDim $CB2 $FSm))

    # Elevate to Admin button (only shown if not already admin)
    $isAdmin=([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrator')
    if(-not $isAdmin){
        $btnElev=New-Btn 'Run as Admin' 500 10 148 36 $CRed $CRedH $CWht $FSmB 'Relaunch this script elevated as Administrator - required for some remote actions' $tip
        $btnElev.Add_Click({
            $ans=[Windows.Forms.MessageBox]::Show("Relaunch AppHunter as Administrator?`n`nThis will close the current window and open a new elevated session.",'Elevate?','YesNo','Question')
            if($ans-eq'Yes'){
                Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
                $frm.Close()
            }
        })
        $hdr.Controls.Add($btnElev)
    }

    # Credentials button - alternate credentials for remote connections
    $script:BtnCred=New-Btn 'Run As: Current User' 656 10 260 36 $CB3 $CAcc $CAcc $FSm 'Click to use alternate domain credentials for remote queries (e.g. admin account). Currently using your logged-in session.' $tip
    $script:BtnCred.FlatAppearance.BorderSize=1; $script:BtnCred.FlatAppearance.BorderColor=$CAcc
    $script:BtnCred.TextAlign='MiddleCenter'
    $hdr.Controls.Add($script:BtnCred)

    $script:BtnCred.Add_Click({
        if ($script:UseAltCreds) {
            $script:UseAltCreds=$false; $script:AltCred=$null
            $script:BtnCred.Text='Run As: Current User'; $script:BtnCred.ForeColor=$CAcc; $script:BtnCred.FlatAppearance.BorderColor=$CAcc
            Set-Status 'Using current session credentials.' 'Info'
        } else {
            $c=Get-Credential -Message 'Enter credentials for remote connections' -EA SilentlyContinue
            if ($c) {
                $script:UseAltCreds=$true; $script:AltCred=$c
                $script:BtnCred.Text="Run As: $($c.UserName)"; $script:BtnCred.ForeColor=$CYel; $script:BtnCred.FlatAppearance.BorderColor=$CYel
                Set-Status "Alternate credentials: $($c.UserName)" 'Warn'
            }
        }
    })

    #-- Sidebar
    $side=New-Object Windows.Forms.Panel; $side.Location=[Drawing.Point]::new(0,56)
    $side.Size=[Drawing.Size]::new(232,784); $side.BackColor=$CB1; $frm.Controls.Add($side)
    $side.Anchor='Top,Bottom,Left'; $side.Controls.Add((New-Lbl 'SEARCH TYPE' 12 12 208 16 $CDim $CB1 $FSm))
    $side.Controls.Add((New-HR 8 30 216))

    $tabDefs=[ordered]@{'Software'='Installed Software';'Services'='Windows Services';'Processes'='Running Processes';'Tasks'='Scheduled Tasks'}
    $ty=36
    foreach ($k in $tabDefs.Keys) {
        $tb=New-Object Windows.Forms.Button; $tb.Text=$tabDefs[$k]
        $tb.Location=[Drawing.Point]::new(4,$ty); $tb.Size=[Drawing.Size]::new(224,42)
        $tb.FlatStyle='Flat'; $tb.FlatAppearance.BorderSize=0
        $tb.BackColor=$CB1; $tb.ForeColor=$CDim; $tb.Font=$FUI
        $tb.TextAlign='MiddleLeft'; $tb.Padding=[Windows.Forms.Padding]::new(14,0,0,0)
        $tb.Cursor=[Windows.Forms.Cursors]::Hand
        $script:TabBtns[$k]=$tb; $side.Controls.Add($tb); $ty+=46
    }

    function Set-ActiveTab { param([string]$N)
        $script:ActiveTab=$N
        foreach ($k in $script:TabBtns.Keys) {
            if ($k-eq$N) { $script:TabBtns[$k].BackColor=$CAcc; $script:TabBtns[$k].ForeColor=$CB0; $script:TabBtns[$k].Font=$FUIB }
            else          { $script:TabBtns[$k].BackColor=$CB1;  $script:TabBtns[$k].ForeColor=$CDim; $script:TabBtns[$k].Font=$FUI }
        }
    }
    Set-ActiveTab 'Software'
    $script:TabBtns['Software'].Add_Click({ Set-ActiveTab 'Software' })
    $script:TabBtns['Services'].Add_Click({ Set-ActiveTab 'Services' })
    $script:TabBtns['Processes'].Add_Click({ Set-ActiveTab 'Processes' })
    $script:TabBtns['Tasks'].Add_Click({ Set-ActiveTab 'Tasks' })

    $tip.SetToolTip($script:TabBtns['Software'], 'Search installed apps (registry / Programs and Features)')
    $tip.SetToolTip($script:TabBtns['Services'], 'Search Windows services by name or display name')
    $tip.SetToolTip($script:TabBtns['Processes'],'Search running processes by executable name')
    $tip.SetToolTip($script:TabBtns['Tasks'],    'Search Task Scheduler tasks by name')

    $side.Controls.Add((New-HR 8 226 216))
    $side.Controls.Add((New-Lbl 'TARGET SCOPE' 12 234 208 16 $CDim $CB1 $FSm))

    $scopeDefs=[ordered]@{'Local'='Local Machine';'Remote'='Remote Host';'Domain'='Entire AD Domain';'OU'='Specific OU';'Group'='AD Group'}
    $sy=254
    foreach ($k in $scopeDefs.Keys) {
        $rb=New-Object Windows.Forms.RadioButton; $rb.Text=$scopeDefs[$k]
        $rb.Location=[Drawing.Point]::new(8,$sy); $rb.Size=[Drawing.Size]::new(216,28)
        $rb.ForeColor=$CTxt; $rb.BackColor=$CB1; $rb.Font=$FUI; $rb.Cursor=[Windows.Forms.Cursors]::Hand
        $script:ScopeRBs[$k]=$rb; $side.Controls.Add($rb); $sy+=32
    }
    $script:ScopeRBs['Local'].Checked=$true
    $tip.SetToolTip($script:ScopeRBs['Local'],  'Query this local machine only - instant, no network needed')
    $tip.SetToolTip($script:ScopeRBs['Remote'], 'Query one remote machine - enter hostname or IP below')
    $tip.SetToolTip($script:ScopeRBs['Domain'], 'Query ALL AD computers - requires RSAT AD module')
    $tip.SetToolTip($script:ScopeRBs['OU'],     'Query all computers in an OU - paste Distinguished Name below')
    $tip.SetToolTip($script:ScopeRBs['Group'],  'Query computer members of an AD Security Group')

    $side.Controls.Add((New-HR 8 422 216))
    $side.Controls.Add((New-Lbl 'Target (hostname / OU DN / Group):' 10 430 212 16 $CDim $CB1 $FSm))
    $script:TxtTarget=New-Object Windows.Forms.TextBox
    $script:TxtTarget.Location=[Drawing.Point]::new(8,448); $script:TxtTarget.Size=[Drawing.Size]::new(214,26)
    $script:TxtTarget.BackColor=$CB3; $script:TxtTarget.ForeColor=$CDim; $script:TxtTarget.Font=$FMono; $script:TxtTarget.BorderStyle='FixedSingle'
    $script:TxtTarget.Text='e.g. PC-12345'
    $script:TxtTarget.Add_GotFocus({ if($script:TxtTarget.ForeColor-eq$CDim){$script:TxtTarget.Text='';$script:TxtTarget.ForeColor=$CTxt} })
    $script:TxtTarget.Add_LostFocus({ if($script:TxtTarget.Text.Trim()-eq''){$script:TxtTarget.Text='e.g. PC-12345';$script:TxtTarget.ForeColor=$CDim} })
    $side.Controls.Add($script:TxtTarget)
    $tip.SetToolTip($script:TxtTarget,'Hostname/IP for Remote scope, OU Distinguished Name for OU, or Group name for Group scope')

    $side.Controls.Add((New-HR 8 482 216))
    $side.Controls.Add((New-Lbl 'DOMAIN OPTIONS' 12 490 208 16 $CDim $CB1 $FSm))
    $side.Controls.Add((New-Lbl 'Parallel threads:' 10 510 140 16 $CDim $CB1 $FSm))

    $lblThr=New-Lbl '20' 178 508 44 20 $CAcc $CB1 $FSmB 'MiddleRight'; $side.Controls.Add($lblThr)
    $script:TrkThr=New-Object Windows.Forms.TrackBar
    $script:TrkThr.Location=[Drawing.Point]::new(6,526); $script:TrkThr.Size=[Drawing.Size]::new(218,30)
    $script:TrkThr.Minimum=1; $script:TrkThr.Maximum=50; $script:TrkThr.Value=20; $script:TrkThr.TickFrequency=10; $script:TrkThr.BackColor=$CB1
    $script:TrkThr.Add_ValueChanged({ $lblThr.Text="$($script:TrkThr.Value)" })
    $side.Controls.Add($script:TrkThr)
    $tip.SetToolTip($script:TrkThr,'Machines to query simultaneously during domain-wide searches')

    $side.Controls.Add((New-Lbl 'Timeout per machine (sec):' 10 560 180 16 $CDim $CB1 $FSm))
    $script:NumTO=New-Object Windows.Forms.NumericUpDown
    $script:NumTO.Location=[Drawing.Point]::new(8,578); $script:NumTO.Size=[Drawing.Size]::new(72,26)
    $script:NumTO.Minimum=5; $script:NumTO.Maximum=120; $script:NumTO.Value=30
    $script:NumTO.BackColor=$CB3; $script:NumTO.ForeColor=$CTxt; $script:NumTO.Font=$FMono
    $side.Controls.Add($script:NumTO)
    $tip.SetToolTip($script:NumTO,'Max seconds to wait per machine before skipping')

    $script:ChkPing=New-Object Windows.Forms.CheckBox; $script:ChkPing.Text='Skip offline machines (ping)'
    $script:ChkPing.Location=[Drawing.Point]::new(8,608); $script:ChkPing.Size=[Drawing.Size]::new(216,22)
    $script:ChkPing.Checked=$true; $script:ChkPing.ForeColor=$CTxt; $script:ChkPing.BackColor=$CB1; $script:ChkPing.Font=$FSm
    $side.Controls.Add($script:ChkPing)
    $tip.SetToolTip($script:ChkPing,'Ping each machine before connecting - skips offline ones to avoid timeouts')

    #-- Main content
    $main=New-Object Windows.Forms.Panel; $main.Location=[Drawing.Point]::new(232,56)
    $main.Size=[Drawing.Size]::new(992,784); $main.BackColor=$CB0; $frm.Controls.Add($main)
    $main.Anchor='Top,Bottom,Left,Right'  # resize with form

    #-- Search bar panel
    $sb=New-Object Windows.Forms.Panel; $sb.Dock='Top'; $sb.Height=70; $sb.BackColor=$CB1
    $main.Controls.Add($sb)
    $sb.Controls.Add((New-Lbl 'SEARCH QUERY' 14 6 180 16 $CDim $CB1 $FSm))

    $script:TxtQ=New-Object Windows.Forms.TextBox
    $script:TxtQ.Location=[Drawing.Point]::new(14,26); $script:TxtQ.Size=[Drawing.Size]::new(500,32)
    $script:TxtQ.BackColor=$CB3; $script:TxtQ.ForeColor=$CTxt
    $script:TxtQ.Font=New-Object Drawing.Font('Consolas',11,[Drawing.FontStyle]::Regular)
    $script:TxtQ.BorderStyle='FixedSingle'; $sb.Controls.Add($script:TxtQ)
    $tip.SetToolTip($script:TxtQ,'Type name or partial name. Wildcards auto-added: "Airlock" searches as "*Airlock*". Case-insensitive.')

    $script:BtnSearch=New-Btn 'Search' 522 22 120 38 $CAcc $CAccH $CB0 $FUIB 'Run search (or press Enter). Wildcards added automatically.' $tip
    $sb.Controls.Add($script:BtnSearch)

    $btnClear=New-Btn 'Clear' 650 22 80 38 $CB3 $CBdr $CDim $FUI 'Clear results and query' $tip; $sb.Controls.Add($btnClear)
    $btnCSV=New-Btn 'Export CSV' 738 22 106 38 $CB3 $CBdr $CTxt $FUI 'Save results to CSV file' $tip; $sb.Controls.Add($btnCSV)

    $script:LblCount=New-Object Windows.Forms.Label
    $script:LblCount.Text='0 results'; $script:LblCount.Location=[Drawing.Point]::new(854,28)
    $script:LblCount.Size=[Drawing.Size]::new(128,20); $script:LblCount.ForeColor=$CDim; $script:LblCount.BackColor=$CB1
    $script:LblCount.Font=$FSm; $script:LblCount.TextAlign='MiddleRight'; $sb.Controls.Add($script:LblCount)

    #-- Progress bar
    $script:ProgBar=New-Object Windows.Forms.ProgressBar
    $script:ProgBar.Location=[Drawing.Point]::new(0,68); $script:ProgBar.Size=[Drawing.Size]::new(992,3)
    $script:ProgBar.Style='Marquee'; $script:ProgBar.MarqueeAnimationSpeed=18; $script:ProgBar.Visible=$false
    $main.Controls.Add($script:ProgBar); $script:ProgBar.BringToFront()

    #-- Status bar - PANEL containing a LABEL (Label.Text always works)
    $script:StatPanel=New-Object Windows.Forms.Panel
    $script:StatPanel.Dock='Bottom'; $script:StatPanel.Height=26; $script:StatPanel.BackColor=$CB2
    $main.Controls.Add($script:StatPanel)

    $script:StatLabel=New-Object Windows.Forms.Label
    $script:StatLabel.Dock='Fill'; $script:StatLabel.BackColor=$CB2; $script:StatLabel.ForeColor=$CDim
    $script:StatLabel.Font=$FSm; $script:StatLabel.TextAlign='MiddleLeft'
    $script:StatLabel.Text='  Ready  |  Choose search type, enter a name, select scope, click Search'
    $script:StatPanel.Controls.Add($script:StatLabel)

    #-- Grid
    $script:Grid=New-Object Windows.Forms.DataGridView
    $script:Grid.Location=[Drawing.Point]::new(0,72); $script:Grid.Size=[Drawing.Size]::new(992,576)
    $script:Grid.Anchor='Top,Bottom,Left,Right'; $script:Grid.BackgroundColor=$CB0; $script:Grid.GridColor=$CBdr; $script:Grid.BorderStyle='None'
    $script:Grid.ColumnHeadersDefaultCellStyle.BackColor=$CB2
    $script:Grid.ColumnHeadersDefaultCellStyle.ForeColor=$CWht
    $script:Grid.ColumnHeadersDefaultCellStyle.Font=$FUIB
    $script:Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor=$CB2
    $script:Grid.ColumnHeadersHeight=30; $script:Grid.EnableHeadersVisualStyles=$false
    $script:Grid.ColumnHeadersBorderStyle='Single'
    $script:Grid.DefaultCellStyle.BackColor=$CB0; $script:Grid.DefaultCellStyle.ForeColor=$CTxt
    $script:Grid.DefaultCellStyle.Font=$FMono
    $script:Grid.DefaultCellStyle.SelectionBackColor=$CSel; $script:Grid.DefaultCellStyle.SelectionForeColor=$CWht
    $script:Grid.AlternatingRowsDefaultCellStyle.BackColor=$CAlt
    $script:Grid.AutoSizeColumnsMode='Fill'; $script:Grid.SelectionMode='FullRowSelect'; $script:Grid.MultiSelect=$true
    $script:Grid.ReadOnly=$false; $script:Grid.AllowUserToAddRows=$false; $script:Grid.AllowUserToDeleteRows=$false
    $script:Grid.RowHeadersVisible=$false; $script:Grid.CellBorderStyle='SingleHorizontal'
    $main.Controls.Add($script:Grid)

    #-- Action panel
    $act=New-Object Windows.Forms.Panel; $act.Location=[Drawing.Point]::new(0,648)
    $act.Size=[Drawing.Size]::new(992,110);
    $act.Anchor='Bottom,Left,Right'; $act.BackColor=$CB2; $main.Controls.Add($act)
    $act.Controls.Add((New-Lbl 'ACTIONS  -  check box(es) in first column OR highlight row(s), then click action  |  Right-click for more options' 8 3 860 16 $CDim $CB2 $FSm))
    $act.Controls.Add((New-HR 0 20 992))

    $aY=26
    $btnUninst =New-Btn 'Uninstall'   6    $aY 148 38 $CRed  $CRedH $CWht $FUIB 'Uninstall selected app (uses built-in uninstaller, asks confirmation)' $tip
    $btnCpPath =New-Btn 'Copy Path'   160  $aY 126 38 $CB3   $CBdr  $CTxt $FUI  'Copy install path to clipboard' $tip
    $act.Controls.Add((New-HR 294 22 1 68 $CBdr))
    $btnSvcStart  =New-Btn 'Start'     300  $aY  90 38 $CGrn  $CGrnH $CB0  $FUIB 'Start selected Windows service' $tip
    $btnSvcStop   =New-Btn 'Stop'      396  $aY  82 38 $CYel  $CYelH $CB0  $FUIB 'Stop selected Windows service (asks confirmation)' $tip
    $btnSvcRestart=New-Btn 'Restart'  484  $aY 104 38 $CAccD $CAcc  $CWht $FUIB 'Restart selected service (Stop then Start)' $tip
    $btnSvcDel    =New-Btn 'Remove Service'   594  $aY 120 38 $CRed  $CRedH $CWht $FUIB 'Permanently delete service (sc.exe delete) - CANNOT be undone!' $tip
    $act.Controls.Add((New-HR 722 22 1 68 $CBdr))
    $btnKill   =New-Btn 'Kill Process'   728  $aY 126 38 $CRed  $CRedH $CWht $FUIB 'Force-terminate selected process immediately' $tip
    $btnReRun  =New-Btn 'Re-Run'       860  $aY 124 38 $CB3   $CBdr  $CTxt $FUI  'Re-run last search to refresh results' $tip

    $act.Controls.AddRange(@($btnUninst,$btnCpPath,$btnSvcStart,$btnSvcStop,$btnSvcRestart,$btnSvcDel,$btnKill,$btnReRun))
    $act.Controls.Add((New-Lbl '--- Software ---'   6   72 288 18 $CDim $CB2 $FSm 'MiddleCenter'))
    $act.Controls.Add((New-Lbl '--- Services ---'   300 72 422 18 $CDim $CB2 $FSm 'MiddleCenter'))
    $act.Controls.Add((New-Lbl '--- Process ---'    728 72 120 18 $CDim $CB2 $FSm 'MiddleCenter'))

    # Context menu with full actions
    $ctx=New-Object Windows.Forms.ContextMenuStrip; $ctx.BackColor=$CB2; $ctx.ForeColor=$CTxt; $ctx.Font=$FUI
    $miCpRow    = $ctx.Items.Add('Copy Row')
    $miCpCell   = $ctx.Items.Add('Copy Cell')
    $ctx.Items.Add((New-Object Windows.Forms.ToolStripSeparator))|Out-Null
    $miCtxUninst= $ctx.Items.Add('Uninstall Selected')
    $miCtxStart = $ctx.Items.Add('Start Service')
    $miCtxStop  = $ctx.Items.Add('Stop Service')
    $miCtxKill  = $ctx.Items.Add('Kill Process')
    $ctx.Items.Add((New-Object Windows.Forms.ToolStripSeparator))|Out-Null
    $miCheckAll  = $ctx.Items.Add('Check All Rows')
    $miUncheckAll= $ctx.Items.Add('Uncheck All Rows')
    $script:Grid.ContextMenuStrip=$ctx

    # Click checkbox column header to toggle all rows
    $script:Grid.Add_ColumnHeaderMouseClick({
        if($_.ColumnIndex -eq 0){
            $anyUnchecked=$false
            foreach($r in $script:Grid.Rows){ if($r.Cells['_Select'].Value -ne $true){$anyUnchecked=$true;break} }
            foreach($r in $script:Grid.Rows){ $r.Cells['_Select'].Value=$anyUnchecked }
            $script:Grid.RefreshEdit()
            Set-Status "All rows $(if($anyUnchecked){'checked'}else{'unchecked'}). Click [/] header again to toggle." 'Info'
        }
    })

    #-- SEARCH handler
    $script:BtnSearch.Add_Click({
        $q=$script:TxtQ.Text.Trim()
        if ($q -eq '') { [Windows.Forms.MessageBox]::Show('Enter a search term first.','No Query','OK','Information')|Out-Null; return }

        $scope='Local'
        foreach ($k in $script:ScopeRBs.Keys) { if ($script:ScopeRBs[$k].Checked) { $scope=$k; break } }
        $tIn=$script:TxtTarget.Text.Trim()
        if ($tIn -eq 'e.g. PC-12345') { $tIn='' }
        if ($scope -in 'Remote','OU','Group' -and $tIn -eq '') {
            [Windows.Forms.MessageBox]::Show("Enter a target for scope: $scope",'Missing Target','OK','Warning')|Out-Null; return
        }

        $script:BtnSearch.Enabled=$false; $script:ProgBar.Visible=$true
        $script:Grid.Columns.Clear(); $script:Grid.Rows.Clear(); $script:AllResults.Clear()
        $script:LblCount.Text='Searching...'
        Set-Status "Searching '$q' [$($script:ActiveTab)] on [$scope]..." 'Info'

        $allR=[System.Collections.ArrayList]::new()
        try {
            $targets=Get-ScopeTargets -Scope $scope -Input $tIn
            if (-not $targets -or $targets.Count -eq 0) { Set-Status 'No targets found. Check AD module.' 'Err'; return }

            if ($scope -in @('Local','Remote')) {
                foreach ($comp in $targets) {
                    Set-Status "Querying $comp..." 'Info'
                    $r=@()
                    if ($comp -eq 'LOCAL') {
                        if ($script:ActiveTab -eq 'Software')  { $r=@(Search-LocalSoftware  $q) }
                        if ($script:ActiveTab -eq 'Services')  { $r=@(Search-LocalServices  $q) }
                        if ($script:ActiveTab -eq 'Processes') { $r=@(Search-LocalProcesses $q) }
                        if ($script:ActiveTab -eq 'Tasks')     { $r=@(Search-LocalTasks     $q) }
                    } else {
                        if ($script:ActiveTab -eq 'Software')  { $r=@(Search-RemoteSoftware  $comp $q) }
                        if ($script:ActiveTab -eq 'Services')  { $r=@(Search-RemoteServices  $comp $q) }
                        if ($script:ActiveTab -eq 'Processes') { $r=@(Search-RemoteProcesses $comp $q) }
                        if ($script:ActiveTab -eq 'Tasks')     { $r=@(Search-RemoteTasks     $comp $q) }
                    }
                    foreach ($item in $r) { if ($item) { [void]$allR.Add($item) } }
                }
            } else {
                # Parallel domain search using Start-Job
                $bSz=[int]$script:TrkThr.Value; $tOut=[int]$script:NumTO.Value
                $doPing=$script:ChkPing.Checked; $tabNow=$script:ActiveTab
                $credNow=if($script:UseAltCreds -and $script:AltCred){$script:AltCred}else{$null}
                $q2=$q
                $queue=[System.Collections.Queue]::new($targets)
                $running=[System.Collections.ArrayList]::new()
                $done=0; $total=$targets.Count

                while ($queue.Count -gt 0 -or $running.Count -gt 0) {
                    while ($queue.Count -gt 0 -and $running.Count -lt $bSz) {
                        $comp=$queue.Dequeue()
                        if ($doPing -and -not (Test-Connection $comp -Count 1 -Quiet -TimeToLive 10 -EA SilentlyContinue)) { $done++; continue }
                        $job=Start-Job -ScriptBlock {
                            param($c,$q,$tab,$to,$cred)
                            $ErrorActionPreference='SilentlyContinue'
                            $so=New-PSSessionOption -OperationTimeout($to*1000) -OpenTimeout($to*1000)
                            $cp=@{ComputerName=$c;ErrorAction='Stop';SessionOption=$so}
                            if($cred){$cp['Credential']=$cred}
                            $f=if($q-notmatch'\*'){"*$q*"}else{$q}
                            $blk=if($tab-eq'Software'){{param($f)$keys=@('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*');Get-ItemProperty $keys -EA SilentlyContinue|Where-Object{$_.DisplayName -and $_.DisplayName-like$f}|Select-Object @{N='Name';E={"$($_.DisplayName)"}},@{N='Version';E={"$($_.DisplayVersion)"}},@{N='Publisher';E={"$($_.Publisher)"}},@{N='InstallPath';E={"$($_.InstallLocation)"}}}}elseif($tab-eq'Services'){{param($f)Get-WmiObject Win32_Service -EA SilentlyContinue|Where-Object{$_.Name-like$f -or $_.DisplayName-like$f}|Select-Object Name,DisplayName,State,StartMode}}elseif($tab-eq'Processes'){{param($f)Get-Process -EA SilentlyContinue|Where-Object{$_.Name-like$f}|Select-Object @{N='PID';E={$_.Id}},Name,@{N='CPU_s';E={[math]::Round($_.CPU,1)}},@{N='Mem_MB';E={[math]::Round($_.WorkingSet64/1MB,1)}}}}else{{param($f)Get-ScheduledTask -EA SilentlyContinue|Where-Object{$_.TaskName-like$f}|Select-Object TaskName,@{N='Path';E={$_.TaskPath}},@{N='State';E={"$($_.State)"}}}}
                            try{$r=Invoke-Command @cp -ScriptBlock $blk -ArgumentList $f;if($r){$r|Select-Object *,@{N='Computer';E={$c}}}}catch{}
                        } -ArgumentList $comp,$q2,$tabNow,$tOut,$credNow
                        [void]$running.Add(@{Job=$job;Comp=$comp})
                    }
                    $still=[System.Collections.ArrayList]::new()
                    foreach ($entry in $running) {
                        if ($entry.Job.State -in 'Completed','Failed','Stopped') {
                            $r=Receive-Job $entry.Job -EA SilentlyContinue
                            if ($r) { foreach ($item in @($r)) { if($item){[void]$allR.Add($item)} } }
                            Remove-Job $entry.Job -Force -EA SilentlyContinue
                            $done++; Set-Status "Progress: $done/$total machines..." 'Info'
                        } else { [void]$still.Add($entry) }
                    }
                    $running=$still
                    if ($running.Count -ge $bSz -or $queue.Count -eq 0) { Start-Sleep -Milliseconds 200 }
                    [Windows.Forms.Application]::DoEvents()
                }
            }
            Update-Grid @($allR)
        } catch {
            Set-Status "Error: $_" 'Err'
        } finally {
            $script:BtnSearch.Enabled=$true; $script:ProgBar.Visible=$false
        }
    })

    $script:TxtQ.Add_KeyDown({ if($_.KeyCode-eq'Return'){$script:BtnSearch.PerformClick()} })
    $btnReRun.Add_Click({ $script:BtnSearch.PerformClick() })

    $btnClear.Add_Click({
        $script:Grid.Columns.Clear(); $script:Grid.Rows.Clear(); $script:AllResults.Clear()
        $script:LblCount.Text='0 results'; $script:TxtQ.Text=''
        Set-Status 'Cleared.' 'Info'
    })

    $btnCSV.Add_Click({
        if ($script:AllResults.Count -eq 0) { Set-Status 'No results to export.' 'Warn'; return }
        $sv=New-Object Windows.Forms.SaveFileDialog; $sv.Title='Export to CSV'; $sv.Filter='CSV (*.csv)|*.csv'
        $sv.FileName="AppHunter_$($script:ActiveTab)_$(Get-Date -f 'yyyyMMdd_HHmmss').csv"
        $sv.InitialDirectory=[Environment]::GetFolderPath('Desktop')
        if ($sv.ShowDialog()-eq'OK') {
            @($script:AllResults)|Export-Csv $sv.FileName -NoTypeInformation -Encoding UTF8
            Set-Status "Exported $($script:AllResults.Count) rows to $($sv.FileName)" 'OK'
        }
    })

    $miCpRow.Add_Click({
        $rows=Get-SelectedRows
        if ($rows) {
            $lines=$rows|ForEach-Object{($_.PSObject.Properties|Where-Object{$_.Name-notmatch'^_'}|ForEach-Object{"$($_.Name)=$($_.Value)"})-join' | '}
            [Windows.Forms.Clipboard]::SetText(($lines-join"`n")); Set-Status 'Row(s) copied.' 'OK'
        }
    })
    $miCpCell.Add_Click({
        $cell=$script:Grid.CurrentCell
        if($cell -and $null-ne $cell.Value){[Windows.Forms.Clipboard]::SetText("$($cell.Value)");Set-Status 'Cell copied.' 'OK'}
    })
    $miCtxUninst.Add_Click({ $btnUninst.PerformClick() })
    $miCtxStart.Add_Click({  $btnSvcStart.PerformClick() })
    $miCtxStop.Add_Click({   $btnSvcStop.PerformClick() })
    $miCtxKill.Add_Click({   $btnKill.PerformClick() })
    $miCheckAll.Add_Click({
        foreach ($gridRow in $script:Grid.Rows) {
            $c=$gridRow.Cells['_Select']; if($c){$c.Value=$true}
        }
        $script:Grid.RefreshEdit(); Set-Status "All rows checked." 'Info'
    })
    $miUncheckAll.Add_Click({
        foreach ($gridRow in $script:Grid.Rows) {
            $c=$gridRow.Cells['_Select']; if($c){$c.Value=$false}
        }
        $script:Grid.RefreshEdit(); Set-Status "All rows unchecked." 'Info'
    })

    $btnUninst.Add_Click({
        $rows=Get-SelectedRows; if(-not $rows){Set-Status 'Select a software row.' 'Warn';return}
        foreach($r in $rows){
            $name=$r.Name; $comp=$r.Computer
            if(-not $name){continue}
            if([Windows.Forms.MessageBox]::Show("Uninstall '$name' on '$comp'?`nCannot be undone.",'Confirm','YesNo','Warning')-ne'Yes'){continue}
            $uStr=if($r._Quiet){$r._Quiet}elseif($r._Uninstall){$r._Uninstall}else{$null}
            if(-not $uStr){Set-Status "No uninstall string for '$name'" 'Err';continue}
            try{
                Set-Status "Uninstalling '$name' on '$comp' - please wait..." 'Info'
                [Windows.Forms.Application]::DoEvents()
                if($comp-eq$env:COMPUTERNAME){
                    $proc=Start-Process cmd.exe -ArgumentList "/c $uStr" -WindowStyle Hidden -Wait -PassThru
                    if($proc.ExitCode -eq 0){
                        [Windows.Forms.MessageBox]::Show("Successfully uninstalled:`n$name`n`nOn: $comp`nExit code: $($proc.ExitCode)",'Uninstall Complete','OK','Information')|Out-Null
                        Set-Status "Uninstalled: '$name' on '$comp' (exit code 0)" 'OK'
                    } else {
                        [Windows.Forms.MessageBox]::Show("Uninstall finished with exit code $($proc.ExitCode):`n$name`n`nOn: $comp`n`nSome uninstallers use non-zero codes even on success.`nVerify in Programs and Features.",'Uninstall Result','OK','Warning')|Out-Null
                        Set-Status "Uninstall exit code $($proc.ExitCode): '$name' on '$comp'" 'Warn'
                    }
                } else {
                    Invoke-Remote $comp {param($u)Start-Process cmd.exe "/c $u" -WindowStyle Hidden -Wait} @($uStr)|Out-Null
                    [Windows.Forms.MessageBox]::Show("Uninstall command sent to: $comp`nApp: $name`n`nVerify removal on the remote machine.",'Remote Uninstall Sent','OK','Information')|Out-Null
                    Set-Status "Uninstall command sent: '$name' on '$comp'" 'OK'
                }
            }catch{
                [Windows.Forms.MessageBox]::Show("Uninstall FAILED:`n$name`nOn: $comp`n`nError: $_",'Uninstall Failed','OK','Error')|Out-Null
                Set-Status "Uninstall failed: $_" 'Err'
            }
        }
    })
    $btnCpPath.Add_Click({
        $rows=Get-SelectedRows; if(-not $rows){Set-Status 'Select a row.' 'Warn';return}
        $paths=($rows|ForEach-Object{$_.InstallPath}|Where-Object{$_})-join"`n"
        if($paths){[Windows.Forms.Clipboard]::SetText($paths);Set-Status 'Path(s) copied.' 'OK'}else{Set-Status 'No path found.' 'Warn'}
    })
    $btnSvcStart.Add_Click({
        $rows=Get-SelectedRows; if(-not $rows){Set-Status 'Select a service row.' 'Warn';return}
        foreach($r in $rows){$n=$r.Name;$c=$r.Computer
            try{if($c-eq$env:COMPUTERNAME){Start-Service $n -EA Stop}else{Invoke-Remote $c {param($s)Start-Service $s} @($n)|Out-Null};Set-Status "Started: '$n'" 'OK'}
            catch{Set-Status "Failed to start '$n': $_" 'Err'}}
    })
    $btnSvcStop.Add_Click({
        $rows=Get-SelectedRows; if(-not $rows){Set-Status 'Select a service row.' 'Warn';return}
        foreach($r in $rows){$n=$r.Name;$c=$r.Computer
            if([Windows.Forms.MessageBox]::Show("Stop '$n' on '$c'?",'Confirm','YesNo','Warning')-ne'Yes'){continue}
            try{if($c-eq$env:COMPUTERNAME){Stop-Service $n -Force -EA Stop}else{Invoke-Remote $c {param($s)Stop-Service $s -Force} @($n)|Out-Null};Set-Status "Stopped: '$n'" 'OK'}
            catch{Set-Status "Failed to stop '$n': $_" 'Err'}}
    })
    $btnSvcRestart.Add_Click({
        $rows=Get-SelectedRows; if(-not $rows){Set-Status 'Select a service row.' 'Warn';return}
        foreach($r in $rows){$n=$r.Name;$c=$r.Computer
            try{if($c-eq$env:COMPUTERNAME){Restart-Service $n -Force -EA Stop}else{Invoke-Remote $c {param($s)Restart-Service $s -Force} @($n)|Out-Null};Set-Status "Restarted: '$n'" 'OK'}
            catch{Set-Status "Failed to restart '$n': $_" 'Err'}}
    })
    $btnSvcDel.Add_Click({
        $rows=Get-SelectedRows; if(-not $rows){Set-Status 'Select a service row.' 'Warn';return}
        foreach($r in $rows){$n=$r.Name;$c=$r.Computer
            if([Windows.Forms.MessageBox]::Show("[!!] DELETE service '$n' on '$c'?`nCannot be undone!",'Confirm','YesNo','Warning')-ne'Yes'){continue}
            try{if($c-eq$env:COMPUTERNAME){& sc.exe delete $n|Out-Null}else{Invoke-Remote $c {param($s)& sc.exe delete $s} @($n)|Out-Null};Set-Status "Deleted: '$n'" 'OK'}
            catch{Set-Status "Failed to delete '$n': $_" 'Err'}}
    })
    $btnKill.Add_Click({
        $rows=Get-SelectedRows; if(-not $rows){Set-Status 'Select a process row.' 'Warn';return}
        foreach($r in $rows){$pid2=$r.PID;$n=$r.Name;$c=$r.Computer
            if([Windows.Forms.MessageBox]::Show("Kill '$n' (PID $pid2) on '$c'?",'Confirm','YesNo','Warning')-ne'Yes'){continue}
            try{if($c-eq$env:COMPUTERNAME){Stop-Process -Id $pid2 -Force -EA Stop}else{Invoke-Remote $c {param($p)Stop-Process -Id $p -Force} @($pid2)|Out-Null};Set-Status "Killed: '$n' PID $pid2" 'OK'}
            catch{Set-Status "Failed to kill '$n': $_" 'Err'}}
    })

    Set-Status 'Ready  |  Select search type, type a name, choose scope, click Search' 'Info'
    [void]$frm.ShowDialog()
}
#endregion

#region ENTRY POINT
if (Show-PreFlight) { Show-Main }
#endregion
