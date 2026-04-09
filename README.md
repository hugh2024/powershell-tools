# powershell-tools

A collection of PowerShell GUI utilities for Windows IT and security operations. Built for sysadmins and security engineers managing Active Directory domain environments.

> **Author:** Hugh Gozlou &nbsp;|&nbsp; **License:** MIT &nbsp;|&nbsp; **GitHub:** [hugh2024/powershell-tools](https://github.com/hugh2024/powershell-tools)

---

## Tools

| Tool | Version | Description |
|------|---------|-------------|
| [ADLockoutManager](#adlockoutmanager) | v1.0 | Real-time AD account lockout investigation and response |
| [LocalAdminAuditor](#localadminauditor) | v9.0 | Audit local Administrators group across your AD domain |
| [NetworkSpeedChecker](#networkspeedchecker) | v4.4 | Test network speed to local and remote machines |
| [ServerHealthCheck](#serverhealthcheck) | v6.0 | 12-point health diagnostic for Windows Servers |
| [SecurityChecker](#securitychecker) | v2.6.9 | Scan subnets for security misconfigurations |
| [AppHunter Pro](#apphunter-pro) | v5.0 | Enterprise software & service discovery across your domain |

---

## Quick Run (no install)

Run any tool directly from PowerShell - no download or installation needed:

```powershell
# AD Lockout Manager
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/ADLockoutManager.ps1 | iex

# Local Administrator Auditor
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/LocalAdminAuditor-GUI.ps1 | iex

# Network Speed Checker
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/NetworkSpeedChecker_v4_4.ps1 | iex

# Server Health Check
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/ServerHealthCheck_6.ps1 | iex

# Security Checker
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/SecurityChecker-GUI-v2_6_9.ps1 | iex

# AppHunter Pro
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/AppHunter-Pro.ps1 | iex
```

> **Note:** Run PowerShell as Administrator for full functionality.

---

## AppHunter Pro

**Version:** v5.0 &nbsp;|&nbsp; **Requires:** PowerShell 5.1+, WinRM (remote queries), RSAT AD module (domain/OU/group scope)

A dark-themed PowerShell GUI tool for discovering, auditing, and managing installed software, Windows services, running processes, and scheduled tasks — across your local machine, a single remote host, or your entire Active Directory domain simultaneously.

Useful when you need to quickly find what is installed across your environment, check service states, or take action on multiple machines without scripting.

### What it does

- Searches installed software (registry / Programs and Features equivalent), Windows services, running processes, and scheduled tasks
- Targets local machine, single remote host, entire AD domain, specific OU, or AD security group
- Parallel domain-wide queries using background jobs — configurable thread count (1–50) and per-machine timeout
- Auto-wildcard search — type `airlock` or `zoom` and it searches as `*airlock*` automatically (case-insensitive)
- Checkbox column — check specific rows across multiple machines, then act on only those
- Uninstall apps remotely with exit code feedback and success/failure dialog
- Start, Stop, Restart, or permanently remove Windows services
- Force-terminate running processes
- Pre-flight dependency check on every launch (WinRM, RSAT, elevation status)
- One-click UAC elevation — relaunch as Administrator without closing
- Alternate credentials support for remote connections
- Right-click context menu with full action set
- Export all results to CSV

### Key Features

| Feature | Details |
|---------|---------|
| **Search Types** | Installed Software, Windows Services, Running Processes, Scheduled Tasks |
| **Target Scopes** | Local, Remote Host, Entire AD Domain, Specific OU, AD Security Group |
| **Parallel Queries** | Background jobs, 1–50 configurable threads, per-machine timeout (5–120s) |
| **Smart Search** | Wildcards auto-added, case-insensitive, partial name matching |
| **Checkbox Selection** | Check specific rows, act on only those — safe multi-machine operations |
| **Software Actions** | Uninstall with exit code dialog, copy install path |
| **Service Actions** | Start, Stop, Restart, Delete (sc.exe delete with confirmation) |
| **Process Actions** | Force-kill with confirmation |
| **Credentials** | Alternate domain credentials for remote connections |
| **Elevation** | One-click UAC relaunch as Administrator |
| **Export** | CSV with timestamp, full result set |
| **Pre-flight** | Auto-checks WinRM, RSAT, PSExec, elevation — offers to fix missing deps |

---

## ADLockoutManager

**Version:** v1.0 &nbsp;|&nbsp; **Requires:** RSAT ActiveDirectory module, Domain Admin or Account Operators

A modern dark-themed PowerShell GUI tool for investigating, tracking, and resolving Active Directory account lockouts in real time. Replaces LockoutStatus.exe with a full investigation and response workflow in one window.

### What it does

- Queries all Domain Controllers simultaneously for locked accounts
- Identifies the PDC Emulator and marks it in the DC list (`[PDC]`)
- Investigates lockout source via Event ID 4740 (which machine) and 4625 (which process/service)
- One-click unlock of individual accounts or bulk unlock all
- Real-time auto-monitor with configurable refresh interval
- Teams and Email alerting when new lockouts are detected
- Password policy viewer including Fine-Grained Password Policies (FGPP)
- Full audit log of every action with CSV export
- Hover tooltips on every control
- Auto-detects missing RSAT and offers to install it

### Event IDs Used

| Event ID | Source | What it reveals |
|----------|--------|----------------|
| 4740 | PDC Emulator | Account lockout - which machine sent bad credentials |
| 4625 | Source machine | Failed logon - which process or service used stale credentials |

### Key Features

- **Search** - Find any AD account by username or display name
- **Show All Locked** - List all currently locked accounts across all DCs
- **Investigate Source** - Trace lockout to source machine and exact process
- **Unlock Selected / Unlock All** - One-click resolution with confirmation
- **Auto-Monitor** - Continuous refresh with Teams/Email alerts on new lockouts
- **PDC Aware** - All Event 4740 queries run against the PDC Emulator
- **Password Policy** - View domain policy and all FGPPs in one screen
- **Audit Log** - Every search, unlock, and investigation logged with operator and timestamp

### Prerequisites

| Requirement | Notes |
|-------------|-------|
| PowerShell 5.1+ | Built into Windows 10/11 and Server 2016+ |
| RSAT - ActiveDirectory module | Auto-installs if missing (requires internet + admin) |
| Domain Admin or Account Operators | Required to read lockout events and unlock accounts |

### Quick Start

```powershell
# Run as Administrator
.\ADLockoutManager.ps1

# Or one-liner (no download needed)
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/ADLockoutManager.ps1 | iex
```

1. Tool opens and detects your AD environment automatically
2. Click **Show All Locked** to see all currently locked accounts
3. Click a row to see full account details in the panel below
4. Click **Investigate Source** to find what machine and process caused the lockout
5. Click **Unlock Selected** to resolve it — action is logged automatically

---

## LocalAdminAuditor

**Version:** v9.0 &nbsp;|&nbsp; **Requires:** Administrator, RSAT, PowerShell Remoting on targets

A WPF GUI tool that audits the local Administrators group across all Windows computers in your Active Directory domain. Find unexpected local admins, assess risk level, and remediate directly from the interface - without touching each machine individually.

### What it does

- Queries Active Directory for all servers and/or workstations in your domain
- Connects to each machine via PowerShell Remoting (WinRM) and reads the local Administrators group
- Scores each finding by risk level (HIGH / MEDIUM / LOW) based on account type and machine type
- Lets you remove accounts remotely with a right-click - with double confirmation
- Tracks findings over time with trend analysis and scan comparison

### Risk Levels

| Level | Color | Meaning |
|-------|-------|---------|
| HIGH | Red | Unknown local user account on a server |
| MEDIUM | Yellow | Domain user on a workstation, or service account on a server |
| LOW | Teal | Group membership - lower risk, still worth documenting |

### Key Features

- **Scan Targets:** Servers Only, Workstations Only, Both, or a specific list of computers
- **Authenticate:** Supply alternate domain admin credentials without restarting PowerShell
- **Exclusions:** Filter out expected accounts (Domain Admins, built-in Administrator, your secondary admin account prefix)
- **Remote Remediation:** Remove accounts from local Administrators group directly from the UI - requires double confirmation, all actions logged with operator identity
- **Export:** Executive Summary (HTML), CSV, PDF, clipboard
- **Trends:** Historical analysis across multiple scans - see if findings are increasing or decreasing
- **Compare:** Diff two scans to see what is NEW or REMOVED since last time
- **Resume:** If a scan is cancelled mid-way, state is saved and offered on next launch
- **Auto prereq setup:** Detects missing RSAT and WinRM on launch and offers to install/enable automatically

### Prerequisites

| Requirement | Notes |
|-------------|-------|
| PowerShell 5.1+ | Built into Windows 10/11 and Server 2016+ |
| RSAT - ActiveDirectory module | Auto-installs if missing (requires internet + admin) |
| WinRM enabled locally | Auto-enables if missing (requires admin) |
| WinRM enabled on targets | Required on every machine you want to scan - typically via GPO |
| Domain admin credentials | Or use the Authenticate button for alternate credentials |

### Quick Start

```powershell
# Run as Administrator
.\LocalAdminAuditor-GUI.ps1

# Or one-liner (no download needed)
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/LocalAdminAuditor-GUI.ps1 | iex
```

1. Click **Authenticate** and enter your domain admin credentials
2. Select **Servers Only** as the scan target (fastest, most important)
3. Click **Start Scan**
4. Review results sorted by risk level
5. Right-click any finding to remediate, exclude, or investigate

### Settings

- **Secondary Admin Account Prefix** - your secondary admin naming convention (e.g. `Fadmin`, `ladmin`, `secadmin`) - accounts matching this prefix are excluded from results
- **Domain Admins / Built-in Administrator** - excluded by default (expected accounts)
- **Custom Exclusion Patterns** - regex supported, one per line
- **Computer Exclusions** - honeypots, test systems, and dev machines to skip entirely
- **Output Directory** - where exports and scan history are saved
- **Auto-export** - automatically save CSV after every scan

> **Remediation Warning:** The Remove Admin Account right-click action connects to the remote machine via WinRM and removes the selected account from the local Administrators group. This takes effect immediately and cannot be easily undone. Always verify the account before removing. All removals are logged in the Activity Log with the operator's Windows identity and timestamp.

---

## NetworkSpeedChecker

**Version:** v4.4 &nbsp;|&nbsp; **Requires:** Administrator recommended, WinRM for remote testing

A GUI network speed testing tool for measuring throughput between your machine and remote servers or workstations on the network.

### What it does

- Tests upload and download speed to local and remote machines
- Supports batch testing across multiple targets simultaneously
- Parallel execution with configurable throttle
- Traceroute integration to diagnose network path issues
- Exports results to CSV for reporting

### Key Features

- **Local Test:** Measure speed on the local machine
- **Remote Test:** Test speed to a specific computer via WinRM
- **Batch Mode:** Import a list of computers and test all in parallel
- **Traceroute:** Built-in traceroute per target
- **CSV Export:** Save all results with timestamps

### Quick Start

```powershell
.\NetworkSpeedChecker_v4_4.ps1

# Or one-liner
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/NetworkSpeedChecker_v4_4.ps1 | iex
```

---

## ServerHealthCheck

**Version:** v6.0 &nbsp;|&nbsp; **Requires:** Administrator, WinRM on target servers

A GUI diagnostic tool that runs a 12-point health check on Windows Servers. Designed for quick triage during incidents or as part of a regular health review.

### What it checks

| Check | What it looks for |
|-------|-------------------|
| Disk Space | Drives below threshold (default 20% free) |
| CPU Usage | High sustained CPU load |
| Memory | Available RAM and page file usage |
| Event Log | Recent critical and error events (System + Application) |
| Services | Stopped services that should be running |
| Windows Updates | Pending updates and last install date |
| Uptime | Last reboot time and uptime duration |
| Network | Adapter status, IP config, connectivity |
| DNS | Resolution tests for key hostnames |
| Security | Firewall state, RDP status, audit policy |
| Shares | Open shares and connected sessions |
| Processes | Top CPU and memory consumers |

### Key Features

- Run against local machine or any remote server via WinRM
- Color-coded results (green / yellow / red) per check
- Export full diagnostic report to HTML or CSV
- Batch mode - run against multiple servers and compare

### Quick Start

```powershell
.\ServerHealthCheck_6.ps1

# Or one-liner
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/ServerHealthCheck_6.ps1 | iex
```

---

## SecurityChecker

**Version:** v2.6.9 &nbsp;|&nbsp; **Requires:** Administrator

A GUI subnet security scanner that checks Windows machines for common security misconfigurations. Useful for internal audits and hardening reviews.

### What it checks

| Check | Risk |
|-------|------|
| LLMNR enabled | Medium - enables credential capture attacks |
| mDNS enabled | Medium - enables credential capture attacks |
| NetBIOS enabled | Medium - enables credential capture attacks |
| SMBv1 enabled | High - legacy protocol with known critical vulnerabilities |
| SMB Signing disabled | Medium - enables man-in-the-middle attacks |
| WPAD enabled | Medium - enables proxy hijacking |
| IPv6 enabled (unmanaged) | Low - can enable rogue DHCPv6 attacks |
| RDP exposed | Informational - verify it is intentional and patched |

### Key Features

- Scan a single IP, hostname, or entire subnet (CIDR notation)
- Color-coded findings by severity
- Export results to CSV for audit reporting
- Remediation guidance per finding

### Quick Start

```powershell
# Run as Administrator
.\SecurityChecker-GUI-v2_6_9.ps1

# Or one-liner
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/SecurityChecker-GUI-v2_6_9.ps1 | iex
```

---

## Requirements Summary

| Requirement | ADLockoutManager | LocalAdminAuditor | NetworkSpeedChecker | ServerHealthCheck | SecurityChecker | AppHunter Pro |
|-------------|:----------------:|:-----------------:|:-------------------:|:-----------------:|:---------------:|:-------------:|
| PowerShell 5.1+ | Yes | Yes | Yes | Yes | Yes | Yes |
| Run as Administrator | Recommended | Yes | Recommended | Yes | Yes | Recommended |
| RSAT / AD Module | Yes (auto-installs) | Yes (auto-installs) | No | No | No | Domain scope only |
| WinRM on this machine | No | Yes (auto-enables) | Yes | Yes | No | Remote queries |
| WinRM on targets | No | Yes | Yes | Yes | No | Remote queries |
| Domain membership | Yes | Yes | No | No | No | Domain scope only |

---

## General Notes

- All tools are standalone `.ps1` files - no installation, no dependencies to download manually
- All tools are dark-themed GUI applications - they open a window, not a console
- Tested on Windows 10, Windows 11, Windows Server 2016, 2019, 2022
- PowerShell 7+ is not required but is supported

---

## License

MIT - free to use, modify, and distribute with attribution. See [LICENSE](LICENSE) for full text.
