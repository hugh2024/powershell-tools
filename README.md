# powershell-tools
A collection of PowerShell GUI utilities for Windows IT and security operations.
> **Author:** Hugh Gozlou  
> **License:** MIT
---
## Tools
| Tool | Description |
|------|-------------|
| [NetworkSpeedChecker](./NetworkSpeedChecker_v4_4.ps1) | GUI network speed tester — local, remote, batch, traceroute, CSV export |
| [ServerHealthCheck](./ServerHealthCheck_6.ps1) | GUI 12-point Windows Server health diagnostic — events, disk, CPU, security |
| [SecurityChecker](./SecurityChecker-GUI-v2_6_9.ps1) | GUI subnet security scanner — LLMNR, mDNS, NetBIOS, SMBv1, SMB Signing, WPAD, IPv6, RDP |
---
## Quick Run (no install)
```powershell
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/NetworkSpeedChecker_v4_4.ps1 | iex
irm https://raw.githubusercontent.com/hugh2024/powershell-tools/main/SecurityChecker-GUI-v2_6_9.ps1 | iex
```
## Requirements
- Windows 10 / 11 or Windows Server 2016+
- PowerShell 5.1+
- Run as Administrator (SecurityChecker)
---
## License
MIT — free to use and modify with attribution.
