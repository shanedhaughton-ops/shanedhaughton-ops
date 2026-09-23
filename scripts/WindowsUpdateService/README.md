# Set-WindowsUpdateService.ps1

Checks the status and startup type of the **Windows Update** service (`wuauserv`) and sets its startup type. It works on the local computer and on remote computers through PowerShell Remoting.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- To make changes, run it **as Administrator** (on remote targets, the account must be a local admin)
- To reach remote computers, PowerShell Remoting (WinRM) must be enabled on them (`Enable-PSRemoting -Force`, or through GPO)

## Usage

```powershell
# Report only (no changes) on this computer
.\Set-WindowsUpdateService.ps1 -ReportOnly

# Set to Automatic on this computer and start the service
.\Set-WindowsUpdateService.ps1 -StartService

# Remote computers, with a CSV report
.\Set-WindowsUpdateService.ps1 -ComputerName PC01, PC02 -StartService -CsvPath .\wu-report.csv

# A list of computers from a file (one per line); preview with -WhatIf first
.\Set-WindowsUpdateService.ps1 -ComputerListPath .\workstations.txt -WhatIf
.\Set-WindowsUpdateService.ps1 -ComputerListPath .\workstations.txt -StartService -Credential (Get-Credential)

# Pipe from Active Directory, using Automatic (Delayed Start)
Get-ADComputer -Filter 'OperatingSystem -like "*Windows 1*"' |
    .\Set-WindowsUpdateService.ps1 -StartupType AutomaticDelayedStart
```

If the script is blocked by execution policy, run it with `powershell.exe -ExecutionPolicy Bypass -File .\Set-WindowsUpdateService.ps1 ...`.

## Output

The script returns one object per computer:

| Property | Meaning |
|---|---|
| `Reachable` | `False` if the remote session could not be opened (see `Error`) |
| `Status` | Running / Stopped |
| `StartupTypeBefore` / `StartupType` | Startup type before and after the run |
| `Compliant` | The startup type matches `-StartupType` |
| `Action` | What was changed, or what would have changed |
| `RebootRequired` | `sc.exe` was refused, so the registry was edited instead; the change takes effect after a reboot |
| `PolicyWarnings` | Group Policy settings that block Windows Update even when the service is running |

## Notes

- The Windows 10/11 default for this service is **Manual (Trigger Start)**. Setting it to **Automatic** is supported. If a GPO or management tool keeps setting it to Disabled, it will undo this script's change; check `PolicyWarnings` and your GPOs.
- To enforce this across a whole domain without running the script again, you can also use a GPO: *Computer Configuration → Policies → Windows Settings → Security Settings → System Services → Windows Update → Automatic*.
