#Requires -Version 5.1
<#
.SYNOPSIS
    Reports on and enforces the startup type of the Windows Update service (wuauserv)
    on the local computer and/or remote computers.

.DESCRIPTION
    For each target computer the script:
      1. Reads the current state of the Windows Update service (status, startup type,
         delayed-start flag) plus any Group Policy settings that block Windows Update.
      2. If the startup type differs from -StartupType, changes it (unless -ReportOnly
         or -WhatIf is used).
      3. Optionally starts the service (-StartService).
      4. Re-reads the service and returns one result object per computer.

    Local targets run in-process. Remote targets run through PowerShell Remoting
    (WinRM / Invoke-Command), in parallel.

    The startup type is changed with sc.exe, which works on Windows PowerShell 5.1
    and supports "Automatic (Delayed Start)". If sc.exe is refused, the script falls
    back to writing the service's registry values directly; that change takes effect
    after a reboot and is reported as such.

.PARAMETER ComputerName
    One or more computers to process. Defaults to the local computer. Accepts
    pipeline input (strings, or objects with a Name/ComputerName/DNSHostName
    property such as the output of Get-ADComputer).

.PARAMETER ComputerListPath
    Path to a text file with one computer name per line. Blank lines and lines
    starting with # are ignored.

.PARAMETER StartupType
    The startup type to enforce: Automatic (default), AutomaticDelayedStart,
    Manual, or Disabled.

.PARAMETER ReportOnly
    Only report the current state. Nothing is changed.

.PARAMETER StartService
    Also start the service if it is not running (ignored when -StartupType is Disabled).

.PARAMETER Credential
    Credential used for remote computers.

.PARAMETER CsvPath
    Optional path of a CSV file to export the results to.

.PARAMETER ThrottleLimit
    Maximum number of remote computers processed at the same time. Default 32.

.EXAMPLE
    .\Set-WindowsUpdateService.ps1 -ReportOnly
    Shows the Windows Update service status and startup type on this computer.

.EXAMPLE
    .\Set-WindowsUpdateService.ps1
    Sets Windows Update to Automatic on this computer (run as Administrator).

.EXAMPLE
    .\Set-WindowsUpdateService.ps1 -ComputerName PC01, PC02 -StartService -CsvPath .\wu-report.csv
    Sets Windows Update to Automatic on PC01 and PC02, starts it, and saves a CSV report.

.EXAMPLE
    .\Set-WindowsUpdateService.ps1 -ComputerListPath .\workstations.txt -WhatIf
    Shows what would change on every computer in workstations.txt without changing anything.

.EXAMPLE
    Get-ADComputer -Filter 'OperatingSystem -like "*Windows 1*"' |
        .\Set-WindowsUpdateService.ps1 -StartupType AutomaticDelayedStart
    Enforces Automatic (Delayed Start) on all Windows 10/11 computers in Active Directory.

.OUTPUTS
    PSCustomObject per computer with: ComputerName, Reachable, Status,
    StartupTypeBefore, StartupType, DesiredStartupType, Compliant, Action,
    RebootRequired, PolicyWarnings, Error.

.NOTES
    - Changing the startup type requires local Administrator rights on the target.
    - Remote computers need PowerShell Remoting enabled (Enable-PSRemoting / GPO).
    - On Windows 10/11 the Microsoft default for this service is
      "Manual (Trigger Start)"; Windows starts it on demand. Forcing Automatic is a
      valid choice, but the most common reason Windows Update is broken is the
      service being Disabled, or a Group Policy blocking it (reported in
      PolicyWarnings).
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'ByName')]
param(
    [Parameter(ParameterSetName = 'ByName', Position = 0,
               ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('CN', 'Name', 'DNSHostName', 'HostName')]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerName = $env:COMPUTERNAME,

    [Parameter(ParameterSetName = 'ByFile', Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ComputerListPath,

    [ValidateSet('Automatic', 'AutomaticDelayedStart', 'Manual', 'Disabled')]
    [string]$StartupType = 'Automatic',

    [switch]$ReportOnly,

    [switch]$StartService,

    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [string]$CsvPath,

    [ValidateRange(1, 256)]
    [int]$ThrottleLimit = 32
)

begin {
    $ServiceName = 'wuauserv'
    $targets = New-Object System.Collections.Generic.List[string]

    # Runs ON the target computer (locally or through Invoke-Command), so it must be
    # self-contained and compatible with Windows PowerShell 5.1.
    $worker = {
        param(
            [string]$ServiceName,
            [string]$DesiredStartupType,
            [bool]$Apply,
            [bool]$StartService
        )

        $ErrorActionPreference = 'Stop'
        $svcKey = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"

        function Get-WuState {
            $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='$ServiceName'"
            if (-not $svc) { throw "Service '$ServiceName' was not found." }

            $delayed = $false
            $regValues = Get-ItemProperty -LiteralPath $svcKey -ErrorAction SilentlyContinue
            if ($regValues -and $regValues.PSObject.Properties['DelayedAutostart']) {
                $delayed = [int]$regValues.DelayedAutostart -eq 1
            }

            $type = switch ($svc.StartMode) {
                'Auto'     { if ($delayed) { 'AutomaticDelayedStart' } else { 'Automatic' } }
                'Manual'   { 'Manual' }
                'Disabled' { 'Disabled' }
                default    { [string]$svc.StartMode }
            }

            [pscustomobject]@{
                Status      = [string]$svc.State
                StartupType = $type
            }
        }

        function Get-PolicyWarnings {
            $warnings = @()
            $wuPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
            $au = Get-ItemProperty -LiteralPath "$wuPolicy\AU" -ErrorAction SilentlyContinue
            $wu = Get-ItemProperty -LiteralPath $wuPolicy -ErrorAction SilentlyContinue

            if ($au -and $au.PSObject.Properties['NoAutoUpdate'] -and [int]$au.NoAutoUpdate -eq 1) {
                $warnings += 'Policy: automatic updates disabled (AU\NoAutoUpdate=1)'
            }
            if ($wu -and $wu.PSObject.Properties['DisableWindowsUpdateAccess'] -and [int]$wu.DisableWindowsUpdateAccess -eq 1) {
                $warnings += 'Policy: Windows Update access removed (DisableWindowsUpdateAccess=1)'
            }
            if ($wu -and $wu.PSObject.Properties['DoNotConnectToWindowsUpdateInternetLocations'] -and [int]$wu.DoNotConnectToWindowsUpdateInternetLocations -eq 1) {
                $warnings += 'Policy: blocked from Windows Update internet locations'
            }
            $warnings -join '; '
        }

        $result = [ordered]@{
            ComputerName       = $env:COMPUTERNAME
            Reachable          = $true
            Status             = $null
            StartupTypeBefore  = $null
            StartupType        = $null
            DesiredStartupType = $DesiredStartupType
            Compliant          = $false
            Action             = 'None'
            RebootRequired     = $false
            PolicyWarnings     = $null
            Error              = $null
        }

        try {
            $before = Get-WuState
            $result.StartupTypeBefore = $before.StartupType
            $result.PolicyWarnings = Get-PolicyWarnings
            $actions = @()

            if ($before.StartupType -ne $DesiredStartupType) {
                if (-not $Apply) {
                    $actions += "Would change startup type to $DesiredStartupType"
                }
                else {
                    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
                        [Security.Principal.WindowsBuiltInRole]::Administrator)
                    if (-not $isAdmin) {
                        throw 'Administrator rights are required to change the startup type (run PowerShell as Administrator).'
                    }

                    $scStart = @{
                        Automatic             = 'auto'
                        AutomaticDelayedStart = 'delayed-auto'
                        Manual                = 'demand'
                        Disabled              = 'disabled'
                    }[$DesiredStartupType]

                    $scOutput = & "$env:SystemRoot\System32\sc.exe" config $ServiceName start= $scStart 2>&1
                    $scExit = $LASTEXITCODE

                    if ($scExit -eq 0) {
                        # sc.exe "auto" does not always clear an existing delayed-start flag.
                        if ($DesiredStartupType -eq 'Automatic' -and (Get-WuState).StartupType -ne 'Automatic') {
                            Set-ItemProperty -LiteralPath $svcKey -Name DelayedAutostart -Value 0 -Type DWord
                        }
                        $actions += "Changed startup type $($before.StartupType) -> $DesiredStartupType"
                    }
                    else {
                        # Fallback: write the service configuration directly. The Service
                        # Control Manager only reads this at boot, so a reboot is needed.
                        $startValue = @{ Automatic = 2; AutomaticDelayedStart = 2; Manual = 3; Disabled = 4 }[$DesiredStartupType]
                        $delayedValue = [int]($DesiredStartupType -eq 'AutomaticDelayedStart')
                        try {
                            Set-ItemProperty -LiteralPath $svcKey -Name Start -Value $startValue -Type DWord
                            Set-ItemProperty -LiteralPath $svcKey -Name DelayedAutostart -Value $delayedValue -Type DWord
                        }
                        catch {
                            throw "sc.exe failed (exit $scExit): $(($scOutput | Out-String).Trim()) Registry fallback also failed: $($_.Exception.Message)"
                        }
                        $result.RebootRequired = $true
                        $actions += "Set startup type $DesiredStartupType in registry (sc.exe exit $scExit); takes effect after reboot"
                    }
                }
            }

            if ($StartService -and $DesiredStartupType -ne 'Disabled') {
                $current = Get-WuState
                if ($current.Status -ne 'Running') {
                    if (-not $Apply) {
                        $actions += 'Would start service'
                    }
                    elseif ($current.StartupType -eq 'Disabled') {
                        $actions += 'Cannot start service until reboot (still Disabled)'
                    }
                    else {
                        Start-Service -Name $ServiceName
                        $actions += 'Started service'
                    }
                }
            }

            $after = Get-WuState
            $result.Status = $after.Status
            $result.StartupType = $after.StartupType
            $result.Compliant = ($after.StartupType -eq $DesiredStartupType)
            if ($result.RebootRequired -and -not $result.Compliant) {
                $result.StartupType = "$($after.StartupType) (pending reboot: $DesiredStartupType)"
            }
            if ($DesiredStartupType -eq 'Disabled' -and $after.Status -eq 'Running') {
                $actions += 'Service is still running until stopped or rebooted'
            }
            if ($actions) { $result.Action = $actions -join '; ' }
        }
        catch {
            $result.Error = $_.Exception.Message
            if (-not $result.Status) {
                try {
                    $state = Get-WuState
                    $result.Status = $state.Status
                    $result.StartupType = $state.StartupType
                    $result.Compliant = ($state.StartupType -eq $DesiredStartupType)
                }
                catch { }
            }
        }

        [pscustomobject]$result
    }

    function Test-IsLocalComputer([string]$Name) {
        $short = ($Name -split '\.')[0]
        $Name -in @('.', 'localhost', '127.0.0.1', '::1') -or $short -eq $env:COMPUTERNAME
    }
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'ByFile') { return }
    foreach ($name in $ComputerName) {
        $trimmed = $name.Trim()
        if ($trimmed) { $targets.Add($trimmed) }
    }
}

end {
    if ($PSCmdlet.ParameterSetName -eq 'ByFile') {
        Get-Content -LiteralPath $ComputerListPath |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') } |
            ForEach-Object { $targets.Add($_) }
    }

    $uniqueTargets = $targets | Sort-Object -Unique
    if (-not $uniqueTargets) {
        Write-Warning 'No computers to process.'
        return
    }

    # Decide per computer whether changes are allowed (-ReportOnly, -WhatIf, -Confirm).
    $applyTargets = @()
    $reportTargets = @()
    foreach ($target in $uniqueTargets) {
        if (-not $ReportOnly -and
            $PSCmdlet.ShouldProcess($target, "Set Windows Update service startup type to $StartupType")) {
            $applyTargets += $target
        }
        else {
            $reportTargets += $target
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($group in @(
            @{ Apply = $true;  Computers = $applyTargets },
            @{ Apply = $false; Computers = $reportTargets })) {

        if (-not $group.Computers) { continue }
        $workerArgs = @($ServiceName, $StartupType, $group.Apply, [bool]$StartService)

        $local = @($group.Computers | Where-Object { Test-IsLocalComputer $_ })
        $remote = @($group.Computers | Where-Object { -not (Test-IsLocalComputer $_) })

        if ($local) {
            Write-Verbose "Processing local computer ($env:COMPUTERNAME)"
            $results.Add((& $worker @workerArgs))
        }

        if ($remote) {
            Write-Verbose "Processing $($remote.Count) remote computer(s): $($remote -join ', ')"
            $invokeParams = @{
                ComputerName  = $remote
                ScriptBlock   = $worker
                ArgumentList  = $workerArgs
                ThrottleLimit = $ThrottleLimit
                ErrorAction   = 'SilentlyContinue'
                ErrorVariable = 'remoteErrors'
            }
            if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
                $invokeParams.Credential = $Credential
            }

            $remoteErrors = @()
            try {
                $remoteResults = @(Invoke-Command @invokeParams)
            }
            catch {
                $remoteResults = @()
                $remoteErrors = @($_)
            }
            $answered = @{}
            foreach ($r in $remoteResults) {
                $answered[$r.PSComputerName.ToLowerInvariant()] = $true
                $results.Add(($r | Select-Object -Property * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName))
            }

            # Any computer that returned nothing could not be reached; report why.
            foreach ($computer in $remote) {
                if ($answered.ContainsKey($computer.ToLowerInvariant())) { continue }
                $err = $remoteErrors | Where-Object {
                    ($_.TargetObject -and "$($_.TargetObject)" -eq $computer) -or
                    ($_.OriginInfo -and $_.OriginInfo.PSComputerName -eq $computer)
                } | Select-Object -First 1
                if (-not $err -and @($remoteErrors).Count -eq 1 -and $remote.Count -eq 1) { $err = $remoteErrors[0] }
                $message = if ($err) { $err.Exception.Message } else { 'No response from computer (PowerShell Remoting unavailable?)' }

                $results.Add([pscustomobject][ordered]@{
                    ComputerName       = $computer
                    Reachable          = $false
                    Status             = $null
                    StartupTypeBefore  = $null
                    StartupType        = $null
                    DesiredStartupType = $StartupType
                    Compliant          = $false
                    Action             = 'None'
                    RebootRequired     = $false
                    PolicyWarnings     = $null
                    Error              = $message.Trim()
                })
            }
        }
    }

    $sorted = $results | Sort-Object ComputerName

    if ($CsvPath) {
        $sorted | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Verbose "Results exported to $CsvPath"
    }

    $sorted

    $total = @($sorted).Count
    $compliant = @($sorted | Where-Object Compliant).Count
    $unreachable = @($sorted | Where-Object { -not $_.Reachable }).Count
    $failed = @($sorted | Where-Object { $_.Reachable -and $_.Error }).Count
    Write-Host ("Windows Update service ({0}): {1}/{2} compliant, {3} unreachable, {4} with errors." -f
        $StartupType, $compliant, $total, $unreachable, $failed) -ForegroundColor Cyan
}
