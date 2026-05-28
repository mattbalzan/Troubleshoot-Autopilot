<#
.SYNOPSIS
    Post-build Autopilot X-Ray snapshot - captures first-event timings for an Autopilot deployment.

.DESCRIPTION
    Runs immediately after Autopilot/ESP completes to capture a timestamped timeline of:
    Autopilot profile JSON download, device registration, MDM enrollment, policies applied,
    apps installed (Win32 & UWP), certificates added, Intune Windows Agent (sidecar) install,
    platform scripts, remediation scripts and Office 365 / M365 Apps installation.

    Data is collected from registry last-write timestamps (via getRegTime), Windows event logs
    and the Intune Management Extension (IME) logs (which are also used to map app GUIDs to
    friendly app names). The unified timeline is exported to a local log, CSV and JSON under
    C:\ProgramData\AutopilotXRay.

.NOTES
    +------------+---------+---------+---------------------------------------------------------------------+
    | Date       | Author  | Version | Changes                                                             |
    |------------+---------+---------+---------------------------------------------------------------------|
    | 2025-02-04 | mattGPT | 1.0     | Original.                                                           |
    | 2025-02-05 | mattGPT | 1.1     | Added parse app names using GUID.                                   |
    | 2025-02-06 | mattGPT | 1.2     | Joined results: Timestamp (local) - Name - Status - Property.       |
    | 2025-02-07 | mattGPT | 1.3     | Added First Events, Intune Win Agent, Office & Autopilot tables.    |
    | 2026-05-28 | mattGPT | 2.0     | Added platform & remediation script timings via IME registry/logs.  |
    | 2026-05-28 | mattGPT | 2.0     | Added device registration & MDM enrollment event log capture.       |
    | 2026-05-28 | mattGPT | 2.0     | Added Office 365 install timeline via OfficeCSP.                    |
    | 2026-05-28 | mattGPT | 2.0     | Added IME log parsing to map App GUIDs to friendly names.           |
    | 2026-05-28 | mattGPT | 2.0     | Added unified export to .log, .csv & .json in ProgramData.          |
    | 2026-05-28 | mattGPT | 2.0     | Added Autopilot snapshot duration (first event to last event).      |
    +------------+---------+---------+---------------------------------------------------------------------+

#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "SilentlyContinue"

function getRegTime($regPath) {
    $signature = '[DllImport("advapi32.dll", CharSet = CharSet.Auto)]
public static extern Int32 RegQueryInfoKey(
Microsoft.Win32.SafeHandles.SafeRegistryHandle hKey,
StringBuilder lpClass,
Int32 lpCls, Int32 spare, Int32 subkeys,
Int32 skLen, Int32 mcLen, Int32 values,
Int32 vNLen, Int32 mvLen, Int32 secDesc,
out System.Runtime.InteropServices.ComTypes.FILETIME lpftLastWriteTime
);'

    if (-not ([System.Management.Automation.PSTypeName]'RegQueryInfoKey.GetRegData').Type) {
        $null = Add-Type $signature -Name GetRegData -Namespace RegQueryInfoKey -Using System.Text -PassThru
    }

    $reg = Get-Item -Path $regPath -Force
    if (-not $reg -or -not $reg.Handle) {
        return $null
    }

    $time = New-Object System.Runtime.InteropServices.ComTypes.FILETIME
    $result = [RegQueryInfoKey.GetRegData]::RegQueryInfoKey($reg.Handle, $null, 0, 0, 0, 0, 0, 0, 0, 0, 0, [ref]$time)
    if ($result -ne 0) {
        return $null
    }

    $low = [uint32]0 -bor $time.dwLowDateTime
    $high = [uint32]0 -bor $time.dwHighDateTime
    $timeValue = ([int64]$high -shl 32) -bor $low
    return [datetime]::FromFileTime($timeValue)
}

function New-TimelineRecord {
    param(
        [datetime]$Timestamp,
        [string]$Name,
        [string]$Status,
        [string]$Property,
        [string]$Source,
        [string]$Details
    )

    if (-not $Timestamp) {
        return $null
    }

    [PSCustomObject]@{
        Timestamp = $Timestamp
        Name = $Name
        Status = $Status
        Property = $Property
        Source = $Source
        Details = $Details
    }
}

function New-LocalLogContext {
    $root = Join-Path $env:ProgramData "AutopilotXRay"
    $null = New-Item -Path $root -ItemType Directory -Force

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $base = "Autopilot_XRay_$($env:COMPUTERNAME)_$stamp"

    [PSCustomObject]@{
        Root = $root
        TextLog = Join-Path $root "$base.log"
        CsvLog = Join-Path $root "$base.csv"
        JsonLog = Join-Path $root "$base.json"
    }
}

function Write-LocalLogLine {
    param(
        [string]$Path,
        [string]$Message
    )
    "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") $Message" | Out-File -FilePath $Path -Append -Encoding utf8
}

# --[ Console colour map per Property/category ]
$Script:CategoryColors = @{
    "Autopilot Profile"   = "Cyan"
    "Enrollment"          = "Yellow"
    "Policy"              = "Green"
    "IME"                 = "Magenta"
    "ESP"                 = "DarkCyan"
    "Win32App"            = "White"
    "Device UWP"          = "Gray"
    "User UWP"            = "Gray"
    "Device Certificate"  = "DarkYellow"
    "User Certificate"    = "DarkYellow"
    "Platform Script"     = "Blue"
    "Remediation Script" = "DarkMagenta"
    "M365 Apps"           = "DarkGreen"
}

function Get-CategoryColor {
    param([string]$Property)

    if ($Property -and $Script:CategoryColors.ContainsKey($Property)) {
        return $Script:CategoryColors[$Property]
    }
    return "Gray"
}

function Get-StatusColor {
    param([string]$Status)

    switch -Regex ($Status) {
        "(?i)fail|error|connectivity issue"        { return "Red" }
        "(?i)success|completed|installed|succeeded|downloaded|processed|applied" { return "Green" }
        "(?i)in progress|processing|waiting|pending|started|retry" { return "Yellow" }
        default                                     { return "Gray" }
    }
}

function Write-TimelineConsole {
    param([array]$Timeline)

    $col = [PSCustomObject]@{
        Timestamp = 19
        Name      = [Math]::Min(40, [Math]::Max(20, ($Timeline.Name      | Where-Object { $_ } | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum))
        Status    = [Math]::Min(28, [Math]::Max(10, ($Timeline.Status    | Where-Object { $_ } | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum))
        Property  = [Math]::Min(22, [Math]::Max(10, ($Timeline.Property  | Where-Object { $_ } | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum))
        Source    = 32
    }

    $headerFmt = "{0,-$($col.Timestamp)}  {1,-$($col.Name)}  {2,-$($col.Status)}  {3,-$($col.Property)}  {4,-$($col.Source)}"
    $separator = ("-" * ($col.Timestamp + $col.Name + $col.Status + $col.Property + $col.Source + 8))

    Write-Host ""
    Write-Host ($headerFmt -f "Timestamp", "Name", "Status", "Property", "Source") -ForegroundColor White
    Write-Host $separator -ForegroundColor DarkGray

    foreach ($row in $Timeline) {
        $catColor    = Get-CategoryColor -Property $row.Property
        $statusColor = Get-StatusColor   -Status   $row.Status
        $name        = if ($row.Name)   { $row.Name }   else { "" }
        $status      = if ($row.Status) { $row.Status } else { "" }
        $property    = if ($row.Property) { $row.Property } else { "" }
        $source      = if ($row.Source) { $row.Source } else { "" }
        if ($name.Length     -gt $col.Name)     { $name     = $name.Substring(0, $col.Name - 1) + "…" }
        if ($status.Length   -gt $col.Status)   { $status   = $status.Substring(0, $col.Status - 1) + "…" }
        if ($property.Length -gt $col.Property) { $property = $property.Substring(0, $col.Property - 1) + "…" }
        if ($source.Length   -gt $col.Source)   { $source   = $source.Substring(0, $col.Source - 1) + "…" }

        Write-Host ($row.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")).PadRight($col.Timestamp) -ForegroundColor DarkGray -NoNewline
        Write-Host "  " -NoNewline
        Write-Host $name.PadRight($col.Name)         -ForegroundColor $catColor    -NoNewline
        Write-Host "  " -NoNewline
        Write-Host $status.PadRight($col.Status)     -ForegroundColor $statusColor -NoNewline
        Write-Host "  " -NoNewline
        Write-Host $property.PadRight($col.Property) -ForegroundColor $catColor    -NoNewline
        Write-Host "  " -NoNewline
        Write-Host $source.PadRight($col.Source)     -ForegroundColor DarkGray
    }

    Write-Host $separator -ForegroundColor DarkGray
}

function Write-CategoryLegend {
    Write-Host ""
    Write-Host "Category legend:" -ForegroundColor White
    foreach ($key in ($Script:CategoryColors.Keys | Sort-Object)) {
        Write-Host (" * {0}" -f $key) -ForegroundColor $Script:CategoryColors[$key]
    }
    Write-Host ""
}

function Convert-IMELogTimestamp {
    param(
        [string]$DatePart,
        [string]$TimePart
    )

    if (-not $DatePart -or -not $TimePart) {
        return $null
    }

    $cleanTime = ($TimePart -split "\+")[0]
    $value = "$DatePart $cleanTime"
    $formats = @(
        "M-d-yyyy H:mm:ss.fffffff",
        "M-d-yyyy HH:mm:ss.fffffff",
        "M-d-yyyy H:mm:ss",
        "M-d-yyyy HH:mm:ss",
        "MM-dd-yyyy HH:mm:ss.fffffff",
        "MM-dd-yyyy HH:mm:ss"
    )

    foreach ($format in $formats) {
        $parsed = $null
        if ([datetime]::TryParseExact($value, $format, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
            return $parsed
        }
    }

    return (Get-Date $value)
}

function Get-IMELogEntries {
    $entries = @()
    $logFolder = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
    if (-not (Test-Path $logFolder)) {
        return $entries
    }

    $logFiles = Get-ChildItem -Path $logFolder -File -Filter "*.log"
    $pattern = '<!\[LOG\[(?<msg>.*?)\]LOG\]!><time="(?<time>[^"]+)" date="(?<date>[^"]+)" component="(?<component>[^"]*)" context="(?<context>[^"]*)" type="(?<type>[^"]*)" thread="(?<thread>[^"]*)" file="(?<file>[^"]*)">'

    foreach ($file in $logFiles) {
        foreach ($line in (Get-Content -Path $file.FullName)) {
            $match = [regex]::Match($line, $pattern)
            if (-not $match.Success) {
                continue
            }

            $entries += [PSCustomObject]@{
                Timestamp = Convert-IMELogTimestamp -DatePart $match.Groups["date"].Value -TimePart $match.Groups["time"].Value
                Message = $match.Groups["msg"].Value
                Component = $match.Groups["component"].Value
                File = $file.Name
            }
        }
    }

    return ($entries | Where-Object { $_.Timestamp } | Sort-Object Timestamp)
}

function Get-AppNameMapFromIME {
    param([array]$IMEEntries)

    $map = @{}
    foreach ($entry in $IMEEntries) {
        $m = [regex]::Match($entry.Message, '"Id":"(?<id>[0-9a-fA-F-]{36})","Name":"(?<name>.*?)"')
        if ($m.Success -and -not $map.ContainsKey($m.Groups["id"].Value)) {
            $map[$m.Groups["id"].Value] = $m.Groups["name"].Value
        }
    }
    return $map
}

function Resolve-AppName {
    param(
        [string]$AppId,
        [hashtable]$AppMap
    )

    if ($AppMap.ContainsKey($AppId)) {
        return $AppMap[$AppId]
    }

    return "Unknown App ($AppId)"
}

function Get-AutopilotProfileEvent {
    $jsonFile = "$($env:WINDIR)\ServiceState\wmansvc\AutopilotDDSZTDFile.json"
    if (-not (Test-Path $jsonFile)) {
        return $null
    }

    $json = Get-Content -Path $jsonFile -Raw | ConvertFrom-Json
    return (New-TimelineRecord -Timestamp ([datetime]$json.PolicyDownloadDate) -Name $json.DeploymentProfileName -Status "Downloaded" -Property "Autopilot Profile" -Source "JSON" -Details $jsonFile)
}

function Get-DeviceRegistrationEvents {
    $results = @()
    $eventSources = @(
        @{ Log = "Microsoft-Windows-User Device Registration/Admin"; Ids = @(100, 101, 104, 106, 107, 109, 111, 306) },
        @{ Log = "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin"; Ids = @(70, 71, 72, 75) }
    )

    $messages = @{
        70 = @{ Name = "MDM Enrollment"; Status = "Started" }
        71 = @{ Name = "MDM Enrollment"; Status = "In Progress" }
        72 = @{ Name = "MDM Enrollment"; Status = "Succeeded" }
        75 = @{ Name = "MDM Enrollment"; Status = "Failed" }
        100 = @{ Name = "Device Registration"; Status = "Connectivity Issue" }
        101 = @{ Name = "Device Registration"; Status = "Service Connection Point successful" }
        104 = @{ Name = "Device Registration"; Status = "Join task started" }
        106 = @{ Name = "Device Registration"; Status = "Join in progress" }
        107 = @{ Name = "Device Registration"; Status = "ODJ blob applied" }
        109 = @{ Name = "Device Registration"; Status = "Join completed" }
        111 = @{ Name = "Device Registration"; Status = "Waiting for ODJ blob" }
        306 = @{ Name = "Device Registration"; Status = "Hybrid AADJ registration succeeded" }
    }

    foreach ($source in $eventSources) {
        $events = Get-WinEvent -LogName $source.Log -Oldest | Where-Object { $_.Id -in $source.Ids }
        foreach ($evt in $events) {
            $details = if ($evt.Message.Length -gt 180) { $evt.Message.Substring(0, 180) } else { $evt.Message }
            $results += New-TimelineRecord -Timestamp $evt.TimeCreated -Name $messages[$evt.Id].Name -Status $messages[$evt.Id].Status -Property "Enrollment" -Source $source.Log -Details $details
        }
    }

    return $results
}

function Get-PolicyEvents {
    $results = @()
    $policyRoot = "HKLM:\SOFTWARE\Microsoft\PolicyManager\Providers"

    if (-not (Test-Path $policyRoot)) {
        return $results
    }

    $providerRoots = Get-ChildItem -Path $policyRoot
    foreach ($provider in $providerRoots) {
        $devicePath = Join-Path $provider.PSPath "default\Device"
        if (-not (Test-Path $devicePath)) {
            continue
        }

        foreach ($key in (Get-ChildItem -Path $devicePath)) {
            $regPath = $key.Name -replace "HKEY_LOCAL_MACHINE", "HKLM:"
            $timestamp = getRegTime -regPath $regPath
            $results += New-TimelineRecord -Timestamp $timestamp -Name $key.PSChildName -Status "Processed" -Property "Policy" -Source "Registry" -Details $provider.PSChildName
        }
    }

    return $results
}

function Get-SidecarInstallEvent {
    $path = "HKLM:\SOFTWARE\Microsoft\IntuneWindowsAgent"
    if (-not (Test-Path $path)) {
        return $null
    }

    $timestamp = getRegTime -regPath $path
    return (New-TimelineRecord -Timestamp $timestamp -Name "Intune Windows Agent" -Status "Installed" -Property "IME" -Source "Registry" -Details $path)
}

function Get-ESPStatusEvents {
    param([hashtable]$AppMap)

    $results = @()
    $espRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics\Expected*",
        "HKLM:\SOFTWARE\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics\S-*\Expected*"
    )

    $espState = @{ "0" = "Not Processed"; "1" = "Processed"; "60" = "Failed"; "70" = "Completed" }

    foreach ($root in $espRoots) {
        foreach ($key in (Get-ChildItem -Path $root -Recurse)) {
            $stamp = $null
            if (-not [datetime]::TryParse($key.PSChildName, [ref]$stamp)) {
                continue
            }

            $props = Get-ItemProperty -Path $key.PSPath
            foreach ($prop in $props.PSObject.Properties) {
                if ($prop.Name -match "^PS") {
                    continue
                }

                $propertyGroup = "ESP"
                $name = $prop.Name
                if ($prop.Name -like "./User/Vendor/MSFT/ClientCertificateInstall/*") {
                    $name = [URI]::UnescapeDataString(($prop.Name.Split("/"))[6])
                    $propertyGroup = "User Certificate"
                } elseif ($prop.Name -like "./Device/Vendor/MSFT/ClientCertificateInstall/*") {
                    $name = [URI]::UnescapeDataString(($prop.Name.Split("/"))[6])
                    $propertyGroup = "Device Certificate"
                } elseif ($prop.Name -like "./Vendor/MSFT/Office/Installation*") {
                    $name = "Office Installation"
                    $propertyGroup = "M365 Apps"
                } elseif ($prop.Name -like "./Device/Vendor/MSFT/EnterpriseModernAppManagement/AppManagement/*") {
                    $raw = [URI]::UnescapeDataString(($prop.Name.Split("/"))[7])
                    if ($raw -match "[0-9a-fA-F-]{36}") {
                        $name = Resolve-AppName -AppId $matches[0] -AppMap $AppMap
                    } else {
                        $name = $raw
                    }
                    $propertyGroup = "Device UWP"
                } elseif ($prop.Name -like "./User/Vendor/MSFT/EnterpriseModernAppManagement/AppManagement/*") {
                    $name = [URI]::UnescapeDataString(($prop.Name.Split("/"))[7])
                    $propertyGroup = "User UWP"
                }

                $status = if ($espState.ContainsKey([string]$prop.Value)) { $espState[[string]$prop.Value] } else { [string]$prop.Value }
                $results += New-TimelineRecord -Timestamp $stamp -Name $name -Status $status -Property $propertyGroup -Source "ESP Registry" -Details $prop.Name
            }
        }
    }

    return $results
}

function Get-Win32AppEvents {
    param([hashtable]$AppMap)

    $results = @()
    $sidecarPath = "HKLM:\SOFTWARE\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics\Sidecar"
    $statusMap = @{ "1" = "Not Installed"; "2" = "In Progress"; "3" = "Completed"; "4" = "Error" }

    if (-not (Test-Path $sidecarPath)) {
        return $results
    }

    foreach ($key in (Get-ChildItem -Path $sidecarPath -Recurse | Where-Object { $_.PSChildName -ne "LastLoggedState" })) {
        $stamp = $null
        if (-not [datetime]::TryParse($key.PSChildName, [ref]$stamp)) {
            continue
        }

        $props = Get-ItemProperty -Path $key.PSPath
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Name -notmatch "Win32App_" -or $prop.Name -match "^PS") {
                continue
            }

            $id = $prop.Name -replace "^.*Win32App_", "" -replace "_\d+$", ""
            $appName = Resolve-AppName -AppId $id -AppMap $AppMap
            $status = if ($statusMap.ContainsKey([string]$prop.Value)) { $statusMap[[string]$prop.Value] } else { [string]$prop.Value }

            $results += New-TimelineRecord -Timestamp $stamp -Name $appName -Status $status -Property "Win32App" -Source "ESP Sidecar" -Details $id
        }
    }

    return $results
}

function Get-PlatformScriptEvents {
    $results = @()
    $root = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Policies"

    if (-not (Test-Path $root)) {
        return $results
    }

    $keys = Get-ChildItem -Path $root -Recurse | Where-Object { $_.PSChildName -match "^[0-9a-fA-F-]{36}$" }
    foreach ($key in $keys) {
        $timestamp = getRegTime -regPath ($key.Name -replace "HKEY_LOCAL_MACHINE", "HKLM:")
        $results += New-TimelineRecord -Timestamp $timestamp -Name $key.PSChildName -Status "Processed" -Property "Platform Script" -Source "IME Registry" -Details $key.Name
    }

    return $results
}

function Get-RemediationScriptEvents {
    $results = @()
    $root = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\SideCarPolicies\Scripts\Reports"

    if (-not (Test-Path $root)) {
        return $results
    }

    $resultKeys = Get-ChildItem -Path $root -Recurse | Where-Object { $_.PSChildName -eq "Result" }
    foreach ($key in $resultKeys) {
        $timestamp = getRegTime -regPath ($key.Name -replace "HKEY_LOCAL_MACHINE", "HKLM:")
        $parent = Split-Path -Path $key.PSPath -Parent
        $scriptNode = Split-Path -Path $parent -Leaf
        $assignmentNode = Split-Path -Path (Split-Path -Path $parent -Parent) -Leaf

        $statusText = "Completed"
        $values = Get-ItemProperty -Path $key.PSPath
        if ($values.PSObject.Properties.Name -contains "ErrorCode" -and [int]$values.ErrorCode -ne 0) {
            $statusText = "Failed"
        }

        $results += New-TimelineRecord -Timestamp $timestamp -Name "$assignmentNode / $scriptNode" -Status $statusText -Property "Remediation Script" -Source "IME Registry" -Details $key.Name
    }

    return $results
}

function Get-IMEProcessingEvents {
    param(
        [array]$IMEEntries,
        [hashtable]$AppMap
    )

    $results = @()
    foreach ($entry in $IMEEntries) {
        $msg = $entry.Message

        if ($msg -match '(?i)processing policy with id\s*=\s*(?<id>[0-9a-fA-F-]{36})') {
            $results += New-TimelineRecord -Timestamp $entry.Timestamp -Name $matches["id"] -Status "Processing" -Property "Platform Script" -Source $entry.File -Details $msg
            continue
        }

        if ($msg -match '(?i)remediation' -and $msg -match '(?<id>[0-9a-fA-F-]{36})') {
            $status = if ($msg -match '(?i)fail|error') { "Failed" } elseif ($msg -match '(?i)success|completed') { "Completed" } else { "Processing" }
            $results += New-TimelineRecord -Timestamp $entry.Timestamp -Name $matches["id"] -Status $status -Property "Remediation Script" -Source $entry.File -Details $msg
            continue
        }

        if ($msg -match '(?i)win32app|app installation|enforcement') {
            $guidMatch = [regex]::Match($msg, '(?<id>[0-9a-fA-F-]{36})')
            if ($guidMatch.Success) {
                $appName = Resolve-AppName -AppId $guidMatch.Groups["id"].Value -AppMap $AppMap
                $status = if ($msg -match '(?i)fail|error') { "Failed" } elseif ($msg -match '(?i)success|completed|installed') { "Completed" } else { "Processing" }
                $results += New-TimelineRecord -Timestamp $entry.Timestamp -Name $appName -Status $status -Property "Win32App" -Source $entry.File -Details $msg
            }
        }
    }

    return $results
}

function Get-OfficeInstallEvents {
    $results = @()
    $officeRoot = "HKLM:\SOFTWARE\Microsoft\OfficeCSP"
    if (-not (Test-Path $officeRoot)) {
        return $results
    }

    $statusMap = @{
        "0" = "None"
        "10" = "Initialized"
        "20" = "Download In Progress"
        "25" = "Pending Download Retry"
        "30" = "Download Failed"
        "40" = "Download Completed"
        "48" = "Pending User Session"
        "50" = "Enforcement In Progress"
        "55" = "Pending Enforcement Retry"
        "60" = "Enforcement Failed"
        "70" = "Success"
    }

    foreach ($key in (Get-ChildItem -Path $officeRoot)) {
        $stamp = getRegTime -regPath ($key.Name -replace "HKEY_LOCAL_MACHINE", "HKLM:")
        $state = (Get-ItemProperty -Path $key.PSPath -Name FinalStatus).FinalStatus
        $status = if ($statusMap.ContainsKey([string]$state)) { $statusMap[[string]$state] } else { "Unknown" }

        $results += New-TimelineRecord -Timestamp $stamp -Name "Office 365 Apps" -Status $status -Property "M365 Apps" -Source "OfficeCSP" -Details $key.PSChildName
    }

    return $results
}

function Get-AutopilotJSONSettings {
    $path = "HKLM:\Software\Microsoft\Provisioning\Diagnostics\AutoPilot"
    if (-not (Test-Path $path)) {
        return $null
    }

    $values = Get-ItemProperty -Path $path
    [PSCustomObject]@{
        CloudAssignedTenantDomain = $values.CloudAssignedTenantDomain
        CloudAssignedTenantId = $values.CloudAssignedTenantId
        CloudAssignedAadServerData = $values.CloudAssignedAadServerData
        CloudAssignedForcedEnrollment = $values.CloudAssignedForcedEnrollment
        CloudAssignedOobeConfig = $values.CloudAssignedOobeConfig
        IsAutopilotDisabled = $values.IsAutopilotDisabled
    }
}

function Main {
    $log = New-LocalLogContext
    Write-LocalLogLine -Path $log.TextLog -Message "Autopilot X-Ray Report v2.0 started"

    Write-Host ""
    Write-Host "Autopilot X-Ray Report v2.0" -ForegroundColor Cyan
    Write-Host ("Computer: {0}    Started: {1}" -f $env:COMPUTERNAME, (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray

    $timeline = @()
    $autopilotEvent = Get-AutopilotProfileEvent
    if (-not $autopilotEvent) {
        Write-Host "This is not an Autopilot device (AutopilotDDSZTDFile.json not found)." -ForegroundColor Red
        Write-LocalLogLine -Path $log.TextLog -Message "Autopilot profile JSON not found. Exiting."
        return
    }

    $imeEntries = Get-IMELogEntries
    $appMap = Get-AppNameMapFromIME -IMEEntries $imeEntries

    $timeline += $autopilotEvent
    $timeline += Get-DeviceRegistrationEvents
    $timeline += Get-PolicyEvents
    $timeline += Get-SidecarInstallEvent
    $timeline += Get-ESPStatusEvents -AppMap $appMap
    $timeline += Get-Win32AppEvents -AppMap $appMap
    $timeline += Get-PlatformScriptEvents
    $timeline += Get-RemediationScriptEvents
    $timeline += Get-IMEProcessingEvents -IMEEntries $imeEntries -AppMap $appMap
    $timeline += Get-OfficeInstallEvents

    $timeline = $timeline | Where-Object { $_ } | Sort-Object Timestamp, Property, Name -Unique

    if (-not $timeline -or $timeline.Count -eq 0) {
        Write-Host "No timeline events were found." -ForegroundColor Red
        Write-LocalLogLine -Path $log.TextLog -Message "No timeline events returned from logs/registry."
        return
    }

    Write-TimelineConsole -Timeline $timeline
    Write-CategoryLegend

    $duration = $timeline[-1].Timestamp - $timeline[0].Timestamp
    $durationText = "{0}h {1}m {2}s" -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds
    Write-Host ""
    Write-Host "Autopilot Snapshot Duration: " -ForegroundColor White -NoNewline
    Write-Host $durationText -ForegroundColor Yellow
    Write-Host ""

    $timeline | Export-Csv -Path $log.CsvLog -NoTypeInformation -Encoding utf8
    $timeline | ConvertTo-Json -Depth 5 | Out-File -FilePath $log.JsonLog -Encoding utf8

    Write-LocalLogLine -Path $log.TextLog -Message "Autopilot Snapshot Duration: $durationText"
    foreach ($row in $timeline) {
        Write-LocalLogLine -Path $log.TextLog -Message ("[{0}] {1} | {2} | {3} | {4}" -f $row.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"), $row.Name, $row.Status, $row.Property, $row.Source)
    }

    $settings = Get-AutopilotJSONSettings
    if ($settings) {
        Write-LocalLogLine -Path $log.TextLog -Message "Autopilot JSON Settings:"
        ($settings | ConvertTo-Json -Depth 3) | Out-File -FilePath $log.TextLog -Append -Encoding utf8
    }

    Write-LocalLogLine -Path $log.TextLog -Message "Report complete"
    Write-Host "Exports:" -ForegroundColor White
    Write-Host "  CSV:  " -ForegroundColor DarkGray -NoNewline; Write-Host $log.CsvLog  -ForegroundColor Green
    Write-Host "  JSON: " -ForegroundColor DarkGray -NoNewline; Write-Host $log.JsonLog -ForegroundColor Green
    Write-Host "  LOG:  " -ForegroundColor DarkGray -NoNewline; Write-Host $log.TextLog -ForegroundColor Green
}

Main
