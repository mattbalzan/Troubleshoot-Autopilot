#Requires -Version 5.1
<#
.SYNOPSIS
    Collect an Autopilot diagnostics CAB from the local device and produce an HTML
    report using the Get-AutopilotDiagnostics (Michael Niehaus) parsing approach.

.DESCRIPTION
    This script is a self-contained re-implementation that:
      1. Collects a diagnostics CAB locally with MdmDiagnosticsTool.exe -area Autopilot.
      2. Parses it exactly the way Get-AutopilotDiagnostics.ps1 does:
           - Extracts MdmDiagReport_RegistryDump.reg + the two .evtx logs + the
             AutopilotDDSZTDFile.json + hardware-hash CSV.
           - Edits and imports the registry dump into HKCU:\ESPStatus.tmp.
           - Reads the Autopilot profile / OOBE config / ESP settings.
           - Builds an "observed timeline" via RecordStatus from device-management and
             user-device-registration events and the ESP tracking (policies, MSI apps,
             modern apps, Win32 sidecar apps, SCEP certs) for the device + user sessions.
      3. Resolves Win32 app friendly names from the IME (AppWorkload) logs if present.
      4. Renders a single HTML report (no Markdown).
      5. Deletes the created CAB and all extracted files afterwards (only the HTML
         report + parser log remain).

    Based on Get-AutopilotDiagnostics.ps1 5.6 by Michael Niehaus (Microsoft).

.PARAMETER OutputPath
    Root folder for report runs. Each run is written to
    <OutputPath>\Run_<COMPUTERNAME>_<timestamp>\AutopilotReport.html.
    Default: .\AutopilotReport

.PARAMETER AllSessions
    Include every ESP tracking session instead of just the final one.

.EXAMPLE
    .\Autopilot-CAB-Report.ps1 -OutputPath 'C:\temp\AutopilotReport'
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path (Get-Location) 'AutopilotReport'),
    [switch]$AllSessions
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web | Out-Null

# ---------------------------------------------------------------------------
# region  Per-run folder layout
# ---------------------------------------------------------------------------
$runStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$RunRoot  = Join-Path $OutputPath ('Run_{0}_{1}' -f $env:COMPUTERNAME, $runStamp)
$Extract  = Join-Path $RunRoot 'extracted'
New-Item -Path $Extract -ItemType Directory -Force | Out-Null
$script:Log     = Join-Path $RunRoot 'parser.log'
$HtmlReport     = Join-Path $RunRoot 'AutopilotReport.html'
$CabPath        = Join-Path $RunRoot ('AutopilotDiag_{0}_{1}.cab' -f $env:COMPUTERNAME, $runStamp)

function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR')]$Level='INFO')
    $line = '{0}  [{1}]  {2}' -f (Get-Date).ToString('s'), $Level, $Message
    Write-Verbose $line
    $line | Out-File -FilePath $script:Log -Append -Encoding utf8
}
function HtmlEnc { param([string]$s) if ($null -eq $s) { return '' } [System.Web.HttpUtility]::HtmlEncode($s) }

# ---------------------------------------------------------------------------
# region  Status / colour constants (from Get-AutopilotDiagnostics)
# ---------------------------------------------------------------------------
$script:officeStatus = @{'0'='None';'10'='Initialized';'20'='Download In Progress';'25'='Pending Download Retry';
    '30'='Download Failed';'40'='Download Completed';'48'='Pending User Session';'50'='Enforcement In Progress';
    '55'='Pending Enforcement Retry';'60'='Enforcement Failed';'70'='Success / Enforcement Completed'}
$script:espStatus    = @{'1'='Not Installed';'2'='Downloading / Installing';'3'='Success / Installed';'4'='Error / Failed'}
$script:policyStatus = @{'0'='Not Processed';'1'='Processed'}

# Consolidated timeline (Date, Detail, Status, Color) + structured detail rows
$script:observedTimeline = @()
$script:rowPolicies = @()
$script:rowApps     = @()
$script:rowCerts    = @()

function RecordStatus {
    param([Parameter(Mandatory)][string]$detail,
          [Parameter(Mandatory)][string]$status,
          [Parameter(Mandatory)][string]$color,
          [Parameter(Mandatory)][datetime]$date)
    $found = $script:observedTimeline | Where-Object { $_.Detail -eq $detail -and $_.Status -eq $status }
    if (-not $found) {
        $script:observedTimeline += [pscustomobject]@{ Date=$date; Detail=$detail; Status=$status; Color=$color }
    }
}
function ConvertTo-Date { param($v) try { [datetime]$v } catch { Get-Date } }

# ---------------------------------------------------------------------------
# region  1. Collect the diagnostics CAB locally
# ---------------------------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Write-Warning 'Not running elevated - MdmDiagnosticsTool may collect an incomplete CAB. Re-run as Administrator for best results.' }

$tool = Join-Path $env:windir 'system32\MdmDiagnosticsTool.exe'
if (-not (Test-Path $tool)) { throw "MdmDiagnosticsTool.exe not found at $tool" }

Write-Host "Collecting MDM diagnostics on $env:COMPUTERNAME ... (this can take a minute or two)" -ForegroundColor Cyan
Write-Log "Collecting CAB via: $tool -area Autopilot -cab $CabPath"
& $tool -area Autopilot -cab $CabPath | Out-Null
if (-not (Test-Path $CabPath)) { throw "MdmDiagnosticsTool did not produce a CAB at $CabPath" }
Write-Host "CAB collected: $CabPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# region  2. Extract the CAB
# ---------------------------------------------------------------------------
Write-Log "Expanding CAB into $Extract"
& expand.exe "$CabPath" -F:* -R "$Extract" | Out-Null
# Some CABs embed a nested LicensingDiag.cab - expand it too if present
Get-ChildItem -Path $Extract -Recurse -Filter '*.cab' -ErrorAction SilentlyContinue | ForEach-Object {
    & expand.exe "$($_.FullName)" -F:* -R "$Extract" | Out-Null
}

function Find-File { param([string]$Pattern)
    Get-ChildItem -Path $Extract -Recurse -Filter $Pattern -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
}
$regFile   = Find-File 'MdmDiagReport_RegistryDump.reg'
$devEvtx   = Find-File 'microsoft-windows-devicemanagement-enterprise-diagnostics-provider-admin.evtx'
$udrEvtx   = Find-File 'microsoft-windows-user device registration-admin.evtx'
$jsonFile  = Find-File 'AutopilotDDSZTDFile.json'
$csvFile   = Get-ChildItem -Path $Extract -Recurse -Filter '*.csv' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
if (-not $regFile) { throw "MdmDiagReport_RegistryDump.reg not found in the CAB - cannot continue." }

$hash = $null
if ($csvFile) { try { $hash = (Get-Content $csvFile | ConvertFrom-Csv).'Hardware Hash' } catch {} }

# ---------------------------------------------------------------------------
# region  3. Import the registry dump into HKCU:\ESPStatus.tmp  (GAD approach)
# ---------------------------------------------------------------------------
$editedReg = Join-Path $Extract 'MdmDiagReport_Edited.reg'
$content = Get-Content -Path $regFile
$content = $content -replace '\[HKEY_CURRENT_USER\\', '[HKEY_CURRENT_USER\ESPStatus.tmp\USER\'
$content = $content -replace '\[HKEY_LOCAL_MACHINE\\', '[HKEY_CURRENT_USER\ESPStatus.tmp\MACHINE\'
$content = $content -replace '^    "', '"'
$content = $content -replace '^    @', '@'
$content = $content -replace 'DWORD:', 'dword:'
"Windows Registry Editor Version 5.00`n" | Set-Content -Path $editedReg -Encoding Unicode
$content | Add-Content -Path $editedReg -Encoding Unicode

if (Test-Path 'HKCU:\ESPStatus.tmp') { Remove-Item -Path 'HKCU:\ESPStatus.tmp' -Recurse -Force }
# reg.exe writes its success banner to stderr. Under $ErrorActionPreference='Stop'
# (the default in Windows PowerShell 5.1 here) any native command stderr surfaces as a
# terminating NativeCommandError, which would abort before the report is written.
# Suppress stderr inside cmd so PowerShell never sees it, and relax EAP just in case.
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
cmd.exe /c "reg.exe IMPORT `"$editedReg`" >nul 2>&1" | Out-Null
$ErrorActionPreference = $prevEAP
Write-Log "Imported registry dump into HKCU:\ESPStatus.tmp"

# Constants pointing at the imported (offline) hive
$provisioningPath  = 'HKCU:\ESPStatus.tmp\MACHINE\software\microsoft\provisioning'
$autopilotDiagPath = "$provisioningPath\Diagnostics\Autopilot"
$omadmPath         = "$provisioningPath\OMADM"
$espPath           = 'HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics'
$msiPath           = 'HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\EnterpriseDesktopAppManagement'
$officePath        = 'HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\OfficeCSP'
$sidecarPath       = 'HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\IntuneManagementExtension\Win32Apps'
$enrollmentsPath   = 'HKCU:\ESPStatus.tmp\MACHINE\Software\Microsoft\enrollments'

# ---------------------------------------------------------------------------
# region  4. Win32 app friendly names from the IME (AppWorkload) logs
# ---------------------------------------------------------------------------
$script:AppNameMap = @{}
$intentText = @{ '1'='Available'; '3'='Required'; '4'='Uninstall' }
$awLogs = Get-ChildItem -Path $Extract -Recurse -Filter 'AppWorkload*.log' -ErrorAction SilentlyContinue
foreach ($awl in $awLogs) {
    $awRaw = Get-Content $awl.FullName -Raw -ErrorAction SilentlyContinue
    if (-not $awRaw) { continue }
    foreach ($m in [regex]::Matches($awRaw,
        '"Id":"(?<id>[0-9a-fA-F-]{8,})","Name":"(?<name>(?:[^"\\]|\\.)*)","Version":\d+,"Intent":(?<intent>\d+)')) {
        $gid = $m.Groups['id'].Value.ToLower()
        if ($script:AppNameMap.ContainsKey($gid)) { continue }
        $nm = $m.Groups['name'].Value -replace '\\"','"' -replace '\\\\','\'
        $it = $m.Groups['intent'].Value
        $script:AppNameMap[$gid] = [pscustomobject]@{ Name=$nm; Intent=$intentText[$it] }
    }
}
if ($script:AppNameMap.Count) { Write-Log "App name map: $($script:AppNameMap.Count) Win32 app names parsed from $($awLogs.Count) AppWorkload log(s)" }
function Get-AppName { param([string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
    $g = [regex]::Match($Id,'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}')
    $key = if ($g.Success) { $g.Value.ToLower() } else { $Id.ToLower() }
    $script:AppNameMap[$key]
}

# ---------------------------------------------------------------------------
# region  5. ESP tracking processors (adapted from Get-AutopilotDiagnostics)
# ---------------------------------------------------------------------------
# Returns the session subkeys to process: the last (final) one, or all with -AllSessions
function Get-SessionKeys { param([string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    $items = @(Get-ChildItem $Path | Sort-Object PSChildName)
    if (-not $items.Count) { return @() }
    if ($AllSessions) { return $items }
    return ,($items[-1])
}

function Read-Apps { param($Key,[string]$CurrentUser)
    $date = ConvertTo-Date $Key.PSChildName
    foreach ($p in $Key.Property) {
        if ($p.StartsWith('./Device/Vendor/MSFT/EnterpriseDesktopAppManagement/MSI/')) {
            $msiKey = [URI]::UnescapeDataString(($p.Split('/'))[6])
            $full = "$msiPath\$CurrentUser\MSI\$msiKey"
            $status = 0; $msiFile = $null
            if (Test-Path $full) { $status = (Get-ItemProperty $full).Status; $msiFile = (Get-ItemProperty $full).CurrentDownloadUrl }
            if ($null -eq $status -or $status -eq '') { $status = 0 }
            $name = $msiKey
            if ($msiFile -match 'IntuneWindowsAgent.msi') { $name = "Intune Management Extension" }
            $txt = $officeStatus["$status"]; if (-not $txt) { $txt = "$status" }
            $color = if ($status -eq 70) {'Green'} elseif ($status -eq 60) {'Red'} else {'Yellow'}
            $script:rowApps += [pscustomobject]@{ Type='MSI'; Name=$name; Id=$msiKey; Status=$txt; Color=$color }
            RecordStatus -detail "MSI $name" -status $txt -color $color -date $date
        }
        elseif ($p.StartsWith('./Vendor/MSFT/Office/Installation/')) {
            $status = Get-ItemPropertyValue -Path $Key.PSPath -Name $p
            $officeKey = [URI]::UnescapeDataString(($p.Split('/'))[5])
            $oStatus = 'None'
            $full = "$officePath\$officeKey"
            if (Test-Path $full) {
                $oStatus = (Get-ItemProperty $full).FinalStatus
                if ($null -eq $oStatus) { $oStatus = (Get-ItemProperty $full).Status }
                if ($null -eq $oStatus) { $oStatus = 'None' }
            }
            $oTxt = $officeStatus["$oStatus"]; if (-not $oTxt) { $oTxt = "$oStatus" }
            $combined = "$($policyStatus["$status"]) / $oTxt"
            $color = if ($status -eq 1) {'Green'} else {'Yellow'}
            $script:rowApps += [pscustomobject]@{ Type='Office'; Name=$officeKey; Id=$officeKey; Status=$combined; Color=$color }
            RecordStatus -detail "Office $officeKey" -status $combined -color $color -date $date
        }
    }
}

function Read-ModernApps { param($Key,[string]$CurrentUser)
    $date = ConvertTo-Date $Key.PSChildName
    foreach ($p in $Key.Property) {
        $status = (Get-ItemPropertyValue -Path $Key.PSPath -Name $p).ToString()
        if ($p.StartsWith('./User/Vendor/MSFT/EnterpriseModernAppManagement/AppManagement/'))   { $appID=[URI]::UnescapeDataString(($p.Split('/'))[7]); $type='User UWP' }
        elseif ($p.StartsWith('./Device/Vendor/MSFT/EnterpriseModernAppManagement/AppManagement/')) { $appID=[URI]::UnescapeDataString(($p.Split('/'))[7]); $type='Device UWP' }
        else { $appID = $p; $type='UWP' }
        $txt = $policyStatus["$status"]; if (-not $txt) { $txt = "$status" }
        if ($status -eq '1') {
            $script:rowApps += [pscustomobject]@{ Type=$type; Name=$appID; Id=$appID; Status=$txt; Color='Green' }
            RecordStatus -detail "UWP $appID" -status $txt -color 'Green' -date $date
        } else {
            $script:rowApps += [pscustomobject]@{ Type=$type; Name=$appID; Id=$appID; Status=$txt; Color='Yellow' }
        }
    }
}

function Read-Sidecar { param($Key,[string]$CurrentUser)
    $date = ConvertTo-Date $Key.PSChildName
    foreach ($p in $Key.Property) {
        $win32Key = [URI]::UnescapeDataString(($p.Split('/'))[9])
        $status = Get-ItemPropertyValue -Path $Key.PSPath -Name $p
        $named = Get-AppName $win32Key
        $display = if ($named) { $named.Name } else { $win32Key }
        $exitCode = $null
        $g = [regex]::Match($win32Key,'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}')
        if ($g.Success -and (Test-Path "$sidecarPath\$CurrentUser")) {
            $appKey = Get-ChildItem "$sidecarPath\$CurrentUser" -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match $g.Value } | Select-Object -First 1
            if ($appKey) { $exitCode = (Get-ItemProperty $appKey.PSPath).ExitCode }
        }
        $txt = $espStatus["$status"]; if (-not $txt) { $txt = "$status" }
        if ($null -ne $exitCode) { $txt = "$txt (rc=$exitCode)" }
        $color = if ($status -eq '3') {'Green'} elseif ($status -eq '4') {'Red'} else {'Yellow'}
        $script:rowApps += [pscustomobject]@{ Type='Win32'; Name=$display; Id=$win32Key; Status=$txt; Color=$color }
        if ($status -ne '1') { RecordStatus -detail "Win32 $display" -status $txt -color $color -date $date }
    }
}

function Read-Policies { param($Key)
    $date = ConvertTo-Date $Key.PSChildName
    foreach ($p in $Key.Property) {
        $status = Get-ItemPropertyValue -Path $Key.PSPath -Name $p
        $txt = $policyStatus["$status"]; if (-not $txt) { $txt = "$status" }
        $color = if ($status -eq '1') {'Green'} else {'Yellow'}
        $script:rowPolicies += [pscustomobject]@{ Name=$p; Status=$txt; Color=$color }
        RecordStatus -detail "Policy $p" -status $txt -color $color -date $date
    }
}

function Read-Certs { param($Key)
    $date = ConvertTo-Date $Key.PSChildName
    foreach ($p in $Key.Property) {
        $certKey = [URI]::UnescapeDataString(($p.Split('/'))[6])
        $status = Get-ItemPropertyValue -Path $Key.PSPath -Name $p
        $txt = $policyStatus["$status"]; if (-not $txt) { $txt = "$status" }
        $color = if ($status -eq '1') {'Green'} else {'Yellow'}
        $script:rowCerts += [pscustomobject]@{ Name=$certKey; Status=$txt; Color=$color }
        RecordStatus -detail "Cert $certKey" -status $txt -color $color -date $date
    }
}

# ---------------------------------------------------------------------------
# region  6. Event-log processing (ODJ / enrollment / app / device registration)
# ---------------------------------------------------------------------------
function Read-Events {
    $productCode = 'IME-Not-Yet-Installed'
    $imeMsiRoot = "$msiPath\S-0-0-00-0000000000-0000000000-000000000-000\MSI"
    if (Test-Path $imeMsiRoot) {
        $found = Get-ChildItem $imeMsiRoot | Where-Object { (Get-ItemProperty $_.PSPath).CurrentDownloadUrl -match 'IntuneWindowsAgent.msi' } | Select-Object -First 1
        if ($found) { $productCode = (Get-ItemProperty $found.PSPath).ProductCode }
    }
    if ($devEvtx -and (Test-Path $devEvtx)) {
        $events = Get-WinEvent -Path $devEvtx -Oldest -ErrorAction SilentlyContinue |
            Where-Object { ($_.Message -match $productCode -and $_.Id -in 1905,1906,1920,1922) -or $_.Id -in 72,100,107,109,110,111 }
        foreach ($e in $events) {
            $message = $e.Message; $detail = 'Sidecar'; $color = 'Yellow'
            switch ($e.Id) {
                {$_ -in 110,109} {
                    $detail = 'Offline Domain Join'
                    switch ($e.Properties[0].Value) {
                        0 { $message = 'Offline domain join not configured' }
                        1 { $message = 'Waiting for ODJ blob' }
                        2 { $message = 'Processed ODJ blob' }
                        3 { $message = 'Timed out waiting for ODJ blob or connectivity' }
                    }
                }
                111 { $detail='Offline Domain Join'; $message='Starting wait for ODJ blob' }
                107 { $detail='Offline Domain Join'; $message='Successfully applied ODJ blob' }
                100 { $detail='Offline Domain Join'; $message='Could not establish connectivity'; $color='Red' }
                72  { $detail='MDM Enrollment' }
                1905 { $message='Download started' }
                1906 { $message='Download finished' }
                1920 { $message='Installation started' }
                1922 { $message='Installation finished' }
                {$_ -in 1922,72} { $color='Green' }
            }
            RecordStatus -detail $detail -status $message -color $color -date $e.TimeCreated
        }
    }
    if ($udrEvtx -and (Test-Path $udrEvtx)) {
        $events = Get-WinEvent -Path $udrEvtx -Oldest -ErrorAction SilentlyContinue | Where-Object { $_.Id -in 306,101 }
        foreach ($e in $events) {
            $message = $e.Message; $detail = 'Device Registration'; $color = 'Yellow'
            switch ($e.Id) {
                101 { $message='SCP discovery successful.' }
                306 { $message='Hybrid Azure AD device registration succeeded.'; $color='Green' }
            }
            RecordStatus -detail $detail -status $message -color $color -date $e.TimeCreated
        }
    }
}

# ---------------------------------------------------------------------------
# region  7. Autopilot profile / OOBE config / scenario / ESP settings
# ---------------------------------------------------------------------------
$summary  = [ordered]@{}
$oobeFlags = @()
$espSettings = [ordered]@{}
$hardware = [ordered]@{}
$isAutopilot = $false

# The OOBE-era values under Diagnostics\Autopilot are populated during provisioning
# and cleared afterwards, so on an already-provisioned device they can be empty even
# though the device IS an Autopilot device. Treat the presence of the diag key (or a
# cached profile / the DDSZTD JSON) as sufficient evidence and surface whatever exists.
$values = Get-ItemProperty $autopilotDiagPath -ErrorAction SilentlyContinue
$apCache = Get-ItemProperty "$provisioningPath\AutopilotPolicyCache" -ErrorAction SilentlyContinue
if ($values -or $apCache -or ($jsonFile -and (Test-Path $jsonFile))) {
    $isAutopilot = $true
    if ($values) {
        if ($values.DeploymentProfileName)     { $summary['Profile']      = $values.DeploymentProfileName }
        if ($values.CloudAssignedTenantDomain) { $summary['Tenant domain']= $values.CloudAssignedTenantDomain }
        if ($values.CloudAssignedTenantId)     { $summary['Tenant ID']    = $values.CloudAssignedTenantId }
        $corr = Get-ItemProperty "$autopilotDiagPath\EstablishedCorrelations" -ErrorAction SilentlyContinue
        if ($corr) {
            if ($corr.ZTDRegistrationID) { $summary['ZTDID']   = $corr.ZTDRegistrationID }
            if ($corr.EntDMID)           { $summary['EntDMID'] = $corr.EntDMID }
        }
        if ($null -ne $values.CloudAssignedOobeConfig -and $values.CloudAssignedOobeConfig -ne '') {
            $cfg = [int]$values.CloudAssignedOobeConfig
            $summary['OobeConfig'] = $cfg
            $flagMap = [ordered]@{
                'Skip keyboard'=1024; 'Enable patch download'=512; 'Skip Windows upgrade UX'=256; 'AAD TPM required'=128;
                'AAD device auth'=64; 'TPM attestation'=32; 'Skip EULA'=16; 'Skip OEM registration'=8;
                'Skip express settings'=4; 'Disallow admin'=2
            }
            foreach ($k in $flagMap.Keys) { $oobeFlags += [pscustomobject]@{ Name=$k; Value=(([int]$cfg -band $flagMap[$k]) -gt 0) } }
        }
    }
    if (-not $summary['Profile'] -and $apCache -and $apCache.DeploymentProfileName) { $summary['Profile'] = $apCache.DeploymentProfileName }
}

# Scenario + profile-download timeline entry from the JSON
if ($jsonFile -and (Test-Path $jsonFile)) {
    try {
        $json = Get-Content $jsonFile | ConvertFrom-Json
        if ($json.PolicyDownloadDate) { RecordStatus -date (ConvertTo-Date $json.PolicyDownloadDate) -detail 'Autopilot profile' -status 'Profile downloaded' -color 'Yellow' }
        if (-not $summary['Tenant ID']     -and $json.CloudAssignedTenantId)     { $summary['Tenant ID']     = $json.CloudAssignedTenantId }
        if (-not $summary['Tenant domain'] -and $json.CloudAssignedTenantDomain) { $summary['Tenant domain'] = $json.CloudAssignedTenantDomain }
        if (-not $summary['Profile']       -and $json.DeploymentProfileName)     { $summary['Profile']       = $json.DeploymentProfileName }
        if ($json.CloudAssignedDomainJoinMethod -eq 1) {
            $summary['Scenario'] = 'Hybrid Azure AD Join'
            $summary['ODJ applied'] = (Test-Path "$omadmPath\SyncML\ODJApplied")
        } else {
            $summary['Scenario'] = 'Azure AD Join'
        }
    } catch { Write-Log "Failed to parse AutopilotDDSZTDFile.json: $($_.Exception.Message)" 'WARN' }
} else {
    $summary['Scenario'] = 'Not available (JSON not found - device already provisioned)'
}

# ESP settings (FirstSync)
Get-ChildItem $enrollmentsPath -ErrorAction SilentlyContinue | Where-Object { Test-Path "$($_.PSPath)\FirstSync" } | ForEach-Object {
    $pr = Get-ItemProperty "$($_.PSPath)\FirstSync"
    $espSettings['Device ESP enabled'] = ($pr.SkipDeviceStatusPage -eq 0)
    $espSettings['User ESP enabled']   = ($pr.SkipUserStatusPage -eq 0)
    $espSettings['ESP timeout (sec)']  = $pr.SyncFailureTimeout
    if ($pr.BlockInStatusPage -eq 0) { $espSettings['ESP blocking'] = 'No' }
    else {
        $espSettings['ESP blocking'] = 'Yes'
        $espSettings['Allow reset']     = [bool]($pr.BlockInStatusPage -band 1)
        $espSettings['Allow try again'] = [bool]($pr.BlockInStatusPage -band 2)
        $espSettings['Continue anyway'] = [bool]($pr.BlockInStatusPage -band 4)
    }
}

# Hardware info via WMI/CIM (best effort)
try {
    $cs   = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
    $os   = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $tpm  = Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction SilentlyContinue
    if ($os)   { $hardware['OS build']      = $os.BuildNumber }
    if ($cs)   { $hardware['Manufacturer']  = $cs.Manufacturer }
    if ($cs)   { $hardware['Model']         = $cs.Model }
    if ($bios) { $hardware['Serial number'] = $bios.SerialNumber }
    if ($tpm -and $tpm.SpecVersion) { $hardware['TPM version'] = ($tpm.SpecVersion -split ',')[0].Trim() }
} catch { Write-Log "WMI/CIM hardware query failed: $($_.Exception.Message)" 'WARN' }

# ---------------------------------------------------------------------------
# region  8. Run the processors
# ---------------------------------------------------------------------------
Read-Events

if (Test-Path $espPath) {
    # Device session
    foreach ($k in (Get-SessionKeys "$espPath\ExpectedPolicies"))        { Read-Policies   -Key $k }
    foreach ($k in (Get-SessionKeys "$espPath\ExpectedMSIAppPackages"))  { Read-Apps       -Key $k -CurrentUser 'S-0-0-00-0000000000-0000000000-000000000-000' }
    foreach ($k in (Get-SessionKeys "$espPath\ExpectedModernAppPackages")) { Read-ModernApps -Key $k -CurrentUser 'S-0-0-00-0000000000-0000000000-000000000-000' }
    if (Test-Path "$espPath\Sidecar") {
        $sc = @(Get-ChildItem "$espPath\Sidecar" | Where-Object { $_.Property -match './Device' } | Sort-Object PSChildName)
        if ($sc.Count) { $set = if ($AllSessions) { $sc } else { ,($sc[-1]) }; foreach ($k in $set) { Read-Sidecar -Key $k -CurrentUser '00000000-0000-0000-0000-000000000000' } }
    }
    foreach ($k in (Get-SessionKeys "$espPath\ExpectedSCEPCerts"))       { Read-Certs      -Key $k }

    # User sessions
    Get-ChildItem $espPath -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName.StartsWith('S-') } | ForEach-Object {
        $userPath = $_.PSPath; $userSid = $_.PSChildName
        foreach ($k in (Get-SessionKeys "$userPath\ExpectedPolicies"))        { Read-Policies   -Key $k }
        foreach ($k in (Get-SessionKeys "$userPath\ExpectedMSIAppPackages"))  { Read-Apps       -Key $k -CurrentUser $userSid }
        foreach ($k in (Get-SessionKeys "$userPath\ExpectedModernAppPackages")) { Read-ModernApps -Key $k -CurrentUser $userSid }
        if (Test-Path "$userPath\Sidecar") {
            $sc = @(Get-ChildItem "$espPath\Sidecar" | Where-Object { $_.Property -match './User' } | Sort-Object PSChildName)
            if ($sc.Count) { $set = if ($AllSessions) { $sc } else { ,($sc[-1]) }; foreach ($k in $set) { Read-Sidecar -Key $k -CurrentUser $userSid } }
        }
        foreach ($k in (Get-SessionKeys "$userPath\ExpectedSCEPCerts"))       { Read-Certs      -Key $k }
    }
}

# ---------------------------------------------------------------------------
# region  9. Verdict + timeline ordering
# ---------------------------------------------------------------------------
$timeline = $script:observedTimeline | Sort-Object Date
$redCount   = @($timeline | Where-Object { $_.Color -eq 'Red' }).Count
$greenCount = @($timeline | Where-Object { $_.Color -eq 'Green' }).Count
$times = @($timeline | Where-Object { $_.Date } | Select-Object -ExpandProperty Date)
$apStart = if ($times.Count) { ($times | Sort-Object | Select-Object -First 1) } else { $null }
$apEnd   = if ($times.Count) { ($times | Sort-Object | Select-Object -Last 1) } else { $null }
$durationText = if ($apStart -and $apEnd -and $apEnd -ne $apStart) {
    $d = $apEnd - $apStart; ('{0}h {1}m {2}s' -f [int]$d.TotalHours, $d.Minutes, $d.Seconds)
} else { 'Indeterminate (single timeline entry retained in this CAB)' }

if (-not $isAutopilot) { $verdict = 'This is not an Autopilot device (no CloudAssignedTenantId).'; $verdictColor = '#797775' }
elseif ($redCount -gt 0) { $verdict = "Errors recorded during provisioning ($redCount failed item(s))"; $verdictColor = '#d13438' }
else { $verdict = 'No errors recorded in the provisioning timeline'; $verdictColor = '#107c10' }

# ---------------------------------------------------------------------------
# region  10. HTML report (same styling as before)
# ---------------------------------------------------------------------------
$colorHex = @{ Green='#107c10'; Red='#d13438'; Yellow='#ca5010' }
function Badge { param($Color,$Text) $hex = $colorHex[$Color]; if (-not $hex) { $hex='#8a8886' } "<span class='badge' style='background:$hex'>$(HtmlEnc $Text)</span>" }

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8"><title>Autopilot Diagnostics Report</title><style>')
[void]$sb.AppendLine('body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f3f2f1;color:#201f1e}')
[void]$sb.AppendLine('.wrap{max-width:1100px;margin:0 auto;padding:24px}')
[void]$sb.AppendLine('h1{font-size:24px}h2{margin-top:32px;border-bottom:2px solid #edebe9;padding-bottom:6px}')
[void]$sb.AppendLine('.card{background:#fff;border-radius:8px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,.1);margin-bottom:16px}')
[void]$sb.AppendLine('.verdict{font-size:20px;font-weight:600;color:#fff;padding:14px 18px;border-radius:8px}')
[void]$sb.AppendLine('table{border-collapse:collapse;width:100%;background:#fff;border-radius:8px;overflow:hidden}')
[void]$sb.AppendLine('th,td{text-align:left;padding:10px 12px;border-bottom:1px solid #edebe9;vertical-align:top;font-size:14px}')
[void]$sb.AppendLine('th{background:#faf9f8;font-weight:600}')
[void]$sb.AppendLine('.badge{display:inline-block;padding:2px 10px;border-radius:12px;color:#fff;font-size:12px;font-weight:600}')
[void]$sb.AppendLine('.kv{display:grid;grid-template-columns:240px 1fr;gap:4px 12px;font-size:14px}.kv b{color:#605e5c;font-weight:600}')
[void]$sb.AppendLine('small{color:#605e5c}</style></head><body><div class="wrap">')
[void]$sb.AppendLine('<h1>Windows Autopilot Diagnostics Report</h1>')
[void]$sb.AppendLine("<div class='card'><div class='verdict' style='background:$verdictColor'>$(HtmlEnc $verdict)</div>")
[void]$sb.AppendLine("<p><small>Device:</small> <b>$(HtmlEnc $env:COMPUTERNAME)</b> &nbsp; <small>Generated:</small> $(Get-Date -Format 'u')<br><small>Source:</small> diagnostics CAB collected locally and parsed (CAB + extracted files removed after parsing)</p>")
[void]$sb.AppendLine("<div class='kv'><b>First timeline entry</b><span>$(if($apStart){$apStart.ToString('yyyy-MM-dd HH:mm:ss')}else{'-'})</span><b>Last timeline entry</b><span>$(if($apEnd){$apEnd.ToString('yyyy-MM-dd HH:mm:ss')}else{'-'})</span><b>Elapsed</b><span>$(HtmlEnc $durationText)</span></div>")
[void]$sb.AppendLine('</div>')

# Autopilot summary
[void]$sb.AppendLine('<h2>Autopilot profile</h2><div class="card"><div class="kv">')
foreach ($kv in $summary.GetEnumerator()) { [void]$sb.AppendLine("<b>$(HtmlEnc $kv.Key)</b><span>$(HtmlEnc ([string]$kv.Value))</span>") }
[void]$sb.AppendLine('</div></div>')

# OOBE flags
if ($oobeFlags.Count) {
    [void]$sb.AppendLine('<h2>OOBE configuration</h2><table><tr><th>Setting</th><th>Enabled</th></tr>')
    foreach ($f in $oobeFlags) { [void]$sb.AppendLine("<tr><td>$(HtmlEnc $f.Name)</td><td>$(HtmlEnc ([string]$f.Value))</td></tr>") }
    [void]$sb.AppendLine('</table>')
}

# ESP settings
if ($espSettings.Count) {
    [void]$sb.AppendLine('<h2>Enrollment Status Page settings</h2><div class="card"><div class="kv">')
    foreach ($kv in $espSettings.GetEnumerator()) { [void]$sb.AppendLine("<b>$(HtmlEnc $kv.Key)</b><span>$(HtmlEnc ([string]$kv.Value))</span>") }
    [void]$sb.AppendLine('</div></div>')
}

# Hardware
if ($hardware.Count) {
    [void]$sb.AppendLine('<h2>Hardware</h2><div class="card"><div class="kv">')
    foreach ($kv in $hardware.GetEnumerator()) { [void]$sb.AppendLine("<b>$(HtmlEnc $kv.Key)</b><span>$(HtmlEnc ([string]$kv.Value))</span>") }
    [void]$sb.AppendLine('</div></div>')
}

# Observed timeline (same method/layout as Get-AutopilotDiagnostics: Date | Status | Detail, in date order)
[void]$sb.AppendLine('<h2>Observed timeline</h2><table><tr><th>Date / time</th><th>Status</th><th>Detail</th></tr>')
if ($timeline.Count) {
    foreach ($t in $timeline) { [void]$sb.AppendLine("<tr><td><small>$($t.Date.ToString('yyyy-MM-dd HH:mm:ss'))</small></td><td>$(Badge $t.Color $t.Status)</td><td>$(HtmlEnc $t.Detail)</td></tr>") }
} else { [void]$sb.AppendLine('<tr><td colspan="3">No timeline events were recorded in this CAB.</td></tr>') }
[void]$sb.AppendLine('</table>')

# Policies
[void]$sb.AppendLine('<h2>Policies tracked by the ESP</h2><table><tr><th>Policy</th><th>Status</th></tr>')
if ($script:rowPolicies.Count) { foreach ($r in $script:rowPolicies) { [void]$sb.AppendLine("<tr><td><small>$(HtmlEnc $r.Name)</small></td><td>$(Badge $r.Color $r.Status)</td></tr>") } }
else { [void]$sb.AppendLine('<tr><td colspan="2">No policies tracked by the ESP in this CAB.</td></tr>') }
[void]$sb.AppendLine('</table>')

# Applications (MSI / Office / UWP / Win32)
[void]$sb.AppendLine('<h2>Applications tracked by the ESP</h2><table><tr><th>Type</th><th>App name</th><th>App ID</th><th>Status</th></tr>')
if ($script:rowApps.Count) { foreach ($r in $script:rowApps) { [void]$sb.AppendLine("<tr><td>$(HtmlEnc $r.Type)</td><td><b>$(HtmlEnc $r.Name)</b></td><td><small>$(HtmlEnc $r.Id)</small></td><td>$(Badge $r.Color $r.Status)</td></tr>") } }
else { [void]$sb.AppendLine('<tr><td colspan="4">No applications tracked by the ESP in this CAB.</td></tr>') }
[void]$sb.AppendLine('</table>')

# Certificates
[void]$sb.AppendLine('<h2>SCEP certificates tracked by the ESP</h2><table><tr><th>Certificate</th><th>Status</th></tr>')
if ($script:rowCerts.Count) { foreach ($r in $script:rowCerts) { [void]$sb.AppendLine("<tr><td>$(HtmlEnc $r.Name)</td><td>$(Badge $r.Color $r.Status)</td></tr>") } }
else { [void]$sb.AppendLine('<tr><td colspan="2">No SCEP certificates tracked by the ESP in this CAB.</td></tr>') }
[void]$sb.AppendLine('</table>')

[void]$sb.AppendLine("<p style='margin-top:32px'><small>Win32 app names resolved from IME AppWorkload logs.</small></p>")
[void]$sb.AppendLine('</div></body></html>')

[System.IO.File]::WriteAllText($HtmlReport, $sb.ToString(), [System.Text.Encoding]::UTF8)

# ---------------------------------------------------------------------------
# region  11. Cleanup - remove imported hive, extracted files and the CAB
# ---------------------------------------------------------------------------
if (Test-Path 'HKCU:\ESPStatus.tmp') { Remove-Item -Path 'HKCU:\ESPStatus.tmp' -Recurse -Force -ErrorAction SilentlyContinue }
Remove-Item -Path $Extract -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path $CabPath -Force -ErrorAction SilentlyContinue
Write-Log "Cleanup complete - removed extracted files and CAB; HTML report retained."

# ---------------------------------------------------------------------------
# region  12. Console summary + open report
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host '==== AUTOPILOT DIAGNOSTICS REPORT ====' -ForegroundColor Magenta
Write-Host ("Device   : {0}" -f $env:COMPUTERNAME)
Write-Host ("Verdict  : {0}" -f $verdict)
Write-Host ("Timeline : {0} entr(ies)  ({1} green / {2} red)" -f $timeline.Count, $greenCount, $redCount)
Write-Host ("Apps     : {0}   Policies: {1}   Certs: {2}" -f $script:rowApps.Count, $script:rowPolicies.Count, $script:rowCerts.Count)
Write-Host ''
Write-Host ("HTML report : {0}" -f $HtmlReport) -ForegroundColor Green
Write-Host ("Parser log  : {0}" -f $script:Log)
Write-Host ("Run folder  : {0}" -f $RunRoot)

try { Start-Process $HtmlReport } catch {}

[pscustomobject]@{
    Device        = $env:COMPUTERNAME
    Verdict       = $verdict
    TimelineCount = $timeline.Count
    AppCount      = $script:rowApps.Count
    PolicyCount   = $script:rowPolicies.Count
    CertCount     = $script:rowCerts.Count
    HtmlReport    = $HtmlReport
    RunFolder     = $RunRoot
}
