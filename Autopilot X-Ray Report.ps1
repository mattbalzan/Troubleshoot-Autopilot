# --[ Autopilot X-Ray Report      ]
# --[ Matt Balzan | mattGPT.co.uk ]

<#

    04.02.25 | Version 1.0 | Original.
    05.02.25 | Version 1.1 | Added parse app names using GUID.
    06.02.25 | Version 1.2 | Joined all results to show: Timestamp (local time) - Name - Status - Property.
    07.02.25 | Version 1.3 | Added First Events, Intune Win Agent, Office & Autopilot data tables & Autopilot Deployment time
    07.02.25 | Version 1.x | To-Do: 
                             
                             Platform scripts [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\Policies\*\a42b83a8-bab5-4eba-9cee-7424678a5d62]
                             Remediation scripts [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\SideCarPolicies\Scripts\Reports\cfe1502e-51da-4e14-8fec-bdd040a29dde\50b112a5-65e7-4633-9363-d7dd6004b86c_1\Result]
                             <![LOG[[PowerShell] Processing policy with id = a42b83a8-bab5-4eba-9cee-7424678a5d62 for user 00000000-0000-0000-0000-000000000000]LOG]!><time="15:09:33.9450446" date="2-4-2025" component="IntuneManagementExtension" context="" type="1" thread="14" file="">
                             <![LOG[[PowerShell] Policy body = ﻿# --[ Debloat AppX Packages from Windows Image ]
                             
                             Test UWP User & Device deployed apps
                             Get culture & OS, Patch info

#>


function getRegTime($regPath) {

# --[ Dll import for RegQuery ]
$signature = '[DllImport("advapi32.dll", CharSet = CharSet.Auto)]
public static extern Int32 RegQueryInfoKey(
Microsoft.Win32.SafeHandles.SafeRegistryHandle hKey,
StringBuilder lpClass,
Int32 lpCls, Int32 spare, Int32 subkeys,
Int32 skLen, Int32 mcLen, Int32 values,
Int32 vNLen, Int32 mvLen, Int32 secDesc,
out System.Runtime.InteropServices.ComTypes.FILETIME lpftLastWriteTime
);'


    $typename = "RegQueryInfoKey.GetRegData"

    # --[ Check if the type already exists ]
    #if (-not [System.Management.Automation.PSTypeName]::new($typename).Type) {
        
    # --[ Add the type if it does not exist ]
    $regData = Add-Type $signature -Name GetRegData -Namespace RegQueryInfoKey -Using System.Text -PassThru
    #}
            $reg = Get-Item $regPath -force

                if ($reg.handle) {

                $time = New-Object System.Runtime.InteropServices.ComTypes.FILETIME

                $result = $regData::RegQueryInfoKey($reg.Handle, $null, 0,0,0,0,0,0,0,0,0, [ref]$time)

                            if ($result -eq 0) {
                            
                                $low = [uint32]0 -bor $time.dwLowDateTime
                                $high = [uint32]0 -bor $time.dwHighDateTime
                                $timeValue = ([int64]$high -shl 32) -bor $low

                            return [datetime]::FromFileTime($timeValue)
                            }
                }
}


# --[ Gather Autopilot JSON ]
function Get-AutopilotStart{
    $results = @()
    
    # --[ Import Autopilot Profile ]
    $jsonFile = "$($env:WINDIR)\ServiceState\wmansvc\AutopilotDDSZTDFile.json" 
    
    
    if(Test-Path $jsonFile){
    
    $json = Get-Content $jsonFile | ConvertFrom-Json
    $date = $json.PolicyDownloadDate
    
    $results = [PSCustomObject]@{
                                    Timestamp = [datetime]$date
                                    Name = $json.DeploymentProfileName
                                    Status = "Downloaded"
                                    Property = "Autopilot Profile"
                                }
    
    return $results
    }

    else { return $false }
}


# --[ Get First Events data ]
function Get-FirstEvents{
$events = @()
$results = @()

# --[ Define Event Log Paths ]
$eventLogs = @("microsoft-windows-devicemanagement-enterprise-diagnostics-provider/admin", "microsoft-windows-user device registration/admin")

# --[ Event Message Mapping ]
$messageMap = @{
    101  = "Device Registration,Service Connection Point successful"
    306  = "Device Registration,Hybrid AADJ device registration succeeded"
    72   = "MDM Enrollment,Succeeded"
    111  = "Device Registration,Starting wait for ODJ blob"
    107  = "Device Registration,Successfully applied ODJ blob"
    100  = "Device Registration,Could not establish connectivity"
}

# --[ Collect Events ]
$eventLogs | % {
    $events += Get-WinEvent -LogName $_ -Oldest | Where-Object { $_.Id -in $messageMap.Keys }
}

# --[ Track and Omit Event 100 if Event 101 Found ]
$processedEvents = @()
$event101Found = $false

$events | % {

    $messageParts = $messageMap[$_.Id] -split ",", 2  # Split into two parts
    $results += [PSCustomObject]@{
        Timestamp = $_.TimeCreated
        Name      = $messageParts[0]  # The first part is Name
        Status    = $messageParts[1]  # The second part is Status
        Property  = "Device Phase 1"
    }
}

    return $results

}

# --[ Get Sidecar installation time ]
function Get-SidecarData{

$sidecarRegKey = Get-Item -Path "HKLM:\SOFTWARE\Microsoft\IntuneWindowsAgent"

$Timecreated = getRegTime ($sidecarRegKey -replace "HKEY_LOCAL_MACHINE","HKLM:")

    $results = [PSCustomObject]@{
                                        Timestamp = [datetime]$Timecreated
                                        Name = "Intune Windows Agent"
                                        Status = "Installed"
                                        Property = "Sidecar"
                                    }
    
        return $results

}


function Get-AutopilotJSONdata{
    # --[ Retrieve AutoPilot Configuration Values ]
    $values = Get-ItemProperty "HKLM:\Software\Microsoft\Provisioning\Diagnostics\AutoPilot"
    $flags = $values.CloudAssignedOobeConfig

    # --[ Define Configuration Settings with Bitwise Masks ]
    $configurations = @(
        @{ Name = "Skip keyboard";           Mask = 1024 }
        @{ Name = "Enable patch download";   Mask = 512 }
        @{ Name = "Skip Windows upgrade UX"; Mask = 256 }
        @{ Name = "AAD TPM Required";        Mask = 128 }
        @{ Name = "AAD device auth";         Mask = 64 }
        @{ Name = "TPM attestation";         Mask = 32 }
        @{ Name = "Skip EULA";               Mask = 16 }
        @{ Name = "Skip OEM registration";   Mask = 8 }
        @{ Name = "Skip express settings";   Mask = 4 }
        @{ Name = "Disallow admin";          Mask = 2 }
    )

    # --[ Generate ASCII Table Header ]
    $border = "+----------------------------+---------+---------+"
    $header = "| Autopilot JSON Settings    | Enabled | BitMask |"
    Write-Host $border
    Write-Host $header
    Write-Host $border

    # --[ Generate and Display ASCII Table Rows ]
    $configurations | ForEach-Object {
        $enabled = If (($flags -band $_.Mask) -gt 0) { "Yes" } Else { "No" }
        $bitMask = "$(($_.Mask).ToString("X").PadLeft(4,'0'))"
        Write-Host ("| {0,-26} | {1,-7} | {2,-7} |" -f $_.Name, $enabled, $bitMask)
    }

    Write-Host $border

}

# --[ Office data ]
function Get-OfficeInstallData{
# --[ Retrieve OfficeCSP Configuration Values ]
$OfficeCSP = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\OfficeCSP" -ErrorAction SilentlyContinue

if($OfficeCSP){

$OfficeStatus = @{
    "0" = "None"; "10" = "Initialized"; "20" = "Download In Progress"; "25" = "Pending Download Retry";
    "30" = "Download Failed"; "40" = "Download Completed"; "48" = "Pending User Session"; "50" = "Enforcement In Progress"; 
    "55" = "Pending Enforcement Retry"; "60" = "Enforcement Failed"; "70" = "Success / Enforcement Completed"
}

# --[ Generate ASCII Table Header ]
$border = "+----------------------+------------------------------------------------------------------+"
$header = "| Office Config XML    | Value                                                            |"
Write-Host $border
Write-Host $header
Write-Host $border

foreach ($key in $OfficeCSP) {
    $OfficeState = Get-ItemProperty -Path $key.PSPath -Name FinalStatus -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FinalStatus
    $OfficeConfig = Get-ItemProperty -Path $key.PSPath -Name "(default)" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "(default)"
    
    $OfficeStateText = if ($OfficeState) { $OfficeStatus[$OfficeState.ToString()] } else { "Unknown" }
    
    Write-Host ("| {0,-20} | {1,-64} |" -f "Office Status", $OfficeStateText)
    
    if ($OfficeConfig) {
        [xml]$ConfigXml = $OfficeConfig
        
        # --[ Extract Values from XML ]
        $channel = $ConfigXml.Configuration.Add.Channel
        $officeEdition = $ConfigXml.Configuration.Add.OfficeClientEdition
        $migrateArch = $ConfigXml.Configuration.Add.MigrateArch
        $productID = $ConfigXml.Configuration.Add.Product.ID
        $excludedApps = ($ConfigXml.Configuration.Add.Product.ExcludeApp | ForEach-Object { $_.ID }) -join ", "
        $languages = ($ConfigXml.Configuration.Add.Product.Language | ForEach-Object { $_.ID }) -join ", "
        $displayLevel = $ConfigXml.Configuration.Display.Level
        $acceptEULA = $ConfigXml.Configuration.Display.AcceptEULA
        $appSettings = ($ConfigXml.Configuration.AppSettings.User.Id) -join ", "
        $removeMSI = [bool]$ConfigXml.Configuration.RemoveMSI
        $sharedComputerLicensing = $ConfigXml.Configuration.Property | Where-Object { $_.Name -eq "SharedComputerLicensing" } | Select-Object -ExpandProperty Value
        $lastExecuteTime = $ConfigXml.Configuration.LastExecuteTime.'#text'
        
        Write-Host ("| {0,-20} | {1,-64} |" -f "Channel", $channel)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Office Edition", $officeEdition)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Migrate Arch", $migrateArch)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Product ID", $productID)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Excluded Apps", $excludedApps)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Languages", $languages)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Display Level", $displayLevel)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Accept EULA", $acceptEULA)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Remove MSI", $removeMSI)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Shared Licensing", $sharedComputerLicensing)
        Write-Host ("| {0,-20} | {1,-64} |" -f "App Settings", $appSettings)
        Write-Host ("| {0,-20} | {1,-64} |" -f "Last Execute Time", $lastExecuteTime)
        Write-Host $border
    }
    
  }

}

else { return "No M365 Apps deployed "}
}

# --[ Function to parse policies from registry ]
function Get-PoliciesData{
    $results = @()

    # --[ Gather ID data from registry ]
    $EnrollmentID = (Get-ChildItem -Path "C:\ProgramData\Microsoft\DMClient\*").Name
    $UserSID = (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$EnrollmentID" | Select -Last 1).PSChildName
    $regPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager"

    $PolicyPaths = @( "$regPath\providers\$EnrollmentID\default\Device\*")
        $PolicyPaths | % {
                $regKeys = Get-ChildItem -Path $_
                    foreach($reg in $regKeys.Name) {
            
                        $policyProcessed  = getRegTime ($reg -replace "HKEY_LOCAL_MACHINE","HKLM:")
                        $policyName = [regex]::Match($reg, '[^\\]+$').Value 
                     
                            $results += [PSCustomObject]@{
                                    Timestamp = [datetime]$policyProcessed
                                    Name = $policyName
                                    Status = "Processed"
                                    Property = "Policy"
                                }
                }
        }
        return $results
}

# --[ Define registry base keys to scan ]
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics\Expected*",
    "HKLM:\SOFTWARE\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics\S-*\Expected*"
)

$sidecarPath = "HKLM:\SOFTWARE\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics\Sidecar"

# --[ Define status mappings ]
$statusMapping = @{
    "1" = "Not Installed"
    "2" = "In Progress"
    "3" = "Completed"
    "4" = "Error"
}

$espStatus = @{
    "0" = "Not Processed"
    "1" = "Processed"
    "60" = "Failed"
    "70" = "Completed"
}

# --[ Function to parse Sidecar registry data ]
function Get-AppsData {
    $results = @()
    if (Test-Path $sidecarPath) {
        Get-ChildItem -Path $sidecarPath -Recurse | Where-Object { $_.PSChildName -ne "LastLoggedState" } | ForEach-Object {
            $timestamp = [datetime]::Parse($_.PSChildName)
            Get-ItemProperty -Path $_.PsPath | ForEach-Object {
                $_.PSObject.Properties | Where-Object { $_.Name -match "Win32App_" -and $_.Name -notmatch "PS(Path|ParentPath|ChildName|Provider)" } | ForEach-Object {
                    $guid = $_.Name -replace "^.*Win32App_", "" -replace "_\d+$", ""
                    $results += [PSCustomObject]@{
                        Timestamp = $timestamp
                        Name = Get-AppName $guid
                        Status = $statusMapping[[string]$_.Value]
                        Property = "Win32App"
                    }
                }
            }
        }
    }
    return $results
}

# --[ Function to parse ESP registry data ]
function Get-ESPStatusData {
    param(
        [string]$basePath
    )
    
    $results = @()
    Get-ChildItem -Path $basePath -Recurse | ForEach-Object {
        $timestamp = [datetime]::Parse($_.PSChildName)
        Get-ItemProperty -Path $_.PsPath | ForEach-Object {
            $_.PSObject.Properties | Where-Object { $_.Name -notmatch "PS(Path|ParentPath|ChildName|Provider)" } | ForEach-Object {
                $propertyName = $_.Name
                $propertyValue = $_.Value
                
                $name = switch -Wildcard ($propertyName) {
                    "./Device/Vendor/MSFT/EnterpriseDesktopAppManagement/MSI/" { [URI]::UnescapeDataString(($propertyName.Split("/"))[6]);$prop = "track this unknown"} 
                   
                    "./User/Vendor/MSFT/EnterpriseModernAppManagement/AppManagement/*" { [URI]::UnescapeDataString(($propertyName.Split("/"))[7]); $prop = "USER UWP" }
                    "./Device/Vendor/MSFT/EnterpriseModernAppManagement/AppManagement/*" {[URI]::UnescapeDataString(($propertyName.Split("/"))[7]); $prop ="DEVICE UWP" }
                    "./User/Vendor/MSFT/ClientCertificateInstall/*" {[URI]::UnescapeDataString(($propertyName.Split("/"))[6]) ; $prop = "User Certificate"}
                    "./Device/Vendor/MSFT/ClientCertificateInstall/*" {[URI]::UnescapeDataString(($propertyName.Split("/"))[6]) ; $prop = "Device Certificate"}
                    "./Vendor/MSFT/Office/Installation*" { "Office Installation" ; $prop = "M365 APPS"}
                    "./Device/Vendor/MSFT/WiFi/Profile/*" { $propertyName -replace "^./Device/Vendor/MSFT/WiFi/Profile/", ""; $prop = "WiFi Profile" }
                    "./Vendor/MSFT/DMClient/Provider/MS%20DM%20Server/EntDMID" `
                    {$p++; if($p -ge 1){$propertyName -replace "./Vendor/MSFT/DMClient/Provider/MS%20DM%20Server/EntDMID","Device Phase Completed";$prop = "ESP Phase"}else{$propertyName -replace "./Vendor/MSFT/DMClient/Provider/MS%20DM%20Server/EntDMID","Account Phase Completed";$prop = "ESP Phase"}}
                    Default { $propertyName }
                }

                $results += [PSCustomObject]@{
                    Timestamp = $timestamp
                    Name = $name
                    Status = $espStatus[[string]$propertyValue]
                    Property = $prop
                }
            }
        }
    }
    return $results
}

# --[ Search logs for application names ]
$logContent = Get-Content -Path "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AppW*.log" -Raw
function Get-AppName {
    param([string]$guid)
    $pattern = '"Id":"' + [regex]::Escape($guid) + '","Name":"(.*?)"'
    if ($logContent -match $pattern) {
        return $matches[1]
    }
    return "Unknown"
}


# --[ Gather Autopilot Profile data ]
$autopilotResults = Get-AutopilotStart

if ($autopilotResults -ne $false){

        # --[ Gathet First Events data ]
        $firsteventsResults = Get-FirstEvents

        # --[ Gather Policies registry data ]
        $policyResults = Get-PoliciesData

        # --[ Gather Sidecar registry data ]
        $sidecarResults = Get-SidecarData

        # --[ Gather Apps registry data ]
        $appsResults = Get-AppsData

        # --[ Gather ESP status data ]
        $espResults = @()
        foreach ($path in $registryPaths) {
            $espResults += Get-ESPStatusData -basePath $path
        }

        # --[ Output results sorted by timestamp ]
        $finalResults = @()
        $finalResults += $autopilotResults
        $finalResults += $firsteventsResults
        $finalResults += $appsResults
        $finalResults += $policyResults
        $finalResults += $sidecarResults
        $finalResults += $espResults


        $finalResults = $finalResults | Sort-Object Timestamp

        $finalResults | Format-Table Timestamp, Name, Status, Property -AutoSize
        "`n"

        # --[ Calculate the Total Autopilot time ]
        $timeDifference = $finalResults[-1].Timestamp - $finalResults[0].Timestamp
        
        $AutopilotDeploymentTime = 
        $timeDifference | % { "$($_.Hours)h $($_.Minutes)m $($_.Seconds)s" }

        Write-Host "Autopilot Deployment Time: $AutopilotDeploymentTime`n`n"

        # --[ Output Autopilot JSON data ]
        Get-AutopilotJSONdata
        "`n"

        # --[ Output M365 Apps data ]
        Get-OfficeInstallData
    }
else 
    {
        Write-Host "This is not an Autopilot device."
    }

# --[ End of script ]