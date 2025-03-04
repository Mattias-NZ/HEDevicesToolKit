<#
  .SYNOPSIS
  Provides extended visibility of Hubitat Elevation devices spread across multiple hubs on one 
  pane of glass. Devices shared between hubs with Hub Mesh can be displayed with their source/
  remote relationships, complete with "In Use By" information for each device, and hyperlinks.
  Output is available to screen, HTML or CSV.

  .DESCRIPTION
  HEDevicesToolKit.ps1 brings together devices from multiple Hubitat Elevation hubs to provide 
  visibility of all these devices from one source.
  The script uses web endpoints to query each hub and from that builds a database of devices 
  and hubs. The database can then be used to:
  * Provide a list of all hubs.
  * Provide a list of all devices with their parent/child relationships and "In Use By" apps.
  * Provide a list of all Hub Mesh devices and their source/remote device relationships and "In 
    Use By" apps.
  * Checking all devices for potential issues.
  * Providing search capabilities.

  The script can be run either interactively through a text based menu system, or non-
  interactively through the use of parameters. All functionality except for changing 
  configuration settings is available non-interactively. When running non-interactively, it's 
  using the same configuration settings as when running interactively.

  The output can be displayed on the screen and/or be sent to an HTML file and/or be sent to a
  CSV file. HTML is recommended for its usability, whereas CSV is recommended for the most 
  complete dataset with the possibility of further analysis using external tools, such as 
  Microsoft Excel.

  ******** Issues Detection ********
  HEDevicesToolKit can identify the following issues:
  * Low battery charge - Checks all devices with the attribute 'battery' and reports on 
    devices that have a charge lower than the set threshold (user configurable).
  * Inactive devices - Checks to see if the last activity recorded for a device is older than
    than the set threshold (user configurable).
  * Offline devices - Checks for devices that are reported as being offline.
  * Hub Mesh - Orphaned remote devices - Checks all Hub Mesh devices to identify any remote 
    devices that have been orphaned by the source device having been removed.
  * Hub Mesh - Hub Mesh disabled on source device - Checks all Hub Mesh devices for source
    devices that have been disabled but still have remote devices on other hubs.
  * Hub Mesh - No remote devices - Checks all Hub Mesh devices to identify any source devices
    that do not have any remote devices.
  
  ******** Files ********
  The following files are by default located in the same folder as HEDevicesToolKit.ps1:
  * HEDevicesToolKit.ps1 - This file.
  * Hubs.txt - Can be used to provide HEDevicesTooLkit.ps1 with a list of Hubitat Elevation hub
    IPv4 addresses, one address per line.
  * Data.json - Contains configuration settings and data from the previous scan.
  HTML files (Only produced when output device is set to HTML).
  * DeviceList.html - Contains the output from the last run of "List all devices" or "Search 
    for devices". Configurable name and path.
  * HubList.html - Contains the output from the last run of "List all hubs". Configurable name 
    and path.
  * HubMeshDeviceList.html - Contains the output from the last run of "List all Hub Mesh 
    devices" or "Search for Hub Mesh devices". Configurable name and path.
  * IssuesList.html - Contains the output from the last run of "List all hubs". Configurable
    name and path.
  CSV files (Only produced when output device is set to CSV).
  * DeviceList.csv - Contains the output from the last run of "List all devices" or "Search 
    for devices". Configurable name and path.
  * HubList.csv - Contains the output from the last run of "List all hubs". Configurable name 
    and path.
  * HubMeshDeviceList.csv - Contains the output from the last run of "List all Hub Mesh 
    devices" or "Search for Hub Mesh devices". Configurable name and path.
  * IssuesList.csv - Contains the output from the last run of "List all hubs". Configurable
    name and path.

  ******** Configuration Settings ********
  Configuration settings can only be accessed in interactive mode, but is used for both 
  interactive and non-interactive modes. The following settings are available:
  * autoLaunchWebBrowser - When True and outputDevice is HTML, the web browser will be 
    launched automatically every time output has been produced to display the output. The
    default is False.
  * Change file names or paths - Change the filename and file path of the output files to a
    different setting from their defaults.
  * dateTimeFormat - The format used with timestamps. The default is 'MMM dd, yyyy @ HH:mm' 
    which produces a timestamp like 'Feb 22, 2025 @ 13:30'.
  * Devices Excluded From Issues Reporting - Devices can be added to this list in order to be 
    excluded from the device issues reporting. This can be beneficial for devices that are
    expected to be in a certain state that would normally trigger a device warning, for 
    example a home battery system that has a battery charge lower than the 
    lowBatteryChargeThreshold. For this device, this would be normal and thus the device can
    be excluded from the issue reporting. Devices are excluded per issue category, meaning
    that in the example above, the home battery system wouldn't show up in any issues 
    reporting regarding low battery charge, but it would not be excluded from any of the 
    other checks.
  * inactivityThresholdMinutes - The threshold in minutes after which a device will be 
    considered being inactive. The default is 1440 minutes (1 day).
  * internetProtocol - The Internet protocol used when accessing the Hubitat Elevation hubs.
    The default is "https://", but can be changed to "http://"
  * lowBatteryChargeThreshold - The threshold in % below which a device will be reported as 
    having a low battery charge. The default is 20%.
  * outputDevice - Which output will be used for the output from HEDevicesToolKit.ps1. Any 
    combination of output devices can be used. Available options are 'screen', 'HTML' and 'CSV'.
    The default is 'screen'.
  * sendWebCallOnIssue - This setting is used to configure web call URL(s) for the different
    categories of issue that can be detected. This could for example be a Hubitat Elevation 
    Maker API endpoint for a virtual switch or a Rule Machine endpoint. 
    For the web call to work, sendWebCallOnIssue must be enabled, as well as each of the sub 
    sendWebCallOnIssue that you want the web call to take place on. A generic URL can be 
    provided which will be used for any enabled sendWebCallOnIssue categories that haven't got
    a specific URL defined. When an issue with a device is detected during a check for issues 
    scan, the specified web call URL will be invoked.
  * textColour - Which foreground colour will be used for HEDevicesToolKit.ps1. It takes any of 
    the standard PowerShell colours (Black DarkBlue DarkGreen DarkCyan DarkRed DarkMagenta 
    DarkYellow Gray DarkGray Blue Green Cyan Red Magenta Yellow White). The default is 'White'.

    
  .PARAMETER NonInteractive
  Is required for running HEDevicesToolKit.ps1 in non-interactive mode. None of the other 
  parameters will work without including this one.

  .PARAMETER RunNewScan
  Used together with -NonInteractive. Starts a new scan of devices. If the parameter is not 
  set, HEDevicesToolKit.ps1 will use data from the last time the scan was performed. 
  When running the scan, HEDevicesToolKit.ps1 will firstly look for Hubitat Elevation hub IP 
  addresses provided with the optional -HubIPAddress <String[]> parameter. If no IP addresses 
  are provided, it will check hubs.txt in the script folder for a list of IP addresses to scan. 
  If that also fails, the script will give the option of entering interactive mode for more 
  scan options.

  .PARAMETER HubIPAddress
  Used together with -NonInteractive and -RunNewScan. Expects a comma separated list of IP 
  addresses to the Hubitat Elevation hubs to be scanned.
  
  .PARAMETER ListAllDevices
  Used together with -NonInteractive. Provides a list of all devices loaded to the configured 
  output device(s). Parent/child relationship between devices is clearly and intuitively shown 
  as well as In Use By apps and other status information.
  
  .PARAMETER ListAllHubs
  Used together with -NonInteractive. Provides a list of all hubs loaded to the configured 
  output device(s).
  
  .PARAMETER ListAllHubMeshDevices
  Used together with -NonInteractive. Provides a list of all Hub Mesh devices loaded to the 
  configured output device(s). Source/remote device relationship between devices is clearly 
  and intuitively shown as well as In Use By apps and other status information.
  
  .PARAMETER CheckForDeviceIssues
  Used together with -NonInteractive. Checks all loaded devices for a range of possible issues 
  and provides a list of these to the configured output device(s). Provides an explaination of 
  each category of issues and possible ways of rectifying the issue.
  
  .PARAMETER SearchForDeviceByName
  Used together with -NonInteractive. Will perform a search of devices by name for the search 
  term provided in the optional -SearchTerm <String> parameter and provide a list of search 
  results to the configured output device(s). The search performs a wildcard search of the 
  search term, i.e. *<search term>*. For example, a search for "wit" will find devices 
  containing "switch" and "with" in their names.
  
  .PARAMETER SearchForHubMeshDeviceByName
  Used together with -NonInteractive. Will perform a search of Hub Mesh devices by name for 
  the search term provided in the optional -SearchTerm <String> parameter and provide a list 
  of search results to the configured output device(s). The search performs a wildcard search 
  of the search term, i.e. *<search term>*. For example, a search for "wit" will find devices 
  containing "switch" and "with" in their names.
  
  .PARAMETER SearchTerm
  Used together with -NonInteractive and -SearchForDeviceByName and/or 
  -SearchForHubMeshDeviceByName. Expects a string of the search term to be used for searching. 
  If the parameter is omitted, HEDevicesToolKit.ps1 will ask for a search term when running.
  
  

  .INPUTS
  None. You can't pipe objects to HEDevicesToolKit.ps1.

  .OUTPUTS
  Optional. HEDevicesToolKit.ps1 outputs to screen and/or HTML file and/or CSV file, depending 
  on what has been configured.

  .EXAMPLE
  PS> .\HEDevicesToolKit.ps1

  Starts HEDevicesToolKit.ps1 in interactive mode.

  .EXAMPLE
  PS> .\HEDevicesToolKit.ps1 -NonInteractive -RunNewScan

  Starts a new non-interactive scan using the Hubitat Elevation hub IP addresses found in the
  hubs.txt file. If there are no valid IP addresses in the file, HEDevicesToolKit.ps1 will 
  prompt for IP addresses to scan.

  .EXAMPLE
  PS> .\HEDevicesToolKit.ps1 -NonInteractive -RunNewScan -HubIPAddress "192.168.1.1","192.168.1.88"

  Starts a new non-interactive scan using the Hubitat Elevation hub IP addresses of 
  "192.168.1.1" and "192.168.1.88"

  .EXAMPLE
  PS> .\HEDevicesToolKit.ps1 -NonInteractive -RunNewScan -HubIPAddress "192.168.1.1","192.168.1.88" 
       -CheckForDeviceIssues -ListAllHubMeshDevices -ListAllDevices

  Starts a new non-interactive scan using the Hubitat Elevation hub IP addresses of "192.168.1.1" 
  and "192.168.1.88" and provides a list of potential issues found, a list of all Hub Mesh 
  devices and a list of all devices

  .EXAMPLE
  PS> .\HEDevicesToolKit.ps1 -NonInteractive -CheckForDeviceIssues -ListAllHubMeshDevices 
       -ListAllDevices

  Uses data stored on disk from a previous scan and provides a list of potential issues found, 
  a list of all Hub Mesh devices and a list of all devices

  .EXAMPLE
  PS> .\HEDevicesToolKit.ps1 -NonInteractive -SearchForDeviceByName -SearchTerm "Contact Sensor"

  Uses data stored on disk from a previous scan and performs a search for devices with 
  "Contact Sensor" in their name

  .NOTES
  Author: Mattias Salomonsson
  License type: GNU AGPL v3.0

  System requirements:
  PowerShell 5.1 or newer
  Hubitat Elevation platform version 2.3.9 or newer
  
  ---Release Notes---
  2025.03.04.1826
  First public version

  

  ----------------------------

#>


param (
    [switch]$NonInteractive,
    [switch]$RunNewScan,
    [string[]]$HubIPAddress, #Comma delimited string of IP addresses
    [switch]$ListAllDevices,
    [switch]$ListAllHubs,
    [switch]$ListAllHubMeshDevices,
    [switch]$CheckForDeviceIssues,
    [switch]$SearchForDeviceByName,
    [switch]$SearchForHubMeshDeviceByName,
    [string]$SearchTerm
)
$Version = "2025.03.04.1826"

function ScanForData { #Takes an array of hub IP addresses as the input and queries them and their devices for data. Wipes existing data
    param (
        [string[]]$HubIPAddressList
    )
    
    $AddressCount = $HubIPAddressList.Count
    if ($AddressCount -ge 1) {
        Write-Host
        
        $i = 0
        $ValidHEAddressFound = $false
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = [SSLHandler]::GetSSLHandler() #Ignore all SSL cert issues
        foreach ($HubIPAddress in $HubIPAddressList) {
            $i++
            $ValidatedAddress = ValidateHubIPAddress -IPaddress $HubIPAddress
            if ($ValidatedAddress) {
                Write-Host ("Getting hub details from IP address '{0}' (address {1} of {2})" -f (($ValidatedAddress),($i),($AddressCount)))
                if (-NOT $ValidHEAddressFound) { #This is the first address found that appears to be a valid HE hub address. Let's destroy previously loaded data now
                    $script:DataHashTable.hubs = [ordered]@{}
                    $script:DataHashTable.devices = [ordered]@{}
                    $script:DataHashTable.apps = [ordered]@{}
                    $script:DataHashTable.hubMeshSourceDevices = @{}
                    $script:DataHashTable.config.deviceListLastUpdated = ""
                }
                $ValidHEAddressFound = $true
                $HubDetails = (Invoke-WebRequest -Uri ($script:DataHashTable.config.internetProtocol + $ValidatedAddress + "/hub/details/json")).Content | ConvertFrom-Json
                
                $script:DataHashTable.hubs.Add($ValidatedAddress,[ordered]@{    hubName = $HubDetails.hubName
                                                                            hubIPAddress = $ValidatedAddress
                                                                            hubURL = $ValidatedAddress
                                                                            platformVersion = $HubDetails.platformVersion
                                                                            hardwareVersion = $HubDetails.hardwareVersion
                })

                Write-Host ("{0} found. Getting device data" -f ($HubDetails.hubName))
                $URL = $script:DataHashTable.config.internetProtocol + $ValidatedAddress + "/hub2/devicesList"
                ScanDevice ((ConvertFrom-Json (Invoke-WebRequest -Uri $URL).Content).devices) $ValidatedAddress
                Write-Host 
                $script:DataHashTable.config.deviceListLastUpdated = (Get-Date).DateTime
                SortDataHashTableByName
                WriteDataToDisk
                Write-Host "Done"
            } else {
                Write-Host ("Getting hub details from IP address '{0}' (address {1} of {2})" -f (($HubIPAddress),($i),($AddressCount)))
                Write-Host "Invalid IP address or no HE hub detected on address. Skipping." -ForegroundColor "Red"
            }
            Write-Host
            
        }
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null #Stop ignoring SSL cert issues
        if (-NOT $ValidHEAddressFound) { #No valid HE hub address was found, so no new devices were ever scanned, and the old data was never removed.
            Write-Host
            Write-Host "No valid Hubitat Elevation hub IP addresses were found. No new scan has taken place." -ForegroundColor "Red"
            if ($script:DataHashTable.config.deviceListLastUpdated) {
                Write-Host "The data that was already present before attempting to run the scan is intact."
            }
            if (-not $NonInteractive) {
                Write-Host
            }
        }

    }
}

function ScanDevice { #Called from ScanForData to query each found device for data
    param (
        $Devices,
        $HubIPaddress
    )

    foreach ($Device in $Devices) {
        Write-Host "." -NoNewline
        $VerboseDeviceDetails = ((Invoke-WebRequest -Uri ($script:DataHashTable.config.internetProtocol + $HubIPaddress + "/device/fullJson/" + $Device.data.id)).Content).ToLower() | ConvertFrom-Json
        $DeviceId = $HubIPaddress + "-" + $Device.data.id
        $DeviceURL = $HubIPaddress + "/device/edit/" + $Device.data.id
        $DeviceStatus = "Enabled"
        $WirelessProtocol = $null
        $InUseBy = $null
        $linked = $false
        $sourceDeviceURL = $null
        $sourceDeviceID = $null
        $Children = $null
        $ParentDeviceID = $null
        $ParentAppName = $null
        
        if ($VerboseDeviceDetails.device.zigbee) {
            $WirelessProtocol = "Zigbee"
        } elseif ($VerboseDeviceDetails.device.ZWave) {
            $WirelessProtocol = "ZWave"
        } elseif ($VerboseDeviceDetails.device.matter) {
            $WirelessProtocol = "Matter"
        }

        if ($Device.data.disabled -eq $true) {
            $DeviceStatus = "Disabled"
        }

        if ($VerboseDeviceDetails.appsusingcount -ge 1) {
            $InUseBy = foreach ($AppUsing in $VerboseDeviceDetails.appsusing){
                $AppID = $HubIPaddress + "-" + $AppUsing.id
                $AppName = (Get-Culture).TextInfo.ToTitleCase($AppUsing.name)
                $AppLabel = (Get-Culture).TextInfo.ToTitleCase($AppUsing.truelabel)
                $AppStatus = "Enabled"
                $AppURL = $HubIPaddress + "/installedapp/configure/" + $AppUsing.id
                
                if ($AppLabel.indexOf(" <Span") -gt 1) { #Remove HTML SPAN tags which are added to the end of stopped or paused app names
                    $AppLabel = $AppLabel.Substring(0,$AppLabel.indexOf(" <Span"))
                }

                if ($AppUsing.disabled -eq $true) {
                    $AppStatus = "Disabled"
                }
                if (-NOT $script:DataHashTable.apps.$AppID) {
                    $script:DataHashTable.apps.Add($AppID,[ordered]@{   appName=$AppName
                                                                        appLabel=$AppLabel
                                                                        appStatus=$AppStatus
                                                                        appURL=$AppURL
                                                                        hubIPAddress=$HubIPaddress
                    })
                }
                $AppID
            }
        }
        
        if ($Device.data.source -like "Linked") { #Indicates a remote Hub Mesh device
            $linked = $true
            $sourceDeviceURL = $VerboseDeviceDetails.device.remoteDeviceUrl
            $sourceDeviceURL = $sourceDeviceURL.substring(($sourceDeviceURL.indexof("://")+3)) #Get rid of the internet protocol from the string
            $sourceDeviceID = ($sourceDeviceURL.split("/",4))[0] + "-" + ($sourceDeviceURL.split("/"))[-1]
            $script:DataHashTable.hubMeshSourceDevices[$sourceDeviceID] += @($DeviceId)
        } 
        if ($VerboseDeviceDetails.device.meshenabled) { #Indicates a source Hub Mesh device
            $script:DataHashTable.hubMeshSourceDevices[$DeviceId] += @()
        }

        if ($Device.children.count -ge 1) {
            foreach ($Child in $Device.children) {ScanDevice $Child $HubIPaddress}
            $Children = foreach ($ChildID in $Device.children.data.id){$HubIPaddress + "-" + $ChildID}
        }

        if ($Device.child) {
            $ParentDeviceID = $HubIPaddress + "-" + $VerboseDeviceDetails.device.parentdeviceid
        }

        $script:DataHashTable.devices.Add($DeviceId, [ordered]@{    deviceName=$Device.data.name
                                                                    deviceURL=$DeviceURL
                                                                    deviceStatus=$DeviceStatus
                                                                    lastActivity = $VerboseDeviceDetails.device.lastActivityTime
                                                                    wirelessProtocol=$WirelessProtocol
                                                                    batteryCharge=$VerboseDeviceDetails.device.currentStates.battery.value
                                                                    hubIPAddress=$HubIPaddress
                                                                    inUseBy=$InUseBy
                                                                    linked=$linked
                                                                    hubMeshEnabled=$VerboseDeviceDetails.device.meshenabled
                                                                    sourceDeviceURL=$sourceDeviceURL
                                                                    sourceDeviceID=$sourceDeviceID
                                                                    parent=$Device.parent
                                                                    child=$Device.child
                                                                    children=$Children
                                                                    parentDeviceID=$ParentDeviceID
                                                                    parentAppName=$ParentAppName
        })
    }
}

function ValidateHubIPAddress { #Validates provided IP address. Checks that the address is valid and that there is a responding HE hub behind it. Returns a sanitised IP address if valid or false otherwise
    param (
        [string]$IPaddress
    )
    $IPaddress = $IPaddress.trim()
    $TestResult = $true
    $Octets = $IPaddress -split "\."
    if ($Octets.count -ne 4) { #Make sure there are only 4 octets
        $TestResult = $false
    } else { #Yes, there are 4 octets. Now check that each octet is a number between 0 and 255
        $SanitisedOctets = @()
        foreach ($Octet in $Octets) {
            if (($Octet -match '^[0-9]+$' -AND [int]$Octet -le 255 -AND [int]$Octet -ge 0) -eq $false){
                $TestResult = $false
            } else {
                $SanitisedOctets += [int]$Octet
            }
        }
        if ($TestResult) {
            $IPaddress = $SanitisedOctets -join "."
        }
    }

    if ($TestResult) { #The IP address is valid, now check if an HE hub can be connected to on well known HE ports
        $Ports = 8081,39501
        foreach ($Port in $Ports) {
            $ConnectionTest = (New-Object System.Net.Sockets.TcpClient)
            if ($ConnectionTest.ConnectAsync($IPaddress, $Port).Wait(750) -eq $false) {
                $TestResult = $false
            }
            $ConnectionTest.Close()
        }
    }
    if ($TestResult) {
        $IPaddress
    } else {
        $false
    }
}

function WriteListOfDevicesToOutputDevice { #Used to produce a list of devices that is sent to the selected output devices
    param (
        [string[]]$DeviceIDList = $script:DataHashTable.devices.keys,
        [switch]$ShowHierarchy, #If used, only non-child devices are shown with children devices listed underneath their parents. If a deviceID of a child device is provided without the deviceID of its parent being provided as well, the child device will not be included in the output
        [switch]$ShowHubMeshHierarchy #If used, only devices involved with Hub Mesh are included in the output. If the ID of a remote device is provided, then its source device with all the source devices' remote devices will be included in the output.
    )
    if ($ShowHubMeshHierarchy) {
        WriteTopOfPage "List of Hub Mesh devices"
    } else {
        WriteTopOfPage "List of devices"
    }
    
    if ($DeviceIDList.count -ge 1 -AND ($script:DataHashTable.config.outputDevice.screen -OR $script:DataHashTable.config.outputDevice.HTML -OR $script:DataHashTable.config.outputDevice.CSV)) {
        if ($script:DataHashTable.config.outputDevice.HTML) {
            $HTMLFilePath = $script:DataHashTable.config.filePath.DeviceListHTML
            $HTML = SetHTMLHeader "List of devices"
            if ($ShowHubMeshHierarchy) {
                $HTMLFilePath = $script:DataHashTable.config.filePath.HubMeshDeviceListHTML
                $HTML = SetHTMLHeader "List of Hub Mesh devices"    
            }
        }
        if ($script:DataHashTable.config.outputDevice.CSV) {
            $CSVFilePath = $script:DataHashTable.config.filePath.DeviceListCSV
            $CSVHeader = [ordered]@{deviceName="Device Name"
                                    deviceID="Device ID"
                                    HEDeviceID="HE Device ID"
                                    hub="Hub"
                                    deviceURL="Device URL"
                                    isDisabled="Is Disabled?"
                                    isParent="Is Parent?"
                                    children="Child Devices"
                                    isChild="Is Child?"
                                    parentDevice="Parent Device"
                                    lastActivity="Last Activity"
                                    wirelessProtocol="Wireless Protocol"
                                    batteryCharge="Battery Charge (%)"
                                    inUseBy="In Use By"
                                    hubMeshEnabled="Hub Mesh Enabled?"
                                    hubMeshRemoteDevices="Hub Mesh Remote Devices"
                                    hubMeshSourceDevice="Hub Mesh Source Device"             
            }
            if ($ShowHubMeshHierarchy) {
                $CSVFilePath = $script:DataHashTable.config.filePath.HubMeshDeviceListCSV   
            }
            Remove-Item $CSVFilePath -Force -ErrorAction SilentlyContinue
            #Print header
            $CSVContent = ""
            $i = 0
            foreach ($Column in $CSVHeader.keys) {
                $i++
                $CSVContent += '"' + $CSVHeader.$Column.replace('"','""') + '"'
                if ($i -lt $CSVHeader.count) {
                    $CSVContent += ","
                }
            }
            $CSVContent += "`n"
        }
        
        if ($ShowHierarchy) { 
            $DeviceIDList = $DeviceIDList | Where-Object {$script:DataHashTable.devices.$_.child -eq $false}
        } elseif ($ShowHubMeshHierarchy) {
            $DeviceIDList = FindSourceDevice $DeviceIDList
        }

        foreach ($DeviceID in $DeviceIDList) {
            $HTML += "<div id=`"$DeviceID`">"
            $DeviceData = @{}
            if ($script:DataHashTable.config.outputDevice.CSV) {
                $CSV = @{}
                foreach ($key in $CSVHeader.keys) {
                    $CSV.Add($key,"")
                }
                $CSV.deviceID=$DeviceID
                $CSV.HEDeviceID=$DeviceID.Substring($DeviceID.indexOf("-")+1,($DeviceID.length-$DeviceID.indexOf("-")-1))
                $i = 0
                foreach ($RemoteID in $script:DataHashTable.hubMeshSourceDevices.$DeviceID) {
                    $i++
                    $CSV.hubMeshRemoteDevices += $script:DataHashTable.devices.$RemoteID.deviceName + " (" + $RemoteID + " - " + ($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$RemoteID.deviceURL) + ") " + " on " + $script:DataHashTable.hubs.($script:DataHashTable.devices.$RemoteID.hubIPAddress).hubName
                                if ($i -lt ($script:DataHashTable.hubMeshSourceDevices.$DeviceID).count) {
                        $CSV.hubMeshRemoteDevices += "`n"
                    }
                }
            }
            if (-NOT $script:DataHashTable.devices.$DeviceID.deviceURL) { #The device doesn't have a deviceURL which indicates that the device doesn't exist. This happens when a source hubmesh device has been removed but one or more remote devices remain
                $Message = "Source device is missing!"
                if (-NOT ($script:DataHashTable.hubs.($DeviceID.substring(0,$DeviceID.indexOf("-"))).hubName)) { #Unable to resolve the name of the missing device's hub. This points to the hub not having been scanned rather than the source device missing
                    $Message = "Source device is missing, because the hub it's on wasn't included in the device scan!"
                }
                $HTML += GenerateOutput @{  settings=@{ useFieldDescriptions=$true
                                                        firstLineIsTitle=$true}
                                            one=@{      one=@{  displayValue=$Message
                                                                warning=$true}}}
                if ($script:DataHashTable.config.outputDevice.CSV) {
                    $CSV.deviceName="Device is missing!"
                    $CSV.hub=$script:DataHashTable.hubs.($DeviceID.substring(0,$DeviceID.indexOf("-"))).hubName
                    if (-NOT $CSV.hub) { #Unable to resolve hub name from the IP address (HE hub was probably not included in the device scan so is not part of the dataset)
                        $CSV.hub=$DeviceID.substring(0,$DeviceID.indexOf("-")) + " - not able to resolve hub name. Was the hub IP address included when doing the device scan?"
                    }
                }
            } else {
                if ($script:DataHashTable.config.outputDevice.CSV) {
                    $CSV.deviceName=$script:DataHashTable.devices.$DeviceID.deviceName
                    $CSV.deviceURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$DeviceID.deviceURL)
                    $CSV.hub=$script:DataHashTable.hubs.($script:DataHashTable.devices.$DeviceID.hubIPAddress).hubName
                    $CSV.hubMeshEnabled=($script:DataHashTable.devices.$DeviceID.hubMeshEnabled).ToString()
                    $CSV.isParent=($script:DataHashTable.devices.$DeviceID.parent).ToString()
                    $CSV.isChild=($script:DataHashTable.devices.$DeviceID.child).ToString()
                    if ($script:DataHashTable.devices.$DeviceID.lastActivity) {
                        $CSV.lastActivity=$script:DataHashTable.devices.$DeviceID.lastActivity
                    }
                    if ($script:DataHashTable.devices.$DeviceID.wirelessProtocol) {
                        $CSV.wirelessProtocol=$script:DataHashTable.devices.$DeviceID.wirelessProtocol
                    }
                    if ($script:DataHashTable.devices.$DeviceID.batteryCharge) {
                        $CSV.batteryCharge=$script:DataHashTable.devices.$DeviceID.batteryCharge
                    }
                    if ($script:DataHashTable.devices.$DeviceID.parentDeviceID) {
                        $CSV.parentDevice=$script:DataHashTable.devices.($script:DataHashTable.devices.$DeviceID.parentDeviceID).deviceName + " (" + $script:DataHashTable.devices.$DeviceID.parentDeviceID + ")"
                    }
                    
                }

                $Disabled = $false
                if ($script:DataHashTable.devices.$DeviceID.deviceStatus -eq "Disabled") {
                    $Disabled = $true
                }
                if ($script:DataHashTable.config.outputDevice.CSV) {$CSV.isDisabled=$Disabled.ToString()}
                    
                $Linked = $false
                if ($script:DataHashTable.devices.$DeviceID.linked -eq $true){
                    if ($ShowHubMeshHierarchy -eq $false) {
                        $Linked = "#$($script:DataHashTable.devices.$DeviceID.sourceDeviceID)"
                    }
                    if ($script:DataHashTable.config.outputDevice.CSV) {
                        if ($script:DataHashTable.devices.($script:DataHashTable.devices.$DeviceID.sourceDeviceID).deviceName){ #If the source device has a name, i.e. it exists
                            $CSV.hubMeshSourceDevice=$script:DataHashTable.devices.($script:DataHashTable.devices.$DeviceID.sourceDeviceID).deviceName + " (" + $script:DataHashTable.devices.$DeviceID.sourceDeviceID + ")"
                        } else { #Source device doesn't have a name, so it doesn't exist
                            $CSV.hubMeshSourceDevice="Source device $($script:DataHashTable.devices.$DeviceID.sourceDeviceID) is missing"
                        }
                    }
                }
                
                $SecondDisplayName = "Hub"
                if ($ShowHubMeshHierarchy) {
                    $SecondDisplayName = "Source hub"
                }

                if (($script:DataHashTable.devices.$DeviceID.children).count -ge 1) {
                    $Children = @{}
                    $i=0
                    foreach ($ChildDeviceID in $script:DataHashTable.devices.$DeviceID.children) { 
                        $Children["$ChildDeviceID"] = @{childName=$script:DataHashTable.devices.$ChildDeviceID.deviceName
                                                        childURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$ChildDeviceID.deviceURL)}
                        if ($script:DataHashTable.config.outputDevice.CSV) {
                            $i++
                            $CSV.children+=$script:DataHashTable.devices.$ChildDeviceID.deviceName + " (" + $ChildDeviceID + " - " + ($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$ChildDeviceID.deviceURL) + ")"
                            if ($i -lt ($script:DataHashTable.devices.$DeviceID.children).count) {
                                $CSV.children += "`n"
                            }
                        }
                    }
                }
                if (($script:DataHashTable.devices.$DeviceID.inUseBy).count -ge 1) {
                    $InUseBy = @{}
                    $i=0
                    foreach ($AppID in $script:DataHashTable.devices.$DeviceID.inUseBy) { 
                        if ($script:DataHashTable.apps.$AppID.appLabel) { #If the app has an AppLabel, use that as it's name, otherwise use the appName
                            $AppName = $script:DataHashTable.apps.$AppID.appLabel
                        } else {
                            $AppName = $script:DataHashTable.apps.$AppID.appName
                        }
                        if ($script:DataHashTable.apps.$AppID.appName) { #If the app has an appName, add it to the display name unless it's the same as the current display name ($AppName)
                            if (-NOT (($script:DataHashTable.apps.$AppID.appName).ToLower() -eq $AppName.ToLower())) { #check if appName and $AppName are different, if so, add appName to $AppName
                                $AppName += " (" + $script:DataHashTable.apps.$AppID.appName + ")"   
                            }
                        }
                        $InUseBy["$AppID"] = @{appName=$AppName
                                            appURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.apps.$AppID.appURL)}
                        if ($script:DataHashTable.config.outputDevice.CSV) {
                            $i++
                            $CSV.inUseBy+=$AppName + " (" + ($script:DataHashTable.config.internetProtocol + $script:DataHashTable.apps.$AppID.appURL) + ")"
                            if ($i -lt ($script:DataHashTable.devices.$DeviceID.inUseBy).count) {
                                $CSV.inUseBy += "`n"
                            }
                        }
                    }
                }
                $DeviceData = @{one=@{  displayValue=$script:DataHashTable.devices.$DeviceID.deviceName
                                        URL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$DeviceID.deviceURL)
                                        disabled=$Disabled
                                        linked=$Linked}
                                two=@{  displayValue=$script:DataHashTable.hubs.($script:DataHashTable.devices.$DeviceID.hubIPAddress).hubName
                                        displayName=$SecondDisplayName}}
                if ($ShowHierarchy -eq $false -AND $ShowHubMeshHierarchy -eq $false) {
                    if ($script:DataHashTable.devices.$DeviceID.child) {
                        $DeviceData += @{three=@{   displayValue=$script:DataHashTable.devices.$($script:DataHashTable.devices.$DeviceID.parentDeviceID).deviceName
                                                    displayName="Parent device"
                                                    URL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$($script:DataHashTable.devices.$DeviceID.parentDeviceID).deviceURL)}}
                    }
                    if (($script:DataHashTable.devices.$DeviceID.children).count -ge 1) {
                        $NextIterationValue=0
                        for ($NextIterationValue;$DeviceData.($script:IterationArray[$NextIterationValue]);$NextIterationValue++){}
                        $DeviceData += @{$($script:IterationArray[$NextIterationValue])=@{  displayValue=$Children
                                                                                            displayName="Child devices"
                                                                                            children=$true}}
                    }    
                }
                if (($script:DataHashTable.devices.$DeviceID.inUseBy).count -ge 1) {
                    $NextIterationValue=0
                    for ($NextIterationValue;$DeviceData.($script:IterationArray[$NextIterationValue]);$NextIterationValue++){}
                    $DeviceData += @{$($script:IterationArray[$NextIterationValue])=@{  displayValue=$InUseBy
                                                                                        displayName="In use by"
                                                                                        inUseBy=$true}}
                }
                $HTML += GenerateOutput @{settings=@{useFieldDescriptions=$true
                                                            firstLineIsTitle=$true}
                                                one=$DeviceData}
            }
            
            if ($script:DataHashTable.config.outputDevice.CSV) {
                $i = 0
                foreach ($Column in $CSVHeader.keys) {
                    $i++
                    $CSVContent += '"' + $CSV.$Column.replace('"','""') + '"'
                    if ($i -lt $CSVHeader.count) {
                        $CSVContent += ","
                    }
                }
                $CSVContent += "`n"
            }
            
            if (($script:DataHashTable.devices.$DeviceID.parent -eq $true -AND $ShowHierarchy -eq $true) -OR $ShowHubMeshHierarchy -eq $true){
                if ($script:DataHashTable.devices.$DeviceID.parent -eq $true -AND $ShowHierarchy -eq $true){
                    $SecondaryDeviceIDList = $script:DataHashTable.devices.$DeviceID.children
                } elseif ($ShowHubMeshHierarchy -eq $true) {
                    $SecondaryDeviceIDList = $script:DataHashTable.hubMeshSourceDevices.$DeviceID
                }
                foreach ($SecondaryDeviceID in $SecondaryDeviceIDList) {
                    if ($script:DataHashTable.config.outputDevice.CSV) {
                        $CSV = @{}
                        foreach ($key in $CSVHeader.keys) {
                            $CSV.Add($key,"")
                        }
                        $CSV.deviceName=$script:DataHashTable.devices.$SecondaryDeviceID.deviceName
                        $CSV.deviceID=$SecondaryDeviceID
                        $CSV.HEDeviceID=$SecondaryDeviceID.Substring($SecondaryDeviceID.indexOf("-")+1,($SecondaryDeviceID.length-$SecondaryDeviceID.indexOf("-")-1))
                        $CSV.hub=$script:DataHashTable.hubs.($script:DataHashTable.devices.$SecondaryDeviceID.hubIPAddress).hubName
                        $CSV.deviceURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$SecondaryDeviceID.deviceURL)
                        $CSV.hubMeshEnabled=($script:DataHashTable.devices.$SecondaryDeviceID.hubMeshEnabled).ToString()
                        $CSV.isParent=($script:DataHashTable.devices.$SecondaryDeviceID.parent).ToString()
                        $CSV.isChild=($script:DataHashTable.devices.$SecondaryDeviceID.child).ToString()
                        if ($script:DataHashTable.devices.$SecondaryDeviceID.lastActivity) {
                            $CSV.lastActivity=$script:DataHashTable.devices.$SecondaryDeviceID.lastActivity
                        }
                        if ($script:DataHashTable.devices.$SecondaryDeviceID.wirelessProtocol) {
                            $CSV.wirelessProtocol=$script:DataHashTable.devices.$SecondaryDeviceID.wirelessProtocol
                        }
                        if ($script:DataHashTable.devices.$SecondaryDeviceID.batteryCharge) {
                            $CSV.batteryCharge=$script:DataHashTable.devices.$SecondaryDeviceID.batteryCharge
                        }
                        if ($script:DataHashTable.devices.$SecondaryDeviceID.parentDeviceID) {
                            $CSV.parentDevice=$script:DataHashTable.devices.($script:DataHashTable.devices.$SecondaryDeviceID.parentDeviceID).deviceName + " (" + $script:DataHashTable.devices.$SecondaryDeviceID.parentDeviceID + ")"
                        }
                        if (($script:DataHashTable.hubMeshSourceDevices.$SecondaryDeviceID).count -ge 1) { #Remote devices exist
                            $i = 0
                            foreach ($RemoteID in $script:DataHashTable.hubMeshSourceDevices.$SecondaryDeviceID) {
                                $i++
                                $CSV.hubMeshRemoteDevices += $script:DataHashTable.devices.$RemoteID.deviceName + " (" + $RemoteID + " - " + ($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$RemoteID.deviceURL) + ") " + " on " + $script:DataHashTable.hubs.($script:DataHashTable.devices.$RemoteID.hubIPAddress).hubName
                                if ($i -lt ($script:DataHashTable.hubMeshSourceDevices.$SecondaryDeviceID).count) {
                                    $CSV.hubMeshRemoteDevices += "`n"
                                }
                            }
                        }
                        if (($script:DataHashTable.devices.$SecondaryDeviceID.children).count -ge 1) { #Children exist
                            $i = 0
                            foreach ($ChildDeviceID in $script:DataHashTable.devices.$SecondaryDeviceID.children) {
                                $i++
                                $CSV.children+=$script:DataHashTable.devices.$ChildDeviceID.deviceName + " (" + $ChildDeviceID + " - " + ($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$ChildDeviceID.deviceURL) + ")"
                                if ($i -lt ($script:DataHashTable.devices.$SecondaryDeviceID.children).count) {
                                    $CSV.children += "`n"
                                }
                            }
                        }
                    }
                    $HTML += "<div id=`"$SecondaryDeviceID`">"
                    $SecondaryDisabled = $false
                    if ($script:DataHashTable.devices.$SecondaryDeviceID.deviceStatus -eq "Disabled") {
                        $SecondaryDisabled = $true
                    }
                    if ($script:DataHashTable.config.outputDevice.CSV) {$CSV.isDisabled=$SecondaryDisabled.ToString()}

                    $SecondaryLinked = $false
                    if ($script:DataHashTable.devices.$SecondaryDeviceID.linked -eq $true) {
                        if ($ShowHubMeshHierarchy -eq $false) {
                            $SecondaryLinked = "#$($script:DataHashTable.devices.$SecondaryDeviceID.sourceDeviceID)"
                        }
                        if ($script:DataHashTable.config.outputDevice.CSV) {
                            if ($script:DataHashTable.devices.($script:DataHashTable.devices.$SecondaryDeviceID.sourceDeviceID).deviceName){ #If the source device has a name, i.e. it exists
                                $CSV.hubMeshSourceDevice=$script:DataHashTable.devices.($script:DataHashTable.devices.$SecondaryDeviceID.sourceDeviceID).deviceName + " (" + $script:DataHashTable.devices.$SecondaryDeviceID.sourceDeviceID + ")"
                            } else { #Source device doesn't have a name, so it doesn't exist
                                $CSV.hubMeshSourceDevice="Source device $($script:DataHashTable.devices.$SecondaryDeviceID.sourceDeviceID) is missing"
                            }
                        }
                    }

                    $SecondaryFirstDisplayName = "Child device"
                    $SecondarySecondDisplayName = "Hub"
                    if ($ShowHubMeshHierarchy) {
                        $SecondaryFirstDisplayName = "Remote device"
                        $SecondarySecondDisplayName = "Remote hub"
                    }
                    if (($script:DataHashTable.devices.$SecondaryDeviceID.inUseBy).count -ge 1) {
                        $i=0
                        $SecondaryInUseBy = @{}
                        foreach ($AppID in $script:DataHashTable.devices.$SecondaryDeviceID.inUseBy) { 
                            if ($script:DataHashTable.apps.$AppID.appLabel) { #If the app has an AppLabel, use that as it's name, otherwise use the appName
                                $AppName = $script:DataHashTable.apps.$AppID.appLabel
                            } else {
                                $AppName = $script:DataHashTable.apps.$AppID.appName
                            }
                            if (-NOT (($script:DataHashTable.apps.$AppID.appName).ToLower() -eq $AppName.ToLower())) { #Only write the appName if the appName and the name already given to the device differ 
                                $AppName += " (" + $script:DataHashTable.apps.$AppID.appName + ")"   
                            }
                            $SecondaryInUseBy["$AppID"] = @{appName=$AppName
                                                        appURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.apps.$AppID.appURL)}
                            if ($script:DataHashTable.config.outputDevice.CSV) {
                                $i++
                                $CSV.inUseBy+=$AppName + " (" + ($script:DataHashTable.config.internetProtocol + $script:DataHashTable.apps.$AppID.appURL) + ")"
                                if ($i -lt ($script:DataHashTable.devices.$SecondaryDeviceID.inUseBy).count) {
                                    $CSV.inUseBy += "`n"
                                }
                            }
                        }
                    }
                    $SecondaryDeviceData = @{one=@{ displayValue=$script:DataHashTable.devices.$SecondaryDeviceID.deviceName
                                                displayName=$SecondaryFirstDisplayName
                                                URL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$SecondaryDeviceID.deviceURL)
                                                disabled=$SecondaryDisabled
                                                linked=$SecondaryLinked}
                                        two=@{  displayValue=$script:DataHashTable.hubs.($script:DataHashTable.devices.$SecondaryDeviceID.hubIPAddress).hubName
                                                displayName=$SecondarySecondDisplayName}}
                    if (($script:DataHashTable.devices.$SecondaryDeviceID.inUseBy).count -ge 1) {
                        $SecondaryDeviceData += @{three=@{  displayValue=$SecondaryInUseBy
                                                        displayName="In use by"
                                                        inUseBy=$true}}
                    }
                    $HTML += GenerateOutput @{settings=@{useFieldDescriptions=$true
                                                                indentationLevel=2}
                                                    one=$SecondaryDeviceData}
                    $HTML += "</div>"
                    
                    if ($script:DataHashTable.config.outputDevice.CSV) {
                        $i = 0
                        foreach ($Column in $CSVHeader.keys) {
                            $i++
                            #Write-Host "$DeviceID - $Column"
                            $CSVContent += '"' + $CSV.$Column.replace('"','""') + '"'
                            if ($i -lt $CSVHeader.count) {
                                $CSVContent += ","
                            }
                        }
                        $CSVContent += "`n"
                    }
                }
            }
            
            if ($script:DataHashTable.config.outputDevice.screen) {Write-Host}
            if ($script:DataHashTable.config.outputDevice.HTML) {$HTML += "</div><div class=`"divider`"></div>`n"}
            if ($script:DataHashTable.config.outputDevice.CSV) {$CSVContent += "`n"}
        }
        if ($script:DataHashTable.config.outputDevice.HTML) {$HTML += SetHTMLFooter;WriteHTMLtoDisk -HTMLcode $HTML -FilePath $HTMLFilePath}
        if ($script:DataHashTable.config.outputDevice.screen) {Write-Host}
        if ($script:DataHashTable.config.outputDevice.CSV) {Add-Content -Path $CSVFilePath -Force -Value $CSVContent}
    } else {
        if ($DeviceIDList.count -lt 1) {
            Write-Host "No devices found to be listed" -ForegroundColor "Red"
        } else {
            Write-Host "No supported output device selected." -ForegroundColor "Red"
        }
    }
}

function WriteHubListToOutputDevice { #Used to produce a list of hubs that is sent to the selected output devices
    WriteTopOfPage "List of hubs"
    
    if ($script:DataHashTable.hubs.count -ge 1 -AND ($script:DataHashTable.config.outputDevice.screen -OR $script:DataHashTable.config.outputDevice.HTML -OR $script:DataHashTable.config.outputDevice.CSV)) {
        if ($script:DataHashTable.config.outputDevice.HTML) {
            $HTML = SetHTMLHeader "List of hubs"
            if ($script:DataHashTable.hubs.count -eq 1){
                $HTML += "<p>Listing 1 hub:</p>"
            } else {
                $HTML += "<p>Listing $($script:DataHashTable.hubs.count) hubs</p>"
            }
        }
        if ($script:DataHashTable.config.outputDevice.screen) {
            if ($script:DataHashTable.hubs.count -eq 1){
                Write-Host "Listing 1 hub"
            } else {
                Write-Host "Listing $($script:DataHashTable.hubs.count) hubs"
            }
            Write-Host
        }
        if ($script:DataHashTable.config.outputDevice.CSV) {
            $CSVFilePath = $script:DataHashTable.config.filePath.HubListCSV
            $CSVHeader = [ordered]@{hubName="Hub Name"
                                    hubID="Hub ID"
                                    hubURL="Hub URL"
                                    hardwareVersion="Hardware Version"
                                    platformVersion="Platform Version"           
            }
            Remove-Item $CSVFilePath -Force -ErrorAction SilentlyContinue
            #Print header
            $CSVContent = ""
            $i = 0
            foreach ($Column in $CSVHeader.keys) {
                $i++
                $CSVContent += '"' + $CSVHeader.$Column.replace('"','""') + '"'
                if ($i -lt $CSVHeader.count) {
                    $CSVContent += ","
                }
            }
            $CSVContent += "`n"
        }

        foreach ($HubID in $script:DataHashTable.hubs.keys) {
            $URL = $script:DataHashTable.config.internetProtocol + $script:DataHashTable.hubs.$HubID.hubURL
            $HTML += "<div class=`"$HubID`">"
            $HTML += GenerateOutput @{settings=@{useFieldDescriptions=$true
                                                        firstLineIsTitle=$true}
                                                one=@{  one=@{displayValue=$script:DataHashTable.hubs.$HubID.hubName
                                                            displayName="Hub"
                                                            URL=$URL}
                                                        two=@{displayValue=$HubID
                                                            displayName="IP address"}
                                                        three=@{displayValue=$script:DataHashTable.hubs.$HubID.hardwareVersion
                                                            displayName="Hardware version"}
                                                        four=@{displayValue=$script:DataHashTable.hubs.$HubID.platformVersion
                                                            displayName="Platform version"}}
                                            }
            if ($script:DataHashTable.config.outputDevice.screen) {Write-Host}
            if ($script:DataHashTable.config.outputDevice.HTML) {$HTML += "</div><div class=`"divider`"></div>"}
            if ($script:DataHashTable.config.outputDevice.CSV) {
                $CSV = @{hubName=$script:DataHashTable.hubs.$HubID.hubName
                         hubID=$HubID
                         hubURL=$URL
                         platformVersion=$script:DataHashTable.hubs.$HubID.platformVersion
                         hardwareVersion=$script:DataHashTable.hubs.$HubID.hardwareVersion}
                $i = 0
                foreach ($Column in $CSVHeader.keys) {
                    $i++
                    $CSVContent += '"' + $CSV.$Column.replace('"','""') + '"'
                    if ($i -lt $CSVHeader.count) {
                        $CSVContent += ","
                    }
                }
                $CSVContent += "`n"
            }
        }
        if ($script:DataHashTable.config.outputDevice.HTML) {$HTML += SetHTMLFooter; WriteHTMLtoDisk -HTMLcode $HTML -FilePath $script:DataHashTable.config.filePath.HubListHTML}
        if ($script:DataHashTable.config.outputDevice.screen) {Write-Host}
        if ($script:DataHashTable.config.outputDevice.CSV) {Add-Content -Path $CSVFilePath -Force -Value $CSVContent}
    } else {
        if ($script:DataHashTable.hubs.count -lt 1) {
            Write-Host "No hubs found to be listed" -ForegroundColor "Red"
        } else {
            Write-Host "No supported output device selected." -ForegroundColor "Red"
        }
    }
}

function WriteDataToDisk { #Writes data and configuration to disk
    $script:DataHashTable | ConvertTo-Json -Depth 10 | Out-File ($script:DataJSONFilePath) -Force
}

function WriteHTMLtoDisk {
    param (
        [string]$HTMLcode,
        [string]$FilePath
    )
    if ($HTMLcode -AND $FilePath) {
        $HTMLcode | Out-File $FilePath -Force
        if ($script:DataHashTable.config.autoLaunchWebBrowser) {
            Start-Process $FilePath
        }
    }
}

function ReadDataFromDisk { #Reads data and configuration data from disk
    $PSCObject = Get-Content ($script:DataJSONFilePath) | Out-String | ConvertFrom-Json
    ReadConfigFromDisk $PSCObject

    foreach ($propertyLevel1 in ($PSCObject.psobject.properties.name | Where-Object {$_ -ne "config"})) { #Foreach propery name, except for 'config'
        $script:DataHashTable.$propertyLevel1 = @{}
        foreach ($propertyLevel2 in $PSCObject.$propertyLevel1.psobject.properties.name){
            $script:DataHashTable.$propertyLevel1.Add($propertyLevel2,$PSCObject.$propertyLevel1.$propertyLevel2)
        }
    }
    
    SortDataHashTableByName
}

function ReadConfigFromDisk { #Reads configuration data from disk
    param (
        $PSCObject = (Get-Content ($script:DataJSONFilePath) | Out-String | ConvertFrom-Json)
    )
    foreach ($propertyLevel2 in $PSCObject.config.psobject.properties.name){
        if ($propertyLevel2 -eq "sendWebCallOnIssue" -OR $propertyLevel2 -eq "filePath") {
            foreach ($propertyLevel3 in $PSCObject.config.$propertyLevel2.psobject.properties.name){
                if ($script:DataHashTable.config.sendWebCallOnIssue.categories -contains $propertyLevel3) {
                    foreach ($propertyLevel4 in $PSCObject.config.$propertyLevel2.$propertyLevel3.psobject.properties.name){
                        $script:DataHashTable.config.$propertyLevel2.$propertyLevel3[$propertyLevel4]=$PSCObject.config.$propertyLevel2.$propertyLevel3.$propertyLevel4
                    }
                } else {
                    $script:DataHashTable.config.$propertyLevel2[$propertyLevel3]=$PSCObject.config.$propertyLevel2.$propertyLevel3
                }
            }
        } else {
            $script:DataHashTable.config[$propertyLevel2]=$PSCObject.config.$propertyLevel2
        }
    }
    if ($script:ValidColours -contains $script:DataHashTable.config.textColour){
        $script:PSDefaultParameterValues['*:ForegroundColor'] = $script:DataHashTable.config.textColour
    } else {
        $script:DataHashTable.config.textColour = $script:DefaultTextColour
    }
}

function ReadHubsTXTListFromDisk { #Reads the hubs.txt file and returns $false or list of valid IP addresses
    $Contents = Get-Content $script:HubListFilePath -ErrorAction SilentlyContinue
    if ($Contents) {
        $ValidIPAddresses = @()
        foreach ($Line in ($Contents | Where-Object {$_.trim().length -ge 1} | Where-Object {$_.trim().substring(0,1) -ne "#"})) { #For each line that doesn't start with "#"
            if ($Line.indexOf("#") -ge 1) { #Check if line contains a comment, it so - cut out the comments
                $Line = $Line.substring(0,$Line.indexOf("#"))
            }
            $Line = $Line.trim()
            $Line = ValidateHubIPAddress $Line
            if ($Line) {
                $ValidIPAddresses += $Line
            }
        }
        if ($ValidIPAddresses.count -ge 1) {
            $ValidIPAddresses
        } else {
            $false
        }
    } else {
        $false
    }
}

function SortDataHashTableByName { #Used to sort the devices and hubs found in $DataHashTable by name 
    $DeviceIDHashTableListSorted = [ordered]@{}
    $ListOfNames = @()
    foreach ($DeviceID in $script:DataHashTable.devices.keys) {
        $ListOfNames += @($script:DataHashTable.devices.$DeviceID.deviceName + "<<<SPLIT>>>" + $DeviceID)
    }
    foreach ($sortedName in ($ListOfNames | Sort-Object)){
        $DeviceIDHashTableListSorted.Add((($sortedName -split "<<<SPLIT>>>")[1]),$script:DataHashTable.devices.(($sortedName -split "<<<SPLIT>>>")[1]))
    }
    $script:DataHashTable.devices = [ordered]@{}
    $script:DataHashTable.devices = $DeviceIDHashTableListSorted
    
    
    $HubIDHashTableListSorted = [ordered]@{}
    $ListOfNames = @()
    foreach ($HubID in $script:DataHashTable.hubs.keys) {
        $ListOfNames += @($script:DataHashTable.hubs.$HubID.hubName + "<<<SPLIT>>>" + $HubID)
    }
    foreach ($sortedName in ($ListOfNames | Sort-Object)){
        $HubIDHashTableListSorted.Add((($sortedName -split "<<<SPLIT>>>")[1]),$script:DataHashTable.hubs.(($sortedName -split "<<<SPLIT>>>")[1]))
    }
    $script:DataHashTable.hubs = [ordered]@{}
    $script:DataHashTable.hubs = $HubIDHashTableListSorted
}

function PauseHEDevicesToolKit { #Overrides the native pause command, just to get the configured text colour on the pause text as well
    Write-Host "Press ENTER to continue" -NoNewline
    Read-Host
}

function EraseConfigurationSettingsAndSetDefaults { #Erases all config settings and repopulates it with default settings
    $deviceListLastUpdated = $script:DataHashTable.config.deviceListLastUpdated
    $script:DataHashTable.config = [ordered]@{}
    $script:DataHashTable.config.Add("autoLaunchWebBrowser",$script:DefaultAutoLaunchWebBrowser)
    $script:DataHashTable.config.Add("dateTimeFormat",$script:DefaultDateTimeFormat)
    $script:DataHashTable.config.Add("inactivityThresholdMinutes",$script:DefaultInactivityThresholdMinutes)
    $script:DataHashTable.config.Add("internetProtocol",$script:DefaultInternetProtocol)
    $script:DataHashTable.config.Add("lowBatteryChargeThreshold",$script:DefaultLowBatteryChargeThreshold)
    $script:DataHashTable.config.Add("outputDevice",$script:DefaultOutputDevice)
    $script:DataHashTable.config.Add("textColour",$script:DefaultTextColour)
    $script:DataHashTable.config.deviceListLastUpdated = $deviceListLastUpdated
    $script:DataHashTable.config.excludedDevicesFromIssuesReporting = [ordered]@{}
    $script:DataHashTable.config.filePath = [ordered]@{}
    $script:DataHashTable.config.filePath.Add("DeviceListHTML",$script:DefaultFilePathDeviceListHTML)
    $script:DataHashTable.config.filePath.Add("DeviceListCSV",$script:DefaultFilePathDeviceListCSV)
    $script:DataHashTable.config.filePath.Add("DeviceIssuesListHTML",$script:DefaultFilePathDeviceIssuesListHTML)
    $script:DataHashTable.config.filePath.Add("DeviceIssuesListCSV",$script:DefaultFilePathDeviceIssuesListCSV)
    $script:DataHashTable.config.filePath.Add("HubListHTML",$script:DefaultFilePathHubListHTML)
    $script:DataHashTable.config.filePath.Add("HubListCSV",$script:DefaultFilePathHubListCSV)
    $script:DataHashTable.config.filePath.Add("HubMeshDeviceListHTML",$script:DefaultFilePathHubMeshDeviceListHTML)
    $script:DataHashTable.config.filePath.Add("HubMeshDeviceListCSV",$script:DefaultFilePathHubMeshDeviceListCSV)
    $script:DataHashTable.config.sendWebCallOnIssue = [ordered]@{}
    $script:DataHashTable.config.sendWebCallOnIssue.Add("categories",@("lowBattery","inactiveDevices","offlineDevices","hubMesh_orphanedDevices","hubMesh_disabledOnSourceDevice","hubMesh_noRemoteDevice"))
    $script:DataHashTable.config.sendWebCallOnIssue.Add("globalWebCallURLStatus",$script:DefaultGlobalWebCallURLStatus)
    $script:DataHashTable.config.sendWebCallOnIssue.Add("status",$script:DefaultSendWebCallOnIssueStatus)
    foreach ($Category in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
        $script:DataHashTable.config.excludedDevicesFromIssuesReporting.$Category = @()
        $script:DataHashTable.config.sendWebCallOnIssue.$Category = [ordered]@{}
        $script:DataHashTable.config.sendWebCallOnIssue.$Category.Add("status",$script:DefaultSendWebCallOnIssueStatus)
    }
    $script:DataHashTable.config.textColour = $script:DefaultTextColour

    $script:PSDefaultParameterValues['*:ForegroundColor'] = $script:DataHashTable.config.textColour    
}

function InitialiseDB { #Checks if data already exists on disk and provides options to the user on how to proceed
    if (Get-ChildItem $script:DataJSONFilePath -ErrorAction SilentlyContinue) {
        Write-Host "Reading config from disk..."
        ReadConfigFromDisk
        WriteTopOfPage "Data file detected on disk"
        Write-Host "Device data from a previous iteration of the script has been detected."
        if ($script:DataHashTable.config.deviceListLastUpdated) {#There appears to be valid data in the file
            Write-Host "The data was last updated $(Get-Date $script:DataHashTable.config.deviceListLastUpdated -Format $script:DataHashTable.config.dateTimeFormat)`n"
            $MenuContents = (   "What would you like to do?",
                                "Use the saved data", #1
                                "Run a new scan", #2
                                "Exit program" #99
            )
        } else {#No deviceListLastUpdated value detected which indicates that a device scan has never taken place. There is most likely no device data in this file
            Write-Host "However, there doesn't appear to be any valid device data in the file.`n" -ForegroundColor "Yellow"
            $MenuContents = (   "What would you like to do?",
                                "Use the saved data anyway", #1
                                "Run a new scan", #2
                                "Exit program" #99
            )
        }
        $Option = WriteMenuToHost $MenuContents
        switch ($Option) {
            1 {
                WriteTopOfPage "Reading data from disk..."
                ReadDataFromDisk
                Clear-Host
                Break
            }
            2 {
                RunANewScanMenu
                Break
            }
            99 {
                Write-Host "Exiting"
                Exit
            }
        }
    } else {
        WriteTopOfPage "No saved data detected"
        Write-Host "No saved data has been detected."
        $MenuContents = (   "What would you like to do?",
                            "Run a new scan", #1
                            "Exit program" #99
        )
        $Option = WriteMenuToHost $MenuContents
        switch ($Option) {
            1 {
                RunANewScanMenu
                Break
            }
            99 {
                Write-Host "Exiting"
                Exit
            }
        }
    }
}

function RunANewScanMenu { #Displays the Run a new scan menu
    WriteTopOfPage "Run a new scan"
    Write-Host "The scan uses a list of hub IP addresses to query."
    $MenuContents = (   "How would you like to provide this list?",
                        "Use the contents of '$($script:HubListFilePath)'", #1
                        "Enter IP address(es) manually", #2
                        "Scan network for hubs", #3
                        "To main menu" #99
    ) 
    $Option = WriteMenuToHost $MenuContents
    Write-Host
    switch ($Option) {
        1 {
            if (ReadHubsTXTListFromDisk) {
                ScanForData -HubIPAddressList (ReadHubsTXTListFromDisk)
                PauseHEDevicesToolKit
                Break
            } else {
                Write-Host "`nNo IP addresses for Hubitat Elevation hubs found in '$script:HubListFilePath'." -ForegroundColor "Red"
                PauseHEDevicesToolKit
                RunANewScanMenu
                Break
            }
            
        }
        2 {
            Write-Host
            Write-Host
            Write-Host "Enter the IP address(es) for the HE hub(s) (comma separated): " -NoNewline
            ScanForData -HubIPAddressList ((Read-Host) -split ",")
            PauseHEDevicesToolKit
            Break
        }
        3 {
            Write-Host
            Write-Host
            $ComputerIPaddresses = (Get-NetIPAddress -AddressFamily IPv4 -SuffixOrigin Dhcp -AddressState Preferred).IPaddress
            if ($ComputerIPaddresses.count -lt 1) {#No IP address detected
                Write-Host "Unabled to determine which network this computer is on" -ForegroundColor "Red"
                PauseHEDevicesToolKit
                RunANewScanMenu   
            } elseif ($ComputerIPaddresses.count -ge 2) {# Two or more IP addresses detected
                $MenuContents = "Two or more IP addresses were detected on this computer. `nWhich one would you like to use for the scan?"
                for ($i=0;$i -lt $ComputerIPaddresses.count;$i++) {
                    $MenuContents += "<<<SPLIT>>>" + $ComputerIPaddresses[$i]
                }
                $MenuContents += "<<<SPLIT>>>Back to run a new scan menu" #99
                $MenuContents = $MenuContents -split "<<<SPLIT>>>"
                $OptionTwo = WriteMenuToHost $MenuContents
                switch ($OptionTwo) {
                    99 {
                        RunANewScanMenu
                        Break
                    }
                    default {
                        $ComputerIPaddresses = $ComputerIPaddresses[$OptionTwo-1]
                    }
                }
                Write-Host
            }
            $Subnet = ($ComputerIPaddresses -split "\.")[0]+"."+($ComputerIPaddresses -split "\.")[1]+"."+($ComputerIPaddresses -split "\.")[2]+"."
            $IPaddressesToScan = for ($i=0;$i -lt 256;$i++) {$address = $Subnet + $i;if ($address -ne $ComputerIPaddresses){$address}}
            ScanForData -HubIPAddressList $IPaddressesToScan
            PauseHEDevicesToolKit
            Break
        }
        99 {
            MainMenu
            Break
        }
    }
}

function SearchForDeviceMenu { #Displays search menus on the screen
    WriteTopOfPage "Search for device"
    $SearchTypeMenuContents = ( "What type of search would you like to perform?",
                                "Search for devices", #1
                                "Search for Hub Mesh devices and display their Hub Mesh hierarchy", #2
                                "Back to main menu" #99
    )
    $SearchType = WriteMenuToHost $SearchTypeMenuContents

    if ($SearchType -eq 99) {MainMenu;Break}
    
    WriteTopOfPage $SearchTypeMenuContents[$SearchType]
    $SearchByMenuContents = (   "What would you like to search by?",
                                "Device name", #1
                                "Back to search menu" #99
    )
    $SearchBy = WriteMenuToHost $SearchByMenuContents

    if ($SearchBy -eq 99) {SearchForDeviceMenu;Exit}
    
    WriteTopOfPage "Search"
    Write-Host ("Search selected: {0}" -f $SearchTypeMenuContents[$SearchType])
    switch ($SearchBy) {
        1   {
                Write-Host "Search by name"
                Write-Host
                Write-Host "The search is case insensitive and uses wildcards on both sides"
                Write-Host "of the search term, so a search for 'WiT' will return hits for"
                Write-Host "devices with 'switch' in their name."
                Write-Host
                Write-Host "Enter the search term: " -NoNewline
                $SearchTerm = "*" + (Read-Host).Trim().ToLower() + "*"
                Write-Host
                Write-Host
                $SearchResult = $script:DataHashTable.devices.keys | Where-Object {$script:DataHashTable.devices.$_.deviceName -like $SearchTerm}
                switch ($SearchType) {
                    1 {  
                        WriteListOfDevicesToOutputDevice $SearchResult
                    }
                    2 {
                        WriteListOfDevicesToOutputDevice $SearchResult -ShowHubMeshHierarchy
                    }
                }
                
                Write-Host
                PauseHEDevicesToolKit
                SearchForDeviceMenu
                Exit
            }
        2   {
                Break
            }
    }
}

function FindSourceDevice { #Hub mesh - Locates the source device ID of each device supplied and returns a list of only source device IDs
    param (
        [string[]]$DeviceIDList
    )

    if ($DeviceIDList.count -ge 1) {
        $SourceDevicesFound = [ordered]@{}
        foreach ($DeviceID in $DeviceIDList) {
            if ($script:DataHashTable.devices.$DeviceID.hubMeshEnabled -eq $true) { # Device $DeviceID is a heb mesh enabled source device. Add it to $SourceDevicesFound
                $SourceDevicesFound[$DeviceID] = @{found=1}
            } elseif ($script:DataHashTable.devices.$DeviceID.sourceDeviceID) { #Remote device found. Add its source device to $SourceDevicesFound
                $SourceDevicesFound[($script:DataHashTable.devices.$DeviceID.sourceDeviceID)] = @{found=1}
            }
        }
        $SourceDevicesFound.keys
    }
}

function DetectIssuesWithDevices { #Runs through a list of checks against the supplied list of device IDs to detect any issues with them
    param (
        [string[]]$DeviceIDList = $null,
        [ValidateSet("DeviceList","SourceList", IgnoreCase = $true)]
        [string]$TypeOfList = "DeviceList",
        [switch]$NoSendWebCallFunction #Prevents the sendWebCallFunction from executing
    )
    if ($DeviceIDList.count -ge 1) {
        $script:DevicesWithIssuesFound = @{}
        $script:DevicesWithIssuesHidden = @{}
        foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
            $script:DevicesWithIssuesHidden.Add($IssueCategory,0)
        }
        $InactiveDeviceTime = (Get-Date).AddMinutes(-$script:DataHashTable.config.inactivityThresholdMinutes)
        
        ##### Run through the device list provided and check for general issues #####
        foreach ($DeviceID in $DeviceIDList) {
            if ($script:DataHashTable.devices.$DeviceID.deviceName) { #Check that the device has a name first. A device without a name happens when a source hubmesh device has been removed but one or more remote devices remain
                if (($script:DataHashTable.devices.$DeviceID.deviceName).ToLower() -like "offline*") { #Check if the device is offline
                    if (($script:DataHashTable.config.excludedDevicesFromIssuesReporting.offlineDevices) -notcontains $DeviceID) { #Check that the device hasn't been excluded from issue category   
                        $script:DevicesWithIssuesFound[$DeviceID] += @{offlineDevices = $true}
                    } else { #Device is excluded from issue category - add 1 to the tally of devices hidden from display
                        $script:DevicesWithIssuesHidden.offlineDevices ++
                    }
                }
            }
            if ($script:DataHashTable.devices.$DeviceID.batteryCharge -AND $script:DataHashTable.devices.$DeviceID.linked -eq $false){ #Check if a battery charge value is recorded for the device but don't check devices that are remote Hub Mesh devices
                if ([int]$script:DataHashTable.devices.$DeviceID.batteryCharge -lt $script:DataHashTable.config.lowBatteryChargeThreshold) { #Check if the battery charge is less than the threshold
                    if (($script:DataHashTable.config.excludedDevicesFromIssuesReporting.lowBattery) -notcontains $DeviceID) { #Check that the device hasn't been excluded from issue category 
                        $script:DevicesWithIssuesFound[$DeviceID] += @{lowBattery = $true}
                    } else { #Device is excluded from issue category - add 1 to the tally of devices hidden from display
                        $script:DevicesWithIssuesHidden.lowBattery ++
                    }
                }
            }
            if ($script:DataHashTable.devices.$DeviceID.wirelessProtocol) { #Check if device is running either Zigbee, ZWave or Matter
                if ($script:DataHashTable.devices.$DeviceID.lastActivity) { #Check if last activity is recorded for the device
                    if ((Get-Date ($script:DataHashTable.devices.$DeviceID.lastActivity)) -lt $InactiveDeviceTime) { #Check if the last activity timestamp is older than the threshold
                        if (($script:DataHashTable.config.excludedDevicesFromIssuesReporting.inactiveDevices) -notcontains $DeviceID) { #Check that the device hasn't been excluded from issue category
                            $script:DevicesWithIssuesFound[$DeviceID] += @{inactiveDevices = $true}
                        } else { #Device is excluded from issue category - add 1 to the tally of devices hidden from display
                            $script:DevicesWithIssuesHidden.inactiveDevices ++
                        }
                    }
                }
            }
        }
        
        ##### Populate list of Hub Mesh source devices and check for Hub Mesh specific issues #####
        if ($TypeOfList.ToLower() -eq "devicelist") { #If the provided list is just a list of device ID's, then run it through the FindSourceDevice function first
            $HubMeshSourceDevices = FindSourceDevice $DeviceIDList
        } else { #The provided list is already a list of source devices
            $HubMeshSourceDevices = $DeviceIDList
        }
        foreach ($HubMeshSourceDeviceID in $HubMeshSourceDevices) {
            if (-NOT ($script:DataHashTable.devices.$HubMeshSourceDeviceID.deviceName)) {#The $SourceDeviceID doesn't contain a deviceName. This happens when a source hubmesh device has been removed but one or more remote devices remain
                if ($script:DataHashTable.hubMeshSourceDevices.$HubMeshSourceDeviceID) {#Find each orphaned remote hub mesh device to this removed source device
                    foreach ($DeviceID in $script:DataHashTable.hubMeshSourceDevices.$HubMeshSourceDeviceID) {
                        if (($script:DataHashTable.config.excludedDevicesFromIssuesReporting.hubMesh_orphanedDevices) -notcontains $DeviceID) { #Unless device has been excluded from issue category
                            $script:DevicesWithIssuesFound[$DeviceID] += @{hubMesh_orphanedDevices = $true}
                        } else { #Device is excluded from issue category - add 1 to the tally of devices hidden from display
                            $script:DevicesWithIssuesHidden.hubMesh_orphanedDevices ++
                        }
                    }
                }
            } else { #For all other source devices that actually do contain a valid name (and thus exist)
                if (($script:DataHashTable.devices.$HubMeshSourceDeviceID.deviceName).ToLower() -like "offline*" -AND $script:DevicesWithIssuesFound[$HubMeshSourceDeviceID].offlineDevices -eq $false) { #Check if the source device is offline and issue hasn't already been logged
                    if (($script:DataHashTable.config.excludedDevicesFromIssuesReporting.offlineDevices) -notcontains $HubMeshSourceDeviceID) { #Check that the device hasn't been excluded from issue category
                        $script:DevicesWithIssuesFound[$HubMeshSourceDeviceID] += @{offlineDevices = $true}
                    } else { #Device is excluded from issue category - add 1 to the tally of devices hidden from display
                        $script:DevicesWithIssuesHidden.offlineDevices ++
                    }
                }
            }
            
            if (-NOT $script:DataHashTable.hubMeshSourceDevices.$HubMeshSourceDeviceID) { #No remote devices detected. Hub mesh can be disabled on the source device
                if (($script:DataHashTable.config.excludedDevicesFromIssuesReporting.hubMesh_noRemoteDevice) -notcontains $HubMeshSourceDeviceID) { #Check that the device hasn't been excluded from issue category
                    $script:DevicesWithIssuesFound[$HubMeshSourceDeviceID] += @{hubMesh_noRemoteDevice = $true}
                } else { #Device is excluded from issue category - add 1 to the tally of devices hidden from display
                    $script:DevicesWithIssuesHidden.hubMesh_noRemoteDevice ++
                }
            } else { #Remote devices detected.
                if ($script:DataHashTable.devices.$HubMeshSourceDeviceID.hubMeshEnabled -eq $false) { #Remote device exists, but hub mesh is disabled on the source device
                    if (($script:DataHashTable.config.excludedDevicesFromIssuesReporting.hubMesh_disabledOnSourceDevice) -notcontains $HubMeshSourceDeviceID) { #Check that the device hasn't been excluded from issue category
                        $script:DevicesWithIssuesFound[$HubMeshSourceDeviceID] += @{hubMesh_disabledOnSourceDevice = $true}
                    } else { #Device is excluded from issue category - add 1 to the tally of devices hidden from display
                        $script:DevicesWithIssuesHidden.hubMesh_disabledOnSourceDevice ++
                    }
                }
                foreach ($RemoteDeviceID in $script:DataHashTable.hubMeshSourceDevices.$HubMeshSourceDeviceID) {
                    if (($script:DataHashTable.devices.$RemoteDeviceID.deviceName).ToLower() -like "offline*" -AND $script:DevicesWithIssuesFound[$RemoteDeviceID].offlineDevices -eq $false) { #Check if the remote device is offline and issue hasn't already been logged
                        if (($script:DataHashTable.config.excludedDevicesFromIssuesReporting.offlineDevices) -notcontains $RemoteDeviceID) { #Check that the device hasn't been excluded from issue category
                            $script:DevicesWithIssuesFound[$RemoteDeviceID] += @{offlineDevices = $true}
                        } else { #Device is excluded from issue category - add 1 to the tally of devices hidden from display
                            $script:DevicesWithIssuesHidden.offlineDevices ++
                        }
                    }
                }
            }
        }

        ##### sendWebCall #####
        if ($NoSendWebCallFunction -eq $false -AND $script:DataHashTable.config.sendWebCallOnIssue.status -eq "Enabled" -AND $script:DevicesWithIssuesFound.count -ge 1){
            $CategoriesWithIssuesFound = [ordered]@{}
            foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
                $CategoriesWithIssuesFound[$IssueCategory] = $script:DevicesWithIssuesFound.keys | Where-Object {$script:DevicesWithIssuesFound[$_].$IssueCategory -eq $true}
            }
            foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
                if ($script:DataHashTable.config.sendWebCallOnIssue.$IssueCategory.status -eq "Enabled" -AND $CategoriesWithIssuesFound.$IssueCategory.count -ge 1){
                    $WebCallURL = ""
                    if ($script:DataHashTable.config.sendWebCallOnIssue.$IssueCategory.webCallURLStatus -eq "Enabled" -AND $script:DataHashTable.config.sendWebCallOnIssue.$IssueCategory.webCallURL) {
                        $WebCallURL = $script:DataHashTable.config.sendWebCallOnIssue.$IssueCategory.webCallURL
                    } elseif ($script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURLStatus -eq "Enabled" -AND $script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURL) {
                        $WebCallURL = $script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURL
                    }
                    if ($WebCallURL) {
                        try {
                            $null = Invoke-WebRequest -Uri $WebCallURL -TimeoutSec 5 -ErrorAction SilentlyContinue
                        }
                        catch {
                            Write-Host
                            Write-Host "Failed to send a web call for the issue category $IssueCategory using the URL $WebCallURL with error: `n`n$_" -ForegroundColor "Red"
                            Write-Host
                            PauseHEDevicesToolKit
                        }
                    }
                }
            }
        }
    }
}

function GenerateOutput { #Returns the HTML code for creating an HTML table with device data and displays the data on screen
    param (
        $DeviceDataHashTable = @{}
    )
    $MaxTextWidthOfDescriptionColumn = 9
    $MaxHTMLWidthOfDescriptionColumn = 4.0
    $MaxTextWidthOfIndentationColumn = 4

    $IndentationLevel = 0
    if ($DeviceDataHashTable.settings.identationLevel) {
        switch ($DeviceDataHashTable.settings.identationLevel) {
            {$_ -lt 0} {$IndentationLevel=0;break}
            {$_ -gt 3} {$IndentationLevel=3;break}
            default {$IndentationLevel=$DeviceDataHashTable.settings.identationLevel}
        }
    }
    
    $NbrOfColumns = $IndentationLevel +1 #The number of columns for the table is the indentation level +1
    if ($DeviceDataHashTable.settings.useFieldDescriptions) {  #If field descriptions is used, we'll need to add another column to the table to accommodate this
        $NbrOfColumns++
        #### Determine how many characters the longest displayName is to properly size up the column ####
        foreach ($Classification in ($script:IterationArray | Where-Object {($DeviceDataHashTable.$_.keys).count -gt 0})) {
            foreach ($LineItem in ($script:IterationArray | Where-Object {($DeviceDataHashTable.$Classification.$_.displayName)})) {
                if (($DeviceDataHashTable.$Classification.$LineItem.displayName).length -gt $MaxTextWidthOfDescriptionColumn-2) {
                    $MaxTextWidthOfDescriptionColumn = ($DeviceDataHashTable.$Classification.$LineItem.displayName).length+2
                    $MaxHTMLWidthOfDescriptionColumn = $MaxTextWidthOfDescriptionColumn * 0.47
                }
            }
        }
    }

    if ($DeviceDataHashTable.two.one.displayValue) { #If the data contains any secondary fields, add another column to accomodate that
        $NbrOfColumns++
    }
    
    $HTML = "<table class=`"minimalBottomMargin`">"
    $Screen = ""
    
    foreach ($Classification in ($script:IterationArray | Where-Object {($DeviceDataHashTable.$_.keys).count -gt 0})) {
        foreach ($LineItem in ($script:IterationArray | Where-Object {($DeviceDataHashTable.$Classification.$_.displayValue)})) {
            $NbrOfIndentationColumns = $NbrOfColumns -1
            $SpanColumns = 1 #The number of columns that the primary name field will have to span
            
            if ($Classification -eq "one") {
                $HTML += "<tr>"
                if ($DeviceDataHashTable.two.one.displayValue) { #If the data contains any secondary fields
                    $NbrOfIndentationColumns --
                    $SpanColumns ++
                }
            } elseif ($Classification -eq "two") {
                $HTML += "<tr class=`"secondaryClassification`">"
            }
            
            if ($DeviceDataHashTable.settings.firstLineIsTitle -AND $Classification -eq "one" -AND $LineItem -eq "one") {#Check if firstLineIsTitle is true and if this is the first line
                $SpanColumns = $NbrOfColumns
                $NbrOfIndentationColumns = 0
            } else {
                if ($DeviceDataHashTable.settings.useFieldDescriptions) {
                    $NbrOfIndentationColumns --
                }
                if ($NbrOfIndentationColumns -gt 0) {
                    for ($i=0;$i -lt $NbrOfIndentationColumns;$i++) {
                        $HTML += "<td class=`"indentationColumn`"></td>"
                    }
                }
                if ($DeviceDataHashTable.settings.useFieldDescriptions) {
                    $HTML += "<td style=`"width: ${MaxHTMLWidthOfDescriptionColumn}em;`" class=`"descriptionColumn`">$($DeviceDataHashTable.$Classification.$LineItem.displayName):</td>"
                    $Screen += "$($DeviceDataHashTable.$Classification.$LineItem.displayName)`: ".padleft(($NbrOfIndentationColumns*$MaxTextWidthOfIndentationColumn+$MaxTextWidthOfDescriptionColumn)," ")
                } else {
                    $Screen += (" " * ($NbrOfIndentationColumns*$MaxTextWidthOfIndentationColumn))
                }
            }
            
            $HTML += "<td colspan=`"$SpanColumns`">"
            if ($DeviceDataHashTable.$Classification.$LineItem.inUseBy) {
                $HTML += "<p class=`"inUseBy`">"
                foreach ($AppHashTable in $DeviceDataHashTable.$Classification.$LineItem.displayValue) {
                    $i = 0
                    foreach ($AppID in $AppHashTable.keys) {
                        $i++
                        if ($AppHashTable.$AppID.appURL) {
                            $HTML += "<a href=`"$($AppHashTable.$AppID.appURL)`" target=`"_blank`">$($AppHashTable.$AppID.appName)</a>"
                        } else {
                            $HTML += "$($AppHashTable.$AppID.appName)"
                        }
                        if ($i -eq 2) { #When printing the second and subsequent InUseBy app to screen, replace the description column text with white space
                            $Screen = (" " * $Screen.Length)
                        }
                        if ($script:DataHashTable.config.outputDevice.screen){Write-Host ("{0}{1} {2}" -f $Screen,[char]0x2022,$AppHashTable.$AppID.appName) -NoNewline}
                        if ($script:DataHashTable.apps.$AppID.appStatus -eq "Disabled") {
                            $HTML += "<span class=`"red`"> *** Disabled *** </span>"
                        }
                        $HTML += "<br>"
                        if ($script:DataHashTable.config.outputDevice.screen){
                            if ($AppHashTable.$AppID.appURL) {
                                Write-Host "  $($AppHashTable.$AppID.appURL)" -NoNewline
                            }
                            if ($script:DataHashTable.apps.$AppID.appStatus -eq "Disabled"){
                                Write-Host " *** Disabled *** " -NoNewline -ForegroundColor "Red"
                            }
                            Write-Host
                        }
                    }
                }
                $HTML += "</p>"
            } elseif ($DeviceDataHashTable.$Classification.$LineItem.children) {
                $HTML += "<p class=`"children`">"
                foreach ($ChildHashTable in $DeviceDataHashTable.$Classification.$LineItem.displayValue) {
                    $i = 0
                    foreach ($ChildID in $ChildHashTable.keys) {
                        $i++
                        if ($ChildHashTable.$ChildID.childURL) {
                            $HTML += "<a href=`"$($ChildHashTable.$ChildID.childURL)`" target=`"_blank`">$($ChildHashTable.$ChildID.childName)</a>"
                        } else {
                            $HTML += "$($ChildHashTable.$ChildID.childName)"
                        }
                        if ($i -eq 2) { #When printing the second and subsequent child to screen, replace the description column text with white space
                            $Screen = (" " * $Screen.Length)
                        }
                        if ($script:DataHashTable.config.outputDevice.screen){Write-Host ("{0}{1} {2}" -f $Screen,[char]0x2022,$ChildHashTable.$ChildID.childName) -NoNewline}
                        if ($script:DataHashTable.devices.$ChildID.deviceStatus -eq "Disabled") {
                            $HTML += "<span class=`"red`"> *** Disabled *** </span>"
                        }
                        $HTML += "<br>"
                        if ($script:DataHashTable.config.outputDevice.screen){
                            if ($ChildHashTable.$ChildID.childURL) {
                                Write-Host "  $($ChildHashTable.$ChildID.childURL)" -NoNewline
                            }
                            if ($script:DataHashTable.devices.$ChildID.deviceStatus -eq "Disabled"){
                                Write-Host " *** Disabled *** " -NoNewline -ForegroundColor "Red"
                            }
                            Write-Host
                        }
                    }
                }
                $HTML += "</p>"
            } else {
                $Screen += $DeviceDataHashTable.$Classification.$LineItem.displayValue
                
                if ($DeviceDataHashTable.settings.firstLineIsTitle -AND $Classification -eq "one" -AND $LineItem -eq "one") {#Check if firstLineIsTitle is true and if this is the first line
                    if ($DeviceDataHashTable.$Classification.$LineItem.warning) {
                        $HTML += "<h4 class=`"listOfDevicesWarning`">"
                    } else {
                        $HTML += "<h4 class=`"listOfDevices`">"
                    }
                }
                if ($DeviceDataHashTable.$Classification.$LineItem.URL) {
                    $HTML += "<a href=`"$($DeviceDataHashTable.$Classification.$LineItem.URL)`" target=`"_blank`">$($DeviceDataHashTable.$Classification.$LineItem.displayValue)</a>"
                } else {
                    $HTML += "$($DeviceDataHashTable.$Classification.$LineItem.displayValue)"
                }
                if ($script:DataHashTable.config.outputDevice.screen){
                    if ($DeviceDataHashTable.$Classification.$LineItem.warning){
                        Write-Host $Screen -NoNewline -ForegroundColor "Red"
                    } else {
                        Write-Host $Screen -NoNewline
                    }
                }
                if ($DeviceDataHashTable.$Classification.$LineItem.disabled) {
                    if ($script:DataHashTable.config.outputDevice.screen){Write-Host " *** Disabled *** " -NoNewline -ForegroundColor "Red"}
                    $HTML += "<span class=`"red`"> *** Disabled *** </span>"
                }
                if ($DeviceDataHashTable.$Classification.$LineItem.linked) {
                    if ($script:DataHashTable.config.outputDevice.screen){Write-Host " <Linked>" -NoNewline -ForegroundColor "Blue"}
                    $HTML += "<a class=`"blue`" href=`"$($DeviceDataHashTable.$Classification.$LineItem.linked)`"> &#60;Linked&#62; </a>"
                }
                if ($DeviceDataHashTable.settings.firstLineIsTitle -AND $Classification -eq "one" -AND $LineItem -eq "one") {
                    $HTML += "</h4>"
                }
                
                if ($script:DataHashTable.config.outputDevice.screen){
                    Write-Host
                    if ($DeviceDataHashTable.$Classification.$LineItem.URL) {
                        if ($DeviceDataHashTable.settings.useFieldDescriptions) {
                            $Screen = "URL: ".padleft(($NbrOfIndentationColumns*$MaxTextWidthOfIndentationColumn+$MaxTextWidthOfDescriptionColumn)," ")
                        } else {
                            $Screen = (" " * ($NbrOfIndentationColumns*$MaxTextWidthOfIndentationColumn))
                        } 
                        $Screen += $DeviceDataHashTable.$Classification.$LineItem.URL
                        Write-Host $Screen
                    }
                }
            }
            $Screen = ""
            $HTML += "</td></tr>"
        }
    }
    $HTML += "</table>"
    if ($script:DataHashTable.config.outputDevice.HTML) {$HTML}
}

function WriteIssuesListToOutputDevice { #Sends the contents of the $DevicesWithIssuesFound variable to the selected output device(s)
    param (
        [switch]$ExcludeDevicesMode
    )
    
    WriteTopOfPage "Device issues"
    if ($ExcludeDevicesMode) {
        $NbrOfOptions = 0
        $HashTableOfDevicesWithIssues = @{}
        $CaptureOutputDeviceState = @{screen=$script:DataHashTable.config.outputDevice.screen
                                    HTML=$script:DataHashTable.config.outputDevice.HTML
                                    CSV=$script:DataHashTable.config.outputDevice.CSV}
        $script:DataHashTable.config.outputDevice.screen = $true
        $script:DataHashTable.config.outputDevice.HTML = $false
        $script:DataHashTable.config.outputDevice.CSV = $false
    }
    
    if ($script:DataHashTable.config.outputDevice.screen -OR $script:DataHashTable.config.outputDevice.HTML -OR $script:DataHashTable.config.outputDevice.CSV) {
        if ($script:DataHashTable.config.outputDevice.HTML) {
            $HTML = SetHTMLHeader "Device issues"
        }
        if ($script:DataHashTable.config.outputDevice.CSV) {
            $CSVFilePath = $script:DataHashTable.config.filePath.DeviceIssuesListCSV
            Remove-Item $CSVFilePath -Force -ErrorAction SilentlyContinue
            $CSVContent = ""
        }

        if ($script:DevicesWithIssuesFound.count -ge 1) {
            $CategoryText = @{}
            $CategoryText["lowBattery"] = @{title="DEVICES WITH LOW BATTERY CHARGE"
                                                            text="The following devices have a battery charge of less than $($script:DataHashTable.config.lowBatteryChargeThreshold)%:"}
            $CategoryText["inactiveDevices"] = @{title="INACTIVE DEVICES"
                                                            text="An inactive device is a Zigbee, ZWave or Matter device that hasn't reported any activity for the last $($script:DataHashTable.config.inactivityThresholdMinutes) minutes. Causes for this include device dropping off the mesh, the device is non-functional (e.g. dead or dead battery), or there may just not have been any activity for the device to report and it has not responded to the hub's periodic pings.","The following inactive devices were detected:"}
            $CategoryText["offlineDevices"] = @{title="OFFLINE DEVICES"
                                                            text="An offline device is a device that HE no longer can contact. It could for example be a Hub Mesh remote device where Hub Mesh has been disabled on the source device, leading to the remote device going offline. It could also be a LAN device that has been disconnected from the network or it has received a new IP address and is no longer accessible on the original IP address.","The following offline devices were detected:"}
            $CategoryText["hubMesh_orphanedDevices"] = @{title="HUB MESH - ORPHANED DEVICES"
                                                            text="An orphaned device is a Hub Mesh remote device whose source device has been removed.","Check first that all hubs in your Hub Mesh environment have been added to this program. This error is expected for each remote device of source devices that are installed on hubs that haven't been included.","If all hubs are accounted for, the only way to resolve this issue is to remove the orphaned device and replace it with a working device.","The following orphaned devices were detected:"}
            $CategoryText["hubMesh_disabledOnSourceDevice"] = @{title="HUB MESH - HUB MESH DISABLED ON SOURCE DEVICE"
                                                            text="The problem with these devices is that Hub Mesh has been disabled on the source device even though remote devices exist. The remote devices will be offline and non-functional at this stage.","Resolve the issue by enabling Hub Mesh on the source device or removing the remote device.","The following devices with this issue were detected:"}
            $CategoryText["hubMesh_noRemoteDevice"] = @{title="HUB MESH - NO REMOTE DEVICES DETECTED"
                                                            text="Hub Mesh is enabled on these source devices, but no remote devices were detected on any other hubs. It's not an issue per se, but to reduce system resorce usage, any Hub Mesh enabled device not linked to a remote device should have Hub Mesh disabled.","But, check first that all hubs in your Hub Mesh environment have been added to this program. This error is expected for each source device that have remote devices on hubs that haven't been included.","The following source devices without remote devices were detected:"}
            
            foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) { #If there are devices excluded from an issue category, add some text to the end of categoryText to say how many were excluded
                if ($script:DevicesWithIssuesHidden.$IssueCategory -eq 1) {
                    ($CategoryText.$IssueCategory["text"])[$CategoryText.$IssueCategory["text"].count-1] = ($CategoryText.$IssueCategory["text"])[$CategoryText.$IssueCategory["text"].count-1].Substring(0,($CategoryText.$IssueCategory["text"])[$CategoryText.$IssueCategory["text"].count-1].length-1) + " (1 device with this issue has been excluded from this list):"
                } elseif ($script:DevicesWithIssuesHidden.$IssueCategory -ge 2) {
                    ($CategoryText.$IssueCategory["text"])[$CategoryText.$IssueCategory["text"].count-1] = ($CategoryText.$IssueCategory["text"])[$CategoryText.$IssueCategory["text"].count-1].Substring(0,($CategoryText.$IssueCategory["text"])[$CategoryText.$IssueCategory["text"].count-1].length-1) + " ($($script:DevicesWithIssuesHidden.$IssueCategory) devices with this issue have been excluded from this list):"
                }
            }
            
            $CategoriesWithIssuesFound = [ordered]@{}
            foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
                $CategoriesWithIssuesFound[$IssueCategory] = $script:DevicesWithIssuesFound.keys | Where-Object {$script:DevicesWithIssuesFound[$_].$IssueCategory -eq $true}
            }    
            foreach ($Category in ($CategoriesWithIssuesFound.keys | Where-Object {$CategoriesWithIssuesFound[$_].count -ge 1})) {
                if ($script:DataHashTable.config.outputDevice.HTML) {
                    $HTML += "<div><h4 class=`"deviceIssues`">$($CategoryText[$Category].title)</h4></div>
                            <div><p style=`"margin-top: 0px;margin-bottom: 0px;`">"
                }
                if ($script:DataHashTable.config.outputDevice.screen) {
                    Write-Host $CategoryText[$Category].title -ForegroundColor "Yellow"
                }
                if ($script:DataHashTable.config.outputDevice.CSV) {
                    $CSVContent += "`"*** " + $CategoryText[$Category].title.replace('"','""') + " ***`"`n"
                }
                foreach ($line in $CategoryText[$Category].text) {
                    if ($script:DataHashTable.config.outputDevice.HTML) {
                        $HTML += "$line<br>"
                    }
                    if ($script:DataHashTable.config.outputDevice.screen) {
                        Write-Host $line
                    }
                    if ($script:DataHashTable.config.outputDevice.CSV) {
                        $CSVContent += '"' + $line.replace('"','""') + "`"`n"
                    }
                }
                if ($script:DataHashTable.config.outputDevice.HTML) {
                    $HTML += "</p></div>
                            <div class=`"$Category`">"
                }
                if ($script:DataHashTable.config.outputDevice.CSV) {
                    if ($Category -eq "lowBattery") {
                        $CSVHeader = [ordered]@{deviceName="Device Name"
                                                deviceURL="URL"
                                                hub="Hub"
                                                batteryCharge="Battery Charge (%)"          
                        }
                    } elseif ($Category -eq "inactiveDevices") {
                        $CSVHeader = [ordered]@{deviceName="Device Name"
                                                deviceURL="URL"
                                                hub="Hub"
                                                lastActivityTime="Last Activity Time"          
                        }
                    } elseif (@("offlineDevices","hubMesh_orphanedDevices","hubMesh_noRemoteDevice") -contains $Category) {
                        $CSVHeader = [ordered]@{deviceName="Device Name"
                                                deviceURL="URL"
                                                hub="Hub"          
                        }
                    } elseif ($Category -eq "hubMesh_disabledOnSourceDevice") {
                        $CSVHeader = [ordered]@{sourceDeviceName="Source Device Name"
                                                sourceURL="Source URL"
                                                sourceHub="Source Hub"
                                                remoteDevices="Remote Devices"          
                        }
                    }
                    #Print header
                    $i = 0
                    $CSVContent += "`n"
                    foreach ($Column in $CSVHeader.keys) {
                        $i++
                        $CSVContent += '"' + $CSVHeader.$Column.replace('"','""') + '"'
                        if ($i -lt $CSVHeader.count) {
                            $CSVContent += ","
                        }
                    }
                    $CSVContent += "`n"
                }
                foreach ($Device in $CategoriesWithIssuesFound[$Category]) {
                    if ($ExcludeDevicesMode) {
                        $NbrOfOptions ++
                        $HashTableOfDevicesWithIssues["$NbrOfOptions"] = @{deviceID=$Device
                                                                        issueCategory=$Category}
                    }        
                    if ($script:DataHashTable.config.outputDevice.screen) {
                        Write-Host
                        if ($ExcludeDevicesMode) {
                            Write-Host "[$NbrOfOptions]" -NoNewline
                        }
                    }
                    if ($script:DataHashTable.config.outputDevice.HTML) {
                        $HTML+="<BR>"
                    }
                    switch ($Category) {
                        "lowBattery" {
                            if ($script:DataHashTable.config.outputDevice.screen -OR $script:DataHashTable.config.outputDevice.HTML) {
                                $HTML += GenerateOutput @{settings=@{useFieldDescriptions=$true}
                                                        one=@{     one=@{displayValue=$script:DataHashTable.devices.$Device.deviceName
                                                                        displayName="Name"
                                                                        URL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$Device.deviceURL)}
                                                                    two=@{displayValue=$script:DataHashTable.devices.$Device.batteryCharge + "%"
                                                                        displayName="Charge"}}}
                            }
                            if ($script:DataHashTable.config.outputDevice.CSV) {
                                $CSV = @{}
                                $CSV.deviceName=$script:DataHashTable.devices.$Device.deviceName
                                $CSV.deviceURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$Device.deviceURL)
                                $CSV.hub=$script:DataHashTable.hubs.($script:DataHashTable.devices.$Device.hubIPAddress).hubName
                                $CSV.batteryCharge=$script:DataHashTable.devices.$Device.batteryCharge         
                            }
                            break
                        }
                        "inactiveDevices" {
                            if ($script:DataHashTable.config.outputDevice.screen -OR $script:DataHashTable.config.outputDevice.HTML) {
                                $HTML += GenerateOutput @{settings=@{useFieldDescriptions=$true}
                                                          one=@{     one=@{displayValue=$script:DataHashTable.devices.$Device.deviceName
                                                                           displayName="Name"
                                                                           URL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$Device.deviceURL)}
                                                                     two=@{displayValue=(Get-Date ($script:DataHashTable.devices.$Device.lastActivity) -Format $script:DataHashTable.config.dateTimeFormat)
                                                                           displayName="Time"}}}
                            }
                            if ($script:DataHashTable.config.outputDevice.CSV) {
                                $CSV = @{}
                                $CSV.deviceName=$script:DataHashTable.devices.$Device.deviceName
                                $CSV.deviceURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$Device.deviceURL)
                                $CSV.hub=$script:DataHashTable.hubs.($script:DataHashTable.devices.$Device.hubIPAddress).hubName
                                $CSV.lastActivityTime=$script:DataHashTable.devices.$Device.lastActivity        
                            }
                            break
                        }
                        {@("offlineDevices","hubMesh_orphanedDevices","hubMesh_noRemoteDevice") -contains $_} {
                            if ($script:DataHashTable.config.outputDevice.screen -OR $script:DataHashTable.config.outputDevice.HTML) {
                                $HTML += GenerateOutput  @{settings=@{useFieldDescriptions=$true}
                                                           one=@{     one=@{displayValue=$script:DataHashTable.devices.$Device.deviceName
                                                                            displayName="Name"
                                                                            URL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$Device.deviceURL)}
                                                                      two=@{displayValue=$script:DataHashTable.hubs.($script:DataHashTable.devices.$Device.hubIPAddress).hubName
                                                                            displayName="Hub"}}}
                            }       
                            if ($script:DataHashTable.config.outputDevice.CSV) {
                                $CSV = @{}
                                $CSV.deviceName=$script:DataHashTable.devices.$Device.deviceName
                                $CSV.deviceURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$Device.deviceURL)
                                $CSV.hub=$script:DataHashTable.hubs.($script:DataHashTable.devices.$Device.hubIPAddress).hubName        
                            }
                            break
                        }
                        "hubMesh_disabledOnSourceDevice" {
                            if ($script:DataHashTable.config.outputDevice.screen -OR $script:DataHashTable.config.outputDevice.HTML) {
                                $HTML += GenerateOutput @{settings=@{useFieldDescriptions=$true}
                                                          one=@{     one=@{displayValue=$script:DataHashTable.devices.$Device.deviceName
                                                                           displayName="Source device"
                                                                           URL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$Device.deviceURL)}
                                                                     two=@{displayValue=$script:DataHashTable.hubs.($script:DataHashTable.devices.$Device.hubIPAddress).hubName
                                                                           displayName="Hub"}}}
                            }
                            if ($script:DataHashTable.config.outputDevice.CSV) {
                                $i = 0
                                $CSV = @{}
                                $CSV.sourceDeviceName=$script:DataHashTable.devices.$Device.deviceName
                                $CSV.sourceURL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$Device.deviceURL)
                                $CSV.sourceHub=$script:DataHashTable.hubs.($script:DataHashTable.devices.$Device.hubIPAddress).hubName        
                            }
                            foreach ($RemoteDeviceID in $script:DataHashTable.hubMeshSourceDevices.$Device) {
                                if ($script:DataHashTable.config.outputDevice.screen -OR $script:DataHashTable.config.outputDevice.HTML) {
                                    $HTML += GenerateOutput @{settings=@{useFieldDescriptions=$true
                                                                         identationLevel=1}
                                                              one=@{     one=@{displayValue=$script:DataHashTable.devices.$RemoteDeviceID.deviceName
                                                                               displayName="Remote device"
                                                                               URL=($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$RemoteDeviceID.deviceURL)}
                                                                         two=@{displayValue=$script:DataHashTable.hubs.($script:DataHashTable.devices.$RemoteDeviceID.hubIPAddress).hubName
                                                                               displayName="Remote hub"}}}
                                }
                                if ($script:DataHashTable.config.outputDevice.CSV) {
                                    $i++
                                    $CSV.remoteDevices += $script:DataHashTable.devices.$RemoteDeviceID.deviceName + " (" + $RemoteID + " - " + ($script:DataHashTable.config.internetProtocol + $script:DataHashTable.devices.$RemoteDeviceID.deviceURL) + ") " + " on " + $script:DataHashTable.hubs.($script:DataHashTable.devices.$RemoteDeviceID.hubIPAddress).hubName
                                    if ($i -lt ($script:DataHashTable.hubMeshSourceDevices.$Device).count) {
                                        $CSV.remoteDevices += "`n"
                                    }
                                }
                            }
                            break
                        }
                    }
                    if ($script:DataHashTable.config.outputDevice.CSV) {
                        $i = 0
                        foreach ($Column in $CSVHeader.keys) {
                            $i++
                            $CSVContent += '"' + $CSV.$Column.replace('"','""') + '"'
                            if ($i -lt $CSVHeader.count) {
                                $CSVContent += ","
                            }
                        }
                        $CSVContent += "`n"
                    }
                }
                if ($script:DataHashTable.config.outputDevice.screen) {
                    Write-Host
                }
                if ($script:DataHashTable.config.outputDevice.HTML) {
                    $HTML += "</div>`n<div class=`"divider`"></div>"
                }
                if ($script:DataHashTable.config.outputDevice.CSV) {
                    $CSVContent += "`n`n"
                }
            }
            
            ##### Check if there are additional devices excluded from issue categories that weren't listed above #####
            $NbrOfDevicesInNonDisplayedCategories = 0
            foreach ($IssueCategory in ($CategoriesWithIssuesFound.keys | Where-Object {$CategoriesWithIssuesFound.$_.count -lt 1})) {
                $NbrOfDevicesInNonDisplayedCategories += $script:DevicesWithIssuesHidden.$IssueCategory
            }
            if ($NbrOfDevicesInNonDisplayedCategories -ge 1) {
                $AdditionalExcludedDevicesText = "An additional $NbrOfDevicesInNonDisplayedCategories devices with issues belonging to categories not listed above were excluded from the list."
                if ($NbrOfDevicesInNonDisplayedCategories -eq 1) {
                    $AdditionalExcludedDevicesText = "An additional 1 device with issues belonging to a category not listed above was excluded from the list."
                } 
                    
                if ($script:DataHashTable.config.outputDevice.screen) {
                    Write-Host
                    Write-Host $AdditionalExcludedDevicesText -ForegroundColor "Yellow"
                    Write-Host
                }
                if ($script:DataHashTable.config.outputDevice.HTML) {
                    $HTML += "<div><p style=`"margin-top: 0px;margin-bottom: 0px;`">
                              $AdditionalExcludedDevicesText
                              </p></div>"
                }
                if ($script:DataHashTable.config.outputDevice.CSV) {
                    $CSVContent += $AdditionalExcludedDevicesText + "`n`n"
                }
            }
            if ($ExcludeDevicesMode) {
                $script:DataHashTable.config.outputDevice.screen = $CaptureOutputDeviceState.screen
                $script:DataHashTable.config.outputDevice.HTML = $CaptureOutputDeviceState.HTML
                $script:DataHashTable.config.outputDevice.CSV = $CaptureOutputDeviceState.CSV
                Write-Host
                Write-Host
                Write-Host
                Write-Host "Each device in the issue list above is prefixed with a [<number>]. To exclude a device from being listed in the device list, enter the number of the device(s) below."
                Write-Host "NOTE: The device exclusion is per issue category. E.g, if a device is excluded from the lowBattery category, it will still be included in all other categories.`n"
                Write-Host "Enter the number of the device(s) to be excluded (comma separated) or type 'All' to exclude all devices in the list.`n(Leave blank and hit ENTER to cancel): " -NoNewline
                $DevicesToBeExcludedList = (Read-Host).split(",").trim()
                Write-Host
                Write-Host
                if ($DevicesToBeExcludedList) {
                    if ($DevicesToBeExcludedList.ToLower() -eq "all") {
                        foreach ($Option in $HashTableOfDevicesWithIssues.keys) {
                            $script:DataHashTable.config.excludedDevicesFromIssuesReporting.($HashTableOfDevicesWithIssues["$Option"].issueCategory) += $HashTableOfDevicesWithIssues["$Option"].deviceID
                            Write-Host "'$($script:DataHashTable.devices.($HashTableOfDevicesWithIssues["$Option"].deviceID).deviceName)' excluded from '$($HashTableOfDevicesWithIssues["$Option"].issueCategory)'"  -ForegroundColor "Green"
                        }
                    } else {
                        foreach ($Option in $DevicesToBeExcludedList) {
                            if ($HashTableOfDevicesWithIssues["$Option"]){
                                $script:DataHashTable.config.excludedDevicesFromIssuesReporting.($HashTableOfDevicesWithIssues["$Option"].issueCategory) += $HashTableOfDevicesWithIssues["$Option"].deviceID
                                Write-Host "'$($script:DataHashTable.devices.($HashTableOfDevicesWithIssues["$Option"].deviceID).deviceName)' excluded from '$($HashTableOfDevicesWithIssues["$Option"].issueCategory)'"  -ForegroundColor "Green"
                            } else {
                                Write-Host "Couldn't resolve option '$Option'. Did you enter the correct number?" -ForegroundColor "Red"
                            }
                        }
                    }
                    Write-Host
                    WriteDataToDisk
                    PauseHEDevicesToolKit
                }
            }
        } else { #No issues detected
            $NbrOfExcludedDevices = 0
            foreach ($IssueCategory in $script:DevicesWithIssuesHidden.keys) {
                $NbrOfExcludedDevices += $script:DevicesWithIssuesHidden.$IssueCategory
            }
            $NoIssuesTitle = "No issues detected!"
            if ($NbrOfExcludedDevices -eq 1) {
                $NoIssuesText = "1 device with issues was excluded from this list"
            } elseif ($NbrOfExcludedDevices -ge 2) {
                $NoIssuesText = "$NbrOfExcludedDevices devices with issues were excluded from this list"
            }
            if ($script:DataHashTable.config.outputDevice.screen) {
                Write-Host $NoIssuesTitle
                if ($NbrOfExcludedDevices -ge 1) {
                    Write-Host $NoIssuesText
                }
                Write-Host
                Write-Host
            }
            if ($script:DataHashTable.config.outputDevice.HTML) {
                $HTML += "<div><h4 class=`"deviceIssues`">$NoIssuesTitle</h4></div>"
                if ($NbrOfExcludedDevices -ge 1) {
                    $HTML += "<div><p style=`"margin-top: 0px;margin-bottom: 0px;`">
                            $NoIssuesText
                            </p></div>"
                }
            }
            if ($script:DataHashTable.config.outputDevice.CSV) {
                $CSVContent += "*** $NoIssuesTitle ***`n$NoIssuesText`n`n"
            }
            if ($ExcludeDevicesMode) {
                $script:DataHashTable.config.outputDevice.screen = $CaptureOutputDeviceState.screen
                $script:DataHashTable.config.outputDevice.HTML = $CaptureOutputDeviceState.HTML
                $script:DataHashTable.config.outputDevice.CSV = $CaptureOutputDeviceState.CSV
                PauseHEDevicesToolKit
            }
            
            
        }
        if ($script:DataHashTable.config.outputDevice.HTML -AND -NOT $ExcludeDevicesMode) {$HTML += SetHTMLFooter; WriteHTMLtoDisk -HTMLcode $HTML -FilePath $script:DataHashTable.config.filePath.DeviceIssuesListHTML}
        if ($script:DataHashTable.config.outputDevice.screen) {Write-Host}
        if ($script:DataHashTable.config.outputDevice.CSV -AND -NOT $ExcludeDevicesMode) {Add-Content -Path $CSVFilePath -Force -Value $CSVContent}
    } else {
        Write-Host "No supported output device selected." -ForegroundColor "Red"
    }
}

function ConfigMenu { #Displays menus for listing and changing configuration settings
    WriteTopOfPage "Configuration settings"
    Write-Host "Current configuration settings are listed below."
    
    $MenuContents = (   "If you want to change a setting, enter the corresponding option number",
                        "autoLaunchWebBrowser: '$($script:DataHashTable.config.autoLaunchWebBrowser)'", #1
                        "Change file names or file paths", #2
                        "dateTimeFormat: '$($script:DataHashTable.config.dateTimeFormat)'", #3
                        "Devices excluded from issues reporting", #4
                        "inactivityThresholdMinutes: '$($script:DataHashTable.config.inactivityThresholdMinutes)'", #5
                        "internetProtocol: '$($script:DataHashTable.config.internetProtocol)'", #6
                        "lowBatteryChargeThreshold: '$($script:DataHashTable.config.lowBatteryChargeThreshold)'", #7
                        "outputDevice: '$(($script:DataHashTable.config.outputDevice.psobject.properties | Where-Object {$_.Value} | ForEach-Object {$_.Name}) -join ", ")'", #8
                        "Send web call on issue: '$($DataHashTable.config.sendWebCallOnIssue.status)'",#9
                        "textColour: '$($script:DataHashTable.config.textColour)'", #10
                        "Divider",
                        "Reset all configuration settings to default settings", #11 
                        "Back to main menu" #99
    )
    $Option = WriteMenuToHost $MenuContents
    $Setting = ""
    switch ($Option) {
        1 {
            $Setting = "autoLaunchWebBrowser"
            Break
        }
        2 {
            ChangeFilePathMenu
            Break
        }
        3 {
            $Setting = "dateTimeFormat"
            Break
        }
        4 {
            DevicesExcludedFromIssuesReportingMenu
            Break
        }
        5 {
            $Setting = "inactivityThresholdMinutes"
            Break
        }
        6 {
            $Setting = "internetProtocol"
            Break
        }
        7 {
            $Setting = "lowBatteryChargeThreshold"
            Break
        }
        8 {
            SetOutputDevice
            Break
        }
        9 {
            SendWebCallOnIssueMainMenu
            Break
        }
        10 {
            $Setting = "textColour"
            Break
        }
        11 {
            EraseConfigurationSettingsAndSetDefaults
            WriteDataToDisk
            ConfigMenu
            Break
        }
        99{
            MainMenu
            Break
        }
        default {
            ConfigMenu
        }
    }
    WriteTopOfPage $Setting
    
    $DefaultSetting = ""
    $Type = "String"
    if ($Setting -eq "autoLaunchWebBrowser") {
        $DefaultSetting = $script:DefaultAutoLaunchWebBrowser
        $Type = "Toggle"
    } elseif ($Setting -eq "dateTimeFormat") {
        $DefaultSetting = $script:DefaultDateTimeFormat
    } elseif ($Setting -eq "inactivityThresholdMinutes"){
        $DefaultSetting = $script:DefaultInactivityThresholdMinutes
        $Type = "Int"
    } elseif ($Setting -eq "internetProtocol"){
        $DefaultSetting = $script:DefaultInternetProtocol
        $Type = "Toggle"
    } elseif ($Setting -eq "lowBatteryChargeThreshold"){
        $DefaultSetting = $script:DefaultLowBatteryChargeThreshold
        $Type = "Int"
    } elseif ($Setting -eq "textColour"){
        $DefaultSetting = $script:DefaultTextColour
    }
    $MenuContents = 
    $MenuContents = (   "The current setting for $Setting is '$($script:DataHashTable.config.$Setting)'.`nWhat would like to do?",
                        "Enter a new value for $Setting", #1                   
                        "Reset $Setting to the default setting ('$DefaultSetting')", #2
                        "Back to configuration settings" #99
    )
    $Option = WriteMenuToHost $MenuContents
    switch ($Option) {
        1 {
            if ($Type -eq "Toggle"){
                if ($Setting -eq "autoLaunchWebBrowser") { #Toggle between $True and $False
                    $Toggle1 = $true
                    $Toggle2 = $false
                } elseif ($Setting -eq "internetProtocol") { #Toggle between "http://" and "https://"
                    $Toggle1 = "http://"
                    $Toggle2 = "https://"
                }
                if ($script:DataHashTable.config.$Setting -eq $Toggle1) {
                    $script:DataHashTable.config.$Setting = $Toggle2
                } else {
                    $script:DataHashTable.config.$Setting = $Toggle1
                }
                Write-Host
                Write-Host "$Setting has been changed to " -NoNewline
                Write-Host $script:DataHashTable.config.$Setting -ForegroundColor "Green"
                WriteDataToDisk
                Start-Sleep 3
                Break
            } else {
                Write-Host
                if ($Setting -eq "textColour"){
                    Write-Host "Valid colours are: $($script:ValidColours)"
                    Write-Host "New text colour: " -NoNewline
                    $Response = Read-Host
                    if ($script:ValidColours -contains $Response) {
                        $script:DataHashTable.config.$Setting = (Get-Culture).TextInfo.ToTitleCase([string]$Response.ToLower())
                        $script:PSDefaultParameterValues['*:ForegroundColor'] = $script:DataHashTable.config.$Setting
                    } else {
                        Write-Host
                        Write-Host "Invalid colour." -ForegroundColor "Red"
                        Write-Host
                        PauseHEDevicesToolKit
                        Break   
                    }
                } elseif ($Type -eq "String"){
                    Write-Host "New value for $($Setting): " -NoNewline
                    $script:DataHashTable.config.$Setting = [string](Read-Host)
                } elseif ($Type -eq "Int"){
                    Write-Host "New value for $($Setting): " -NoNewline
                    $script:DataHashTable.config.$Setting = [int](Read-Host)
                }
                WriteDataToDisk
            }
            Break
        }
        2 {
            $script:DataHashTable.config.$Setting = $DefaultSetting
            if ($Setting -eq "textColour"){
                $script:PSDefaultParameterValues['*:ForegroundColor'] = $DefaultSetting
            }
            WriteDataToDisk
            Break
        }
        99{
            Break
        }
    }
    ConfigMenu
    
}

function MainMenu { #Displays the main menu
    WriteTopOfPage "main menu"
    Write-Host ("Hubs: {0}" -f $($script:DataHashTable.hubs.count)) 
    Write-Host ("Devices: {0}" -f $($script:DataHashTable.devices.count))
    if ($script:DataHashTable.config.deviceListLastUpdated){Write-Host ("Data last updated: {0}" -f (Get-Date $($script:DataHashTable.config.deviceListLastUpdated) -Format $script:DataHashTable.config.dateTimeFormat))}
    Write-Host ("Output device(s): $(($script:DataHashTable.config.outputDevice.psobject.properties | Where-Object {$_.Value} | ForEach-Object {$_.Name}) -join ", ")")
    Write-Host
    $MenuContents = (   "What would you like to do?",
                        "List all hubs", #1
                        "List all devices", #2
                        "List all Hub Mesh devices", #3
                        "Search for devices", #4
                        "Check for device issues", #5
                        "Divider",
                        "Configuration settings", #6
                        "Divider",
                        "Reset all device data and run a new scan", #7
                        "Exit program" #99
    )
    $Option = WriteMenuToHost $MenuContents
    switch ($Option) {
        1{
            WriteHubListToOutputDevice
            Write-Host     
            PauseHEDevicesToolKit
            MainMenu
            Break
        }
        2{
            WriteListOfDevicesToOutputDevice -ShowHierarchy
            Write-Host
            PauseHEDevicesToolKit
            MainMenu
            Break
        }
        3{
            WriteListOfDevicesToOutputDevice -ShowHubMeshHierarchy
            Write-Host
            PauseHEDevicesToolKit
            MainMenu
            Break
        }
        4{
            SearchForDeviceMenu
            PauseHEDevicesToolKit
            MainMenu
            Break
        }
        5{
            WriteTopOfPage "Checking devices for issues"
            if ($script:DataHashTable.config.outputDevice.HTML -OR $script:DataHashTable.config.outputDevice.screen) {
                Write-Host "Processing..."
                DetectIssuesWithDevices -DeviceIDList $script:DataHashTable.devices.keys -TypeOfList DeviceList
                WriteIssuesListToOutputDevice
            } else {
                Write-Host "No supported output device selected" -ForegroundColor "Red"
                Write-Host
                PauseHEDevicesToolKit
                SetOutputDevice
            }
            PauseHEDevicesToolKit
            MainMenu
            Break
        }
        6{
            ConfigMenu
            PauseHEDevicesToolKit
            MainMenu
            Break
        }
        7{
            RunANewScanMenu
            MainMenu
            Break
        }
        99{
            Write-Host "`nGood bye!`n"
            Exit
        }
    }
}

function SetOutputDevice { #Displays menu for setting the output devices 
    WriteTopOfPage "Set output device"
    Write-Host "The current output device settings are:"
    foreach ($OutputDevice in ($script:DataHashTable.config.outputDevice.psobject.properties).Name) {
        Write-Host "$($OutputDevice): " -NoNewline
        if ($script:DataHashTable.config.outputDevice.$OutputDevice) {
            Write-Host "Enabled" -ForegroundColor "Green"
        } else {
            Write-Host "Disabled" -ForegroundColor "Red"
        }
    }
    Write-Host
    $MenuContents = (   "What would you like to do?",
                        "Toggle screen output", #1
                        "Toggle HTML output", #2
                        "Toggle CSV output", #3
                        "Back to configuration settings" #99
    )
    $Option = WriteMenuToHost $MenuContents
    $Setting = ""
    switch ($Option) {
        1{
            $Setting = "screen"
            Break
        }
        2{
            $Setting = "HTML"
            Break
        }
        3{
            $Setting = "CSV"
            Break
        }
        99{
            ConfigMenu
            Exit
        }
    }
    if ($script:DataHashTable.config.outputDevice.$Setting) {
        $script:DataHashTable.config.outputDevice.$Setting = $false
    } else {
        $script:DataHashTable.config.outputDevice.$Setting = $true
    }
    WriteDataToDisk
    SetOutputDevice
}

function DevicesExcludedFromIssuesReportingMenu { #Displays menu for managing devices exclusion from issues reporting
    WriteTopOfPage "Devices excluded from issues reporting"
    Write-Host "Adding a device to this list will exclude the device from the device issues reporting, `nincluding the sendWebCallOnIssue. The exclusion is per issue category so a device `nthat is exluded for one issue category will be included in the other categories."
    $MenuContents = @()
    $MenuContents += ("What would you like to do?",
                      "Add device to exclusion list", #1
                      "Divider")
    $NbrOfOptions = 1
    $OptionsLookup = [ordered]@{}
    foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
        foreach ($ExcludedDeviceID in $script:DataHashTable.config.excludedDevicesFromIssuesReporting.$IssueCategory) {
            $NbrOfOptions ++
            $MenuContents += ("Remove '$($script:DataHashTable.devices.$ExcludedDeviceID.deviceName)' from $IssueCategory")
            $OptionsLookup.Add("$NbrOfOptions",@{issuesCategory=$IssueCategory
                                                 deviceID=$ExcludedDeviceID})
        }   
    }                 
    if ($NbrOfOptions -gt 1) {
        $MenuContents += ("Divider")
    }
    $MenuContents += ("Remove all devices from the exclusion list", #$NbrOfOptions
                      "Back to configuration settings") #99
    $NbrOfOptions ++            
    $Option = WriteMenuToHost $MenuContents
    switch ($Option) {
        1 {
            WriteTopOfPage "Checking devices for issues"
            Write-Host "Processing..."
            DetectIssuesWithDevices -DeviceIDList $script:DataHashTable.devices.keys -TypeOfList DeviceList -NoSendWebCallFunction
            WriteIssuesListToOutputDevice -ExcludeDevicesMode
            DevicesExcludedFromIssuesReportingMenu
            Break
        }
        $NbrOfOptions {
            foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
                $script:DataHashTable.config.excludedDevicesFromIssuesReporting.$IssueCategory = @()
            }
            WriteDataToDisk
            DevicesExcludedFromIssuesReportingMenu
            Break
        }
        99{
            ConfigMenu
            Break
        }
        default {
            Write-Host
            Write-Host 
            $newArray = @()
            $newArray += $script:DataHashTable.config.excludedDevicesFromIssuesReporting.($OptionsLookup."$Option".issuesCategory) | Where-Object {$_ -ne ($OptionsLookup."$Option".deviceID)}
            if ($newArray) {
                $script:DataHashTable.config.excludedDevicesFromIssuesReporting.($OptionsLookup."$Option".issuesCategory) = $newArray
            } else {
                $script:DataHashTable.config.excludedDevicesFromIssuesReporting.($OptionsLookup."$Option".issuesCategory) = @()
            }
            Write-Host "$($script:DataHashTable.devices.($OptionsLookup."$Option".deviceID).deviceName) has been removed from the exclusion list"
            WriteDataToDisk
            Start-Sleep 3
            DevicesExcludedFromIssuesReportingMenu
            Break
        }
    }
}

function SendWebCallOnIssueMainMenu { #Displays menu for configuring SendWebCallOnIssue feature 
    $CategoryName = "sendWebCallOnIssue"
    WriteTopOfPage ("$CategoryName - main menu")
    $MenuContents = @()
    if ($script:DataHashTable.config.sendWebCallOnIssue.status -eq "Enabled")  {
        $MenuContents += ("$CategoryName is currently enabled with the following settings.`nWhat would like to do?")
        $NbrOfOptions = 0
        $OptionsLookup = [ordered]@{}
        if ($script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURLStatus -eq "Disabled"){
            $NbrOfOptions = 1
            $MenuContents += ("Enable use of a generic web call URL for all issues with no specific URL configured (only where the web call for the category of issue has been enabled)") #1
        } else {
            $NbrOfOptions = 2
            $MenuContents += ("Disable use of the generic web call URL for all issues with no specific URL configured", #1
                              "Set generic web call URL ('$($script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURL)')") #2
        }
        
        $MenuContents += ("Divider",
                          "Disable $CategoryName", #$NbrOfOptions - ($script:DataHashTable.config.sendWebCallOnIssue.categories).count -1
                          "Reset $CategoryName to the default setting ('$($script:DefaultSendWebCallOnIssueStatus)')", #$NbrOfOptions - ($script:DataHashTable.config.sendWebCallOnIssue.categories).count
                          "Divider")
        $NbrOfOptions += 2
        foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
            $NbrOfOptions ++
            $MenuContents += ("Configure settings for '$IssueCategory' ($($script:DataHashTable.config.sendWebCallOnIssue.$IssueCategory.status))")
            $OptionsLookup.Add("$NbrOfOptions",$IssueCategory)
        }
        $MenuContents += ("Back to configuration settings") #99
    
        
        $Option = WriteMenuToHost $MenuContents
        switch ($Option) {
            ($NbrOfOptions - ($script:DataHashTable.config.sendWebCallOnIssue.categories).count -1) {
                $script:DataHashTable.config.sendWebCallOnIssue.status = "Disabled"
                foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
                    $script:DataHashTable.config.sendWebCallOnIssue.$IssueCategory.status = "Disabled"
                }
                WriteDataToDisk
                SendWebCallOnIssueMainMenu
                Break
            }
            ($NbrOfOptions - ($script:DataHashTable.config.sendWebCallOnIssue.categories).count) {
                $script:DataHashTable.config.sendWebCallOnIssue.status = $script:DefaultSendWebCallOnIssueStatus
                if ($script:DefaultSendWebCallOnIssueStatus -eq "Disabled") {
                    foreach ($IssueCategory in $script:DataHashTable.config.sendWebCallOnIssue.categories) {
                        $script:DataHashTable.config.sendWebCallOnIssue.$IssueCategory.status = "Disabled"
                    }
                }
                WriteDataToDisk
                SendWebCallOnIssueMainMenu
                Break
            }
            1 {
                if ($script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURLStatus -eq "Disabled") {
                    $script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURLStatus = "Enabled"
                } else {
                    $script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURLStatus = "Disabled"
                }
                WriteDataToDisk
                SendWebCallOnIssueMainMenu
                Break
            }
            2 {
                Write-Host
                Write-Host "The current generic web call URL is set to:"
                Write-Host $script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURL
                Write-Host
                Write-Host "To exit without making any changes, just hit ENTER."
                Write-Host "To erase the current URL, type 'erase' and hit ENTER."
                Write-Host "To enter a new URL, type the URL in full and hit ENTER."
                Write-Host
                Write-Host "New URL: " -NoNewline
                $Response = Read-Host
                if ($Response.ToLower() -eq "erase") {
                    $script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURL = ""
                    WriteDataToDisk
                    Write-Host
                    Write-Host
                    Write-Host "Generic URL has been erased"
                    Start-Sleep 3
                } elseif ($Response -ne "") {
                    $script:DataHashTable.config.sendWebCallOnIssue.globalWebCallURL = $Response
                    WriteDataToDisk
                    Write-Host
                    Write-Host
                    Write-Host "Generic URL has been set to $Response"
                    Start-Sleep 3
                }
                SendWebCallOnIssueMainMenu
                Break
            }
            99 {
                ConfigMenu
                Break
            }
            default {
                SendWebCallOnIssueSubMenu -Category ($OptionsLookup.$Option)
                Break
            }
        }
    } else {
        $MenuContents = ("$CategoryName is currently disabled.`nWhat would like to do?",
                        "Enable $CategoryName", #1                   
                        "Reset $CategoryName to the default setting ('$($script:DefaultSendWebCallOnIssueStatus)')", #2
                        "Back to configuration settings" #99
        )
        $Option = WriteMenuToHost $MenuContents
        switch ($Option) {
            1 {
                $script:DataHashTable.config.sendWebCallOnIssue.status = "Enabled"
                WriteDataToDisk
                SendWebCallOnIssueMainMenu
                Break
            }
            2 {
                $script:DataHashTable.config.sendWebCallOnIssue.status = $script:DefaultSendWebCallOnIssueStatus
                WriteDataToDisk
                SendWebCallOnIssueMainMenu
                Break
            }
            99 {
                ConfigMenu
                Break
            }
        }
    }
}

function SendWebCallOnIssueSubMenu { #Displays menu for configuring SendWebCallOnIssue feature 
    param (
        $Category = ""
    )
    
    if ($Category -ne "") {
        $CategoryName = "sendWebCallOnIssue - $Category"
        WriteTopOfPage $CategoryName
        $MenuContents = @()
        if ($script:DataHashTable.config.sendWebCallOnIssue.$Category.status -eq "Enabled") {
            $MenuContents += ("$CategoryName is currently enabled with the following settings.`nWhat would like to do?")
            $NbrOfOptions = 0
            if ($script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURLStatus -eq "Enabled"){
                $NbrOfOptions = 2
                $MenuContents += ("Disable use of the web call URL for the $Category category", #1
                                  "Set web call URL ('$($script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURL)')") #2
            } else {
                $NbrOfOptions = 1
                $MenuContents += ("Enable use of a web call URL for the $Category category") #1
            }
            $MenuContents += ("Divider",
                              "Disable $CategoryName", #NbrOfOptions -1
                              "Reset $CategoryName to the default setting ('$($script:DefaultSendWebCallOnIssueStatus)')", #NbrOfOptions
                              "Go back") #99
            $NbrOfOptions = $NbrOfOptions +2
            $Option = WriteMenuToHost $MenuContents
            switch ($Option) {
                ($NbrOfOptions - 1) {
                    $script:DataHashTable.config.sendWebCallOnIssue.$Category.status = "Disabled"
                    $script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURLStatus = "Disabled"
                    WriteDataToDisk
                    SendWebCallOnIssueSubMenu -Category $Category
                    Break
                }
                $NbrOfOptions {
                    $script:DataHashTable.config.sendWebCallOnIssue.$Category.status = $script:DefaultSendWebCallOnIssueStatus
                    if ($script:DefaultSendWebCallOnIssueStatus -eq "Disabled") {
                        $script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURLStatus = "Disabled"
                    }
                    WriteDataToDisk
                    SendWebCallOnIssueSubMenu -Category $Category
                    Break
                }
                1 {
                    if ($script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURLStatus -eq "Enabled") {
                        $script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURLStatus = "Disabled"
                    } else {
                        $script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURLStatus = "Enabled"
                    }
                    WriteDataToDisk
                    SendWebCallOnIssueSubMenu -Category $Category
                    Break
                }
                2 {
                    Write-Host
                    Write-Host "The current web call URL is set to:"
                    Write-Host $script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURL
                    Write-Host
                    Write-Host "To exit without making any changes, just hit ENTER."
                    Write-Host "To erase the current URL, type 'erase' and hit ENTER."
                    Write-Host "To enter a new URL, type the URL in full and hit ENTER."
                    Write-Host
                    Write-Host "New URL: " -NoNewline
                    $Response = Read-Host
                    if ($Response.ToLower() -eq "erase") {
                        $script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURL = ""
                        WriteDataToDisk
                        Write-Host
                        Write-Host
                        Write-Host "Generic URL has been erased"
                        Start-Sleep 3
                    } elseif ($Response -ne "") {
                        $script:DataHashTable.config.sendWebCallOnIssue.$Category.webCallURL = $Response
                        WriteDataToDisk
                        Write-Host
                        Write-Host
                        Write-Host "Generic URL has been set to $Response"
                        Start-Sleep 3
                    }
                    SendWebCallOnIssueSubMenu -Category $Category
                    Break
                }
                99 {
                    SendWebCallOnIssueMainMenu
                    Break
                }
                default {
                    SendWebCallOnIssueSubMenu -Category $Category
                    Break
                }
            }
        } else {
            $MenuContents += ("$CategoryName is currently disabled.`nWhat would like to do?",
                              "Enable $CategoryName", #1
                              "Reset $CategoryName to the default setting ('$($script:DefaultSendWebCallOnIssueStatus)')", #2
                              "Go back") #99
            $Option = WriteMenuToHost $MenuContents
            switch ($Option) {
                1 {
                    $script:DataHashTable.config.sendWebCallOnIssue.$Category.status = "Enabled"
                    WriteDataToDisk
                    SendWebCallOnIssueSubMenu -Category $Category
                    Break
                }
                2 {
                    $script:DataHashTable.config.sendWebCallOnIssue.$Category.status = $script:DefaultSendWebCallOnIssueStatus
                    WriteDataToDisk
                    SendWebCallOnIssueSubMenu -Category $Category
                    Break
                }
                99 {
                    SendWebCallOnIssueMainMenu
                    Break
                }
                default {
                    SendWebCallOnIssueSubMenu -Category $Category
                    Break
                }
            }
        }
        
        
        
    } else {
        SendWebCallOnIssueMainMenu
    }
}

function ChangeFilePathMenu {
    WriteTopOfPage "Change file names or file paths"
    $MenuContents = (   "The current settings are listed below. `nWhat would you like to do?",
                        "Change 'Device' list HTML path ('$($script:DataHashTable.config.filePath.DeviceListHTML)')", #1
                        "Change 'Device' list CSV path ('$($script:DataHashTable.config.filePath.DeviceListCSV)')", #2
                        "Divider",
                        "Change 'Device Issues' list HTML path ('$($script:DataHashTable.config.filePath.DeviceIssuesListHTML)')", #3
                        "Change 'Device Issues' list CSV path ('$($script:DataHashTable.config.filePath.DeviceIssuesListCSV)')", #4
                        "Divider",
                        "Change 'Hub' list HTML path ('$($script:DataHashTable.config.filePath.HubListHTML)')", #5
                        "Change 'Hub' list CSV path ('$($script:DataHashTable.config.filePath.HubListCSV)')", #6
                        "Divider",
                        "Change 'Hub Mesh Device' list HTML path ('$($script:DataHashTable.config.filePath.HubMeshDeviceListHTML)')", #7
                        "Change 'Hub Mesh Device' list CSV path ('$($script:DataHashTable.config.filePath.HubMeshDeviceListCSV)')", #8
                        "Divider",
                        "Reset all to default settings", #9
                        "Back to configuration settings" #99
    )
    
    $Option = WriteMenuToHost $MenuContents
    switch ($Option) {
        1{
            ChangeFilePathSubMenu @{setting="DeviceListHTML";name="Device list HTML path";defaultSetting=$script:DefaultFilePathDeviceListHTML}
            Break
        }
        2{
            ChangeFilePathSubMenu @{setting="DeviceListCSV";name="Device list CSV path";defaultSetting=$script:DefaultFilePathDeviceListCSV}
            Break
        }
        3{
            ChangeFilePathSubMenu @{setting="DeviceIssuesListHTML";name="Device Issues list HTML path";defaultSetting=$script:DefaultFilePathDeviceIssuesListHTML}
            Break
        }
        4{
            ChangeFilePathSubMenu @{setting="DeviceIssuesListCSV";name="Device Issues list CSV path";defaultSetting=$script:DefaultFilePathDeviceIssuesListCSV}
            Break
        }
        5{
            ChangeFilePathSubMenu @{setting="HubListHTML";name="Hub list HTML path";defaultSetting=$script:DefaultFilePathHubListHTML}
            Break
        }
        6{
            ChangeFilePathSubMenu @{setting="HubListCSV";name="Hub list CSV path";defaultSetting=$script:DefaultFilePathHubListCSV}
            Break
        }
        7{
            ChangeFilePathSubMenu @{setting="HubMeshDeviceListHTML";name="Hub Mesh device list HTML path";defaultSetting=$script:DefaultFilePathHubMeshDeviceListHTML}
            Break
        }
        8{
            ChangeFilePathSubMenu @{setting="HubMeshDeviceListCSV";name="Hub Mesh device list CSV path";defaultSetting=$script:DefaultFilePathHubMeshDeviceListCSV}
            Break
        }
        9{
            $script:DataHashTable.config.filePath = [ordered]@{}
            $script:DataHashTable.config.filePath.Add("DeviceListHTML",$script:DefaultFilePathDeviceListHTML)
            $script:DataHashTable.config.filePath.Add("DeviceListCSV",$script:DefaultFilePathDeviceListCSV)
            $script:DataHashTable.config.filePath.Add("DeviceIssuesListHTML",$script:DefaultFilePathDeviceIssuesListHTML)
            $script:DataHashTable.config.filePath.Add("DeviceIssuesListCSV",$script:DefaultFilePathDeviceIssuesListCSV)
            $script:DataHashTable.config.filePath.Add("HubListHTML",$script:DefaultFilePathHubListHTML)
            $script:DataHashTable.config.filePath.Add("HubListCSV",$script:DefaultFilePathHubListCSV)
            $script:DataHashTable.config.filePath.Add("HubMeshDeviceListHTML",$script:DefaultFilePathHubMeshDeviceListHTML)
            $script:DataHashTable.config.filePath.Add("HubMeshDeviceListCSV",$script:DefaultFilePathHubMeshDeviceListCSV)
            WriteDataToDisk
            ChangeFilePathMenu
            Break
        }
        99{
            ConfigMenu
            Exit
        }
    }
}

function ChangeFilePathSubMenu {
    param (
        $Settings=@{}
    )
    WriteTopOfPage $Settings.name

    Write-Host "Current settings are:"
    Write-Host "File name: $(Split-Path $script:DataHashTable.config.filePath.($Settings.setting) -Leaf)"
    Write-Host "File path: $(Split-Path $script:DataHashTable.config.filePath.($Settings.setting) -Parent)"
    Write-Host
    $MenuContents = (   "What would you like to do?",
                        "Change the file name", #1
                        "Change the file path", #2
                        "Divider",
                        "Reset both settings to default settings", #3
                        "Back to change file names or file paths" #99
    )
    Write-Host
    $Option = WriteMenuToHost $MenuContents
    Write-Host
    Write-Host
    
    switch ($Option) {
        1{
            Write-Host "Enter new file name (include extension): " -NoNewline
            $Response = (Read-Host).trim()
            if ($Response.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -eq -1) { #Checks that there are no invalid characters in the response
                $script:DataHashTable.config.filePath.($Settings.setting) = (Split-Path $script:DataHashTable.config.filePath.($Settings.setting) -Parent) + "\$Response"
                WriteDataToDisk         
            } else {
                Write-Host
                Write-Host "'$Response' is not a valid filename" -ForegroundColor "Red"
                Write-Host
                PauseHEDevicesToolKit
            }
            Break
        }
        2{
            Write-Host "Enter new file path (don't include the filename): " -NoNewline
            $Response = (Read-Host).trim()
            if ($Response.Substring($Response.length-1,1) -eq "\" -OR $Response.Substring($Response.length-1,1) -eq "/") { #If response includes trailing \ or / then remove it
                $Response = $Response.Substring(0,$Response.length-1)
            }
            if ((Test-Path $Response -PathType Container -ErrorAction SilentlyContinue)) { #Checks if the folder path exists
                $script:DataHashTable.config.filePath.($Settings.setting) = "$Response\" + (Split-Path $script:DataHashTable.config.filePath.($Settings.setting) -Leaf)
                WriteDataToDisk         
            } else {
                Write-Host
                Write-Host "'$Response' doesn't exist. Please ensure the path exists and then try again." -ForegroundColor "Red"
                Write-Host
                PauseHEDevicesToolKit
            }
            Break
        }
        3{
            $script:DataHashTable.config.filePath.($Settings.setting) = $Settings.defaultSetting
            WriteDataToDisk
        }
        99{
            ChangeFilePathMenu
            Exit
        }
    }
    ChangeFilePathSubMenu @{setting=$Settings.setting;name=$Settings.name;defaultSetting=$Settings.defaultSetting}
}

function WriteMenuToHost { #Displays a menu with the supplied metadata and validates input before returning the menu option selected
    #$MenuOptions[0] contains the heading of the menu. Any subsequent elements are the menu options
    #If $ExitOption is set to true, the last element in $MenuOptions will be used as an option to exit or back out of the current operation. Option #99 will be used
    param (
        [string[]]$MenuOptions = $null,
        [bool]$ExitOption = $true 
    )
    
    $ValidOptionPicked = $false
    while ($ValidOptionPicked -eq $false) {
        $OptionsCount = $MenuOptions.count -1
        $NbrOfDividersSoFar = 0
        [int[]]$ValidResponses = $null
        $TableWidth = 25 #Set minimum width of the table
        for ($i=1;$i -le $OptionsCount;$i++) { #Check the length of every menu option and adjust the table width to fit
            if ($MenuOptions[$i].length+1 -gt $TableWidth) {
                $TableWidth = $MenuOptions[$i].length+1
            }
        }
        if ($TableWidth -gt 100) { #Set maximum width of the table
            $TableWidth = 100
        }
        Write-Host $MenuOptions[0]
        DrawTableTopBoundary $TableWidth
        Write-Host ("{2}{0} {3} {1}{2}" -f (("  #"),("Options".PadRight($TableWidth," ")),([char]0x2503),([char]0x2502)))
        DrawTableHorizontalDivider $TableWidth
        for ($i=1;$i -le $OptionsCount;$i++) {
            if ($ExitOption -AND $i -eq $OptionsCount) {
                $OptionsCount--
                $ValidResponses += 99
                DrawTableHorizontalDivider $TableWidth
                Write-Host ("{2}{0} {3} {1}{2}" -f ((" 99"),($MenuOptions[$i].PadRight($TableWidth, " ")),([char]0x2503),([char]0x2502)))
            } elseif ($MenuOptions[$i] -eq "Divider"){
                $NbrOfDividersSoFar ++
                Write-Host ("{1}    {2} {0}{1}" -f ((" ".PadRight($TableWidth, " ")),([char]0x2503),([char]0x2502)))
            } else {
                $ValidResponses += $i -$NbrOfDividersSoFar
                $NbrOfOptionsText = ($i -$NbrOfDividersSoFar).ToString()
                do {
                    if ($MenuOptions[$i].length -le $TableWidth-1) {
                        Write-Host ("{2}{0} {3} {1} {2}" -f (($NbrOfOptionsText.padleft(3," ")),($MenuOptions[$i].PadRight($TableWidth-1, " ")),([char]0x2503),([char]0x2502)))
                        $Repeat = $false
                    } else {
                        Write-Host ("{2}{0} {3} {1} {2}" -f (($NbrOfOptionsText.padleft(3," ")),(($MenuOptions[$i]).Substring(0, $TableWidth-2) + "~"),([char]0x2503),([char]0x2502)))
                        $Repeat = $true
                        $MenuOptions[$i] = " ~" + $MenuOptions[$i].Substring($TableWidth-2,$MenuOptions[$i].Length-$TableWidth+2)
                    }
                    $NbrOfOptionsText = ""
                } while ($Repeat)
            }
        }
        DrawTableBottomBoundary $TableWidth
        Write-Host
        Write-Host "Enter option (1" -NoNewline
        if ($OptionsCount-$NbrOfDividersSoFar -gt 1) {
            Write-Host ("-{0}" -f ($OptionsCount-$NbrOfDividersSoFar)) -NoNewline
        }
        if ($ExitOption) {
            Write-Host ",99" -NoNewline
        }
        Write-Host "): " -NoNewline
                
        $Response = Read-Host
        if ($ValidResponses -contains $Response) {
            $ValidOptionPicked = $true
            $Response
        } else {
            Clear-Host
            Write-Host
            Write-Host "$Response is NOT a valid option" -ForegroundColor "Red"
        }
    }
}

function WriteTopOfPage { #Prints the header of the page to the screen
    param (
        $Title,
        $colour = $script:DataHashTable.config.textColour
    )

    if ($script:NonInteractive -eq $false) {
        Clear-Host
    }
    Write-Host
    if ($Title) {
        Write-Host ("**** {0} ****" -f $Title.toUpper()) -NoNewline -ForegroundColor $colour
        if ($Title -like "main menu") {
            Write-Host "           ver $script:Version" -NoNewline -ForegroundColor $colour
        }
        Write-Host
        Write-Host
    }
    Write-Host
}

function DrawTableTopBoundary { #Used for box drawing for the menus on the screen
    param (
        [int]$width=1,
        [string]$colour = $script:DataHashTable.config.textColour
    )
    Write-Host ("{0}" -f [char]0x250F) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2501 * 4) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x252F) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2501 * ($width+1)) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2513) -ForegroundColor $colour
}

function DrawTableHorizontalDivider { #Used for box drawing for the menus on the screen
    param (
        [int]$width=1,
        [string]$colour = $script:DataHashTable.config.textColour
    )
    Write-Host ("{0}" -f [char]0x2520) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2500 * 4) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x253C) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2500 * ($width+1)) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2528) -ForegroundColor $colour
}

function DrawTableBottomBoundary { #Used for box drawing for the menus on the screen
    param (
        [int]$width=1,
        [string]$colour = $script:DataHashTable.config.textColour
    )
    Write-Host ("{0}" -f [char]0x2517) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2501 * 4) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2537) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x2501 * ($width+1)) -NoNewline -ForegroundColor $colour
    Write-Host ("{0}" -f [char]0x251B) -ForegroundColor $colour
}

function SetHTMLHeader { #Returns the HTML code for the beginning of the HTML file
    param (
        $Title = "HTML output"
    )
    $HTML = "<!DOCTYPE html>
            <html lang=`"en`">
            <head>
            <title>$Title</title>
            <meta charset=`"UTF-8`">
            <meta name=`"viewport`" content=`"width=device-width, initial-scale=1`">
            <style>
            `thtml {background-color: #1abc9c;}
            `tbody {background-color: white; font-size: 16px;font-family: Arial, Helvetica, sans-serif;color: #107b68; max-width: 1200px; margin-top: 0px; margin-left: auto; margin-right: auto; margin-top: 0px; margin-bottom: 0px; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 1px; border-radius: 20px;}
            `t.header {padding: 5px; text-align: left; background: #1abc9c; color: white; border-radius: 20px 60px;margin-bottom: 20px;}
            `t.footer {border-top: 2px solid #1abc9c;margin-top: 20px;}
            `t.minimalBottomMargin {margin-bottom: 2px;}
            `t.divider {margin-bottom: 2.5em;}
            `t.secondaryClassification {font-style: italic;}
            `t.red {color:red}
            `t.blue {color:blue}
            `t.blue:hover {color: #107b68}
            `th1 {margin-left: 30px;font-size: 2em;}
            `th4 {font-size: 1em;}
            `th4.listOfDevices {margin-top: 2px;margin-bottom: 2px;}
            `th4.listOfDevicesWarning {margin-top: 2px;margin-bottom: 2px;color: red}
            `th4.deviceIssues {margin-bottom: 0px;}
            `ttable {width: 100%;margin-bottom: 10px;}
            `ttable:hover {background-color: #1abc9c; color: white;}
            `ttd.descriptionColumn {text-align: right;vertical-align: top;}
            `ttd.indentationColumn {width: 3em; text-align: right;}
            `ta {text-decoration: none;color:inherit}
            `ta:active,a:hover {color: #107b68;text-decoration: underline;}
            `tp.inUseBy,p.children {margin-top: 0px; margin-bottom: 0px; padding-left: 0em;}
            `tp.timestamp {font-size: 0.7em;text-align: right;margin-right: 10px}
            </style>
            </head>
            <body>
            <div class=`"header`">
            `t<h1>$Title</h1>
            </div>`n"
    $HTML
}

function SetHTMLFooter { #Returns the HTML code for the end of the HTML file
    $HTML = "<div class=`"footer`">
            `t<p class=`"timestamp`">Generated $(Get-Date -Format $script:DataHashTable.config.dateTimeFormat)</p>
            </div>
            </body>
            </html>"
    $HTML
}

$FilePath = (Split-Path ($MyInvocation.MyCommand.Path) -Parent)
$DataJSONFilePath = $FilePath + "\Data.json"
$HubListFilePath = $FilePath + "\hubs.txt"

$DefaultAutoLaunchWebBrowser = $false
$DefaultDateTimeFormat = "MMM dd, yyyy @ HH:mm"
$DefaultGlobalWebCallURLStatus = "Disabled"
$DefaultFilePathDeviceListHTML = $FilePath + "\DeviceList.html"
$DefaultFilePathDeviceListCSV = $FilePath + "\DeviceList.csv"
$DefaultFilePathDeviceIssuesListHTML = $FilePath + "\DeviceIssuesList.html"
$DefaultFilePathDeviceIssuesListCSV = $FilePath + "\DeviceIssuesList.csv"
$DefaultFilePathHubListHTML = $FilePath + "\HubList.html"
$DefaultFilePathHubListCSV = $FilePath + "\HubList.csv"
$DefaultFilePathHubMeshDeviceListHTML = $FilePath + "\HubMeshDeviceList.html"
$DefaultFilePathHubMeshDeviceListCSV = $FilePath + "\HubMeshDeviceList.csv"
$DefaultInactivityThresholdMinutes = 1440
$DefaultInternetProtocol = "https://"
$DefaultLowBatteryChargeThreshold = 20
$DefaultOutputDevice = [PSCustomObject]@{screen=$true;HTML=$false;CSV=$false}
$DefaultSendWebCallOnIssueStatus = "Disabled"
$DefaultTextColour = "White"
$ValidColours = "Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White"
$IterationArray = ("one","two","three","four","five")

$DataHashTable = [ordered]@{}
$DevicesWithIssuesFound = @{}
$DevicesWithIssuesHidden = @{}
$PSDefaultParameterValues = @{}
    
EraseConfigurationSettingsAndSetDefaults

#C# class to create callback to ignore SSL certificate issues
$Ccode = @"
public class SSLHandler
{
    public static System.Net.Security.RemoteCertificateValidationCallback GetSSLHandler()
    {

        return new System.Net.Security.RemoteCertificateValidationCallback((sender, certificate, chain, policyErrors) => { return true; });
    }
    
}
"@
#compile the class
Add-Type -TypeDefinition $Ccode

if (-NOT (Test-Path $HubListFilePath -PathType Leaf -ErrorAction SilentlyContinue)) { #Checks if the hubs.txt file doesn't exists
    $hubsTXTContents = "#Enter the IP addresses of each Hubitat Elevation hub to be scanned below, one per line. A comment may be added to each line, as long as the comment is placed after the IP address and is preceded by `"#`".`n#Example (remove the first `"#`"):`n#192.168.1.2`n#192.168.1.87 #Garage hub"
    $hubsTXTContents | Out-File $HubListFilePath -Force
}

Clear-Host
if ($NonInteractive) {
    WriteTopOfPage "Non-interactive mode"
    Write-Host "--- ver $Version ---"
    Write-Host "* Checking if configuration data exists on disk..."
    if (Get-ChildItem $script:DataJSONFilePath -ErrorAction SilentlyContinue) {
        Write-Host "* Configuration file found. Reading file..."
        ReadConfigFromDisk
        Write-Host "* Configuration from disk has been applied."
    } else {
        Write-Host "* No configuration settings found on disk. Using default settings."
        if ($RunNewScan -eq $false) {
            Write-Host "* No data file detected and a new scan has not been requested by using the switch '-RunNewScan'.`n     There's no point continuing at this point unless a new scan is requested.`n     Hit ENTER to run a new scan, otherwise CTRL+C to terminate the program." -NoNewline -ForegroundColor "Yellow"
            Read-Host
            $RunNewScan = $true
        }
    }
    Write-Host ("* Output device(s): $(($script:DataHashTable.config.outputDevice.psobject.properties | Where-Object {$_.Value} | ForEach-Object {$_.Name}) -join ", ")")
    if ($RunNewScan -eq $false) {
        Write-Host "* Using saved device data."
        if ($script:DataHashTable.config.deviceListLastUpdated) {#There appears to be valid device data in the data file
            Write-Host "* Device data appears to be valid"
        } else { #No deviceListLastUpdated value detected which indicates that a device scan has never taken place. There is most likely no device data in this file
            Write-Host "     However, there doesn't appear to be any valid device data in the file.`n" -ForegroundColor "Yellow"
            $MenuContents = (   "  What would you like to do?",
                                "Use the saved data anyway", #1
                                "Run a new scan" #2
            )
            $Option = WriteMenuToHost $MenuContents -ExitOption $false
            if ($Option -eq 2) {
                $RunNewScan = $true
            }
        }
    }
    if ($RunNewScan) {
        Write-Host "* A new scan has been requested."
        if ($HubIPAddress) {
            Write-Host "* Starting a new scan."
            ScanForData -HubIPAddressList ($HubIPAddress -split ",")
        } else {
            Write-Host "* No HubIPAddress supplied. Checking disk to see if '$(Split-Path ($script:HubListFilePath) -Leaf)' has any addresses in it."
            if (ReadHubsTXTListFromDisk) {
                ScanForData -HubIPAddressList (ReadHubsTXTListFromDisk)
            } else {
                Write-Host "* No IP addresses for Hubitat Elevation hubs found in '$script:HubListFilePath'."
                Write-Host "* Unable to find a source of hub IP addresses. Press ENTER to enter interactive mode." -NoNewline -ForegroundColor "Red"
                Read-Host
                $NonInteractive = $false
                if ($script:DataHashTable.config.deviceListLastUpdated) {
                    WriteTopOfPage "Reading data from disk..."
                    ReadDataFromDisk
                }
                RunANewScanMenu 
                Break               
            }
        }
    } else {
        Write-Host "* Using saved data.`n* Reading device data..."
        ReadDataFromDisk
        Write-Host "* Device data loaded."
    }
    if ($script:DataHashTable.config.deviceListLastUpdated) {
        Write-Host "* Valid data appears to have been loaded."
    } else {
        Write-Host "* No valid device data found. Press ENTER to enter interactive mode." -NoNewline -ForegroundColor "Red"
        Read-Host
        $NonInteractive = $false
        RunANewScanMenu
        Break
    }
    
    if ($ListAllDevices) {
        Write-Host "* Listing all devices."
        WriteListOfDevicesToOutputDevice -ShowHierarchy
        Write-Host
    }
    if ($ListAllHubs){
        Write-Host "* Listing all hubs."
        WriteHubListToOutputDevice
        Write-Host
    }
    if ($ListAllHubMeshDevices){
        Write-Host "* Listing all Hub Mesh devices."
        WriteListOfDevicesToOutputDevice -ShowHubMeshHierarchy
        Write-Host
    }
    if ($CheckForDeviceIssues){
        Write-Host "* Checking for device issues."
        DetectIssuesWithDevices -DeviceIDList $script:DataHashTable.devices.keys -TypeOfList DeviceList
        WriteIssuesListToOutputDevice
        Write-Host
    }
    if ($SearchForDeviceByName){
        Write-Host "* Initiating searching for device by name."
        if (-NOT $SearchTerm) {
            Write-Host "* No SearchTerm provided. Enter a search term now: " -NoNewline -ForegroundColor "Yellow"
            $SearchTerm = Read-Host
        }
        Write-Host "* Searching for $SearchTerm..."
        $SearchTerm = "*" + $SearchTerm.Trim().ToLower() + "*"
        WriteListOfDevicesToOutputDevice ($script:DataHashTable.devices.keys | Where-Object {$script:DataHashTable.devices.$_.deviceName -like $SearchTerm})
        Write-Host
    }
    if ($SearchForHubMeshDeviceByName){
        Write-Host "* Initiating searching for Hub Mesh device by name."
        if (-NOT $SearchTerm) {
            Write-Host "* No SearchTerm provided. Enter a search term now: " -NoNewline -ForegroundColor "Yellow"
            $SearchTerm = Read-Host
        }
        Write-Host "* Searching for $SearchTerm..."
        $SearchTerm = "*" + $SearchTerm.Trim().ToLower() + "*"
        WriteListOfDevicesToOutputDevice ($script:DataHashTable.devices.keys | Where-Object {$script:DataHashTable.devices.$_.deviceName -like $SearchTerm})  -ShowHubMeshHierarchy
        Write-Host
    }
    
    Write-Host "Finished. Have a good day! "
    if ($DataHashTable.config.outputDevice.screen) { #If screen is being used as an output device, don't automatically terminate the program
        Write-Host "Press ENTER to terminate the program" -NoNewline
        Read-Host
    } else {
        Write-Host "(Terminating automatically in 10 seconds)"
        Start-Sleep 10
    }
} else {
    InitialiseDB
    MainMenu
}
