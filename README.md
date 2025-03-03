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
    The default is "http://", but can be changed to "https://"
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
