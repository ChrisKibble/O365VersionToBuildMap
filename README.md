# O365VersionToBuildMap

PowerShell Script to Capture Office 365 Build to Version Mappings.

The purpose of this function is to go to the Microsoft support sites and download the Version and Build information.  It will then map it into an array that can be referenced for other things or exported to a CSV.

## Example Function Calls

* Get-Office365BuildToVersionMap | Sort-Object versionNumber
* Export-CSV $env:temp\Office365BuildToVersionMap.csv -NoTypeInformation

## Example Output

BuildNumber|VersionNumber |Source
---|---|---
1509|16.0.4229.1024|https://docs.microsoft.com/en-us/officeupdates/monthly-channel-2015
1509|16.0.4229.1024|https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-targeted-2015
1509|16.0.4229.1024|https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-2015
...|...|...
1806|16.0.10228.20080|https://docs.microsoft.com/en-us/officeupdates/monthly-channel-2018
1806|16.0.10228.20104|https://docs.microsoft.com/en-us/officeupdates/monthly-channel-2018