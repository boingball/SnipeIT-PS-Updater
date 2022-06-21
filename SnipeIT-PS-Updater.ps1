#Script to update / create Asset in SnipeIT from NCentral AMP or Standalone Powershell (Replacing Recon and Marksman)
#Automation Policy Version
#v1.7 - 21/06/2022 - Darren Banfi
#v1.7 - 21/06/22 - Laptop or Desktop Check for Model Creation, EXO5 Service Check
#1.6 -- 20/06/22 - Added AV Status, Office and Windows Activation Check, Laptop or Desktop
#1.5 - 20/06/22 - Added Logging Turn On and Off, Get Current Logged in User and GeoLocation Lookup
#1.4 -- 20/06/22 - Converted WMI to CIM Querys - Added Correct MAC Address Lookup and Extended Details of Interface Adapter that is connected
#1.3 -- 17/06/22 - Added IP to Location Hash Table
#Req: Powershell v5.1 to run correctly

#CustomFields - These will need changed to fit your custom fields if required on Line 213 and 248 - If not required take them out of the Statement
# "_snipeit_agent_name_8" = $BuilderName; "Sends N-Central Agent"
# "_snipeit_total_memory_size_mb_21" = $MemorySize; - Total Memory Size
# "_snipeit_os_type_19" = $OSCaption + " " + $OSVersion; - OS Version
# "_snipeit_last_online_16" = $Date; - Todays Date
# "_snipeit_bios_updated_26" = $BIOSDateConverted; - BIOS Updated Date 
# "_snipeit_build_date_7" = $BuildDateConverted;  - Windows Install Date
# "_snipeit_windows_update_32" = $InstalledOnConverted;  - Windows Updates last installed
# "_snipeit_mac_address_1" = $NetworkMACAddress;  - Active Network MAC Address
# "_snipeit_domain_name_31" = $Domain; - Domain Name
# "_snipeit_cpu_20" = $CPUName;  - CPU Name 
# "_snipeit_ip_address_30" = $ExternalIP;  - External IP of computers running this script
# "_snipeit_network_link_speed_39" = $NetworkLinkSpeed;  - Network Link Speed
# "_snipeit_network_adapter_name_40" = $NetworkIF;  - Network Adapter Name
# "_snipeit_current_logged_in_user_41" = $LoggedInUser; - Current Logged in User
# "_snipeit_geo_location_42" = $GeoLocation;  - GeoLocation from IP
# "_snipeit_isp_name_43" = $ISPName; - ISP Name
# "_snipeit_av_25" = $AVStatus; - AV Status
# "_snipeit_office_activated_and_updated_3" = $OfficeProduct; - Office Version / Activation Check (Supports 2013 and Office 365, Office 2016/2019/2022 have the same version number as Office 365)
# "_snipeit_windows_activated_4" = $WindowsActivation; - Windows Activation Check
# "_snipeit_exo5_5" = $EXO5Result; - EXO5 Service Running

#Uncomment the next line for Standalone-mode - Leave Commited for Scripting in NCentral AMP
#Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force

#User Settings Required for running
#URL of your SnipeIT Installation
$SnipeITURL = ''
#APIKey for SnipeIT
$APIKey = ''
#Turn on Logging - This creates a logfile in the C:\Temp Folder by default - $true to turn on / $false to turn off
$LoggingRequired = $false
#GeoLocation Vendor - List of Vendors here - https://github.com/public-apis/public-apis#geocoding - Tested with IPGeo and ip-api.com
$GeoLocationVendor = "https://api.techniknews.net/ipgeo/"

#Location Lookup Table - IP Address, SnipeIT Location ID
$LocationTable = @{}
$LocationTable.add( '194.0.0.0', 1 <# Location 1 #>)
$LocationTable.add( '81.0.0.0', 2 <#  Location 2 #>)

#Lets get Todays Date and Make it in a Friendly SnipeIT Format
$Date = (Get-Date -DisplayHint Date).ToString("yyyy-MM-dd")
#Lets setup some logging - if LoggingRequired = $true
#Set up Temp Folder if it's missing
if($true -eq $LoggingRequired) {
$LogFolderName ="C:\Temp\"
if (Test-Path $LogFolderName) {} else { New-Item $LogFolderName -ItemType Directory}
$LogFile = "C:\Temp\log-$(Get-Date -Format 'yyMMdd-HHmmss').log"
Write-Output $Date "Starting the Log-Process" | Out-File -FilePath $LogFile -Append
}

#Lets enable TLS1.2 for this session - This could be missing on systems that have not used Powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$LogData = [Net.ServicePointManager]::SecurityProtocol
if($true -eq $LoggingRequired){ Write-Output $Date "TLS Protocol has been set to" $LogData | Out-File -FilePath $LogFile -Append }
# Install PackageManagement first so we can use NuGet and PSGallery (Missing on devices with PS3.0)
Install-Module -Name PackageManagement -Force -ErrorAction SilentlyContinue | Out-Null

# Install NuGet package repository so we can use the PSGallery
Install-PackageProvider "NuGet" -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null

# Check to see if SnipeitPS is there, if not go to the PSGallery and get a copy
If(-not(Get-InstalledModule SnipeitPS -ErrorAction SilentlyContinue))
{
    if($true -eq $LoggingRequired){Write-Output $Date "Trying to install SnipeitPS Module" | Out-File -FilePath $LogFile -Append}
   Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted 
   Install-Module -Name SnipeitPS
}

# We are ready to import the SnipeitPS to run the rest of the script
Import-Module -Name SnipeitPS

# Lets make some logging - This creates a file in C:\temp\
Connect-SnipeitPS -URL $SnipeITURL -apiKey $APIKey

#Search for Asset
$AssetTag = hostname
$cimserial = Get-CimInstance -Class Win32_BaseBoard
$SerialNumber = (Get-CimInstance -Class win32_bios).SerialNumber
#If the Serial Number is Generic in the BIOS - Change the Serial number to the Motherboard number
if ($SerialNumber -ceq "To Be Filled By O.E.M." -or $SerialNumber -ceq "System Serial Number") { $SerialNumber = $cimserial.SerialNumber}
$BuilderName = "NCentral Update"
#Test Serial
#$SerialNumber = "1234432112"

#Search the computer for Details to update WMIQuerys Setup
$wmios = Get-CimInstance -Class Win32_OperatingSystem
$wmics = Get-CimInstance -Class Win32_ComputerSystem
$wmina = Get-CimInstance -Class Win32_NetworkAdapter
$wmibios = Get-CimInstance -Class Win32_BIOS
$wmicpu = Get-CimInstance -Class Win32_Processor
$wmiqfe = Get-CimInstance -Class Win32_QuickFixEngineering

#Values returned from WMIQuery
$OSCaption = $wmios.Caption
$OSVersion = $wmios.Version
#BuildDate needs converted to SnipeIT Friendly Format
$BuildDate = $wmios.InstallDate
#$BuildDateConverted = [datetime]::ParseExact($BuildDate, 'dd/MM/yyyy hh:mm:ss', $null).ToString("yyyy-MM-dd")
$BuildDateConverted = $BuildDate.ToString("yyyy-MM-dd")
#MemorySize is truncated down 3 decimal places
$MemorySize = ($wmios.TotalVisibleMemorySize).ToString()
$MemorySize = $MemorySize.Substring(0, $MemorySize.Length -3)
#ComputerName should be the same as hostname
$ComputerName = $wmics.Name
$Manufacturer = $wmics.Manufacturer
#Model is Truncated down by 1 place due to trailing space
$Model = $wmics.Model
$Model = $Model.Substring(0, $Model.Length -1)
$Domain = $wmics.Domain
$SystemSKUNumber = $wmics.SystemSKUNumber
#BIOSDate needs converted to a correct format
$BIOSDate = $wmibios.ReleaseDate
#$BIOSDateConverted = [datetime]::ParseExact($BIOSDate, 'MM/dd/yyyy hh:mm:ss', $null).ToString("yyyy-MM-dd")
$BIOSDateConverted = $BuildDateConverted = $BIOSDate.ToString("yyyy-MM-dd")
#MACAddress - Different way to pull this due to the inconsitancy of the WMI Query - Now finds the Network Adapter with the status Up
$NETWORKDetails = Get-NetAdapter | Select-Object InterfaceDescription, Status, MacAddress, LinkSpeed -Unique | Where-Object{$_.Status -like "*up" -and $_.InterfaceDescription -notlike "*Hyper*" -and $_.InterfaceDescription -notlike "*VPN*"}
$NetworkMACAddress = $NETWORKDetails.MacAddress
$NetworkMACAddress = $NetworkMACAddress -replace '-',':'
#Check with have a valid MACAddress - 17 Digits with : - if not just blank the MAC
if ($MACAddress.Length -lt 17) { $MACAddress = "00:00:00:00:00:00"}
#NetworkLink Speed
$NetworkLinkSpeed = $NETWORKDetails.LinkSpeed
#NetworkInterfaceName
$NetworkIF = $NETWORKDetails.InterfaceDescription
#Current Logged in User
$LoggedInUser = (Get-CimInstance -ClassName Win32_ComputerSystem).Username
if ($null -eq $LoggedInUser){$LoggedInUser = "No Logged in User detected"}
#AntiVirus-Status - Checks for which products are activated in SecurityCenter
$AntiVirusStatus = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object displayName,timestamp,productstate -Last 1 | Where-Object {$_.productState -eq "266240"}
If($AntiVirusStatus.displayName -like "*Kaspersky*" -or $AntiVirusStatus.displayName -like "*LogMeIn*" -or $AntiVirusStatus.displayName -like "*BitDefender*" -and $AntiVirusStatus.productstate -eq "266240") { 
    $AVStatus = $AntiVirusStatus.displayName 
} else {
    $AVStatus = "AntiVirus Not Found"
}
#Office Activation Check
$OfficeActivationCheck = Get-CimInstance SoftwareLicensingProduct| Where-Object {$_.name -like "*office*" -and $_.licensestatus -eq "1"}|Select-Object name,licensestatus -First 1
if($OfficeActivationCheck.name -like "*Office 16*") { $OfficeProduct = "Office 365 - Activated"}
if($OfficeActivationCheck.name -like "*Office 15*") { $OfficeProduct = "Office 2013 - Activated"}
if($OfficeActivationCheck.name -like "*Office 14*") { $OfficeProduct = "Office 2010 - Activated"}

#Windows Activation Check
$WindowsActivationCheck = Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" | Where-Object { $_.PartialProductKey } | Select-Object Description, LicenseStatus
If ($WindowsActivationCheck.LicenseStatus -like "*1*"){ $WindowsActivation = "Activated"} else { $WindowsActivation = "Not Activated"}

#CPU Name
$CPUName = $wmicpu.Name
#InstalledOn returns multiple entrys, picking the first unique one from the list and then convert the date
$InstalledOn = $wmiqfe.InstalledOn | Select-Object -Unique -Index 0
$InstalledOnConverted = $InstalledOn.ToString("yyyy-MM-dd")
#ExternalIP Address Check with Location Check
$ExternalIP = (Invoke-WebRequest -uri "http://ifconfig.me/ip" -UseBasicParsing).Content
#Check against LocationTable and Set Location if avilable
If ($ExternalIP.Contains(".")) { $LocationCheck = $LocationTable[$ExternalIP]}
If ($null -ne $LocationCheck) { 
   $params = @{
      rtd_location_id = $LocationCheck
   }
} else {
    $params = @{}
}

#GeoLocationCheck
$GeoLocationCheck = Invoke-RestMethod -Method Get -Uri "$GeoLocationVendor$ExternalIP"
if($null -eq $GeoLocationCheck) {

    $GeoLocationCheck = "Not Found";
    $ISPName = "Not Found";

} else {
    $GeoLocation = $GeoLocationCheck.city + "," + $GeoLocationCheck.regionName + "," + $GeoLocationCheck.country + "," + $GeoLocationCheck.zip
    $ISPName = $GeoLocationCheck.isp
}
#Laptop or Desktop?
$isLaptop = $false
$isDesktop = $true
$CategoryID = 5 #Desktop ID
if(Get-WmiObject -Class win32_systemenclosure | Where-Object { $_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 14}) { 
    $isLaptop = $true 
    $isDesktop = $false
    $CategoryID = 2 #Laptop ID
}
 if(Get-WmiObject -Class win32_battery) {
     $isLaptop = $true
     $isDesktop = $false
     $CategoryID = 2 #Laptop ID
    }

#EXO5 Service Check
$EXO5Status = Get-Service "wpagnt" -ErrorAction SilentlyContinue
if($EXO5Status.Status -like "*Running*"){
  $EXO5Result = "Installed and Running"
} else {
  $EXO5Result = "Not Found"
}
#Now to invoke the SnipeIT Connection and see if we can find the Asset
if($true -eq $LoggingRequired){Write-Output  $Date " -- Searching for asset" | Out-File -FilePath $LogFile -Append}
$CheckAsset = Get-SnipeitAsset -serial $SerialNumber 
#Check for Duplicated AssetName - If So, Create a new Asset with .Duplicated on the end
if ($null -ne $CheckAsset){
#Update Asset with Details
if($true -eq $LoggingRequired){Write-Output  $Date $CheckAsset " -- Asset has been found in SnipeIT" | Out-File -FilePath $LogFile -Append}
Get-SnipeitAsset -serial $SerialNumber | Set-SnipeitAsset -Name $AssetTag @params -customfields @{ "_snipeit_agent_name_8" = $BuilderName; "_snipeit_total_memory_size_mb_21" = $MemorySize; "_snipeit_os_type_19" = $OSCaption + " " + $OSVersion;  "_snipeit_last_online_16" = $Date;  "_snipeit_bios_updated_26" = $BIOSDateConverted;  "_snipeit_build_date_7" = $BuildDateConverted;  "_snipeit_windows_update_32" = $InstalledOnConverted; "_snipeit_mac_address_1" = $NetworkMACAddress; "_snipeit_domain_name_31" = $Domain; "_snipeit_cpu_20" = $CPUName; "_snipeit_ip_address_30" = $ExternalIP; "_snipeit_network_link_speed_39" = $NetworkLinkSpeed; "_snipeit_network_adapter_name_40" = $NetworkIF; "_snipeit_current_logged_in_user_41" = $LoggedInUser; "_snipeit_geo_location_42" = $GeoLocation; "_snipeit_isp_name_43" = $ISPName; "_snipeit_av_25" = $AVStatus; "_snipeit_office_activated_and_updated_3" = $OfficeProduct; "_snipeit_windows_activated_4" = $WindowsActivation; "_snipeit_exo5_5" = $EXO5Result; }
 }
 else
 {
#Create a new asset
if($true -eq $LoggingRequired){Write-Output  $Date " -- Asset has not been found - trying to create a new one" | Out-File -FilePath $LogFile -Append}
#Check for the ID for the model
$id = Get-SnipeitModel -search $Model -all | Select-Object -ExpandProperty id
    #If Model is not found in SnipeIT - Lets try and create a new one
 if ($null -eq $id) { 

    if($true -eq $LoggingRequired){Write-Output  $Date "Model was not found - Trying to create a new one" | Out-File -FilePath $LogFile -Append}
      #Check the Manufacturer is there
     $ManufacturerID = Get-SnipeitManufacturer -search $Manufacturer | Select-Object -ExpandProperty id
     if ($null -ne $ManufacturerID) {  $ManufacturerID = $ManufacturerID.Item(0) }
     else
     {
        #If this is a brand new manufacturer - lets add them to the SnipeIT Database
        if($true -eq $LoggingRequired){Write-Output  $Date $Manufacturer "Manufacturer was not found - Trying to create a new one" | Out-File -FilePath $LogFile -Append}
        New-SnipeitManufacturer -name $Manufacturer
        $ManufacturerID = Get-SnipeitManufacturer -search $Manufacturer | Select-Object -ExpandProperty id
        $ManufacturerID = $ManufacturerID.Item(0)
     }
    
     New-SnipeitModel -name $Model -manufacturer_id $ManufacturerID -fieldset_id 3 -category_id $CategoryID
     if($true -eq $LoggingRequired){Write-Output  $Date $Manufacturer $Model "Has been selected" | Out-File -FilePath $LogFile -Append}
     $id = Get-SnipeitModel -search $Model -all | Select-Object -ExpandProperty id

  }
  #Check for a ducplicate Asset Name in SnipeIT, if found add the .Duplicate tag on to it (TBC and email someone)
 $CheckAssetHost = Get-SnipeitAsset -asset_tag $AssetTag
 if ($null -eq $Domain) { $Domain = "No-DomainName"}
        if ($null -ne $CheckAssetHost) {$AssetTag = $AssetTag + $Domain + ".Check" }
 New-SnipeitAsset -status_id 9 -model_id $id -asset_tag $AssetTag -serial $SerialNumber -company_id 1 @params
 if($true -eq $LoggingRequired){Write-Output  $Date $AssetTag "New Asset Created in SnipeIT" | Out-File -FilePath $LogFile -Append}
 Get-SnipeitAsset -serial $SerialNumber | Set-SnipeitAsset -Name $AssetTag -customfields @{ "_snipeit_agent_name_8" = $BuilderName; "_snipeit_total_memory_size_mb_21" = $MemorySize; "_snipeit_os_type_19" = $OSCaption + " " + $OSVersion;  "_snipeit_last_online_16" = $Date;  "_snipeit_bios_updated_26" = $BIOSDateConverted;  "_snipeit_build_date_7" = $BuildDateConverted;  "_snipeit_windows_update_32" = $InstalledOnConverted; "_snipeit_mac_address_1" = $NetworkMACAddress; "_snipeit_domain_name_31" = $Domain; "_snipeit_cpu_20" = $CPUName; "_snipeit_ip_address_30" = $ExternalIP; "_snipeit_network_link_speed_39" = $NetworkLinkSpeed; "_snipeit_network_adapter_name_40" = $NetworkIF; "_snipeit_current_logged_in_user_41" = $LoggedInUser; "_snipeit_geo_location_42" = $GeoLocation; "_snipeit_isp_name_43" = $ISPName;  "_snipeit_av_25" = $AVStatus; "_snipeit_office_activated_and_updated_3" = $OfficeProduct; "_snipeit_windows_activated_4" = $WindowsActivation; "_snipeit_exo5_5" = $EXO5Result; }
 }

 if($true -eq $LoggingRequired){Write-Output  $Date " -- All Done - Closing" | Out-File -FilePath $LogFile -Append}

Exit
