#Script to update / create Asset in SnipeIT from NCentral AMP or Standalone Powershell (Replacing Recon and Marksman)
#Automation Policy Version
#v1.3 - 17/06/2022 - Darren Banfi
#1.3 -- Added IP to Location Hash Table
#1.2 --Fixing all missing dependancys PackageManagement / NuGet / PSGallery
#1.1 --Intial Release
#Uncomment the next line for Standalone-mode - Leave Commited for Scripting in NCentral AMP
#Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force

#User Settings Required for running
#URL of your SnipeIT Installation
$SnipeITURL = ''
#APIKey for SnipeIT
$APIKey = ''
#Location Lookup Table - IP Address, SnipeIT Location ID
$LocationTable = @{}
$LocationTable.add( '194.0.0.0', 1 <# Location 1 #>)
$LocationTable.add( '81.0.0.0', 2 <#  Location 2 #>)


#Lets get Todays Date and Make it in a Friendly SnipeIT Format
$Date = (Get-Date -DisplayHint Date).ToString("yyyy-MM-dd")
#Lets setup some logging for debugging - Set up Temp Folder if it's missing
$LogFolderName ="C:\Temp\"
if (Test-Path $LogFolderName) {} else { New-Item $LogFolderName -ItemType Directory}
$LogFile = "C:\Temp\log-$(Get-Date -Format 'yyMMdd-HHmmss').log"
Write-Output $Date "Starting the Log-Process" | Out-File -FilePath $LogFile -Append

#Lets enable TLS1.2 for this session - This could be missing on systems that have not used Powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$LogData = [Net.ServicePointManager]::SecurityProtocol
Write-Output $Date "TLS Protocol has been set to" $LogData | Out-File -FilePath $LogFile -Append
# Install PackageManagement first so we can use NuGet and PSGallery (Missing on devices with PS3.0)
Install-Module -Name PackageManagement -Force -ErrorAction SilentlyContinue | Out-Null

# Install NuGet package repository so we can use the PSGallery
Install-PackageProvider "NuGet" -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null

# Check to see if SnipeitPS is there, if not go to the PSGallery and get a copy
If(-not(Get-InstalledModule SnipeitPS -ErrorAction SilentlyContinue))
{
   Write-Output $Date "Trying to install SnipeitPS Module" | Out-File -FilePath $LogFile -Append
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
$SerialNumber = (Get-WmiObject -class win32_bios).SerialNumber
#If the Serial Number is Generic in the BIOS - Change the Serial number to the Motherboard number
if ($SerialNumber -ceq "To Be Filled By O.E.M." -or $SerialNumber -ceq "System Serial Number") { $SerialNumber = $cimserial.SerialNumber}
#$SerialNumber = 12344321
$BuilderName = "NCentral Update"

#Search the computer for Details to update WMIQuerys Setup
$wmios = Get-WmiObject -Class Win32_OperatingSystem
$wmics = Get-WmiObject -Class Win32_ComputerSystem
$wmina = Get-WmiObject -Class Win32_NetworkAdapter
$wmibios = Get-WmiObject -Class Win32_BIOS
$wmicpu = Get-WmiObject -Class Win32_Processor
$wmiqfe = Get-WmiObject -Class Win32_QuickFixEngineering

#Values returned from WMIQuery
$OSCaption = $wmios.Caption
$OSVersion = $wmios.Version
#BuildDate needs converted to SnipeIT Friendly Format
$BuildDate = $wmios.InstallDate
$BuildDate = $BuildDate.Substring(0,8)
$BuildDateConverted = [datetime]::ParseExact($BuildDate, 'yyyyMMdd', $null).ToString("yyyy-MM-dd")
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
$BIOSDate = $BIOSDate.Substring(0,8)
$BIOSDateConverted = [datetime]::ParseExact($BIOSDate, 'yyyyMMdd', $null).ToString("yyyy-MM-dd")
#MACAddress can be multiple entrys - so picking the first unique one
$MACAddress = $wmina.MACAddress | Select-Object -Unique -Index 1
#Check with have a valid MACAddress - 17 Digits with : - if not just blank the MAC
if ($MACAddress.Length -lt 17) { $MACAddress = "00:00:00:00:00:00"}
$CPUName = $wmicpu.Name
#InstalledOn returns multiple entrys, picking the first unique one from the list and then convert the date
$InstalledOn = $wmiqfe.InstalledOn | Select-Object -Unique -Index 0
$InstalledOnConverted = $InstalledOn.ToString("yyyy-MM-dd")
#ExternalIP Address Check with Location Check
$ExternalIP = (Invoke-WebRequest -uri "http://ifconfig.me/ip" -UseBasicParsing).Content
If ($null -ne $ExternalIP) { $LocationCheck = $LocationTable[$ExternalIP]}
If ($null -ne $LocationCheck) { 
   $params = @{
      rtd_location_id = $LocationCheck
   }
}

#Now to invoke the SnipeIT Connection and see if we can find the Asset
Write-Output  $Date " -- Searching for asset" | Out-File -FilePath $LogFile -Append
$CheckAsset = Get-SnipeitAsset -serial $SerialNumber 
#Check for Duplicated AssetName - If So, Create a new Asset with .Duplicated on the end
if ($null -ne $CheckAsset){
#Update Asset with Details
Write-Output  $Date $CheckAsset " -- Asset has been found in SnipeIT" | Out-File -FilePath $LogFile -Append
Get-SnipeitAsset -serial $SerialNumber | Set-SnipeitAsset -Name $AssetTag @params -customfields @{ "_snipeit_agent_name_8" = $BuilderName; "_snipeit_total_memory_size_mb_21" = $MemorySize; "_snipeit_os_type_19" = $OSCaption + " " + $OSVersion;  "_snipeit_last_online_16" = $Date;  "_snipeit_bios_updated_26" = $BIOSDateConverted;  "_snipeit_build_date_7" = $BuildDateConverted;  "_snipeit_windows_update_32" = $InstalledOnConverted; "_snipeit_mac_address_1" = $MACAddress; "_snipeit_domain_name_31" = $Domain; "_snipeit_cpu_20" = $CPUName; "_snipeit_ip_address_30" = $ExternalIP; }
 }
 else
 {
#Create a new asset
Write-Output  $Date " -- Asset has not been found - trying to create a new one" | Out-File -FilePath $LogFile -Append
#Check for the ID for the model
$id = Get-SnipeitModel -search $Model -all | Select-Object -ExpandProperty id
    #If Model is not found in SnipeIT - Lets try and create a new one
 if ($null -eq $id) { 

     Write-Output  $Date "Model was not found - Trying to create a new one" | Out-File -FilePath $LogFile -Append
      #Check the Manufacturer is there
     $ManufacturerID = Get-SnipeitManufacturer -search $Manufacturer | Select-Object -ExpandProperty id
     if ($null -ne $ManufacturerID) {  $ManufacturerID = $ManufacturerID.Item(0) }
     else
     {
        #If this is a brand new manufacturer - lets add them to the SnipeIT Database
        Write-Output  $Date $Manufacturer "Manufacturer was not found - Trying to create a new one" | Out-File -FilePath $LogFile -Append
        New-SnipeitManufacturer -name $Manufacturer
        $ManufacturerID = Get-SnipeitManufacturer -search $Manufacturer | Select-Object -ExpandProperty id
        $ManufacturerID = $ManufacturerID.Item(0)
     }
    
     New-SnipeitModel -name $Model -manufacturer_id $ManufacturerID -fieldset_id 3 -category_id 5
     Write-Output  $Date $Manufacturer $Model "Has been selected" | Out-File -FilePath $LogFile -Append
     $id = Get-SnipeitModel -search $Model -all | Select-Object -ExpandProperty id

  }
  #Check for a ducplicate Asset Name in SnipeIT, if found add the .Duplicate tag on to it (TBC and email someone)
 $CheckAssetHost = Get-SnipeitAsset -asset_tag $AssetTag
 if ($null -eq $Domain) { $Domain = "No-DomainName"}
        if ($null -ne $CheckAssetHost) {$AssetTag = $AssetTag + $Domain + ".Check" }
 New-SnipeitAsset -status_id 9 -model_id $id -asset_tag $AssetTag -serial $SerialNumber -company_id 1 @params
 Write-Output  $Date $AssetTag "New Asset Created in SnipeIT" | Out-File -FilePath $LogFile -Append
 Get-SnipeitAsset -serial $SerialNumber | Set-SnipeitAsset -Name $AssetTag -customfields @{ "_snipeit_agent_name_8" = $BuilderName; "_snipeit_total_memory_size_mb_21" = $MemorySize; "_snipeit_os_type_19" = $OSCaption + " " + $OSVersion;  "_snipeit_last_online_16" = $Date;  "_snipeit_bios_updated_26" = $BIOSDateConverted;  "_snipeit_build_date_7" = $BuildDateConverted;  "_snipeit_windows_update_32" = $InstalledOnConverted; "_snipeit_mac_address_1" = $MACAddress; "_snipeit_domain_name_31" = $Domain; "_snipeit_cpu_20" = $CPUName; "_snipeit_ip_address_30" = $ExternalIP; }
 }

 Write-Output  $Date " -- All Done - Closing" | Out-File -FilePath $LogFile -Append

Exit

