# SnipeIT-PS-Updater
PowerShell to Update SnipeIT

Script to update / create Asset in SnipeIT from NCentral AMP or Standalone Powershell (Replacing Recon and Marksman)
Automation Policy Version
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

#Uncomment this line for Standalone-mode - Leave Commited for Scripting in NCentral AMP
#Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force
