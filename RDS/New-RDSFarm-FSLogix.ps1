<#Author       : Chris Williams
# Creation Date: 4-21-2020
# Usage        : Deploy RDS Farm with FSLogix Profile Containers

#********************************************************************************
# Date                         Version      Changes
#------------------------------------------------------------------------
# 04-21-2020                     1.0        Initial Version
# 07-23-2020                     1.1        Replaced set-acl with icacls
# 07-28-2020                     1.2        FSLogix Install
# 12-01-2020                     1.3        Group Creation, fixed SMB share creation, SSL deploy
# 12-02-2020                     1.4        Add OneDrive and Office admx file import
# 01-21-2021                     1.5        Add Script block for Certify and ADPasswordProtection Script
# 03-16-2021                     1.6        Add Chrome and admx file import
# 06-03-2021                     1.7        Optimized and added Compact-UPD task
# 07-22-2022                     1.8        Updated FSLogix Download URI
#
#********************************************************************************
# Run from the Primary Domain Controller
#>

### Current SPLA agreement is 7206943
$rdsCB                 = 'Compass-GTWY'
$rdsWAS                = 'Compass-GTWY'
$rdsSH1                = 'Compass-RDS1'
$rdsSH2                = 'Compass-RDS2'
$rdsGTWY               = 'Compass-GTWY'
$rdsLS                 = 'Compass-GTWY'
$rdsFILE               = 'Compass-GTWY'
$internalFQDN          = 'Compass.local'
$externalFQDN          = 'Compass.compassmsp.com'
$collectionName        = 'RDS Collection'
$collectionDescription = 'RDS Collection Description'

$rdsLSAD = $rdsLS + "$"
$rdsCB   = $rdsCB   + "." + $internalFQDN
$rdsWAS  = $rdsWAS  + "." + $internalFQDN
$rdsSH1  = $rdsSH1  + "." + $internalFQDN
$rdsSH2  = $rdsSH2  + "." + $internalFQDN
$rdsLS   = $rdsLS   + "." + $internalFQDN
$rdsGTWY = $rdsGTWY + "." + $internalFQDN

### Installs/Configures ADPasswordProtection
$keyPath = 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Main'
if (!(Test-Path $keyPath)) { New-Item $keyPath -Force | Out-Null }
Set-ItemProperty -Path $keyPath -Name "DisableFirstRunCustomize" -Value 1

Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Install-ADPasswordProtection.ps1'); Install-ADPasswordProtection -StoreFilesInDBFormatLink 'https://rmm.compassmsp.com/softwarepackages/ADPasswordProtectionStore.zip' -NotificationEmail 'ADPassProtectionNotifications@compassmsp.com' -SMTPRelay 'compassmsp-com.mail.protection.outlook.com' -FromEmail 'ADPasswordNotifications@compassmsp.com'

### Imports required modules
Import-Module -Name RemoteDesktop, BitsTransfer

### Looks for RDS Users Group creates if needed
try{
    Get-ADGroup Sec_RDS_Users -ErrorAction Stop
}
catch{
    New-ADGroup -Name Sec_RDS_Users -GroupCategory Security -GroupScope Global
}

### 
try{
    Get-ADGroup "Terminal Server License Servers"  -ErrorAction Stop
    ADD-ADGroupMember "Terminal Server License Servers" –members $rdsLSAD
}    
catch{ 
    Write-Output "Terminal Server License Servers Group could not be found." 
} 

### Formats and Configures FSLogix Disk on File Server ### NOTE DISK MUST BE PRESENT ###
Invoke-Command -ComputerName "$rdsFILE" -ScriptBlock {

    $keyPath = 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Main'
    if (!(Test-Path $keyPath)) { New-Item $keyPath -Force | Out-Null }
    Set-ItemProperty -Path $keyPath -Name "DisableFirstRunCustomize" -Value 1

    $FSLDiskLabel = 'FSLogixDisks'

    $AllDiskInfo = Get-Disk | Select-Object Number,OperationalStatus,@{Name="Size";Expression={$_.size/1GB}}

    $AllDiskInfo | Out-String
    Write-Host 'All Data on this drive will be deleted' -ForegroundColor Red

    $DiskNumber = Read-Host "Enter disk number to initialize and format"

    $WorkingDisk = Get-Disk -Number $DiskNumber

    #Set the disk to Online
    if ($WorkingDisk.OperationalStatus -eq 'Offline'){
        Set-Disk -Number $DiskNumber -IsOffline $false
    }

    #Check to see if any partitions exist. Exit if there are active partitions
    if (($WorkingDisk | Get-Partition).count -gt 0){
        Write-Output 'The selected disk already contains a partition. As a safety measure, the script only works if the disk does not contain any partitions'
        Exit
    }
    else{
        #Clear the disk if it was previously initialized
        if ($WorkingDisk.PartitionStyle -ne 'RAW'){
            Clear-Disk -Number $DiskNumber -RemoveData -Confirm:$false
        }

        Initialize-Disk -Number $DiskNumber -PartitionStyle GPT

        #Create a partition using the next available drive letter
        $NewPartition = New-Partition -DiskNumber $DiskNumber -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem ReFS -AllocationUnitSize 64KB -NewFileSystemLabel $FSLDiskLabel
    }

    #Get Path to new folder
    $FolderDirectory = "$($NewPartition.DriveLetter):\$($FSLDiskLabel)"

    #Create Folder
    New-Item -Path $FolderDirectory -ItemType Directory

    #Clear all Explicit Permissions on the folder
    ICACLS ("$FolderDirectory") /reset

    #Add CREATOR OWNER permission
    ICACLS ("$FolderDirectory") /grant ("CREATOR OWNER" + ':(OI)(CI)(IO)F')

    #Add SYSTEM permission
    ICACLS ("$FolderDirectory") /grant ("SYSTEM" + ':(OI)(CI)F')

    #Give Domain Admins Full Control
    ICACLS ("$FolderDirectory") /grant ("Domain Admins" + ':(OI)(CI)F')

    #Apply Create Folder/Append Data, List Folder/Read Data, Read Attributes, Traverse Folder/Execute File, Read permissions to this folder only. Synchronize is required in order for the permissions to work
    ICACLS ("$FolderDirectory") /grant ("Sec_RDS_Users" + ':(AD,REA,RA,X,RC,RD,S)')

    #Disable Inheritance on the Folder. This is done last to avoid permission errors.
    ICACLS ("$FolderDirectory") /inheritance:r

    #Create SmbShare
    New-SmbShare -Path $FolderDirectory -CachingMode None -Name $FSLDiskLabel -FullAccess @("Sec_RDS_Users", 'Domain Admins') -FolderEnumerationMode Unrestricted

    mkdir c:\scripts

    (New-Object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/RDS/FSLogix/Compact-UPDs.ps1', 'c:\scripts\Compact-UPDs.ps1')

    $action = New-ScheduledTaskAction -Execute 'powershell' -Argument '-File C:\Scripts\Compact-UPDs.ps1'
    $trigger = New-ScheduledTaskTrigger -Daily -At 4am
    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "Compact UPDs" -Description "Compact FSLogix Profile Disk"
    }

Write-Output "Created FSLogix Profile Share on: $rdsFILE. Please edit C:\Scripts\Compact-UPDs.ps1 with path."
Start-Process \\$rdsFILE\c$\scripts\Compact-UPDs.ps1

### Add RD Web Access, RD Broker and RD Host Servers
New-RDSessionDeployment -ConnectionBroker $rdsCB `
                        -SessionHost @($rdsSH1,$rdsSH2) `
                        -WebAccessServer $rdsWAS

### Add RD Licensing Server and a RD Gateway Server
Add-RDServer -Server $rdsLS `
             -Role RDS-Licensing `
             -ConnectionBroker $rdsCB

Add-RDServer -Server $rdsGTWY `
             -Role RDS-Gateway `
             -GatewayExternalFqdn $externalFQDN `
             -ConnectionBroker $rdsCB

### Create Collection
New-RDSessionCollection -CollectionName $collectionName `
                        -CollectionDescription $collectionDescription `
                        -SessionHost @($rdsSH1,$rdsSH2) `
                        -ConnectionBroker $rdsCB

### Set RD License Mode
Set-RDLicenseConfiguration -LicenseServer $rdsLS `
                           -Mode PerUser `
                           -ConnectionBroker $rdsCB

### Configure RD Collection
Set-RDSessionCollectionConfiguration -CollectionName $collectionName `
                                     -UserGroup "Sec_RDS_Users" `
                                     -EncryptionLevel ClientCompatible `
                                     -IdleSessionLimitMin 180 `
                                     -DisconnectedSessionLimitMin 180 `
                                     -AuthenticateUsingNLA 0 `
                                     -SecurityLayer SSL `
                                     -ConnectionBroker $rdsCB

### Download and install Certify /w RDGCert deploy script ### Manual Configuration of Certify Required after install
Invoke-Command -ComputerName "$rdsGTWY" -ScriptBlock {
        $keyPath = 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Main'
        if (!(Test-Path $keyPath)) { New-Item $keyPath -Force | Out-Null }
        Set-ItemProperty -Path $keyPath -Name "DisableFirstRunCustomize" -Value 1
        mkdir c:\BIN\CertRenew
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $LatestVersionUrl = (Invoke-WebRequest https://certifytheweb.com/home/download -MaximumRedirection 0).Links | Where-Object {$_.innerText -eq "click here"} | Select-Object -expand href
        (New-Object System.Net.WebClient).DownloadFile("$LatestVersionUrl", 'C:\Windows\temp\CertifyTheWebSetup.exe')
        (New-Object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/LetsEncrypt/Install-LeRdgCertificate.ps1', 'c:\BIN\CertRenew\Install-LeRdgCertificate.ps1')
        $certify_deploy_status = Start-Process -FilePath 'C:\Windows\temp\CertifyTheWebSetup.exe' -ArgumentList "/verysilent" -Wait -Passthru
        (New-Object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/FSLogix/Invoke-FslShrinkDisk/master/Invoke-FslShrinkDisk.ps1', 'c:\Scripts\Invoke-FslShrinkDisk.ps1')
    }

Write-Output "Certify has been installed on: $rdsGTWYT. Please login to the server and configure the SSL before deployment."

### Download, installs, and configure FSLogix on Session Host 1 and 2
Invoke-Command -ComputerName "$rdsSH1", "$rdsSH2" -ScriptBlock {
    
    $keyPath = 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Main'
    if (!(Test-Path $keyPath)) { New-Item $keyPath -Force | Out-Null }
    Set-ItemProperty -Path $keyPath -Name "DisableFirstRunCustomize" -Value 1

    reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy" /v DeleteUserAppContainersOnLogoff /t REG_DWORD /d 1 /f
    [string]$XMLTask = (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/RDS/FSLogix/2019SearchFix/2019ResetSearchOnLogoff.xml') | Out-String
    Register-ScheduledTask -XML $XMLTask -TaskName 'Reset Search on Logoff'

    ## Install FSLogix
    $FSLogixAppsSetupURI = "https://aka.ms/fslogix/download"
    $FSLogixAppsSetup = "FSLogixAppsSetup.zip"
    (New-Object System.Net.WebClient).DownloadFile("$FSLogixAppsSetupURI","C:\windows\temp\$FSLogixAppsSetup")
    Expand-Archive -LiteralPath 'C:\Windows\temp\FSLogixAppsSetup.zip' -DestinationPath 'C:\Windows\temp\FSLogix' -Force -Verbose
    Start-Process -FilePath 'C:\Windows\temp\FSLogix\x64\Release\FSLogixAppsSetup.exe' -ArgumentList "/install /quiet /norestart" -Wait -Passthru
    
    ## Install Chrome
    $ChromeURI = "https://dl.google.com/tag/s/appguid%253D%257B8A69D345-D564-463C-AFF1-A69D9E530F96%257D%2526iid%253D%257BBEF3DB5A-5C0B-4098-B932-87EC614379B7%257D%2526lang%253Den%2526browser%253D4%2526usagestats%253D1%2526appname%253DGoogle%252520Chrome%2526needsadmin%253Dtrue%2526ap%253Dx64-stable-statsdef_1%2526brand%253DGCEB/dl/chrome/install/GoogleChromeEnterpriseBundle64.zip?_ga%3D2.8891187.708273100.1528207374-1188218225.1527264447"
    (New-Object System.Net.WebClient).DownloadFile("$ChromeURI","C:\windows\temp\GoogleChromeEnterpriseBundle64.zip")
    Expand-Archive -LiteralPath 'C:\Windows\temp\GoogleChromeEnterpriseBundle64.zip' -DestinationPath 'C:\Windows\temp\GoogleChromeEnterpriseBundle64' -Force -Verbose
    Start-Process -FilePath 'C:\Windows\temp\GoogleChromeEnterpriseBundle64\Installers\GoogleChromeStandaloneEnterprise64.msi' -ArgumentList "/quiet /m" -Wait -Passthru
    ## Install 7Zip
    $7ZipURI = "https://www.7-zip.org/a/7z1900-x64.msi"
    $7Zip = "7z1900-x64.msi"
    (New-Object System.Net.WebClient).DownloadFile("$7ZipURI","C:\windows\temp\$7Zip")
    Start-Process -FilePath C:\windows\temp\$7Zip -Wait -ArgumentList "/q";
    ## Install Notepad++
    $nppURI = "https://sourceforge.net/projects/notepadmsi/files/latest/download"
    $npp = "Notepad++7_9_1.msi"
    (New-Object System.Net.WebClient).DownloadFile("$nppURI","C:\windows\temp\$npp")
    Start-Process -FilePath C:\windows\temp\$npp -Wait -ArgumentList "/q";
    ## Install Firefox
    $FirefoxURI = "https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US"
    $Firefox = "Firefox.msi"
    (New-Object System.Net.WebClient).DownloadFile("$FirefoxURI","C:\windows\temp\$Firefox")
    Start-Process -FilePath C:\windows\temp\$Firefox -Wait -ArgumentList "/q";
}

### Downloads Office Admin Templates
Start-BitsTransfer -Source "https://download.microsoft.com/download/2/E/E/2EEEC938-C014-419D-BB4B-D184871450F1/admintemplates_x64_5098-1000_en-us.exe" -Destination "C:\Windows\temp\admintemplates_x64_5098-1000_en-us.exe"
Start-Process -FilePath "C:\Windows\temp\admintemplates_x64_5098-1000_en-us.exe" -ArgumentList "/extract:C:\Windows\temp\office_admx /passive /quiet" -Wait -Passthru

### Copies GPO Templates to Policies SYSVOLShare
Copy-Item "\\$rdsSH1\c$\Windows\temp\FSLogix\fslogix.admx" -Destination C:\Windows\PolicyDefinitions
Copy-Item "\\$rdsSH1\c$\Windows\temp\FSLogix\fslogix.adml" -Destination C:\Windows\PolicyDefinitions\en-US
Copy-Item "\\$rdsSH1\c$\Program Files (x86)\Microsoft OneDrive\*\adm\OneDrive.admx" -Destination C:\Windows\PolicyDefinitions
Copy-Item "\\$rdsSH1\c$\Program Files (x86)\Microsoft OneDrive\*\adm\OneDrive.adml" -Destination C:\Windows\PolicyDefinitions\en-US
Copy-Item "\\$rdsSH1\c$\Windows\temp\GoogleChromeEnterpriseBundle64\Configuration\*.admx" C:\Windows\PolicyDefinitions
Copy-Item "\\$rdsSH1\c$\Windows\temp\GoogleChromeEnterpriseBundle64\Configuration\en-US\*.adml" C:\Windows\PolicyDefinitions\en-US
Copy-Item "C:\Windows\temp\office_admx\admx\*.admx" C:\Windows\PolicyDefinitions
Copy-Item "C:\Windows\temp\office_admx\admx\en-us\*.adml" C:\Windows\PolicyDefinitions\en-US

### Install HTML5 on Web Access Server
#Enter-PSSession $rdsWAS
#Install-Module -Name PowerShellGet -Force -Confirm:$False
#Install-Module -Name RDWebClientManagement -Confirm:$False
#Import-RDWebClientBrokerCert C:\Windows\temp\broker.cer
#Install-RDWebClientPackage
#Publish-RDWebClientPackage -Type Production -Latest
#Exit-PSSession

### https://RDS-WEB/RDWeb/WebClient/Index.html
