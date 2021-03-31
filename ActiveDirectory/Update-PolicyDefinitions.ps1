
<#
Downloads the latest ADMX files and copies them to the PDC sysvol share. 

Andy Morales
#>
function Expand-ZIP {
    <#
    Extracts a ZIP file to a directory. The contents of the destination will be deleted if they already exist.

    Andy Morales
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [String]$ZipFile,

        [parameter(Mandatory = $true)]
        [String]$OutPath
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -Path $OutPath) {
        Remove-Item $OutPath -Recurse -Force
    }

    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFile, $OutPath)
}

#Check for PDC
if ((Get-WmiObject Win32_ComputerSystem).domainRole -eq 5){
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    (New-Object System.Net.WebClient).DownloadFile('https://github.com/CompassMSP/PublicScripts/raw/master/ActiveDirectory/PolicyDefinitions.zip','C:\Windows\Temp\PolicyDefinitions.zip')

    Expand-ZIP -ZipFile 'C:\Windows\Temp\PolicyDefinitions.zip' -OutPath 'C:\Windows\Temp\PolicyDefinitions'

    if(Test-Path 'C:\Windows\SYSVOL\domain\Policies'){
        ROBOCOPY "C:\Windows\Temp\PolicyDefinitions\PolicyDefinitions" "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" /R:0 /W:0 /E /xo /dcopy:t /MT:32 /np
    }
    elseif(Test-Path 'C:\Windows\SYSVOL_DFSR\domain\Policies'){
        ROBOCOPY "C:\Windows\Temp\PolicyDefinitions\PolicyDefinitions" "C:\Windows\SYSVOL_DFSR\domain\Policies\PolicyDefinitions" /R:0 /W:0 /E /xo /dcopy:t /MT:32 /np
    }

    Remove-Item -Path 'C:\Windows\Temp\PolicyDefinitions.zip', 'C:\Windows\Temp\PolicyDefinitions' -Force -Recurse
}