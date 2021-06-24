<#

This script adds an IP printer to windows. Run as administrator/SYSTEM for best results.

.PARAMETER PrinterName
Name of the printer

.PARAMETER IPAddress
IP Address/hostname of printer

.PARAMETER DriverDownloadLink
ZIP file that contains the driver

.PARAMETER DriverInfFolder
Location where the .inf file is located

TO FIND: Extract the zip and find the location (usually under an x64 folder)

.PARAMETER DriverName
Name of the driver that will be used

TO FIND: Open the .inf file and find the exact driver name

.PARAMETER DriverInfPath

Location in windows where the .inf is located

TO FIND:

Run the pnputil.exe command
Go to "C:\Windows\System32\DriverStore\FileRepository" and sort by newest folder.
The .inf will usually be in there

.LINK
https://www.pdq.com/blog/using-powershell-to-install-printers/

Andy Morales
#>
#Requires -RunAsAdministrator

$PrinterName = 'TASKalfa 6003i'
$IPAddress = '10.8.11.239'
$DriverDownloadLink = 'https://cdn.kyostatics.net/dlc/eu/driver/all/kx702415_upd_signed.-downloadcenteritem-Single-File.downloadcenteritem.tmp/KX_Universal_Pr...nter_Driver.zip'
$DriverInfFolder = 'C:\Windows\Temp\PrintDriver\Kx_8.1.1109_UPD_Signed_EU\en\64bit\*.inf'
$DriverName = 'Kyocera TASKalfa 6003i KX'
$DriverInfPath = 'C:\Windows\System32\DriverStore\FileRepository\oemsetup.inf_amd64_6bff917e8a9060a5\OEMSETUP.INF'


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

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region driverInstall
Write-Output "Downloading driver from $($DriverDownloadLink)"

(New-Object System.Net.WebClient).DownloadFile($DriverDownloadLink , 'C:\Windows\Temp\PrintDriver.zip')

Expand-ZIP -ZipFile 'C:\Windows\Temp\Driver.zip' -OutPath 'C:\Windows\Temp\PrintDriver'

Write-Output "Installing driver from $($DriverInfFolder)"

pnputil.exe /a $DriverInfFolder

Add-PrinterDriver -Name $DriverName -InfPath $DriverInfPath

$PortName = $IPAddress

$CurrentPorts = @(Get-PrinterPort | Where-Object { $_.Name -eq $IPAddress })

if ($CurrentPorts.count -eq 1) {
    if ($CurrentPorts.PrinterHostAddress -eq $IPAddress -or $CurrentPorts.PrinterHostIP -eq $IPAddress) {
        Write-Output 'Port already exists.'
    }
    else {
        $PortName = $PortName + '-' + (Get-Random -Maximum 99)
        Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress
    }
}
else{
    Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress
}

Add-Printer -DriverName $DriverName -Name $PrinterName -PortName $PortName