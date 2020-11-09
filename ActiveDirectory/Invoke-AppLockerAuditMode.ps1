<#
This script will deploy AppLocker in Audit mode to the root of a domain.



TO DO*************
Check if there is currently an applocker policy in place


Andy Morales
#>
#requires -Modules ActiveDirectory


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

    if (Test-Path -Path $OutPath){
        Remove-Item $OutPath -Recurse -Force
    }

    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFile, $OutPath)
}

$ZipPath = 'C:\windows\temp\AppLockerRDSAuditOnly.zip'
$GPOFolder = $ZipPath.Replace('.zip', '')

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

(New-Object System.Net.WebClient).DownloadFile('https://github.com/CompassMSP/PublicScripts/raw/master/ActiveDirectory/GPOBackups/AppLocker%20RDS%20AUDIT%20ONLY.zip', $ZipPath)

Expand-ZIP -ZipFile $ZipPath -OutPath $GPOFolder

Import-Module ActiveDirectory

$GPOReportPath = Get-ChildItem $GPOFolder -Recurse | Where-Object name -EQ gpreport.xml

#Get the Name of the GPO from the content of the XML
[XML]$GPOReportXML = Get-Content -Path $GPOReportPath.FullName
[string]$GPOBackupName = $GPOReportXML.GPO.Name
$GPOPrefixedName = "_$GPOBackupName"

$GPOContentsFolder = (Get-ChildItem $GPOFolder).fullname

New-GPO -Name $GPOPrefixedName -ErrorAction SilentlyContinue
Import-GPO -Path $GPOContentsFolder -TargetName $GPOPrefixedName -BackupGpoName $GPOBackupName -ErrorAction Stop

New-GPLink -Name $GPOPrefixedName -Target "$((Get-ADDomain).DistinguishedName)" -LinkEnabled Yes -ErrorAction Stop



"Successfully imported GPO $GPOPrefixedName" | Out-File $ErrorLogLocation -Append
$SuccessfullyImportedGPOs += $GPOPrefixedName