#This script is used to determine if a computer can have WMF upgraded to 5.1
#https://docs.microsoft.com/en-us/powershell/scripting/wmf/whats-new/compatibility?view=powershell-7
$WMFShouldBeUpdated = $true

if ($PSVersionTable.PSVersion -gt [version]('{0}.{1}.{2}.{3}' -f '5.1.0.0'.split('.'))) {
    Write-Output 'PS 5.1 is already installed'
    $WMFShouldBeUpdated = $false
    exit
}

#Check to see if Exchange is installed
if ((Get-PSSnapin -Registered | Select-Object -ExpandProperty name) -match 'Microsoft.Exchange.Management.PowerShell') {
    Write-Output 'Exchange is installed. Do not upgrade PowerShell'
    $WMFShouldBeUpdated = $false
    exit
}

# Check if .Net 4.5 or above is installed

$release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Release -ErrorAction SilentlyContinue -ErrorVariable evRelease).release
$installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Install -ErrorAction SilentlyContinue -ErrorVariable evInstalled).install

if (($installed -ne 1) -or ($release -lt 378389)) {
    Write-Output "Setup requires .Net 4.5."
    $WMFShouldBeUpdated = $false
    exit
}

#Check if Sharepoint is installed
#This check is not perfect since I do not have a server to test with
if (Test-Path "$env:ProgramFiles\Common Files\Microsoft Shared\Web Server Extensions"){
    Write-Output "WMF Should not be upgraded when Sharepoint is installed"
    $WMFShouldBeUpdated = $false
    exit
}

#Check if Lync is installed
#This check has not been tested since I do not have a server to test with
if (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Real-Time Communications\{A593FD00-64F1-4288-A6F4-E699ED9DCA35}'){
    Write-Output "WMF Should not be upgraded when Lync is installed"
    $WMFShouldBeUpdated = $false
    exit
}

#Check if System Center is installed
#This check is not perfect since I do not have a server to test with
if ((Test-Path "$env:ProgramFiles\Microsoft System Center 2012") -or (Test-Path "$env:ProgramFiles\Microsoft System Center 2012 R2") -or (Test-Path "$env:ProgramFiles\Microsoft Configuration Manager")) {
    Write-Output "WMF Should not be upgraded when System Center is installed"
    $WMFShouldBeUpdated = $false
    exit
}

#If no compatibility issues found
if ($WMFShouldBeUpdated){
    Return 'InstallWMF5.1'
}