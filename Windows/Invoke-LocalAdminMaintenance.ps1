<#
This script must execute as a 64 bit PowerShell command. Make sure that the RMM command is set to execute as x64.

This script will do the following:
    Create Local admin account
    Makes sure it is a member of local admins
    Make sure it is enabled
    Do nothing if LAPS is enabled
    Reset the local admin password

Andy Morales
#>
$LocalAdminAccountName = 'CMSP'

#The "Administrator" account should not be added to the array since it will not be properly removed from the administrators group
$LocalAccountsToDisable = @(
    'WDLocal'
)

$EventLogSourceName = 'LocalAdminMaintScript'

#Password Generate Function
function Get-RandomCharacters {
    #https://activedirectoryfaq.com/2017/08/creating-individual-random-passwords/
    $RandomPassword = ''
    $Length = '12'

    $LowerCaseChars = 'abcdefghkmnoprstuvwxyz'
    $UpperCaseChars = 'ABCDEFGHKMNPRSTUVWXYZ'
    $NumberChars = '23456789'
    $SpecialChars = '@#$%-+*_=?:<>^&'
    $AllCharacters = $LowerCaseChars + $UpperCaseChars + $NumberChars + $SpecialChars

    #generate random strings until they contain one of each character type. Put a safety on the loop so it only runs 10 times
    DO{
        $random = 1..$length | ForEach-Object { Get-Random -Maximum $AllCharacters.length }
        $private:ofs = ""

        $RandomCharacters = [String]$AllCharacters[$random]

        $count++
    }Until (($RandomCharacters -cmatch '(?=.*[a-z])(?=.*[A-Z])(?=.*[@#$%-+*_=?:<>^&])(?=.*\d)') -or ($count -ge 10))

    return $RandomCharacters
}

function Test-RegistryValue {
    #Modified version of the function below
    #https://www.jonathanmedd.net/2014/02/testing-for-the-presence-of-a-registry-key-and-value.html
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
            Position = 1,
            HelpMessage = 'Registry::HKEY_LOCAL_MACHINE\SYSTEM')]
        [ValidatePattern('Registry::.*')]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [parameter(Mandatory = $true,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [parameter(Position = 3)]
        [ValidateNotNullOrEmpty()]$ValueData
    )
    try {
        if ($ValueData) {
            if ((Get-ItemProperty -Path $Path -ErrorAction Stop | Select-Object -ExpandProperty $Name -ErrorAction Stop) -eq $ValueData) {
                return $true
            }
            else {
                return $false
            }
        }
        else {
            Get-ItemProperty -Path $Path -ErrorAction Stop | Select-Object -ExpandProperty $Name -ErrorAction Stop | Out-Null
            return $true
        }
    }
    catch {
        return $false
    }
}

#Create Event Log source if it does not exist
if ([System.Diagnostics.EventLog]::SourceExists("$EventLogSourceName") -eq $false) {
    [System.Diagnostics.EventLog]::CreateEventSource("$EventLogSourceName", 'System')
}

#Check to see if computer is a DC
if ((Get-WmiObject -Class Win32_OperatingSystem).productType -eq '2') {
    Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8965 -EntryType 'Information' -Message "Computer is a DC script will not run."
    Write-Output 'Computer is a DC'
    Exit
}

#Check to see if required commands are available
Try {
    Get-Command -Name 'Get-LocalUser' -ErrorAction Stop | Out-Null
}
catch {
    Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8569 -EntryType 'Error' -Message "WMF 5.1+ is required for Maintenance Script"
    Write-Output 'WMF 5.1 Not Installed'
    Exit
}

#Check if Local Admin Account Exists
Try {
    $LocalAdminAccount = Get-LocalUser $LocalAdminAccountName -ErrorAction Stop
    #$LocalAdminAccount = Get-WmiObject -class Win32_UserAccount | Where-Object { $_.Name -eq "$LocalAdminAccountName" -and $_.LocalAccount -eq $true }

    if ( $LocalAdminAccount.Enabled -eq $true) {
        Write-verbose 'Account Exists and it is enabled'
    }
    else {
        Enable-LocalUser -Name $LocalAdminAccount

        Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 3625 -EntryType 'Warning' -Message "$($LocalAdminAccountName) account has been enabled."
    }
}
Catch {
    #Create the account if it is not Found
    $RandomPassword = Get-RandomCharacters
    try {
        New-LocalUser -Name "$LocalAdminAccountName" -Password (ConvertTo-SecureString -string $RandomPassword -AsPlainText -force) | Add-LocalGroupMember -Group Administrators
    }
    Catch [Microsoft.PowerShell.Commands.InvalidPasswordException] {

    }
    catch {

    }

    Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8541 -EntryType 'Warning' -Message "$($LocalAdminAccountName) account has been created."

}

#Check if Local Admin Account is a member of Admins
#Get all local Administrators. Remove the computer/Domain name from all of the names.
$AdministratorsMembers = (Get-LocalGroupMember administrators -ErrorAction Ignore ).name -replace ".*\\", ""

if ($AdministratorsMembers -contains "$LocalAdminAccountName") {
    Write-Verbose 'Local Admin is member of the Administrators group'
}
Else {
    Add-LocalGroupMember -Group 'Administrators' -Member "$LocalAdminAccountName" -ErrorAction Ignore

    Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8526 -EntryType 'Warning' -Message "$($LocalAdminAccountName) account has been added to the Administrators Group."
}

#Change Local Admin password if LAPS is not installed
#Check to see if LAPS is installed, enabled, and configured to use the same local admin as this script
if ((Test-Path -Path "$env:ProgramFiles\LAPS\CSE\AdmPwd.dll") -and
    (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft Services\AdmPwd' -Name 'AdmPwdEnabled' -ValueData '1') -and
    (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft Services\AdmPwd' -Name 'AdminAccountName' -ValueData "$LocalAdminAccountName")
    ) {
    Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8526 -EntryType 'Information' -Message "LAPS was detected. $($LocalAdminAccountName) account password will not be changed"

    Write-Output "LAPS"
}
else {
    try{
        $RandomPassword = Get-RandomCharacters

        Set-LocalUser $LocalAdminAccountName -Password (ConvertTo-SecureString -string $RandomPassword -AsPlainText -force) -ErrorAction Stop

        Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8526 -EntryType 'Information' -Message "LAPS was not detected. $($LocalAdminAccountName) account password has been changed and sent to RMM"

        Write-Output $RandomPassword
    }
    catch{
        Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8526 -EntryType 'Error' -Message "Unable to set password for $($LocalAdminAccountName)"

        Write-Output "Could not set password for $LocalAdminAccountName"
    }

}

#Disable unwanted accounts
Foreach($Account in $LocalAccountsToDisable){
    #Check to see if the account exists
    $LocalAccount = Get-LocalUser -Name $Account -ErrorAction Ignore
    #Check to see if the account is enabled
    if ($LocalAccount.Enabled){
        Disable-LocalUser -Name $Account -ErrorAction Ignore
        Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8469 -EntryType 'Information' -Message "The account $($Account) was disabled"
    }
    #Check to see if the account is a member of the Administrators group
    if ((Get-LocalGroupMember -Group 'Administrators' -ErrorAction ignore).name -match $Account){
        Remove-LocalGroupMember -Group 'Administrators' -Member $Account -ErrorAction Ignore
        Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8469 -EntryType 'Information' -Message "The account $($Account) has been removed from the Administrators group."
    }
}

Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 2632 -EntryType 'Information' -Message "The $($LocalAdminAccountName) Maintance script has completed"