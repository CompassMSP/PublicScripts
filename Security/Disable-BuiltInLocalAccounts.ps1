<#
This script disables and renames built-in local accounts.

Accounts will be matched using their SID to ensure that they are always found.

The script will only run if the custom local admin account $LocalAdminAccountName has been created.

Andy Morales
#>
$LocalAdminAccountName = 'Compass'

$LocalAccountsToDisable = @()

$CurrentLocalAccounts = Get-WmiObject Win32_UserAccount -Filter 'LocalAccount=True'

Foreach ($LocalAccount in $CurrentLocalAccounts) {
    #Built-in Administrator account
    if ($LocalAccount.SID -like '*-500') {
        $LocalAccountsToDisable += $LocalAccount.Name
        $AdministratorAccountName = $LocalAccount.Name
    }
    #Built-in Guest Account
    elseif ($LocalAccount.SID -like '*-501') {
        $LocalAccountsToDisable += $LocalAccount.Name
        $GuestAccountName = $LocalAccount.Name
    }
}

$EventLogSourceName = 'DisableBuiltInAccountsScript'

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

    if ( $LocalAdminAccount.Enabled -eq $true) {
        Write-Verbose 'Account Exists and it is enabled'

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

        $DisableBuiltInAccount = $true

    }
    else {
        Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 3625 -EntryType 'Error' -Message "$($LocalAdminAccountName) account exists but it is not enabled."
        $DisableBuiltInAccount = $false
    }
}
Catch {
    Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8541 -EntryType 'Error' -Message "$($LocalAdminAccountName) account does not exist. The script will exit"
    $DisableBuiltInAccount = $false
}

if ($DisableBuiltInAccount ) {
    $CurrentLocalAdministrators = Get-LocalGroupMember -Group 'Administrators' -ErrorAction Ignore

    #Disable unwanted accounts
    Foreach ($Account in $LocalAccountsToDisable) {
        #Check to see if the account exists
        $LocalAccount = Get-LocalUser -Name $Account -ErrorAction Ignore
        #Check to see if the account is enabled
        if ($LocalAccount.Enabled) {
            Disable-LocalUser -Name $Account -ErrorAction Ignore
            Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8469 -EntryType 'Information' -Message "The account $($Account) was disabled"
        }
        #Check to see if the account is a member of the Administrators group
        #The built-in "Administrator" account cannot be removed from the Administrators group
        if ($CurrentLocalAdministrators.name -match $Account) {
            Remove-LocalGroupMember -Group 'Administrators' -Member $Account -ErrorAction Ignore
            Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 8469 -EntryType 'Information' -Message "The account $($Account) has been removed from the Administrators group."
        }
    }

    #Rename Built-in accounts
    if ($AdministratorAccountName -ne 'Robert Bell') {
        Rename-LocalUser -Name $AdministratorAccountName -NewName 'Robert Bell'
        Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 2688 -EntryType 'Information' -Message "Renamed $($AdministratorAccountName) to Robert Bell"
    }
    if ($GuestAccountName -ne 'guestRenamed') {
        Rename-LocalUser -Name $GuestAccountName -NewName 'guestRenamed'
        Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 2688 -EntryType 'Information' -Message "Renamed $($GuestAccountName) to guestRenamed"
    }
}

Write-EventLog -LogName System -Source "$EventLogSourceName" -EventId 2632 -EntryType 'Information' -Message "The disable built-in accounts script has completed"