#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync -RunAsAdministrator

<#Author       : Chris Williams
# Creation Date: 3-02-2022
# Usage        : Copies user template and creates new user with groups and licenses 

#********************************************************************************
# Date                     Version      Changes
#--------------------------------------------------------------------------------
# 03-02-2022                 1.0         Initial Version
# 03-04-2022                 1.1         Add Checks For Duplicate Attributes 
# 03-06-2022                 1.2         Add Check Loop for AD Sync
# 06-27-2022                 1.3         Change Group Lookup and Member Add
# 09-29-2022                 1.4         Add fax attributes copy
# 10-07-2022                 1.5         Add check for duplicate SamAccountName attributes
#********************************************************************************
#
# Run from the Primary Domain Controller with AD Connect installed
#
# The following modules must be installed
# Install-Module ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Online.Sharepoint.PowerShell
#
# Azure licenses Sku - Selected Sku must have free licenses available. This MUST be set in the portal before running the script
#
# Exchange Online = EXCHANGESTANDARD
# Microsoft 365 Business Basic = O365_BUSINESS_ESSENTIALS
# Microsoft 365 E3 = SPE_E3
# Microsoft 365 Business Premium = SPB
# Office 365 E3 = ENTERPRISEPACK
#
# .\Invoke-MgNewUserRequest.ps1 -UserToCopy "Copy User" -NewUser "Chris Williams" -Phone "555-555-5555"
#>

Param (
    [Parameter(Mandatory = $False)]
    [String]$UserToCopy,
    [Parameter(Mandatory = $False)]
    [String]$NewUser,
    [Parameter(Mandatory = $False)]
    [String]$Phone,
    [Parameter(Mandatory = $False)]
    [String]$SkipAz
)

IF ([string]::IsNullOrEmpty($UserToCopy)) {
    $UserToCopy = Read-Host "Please enter the DisplayName of the user template you want to copy. EX: 'Chris Williams'"
}
IF ([string]::IsNullOrEmpty($NewUser)) {
    $NewUser = Read-Host "Please enter the First and Last name of the new user. EX: 'New User'"
}
IF ([string]::IsNullOrEmpty($Phone)) {
    $Phone = Read-Host "Please enter mobile phone number of the new user."
}

if ($SkipAz -ne 'y') {
    Write-Output 'Logging into 365 services.'
    Connect-ExchangeOnline
    Connect-MgGraph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All", "Directory.AccessAsUser.All", "Group.ReadWrite.All", "GroupMember.Read.All", "Organization.Read.All"
    Connect-SPOService -Url "https://compassmsp-admin.sharepoint.com"
}

if ($Sku) { 
    try { 
        $GetLic = Get-MgSubscribedSku | Where-Object { ($_.SkuPartNumber -eq $Sku) } -ErrorAction stop
    }
    catch { 
        Write-Output "License Sku could not be found. Or no Sku was selected."
        $Sku = $NULL
    }
}

if (!$Sku) { 
    $License = Get-MgSubscribedSku | Select-Object SkuPartNumber, ConsumedUnits, SkuId | Where-Object { $_.SkuPartNumber -eq 'EXCHANGESTANDARD' -or $_.SkuPartNumber -eq 'O365_BUSINESS_ESSENTIALS' -or $_.SkuPartNumber -eq 'SPE_E3' -or $_.SkuPartNumber -eq 'SPB' -or $_.SkuPartNumber -eq 'ENTERPRISEPACK' }

    $GridArguments = @{
        OutputMode = 'Single'
        Title      = 'Please select a license and click OK'
    }
        
    $GetLic = $License | Out-GridView @GridArguments | ForEach-Object {
        $_
    } 
}

try {
    $UserToCopyUPN = Get-ADUser -Filter "DisplayName -eq '$($UserToCopy)'" -Properties Title, Fax, wWWHomePage, physicalDeliveryOfficeName, Office, Manager, Description, Department, Company
    if ($UserToCopyUPN.Count -gt 1) {  
        Write-Host "UserToCopy has multiple values. Please check AD for accounts with duplicate DisplayName attributes."
        exit
    } 
}
catch {
    Write-Output "Could not find user $($UserToCopy) in AD to copy from."
    exit
}

$Domain = $($UserToCopyUPN.UserPrincipalName -replace '.+?(?=@)')
$NewUserFirstName = $($NewUser.split(' ')[-2])
$NewUserLastName = $($NewUser -replace '.+\s')
$NewUserSamAccountName = $(($NewUser -replace '(?<=.{1}).+') + ($NewUser -replace '.+\s')).ToLower()
$NewUserEmail = $($NewUserSamAccountName + $Domain).ToLower()

$CheckNewUserUPN = $(try { Get-ADUser -Identity $NewUserSamAccountName } catch { $null })
if ($null -ne $CheckNewUserUPN) {
    Write-Host "SamAccountName exist for user $NewUser. Please check AD for accounts with duplicate SamAccountName attributes."
    exit
} 

function Get-NewPassword { -join ('abcdefghkmnrstuvwxyzABCDEFGHKLMNPRSTUVWXYZ23456789$%&*#'.ToCharArray() | Get-Random -Count 16) }

$Password = Get-NewPassword

$Confirmation = Read-Host -Prompt "The user below will be created:`n
Display Name = $($NewUser)
Email Address = $($NewUserEmail)
Password = $($Password)
First Name = $($NewUserFirstName)
Last Name = $($NewUserLastName)
SamAccountName = $($NewUserSamAccountName)
Destination OU = $($UserToCopyUPN.DistinguishedName.split(",",2)[1])`n
Template User to Copy = $($UserToCopy)`n
Continue? (Y/N)`n"

if ($Confirmation -ne 'y') {
    Write-Output 'User did not enter "Y"'
    exit
}

try {
    New-ADUser -Name $NewUser `
        -SamAccountName $NewUserSamAccountName `
        -UserPrincipalName $NewUserEmail `
        -DisplayName $NewUser `
        -GivenName $NewUserFirstName `
        -Surname $NewUserLastName  `
        -MobilePhone $Phone `
        -EmailAddress $NewUserEmail `
        -OtherAttributes @{ 'proxyAddresses' = "SMTP:$($NewUserEmail)" } `
        -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
        -Path $($UserToCopyUPN.DistinguishedName.split(",", 2)[1]) `
        -Instance $UserToCopyUPN `
        -Enabled $True
}
catch { 
    Write-Host "New User creation was not successful."
    exit
}

Write-Output 'AD User has been created.'

Write-Output 'Adding AD Groups to new user.'

$CopyFromUser = Get-ADUser -Filter "DisplayName -eq '$($UserToCopy)'" -prop MemberOf
$CopyToUser = Get-ADUser -Filter "DisplayName -eq '$($NewUser)'" -prop MemberOf
$CopyFromUser.MemberOf | Where-Object { $CopyToUser.MemberOf -notcontains $_ } | Add-ADGroupMember -Members $CopyToUser

Write-Output 'Starting AD Sync'

powershell.exe -command Start-ADSyncSyncCycle -PolicyType Delta

Write-Output 'Waiting 90 seconds for AD Connect sync process.'

Start-Sleep -Seconds 90

$Stoploop = $false
[int]$Retrycount = "0"
 
do {
    try {
        $NewMgUser = Get-MgUser -UserId $NewUserEmail -ErrorAction Stop
        Write-Output "User $NewUser has synced to Azure. Script will now continue."
        $Stoploop = $true
    }
    catch {
        if ($Retrycount -gt 3) {
            Write-Host "Could not sync AD User to 365 after 3 retries."
            $Stoploop = $true
        }
        else {
            Write-Host "Could not sync AD User to 365 retrying in 60 seconds..."
            Start-Sleep -Seconds 60
            $Retrycount = $Retrycount + 1
        }
    }
} while ($Stoploop -eq $false)

if (!$NewMgUser) { 
    $ADSyncCompleteYesorExit = Read-Host -Prompt 'AD Sync has not completed within allotted time frame. Please wait for AD sync. To resume type yes or exit'
} while ("yes", "exit" -notcontains $ADSyncCompleteYesorExit ) { 
    $ADSyncCompleteYesorExit = Read-Host "Please enter your response (yes/exit)"
}

if ($ADSyncCompleteYesorExit -eq 'exit') {
    Write-Output 'You will need to set the license and add Office 365 groups via the portal. Script will now exit'
    exit
}

if ($ADSyncCompleteYesorExit -eq 'yes') {

    $NewMgUser = Get-MgUser -UserId $NewUserEmail -ErrorAction Stop
    if (!$NewMgUser) { 
        Write-Output 'Script cannot find new user. You will need to set the license and add Office 365 groups via the portal.'
        exit
    }

    Write-Output 'Script now will resume'

    Write-Output 'Setting Usage Location for new user'

    Update-MgUser -UserId $NewMgUser.Id -UsageLocation US

    Write-Output 'Adding Office 365 Groups to new user.'

    if ($GetLic) { 
        try {
            Set-MgUserLicense -UserId $NewMgUser.Id -AddLicenses @{SkuId = $GetLic.SkuId } -RemoveLicenses @() -ErrorAction stop
            Write-Output 'License added.'
        }
        catch {
            Write-Output 'License could not be added. You will need to set the license and add Office 365 groups via the portal.'
            exit
        }
    }

    $All365Groups = Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserToCopyUPN.UserPrincipalName).Id  | Where-Object {
        $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' } | ForEach-Object { 
        @{ GroupId = $_.Id } } | Get-MgGroup | Where-Object { $_.OnPremisesSyncEnabled -eq $NULL -and $_.DisplayName -ne 'All Users' } | Select-Object DisplayName, SecurityEnabled, Mail, Id

    Foreach ($365Group in $All365Groups) {
        try {
            New-MgGroupMember -GroupId $365Group.Id -DirectoryObjectId $(Get-MgUser -UserId $NewUserEmail).Id -ErrorAction Stop  
        }
        catch {
            Add-DistributionGroupMember -Identity $365Group.DisplayName -Member $NewUserEmail -BypassSecurityGroupManagerCheck -Confirm:$false
        }
    }

    $CopyUserGroupCount = (Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserToCopyUPN.UserPrincipalName).Id).Count
    $NewUserGroupCount = (Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $NewUserEmail).Id).Count

    ## Assigns US as UsageLocation
    Update-MgUser -UserId $NewUserEmail -UsageLocation US

    ## Creates OneDrive
    Request-SPOPersonalSite -UserEmails $NewUserEmail -NoWait
    
    #Adds user to All Company group.
    New-MgGroupMember -GroupId (Get-MgGroup -Filter "DisplayName eq 'All Company'").Id -DirectoryObjectId $(Get-MgUser -UserId $NewUserEmail).Id

    Write-Output "User $($NewUser) should now be created unless any errors occurred during the process."
    Write-Output "Copy User group count: $($CopyUserGroupCount)"
    Write-Output "New User group count: $($NewUserGroupCount)"

    $AddLic = Read-Host "Would you like to add additional licenses? (Y/N)"

    if ($AddLic -ne 'y') { Write-Output 'Goodbye!' }

    if ($AddLic -eq 'y') { 
        $License2 = Get-MgSubscribedSku | Select-Object SkuPartNumber, ConsumedUnits, SkuId

        $GridArguments = @{
            OutputMode = 'Multiple'
            Title      = 'Please select licenses and click OK'
        }
    
        $GetLic2 = $License2 | Out-GridView @GridArguments | ForEach-Object {
            $_
        } 
    }

    if ($GetLic2) { 
        $GetLic2 | ForEach-Object {
            try {
                Set-MgUserLicense -UserId $NewMgUser.Id -AddLicenses @{SkuId = $_.SkuId } -RemoveLicenses @() -ErrorAction stop
                Write-Output "$($_.SkuPartNumber) License added."
            }
            catch {
                Write-Output "$($_.SkuPartNumber) License could not be added."
            }   
        }
    }
}
