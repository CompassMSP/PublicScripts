#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync -RunAsAdministrator

<#Author       : Chris Williams
# Creation Date: 03-02-2022
# Usage        : Copies user template and creates new user with groups and licenses 

#********************************************************************************
# Date                     Version      Changes
#--------------------------------------------------------------------------------
# 03-02-2022                    1.0         Initial Version
# 03-04-2022                    1.1         Add Checks For Duplicate Attributes 
# 03-06-2022                    1.2         Add Check Loop for AD Sync
# 06-27-2022                    1.3         Change Group Lookup and Member Add
# 09-29-2022                    1.4         Add fax attributes copy
# 10-07-2022                    1.5         Add check for duplicate SamAccountName attributes
# 02-12-2024                    1.6         Add AppRoleAssignment for KnowBe4 SCIM App
# 02-14-2024                    1.7         Fix issues with copy groups function and code cleanup
# 02-19-2024                    1.8         Changes to Get-MgUserMemberOf function 
# 03-08-2024                    1.9         Cleaned up licenses select display output
# 05-08-2024                    2.0         Add input box for Variables
# 05-21-2024                    2.1         Added stop for if UserToCopy cannot be found
#********************************************************************************
#
# Run from the Primary Domain Controller with AD Connect installed
#
# The following modules must be installed
# Install-Module ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, PnP.PowerShell
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

#Import-Module adsync -UseWindowsPowerShell

Param (
    [Parameter(Mandatory = $False)]
    [String]$SkipAz
)


#Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show("For the 'NewUser' and 'UserToCopy' please enter in a DiplayName format: 'FirstName LastName'", 'Compass New User Request')

function CompassNewUserRequest {
    param (
        [Parameter(Mandatory)]
        [string]$NewUser,
        [string]$NewUserMobile,
        [Parameter(Mandatory)]
        [string]$UserToCopy,
        [validateset('Exchange Online (Plan 1)', 'Microsoft 365 Business Basic', 'Microsoft 365 E3', 'Microsoft 365 Business Premium', 'Office 365 E3')]
        [string]$SelectLicenseSku
    )
    [pscustomobject]@{
        InputNewUser    = $NewUser
        InputNewMobile  = $NewUserMobile
        InputUserToCopy = $UserToCopy
        InputSku        = $SelectLicenseSku
    }
}

$result = Invoke-Expression (Show-Command CompassNewUserRequest -PassThru)

$NewUser = $result.InputNewUser
$Phone = $result.InputNewMobile
$UserToCopy = $result.InputUserToCopy

$UserToCopyUPN = Get-ADUser -Filter "DisplayName -eq '$($UserToCopy)'" -Properties Title, Fax, wWWHomePage, physicalDeliveryOfficeName, Office, Manager, Description, Department, Company 
    
if ($UserToCopyUPN.Count -gt 1) {  
    Write-Host "UserToCopy has multiple values. Please check AD for accounts with duplicate DisplayName attributes."
    exit
} elseif ($NULL -eq $UserToCopyUPN) {
    Write-Output "Could not find user $($UserToCopy) in AD to copy from."
    exit
}

if (!$result.InputSku) { 
    Write-Host 'License Sku not selected.'
} else {
    if ($result.InputSku -eq 'Exchange Online (Plan 1)') { $Sku = "EXCHANGESTANDARD" }
    if ($result.InputSku -eq 'Microsoft 365 Business Basic') { $Sku = "O365_BUSINESS_ESSENTIALS" }
    if ($result.InputSku -eq 'Microsoft 365 E3') { $Sku = "SPE_E3" }
    if ($result.InputSku -eq 'Microsoft 365 Business Premium') { $Sku = "SPB" }
    if ($result.InputSku -eq 'Office 365 E3') { $Sku = "ENTERPRISEPACK" }
}

if ($SkipAz -ne 'y') {
    Write-Output 'Logging into 365 services.'
    $Scopes = @(
        "Directory.ReadWrite.All",
        "User.ReadWrite.All",
        "Directory.AccessAsUser.All",
        "Group.ReadWrite.All",
        "GroupMember.Read.All", 
        "Organization.Read.All",
        "AppRoleAssignment.ReadWrite.All")
    Connect-MgGraph -Scopes $Scopes -NoWelcome
    Connect-ExchangeOnline -ShowBanner:$false 
}

if ($Sku) { 
    try {
        $SelectObjectPropertyList = @(
            "SkuPartNumber"
            "SkuId"
            @{
                n = "Available"
                e = { (($_.PrepaidUnits).Enabled - $_.ConsumedUnits) }
            }
        )

        $getLicCount = Get-MgSubscribedSku | Where-Object { ($_.SkuPartNumber -eq $Sku) } | Select-Object $SelectObjectPropertyList

        if ($getLicCount.Available -gt 0) {
            $getLic = $getLicCount
        } else {
            Write-Output "No available license for '$($result.InputSku)'. Please add additional licenses via the Microsoft Portal."
            $Sku = $NULL
        }
    } catch { 
        Write-Output "License Sku could not be found. Or no Sku was selected."
        $Sku = $NULL
    }
}

if (!$Sku) {

    $SelectObjectPropertyList = @(
        "SkuPartNumber"
        "SkuId"
        @{
            n = "ActiveUnits"
            e = { ($_.PrepaidUnits).Enabled }
        }
        "ConsumedUnits"
    )

    $WhereObjectFilter = {
        ($_.SkuPartNumber -eq 'EXCHANGESTANDARD') -or 
        ($_.SkuPartNumber -eq 'O365_BUSINESS_ESSENTIALS') -or 
        ($_.SkuPartNumber -eq 'SPE_E3') -or 
        ($_.SkuPartNumber -eq 'SPB') -or 
        ($_.SkuPartNumber -eq 'ENTERPRISEPACK')
    }

    $selectLicense = Get-MgSubscribedSku | Select-Object $SelectObjectPropertyList | Where-Object -FilterScript $WhereObjectFilter | `
        ForEach-Object {
        [PSCustomObject]@{
            DisplayName   = switch -Regex ($_.SkuPartNumber) {
                "EXCHANGESTANDARD" { "Exchange Online (Plan 1)" }
                "O365_BUSINESS_ESSENTIALS" { "Microsoft 365 Business Basic" }
                "SPE_E3" { "Microsoft 365 E3" }
                "SPB" { "Microsoft 365 Business Premium" }
                "ENTERPRISEPACK" { "Office 365 E3" }
            }
            SkuPartNumber = $_.SkuPartNumber
            SkuId         = $_.SkuId
            #NumberTotal   = $_.ActiveUnits
            #NumberUsed    = $_.ConsumedUnits
            Available     = ($_.ActiveUnits - $_.ConsumedUnits)
        }
    } | Sort-Object DisplayName
    
    $GridArguments = @{
        OutputMode = 'Single'
        Title      = 'Please select a license and click OK'
    }
    
    $selectLicenseTEMP = $selectLicense | ForEach-Object { $_ | Select-Object -Property 'DisplayName', 'Available' } | Out-GridView @GridArguments 
    $getLic = $selectLicense | Where-Object { $_.DisplayName -in $selectLicenseTEMP.DisplayName }

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
} catch { 
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
        $ADSyncCompleteYesorExit = 'yes'
    } catch {
        if ($Retrycount -gt 3) {
            Write-Host "Could not sync AD User to 365 after 3 retries."
            $Stoploop = $true
        } else {
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

    Write-Output 'Adding Office 365 Groups to new user.'

    if ($getLic) { 
        try {
            Set-MgUserLicense -UserId $NewMgUser.Id -AddLicenses @{SkuId = $getLic.SkuId } -RemoveLicenses @() -ErrorAction stop
            Write-Output 'License added.'
        } catch {
            Write-Output 'License could not be added. You will need to set the license and add Office 365 groups via the portal.'
            exit
        }
    }

    $All365Groups = Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserToCopyUPN.UserPrincipalName).Id | `
        Where-Object { $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' -and $_.AdditionalProperties.membershipRule -eq $NULL -and $_.onPremisesSyncEnabled -ne 'False' } | `
        Select-Object Id, @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }, @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }

    Foreach ($365Group in $All365Groups) {
        try {
            New-MgGroupMember -GroupId $365Group.Id -DirectoryObjectId $(Get-MgUser -UserId $NewUserEmail).Id -ErrorAction Stop  
        } catch {
            Add-DistributionGroupMember -Identity $365Group.DisplayName -Member $NewUserEmail -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction 'SilentlyContinue'
        }
    }

    $CopyUserGroupCount = (Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserToCopyUPN.UserPrincipalName).Id).Count
    $NewUserGroupCount = (Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $NewUserEmail).Id).Count

    Write-Output 'Setting Usage Location for new user'
    
    ## Assigns US as UsageLocation
    Update-MgUser -UserId $NewUserEmail -UsageLocation US

    #Adds user to All Company group.
    New-MgGroupMember -GroupId (Get-MgGroup -Filter "DisplayName eq 'All Company'").Id -DirectoryObjectId $(Get-MgUser -UserId $NewUserEmail).Id
    New-MgGroupMember -GroupId (Get-MgGroup -Filter "DisplayName eq 'Exclaimer Default'").Id -DirectoryObjectId $(Get-MgUser -UserId $NewUserEmail).Id
    New-MgGroupMember -GroupId (Get-MgGroup -Filter "DisplayName eq 'Exclaimer Add-in'").Id -DirectoryObjectId $(Get-MgUser -UserId $NewUserEmail).Id

    Write-Output "User $($NewUser) should now be created unless any errors occurred during the process."
    Write-Output "Copy User group count: $($CopyUserGroupCount)"
    Write-Output "New User group count: $($NewUserGroupCount)"

    $AddLic = Read-Host "Would you like to add additional licenses? (Y/N)"

    if ($AddLic -ne 'y') { Write-Output 'Goodbye!' }

    if ($AddLic -eq 'y') { 

        $SelectObjectPropertyList = @(
            "SkuPartNumber"
            "SkuId"
            @{
                n = "ActiveUnits"
                e = { ($_.PrepaidUnits).Enabled }
            }
            "ConsumedUnits"
        )
    
        $WhereObjectFilter = {
            ($_.SkuPartNumber -notlike 'EXCHANGESTANDARD') -and 
            ($_.SkuPartNumber -notlike 'O365_BUSINESS_ESSENTIALS') -and 
            ($_.SkuPartNumber -notlike 'SPE_E3') -and 
            ($_.SkuPartNumber -notlike 'SPB') -and
            ($_.SkuPartNumber -notlike 'ENTERPRISEPACK') -and
            ($_.SkuPartNumber -notlike 'PROJECT_MADEIRA_PREVIEW_IW_SKU') -and
            ($_.SkuPartNumber -notlike 'POWERAUTOMATE_ATTENDED_RPA') -and
            ($_.SkuPartNumber -notlike 'RMSBASIC') -and
            ($_.SkuPartNumber -notlike 'MCOPSTNC') -and
            ($_.SkuPartNumber -notlike 'CCIBOTS_PRIVPREV_VIRAL') -and
            ($_.SkuPartNumber -notlike 'MCOPSTN1') -and
            ($_.SkuPartNumber -notlike 'WINDOWS_STORE') -and
            ($_.SkuPartNumber -notlike 'STREAM') -and
            ($_.SkuPartNumber -notlike 'POWERAPPS_DEV') -and
            ($_.SkuPartNumber -notlike 'RIGHTSMANAGEMENT_ADHOC') -and
            ($_.SkuPartNumber -notlike 'MCOMEETADV') -and
            ($_.SkuPartNumber -notlike 'MEETING_ROOM') -and
            ($_.SkuPartNumber -notlike 'VISIO_PLAN1_DEPT') -and
            ($_.SkuPartNumber -notlike 'FLOW_FREE') -and
            ($_.SkuPartNumber -notlike 'MICROSOFT_BUSINESS_CENTER') -and
            ($_.SkuPartNumber -notlike 'PHONESYSTEM_VIRTUALUSER') -and
            ($_.SkuPartNumber -notlike 'Microsoft_Copilot_for_Finance_trial') -and
            ($_.SkuPartNumber -notlike 'POWERAPPS_VIRAL') -and
            ($_.SkuPartNumber -notlike 'Microsoft_Teams_Exploratory_Dept') -and
            ($_.SkuPartNumber -notlike 'POWERAPPS_PER_USER') -and
            ($_.SkuPartNumber -notlike 'Power BI Standard')
        }
    
        $selectLicense2 = Get-MgSubscribedSku | Select-Object $SelectObjectPropertyList | Where-Object -FilterScript $WhereObjectFilter | `
            ForEach-Object {
            [PSCustomObject]@{
                DisplayName   = switch -Regex ($_.SkuPartNumber) {
                    "PROJECT_P1" { "Project Plan 1" }
                    "PROJECTPROFESSIONAL" { "Project Plan 3" }
                    "VISIOCLIENT" { "Visio Plan 2" }
                    "Microsoft_Teams_Audio_Conferencing_select_dial_out" { "Microsoft Teams Audio Conferencing with dial-out to USA/CAN" }
                    "POWER_BI_PRO" { "Power BI Pro" }
                    "Microsoft_365_Copilot" { "Microsoft 365 Copilot" }
                    "Microsoft_Teams_Premium" { "Microsoft Teams Premium" }
                    "MCOEV" { "Microsoft Teams Phone Standard" }
                    "AAD_PREMIUM_P2" { "Microsoft Entra ID P2" }
                    "POWER_BI_STANDARD" { "Power BI Standard" }
                    "Microsoft365_Lighthouse" { "Microsoft 365 Lighthouse" }
                }
                SkuPartNumber = $_.SkuPartNumber
                SkuId         = $_.SkuId
                #NumberTotal   = $_.ActiveUnits
                #NumberUsed    = $_.ConsumedUnits
                Available     = ($_.ActiveUnits - $_.ConsumedUnits)
            }
        } | Sort-Object DisplayName
    
        $GridArguments = @{
            OutputMode = 'Multiple'
            Title      = 'Please select licenses and click OK (Hold CTRL to select multiple licenses)'
        }
    
        $selectLicenseTEMP2 = $selectLicense2 | ForEach-Object { $_ | Select-Object -Property 'DisplayName', 'Available' } | Out-GridView @GridArguments 
        $getLic2 = $selectLicense2 | Where-Object { $_.DisplayName -in $selectLicenseTEMP2.DisplayName }
    }

    if ($GetLic2) { 
        $GetLic2 | ForEach-Object {
            try {
                Set-MgUserLicense -UserId $NewMgUser.Id -AddLicenses @{SkuId = $_.SkuId } -RemoveLicenses @() -ErrorAction stop
                Write-Output "$($_.SkuPartNumber) License added."
            } catch {
                Write-Output "$($_.SkuPartNumber) License could not be added."
            }   
        }
    }

    #Disconnect from Exchange and Graph
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-Graph

    ## Creates OneDrive
    Connect-PnPOnline -Url compassmsp-admin.sharepoint.com -ClientId '24e3c6ad-9658-4a0d-b85f-82d67d148449' -Tenant compassmsp.onmicrosoft.com -Thumbprint '3b51fcc465d26593303453c8a636b13587e0dc81'
    Request-PnPPersonalSite -UserEmails $NewUserEmail -NoWait
    Disconnect-PnPOnline
    
}