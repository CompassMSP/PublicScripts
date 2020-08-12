<#
This script runs an audit of AD and checks for any accounts that have known bad passwords.

Lithnet AD Password Protection must be enabled and configured on the domain. The DC must have been restarted after the install for this script to work.

Andy Morales
#>
Function Invoke-ADPasswordAudit {
    #Requires -Modules LithnetPasswordProtection, ActiveDirectory

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$NotificationEmail,

        [Parameter(Mandatory = $true)]
        [string]$FromEmail,

        [Parameter(Mandatory = $true)]
        [string]$SMTPRelay
    )

    Import-Module LithnetPasswordProtection
    Import-Module ActiveDirectory

    $CompromisedUsers = @()

    $AllActiveADUsers = Get-ADUser -Filter 'enabled -eq "true"' -Properties PwdLastSet, lastLogonTimeStamp, mail

    #Identify any compromised users
    Foreach ($user in $AllActiveADUsers) {
        if (Test-IsADUserPasswordCompromised -AccountName $user.SamAccountName -Server $env:COMPUTERNAME) {
            $CompromisedUsers += [PSCustomObject]@{
                Name              = $user.Name
                SamAccountName    = $user.SamAccountName
                Mail              = $user.mail
                UserPrincipalName = $user.UserPrincipalName
                PwdLastSet        = [DateTime]::FromFileTimeUtc($user.PwdLastSet).ToShortDateString()
            }
        }
    }

    #Send message if any compromised users are found
    If ($CompromisedUsers.count -gt 0) {
        $EmailBody = $CompromisedUsers | ConvertTo-Html -Fragment -As Table | Out-String

        $SendMailMessageParams = @{
            To         = $NotificationEmail
            From       = $FromEmail
            Subject    = "Compromised passwords found on domain $((Get-WmiObject Win32_ComputerSystem).Domain)"
            Body       = $EmailBody
            BodyAsHtml = $true
            SmtpServer = $SMTPRelay
        }

        Send-MailMessage @SendMailMessageParams
    }
}