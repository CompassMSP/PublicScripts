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

    function New-SecureFolder {
        <#
    #This script creates a folder that only administrators and system have access to.
    #It is a best practice to wipe the folder once you are done running the script that relies on it.

    Andy Morales
    #>
        [CmdletBinding()]
        param (
            [parameter(Mandatory = $true,
                HelpMessage = 'C:\temp')]
            [String]$Path
        )

        #Delete the folder if it already exists
        if (Test-Path -Path $Path) {
            try {
                Remove-Item -Path $Path -Force -Recurse -ErrorAction Stop
            }
            catch {
                Write-Output "Could not clear the contents of $($Path). Script will exit."
                EXIT
            }
        }

        #Create the folder
        New-Item -Path $Path -ItemType Directory -Force

        #Remove all explicit permissions
        ICACLS ("$Path") /reset | Out-Null

        #Add SYSTEM permission
        ICACLS ("$Path") /grant ("SYSTEM" + ':(OI)(CI)F') | Out-Null

        #Give Administrators Full Control
        ICACLS ("$Path") /grant ("Administrators" + ':(OI)(CI)F') | Out-Null

        #Disable Inheritance on the Folder. This is done last to avoid permission errors.
        ICACLS ("$Path") /inheritance:r | Out-Null
    }

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
            Subject    = "Compromised passwords found on DC: $((Get-WmiObject Win32_Computersystem).Name)"
            Body       = $EmailBody
            BodyAsHtml = $true
            SmtpServer = $SMTPRelay
        }

        Send-MailMessage @SendMailMessageParams

        #Export CSV
        $CSVFolder = 'C:\Temp\ADPasswordAudit'
        New-SecureFolder -Path $CSVFolder
        $CompromisedUsers | Export-Csv -Path "$($CSVFolder)\CompromisedUsers.csv"

        Write-Output 'CSV exported to C:\Temp\ADPasswordAudit\CompromisedUsers.csv'
    }
}