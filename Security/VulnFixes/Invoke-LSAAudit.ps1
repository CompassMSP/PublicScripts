<#
This script will check the event log to make sure that LSA protection can be enabled.

Andy Morales
#>
Function Invoke-LSAAudit {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$NotificationEmail,

        [Parameter(Mandatory = $true)]
        [string]$FromEmail,

        [Parameter(Mandatory = $true)]
        [string]$SMTPRelay
    )

    $WinEventHashTable = @{
        ProviderName = 'Microsoft-Windows-CodeIntegrity';
        Id           = '3065', '3066'
    }

    $Events = Get-WinEvent -FilterHashTable  $WinEventHashTable

    if ($Events.count -gt 0) {

        $EventsToExclude = @(
            'CyMemDef64.dll',
            'Signature information for another event. Match using the Correlation Id.'
        )

        #Filter out known events
        $FilteredEvents = @()

        Foreach ($Event in $Events) {
            if ($Event.Message -notMatch ($EventsToExclude -join '|')) {
                $FilteredEvents += $Event
            }
        }

        if ($FilteredEvents.Count -gt 0) {
            $EmailBody = $FilteredEvents[0..10] | ConvertTo-Html -Fragment -As Table | Out-String

            $SendMailMessageParams = @{
                To         = $NotificationEmail
                From       = $FromEmail
                Subject    = "Found LSA issue on $($ENV:COMPUTERNAME)"
                Body       = $EmailBody
                BodyAsHtml = $true
                SmtpServer = $SMTPRelay
            }

            Send-MailMessage @SendMailMessageParams
        }
        else {
            Write-Output 'No Events Found'
        }
    }
    else {
        Write-Output 'No events found'
    }
}