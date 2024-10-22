<#
This script searches through the PowerShell event logs for events that might be suspicious

Andy Morales
#>

$csvDir = 'c:\'

$Events = Get-WinEvent -LogName 'Windows PowerShell'

#Find events that contain ies or invoke expression
$invokeEvents = @()
$downloadEvents = @()
$encodedEvents = @()

foreach ($Event in $Events) {
    if ($Event.Message -like '*iex*' -or $Event.Message -like '*invoke-expression*') {
        $invokeEvents += $Event
    }
    if ($Event.Message -like '*downloadstring*') {
        $downloadEvents += $Event
    }
    if ($Event.Message -like '*powershell*-e*') {
        $encodedEvents += $Event
    }
}

if ($invokeEvents.count -gt 0) {
    $invokeEvents | Export-Csv -Path "$($csvDir)iexEvents.csv" -NoTypeInformation
    Write-Output "IEX events exported to $($csvDir)iexEvents.csv"
}
else {
    Write-Output 'No iex events found'
}
if ($downloadEvents.count -gt 0) {
    $downloadEvents  | Export-Csv -Path "$($csvDir)downloadstringEvents.csv" -NoTypeInformation
    Write-Output "Download events exported to $($csvDir)downloadstringEvents.csv"
}
else {
    Write-Output 'No download events found'
}
if ($encodedEvents.count -gt 0) {
    $encodedEvents  | Export-Csv -Path "$($csvDir)\encodedEvents.csv" -NoTypeInformation
    Write-Output "Encoded events exported to $($csvDir)encodedEvents.csv"
}
else {
    Write-Output 'No encoded events found'
}