<#
This script searches for WMI subscriptions that might contain entries used for malware persistence.

.LINK
https://pentestlab.blog/2020/01/21/persistence-wmi-event-subscription/
#>

$ClassNames = @(
    '__EventFilter',
    '__FilterToConsumerBinding',
    '__EventConsumer',
    'CommandLineEventConsumer'
)

$ExeWhiteList = @(
    '*WSCEAA.exe*'
)

# Turn wildcards into regex
# First escape all characters that might cause trouble in regex (leaving out those we care about)
$escaped = $ExeWhiteList -replace '[ #$()+.[\\^{]', '\$&' # list taken from Regex.Escape
# replace wildcards with their regex equivalents
$regexStrings = $escaped -replace '\*', '.*' -replace '\?', '.'
# combine them into one regex
$ExeWhiteListRegex = ($regexStrings | ForEach-Object { '^' + $_ + '$' }) -join '|'

$WmiSubs = @()

Foreach ($class in $ClassNames) {
    $WmiSubs += Get-WmiObject -Namespace root\Subscription -Class $class
}

$SuspiciousSubFound = $false
$SuspiciousSubs = @()

foreach ($sub in $WmiSubs) {

    if ($sub.CommandLineTemplate -like '*powershell*' -or ($sub.CommandLineTemplate -like '*.exe*' -and $sub.CommandLineTemplate -notmatch $ExeWhiteListRegex)) {
        $SuspiciousSubFound = $true
        $SuspiciousSubs += $sub
    }
}

if ($SuspiciousSubFound) {
    Write-Output 'SuspiciousSubsFound'
    $SuspiciousSubs | Format-List Name,CommandLineTemplate
}
else{
    Write-Output 'NothingFound'
}