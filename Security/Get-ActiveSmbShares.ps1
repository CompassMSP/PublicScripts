<#
Returns a list of non-standard shares on a machine

Andy Morales
#>
$ShareDescriptionExclusions = @(
    'Remote Admin',
    'Default share',
    'Remote IPC',
    'Logon server share',
    'Printer Drivers'
)

$ShareNameExclusions = @(
    'SYSVOL',
    'NETLOGON'
)

$AllShares = Get-WmiObject Win32_Share

#Description exclusions
$FilteredShares = $AllShares | Where-Object { $ShareDescriptionExclusions -notcontains $_.Description }

#Name exclusions
$FilteredShares = $FilteredShares | Where-Object { $ShareNameExclusions -notcontains $_.Name }

#Remove Printers
$FilteredShares = $FilteredShares | Where-Object { $_.Path -notLike '*LocalsplOnly*' }

RETURN ($FilteredShares.Name | Sort-Object ) -join ','