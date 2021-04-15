<#
This script enables additional detailed auditing on a computer.

IMPORTANT: The secpol gui will not reflect these changes. Use the command below to confirm:
    auditpol /get /category:*

Andy Morales
#>

$SuccessAndFail = @(
    'Credential Validation',
    'Other Account Management Events',
    'Security Group Management',
    'User Account Management',
    'Process Creation',
    'Account Lockout',
    'Other Logon/Logoff Events',
    'Special Logon',
    'Other Object Access Events',
    'Audit Policy Change',
    'Authentication Policy Change',
    'Sensitive Privilege Use',
    'IPsec Driver',
    'Other System Events',
    'Security State Change',
    'System Integrity',
    'Kerberos Authentication Service',
    'Computer Account Management',
    'Directory Service Access',
    'Directory Service Changes'
)

$SuccessOnly = @(
    'Plug and Play Events',
    'Group Membership',
    'Authorization Policy Change',
    'Security System Extension',
    'Application Group Management',
    'Distribution Group Management'
)

$FailOnly = @(
    'Detailed File Share',
    'Other Policy Change Events',
    'Kerberos Service Ticket Operations',
    'Other Account Logon Events'
)

#add specific audit rules for DCs/non-DCs
if ((Get-WmiObject Win32_ComputerSystem).domainRole -ge 4) {
    #DomainController
}
else {
    #non-domain controller
    $SuccessAndFail += @(
        'Logon',
        'Logoff'
    )
}

Foreach ($sf in $SuccessAndFail){
    C:\Windows\System32\auditpol.exe /set /subcategory:"$($sf)" /success:enable /failure:enable
}
Foreach ($so in $SuccessOnly) {
    C:\Windows\System32\auditpol.exe /set /subcategory:"$($so)" /success:enable /failure:disable
}
Foreach ($fo in $FailOnly) {
    C:\Windows\System32\auditpol.exe /set /subcategory:"$($so)" /success:disable /failure:enable
}

#Force the use of advanced audit policies
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Lsa' -Name SCENoApplyLegacyAuditPolicy -Value 1

#region increaseLogSize
$LogsToIncrease = @(
    'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-TerminalServices-LocalSessionManager/Admin'
    'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
    'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-TerminalServices-RemoteConnectionManager/Admin'
    'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'
)

Foreach ($log in $LogsToIncrease){
    reg add $log /v MaxSize /t REG_DWORD /d 512032768 /f
}

#max size value is different for Policies
$PolicyLogsToIncrease = @(
    'HKLM\SOFTWARE\Policies\Microsoft\Windows\EventLog\Application'
    'HKLM\SOFTWARE\Policies\Microsoft\Windows\EventLog\Security'
    'HKLM\SOFTWARE\Policies\Microsoft\Windows\EventLog\System'
)
Foreach ($pLog in $PolicyLogsToIncrease) {
    reg add $pLog /v MaxSize /t REG_DWORD /d 500032 /f
}
#endregion increaseLogSize