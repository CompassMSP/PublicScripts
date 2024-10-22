<#
This will disable NetBIOS on all network adapters. It should be configured as a startup script so that any new adapters also have NetBIOS disbled.
#>
Set-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\NetBT\Parameters\Interfaces\tcpip* -Name NetbiosOptions -Value 2