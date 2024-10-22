<#
Enables recommended security options


.LINK
https://support.microsoft.com/en-us/topic/client-service-and-program-issues-can-occur-if-you-change-security-settings-and-user-rights-assignments-0cb6901b-dcbf-d1a9-e9ea-f1b49a56d53a

Andy Morales
#>
reg add "HKLM\System\CurrentControlSet\Control\Lsa" /v "LimitBlankPasswordUse" /t REG_DWORD /d 1 /f

#anonymous access disable
reg add "HKLM\System\CurrentControlSet\Control\Lsa" /v "RestrictAnonymousSAM" /t REG_DWORD /d 1 /f
reg add "HKLM\System\CurrentControlSet\Control\Lsa" /v "RestrictAnonymous" /t REG_DWORD /d 1 /f
reg add "HKLM\System\CurrentControlSet\Services\LanmanServer\Parameters" /v "RestrictNullSessAccess" /t REG_DWORD /d 1 /f

reg add "HKLM\System\CurrentControlSet\Control\Lsa\MSV1_0" /v "allowNullSessionFallback" /t REG_DWORD /d 0 /f