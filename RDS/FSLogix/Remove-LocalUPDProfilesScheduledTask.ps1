<#

.SYNOPSIS
This script deletes local profile data on an RDS server that has UPDs. When UPDs are enabled user files will sometimes be left behind. This can cause some unexpected resuults with applications.

The script has a failsafe that does not create the task if UPDs become disabled at any point.

The script needs Remove-LocalUPDProfiles.ps1 to work correctly.

.PARAMETER ScheduledTaskName
This is the name that will be given to the ScheduledTask

Andy Morales
#>


$ScheduledTaskName = 'Delete UPD Local Files'

function Test-RegistryValue {
    <#
    Checks if a reg key/value exists

    #Modified version of the function below
    #https://www.jonathanmedd.net/2014/02/testing-for-the-presence-of-a-registry-key-and-value.html

    Andy Morales
    #>

    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
            Position = 1,
            HelpMessage = 'HKEY_LOCAL_MACHINE\SYSTEM')]
        [ValidatePattern('Registry::.*|HKEY_')]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [parameter(Mandatory = $true,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [parameter(Position = 3)]
        $ValueData
    )

    Set-StrictMode -Version 2.0

    #Add Regdrive if it is not present
    if ($Path -notmatch 'Registry::.*'){
        $Path = 'Registry::' + $Path
    }

    try {
        #Reg key with value
        if ($ValueData) {
            if ((Get-ItemProperty -Path $Path -ErrorAction Stop | Select-Object -ExpandProperty $Name -ErrorAction Stop) -eq $ValueData) {
                return $true
            }
            else {
                return $false
            }
        }
        #Key key without value
        else {
            $RegKeyCheck = Get-ItemProperty -Path $Path -ErrorAction Stop | Select-Object -ExpandProperty $Name -ErrorAction Stop
            if ($null -eq $RegKeyCheck) {
                #if the Key Check returns null then it probably means that the key does not exist.
                return $false
            }
            else {
                return $true
            }
        }
    }
    catch {
        return $false
    }
}

Write-Output "Check if the computer is an RDS Server"

$TSMode = $string = Get-WmiObject -Namespace "root\CIMV2\TerminalServices" -Class "Win32_TerminalServiceSetting"  | select -ExpandProperty TerminalServerMode

if ($TSMode -eq '1') {
    #Check to make sure UPDs are enabled
    if((Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Terminal Server\ClusterSettings' -Name UvhdEnabled -Value 1) -or
	(Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\FSLogix\Profiles' -Name 'Enabled' -ValueData 1)){

        #region createScheduledTask
        $ScheduledTaskAction = New-ScheduledTaskAction -Execute '%SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument '-ExecutionPolicy Bypass C:\BIN\DelProf2\Remove-LocalUPDProfiles.ps1'

        $ScheduledTaskTrigger = @(
            $(New-ScheduledTaskTrigger -Daily -At 4AM),
            $(New-ScheduledTaskTrigger -AtStartup )
        )

        $ScheduledTaskPrincipal = New-ScheduledTaskPrincipal -UserId 'SYSTEM'

        $ScheduledTaskSettingsSet = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 1)

        $ScheduledTask = New-ScheduledTask -Action $ScheduledTaskAction -Trigger $ScheduledTaskTrigger -Principal $ScheduledTaskPrincipal -Settings $ScheduledTaskSettingsSet

        Register-ScheduledTask -TaskName $ScheduledTaskName -InputObject $ScheduledTask
        #endregion createScheduledTask
        
        Write-Host 'Make sure that C:\BIN\DelProf2\Remove-LocalUPDProfiles.ps1 and C:\BIN\DelProf2\DelProf2.exe exist' -BackgroundColor Yellow -ForegroundColor Black
    }
    else {
        Write-Output "UPDs are not enabled. The script will not run"
    }
}
else {
    Write-Output "The computer is not an RDS Server. The script will not run"
}
