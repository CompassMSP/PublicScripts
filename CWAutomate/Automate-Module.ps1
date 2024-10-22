<#
.SYNOPSIS
    These PowerShell Functions will Install, Push, Uninstall, and Confirm ConnectWise Automate installations.

.DESCRIPTION
    Functions Included:
        Confirm-Automate
        Uninstall-Automate
        Install-Automate
        Push-Automate
        Get-ADComputerNames
        Install-Chrome
        Install-Manage
        Scan-Network

        New-IPRange
        http://powershell.com/cs/media/p/9437.aspx

        Invoke-Ping
        https://gallery.technet.microsoft.com/scriptcenter/Invoke-Ping-Test-in-b553242a

        Get-IPv4Subnet
        https://github.com/briansworth/GetIPv4Address/blob/master/GetIPv4Subnet.psm1

.LINK
    https://github.com/Braingears/PowerShell

.NOTES
    File Name      : Automate-Module.psm1
    Author         : Chuck Fowler (Chuck@Braingears.com)
    Version        : 1.0
    Creation Date  : 11/10/2019
    Purpose/Change : Initial script development
    Prerequisite   : PowerShell V2

    Version        : 1.1
    Date           : 11/15/2019
    Changes        : Add $Automate.InstFolder and $Automate.InstRegistry and check for both to be consdered for $Automate.Installed
                     It was found that the Automate Uninstaller EXE is leaving behind the LabTech registry keys and it was not being detected properly.


.EXAMPLE
    Confirm-Automate [-Silent]

    Confirm-Automate [-Show]

.EXAMPLE
    Uninstall-Automate [-Silent]

.EXAMPLE
    Install-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 [-Show]

.Example
    To push a single Automate Agent:
    Push-Automate -Computer 'ComputerName' -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'

    For multiple computers, use a | "pipe" into Push-Automate function:
    $Computers | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'
    - or -
     Scan-Network | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'
    - or -
    Get-ADComputerNames | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'
    - or -
    "Computer1", "Computer2" | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'

#>
Function Confirm-Automate {
<#
.SYNOPSIS
    This PowerShell Function will confirm If Automate is installed, services running, and checking-in.

.DESCRIPTION
    This function will automatically start the Automate services (If stopped). It will collect Automate information from the registry.

.PARAMETER Raw
    This will show the Automate registry entries

.PARAMETER Show
    This will display $Automate object

.PARAMETER Silent
    This will hide all output

.LINK
    https://github.com/Braingears/PowerShell

.NOTES
    Version        : 1.0
    Author         : Chuck Fowler
    Creation Date  : 08/16/2019
    Purpose/Change : Initial script development

    Version        : 1.1
    Date           : 11/15/2019
    Changes        : Add $Automate.InstFolder and $Automate.InstRegistry and check for both to be consdered for $Automate.Installed
                     It was found that the Automate Uninstaller EXE is leaving behind the LabTech registry keys and it was not being detected properly.


.EXAMPLE
    Confirm-Automate [-Silent]

    Confirm-Automate [-Show]

    ServerAddress : https://yourserver.hostedrmm.com
    ComputerID    : 321
    ClientID      : 1
    LocationID    : 2
    Version       : 190.221
    Service       : Running
    Online        : True
    LastHeartbeat : 29
    LastStatus    : 36

    $Automate
    $Global:Automate
    This output will be saved to $Automate as an object to be used in other functions.


#>
 [CmdletBinding(SupportsShouldProcess=$True)]
    Param (
        [switch]$Raw    = $False,
        [switch]$Show   = $False,
        [switch]$Silent = $False
    )
    $ErrorActionPreference = 'SilentlyContinue'
    $Online = If ((Test-Path "HKLM:\SOFTWARE\LabTech\Service") -and ((Get-Service ltservice).status) -eq "Running") {((( (Get-Date) - [System.DateTime](Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service").LastSuccessStatus).TotalSeconds) -lt 600)} Else {Write $False}
    If (Test-Path "HKLM:\SOFTWARE\LabTech\Service") {
        $Global:Automate = New-Object -TypeName psobject
        $Global:Automate | Add-Member -MemberType NoteProperty -Name ComputerName -Value $env:ComputerName
        $Global:Automate | Add-Member -MemberType NoteProperty -Name ServerAddress -Value ((Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service").'Server Address')
        $Global:Automate | Add-Member -MemberType NoteProperty -Name ComputerID -Value ((Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service").ID)
        $Global:Automate | Add-Member -MemberType NoteProperty -Name ClientID -Value ((Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service").ClientID)
        $Global:Automate | Add-Member -MemberType NoteProperty -Name LocationID -Value ((Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service").LocationID)
        $Global:Automate | Add-Member -MemberType NoteProperty -Name Version -Value ((Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service").Version)
        $Global:Automate | Add-Member -MemberType NoteProperty -Name InstFolder -Value (Test-Path "$($env:windir)\ltsvc")
        $Global:Automate | Add-Member -MemberType NoteProperty -Name InstRegistry -Value $True
        $Global:Automate | Add-Member -MemberType NoteProperty -Name Installed -Value (Test-Path "$($env:windir)\ltsvc")
        $Global:Automate | Add-Member -MemberType NoteProperty -Name Service -Value ((Get-Service LTService).Status)
        $Global:Automate | Add-Member -MemberType NoteProperty -Name Online -Value $Online
        $Global:Automate | Add-Member -MemberType NoteProperty -Name LastHeartbeat -Value (((Get-Date) - [System.DateTime](Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service").HeartbeatLastReceived).TotalSeconds)
        $Global:Automate | Add-Member -MemberType NoteProperty -Name LastStatus -Value (((Get-Date) - [System.DateTime](Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service").LastSuccessStatus).TotalSeconds)
        Write-Verbose $Global:Automate
        If ($Show) {
            $Global:Automate
        } Else {
            If (!$Silent) {
                Write "Server Address checking-in to    $($Global:Automate.ServerAddress)"
                Write "ComputerID:                      $($Global:Automate.ComputerID)"
                Write "The Automate Agent Online        $($Global:Automate.Online)"
                Write "Last Successful Heartbeat        $($Global:Automate.LastHeartbeat) seconds"
                Write "Last Successful Status Update    $($Global:Automate.LastStatus) seconds"
            } # End Not Silent
        } # End If
        If ($Raw -eq $True) {Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service"}
    } Else {
        $Global:Automate = New-Object -TypeName psobject
        $Global:Automate | Add-Member -MemberType NoteProperty -Name ComputerName -Value $env:ComputerName
        $Global:Automate | Add-Member -MemberType NoteProperty -Name InstFolder -Value (Test-Path "$($env:windir)\ltsvc")
        $Global:Automate | Add-Member -MemberType NoteProperty -Name InstRegistry -Value $False
        $Global:Automate | Add-Member -MemberType NoteProperty -Name Installed -Value ((Test-Path "$($env:windir)\ltsvc") -and (Test-Path "HKLM:\SOFTWARE\LabTech\Service"))
        $Global:Automate | Add-Member -MemberType NoteProperty -Name Service -Value ((Get-Service ltservice ).status)
        $Global:Automate | Add-Member -MemberType NoteProperty -Name Online -Value $Online
        Write-Verbose $Global:Automate
    } #End If Registry Exists
    If (!$Global:Automate.InstFolder -and !$Global:Automate.InstRegistry) {If ($Silent -eq $False) {Write "Automate is NOT Installed"}}
} #End Function Confirm-Automate
########################
Set-Alias -Name LTC -Value Confirm-Automate -Description 'Confirm If Automate is running properly'
########################
Function Uninstall-Automate {
<#
.SYNOPSIS
    This PowerShell Function Uninstall Automate.

.DESCRIPTION
    This function will download the Automate Uninstaller from Connectwise and completely remove the Automate / LabTech Agent.

.PARAMETER Silent
    This will hide all output

.LINK
    https://github.com/Braingears/PowerShell

.NOTES
    Version        : 1.0
    Author         : Chuck Fowler
    Website        : braingears.com
    Creation Date  : 8/2019
    Purpose        : Create initial function script

    Version        : 1.1
    Date           : 11/15/2019
    Changes        : Add $Automate.InstFolder and $Automate.InstRegistry and check for both to be consdered for $Automate.Installed
                     It was found that the Automate Uninstaller EXE is leaving behind the LabTech registry keys and it was not being detected properly.
                     If the LTSVC Folder or Registry keys are found after the uninstaller runs, the script now performs a manual gutting via PowerShell.


.EXAMPLE
    Uninstall-Automate [-Silent]


#>
[CmdletBinding(SupportsShouldProcess=$True)]
    Param (
     [switch]$Force,
     [switch]$Raw,
     [switch]$Show,
     [switch]$Silent = $False
     )
$ErrorActionPreference = 'SilentlyContinue'
$Verbose = If ($PSBoundParameters.Verbose -eq $True) { $True } Else { $False }
$DownloadPath = "https://s3.amazonaws.com/assets-cp/assets/Agent_Uninstall.exe"
If (([int]((Get-WmiObject Win32_OperatingSystem).BuildNumber) -gt 6000) -and ((get-host).Version.ToString() -ge 3)) {
    $DownloadPath = "https://s3.amazonaws.com/assets-cp/assets/Agent_Uninstall.exe"
} Else {
    $DownloadPath = "http://s3.amazonaws.com/assets-cp/assets/Agent_Uninstall.exe"
}
$SoftwarePath = "C:\Support\Automate"
Write-Debug "Checking if Automate Installed"
Confirm-Automate -Silent -Verbose:$Verbose
    If (($Global:Automate.InstFolder) -or ($Global:Automate.InstRegistry) -or ($Force)) {
    $Filename = [System.IO.Path]::GetFileName($DownloadPath)
    $SoftwareFullPath = "$($SoftwarePath)\$Filename"
    If (!(Test-Path $SoftwarePath)) {md $SoftwarePath | Out-Null}
    Set-Location $SoftwarePath
    If ((Test-Path $SoftwareFullPath)) {Remove-Item $SoftwareFullPath | Out-Null}
    $WebClient = New-Object System.Net.WebClient
    $WebClient.DownloadFile($DownloadPath, $SoftwareFullPath)
    If (!$Silent) {Write-Host "Removing Existing Automate Agent..."}
    Write-Verbose "Closing Open Applications and Stopping Services"
    Stop-Process -Name "ltsvcmon","lttray","ltsvc","ltclient" -Force
    Stop-Service ltservice,ltsvcmon -Force
    $UninstallExitCode = (Start-Process "cmd" -ArgumentList "/c $($SoftwareFullPath)" -NoNewWindow -Wait -PassThru).ExitCode
    If (!$Silent) {
        If ($UninstallExitCode -eq 0) {
          # Write-Host "The Automate Agent Uninstaller Executed Without Errors" -ForegroundColor Green
            Write-Verbose "The Automate Agent Uninstaller Executed Without Errors"
        } Else {
            Write-Host "Automate Uninstall Exit Code: $($UninstallExitCode)" -ForegroundColor Red
            Write-Verbose "Automate Uninstall Exit Code: $($UninstallExitCode)"
        }
    }
    Write-Verbose "Checking For Removal - Loop 3X"
    While ($Counter -ne 3) {
        $Counter++
        Start-Sleep 10
        Confirm-Automate -Silent -Verbose:$Verbose
        If ((!$Global:Automate.InstFolder) -and (!$Global:Automate.InstRegistry)) {
            Write-Verbose "Automate Uninstaller Completed Successfully"
            Break
        }
    }# end While
    If (($Global:Automate.InstFolder) -or ($Global:Automate.InstRegistry)) {
        Write-Verbose "Uninstaller Failed"
        Write-Verbose "Manually Gutting Automate..."
        Stop-Process -Name "ltsvcmon","lttray","ltsvc","ltclient" -Force
        Stop-Service ltservice,ltsvcmon -Force
        Write-Verbose "Uninstalling LabTechAD Package"
        Start-Process "msiexec.exe" -ArgumentList "/x {3F460D4C-D217-46B4-80B6-B5ED50BD7CF5} /qn" -NoNewWindow -Wait -PassThru | Out-Null
        Remove-Item "$($env:windir)\ltsvc" -Recurse -Force
        Get-ItemProperty "HKLM:\SOFTWARE\LabTech\LabVNC" | Remove-Item -Recurse -Force
        Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" | Remove-Item -Recurse -Force
        Start-Process "cmd" -ArgumentList "/c $($SoftwareFullPath)" -NoNewWindow -Wait -PassThru | Out-Null
        Confirm-Automate -Silent -Verbose:$Verbose
        If ($Global:Automate.InstFolder) {
            If (!$Silent) {
                Write-Host "Automate Uninstall Failed" -ForegroundColor Red
                Write-Host "$($env:windir)\ltsvc folder still exists" -ForegroundColor Red
            } else {
                Write-Verbose "Automate Uninstall Failed"
                Write-Verbose "$($env:windir)\ltsvc folder still exists"
            }
            If ($Global:Automate.InstRegistry) {
                Write-Host "Automate Uninstall Failed" -ForegroundColor Red
                Write-Host "HKLM:\SOFTWARE\LabTech\Service Registry keys still exists" -ForegroundColor Red
            } else {
                Write-Verbose "Automate Uninstall Failed"
                Write-Verbose "HKLM:\SOFTWARE\LabTech\Service Registry keys still exists"
            }
        }
    } Else {
        If (!$Silent) {Write-Host "The Automate Agent Uninstalled Successfully" -ForegroundColor Green}
        Write-Verbose "The Automate Agent Uninstalled Successfully"
    }
} # If Test-Install
    Confirm-Automate -Silent:$Silent
} # Function Uninstall-Automate
########################
Set-Alias -Name LTU -Value Uninstall-Automate -Description 'Uninstall Automate Agent'
########################
Function Install-Automate {
<#
.SYNOPSIS
    This PowerShell Function is for Automate Deployments

.DESCRIPTION
    Install the Automate Agent.

    This function will qualIfy the If another Autoamte agent is already
    installed on the computer. If the existing agent belongs to dIfferent
    Automate server, it will automatically "Rip & Replace" the existing
    agent. This comparison is based on the server's FQDN.

    This function will also verIfy If the existing Automate agent is
    checking-in. The Confirm-Automate Function will verIfy the Server
    address, LocationID, and Heartbeat/Check-in. If these entries are
    missing or not checking-in properly; this function will automatically
    attempt to restart the services, and then "Rip & Replace" the agent to
    remediate the agent.

    $Automate
    $Global:Automate
    The output will be saved to $Automate as an object to be used in other functions.

    Example:
    Install-Automate -Server YOURSERVER.DOMAIN.COM -LocationID 2 -Transcript


    Tested OS:      Windows XP (with .Net 3.5.1 and PowerShell installed)
                    Windows Vista
                    Windows 7
                    Windows 8
                    Windows 10
                    Windows 2003R2
                    Windows 2008R2
                    Windows 2012R2
                    Windows 2016
                    Windows 2019

.PARAMETER Server
    This is the URL to your Automate server.

        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2

.PARAMETER LocationID
    Use LocationID to install the Automate Agent directly to the appropieate client's location / site.
    If parameter is not specIfied, it will automatically assign LocationID 1 (New Computers).

.PARAMETER Force
    This will force the Automate Uninstaller prior to installation.
    Essentually, this will be a fresh install and a fresh check-in to the Automate server.

        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2 -Force

.PARAMETER Silent
    This will hide all output (except a failed installation when Exit Code -ne 0)
    The function will exit once the installer has completed.

        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2 -Silent

.PARAMETER Transcript
    This parameter will save the entire transcript and responsed to:
    $($env:windir)\Temp\AutomateLogon.txt

        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2 -Transcript -Verbose

.LINK
    https://github.com/Braingears/PowerShell

.NOTES
    Version        : 1.0
    Author         : Chuck Fowler
    Creation Date  : 08/2019
    Purpose/Change : Initial script development

    Version        : 1.1
    Date           : 11/15/2019
    Changes        : Add $Automate.InstFolder and $Automate.InstRegistry and check for both to be consdered for $Automate.Installed
                     It was found that the Automate Uninstaller EXE is leaving behind the LabTech registry keys and it was not being detected properly.
                     If the LTSVC Folder or Registry keys are found after the uninstaller runs, the script now performs a manual gutting via PowerShell.

    Version        : 1.2
    Date           : 02/17/2020
    Changes        : Add MSIEXEC Log Files to C:\Windows\Temp\Automate_Agent_(Date).log

.EXAMPLE
    Install-Automate -Server 'automate.domain.com' -LocationID 42
    This will install the LabTech agent using the provided Server URL, and LocationID.


#>
[CmdletBinding(SupportsShouldProcess=$True)]
    Param(
        [Parameter(ValueFromPipelineByPropertyName = $True, Position=0)]
        [Alias("FQDN","Srv")]
        [string[]]$Server = $Null,
        [Parameter(ValueFromPipelineByPropertyName = $True)]
        [AllowNull()]
        [Alias('LID','Location')]
        [int]$LocationID = '1',
        [Parameter()]
        [AllowNull()]
        [switch]$Force,
        [Parameter()]
        [AllowNull()]
        [switch]$Show = $False,
        [switch]$Silent,
        [Parameter()]
        [AllowNull()]
        [switch]$Transcript = $False
    )
    $ErrorActionPreference = 'SilentlyContinue'
    $Verbose = If ($PSBoundParameters.Verbose -eq $True) { $True } Else { $False }
    $Error.Clear()
    If ($Transcript) {Start-Transcript -Path "$($env:windir)\Temp\Automate_Deploy.txt" -Force}
    Write-Verbose "Checking Operating System (WinXP and Older) for HTTP vs HTTPS"
    If (([int]((Get-WmiObject Win32_OperatingSystem).BuildNumber) -gt 6000) -and ((get-host).Version.ToString() -ge 3)) {$AutomateURL = "https://$($Server)"} Else {$AutomateURL = "http://$($Server)"}
    $AutomateURLTest = "$($AutomateURL)/LabTech/"
    $SoftwarePath = "C:\Support\Automate"
    $DownloadPath = "$($AutomateURL)/Labtech/Deployment.aspx?Probe=1&installType=msi&MSILocations=$($LocationID)"
    $Filename = "Automate_Agent.msi"
    $SoftwareFullPath = "$SoftwarePath\$Filename"
    Write-Verbose "Checking if Automate Server URL is active. Server entered: $($Server)"
    Try {
        If ((get-host).Version.ToString() -ge 3) {
            $TestURL = (New-Object Net.WebClient).DownloadString($AutomateURLTest)
            Write-Verbose "$AutomateURL is Active"
        }
    }
    Catch {
        Write-Host "The Automate Server Parameter Was Not Entered or Inaccessable" -ForegroundColor Red
        Write-Host "Help: Get-Help Install-Automate -Full"
        Write-Host " "
        Confirm-Automate -Show
        Break
        }
    Confirm-Automate -Silent -Verbose:$Verbose
    Write-Verbose "If ServerAddress matches, the Automate Agent is currently Online, and Not forced to Rip & Replace then Automate is already installed."
    Write-Verbose (($Global:Automate.ServerAddress -like "*$($Server)*") -and ($Global:Automate.Online) -and !($Force))
    If (($Global:Automate.ServerAddress -like "*$($Server)*") -and $Global:Automate.Online -and !$Force) {
        If (!$Silent) {
            If ($Show) {
              $Global:Automate
            } Else {
              Write-Host "The Automate Agent is already installed and checked-in $($Global:Automate.LastStatus) seconds ago to $($Global:Automate.ServerAddress)" -ForegroundColor Green
            }
        }
    } Else {
        If (!$Silent -and $Global:Automate.Online -and (!($Global:Automate.ServerAddress -like "*$($Server)*"))) {
            Write-Host "The Existing Automate Server Does Not Match The Target Automate Server." -ForegroundColor Red
            Write-Host "Current Automate Server: $($Global:Automate.ServerAddress)" -ForegroundColor Red
            Write-Host "New Automate Server:     $($AutomateURL)" -ForegroundColor Green
        } # If Different Server
        Write-Verbose "Removing Existing Automate Agent"
        Uninstall-Automate -Force:$Force -Silent:$Silent -Verbose:$Verbose
        Write-Verbose "Installing Automate Agent on $($AutomateURL)"
            If (!(Test-Path $SoftwarePath)) {md $SoftwarePath | Out-Null}
            Set-Location $SoftwarePath
            If ((test-path $SoftwareFullPath)) {Remove-Item $SoftwareFullPath | Out-Null}
            $WebClient = New-Object System.Net.WebClient
            $WebClient.DownloadFile($DownloadPath, $SoftwareFullPath)
            If (!$Silent) {Write-Host "Installing Automate Agent to $AutomateURL"}
            Stop-Process -Name "ltsvcmon","lttray","ltsvc","ltclient" -Force -PassThru
            $Date = (get-date -UFormat %Y-%m-%d_%H-%M-%S)
            $LogFullPath = "$env:windir\Temp\Automate_Agent_$Date.log"
            $InstallExitCode = (Start-Process "msiexec.exe" -ArgumentList "/i $($SoftwareFullPath) /quiet /norestart LOCATION=$($LocationID) /L*V $($LogFullPath)" -NoNewWindow -Wait -PassThru).ExitCode
            Write-Verbose "MSIEXEC Log Files: $LogFullPath"
            If ($InstallExitCode -eq 0) {
                If (!$Silent) {Write-Verbose "The Automate Agent Installer Executed Without Errors"}
            } Else {
                Write-Host "Automate Installer Exit Code: $InstallExitCode" -ForegroundColor Red
                Write-Host "Automate Installer Logs: $LogFullPath" -ForegroundColor Red
                Write-Host "The Automate MSI failed. Waiting 15 Seconds..." -ForegroundColor Red
                Start-Sleep -s 15
                Write-Host "Installer will execute twice (KI 12002617)" -ForegroundColor Yellow
                $Date = (get-date -UFormat %Y-%m-%d_%H-%M-%S)
                $LogFullPath = "$env:windir\Temp\Automate_Agent_$Date.log"
                $InstallExitCode = (Start-Process "msiexec.exe" -ArgumentList "/i $($SoftwareFullPath) /quiet /norestart LOCATION=$($LocationID) /L*V $($LogFullPath)" -NoNewWindow -Wait -PassThru).ExitCode
                Write-Host "Automate Installer Exit Code: $InstallExitCode" -ForegroundColor Yellow
                Write-Host "Automate Installer Logs: $LogFullPath" -ForegroundColor Yellow
            }# End Else
        While ($Counter -ne 30) {
            $Counter++
            Start-Sleep 10
            Confirm-Automate -Silent -Verbose:$Verbose
            If ($Global:Automate.Online -and $Global:Automate.ComputerID -ne $Null) {
                If (!$Silent) {
                    Write-Host "The Automate Agent Has Been Successfully Installed" -ForegroundColor Green
                    $Global:Automate
                }#end If Silent
                Break
            } # end If
        }# end While
    } # End
    If ($Transcript) {Stop-Transcript}
} #End Function Install-Automate
########################
Set-Alias -Name LTI -Value Install-Automate -Description 'Install Automate Agent'
########################
Function Push-Automate
{
<#
.SYNOPSIS
    This PowerShell Function is for pushing Automate Deployments

.DESCRIPTION
    Install the Automate Agent.

    This function will qualIfy the If another Autoamte agent is already
    installed on the computer. If the existing agent belongs to dIfferent
    Automate server, it will automatically "Rip & Replace" the existing
    agent. This comparison is based on the server's FQDN.

    This function will also verIfy If the existing Automate agent is
    checking-in. The Confirm-Automate Function will verIfy the Server
    address, LocationID, and Heartbeat/Check-in. If these entries are
    missing or not checking-in properly; this function will automatically
    attempt to restart the services, and then "Rip & Replace" the agent to
    remediate the agent.

    $AutoResults
    $Global:AutoResults
    The output will be saved to $AutoResults as an object to be used in other functions.

    Example:
    To push a single Automate Agent:
    Push-Automate -Computer 'Computername' -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'

    For multiple computers, use a | "pipe" into Push-Automate function:
    $Computers | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'
    - or -
    Get-ADComputerNames | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'
    - or -
    "Computer1", "Computer2" | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'

    When pushing to multiple computers, use the actual computer names. If you use IP Address, it will fail when using WINRM Protocols (and use WMI/RCP instead).

.PARAMETER Server
    This is the URL to your Automate server.

        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2

.PARAMETER LocationID
    Use LocationID to install the Automate Agent directly to the appropieate client's location / site.
    If parameter is not specIfied, it will automatically assign LocationID 1 (New Computers).
        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2

.PARAMETER Username
    Enter username with Domain Admin rights. When entering username, use 'DOMAIN\USERNAME'

    The function will accept PSCredentials saved to $Credentials prior to running this function.

.PARAMETER Password
    Enter Password for Domain Admin account.

    The function will accept PSCredentials saved to $Credentials prior to running this function.

.PARAMETER Force
    >>> This Function Is Currently Disabled <<<
    This will force the Automate Uninstaller prior to installation.
    Essentually, this will be a fresh install and a fresh check-in to the Automate server.

        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2 -Force

.PARAMETER Silent
    >>> This Function Is Currently Disabled <<<
    This will hide all output (except a failed installation when Exit Code -ne 0)
    The function will exit once the installer has completed.

        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2 -Silent

.PARAMETER Transcript
    >>> This Function Is Currently Disabled <<<
    This parameter will save the entire transcript and responsed to:
    $($env:windir)\Temp\AutomateLogon.txt

        Install-Automate -Server 'server.hostedrmm.com' -LocationID 2 -Transcript -Verbose

.LINK
    https://github.com/Braingears/PowerShell

.NOTES
    Version        : 1.0
    Author         : Chuck Fowler
    Creation Date  : 08/2019
    Purpose/Change : Initial script development

    Version        : 1.1
    Date           : 11/15/2019
    Changes        : Add $Automate.InstFolder and $Automate.InstRegistry and check for both to be consdered for $Automate.Installed
                     It was found that the Automate Uninstaller EXE is leaving behind the LabTech registry keys and it was not being detected properly.
                     If the LTSVC Folder or Registry keys are found after the uninstaller runs, the script now performs a manual gutting via PowerShell.


.EXAMPLE
    Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd' -Computer

    Use the -Computer parameter for single computers.

.EXAMPLE
    $Computers | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'

    Use Array to pipe multiple computers into Push=Automate function.

.EXAMPLE
    Get-ADComputerNames | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'

    Use another function to pipe multiple computers into Push=Automate function. Select only computer names.

.EXAMPLE
    "Computer1", "Computer2" | Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2 -Username 'DOMAIN\USERNAME' -Password 'Ch@ng3P@ssw0rd'

    When pushing to multiple computers, use the actual computer names. If you use IP Address, it will fail when using WINRM Protocols (and use WMI/RCP instead).
    This will install the LabTech agent using the provided Server URL, and LocationID.

.EXAMPLE
    $Credential = Get-Credential
    Push-Automate -Server 'YOURSERVER.DOMAIN.COM' -LocationID 2

    You can proactivly load PSCredential, then use the Push-Automate function within the same Powershell session.


#>
[CmdletBinding()]
Param
(
    [Parameter(ValueFromPipeline=$True)]
    [string[]]$Computer = $env:COMPUTERNAME,
    [Parameter()]
    [Alias("FQDN","Srv")]
    [string[]]$Server = $Null,
    [Parameter()]
    [AllowNull()]
    [Alias('LID','Location')]
    [int]$LocationID = '1',
    [Parameter()]
    [AllowNull()]
    [Alias('User')]
    [string[]]$Username,
    [Parameter()]
    [AllowNull()]
    [Alias('Pass')]
    [string[]]$Password,
    [Parameter()]
    [AllowNull()]
    [switch]$Force = $False,
    [Parameter()]
    [AllowNull()]
    [switch]$Show = $False,
    [Parameter()]
    [AllowNull()]
    [switch]$Silent = $False,
    [Parameter()]
    [AllowNull()]
    [switch]$Transcript = $False
)
BEGIN
{
    $ErrorActionPreference = "SilentlyContinue"
    $Verbose = If ($PSBoundParameters.Verbose -eq $True) { $True } Else { $False }
    Write-Verbose "Checking if Automate Server URL is active. Server entered: $($Server)"
    $AutomateURLTest = "https://$($Server)/LabTech/"
    Write-Verbose "$AutomateURLTest"
    Try {
        $TestURL = (New-Object Net.WebClient).DownloadString($AutomateURLTest)
        Write-Verbose "https://$($Server) is Active"
    }
    Catch {
        Write-Host "The Automate Server Parameter Was Not Entered or Inaccessable" -ForegroundColor Red
        Write-Host "Help: Get-Help Push-Automate -Full"
        Break
        }
    $Whoami = whoami
    Write-Verbose "Running Script as: $whoami"
    If (($Username -eq $Null) -and ($Password -eq $Null) -and ($Credential -eq $Null) -and !((whoami) -eq 'nt authority\system'))
        {$Credential = Get-Credential -Message "Enter Domain Admin Credentials for Remote Automate Push"}
    If (($Username -ne $Null) -and ($Password -ne $Null)) {
        $Pass = $Password | ConvertTo-SecureString -asPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential($Username,$Pass)
    }
    If ($Credential -eq $Null) {
        If ((whoami) -eq 'nt authority\system') {Write-Host "Running function as $($Whoami)"}
        Write-Host "Credentials Are Missing!" -ForegroundColor Red
        Clear-Variable Computer, Server, Force, Silent
        Break
    }
    Write-Verbose "Credential loaded: $($Credential.Username)"
    $Global:AutoChecks = @()
} #End Begin
PROCESS
{
    # Variables
    $Time = Date
    $CheckAutomateWinRM = {
        Write-Verbose "Invoke Confirm-Automate -Silent"
        Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/CWAutomate/Automate-Module.psm1')
        Confirm-Automate -Silent
        Write $Global:Automate
    }
    $InstallAutomateWinRM = {
        $Server = $Args[0]
        $LocationID = $Args[1]
        $Force = $Args[2]
        $Silent = $Args[3]
        $Transcript = $Args[4]
        Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/CWAutomate/Automate-Module.psm1')
        Install-Automate -Server $Server -LocationID $LocationID -Transcript
    }
    $WMICMD = 'powershell.exe -Command "Invoke-Expression(New-Object Net.WebClient).DownloadString(''https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/CWAutomate/Automate-Module.psm1''); '
    $WMIPOSH = "Install-Automate -Server $Server -LocationID $LocationID -Transcript"
    $WMIArg = Write-Output "$WMICMD$WMIPOSH"""
    $WinRMConectivity = "N/A"
    $WMICConectivity = "N/A"
    $WinRMDeployed = $False
    $WMIDeployed = $False
    Clear-Variable Automate, ProcessErrorWinRM, ProcessErrorWMIC
    # End Variables
    # Now Trying WinRM
    If ($Computer -eq $env:COMPUTERNAME) {
        Write-Verbose "Installing Automate on Local Computer - $Computer"
        Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/CWAutomate/Automate-Module.psm1')
        Install-Automate -Server $Server -LocationID $LocationID -Show:$Show -Transcript:$Transcript
    } Else {        # Remote Computer
        If (!$Silent) {Write-Host "$($Time) - Now Checking $($COMPUTER)"}
        Write-Verbose "Ping Connectivity  - Testing..."
        If (Test-Connection -ComputerName $COMPUTER -Count 1 -Quiet) {
            Write-Verbose "Ping Connectivity  - Passed"
            $PingTest = $True
            Write-Verbose "IP or NetBIOS Name - Testing..."
            If ($Computer -notmatch "[a-z]") {
                Write-Verbose "$Computer is IP Address"
                $ComputerNetBIOS = nbtstat -A $Computer | Where-Object { $_ -match '^\s*([^<\s]+)\s*<00>\s*UNIQUE' } | ForEach-Object { $matches[1] }
                if ($ComputerNetBIOS -eq $Null) {
                    Write-Verbose "$Computer could not query NetBIOS Name"
                    Write-Verbose "$Computer as an IP Address will likely fail WinRM Connectivity"
                } else {
                    Write-Verbose "Replacing $Computer with $ComputerNetBIOS"
                    $Computer = $ComputerNetBIOS
                }
            }
            Try {
                Write-Verbose "Proactively Remote Starting WinRM Service"
                Get-Service WinRM -ComputerName $Computer -ErrorAction Stop | Start-Service
            }
            Catch {Write-Verbose "Start Service      - Failed"}
            Try {
                # The $WinRMConectivity will change to $True if the Invoke-Command has no $Errors
                $WinRMConectivity = $False
                $WinRMFailed = $True
                Write-Verbose "WinRM Connectivity - Testing..."
                $Global:Automate = (Invoke-Command $COMPUTER -Credential $Credential -ScriptBlock $CheckAutomateWinRM -ErrorAction Stop -ErrorVariable ProcessErrorWinRM)
                Write-Verbose "WinRM Connectivity - Passed"
                Write-Verbose "Global Automate: $($Global:Automate)"
                $WinRMConectivity = $True
                $WinRMFailed = $False
            }
            Catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
                If ($($ProcessErrorWinRM) -like "*Logon failure*") {
                    Write-Verbose "WinRM Connectivity - Credentials Failed"
                    Write-Host    "WinRM Connectivity - Credentials Failed" -ForegroundColor Red
                } else {
                    Write-Verbose "WinRM Connectivity - Failed"
                }
            }
            Catch {
                Write-Verbose "WinRM Connectivity - Failed"
                Write-Verbose "WinRM Errors: $ProcessErrorWinRM.Exception"
                $ProcessErrorWinRM.Exception | Select -Property *
            }
            If (($Global:Automate.ServerAddress -like "*$($Server)*") -and $WinRMConectivity -and $Global:Automate.Online -and !$Force) {
                If ($Show) {
                    $Global:Automate
                } Else {
                    Write-Host "The Automate Agent is already installed and checked-in $($Global:Automate.LastStatus) seconds ago to $($Global:Automate.ServerAddress)" -ForegroundColor Green
                }
            } Else {
                If ($WinRMConectivity) {
                    Write-Verbose "WinRM Connectivity - Passed"
                    Write-Verbose "Installing Automate..."
                    Invoke-Command $COMPUTER -Credential $Credential -ScriptBlock $InstallAutomateWinRM -ArgumentList $Server, $LocationID, $Force, $Silent, $Transcript -ErrorAction SilentlyContinue
                    $Global:Automate = (Invoke-Command $COMPUTER -Credential $Credential -ScriptBlock $CheckAutomateWinRM -ErrorAction SilentlyContinue)
                    Write-Verbose "Local Automate:  $($Automate)"
                    Write-Verbose "Global Automate: $($Global:Automate)"
                    $WinRMDeployed = $True
                }
            }
        #### Now Trying RPC
            If (!$Global:Automate.Online) {
                Write-Verbose "WMIC Connectivity  - Testing..."
                Try {
                    $WMICFailed = $True
                    $WMICConectivity = $False
                    $ComputerWMI = ((Get-WmiObject -ComputerName $Computer -Class Win32_ComputerSystem -Credential $Credential -ErrorAction Stop -ErrorVariable ProcessErrorWMIC).Name)
                    Write-Verbose "WMIC Connectivity  - Passed"
                    $WMICConectivity = $True
                    $WMICFailed = $False
                }
                Catch [System.Runtime.InteropServices.COMException] {
                    Write-Verbose "WMIC Connectivity  - RPC Server is Unavailable"
                }
                Catch [System.UnauthorizedAccessException] {
                    Write-Verbose "WMIC Connectivity  - Credentials Failed"
                    Write-Host    "WMIC Connectivity  - Credentials Failed" -ForegroundColor Red
                }
                Catch {
                    Write-Verbose "WMIC Connectivity  - Failed"
                    Write-Verbose "WMIC Errors: $ProcessErrorWMIC"
                }
                If   ($WMICConectivity) {
                    $Reg = Get-WmiObject -List StdRegProv -Namespace root\default -ComputerName $Computer -Credential $Credential
                    $HKLM = 2147483650
                    $Key = 'SOFTWARE\LabTech\Service\'
                    $Values = $Reg.EnumValues($HKLM,$Key)
                    # Registry types enumerations:
                    $RegTypes = @{
                        1 = 'REG_SZ'
                        2 = 'REG_EXPAND_SZ'
                        3 = 'REG_BINARY'
                        4 = 'REG_DWORD'
                        7 = 'REG_MULTI_SZ'
                    }
                    # Use a for loop to go through the values
                    $Results = @(
                        for ($i = 0; $i -lt $Values.sNames.count; $i++) {
                            $Name = $Values.sNames[$i]
                            $Type = $RegTypes[$Values.Types[$i]]
                            switch ($Values.Types[$i]) {
                                1 {$Value = $Reg.GetStringValue($HKLM,$Key,$Name).sValue}
                                2 {$Value = $Reg.GetExpandedStringValue($HKLM,$Key,$Name).sValue}
                                3 {$Value = $Reg.GetBinaryValue($HKLM,$Key,$Name).uValue}
                                4 {$Value = $Reg.GetDWORDValue($HKLM,$Key,$Name).uValue}
                                7 {$Value = $Reg.GetMultiStringValue($HKLM,$Key,$Name).sValue}
                            }
                            [pscustomobject]@{
                                Name = $Name
                                Type = $Type
                                Data = $Value
                            }
                        }
                    ) # $Results - Registry
                    If ($Results) {
                        Write-Verbose "Confirm Install    - Automate Installed - Registry Keys Found"
                        $Global:Automate = New-Object -TypeName psobject
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name ComputerName -Value $ComputerWMI
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name ServerAddress -Value (($Results | Where-Object -Property Name -eq 'Server Address').Data)
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name ComputerID -Value (($Results | Where-Object -Property Name -eq 'ID').Data)
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name ClientID -Value (($Results | Where-Object -Property Name -eq 'ClientID').Data)
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name LocationID -Value (($Results | Where-Object -Property Name -eq 'LocationID').Data)
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name Version -Value (($Results | Where-Object -Property Name -eq 'Version').Data)
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name InstFolder -Value (Test-Path "$($env:windir)\ltsvc")
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name InstRegistry -Value $True
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name Installed -Value (Test-Path "$($env:windir)\ltsvc")
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name Service -Value ((Get-WmiObject -ComputerName $Computer -Class Win32_Service -Filter "Name='LTService'" -Credential $Credential -ErrorAction SilentlyContinue -ErrorVariable ProcessErrorWMIC).State)
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name LastHeartbeat -Value (((Get-Date) - [System.DateTime]($Results | Where-Object -Property Name -eq 'HeartbeatLastReceived').Data).TotalSeconds)
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name LastStatus -Value (((Get-Date) - [System.DateTime]($Results | Where-Object -Property Name -eq 'LastSuccessStatus').Data).TotalSeconds)
                        $Global:Automate | Add-Member -MemberType NoteProperty -Name Online -Value ($Global:Automate.InstFolder -and ($Global:Automate.Service -eq "Running"))
                        Write-Verbose $Global:Automate
                        If (($Global:Automate.ServerAddress -like "*$($Server)*") -and $Global:Automate.Online -and !$Force) {
                            If ($Show) {
                                $Global:Automate
                            } Else {
                                Write-Host "The Automate Agent is already installed and checked-in $($Global:Automate.LastStatus) seconds ago to $($Global:Automate.ServerAddress)" -ForegroundColor Green
                            }
                        } Else {
                            IF (!($Global:Automate.ServerAddress -like "*$($Server)*")) {
                                Write-Host "The Existing Automate Server Does Not Match The Target Automate Server." -ForegroundColor Red
                                Write-Host "Current Automate Server: $($Global:Automate.ServerAddress)" -ForegroundColor Red
                                }
                            Write-Verbose "Installing Automate..."
                            $WMIExitCode = Invoke-WmiMethod -class Win32_process -name Create -ArgumentList $WMIArg -ComputerName $Computer -Impersonation 3 -EnableAllPrivileges -Credential $Credential -ErrorAction SilentlyContinue
                            If ($WMIExitCode.ReturnValue -eq 0) {
                                Write-Host "Installing Automate Agent to https://$($Server) - WMI" -ForegroundColor Green
                                Write-Verbose "When pushing via WMI/RPC, the function will not wait and confirm the installation. "
                                $WMIDeployed = $True
                            } Else {
                                Write-Host "WMI Did NOT Execute Properly." -ForegroundColor Red
                                Write-Host "WMI Return Value: $($WMIExitCode.ReturnValue)" -ForegroundColor Red
                            }
                        }
                    } else {
                        Write-Verbose "Confirm Install    - Automate NOT Installed"
                        Write-Verbose "Installing Automate..."
                        $WMIExitCode = Invoke-WmiMethod -class Win32_process -name Create -ArgumentList $WMIArg -ComputerName $Computer -Impersonation 3 -EnableAllPrivileges -Credential $Credential -ErrorAction SilentlyContinue
                        If ($WMIExitCode.ReturnValue -eq 0) {
                            Write-Host "Installing Automate Agent to https://$($Server) - WMI" -ForegroundColor Green
                            Write-Verbose "When pushing via WMI, the function will not wait and confirm the installation. "
                            $WMIDeployed = $True
                        } Else {
                            Write-Host "WMI Did NOT Execute Properly." -ForegroundColor Red
                            Write-Host "WMI Return Value: $($WMIExitCode.ReturnValue)" -ForegroundColor Red
                        }
                    }
                } #End WMI Connectivity
                If (!$WinRMConectivity -and !$WMICConectivity) {Write-Host "Could Ping, but all protocols are inaccessible on $Computer. Deployment and Confirmations Failed" -ForegroundColor Yellow}
            }
        } Else {
            Write-Verbose "Ping Connectivity  - Failed"
            Write-Host "                      Ping Connectivity  - Failed" -ForegroundColor Yellow
        $PingTest = $False
        } # End Ping Test
    } # End Else Local Computer
        $Global:AutoChecks += New-Object psobject -Property @{
        Computer      = ($COMPUTER)
        ServerAddress = $Global:Automate.ServerAddress
        ComputerID    = $Global:Automate.ComputerID
        ClientID      = $Global:Automate.ClientID
        Version       = $Global:Automate.Version
        Online        = $Global:Automate.Online
        Ping          = $PingTest
        WinRM         = $WinRMConectivity
        WMI           = $WMICConectivity
        DeployedWinRm = $WinRMDeployed
        DeployedWMI   = $WMIDeployed
        } # End $Global:AutoChecks
        Clear-Variable Automate
        Clear-Variable Automate -Scope Global
        If (!$Silent) {Write-Host " "}
} # End Process

END
{
    Clear-Variable Username, Pass -ErrorAction SilentlyContinue | Out-Null
    $Global:AutoResults = ($Global:AutoChecks | Select-Object Computer, Online, ServerAddress, ComputerID, ClientID, Ping, WinRM, WMI, DeployedWinRM, DeployedWMI)
    Write-Verbose 'Results have been saved to $Global:AutoResults'
    $ErrorActionPreference = "Continue"
} # End END
} # End Function Push-Automate
########################
Set-Alias -Name LTP -Value Push-Automate -Description 'Push Automate Agent to Remote Computers'