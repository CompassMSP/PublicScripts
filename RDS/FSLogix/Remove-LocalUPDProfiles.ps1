<#

This script deletes local profile data on an RDS server that has UPDs or FSL Profiles. When UPDs are enabled user files will sometimes be left behind. This can cause some unexpected results with applications.

The script has a failsafe that does not run if UPDs or FSL Profiles become disabled at any point.

Andy Morales
#>
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
Function Get-RDSActiveSessions {
    <#
    .SYNOPSIS
        Returns open sessions of a local workstation
    .DESCRIPTION
        Get-ActiveSessions uses the command line tool qwinsta to retrieve all open user sessions on a computer regardless of how they are connected.
    .OUTPUTS
        A custom object with the following members:
            UserName: [string]
            SessionName: [string]
            ID: [string]
            Type: [string]
            State: [string]
    .NOTES
        Author: Anthony Howell
    .LINK
        qwinsta
        http://stackoverflow.com/questions/22155943/qwinsta-error-5-access-is-denied
        https://theposhwolf.com
    #>
    Begin {
        $Name = $env:COMPUTERNAME
        $ActiveUsers = @()
    }
    Process {
        $result = qwinsta /server:$Name
        If ($result) {
            ForEach ($line in $result[1..$result.count]) {
                #avoiding the line 0, don't want the headers
                $tmp = $line.split(" ") | Where-Object { $_.length -gt 0 }
                If (($line[19] -ne " ")) {
                    #username starts at char 19
                    If ($line[48] -eq "A") {
                        #means the session is active ("A" for active)
                        $ActiveUsers += New-Object PSObject -Property @{
                            "ComputerName" = $Name
                            "SessionName"  = $tmp[0]
                            "UserName"     = $tmp[1]
                            "ID"           = $tmp[2]
                            "State"        = $tmp[3]
                            "Type"         = $tmp[4]
                        }
                    }
                    Else {
                        $ActiveUsers += New-Object PSObject -Property @{
                            "ComputerName" = $Name
                            "SessionName"  = $null
                            "UserName"     = $tmp[0]
                            "ID"           = $tmp[1]
                            "State"        = $tmp[2]
                            "Type"         = $null
                        }
                    }
                }
            }
        }
        Else {
            Write-Error "Unknown error, cannot retrieve logged on users"
        }
    }
    End {
        Return $ActiveUsers
    }
}
function Get-RecursiveLocalGroupMembers {
    <#
    Modified version of this:
        https://gist.githubusercontent.com/mr64bit/874b773795f25673145973a84afd7741/raw/8e5ffe804182ba3ab97ab6cc48da4708dcc76fd7/Get-RecursiveGroupMembers.ps1
    #>
    [CmdletBinding()]
    param(
        [String]$Group
    )

    #Private function for recursion
    function Get-Members {
        param(
            [String]$Domain,
            [String]$Group,
            [String[]]$Path
        )

        $Members = Get-WmiObject Win32_GroupUser -Filter "GroupComponent=`"Win32_Group.Domain='$Domain',Name='$Group'`"" | ForEach-Object { $_.PartComponent }

        Write-Verbose "Filter: GroupComponent=`"Win32_Group.Domain='$Domain',Name='$Group'`""
        Foreach ($Member in $Members) {
            Write-Verbose $Member
            #Parsing the string is faster than resolving to a WMI object
            $Regex = [Regex]::Match($Member, '(?i)cimv2:(.+?)\.Domain="(.+?)",Name="(.+?)"')
            $Class = $Regex.Groups[1].Value
            $Domain = $Regex.Groups[2].Value
            $Name = $Regex.Groups[3].Value

            If ($Script:Done -notContains $Name) {
                $Script:Done.add($Name)
                If ($Class -like "Win32_Group") {
                    #Don't loop into the current group
                    If ($Name -notlike $Group) {
                        Write-Verbose "Calling get-members with group $Name"
                        Get-Members -Domain $Domain -Group $Name -Path ($Path + "$Domain\$Name")
                    }
                }
                ElseIf ($Name) {
                    New-Object PSObject -Property @{
                        User       = "$Domain\$Name";
                        MemberPath = ($Path -join ("->"))
                    }
                }
            }
        }
    }

    #create array to store user/group objects
    $Script:Done = New-Object System.Collections.Generic.List[System.Object]

    #Get members of a group
    Get-Members -Domain $env:COMPUTERNAME -Group $Group -Path @($Group)
}

#Array that will store the command and all parameters
$CommandToExecute = @()

#The basic command
$CommandToExecute += 'C:\BIN\DelProf2\DelProf2.exe /u'

#region find members of the FSL exclude groups
$ExcludeMembers = @()
$ExcludeMembers += Get-RecursiveLocalGroupMembers -Group 'FSLogix ODFC Exclude List'
$ExcludeMembers += Get-RecursiveLocalGroupMembers -Group 'FSLogix Profile Exclude List'

$UsersToExclude = @(
    'TWBAIT'
)

if ($ExcludeMembers.count -gt 0) {
    Foreach ($user in $ExcludeMembers.user) {
        $UsersToExclude += ($user -split '\\')[-1]
    }
}
#endregion

#Add current sessions to the exclude variable
$UsersToExclude += (Get-RDSActiveSessions).username

#UPDs
if (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Terminal Server\ClusterSettings' -Name UvhdEnabled -Value 1) {
    #UvhdCleanupBin is also excluded since it is unknown what deleting it will do
    $CommandToExecute += '/ed:UvhdCleanupBin'

    #Exclude logged in users
    $CommandToExecute += "/ed:$($UsersToExclude -join ' /ed:')"
}

#FSL Profiles
if (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\FSLogix\Profiles' -Name 'Enabled' -ValueData 1) {
    #Exclude Currently logged in users
    $CommandToExecute += "/ed:$($UsersToExclude -join ' /ed:')"

    #Exclude local_ folder of logged in users
    $CommandToExecute += "/ed:local_$($UsersToExclude -join ' /ed:Local_')"
}

#merge commands and run
Invoke-Expression -Command ($CommandToExecute -join ' ')