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