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
        $return = @()
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