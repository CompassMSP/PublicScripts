# WMI Service check and start/auto
Function serviceCheck($service) {
    Write-Verbose "Checking $service"
    Try {
        $svc_info = Get-WmiObject win32_service | Where-Object { $_.name -eq $service }
        if ($null -ne $svc_info.State) {
            @{'Status' = $svc_info.State; 'Start Mode' = $svc_info.StartMode; 'User' = $svc_info.StartName }
        } else {
            @{'Status' = 'Not Detected'; 'Start Mode' = ''; 'User' = '' }
        }
    } Catch {
        Write-Verbose $Error[0].exception.GetType().fullname
        @{'Status' = 'WMI Error'; 'Start Mode' = ''; 'User' = '' }
    }
}

# Check PS Version
Function Get-PSVersion {
    if (Test-Path variable:psversiontable) {
        $psversiontable.psversion
    } else {
        [version]'1.0.0.0'
    }
}

function extractHostname($url) {
    if ($url -eq '') {
        $false
    } elseif ($url -match '^http.+$') {
	    ([System.Uri]"$url").Authority
    } else {
        Write-Verbose 'Warning, server address does not supply http(s). Modify your agent template accordingly and run Update Config.'
        $url
    }
}

# Author: Joakim Borger Svendsen, 2017. 
# JSON info: http://www.json.org
# Svendsen Tech. MIT License. Copyright Joakim Borger Svendsen / Svendsen Tech. 2016-present.

# Take care of special characters in JSON (see json.org), such as newlines, backslashes
# carriage returns and tabs.
# '\\(?!["/bfnrt]|u[0-9a-f]{4})'
function EscapeJson {
    param(
        [String] $String)
    # removed: #-replace '/', '\/' `
    # This is returned 
    $String -replace '\\', '\\' -replace '\n', '\n' `
        -replace '\u0008', '\b' -replace '\u000C', '\f' -replace '\r', '\r' `
        -replace '\t', '\t' -replace '"', '\"'
}

# Meant to be used as the "end value". Adding coercion of strings that match numerical formats
# supported by JSON as an optional, non-default feature (could actually be useful and save a lot of
# calculated properties with casts before passing..).
# If it's a number (or the parameter -CoerceNumberStrings is passed and it 
# can be "coerced" into one), it'll be returned as a string containing the number.
# If it's not a number, it'll be surrounded by double quotes as is the JSON requirement.
function GetNumberOrString {
    param(
        $InputObject)
    if ($InputObject -is [System.Byte] -or $InputObject -is [System.Int32] -or `
        ($env:PROCESSOR_ARCHITECTURE -imatch '^(?:amd64|ia64)$' -and $InputObject -is [System.Int64]) -or `
            $InputObject -is [System.Decimal] -or `
        ($InputObject -is [System.Double] -and -not [System.Double]::IsNaN($InputObject) -and -not [System.Double]::IsInfinity($InputObject)) -or `
            $InputObject -is [System.Single] -or $InputObject -is [long] -or `
        ($Script:CoerceNumberStrings -and $InputObject -match $Script:NumberRegex)) {
        Write-Verbose -Message 'Got a number as end value.'
        "$InputObject"
    } else {
        Write-Verbose -Message "Got a string (or 'NaN') as end value."
        """$(EscapeJson -String $InputObject)"""
    }
}

function ConvertToJsonInternal {
    param(
        $InputObject, # no type for a reason
        [Int32] $WhiteSpacePad = 0)
    
    [String] $Json = ''
    
    $Keys = @()
    
    Write-Verbose -Message "WhiteSpacePad: $WhiteSpacePad."
    
    if ($null -eq $InputObject) {
        Write-Verbose -Message "Got 'null' in `$InputObject in inner function"
        $null
    }
    
    elseif ($InputObject -is [Bool] -and $InputObject -eq $true) {
        Write-Verbose -Message "Got 'true' in `$InputObject in inner function"
        $true
    }
    
    elseif ($InputObject -is [Bool] -and $InputObject -eq $false) {
        Write-Verbose -Message "Got 'false' in `$InputObject in inner function"
        $false
    }
    
    elseif ($InputObject -is [DateTime] -and $Script:DateTimeAsISO8601) {
        Write-Verbose -Message 'Got a DateTime and will format it as ISO 8601.'
        """$($InputObject.ToString('yyyy\-MM\-ddTHH\:mm\:ss'))"""
    }
    
    elseif ($InputObject -is [HashTable]) {
        $Keys = @($InputObject.Keys)
        Write-Verbose -Message "Input object is a hash table (keys: $($Keys -join ', '))."
    }
    
    elseif ($InputObject.GetType().FullName -eq 'System.Management.Automation.PSCustomObject') {
        $Keys = @(Get-Member -InputObject $InputObject -MemberType NoteProperty |
            Select-Object -ExpandProperty Name)

        Write-Verbose -Message "Input object is a custom PowerShell object (properties: $($Keys -join ', '))."
    }
    
    elseif ($InputObject.GetType().Name -match '\[\]|Array') {
        
        Write-Verbose -Message 'Input object appears to be of a collection/array type. Building JSON for array input object.'
        
        $Json += "[`n" + (($InputObject | ForEach-Object {
            
                    if ($null -eq $_) {
                        Write-Verbose -Message 'Got null inside array.'

                        ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + 'null'
                    }
            
                    elseif ($_ -is [Bool] -and $_ -eq $true) {
                        Write-Verbose -Message "Got 'true' inside array."

                        ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + 'true'
                    }
            
                    elseif ($_ -is [Bool] -and $_ -eq $false) {
                        Write-Verbose -Message "Got 'false' inside array."

                        ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + 'false'
                    }
            
                    elseif ($_ -is [DateTime] -and $Script:DateTimeAsISO8601) {
                        Write-Verbose -Message 'Got a DateTime and will format it as ISO 8601.'

                        ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + """$($_.ToString('yyyy\-MM\-ddTHH\:mm\:ss'))"""
                    }
            
                    elseif ($_ -is [HashTable] -or $_.GetType().FullName -eq 'System.Management.Automation.PSCustomObject' -or $_.GetType().Name -match '\[\]|Array') {
                        Write-Verbose -Message 'Found array, hash table or custom PowerShell object inside array.'

                        ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + (ConvertToJsonInternal -InputObject $_ -WhiteSpacePad ($WhiteSpacePad + 4)) -replace '\s*,\s*$'
                    }
            
                    else {
                        Write-Verbose -Message 'Got a number or string inside array.'

                        $TempJsonString = GetNumberOrString -InputObject $_
                        ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + $TempJsonString
                    }

                }) -join ",`n") + "`n$(' ' * (4 * ($WhiteSpacePad / 4)))],`n"

    } else {
        Write-Verbose -Message 'Input object is a single element (treated as string/number).'

        GetNumberOrString -InputObject $InputObject
    }
    if ($Keys.Count) {

        Write-Verbose -Message 'Building JSON for hash table or custom PowerShell object.'

        $Json += "{`n"

        foreach ($Key in $Keys) {

            # -is [PSCustomObject]) { # this was buggy with calculated properties, the value was thought to be PSCustomObject

            if ($null -eq $InputObject.$Key) {
                Write-Verbose -Message "Got null as `$InputObject.`$Key in inner hash or PS object."
                $Json += ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": null,`n"
            }

            elseif ($InputObject.$Key -is [Bool] -and $InputObject.$Key -eq $true) {
                Write-Verbose -Message "Got 'true' in `$InputObject.`$Key in inner hash or PS object."
                $Json += ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": true,`n"
            }

            elseif ($InputObject.$Key -is [Bool] -and $InputObject.$Key -eq $false) {
                Write-Verbose -Message "Got 'false' in `$InputObject.`$Key in inner hash or PS object."
                $Json += ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": false,`n"
            }

            elseif ($InputObject.$Key -is [DateTime] -and $Script:DateTimeAsISO8601) {
                Write-Verbose -Message 'Got a DateTime and will format it as ISO 8601.'
                $Json += ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": ""$($InputObject.$Key.ToString('yyyy\-MM\-ddTHH\:mm\:ss'))"",`n"
                
            }

            elseif ($InputObject.$Key -is [HashTable] -or $InputObject.$Key.GetType().FullName -eq 'System.Management.Automation.PSCustomObject') {
                Write-Verbose -Message "Input object's value for key '$Key' is a hash table or custom PowerShell object."
                $Json += ' ' * ($WhiteSpacePad + 4) + """$Key"":`n$(' ' * ($WhiteSpacePad + 4))"
                $Json += ConvertToJsonInternal -InputObject $InputObject.$Key -WhiteSpacePad ($WhiteSpacePad + 4)
            }

            elseif ($InputObject.$Key.GetType().Name -match '\[\]|Array') {

                Write-Verbose -Message "Input object's value for key '$Key' has a type that appears to be a collection/array."
                Write-Verbose -Message "Building JSON for ${Key}'s array value."

                $Json += ' ' * ($WhiteSpacePad + 4) + """$Key"":`n$(' ' * ((4 * ($WhiteSpacePad / 4)) + 4))[`n" + (($InputObject.$Key | ForEach-Object {

                            if ($null -eq $_) {
                                Write-Verbose -Message 'Got null inside array inside inside array.'
                                ' ' * ((4 * ($WhiteSpacePad / 4)) + 8) + 'null'
                            }

                            elseif ($_ -is [Bool] -and $_ -eq $true) {
                                Write-Verbose -Message "Got 'true' inside array inside inside array."
                                ' ' * ((4 * ($WhiteSpacePad / 4)) + 8) + 'true'
                            }

                            elseif ($_ -is [Bool] -and $_ -eq $false) {
                                Write-Verbose -Message "Got 'false' inside array inside inside array."
                                ' ' * ((4 * ($WhiteSpacePad / 4)) + 8) + 'false'
                            }

                            elseif ($_ -is [DateTime] -and $Script:DateTimeAsISO8601) {
                                Write-Verbose -Message 'Got a DateTime and will format it as ISO 8601.'
                                ' ' * ((4 * ($WhiteSpacePad / 4)) + 8) + """$($_.ToString('yyyy\-MM\-ddTHH\:mm\:ss'))"""
                            }

                            elseif ($_ -is [HashTable] -or $_.GetType().FullName -eq 'System.Management.Automation.PSCustomObject' `
                                    -or $_.GetType().Name -match '\[\]|Array') {
                                Write-Verbose -Message 'Found array, hash table or custom PowerShell object inside inside array.'
                                ' ' * ((4 * ($WhiteSpacePad / 4)) + 8) + (ConvertToJsonInternal -InputObject $_ -WhiteSpacePad ($WhiteSpacePad + 8)) -replace '\s*,\s*$'
                            }

                            else {
                                Write-Verbose -Message 'Got a string or number inside inside array.'
                                $TempJsonString = GetNumberOrString -InputObject $_
                                ' ' * ((4 * ($WhiteSpacePad / 4)) + 8) + $TempJsonString
                            }

                        }) -join ",`n") + "`n$(' ' * (4 * ($WhiteSpacePad / 4) + 4 ))],`n"

            } else {

                Write-Verbose -Message 'Got a string inside inside hashtable or PSObject.'
                # '\\(?!["/bfnrt]|u[0-9a-f]{4})'

                $TempJsonString = GetNumberOrString -InputObject $InputObject.$Key
                $Json += ' ' * ((4 * ($WhiteSpacePad / 4)) + 4) + """$Key"": $TempJsonString,`n"

            }

        }

        $Json = $Json -replace '\s*,$' # remove trailing comma that'll break syntax
        $Json += "`n" + ' ' * $WhiteSpacePad + "},`n"

    }

    $Json

}

function ConvertTo-STJson {
    [CmdletBinding()]
    #[OutputType([Void], [Bool], [String])]
    Param(
        [AllowNull()]
        [Parameter(Mandatory = $True,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True)]
        $InputObject,
        [Switch] $Compress,
        [Switch] $CoerceNumberStrings = $False,
        [Switch] $DateTimeAsISO8601 = $False)
    Begin {

        $JsonOutput = ''
        $Collection = @()
        # Not optimal, but the easiest now.
        [Bool] $Script:CoerceNumberStrings = $CoerceNumberStrings
        [Bool] $Script:DateTimeAsISO8601 = $DateTimeAsISO8601
        [String] $Script:NumberRegex = '^-?\d+(?:(?:\.\d+)?(?:e[+\-]?\d+)?)?$'
        #$Script:NumberAndValueRegex = '^-?\d+(?:(?:\.\d+)?(?:e[+\-]?\d+)?)?$|^(?:true|false|null)$'

    }

    Process {

        # Hacking on pipeline support ...
        if ($_) {
            Write-Verbose -Message "Adding object to `$Collection. Type of object: $($_.GetType().FullName)."
            $Collection += $_
        }

    }

    End {
        
        if ($Collection.Count) {
            Write-Verbose -Message "Collection count: $($Collection.Count), type of first object: $($Collection[0].GetType().FullName)."
            $JsonOutput = ConvertToJsonInternal -InputObject ($Collection | ForEach-Object { $_ })
        }
        
        else {
            $JsonOutput = ConvertToJsonInternal -InputObject $InputObject
        }
        
        if ($null -eq $JsonOutput) {
            Write-Verbose -Message "Returning `$null."
            return $null # becomes an empty string :/
        }
        
        elseif ($JsonOutput -is [Bool] -and $JsonOutput -eq $true) {
            Write-Verbose -Message "Returning `$true."
            [Bool] $true # doesn't preserve bool type :/ but works for comparisons against $true
        }
        
        elseif ($JsonOutput -is [Bool] -and $JsonOutput -eq $false) {
            Write-Verbose -Message "Returning `$false."
            [Bool] $false # doesn't preserve bool type :/ but works for comparisons against $false
        }
        
        elseif ($Compress) {
            Write-Verbose -Message 'Compress specified.'
            (
                ($JsonOutput -split '\n' | Where-Object { $_ -match '\S' }) -join "`n" `
                    -replace '^\s*|\s*,\s*$' -replace '\ *\]\ *$', ']'
            ) -replace ( # these next lines compress ...
                '(?m)^\s*("(?:\\"|[^"])+"): ((?:"(?:\\"|[^"])+")|(?:null|true|false|(?:' + `
                    $Script:NumberRegex.Trim('^$') + `
                    ')))\s*(?<Comma>,)?\s*$'), "`${1}:`${2}`${Comma}`n" `
                -replace '(?m)^\s*|\s*\z|[\r\n]+'
        }
        
        else {
            ($JsonOutput -split '\n' | Where-Object { $_ -match '\S' }) -join "`n" `
                -replace '^\s*|\s*,\s*$' -replace '\ *\]\ *$', ']'
        }
    
    }

}

Function Test-CommandExists {
    Param ($command)
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'stop'
    try {
        if (Get-Command $command ) {
            $true
        }
    } Catch {
        $false
    } Finally {
        $ErrorActionPreference = $oldPreference
    }
}

Function Test-JanusLoaded {
    try { 
        $janus = Get-Content $env:windir\ltsvc\lterrors.txt | Select-String 'Janus' | Select-Object -Last 1
        if ($janus -match 'Janus enabled') {
            $true
        } else {
            $false
        }
    } catch {
        $false
    }
}

Function Test-FailedSignup {
    try { 
        $signup = Get-Content $env:windir\ltsvc\lterrors.txt | Select-String 'Failed Signup' | Select-Object -Last 1
        if ($signup -match 'Failed Signup') {
            $true
        } else {
            $false
        }
    } catch {
        $false
    }
}
Function Test-CryptoFailed {
    try {
        $signup = Get-Content $env:windir\ltsvc\lterrors.txt | Select-String 'Unable to initialize remote agent security.' | Select-Object -Last 1
        if ($signup -match 'Unable to initialize remote agent security.') {
            $true
        } else {
            $false
        }
    } catch {
        $false
    }
}

Function Invoke-CheckIn {
    $servicecmd = (Join-Path $env:windir '\system32\sc.exe')
    # Force check-in
    Try {
        & $servicecmd control ltservice 136 | Out-Null
    } catch {
        Write-Verbose 'Error sending checkin'
    }
}

Function Start-AutomateDiagnostics {
    Param(
        $ltposh = 'http://bit.ly/LTPoSh',
        $automate_server = '',
        [switch]$verbose,
        [switch]$include_lterrors,
        [switch]$skip_updates,
        [switch]$use_sockets
    )
    # Force $skip_updates to be $true
    $skip_updates = $true
    if ($verbose) {
        $VerbosePreference = 'Continue'
    }

    $signup_failure = $false
    $agent_crypto_failure = $false
    $janus_res = $false
    # Initial checkin
    Invoke-CheckIn

    # Get powershell version
    $psver = Get-PSVersion
    
    # 2023.01.16 -- Joe McCall
    # Simplified the checks for the $ltposh_loaded variable

    Write-Verbose 'Loading LTPosh'
    Try { 
		(New-Object Net.WebClient).DownloadString($ltposh) | Invoke-Expression
        $ltposh_loaded = [bool](Get-Command -ListImported -Name Get-LTServiceInfo)
    } Catch {
        Write-Verbose $Error[0].exception.GetType().fullname
        Write-Verbose $_.Exception.Message
        $ltposh_loaded = $false
    }
    If ($ltposh_loaded -eq $false -and $ltposh -ne 'http://bit.ly/LTPoSh') {
        Write-Output 'LTPosh failed to load, failing back to bit.ly link'
        $ltposh = 'http://bit.ly/LTPoSh'
        Try {
			(New-Object Net.WebClient).DownloadString($ltposh) | Invoke-Expression
            $ltposh_loaded = [bool](Get-Command -ListImported -Name Get-LTServiceInfo)
        } Catch {
            Write-Verbose $Error[0].exception.GetType().fullname
            Write-Verbose $_.Exception.Message
            $ltposh_loaded = $false
        }
    }

    # Check services
    $ltservice_check = serviceCheck('LTService')
    $ltsvcmon_check = serviceCheck('LTSVCMon')

    # Check LTSVC path and lterrors.txt
    $ltsvc_path_exists = Test-Path -Path (Join-Path $env:windir '\ltsvc')
    $lterrors_exists = Test-Path -Path (Join-Path $env:windir '\ltsvc\lterrors.txt')
    
    # 2023.01.16 -- Joe McCall
    # Strip the leading LTService version header from each line, and ignore the Disk Clean access denied errors to improve readability of the log output in the browser
    if ($include_lterrors) {
        $lterrors = if ($lterrors_exists) {
            (Get-Content -Path (Join-Path $env:windir '\ltsvc\lterrors.txt') | Select-String 'Disk Clean Dir Error: Access to the path' -NotMatch) -replace 'LTService\s*v\d+\.\d+\s*-\s*', '' -join "`n"
        } else {
            ''
        }
        $lterrors_enc = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($lterrors))
    } else {
        $lterrors_enc = ''
    }

    # Get reg keys in case LTPosh fails
    $locationid = Try {
 (Get-ItemProperty -Path hklm:\software\labtech\service -ErrorAction Stop).locationid
    } Catch {
        $null
    }
    $clientid = Try {
 (Get-ItemProperty -Path hklm:\software\labtech\service -ErrorAction Stop).clientid
    } Catch {
        $null
    }
    $id = Try {
 (Get-ItemProperty -Path hklm:\software\labtech\service -ErrorAction Stop).id
    } Catch {
        $null
    }
    $version = Try {
 (Get-ItemProperty -Path hklm:\software\labtech\service -ErrorAction Stop).version
    } Catch {
        $null
    }
    $server = Try {
 ((Get-ItemProperty -Path hklm:\software\labtech\service -ErrorAction Stop).'Server Address')
    } Catch {
        $null
    }
    if ($ltsvc_path_exists -and $lterrors_exists) {
        $janus_res = Test-JanusLoaded
        $signup_failure = Test-FailedSignup
        $agent_crypto_failure = Test-CryptoFailed
    }

    if ($ltposh_loaded) {
        # Get ltservice info
        Try {
            $info = Get-LTServiceInfo

            # Get checkin / heartbeat times to DateTime
            $lasthbsent = Get-Date $info.HeartbeatLastSent
            # This could throw exception because it is only returned if there has been a SuccessStatus
            if ($info.LastSuccessStatus) {
                $lastsuccess = Get-Date $info.LastSuccessStatus
            }
            $lasthbrcv = Get-Date $info.HeartbeatLastReceived

            # Check online and heartbeat statuses
            $online_threshold = (Get-Date).AddMinutes(-5)
            $heartbeat_threshold = (Get-Date).AddMinutes(-5)
            
            # Split server list
            $servers = ($info.'Server Address').Split('|')

            $online = $lastsuccess -ge $online_threshold
            $heartbeat_rcv = $lasthbrcv -ge $heartbeat_threshold 
            $heartbeat_snd = $lasthbsent -ge $heartbeat_threshold
            $heartbeat = $heartbeat_rcv -or $heartbeat_snd
            $heartbeat_status = $info.HeartbeatLastSent
            
            # Check for persistent TCP connection
            Try {
                if ($psver -ge [version]'3.0.0.0' -and $heartbeat -eq $false) {
                    # Check network sockets for established connection from ltsvc to server
                    Write-Verbose 'Heartbeat failed, checking for TCP tunnel'
                    $socket = Get-Process -processname 'ltsvc' | ForEach-Object { $process = $_.ID; Get-NetTCPConnection | Where-Object { $_.State -eq 'Established' -and $_.RemotePort -eq 443 -and $_.OwningProcess -eq $process } }
                    if ($socket.State -eq 'Established') {
                        $heartbeat = $true
                        $heartbeat_status = 'Socket Established'
                    }
                }
            } Catch {
            }

            # If services are stopped, use Restart-LTService to get them working again
            If ($ltservice_check.Status -eq 'Stopped' -or $ltsvcmon_check -eq 'Stopped' -or !($heartbeat) -or !($online)) {
                Write-Verbose 'Issuing Restart-LTService and sending checkin'
                Restart-LTService
                Invoke-CheckIn
                Start-Sleep -Seconds 10
                $info = Get-LTServiceInfo
                $ltservice_check = serviceCheck('LTService')
                $ltsvcmon_check = serviceCheck('LTSVCMon')
                # Get checkin / heartbeat times to DateTime
                # This could throw exception because it is only returned if there has been a SuccessStatus
                if ($info.LastSuccessStatus) {
                    $lastsuccess = Get-Date $info.LastSuccessStatus
                }
                $lasthbsent = Get-Date $info.HeartbeatLastSent
                $lasthbrcv = Get-Date $info.HeartbeatLastReceived
                $online = $lastsuccess -ge $online_threshold
                $heartbeat_rcv = $lasthbrcv -ge $heartbeat_threshold 
                $heartbeat_snd = $lasthbsent -ge $heartbeat_threshold
                $heartbeat = $heartbeat_rcv -or $heartbeat_snd
                $heartbeat_status = $info.HeartbeatLastSent
            }

            # Get server list
            $server_test = $false
            foreach ($server in $servers) {
                Write-Verbose "Server: $server"
                $hostname = extractHostname($server)
                Write-Verbose "Hostname: $hostname"
                if (!($hostname) -or $hostname -eq '' -or $null -eq $hostname) {
                    Write-Verbose 'Error with hostname'
                    continue
                } else {
                    $compare_test = if (($hostname -eq $automate_server -and $automate_server -ne '') -or $automate_server -eq '') {
                        $true
                    } else {
                        $false
                    }
                    if (Test-CommandExists -Command 'Test-NetConnection' -and $use_sockets) {
                        Try {
                            $conn_test = Test-NetConnection -ComputerName $hostname -Port 443 -ErrorAction Stop
                        } Catch {
                            $timeout = 1000
                            $netping = New-Object System.Net.NetworkInformation.Ping
                            $ping = $netping.Send($hostname, $timeout)
                            if ($ping.Status -eq 'Success') {
                                Write-Verbose 'Fallback to .net ping succeeded'
                                $conn_test = $true
                            } else {
                                Write-Verbose 'Port test failed'
                                $conn_test = $false
                            }
                        }
                    } else {
                        Try {
                            $conn_test = Test-Connection -ComputerName $hostname -Count 1 -ErrorAction Stop
                        } Catch {
                            $timeout = 1000
                            $netping = New-Object System.Net.NetworkInformation.Ping
                            $ping = $netping.Send($hostname, $timeout)
                            if ($ping.Status -eq 'Success') {
                                Write-Verbose 'Fallback to .net ping succeeded'
                                $conn_test = $true
                            } else {
                                Write-Verbose 'Ping test failed'
                                $conn_test = $false 
                            }
                        }
                    }

                    Try { 
                        $ver_test = (New-Object Net.WebClient).DownloadString("$($server)/labtech/agent.aspx")
                        $target_version = $ver_test.Split('|')[6]
                    } Catch {
                        Write-Verbose 'Unable to obtain target version'
                        $target_version = ''
                    }
                    if ($conn_test -and $target_version -ne '' -and $compare_test) {
                        $server_test = $true
                        $server_msg = "$server passed all checks"
                        break
                    }
                }
            }
            if ($server_test -eq $false -and $servers.Count -eq 0) {
                $server_msg = 'No automate servers detected'
            } elseif ($server_test -eq $false -and $servers.Count -gt 0) {
                $server_msg = 'Error'
                if (!($compare_test)) {
                    $server_msg = $server_msg + " | Server address not matched ($hostname)"
                }
                if (!($conn_test)) {
                    $server_msg = $server_msg + " | Ping failure ($hostname)"
                }
                if ($target_version -eq '') {
                    $server_msg = $server_msg + " | Version check fail ($hostname)"
                }
            }

            # Check updates
            $current_version = $info.Version
            if ($target_version -eq $info.Version) {
                $update_text = 'Version {0} - Latest' -f $info.Version
            } elseif (!$skip_updates) {
                Write-Verbose 'Starting update'
                taskkill /im ltsvc.exe /f
                taskkill /im ltsvcmon.exe /f
                taskkill /im lttray.exe /f
                Try {
                    Update-LTService -WarningVariable updatewarn
                    Start-Sleep -Seconds 30
                    Try {
                        Restart-LTService -Confirm:$false
                    } Catch {
                    }
                    Invoke-CheckIn
                    Start-Sleep -Seconds 30
                    $info = Get-LTServiceInfo
                    $janus_res = Test-JanusLoaded
                    if ([version]$info.Version -gt [version]$current_version) {
                        $update_text = 'Updated from {1} to {0}' -f $info.Version, $current_version
                    } else {
                        $update_text = 'Error updating, still on {0}' -f $info.Version
                    }
                } Catch {
                    $update_text = 'Error: Update-LTService failed to run'
                }

            } else {
                $update_text = 'Updates needed, available version {0}' -f $target_version
            }
            # Collect diagnostic data into hashtable
            $diag = @{
                'id'                   = $info.id
                'version'              = $info.Version
                'server_addr'          = $server_msg
                'server_match'         = $compare_test
                'online'               = $online
                'heartbeat'            = $heartbeat
                'update'               = $update_text
                'lastcontact'          = $info.LastSuccessStatus
                'heartbeat_sent'       = $heartbeat_status
                'heartbeat_rcv'        = $info.HeartbeatLastReceived
                'svc_ltservice'        = $ltservice_check
                'svc_ltsvcmon'         = $ltsvcmon_check
                'ltposh_loaded'        = $ltposh_loaded
                'clientid'             = $info.ClientID
                'locationid'           = $info.LocationID
                'ltsvc_path_exists'    = $ltsvc_path_exists
                'janus_status'         = $janus_res
                'signup_failure'       = $signup_failure
                'agent_crypto_failure' = $agent_crypto_failure
                'lterrors'             = $lterrors_enc
            }
        } Catch {
            # LTPosh loaded, issue with agent
            $exception = $_.Exception.Message
            $repair = if (-not ($ltsvc_path_exists) -or $ltsvcmon_check.Status -eq 'Not Detected' -or $ltservice_check.Status -eq 'Not Detected' -or $null -eq $id -or $janus_res -eq $false) {
                'Reinstall'
            } else {
                'Restart'
            }
            if ($null -eq $version -or $ltsvc_path_exists -eq $false) {
                $version = 'Agent error'
            }
            if ($janus_res -eq $false) {
                $version = 'Janus failure'
            }
            $diag = @{
                'id'                   = $id
                'svc_ltservice'        = $ltservice_check
                'svc_ltsvcmon'         = $ltsvcmon_check
                'ltposh_loaded'        = $ltposh_loaded
                'server_addr'          = $server
                'version'              = $version
                'ltsvc_path_exists'    = $ltsvc_path_exists
                'locationid'           = $locationid
                'clientid'             = $clientid
                'repair'               = $repair
                'janus_status'         = $janus_res
                'signup_failure'       = $signup_failure
                'agent_crypto_failure' = $agent_crypto_failure
                'lterrors'             = $lterrors_enc
                'exception'            = $exception
            }
        }
    } else {
        # LTPosh Failure, show basic settings
        $diag = @{
            'id'                   = $id
            'ltsvc_path_exists'    = $ltsvc_path_exists
            'server_addr'          = $server
            'locationid'           = $locationid
            'clientid'             = $clientid
            'svc_ltservice'        = $ltservice_check
            'svc_ltsvcmon'         = $ltsvcmon_check
            'ltposh_loaded'        = $ltposh_loaded
            'version'              = $version
            'janus_status'         = $janus_res
            'signup_failure'       = $signup_failure
            'agent_crypto_failure' = $agent_crypto_failure
            'lterrors'             = $lterrors_enc
        }
    }

    # Output diagnostic data in JSON format - ps2.0 compatible
    if ($psver -ge [version]'3.0.0.0') {
        $output = $diag | ConvertTo-Json -Depth 2
    } else {
        $output = $diag | ConvertTo-STJson
    }
    Write-Host '!---BEGIN JSON---!'
    Write-Host $output
}
