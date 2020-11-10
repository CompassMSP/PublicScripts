#Requires -Module ActiveDirectory -RunAsAdministrator

$LogPath = 'C:\Windows\Temp\krbScript.txt'

function Write-Log {
    <#
    .Synopsis
    Write-Log writes a message to a specified log file with the current time stamp.
    .DESCRIPTION
    The Write-Log function is designed to add logging capability to other scripts.
    In addition to writing output and/or verbose you can write to a log file for
    later debugging.
    .NOTES
    Created by: Jason Wasser @wasserja
    Modified: 11/24/2015 09:30:19 AM

    Changelog:
        * Code simplification and clarification - thanks to @juneb_get_help
        * Added documentation.
        * Renamed LogPath parameter to Path to keep it standard - thanks to @JeffHicks
        * Revised the Force switch to work as it should - thanks to @JeffHicks

    To Do:
        * Add error handling if trying to create a log file in a inaccessible location.
        * Add ability to write $Message to $Verbose or $Error pipelines to eliminate
        duplicates.
    .PARAMETER Message
    Message is the content that you wish to add to the log file.
    .PARAMETER Path
    The path to the log file to which you would like to write. By default the function will
    create the path and file if it does not exist.
    .PARAMETER Level
    Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational)
    .PARAMETER NoClobber
    Use NoClobber if you do not wish to overwrite an existing file.
    .EXAMPLE
    Write-Log -Message 'Log message'
    Writes the message to c:\Logs\PowerShellLog.log.
    .EXAMPLE
    Write-Log -Message 'Restarting Server.' -Path c:\Logs\ScriptOutput.log
    Writes the content to the specified log file and creates the path and file specified.
    .EXAMPLE
    Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error
    Writes the message to the specified log file as an error message, and writes the message to the error pipeline.
    .LINK
    https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [Alias('LogPath')]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warn", "Info")]
        [string]$Level = "Info",

        [Parameter(Mandatory = $false)]
        [switch]$NoClobber,

        [Parameter(Mandatory = $false)]
        [switch]$DailyMode
    )

    Begin {
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'Continue'
        if ($DailyMode) {
            $Path = $Path.Replace('.', "-$(Get-Date -UFormat "%Y%m%d").")
        }
    }
    Process {
        # If the file already exists and NoClobber was specified, do not write to the log.
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
        }

        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            New-Item $Path -Force -ItemType File
        }

        else {
            # Nothing to see here yet.
        }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Error $Message
                $LevelText = 'ERROR:'
            }
            'Warn' {
                Write-Warning $Message
                $LevelText = 'WARNING:'
            }
            'Info' {
                Write-Verbose $Message
                $LevelText = 'INFO:'
            }
        }

        # Write log entry to $Path
        #try to write to the log file. Retry if it is locked
        $StopWriteLogLoop = $false
        [int]$WriteLogRetryCount = "0"
        do {
            try {
                "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append -ErrorAction Stop
                $StopWriteLogLoop = $true
            }
            catch {
                if ($WriteLogRetryCount -gt 5) {
                    $StopWriteLogLoop = $true
                }
                else {
                    Start-Sleep -Milliseconds 500
                    $WriteLogRetryCount++
                }
            }
        }While ($StopWriteLogLoop -eq $false)
    }
    End {
    }
}

function Get-RandomCharacters {
    #https://activedirectoryfaq.com/2017/08/creating-individual-random-passwords/
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $false)]
        [String]$Length = '12',

        [parameter(Mandatory = $false)]
        [switch]$AsSecureString
    )

    $RandomPassword = ''

    $LowerCaseChars = 'abcdefghkmnoprstuvwxyz'
    $UpperCaseChars = 'ABCDEFGHKMNPRSTUVWXYZ'
    $NumberChars = '23456789'
    $SpecialChars = '@#$%-+*_=?:<>^&'
    $AllCharacters = $LowerCaseChars + $UpperCaseChars + $NumberChars + $SpecialChars

    #generate random strings until they contain one of each character type. Put a limit on the loop so it only runs 10 times
    DO {
        $bytes = New-Object "System.Byte[]" $Length
        $rnd = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
        $rnd.GetBytes($bytes)
        $RandomCharacters = ""

        for ( $i = 0; $i -lt $Length; $i++ ) {
            $RandomCharacters += $AllCharacters[ $bytes[$i] % $AllCharacters.Length ]
        }

        $count++
    }Until (($RandomCharacters -cmatch '(?=.*[a-z])(?=.*[A-Z])(?=.*[@#$%-+*_=?:<>^&])(?=.*\d)') -or ($count -ge 10))

    if ($AsSecureString) {
        [securestring]$RandomCharacters = ConvertTo-SecureString $RandomCharacters -AsPlainText -Force
    }

    return $RandomCharacters
}

Function Test-AllDcConnection {
    #Checks to see if all DCs can be reached

    #Find all DCs
    $AllDomainControllers = (Get-ADDomainController -Filter *).name

    try{
        Test-Connection -ComputerName $AllDomainControllers -Count 2 -ErrorAction Stop
        $AllConnectionsSucceeded = $true
    }
    catch{
        $AllConnectionsSucceeded = $false
    }

    Return $AllConnectionsSucceeded
}

#Check to see if the computer is the PDC
if ((Get-WmiObject Win32_ComputerSystem).domainRole -eq 5) {
    try{
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch{
        Write-Log -Path $LogPath -Level Error -Message "Could not load ActiveDirectory Module"
        break
    }

    #Search for the account using its SID
    $krbtgtAccount = Get-ADUser -Identity "$((Get-ADDomain).domainSID.Value)-502" -Properties passwordLastSet

    #Check to see if the password has been changed in the past 180 days
    if ($krbtgtAccount.PasswordLastSet -lt (Get-Date).AddDays(-180)) {

        #Check the forest/domain functional level
        if ((Get-ADForest).ForestMode.ToString().Substring(7, 4) -ge '2008' -and (Get-ADDomain).DomainMode.ToString().Substring(7, 4) -ge '2008') {

            #Find all domain controllers
            if (Test-AllDcConnection){
                #Change the krbtgt password
                Set-ADAccountPassword -Identity $krbtgtAccount -NewPassword (Get-RandomCharacters -Length 64 -AsSecureString) -Reset

                #Force AD Replication
                Foreach ($DC in $AllDomainControllers) {
                    repadmin /syncall $DC (Get-ADDomain).DistinguishedName /e /A | Out-Null
                }
            }
            else{
                Write-Log -Path $LogPath -Level Error -Message "Could not connect to all domain controllers. Exiting the script to avoid replication issues."
            }
        }
        else {
            Write-Log -Path $LogPath -Level Error -Message 'The domain/forest functional level is too low. Raise to at least 2008 to continue.'
        }
    }
    else{
        Write-Log -Path $LogPath -Level Information -Message "The password was changed less than 180 days ago. No need to update it."
    }
}
else {
    Write-Log -Path $LogPath -Level Info -Message 'Computer is not a DC. Script will exit.'
}
