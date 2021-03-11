<#
Downloads autoruns and checks for suspicious files. The threshold percentage can be adjusted with $VtThresholdPercent

Andy Morales
#>

[CmdletBinding()]
param (
    [parameter(Mandatory = $false)]
    [String]$VtThresholdPercent = .01
)

$AutoRunsFolder = 'C:\windows\temp\autoruns'

function Expand-ZIP {
    <#
    Extracts a ZIP file to a directory. The contents of the destination will be deleted if they already exist.

    Andy Morales
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [String]$ZipFile,

        [parameter(Mandatory = $true)]
        [String]$OutPath
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -Path $OutPath) {
        Remove-Item $OutPath -Recurse -Force
    }

    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFile, $OutPath)
}
function New-SecureFolder {
    <#
    #This script creates a folder that only administrators and system have access to.
    #It is a best practice to wipe the folder once you are done running the script that relies on it.

    Andy Morales
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
            HelpMessage = 'C:\temp')]
        [String]$Path
    )

    #Delete the folder if it already exists
    if (Test-Path -Path $Path) {
        try {
            Remove-Item -Path $Path -Force -Recurse -ErrorAction Stop
        }
        catch {
            Write-Output "Could not clear the contents of $($Path). Script will exit."
            EXIT
        }
    }

    #Create the folder
    New-Item -Path $Path -ItemType Directory -Force | Out-Null

    #Remove all explicit permissions
    ICACLS ("$Path") /reset | Out-Null

    #Add SYSTEM permission
    ICACLS ("$Path") /grant ("SYSTEM" + ':(OI)(CI)F') | Out-Null

    #Give Administrators Full Control
    ICACLS ("$Path") /grant ("Administrators" + ':(OI)(CI)F') | Out-Null

    #Disable Inheritance on the Folder. This is done last to avoid permission errors.
    ICACLS ("$Path") /inheritance:r | Out-Null
}

#region pre-reqs
New-SecureFolder -Path $AutoRunsFolder

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$AutorunsZip = "$($AutoRunsFolder)\autoruns.zip"
$AutoRunsBin = "$($AutoRunsFolder)\BIN"

(New-Object System.Net.WebClient).DownloadFile('https://download.sysinternals.com/files/Autoruns.zip', "$AutorunsZip")

Expand-ZIP -ZipFile "$($AutoRunsFolder)\autoruns.zip" -OutPath $AutoRunsBin
#endregion pre-reqs

#region runApp
$AutorunsArgs = @(
    "$($AutoRunsBin)\autorunsc.exe",
    "-a * -nobanner -c -accepteula -vt -o $($AutoRunsBin)\output.csv"
)

Invoke-Expression ($AutorunsArgs -Join ' ')
#endregion runApp

#region ReviewResults
$result = Import-Csv -Path "$($AutoRunsBin)\output.csv"

$FoundThreats = @()

Foreach ($item in $result) {
    #Make sure result is not empty or unknown
    if ((![string]::IsNullOrEmpty($item.'VT detection')) -and ($item.'VT detection' -ne 'Unknown')) {
        $VTResult = $item.'VT detection'.Split('|')

        #add to array if VT Result is over threshold
        if (($VTResult[0] / $VTResult[1]) -gt $VtThresholdPercent) {
            $FoundThreats += $item
        }
    }
}
#endregion ReviewResults

#Return detected items if any exist
if ($FoundThreats.Count -gt 0) {
    $Result = 'DETECTED:'
    $Result += $FoundThreats | Select-Object 'Image Path', 'VT detection' | Out-String

    $Result | Out-File -FilePath "$($AutoRunsBin)\result.txt"

    RETURN $Result
}
else {
    RETURN 'Nothing Found'
}