#POC ONLY! still needs work
#Andy Morales

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

    if (Test-Path -Path $OutPath){
        Remove-Item $OutPath -Recurse -Force
    }

    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFile, $OutPath)
}

mkdir C:\windows\temp\autoruns -Force

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

(New-Object System.Net.WebClient).DownloadFile('https://download.sysinternals.com/files/Autoruns.zip', 'C:\windows\temp\autoruns.zip')

Expand-ZIP -ZipFile 'C:\windows\temp\autoruns.zip' -OutPath 'C:\windows\temp\autoruns'

$AutorunsArgs = @(
        "C:\Windows\Temp\autoruns\autorunsc.exe",
        "-a * -nobanner -c -accepteula -vt -o 'C:\Windows\Temp\autoruns\output.csv'"
)

Invoke-Expression ($AutorunsArgs -Join ' ')

$result = Import-Csv -Path "C:\Windows\Temp\autoruns\output.csv"

$VtThresholdPercent = .001
$FoundThreats = @()

Foreach ($item in $result){
    #Make sure result is not empty or unkown
    if((![string]::IsNullOrEmpty($item.'VT detection')) -and ($item.'VT detection' -ne 'Unknown')){
        $VTResult = $item.'VT detection'.Split('|')

        #add to aray if VT Result is over threshold
        if (($VTResult[0] / $VTResult[1]) -gt $VtThresholdPercent){
            $FoundThreats += $item
        }
    }
}

#Remove-item -path "C:\Windows\Temp\autoruns\output.csv" -force

$FoundThreats