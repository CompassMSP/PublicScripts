function Get-GitFilesFromRepo {
    #Import All GPOs from GitHub
    #https://gist.github.com/chrisbrownie/f20cb4508975fb7fb5da145d3d38024a
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
            HelpMessage = 'Username')]
        [String]$Owner,
        [parameter(Mandatory = $true)]
        [String]$Repository,
        [parameter(Mandatory = $true,
            HelpMessage = 'Scripts\Example')]
        [String]$Path,
        [parameter(Mandatory = $true,
            HelpMessage = 'C:\Temp')]
        [String]$DestinationPath
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Add-Type -AssemblyName System.Web

    $baseUri = "https://api.github.com/"
    $WebArgs = "repos/$Owner/$Repository/contents/$Path"
    $wr = Invoke-WebRequest -Uri $($baseUri + $WebArgs) -UseBasicParsing
    $objects = $wr.Content | ConvertFrom-Json
    $files = $objects | Where-Object { $_.type -eq "file" } | Select-Object -exp download_url

    if (-not (Test-Path $DestinationPath)) {
        #Destination path does not exist, create it
        try {
            New-Item -Path $DestinationPath -ItemType Directory -ErrorAction Stop
        }
        catch {
            throw "Could not create path '$DestinationPath'!"
        }
    }

    foreach ($file in $files) {
        $fileDestination = [System.Web.HttpUtility]::UrlDecode((Join-Path $DestinationPath (Split-Path $file -Leaf)))
        try {
            Invoke-WebRequest -Uri $file -OutFile $fileDestination -UseBasicParsing -ErrorAction Stop -Verbose
            "Grabbed '$($file)' to '$fileDestination'"
        }
        catch {
            throw "Unable to download '$($file.path)'"
        }
    }
}