function Get-GitFilesFromRepo {
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

    $baseUri = "https://api.github.com/"
    $WebArgs = "repos/$Owner/$Repository/contents/$Path"
    $wr = Invoke-WebRequest -Uri $($baseuri + $WebArgs)
    $objects = $wr.Content | ConvertFrom-Json
    $files = $objects | Where-Object { $_.type -eq "file" } | Select-Object -exp download_url
    $directories = $objects | Where-Object { $_.type -eq "dir" }

    $directories | ForEach-Object {
        DownloadFilesFromRepo -Owner $Owner -Repository $Repository -Path $_.path -DestinationPath $($DestinationPath + $_.name)
    }


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
        $fileDestination = Join-Path $DestinationPath (Split-Path $file -Leaf)
        try {
            Invoke-WebRequest -Uri $file -OutFile $fileDestination -ErrorAction Stop -Verbose
            "Grabbed '$($file)' to '$fileDestination'"
        }
        catch {
            throw "Unable to download '$($file.path)'"
        }
    }
}