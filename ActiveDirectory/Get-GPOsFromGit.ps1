$DestinationFolder = 'C:\Temp\GPOs'
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

Get-GitFilesFromRepo -owner 'CompassMSP' -Repository 'PublicScripts' -Path 'ActiveDirectory/GPOBackups' -DestinationPath $DestinationFolder

$GPORootFolder = $DestinationFolder
$ErrorLogLocation = "$GPORootFolder\errors.txt"

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function Unzip {
        param([string]$zipfile, [string]$outpath)

        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
    }


    $AllZipFiles = Get-ChildItem $GPORootFolder -Filter *.zip

    foreach ($zip in $AllZipFiles) {
        Unzip -zipfile $zip.FullName -outpath $zip.FullName.Replace('.zip', '')
    }

    try {

        Import-Module activedirectory

        $GPORoot = Get-ChildItem -Path $GPORootFolder -Attributes Directory

        $SucessfullyImportedGPOs = @()

        foreach ($folder in $GPORoot) {

            try {
                Remove-Variable GPOFolder, GPOReportPath, GPOReportXML, GPOBackupName -ErrorAction SilentlyContinue

                $GPOFolder = (Get-ChildItem $folder.fullname).FullName
                $GPOReportPath = Get-ChildItem $folder.FullName -Recurse | Where-Object name -EQ gpreport.xml

                #Get the Name of the GPO from the content of the XML
                [XML]$GPOReportXML = Get-Content -Path $GPOReportPath.FullName
                [string]$GPOBackupName = $GPOReportXML.GPO.Name
                $GPOPrefixedName = "_$GPOBackupName"

                New-GPO -Name $GPOPrefixedName -ErrorAction SilentlyContinue
                Import-GPO -Path $GPOFolder -TargetName $GPOPrefixedName -BackupGpoName $GPOBackupName -ErrorAction Stop

                "Sucessfully imported GPO $GPOPrefixedName" | Out-File $ErrorLogLocation -Append
                $SucessfullyImportedGPOs += $GPOPrefixedName

            }
            catch {
                "Error with GPO folder $folder" | Out-File $ErrorLogLocation -Append
                $_ | Out-File $ErrorLogLocation -Append
            }
        }
    }
    Catch {
        $_ | Out-File $ErrorLogLocation -Append
    }

    Write-Output "The following GPOs have been imported sucessfully:"
    $SucessfullyImportedGPOs
}
catch {
    Write-Output 'Ran into error unzipping files. Error has been written to file'
    $_ | Out-File $ErrorLogLocation -Append
}