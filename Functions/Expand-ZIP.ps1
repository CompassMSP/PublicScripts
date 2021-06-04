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