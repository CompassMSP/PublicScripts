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
        catch{
            Write-Output "Could not clear the contents of $($Path). Script will exit."
            EXIT
        }
    }

    #Create the folder
    New-Item -Path $Path -ItemType Directory -Force

    #Remove all explicit permissions
    ICACLS ("$Path") /reset | Out-Null

    #Add SYSTEM permission
    ICACLS ("$Path") /grant ("SYSTEM" + ':(OI)(CI)F') | Out-Null

    #Give Administrators Full Control
    ICACLS ("$Path") /grant ("Administrators" + ':(OI)(CI)F') | Out-Null

    #Disable Inheritance on the Folder. This is done last to avoid permission errors.
    ICACLS ("$Path") /inheritance:r | Out-Null
}  