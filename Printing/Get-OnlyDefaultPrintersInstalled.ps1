<#
This script enumerates all the printers on a computer to see if only the default printers exist.

.LINK
https://github.com/gtworek/PSBits/blob/master/Misc/StopAndDisableDefaultSpoolers.ps1

Andy Morales
#>

$service = Get-Service -Name Spooler -ErrorAction SilentlyContinue

if (!$service) {
    #Cannot connect to Spooler Service
    EXIT
}

if ($service.Status -ne "Running") {
    Write-Output "SpoolerNotRunning"
    EXIT
}

$printers = Get-WmiObject -Class Win32_printer

if (!$printers) {
    Write-Output "CannotEnumeratePrinters"
    EXIT
}

$OnlyDefaultPrintersExist = $true

foreach ($DriverName in ($printers.DriverName)) {
    if (($DriverName -notmatch 'Microsoft XPS Document Writer') -and ($DriverName -notmatch 'Microsoft Print To PDF')) {
        $OnlyDefaultPrintersExist = $false
        BREAK
    }
}

if ($OnlyDefaultPrintersExist) {
    Write-Output "OnlyDefaultPrintersFound"
}
else {
    Write-Output "NonDefaultPrintersFound"
}