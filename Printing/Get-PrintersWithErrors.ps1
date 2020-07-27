<#
Returns a list of printers that are paused or in an error state.

Andy Morales
#>
$ExcludedPrinterNames = @(
    'Example',
    'Example2'
)

$ExcludedDrivers = @(
    'Example Driver',
    'Example Driver2'
)

$AllPrinters = Get-Printer | Where-Object { $ExcludedPrinterNames -notcontains $_.name -and $ExcludedDrivers -notcontains $_.DriverName }

$PrintersWithErrors = @()

Foreach ($Printer in $AllPrinters) {

    if ($Printer.PrinterStatus -eq 'Paused') {
        $PrintersWithErrors += [PSCustomObject]@{
            Comp      = $Env:COMPUTERNAME
            PrintName = $Printer.name
            DocName   = 'NA'
            JobStatus = 'Printer is paused'
        }
    }

    $PrintJobs = $Printer | Get-PrintJob

    Foreach ($PrintJob in $PrintJobs) {
        If ($PrintJob.jobStatus -like '*Error*') {
            $PrintersWithErrors += [PSCustomObject]@{
                Comp      = $Env:COMPUTERNAME
                PrintName = $PrintJob.PrinterName
                Docname   = $PrintJob.DocumentName
                JobStatus = $PrintJob.jobStatus
            }
        }
    }
}

if ($PrintersWithErrors.count -gt 0) {
    RETURN "Errors: `n$($PrintersWithErrors | Out-String)"
}