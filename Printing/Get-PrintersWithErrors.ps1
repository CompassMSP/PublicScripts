<#
Returns a list of printers that are paused or in an error state.

The script is also able to delete jobs in an error state by setting $ClearErrorJobs  to true

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

#Setting this to true will attempt to delete jobs in an error state
$ClearErrorJobs = $false

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
            
            if ($ClearErrorJobs) {
                $PrintJob | Remove-PrintJob
            }
        }
    }
}

if ($PrintersWithErrors.count -gt 0) {
    RETURN "Errors: `n$($PrintersWithErrors | Out-String)"
}