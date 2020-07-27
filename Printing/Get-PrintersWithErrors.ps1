<#
Returns a list of printers that are paused or in an error state.

Andy Morales
#>
$ExcludedPrinters = @(
    'Example',
    'Example2'
)

$AllPrinters = Get-Printer | Where-Object { $ExcludedPrinters -notcontains $_.name }

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