Function Get-PrintersWithErrors {
    <#
    Returns a list of printers that are paused or in an error state.

    The script is also able to delete jobs in an error state by setting $ClearErrorJobs  to true

    Andy Morales
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$ClearErrorJobs,

        [Parameter(Mandatory = $false)]
        [int]$OldJobDaysThreshold = 3
    )


    $ExcludedPrinterNames = @(
        'Example',
        'Example2'
    )

    $ExcludedDrivers = @(
        'Example Driver',
        'Example Driver2'
    )

    $AllPrinters = Get-Printer | Where-Object { $ExcludedPrinterNames -notContains $_.name -and $ExcludedDrivers -notContains $_.DriverName }

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
            #Identify jobs in an error state, or that have been queued for too many days
            If (($PrintJob.jobStatus -like '*Error*') -or ($PrintJob.SubmittedTime -lt (Get-Date).AddDays(-$OldJobDaysThreshold))) {
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
}