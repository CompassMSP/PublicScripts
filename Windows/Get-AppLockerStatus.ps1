try {
    if (
        ((Get-AppLockerPolicy -Effective -ErrorAction Stop).rulecollections.count -gt 0) -and
        ((Get-Service -Name AppIDSvc).Status -eq 'Running')
    ) {
        Return 'AppLocker Enabled'
    }
    else {
        Return 'AppLocker Inactive'
    }
}
catch {
    Write-Verbose 'Could not determine AppLocker Status. Computer OS is probably too old.'
}