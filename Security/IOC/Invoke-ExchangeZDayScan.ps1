<#

Merge of the two scripts below
https://github.com/mr-r3b00t/ExchangeMarch2021IOCHunt/blob/main/fastcheck.ps1
https://github.com/microsoft/CSS-Exchange/blob/main/Security/Test-ProxyLogon.ps1

#>

Function Write-Log {
    param(
        [Parameter(Mandatory = $true)][String]$msg
    )
    Add-Content "c:\exchangeLog.txt" $msg
}

#Check to see if Exchange 2013+ is installed
if ($env:exchangeInstallPath) {

    if(Test-Path 'c:\exchangeLog.txt'){
        Rename-Item -Path 'c:\exchangeLog.txt' -NewName "exchangeLog$((Get-ItemProperty 'c:\exchangeLog.txt').LastWriteTime.ToString("yyyyMMdd-hh-mm-ss")).txt"
    }

    $errorFound = 0

    #region CheckHashes
    $badHashes = @(
        "b75f163ca9b9240bf4b37ad92bc7556b40a17e27c2b8ed5c8991385fe07d17d0",
        "097549cf7d0f76f0d99edf8b2d91c60977fd6a96e4b8c3c94b0b1733dc026d3e",
        "2b6f1ebb2208e93ade4a6424555d6a8341fd6d9f60c25e44afe11008f5c1aad1",
        "65149e036fff06026d80ac9ad4d156332822dc93142cf1a122b1841ec8de34b5",
        "511df0e2df9bfa5521b588cc4bb5f8c5a321801b803394ebc493db1ef3c78fa1",
        "4edc7770464a14f54d17f36dc9d0fe854f68b346b27b35a6f5839adf1f13f8ea",
        "811157f9c7003ba8d17b45eb3cf09bef2cecd2701cedb675274949296a6a183d",
        "1631a90eb5395c4e19c7dbcbf611bbe6444ff312eb7937e286e4637cb9e72944"
    )

    Write-Log "Checking C:\inetpub\wwwroot\aspnet_client for extra files"

    $enumFiles = Get-ChildItem -Path C:\inetpub\wwwroot\aspnet_client -Recurse -File -Force

    foreach ($file in $enumFiles) {

        $fileHash = Get-FileHash -Path $file.FullName -Algorithm SHA256

        if ($badHashes.Contains($fileHash.Hash)) {
            $errorFound = 1
            Write-Log $file.DirectoryName
            Write-Log $file.FullName
            Write-Log $file.Name
            Write-Log " "
            Write-Log "BAD HASH DETECTED ASSUME BREACH"
        }
    }

    Write-Log " "

    #endregion CheckHashes

    #region SuspiciousFiles
    Write-Log "`r`nChecking for suspicious files"
    $lsassFiles = @(Get-ChildItem -Recurse -Path "$env:WINDIR\temp\lsass.*dmp" -Force)
    $lsassFiles += @(Get-ChildItem -Recurse -Path "c:\root\lsass.*dmp"  -Force)
    if ($lsassFiles.Count -gt 0) {
        Write-Warning "lsass.exe dumps found, please verify these are expected:"
        $lsassFiles.FullName
    }
    else {
        Write-Log "No suspicious lsass dumps found."
    }

    $zipFiles = @(Get-ChildItem -Recurse -Path "$env:ProgramData" -ErrorAction SilentlyContinue -Force | Where-Object { $_.Extension -match ".7z| .zip | .rar" } | Where-Object { $_.FullName -notMatch 'Kaseya|LTDatabase|VMware Tools|LabTech|chocolatey|Norton|Symantec' })

    if ($zipFiles.Count -gt 0) {
        Write-Log "`r`nZipped files found in $env:ProgramData, please verify these are expected:"
        Write-Log ($zipFiles.FullName | Out-String)
    }
    else {
        Write-Log "`r`nNo suspicious zip files found."
    }
    #endregion SuspiciousFiles

    #region oddASPX
    Write-Log " "
    Write-Log "looking for odd aspx files"


    $oddASPX = @()
    $oddASPX += (Get-ChildItem -Path C:\inetpub\wwwroot\aspnet_client\ -Recurse -Filter "*.aspx" -Force)

    if ($oddASPX.count -gt 0) {
        Write-Log ($oddASPX | Out-String)

        $errorFound = 1
    }

    Write-log "looking for odd aspx files (default names are 'errorFE.aspx', 'ExpiredPassword.aspx', 'frowny.aspx', 'logoff.aspx', 'logon.aspx', 'OutlookCN.aspx'.'RedirSuiteServiceProxy.aspx', 'signout.aspx'"

    $owaAspx = @()
    $owaAspx += (Get-ChildItem -Path "$($env:exchangeInstallPath)FrontEnd\HttpProxy\owa\auth\" -Recurse -Filter "*.aspx*" -Force | Select-Object -ExpandProperty fullName | Where-Object { $_ -notMatch 'errorFE|SvmFeedback|ExpiredPassword|frowny|logoff|logon.aspx|OutlookCN|RedirSuiteServiceProxy|signout' })

    if ($owaAspx.count -gt 0) {

        $errorFound = 1

        foreach ($owaFile in $owaAspx) {
            Write-Log $owaFile
        }
    }
    else {
        Write-Log "No odd aspx files"
    }

    Write-Log " "

    #endregion #region oddASPX

    #region 26855
    Write-Log "Checking for CVE-2021-26855 in the HttpProxy logs"
    $files = (Get-ChildItem -Recurse -Path "$($env:exchangeInstallPath)Logging\HttpProxy" -Filter '*.log' -Force).FullName
    $count = 0
    $allResults = @()
    $sw = New-Object System.Diagnostics.Stopwatch
    $sw.Start()
    $files | ForEach-Object {
        $count++
        if ($sw.ElapsedMilliseconds -gt 500) {
            Write-Progress -Activity "Checking for CVE-2021-26855 in the HttpProxy logs" -Status "$count / $($files.Count)" -PercentComplete ($count * 100 / $files.Count)
            $sw.Restart()
        }
        if ((Get-ChildItem $_ -ErrorAction SilentlyContinue -Force | Select-String "ServerInfo~").Count -gt 0) {
            $fileResults = @(Import-Csv -Path $_ -ErrorAction SilentlyContinue | Where-Object { $_.AnchorMailbox -like 'ServerInfo~*/*' })
            $fileResults | ForEach-Object {
                $allResults += $_
            }
        }
    }

    Write-Progress -Activity "Checking for CVE-2021-26855 in the HttpProxy logs" -Completed

    if ($allResults.Length -gt 0) {
        Write-Log "Suspicious entries found in $($env:exchangeInstallPath)Logging\HttpProxy.  Check the .\CVE-2021-26855.csv log for specific entries."

        write-log ($allResults | Select-Object DateTime, RequestId, ClientIPAddress, UrlHost, UrlStem, RoutingHint, UserAgent, AnchorMailbox, HttpStatus | Out-String)

        $errorFound = 1

    }
    else {
        Write-Log "No suspicious entries found."
    }
    #endregion 26855

    #region 26858
    Write-Log "`r`nChecking for CVE-2021-26858 in the OABGenerator logs"
    $OABLogs = Get-ChildItem -Recurse -Path "$($env:exchangeInstallPath)Logging\OABGeneratorLog" -Force | Select-String "Download failed and temporary file" -List | Select-Object Path
    if ($OABLogs.Path.Count -gt 0) {
        Write-Log "Suspicious OAB download entries found in the following logs, please review them for `"Download failed and temporary file`" entries:"
        Write-Log ($OABLogs.Path | Out-String)

        $errorFound = 1
    }
    else {
        Write-Log "No suspicious entries found."
    }
    #endregion 26858

    #region 26857
    Write-log "`r`nChecking for CVE-2021-26857 in the Event Logs"
    $eventLogs = @(Get-WinEvent -FilterHashtable @{LogName = 'Application'; ProviderName = 'MSExchange Unified Messaging'; Level = '2' } -ErrorAction SilentlyContinue | Where-Object { $_.Message -like "*System.InvalidCastException*" })
    if ($eventLogs.Count -gt 0) {
        Write-log "Suspicious event log entries for Source `"MSExchange Unified Messaging`" and Message `"System.InvalidCastException`" were found.  These may indicate exploitation.  Please review these event log entries for more details."
    }
    else {
        Write-log "No suspicious entries found."
    }
    #endregion 26857

    #region 27065
    Write-log "`r`nChecking for CVE-2021-27065 in the ECP Logs"
    $ECPLogs = Get-ChildItem -Recurse -Path "$($env:exchangeInstallPath)Logging\ECP\Server\*.log" -Force | Select-String "Set-.*VirtualDirectory" -List | Select-Object Path
    if ($ECPLogs.Path.Count -gt 0) {
        Write-Log "Suspicious virtual directory modifications found in the following logs, please review them for `"Set-*VirtualDirectory`" entries:"
        Write-Log ($ECPLogs.Path | Out-String)

        $errorFound = 1
    }
    else {
        Write-Log "No suspicious entries found."
    }
    #endregion 27065

    #region IIS-W3SVC-WP

    #there should be no events
    Write-Log "Checking IIS-W3SVC-WP event logs"

    Try {
        Write-Log (Get-Event -LogName Application -Source IIS-W3SVC-WP -InstanceId 2303 -ErrorAction stop)

        $errorFound = 1
    }
    Catch {
        Write-Log "No Event logs with source IIS-W3SVC-WP"
    }
    #endregion IIS-W3SVC-WP

    #region IIS-APPHOSTSVC
    Write-Log " "
    Write-Log "Checking IIS-APPHOSTSVC event logs"

    Try {
        Write-Log (Get-Event -LogName Application -Source IIS-APPHOSTSVC -InstanceId 9009 -ErrorAction Stop)
        $errorFound = 1
    }
    Catch {
        Write-Log "No Event logs with source IIS-APPHOSTSVC"
    }

    Write-Log " "
    #endregion IIS-APPHOSTSVC

    #region OABGenerator
    #there should be no entries
    Write-Log "Checking OABGenerator logs"

    Try {
        Write-Log (findstr /snip /c:"Download failed and temporary file" "$($env:exchangeInstallPath)Logging\OABGeneratorLog\*.log")
        $errorFound = 1
    }
    Catch {
        Write-Log "No OABGenerator logs"
    }

    Write-Log " "

    #endregion OABGenerator

    #region UnifiedMessage
    #there should be no events
    Write-Log "Checking Unified Message event logs"

    Try {
        Write-Log (Get-Event -LogName Application -Source "MSExchange Unified Messaging" -EntryType Error -ErrorAction stop | Where-Object { $_.Message -like "*System.InvalidCastException*" } )
        $errorFound = 1
    }
    Catch {
        Write-Log "No Unified Message event logs"
    }

    Write-Log " "
    #endregion UnifiedMessage

    #region SetVirtualDirectory
    #this should be blank
    Write-Log "Checking for Set-VirtualDirectory indicators"
    try {
        Write-Log (Select-String -Path "$($env:exchangeInstallPath)Logging\ECP\Server\*.log" -Pattern 'Set-.+VirtualDirectory' -ErrorAction silentlyContinue)
        $errorFound = 1
    }
    catch {
        Write-Log "No Set-VirtualDirectory indicators"
    }

    Write-Log " "
    #endregion SetVirtualDirectory

    #region IIS
    #read all the IIS logs looking for POST requests to /owa/auth/Current/themes/resources/
    Write-Log "Checking for theme resource indicators"
    $parse1 = Select-String -Path "C:\inetpub\logs\LogFiles\W3SVC1\*.log" -Pattern 'POST /owa/auth/Current/themes/resources/'
    if ($parse1 -eq "") {
        $errorFound = 1
    }
    else {
        foreach ($line in $parse1) {
            Write-Log "Might want to investigate this"
            Write-Log $line
        }
    }

    Write-Log " "
    #endregion IIS

    Write-Log "all testing completed, please review the above Write-Log for any suspicious files or activity"
    $errorFound
}
else {
    Write-Log 'Exchange 2013+ not detected'
}