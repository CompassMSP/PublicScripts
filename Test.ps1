$LatestVersionUrl = (Invoke-WebRequest https://haveibeenpwned.com/Passwords -MaximumRedirection 0).Links | Where-Object {$_.href -like "*pwned-passwords-ntlm-ordered-by-hash-v*.7z"} | Select-Object -expand href

#Variables built out for script 
$LatestVersionZip = $($LatestVersionUrl -replace '[a-zA-Z]+://[a-zA-Z]+\.[a-zA-Z]+\.[a-zA-Z]+/[a-zA-Z]+/')
$LatestVersionTXT = $($LatestVersionZip -replace '.7z') + '.txt'
$LatestVersionLog = $($LatestVersionZip -replace 'pwned-passwords-ntlm-ordered-by-hash-') 
$LatestVersionLog = $($LatestVersionLog -replace '.7z')

$PassProtectionPath = 'C:\Program Files\Lithnet\Active Directory Password Protection\'

Write-Output "URL is $LatestVersionUrl"
Write-Output "Log version is $LatestVersionLog."
Start-Sleep 10
Write-Output "TXT is $($LatestVersionTXT)."
