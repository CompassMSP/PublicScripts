<#
This script will run through all established connections on a server and identify the country of the IP address.

Andy Morales
#>
#get all open network connections
$AllEstablished = Get-NetTCPConnection -State Established

#get all running processes
try {
    $AllProcesses = Get-Process -IncludeUserName -ErrorAction stop
}
catch {
    $AllProcesses = Get-Process
}

$rfc1918regex = '(192\.168\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5]))|(172\.([1][6-9]|[2][0-9]|[3][0-1])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5]))|(10\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5]))|(127\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5]))'
$IpAddressRegex = '\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'

#filter connections to only Public IPs
$InternetConnections = @()
foreach ($connection in $AllEstablished) {
    if (($connection.RemoteAddress -notmatch $rfc1918regex) -and ($connection.RemoteAddress -match $IpAddressRegex)) {
        $InternetConnections += $connection
    }
}

$IPCountry = @()

#get country info
Foreach ($IP in $InternetConnections) {

    Clear-Variable countryInfo -ErrorAction ignore

    try {
        $CountryInfo = Invoke-RestMethod -Method Get -Uri "http://ip-api.com/json/$($IP.RemoteAddress)" -ErrorAction Stop
    }
    catch {
        #create an error object. Probably too many API requests.
        $CountryInfo = [PSCustomObject]@{
            country    = 'API Error'
            city       = 'API Error'
            org        = 'API Error'
            regionName = 'API Error'
        }
    }

    #get the matching process
    Foreach ($process in $AllProcesses) {
        If ($process.id -eq $IP.OwningProcess) {
            $CurrentProcInfo = $process
            BREAK
        }
    }

    #return results
    $IPCountry += [PSCustomObject]@{
        IPAddress   = $IP.RemoteAddress
        ProcessID   = $IP.OwningProcess
        ProcessName = $CurrentProcInfo.ProcessName
        Username    = $CurrentProcInfo.Username
        LocalPort   = $IP.LocalPort
        RemotePort  = $IP.RemotePort
        Country     = $CountryInfo.country
        regionName  = $CountryInfo.regionName
        city        = $CountryInfo.city
        org         = $CountryInfo.org
    }
}

$IPCountry | Sort-Object Country, regionName, City | ft