$orgName = Read-Host "Please enter your tenant name to connect to SharePoint. EX: https://contoso.sharepoint.com"

$orgSharePointURL = "https://$($orgName)-admin.sharepoint.com"
Connect-SPOService -Url $orgSharePointURL

Get-SPOSite -Limit all -Filter "Url -like 'sharepoint.com/sites'" | here-Object {$_.ListsShowHeaderAndNavigation -eq $False} | ForEach-Object { Set-SPOSite -Identity $_.Url -ListsShowHeaderAndNavigation $true }