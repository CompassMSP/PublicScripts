#CSV Headers - EmailAddress,Title,Department,MobilePhone,OfficePhone,ManagerEmailAddress


# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Directory.ReadWrite.All"

Import-CSV -Path "C:\scripts\user_data.csv" | Foreach-Object {

    $UserId = (Get-MgUser -Filter "EmailAddress eq '$($_.EmailAddress)'").Id
    
    if ($_.Title) {     
        Update-MgUser -UserId $UserId -JobTitle $_.Title 
        Write-Host "Updated Title for $($ADUser.SamAccountName)"
    }

    if ($_.Department) {     
        Update-MgUser -UserId $UserId Department $_.Department
        Write-Host "Updated Department for $($ADUser.SamAccountName)"
    }
    
    if ($_.MobilePhone) {     
        Update-MgUser -UserId $UserId -Mobile $_.MobilePhone
        Write-Host "Updated Mobile for $($ADUser.SamAccountName)"
    }

    if ($_.OfficePhone) {     
        Update-MgUser -UserId $UserId –OfficePhone $_.OfficePhone
        Write-Host "Updated OfficePhone for $($ADUser.SamAccountName)"
    }

    if ($_.ManagerEmailAddress) {
        $ManagerBinding = @{
            "@odata.bind" = @("https://graph.microsoft.com/v1.0/users/"+$(Get-MgUser -Filter "UserPrincipalName eq '$($Manager)'").Id)
        }
        Set-MgUserManagerByRef -UserId $UserId -AdditionalProperties $ManagerBinding
        Write-Host "Updated Manager for $($ADUser.SamAccountName)"
    }

}
