#CSV Headers - EmailAddress,Title,Department,MobilePhone,OfficePhone,ManagerEmailAddress


Import-CSV -Path "C:\scripts\user_data.csv" | Foreach-Object {

    $ADUser = Get-ADUser -Filter "EmailAddress -eq '$($_.EmailAddress)'" -Properties * 
    
    if ($_.Title) {     
        $ADUser | Set-ADUser -Title $_.Title 
        Write-Host "Updated Title for $($ADUser.SamAccountName)"
    }

    if ($_.Department) {     
        $ADUser | Set-ADUser -Department $_.Department
        Write-Host "Updated Department for $($ADUser.SamAccountName)"
    }
    
    if ($_.MobilePhone) {     
        $ADUser | Set-ADUser -Mobile $_.MobilePhone
        Write-Host "Updated Mobile for $($ADUser.SamAccountName)"
    }

    if ($_.OfficePhone) {     
        $ADUser | Set-ADUser –OfficePhone $_.OfficePhone
        Write-Host "Updated OfficePhone for $($ADUser.SamAccountName)"
    }

    if ($_.ManagerEmailAddress) {
        $managerDN = (Get-ADUser -Filter "EmailAddress -eq '$($_.ManagerEmailAddress)'").DistinguishedName
        $ADUser | Set-ADUser -Manager $managerDN
        Write-Host "Updated Manager for $($ADUser.SamAccountName)"
    }

}

