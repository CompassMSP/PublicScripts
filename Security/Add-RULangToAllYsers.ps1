<#

This script is intended to help prevent ransomware attacks where it does not infect a computer if the RU keyboard layout is in use.

This will not replace the default, but rather add a second keyboard.

The script will run through all the currently signed in users and add RU as second keyboard. Users will need to sign out for the change to take effect. Running the script repeatedly is recomended.

Andy Morales
#>

#only get reg keys belonging to signed in users
$UserKeys = Get-ChildItem -Path registry::HKEY_USERS | Where-Object {$_.name.Length -eq 57}


Foreach ($User in $UserKeys){

    $CurrentLangs =  (Get-Item -path "registry::$($user.name)\Keyboard Layout\Preload" | Select-Object -ExpandProperty Property | Where-Object {$_.Length -eq 1} | measure -Maximum).Maximum

    Foreach ($lang in $CurrentLangs){
        if((Get-ItemProperty -path "registry::$($user.name)\Keyboard Layout\Preload" -name $lang).$lang -eq '00000419'){

            Write-Output 'Lang is already installed. Exit script'
            Exit

        }
    }

    $HighestLangNumber = ($CurrentLangs| measure -Maximum).Maximum

    $NewLangNumber = $HighestLangNumber + 1

    Set-ItemProperty -Path "registry::$($user.name)\Keyboard Layout\Preload" -Name $NewLangNumber -Value '00000419' -Type String

}