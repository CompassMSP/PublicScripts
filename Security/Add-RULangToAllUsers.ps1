<#
This script is intended to help prevent ransomware attacks where it does not infect a computer if the RU keyboard layout is in use.

This will not replace the default, but rather add a second keyboard.

The ES and FR languages will also be added to avoid alarming users when they see the RU language option.

The script will run through all the currently signed in users and add RU as second keyboard. Users will need to sign out for the change to take effect. Running the script repeatedly is recommended.

Andy Morales
#>

$NewLanguageCodes = @(
    '00000419',
    '0000040a',
    '0000040c'
)

#only get reg keys belonging to signed in users
$UserKeys = Get-ChildItem -Path registry::HKEY_USERS | Where-Object { $_.name.Length -eq 57 }

Foreach ($User in $UserKeys) {

    :LangCodes Foreach ($code in $NewLanguageCodes) {

        $InstallLang = $true

        $CurrentLangs = Get-Item -Path "registry::$($user.name)\Keyboard Layout\Preload" | Select-Object -ExpandProperty Property | Where-Object { $_.Length -eq 1 }

        Foreach ($lang in $CurrentLangs) {
            if ((Get-ItemProperty -Path "registry::$($user.name)\Keyboard Layout\Preload" -Name $lang).$lang -eq $code) {
                Write-Output "Lang $($code) is already installed for $($user.name)."

                $InstallLang = $false

                #Move to next Lang
                Continue LangCodes
            }
        }

        if ($InstallLang) {

            $HighestLangNumber = ($CurrentLangs | Measure-Object -Maximum).Maximum

            $NewLangNumber = $HighestLangNumber + 1

            Set-ItemProperty -Path "registry::$($user.name)\Keyboard Layout\Preload" -Name $NewLangNumber -Value $code -Type String
        }
    }
}