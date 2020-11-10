function Get-RandomCharacters {
    #https://activedirectoryfaq.com/2017/08/creating-individual-random-passwords/
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $false)]
        [String]$Length = '12',
        
        [parameter(Mandatory = $false)]
        [switch]$AsSecureString
    )

    $RandomPassword = ''

    $LowerCaseChars = 'abcdefghkmnoprstuvwxyz'
    $UpperCaseChars = 'ABCDEFGHKMNPRSTUVWXYZ'
    $NumberChars = '23456789'
    $SpecialChars = '@#$%-+*_=?:<>^&'
    $AllCharacters = $LowerCaseChars + $UpperCaseChars + $NumberChars + $SpecialChars

    #generate random strings until they contain one of each character type. Put a limit on the loop so it only runs 10 times
    DO {
        $bytes = New-Object "System.Byte[]" $Length
        $rnd = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
        $rnd.GetBytes($bytes)
        $RandomCharacters = ""

        for ( $i = 0; $i -lt $Length; $i++ ) {
            $RandomCharacters += $AllCharacters[ $bytes[$i] % $AllCharacters.Length ]
        }

        $count++
    }Until (($RandomCharacters -cmatch '(?=.*[a-z])(?=.*[A-Z])(?=.*[@#$%-+*_=?:<>^&])(?=.*\d)') -or ($count -ge 10))

    if ($AsSecureString){
        [securestring]$RandomCharacters = ConvertTo-SecureString $RandomCharacters -AsPlainText -Force
    }

    return $RandomCharacters
}