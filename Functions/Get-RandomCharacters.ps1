function Get-RandomCharacters {
    #https://activedirectoryfaq.com/2017/08/creating-individual-random-passwords/
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $false)]
        [String]$Length = '12'
    )

    $RandomPassword = ''

    $LowerCaseChars = 'abcdefghkmnoprstuvwxyz'
    $UpperCaseChars = 'ABCDEFGHKMNPRSTUVWXYZ'
    $NumberChars = '23456789'
    $SpecialChars = '@#$%-+*_=?:<>^&'
    $AllCharacters = $LowerCaseChars + $UpperCaseChars + $NumberChars + $SpecialChars

    #generate random strings until they contain one of each character type. Put a safety on the loop so it only runs 10 times
    DO {
        $random = 1..$length | ForEach-Object { Get-Random -Maximum $AllCharacters.length }
        $private:ofs = ""

        $RandomCharacters = [String]$AllCharacters[$random]

        $count++
    }Until (($RandomCharacters -cmatch '(?=.*[a-z])(?=.*[A-Z])(?=.*[@#$%-+*_=?:<>^&])(?=.*\d)') -or ($count -ge 10))

    return $RandomCharacters
}