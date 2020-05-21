#This checks the output of the Test-RegistryValue function
#Andy Morales

#Directory of the script
$script_dir = Split-Path -Parent $MyInvocation.MyCommand.Path

#Load the function
. "$script_dir\Get-RandomCharacters.ps1"

$RegexPassword = '(?=.*[a-z])(?=.*[A-Z])(?=.*[@#$%-+*_=?:<>^&])(?=.*\d)'

$TestData = @(
    @{
        'TestName'       = '12 Characters';
        'StringLength'   = '12';
        'ExpectedResult' = $RegexPassword
    },
    @{
        'TestName'       = '50 Characters';
        'StringLength'   = '50';
        'ExpectedResult' = $RegexPassword
    },
    @{
        'TestName'       = '8 Characters';
        'StringLength'   = '8';
        'ExpectedResult' = $RegexPassword
    },
    @{
        'TestName'       = '4 Characters';
        'StringLength'   = '8';
        'ExpectedResult' = $RegexPassword
    }
)

Describe 'Test-RegistryValue Function' {
    #Loop through and identify any output that does not match the expected string content
    for ($i = 0; $i -lt 10; $i++) {
        It "<TestName> Character Test" -TestCases $TestData {
            Get-RandomCharacters -Length $StringLength | Should -Match $ExpectedResult
        }
    }

    #Check to make sure that the output matches the expected length
    It "<TestName> Character Length" -TestCases $TestData {
        (Get-RandomCharacters -Length $StringLength).Length | Should -be $StringLength
    }
}