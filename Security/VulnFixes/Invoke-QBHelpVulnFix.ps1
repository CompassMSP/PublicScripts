<#
This script addresses the vulnerability found here: https://www.securityfocus.com/archive/1/522138

https://www.securityfocus.com/archive/1/522138

Andy Morales
#>
function Get-AuthenticodeSignatureEx {
    <#
    .ForwardHelpTargetName Get-AuthenticodeSignature
    https://stackoverflow.com/questions/15515134/get-signing-timetime-stamp-of-a-digital-signature-using-powershell
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String[]]$FilePath
    )
    begin {
        $signature = @"
[DllImport("crypt32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern bool CryptQueryObject(
    int dwObjectType,
    [MarshalAs(UnmanagedType.LPWStr)]string pvObject,
    int dwExpectedContentTypeFlags,
    int dwExpectedFormatTypeFlags,
    int dwFlags,
    ref int pdwMsgAndCertEncodingType,
    ref int pdwContentType,
    ref int pdwFormatType,
    ref IntPtr phCertStore,
    ref IntPtr phMsg,
    ref IntPtr ppvContext
);
[DllImport("crypt32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern bool CryptMsgGetParam(
    IntPtr hCryptMsg,
    int dwParamType,
    int dwIndex,
    byte[] pvData,
    ref int pcbData
);
[DllImport("crypt32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern bool CryptMsgClose(
    IntPtr hCryptMsg
);
[DllImport("crypt32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern bool CertCloseStore(
    IntPtr hCertStore,
    int dwFlags
);
"@
        Add-Type -AssemblyName System.Security
        Add-Type -MemberDefinition $signature -Namespace PKI -Name Crypt32
    }
    process {
        Get-AuthenticodeSignature @PSBoundParameters | ForEach-Object {
            $Output = $_
            if ($Output.SignerCertificate -ne $null) {
                $pdwMsgAndCertEncodingType = 0
                $pdwContentType = 0
                $pdwFormatType = 0
                [IntPtr]$phCertStore = [IntPtr]::Zero
                [IntPtr]$phMsg = [IntPtr]::Zero
                [IntPtr]$ppvContext = [IntPtr]::Zero
                $return = [PKI.Crypt32]::CryptQueryObject(
                    1,
                    $Output.Path,
                    16382,
                    14,
                    $null,
                    [ref]$pdwMsgAndCertEncodingType,
                    [ref]$pdwContentType,
                    [ref]$pdwFormatType,
                    [ref]$phCertStore,
                    [ref]$phMsg,
                    [ref]$ppvContext
                )
                $pcbData = 0
                $return = [PKI.Crypt32]::CryptMsgGetParam($phMsg, 29, 0, $null, [ref]$pcbData)
                $pvData = New-Object byte[] -ArgumentList $pcbData
                $return = [PKI.Crypt32]::CryptMsgGetParam($phMsg, 29, 0, $pvData, [ref]$pcbData)
                $SignedCms = New-Object Security.Cryptography.Pkcs.SignedCms
                $SignedCms.Decode($pvData)
                foreach ($Infos in $SignedCms.SignerInfos) {
                    foreach ($CounterSignerInfos in $Infos.CounterSignerInfos) {
                        $sTime = ($CounterSignerInfos.SignedAttributes | ? { $_.Oid.Value -eq "1.2.840.113549.1.9.5" }).Values | `
                            Where-Object { $_.SigningTime -ne $null }
                    }
                }
                $Output | Add-Member -MemberType NoteProperty -Name SigningTime -Value $sTime.SigningTime.ToLocalTime() -PassThru -Force
                [void][PKI.Crypt32]::CryptMsgClose($phMsg)
                [void][PKI.Crypt32]::CertCloseStore($phCertStore, 0)
            }
            else {
                $Output
            }
        }
    }
    end {
    }
}

$VulnerableFile = 'C:\Program Files (x86)\Intuit\*\HelpAsyncPluggableProtocol.dll'

if (Test-Path $VulnerableFile) {
    $Signature = Get-Item -Path $VulnerableFile | Get-AuthenticodeSignatureEx

    if ($Signature.SigningTime -lt (Get-Date -Year 2012 -Month 4 -Day 27)){
        Write-Output 'Vulnerable version of the file were found.'

        $PossibleRegKeys = @(
            'HKLM:\SOFTWARE\WOW6432Node\Classes\PROTOCOLS\Handler\intu-help-qb1',
            'HKLM:\SOFTWARE\WOW6432Node\Classes\PROTOCOLS\Handler\intu-help-qb2',
            'HKLM:\SOFTWARE\WOW6432Node\Classes\PROTOCOLS\Handler\intu-help-qb3',
            'HKLM:\SOFTWARE\WOW6432Node\Classes\PROTOCOLS\Handler\intu-help-qb4',
            'HKLM:\SOFTWARE\WOW6432Node\Classes\PROTOCOLS\Handler\intu-help-qb5'
        )

        Foreach ($Key in $PossibleRegKeys) {
            if (Test-Path -Path $key) {
                Rename-Item -Path $Key -NewName "$($Key.Split('\')[-1]).renamed"
                Write-Output "The key $($Key) has been renamed"
            }
        }
    }
}
else {
    Write-Output 'No Vulnerable Files Found'
}
