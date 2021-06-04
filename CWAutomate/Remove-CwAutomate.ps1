Function Remove-CWAutomate {
    <#
    This function removed CW automate from a computer.

    .DESCRIPTION

    The Agent_Uninstaller.zip will be downloaded an ran on the computer. Afterwards a manual cleanup of any leftover items will run.

    Andy Morales
    #>

    function Expand-ZIP {
        <#
        Extracts a ZIP file to a directory. The contents of the destination will be deleted if they already exist.

        Andy Morales
        #>
        [CmdletBinding()]
        param (
            [parameter(Mandatory = $true)]
            [String]$ZipFile,

            [parameter(Mandatory = $true)]
            [String]$OutPath
        )
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        if (Test-Path -Path $OutPath) {
            Remove-Item $OutPath -Recurse -Force
        }

        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFile, $OutPath)
    }

    Write-Output 'Downloading and running Agent_Uninstaller.'

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    (New-Object System.Net.WebClient).DownloadFile('https://s3.amazonaws.com/assets-cp/assets/Agent_Uninstaller.zip', 'C:\Windows\Temp\LT_Agent_Uninstaller.zip')

    #unzip
    Expand-ZIP -ZipFile 'C:\Windows\Temp\LT_Agent_Uninstaller.zip' -OutPath 'C:\Windows\Temp\LTAgentUninstaller'

    #run uninstaller
    Start-Process -FilePath "C:\Windows\Temp\LTAgentUninstaller\Agent_Uninstall.exe" -Wait

    Write-Output 'Manually removing any items that may have been left over.'

    #region processes
    $Processes = @(
        'lttray',
        'ltservice',
        'ltsvc',
        'ltsvcmon'
    )

    $Processes | ForEach-Object { Stop-Process -Name $_ -Force -ErrorAction SilentlyContinue }

    #EndRegion processes

    Start-Process "$($env:SystemRoot)\System32\regsvr32.exe" -ArgumentList "/u C:\WINDOWS\LTSvc\wodVPN.dll /s" -Wait

    #region Services
    $LtServicesName = @(
        'ltservice',
        'ltsvc',
        'ltsvcmon'
    )

    $LtServices = Get-WmiObject -Class Win32_Service | Where-Object { $LtServicesName -contains $_.name }

    Foreach ($service in $LtServices) {

        if ($service.ProcessId -gt 0) {
            Start-Process "$($env:SystemRoot)\System32\taskkill.exe" -ArgumentList "/pid $($service.ProcessId) /f"  -Wait
        }

        Start-Process "$($env:SystemRoot)\System32\sc.exe" -ArgumentList "delete $($service.name)"  -Wait
    }
    #endregion Services

    #region RegKeys
    $RegKeys = @(
        'Registry::HKEY_LOCAL_MACHINE\software\labtech',
        'Registry::HKEY_LOCAL_MACHINE\software\wow6432node\labtech',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{3426921d-9ad5-4237-9145-f15dee7e3004}',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Appmgmt\{40bf8c82-ed0d-4f66-b73e-58a3d7ab6582}',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed\Installer\Products\C4D064F3712D4B64086B5BDE05DBC75F',
        'Registry::HKEY_CLASSES_ROOT\Installer\Dependencies\{3426921d-9ad5-4237-9145-f15dee7e3004}',
        'Registry::HKEY_CLASSES_ROOT\Installer\Dependencies\{3F460D4C-D217-46B4-80B6-B5ED50BD7CF5}',
        'Registry::HKEY_CLASSES_ROOT\Installer\Products\C4D064F3712D4B64086B5BDE05DBC75F',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{09DF1DCA-C076-498A-8370-AD6F878B6C6A}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{15DD3BF6-5A11-4407-8399-A19AC10C65D0}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{3C198C98-0E27-40E4-972C-FDC656EC30D7}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{459C65ED-AA9C-4CF1-9A24-7685505F919A}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{7BE3886B-0C12-4D87-AC0B-09A5CE4E6BD6}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{7E092B5C-795B-46BC-886A-DFFBBBC9A117}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{9D101D9C-18CC-4E78-8D78-389E48478FCA}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{B0B8CDD6-8AAA-4426-82E9-9455140124A1}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{B1B00A43-7A54-4A0F-B35D-B4334811FAA4}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{BBC521C8-2792-43FE-9C91-CCA7E8ACBCC9}',
        'Registry::HKEY_CLASSES_ROOT\CLSID\{C59A1D54-8CD7-4795-AEDD-F6F6E2DE1FE7}',
        'Registry::HKEY_CLASSES_ROOT\Installer\Products\C4D064F3712D4B64086B5BDE05DBC75F'
    )

    foreach ($key in $RegKeys) {
        Remove-Item -Path $key -Force -Verbose -Recurse -ErrorAction SilentlyContinue
    }

    $uninstallKeys = @(
        '3426921d-9ad5-4237-9145-f15dee7e3004',
        '3F460D4C-D217-46B4-80B6-B5ED50BD7CF5',
        '02ff82a3-f67d-4d3f-bc33-26c877c793a7',
        'b5ff5d67-cb77-4a86-8398-d20c2acef43a'
    )

    foreach ($uKey in $uninstallKeys) {
        Remove-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\UNINSTALL\$($key)" -Force -Verbose -Recurse -ErrorAction SilentlyContinue
        Remove-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432NODE\MICROSOFT\WINDOWS\CURRENTVERSION\UNINSTALL\$($key)" -Force -Verbose -Recurse -ErrorAction SilentlyContinue
    }
    #endRegion RegKeys
    if (Test-Path -Path "$env:windir\ltsvc") {
        Remove-Item -Path "$env:windir\ltsvc" -Force -Recurse -ErrorAction SilentlyContinue
    }
}

Remove-CWAutomate