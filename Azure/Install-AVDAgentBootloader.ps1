# Download Azure Virtual Desktop agent 
Invoke-WebRequest -URI https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RWrmXv -OutFile "C:\Windows\temp\RDAgent.msi"
# Download Azure Virtual Desktop agent bootloader
Invoke-WebRequest -URI https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RWrxrH -OutFile "C:\Windows\temp\RDBootstrap.msi"
# Install Azure Virtual Desktop agent - use your registration key for REGISTRATIONTOKEN
Start-Process -FilePath msiexec -ArgumentList "/i C:\Windows\Temp\RDAgent.msi REGISTRATIONTOKEN=$($RegistrationToken.token) /qn /norestart /passive /lv* C:\Windows\Temp\rdagentinstall.log"
# Install Azure Virtual Desktop bootstrap agent
Start-Process -FilePath msiexec -ArgumentList "/i C:\Windows\Temp\RDBootstrap.msi /qn /norestart /passive /lv* C:\Windows\Temp\RDBootstrapinstall.log"
