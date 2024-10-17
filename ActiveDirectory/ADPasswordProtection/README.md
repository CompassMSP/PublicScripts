# Requirements

- The application must be installed on every DC at a site
- DCs must have at least 45GB free during setup, and 15GB free once the setup is complete. This space is required for the HIBP database
- DCs must be 2016+ (or have WMF 5.1 (Powershell 5.1) installed if 2012 R2) for the automated script to run
  - To check for WMF version open Powershell and run "$PSVersionTable.PSVersion"
    - WMF 5.1 for 2012 R2 link: https://docs.microsoft.com/en-us/powershell/scripting/windows-powershell/wmf/setup/install-configure
  - Older versions will require a manual install if you cannot upgrade to WMF 5.1

# Deployment
The full deployment instructions can be found [here](https://blog.lithnet.io/2019/01/lppad-1.html) and [here](https://github.com/lithnet/ad-password-protection), but the script Install-ADPasswordProtection will do all of the work for you. The script can be run manually or through an RMM.

Here’s an overview of what it does when run on a DC:
1. Download the HIBP hashes database to the DC
2. Install the password protection application
3. Create the password protection GPO
4. Send out a notification if any errors occur during execution
5. Creates scheduled task on the PDC to check password hashes against the database and if any are found to email the address provided daily at 4 AM. 

Once the script runs the server will need to be rebooted (the script will not reboot the server) for the changes to take effect. These restrictions will be enforced the next time a user tries to change their password, existing passwords will not be affected.

# Running the script

Run the command below as admin on a PowerShell terminal. This will perform the actual install. Make sure to replace the parameters with the correct URL and emails.

The script will run a password audit on the primary domain controller once it installs successfully.

You will need to run this on all domain controllers.

First you will need to enable TLS1.2 inside of Powershell and put the AV/EDR into a monitor/read only mode temporarily

````powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
````
Then you can run the following command to start the setup process. 

````powershell
Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Install-ADPasswordProtection.ps1'); Install-ADPasswordProtection -StoreFilesInDBFormatLink '<zipFileURL>' -NotificationEmail '<email>' -SMTPRelay '<smtpServer>' -FromEmail '<fromEmail>'
````

# Parameters

  NotificationEmail is the email you want to receive the error logs as well as the password audit notification.

    -NotificationEmail 'cwilliams@compassmsp.com'

  SMTPRelay is the smtp relay address you wish to use to send emails.

    -SMTPRelay 'compassmsp-com.mail.protection.outlook.com'

  FromEmail is is the email you will want to send as for the error logs as well as the password audit notification. 

    -FromEmail 'cwilliams@compassmsp.com'

# Updates
To update the HIBP hash database you will run the following command. This will grab the update NTLM hashes file from HBIP and update database on the DC. Make sure to replace the parameters with the correct URL and emails.

The script will run a password audit on the primary domain controller once it installs successfully.

Please make sure to run this on all domain controllers.

````powershell
Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Update-ADPasswordProtection.ps1'); Update-ADPasswordProtection
````

# Verify Install
Once the install completes you can check for 3 things:

- Add/Remove Programs
  - ![](https://i.imgur.com/KcobD6H.png)
- Group Policy Management (run on a PDC for it to show up)
  - ![](https://i.imgur.com/IgMRMk6.png)
- DB Files
  - Go to "C:\Program Files\Lithnet\Active Directory Password Protection\Store\v3\p" and there should be a ton of DB files. If 0000.db and FFFF.db exist then chances are all the files are there.
  - ![](https://i.imgur.com/3hJMbKy.png)

# Reporting
The script Invoke-ADPasswordAudit will go through all AD users and compare their password hashes with known compromised hashes. This will be set to run on a schedule, and it will notify us if any users currently have compromised passwords.

````powershell
Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Invoke-ADPasswordAudit.ps1'); Invoke-ADPasswordAudit -NotificationEmail '<email>' -SMTPRelay '<smtpServer>' -FromEmail '<fromEmail>'
````

# Limitations
The application will not tell users why their password was rejected, it will only tell them that it does not meet the complexity requirements. As a result, users should be made aware that this is being put in place.

![Error Message](https://i.imgur.com/a0nIGtR.png)

# Other Features
User passwords can also be rejected for the following reasons:
- They contain the user’s name
- They contain a predefined pattern

Failed requests will show up on event viewer so alerting could be setup if required.

![](https://i.imgur.com/DmwpoFn.png)

# Removal
If this application needs to be removed for any reason it can be done through add/remove programs without having to take any special steps.
