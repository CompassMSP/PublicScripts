# Requirements

- The application must be installed on every DC at a site
- DCs must have at least 20GB free during setup, and 8GB free once the setup is complete. This space is required for the HIBP database
- DCs must be 2016+ (or have WMF 5.1 installed if 2012 R2) for the automated script to run
  - Older versions will require a manual install

# Deployment
The full deployment instructions can be found [here](https://blog.lithnet.io/2019/01/lppad-1.html) and [here](https://github.com/lithnet/ad-password-protection), but the script Install-ADPasswordProtection will do all of the work for you. The script can be run manually or through VSA/Automate. Additional parameters (ZIP File location, and notification email) will be required if the script is run manually.

Here’s an overview of what it does when run on a DC:
1. Download the HIBP hashes database to the DC
2. Install the password protection application
3. Create the password protection GPO
4. Send out a notification if any errors occur during execution

Once the script runs the server will need to be rebooted (the script will not reboot the server) for the changes to take effect. These restrictions will be enforced the next time a user tries to change their password, existing passwords will not be affected.

# Reporting
The script Invoke-ADPasswordAudit will go through all AD users and compare their password hashes with known compromised hashes. This will be set to run on a schedule, and it will notify us if any users currently have compromised passwords.

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
