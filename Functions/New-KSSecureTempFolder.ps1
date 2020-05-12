#This script creates a folder that only administrators and system have access to.
#It is a best practice to wipe the folder once you are done running the script that relies on it.
$FolderDirectory = 'C:\Windows\Temp\KSSecure'

New-Item -Path $FolderDirectory -ItemType Directory -Force | Out-Null

#Remove all explicit permissions
ICACLS ("$FolderDirectory") /reset | Out-Null

#Add SYSTEM permission
ICACLS ("$FolderDirectory") /grant ("SYSTEM" + ':(OI)(CI)F') | Out-Null

#Give Administrators Full Control
ICACLS ("$FolderDirectory") /grant ("Administrators" + ':(OI)(CI)F') | Out-Null

#Disable Inheritance on the Folder. This is done last to avoid permission errors.
ICACLS ("$FolderDirectory") /inheritance:r | Out-Null