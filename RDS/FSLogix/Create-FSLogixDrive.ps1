<#
This script will initialize the FSLogix drive, create the folder, and create the share.

All data on the selected drive will be deleted

Andy Morales
#>
#Requires -Modules ActiveDirectory -Version 5 -RunAsAdministrator

$FSLDiskLabel = 'FSLogixDisks'
$AclAdGroupName = 'ACL_FSLogix_FullControl'

#region CreateDisk
$AllDiskInfo = Get-Disk | Select-Object Number,OperationalStatus,@{Name="Size";Expression={$_.size/1GB}}

$AllDiskInfo | Out-String
Write-Host 'Make sure that the disk is on its own VMware SCSI controller' -ForegroundColor Yellow
Write-Host 'All Data on this drive will be deleted' -ForegroundColor Red

$DiskNumber = Read-Host "Enter disk number to initialize and format"

$WorkingDisk = Get-Disk -Number $DiskNumber

#Set the disk to Online
if ($WorkingDisk.OperationalStatus -eq 'Offline'){
    Set-Disk -Number $DiskNumber -IsOffline $false
}

#Check to see if any partitions exist. Exit if there are active partitions
if (($WorkingDisk | Get-Partition).count -gt 0){
    Write-Output 'The selected disk already contains a partition. As a safety measure, the script only works if the disk does not contain any partitions'
    Exit
}
else{
    #Clear the disk if it was previously initialized
    if ($WorkingDisk.PartitionStyle -ne 'RAW'){
        Clear-Disk -Number $DiskNumber -RemoveData -Confirm:$false
    }

    Initialize-Disk -Number $DiskNumber -PartitionStyle GPT

    #Create a partition using the next available drive letter
    $NewPartition = New-Partition -DiskNumber $DiskNumber -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem ReFS -AllocationUnitSize 64KB -NewFileSystemLabel $FSLDiskLabel
}
#endregion CreateDisk

#Get Path to new folder
$FolderDirectory = "$($NewPartition.DriveLetter):\$($FSLDiskLabel)"

#region createACLGroup

#Create the ACL group if it does not exist in AD
try{
    Get-ADGroup $AclAdGroupName -ErrorAction Stop
}
catch{
    New-ADGroup $AclAdGroupName -GroupScope DomainLocal
    Set-ADGroup $AclAdGroupName -Replace @{info = "$FolderDirectory" } -Description "\\$($env:COMPUTERNAME)\$($FSLDiskLabel)"
    Add-ADGroupMember -Members $env:USERNAME -Identity $AclAdGroupName
}
#endregion createACLGroup

#region createFolder
New-Item -Path $FolderDirectory -ItemType Directory

#Clear all Explicit Permissions on the folder
ICACLS ("$FolderDirectory") /reset

#Add CREATOR OWNER permission
ICACLS ("$FolderDirectory") /grant ("CREATOR OWNER" + ':(OI)(CI)(IO)F')

#Add SYSTEM permission
ICACLS ("$FolderDirectory") /grant ("SYSTEM" + ':(OI)(CI)F')

#Give Domain Admins Full Control
ICACLS ("$FolderDirectory") /grant ("Domain Admins" + ':(OI)(CI)F')

#Apply Create Folder/Append Data, List Folder/Read Data, Read Attributes, Traverse Folder/Execute File, Read permissions to this folder only. Synchronize is required in order for the permissions to work
ICACLS ("$FolderDirectory") /grant ("Domain Users" + ':(AD,REA,RA,X,RC,RD,S)')

#Give ACL Group Full Control
ICACLS ("$FolderDirectory") /grant ("$AclAdGroupName" + ':(OI)(CI)F')

#Disable Inheritance on the Folder. This is done last to avoid permission errors.
ICACLS ("$FolderDirectory") /inheritance:r
#endregion CreateFolder

#region CreateSmbShare
New-SmbShare -Path $FolderDirectory -CachingMode None -Name $FSLDiskLabel -FullAccess @('Domain Users', 'Everyone') -FolderEnumerationMode Unrestricted
#endregion CreateSmbShare