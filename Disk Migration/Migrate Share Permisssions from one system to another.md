# Migrate Share Permissions from one system to another (Using same disk image in vSphere, or local hard disk, SSD)

## Migrate the share permissions:
On the OLD server > Start > Run > Regedit Run as Administrator - > Hit Enter

## Navigate  to
HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Lanmanserver\Shares -> Right click the entire “Shares” Key and export ## it to a file.

## Next
Save it somewhere on the new system, attach the drive in vSphere, Open -> Disk Management, find the newly attached disk.

## Example: 
### Disk 0 is usually the OS, right click on the Disk Number, and click Online, you should now see the drive showing up in disk management.

You will need to assign it a drive letter, make sure you make not of the drive letter from the old system, you want to use that one.

F: should be showing Online and a Healthy (Basic Partition)

Navigate to your saved reg-key backup, double click on it and then hit OK, I suggest rebooting, login and have a user test their access.
