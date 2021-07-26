# Export-ClassInstances
A script to export class instance properties, relationships and file attachment from System Center Service Manager (Work Item or Configuration item based classes).
    
This script could be usful if you need to export class instances in bulk for archival purposes, or if you need to make changes to a custom class that are not upgrade compatible for later import with Import-ClassInstances.ps1 (https://github.com/bennyguk/Import-ClassInstances).

## To use


# Add-AOVPNTunnels
A PowerShell script to deploy and manage Always On VPN Device and User tunnels using Group Policy as an alternative to Microsoft Intune.

The script creates hashes of the ProfileXML files to detect changes that need to be applied. If you need to update the VPN client properties, just copy the new ProfileXML files to your network location and the script will do the rest.

The script uses Group Policy Preferences to copy files and create a scheduled task to run the script.

## To use
There are a few prerequisites to use this script. These are:
1. Valid ProfileXML files for Device and User tunnels. I recommend testing these profiles with PowerShell on your devices locally before using this script for deployment:
   * [Microsoft's AOVPN documentation](https://docs.microsoft.com/en-us/windows-server/remote/remote-access/vpn/always-on-vpn/deploy/vpn-deploy-client-vpn-connections)
   * Example ProfileXML files can be downloaded from Richard Hicks' github page here:
     * [Device profileXML example](https://github.com/richardhicks/aovpn/blob/master/ProfileXML_Device.xml)
     * [User ProfileXML example](https://github.com/richardhicks/aovpn/blob/master/ProfileXML_User.xml)  
3. The script depends on [New-AovpnConnection.ps1](https://github.com/richardhicks/aovpn/blob/master/New-AovpnConnection.ps1) created by Richard Hicks.

3. Create a new Group Policy Object that is enabled for computer settings and is linked to OUs that contain computer objects that you wish to delpoy the VPN profiles to. You may optionally chose to also use a computer group to filter the policy so that only specific computers will receive the policy.  
4. Copy the files ([Add-OAVPNTunnels.ps1](https://github.com/bennyguk/Add-AOVPNTunnels/blob/main/Add-AOVPNTunnels.ps1), [New-AovpnConnection.ps1](https://github.com/richardhicks/aovpn/blob/master/New-AovpnConnection.ps1), [profileXML_device.xml](https://github.com/richardhicks/aovpn/blob/master/ProfileXML_Device.xml) and [profileXML_device.xml](https://github.com/richardhicks/aovpn/blob/master/ProfileXML_User.xml)) to a network location that client devices can access to copy the files locally. I have chosen to use the folder that stores that Group Policy created earlier for central mangement and fault tolerance as the files will be replicated to all domain controllers. (\\*domain.com*\\SYSVOL\\*domain.com*\\Policies\\*{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}*\\Machine\\Scripts\\*).

   To find your policy folder, use the details tab on the GPO to find the Unique ID and substitute it for *{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}* in the line above.

5. Configure a Files Preference in the new policy:
   * Computer Configuration -> Preferences -> Windows Settings -> Files. Create a new file:
     * [General tab](/images/GPPCreateFileGeneral.JPG?raw=true "GPP Files general tab"):
       * Configure the source folder for your script and ProfileXML files followed by '\\\*'. This will copy all files in the folder.
       * Specify a local destination folder. I have chosen to create a new folder under the local Windows directory using the %windir% environment variable. GPP will automaticall create the folder if it is missing.
       * Action: Replace.
     * [Common tab](/images/GPPCreateFileCommon.JPG?raw=true "GPP Files common tab")
       * Check the box 'Remove this item when it is no longer required'.

6. Configure a Scheduled Tasks Preference in the new policy:
   * Computer Configuration -> Preferences -> Control Panel Settings -> Scueduled Tasks. New Scheduled Tasks (At least Windows 7):
     * [General Tab:](/images/GPPTasksGeneral.JPG?raw=true "GPP Files general tab")
       * Action: Replace
       * Give the task a name.
       * Use the NT AUTHORITY\SYSTEM account.
       * Check the box 'Run with highest privileges'.
       * Configure for: Windows 7, Windows Server 2008 R2. (If there is a later OS, choose that instead).  
     * [Triggers:](/images/GPPTasksTriggers.JPG?raw=true "GPP Files common tab")
       * Add a new trigger to run at log on (I tried with 'at startup', but could not get it to run reliably).
       * Configure the task to run for any user.  
     * [Actions:](/images/GPPTasksActions.JPG?raw=true "GPP Files common tab")
       * Action: Start a program
       * Program/Script: PowerShell
       * Add arguments(optional): `-ExecutionPolicy Bypass -File "%windir%\AOVPN\AddAOVPNTunnels.ps1"`  
     * [Settings:](/images/GPPTasksSettings.JPG?raw=true "GPP Files common tab")
       * Tick 'Allow task to be run on demand' (for troubleshooting).  
     * [Common:](/images/GPPTasksCommon.JPG?raw=true "GPP Files common tab")
       * Tick 'Remove this item when it is no longer applied'.  

## More information
**Why Scheduled tasks and not a startup script?**  
Startup scripts require a network connection to work because the script files must be stored in the GPO folder on the domain controller. This is not always poossible on portable devices and I found it unreliable on wireless connections.

I have tested the script on Windows 10 version 1909, 2004 and 20H2.
