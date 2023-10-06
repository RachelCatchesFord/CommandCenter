# CommandCenter
Powershell UI to run certain things as Local Admin (LA) accounts without having to repeatedly type in password.  
When running this script, do not run as admin. It will prompt for admin credentials for the certain functions within that require it.  

# Contents:
  *.ico file* - Icon for the UI Title Bar. Name and path must be changed on line 10 of CommandCenter.ps1  
  *.jpg file* - background for UI. Name and path must be changed on line 12 of CommandCenter.ps1  
  *RemoteRegistry.ps1* - workaround to start remote registry on a remote computer, not included in this UI  
  *Start Command Center.vbs* - VBS Script to launch CommandCenter.ps1 using Powershell. This is done so that an Icon can be created on the desktop.  
  *Command Center shortcut* - Shortcut to the VBS file above. Location and icon must be changed.         
  ## CMClient Folder  
      - Copy in the Client Install files for MECM with the latest patch (MP Install Path - usually MPServer\C$\Program Files\Microsoft Configuration Manager\Client)   
      - ccmclean.exe - Microsoft's exe to clean / uninstall the MECM client. 30% of the time it works 100% of the time. It is a Microsoft product, afterall... :)  
      - WMIFix01.bat - Additional commands to fix WMI if it becomes corrupt upon MECM client uninstall. This is included in `Uninstall-CMClient` and shouldn't need to be ran separately  
      - EMTCustomModules.ps1 - Custom modules that get copied to the client machine and imported into PS. Some of the commands need to be run on the local computer.  

# Notes
There are comments that start with #!!!!!!!!# throughout the CommandCenter.ps1 file that will need to be updated for your environment.  
There are functions listed throughout the CommandCenter.ps1 file that are not used. Some of them are called at the end of the file and will result in a failure, but will continue on error and the script / UI will still function. 

