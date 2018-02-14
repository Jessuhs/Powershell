<b>Restore Script Instructions</b>
This script is meant to be run after using the Backup script in this same repository.

1.	Run the script from the user’s profile.

2.	Enter the parent folder of your backup location next. For example, if the Backups folder sits at the root of the external E drive, the proper syntax would be <b>E:</b> then hit Enter

3.	The script will now run through and restore all the user’s data.

4.	Currently there are a few issues with the script that are being worked on:

-The Outlook profile is not automatically created when Office 2016 is present.  You will need to manually click through the profile creation.

-ODBC connections may not be remapped

-The user’s taskbar will be restored completely on the next login

-They will only be placed into the Remote Desktop User’s group if they have admin rights.
