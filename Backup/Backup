###############################################################
# Run with elevated rights. Make sure all programs are closed.#
#       Create a Backups folder and share it out first.       #
###############################################################

#Changelog
#11/16/17 - added Excel macros
#12/28/17 - added reg export for pinned taskbar items
#01/25/18 - added Sticky Notes backup


#Begin Script If/Then Statement#
Write-Host ""
Write-Host ""
$FileExists = Read-Host "Do you already have a backup folder? (Y/N)"
If ($FileExists -eq "n")
{
    #Creating the Backup folder path
    Write-Host ""
    Write-Host ""
    $BackupPath = Read-Host "Enter the path to your Backups folder (Example - K: or \\servername01)" 
    new-item -type directory -path $BackupPath\Backups\ | out-null
    new-item -type directory -path $BackupPath\Backups\$env:username\ | out-null
}
Else
{
    #If you already have a Backups folder
    Write-Host ""
    Write-Host ""
    $BackupPath = Read-Host "Enter the Path of the Backups folder. (Path to shared drive)"
    new-item -type directory -path $BackupPath\Backups\$env:username | out-null
}

Write-Host ""
Write-Host ""
$PST = Read-Host "Do you want to copy PSTs?"

#User Folder data
Write-Host ""
Write-Host ""
Write-Host "Copying Desktop..."
robocopy C:\Users\$env:username\Desktop $BackupPath\Backups\$env:username\Desktop /E /COPY:DAT /R:3 /W:1 /LOG:$BackupPath\Backups\$env:username\backup.log | out-null
Write-Host ""
Write-Host ""
Write-Host "Copying Documents..."
robocopy C:\Users\$env:username\Documents $BackupPath\Backups\$env:username\Documents /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log "C:\Users\$env:username\Documents\My Pictures" "C:\Users\$env:username\Documents\My Music" "C:\Users\$env:username\Documents\My Videos" | out-null
Write-Host ""
Write-Host ""
Write-Host "Copying Favorites..."
robocopy C:\Users\$env:username\Favorites $BackupPath\Backups\$env:username\Favorites /E /COPY:DAT /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null

#Sticky Notes
New-Item -Path $BackupPath\Backups\$env:username\SNotes -ItemType directory | Out-Null
$StickyNotes = Test-Path -Path "C:\Users\$env:username\AppData\Roaming\Microsoft\Sticky Notes"
If ($StickyNotes -eq $true)
{
    Write-Host ""
    Write-Host ""
    Write-Host "Sticky Notes found!  Copying sticky notes..."
    robocopy C:\Users\$env:username\AppData\Roaming\Microsoft\Sticky Notes\StickyNotes.snt $BackupPath\Backups\$env:username\SNotes /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null
}
$StickyNotes10 = Test-Path -Path "C:\Users\$env:username\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe"
If ($StickyNotes10 -eq $true)
{
    Write-Host ""
    Write-Host ""
    Write-Host "Sticky Notes found!  Copying sticky notes..."
    robocopy C:\Users\$env:username\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState $BackupPath\Backups\$env:username\SNotes /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null
}

#Office Data
Write-Host ""
Write-Host ""
Write-Host "Copying Office templates..."
robocopy C:\Users\$env:username\AppData\Roaming\Microsoft\Templates "$BackupPath\Backups\$env:username\Office Templates" /E /COPY:DAT /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null

#Excel Queries
Write-Host ""
Write-Host ""
Write-Host "Copying Excel queries..."
robocopy C:\Users\$env:username\AppData\Roaming\Microsoft\Queries "$BackupPath\Backups\$env:username\Queries" /E /COPY:DAT /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null

#Excel Macros
Write-Host ""
Write-Host ""
Write-Host "Copying Excel macros..."
robocopy C:\Users\$env:username\AppData\Roaming\Microsoft\Excel\XLSTART "$BackupPath\Backups\$env:username\Macros" /E /COPY:DAT /R:3 /W:1 /LOG+$BackupPath\Backups\$env:username\backup.log | out-null

if (($PST -eq "y" ) -or ($PST -eq "Y"))
{
    #PSTs
    Write-Host ""
    Write-Host ""
    Write-Host "Copying PSTs..."
    robocopy C:\users\$env:username\AppData\Local\Microsoft\Outlook "$BackupPath\Backups\$env:username\Outlook pst" *.pst /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log "C:\users\$env:username\AppData\Local\Microsoft\Outlook\RoamCache" "C:\users\$env:username\AppData\Local\Microsoft\Outlook\Offline Address Books" | out-null
}

#Outlook Signatures
Write-Host ""
Write-Host ""
Write-Host "Copying Signatures..."
robocopy C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures "$BackupPath\Backups\$env:username\Outlook Signatures" /E /COPY:DAT /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null

#Other App Data
Write-Host ""
Write-Host ""
Write-Host "Copying Chrome Data..."
robocopy "C:\Users\$env:username\AppData\Local\Google\Chrome\User Data\Default" "$BackupPath\Backups\$env:username\Chrome Favorites" Bookmarks.* /S /COPY:DAT /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null

#Windows Data
Write-Host ""
Write-Host ""
Write-Host "Copying Windows Data"
robocopy "C:\Users\$env:username\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch" "$BackupPath\Backups\$env:username\Quick Launch" /E /COPY:DAT /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null
robocopy "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" "$BackupPath\Backups\$env:username\Startup" /E /COPY:DAT /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\backup.log | out-null
Reg Export "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" "$BackupPath\Backups\$env:username\Taskbar.reg" /y


#Registry Keys
Reg Export "HKCU\Software\ODBC" "$BackupPath\Backups\$env:username\ODBC\USERDSN.reg" /y
Reg Export "HKLM\Software\Wow6432Node\ODBC\ODBC.INI" "$BackupPath\Backups\$env:username\ODBC\SYSTEMDSN.reg" /y

#OneNote Notebook Connections
$testMail10 = Test-Path "C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"
$testMail13 = Test-Path "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE"
$testMail16 = Test-Path "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"

if ($testMail10 -eq $true)
{
    Reg Export "HKCU\Software\Microsoft\Office\14.0\OneNote\OpenNotebooks" "$BackupPath\Backups\$env:username\OpenNotebooks.reg" /y
}
elseif ($testMail13 -eq $true)
{
    Reg Export "HKCU\Software\Microsoft\Office\15.0\OneNote\OpenNotebooks" "$BackupPath\Backups\$env:username\OpenNotebooks.reg" /y
}
elseif ($testMail16 -eq $true)
{
    Reg Export "HKCU\Software\Microsoft\Office\16.0\OneNote\OpenNotebooks" "$BackupPath\Backups\$env:username\OpenNotebooks.reg" /y
}

#Printers
Write-Host ""
Write-Host ""
Write-Host "Copying Printers"
Get-WMIObject -class Win32_Printer | Select-Object Name | Export-CSV $BackupPath\Backups\$env:username\printers.csv -Force

#Default Printer
Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true" | Select-Object Name | Export-CSV $BackupPath\Backups\$env:username\default.csv -force

#Get mapped drives
Write-Host ""
Write-Host ""
Write-Host "Getting mapped drives"
Get-PSDrive -PSProvider FileSystem | Select-Object -property Root,DisplayRoot | Export-CSV $BackupPath\Backups\$env:username\drives.csv -force

Write-Host ""
Write-Host ""
Read-Host -Prompt "Script complete.  Press Enter to exit"
