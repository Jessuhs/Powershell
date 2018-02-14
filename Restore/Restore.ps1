#Changelog
#11/16/17 - Changed Outlook 2016 restore to reflect new GPO change.
#12/28/17 - Added reg import for taskbar pinned items
#01/25/18 - Added Sticky Notes restore

#TODO
#Outlook 2016 profile auto-create fix


$BackupPath = Read-Host "Please enter the path to your backup folder"
$OGPath = Get-Location

#Get-ChildItem -Path $BackupPath

#Test the path provided and prompt again if invalid
$testPath = Test-Path -path "$BackupPath\Backups\$env:username\"
while ($testPath -eq $false)
{
    Write-Host ""
    Write-Host ""
    $BackupPath = Read-Host "Path invalid!  Please enter a valid path"
    $testPath = Test-Path -path "$BackupPath\Backups\$env:username\"
}

Write-Host ""
Write-Host ""
Write-Host "Path valid.  Starting restore"

#User folder data
Write-Host ""
Write-Host ""
Write-Host "Restoring Documents..."
robocopy $BackupPath\Backups\$env:username\Documents C:\Users\$env:username\Documents /E /COPY:DAT /XD /R:3 /W:1 /LOG:$BackupPath\Backups\$env:username\restore.log | out-null
Write-Host ""
Write-Host ""
Write-Host "Restoring Desktop..."
robocopy $BackupPath\Backups\$env:username\Desktop C:\Users\$env:username\Desktop /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null
Write-Host ""
Write-Host ""
Write-Host "Restoring Favorites..."
robocopy $BackupPath\Backups\$env:username\Favorites C:\Users\$env:username\Favorites /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null

#Sticky Notes
$stick7 = Test-Path -Path "$BackupPath\Backups\$env:username\SNotes\StickyNotes.snt"
$is10 = Test-Path -Path "C:\Users\$env:username\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes8wekyb3d8bbwe"
$stick10 = Test-Path -Path "$BackuPath\Backups\$env:username\SNotes\plum.sqlite"
#7 to 10
If ($stick7 -eq $true -and $is10 -eq $true)
{
    Write-Host ""
    Write-Host ""
    Write-Host "Sticky Notes found!  Restoring Sticky Notes..."
    New-Item -Path C:\Users\$env:username\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\Legacy -ItemType directory
    robocopy $BackupPath\Backups\$env:username\SNotes C:\Users\$env:username\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\Legacy /E /COPY:DAT /XD /R:3 /W:1 /LOG:$BackupPath\Backups\$env:username\restore.log | out-null
    Rename-Item C:\Users\$env:username\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\Legacy\StickyNotes.snt ThresholdNotes.snt
}
#7 to 7
elseif ($stick7 -eq $true)
{
    Write-Host ""
    Write-Host ""
    Write-Host "Sticky Notes found! Restoring Sticky Notes..."
    stikynot
    Stop-Process -name StikyNot
    robocopy $BackupPath\Backups\$env:username\SNotes "C:\Users\$env:username\AppData\Roaming\Microsoft\Sticky Notes" /E /COPY:DAT /XD /R:3 /W:1 /LOG:$BackupPath\Backups\$env:username\restore.log | out-null
    stikynot
}
#10 to 10
elseif ($is10 -eq $true -and $stick10 -eq $true)
{
    Write-Host ""
    Write-Host ""
    Write-Host "Sticky Notes found!  Restoring Sticky Notes..."
    robocopy $BackupPath\Backups\$env:username\SNotes\plum.sqlite C:\Users\$env:username\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState /E /COPY:DAT /XD /R:3 /W:1 /LOG:$BackupPath\Backups\$env:username\restore.log | out-null
}

#Office Data
Write-Host ""
Write-Host ""
Write-Host "Restoring Office templates..."
robocopy "$BackupPath\Backups\$env:username\Office Templates" C:\Users\$env:username\AppData\Roaming\Microsoft\Templates /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null

#Excel Queries
Write-Host ""
Write-Host ""
Write-Host "Restoring Excel queries..."
robocopy "$BackupPath\Backups\$env:username\Queries" C:\Users\$env:username\AppData\Roaming\Microsoft\Queries /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null

#Excel Macros
Write-Host ""
Write-Host ""
Write-Host "Restoring Excel macros..."
robocopy "$BackupPath\Backups\$env:username\Macros" C:\Users\$env:username\AppData\Roaming\Microsoft\Excel\XLSTART /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null

#Other App Data
Write-Host ""
Write-Host ""
Write-Host "Restoring Chrome Data..."
robocopy "$BackupPath\Backups\$env:username\Chrome Favorites" "C:\Users\$env:username\AppData\Local\Google\Chrome\User Data\Default" Bookmarks.* /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null

#Signatures
Write-Host ""
Write-Host ""
Write-Host "Restoring Outlook Signatures..."
robocopy "$BackupPath\Backups\$env:username\Outlook Signatures" C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null

#Windows Data
Write-Host ""
Write-Host ""
Write-Host "Restoring Windows Data..."
robocopy "$BackupPath\Backups\$env:username\Quick Launch" "C:\Users\$env:username\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch" /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null
robocopy "$BackupPath\Backups\$env:username\Startup" "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" /E /COPY:DAT /XD /R:3 /W:1 /LOG+:$BackupPath\Backups\$env:username\restore.log | out-null
Reg Import "$BackupPath\Backups\$env:username\Taskbar.reg"

#ODBC
Write-Host ""
Write-Host ""
Write-Host "Restoring USER ODBC connections..."
Reg Import "$BackupPath\Backups\$env:username\ODBC\USERDSN.reg"
Write-Host ""
Write-Host ""
Write-Host "Attempting to restore SYSTEM ODBC connections..."
Write-Host "If this fails, the user does not have admin rights." -ForegroundColor Cyan 
Reg Import "$BackupPath\Backups\$env:username\ODBC\SYSTEMDSN.reg"

#Printers
Write-Host ""
Write-Host ""
Write-Host "Adding Printers..."
$PrinterTable = import-csv $BackupPath\Backups\$env:username\printers.csv
foreach ($item in $PrinterTable)
{
    if($item.Name -like "*\\*")
    {
        cd C:\Windows\System32\Printing_Admin_Scripts\en-US
        Write-Host "Adding: $item"
        cscript //b prnmngr.vbs -ac -p $item.Name
    }
    else
    {
        continue
    }
}
cd $OGPath

#Map Drives
Write-Host ""
Write-Host ""
Write-Host "Mapping Drives..."
$DriveTable = import-csv $BackupPath\Backups\$env:username\drives.csv
foreach ($item in $DriveTable)
{
    $root = $item.root.trim("\")
    $drive = $item.DisplayRoot

    net use $root $drive
}

#Add user to Remote Desktop Users group (only works if user has admin rights)
Write-Host ""
Write-Host ""
Write-Host "Adding user to Remote Desktop Users group."
Write-Host ""
Write-Host "If this fails, the user does not have admin rights." -ForegroundColor Cyan 
net localgroup "Remote Desktop Users" $env:username /add

#OneNote Notebook Connections
Reg Import "$BackupPath\Backups\$env:username\OpenNotebooks.reg"

#Outlook profile
#Find the version of Office installed
$testMail10 = Test-Path "C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"
$testMail13 = Test-Path "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE"
$testMail16 = Test-Path "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
if ($testMail10 -eq $true)
{
    Write-Host ""
    Write-Host ""
    Write-Host "Outlook 2010 found!  Creating Profile."

    #Open Word to initialize registry key
    Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.EXE" -WindowStyle Hidden
    Start-Sleep -s 10
    Get-Process -Name WINWORD | Stop-Process

    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\14.0\Outlook\AutoDiscover" -Name "ZeroConfigExchange" -Value "1" -PropertyType DWORD -Force | Out-Null

    Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"
}
elseif ($testMail13 -eq $true)
{
    Write-Host ""
    Write-Host ""
    Write-Host "Outlook 2013 found!  Creating Profile."

    #Open Word to initialize registry key
    Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE" -WindowStyle Hidden
    Start-Sleep -s 10
    Get-Process -Name WINWORD | Stop-Process

    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\AutoDiscover" -Name "ZeroConfigExchange" -Value "1" -PropertyType DWORD -Force | Out-Null

    Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE"
}
elseif ($testMail16 -eq $true)
{
    Write-Host ""
    Write-Host ""
    Write-Host "Outlook 2016 found!  Creating Profile."
    
    start-process -FilePath "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
}
else
{
    Write-Host ""
    Write-Host ""
    Write-Host "No valid version of Outlook found!  Skipping profile setup."
}

#Open Bluezone
start-process -filepath "http://stifelbluezone.stifel.com"

Write-Host ""
Write-Host ""
Read-Host -Prompt "Script complete.  Press Enter to exit"
