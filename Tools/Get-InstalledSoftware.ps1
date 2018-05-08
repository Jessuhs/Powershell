<#
.SYNOPSIS
	Get-InstalledSoftware retrieves a list of installed software
.DESCRIPTION
	Get-InstalledSoftware opens up the specified (remote) registry and scours it for installed software. When found it returns a list of the software and it's version.
.PARAMETER ComputerName
	The computer from which you want to get a list of installed software. Defaults to the local host.
.EXAMPLE
	Get-InstalledSoftware DC1
	
	This will return a list of software from DC1. Like:
	Name			Version		Computer  UninstallCommand
	----			-------     --------  ----------------
	7-Zip 			9.20.00.0	DC1       MsiExec.exe /I{23170F69-40C1-2702-0920-000001000000}
	Google Chrome	65.119.95	DC1       MsiExec.exe /X{6B50D4E7-A873-3102-A1F9-CD5B17976208}
	Opera			12.16		DC1		  "C:\Program Files (x86)\Opera\Opera.exe" /uninstall
.EXAMPLE
	Import-Module ActiveDirectory
	Get-ADComputer -filter 'name -like "DC*"' | Get-InstalledSoftware
	
	This will get a list of installed software on every AD computer that matches the AD filter (So all computers with names starting with DC)
.INPUTS
	[string[]]Computername
.OUTPUTS
	PSObject with properties: Name,Version,Computer,UninstallCommand
.NOTES
	Author: Anthony Howell
	
	To add directories, add to the LMkeys (LocalMachine)
.LINK
	[Microsoft.Win32.RegistryHive]
	[Microsoft.Win32.RegistryKey]
#>
Function Get-InstalledSoftware
{
	Param
	(
		[Alias('Computer','ComputerName','HostName')]
		[Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
		[string[]]$Name = $env:COMPUTERNAME
	)
	Begin
	{
		$LMkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
		$LMtype = [Microsoft.Win32.RegistryHive]::LocalMachine
		$CUkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
		$CUtype = [Microsoft.Win32.RegistryHive]::CurrentUser
		
	}
	Process
	{
		ForEach($Computer in $Name)
		{
			$MasterKeys = @()
			If(!(Test-Connection -ComputerName $Computer -count 1 -quiet))
			{
				Write-Error -Message "Unable to contact $Computer. Please verify its network connectivity and try again." -Category ObjectNotFound -TargetObject $Computer
				Break
			}
			$CURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($CUtype,$computer)
			$LMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($LMtype,$computer)
			ForEach($Key in $LMkeys)
			{
				$RegKey = $LMRegKey.OpenSubkey($key)
				If($RegKey -ne $null)
				{
					ForEach($subName in $RegKey.getsubkeynames())
					{
						foreach($sub in $RegKey.opensubkey($subName))
						{
							$MasterKeys += (New-Object PSObject -Property @{
							"ComputerName" = $Computer
							"Name" = $sub.getvalue("displayname")
							"SystemComponent" = $sub.getvalue("systemcomponent")
							"ParentKeyName" = $sub.getvalue("parentkeyname")
							"Version" = $sub.getvalue("DisplayVersion")
							"UninstallCommand" = $sub.getvalue("UninstallString")
							})
						}
					}
				}
			}
			ForEach($Key in $CUKeys)
			{
				$RegKey = $CURegKey.OpenSubkey($Key)
				If($RegKey -ne $null)
				{
					ForEach($subName in $RegKey.getsubkeynames())
					{
						foreach($sub in $RegKey.opensubkey($subName))
						{
							$MasterKeys += (New-Object PSObject -Property @{
							"ComputerName" = $Computer
							"Name" = $sub.getvalue("displayname")
							"SystemComponent" = $sub.getvalue("systemcomponent")
							"ParentKeyName" = $sub.getvalue("parentkeyname")
							"Version" = $sub.getvalue("DisplayVersion")
							"UninstallCommand" = $sub.getvalue("UninstallString")
							})
						}
					}
				}
			}
			$MasterKeys = ($MasterKeys | Where {$_.Name -ne $Null -AND $_.SystemComponent -ne "1" -AND $_.ParentKeyName -eq $Null} | select Name,Version,ComputerName,UninstallCommand | sort Name)
			$MasterKeys
		}
	}
	End
	{
		
	}
}
