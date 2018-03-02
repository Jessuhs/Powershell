Write-Host "`This script will retrieve the MAC address of the computer name(s) given.  The machine(s) must be online for this to work"

$strQuit = "Y"

While ($strQuit -eq "Y")
{
    #String array of computer names read in from user
    [string[]]$Computers = (Read-Host "`Please enter the computer names (separate with comma)").split(",")

    #This function generates a list of Computer Names, their corresponding network adapters, and those adapters' MAC addresses,

    $report = @()
    function Run-MACQuery($Computer)
    {   
        Get-CimInstance -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer |
        Where-Object {$_.IpAddress.Count -ge 1}
    }
        
    #Step through each PC entered by user and send it to the function above for processing
    ForEach ($Computer in $Computers)
    {
        $report += Run-MACQuery($Computer)
    }
    
    #Output the table in a grid view
    $report | Select-Object @{l='ComputerName';e={$_.DNSHostName}},@{l='Description';e={$_.Description}}, @{l='MAC Address';e={$_.MACAddress}} | Out-GridView
    
    #Prompt the user to continue or break the loop
    $strQuit = Read-Host "Would you like to search again? (Y/N)"
}
