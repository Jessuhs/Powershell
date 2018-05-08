function Get-Adapters
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [Alias('hostname')]
        [Alias('cn')]
        [string[]]$ComputerName = $env:COMPUTERNAME
    ) 


    Write-Host ""
    Write-Host ""
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-CimInstance Win32_NetworkAdapter |
        Where-Object { $_.PhysicalAdapter } |
        Select-Object MACAddress,AdapterType,DeviceID,Name,Speed
    }

    Write-Host ""
    Write-Host ""
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        $adaptername = Get-CimInstance Win32_NetworkAdapter |
        Where-Object {($_.PhysicalAdapter)} |
        Where-Object {($_.AdapterType -like "*802.3*" -and $_.Name -notlike "*wireless*" -and $_.Name -notlike "*bluetooth*")} |
        Select-Object Name


        Get-CimInstance Win32_PnPSignedDriver |
        Where-Object -FilterScript { $_.devicename -like $adaptername.Name.ToString() } |
        Select-Object DeviceName, DriverVersion, DriverDate |
        Format-Table -AutoSize
    }
}
