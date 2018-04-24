function Get-MacInfo
{
    [alias("mac")]
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

    PROCESS
    {
        Get-CimInstance -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName |
        Where-Object {$_.IpAddress.Count -ge 1} |
        Select-Object @{l='ComputerName';e={$_.DNSHostName}},@{l='Description';e={$_.Description}}, @{l='MAC Address';e={$_.MACAddress}} 
    }
    
    END{}
}
