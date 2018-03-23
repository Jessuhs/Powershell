function Get-MacInfo
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [Alias('hostname')]
        [Alias('cn')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Position=1,
                    Mandatory=$false)]
        [Alias('runas')]
        [System.Management.Automation.Credential()]$Credential =
        [System.Management.Automation.PSCredential]::Empty
    )

    PROCESS
    {
        Get-CimInstance -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName |
        Where-Object {$_.IpAddress.Count -ge 1} |
        Select-Object @{l='ComputerName';e={$_.DNSHostName}},@{l='Description';e={$_.Description}}, @{l='MAC Address';e={$_.MACAddress}} 
    }
}
