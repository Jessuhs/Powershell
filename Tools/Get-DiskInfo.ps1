function Get-DiskInfo
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
        $Disks = Get-CimInstance -ComputerName $ComputerName -Class CIM_DiskDrive
        $Disks | Select-Object @{l='Computer Name';e={$_.PSComputerName}},@{l='Model';e={$_.Model}}, @{l='Partitions';e={$_.Partitions}}, @{l='Size(GB)';e={$_.Size / 1GB -as [int]}}
    }
}
