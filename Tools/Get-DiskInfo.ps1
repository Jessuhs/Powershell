function Get-DiskInfo
{
    [alias("disks")]
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
        $Disks = Get-CimInstance -ComputerName $ComputerName -Class CIM_DiskDrive
        $Disks | Select-Object @{l='Computer Name';e={$_.PSComputerName}},@{l='Model';e={$_.Model}}, @{l='Partitions';e={$_.Partitions}}, @{l='Size(GB)';e={$_.Size / 1GB -as [int]}}
    }
    
    END{}
}
