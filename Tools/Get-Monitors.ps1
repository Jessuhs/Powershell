function Get-Monitors
{
    [alias("monitors")]
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
        $Monitors = Get-WmiObject -ComputerName $ComputerName WmiMonitorID -Namespace root\wmi

        ForEach ($Monitor in $Monitors)
        { 
            $Name = ($Monitor.UserFriendlyName -notmatch 0 | ForEach{[char]$_}) -join ""
            $Serial = ($Monitor.SerialNumberID -notmatch 0 | ForEach{[char]$_}) -join ""

            $hash = [ordered]@{
                Computer = $ComputerName
                Name = $Name
                Serial = $Serial
            }

            $MonitorTable = New-Object -TypeName psobject -Property $hash
            echo $MonitorTable
        }
    }

    END{}
}
