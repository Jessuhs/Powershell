function Get-Uptime
{
    [Alias("uptime")]
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
        foreach ($Name in $ComputerName)
        {
            $Uptime = Get-CimInstance -Class Win32_PerfFormattedData_PerfOS_System -ComputerName $Name
            $Span = New-TimeSpan -Seconds $Uptime.SystemUpTime
            New-Object -TypeName PSObject -Property @{
                ComputerName = $Name
                Days = $Span.Days
                Hours = $Span.Hours
                Minutes = $Span.Minutes
                Seconds = $Span.Seconds
            } | Format-Table -Property ComputerName,Days,Hours,Minutes,Seconds
        }
    }

    END {}
}
