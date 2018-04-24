function Get-User
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

    PROCESS
    {
        $User = (Get-WmiObject -ComputerName $ComputerName Win32_ComputerSystem).Username
        echo $User
        $Locked = Get-Process -ComputerName $ComputerName -Name *logonui

        if ($Locked.ProcessName -eq "logonui")
        {
            Write-Host "Computer is locked!"
        }
        Else
        {
            Write-Host "Computer is in use!"
        }
    }

    END {}
}
