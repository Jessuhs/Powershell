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
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Position=1,
                    Mandatory=$false)]
        [Alias('runas')]
        [System.Management.Automation.Credential()]$Credential =
        [System.Management.Automation.PSCredential]::Empty
    )   

    PROCESS
    {
        WMIC /NODE: $ComputerName COMPUTERSYSTEM GET USERNAME
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
