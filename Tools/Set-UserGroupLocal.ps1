function Set-UserGroupLocal
{
    [Alias("Set-UGL")]
    [CmdletBinding()]
    param
    (
        [Parameter(Position=1,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [Alias('hostname')]
        [Alias('cn')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Position=1)]
        [Alias('u')]
        [string[]]$User,

        [Parameter(Position=2)]
        [Alias('g')]
        [string[]]$Group
    )

    PROCESS
    {
        $Domain = "contoso"

        $GroupObject = [ADSI]"WinNT://$ComputerName/$Group,group"
        $UserObject = [ADSI]"WinNT://$Domain/$User,user"

        Write-Output "Adding $User to the $Group group on $ComputerName..."
        $GroupObject.Add($UserObject.Path)
    }
    
    END{}
}
