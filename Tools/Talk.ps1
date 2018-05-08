function Talk
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [Alias('hostname')]
        [Alias('cn')]
        [string]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory=$true)]
        [string]$Phrase
    )

    PROCESS
    {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock{
            Add-Type -AssemblyName System.Speech
            $Speech = New-Object System.Speech.Synthesis.SpeechSynthesizer
            $Speech.Speak("$($args[0])")
        } -ArgumentList $Phrase
    }
}
