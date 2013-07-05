 <# 
.ForwardHelpTargetName Stop-Transcript
.ForwardHelpCategory Cmdlet
 #>
function Stop-Transcript {
    [CmdletBinding(HelpUri='http://go.microsoft.com/fwlink/?LinkID=113415')]
    [OutputType('System.String')]
    param ()

    if ($global:isTranscribing) {
        $global:isTranscribing = $false
        Remove-Item function:Out-Default
        "Transcript stopped, output file is $global:TranscriptPath"
    } else {
        throw 'Stop-Transcript : An error occurred stopping transcription: The host is not currently transcribing.'
    }
}
