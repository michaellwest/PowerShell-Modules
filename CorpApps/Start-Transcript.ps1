 <# 
.ForwardHelpTargetName Start-Transcript
.ForwardHelpCategory Cmdlet
 #>
function Start-Transcript {
    [CmdletBinding(DefaultParameterSetName='ByPath', SupportsShouldProcess=$true, ConfirmImpact='Medium', HelpUri='http://go.microsoft.com/fwlink/?LinkID=113408')]
    [OutputType('System.String')]
    param (
        [Parameter(Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('LiteralPath','PSPath')]
        [string]$Path='',

        [switch]$Append,

        [switch]$Force,

        [Alias('NoOverwrite')]
        [switch]$NoClobber
    )

    if($global:isTranscribing) {
        throw 'Start-Transcript : Transcription has already been started. Use the Stop-Transcript command to stop transcription.'
    }

    $timestamp = Get-Date -Format yyyyMMddHHmmss
    $username = "$env:USERDOMAIN\$env:USERNAME"
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $operatingsystem = '{0} ({1} {2} {3})' -f $os.CSName,$os.Caption, $os.Version, $os.CSDVersion

    if (!$Path) {
        $Folder = Split-Path -Path (Split-Path -Path $profile) 
        $File = "PowerShell_transcript.$timestamp.txt"
        $Path = Join-Path -Path $Folder -ChildPath $File
    }
    if ($NoClobber) {
        if (Test-Path -Path $Path) {
            throw "Start-Transcript: File $Path already exists and NoClobber was specified."
        }
    }

    $header = @"
**********************
Windows PowerShell transcript start
Start time: $timestamp
Username  : $username
Machine    : $operatingsystem
**********************
Transcript started, output file is $path
"@
    $header | Out-File -FilePath $Path -Append:$Append -Force:$Force
    Remove-Variable TranscriptContent -Scope Script -ErrorAction SilentlyContinue
    "Transcript started, output file is $path" 
    $global:isTranscribing=$true
    $global:TranscriptPath=$Path

    <#
    .ForwardHelpTargetName Out-Default
    .ForwardHelpCategory Cmdlet
    #>
    function global:Out-Default {
        [CmdletBinding(HelpUri='http://go.microsoft.com/fwlink/?LinkID=113362', RemotingCapability='None')]
        [OutputType()]
        param (
            [Parameter(ValueFromPipeline=$true)]
            [psobject]
            $InputObject
        )

        begin {    
            try {
                $originalCommand = $null
                $script:TranscriptContent = $null 
                $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Out-Default', [System.Management.Automation.CommandTypes]::Cmdlet)

                if ($global:isTranscribing) {
                    $scriptCmd = { Out-StringEx | & $wrappedCmd }
                } else {
                    $scriptCmd = {& $wrappedCmd }
                }

                $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
                $steppablePipeline.Begin($PSCmdlet)
            } catch {
                throw
            }
        }

        process {
            try {
                $steppablePipeline.Process($_)

                if ($global:isTranscribing) {
                    if ($originalCommand -eq $null) {
                        $originalCommand = '{0}{1}' -f (prompt), (Get-Variable -Name MyInvocation -Scope 1).Value.MyCommand.Definition
                    }
                }
            } catch {
                throw
            }
        }

        end {
            try {
                $steppablePipeline.End()
            } catch {
                throw
            } finally {
                if ($global:isTranscribing) {
                    $originalCommand | Out-File -Append -FilePath $global:TranscriptPath
                    $script:TranscriptContent | Out-File -Append -FilePath $global:TranscriptPath
                    $script:TranscriptContent = $null
                }
            }
        }
    }
}
