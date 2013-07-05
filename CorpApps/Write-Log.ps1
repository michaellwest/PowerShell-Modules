function Write-Log {
    [Cmdletbinding()]
    param(
        [Parameter(Position=0, ValueFromPipeline=$true, ParameterSetName="Message")]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Position=0, ValueFromPipeline=$true, ParameterSetName="Exception")]
        [ValidateNotNullOrEmpty()]
        [Exception]$Exception,
        
        [Parameter(Position=1)]
        [string]$Path,

        [switch]$Clobber,

        [switch]$PassThru
    )

    if($Exception) {
        $Message = $Exception.ToString()
    }

    if (!$Path -or !(Test-Path -Path $Path)) {
        $directory = & {
            if($script:MyInvocation.MyCommand.Path) {
                Split-Path -Path $script:MyInvocation.MyCommand.Path
            } else {
                $env:SystemDrive
            }
        }
        $Path = "$(Join-Path -Path ($directory) -ChildPath PowerShellLog.txt)"
    }
    
    Write-Verbose $Message

    if ($LoggingPreference -eq "Continue") {    
        if ($LoggingFilePreference) {
            $LogFile = $LoggingFilePreference
        } else {
            $LogFile = $Path
        }
        
        Write-Output "[$(Get-Date -Format s)] $Message" | Out-File -FilePath $LogFile -Append:(!$Clobber) -Force:($Clobber)
    }

    if($PassThru) {
        $Message
    }
}
