function Send-File {
    <#
        .SYNOPSIS
            Send a file from a local path to the specified remote location.

        .PARAMETER Source
            The path to the local file.

        .PARAMETER Destination
            The path to the remote file.

        .PARAMETER ComputerName
            The computer(s) to receive the file.

        .PARAMETER Session
            The existing session to use to transfer the file.

        .PARAMETER ChunkSize
            The number of bytes of data to transfer at a time. The default is 1MB.

        .EXAMPLE
            PS C:\> Send-File -ComputerName web01 -Source C:\temp\security.xml -Destination C:\security.xml

        .EXAMPLE
            PS C:\> Send-File -ComputerName web01 -Source C:\temp\security.xml -Destination C:\security.xml -ChunkSize 9.9MB
    #>
    [CmdletBinding(DefaultParameterSetName="NewSession")]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [Alias("Path","FilePath", "LiteralPath")]
        [ValidateScript({Test-Path -Path $_})]
        [string]$Source,

        [Parameter(Mandatory=$true)]
        [string]$Destination,

        [Parameter(Mandatory=$false, ParameterSetName="NewSession")]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName=$env:COMPUTERNAME,

        [Parameter(Mandatory=$true, ParameterSetName="ExistingSession")]
        [ValidateNotNull()]
        [System.Management.Automation.Runspaces.PSSession]$Session,

        [Parameter(Mandatory=$false, ParameterSetName="NewSession")]
        [System.Management.Automation.Credential()]
        $Credential=[System.Management.Automation.PSCredential]::Empty,

        [int]$ChunkSize=1MB
    )

    begin {
        Set-StrictMode -Version Latest

        if($ComputerName) {
            $Session = New-PSSession -ComputerName $ComputerName
        }
    }

    process {
        # Open local file
        try {
            [IO.FileStream]$filestream = [IO.File]::OpenRead($Source)
            Write-Progress "Opened local file for reading"
        } catch {
            Write-Error "Could not open local file $Source because:" $_.Exception.ToString()
            return $false
        }

        $props = @{
            Session = $Session
        }
        if($Credential) {
            $props["Credential"] = $Credential
        }

        # Open remote file
        try {
            Invoke-Command @props -ScriptBlock {
                param($remFile)
                [IO.FileStream]$filestream = [IO.File]::Open($remFile, [IO.FileMode]::Create)

                $timer = New-Object System.Timers.Timer
                $timer.Interval = 5000
                $action = {
                    try {                        
                        $filestream.Close()
                        $filestream.Dispose()
                    } catch {
                        Write-Error "Could not close local file $Source because:" $_.Exception.ToString()
                        return $false
                    }
                }
                
                Register-ObjectEvent -InputObject $timer -EventName "Elapsed" -SourceIdentifier "Send-FileTimeout" -Action $action | Out-Null
                $timer.Start()

            } -ArgumentList $Destination
            Write-Progress "Opened remote file for writing"
        } catch {
            Write-Error "Could not open remote file $Destination because:" $_.Exception.ToString()
            return $false
        }
    
        # Copy file in chunks
        [byte[]]$contentchunk = New-Object byte[] $ChunkSize
        $bytesread = 0
        $contentLength = $filestream.Length
        while (($bytesread = $filestream.Read($contentchunk, 0, $ChunkSize)) -ne 0) {
            try {
                $percent = $filestream.Position / $filestream.Length
                Write-Progress -Activity ("Copying {0}, {1:P2} complete, sending {2:N2} MB" -f $Source, $percent, ($contentLength/1MB)) -PercentComplete ($percent * 100)
                Invoke-Command @props -ScriptBlock {
                    param($data, $size)
                    $timer.Stop()
                    $filestream.Write($data, 0, $size)
                    $timer.Start()
                } -ArgumentList $contentchunk,$bytesread
            } catch {
                Write-Error "Could not copy $Source to $Destination because:" $_.Exception.ToString()
                return $false
            }
        }

        # Close remote file
        try {
            Invoke-Command @props -ScriptBlock {
                $filestream.Close()
                $filestream.Dispose()
                $timer.Stop()
                Unregister-Event -SourceIdentifier "Send-FileTimeout"
            }
        } catch {
            Write-Error "Could not close remote file $Destination because:" $_.Exception.ToString()
            return $false
        }

        # Close local file
        try {
            $filestream.Close()
            $filestream.Dispose()
            Write-Progress -Activity "Closed local file, copy complete" -Completed
        } catch {
            Write-Error "Could not close local file $Source because:" $_.Exception.ToString()
            return $false
        }
    }

    end {
        Write-Progress "File transfer complete"
    }
}
