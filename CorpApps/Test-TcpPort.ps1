function Test-TcpPort {
    <#
        .SYNOPSIS
            Determine if computers have the specified ports open.
        
        .EXAMPLE
            PS C:\> Test-TcpPort -ComputerName web01,sql01,dc01 -Port 5985,5986,80,8080,443

        .NOTE
            Example function from PowerShell Deep Dives 2013.
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias("CN","Server","__Server","IPAddress")]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        
        [int[]]$Port = 23,
        
        [int]$Timeout = 5000
    )

    Process {
        foreach ($computer in $ComputerName) {
            foreach ($p in $port) {
                Write-Verbose ("Checking port {0} on {1}" -f $computer, $p)

                $tcpClient = New-Object System.Net.Sockets.TCPClient
                $async = $tcpClient.BeginConnect($computer, $p, $null, $null)
                $wait = $async.AsyncWaitHandle.WaitOne($TimeOut, $false)
                if(-not $Wait) {
                    [PSCustomObject]@{
                        Computername = $ComputerName
                        Port = $P
                        State = 'Closed'
                        Notes = 'Connection timed out'
                    }
                } else {
                    try {
                        $tcpClient.EndConnect($async)
                        [PSCustomObject]@{
                            Computername = $computer
                            Port = $p
                            State = 'Open'
                            Notes = $null
                        }
                    } catch {
                        [PSCustomObject]@{
                            Computername = $computer
                            Port = $p
                            State = 'Closed'
                            Notes = ("{0}" -f $_.Exception.Message)
                        }
                    }
                }
            }
        }
    }
}
