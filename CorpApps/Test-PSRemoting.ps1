function Test-PSRemoting {
    <#
        .SYNOPSIS
            Determines whether all computers have PowerShell Remoting enabled.

        .DESCRIPTION
            The Test-PSRemoting function determines whether all computers have PowerShell Remoting enabled. It returns TRUE ($true)
            for computers with remoting enabled, and FALSE ($false) for computers with remoting disabled.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ComputerName,

        [System.Management.Automation.Credential()]
        $Credential=[System.Management.Automation.PSCredential]::Empty
    )     

    try {
        $props = @{
            ComputerName = $ComputerName
            ScriptBlock = { $true }
        }
        if($Credential) {
            $props["Credential"] = $Credential
        }
        $results = Invoke-Command @props
        foreach($result in $results) {
            $result
        }
    } catch {
        Write-Verbose $_
        $false
    }
} 
