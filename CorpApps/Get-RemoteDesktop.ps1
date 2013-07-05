function Get-RemoteDesktop {
    [CmdletBinding()]
    param(
        [string[]]$ComputerName=$env:COMPUTERNAME,

        [System.Management.Automation.Credential()]
        $Credential=[System.Management.Automation.PSCredential]::Empty
    )

    $props = @{
        Namespace = "root/CIMV2/TerminalServices"
    }

    $cimProps = @{
        ComputerName = $ComputerName
    }

    if($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
        $cimProps["Credential"] = $Credential
    }

    $cimSession = New-CimSession @cimProps

    $tsSetting = Get-CimInstance @props -ClassName Win32_TerminalServiceSetting
    $authenticationSetting = Get-CimInstance @props -ClassName Win32_TSGeneralSetting -Filter 'TerminalName = "RDP-Tcp"'
    
    $settings = [PSCustomObject][Ordered]@{
        Enabled = [bool]($tsSetting.AllowTSConnections)
        RequireNLA = [bool]($authenticationSetting.UserAuthenticationRequired)
    }
    
    $settings
}
