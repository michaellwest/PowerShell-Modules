function Set-RemoteDesktop {
     <#
        .SYNOPSIS
            Configures Remote Desktop settings and requirements on a local or remote computer.

        .DESCRIPTION
            The Set-RemoteDesktop function is used to configure the requirements for Remote Desktop connections on the local 
            or a remote computer. Remote Desktop access can be enabled or disabled and, in Windows Vista or higher, the Windows 
            Firewall configured accordingly. Network Level Authentication requirements can also be configured for Windows Vista and higher, 
            using the RequireNLA parameter.
        
        .PARAMETER Enable
            Enable remote desktop access on the specified computer. This parameter cannot be used when the Disable parameter has been specified.

        .PARAMETER Disable
            Disable remote desktop access on the specified computer. This parameter cannot be used when the Enable parameter has been specified.

        .PARAMETER RequireNLA
            Require Network Level Authentication from clients attempting to access the computer. This parameter can only be used against Windows Vista and Windows 2008 operating systems or higher.

        .PARAMATER ConfigureFirewall
            Configure the firewall in line with whether remote desktop access is enabled on the computer. This parameter can only be used against Windows Vista and Windows 2008 operating systems or higher.

        .PARAMETER ComputerName
            The computer against which to run the function. By default this parameter will be populated with the name of the local computer.

        .PARAMETER Credential
            The credentials under which to run the function. By default this function will run as the current user. Using this parameter and the Get-Credential function you can specify an alternate set of credentials under which to execute this command.
        
        .EXAMPLE
            This command will enable Remote Desktop logins on the local computer:

            PS C:\> Set-RemoteDesktop -Enable

        .EXAMPLE
            This command will enable RemoteDesktop logins on the remote computers web01 and sql01. It also requires that clients connecting
            to the remote desktop of web01 and sql01 use Network Level Authentication:

            PS C:\> Set-RemoteDesktop -Enable -RequireNLA -Computer web01,sql01

        .EXAMPLE
            This command will enable RemoteDesktop logins on the local computer and the Windows Firewall configured to allow remote desktop connections:

            PS C:\> Set-RemoteDesktop -Enable -ConfigureFirewall

        .LINK
            http://msdn.microsoft.com/en-us/library/windows/desktop/aa383644(v=vs.85).aspx

        .LINK
            http://msdn.microsoft.com/en-us/library/windows/desktop/aa383441(v=vs.85).aspx
    #>    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ParameterSetName='Enable')]
        [switch]$Enable,

        [Parameter(Mandatory=$false,ParameterSetName='Enable')]
        [switch]$RequireNLA,

        [Parameter(Mandatory=$true,ParameterSetName='Disable')]
        [switch]$Disable,

        [switch]$ConfigureFirewall,

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
    
    $tsProps = @{
        AllowTSConnections=[int]($Enable.IsPresent)
    }
    $tsProps['ModifyFirewallException'] = [int]($ConfigureFirewall.IsPresent)
    $tsSetting | Invoke-CimMethod -MethodName SetAllowTsConnections -Arguments $tsProps | Out-Null
    
    if($Enable) {
        $authenticationSetting | Invoke-CimMethod -MethodName SetUserAuthenticationRequired -Arguments @{UserAuthenticationRequired=[int]($RequireNLA.IsPresent)} | Out-Null
    }
}
