function Set-MachineAppSetting {
    <#
        .SYNOPSIS
            Sets and application setting in the .NET machine.config file.

        .DESCRIPTION
            The application setting can be set in up to four different machine.config files:

            - .NET 2.0 32-bit (switches -Clr2 -Framework)
            - .NET 2.0 64-bit (switches -Clr2 -Framework64)
            - .NET 4.0 32-bit (switches -Clr4 -Framework)
            - .NET 4.0 64-bit (switches -Clr4 -Framework64)
      
            Any combination of Framework and Clr switch can be used, but you MUST supply one of each.

        .EXAMPLE
            PS C:\> Set-MachineAppSetting -Name Environment -Value Development -Framework -Framework64 -Clr2 -Clr4

            Sets the Environment application setting in the following machine.config files:

            %SYSTEMROOT%\Microsoft.NET\Framework\v2.0.50727\CONFIG\machine.config
            %SYSTEMROOT%\Microsoft.NET\Framework64\v2.0.50727\CONFIG\machine.config
            %SYSTEMROOT%\Microsoft.NET\Framework\v4.0.30319\CONFIG\machine.config
            %SYSTEMROOT%\Microsoft.NET\Framework64\v4.0.30319\CONFIG\machine.config

        .EXAMPLE
            PS C:\> Set-MachineAppSetting -Name Environment -Value Development -Framework64 -Clr4

            Sets the Environment application setting in the following machine.config files:

            %SYSTEMROOT%\Microsoft.NET\Framework64\v4.0.30319\CONFIG\machine.config
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$Name,
        
        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory=$true)]
        
        [string]$Value,
        [switch]$Framework,
        [switch]$Framework64,
        [switch]$Clr2,
        [switch]$Clr4
    )

    $v2x86 = "$($env:SystemRoot)\Microsoft.NET\Framework\v2.0.50727\CONFIG\machine.config"
    $v2x64 = "$($env:SystemRoot)\Microsoft.NET\Framework64\v2.0.50727\CONFIG\machine.config"
    $v4x86 = "$($env:SystemRoot)\Microsoft.NET\Framework\v4.0.30319\CONFIG\machine.config"
    $v4x64 = "$($env:SystemRoot)\Microsoft.NET\Framework64\v4.0.30319\CONFIG\machine.config"

    if(-not ($Framework -or $Framework64)) {
        Write-Error "You must supply one or both of the Framework and Framework64 switches."
        return
    }
    
    if(-not ($Clr2 -or $Clr4)) {
        Write-Error "You must supply one or both of the Clr2 and Clr4 switches."
        return
    }

    $configs = @()
    if($Framework) {
        if ($Clr2) {
            $configs += @($v2x86)
        }
        if ($Clr4) {
            $configs += @($v2x64)
        }
    }
    if($Framework64) {
        if ($Clr2) {
            $configs += @($v4x86)
        }
        if ($Clr4) {
            $configs += @($v4x64)
        }
    }

    foreach ($config in $configs) {
        $xml = New-Object XML
        $xml.load($config)
        $appSettings = $xml.DocumentElement.AppSettings
        if($appSettings -eq $null) {
            $root = $xml.DocumentElement
            $appSettings = $xml.CreateElement("appSettings")
            $root.AppendChild($appSettings)
        }
    
        $found = $false
        foreach($n in $appSettings["add"]) {
            if($n.Key -eq $Name) {
                $found = $true
                $n.Value = $Value 
            }
        }
        if (-not $found) {
            $setting = $xml.CreateElement("add")
            $setting.SetAttribute("key", $Name)
            $setting.SetAttribute("value", $Value)
            [void]$appSettings.AppendChild($setting)
        }
        $xml.Save($config)
    }
}
