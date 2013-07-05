function Get-ScriptDirectory {
    <#
        .SYNOPSIS
            Returns the path of the current executing script.
    #>
    Split-Path $script:MyInvocation.MyCommand.Path
}
