function Test-Administrator { 
    <# 
        .SYNOPSIS
            Tests if the user is an administrator on the local computer.

        .DESCRIPTION
            Returns true if a user is an administrator, false if the user is not an administrator.
                    
        .EXAMPLE 
            PS C:\> Test-Administrator

            True

        .EXAMPLE
            PS C:\> Test-Administrator -UserName "Michael.West"

            True
    #>    
    param(  
        $UserName = [Security.Principal.WindowsIdentity]::GetCurrent() 
    )

    $principal = New-Object Security.Principal.WindowsPrincipal $UserName
    $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator) 
}
