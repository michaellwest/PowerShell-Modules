function Compress-Alias {
    <#
        .SYNOPSIS
            Converts all Cmdlets to their first defined alias form.

        .PARAMETER Path
            The path to a file to convert.

        .PARAMETER Script
            The script text to convert.

        .EXAMPLE
            PS C:\> Compress-Alias -Path "C:\script.ps1"

        .EXAMPLE
            PS C:\> Compress-Alias -Script "Get-Process -Name notepad | Stop-Process"

            gps -Name notes | kill
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true, ParameterSetName="File")]
        [string]$Path,

        [Parameter(Position=0, Mandatory=$true, ParameterSetName="Text")]
        [string]$Script
    )
    
    if($PSCmdlet.ParameterSetName -eq "File") {
        if(Test-Path -Path $Path){
            $Script = (Get-Content -Path $Path -Delimiter ([char]0))            
        }
    }
    
    $cmdlets = Get-Command | Where-Object {$_.CommandType -eq "Cmdlet"} | Select-Object Name
    foreach($cmdlet in $cmdlets) {
        $cmd = $cmdlet.name;

        if($Script -match "\b$cmd\b") {
            $alias = @(Get-Alias | Where-Object {$_.Definition -eq $cmd});   
            $Script = $Script -replace($cmd,$alias[0])
        }
    }
    
    $Script
}
