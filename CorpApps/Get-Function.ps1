function Get-Function {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string[]]$Name
    )

    process {
        $builder = New-Object System.Text.StringBuilder
        foreach($command in (Get-Command -Name $Name -CommandType Function)) {
            $builder.Append("function $($command.Name) {") | Out-Null
            $builder.Append($command.Definition) | Out-Null
            $builder.Append("} ") | Out-Null
        }
        $builder.ToString()
    }
}
