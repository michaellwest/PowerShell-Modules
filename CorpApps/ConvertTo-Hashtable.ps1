function ConvertTo-Hashtable {
    <#
        .SYNOPSIS
            Converts the object to a System.Collections.Hashtable.

        .PARAMETER InputObject
            The object to convert.

        .PARAMETER NoEmpty
            Excludes empty values from the output.

        .EXAMPLE
            PS C:\> ConvertTo-Hashtable -InputObject ([PSCustomObject][Ordered]@{FirstName='Michael';LastName='West'})

            Name                           Value
            ----                           -----
            FirstName                      Michael
            LastName                       West

        .EXAMPLE
            PS C:\> [PSCustomObject][Ordered]@{FirstName='Michael';LastName='West'} | ConvertTo-Hashtable

            Name                           Value
            ----                           -----
            FirstName                      Michael
            LastName                       West
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true, HelpMessage="Please specify an object", ValueFromPipeline=$True)]
        [object]$InputObject,

        [switch]$NoEmpty
    )

    process {
        Write-Verbose "Converting an object of type $($_.GetType().Name)"
        $names = $InputObject | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name 
        $hash = @{}
        $names | foreach-object {
            Write-Verbose "Adding property $_"
            $hash.Add($_,$InputObject.$_)
        } 

        if ($NoEmpty) {
            Write-Verbose "Parsing out empty values"
            $defined=@{}
            foreach($key in $hash.keys) {
                if ($hash[$_]) {
                    $defined[$_,$hash[$_]]
                }
            }

            Write-Verbose "Writing the result to the pipeline"
            $defined
        } else {
            Write-Verbose "Writing the result to the pipeline"
            $hash
        }
    }
}
