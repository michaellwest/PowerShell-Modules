function Test-Command {
    <#
        .SYNOPSIS
            Determines if the command is available.

        .EXAMPLE
            PS C:\> Test-Command -Command Get-Process

            True

        .EXAMPLE
            PS C:\> Test-Command -Command Do-Something

            False
    #>
    param([string]$Command)

    $found = $false
    $match = [Regex]::Match($Command, "(?<Verb>.{3,11})-(?<Noun>.{3,})")
    if($match.Success) {
        if(Get-Command -Verb $match.Groups["Verb"] -Noun $match.Groups["Noun"]) {
            $found = $true
        }
    }

    $found
}
