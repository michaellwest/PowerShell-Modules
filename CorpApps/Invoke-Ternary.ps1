Set-Alias -Name "?:" -Value "Invoke-Ternary" -Description "CorpApps alias"

filter Invoke-Ternary {
    <#
        .SYNOPSIS
            Similar to the C# ? : operator e.g. name = (value != null) ? String.Empty : value

        .DESCRIPTION
            Similar to the C# ? : operator e.g. name = (value != null) ? String.Empty : value.
            The first script block is tested. If it evaluates to $true then the second scripblock
            is evaluated and its results are returned otherwise the third scriptblock is evaluated
            and its results are returned.

        .PARAMETER Condition
            The condition that determines whether the TrueBlock scriptblock is used or the FalseBlock
            is used.

        .PARAMETER TrueBlock
            This block gets evaluated and its contents are returned from the function if the Conditon
            scriptblock evaluates to $true.

        .PARAMETER FalseBlock
            This block gets evaluated and its contents are returned from the function if the Conditon
            scriptblock evaluates to $false.

        .EXAMPLE
            C:\PS> 1..10 | ?: {$_ -gt 5} {"Greater than 5";$_} {"Less than or equal to 5";$_}
            Each input number is evaluated to see if it is > 5.  If it is then "Greater than 5" is
            displayed otherwise "Less than or equal to 5" is displayed.

        .NOTES
            Aliases:  ?:
            Author:   Karl Prosser

    #>
    param([scriptblock]$Condition  = $(throw "Parameter '-condition' (position 1) is required"), 
          [scriptblock]$TrueBlock  = $(throw "Parameter '-trueBlock' (position 2) is required"), 
          [scriptblock]$FalseBlock = $(throw "Parameter '-falseBlock' (position 3) is required"))

    $module = $ExecutionContext.SessionState.Module;
    
    if (& $module.NewBoundScriptBlock($Condition)) { 
        & $module.NewBoundScriptBlock($TrueBlock)
    } 
    else { 
        & $module.NewBoundScriptBlock($FalseBlock)
    }
}
