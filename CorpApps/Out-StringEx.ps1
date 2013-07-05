function Out-StringEx {
    param (
        [Parameter(ValueFromPipeline=$true)]
        $InputObject
    )

    begin {
        $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Out-String', 'Cmdlet')
        $scriptCmd = {&  $wrappedCmd -OutVariable script:TranscriptContent -Stream | Out-Null }
        $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
        $steppablePipeline.Begin($PSCmdlet)
    }

    process {
        $steppablePipeline.Process($_)
        $_
    }

    end {
        $steppablePipeline.End()
    }   
}
