function Remove-LocalUser {
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
        [string]$UserName, 
        
        [string]$ComputerName = $env:ComputerName 
    )
     
    $user = [ADSI]"WinNT://$ComputerName" 
    $user.Delete("User",$UserName) 
}
