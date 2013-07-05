function New-LocalUser { 
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
        [string]$UserName, 
        
        [Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)] 
        [string]$Password, 
        
        [string]$ComputerName = $env:ComputerName, 
        
        [string]$Description = "Created by PowerShell"
    ) 

    $computer = [ADSI]"WinNT://$ComputerName" 
    $user = $computer.Create("User", $UserName) 
    $user.SetPassword($Password)    
    $user.Put("Description", $Description)
    
    $user.SetInfo() 
}
