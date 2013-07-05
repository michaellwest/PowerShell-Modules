function Set-LocalUser {
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
        [string]$UserName, 
        
        [Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='PasswordReset')] 
        [string]$Password, 
        
        [Parameter(ParameterSetName='PasswordReset')] 
        [Parameter(ParameterSetName='EnableUser')] 
        [switch]$Enable, 
        
        [Parameter(ParameterSetName='DisableUser')] 
        [switch]$Disable, 
        
        [string]$ComputerName = $env:ComputerName, 
        
        [string]$Description = "Modified via PowerShell",

        [ValidateSet("Enabled","Disabled")]
        [string]$RequirePasswordChange
    )

    $user = [ADSI]"WinNT://$ComputerName/$UserName,User"

    if($user.Path) {
        if($Disable) {
            $user.UserFlags = 2  # ADS_USER_FLAG_ENUM enumeration value from SDK
        } else {
            $user.SetPassword($Password)

            if($Enable) { 
                $user.UserFlags = 512 # ADS_USER_FLAG_ENUM enumeration value from SDK 
            }
        }

        switch($RequirePasswordChange) {
            "Enabled" {
                $user.PasswordExpired = 1
            }
            "Disabled" {
                $user.PasswordExpired = 0
            }
        }

        $user.Description = $Description 
        $user.SetInfo()
    } else {
        Write-Warning -Message "The user $($UserName) does not exist."
    }
}
