function Test-Credential {
    <#
        .SYNOPSIS
            Test the users ActiveDirectory credentials.

        .EXAMPLE
            PS C:\> Test-Credential -Credential (Get-Credential)

            True

        .EXAMPLE
            PS C:\> Test-Credential -UserName "Michael.West" -Password ******* -Domain "company"

            True
    #>
    [CmdletBinding(DefaultParameterSetName="Default")]
    [OutputType('System.Boolean')]
  param(
        [Parameter(Mandatory=$true, ParameterSetName="Default")]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory=$true, ParameterSetName="NonSecure")]
        [ValidateNotNullOrEmpty()]
        [string]$UserName, 

        [Parameter(Mandatory=$true, ParameterSetName="NonSecure")]
        [ValidateNotNullOrEmpty()]
        [string]$Password,
        
        [Parameter(Mandatory=$false, ParameterSetName="NonSecure")]
        $Domain
    )

    Add-Type -AssemblyName System.DirectoryServices.AccountManagement

    try {
        if($Credential) {
            $split = Split-UserName -UserName $Credential.UserName
            $UserName = $split.UserName
            if($split.Domain) {
                $Domain = $split.Domain
            }
            $Password = $(ConvertTo-PlainText -SecureString $Credential.Password)
        } else {
            $split = Split-UserName -UserName $UserName
            $UserName = $split.UserName
            if($split.Domain) {
                $Domain = $split.Domain
            }
        }
	    $context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, $Domain)
        
		$context.ValidateCredentials($Username, $Password)
    } catch {
        throw (New-Object System.ArgumentException -ArgumentList $_.Exception.Message, "Domain")
    }
}
