function Get-LocalUser {
    <#
       .SYNOPSIS
            Gets all the local accounts for the specified computer(s).

        .DESCRIPTION
            Gets all the local administrators for the specified computer(s).

	    .EXAMPLE
            PS C:\> Get-LocalUser
		
		    This command shows how to list all of local users on local computer.	
	    
        .EXAMPLE
            PS C:\> Get-LocalUser | Export-Csv -Path "D:\LocalUserAccountInfo.csv" -NoTypeInformation
		
		    This command will export report to csv file. If you attach the <NoTypeInformation> parameter with command, it will omits the type information 
		    from the CSV file. By default, the first line of the CSV file contains "#TYPE " followed by the fully-qualified name of the object type.
        
        .EXAMPLE
            PS C:\> Get-LocalUser -AccountName "Administrator","Guest"
		
		    This command shows how to list local Administrator and Guest account information on local computer.
        
        .EXAMPLE
            PS C:\> $cred=Get-Credential
		    PS C:\> Get-LocalUser -Credential $cred -Computername "WINSERVER" 
		
		    This command lists all of local user accounts on the WINSERVER remote computer.
    #>
	[CmdletBinding()]
	param (
		[Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullorEmpty()][Alias('cn')]
        [string[]]$ComputerName=$Env:COMPUTERNAME,
		[Parameter(Position=1,Mandatory=$false)][Alias('un')]
        [string[]]$AccountName,
		[Parameter(Position=2,Mandatory=$false)][Alias('cred')]
        [System.Management.Automation.PsCredential]$Credential
	)
	
	$accounts = @()

    $scriptBlock = {
        param(
            $Credential,
            $GroupName
        )
        $props = @{
            Class = "Win32_UserAccount"
            Namespace = "root\cimv2"
            Filter = "LocalAccount='$True'"
            ErrorAction = "Stop"
        }
        if($Credential) {
            $props["Credential"] = $Credential
		}
        Get-WmiObject @props
    }
	
    $accounts = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $Credential
		
	if($AccountName) {
		foreach($name in $AccountName) {
			$accounts | Where-Object -FilterScript {$_.Name -like "$name"}
		}
	} else {
		$accounts
	}
}
