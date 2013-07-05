function Get-LocalGroupUser {
  [CmdletBinding()]
	param (
		[Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullorEmpty()][Alias('cn')]
        [string[]]$ComputerName=$Env:COMPUTERNAME,
		[Parameter(Position=1,Mandatory=$false)][Alias('un')]
        [string[]]$AccountName,
		[Parameter(Position=2,Mandatory=$false)][Alias('cred')]
        [System.Management.Automation.PsCredential]$Credential,
        [ValidateNotNullOrEmpty()]
        [string]$GroupName="Administrators"
	)

	$accounts = @()

    $scriptBlock = {
        param(
            $Credential,
            $GroupName
        )
        $query="Associators of {Win32_Group.Domain='$Env:COMPUTERNAME',Name='$GroupName'} where Role=GroupComponent"
        if($Credential) {
            Get-WmiObject -Query $query -Credential $Credential -ErrorAction Stop
		} else {
            Get-WmiObject -Query $query -ErrorAction Stop
		}
    }

    $accounts = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $Credential,$GroupName
    
    if($AccountName) {
		foreach($name in $AccountName) {
			$accounts | Where-Object -FilterScript {$_.Name -like "$name"}
		}
	} else {
		$accounts
	}
}
