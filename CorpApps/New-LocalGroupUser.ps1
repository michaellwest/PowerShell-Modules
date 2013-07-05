function New-LocalGroupUser { 
    param(
        [string]$Domain,
        [string]$UserName,        
        [string]$ComputerName = $env:ComputerName,
        [string]$GroupName="Administrators"
    ) 

    $group = [ADSI]"WinNT://$ComputerName/$GroupName,group"
    if($Domain) {
        $group.Add("WinNT://$ComputerName/$Domain/$UserName")
    } else {
        $group.Add("WinNT://$ComputerName/$UserName")
    }
