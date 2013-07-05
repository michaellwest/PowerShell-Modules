Add-type @"
    public enum FirewallProfileType
    {
        Domain = 1,
        Private = 2,
        Public = 4,
        All = 2147483647
    } 
"@

function Get-FirewallRules {
    param(
        [FirewallProfileType]$Profile,
        [switch]$Enabled
    )
    $rules = (New-Object -ComObject HNetCfg.FwPolicy2).Rules
    if($Enabled) {
        $rules = $rules | Where-Object { $_.Enabled -eq $true }
    }
    if($Profile -ne [FirewallProfileType]::All) {
        $rules = $rules | Where-Object { $_.Profiles -bAND $Profile }
    }
    $rules | Sort-Object -Property Name
}
