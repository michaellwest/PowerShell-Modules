function Get-ParameterOption {
    param(
        $Command,
        $Parameter
    )
 
    $parameters = Get-Command -Name $Command | Select-Object -ExpandProperty Parameters
     
    $type = $parameters[$Parameter].ParameterType
    if($type.IsEnum) {
        [System.Enum]::GetNames($type)
    } else {
        $parameters[$Parameter].Attributes.ValidValues
    }
}
