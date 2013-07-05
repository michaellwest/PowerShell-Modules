function Split-UserName {
    param(
        [ValidateNotNullOrEmpty()]
        [string]$UserName
    )

    $split = $UserName -split '\\'
    if($split.Length -gt 1) {
        $Domain = $split[0]
        $UserName = $split[1]
    } else {
        $UserName = $split[0]
    }

    [PSCustomObject]@{
      Domain = $Domain
	    UserName = $UserName
    }
}
