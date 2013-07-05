function ConvertTo-PlainText {
    <#
        .SYNOPSIS
            Converts the System.Security.SecureString to plain text.

        .PARAMETER SecureString
            The encrypted string to convert.

        .EXAMPLE
            PS C:\> ConvertTo-PlainText -SecureString (Get-Credential).Password

        .NOTES
            Author:
            Michael West
    #>
    param(
        [System.Security.SecureString]$SecureString
    )

    [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString));
}
