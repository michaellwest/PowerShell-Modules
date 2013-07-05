function Expand-Archive {
    <#
        .SYNOPSIS
            Extracts the compressed file to to the specified directory.

        .PARAMETER Path
            The path to the compressed file.

        .PARAMETER OutputPath
            The directory to extract the compressed file contents.

        .EXAMPLE
            PS C:\> Expand-Archive -Path "C:\temp.zip" -OutputPath "C:\temp"

            Extracts the entire contents of the compressed file to the specified directory.
    #>
   [CmdletBinding()]
   param(
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [ValidateNotNullOrEmpty()]
        [string]$OutputPath="."
   )

   [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
   [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $OutputPath)
}
