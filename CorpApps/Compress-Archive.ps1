function Compress-Archive {
    <#
        .SYNOPSIS
            Creates a new compressed archive containing the files in the specified destination.

        .PARAMETER Path
            The file or directory path to compress.

        .EXAMPLE
            PS C:\> Compress-Archive -Path "C:\Temp" -Name temp.zip -OutputPath "C:\"

            Compresses the entire directory and outputs to the specified directory.

        .EXAMPLE
            PS C:\> Compress-Archive -Path "C:\Temp\file.txt" -Name file.zip -OutputPath "C:\"

            Compresses the file and outputs to the specified directory.

        .NOTES
            Author:
            Michael West
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Path -Path $_})]
        [string]$Path,

        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Path -Path $_})]
        [string]$OutputPath="."
    )

    $item = Get-Item -Path $Path
    $filter = ""
    if(-not $item.PSIsContainer) {
        $Path = $item.Directory
        $filter = $item.Name
    }

    if($Name -notmatch "zip$") {
        $Name += ".zip"
    }

    [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
    $level = [System.IO.Compression.CompressionLevel]::Optimal
    [System.IO.Compression.ZipFile]::CreateFromDirectory($Path, (Join-Path -Path $OutputPath -ChildPath $Name), $level, $false)
}
