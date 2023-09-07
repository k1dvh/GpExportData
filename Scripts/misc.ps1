function New-DirStruc {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory
        )]
        [string]
        $FolderName,
        [Parameter(
            Mandatory
        )]
        [string]
        $Path
    )

    if (Test-Path "$Path\$FolderName") {
        try {
            $Folder = Get-Item "$Path\$FolderName"
        }
        catch {
            Write-Error $_.Exception.Message
        }    
    }
    else {
        try {
            $Folder = New-Item -Path $Path -Name $FolderName -ItemType Directory -Force
        }
        catch {
            Write-Error $_.Exception.Message
        }
    }

    return $Folder
}