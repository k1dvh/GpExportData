Get-ChildItem -Path "$PSScriptRoot\Scripts\*" -Include *.ps1 | ForEach-Object {
    . $_.FullName
}

$global:Config = @{}
$global:Data = @{}

$Data["Custom"] = @{}


$Config["Result"] = New-DirStruc -FolderName "Results" -Path $PSScriptRoot
$Config["Links"] = New-DirStruc -FolderName "Links" -Path $Config["Result"].FullName
$Config["Delegations"] = New-DirStruc -FolderName "Delegations" -Path $Config["Result"].FullName
$Config["Settings"] = New-DirStruc -FolderName "Settings" -Path $Config["Result"].FullName

$Data["GPOs"] = Get-GPO -All
$Data["Domain"] = Get-ADDomain

foreach ($Gpo in $Data.Gpos) {

    [xml] $Report = Get-GPOReport -Name $Gpo.DisplayName -ReportType Xml

    $CustomGPO = [PSCustomObject]@{
        Data   = $Gpo
        Report = $Report
    }

    Get-GpoGlobalInformations -Gpo $CustomGPO

    Get-GpoLinks -Gpo $CustomGPO

    $Data.Custom[$($Gpo.DisplayName)] += $CustomGPO
    
}