Get-ChildItem -Path "$PSScriptRoot\Scripts\*" -Include *.ps1 | ForEach-Object {
    . $_.FullName
}

$global:Config = @{}
$global:Data = @{}
$FormattedArray = @()

$Data["Custom"] = @()


$Config["Results"] = New-DirStruc -FolderName "Results" -Path $PSScriptRoot
$Config["Links"] = New-DirStruc -FolderName "Links" -Path $Config["Results"].FullName
$Config["Delegations"] = New-DirStruc -FolderName "Delegations" -Path $Config["Results"].FullName
$Config["Settings"] = New-DirStruc -FolderName "Settings" -Path $Config["Results"].FullName

$Data["GPOs"] = Get-GPO "Default Domain Policy"
$Data["Domain"] = Get-ADDomain

foreach ($Gpo in $Data.Gpos) {

    [xml] $Report = Get-GPOReport -Name $Gpo.DisplayName -ReportType Xml

    $CustomGPO = [PSCustomObject]@{
        Data   = $Gpo
        Report = $Report
    }

    Get-GpoGlobalInformations -Gpo $CustomGPO
    Get-GpoLinks -Gpo $CustomGPO
    Get-GpoDelegations -Gpo $CustomGPO
    Get-GpoSettings -Gpo $CustomGPO

    $Data.Custom += $CustomGPO
    
}

$Data.Custom | ForEach-Object {
    $FormattedArray += [PSCustomObject]@{
        Action           = "Todo"
        Name             = $_.Data.DisplayName
        GPOStatus        = $_.Data.GpoStatus
        NoSettings       = $_.NoSettings
        UserSettings     = $_.UserSettings
        ComputerSettings = $_.ComputerSettings
        Linked           = $_.Linked
        Links            = $_.ConcatenatedLinks
        GPOHasTargets    = $_.GpoHasTarget
        Targets          = $_.Targets
        WmiFilters       = $_.Data.WmiFilter
    }
} 

$FormattedArray | Export-Csv -Path "$($Config.Results.FullName)\export.csv" -Force