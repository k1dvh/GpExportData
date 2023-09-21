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
    Get-GpoDelegations -Gpo $CustomGPO
    Get-GpoSettings -Gpo $CustomGPO

    $Data.Custom += $CustomGPO
}



$Data.Custom | ForEach-Object {

    if ($_.Data.GpoStatus -eq "AllsettingsDisabled" `
            -or $_.NoSettings `
            -or !$_.GpoHasTarget `
            -or !$_.Linked) {
        $Action = "To Delete"
    }
    elseif ($_.LinkingError) {
        $Action = "To Analyse"
    }
    else {
        $Action = "To Keep"
    }

    $FormattedArray += [PSCustomObject]@{
        Action           = $Action
        Name             = $_.Data.DisplayName
        GPOStatus        = $_.Data.GpoStatus
        NoSettings       = $_.NoSettings
        UserSettings     = $_.UserSettings
        ComputerSettings = $_.ComputerSettings
        Linked           = $_.Linked
        LinkingError     = $_.LinkingError
        Links            = $_.ConcatenatedLinks
        GPOHasTargets    = $_.GpoHasTarget
        Targets          = $_.Targets
        WmiFilters       = $_.Data.WmiFilter
    }
}

Get-ChildItem $Config.Settings.FullName -Filter "*.csv" | ForEach-Object {
    Import-Csv $_.FullName | Export-Csv "$($Config.Results.FullName)\AllSettings.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Append
}

$FormattedArray | Export-Csv -Path "$($Config.Results.FullName)\GPO-Summary.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force 