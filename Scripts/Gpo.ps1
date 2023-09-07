function Get-GpoLinks {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCustomObject]
        $Gpo
    )

    [System.Object] $Links = $Gpo.Report.GPO.LinksTo
    [string] $ConcatenatedLinks = $null

    [System.Object] $_Links = @() 

    $Gpo | Add-Member -NotePropertyName "Linked" -NotePropertyValue $true

    if (!$Links) {
        $Gpo.Linked = $false
    }
    else {
        foreach ($Link in $Links) {
            [bool] $WellLinkedUser = $true
            [bool] $WellLinkedComputer = $true
            [string] $LinkStatus = "OK"
            
            $ConcatenatedLinks += "|" + $Link.SOMPath

            if ($Link.SOMPath -ne $Data.Domain.DNSRoot) {
                $LinkedOU = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName | Where-Object {
                    $_.CanonicalName -eq $Link.SOMPath
                }

                if ($Gpo.UserSettings -and (Get-ADUser -Filter { Enabled -eq $true } -SearchBase $LinkedOU.DistinguishedName).Count -eq 0) {
                    $WellLinkedUser = $false
                    $LinkStatus = "User Misconfiguration"
                }

                if ($Gpo.ComputerSettings -and (Get-ADComputer -Filter { Enabled -eq $true } -SearchBase $LinkedOU.DistinguishedName).Count -eq 0) {
                    $WellLinkedComputer = $false
                    $LinkStatus = "Computer Misconfiguration"
                }

                if (!$WellLinkedUser -and !$WellLinkedComputer) {
                    $LinkStatus = "Both users and computers are misconfigured"
                }
            }

            $LinksObject = [PSCustomObject]@{
                GPOName   = $Gpo.Data.DisplayName
                Link      = $Link.SOMPath
                Enabled   = $Link.Enabled
                GPOStatus = $LinkStatus    
            }

            # $LinksObject

            $_Links += $LinksObject
        }

        $_Links | Export-Csv -Path "$($Config.Links.FullName)\$($GPO.Data.DisplayName)-Links.csv" `
            -Delimiter ";" `
            -Encoding UTF8 `
            -NoTypeInformation `
            -Force
    }

    # return $LinksObject
}

function Get-GpoGlobalInformations {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCustomObject]
        $Gpo
    )
    $Gpo | Add-Member -NotePropertyName "GpoStatus" -NotePropertyValue $false
    $Gpo | Add-Member -NotePropertyName "ComputerSettings" -NotePropertyValue $false
    $Gpo | Add-Member -NotePropertyName "UserSettings" -NotePropertyValue $false

    if ($Gpo.Data.User.Enabled -and $Gpo.Data.Computer.Enabled) {
        $Gpo.GpoStatus = "AllSettingsEnabled"
    }
    elseif ($Gpo.Data.Computer.Enabled) {
        $Gpo.GpoStatus = "ComputerSettingsEnabled"
    }
    elseif ($Gpo.Data.User.Enabled) {
        $Gpo.GpoStatus = "UserSettingsEnabled"
    }
    else {
        $Gpo.GpoStatus = "AllSettingsDisabled"
    }

    if ($Gpo.Report.Gpo.User.ExtensionData) {
        $Gpo.UserSettings = $true
    }
    if ($Gpo.Report.Gpo.Computer.ExtensionData) {
        $Gpo.ComputerSettings = $true
    }
}