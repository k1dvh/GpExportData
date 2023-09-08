function Get-GpoLinks {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCustomObject]
        $Gpo
    )

    [System.Object] $Links = $Gpo.Report.GPO.LinksTo
    [System.Object] $_Links = @() 

    $Gpo | Add-Member -NotePropertyName "Linked" -NotePropertyValue $true
    $Gpo | Add-Member -NotePropertyName "ConcatenatedLinks" -NotePropertyValue $null

    if (!$Links) {
        $Gpo.Linked = $false
    }
    else {
        foreach ($Link in $Links) {
            [bool] $WellLinkedUser = $true
            [bool] $WellLinkedComputer = $true
            [string] $LinkStatus = "OK"
            
            $Gpo.ConcatenatedLinks += "|" + $Link.SOMPath

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

        if ($Gpo.ConcatenatedLinks.Length -gt 0) {
            $Gpo.ConcatenatedLinks = $Gpo.ConcatenatedLinks.Substring(1)
        }

        $_Links | Export-Csv -Path "$($Config.Links.FullName)\$($GPO.Data.DisplayName)-Links.csv" `
            -Delimiter ";" `
            -Encoding UTF8 `
            -NoTypeInformation `
            -Force
    }

    # return $LinksObject
}

function Get-GpoDelegations {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCustomObject]
        $Gpo
    )

    [System.Object] $_Delegs = @()
    [System.Object] $Perms = Get-GPPermissions -Name $Gpo.Data.DisplayName -All

    $Gpo | Add-Member -NotePropertyName "GpoHasTarget" -NotePropertyValue $true
    $Gpo | Add-Member -NotePropertyName "Targets" -NotePropertyValue $null



    if (!($Perms | Where-Object { $_.Permission -eq "GpoApply" })) {
        $Gpo.GpoHasTarget -eq $false
    }
    else {
        foreach ($Perm in ($Perms | Where-Object { $_.Permission -eq "GpoApply" })) {
            $Gpo.Targets += "|" + $Perm.Trustee.Name
        }
    }

    if ($Gpo.Targets.Length -gt 0) {
        $Gpo.Targets = $Gpo.Targets.Substring(1)
    }

    foreach ($Perm in $Perms) {
        $_Delegs += [PSCustomObject]@{
            GPOName    = $GPO.Data.DisplayName
            Type       = $Perm.Permission
            Trustee    = $Perm.Trustee.Name
            TrusteeSID = $Perm.Trustee.SID
            Inherited  = $Perm.Inherited        
        }
    }

    $_Delegs | Export-Csv -Path "$($Config.Delegations.FullName)\$($GPO.Data.DisplayName)-Delegations.csv" `
        -Delimiter ";" `
        -Encoding UTF8 `
        -NoTypeInformation `
        -Force
}

function Get-GpoGlobalInformations {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCustomObject]
        $Gpo
    )
    $Gpo | Add-Member -NotePropertyName "ComputerSettings" -NotePropertyValue $false
    $Gpo | Add-Member -NotePropertyName "UserSettings" -NotePropertyValue $false
    $Gpo | Add-Member -NotePropertyName "NoSettings" -NotePropertyValue $false

    if ($Gpo.Report.Gpo.User.ExtensionData) {
        $Gpo.UserSettings = $true
    }
    if ($Gpo.Report.Gpo.Computer.ExtensionData) {
        $Gpo.ComputerSettings = $true
    }

    if (!$Gpo.UserSettings -and !$Gpo.ComputerSettings) {
        $Gpo.NoSettings = $true
    }
}