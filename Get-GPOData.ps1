<#
.SYNOPSIS
    This script exports GPO data for analysis purposes.
.DESCRIPTION
    The following data will be exported :
        - Gpo name
        - Gpo status
        - Is it a Computer Gpo, a User Gpo or both?
        - Is the Gpo linked somewhere?
            - If so, where?
        - Does the GPO have targets?
            - If so, who are the targets?
        - Does the GPO have any WmiFilters?
    It also creates to folder with more details on delegations and links.

    At the end, the script give his opinion on the GPO's future (delete, keep, modify)
.NOTES
    For any suggestions for improvement, please contact quentin.mallet@outlook.com
#>

if (!(Test-Path ".\Links")) {
    try {
        New-Item -Path "$PSScriptRoot" -Name "Links" -ItemType Directory -Force
    }
    catch {
        Write-Error $_.Exception.Message
    }
}

if (!(Test-Path ".\Delegations")) {
    try {
        New-Item -Path "$PSScriptRoot" -Name "Delegations" -ItemType Directory -Force
    }
    catch {
        Write-Error $_.Exception.Message
    }
}

$AllResults = @()

$ALlGPOs = Get-GPO -All
$DomainDNS = (Get-ADDomain).DNSRoot

foreach ($GPO in $ALlGPOs) {

    [bool] $isUserSettings = $false
    [bool] $isComputerSettings = $false
    [bool] $NoSettings = $false
    $AllLinks = @()
    [bool] $Linked = $true
    [bool] $LinkingError = $false
    [string] $ConcatenateLinks = ""

    [bool] $GPOTarget = $true
    [string] $Targets = ""
    [xml] $Report = Get-GPOReport $GPO.DisplayName -ReportType Xml

    $Links = $Report.GPO.LinksTo    


    if ($GPO.User.DSVersion -ne 0) {
        $isUserSettings = $true
    }
    if ($GPO.Computer.DSVersion -ne 0) {
        $isComputerSettings = $true
    }

    if ($GPo.Computer.DSVersion -eq 0 -and $GPO.User.DSVersion -eq 0) {
        $NoSettings = $true
    }

    if (!($Links)) {
        $Linked = $false
    }
    else {
        foreach ($Link in $Links) {
            $ConcatenateLinks += "|" + $Link.SOMPath
            [bool] $WellSetUser = $true
            [bool] $WellSetComputer = $true
            [string] $LinkStatus = "OK"

            if ($Link.SOMPath -ne $DomainDNS) {
                $LinkedOU = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName | Where-Object { $_.CanonicalName -eq $Link.SOMPath }
                if ($isUserSettings -and (Get-ADUser -Filter { Enabled -eq $true } -SearchBase $LinkedOU.DistinguishedName).Count -eq 0) {
                    $WellSetUser = $false
                    $LinkStatus = "User misconfiguration"
                }
    
                if ($isComputerSettings -and (Get-ADComputer -Filter { Enabled -eq $true } -SearchBase $LinkedOU.DistinguishedName).Count -eq 0) {
                    $WellSetComputer = $false
                    $LinkStatus = "Computer misconfiguration"
                }
    
                if (!$WellSetUser -and !$WellSetComputer) {
                    $LinkStatus = "Both user and computers are misconfigured"
                }
    
            }
            
            $ResultLinks = [PSCustomObject]@{
                GPOName   = $GPO.DisplayName
                Link      = $Link.SOMPath
                Enabled   = $Link.Enabled
                GPOStatus = $LinkStatus
            }

            $AllLinks += $ResultLinks
        }

        foreach ($link in $AllLinks.GPOisNotApplied) {
            if ($link -ne "OK") {
                $LinkingError = $true
                break
            }
        }

        $AllLinks | Export-Csv -Path ".\Links\$($GPO.DisplayName)-Links.csv" -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
    }

    $GPOPerms = (Get-GPPermission -Name $GPO.DisplayName -All)
    if (!($GPOPerms | Where-Object { $_.Permission -eq "GpoApply" } )) {
        $GPOTarget = $false
    }
    else {
        foreach ($Perm in ($GPOPerms | Where-Object { $_.Permission -eq "GpoApply" })) {
            $Targets += "|" + $Perm.Trustee.Name
        }
    }
    
    if ($Targets.Length -ne "" ) {
        $Targets = $Targets.Substring(1)
    }
    if ($ConcatenateLinks.Length -gt 0) {
        $ConcatenateLinks = $ConcatenateLinks.Substring(1)
    }

    $AllPerms = @()

    foreach ($Perms in $GPOPerms) {
        $AllPerms += [PSCustomObject]@{
            GPOName    = $GPO.DisplayName
            Type       = $Perms.Permission
            Trustee    = $Perms.Trustee.Name
            TrusteeSID = $Perms.Trustee.SID
            Inherited  = $Perms.Inherited
        }
    }

    $AllPerms | Export-Csv -Path ".\Delegations\$($GPO.DisplayName)-Deleg.csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation -Force

    $Action = "To Keep"

    if ($GPO.GpoStatus -eq "AllSettingsDisabled" `
            -or $NoSettings `
            -or !$GPOTarget) {
        $Action = "To Delete"
    }
    
    elseif ($LinkingError) {
        $Action = "To Analyse"
    }

    $ResultGPO = [PSCustomObject]@{
        Action           = $Action
        Name             = $GPO.DisplayName
        GPOStatus        = $GPO.GpoStatus
        NoSettings       = $NoSettings
        UserSettings     = $isUserSettings
        ComputerSettings = $isComputerSettings
        Linked           = $Linked
        LinkingError     = $LinkingError
        Links            = $ConcatenateLinks
        GPOHasTargets    = $GPOTarget
        Targets          = $Targets
        WmiFilter        = $GPO.WmiFilter
    }

    $AllResults += $ResultGPO
}

$AllResults | Export-Csv -Path "$PSScriptRoot\GPO-data.csv" -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force