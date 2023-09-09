function Get-GpoLinks {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory
        )]
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
        [Parameter(
            Mandatory
        )]
        [PSCustomObject]
        $Gpo
    )

    [System.Object] $_Delegs = @()
    [System.Object] $Perms = Get-GPPermissions -Name $Gpo.Data.DisplayName -All

    $Gpo | Add-Member -NotePropertyName "GpoHasTarget" -NotePropertyValue $true
    $Gpo | Add-Member -NotePropertyName "Targets" -NotePropertyValue $null

    if (!($Perms | Where-Object { $_.Permission -eq "GpoApply" })) {
        $Gpo.GpoHasTarget = $false
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

function Get-GpoSettings {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory
        )]
        [PSCustomObject]
        $Gpo
    )
    $_Settings = @()

    if ($Gpo.UserSettings) {
        # Write-Host "$($Gpo.Data.DisplayName) is user gpo"
    }

    if ($Gpo.ComputerSettings) {
        $Settings = $Gpo.Report.GPO.Computer.ExtensionData.Extension

        foreach ($Setting in $Settings) {
            $SettingType = ($Setting | Get-Member -MemberType Properties | Where-Object {
                    $_.Name -notlike "q*" -and $_.Name -ne "type" -and $_.Name -ne "blocked"
                }).Name

            foreach ($Type in $SettingType) {
                $_Settings += Set-SettingType -Setting $Setting -SettingType $Type
            }
        }

        $_Settings | Export-Csv -Path "$($Config.Settings.FullName)\$($GPO.Data.DisplayName)-Settings.csv" `
            -Delimiter ";" `
            -Encoding UTF8 `
            -NoTypeInformation `
            -Force
    }
}

function Set-SettingType {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory
        )]
        $Setting,
        [Parameter(
            Mandatory
        )]
        [string]
        $SettingType
        
    )

    $ExportArray = @()
    # Write-host $SettingType

    switch ($SettingType) {
        "Policy" {
            foreach ($Parameter in $Setting.$SettingType) {
                if ($Parameter.CheckBox -or $Parameter.DropDownList -or $Parameter.Numeric) {
                    if ($Parameter.CheckBox) {
                        foreach ($Checkbox in $Parameter.CheckBox) {
                            $ExportArray += [PSCustomObject]@{
                                GPOName         = $GPO.Data.DisplayName
                                SettingCategory = "Administrative Templates"
                                SettingPath     = $Parameter.Category
                                SettingName     = $Parameter.Name
                                SettingStatus   = $Parameter.State
                                Type            = "Checkbox"
                                Statement       = $Checkbox.Name
                                Value           = $Checkbox.State
                            }
                        }
                    }
                    if ($Parameter.Numeric) {
                        $Parameter.Numeric | ForEach-Object {
                            $ExportArray += [PSCustomObject]@{
                                GPOName         = $GPO.Data.DisplayName
                                SettingCategory = "Administrative Templates"
                                SettingPath     = $Parameter.Category
                                SettingName     = $Parameter.Name
                                SettingStatus   = $Parameter.State
                                Type            = "Numeric"
                                Statement       = $Parameter.Text.Name
                                Value           = $_.Value.Name
                            }
                        }
                    }

                    if ($Parameter.DropDownList) {
                        $Parameter.DropDownList | ForEach-Object {             
                            if ($_.Name) {
                                $ParameterDetail = $_.Name
                            }
                            else {
                                $ParameterDetail = $_.ParentNode.Text.Name
                            }
                            $ExportArray += [PSCustomObject]@{
                                GPOName         = $GPO.Data.DisplayName
                                SettingCategory = "Administrative Templates"
                                SettingPath     = $Parameter.Category
                                SettingName     = $Parameter.Name
                                SettingStatus   = $Parameter.State
                                Type            = "DropDownList"
                                Statement       = $ParameterDetail
                                Value           = $_.Value.Name
                            } 
                        }
                    }
                }
                else {
                    $ExportArray += [PSCustomObject]@{
                        GPOName         = $GPO.Data.DisplayName
                        SettingCategory = "Administrative Templates"
                        SettingPath     = $Parameter.Category
                        SettingName     = $Parameter.Name
                        SettingStatus   = $Parameter.State
                        Type            = $null
                        Statement       = $null
                        Value           = $null
                    }
                }
            }
        }
        "Script" {
            foreach ($Script in $Setting.$SettingType) {
                $ExportArray += [PSCustomObject]@{
                    GPOName         = $GPO.Data.DisplayName
                    SettingCategory = "Scripts"
                    SettingPath     = $null
                    SettingName     = $Script.Command
                    SettingStatus   = $null
                    Type            = $Script.Type
                    Statement       = "Order: $($Script.Order)"
                    Value           = $Script.RunOrder
                }                    
            }
        }
        "Folders" {
            foreach ($Setting in $Setting.$SettingType) {
                foreach ($Folder in $Setting.Folder) {
                    $ExportArray += [PSCustomObject]@{
                        GPOName         = $GPO.Data.DisplayName
                        SettingCategory = "Folders"
                        SettingPath     = $Folder.Properties.path
                        SettingName     = $Folder.Name
                        SettingStatus   = $null
                        Type            = $null
                        Statement       = "Order: $($Folder.GPOSettingOrder)"
                        Value           = "uid: $($Folder.uid)"
                    }       
                } 
            }
        }
        "ScheduledTasks" {
            foreach ($Setting in $Setting.$SettingType) {
                foreach ($Task in $Setting.TaskV2) {
                    $ExportArray += [PSCustomObject]@{
                        GPOName         = $GPO.Data.DisplayName
                        SettingCategory = "Scheduled Task"
                        SettingPath     = $null
                        SettingName     = $Task.name
                        SettingStatus   = $null
                        Type            = $null
                        Statement       = "RunAs: $($Task.Properties.runAs)"
                        Value           = "uid: $($Task.uid)"
                    }       
                }
            }
        }
        "UserRightsAssignment" {
            foreach ($Right in $Setting.UserRightsAssignment) {
                foreach ($Identity in $Right.Member) {
                    $ExportArray += [PSCustomObject]@{
                        GPOName         = $GPO.Data.DisplayName
                        SettingCategory = "User Right Assignement"
                        SettingPath     = $null
                        SettingName     = $Right.Name
                        SettingStatus   = $Identity.Name.'#text'
                        Type            = $null
                        Statement       = $null
                        Value           = "SID: $($Identity.SID.'#text')"
                    }  
                }
            }
        }
        "SecurityOptions" {
            Write-Host "1"
            foreach ($Set in $Setting.SecurityOptions) {
                if ($Set.Display.DisplayFields) {
                    foreach ($Field in $Setting.Display.DisplayFields.Field) {
                        $ExportArray += [PSCustomObject]@{
                            GPOName         = $GPO.Data.DisplayName
                            SettingCategory = "Security Option"
                            SettingPath     = "Reg Path: $($Set.KeyName)"
                            SettingName     = $Set.Display.Name
                            SettingStatus   = $Field.Name.'#text'
                            Type            = $null
                            Statement       = $Field.Name
                            Value           = $Field.Value
                        }    
                    }   
                }
                if ($Set.Display.DisplayBoolean) {
                    Write-Host "2"
                    foreach ($Param in $Set.Display) {
                        Write-Host $Param.Name
                        $ExportArray += [PSCustomObject]@{
                            GPOName         = $GPO.Data.DisplayName
                            SettingCategory = "Security Option"
                            SettingPath     = "Reg Path: $($Set.KeyName)"
                            SettingName     = $Param.Name
                            SettingStatus   = $Param.DisplayBoolean
                            Type            = $null
                            Statement       = $null
                            Value           = $null
                        }
                    }   
                }
            }
        }
        Default {
            Write-Host $SettingType
        }
    }
    return $ExportArray

}

function Get-GpoGlobalInformations {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory
        )]
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