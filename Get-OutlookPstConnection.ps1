Function New-NoteObject {
    Param(
        [String[]]$PropertyName
    )

    $Object = '' | Select-Object -Property $PropertyName
    Return $Object
}

Function ConvertFrom-RegistryUnicode {
    Param(
        [Byte[]]$Bytes
    )

    If ($Null -Eq $Bytes -Or $Bytes.Length -Lt 2) {
        Return ''
    }

    $Builder = ''
    $Limit = $Bytes.Length - ($Bytes.Length % 2)
    For ($Index = 0; $Index -Lt $Limit; $Index += 2) {
        $Code = [Int]$Bytes[$Index] + ([Int]$Bytes[$Index + 1] * 256)
        If ($Code -Eq 0) {
            Break
        }
        $Builder += [Char]$Code
    }

    Return $Builder
}

Function ConvertTo-DecodedRegistryText {
    Param(
        $RawValue
    )

    If ($Null -Eq $RawValue) {
        Return ''
    }

    If ($RawValue -Is [String]) {
        Return ([String]$RawValue).Trim()
    }

    Try {
        Return (ConvertFrom-RegistryUnicode -Bytes ([Byte[]]$RawValue)).Trim()
    } Catch {
        Return ''
    }
}

Function ConvertTo-NormalizedStorePath {
    Param(
        [String]$Path
    )

    $Clean = $Path.Trim()
    If ($Clean.StartsWith('\\?\UNC\')) {
        Return ('\\' + $Clean.Substring(8))
    }
    If ($Clean.StartsWith('\\?\')) {
        Return $Clean.Substring(4)
    }

    Return $Clean
}

Function Get-MappedDrivePath {
    Param(
        [String]$UncPath
    )

    If ([String]::IsNullOrWhiteSpace($UncPath) -Or -Not $UncPath.StartsWith('\\')) {
        Return ''
    }

    $Drives = @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue)
    ForEach ($Drive in $Drives) {
        $Root = [String]$Drive.DisplayRoot
        If ([String]::IsNullOrWhiteSpace($Root) -Or -Not $Root.StartsWith('\\')) {
            Continue
        }

        $Root = $Root.TrimEnd('\')
        $Prefix = $Root.ToLower() + '\'
        If ($UncPath.ToLower().StartsWith($Prefix)) {
            $Remainder = $UncPath.Substring($Root.Length).TrimStart('\')
            Return ($Drive.Name + ':\' + $Remainder)
        }
    }

    Return ''
}

Function Get-StoreLocation {
    Param(
        [String]$NormalizedPath
    )

    $Location = New-NoteObject -PropertyName @('UncPath', 'MappedPath', 'IsNetwork')
    $Location.UncPath = ''
    $Location.MappedPath = ''
    $Location.IsNetwork = $False

    If ([String]::IsNullOrWhiteSpace($NormalizedPath)) {
        Return $Location
    }

    If ($NormalizedPath.StartsWith('\\')) {
        $Location.UncPath = $NormalizedPath
        $Location.IsNetwork = $True
        $Location.MappedPath = Get-MappedDrivePath -UncPath $NormalizedPath
        Return $Location
    }

    If ($NormalizedPath.Length -Ge 2 -And $NormalizedPath.Substring(1, 1) -Eq ':') {
        $DriveName = $NormalizedPath.Substring(0, 1)
        $Drive = Get-PSDrive -Name $DriveName -PSProvider FileSystem -ErrorAction SilentlyContinue
        $Root = ''
        If ($Null -Ne $Drive) {
            $Root = [String]$Drive.DisplayRoot
        }
        If (-Not [String]::IsNullOrWhiteSpace($Root) -And $Root.StartsWith('\\')) {
            $Location.IsNetwork = $True
            $Location.MappedPath = $NormalizedPath
            $Relative = $NormalizedPath.Substring(2).TrimStart('\')
            $Root = $Root.TrimEnd('\')
            $Location.UncPath = $Root + '\' + $Relative
        }
    }

    Return $Location
}

Function Get-OutlookProfileLocation {
    $OfficeRoot = 'HKCU:\Software\Microsoft\Office'
    If (Test-Path -LiteralPath $OfficeRoot) {
        $VersionItems = @(Get-ChildItem -LiteralPath $OfficeRoot -ErrorAction SilentlyContinue)
        ForEach ($VersionItem in $VersionItems) {
            $VersionName = [String]$VersionItem.PSChildName
            If ($VersionName -NotMatch '^\d+\.\d+$') {
                Continue
            }

            $OutlookPath = 'HKCU:\Software\Microsoft\Office\{0}\Outlook' -F $VersionName
            $ProfileRoot = Join-Path -Path $OutlookPath -ChildPath 'Profiles'
            If (-Not (Test-Path -LiteralPath $ProfileRoot)) {
                Continue
            }

            $DefaultProfile = ''
            $OutlookProps = Get-ItemProperty -LiteralPath $OutlookPath -ErrorAction SilentlyContinue
            If ($Null -Ne $OutlookProps -And $Null -Ne $OutlookProps.DefaultProfile) {
                $DefaultProfile = [String]$OutlookProps.DefaultProfile
            }

            $ProfileLocation = New-NoteObject -PropertyName @('OfficeVersion', 'ProfileRoot', 'DefaultProfile')
            $ProfileLocation.OfficeVersion = $VersionName
            $ProfileLocation.ProfileRoot = $ProfileRoot
            $ProfileLocation.DefaultProfile = $DefaultProfile
            Write-Output $ProfileLocation
        }
    }

    $LegacyParent = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem'
    $LegacyRoot = Join-Path -Path $LegacyParent -ChildPath 'Profiles'
    If (Test-Path -LiteralPath $LegacyRoot) {
        $DefaultProfile = ''
        $LegacyProps = Get-ItemProperty -LiteralPath $LegacyParent -ErrorAction SilentlyContinue
        If ($Null -Ne $LegacyProps -And $Null -Ne $LegacyProps.DefaultProfile) {
            $DefaultProfile = [String]$LegacyProps.DefaultProfile
        }

        $ProfileLocation = New-NoteObject -PropertyName @('OfficeVersion', 'ProfileRoot', 'DefaultProfile')
        $ProfileLocation.OfficeVersion = 'Legacy'
        $ProfileLocation.ProfileRoot = $LegacyRoot
        $ProfileLocation.DefaultProfile = $DefaultProfile
        Write-Output $ProfileLocation
    }
}

Function Get-OutlookPstConnection {
    $Seen = @{}
    $Locations = @(Get-OutlookProfileLocation)

    ForEach ($Location in $Locations) {
        $ProfileItems = @(Get-ChildItem -LiteralPath $Location.ProfileRoot -ErrorAction SilentlyContinue)
        ForEach ($ProfileItem in $ProfileItems) {
            $ProfileName = [String]$ProfileItem.PSChildName
            $ProfilePath = Join-Path -Path $Location.ProfileRoot -ChildPath $ProfileName
            $IsDefault = $False
            If (-Not [String]::IsNullOrWhiteSpace($Location.DefaultProfile) -And $ProfileName -Eq $Location.DefaultProfile) {
                $IsDefault = $True
            }

            $StoreItems = @(Get-ChildItem -LiteralPath $ProfilePath -Recurse -ErrorAction SilentlyContinue)
            ForEach ($StoreItem in $StoreItems) {
                $StorePath = [String]$StoreItem.PSPath
                If ([String]::IsNullOrWhiteSpace($StorePath)) {
                    Continue
                }

                $Props = $Null
                Try {
                    $Props = Get-ItemProperty -LiteralPath $StorePath -ErrorAction Stop
                } Catch {
                    Write-Warning ("[WARN] Unable to read registry key: {0}" -F $StorePath)
                    Continue
                }

                $DecodedPath = ConvertTo-DecodedRegistryText -RawValue $Props.'001f6700'
                If ([String]::IsNullOrWhiteSpace($DecodedPath) -Or $DecodedPath -NotLike '*.pst') {
                    Continue
                }

                $NormalizedPath = ConvertTo-NormalizedStorePath -Path $DecodedPath
                $DedupeKey = $ProfileName.ToLower() + '|' + $NormalizedPath.ToLower()
                If ($Seen.ContainsKey($DedupeKey)) {
                    Continue
                }
                $Seen[$DedupeKey] = $True

                $DisplayName = ConvertTo-DecodedRegistryText -RawValue $Props.'001f3001'
                $StoreLocation = Get-StoreLocation -NormalizedPath $NormalizedPath

                $Connection = New-NoteObject -PropertyName @(
                    'DisplayName',
                    'Path',
                    'UncPath',
                    'MappedPath',
                    'IsNetwork',
                    'ProfileName',
                    'IsDefaultProfile',
                    'OfficeVersion',
                    'RegistryPath'
                )
                $Connection.DisplayName = $DisplayName
                $Connection.Path = $NormalizedPath
                $Connection.UncPath = [String]$StoreLocation.UncPath
                $Connection.MappedPath = [String]$StoreLocation.MappedPath
                $Connection.IsNetwork = $StoreLocation.IsNetwork
                $Connection.ProfileName = $ProfileName
                $Connection.IsDefaultProfile = $IsDefault
                $Connection.OfficeVersion = [String]$Location.OfficeVersion
                $Connection.RegistryPath = $DecodedPath
                Write-Output $Connection
            }
        }
    }
}

$LanguageMode = [String]$ExecutionContext.SessionState.LanguageMode
Write-Host ("[INFO] Language mode: {0}" -F $LanguageMode)

$ProfileLocations = @(Get-OutlookProfileLocation)
If ($ProfileLocations.Count -Eq 0) {
    Write-Warning "[WARN] No Outlook profile was found for the current user."
} Else {
    $Connections = @(Get-OutlookPstConnection)
    If ($Connections.Count -Gt 0) {
        Write-Host "[INFO] Outlook has a PST connection: True"
    } Else {
        Write-Host "[INFO] Outlook has a PST connection: False"
    }

    Write-Host ("[INFO] PST connection count: {0}" -F $Connections.Count)

    $Number = 0
    ForEach ($Connection in $Connections) {
        $Number += 1
        Write-Host ("[INFO] --- PST {0} ---" -F $Number)
        Write-Host ("[INFO] Display name: {0}" -F $Connection.DisplayName)
        Write-Host ("[INFO] Path: {0}" -F $Connection.Path)
        If (-Not [String]::IsNullOrWhiteSpace($Connection.MappedPath)) {
            Write-Host ("[INFO] Mapped path: {0}" -F $Connection.MappedPath)
        }
        If (-Not [String]::IsNullOrWhiteSpace($Connection.UncPath)) {
            Write-Host ("[INFO] UNC path: {0}" -F $Connection.UncPath)
        }
        Write-Host ("[INFO] Network location: {0}" -F $Connection.IsNetwork)
        Write-Host ("[INFO] Profile: {0} (Office {1}, default={2})" -F $Connection.ProfileName, $Connection.OfficeVersion, $Connection.IsDefaultProfile)
    }

    Write-Output $Connections
}
