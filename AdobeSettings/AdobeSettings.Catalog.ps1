function Get-AdobeAppDefinitions {
    @(
        [PSCustomObject]@{
            Key = "premiere-pro"
            DisplayName = "Premiere Pro"
            Scanner = {
                param($ProfilePath)
                Get-PremiereCatalogEntries -ProfilePath $ProfilePath
            }
        },
        [PSCustomObject]@{
            Key = "after-effects"
            DisplayName = "After Effects"
            Scanner = {
                param($ProfilePath)
                Get-AfterEffectsCatalogEntries -ProfilePath $ProfilePath
            }
        },
        [PSCustomObject]@{
            Key = "photoshop"
            DisplayName = "Photoshop"
            Scanner = {
                param($ProfilePath)
                Get-PhotoshopCatalogEntries -ProfilePath $ProfilePath
            }
        },
        [PSCustomObject]@{
            Key = "media-encoder"
            DisplayName = "Adobe Media Encoder"
            Scanner = {
                param($ProfilePath)
                Get-MediaEncoderCatalogEntries -ProfilePath $ProfilePath
            }
        }
    )
}

function Get-AdobeAppFilters {
    @{
        "premiere-pro" = @{
            ExcludeNames = @(
                "Recent Directories",
                "Media Browser Provider Exception"
            )
            ExcludeWildcards = @(
                "metadatacache*"
            )
        }
        "after-effects" = @{
            ExcludeNames = @(
                "Cache"
            )
            ExcludeWildcards = @()
        }
        "photoshop" = @{
            ExcludeNames = @(
                "PSErrorLog.txt",
                "FMCache.psp",
                "PluginCache.psp",
                "LaunchEndFlag.psp",
                "QuitEndFlag.psp",
                "MachinePrefs.psp"
            )
            ExcludeWildcards = @(
                "sniffer-out*"
            )
        }
        "media-encoder" = @{
            ExcludeNames = @(
                "logs",
                "SyncBackup",
                "SystemCompatibilityReport",
                "Debug Database.txt",
                "Trace Database.txt",
                "Recent Directories",
                "PresetCache.xml",
                "PresetUserCache.xml",
                "batch.xml",
                "batch.xml.bak",
                "AudioPlugInsScannedOnFirstLaunch_x64"
            )
            ExcludeWildcards = @(
                "AMEEncoding*",
                "*.log"
            )
        }
    }
}

function Get-AdobeAppFilter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppKey
    )

    $filters = Get-AdobeAppFilters
    if ($filters.ContainsKey($AppKey)) {
        return $filters[$AppKey]
    }

    return @{
        ExcludeNames = @()
        ExcludeWildcards = @()
    }
}

function Test-AdobeItemExcluded {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppKey,

        [Parameter(Mandatory = $true)]
        [string]$ItemName
    )

    $filter = Get-AdobeAppFilter -AppKey $AppKey
    foreach ($name in $filter.ExcludeNames) {
        if ($ItemName -ieq $name) {
            return $true
        }
    }

    foreach ($pattern in $filter.ExcludeWildcards) {
        if ($ItemName -like $pattern) {
            return $true
        }
    }

    return $false
}

function New-AdobeCatalogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppKey,

        [Parameter(Mandatory = $true)]
        [string]$AppName,

        [Parameter(Mandatory = $true)]
        [string]$Version,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    [PSCustomObject]@{
        AppKey = $AppKey
        AppName = $AppName
        Version = $Version
        Label = $Label
        SourcePath = $SourcePath
        RelativePath = $RelativePath
        EntryType = "Directory"
    }
}

function Get-RelativePathWithinUserProfile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfilePath,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $normalizedProfilePath = [System.IO.Path]::GetFullPath($ProfilePath)
    $normalizedPath = [System.IO.Path]::GetFullPath($Path)
    if (-not $normalizedPath.StartsWith($normalizedProfilePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Путь $normalizedPath не находится внутри профиля $normalizedProfilePath"
    }

    return $normalizedPath.Substring($normalizedProfilePath.Length).TrimStart('\')
}

function Get-PremiereCatalogEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfilePath
    )

    $entries = @()
    $root = Join-Path $ProfilePath "Documents\Adobe\Premiere Pro"
    if (-not (Test-Path -LiteralPath $root)) {
        return $entries
    }

    foreach ($versionDir in Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name) {
        foreach ($profileDir in Get-ChildItem -LiteralPath $versionDir.FullName -Directory -Filter "Profile-*" -ErrorAction SilentlyContinue | Sort-Object Name) {
            $relativePath = Get-RelativePathWithinUserProfile -ProfilePath $ProfilePath -Path $profileDir.FullName
            $entries += New-AdobeCatalogEntry `
                -AppKey "premiere-pro" `
                -AppName "Premiere Pro" `
                -Version $versionDir.Name `
                -Label ("{0} / {1}" -f $versionDir.Name, $profileDir.Name) `
                -SourcePath $profileDir.FullName `
                -RelativePath $relativePath
        }
    }

    return $entries
}

function Get-AfterEffectsCatalogEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfilePath
    )

    $entries = @()
    $root = Join-Path $ProfilePath "AppData\Roaming\Adobe\After Effects"
    if (-not (Test-Path -LiteralPath $root)) {
        return $entries
    }

    foreach ($versionDir in Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+(\.\d+)*$' } | Sort-Object Name) {
        $relativePath = Get-RelativePathWithinUserProfile -ProfilePath $ProfilePath -Path $versionDir.FullName
        $entries += New-AdobeCatalogEntry `
            -AppKey "after-effects" `
            -AppName "After Effects" `
            -Version $versionDir.Name `
            -Label $versionDir.Name `
            -SourcePath $versionDir.FullName `
            -RelativePath $relativePath
    }

    return $entries
}

function Get-PhotoshopCatalogEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfilePath
    )

    $entries = @()
    $root = Join-Path $ProfilePath "AppData\Roaming\Adobe"
    if (-not (Test-Path -LiteralPath $root)) {
        return $entries
    }

    foreach ($productDir in Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "Adobe Photoshop *" } | Sort-Object Name) {
        $settingsDir = Join-Path $productDir.FullName ("{0} Settings" -f $productDir.Name)
        if (-not (Test-Path -LiteralPath $settingsDir)) {
            continue
        }

        $version = $productDir.Name -replace '^Adobe Photoshop\s+', ''
        $relativePath = Get-RelativePathWithinUserProfile -ProfilePath $ProfilePath -Path $settingsDir
        $entries += New-AdobeCatalogEntry `
            -AppKey "photoshop" `
            -AppName "Photoshop" `
            -Version $version `
            -Label ("{0} / Settings" -f $version) `
            -SourcePath $settingsDir `
            -RelativePath $relativePath
    }

    return $entries
}

function Get-MediaEncoderCatalogEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfilePath
    )

    $entries = @()
    $root = Join-Path $ProfilePath "Documents\Adobe\Adobe Media Encoder"
    if (-not (Test-Path -LiteralPath $root)) {
        return $entries
    }

    foreach ($versionDir in Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+(\.\d+)*$' } | Sort-Object Name) {
        $relativePath = Get-RelativePathWithinUserProfile -ProfilePath $ProfilePath -Path $versionDir.FullName
        $entries += New-AdobeCatalogEntry `
            -AppKey "media-encoder" `
            -AppName "Adobe Media Encoder" `
            -Version $versionDir.Name `
            -Label $versionDir.Name `
            -SourcePath $versionDir.FullName `
            -RelativePath $relativePath
    }

    return $entries
}

function Get-AdobeAppCatalogForProfile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfilePath
    )

    $apps = @()
    foreach ($definition in Get-AdobeAppDefinitions) {
        $entries = @(& $definition.Scanner $ProfilePath)
        if ($entries.Count -eq 0) {
            continue
        }

        $versions = @($entries | Select-Object -ExpandProperty Version -Unique | Sort-Object)
        $apps += [PSCustomObject]@{
            Key = $definition.Key
            DisplayName = $definition.DisplayName
            Entries = $entries
            Versions = $versions
            Summary = "{0}: {1}" -f $definition.DisplayName, ($versions -join ", ")
        }
    }

    return @($apps)
}

function Get-AdobeUserProfiles {
    $excludedNames = @(
        "All Users",
        "Default",
        "Default User",
        "Public",
        "Все пользователи"
    )

    $currentUserProfile = [Environment]::GetFolderPath("UserProfile")
    $result = @()

    foreach ($profile in Get-CimInstance Win32_UserProfile | Sort-Object LocalPath) {
        if ([string]::IsNullOrWhiteSpace($profile.LocalPath)) {
            continue
        }

        if ($profile.LocalPath -notmatch '^C:\\Users\\[^\\]+$') {
            continue
        }

        if ($profile.Special) {
            continue
        }

        $userName = Split-Path -Leaf $profile.LocalPath
        if ($excludedNames -contains $userName) {
            continue
        }

        $item = Get-Item -LiteralPath $profile.LocalPath -ErrorAction SilentlyContinue
        if (-not $item) {
            continue
        }

        $attributes = [string]$item.Attributes
        if ($attributes -match "Hidden" -or $attributes -match "System") {
            continue
        }

        $apps = @(Get-AdobeAppCatalogForProfile -ProfilePath $profile.LocalPath)
        $summary = if ($apps.Count -gt 0) {
            ($apps | ForEach-Object { $_.Summary }) -join " | "
        }
        else {
            "Поддерживаемые каталоги не найдены"
        }

        $result += [PSCustomObject]@{
            UserName = $userName
            ProfilePath = $profile.LocalPath
            IsCurrentUser = [string]::Equals(
                [System.IO.Path]::GetFullPath($profile.LocalPath),
                [System.IO.Path]::GetFullPath($currentUserProfile),
                [System.StringComparison]::OrdinalIgnoreCase
            )
            Loaded = [bool]$profile.Loaded
            Apps = $apps
            DetectedAppCount = $apps.Count
            Summary = $summary
        }
    }

    return $result
}

function Get-AdobeAppDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppKey
    )

    $definition = Get-AdobeAppDefinitions | Where-Object { $_.Key -eq $AppKey } | Select-Object -First 1
    if ($definition) {
        return $definition.DisplayName
    }

    return $AppKey
}
