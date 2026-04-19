function Test-AdobeSettingsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-AdobeSettingsShellPath {
    if ($PSVersionTable.PSEdition -eq "Core") {
        return (Get-Command pwsh -ErrorAction Stop).Source
    }

    return (Get-Command powershell -ErrorAction Stop).Source
}

function Get-AdobeSettingsOperatorDesktop {
    $desktop = [Environment]::GetFolderPath("Desktop")
    if (-not [string]::IsNullOrWhiteSpace($desktop) -and (Test-Path -LiteralPath $desktop)) {
        return $desktop
    }

    $fallback = Join-Path ([Environment]::GetFolderPath("UserProfile")) "Desktop"
    if (Test-Path -LiteralPath $fallback) {
        return $fallback
    }

    return [Environment]::GetFolderPath("MyDocuments")
}

function Get-AdobeSettingsExportRoot {
    Join-Path (Get-AdobeSettingsOperatorDesktop) "AdobeSettingsExports"
}

function Get-AdobeSettingsBackupRoot {
    Join-Path (Get-AdobeSettingsOperatorDesktop) "AdobeSettingsBackups"
}

function New-AdobeSettingsLogger {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $directory = Split-Path -Parent $LogFilePath
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }

    if (Test-Path -LiteralPath $LogFilePath) {
        Remove-Item -LiteralPath $LogFilePath -Force
    }

    return {
        param(
            [string]$Level,
            [string]$Message
        )

        $line = "[{0}] {1}" -f $Level.ToUpperInvariant(), $Message
        Write-Host $line
        Add-Content -LiteralPath $LogFilePath -Value $line
    }.GetNewClosure()
}

function Assert-AdobePathWithinRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$CandidatePath
    )

    $resolvedRoot = [System.IO.Path]::GetFullPath($RootPath)
    $resolvedCandidate = [System.IO.Path]::GetFullPath($CandidatePath)
    if (-not $resolvedCandidate.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Путь $resolvedCandidate выходит за пределы $resolvedRoot"
    }
}

function Copy-AdobeLiteralItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $parentPath = Split-Path -Parent $DestinationPath
    if (-not (Test-Path -LiteralPath $parentPath)) {
        New-Item -ItemType Directory -Path $parentPath -Force | Out-Null
    }

    if (Test-Path -LiteralPath $SourcePath -PathType Container) {
        Copy-Item -LiteralPath $SourcePath -Destination $parentPath -Recurse -Force
    }
    else {
        Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    }
}

function Copy-AdobeFilteredDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppKey,

        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $copiedFiles = New-Object "System.Collections.Generic.List[string]"
    $skippedItems = New-Object "System.Collections.Generic.List[string]"

    $destinationParent = Split-Path -Parent $DestinationPath
    if (-not (Test-Path -LiteralPath $destinationParent)) {
        New-Item -ItemType Directory -Path $destinationParent -Force | Out-Null
    }

    function Copy-AdobeFilteredDirectoryInternal {
        param(
            [string]$WorkingSource,
            [string]$WorkingDestination
        )

        if (-not (Test-Path -LiteralPath $WorkingDestination)) {
            New-Item -ItemType Directory -Path $WorkingDestination -Force | Out-Null
        }

        foreach ($child in Get-ChildItem -LiteralPath $WorkingSource -Force -ErrorAction SilentlyContinue) {
            if (Test-AdobeItemExcluded -AppKey $AppKey -ItemName $child.Name) {
                [void]$skippedItems.Add($child.FullName.Substring($SourcePath.Length).TrimStart('\'))
                continue
            }

            $childDestination = Join-Path $WorkingDestination $child.Name
            if ($child.PSIsContainer) {
                Copy-AdobeFilteredDirectoryInternal -WorkingSource $child.FullName -WorkingDestination $childDestination
            }
            else {
                Copy-Item -LiteralPath $child.FullName -Destination $childDestination -Force
                [void]$copiedFiles.Add($child.FullName.Substring($SourcePath.Length).TrimStart('\'))
            }
        }
    }

    Copy-AdobeFilteredDirectoryInternal -WorkingSource $SourcePath -WorkingDestination $DestinationPath

    return [PSCustomObject]@{
        CopiedFiles = @($copiedFiles)
        SkippedItems = @($skippedItems)
    }
}

function Get-AdobePackageAppsFromManifest {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Manifest
    )

    $apps = @()
    foreach ($app in @($Manifest.Apps)) {
        $entries = @($app.Entries)
        $versions = @($entries | Select-Object -ExpandProperty Version -Unique | Sort-Object)
        $apps += [PSCustomObject]@{
            Key = $app.Key
            DisplayName = $app.DisplayName
            Entries = $entries
            Versions = $versions
            Summary = "{0}: {1}" -f $app.DisplayName, ($versions -join ", ")
        }
    }

    return $apps
}

function Get-AdobeImportPackageCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptRoot
    )

    $candidates = @()
    $seenPaths = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
    $searchRoots = @(
        (Get-AdobeSettingsExportRoot),
        (Get-AdobeSettingsOperatorDesktop),
        $ScriptRoot
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    foreach ($root in $searchRoots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        foreach ($manifestFile in Get-ChildItem -LiteralPath $root -Filter "manifest.json" -Recurse -ErrorAction SilentlyContinue) {
            $packageRoot = Split-Path -Parent $manifestFile.FullName
            if (-not $seenPaths.Add($packageRoot)) {
                continue
            }

            try {
                $manifest = Get-Content -LiteralPath $manifestFile.FullName -Raw | ConvertFrom-Json
                $appsSummary = (@($manifest.Apps) | ForEach-Object { $_.DisplayName }) -join ", "
                $candidates += [PSCustomObject]@{
                    Path = $packageRoot
                    Label = Split-Path -Leaf $packageRoot
                    Description = "{0}; {1}" -f $manifest.SourceUser.Name, $appsSummary
                    Kind = "directory"
                }
            }
            catch {
                $candidates += [PSCustomObject]@{
                    Path = $packageRoot
                    Label = Split-Path -Leaf $packageRoot
                    Description = "Папка пакета без читаемого manifest.json"
                    Kind = "directory"
                }
            }
        }

        foreach ($zipFile in Get-ChildItem -LiteralPath $root -Filter "*.zip" -Recurse -ErrorAction SilentlyContinue) {
            if (-not $seenPaths.Add($zipFile.FullName)) {
                continue
            }

            $candidates += [PSCustomObject]@{
                Path = $zipFile.FullName
                Label = $zipFile.Name
                Description = "Архив пакета"
                Kind = "zip"
            }
        }
    }

    return @($candidates | Sort-Object Label)
}

function Open-AdobeImportPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackagePath
    )

    $resolvedPath = (Resolve-Path -LiteralPath $PackagePath).Path
    if (Test-Path -LiteralPath $resolvedPath -PathType Leaf) {
        $extractRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("adobe-import-" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $extractRoot | Out-Null
        Expand-Archive -LiteralPath $resolvedPath -DestinationPath $extractRoot -Force
        $manifestFile = Get-ChildItem -LiteralPath $extractRoot -Filter "manifest.json" -Recurse -ErrorAction Stop | Select-Object -First 1
        $packageRoot = Split-Path -Parent $manifestFile.FullName
        $manifest = Get-Content -LiteralPath $manifestFile.FullName -Raw | ConvertFrom-Json
        return [PSCustomObject]@{
            SourcePath = $resolvedPath
            PackageRoot = $packageRoot
            CleanupPath = $extractRoot
            Manifest = $manifest
        }
    }

    $manifestPath = Join-Path $resolvedPath "manifest.json"
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        throw "В выбранной папке нет manifest.json: $resolvedPath"
    }

    return [PSCustomObject]@{
        SourcePath = $resolvedPath
        PackageRoot = $resolvedPath
        CleanupPath = $null
        Manifest = (Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json)
    }
}

function Close-AdobeImportPackage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Package
    )

    if (-not [string]::IsNullOrWhiteSpace($Package.CleanupPath) -and (Test-Path -LiteralPath $Package.CleanupPath)) {
        Remove-Item -LiteralPath $Package.CleanupPath -Recurse -Force
    }
}

function Get-AdobeProcessNamesForApp {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppKey
    )

    switch ($AppKey) {
        "premiere-pro" { return @("Adobe Premiere Pro") }
        "after-effects" { return @("AfterFX") }
        "photoshop" { return @("Photoshop") }
        "media-encoder" { return @("Adobe Media Encoder") }
        default { return @() }
    }
}

function Get-RunningAdobeProcesses {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$AppKeys
    )

    $names = @()
    foreach ($appKey in $AppKeys) {
        $names += Get-AdobeProcessNamesForApp -AppKey $appKey
    }

    $names = $names | Select-Object -Unique
    if ($names.Count -eq 0) {
        return @()
    }

    return @(
        Get-CimInstance Win32_Process |
        Where-Object { $names -contains $_.Name.Replace(".exe", "") -or $names -contains $_.Name } |
        Select-Object Name, ProcessId, ExecutablePath
    )
}

function Get-PremiereImportRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetProfilePath,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    if ($RelativePath -notmatch '^Documents\\Adobe\\Premiere Pro\\([^\\]+)\\(Profile-[^\\]+)$') {
        return $RelativePath
    }

    $version = $matches[1]
    $sourceProfileName = $matches[2]
    $targetUserName = Split-Path -Leaf $TargetProfilePath
    $versionDirectory = Join-Path $TargetProfilePath ("Documents\Adobe\Premiere Pro\{0}" -f $version)
    $preferredName = "Profile-{0}" -f $targetUserName
    $preferredRelativePath = "Documents\Adobe\Premiere Pro\{0}\{1}" -f $version, $preferredName

    if (Test-Path -LiteralPath (Join-Path $TargetProfilePath $RelativePath)) {
        return $RelativePath
    }

    if (Test-Path -LiteralPath (Join-Path $TargetProfilePath $preferredRelativePath)) {
        return $preferredRelativePath
    }

    if (Test-Path -LiteralPath $versionDirectory) {
        $profileDirectories = @(Get-ChildItem -LiteralPath $versionDirectory -Directory -Filter "Profile-*" -ErrorAction SilentlyContinue)
        if ($profileDirectories.Count -eq 1) {
            return "Documents\Adobe\Premiere Pro\{0}\{1}" -f $version, $profileDirectories[0].Name
        }
    }

    if ($sourceProfileName -ieq $preferredName) {
        return $RelativePath
    }

    return $preferredRelativePath
}

function Get-ImportDestinationRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetProfilePath,

        [Parameter(Mandatory = $true)]
        [object]$Entry
    )

    if ($Entry.AppKey -eq "premiere-pro") {
        return Get-PremiereImportRelativePath -TargetProfilePath $TargetProfilePath -RelativePath $Entry.RelativePath
    }

    return $Entry.RelativePath
}

function Export-AdobeSettingsPackage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourceProfile,

        [Parameter(Mandatory = $true)]
        [object[]]$SelectedApps
    )

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $packageName = "AdobeSettings-{0}-{1}" -f $SourceProfile.UserName, $timestamp
    $exportRoot = Get-AdobeSettingsExportRoot
    $packageRoot = Join-Path $exportRoot $packageName
    $payloadRoot = Join-Path $packageRoot "payload"
    $logPath = Join-Path $packageRoot "operation.log"

    New-Item -ItemType Directory -Path $payloadRoot -Force | Out-Null
    $log = New-AdobeSettingsLogger -LogFilePath $logPath
    & $log "info" ("Экспорт пользователя {0} из {1}" -f $SourceProfile.UserName, $SourceProfile.ProfilePath)

    $manifestApps = @()
    foreach ($app in $SelectedApps) {
        & $log "info" ("Приложение: {0}" -f $app.DisplayName)
        $manifestEntries = @()
        foreach ($entry in $app.Entries) {
            $destinationPath = Join-Path $payloadRoot $entry.RelativePath
            & $log "info" ("Копирование {0}" -f $entry.RelativePath)
            $copyResult = Copy-AdobeFilteredDirectory -AppKey $app.Key -SourcePath $entry.SourcePath -DestinationPath $destinationPath
            if ($copyResult.SkippedItems.Count -gt 0) {
                & $log "warn" ("Пропущено элементов: {0}" -f $copyResult.SkippedItems.Count)
            }

            $manifestEntries += [PSCustomObject]@{
                AppKey = $entry.AppKey
                AppName = $entry.AppName
                Version = $entry.Version
                Label = $entry.Label
                RelativePath = $entry.RelativePath
                SourcePath = $entry.SourcePath
                EntryType = $entry.EntryType
                CopiedFiles = @($copyResult.CopiedFiles)
                SkippedItems = @($copyResult.SkippedItems)
            }
        }

        $manifestApps += [PSCustomObject]@{
            Key = $app.Key
            DisplayName = $app.DisplayName
            Entries = $manifestEntries
        }
    }

    $manifest = [PSCustomObject]@{
        SchemaVersion = 1
        ExportedAtUtc = [DateTime]::UtcNow.ToString("o")
        SourceUser = [PSCustomObject]@{
            Name = $SourceProfile.UserName
            ProfilePath = $SourceProfile.ProfilePath
        }
        Apps = $manifestApps
        Filters = Get-AdobeAppFilters
    }

    $manifestPath = Join-Path $packageRoot "manifest.json"
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

    $zipPath = Join-Path $exportRoot ($packageName + ".zip")
    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }

    Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath -Force
    & $log "done" ("Архив создан: {0}" -f $zipPath)

    return [PSCustomObject]@{
        PackageRoot = $packageRoot
        ZipPath = $zipPath
        ManifestPath = $manifestPath
        LogPath = $logPath
    }
}

function Import-AdobeSettingsPackage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TargetProfile,

        [Parameter(Mandatory = $true)]
        [object]$Package,

        [Parameter(Mandatory = $true)]
        [string[]]$SelectedAppKeys
    )

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupRoot = Join-Path (Get-AdobeSettingsBackupRoot) ("AdobeSettingsBackup-{0}-{1}" -f $TargetProfile.UserName, $timestamp)
    $backupPayloadRoot = Join-Path $backupRoot "payload"
    $logPath = Join-Path $backupRoot "operation.log"
    New-Item -ItemType Directory -Path $backupPayloadRoot -Force | Out-Null

    $log = New-AdobeSettingsLogger -LogFilePath $logPath
    & $log "info" ("Импорт в профиль {0} ({1})" -f $TargetProfile.UserName, $TargetProfile.ProfilePath)

    $backupApps = @()
    foreach ($app in Get-AdobePackageAppsFromManifest -Manifest $Package.Manifest | Where-Object { $SelectedAppKeys -contains $_.Key }) {
        & $log "info" ("Приложение: {0}" -f $app.DisplayName)
        $backupEntries = @()
        foreach ($entry in $app.Entries) {
            $targetRelativePath = Get-ImportDestinationRelativePath -TargetProfilePath $TargetProfile.ProfilePath -Entry $entry
            $sourcePath = Join-Path (Join-Path $Package.PackageRoot "payload") $entry.RelativePath
            $targetPath = Join-Path $TargetProfile.ProfilePath $targetRelativePath
            Assert-AdobePathWithinRoot -RootPath $TargetProfile.ProfilePath -CandidatePath $targetPath

            if (-not (Test-Path -LiteralPath $sourcePath)) {
                & $log "warn" ("В пакете отсутствует источник: {0}" -f $entry.RelativePath)
                continue
            }

            if (Test-Path -LiteralPath $targetPath) {
                $backupPath = Join-Path $backupPayloadRoot $targetRelativePath
                & $log "info" ("Резервная копия: {0}" -f $targetRelativePath)
                Copy-AdobeLiteralItem -SourcePath $targetPath -DestinationPath $backupPath
                $backupEntries += [PSCustomObject]@{
                    AppKey = $entry.AppKey
                    AppName = $entry.AppName
                    Version = $entry.Version
                    RelativePath = $targetRelativePath
                }

                Remove-Item -LiteralPath $targetPath -Recurse -Force
            }

            & $log "info" ("Установка: {0}" -f $targetRelativePath)
            Copy-AdobeLiteralItem -SourcePath $sourcePath -DestinationPath $targetPath
        }

        if ($backupEntries.Count -gt 0) {
            $backupApps += [PSCustomObject]@{
                Key = $app.Key
                DisplayName = $app.DisplayName
                Entries = $backupEntries
            }
        }
    }

    $backupManifest = [PSCustomObject]@{
        SchemaVersion = 1
        CreatedAtUtc = [DateTime]::UtcNow.ToString("o")
        TargetUser = [PSCustomObject]@{
            Name = $TargetProfile.UserName
            ProfilePath = $TargetProfile.ProfilePath
        }
        BackupOfPackage = $Package.SourcePath
        Apps = $backupApps
    }

    $backupManifestPath = Join-Path $backupRoot "manifest.json"
    $backupManifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $backupManifestPath -Encoding UTF8
    & $log "done" ("Импорт завершён. Резервная копия: {0}" -f $backupRoot)

    return [PSCustomObject]@{
        BackupRoot = $backupRoot
        BackupManifestPath = $backupManifestPath
        LogPath = $logPath
    }
}

function Request-AdobeSettingsElevation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $true)]
        [string]$UserProfilePath,

        [string]$PackagePath,

        [string[]]$SelectedAppKeys
    )

    $shellPath = Get-AdobeSettingsShellPath
    $arguments = @(
        "-ExecutionPolicy", "Bypass",
        "-File", ('"{0}"' -f $ScriptPath),
        "-ResumeAction", $Action,
        "-ResumeUserProfilePath", ('"{0}"' -f $UserProfilePath)
    )

    if (-not [string]::IsNullOrWhiteSpace($PackagePath)) {
        $arguments += @("-ResumePackagePath", ('"{0}"' -f $PackagePath))
    }

    if ($SelectedAppKeys -and $SelectedAppKeys.Count -gt 0) {
        $arguments += @("-ResumeSelectedAppsCsv", ('"{0}"' -f ($SelectedAppKeys -join ";")))
    }

    Start-Process -FilePath $shellPath -ArgumentList $arguments -Verb RunAs -WorkingDirectory (Split-Path -Parent $ScriptPath) | Out-Null
}
