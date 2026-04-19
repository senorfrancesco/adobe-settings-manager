param(
    [string]$ResumeAction,
    [string]$ResumeUserProfilePath,
    [string]$ResumePackagePath,
    [string]$ResumeSelectedAppsCsv
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $PSCommandPath

. (Join-Path $scriptRoot "AdobeSettings\AdobeSettings.Catalog.ps1")
. (Join-Path $scriptRoot "AdobeSettings\AdobeSettings.Ui.ps1")
. (Join-Path $scriptRoot "AdobeSettings\AdobeSettings.Operations.ps1")

Set-StrictMode -Version Latest

function Show-AdobeHostCompatibilityError {
    Show-AdobeInfoScreen `
        -Title "Неподдерживаемый хост" `
        -Lines @(
            "Скрипт рассчитан на обычную консоль PowerShell с поддержкой ReadKey().",
            "Запустите его через powershell.exe или pwsh.exe в консольном окне."
        )
}

function Get-AdobeActionMenuItems {
    @(
        (New-AdobeMenuItem -Id "export" -Label "Экспорт" -Description "Собрать пакет настроек из выбранного профиля Windows."),
        (New-AdobeMenuItem -Id "import" -Label "Импорт" -Description "Установить пакет настроек в выбранный профиль Windows."),
        (New-AdobeMenuItem -Id "exit" -Label "Выход" -Description "Закрыть мастер без изменений.")
    )
}

function Get-AdobeProfileMenuItems {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Profiles
    )

    foreach ($profile in $Profiles) {
        $description = if ($profile.DetectedAppCount -gt 0) {
            "{0} приложений" -f $profile.DetectedAppCount
        }
        else {
            "Поддерживаемые каталоги не найдены"
        }

        New-AdobeMenuItem -Id $profile.ProfilePath -Label $profile.UserName -Description $description -Data $profile
    }
}

function Show-AdobeProfileDetails {
    param(
        [Parameter(Mandatory = $true)]
        [object]$MenuItem
    )

    $profile = $MenuItem.Data
    $lines = @(
        "Путь профиля: {0}" -f $profile.ProfilePath,
        "Текущий пользователь: {0}" -f $(if ($profile.IsCurrentUser) { "да" } else { "нет" })
    )

    if ($profile.Apps.Count -gt 0) {
        $lines += ""
        foreach ($app in $profile.Apps) {
            $lines += ("{0}: {1}" -f $app.DisplayName, ($app.Versions -join ", "))
        }
    }
    else {
        $lines += ""
        $lines += "Каталоги Premiere Pro, After Effects, Photoshop и Media Encoder не найдены."
    }

    return $lines
}

function Get-AdobeAppMenuItems {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Apps
    )

    foreach ($app in $Apps) {
        New-AdobeMenuItem -Id $app.Key -Label $app.DisplayName -Description ($app.Versions -join ", ") -Data $app
    }
}

function Show-AdobeAppDetails {
    param(
        [Parameter(Mandatory = $true)]
        [object]$MenuItem
    )

    $app = $MenuItem.Data
    $lines = @(
        "Приложение: {0}" -f $app.DisplayName,
        "Версии: {0}" -f ($app.Versions -join ", "),
        ""
    )

    foreach ($entry in $app.Entries | Select-Object -First 8) {
        $lines += $entry.RelativePath
    }

    return $lines
}

function Show-AdobeConfirmMenu {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string[]]$DetailLines
    )

    $items = @(
        (New-AdobeMenuItem -Id "run" -Label "Продолжить" -Description "Запустить операцию."),
        (New-AdobeMenuItem -Id "back" -Label "Назад" -Description "Вернуться и изменить выбор.")
    )

    return Show-AdobeSingleSelectMenu -Title $Title -Items $items -DetailScript ({ param($item) $DetailLines }.GetNewClosure())
}

function Get-AdobeSelectedAppKeys {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$SelectedAppItems
    )

    return @($SelectedAppItems | ForEach-Object { $_.Id })
}

function Select-AdobeImportPackage {
    $candidates = @(Get-AdobeImportPackageCandidates -ScriptRoot $scriptRoot)
    $items = @()
    foreach ($candidate in $candidates) {
        $items += New-AdobeMenuItem -Id $candidate.Path -Label $candidate.Label -Description $candidate.Description -Data $candidate
    }
    $items += New-AdobeMenuItem -Id "__manual__" -Label "Указать путь вручную" -Description "Ввести путь к папке пакета или zip-архиву."

    $selection = Show-AdobeSingleSelectMenu `
        -Title "Импорт: выбор пакета" `
        -Items $items `
        -DetailScript {
            param($item)
            if ($item.Id -eq "__manual__") {
                return @(
                    "Вручную можно указать:",
                    "- папку, где лежит manifest.json",
                    "- zip-архив, созданный этим скриптом"
                )
            }

            return @(
                "Путь: {0}" -f $item.Data.Path,
                "Тип: {0}" -f $item.Data.Kind
            )
        }

    if (-not $selection) {
        return $null
    }

    if ($selection.Id -eq "__manual__") {
        $manualPath = Show-AdobeTextPrompt `
            -Title "Импорт: путь к пакету" `
            -Prompt "Введите путь к папке пакета или zip-архиву" `
            -DefaultValue ""

        if ([string]::IsNullOrWhiteSpace($manualPath)) {
            return $null
        }

        return $manualPath
    }

    return $selection.Id
}

function Confirm-AdobeElevationIfNeeded {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TargetProfile,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [string]$PackagePath,

        [string[]]$SelectedAppKeys
    )

    if ($TargetProfile.IsCurrentUser -or (Test-AdobeSettingsAdministrator)) {
        return $true
    }

    $menu = Show-AdobeSingleSelectMenu `
        -Title "Нужны права администратора" `
        -Items @(
            (New-AdobeMenuItem -Id "elevate" -Label "Перезапустить с повышением" -Description "Открыть новый процесс PowerShell с правами администратора."),
            (New-AdobeMenuItem -Id "cancel" -Label "Отмена" -Description "Вернуться без запуска операции.")
        ) `
        -DetailScript ({
            param($item)
            @(
                "Вы выбрали профиль другого пользователя: {0}" -f $TargetProfile.UserName,
                "Для чтения или замены его каталогов настроек нужны права администратора."
            )
        }.GetNewClosure())

    if (-not $menu -or $menu.Id -eq "cancel") {
        return $false
    }

    Request-AdobeSettingsElevation `
        -ScriptPath $PSCommandPath `
        -Action $Action `
        -UserProfilePath $TargetProfile.ProfilePath `
        -PackagePath $PackagePath `
        -SelectedAppKeys $SelectedAppKeys

    return $false
}

function Select-AdobeAppsForExport {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Profile
    )

    $items = @(Get-AdobeAppMenuItems -Apps $Profile.Apps)
    $selected = Show-AdobeMultiSelectMenu `
        -Title ("Экспорт: приложения для {0}" -f $Profile.UserName) `
        -Items $items `
        -InitiallySelectedIds @($items | ForEach-Object { $_.Id }) `
        -DetailScript ${function:Show-AdobeAppDetails}

    if ($null -eq $selected -or $selected.Count -eq 0) {
        return @()
    }

    return @($selected | ForEach-Object { $_.Data })
}

function Select-AdobeAppsForImport {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$PackageApps
    )

    $items = @(Get-AdobeAppMenuItems -Apps $PackageApps)
    $selected = Show-AdobeMultiSelectMenu `
        -Title "Импорт: приложения из пакета" `
        -Items $items `
        -InitiallySelectedIds @($items | ForEach-Object { $_.Id }) `
        -DetailScript ${function:Show-AdobeAppDetails}

    if ($null -eq $selected -or $selected.Count -eq 0) {
        return @()
    }

    return @($selected | ForEach-Object { $_.Data })
}

function Ensure-AdobeAppsClosedBeforeImport {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$SelectedApps
    )

    while ($true) {
        $running = @(Get-RunningAdobeProcesses -AppKeys (@($SelectedApps | ForEach-Object { $_.Key })))
        if ($running.Count -eq 0) {
            return $true
        }

        $menu = Show-AdobeSingleSelectMenu `
            -Title "Закройте приложения Adobe" `
            -Items @(
                (New-AdobeMenuItem -Id "retry" -Label "Проверить снова" -Description "Повторить поиск запущенных процессов."),
                (New-AdobeMenuItem -Id "cancel" -Label "Отмена" -Description "Вернуться без импорта.")
            ) `
            -DetailScript ({
                param($item)
                $lines = @("Сейчас запущены:")
                foreach ($process in $running) {
                    $lines += ("- {0} (PID {1})" -f $process.Name, $process.ProcessId)
                }
                return $lines
            }.GetNewClosure())

        if (-not $menu -or $menu.Id -eq "cancel") {
            return $false
        }
    }
}

function Invoke-AdobeExportFlow {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Profiles
    )

    $profileItem = Show-AdobeSingleSelectMenu `
        -Title "Экспорт: выбор профиля" `
        -Items @(Get-AdobeProfileMenuItems -Profiles $Profiles) `
        -DetailScript ${function:Show-AdobeProfileDetails}

    if (-not $profileItem) {
        return
    }

    $profile = $profileItem.Data
    if ($profile.Apps.Count -eq 0) {
        Show-AdobeInfoScreen `
            -Title "Экспорт невозможен" `
            -Lines @(
                "У пользователя {0} не найдены поддерживаемые каталоги Adobe." -f $profile.UserName,
                "Проверьте, что приложения хотя бы раз запускались под этим профилем."
            )
        return
    }

    $selectedApps = @(Select-AdobeAppsForExport -Profile $profile)
    if ($selectedApps.Count -eq 0) {
        return
    }

    $confirm = Show-AdobeConfirmMenu `
        -Title "Экспорт: подтверждение" `
        -DetailLines @(
            "Действие: экспорт",
            "Профиль: {0}" -f $profile.UserName,
            "Путь профиля: {0}" -f $profile.ProfilePath,
            "Приложения: {0}" -f (($selectedApps | ForEach-Object { $_.DisplayName }) -join ", "),
            "Каталог экспорта: {0}" -f (Get-AdobeSettingsExportRoot)
        )

    if (-not $confirm -or $confirm.Id -ne "run") {
        return
    }

    if (-not (Confirm-AdobeElevationIfNeeded -TargetProfile $profile -Action "export" -SelectedAppKeys (@($selectedApps | ForEach-Object { $_.Key })))) {
        return
    }

    Clear-Host
    Write-Host "Экспорт запущен..." -ForegroundColor Cyan
    Write-Host ""
    $result = Export-AdobeSettingsPackage -SourceProfile $profile -SelectedApps $selectedApps
    Show-AdobeInfoScreen `
        -Title "Экспорт завершён" `
        -Lines @(
            "Папка пакета: {0}" -f $result.PackageRoot,
            "Архив: {0}" -f $result.ZipPath,
            "Manifest: {0}" -f $result.ManifestPath,
            "Журнал: {0}" -f $result.LogPath
        )
}

function Invoke-AdobeImportFlow {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Profiles
    )

    $packagePath = Select-AdobeImportPackage
    if ([string]::IsNullOrWhiteSpace($packagePath)) {
        return
    }

    $package = $null
    try {
        $package = Open-AdobeImportPackage -PackagePath $packagePath
        $packageApps = @(Get-AdobePackageAppsFromManifest -Manifest $package.Manifest)
        if ($packageApps.Count -eq 0) {
            Show-AdobeInfoScreen `
                -Title "Импорт невозможен" `
                -Lines @(
                    "В пакете нет поддерживаемых приложений.",
                    "Проверьте manifest.json."
                )
            return
        }

        $profileItem = Show-AdobeSingleSelectMenu `
            -Title "Импорт: выбор целевого профиля" `
            -Items @(Get-AdobeProfileMenuItems -Profiles $Profiles) `
            -DetailScript ${function:Show-AdobeProfileDetails}

        if (-not $profileItem) {
            return
        }

        $targetProfile = $profileItem.Data
        $selectedApps = @(Select-AdobeAppsForImport -PackageApps $packageApps)
        if ($selectedApps.Count -eq 0) {
            return
        }

        $confirm = Show-AdobeConfirmMenu `
            -Title "Импорт: подтверждение" `
            -DetailLines @(
                "Действие: импорт",
                "Пакет: {0}" -f $package.SourcePath,
                "Целевой профиль: {0}" -f $targetProfile.UserName,
                "Путь профиля: {0}" -f $targetProfile.ProfilePath,
                "Приложения: {0}" -f (($selectedApps | ForEach-Object { $_.DisplayName }) -join ", "),
                "Каталог резервных копий: {0}" -f (Get-AdobeSettingsBackupRoot)
            )

        if (-not $confirm -or $confirm.Id -ne "run") {
            return
        }

        if (-not (Confirm-AdobeElevationIfNeeded `
            -TargetProfile $targetProfile `
            -Action "import" `
            -PackagePath $package.SourcePath `
            -SelectedAppKeys (@($selectedApps | ForEach-Object { $_.Key })))) {
            return
        }

        if (-not (Ensure-AdobeAppsClosedBeforeImport -SelectedApps $selectedApps)) {
            return
        }

        Clear-Host
        Write-Host "Импорт запущен..." -ForegroundColor Cyan
        Write-Host ""
        $result = Import-AdobeSettingsPackage `
            -TargetProfile $targetProfile `
            -Package $package `
            -SelectedAppKeys (@($selectedApps | ForEach-Object { $_.Key }))

        Show-AdobeInfoScreen `
            -Title "Импорт завершён" `
            -Lines @(
                "Резервная копия: {0}" -f $result.BackupRoot,
                "Manifest резервной копии: {0}" -f $result.BackupManifestPath,
                "Журнал: {0}" -f $result.LogPath
            )
    }
    finally {
        if ($package) {
            Close-AdobeImportPackage -Package $package
        }
    }
}

function Invoke-AdobeResumeFlow {
    $profiles = @(Get-AdobeUserProfiles)
    $profile = $profiles | Where-Object { $_.ProfilePath -eq $ResumeUserProfilePath } | Select-Object -First 1
    if (-not $profile) {
        throw "Не найден профиль для возобновления: $ResumeUserProfilePath"
    }

    $selectedAppKeys = @()
    if (-not [string]::IsNullOrWhiteSpace($ResumeSelectedAppsCsv)) {
        $selectedAppKeys = @($ResumeSelectedAppsCsv.Split(";") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    switch ($ResumeAction) {
        "export" {
            $apps = @($profile.Apps | Where-Object { $selectedAppKeys -contains $_.Key })
            if ($apps.Count -eq 0) {
                throw "Не удалось восстановить выбранные приложения для экспорта."
            }

            Clear-Host
            Write-Host "Экспорт запущен после повышения прав..." -ForegroundColor Cyan
            Write-Host ""
            $result = Export-AdobeSettingsPackage -SourceProfile $profile -SelectedApps $apps
            Show-AdobeInfoScreen `
                -Title "Экспорт завершён" `
                -Lines @(
                    "Папка пакета: {0}" -f $result.PackageRoot,
                    "Архив: {0}" -f $result.ZipPath,
                    "Manifest: {0}" -f $result.ManifestPath,
                    "Журнал: {0}" -f $result.LogPath
                )
        }
        "import" {
            $package = $null
            try {
                $package = Open-AdobeImportPackage -PackagePath $ResumePackagePath
                $packageApps = @(Get-AdobePackageAppsFromManifest -Manifest $package.Manifest | Where-Object { $selectedAppKeys -contains $_.Key })
                if (-not (Ensure-AdobeAppsClosedBeforeImport -SelectedApps $packageApps)) {
                    return
                }

                Clear-Host
                Write-Host "Импорт запущен после повышения прав..." -ForegroundColor Cyan
                Write-Host ""
                $result = Import-AdobeSettingsPackage -TargetProfile $profile -Package $package -SelectedAppKeys $selectedAppKeys
                Show-AdobeInfoScreen `
                    -Title "Импорт завершён" `
                    -Lines @(
                        "Резервная копия: {0}" -f $result.BackupRoot,
                        "Manifest резервной копии: {0}" -f $result.BackupManifestPath,
                        "Журнал: {0}" -f $result.LogPath
                    )
            }
            finally {
                if ($package) {
                    Close-AdobeImportPackage -Package $package
                }
            }
        }
        default {
            throw "Неизвестное действие для возобновления: $ResumeAction"
        }
    }
}

function Start-AdobeSettingsManager {
    if (-not (Test-AdobeUiHostSupported)) {
        Show-AdobeHostCompatibilityError
        return
    }

    if (-not [string]::IsNullOrWhiteSpace($ResumeAction)) {
        Invoke-AdobeResumeFlow
        return
    }

    $profiles = @(Get-AdobeUserProfiles)
    if ($profiles.Count -eq 0) {
        Show-AdobeInfoScreen `
            -Title "Профили не найдены" `
            -Lines @(
                "Не удалось найти пригодные пользовательские профили Windows в C:\Users."
            )
        return
    }

    while ($true) {
        $action = Show-AdobeSingleSelectMenu `
            -Title "Adobe Settings Manager" `
            -Items @(Get-AdobeActionMenuItems) `
            -DetailScript {
                param($item)
                @(
                    $item.Description,
                    "",
                    "Поддерживаемые приложения:",
                    "- Premiere Pro",
                    "- After Effects",
                    "- Photoshop",
                    "- Adobe Media Encoder"
                )
            } `
            -Footer "Стрелки: выбор  Enter: открыть  Esc: выход"

        if (-not $action -or $action.Id -eq "exit") {
            break
        }

        switch ($action.Id) {
            "export" { Invoke-AdobeExportFlow -Profiles $profiles }
            "import" { Invoke-AdobeImportFlow -Profiles $profiles }
        }
    }
}

if ($MyInvocation.InvocationName -ne ".") {
    try {
        Start-AdobeSettingsManager
    }
    finally {
        try {
            Clear-Host
        }
        catch {
        }
    }
}
