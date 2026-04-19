function Test-AdobeUiHostSupported {
    if ($Host.Name -like "*ISE*") {
        return $false
    }

    try {
        $null = $Host.UI.RawUI.WindowSize
        return $true
    }
    catch {
        return $false
    }
}

function Get-AdobeUiGlyphSet {
    param(
        [switch]$ForceAscii
    )

    if ($ForceAscii) {
        return [PSCustomObject]@{
            TopLeft = "+"
            TopRight = "+"
            BottomLeft = "+"
            BottomRight = "+"
            Horizontal = "-"
            Vertical = "|"
            Pointer = ">"
        }
    }

    try {
        if ([Console]::OutputEncoding.WebName -match "utf") {
            return [PSCustomObject]@{
                TopLeft = [char]0x250C
                TopRight = [char]0x2510
                BottomLeft = [char]0x2514
                BottomRight = [char]0x2518
                Horizontal = [char]0x2500
                Vertical = [char]0x2502
                Pointer = [char]0x25B6
            }
        }
    }
    catch {
    }

    return [PSCustomObject]@{
        TopLeft = "+"
        TopRight = "+"
        BottomLeft = "+"
        BottomRight = "+"
        Horizontal = "-"
        Vertical = "|"
        Pointer = ">"
    }
}

function Get-AdobeUiWidth {
    try {
        $width = [Console]::WindowWidth
        if ($width -lt 60) {
            return 60
        }

        if ($width -gt 110) {
            return 110
        }

        return $width - 2
    }
    catch {
        return 80
    }
}

function Get-AdobeUiPageSize {
    param(
        [int]$ReserveLines = 12
    )

    try {
        $height = [Console]::WindowHeight
        $pageSize = $height - $ReserveLines
        if ($pageSize -lt 5) {
            return 5
        }

        if ($pageSize -gt 12) {
            return 12
        }

        return $pageSize
    }
    catch {
        return 8
    }
}

function Format-AdobeUiLine {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [int]$Width
    )

    $effectiveWidth = [Math]::Max($Width, 4)
    if ($Text.Length -gt $effectiveWidth) {
        if ($effectiveWidth -le 1) {
            return $Text.Substring(0, $effectiveWidth)
        }

        return $Text.Substring(0, $effectiveWidth - 1) + [char]0x2026
    }

    return $Text.PadRight($effectiveWidth)
}

function New-AdobeMenuItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [string]$Description,

        [object]$Data
    )

    [PSCustomObject]@{
        Id = $Id
        Label = $Label
        Description = $Description
        Data = $Data
    }
}

function Write-AdobeUiBoxLine {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [int]$InnerWidth,

        [Parameter(Mandatory = $true)]
        [object]$Glyphs,

        [switch]$Highlighted
    )

    Write-Host $Glyphs.Vertical -NoNewLine
    $line = Format-AdobeUiLine -Text $Text -Width $InnerWidth
    if ($Highlighted) {
        Write-Host $line -NoNewLine -ForegroundColor $Host.UI.RawUI.BackgroundColor -BackgroundColor $Host.UI.RawUI.ForegroundColor
    }
    else {
        Write-Host $line -NoNewLine
    }
    Write-Host $Glyphs.Vertical
}

function Get-AdobeUiRepeatedGlyph {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Glyph,

        [Parameter(Mandatory = $true)]
        [int]$Count
    )

    return ([string]$Glyph) * [Math]::Max(0, $Count)
}

function Get-AdobeMenuViewport {
    param(
        [int]$Index,
        [int]$TopIndex,
        [int]$ItemCount,
        [int]$PageSize
    )

    if ($ItemCount -le $PageSize) {
        return 0
    }

    if ($Index -lt $TopIndex) {
        return $Index
    }

    if ($Index -ge ($TopIndex + $PageSize)) {
        return [Math]::Max(0, $Index - $PageSize + 1)
    }

    return $TopIndex
}

function Show-AdobeInfoScreen {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [string[]]$Lines,

        [string]$Footer = "Нажмите любую клавишу для продолжения."
    )

    Clear-Host
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ""
    foreach ($line in $Lines) {
        Write-Host $line
    }
    Write-Host ""
    Write-Host $Footer -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Show-AdobeTextPrompt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [string]$DefaultValue
    )

    Clear-Host
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ""
    if ([string]::IsNullOrWhiteSpace($DefaultValue)) {
        return Read-Host $Prompt
    }

    return Read-Host ("{0} [{1}]" -f $Prompt, $DefaultValue)
}

function Show-AdobeSingleSelectMenu {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [object[]]$Items,

        [scriptblock]$DetailScript,

        [string]$Footer = "Стрелки: выбор  Enter: открыть  Esc: назад"
    )

    if ($Items.Count -eq 0) {
        return $null
    }

    $glyphs = Get-AdobeUiGlyphSet
    $contentWidth = Get-AdobeUiWidth
    $innerWidth = $contentWidth - 2
    $pageSize = Get-AdobeUiPageSize
    $index = 0
    $topIndex = 0

    try {
        [Console]::CursorVisible = $false
    }
    catch {
    }

    while ($true) {
        $topIndex = Get-AdobeMenuViewport -Index $index -TopIndex $topIndex -ItemCount $Items.Count -PageSize $pageSize
        $currentItem = $Items[$index]
        $detailLines = @()
        if ($DetailScript) {
            $detailLines = @(& $DetailScript $currentItem)
        }
        elseif (-not [string]::IsNullOrWhiteSpace($currentItem.Description)) {
            $detailLines = @($currentItem.Description)
        }

        Clear-Host
        Write-Host $Title -ForegroundColor Cyan
        Write-Host ""
        Write-Host ($glyphs.TopLeft + (Get-AdobeUiRepeatedGlyph -Glyph $glyphs.Horizontal -Count $innerWidth) + $glyphs.TopRight)
        for ($offset = 0; $offset -lt $pageSize; $offset++) {
            $itemIndex = $topIndex + $offset
            if ($itemIndex -ge $Items.Count) {
                Write-AdobeUiBoxLine -Text "" -InnerWidth $innerWidth -Glyphs $glyphs
                continue
            }

            $item = $Items[$itemIndex]
            $prefix = if ($itemIndex -eq $index) { "{0} " -f $glyphs.Pointer } else { "  " }
            $line = "{0}{1}" -f $prefix, $item.Label
            if (-not [string]::IsNullOrWhiteSpace($item.Description)) {
                $line = "{0} - {1}" -f $line, $item.Description
            }

            Write-AdobeUiBoxLine -Text $line -InnerWidth $innerWidth -Glyphs $glyphs -Highlighted:($itemIndex -eq $index)
        }
        Write-Host ($glyphs.BottomLeft + (Get-AdobeUiRepeatedGlyph -Glyph $glyphs.Horizontal -Count $innerWidth) + $glyphs.BottomRight)

        if ($detailLines.Count -gt 0) {
            Write-Host ""
            foreach ($detailLine in $detailLines | Select-Object -First 8) {
                Write-Host $detailLine
            }
        }

        Write-Host ""
        Write-Host $Footer -ForegroundColor DarkGray

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        switch ($key.VirtualKeyCode) {
            38 {
                if ($index -gt 0) {
                    $index--
                }
            }
            40 {
                if ($index -lt ($Items.Count - 1)) {
                    $index++
                }
            }
            33 {
                $index = [Math]::Max(0, $index - $pageSize)
            }
            34 {
                $index = [Math]::Min($Items.Count - 1, $index + $pageSize)
            }
            36 {
                $index = 0
            }
            35 {
                $index = $Items.Count - 1
            }
            13 {
                return $currentItem
            }
            27 {
                return $null
            }
        }
    }
}

function Show-AdobeMultiSelectMenu {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [object[]]$Items,

        [string[]]$InitiallySelectedIds,

        [scriptblock]$DetailScript,

        [string]$Footer = "Стрелки: выбор  Space: отметить  Enter: подтвердить  Esc: назад"
    )

    if ($Items.Count -eq 0) {
        return @()
    }

    $glyphs = Get-AdobeUiGlyphSet
    $contentWidth = Get-AdobeUiWidth
    $innerWidth = $contentWidth - 2
    $pageSize = Get-AdobeUiPageSize
    $index = 0
    $topIndex = 0
    $selected = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($id in ($InitiallySelectedIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        [void]$selected.Add($id)
    }

    try {
        [Console]::CursorVisible = $false
    }
    catch {
    }

    while ($true) {
        $topIndex = Get-AdobeMenuViewport -Index $index -TopIndex $topIndex -ItemCount $Items.Count -PageSize $pageSize
        $currentItem = $Items[$index]
        $detailLines = @()
        if ($DetailScript) {
            $detailLines = @(& $DetailScript $currentItem)
        }
        elseif (-not [string]::IsNullOrWhiteSpace($currentItem.Description)) {
            $detailLines = @($currentItem.Description)
        }

        Clear-Host
        Write-Host $Title -ForegroundColor Cyan
        Write-Host ""
        Write-Host ($glyphs.TopLeft + (Get-AdobeUiRepeatedGlyph -Glyph $glyphs.Horizontal -Count $innerWidth) + $glyphs.TopRight)
        for ($offset = 0; $offset -lt $pageSize; $offset++) {
            $itemIndex = $topIndex + $offset
            if ($itemIndex -ge $Items.Count) {
                Write-AdobeUiBoxLine -Text "" -InnerWidth $innerWidth -Glyphs $glyphs
                continue
            }

            $item = $Items[$itemIndex]
            $mark = if ($selected.Contains($item.Id)) { "[x]" } else { "[ ]" }
            $prefix = if ($itemIndex -eq $index) { "{0} " -f $glyphs.Pointer } else { "  " }
            $line = "{0}{1} {2}" -f $prefix, $mark, $item.Label
            if (-not [string]::IsNullOrWhiteSpace($item.Description)) {
                $line = "{0} - {1}" -f $line, $item.Description
            }

            Write-AdobeUiBoxLine -Text $line -InnerWidth $innerWidth -Glyphs $glyphs -Highlighted:($itemIndex -eq $index)
        }
        Write-Host ($glyphs.BottomLeft + (Get-AdobeUiRepeatedGlyph -Glyph $glyphs.Horizontal -Count $innerWidth) + $glyphs.BottomRight)

        if ($detailLines.Count -gt 0) {
            Write-Host ""
            foreach ($detailLine in $detailLines | Select-Object -First 8) {
                Write-Host $detailLine
            }
        }

        Write-Host ""
        Write-Host $Footer -ForegroundColor DarkGray

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        switch ($key.VirtualKeyCode) {
            38 {
                if ($index -gt 0) {
                    $index--
                }
            }
            40 {
                if ($index -lt ($Items.Count - 1)) {
                    $index++
                }
            }
            33 {
                $index = [Math]::Max(0, $index - $pageSize)
            }
            34 {
                $index = [Math]::Min($Items.Count - 1, $index + $pageSize)
            }
            36 {
                $index = 0
            }
            35 {
                $index = $Items.Count - 1
            }
            32 {
                if ($selected.Contains($currentItem.Id)) {
                    [void]$selected.Remove($currentItem.Id)
                }
                else {
                    [void]$selected.Add($currentItem.Id)
                }
            }
            13 {
                return @($Items | Where-Object { $selected.Contains($_.Id) })
            }
            27 {
                return $null
            }
        }
    }
}
