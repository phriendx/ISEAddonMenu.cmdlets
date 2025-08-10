Function New-MenuItemFromShortcut {
    <#
    .SYNOPSIS
    Adds a menu item to the PowerShell ISE Add-ons menu from a shortcut, script, executable, URL, or PDF.

    .DESCRIPTION
    Processes supported file types and creates dynamic entries in the specified `$ParentMenu` within the ISE Add-ons menu.
    Supports `.lnk`, `.ps1`, `.bat`, `.cmd`, `.exe`, `.url`, and `.pdf` files.

    - `.lnk`: Extracts and runs the target with optional arguments.
    - `.ps1`: Opens the script in ISE with `psedit`.
    - `.bat` / `.cmd`: Runs using `Start-Process`.
    - `.exe`: Runs the executable directly.
    - `.url`: Opens the URL using Microsoft Edge.
    - `.pdf`: Opens the PDF file using Microsoft Edge.

    .PARAMETER ParentMenu
    The parent ISE Add-ons menu object to which the new menu item will be added.
    This is typically retrieved via 'psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus'.

    .PARAMETER Shortcut
    A file object (from `Get-Item`, etc.) with `.FullName` and `.Name` properties.
    Supported file types: .lnk, .ps1, .bat, .cmd, .exe, .url, .pdf

    .INPUTS
    System.Object

    .OUTPUTS
    None

    .EXAMPLE
    $menu = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Tools", $null, $null)
    New-MenuItemFromShortcut -ParentMenu $menu -Shortcut (Get-Item "C:\Tools\Doc.pdf")

    Adds a menu item to open a PDF in Microsoft Edge.

    .EXAMPLE
    Add-ISEMenuFromDirectory -ShortcutDirectory "C:\Tools\Shortcuts" -RootMenuName "Tools"

    Adds a "Tools" menu to the ISE Add-ons menu with items from the specified folder and subfolders.

    .EXAMPLE
    Add-ISEMenuFromDirectory -ShortcutDirectory "C:\Tools\Shortcuts" -RootMenuName "Tools"

    Adds a "Tools" menu to the ISE Add-ons menu with items from the specified folder and subfolders.

    .EXAMPLE
    Add-ISEMenuFromDirectory -ShortcutDirectory "C:\ProgramData\Microsoft\Windows\Start Menu\Programs" -RootMenuName "Start Menu"

    Adds a top-level menu using shortcuts found in the system Start Menu directory.

    .NOTES
        - Requires helper function: New-MenuItemFromShortcut.
        - Designed for use only within the PowerShell ISE environment.
        - Uses the `New-MenuItemFromShortcut` function to support many file types (see below).
        - Submenus mirror the directory structure beneath the specified root folder.
        - Supported Types
            Extension	Behavior
            .lnk	    Extracts and runs target with args
            .ps1	    Opens in ISE using psedit
            .bat, .cmd	Runs in normal window
            .exe	    Launches directly
            .pdf	    Opens in Microsoft Edge
            .url	    Parses and opens URL in Edge

    .LINK
        New-MenuItemFromShortcut
        Export-ISEMenuToRegistry
        Add-ISEMenuFromRegistry

    #>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    Param (
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('ParentMenuName')]   # allows passing a string name as well
        [object]$ParentMenu,

        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [object]$Shortcut,

        [Parameter()]
        [switch]$Force   # replace existing child menu item with the same name
    )

    Begin {
        if (-not $Shortcut) {
            Write-Error "Shortcut parameter is required."
            return
        }
        if (-not (Test-Path $Shortcut.FullName)) {
            Write-Error "Shortcut '$($Shortcut.FullName)' does not exist."
            return
        }

        $sh = New-Object -ComObject WScript.Shell

        function Resolve-ParentMenu {
            param(
                [object]$InputParent
            )
            # If caller passed a string, try to find an existing submenu by that name; create if missing.
            if ($InputParent -is [string]) {
                $name = $InputParent
                $existing = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus |
                            Where-Object { $_.DisplayName -eq $name } |
                            Select-Object -First 1
                if ($existing) {
                    Write-Verbose "Using existing parent menu: '$name'."
                    return $existing
                } else {
                    Write-Verbose "Parent menu '$name' not found. Creating it."
                    return $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add($name, $null, $null)
                }
            }

            # If caller passed an ISE submenu object, just use it.
            if ($InputParent -and $InputParent.PSObject.Properties['SubMenus']) {
                return $InputParent
            }

            throw "ParentMenu must be an ISE submenu object or the name (string) of an existing/new parent menu."
        }

        try {
            $script:ResolvedParent = Resolve-ParentMenu -InputParent $ParentMenu
        } catch {
            Write-Error $_.Exception.Message
            return
        }
    }

    Process {
        Write-Verbose "Processing shortcut: $($Shortcut.FullName)"
        $appName    = [System.IO.Path]::GetFileNameWithoutExtension($Shortcut.Name)
        $extension  = [System.IO.Path]::GetExtension($Shortcut.Name).ToLowerInvariant()
        $scriptBlock = $null

        switch ($extension) {
            ".lnk" {
                $shortcutObject = $sh.CreateShortcut($Shortcut.FullName)
                $targetPath = $shortcutObject.TargetPath
                $arguments  = if ($shortcutObject.Arguments) { $shortcutObject.Arguments.Replace('"','`"') } else { $null }

                $scriptBlock = if ($arguments) {
                    [scriptblock]::Create("Start-Process -FilePath `"$targetPath`" -ArgumentList `"$arguments`"")
                } else {
                    [scriptblock]::Create("Start-Process -FilePath `"$targetPath`"")
                }
            }

            ".ps1" {
                $scriptBlock = [scriptblock]::Create("psedit -File `"$($Shortcut.FullName)`"")
            }

            { $_ -eq ".bat" -or $_ -eq ".cmd" } {
                $scriptBlock = [scriptblock]::Create("Start-Process -FilePath `"$($Shortcut.FullName)`" -WindowStyle Normal")
            }

            ".exe" {
                $scriptBlock = [scriptblock]::Create("Start-Process -FilePath `"$($Shortcut.FullName)`"")
            }

            ".pdf" {
                $fileUri = [System.Uri]::EscapeUriString(("file:///" + $Shortcut.FullName.Replace('\', '/')))
                $edgePath = "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe"
                if (-not (Test-Path $edgePath)) {
                    $edgePath = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
                }
                if (-not (Test-Path $edgePath)) {
                    Write-Warning "Microsoft Edge not found. Cannot open PDF: $($Shortcut.FullName)"
                    return
                }
                $scriptBlock = [scriptblock]::Create("Start-Process -FilePath `"$edgePath`" -ArgumentList `"$fileUri`"")
            }

            ".url" {
                try {
                    $url = Get-Content -LiteralPath $Shortcut.FullName |
                           Where-Object { $_ -match '^URL=' } |
                           ForEach-Object { $_ -replace '^URL=', '' } |
                           Select-Object -First 1
                    if ($url) {
                        $scriptBlock = [scriptblock]::Create("Start-Process -FilePath 'microsoft-edge:$url'")
                    } else {
                        Write-Warning "URL file does not contain a valid URL: $($Shortcut.FullName)"
                        return
                    }
                } catch {
                    Write-Warning "Failed to parse URL: $_"
                    return
                }
            }

            default {
                Write-Error "Unsupported file type '$extension'. Supported: .lnk, .ps1, .bat, .cmd, .exe, .url, .pdf"
                return
            }
        }

        $functionName = $appName  # use .Replace(' ','') if desired
        $parentLabel  = if ($script:ResolvedParent.DisplayName) { $script:ResolvedParent.DisplayName } else { '<unnamed>' }
        $targetLabel  = "$parentLabel\$functionName"

        # Prevent duplicate child items unless -Force
        $existingChild = $script:ResolvedParent.SubMenus |
                         Where-Object { $_.DisplayName -eq $functionName } |
                         Select-Object -First 1

        if ($existingChild -and -not $Force) {
            Write-Verbose "Menu item '$functionName' already exists under '$parentLabel'. Skipping (use -Force to replace)."
            return
        }

        $action = if ($existingChild) { 'Replace ISE Add-ons menu item' } else { 'Add ISE Add-ons menu item' }

        if ($PSCmdlet.ShouldProcess($targetLabel, $action)) {
            if ($existingChild -and $Force) {
                # Best-effort removal if supported by this collection (works in ISE)
                try {
                    [void]$script:ResolvedParent.SubMenus.Remove($existingChild)
                    Write-Verbose "Removed existing menu item '$functionName' under '$parentLabel'."
                } catch {
                    Write-Warning "Could not remove existing menu item '$functionName'. Re-adding may cause duplication."
                }
            }

            $nsb = $ExecutionContext.InvokeCommand.NewScriptBlock($scriptBlock)
            $null = $script:ResolvedParent.SubMenus.Add($functionName, $nsb, $null)
            Write-Verbose "Added menu item '$functionName' under '$parentLabel'."
        } else {
            Write-Verbose "WhatIf: would add/replace menu item '$functionName' under '$parentLabel'."
        }
    }
}