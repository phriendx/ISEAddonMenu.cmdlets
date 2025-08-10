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

    [cmdletbinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    Param (
        [Parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [object]$ParentMenu, 
        
        [Parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [object]$Shortcut                    
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

        $sh = New-Object -COM WScript.Shell
    }

    Process {
        Write-Verbose "Processing shortcut: $($Shortcut.FullName)"
        $AppName = [System.IO.Path]::GetFileNameWithoutExtension($Shortcut.Name)
        $Extension = [System.IO.Path]::GetExtension($Shortcut.Name).ToLowerInvariant()
        $Scriptblock = $null

        switch ($Extension) {
            ".lnk" {
                $shortcutObject = $sh.CreateShortcut($Shortcut.FullName)
                $targetPath = $shortcutObject.TargetPath
                $Arguments = if ($shortcutObject.Arguments) { $shortcutObject.Arguments.Replace('"','`"') } else { $null }

                $Scriptblock = if ($Arguments) {
                    [scriptblock]::Create("Start-Process -FilePath `"$targetPath`" -ArgumentList `"$Arguments`"")
                } else {
                    [scriptblock]::Create("Start-Process -FilePath `"$targetPath`"")
                }
            }

            ".ps1" {
                $Scriptblock = [scriptblock]::Create("psedit -File `"$($Shortcut.FullName)`"")
            }

             {$_ -eq ".bat" -or $_ -eq ".cmd"} {
                $Scriptblock = [scriptblock]::Create("Start-Process -FilePath `"$($Shortcut.FullName)`" -WindowStyle Normal")
            }

            ".exe" {
                $Scriptblock = [scriptblock]::Create("Start-Process -FilePath `"$($Shortcut.FullName)`"")
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

                $Scriptblock = [scriptblock]::Create("Start-Process -FilePath `"$edgePath`" -ArgumentList `"$fileUri`"")
            }

            ".url" {
                # Try to read the URL from the file contents
                try {
                    $url = Get-Content $Shortcut.FullName | Where-Object { $_ -match '^URL=' } | ForEach-Object { $_ -replace '^URL=', '' }
                    if ($url) {
                        $Scriptblock = [scriptblock]::Create("Start-Process -FilePath 'microsoft-edge:$url'")
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
                Write-Error "Unsupported file type '$Extension'. Supported: .lnk, .ps1, .bat, .cmd, .exe, .url, .pdf"
                return
            }
        }

        Write-Verbose "Registering menu item: $AppName"

       # $Command = @{
        FunctionName = $AppName #.Replace(" ","")
        #    Scriptblock = $Scriptblock
        #}

        $targetLabel  = if ($ParentMenu -and $ParentMenu.DisplayName) { "$($ParentMenu.DisplayName)\$functionName" } else { $functionName }

        if ($PSCmdlet.ShouldProcess($targetLabel, "Add ISE Add-ons menu item")) {
            #$global:functionname = $Command.FunctionName.ToString()
            #$nsb = $ExecutionContext.InvokeCommand.NewScriptBlock($Command.Scriptblock)
            #$ParentMenu.Submenus.Add($functionname, $nsb, $null) | Out-Null
            $nsb = $ExecutionContext.InvokeCommand.NewScriptBlock($scriptBlock)
            $null = $ParentMenu.Submenus.Add($functionName, $nsb, $null)
            Write-Verbose "Added menu item '$functionName' under '$($ParentMenu.DisplayName)'."
        } else {
            Write-Verbose "WhatIf: would add menu item '$functionName' under '$($ParentMenu.DisplayName)'."
        }
    }
}
