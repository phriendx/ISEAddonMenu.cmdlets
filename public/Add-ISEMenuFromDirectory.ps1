Function Add-ISEMenuFromDirectory {
    <#
    .SYNOPSIS
        Adds PowerShell ISE menu items by scanning a directory of shortcuts and scripts.

    .DESCRIPTION
        This function builds a hierarchical PowerShell ISE Add-ons menu by scanning a directory (and its subdirectories)
        for `.lnk` and `.ps1` files. Each shortcut or script is added as a clickable menu item.

        Subdirectories become nested submenus, and menu items are constructed using the `New-MenuItemFromShortcut` function,
        which supports both launching executables and editing PowerShell scripts.

    .PARAMETER ShortcutDirectory
        The root directory to scan for shortcut (`.lnk`) or PowerShell script (`.ps1`) files.
        Defaults to the Start Menu programs path:
        `C:\ProgramData\Microsoft\Windows\Start Menu\Programs`

    .PARAMETER RootMenuName
        The name of the top-level ISE Add-ons menu item. If not specified, the leaf folder name of the directory is used.

    .PARAMETER ParentMenu
        (Used internally during recursion.) Represents the parent ISE menu object under which submenus are created.

    .PARAMETER recursing
        (Used internally.) Indicates whether the function is being invoked recursively to build nested submenus.

    .EXAMPLE
        Add-ISEMenuFromDirectory -ShortcutDirectory "C:\Tools\Shortcuts" -RootMenuName "Tools"

        Adds a "Tools" menu to the ISE Add-ons menu with items from the specified folder and subfolders.

    .EXAMPLE
        Add-ISEMenuFromDirectory

        Adds a top-level menu using shortcuts found in the system Start Menu directory.

    .NOTES
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

    [cmdletbinding()]
    Param (
        [Parameter()][ValidateNotNullOrEmpty()]
        [string]$ShortcutDirectory = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs",

        [Parameter()][ValidateNotNullOrEmpty()]
        [string]$RootMenuName,

        [Parameter()]
        [object]$ParentMenu, 
        
        [Parameter()]
        [switch]$recursing
    )
    
    Begin {
        If (-Not (Test-Path $ShortcutDirectory)) {
            $shell = New-Object -ComObject Wscript.Shell
            $Shell.Popup("Directory not found: $ShortcutDirectory", 0, "ISEAddonMenu.cmdlets: Missing Directory", 64)
            return
        }

        $Folders = Get-ChildItem $ShortcutDirectory -Directory
        $BaseShortcuts = Get-ChildItem $ShortcutDirectory -File
    }
    
    Process {
        # Create root or use parent menu
        $menuParent = if (-not $recursing) {
            If (-not $RootMenuName) {
                $RootMenuName = Split-Path -Path $ShortcutDirectory -Leaf
            }
            $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add($RootMenuName, $null, $null)
        } else {
            $ParentMenu
        }

        # Process subfolders
        foreach ($Folder in $Folders) {
            $menuName = Split-Path -Path $Folder.FullName -Leaf
            If ($menuParent) {$subMenu = $menuParent.Submenus.Add($menuName, $null, $null)}

            $items = Get-ChildItem -Path $Folder.FullName

            foreach ($item in $items) {
                If ($Item.PSIsContainer) {
                    # Recurse into subdirectory
                    Add-ISEMenuFromDirectory -ShortcutDirectory $item.FullName -ParentMenu $subMenu -recursing
                } else {
                    # Add shortcut
                    New-MenuItemFromShortcut -ParentMenu $subMenu -Shortcut $item 
                }
            }

        }

        # Add top-level shortcuts (non-folder files)
        foreach ($Shortcut in $Baseshortcuts) {
            New-MenuItemFromShortcut -ParentMenu $menuParent -Shortcut $Shortcut
        } 
    }
}