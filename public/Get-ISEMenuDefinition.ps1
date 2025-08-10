Function Get-ISEMenuDefinition {
    <#
    .SYNOPSIS
        Retrieves a list of custom menu items defined in the PowerShell ISE Add-ons menu.

    .DESCRIPTION
        This function scans the PowerShell ISE Add-ons menu and returns a list of custom submenu items.
        Each item includes the root menu name, display name, script block, and optionally the associated shortcut key.
        It is useful for exporting, auditing, or reusing custom ISE menu definitions.

    .PARAMETER RootMenuNameFilter
        A wildcard filter (default '*') used to limit which root menus are processed.
        Only submenus under root menu names matching this filter will be returned.

    .PARAMETER IncludeShortcutKeys
        If provided, attempts to include readable shortcut key combinations (e.g., Ctrl+Shift+S) for each menu item that has a defined shortcut.
        Leave blank to omit shortcut key details.

    .OUTPUTS
        [PSCustomObject] with the following properties:
            - RootMenuName: The name of the top-level menu under Add-ons
            - FunctionName: The display name of the menu item
            - Scriptblock: The script block that the menu item invokes
            - Shortcut: A string representing the shortcut key combination (if any)

    .EXAMPLE
        Get-ISEMenuDefinition

        Returns all menu items from all Add-ons root menus.

    .EXAMPLE
        Get-ISEMenuDefinition -RootMenuNameFilter 'MyTools*' -IncludeShortcutKeys "yes"

        Returns menu items from root menus starting with "MyTools" and includes shortcut key definitions.

    .NOTES
        This function is intended for use within the PowerShell ISE environment.
        It relies on the `$psISE` object, which is not available in other hosts.

    #>

    [cmdletbinding()]
    Param (
        [Parameter()]
        [string] $RootMenuNameFilter = '*',
            
        [Parameter()]
        [string] $IncludeShortcutKeys
    )

    Add-Type -AssemblyName PresentationCore
    $converter = New-Object System.Windows.Input.KeyGestureConverter

    $menuItems = @()

    $rootMenus = $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus

    foreach ($submenu in $rootMenus) {
        If ($submenu.DisplayName -notlike $RootMenuNameFilter) { continue }

        foreach ($item in $submenu.Submenus) {
            try {
                If (-not $item.Action) { continue }

                $scriptblock = $ExecutionContext.InvokeCommand.NewScriptBlock($Item.Action)

                $entry = @{
                    RootMenuName = $submenu.DisplayName
                    FunctionName = $Item.DisplayName
                    Scriptblock  = $scriptblock
                    Shortcut     = ""
                }

                IF ($IncludeShortcutKeys -and $Item.Shortcut) {
                    $gesture = $item.Shortcut
                    $Entry.Shortcut = $converter.ConvertToString($gesture)
                }

                $menuItems += [pscustomobject]$entry

            } catch {
                Write-Warning "Could not extract script block for menu item '$($item.DisplayName)'"
            }
        }
    }
}