function Add-ISEMenuFromRegistry {
    <#
    .SYNOPSIS
        Adds a custom menu to the PowerShell ISE Add-ons menu using definitions stored in the Windows Registry.

    .DESCRIPTION
        This function reads a set of menu item definitions from a registry key (previously written using Export-ISEMenuToRegistry)
        and dynamically adds those items to the PowerShell ISE Add-ons menu.

        Each registry value represents a command, where:
          - The value name is the display name of the menu item.
          - The value data is the script block code.
          - An optional companion value `<Name>_ShKey` can define a shortcut key.

    .PARAMETER RegMenuName
        The name of the ISE Add-ons root menu to create or populate (e.g., "DevTools" or "MyScripts").

    .PARAMETER RegKey
        The full registry path where the menu item definitions are stored.
        Example: 'HKCU:\Software\MyCompany\ISEMenus\DevTools'

    .EXAMPLE
        Add-ISEMenuFromRegistry -RegMenuName "DevTools" -RegKey "HKCU:\Software\MyCompany\ISEMenus\DevTools"

        Adds the menu "DevTools" to the ISE Add-ons menu and populates it with items defined in the specified registry key.

    .EXAMPLE
        Add-ISEMenuFromRegistry -RegMenuName "AdminTools" -RegKey "HKCU:\Software\ISEMenus\AdminTools"

        Adds an "AdminTools" menu using definitions exported earlier.

    .NOTES
        - Only works within PowerShell ISE.
        - Script blocks are reconstructed at runtime using `InvokeCommand.NewScriptBlock`.
        - Shortcut keys are optional and must be stored in values named like `<FunctionName>_ShKey`.
        - If the specified registry key does not exist, a popup alert is shown and the function exits.

    .LINK
        Export-ISEMenuToRegistry
        Get-ISEMenuDefinition

    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $RegMenuName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $RegKey
    )

    begin {
        if (-not (Test-Path $RegKey)) {
            $shell = New-Object -ComObject WScript.Shell
            $shell.Popup("Add-ISEMenuFromRegKey failed. Registry key not found:`n$RegKey", 0, "ISEMenuCmdlets", 64)
            break
        }
    }

    process {
        $menuRoot = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add($RegMenuName, $null, $null)

        $valueNames = (Get-Item $RegKey).Property | Where-Object { -not ($_ -like "*_ShKey") }

        foreach ($name in $valueNames) {
            $scriptValue = (Get-ItemProperty -Path $RegKey -Name $name).$name

            if (-not $scriptValue) {
                Write-Warning "Value '$name' is empty or not found in $RegKey. Skipping."
                continue
            }

            $shortcutName = "${name}_ShKey"
            $shortcutValue = (Get-ItemProperty -Path $RegKey -Name $shortcutName -ErrorAction SilentlyContinue).$shortcutName

            $scriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock($scriptValue)
            $displayName = $name
            $shortcutKey = if ($shortcutValue) { $shortcutValue } else { $null }

            $null = $menuRoot.SubMenus.Add($displayName, $scriptBlock, $shortcutKey)
        }
    }
}