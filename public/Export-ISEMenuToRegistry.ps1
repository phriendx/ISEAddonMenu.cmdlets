function Export-ISEMenuToRegistry {
    <#
    .SYNOPSIS
        Exports PowerShell ISE menu items to the Windows Registry for later retrieval or reimport.

    .DESCRIPTION
        This function takes an array of menu command definitions (typically from a function like `Get-ISEMenuDefinition`)
        and writes them into a specified registry path. Each command's script block and optional shortcut key are saved
        as individual registry values under a menu-named subkey.

        Useful for persisting custom ISE menus across sessions, systems, or deployments.

    .PARAMETER RegMenuName
        The name of the menu being exported. This is used as the subkey name under the given registry path.

    .PARAMETER RegistryPath
        The base registry path where the menu definitions should be stored. 
        Example: 'HKCU:\Software\MyCompany\ISEMenus'

    .PARAMETER Commands
        An array of command definitions, each with at least:
            - FunctionName: The name used for the ISE menu entry.
            - Scriptblock: The code to execute.
            - Shortcut: (Optional) A string representing a shortcut key combination.

    .PARAMETER Overwrite
        If specified, the function will clear all existing registry values under the target subkey before writing new values.

    .EXAMPLE
        Export-ISEMenuToRegistry -RegMenuName 'MyTools' -RegistryPath 'HKCU:\Software\ISEMenus' -Commands $myCommands -Overwrite

        Exports the provided `$myCommands` array to the registry under:
        HKCU:\Software\ISEMenus\MyTools

    .EXAMPLE
        $cmds = Get-ISEMenuDefinition -RootMenuNameFilter 'DevTools*' -IncludeShortcutKeys "yes"
        Export-ISEMenuToRegistry -RegMenuName 'DevTools' -RegistryPath 'HKCU:\Software\MyISEMenus' -Commands $cmds

        Retrieves existing DevTools menu items from ISE and saves them to the registry.

    .NOTES
        Shortcut keys (if present) are saved using an additional registry value with the name `<FunctionName>_ShKey`.

        Menu definitions exported this way can be reimported with a complementary function like `Import-ISEMenuFromRegistry`.

    .LINK
        Get-ISEMenuDefinition
        Import-ISEMenuFromRegistry

    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $RegMenuName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $RegistryPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [array] $Commands,

        [Parameter()]
        [switch] $Overwrite
    )

    # Construct the full path
    $fullRegistryPath = Join-Path -Path $RegistryPath -ChildPath $RegMenuName

    # Create the subkey if it doesn't exist
    IF (-not (Test-Path $fullRegistryPath)) {
        try {
            New-Item -Path $fullRegistryPath -Force | out-null
        } catch {
            throw "Failed to create registry key: $fullRegistryPath"
        }
    }


    if ($Overwrite) {
        # Clear existing values under that subkey
        try {
            $existingProps = (Get-Item -Path $fullRegistryPath).Property
            foreach ($Prop in $existingProps) {
                Remove-ItemProperty -Path $fullRegistryPath -Name $Prop -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Warning "Failed to clear old values under: $fullRegistryPath"
        }
    }

    # Write each command to registry
    foreach ($Command in $Commands) {
        $funcName = $Command.FunctionName
        $ScriptText = $Command.Scriptblock.ToString()
        $shortcut = $Command.Shortcut

        Set-ItemProperty -Path $fullRegistryPath -Name $funcName -Value $ScriptText

        If ($shortcut) {
            Set-ItemProperty -Path $fullRegistryPath -Name "$funcName`_ShKey" -Value $shortcut
        }
    }

    Write-Verbose "Exported $($Commands.Count) ISE menu items to: $fullRegistryPath"
}