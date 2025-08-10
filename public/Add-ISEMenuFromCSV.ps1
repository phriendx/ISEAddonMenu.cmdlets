function Add-ISEMenuFromCSV {
    <#
    .SYNOPSIS
        Adds a custom ISE Add-ons menu based on application definitions stored in a CSV file.

    .DESCRIPTION
        This function creates a new top-level menu in the PowerShell ISE Add-ons menu and populates it with
        items defined in a CSV file. Each item is wired to call a `launch` function with the appropriate application
        name and the source CSV file path.

        This is useful for dynamic launch menus in the ISE, driven by a maintained CSV list of applications.

        Each menu item calls:
            launch -Application <App> -AppListFile "<CSVFile>"

    .PARAMETER CSVFile
        The full path to the CSV file containing application launch definitions. 
        The CSV must contain at least the columns:
            - App
            - AppFullName

    .PARAMETER CSVMenuName
        The name of the menu to add under the ISE Add-ons menu.

    .EXAMPLE
        Add-ISEMenuFromCSV -CSVFile "C:\Tools\Apps.csv" -CSVMenuName "Quick Launch"

        Adds a "Quick Launch" menu to ISE with launchable items defined in Apps.csv.

    .EXAMPLE
        Add-ISEMenuFromCSV -CSVFile "$env:USERPROFILE\tools\dev_apps.csv" -CSVMenuName "Dev Tools"

        Adds a "Dev Tools" menu based on a CSV file listing development utilities.

    .NOTES
        - Each menu item created will execute the `launch` function, which must be defined in the ISE session.
        - The CSV file must include at least:
            App         : Short name used in `-Application` parameter
            AppFullName : Friendly display name to appear in the ISE menu
        - If the file is not found, a popup message is shown and execution halts.

    .LINK
        Launch
        Export-ISEMenuToRegistry
        Add-ISEMenuFromRegistry

    #>

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]
        [string] $CSVFile,
            
        [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]
        [string] $CSVMenuName
    )

    Begin {
        If (Test-Path $CSVFile) {
            $AppList = Import-Csv -path $CSVFile
        } Else {
            $promptShell = New-Object -ComObject WScript.Shell		
	        $null = $promptShell.popup("Add-ISEMenuFromCSV Failed. CSV file not found: $CSVFile",0,"ISEMenuCmdlets: Missing CSV File",64)
            Break
        }    
    }

    Process {
        $parentAdded = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add($CSVMenuName,$null,$null)

        $AppCmds = @()

        ForEach ($App in $Applist) {
            $OutputObj  = New-Object -Type PSObject
            $OutputObj | Add-Member -MemberType NoteProperty -Name "FunctionName" –Value $App.AppFullName
            $OutputObj | Add-Member -MemberType NoteProperty -Name "Scriptblock" –Value "launch -Application $($App.App) -AppListFile `"$AppListPath`""
            $AppCmds += $OutputObj  
            $outputObj  
        }

        $AppCmds | ForEach-Object {
            $sb=$executioncontext.InvokeCommand.NewScriptBlock($_.Scriptblock)
            $parentAdded.Submenus.Add($_.FunctionName,$sb,$null)
        }
        
    }
}