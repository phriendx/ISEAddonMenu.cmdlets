function Launch {
    <#
    .SYNOPSIS
        Launches an application using details from a CSV file or lists available applications.

    .DESCRIPTION
        This function reads a list of applications from a CSV file and either:
        - Launches a specified application by name using an associated path and optional command-line arguments.
        - Lists all available applications defined in the CSV file.

        The CSV file must contain at least two columns:
            - App: The friendly name used to identify the application.
            - AppPath: The full file path to the application's executable.

    .PARAMETER Application
        The name of the application to launch. Must match the "App" field in the CSV.

    .PARAMETER AppListFile
        The path to the CSV file containing the application list.

    .PARAMETER AppCmd
        Optional command-line arguments to pass when launching the application.

    .PARAMETER List
        If specified, lists all applications defined in the CSV file instead of launching one.

    .EXAMPLE
        Launch -Application "Notepad" -AppListFile "C:\scripts\applist.csv"

        Launches Notepad using the path defined in the applist.csv file.

    .EXAMPLE
        Launch -Application "Putty" -AppListFile "C:\apps\launchlist.csv" -AppCmd "-load MySession"

        Launches PuTTY with the "-load MySession" command-line argument.

    .EXAMPLE
        Launch -AppListFile "C:\scripts\applist.csv" -List

        Lists all available applications defined in the specified CSV file.

    .NOTES
        The CSV file should be structured like:
            App,AppPath
            Notepad,C:\Windows\System32\notepad.exe
            PuTTY,C:\Tools\putty.exe

        If the specified application is not found in the list, the function attempts to use the value of -Application as a direct path.

    #>
    [cmdletbinding()]
    Param (
        [Parameter()] [ValidateNotNullOrEmpty()]
        [string] $Application,            
    
        [Parameter()] [ValidateNotNullOrEmpty()]
        [string] $AppListFile, 
            
        [Parameter()] [ValidateNotNullOrEmpty()]
        [string] $AppCmd,

        [Parameter()] [ValidateNotNullOrEmpty()]
        [switch] $List        
                
    )
    
    Begin {
        $AppList = Import-Csv -path $AppListFile
    }    
    
    Process {
        Write-output "App: $Application - Command: $AppCmd"
    
        $Apps = @()
    
        ForEach ($App in $Applist) {
            If ($List) {
                $Apps += $App
            } Else {
                If ($Application -ieq $App.App) {
                    $AppPath = $App.AppPath
                }        
            }    
        }
        
        If ($List) {
            Write-Output $Apps | Sort-Object -property @{Expression="App";Ascending=$true}
            Break
        }
        
        If (!$AppPath) {$AppPath = $Application} 
            
        If ($AppCmd) {
            Start-Process -FilePath $AppPath -ArgumentList $AppCmd -PassThru | out-null #-workingdirectory $($AppPath.SubString(0,$AppPath.LastIndexOf("\")))
        } Else {
            Start-Process -FilePath $AppPath -PassThru | out-null
        }    
    }
    
    End {}
}