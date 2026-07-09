<#
.SYNOPSIS
    Pins or unpins applications to or from the Windows Start Menu.
.DESCRIPTION
    This cmdlet programmatically interacts with the Windows Shell Applications folder
    to pin or unpin specified applications to/from the Windows Start Menu. It supports
    both English and German system locales.
.PARAMETER AppName
    An array of application names (as they appear in the Applications folder) to process.
.PARAMETER Unpin
    If specified, the function attempts to unpin the application from the Start Menu.
    If omitted, the function defaults to pinning the application.
.EXAMPLE
    Set-StartPin -AppName "Command Prompt"
    Pins the "Command Prompt" application to the Start Menu.
.EXAMPLE
    Set-StartPin -AppName "Command Prompt", "Calculator" -Unpin
    Unpins both "Command Prompt" and "Calculator" from the Start Menu.
.EXAMPLE
    "Command Prompt", "Calculator" | Set-StartPin -Unpin
    Unpins applications passed via the pipeline.
#>
function Set-StartPin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$AppName,

        [Parameter()]
        [switch]$Unpin
    )

    begin {
        Write-Verbose "Initializing shell application namespace..."
        try {
            $shell = New-Object -ComObject Shell.Application -ErrorAction Stop
            # CLSID for the virtual Applications folder (shell:AppsFolder)
            $appsFolder = $shell.NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}')
            
            if (-not $appsFolder) {
                throw "Could not retrieve the Windows Applications namespace."
            }
        }
        catch {
            Write-Error "Initialization failed: $_"
            return
        }

        # Define localized verb match strings and action descriptions
        if ($Unpin.IsPresent) {
            $actionRegex = 'Von "Start" lösen|Unpin from Start'
            $actionDescription = 'unpinned from'
        }
        else {
            $actionRegex = 'An "Start" anheften|Pin to Start'
            $actionDescription = 'pinned to'
        }
    }

    process {
        # Retrieve all available items in the application folder
        $allApps = $appsFolder.Items()

        $foundApps = [System.Collections.Generic.List[Object]]::new()
        $notFoundApps = [System.Collections.Generic.List[string]]::new()

        # Identify found and missing apps
        foreach ($name in $AppName) {
            # Find matching items (case-insensitive)
            $matchedApp = $allApps | Where-Object { $_.Name -eq $name }
            
            if ($matchedApp) {
                $foundApps.Add($matchedApp)
            }
            else {
                $notFoundApps.Add($name)
            }
        }

        # Handle missing apps early and emit as non-terminating errors
        if ($notFoundApps.Count -gt 0) {
            Write-Error "The following application(s) were not found: $($notFoundApps -join ', ')"
        }

        # Process identified apps
        foreach ($app in $foundApps) {
            try {
                Write-Verbose "Retrieving verbs for application: $($app.Name)"
                $verbs = $app.Verbs()
                
                # Normalize verb name by removing the accelerator prefix (&)
                $targetVerb = $verbs | Where-Object { ($_.Name -replace '&', '') -match $actionRegex }

                if ($targetVerb) {
                    Write-Verbose "Executing action for '$($app.Name)'"
                    $targetVerb.DoIt()
                    Write-Output "App '$($app.Name)' successfully $actionDescription Start."
                }
                else {
                    Write-Warning "App '$($app.Name)' is already $actionDescription Start or the action is not supported."
                }
            }
            catch {
                Write-Error "Failed to change pin state for '$($app.Name)': $_"
            }
        }
    }

    end {
        Write-Verbose "Script execution completed."
    }
}
