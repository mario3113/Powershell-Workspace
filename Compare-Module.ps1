#requires -version 5.0

# https://gist.github.com/jdhitsolutions/7217ed9293f18e8d454e3f88ecb38b67

Function Compare-Module {
    <#
.Synopsis
Compare module versions.
.Description
Use this command to compare module versions between what is installed against an online repository like the PSGallery. Results will be automatically sorted by module name.
.Parameter Name
The name of a module to check. Wildcards are permitted.
.Notes
Version: 1.2

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

.Example
PS C:\> Compare-Module | Where-objject {$_.UpdateNeeded}

Name             : Azure
OnlineVersion    : 1.5.1
InstalledVersion : 1.0.4
PublishedDate    : 6/27/2016 6:50:11 PM
UpdateNeeded     : True

Name             : Azure.Storage
OnlineVersion    : 1.1.4
InstalledVersion : 1.0.4
PublishedDate    : 6/27/2016 6:48:07 PM
UpdateNeeded     : True

Name             : AzureRM
OnlineVersion    : 1.5.1
InstalledVersion : 1.2.0
PublishedDate    : 6/27/2016 7:08:50 PM
UpdateNeeded     : True
...
.Example
PS C:\> Compare-Module | Where UpdateNeeded | Out-Gridview -title "Select modules to update" -outputMode multiple | Foreach { Update-Module $_.name }

Compare modules and send results to Out-Gridview. Use Out-Gridview as an object picker to decide what modules to update.

.Example
PS C:\> compare-module -name xWindows* | format-table


Name                    OnlineVersion InstalledVersion PublishedDate         UpdateNeeded
----                    ------------- ---------------- -------------         ------------
xWindowsEventForwarding 1.0.0.0       1.0.0.0          6/17/2015 9:46:32 PM         False
xWindowsRestore         1.0.0         1.0.0            12/18/2014 4:22:42 AM        False
xWindowsUpdate          2.5.0.0       2.3.0.0          5/18/2016 11:02:47 PM         True

Compare all modules that start with xWindows and display results in a table format.

.Example
PS C:\> get-dscresource cAD* | Select moduleName -Unique | compare-module


Name             : cActiveDirectory
OnlineVersion    : 1.1.1
InstalledVersion : 1.0.1
PublishedDate    : 6/23/2015 9:24:55 PM
UpdateNeeded     : True

Get all DSC Resources that start with cAD and select the corresponding module name. Since the module name will be listed for every resource, get a unique list and pipe that to Compare-Module.
.Link
Find-Module
.Link
Get-Module
.Link
Update-Module
.Inputs
[string]
.Outputs
[PSCustomObject]
#>


    [cmdletbinding()]
    [alias("cmo")]

    Param
    (
        [Parameter(
            Position = 0,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullorEmpty()]
        [Alias("modulename")]
        [string]$Name,
        [ValidateNotNullorEmpty()]
        [string]$Gallery = "PSGallery"
    )

    Begin {

        Write-Verbose "[BEGIN  ] Starting: $($MyInvocation.Mycommand)"

        $progParam = @{
            Activity         = $MyInvocation.MyCommand
            Status           = "Getting installed modules"
            CurrentOperation = "Get-Module -ListAvailable"
            PercentComplete  = 25
        }

        Write-Progress @progParam

    } #begin

    Process {

        $gmoParams = @{
            ListAvailable = $True
        }
        if ($Name) {
            $gmoParams.Add("Name", $Name)
        }

        $installed = Get-Module @gmoParams

        if ($installed) {

            $progParam.Status = "Getting online modules"
            $progParam.CurrentOperation = "Find-Module -repository $Gallery"
            $progParam.PercentComplete = 50
            Write-Progress @progParam

            $fmoParams = @{
                Repository  = $Gallery
                ErrorAction = "Stop"
            }
            if ($Name) {
                $fmoParams.Add("Name", $Name)
            }
            Try {
                $online = Find-Module @fmoParams
            }
            Catch {
                Write-Warning "Failed to find online module(s). $($_.Exception.message)"
            }
            $progParam.status = "Comparing $($installed.count) installed modules to $($online.count) online modules."
            $progParam.percentComplete = 80
            Write-Progress @progParam

            $data = $online | Where-Object {$installed.name -contains $_.name} |
                Select-Object -property Name,
            @{Name = "OnlineVersion"; Expression = {$_.Version}},
            @{Name = "InstalledVersion"; Expression = {
                    #save the name from the incoming online object
                    $name = $_.Name
                    $installed.Where( {$_.name -eq $name}).Version -join ","}
            },
            PublishedDate,
            @{Name = "UpdateNeeded"; Expression = {
                    $name = $_.Name
                    #there could me multiple versions installed
                    $installedVersions = $installed.Where( {$_.name -eq $name}).Version | Sort-Object
                    foreach ($item in $installedVersions) {
                        If ([version]$_.Version -gt [version]$item) {
                            $result = $True
                        }
                        else {
                            $result = $False
                        }
                    }
                    $result
                }
            } | Sort-Object -Property Name

            $progParam.PercentComplete = 100
            $progParam.Completed = $True
            Write-Progress @progparam

            #write the results to the pipeline
            $data
        }
        else {
            Write-Warning "No local module or modules found"
        }

    } #Progress

    End {
        Write-Verbose "[END    ] Ending: $($MyInvocation.Mycommand)"
    } #end

} #close function
