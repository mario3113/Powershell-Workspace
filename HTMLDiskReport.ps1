#requires -version 3.0


Function New-HTMLDiskReport {
<#
.Synopsis
Create a disk utilization report with gradient coloring

.Description
This command will create an HTML report depicting drive utilization through a gradient color bar.

.Parameter Computername
The name(s) of computers to query. They must be running PowerShell 3.0 or later and support CIM queries.
This parameter has an alias of CN.

.Parameter ReportTitle
The HTML title to be for your report. This parameter has an alias of Title.

.Parameter Path
The filename and path for your html report.

.Parameter PreContent
Add any HTML text to insert before the drive utilization table.

.Parameter PostContent
Add any HTML text to insert after the drive utilization table.

.Parameter LogoPath
Specify the path to a PNG or JPG file to use as a logo image. The image will be embedded into the html file.

.Example
PS C:\> New-HTMLDiskReport -passthru


    Directory: C:\Users\Jeff\AppData\Local\Temp


Mode                LastWriteTime         Length Name                                                  
----                -------------         ------ ----                                                  
-a----       10/10/2018   3:58 PM           2493 utilization.htm                                       


Create a report for the local host using default settings.

.Example
PS C:\> get-content c:\work\computers.txt | New-HTMLDiskReport -path c:\work\diskreport.htm

Create a single report for every computer listed in computers.txt. The report will be saved to c:\work\diskreport.htm

.Example
PS C:\> New-HTMLDiskReport -Path c:\work\report.htm -Computername SRV1,SRV2,SRV3 -Precontent "<h3>Company Confidential</h3>" -PostContent "This report is offered as-is. You can verify results with a command like <b>Get-Volume</b>." -logopath c:\scripts\logo.png

Create a report for the specified servers and insert pre- and post-content.

.Notes

Learn more about PowerShell: http://jdhitsolutions.com/blog/essential-powershell-resources/

 .Link
Get-CimInstance

.Link
Get-Volume

.Link
Get-Disk

.Inputs
System.String


#>
    [cmdletbinding(SupportsShouldProcess)]
    [OutputType("None", "System.IO.FileInfo")]

    Param(
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
        [Alias("cn")]
        [string[]]$Computername = $env:Computername,

        [ValidateNotNullorEmpty()]
        [Parameter(HelpMessage = "The report title")]
        [Alias("title")]
        [string]$ReportTitle = "Drive Utilization Report",

        [Parameter(HelpMessage = "The filename and path for the finished HTML report.")]
        [ValidateNotNullorEmpty()]
        [ValidateScript( {
                #get parent
                $parent = Split-Path $_
                if (Test-Path $parent) {
                    $True
                }
                else {
                    Throw "Can't verify part of the file path $_"
                }
            })]
        [string]$Path = "$env:temp\utilization.htm",
        [string[]]$PreContent,
        [string[]]$PostContent,    
        [string]$LogoPath,
        [switch]$Passthru
    )

    Begin {
        
        Write-Verbose "Starting $($MyInvocation.Mycommand)"  

        #define HTML header with style elements. If using a header the title must be inserted here.
        #here strings must be left justified
        $head = @"
<style>
body {
    background-color: #FFFFFF;
    font-family: Tahoma;
    font-size: 12pt;
}

td,
th {
    border: 0px solid black;
    border-collapse: collapse;
}

th {
    color: white;
    background-color: black;
}

table,
tr,
td,
th {
    padding: 2px;
    margin: 0px;
}

tr:nth-child(even) {
    background-color: lightgray
}

table {
    width: 95%;
    margin-left: 10px;
    margin-bottom: 20px;
    table-layout: fixed;
}

.meta {
    width: 25%;
    font-size: 8pt;
    margin-left: 0px;
    table-layout: auto;
}

.top {
    width: 50%;
    margin-left: 0px;
    table-layout: auto;
}

tr.meta {
    background-color: #FFFFFF;
}

.right {
    text-align: right;
    width: 20%;
}

caption {
    background-color: #FFFF66;
    text-align: left;
    font-weight: bold;
    font-size: 14pt;
}

td[tip]:hover {
    color: #ff2283;
    position: relative;
}

td[tip]:hover:after {
    content: attr(tip);
    left: 0;
    top: 100%;
    margin-left: 80px;
    margin-top: 10px;
    width: 400px;
    padding: 3px 8px;
    position: absolute;
    color: #85003a;
    font-family: 'Courier New', Courier, monospace;
    font-size: 10pt;
    background-color: gainsboro;
    white-space: pre-wrap;
}
</style>
<Title>$reportTitle</Title>
"@

        <#
 Define a here string for coloring percentage cells.
 The starting and ending percents will need to provided
 using the -f operator.
 #>
        $gradient = @"
filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, 
StartColorStr=#0A802D, EndColorStr=#FF0011)
background-color: #376C46;
background-image: -mso-linear-gradient(left, #0A802D {0}%, #FF0011 {1}%);
background-image: -ms-linear-gradient(left, #0A802D {0}%, #FF0011 {1}%);
background-image: -moz-linear-gradient(left, #0A802D {0}%, #FF0011 {1}%);
background-image: -o-linear-gradient(left, #0A802D {0}%, #FF0011 {1}%);
background-image: -webkit-linear-gradient(left, #0A802D {0}%, #FF0011 {1}%);
background-image: linear-gradient(left, #0A802D {0}%, #FF0011 {1}%);
color:white;
"@
        If ($LogoPath) {
            if (Test-Path $LogoPath) {
                #insert a graphic
                $ImageBits = [Convert]::ToBase64String((Get-Content $LogoPath -Encoding Byte))
                $ImageFile = Get-Item $LogoPath
                $ImageType = $ImageFile.Extension.Substring(1) #strip off the leading .
                $ImageTag = "<Img src='data:image/$ImageType;base64,$($ImageBits)' Alt='$($ImageFile.Name)' style='float:left' width='120' height='120' hspace=10>"
            }
            else {
                Write-Warning "Could not find image file $LogoPath"
            }
        }
        if ($ImageTag) {
            $top = "<table class='top'><tr><td>$ImageTag</td><td><H1>$ReportTitle</H1></td></table><br>"
        }
        else {
            $top = "<H1>$ReportTitle</H1><br>"
        }

        #define an array to hold HTML fragments
        $fragments = @($top)
        $fragments += $Precontent
        $fragments += "<br><br>"

        #define a parameter hashtable for Write-Progress
        $progParam = @{
            Activity         = $MyInvocation.MyCommand
            Status           = "Gathering disk data"
            CurrentOperation = ""
        }
    } #begin
    Process {

        #get the data for the report
        foreach ($computer in $computername) {
            Write-Verbose "Getting disk information for $Computer"
            $progParam.CurrentOperation = $computer.toUpper()
            Write-Progress @progParam

            Try {
                #create a temporary CIMSession
                $cs = New-CimSession -ComputerName $computer -ErrorAction stop
                #hashtable of parameters to splat to Get-Ciminstance
                
                $paramHash = @{
                    Classname   = "win32_logicaldisk"
                    filter      = "drivetype=3"
                    CimSession  = $cs
                    ErrorAction = "Stop"
                }
                if ($pscmdlet.ShouldProcess($Computer, "Get Disk Information")) {
                    $data = Get-CimInstance @paramHash
                   
                    Write-Verbose "Formatting data"

                    #initialize a hashtable of for phsyical media
                    $hash = @{}
                    #Create a custom object for each drive
                    $drives = foreach ($item in $data) {
                        
                        $Physical = $item | Get-CimAssociatedInstance -ResultClassName Win32_DiskPartition | Get-CimAssociatedInstance -ResultClassName Win32_DiskDrive
                        $hash.Add($item.DeviceID,$physical)

                        $prophash = [ordered]@{
                            Drive         = $item.DeviceID
                            Volume        = $item.VolumeName
                            SizeGB        = $item.size / 1GB -as [int]
                            FreeGB        = "{0:N4}" -f ($item.Freespace / 1GB)
                            PercentFree   = [math]::Round(($item.Freespace / $item.size) * 100, 2)
                        }
                        New-Object PSObject -Property $prophash
                    } #foreach item

                    #convert drive objects to HTML but as an XML document
                    Write-Verbose "Converting to XML"
                    [xml]$html = $drives | ConvertTo-Html -Fragment

                    #add the computer name as the table caption
                    $caption = $html.CreateElement("caption")
                    $html.table.AppendChild($caption) | Out-Null
                    $html.table.caption = $data[0].SystemName
                    $pop = $html.CreateAttribute("title")
                    $pop.value = (Get-Ciminstance -ClassName Win32_OperatingSystem -Property caption -cimsession $cs).caption
                    $html.table.item("caption").attributes.append($pop) | Out-Null

                    #add physical media as a popup for each device
                    for ($i=1; $i -le $html.table.tr.count -1;$i++) {
                        $id = $html.table.tr[$i].ChildNodes[0]."#text"
                        $pop = $html.CreateAttribute("tip")
                        $props = ($hash.Item($id) | Select-Object -property Caption,SerialNumber,FirmwareRevision,Size,InterfaceType,SCSI* | Out-String).trim()     
                        $pop.Value = $props
                        $html.table.tr[$i].ChildNodes[0].Attributes.append($pop) | Out-Null
                    }

                    #go through rows again and add gradient
                    Write-Verbose "Inserting gradient"
                    for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
                        $class = $html.CreateAttribute("style")
                        [int]$start = $html.table.tr[$i].td[-1]
                        #create the gradient using starting and ending values
                        #based on %free
                        $class.value = $Gradient -f $start, (100 - $start)
                        $html.table.tr[$i].ChildNodes[4].Attributes.Append($class) | Out-Null
                    } #for

                    #add the html to the fragments
                    $fragments += $html.InnerXml
                } #should process
            } #Try
            Catch {
                Write-Warning "Failed to get disk information for $computer. $($_.exception.message)"
            } #Catch
            remove-cimsession $cs
        } #foreach computer
    } #process

    End {
        $progParam.currentOperation = "Finalizing report"
        Write-Progress @progParam

        #only proceed if there is data in $fragments
        if ($fragments) {
           
            #add some metadata about this report
            [xml]$metadata = [pscustomobject]@{
                "Report Run" = "$((Get-Date).ToUniversalTime()) UTC"
                "Run By"     = "$($env:USERDOMAIN)\$env:username"
                Originated   = $env:Computername
                Command      = $($myinvocation.invocationname)
                Version      = "2.0"
            } | ConvertTo-html -as List -Fragment
            
            #insert css tags into this table
            $class = $metadata.CreateAttribute("class")
            $class.value = 'meta'
            $metadata.table.Attributes.Append($class) | Out-Null
            for ($i = 0; $i -le $metadata.table.tr.count - 1; $i++) {
                $class = $metadata.CreateAttribute("class")
                $class.value = 'meta'
                $metadata.table.tr[$i].attributes.append($class) | Out-Null
               
                $class = $metadata.CreateAttribute("class")
                $class.value = 'right'
                $metadata.table.tr[$i].item("td").attributes.append($class) | Out-Null        
            }           

            $postcontent += "<br><br>$($metadata.InnerXml)"

            #create the final report
            Write-Verbose "Creating HTML report"
            $paramHash = @{
                head        = $head
                Body        = $($fragments | Out-String)
                PostContent = $PostContent 
            }
            
            if ($pscmdlet.ShouldProcess($path, "Creating HTML file")) {
                ConvertTo-Html @paramHash | Out-File -FilePath $path -Encoding ascii
            }

            Write-Verbose "Report created to $path"

            #if -Passthru write the file object to the pipeline
            if ($Passthru) {
                Get-Item $Path
            }
        }

        Write-Verbose "Ending $($MyInvocation.Mycommand)"

    } #end
} #close New-HTMLDiskReport
