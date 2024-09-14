#requires -version 3.0

Function ConvertTo-HTMLListing {

<#
.Synopsis
Convert text file to HTML listing
.Description
This command will take the contents of a text file and create an HTML document complete with line numbers. The command does not create an actual file. You would need to pipe to Out-File. See examples.
There are options to suppress the line numbers and to skip any blank lines. The command is intended to convert one file at a time although you can pipe a file name to the command. 

The command will attempt to preserve spacing and formatting the best it can. You might have some files where the spacing is slightly off.

Note: Because you are wrapping the existing file in HTML, the converted file will be much larger than the original. You can mitigate this by skipping blank spaces, skipping line numbers and using an external style sheet.
Also do not use the legacy redirection character as this appears to generate a larger file than using Out-File.
.Parameter Path
The path to the text file.
.Parameter CSSUri
You can specify the path or URI to a css stylesheet. Otherwise, a style sheet will be embedded in document head. Using an external style sheet will help keep the file size down.
.Parameter NoFooter
Suppress the default footer.
.Parameter Title
The HTML title.
.Example
PS C:\> ConvertTo-HTMLListing -path c:\scripts\myscript.ps1 | out-file d:\MyScript.htm

Converting a single file.
.Example
PS C:\> dir c:\work\myfile.ps1 | ConvertTo-HTMLListing | Out-file d:\myfile.htm

Converting a single file using a pipelined expression.
.Example
PS C:\> foreach ($file in (dir c:\work\*.txt)) { ConvertTo-HTMLListing $file.fullname -title $file.name | Out-File D:\$($File.basename).htm }

Create an HTML file for each text file in C:\work. Use the file name for the report title.
.Example
PS C:\> ConvertTo-HTMLListing -path c:\work\myfile.txt -SkipBlankLine -NoLineNumber -CSSUri "\\web01\assets\mystyle.css" | out-file '\\Web01\Files$\myfile.htm'

Convert C:\Work\MyFile.txt to an HTML listing, skipping all blank lines, suppressing line numbers and using an external CSS stylesheet.
.Notes
Last Updated: May 6, 2016
Author      : Jeff Hicks 
Version     : 1.3

This command was originally published at:
http://jdhitsolutions.com/blog/powershell/3966/friday-fun-text-to-html/

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/
 
  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************

 .Link
ConvertTo-HTML
#>

[cmdletbinding()]
Param(
[Parameter(
  Position=0,
  Mandatory,
  HelpMessage="Enter the path to the file",
 ValueFromPipeline,
 ValueFromPipelineByPropertyName
 )]
[Alias("PSPath")]
[ValidateScript({Test-Path $_})]
[string]$Path,
[string]$CssUri,
[ValidateNotNullorEmpty()]
[string]$Title = "File Listing",
[switch]$SkipBlankLines,
[switch]$NoLineNumber,
[switch]$NoFooter
)

Begin {
    Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"  

if ($CssUri) {
    Write-Verbose -Message "Using external stylesheet at $CSSUri"
    $myStyle = "<LINK href='$CSSUri' rel='stylesheet' type='text/css'>"
}
else {
   #use a built-in style sheet
  $myStyle = @"
<style>
body { background-color:#FFFFFF;
       font-family:Consolas;
       font-size:10pt;
       white-space:pre;  }
td, th { border:0px solid black; 
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 0px; 
margin: 0px;
white-space:pre;  }
tr:nth-child(odd) {background-color: lightgray}
table { margin-left:25px; }
h2 {
 font-family:Tahoma;
}
.footer 
{ color:green; 
  margin-left:25px; 
  font-family:Tahoma;
  font-size:8pt;
}
</style>
"@
}

Write-Verbose -message "Using title $Title"
#define the html head
$head = @"
<Title>$Title</Title>
$myStyle
"@

} #begin

Process {
    $file = Resolve-Path -Path $Path
    Write-Verbose "Processing $($file.providerpath)"
    $body = "<H2>$($file.providerpath)</H2>"
    $content = Get-Content -Path $file

    if ($SkipBlankLines) {
        #filter out blank lines
        Write-Verbose "Skipping blanks"
        $content = $content | where {$_ -AND $_ -match "\w"}
    }

    Write-Verbose "Converting text to objects"
    $processed = $content | foreach -begin {$i=0} -process { 
      #create a custom object out of each line of text
      $i++
      [pscustomobject]@{Line=$i;"Content"=$_}
    } 

    Write-Verbose "Creating HTML"
    #convert property headings into blanks since they don't need to be displayed
    if ($NoLineNumber) {
        $body+= $processed | ConvertTo-Html -Fragment -Property @{Label="";Expression={$_.Content}}
    }
    else {
        $body+= $processed | ConvertTo-Html -Fragment -Property @{Label="";Expression={$_.Line}},@{Label="";Expression={$_.Content}}
    }

    if ($NoFooter) {
        Write-Verbose -Message "Turning off the default footer"
        $post = "&nbsp;"
    }
    else {
        $post = "<br><div class='footer'>$(Get-Date)</div>"
    }
    #create the HTML output and write to the pipeline
    ConvertTo-HTML -Head $head -Body $body -PostContent $post 
} #process

End {    
    Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
} #end

} #end function

#define an optional alias
Set-Alias -Name chl -Value ConvertTo-HTMLListing
