
#the folder size includes hidden files

Function Get-FSOFolderSize {
[cmdletbinding()]
[OutputType("PSCustomObject")]

param(
[Parameter(Position = 0, Mandatory, HelpMessage = "Enter a filesystem path like C:\Scripts. Do not specify a directory root.")]
[ValidateNotNullOrEmpty()]
[ValidateScript({$_ -notmatch '^[a-zA-Z]:\\$'})]
[string]$Path,
[ValidateNotnullorEmpty()]
[string]$Computername = $env:COMPUTERNAME
)

Write-Verbose "Starting $($MyInvocation.MyCommand)"

Write-Verbose "Measuring $path on $($computername.toUpper())"
Invoke-Command {
if (Test-Path $using:path) {
    $fso = New-Object -ComObject Scripting.FileSystemObject
    $fso.GetFolder($using:path)
}
else {
    Write-Warning "Can't find $($using:path) on $env:computername"
}
} -computername $Computername | 
Select-Object -Property @{Name="Computername";Expression={$_.pscomputername.toUpper()}},
Path,DateCreated,DateLastModified,Size,
@{Name="SizeMB";Expression={$_.size/1mb -as [int]}}

Write-Verbose "Ending $($MyInvocation.MyCommand)"

}

<#

PS C:\> Get-FSOFolderSize C:\scripts 


Computername     : BOVINE320
Path             : C:\scripts
DateCreated      : 7/31/2017 5:06:58 PM
DateLastModified : 10/5/2018 5:12:51 PM
Size             : 859878325
SizeMB           : 820

#>
