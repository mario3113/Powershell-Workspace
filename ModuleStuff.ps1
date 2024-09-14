
# pdate-AllMyModules.ps1
Get-Module -ListAvailable | `
  Where-Object {$null -ne $PSItem.RepositorySourceLocation -and $PSItem.ModuleBase -like "$($env:Userprofile)*"} | `
  Select-Object -Unique -Property Name | `
  Update-Module;

#  Remove-OldModules.ps1
Get-Module -ListAvailable | `
    Where-Object { $null -ne $PSItem.RepositorySourceLocation -and $PSItem.ModuleBase -like "$($env:Userprofile)*" } | `
    Select-Object -Unique -ExpandProperty Name | `
    ForEach-Object {
    Get-Module -ListAvailable -Name $PSItem | `
        Sort-Object -Property Version -Descending | `
        Select-Object -Skip 1 | `
        foreach-object {
        Uninstall-Module -Name $PSItem.Name -RequiredVersion $PSItem.Version
    } 
}

## Update-AllModules.ps1
Get-Module -ListAvailable | `
  Where-Object {$null -ne $PSItem.RepositorySourceLocation} | `
  Select-Object -Unique -Property Name | `
  Update-Module;


# dbatools-spellcheck.json
  "cSpell.enableCompoundWords":true,
    "cSpell.userWords": [
        "MSSQL",
        "Chrissy",
        "gmail",
        "cmdlet",
        "Maire",
        "clemaire",
        "dbatools",
        "computername",
        "interop",
        "sqlcredential",
        "wildcard",
        "failover",
        "subnet",
        "securestring",
        "scriptblock",
        "verifyonly",
        "notin",
        "notlike",
        "notcontains",
        "sqlcollaborative",
        "pscustomobject",
        "pscmdlet",
        "shouldprocess",
        "cmdletbinding",
        "msdb",
        "remoting",
        "HKLM",
        "machinekey",
        "winrm",
        "scred",
        "dcred",
        "netnerds",
        "sqlserver",
        "sqlcluster",
        "filestreams",
        "filegroups",
        "recurse",
        "percentcomplete",
        "tempdb",
        "SSIS",
        "SSISDB",
        "Progressbar",
        "nvarchar",
        "datatable",
        "datarow",
        "bulkcopy",
        "bigint",
        "smallint",
        "tinyint",
        "datetime",
        "guid",
        "uniqueidentifier",
        "wildcard",
        "wildcards",
        "multiprotocol",
        "subnet",
        "subnets",
        "Updateable",
        "OLTP",
        "Abshire",
        "servername",
        "adhoc",
        "SSAS",
        "endregion",
        "parameterclass",
        "paramserver",
        "notmatch",
        "sqlinstance",
        "cmdlets",
        "ldap",
        "contoso",
        "ctrlb"
    ]


    #Updated from the Qliq site's PowerShell testing template
Get-Date
$dnsName= "My_DSN_name"
$user="My_User_ID"
$password=""
$csvPath= "C:\temp\OutFileNameForTestingTheDSN.csv"
#$sqlQuery= "select 2 as f1-- * from DB.dbo.Table"  #hardcode some SQL
$sqlQuery= Get-Content "C:\temp\TestSQL.txt"

$conn = New-Object Data.Odbc.OdbcConnection
$conn.ConnectionString= "dsn=$dnsName;uid=$user;pwd=$password;"
$conn.open()
$command =$conn.CreateCommand();
$command.CommandText=$sqlQuery

$dataAdapter = New-Object System.Data.Odbc.OdbcDataAdapter $command

$dataTable = new-object "System.Data.DataTable"
$dataAdapter.Fill($dataTable)
$conn.close()

$dataTable | Export-csv -Path $csvPath -NoTypeInformation
Get-Date



Start-Process -FilePath "msiexec.exe" -ArgumentList "/quiet", "/passive", "/qn", "/i", "C:\tools\msodbcsql.msi", "IACCEPTMSODBCSQLLICENSETERMS=YES", "ADDLOCAL=ALL" -Passthru |
     Wait-Process. For this to work, I also had to install the Visual C++ Redistributable 2017. 
