#PSRefresh.psd1
#PSModules should be new modules to install
#Scope can be 'CurrentUser' or 'AllUsers'
#vscExtensions are Visual Studio Code extensions to install
@{
    wingetPackages = @(
        'Microsoft.VisualStudioCode',
        'Git.Git',
        'GitHub.cli',
        'Microsoft.WindowsTerminal',
        'github.githubdesktop',
        'Microsoft.PowerShell'
        )
    PSModules = @("PSScriptTools","PSProjectStatus","PSStyle","Platyps")
    vscExtensions = @(
        'github.copilot',
        'github.copilot-chat',
        'github.remotehub',
        'ms-vscode.powershell',
        'davidanson.vscode-markdownlint',
        'inu1255.easy-snippet',
        'gruntfuggly.todo-tree'
    )
    Scope = 'AllUsers'
    Version = '1.2.0'
}
