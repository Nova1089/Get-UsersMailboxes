# functions
function Show-Introduction
{
    Write-Host "This script shows all Shared Mailboxes, Group Mailboxes, and Team Mailboxes that a user has access to in Microsoft 365.`n" -ForegroundColor DarkCyan
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule($moduleName)
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor DarkCyan
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor DarkCyan
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "(?<!\S)y(?!\S)") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
        "1. Open Powershell as admin.`n" +
        "2. CD into script directory.`n" +
        "3. Run .\scriptname`n") -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit
    }
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor DarkCyan
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Write-Warning "Failed to connect to Exchange Online."
            Read-Host "Press Enter to try again."
        }
    }
}

function PromptFor-User
{
    do
    {
        $userPrompt = Read-Host "Enter user's full name or email"
        $mailbox = Get-EXOMailbox -Identity $userPrompt -ErrorAction SilentlyContinue

        if ($null -eq $mailbox)
        {
            Write-Warning "User not found."
        }
    }
    while ($null -eq $mailbox)

    Write-Host "User found: $($mailbox.UserPrincipalName)`n" -ForegroundColor Green

    return $mailbox.UserPrincipalName
}

function Get-ReadManageMailboxes($mailboxesToSearch, $userEmail)
{
    Write-Host "Getting mailboxes where user has `"Read and Manage`" access. Please wait..." -ForegroundColor DarkCyan

    $totalMailboxesWithReadAndManage = 0
    $mailboxesToSearch |
    Write-ProgressInPipeline -activity "Searching mailboxes..." -status "mailboxes searched" |
    Get-MailboxPermission -User $userEmail |
    Where-Object { $_.AccessRights } | # they don't have access if there is no "AccessRights" property"
    ForEach-Object {
        $totalMailboxesWithReadAndManage++
        $_
    } | Format-Table @{Label = "Mailbox"; Expression = { $_.Identity }; Width = 35}, AccessRights

    Write-Host "Found $totalMailboxesWithReadAndManage mailboxes where user has `"Read and Manage`" access.`n" -ForegroundColor DarkCyan
}

function Write-ProgressInPipeline
{
    [Cmdletbinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        [object[]] $inputObjects,
        [string] $activity = "Processing items...",
        [string] $status = "items processed"
    )

    Begin 
    { 
        $itemsProcessed = 0 
    }

    Process
    {
        Write-Progress -Activity $activity -Status "$itemsProcessed $status"
        $itemsProcessed++
        return $_
    }
}

function Get-SendAsMailboxes($mailboxesToSearch, $userEmail)
{
    Write-Host "Getting mailboxes where user has `"Send As`" access. Please wait..." -ForegroundColor DarkCyan

    $totalMailboxesWithSendAs = 0
    $mailboxesToSearch |
    Write-ProgressInPipeline -activity "Searching mailboxes..." -status "mailboxes searched" |
    Get-RecipientPermission -Trustee $userEmail |
    Where-Object { $_.AccessRights } | # they don't have access if there is no "AccessRights" property"
    ForEach-Object {
        $totalMailboxesWithSendAs++
        $_
    } | Format-Table @{Label = "Mailbox"; Expression = { $_.Identity }; Width = 35}, AccessRights

    Write-Host "Found $totalMailboxesWithSendAs mailboxes where user has `"Send As`" access.`n" -ForegroundColor DarkCyan
}

# main
Show-Introduction
Use-Module("ExchangeOnlineManagement")
TryConnect-ExchangeOnline
$userEmail = PromptFor-User
$allMailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox, GroupMailbox, TeamMailbox
Get-ReadManageMailboxes -mailboxesToSearch $allMailboxes -userEmail $userEmail
Get-SendAsMailboxes -mailboxesToSearch $allMailboxes -userEmail $userEmail
Read-Host "Press Enter to exit"