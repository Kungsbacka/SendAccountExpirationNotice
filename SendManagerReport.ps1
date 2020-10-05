Import-Module -Name 'ActiveDirectory'
. "$PSScriptRoot\Functions.ps1"
. "$PSScriptRoot\Config.ps1"

$template = Get-Content -Path "$PSScriptRoot\$($Script:Config.ManagerTemplate)" -Encoding UTF8 | Out-String

$params = @{
    DaysBeforeExpiration = $Script:Config.ManagerDaysBeforeExpiration
    SearchBase = $Script:Config.SearchBase
    UpnDomain = $null
}

# Fetch accounts
$accountsAboutToExpire = New-Object -TypeName 'System.Collections.ArrayList'
foreach ($domain in $Script:Config.UpnDomains) {
    $params.UpnDomain = $domain    
    Get-AccountAboutToExpire @params | ForEach-Object {$null = $accountsAboutToExpire.Add($_)}
}

$accountsAboutToExpire = $accountsAboutToExpire | % {$_.EmailAddress = $null; $_}

# Create messages
$accountsGroupedByManager = $accountsAboutToExpire | Group-Object 'ManagerEmailAddress' | Where-Object Name -ne ''
$messages = New-Object -TypeName 'System.Collections.ArrayList'
foreach ($group in $accountsGroupedByManager) {
    $group.Group | New-ManagerReportMessage -EmailTemplate $template -From $Script:Config.From -To $group.Name -Subject $Script:Config.ManagerSubject | ForEach-Object {
        $null = $messages.Add($_)
    }
}

# Send messages
$messages | Send-AccountExpirationMessage -SmtpServer $Script:Config.SmtpServer
