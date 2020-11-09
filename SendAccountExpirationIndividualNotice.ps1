Import-Module -Name 'ActiveDirectory'
. "$PSScriptRoot\Functions.ps1"
$Script:Config = . "$PSScriptRoot\Config.ps1"

$template = Get-Content -Path "$PSScriptRoot\$($Script:Config.IndividualTemplate)" -Encoding UTF8 | Out-String

$params = @{
    DaysBeforeExpiration = $Script:Config.IndividualDaysBeforeExpiration
    SearchBase = $Script:Config.SearchBase
    UpnDomain = $null
    MailEnabledOnly = $true
}

# Fetch accounts
$accountsAboutToExpire = New-Object -TypeName 'System.Collections.ArrayList'
foreach ($domain in $Script:Config.UpnDomains) {
    $params.UpnDomain = $domain    
    Get-AccountAboutToExpire @params | ForEach-Object {$null = $accountsAboutToExpire.Add($_)}
}

# Create messages
$messages = New-Object -TypeName 'System.Collections.ArrayList'
$params = @{
    EmailTemplate = $template
    From = $Script:Config.From
    Subject = $Script:Config.IndividualSubject
}
foreach ($account in $accountsAboutToExpire) {
    $message = $account | New-AccountExpirationMessage @params
    $null = $messages.Add($message)
}

# Send messages
$messages | Send-AccountExpirationMessage -SmtpServer $Script:Config.SmtpServer
