Import-Module -Name 'ActiveDirectory'
. "$PSScriptRoot\Cmdlets.ps1"
. "$PSScriptRoot\Config.ps1"

$params = @{
    DaysBefore = $Script:Config.DaysBefore
    SearchBase = $Script:Config.SearchBase
}
$accounts = Get-AccountsWithPasswordsAboutToExpire @params

$params = @{
    EmailTemplate = $Script:Config.EmailTemplate
    From = $Script:Config.From
    Subject = $Script:Config.Subject
    SmtpServer = $Script:Config.SmtpServer
}
$accounts | Send-PasswordExpirationNotice @params
