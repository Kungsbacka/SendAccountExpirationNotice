Import-Module -Name 'ActiveDirectory'
. "$PSScriptRoot\Cmdlets.ps1"
. "$PSScriptRoot\Config.ps1"

$params = @{
    DaysBefore = $Script:Config.DaysBefore
    SearchBase = $Script:Config.SearchBase
}
$accounts = Get-AccountsWithPasswordAboutToExpire @params

$params = @{
    EmailTemplate = $Script:Config.EmailTemplate
    From = $Script:Config.From
    Subject = $Script:Config.Subject
    SmtpServer = $Script:Config.SmtpServer
}
$accounts | Send-PasswordExpirationNotice @params

$params = @{
    EmailTemplate = $Script:Config.AdminReportTemplate
    From = $Script:Config.From
    To = $Script:Config.AdminReportRecipient
    Subject = $Script:Config.AdminReportSubject
    SmtpServer = $Script:Config.SmtpServer
}
# Enable admin report to get a full report of all users receiving a notice
# $accounts | Send-AdminReport @params
