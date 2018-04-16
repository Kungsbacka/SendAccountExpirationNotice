$Script:Config = @{
    # Where to base the search
    SearchBase = 'DC=example,DC=com'
    # Find users where the password is about to expire in the next n days
    DaysBefore = 10
    # Array of mail domains that should be included in the search.
    # If the array is empty, all mail domains are included.
    MailDomains = @('example.com','example.net')
    # SMTP server used to relay message
    SmtpServer = 'smtp.example.com'
    # Mail from
    From = 'Admin <noreply@example.com>'
    # Subject text
    Subject = 'Lösenordsbyte'
    # Email temlate. {NAME} = display name, {SAM} SAM account name, {DAYS} = days string (in Swedish)
    EmailTemplate = @"
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
</head>
<body style="font-family:Calibri, Sans-serif; font-size:11pt; line-hight:120%;">
<p>Hej {NAME},</p>

<p>Ditt lösenord för {SAM} går ut {DAYS}.</p>

<p>Vänliga hälsningar<br>
Admin</p>
</body>
</html>
"@
    # Report recipient
    AdminReportRecipient = 'Admin <admin@example.com>'
    # Report subject
    AdminReportSubject = 'Rapport'
    # Email template. {TABLEROWS} = HTML table rows (four columns: Name/Email/Expiration date/Days left)
    AdminReportTemplate = @"
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
</head>
<body style="font-family:Calibri, Sans-serif; font-size:11pt; line-hight:120%;">
<style>
table { color: #333; border-collapse: collapse; border-spacing: 0; }
td, th { border: 1px solid transparent; height: 30px; }
th { background: #DFDFDF; font-weight: bold; }
td { background: #FAFAFA; text-align: center; }
</style>
<p>Hej,</p>
<p>Följande användare har fått ett meddelande om lösenordsbyte:</p>
<table>
<tr><th>Namn</th><th>E-postadress</th><th>Datum</th><th>Dagar kvar</th></tr>
{TABLEROWS}
</table>
<p>Hälsningar,<br>
Admin</p>
</body>
</html>
"@
}
