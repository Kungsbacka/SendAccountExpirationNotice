$Script:Config = @{
    # Where to base the search
    SearchBase = 'DC=example,DC=com'
    # Find users where the password is about to expire in the next n days
    DaysBefore = 10
    # SMTP server used to relay message
    SmtpServer = 'smtp.example.com'
    # Mail from
    From = 'Admin <noreply@example.com>'
    # Subject text
    Subject = 'Lösenordsbyte'
    # Email temlate. {NAME} = display name, {SAM} SAM account name, {DAYS} = days string (in Swedish)
    EmailTemplate = @"
Hej {NAME},

Ditt lösenord för {SAM} går ut {DAYS}.


Vänliga hälsningar
Admin
"@
}
