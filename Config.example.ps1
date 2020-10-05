$Script:Config = @{
    # Where to base the search
    SearchBase = 'DC=example,DC=com'
    # Find users where the password is about to expire in the next n days.
    # This is for the notice that is sent to individual users
    IndividualDaysBeforeExpiration = 10
    # Find users where the password is about to expire in the next n days.
    # This is for the report that is sent to the managers
    ManagerDaysBeforeExpiration = 14
    # Array of UPN domains that should be included in the search. If the array
    # is empty, all domains are included.
    UpnDomains = @('contoso.com','anotherdomain.com')
    # SMTP server used to relay message
    SmtpServer = 'smtp.example.com'
    # From address
    From = 'Admin <noreply@example.com>'
    # Subject text for the notice that is sent to individual users
    IndividualSubject = 'Your account is about to expire'
    # Email HTML template used for the notice that is sent to individual users.
    # The following placeholders will be replaced in the template:
    #   {NAME} => display name
    #   {SAM} => SAM account name
    #   {DAYS} => days string (currently in Swedish)
    IndividualTemplate = "Individual.html"
    # Email HTML temlate used for the report that is sent to the managers.
    # The following placeholders will be replaced:
    #   {TABLE} => table with accounts about to expire that report to the manager
    ManagerTemplate = 'Manager.html'
    # Subject text used for the report sent to the managers
    ManagerSubject = 'Avslut av användarkonto'
}
