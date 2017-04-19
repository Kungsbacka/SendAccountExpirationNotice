Import-Module -Name 'ActiveDirectory'

function Get-AccountsWithPasswordsAboutToExpire
{
    param
    (
        [Parameter(Mandatory=$true,Position=0,ParameterSetName='Search')]
        [ValidateRange(1,10)]
        [int]
        $DaysBeforeExpiration,
        [Parameter(Mandatory=$true,Position=1,ParameterSetName='Search')]
        [string]
        $SearchBase,
        [Parameter(Mandatory=$true,Position=0,ParameterSetName='SingleUser')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Identity
    )
    begin
    {
        if ($PsCmdlet.ParameterSetName -eq 'SingleUser')
        {
            $params = @{
                Properties = @('msDS-UserPasswordExpiryTimeComputed')
                Identity = $Identity
            }
        }
        else
        {
            $passwordPolicy = Get-ADDefaultDomainPasswordPolicy
            $start = (Get-Date).AddDays(-$passwordPolicy.MaxPasswordAge.Days).ToFileTimeUtc()
            $end = (Get-Date).AddDays($DaysBeforeExpiration - $passwordPolicy.MaxPasswordAge.Days + 1).Date.ToFileTimeUtc()
            $params = @{
                Properties = @('msDS-UserPasswordExpiryTimeComputed')
                SearchBase = $SearchBase
                Filter = {
                    Enabled -eq $true
                    -and PasswordNeverExpires -eq $false
                    -and homeMDB -like '*'
                    -and mailNickName -like '*'
                    -and pwdLastSet -ge $start
                    -and pwdLastSet -le $end
                }
            }
        }
        $users = Get-ADUser @params
        foreach ($user in $users)
        {
            $expirationDate = [DateTime]::FromFileTimeUtc($user.'msDS-UserPasswordExpiryTimeComputed').ToLocalTime()
            $out = [pscustomobject]@{
                Name = $user.GivenName
                EmailAddress = $user.UserPrincipalName
                SamAccountName = $user.SamAccountName
                ExpirationDate = $expirationDate
                DaysBeforeExpiration = ($expirationDate.Date - (Get-Date).Date).Days
            }
            Write-Output -InputObject $out
        }
    }
}

function Send-PasswordExpirationNotice
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [object[]]
        $InputObject,
        [Parameter(Mandatory=$true)]
        [string]
        $EmailTemplate,
        [Parameter(Mandatory=$true)]
        [string]
        $From,
        [Parameter(Mandatory=$true)]
        [string]
        $Subject,
        [Parameter(Mandatory=$true)]
        [string]
        $SmtpServer
    )
    begin
    {
        # The reason for not using Send-MailMessage is that there is no way to avoid authentication and
        # I have not found a way to get a gMSA to authenticate successfully against an Exchange connector.
        $smtpClient = New-Object -TypeName 'System.Net.Mail.SmtpClient'
        $smtpClient.UseDefaultCredentials = $false
        $smtpClient.Host = $SmtpServer
    }
    process
    {
        foreach ($item in $InputObject)
        {
            $date = $_.ExpirationDate.ToString('yyyy-MM-dd') + ' klockan ' + $_.ExpirationDate.ToString('HH:mm')
            if ($_.DaysBeforeExpiration -gt 1)
            {
                $msg = "om $($_.DaysBeforeExpiration) dagar ($date)"
            }
            elseif ($_.DaysBeforeExpiration -eq 1)
            {
                $msg = "imorgon ($date)"
            }
            elseif ($_.DaysBeforeExpiration -eq 0)
            {
                $msg = "idag ($date)"
            }
            else
            {
                continue
            }
            $mail = New-Object -TypeName 'System.Net.Mail.MailMessage'
            $mail.BodyEncoding = [System.Text.Encoding]::UTF8
            $mail.SubjectEncoding = [System.Text.Encoding]::UTF8
            $mail.From = $From
            $mail.To.Add($_.EmailAddress)
            $mail.Subject = $Subject
            $mail.Body = $EmailTemplate.Replace('{NAME}', $_.Name).Replace('{SAM}', $_.SamAccountName).Replace('{DAYS}', $msg)
            $smtpClient.Send($mail)
            $mail.Dispose()
        }
    }
    end
    {
        $smtpClient.Dispose()
    }
}
