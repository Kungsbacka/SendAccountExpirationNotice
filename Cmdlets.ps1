Import-Module -Name 'ActiveDirectory'

function Get-AccountsWithPasswordAboutToExpire
{
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateRange(1,10)]
        [int]
        $DaysBeforeExpiration,
        [Parameter(Mandatory=$true,ParameterSetName='Search')]
        [string]
        $SearchBase,
        [Parameter(Mandatory=$true,ParameterSetName='SingleUser')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Identity,
        [Parameter(Mandatory=$false,ParameterSetName='Search')]
        [string[]]
        $MailDomain
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
            $filter = "Enabled -eq 'true'" +
                " -and PasswordNeverExpires -eq 'false'" +
                " -and mail -like '*'" +
                " -and pwdLastSet -ge $start" +
                " -and pwdLastSet -le $end"
            $domainFilter = ''
            foreach ($domain in $MailDomain)
            {
                $domainFilter += "mail -like '*@$domain' -or "
            }
            if ($domainFilter)
            {
                $filter += " -and ($($domainFilter.Substring(0, $domainFilter.Length - 5)))"
            }
            $params = @{
                Properties = @('msDS-UserPasswordExpiryTimeComputed','DisplayName','Mail')
                SearchBase = $SearchBase
                Filter = $filter
            }
        }
        $users = Get-ADUser @params
        foreach ($user in $users)
        {
            $expirationDate = [DateTime]::FromFileTimeUtc($user.'msDS-UserPasswordExpiryTimeComputed').ToLocalTime()
            $out = [pscustomobject]@{
                GivenName = $user.GivenName
                DisplayName = $user.DisplayName
                EmailAddress = $user.Mail
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
            $mail.IsBodyHtml = $true
            $mail.From = $From
            $mail.To.Add($_.EmailAddress)
            $mail.Subject = $Subject
            $mail.Body = $EmailTemplate.Replace('{NAME}', $_.GivenName).Replace('{SAM}', $_.SamAccountName).Replace('{DAYS}', $msg).Replace('{DATE}', $date)
            $smtpClient.Send($mail)
            $mail.Dispose()
        }
    }
    end
    {
        $smtpClient.Dispose()
    }
}

function Send-AdminReport
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
        $To,
        [Parameter(Mandatory=$true)]
        [string]
        $Subject,
        [Parameter(Mandatory=$true)]
        [string]
        $SmtpServer
    )
    begin
    {
        $stringBuilder = New-Object 'System.Text.StringBuilder'
    }
    process
    {
        foreach ($item in $InputObject)
        {
            [void]$stringBuilder.AppendLine("<tr><td>$($item.DisplayName)</td><td>$($item.EmailAddress)</td><td>$($item.ExpirationDate.ToString('yyyy-MM-dd'))</td><td>$($_.DaysBeforeExpiration)</td></tr>")
        }
    }
    end
    {
        # The reason for not using Send-MailMessage is that there is no way to avoid authentication and
        # I have not found a way to get a gMSA to authenticate successfully against an Exchange connector.
        $smtpClient = New-Object -TypeName 'System.Net.Mail.SmtpClient'
        $smtpClient.UseDefaultCredentials = $false
        $smtpClient.Host = $SmtpServer
        $mail = New-Object -TypeName 'System.Net.Mail.MailMessage'
        $mail.BodyEncoding = [System.Text.Encoding]::UTF8
        $mail.SubjectEncoding = [System.Text.Encoding]::UTF8
        $mail.IsBodyHtml = $true
        $mail.From = $From
        $mail.To.Add($To)
        $mail.Subject = $Subject
        $mail.Body = $EmailTemplate.Replace('{TABLEROWS}', $stringBuilder.ToString())
        $smtpClient.Send($mail)
        $mail.Dispose()
        $smtpClient.Dispose()
    }
}
