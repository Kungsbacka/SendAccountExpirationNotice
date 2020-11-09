Import-Module -Name 'ActiveDirectory'

function Get-AccountAboutToExpire
{
    param
    (
        [Parameter(Mandatory=$true)]
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
        $UpnDomain,
        [Parameter(Mandatory=$false)]
        [switch]
        $MailEnabledOnly
    )
    begin
    {
        if ($PsCmdlet.ParameterSetName -eq 'SingleUser')
        {
            $params = @{
                Properties = @('AccountExpirationDate','Manager','DisplayName','Mail')
                Identity = $Identity
            }
        }
        else
        {
            $start = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
            $end = [DateTime]::Now.AddDays($DaysBeforeExpiration).ToString('yyyy-MM-dd HH:mm:ss')
            $filter = "Enabled -eq 'true' -and AccountExpirationDate -ge '$start' -and AccountExpirationDate -le '$end'"
            $domainFilter = ''
            foreach ($domain in $UpnDomain)
            {
                $domainFilter += "userPrincipalName -like '*@$domain' -or "
            }
            if ($domainFilter)
            {
                $filter += " -and ($($domainFilter.Substring(0, $domainFilter.Length - 5)))"
            }
            if ($MailEnabledOnly) {
                $filter += " -and Mail -like '*' -and MailNickname -like '*' -and MsExchRecipientTypeDetails -like '*'"
            }
            $params = @{
                Properties = @('AccountExpirationDate','Manager','DisplayName','Mail','MailNickname','MsExchRecipientTypeDetails')
                SearchBase = $SearchBase
                Filter = $filter
            }
        }
        $users = Get-ADUser @params
        foreach ($user in $users)
        {
            $out = [pscustomobject]@{
                GivenName = $user.GivenName
                DisplayName = $user.DisplayName
                EmailAddress = $null
                ManagerEmailAddress = $null
                SamAccountName = $user.SamAccountName
                ExpirationDate = $user.AccountExpirationDate
                DaysBeforeExpiration = ($user.AccountExpirationDate - (Get-Date).Date).Days - 1
            }
            if ($user.Mail -and $user.MailNickname -and $user.MsExchRecipientTypeDetails) {
                $out.EmailAddress = $user.Mail
            }
            if ($user.Manager) {
                $manager = Get-ADUser -Identity $user.Manager -Properties @('Mail')
                $out.ManagerEmailAddress = $manager.Mail
            }
            Write-Output -InputObject $out
        }
    }
}

function New-AccountExpirationMessage
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
        $Subject
    )
    process
    {
        foreach ($account in $InputObject)
        {
            if (-not $account.EmailAddress) {
                continue
            }
            $date = $account.ExpirationDate.ToString('yyyy-MM-dd')
            if ($account.DaysBeforeExpiration -gt 1)
            {
                $msg = "om $($account.DaysBeforeExpiration) dagar ($date)"
            }
            elseif ($account.DaysBeforeExpiration -eq 1)
            {
                $msg = "imorgon ($date)"
            }
            elseif ($account.DaysBeforeExpiration -eq 0)
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
            $mail.To.Add($account.EmailAddress)
            $mail.Subject = $Subject
            $mail.Body = $EmailTemplate.Replace('{NAME}', $account.GivenName).Replace('{SAM}', $account.SamAccountName).Replace('{DAYS}', $msg).Replace('{DATE}', $date)
            
            $mail
        }
    }
}

function New-ManagerReportMessage
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
        $Subject
    )
    begin
    {
        $accounts = New-Object -TypeName 'System.Collections.ArrayList'
    }
    process
    {
        foreach ($account in $InputObject) {
            $null = $accounts.Add($account)
        }
    }
    end
    {
        $properties = @(
            @{n='Namn';e={$_.DisplayName}}
            @{n='E-postadress';e={if ($_.EmailAddress) {$_.EmailAddress} else {'(Kontot saknar e-post)'}}}
            @{n='Konto'; e={$_.SamAccountName}}
            @{n='Utgångsdatum'; e={$_.ExpirationDate.ToString('yyyy-MM-dd')}}
        )
        $table = $accounts | Select-Object -Property $properties | ConvertTo-Html -As Table -Fragment
        $mail = New-Object -TypeName 'System.Net.Mail.MailMessage'
        $mail.BodyEncoding = [System.Text.Encoding]::UTF8
        $mail.SubjectEncoding = [System.Text.Encoding]::UTF8
        $mail.IsBodyHtml = $true
        $mail.From = $From
        $mail.To.Add($To)
        $mail.Subject = $Subject
        $mail.Body = $EmailTemplate.Replace('{TABLE}', $table)
        
        $mail
    }
}

function Send-AccountExpirationMessage
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [System.Net.Mail.MailMessage[]]
        $InputObject,
        [Parameter(Mandatory=$true)]
        [string]
        $SmtpServer,
        [Parameter(Mandatory=$false)]
        [string]
        $SendAllMessagesTo,
        [Parameter(Mandatory=$false)]
        [string]
        $RunOn,
        [Parameter(Mandatory=$false)]
        [PSCredential]
        $Credential
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
        foreach ($mail in $InputObject) {
            if ($SendAllMessagesTo) {
                $mail.To.Clear()
                $null = $mail.To.Add($SendAllMessagesTo)
            }
            if ($RunOn) {
                $params = @{
                    ComputerName = $RunOn
                }
                if ($Credential) {
                    $Params.Credential = $Credential
                }
                Invoke-Command @params -ScriptBlock {
                    $msg = $using:PSItem
                    $params = @{
                        SmtpServer = $using:SmtpServer
                        From = $msg.From
                        To = $msg.To
                        Subject = $msg.Subject
                        Body = $msg.Body
                        Encoding = 'UTF8'
                        BodyAsHtml = $true
                    }
                    Send-MailMessage @params
                }
            }
            else {
                $smtpClient.Send($mail)
            }
        }
    }
    end
    {
        $smtpClient.Dispose()
    }
}

function Send-ExamapleNotice
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet('Individual','Manager')]
        [string]
        $Type,
        [Parameter(Mandatory=$true)]
        [string]
        $Recipient,
        [Parameter(Mandatory=$false)]
        [string]
        $RunOn,
        [Parameter(Mandatory=$false)]
        [PSCredential]
        $Credential
    )
    begin {
        $config = . "$PSScriptRoot\Config.ps1"
        $accounts = @(
            [pscustomobject]@{GivenName='Anna';DisplayName='Anna Andersson';EmailAddress='anna.andersson@kungsbacka.se';ManagerEmailAddress='lotta.larsson@kungsbacka.se';SamAccountName='annand';ExpirationDate=(Get-Date).AddDays(5);DaysBeforeExpiration=5}
            [pscustomobject]@{GivenName='Pelle';DisplayName='Pelle Persson';EmailAddress='pelle.persson@kungsbacka.se';ManagerEmailAddress='lotta.larsson@kungsbacka.se';SamAccountName='pelper';ExpirationDate=(Get-Date).AddDays(8);DaysBeforeExpiration=8}
            [pscustomobject]@{GivenName='Stina';DisplayName='Stina Svensson';EmailAddress='stina.svennson@kungsbacka.se';ManagerEmailAddress='lotta.larsson@kungsbacka.se';SamAccountName='stisve';ExpirationDate=(Get-Date).AddDays(7);DaysBeforeExpiration=7}
        )
        if ($Type -eq 'Individual') {
            $params = @{
                EmailTemplate = (Get-Content -Path "$PSScriptRoot\$($config.IndividualTemplate)" -Encoding UTF8 | Out-String)
                From = $config.From
                Subject = $config.IndividualSubject
            }
            $messages = $accounts[0] | New-AccountExpirationMessage @params
        }
        else {
            $params = @{
                EmailTemplate = (Get-Content -Path "$PSScriptRoot\$($config.ManagerTemplate)" -Encoding UTF8 | Out-String)
                From = $config.From
                To = $Recipient
                Subject = $config.ManagerSubject
            }
            $messages = $accounts | New-ManagerReportMessage @params
        }
        $params = @{
            SmtpServer = $config.SmtpServer
            SendAllMessagesTo = $Recipient
        }
        if ($RunOn) {
            $params.RunOn = $RunOn
        }
        if ($Credential) {
            $params.Credential = $Credential
        }
        $messages | Send-AccountExpirationMessage @params
    }
}
