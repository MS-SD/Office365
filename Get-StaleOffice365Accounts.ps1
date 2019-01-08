if (!(Get-Command -Name Get-Msoldomain)) {
    "Please connect to Office 365 via Powershell and try again."; break
}

$mailbox = Get-Mailbox -ResultSize unlimited  | Where {$_.RecipientTypeDetails -eq "UserMailbox"}
$count = $mailbox.Count
$company = (Get-MsolCompanyInformation).DisplayName
$result = @()

$i = 0
Foreach ($m in $mailbox) {
    $obj = Get-MailboxStatistics $m.UserPrincipalName -ErrorAction silentlyContinue |
        Where {$_.LastLogonTime -le (Get-Date).AddMonths(-1)} |
        Select-Object @{E = {"$company"}; L = "Company"}, `
    @{E = {$_.DisplayName}; L = "Account Name"}, `
    @{E = {$_.LastLogontime.Date.ToShortDateString()}; L = "Last Logon Date"}, `
    @{E = {$_.LastLogontime.ToShortTimeString()}; L = "Last Logon Time"}

    $result += $obj
    $I = $i + 1
    If ($obj) {
        $activity = $obj | Out-String
        Write-Progress -Id 2 -Activity "Stale user located: $activity" -Status "Mailbox $i/$count"  -PercentComplete ($i / $mailbox.Count * 100)
    }# If
    Else {
        Write-Progress -Id 2 -Activity "Locating stale users in $($t.client) tenant..." -Status "Mailbox $i/$count" -PercentComplete ($i / $mailbox.Count * 100)
    }# else

}# Foreach

$result | Export-Csv "$env:USERPROFILE\Desktop\$company StaleUsers.csv"
Write-Host "File has been created" -f DarkCyan