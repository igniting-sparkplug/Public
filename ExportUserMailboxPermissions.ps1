<#

Purpose: Export SendAs and FullAccess permission for all user mailboxes in Office 365
Version: v1.0
Author: Roshan N.
GitHub: https://github.com/roshannair7/
LinkedIn: https://www.linkedin.com/in/roshannair7/
Twitter: https://twitter.com/IgniteSparkplug

#>

$ST = Get-Date
$path = "C:\PSScripts" #Change the path as per requirement
If(!(test-path $path))
{
    New-Item -ItemType Directory -Force -Path $path
    CD $path
    cls
    Write-Host "Creating new folder in" $path "..." -ForegroundColor Cyan
    Sleep 2
}
else
{
    Write-Host "Moving PS run location to" $path "..." -ForegroundColor Cyan
    sleep 2
    CD $path
}

$username = Read-Host "Enter GA email address"
Read-Host -Prompt “Enter your tenant GA's password” -AsSecureString | ConvertFrom-SecureString | Out-File "$path\TENANTNAME.key"
$password = cat "$path\TENANTNAME.key" | ConvertTo-SecureString

Function Connect-EXOnline
{
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Write-host "Creating new O365 Session" -ForegroundColor Cyan
    $URL = "https://ps.outlook.com/powershell"
    $TenantCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $TenantCredentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
    Import-PSSession $EXOSession
}

Function Disconnect-EXOnline
{
    Write-Host "Disconnecting Office 365 section temporarily" -ForegroundColor Cyan
    Get-PSSession | Remove-PSSession
}

Function Get-ScriptRunTime
{
    $ET = Get-Date
    $diff = New-TimeSpan -Start $ST -End $ET
    Write-Host "Total script run time is" $diff.Days "days" $diff.Hours "hours" $diff.Minutes "minutes &" $diff.Seconds "Seconds" -ForegroundColor Cyan
}

Connect-EXOnline
cls
Write-Host "Calculating the number of User Mailboxes in the environment..." -ForegroundColor Cyan

Get-Mailbox | Select UserPrincipalName | Sort-Object UserPrincipalname | Export-Csv -Path "$path\UPNList.csv"
$MBX = Import-Csv -Path "$path\UPNList.csv"

$TotalCount = $MBX.count
$count = 0
$j=2500

foreach($i in $MBX)
{
    $count++
    $count = $count++
    Write-Host "Extracting for" $i.UserPrincipalName". (" $count "of" $TotalCount ")" -ForegroundColor Green
    Get-MailboxPermission $i.UserprincipalName | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like “NT AUTHORITY\SELF”) } | Select Identity,user,AccessRights | Export-Csv -Path "$path\UserMailboxFullAccessPermission.csv" -Append
    Get-RecipientPermission $i.UserprincipalName | Where {($_.Trustee -ne 'nt authority\self') -and ($_.Trustee -ne 'Null')} | Select Identity, Trustee, AccessRights | Export-Csv -Path "$path\UserMailboxSendAsPermission.csv" -Append
    $j--
        if ($j -eq 1)
        {
            Disconnect-EXOnline
            Write-Host "Refreshing current PS Session" -ForegroundColor Yellow
            sleep -Seconds 20
            Connect-EXOnline
            $j=2500
        }
}
Get-ScriptRunTime