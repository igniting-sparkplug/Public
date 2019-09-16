# Get-ScriptRunTime Initialize
$ST = Get-Date

# Signature
Function Set-Signature
{
    Write-Host "]3 /\ ~|~ |\/| /\ |\|" -ForegroundColor Yellow
    Write-Host  "/? () _\~ |-| /\ |\|" -ForegroundColor Yellow
}

#Create folder for Output and move PS default location
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

# Entering user credentials
$username = Read-Host "Enter GA email address"
# Hash out the next line of code one this is run for the first time on any machine. Passwd is secured and stored and shall be retrived at any time
Read-Host -Prompt “Enter your tenant GA's password” -AsSecureString | ConvertFrom-SecureString | Out-File "$path\TENANTNAME.key"
$password = cat "$path\TENANTNAME.key" | ConvertTo-SecureString

#All functions. Should Ideally move all functions to one seperate section above or below the actual code. Would look better

# Function to connect to Office 365
Function Connect-EXOnline
{
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Write-host "Creating new O365 Session" -ForegroundColor Cyan
    $URL = "https://ps.outlook.com/powershell"
    $TenantCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $TenantCredentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
    Import-PSSession $EXOSession
}

# Function to disconnect Office 365 section
Function Disconnect-EXOnline
{
    Write-Host "Disconnecting Office 365 section temporarily" -ForegroundColor Cyan
    Get-PSSession | Remove-PSSession
}

#Function to get the time taken for the script to run.
Function Get-ScriptRunTime
{
    $ET = Get-Date
    $diff = New-TimeSpan -Start $ST -End $ET
    Write-Host "Total script run time is" $diff.Days "days" $diff.Hours "hours" $diff.Minutes "minutes &" $diff.Seconds "Seconds" -ForegroundColor Cyan
}


# Involing function to connet to office 365
Connect-EXOnline

# Get Input
cls
Write-Host "Calculating the number of User Mailboxes in the environment..." -ForegroundColor Cyan

# Exporting to CSV and then importing so that in case it errors due to any network glitch, we dont have to restart from scratch. Also why its sorted alphabetically
Get-Mailbox | Select UserPrincipalName | Sort-Object UserPrincipalname | Export-Csv -Path "$path\UPNList.csv"
$MBX = Import-Csv -Path "$path\UPNList.csv"

# Counters for various tasks
$TotalCount = $MBX.count #Total Mailbox size
$count = 0 #Count the once completed
$j=2500 #PS reconnect counter

# Main Section
foreach($i in $MBX)
{
    $count++
    $count = $count++
    Write-Host "Extracting for" $i.UserPrincipalName". (" $count "of" $TotalCount ")" -ForegroundColor Green
    #Full Mailbox Permission
    Get-MailboxPermission $i.UserprincipalName | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like “NT AUTHORITY\SELF”) } | Select Identity,user,AccessRights | Export-Csv -Path "$path\UserMailboxFullAccessPermission.csv" -Append
    #Send As Permission
    Get-RecipientPermission $i.UserprincipalName | Where {($_.Trustee -ne 'nt authority\self') -and ($_.Trustee -ne 'Null')} | Select Identity, Trustee, AccessRights | Export-Csv -Path "$path\UserMailboxSendAsPermission.csv" -Append

    # For every 2500 entry, PS would force disconnect and connect again.
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

# Run function to print time taken by the script to run
Get-ScriptRunTime