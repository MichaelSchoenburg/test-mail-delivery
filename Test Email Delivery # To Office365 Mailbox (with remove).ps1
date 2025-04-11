<# 
Script Author: Michael Schönburg
Written: 03/13/2021
#>

Clear-Host

#######################################
# Send email from Gmail to Office 365 #
#######################################

Connect-ExchangeOnline
Connect-IPPSSession

[String]$Recipient = 'willkommen@domain.tld'

# Define clear text string for username and password
[String]$UserName = 'testmailaddress@gmail.com'
[String]$UserPassword = 'Passwort'

# Convert to SecureString
[Securestring]$SecStringPassword = ConvertTo-SecureString $UserPassword -AsPlainText -Force

# Create Credentials
[PSCredential]$Credentials = New-Object System.Management.Automation.PSCredential ( $UserName, $SecStringPassword )

# Define the Send-MailMessage parameters
$Subject = "Test Email Delivery - $( Get-Date -Format 'dd.MM.yyyy HH:mm:ss' )"
$MailParams = 
@{
    SmtpServer                 = 'smtp.gmail.com'
    Port                       = '587'
    UseSSL                     = $true
    Credential                 = $Credentials
    From                       = $UserName
    To                         = $Recipient
    Subject                    = $Subject
    Body                       = 'This is a test email using SMTP Client Submission'
    DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
}

# Send the message
Send-MailMessage @MailParams

###########################################
# Check for email reception in Office 365 #
###########################################

$SearchName = "EmailDeliveryTest_$(  Get-Date -Format 'dd.MM.yyyy_HH:mm:ss' )"
$PurgeName = "$(  $SearchName  )_Purge"
$ContentMatchQuery = 'Subject:'+'"'+$Subject+'"'

# Search Test-Email
do
{
    New-ComplianceSearch `
        -Name $SearchName `
        -ContentMatchQuery $ContentMatchQuery `
        -ExchangeLocation willkommen@domain.tld
    
    Start-ComplianceSearch $SearchName
}
until ( ( Get-ComplianceSearch $SearchName ).Status -eq "Completed" )
$ResultsCount = (Get-ComplianceSearch $SearchName).SuccessResults

if ( $ResultsCount -eq 0 ) 
    {
        Write-Output "Test unsuccessful."
    }
    else
    {
        Write-Output "Test successful."
        Break
    }
    else 
    {
        Write-Output "Test unsuccessful. Found $( $ResultsCount ) results instead of just one result."
    }    

# Delete Test-Email
New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete -Confirm:$false -Force

do
{
    Write-Verbose "Waiting..."
    Start-Sleep -Seconds 1
}
until ( ( Get-ComplianceSearchAction $PurgeName ).Status -eq "Completed" )

$global:CSResult = Get-ComplianceSearchAction $PurgeName

Write-Verbose "Deleting..."
 
Get-ComplianceSearchAction $PurgeName | Remove-ComplianceSearchAction -Confirm:$false
 
Write-Verbose "Checking if deleted..."

try
{
    Get-ComplianceSearchAction $PurgeName -ErrorAction Stop
}
catch
{
     
    $ErrorOccured = $true
}

if ( $ErrorOccured ) {
    Write-Verbose "Loop successful."
}
else
{
    Write-Verbose "Loop unsuccessful. Stopping."
}
