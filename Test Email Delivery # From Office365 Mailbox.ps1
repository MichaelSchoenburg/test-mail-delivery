<# 
Script Author: Michael Schönburg
Written: 03/13/2021
#>

Clear-Host

#######################################
# Send email from Office 365 to Gmail #
#######################################

[String]$Recipient = 'testmailaddress@gmail.com'

# Define clear text string for username and password
[String]$UserName = 'willkommen@domain.tld'
[String]$UserPassword = 'passwort'

# Convert to SecureString
[Securestring]$SecStringPassword = ConvertTo-SecureString $UserPassword -AsPlainText -Force

# Create Credentials
[PSCredential]$Credentials = New-Object System.Management.Automation.PSCredential ( $UserName, $SecStringPassword )

# Define the Send-MailMessage parameters
$Subject = "Test Email Delivery - $( Get-Date -Format 'dd.MM.yyyy HH:mm:ss' )"
$MailParams = 
@{
    SmtpServer                 = 'smtp.office365.com'
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

######################################
# Check for email reception in Gmail #
######################################

# Documentation:
# https://developers.google.com/gmail/gmail_inbox_feed

# Load Atom Gmail Inbox Feed via RSS
$Webclient = New-Object System.Net.WebClient

# Access the rss-feed
$Webclient.Credentials = New-Object System.Net.NetworkCredential ( "testmailaddress@gmail.com", "Passwort" )

$Successful = $false

# Display the table
ForEach ( $Second in ( 0..10 ) ) 
{
    # Download the rss as xml
    [Xml]$Xml= $Webclient.DownloadString( "https://mail.google.com/mail/feed/atom" )

    <#

    # Display only sender name and message title as custom table for debug purposes
    $Format =
        @{Expression={$_.Title};Label="Title"},
        @{Expression={$_.Author.name};Label="Author"}, 
        @{Expression={$_.Issued};Label="Date"}

    $Xml.feed.entry[0..2] | Format-Table $Format # Display the three latest emails in the inbox

    #>

    if ( $Xml.feed.entry[0].title -eq $Subject ) 
    {
        $Successful = $true
        Break
    }
    else {
        Start-Sleep -Seconds 1
    }
}

if ( $Successful ) 
{
    Write-Output "Test successful."
}
else 
{
    Write-Output "Test unsuccessful."
}

