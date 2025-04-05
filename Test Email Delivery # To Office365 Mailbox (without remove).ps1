<# 
Script Author: Michael Schönburg
Last Change: 05/19/2021
#>

<# 
    Variable declaration
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $TestEmailSubjectPrefix = 'Email Deliverability Test',

    [Parameter()]
    [datetime]
    $Timestamp = (Get-Date -Format 'MM/dd/yyyy HH:mm:ss'),

    [Parameter()]
    [string]
    $TestEmailSubject = "$( $TestEmailSubjectPrefix ) - $( $Timestamp )",

    [Parameter(
        Mandatory = $true,
        Position = 0
    )]
    [String[]]
    $TestEmailRecipients,

    [Parameter()]
    [String]
    $TestEmailSender,

    [Parameter()]
    [String]
    $TestEmailSenderPassword
)

# Solarwinds RMM check exit code
# 1 = failed
# 0 = successful
$SolarwindsRMMExitCode = 1

<# 
    Function declaration 
#>

function Send-TestEmail 
{
    Param
    (
        [parameter( Mandatory = $true )] [String] $TestEmailSubject,
        [parameter( Mandatory = $true )] [String] $TestEmailRecipient,
        [parameter( Mandatory = $true )] [String] $TestEmailSender,
        [parameter( Mandatory = $true )] [String] $TestEmailSenderPassword
    )

    # Convert password to SecureString and build PSCredential
    [Securestring]$SecStringPassword = ConvertTo-SecureString $TestEmailSenderPassword -AsPlainText -Force # Convert to SecureString
    [PSCredential]$Credentials = New-Object System.Management.Automation.PSCredential ( $TestEmailSender, $SecStringPassword ) # Create credentials

    # Define the Send-MailMessage parameters
    $MailParams = 
    @{
        SmtpServer                 = 'smtp.gmail.com'
        Port                       = '587'
        UseSSL                     = $true
        Credential                 = $Credentials
        From                       = $TestEmailSender
        To                         = $TestEmailRecipient
        Subject                    = $TestEmailSubject
        Body                       = 'Test'
        DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
    }

    # Send the message
    Send-MailMessage @MailParams
    
    # If no the script got to this point, there has no error occured, which means the email was sent without any error and the function can return $true
    return $true
}

function Remove-TestEmail
{
    Param
    (
        [parameter( Mandatory = $true )] [String] $TestEmailSubject,
        [parameter( Mandatory = $true )] [String] $TestEmailRecipient,
        [parameter( Mandatory = $true )] [String] $TestEmailSender
    )

    # Find the test email in the inbox via Compliance Search

    # Declare Compliance Search parameters
    $global:SearchNamePrefix = 'EmailDeliveryTest_'
    $SearchName = "$( $global:SearchNamePrefix  )$(  Get-Date -Format 'dd.MM.yyyy_HH:mm:ss' )"
    $ContentMatchQuery = 'Subject:' + '"' + $TestEmailSubject + '"'

    # Create Compliance Search
    New-ComplianceSearch `
        -Name $SearchName `
        -ContentMatchQuery $ContentMatchQuery `
        -ExchangeLocation $TestEmailRecipient

    # Start Compliance Search
    Start-ComplianceSearch $SearchName

    # Wait for Compliance Search to be completed
    do
    {
        # Wait
    }
    until ( ( Get-ComplianceSearch $SearchName ).Status -eq "Completed" )

    # Get number of Compliance Search results
    $ComplianceSearch = Get-ComplianceSearch $SearchName    #
    $ResultsCount = $ComplianceSearch                       #
    $ResultsCount = ( $ResultsCount ).SuccessResults        # 
    $ResultsCount = $ResultsCount.Split( ',' )              # Split up the string by comma.
    $ResultsCount = $ResultsCount[1]                        # Get the second item from the split up string.
    $ResultsCount = $ResultsCount.Split( ':' )              # Split up by colon.
    $ResultsCount = $ResultsCount[1]                        # Get the second item from the split up string. 
                                                            # Which would be the number after the colon.
    $ResultsCount = [Int]$ResultsCount                      # Convert string to integer

    if ( $ResultsCount -eq 1) 
    {
        # Soft delete the test email from the inbox
        # Rerouting output into variable $null to suppress text output
        $null = New-ComplianceSearchAction `
                -SearchName $SearchName `
                -Purge `
                -PurgeType SoftDelete `
                -Confirm:$false `
                -Force  
        
        # This automatically creates a ComplianceSearchAction which is automatically 
        # named after the following syntax: <SearchName>_Purge

        $PurgeName = "$(  $SearchName  )_Purge"

        # Wait for the email to be deleted
        do
        {
            # Use sleep to decrease cpu load
            Start-Sleep -Seconds 1
        }
        until ( ( Get-ComplianceSearchAction $PurgeName ).Status -eq "Completed" )

        # Clean up
        Get-ComplianceSearchAction $PurgeName | Remove-ComplianceSearchAction -Confirm:$false
        Get-ComplianceSearch $SearchName | Remove-ComplianceSearch -Confirm:$false

        return $true
    }
    else 
    {
        Write-Host "Test unsuccessful. Found $( $ResultsCount ) results instead of one result."
        
        return $false
    }    
}

function Setup-AlertPolicy
{
    # Make sure the custom policy is active

    $ProtectionAlertCustomName = 'eDiscovery search started or exported - customized'

    try 
    {
        # Rerout text output to $null
        # Try to fetch the custom alert policy
        $null = Get-ProtectionAlert -Identity $ProtectionAlertCustomName -ErrorAction Stop
        Write-Host 'The custom alert policy has been set up already.'
    }
    catch # If the alert policy hasn't been set up yet
    {    
        # Create new alert policy
        # {(emailaddress -like "*@XYZ.de"-and Enabled -eq $true) -or (emailaddress -notlike "*" -and Enabled -eq $true)}
        $Filter = "Activity.Item -notlike '"+$global:SearchNamePrefix+"*'"
        $Filter = '{'+$Filter+'}'

        Write-Host "Filter = $( $Filter )"

        $ProtectionAlertParams =
        @{
            Name = $ProtectionAlertCustomName
            Category = 'ThreatManagement'
            Severity = 'Medium'
            AggregationType = 'None'
            NotifyUser = 'TenantAdmins'
            ThreatType = 'Activity'
            NotifyUserOnFilterMatch = $true
            NotificationEnabled = $true
            Operation = 'eDiscoverySearchStartedOrExported'
            Filter = $Filter
            Comment = 'The alert policy is automatically set up by the script "Solarwinds RMM check: Test Email Deliverability". The alert is triggered when users start content searches or eDiscovery searches or when search results are downloaded or exported -V1.0.0.0.'
        }

        # Create new alert policy
        # Rerouting output into variable $null to suppress text output
        $null = New-ProtectionAlert @ProtectionAlertParams

        Write-Host 'The custom alert policy has been set up.'
    }

    # Make sure the default policy is inactive

    $DefaultPolicyName = 'eDiscovery search started or exported'
    $DefaultPolicy = Get-ProtectionAlert -Identity $DefaultPolicyName

    if ( $DefaultPolicy.Disabled -eq $false )
    {
        Write-Host 'The default alert policy is enabled.'
        Write-Host 'Please disable the alert policy first. This task cannot be automated via PowerShell.'
        $SolarwindsRMMExitCode = 1
        Exit
    }
    elseif ( $DefaultPolicy.Disabled -eq $true )
    {
        Write-Host 'The default alert policy has is disabled. Continuing.'
    }
    else 
    {
        Write-Host 'Unknown error: DefaultPolicy.Disabled is neither true nor false.'
        $SolarwindsRMMExitCode = 1
        Exit
    }
}

<# 
    Script entry point
#>

# Connect to Exchange Online
if (-not ((Get-PSSession).ComputerName -contains "outlook.office365.com")) {
    "doesnt exist"
    Connect-ExchangeOnline -CertificateThumbPrint '18DF396576DCB1F61BBB15F4BCE18FB4C8E5AA4A' -AppID 'd1186226-581c-44e6-a96b-78d7b90cc8cf' -Organization 'itcenterengels.onmicrosoft.com' -ShowBanner:$false # Muss als Administrator ausgeführt werden, da das Zertifikat sonst nicht gefunden wird (oder so)
} else {
    if ((Get-PSSession).Where({$_.ComputerName -eq "outlook.office365.com"}).State -eq "Opened") {
        "is open"
    } else {
        "has been closed"
        Connect-ExchangeOnline -CertificateThumbPrint '18DF396576DCB1F61BBB15F4BCE18FB4C8E5AA4A' -AppID 'd1186226-581c-44e6-a96b-78d7b90cc8cf' -Organization 'itcenterengels.onmicrosoft.com' -ShowBanner:$false
    }
}

<# 
# Connect to Compliance and Security
if (-not ((Get-PSSession).Where({$_.ComputerName -like "*compliance.protection.outlook.com"}))) {
    "doesnt exist"
    Connect-IPPSSession
} else {
    if (((Get-PSSession).Where({$_.ComputerName -like "*compliance.protection.outlook.com"})).State -eq "Opened") {
        "is open"
    } else {
        "has been closed"
        Connect-IPPSSession
    }
} 
#>

# Make sure no alarm will be triggered by the Compliance Search
# Setup-AlertPolicy

Write-Host "TesteEmailRecipients:"
$TestEmailRecipients

$global:HashTests = @{}

foreach ($testEmailRecipient in $TestEmailRecipients) {
    $global:HashTests[$testEmailRecipient] = 0
    
    Write-Host "Test:" -ForegroundColor Yellow
    $global:HashTests

    if (
        Send-TestEmail `
            -TestEmailSubject $TestEmailSubject `
            -TestEmailRecipient $TestEmailRecipient `
            -TestEmailSender $TestEmailSender `
            -TestEmailSenderPassword $TestEmailSenderPassword
    ) 
    {
        Write-Host ""
        Write-Host "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - Test email successfully send from GMail."
        Write-Host "  From: $( $TestEmailSender )"
        Write-Host "  To: $( $TestEmailRecipient )"
        Write-Host "  Subject: $( $TestEmailSubject )."
    }
}

Write-Host "Tests:" -ForegroundColor Yellow
$global:HashTests

Write-Host "All test mails sent."

Pause

# Define how often the script should check if the email was received
$Trys = 120
$Try = 0

# Define how many seconds the script should wait in between tries
$TryInterval = 1

# Wait for reception of test email
do
{
    Clear-Host
    $Try++
    $LastErrors = 0
    Write-Host "Starting try $( $Try ) out of $( $Trys )..."

    foreach ($testEmailRecipient in $TestEmailRecipients) {
        # Get recent Message Trace
        $MessageTrace = Get-MessageTrace `
            -StartDate ( Get-Date ).AddMinutes( -10 ) `
            -EndDate ( Get-Date ) `
            -RecipientAddress $TestEmailRecipient `
            -SenderAddress $TestEmailSender
        
        if ($MessageTrace) {
            Write-Host "Message Trace is not null."
        } else {
            Write-Host "Message Trace is null." -ForegroundColor DarkRed
        }

        foreach ($line in $MessageTrace) {
            Write-Host "Message Trace: Received = $( $line.Received ) # SenderAddress = $( $line.SenderAddress ) # RecipientAddress = $( $line.RecipientAddress ) # Subject = $( $line.Subject )"
        }

        if ( $MessageTrace ) # If the email hasn't arrived yet, the message trace will be empty
        {
            if ( $MessageTrace[0].Subject -eq $TestEmailSubject ) 
            {
                Write-Host "Found $( $testEmailRecipient )" -ForegroundColor DarkGreen
                $global:HashTests[$testEmailRecipient] = 1
                pause
            } else {
                Write-Host "Not found." -ForegroundColor Gray
                $LastErrors++
            }
        } else {
            Write-Host "Not found." -ForegroundColor Gray
            $LastErrors++
        }
    } 

    # Wait for next try
    Start-Sleep -Seconds $TryInterval
} until (($Try -eq $Trys) -or ($global:HashTests -notcontains 0)) 

if ($global:HashTests -contains 0) {
    Write-Host "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - Some test mails still haven't arrived at Exchange Online server after $( $Try ) out of $( $Trys ) tries." -ForegroundColor Red
    $SolarwindsRMMExitCode = 1
} else {
    Write-Host "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - All test emails successfully received on the Exchange Online server." -ForegroundColor Green
    $SolarwindsRMMExitCode = 0
}
