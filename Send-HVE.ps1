# Prompt user for configuration
$TenantId = Read-Host "Enter your Tenant ID"
$ClientId = Read-Host "Enter your Client ID"
$ClientSecret = Read-Host "Enter your Client Secret"
$HveEmail = Read-Host "Enter the HVE sender email address"
$Recipient = Read-Host "Enter the recipient email address"
$smtpServer = "smtp-hve.office365.com"
$smtpPort = "587"

# Prompt for email subject and body
$subject = Read-Host "Enter the email subject"
$bodyText = Read-Host "Enter the email body text"

# Paths to MailKit & MimeKit
$mailKitPath = "C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.12.1\lib\net47\MailKit.dll"
$mimeKitPath = "C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.12.0\lib\net47\MimeKit.dll"

# Load dependencies
Add-Type -Path $mimeKitPath
Add-Type -Path $mailKitPath

# Create the body used for fetching the access token
$AccessTokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://outlook.office365.com/.default"
}

# Retrieve the access token
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $AccessTokenBody -ContentType "application/x-www-form-urlencoded"
$AccessToken = $TokenResponse.access_token
# Build email message
$message = New-Object MimeKit.MimeMessage
$message.From.Add([MimeKit.MailboxAddress]::Parse($hveEmail))
$message.To.Add([MimeKit.MailboxAddress]::Parse($recipient))
$message.Subject = $subject

$builder = New-Object MimeKit.BodyBuilder
$builder.TextBody = $bodyText
$message.Body = $builder.ToMessageBody()

# Send Email using MailKit
$client = New-Object MailKit.Net.Smtp.SmtpClient
try {
    $client.Connect($smtpServer, $SmtpPort, [MailKit.Security.SecureSocketOptions]::StartTls)
    $client.Authenticate([MailKit.Security.SaslMechanismOAuth2]::new($hveEmail, $accessToken))
    $client.Send($message)
    Write-Host "Email sent successfully."
} catch {
    Write-Error "Failed to send email: $_"
} finally {
    $client.Disconnect($true)
    $client.Dispose()
}
