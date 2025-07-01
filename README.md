# Send-HVE
 This PowerShell script demonstrates how to send an email using High Volume Email (HVE) and OAuth2 authentication via MailKit.

## 🔧 Features

- Uses OAuth2 client credentials flow to authenticate with Microsoft identity platform.
- Sends email using MailKit and MimeKit libraries.
- Prompts user for configuration and message details.

## 🛠 Setup

1. Install MailKit and MimeKit via NuGet or download the DLLs:
   - `MailKit.dll`
   - `MimeKit.dll`

You can download the packages using NuGet:

https://www.nuget.org/packages/MailKit/

2. Update the script with the correct paths to the DLLs.

3. Ensure the following permissions are granted to your Azure AD app:
   - `SMTP.Send` (Application permission)
   - `Mail.Send` (if using Graph API)

## ⚙️ Required Configuration

The script will prompt for the following:

- Tenant ID
- Client ID
- Client Secret
- Sender (HVE) email address
- Recipient email address
- Email subject
- Email body text

## ▶️ Example Usage

```powershell
PS> .\Send-HVE-OAuth2.ps1
Enter your Tenant ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
Enter your Client ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
Enter your Client Secret: ********************
Enter the HVE sender email address: sender@example.com
Enter the recipient email address: recipient@example.com
Enter the email subject: Test Email
Enter the email body text: This is a test email sent using OAuth2 and MailKit.
```
