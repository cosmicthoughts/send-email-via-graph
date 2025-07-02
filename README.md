# 📄 Sending PowerShell Reports via Microsoft Graph

This guide documents how to securely send PowerShell-generated reports via email using Microsoft Graph, the recommended method for modern Office 365 tenants.

---

## ✅ Overview

You will:
1. Register an app in Entra ID
2. Grant `Mail.Send` permissions
3. Authenticate with Microsoft Graph
4. Send email with attachments (like Excel reports)

---

## 🔧 Step 1: App Registration in Entra ID (Azure AD)

1. Go to [https://entra.microsoft.com](https://entra.microsoft.com)
2. Navigate to **App registrations** → **New registration**
3. Fill out:
   - **Name**: PowerShell Mailer
   - **Supported account types**: Single tenant
   - **Redirect URI**: Leave blank for now
4. Click **Register**

---

## 🔐 Step 2: Add Microsoft Graph Permissions

1. Go to **API permissions** → **Add a permission**
2. Choose **Microsoft Graph**
3. Select:
   - **Delegated permissions** for interactive use
   - **Application permissions** for automation
4. Add:
   - `Mail.Send`
   - (Optional: `User.Read` for delegated login)
5. Click **Grant admin consent**

---

## 🔑 Step 3: Authentication Options

### Option A: Delegated Auth (Interactive User)
```powershell
Connect-MgGraph -Scopes "Mail.Send"
```

### Option B: App-Only Auth with Certificate (Automation)
```powershell
# Create certificate
$cert = New-SelfSignedCertificate -CertStoreLocation "Cert:\\CurrentUser\\My" -Subject "CN=PowerShell Mailer" -KeySpec KeyExchange

# Export certificate
$password = ConvertTo-SecureString -String "YourStrongPassword" -Force -AsPlainText
Export-PfxCertificate -Cert "Cert:\\CurrentUser\\My\\$($cert.Thumbprint)" -FilePath "C:\\certs\\mailer.pfx" -Password $password
```

Upload the **.cer file** (public key) to **Certificates & secrets → Certificates** in your app registration.

---

## 📧 Step 4: Sending Email with Attachment

### Install Microsoft Graph SDK
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Example Script
```powershell
$to = "recipient@domain.com"
$from = "youruser@domain.com"
$subject = "Monthly PowerShell Report"
$body = @"
Hi,

Please find the attached report.

Regards,
Automation Script
"@

$filePath = "C:\\Reports\\MonthlyReport.xlsx"
$fileBytes = [System.IO.File]::ReadAllBytes($filePath)
$fileBase64 = [System.Convert]::ToBase64String($fileBytes)

$message = @{
  Message = @{
    Subject = $subject
    Body = @{ ContentType = "Text"; Content = $body }
    ToRecipients = @(@{ EmailAddress = @{ Address = $to } })
    Attachments = @(
      @{
        "@odata.type" = "#microsoft.graph.fileAttachment"
        Name = [System.IO.Path]::GetFileName($filePath)
        ContentBytes = $fileBase64
      }
    )
  }
  SaveToSentItems = $true
}

Connect-MgGraph -Scopes "Mail.Send"
Send-MgUserMail -UserId $from -BodyParameter $message
```

---

## 🛡️ App-Only Auth Usage (Automation)

```powershell
Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $certThumb
```

> Use this in scheduled scripts or unattended environments. The app must have **application permissions** and `Mail.Send` granted.

---

## 🕒 Step 5: Automate with Task Scheduler or Azure Automation

- Store report path and cert locally
- Schedule script to run at desired interval
- Log output and errors for auditing

---

## ⚠️ Important Notes

- App-only auth requires mailbox delegation for `SendAs` actions.
- Use `Export-Excel` to create reports in `.xlsx` format.
