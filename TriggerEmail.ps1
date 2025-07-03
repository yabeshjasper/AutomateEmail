# Define Gmail credentials
$gmailUser = "yabeshjasper17@gmail.com"
$gmailPassword = "qikc lvmy kpgh hpas"  # Use App Password(https://myaccount.google.com/apppasswords), not your normal Gmail password

# Load CSV
$recipients = Import-Csv -Path "/Users/YabeshJasper/Powershell/AutomateEmail/email_list.csv"

# SMTP settings
$smtpServer = "smtp.gmail.com"
$smtpPort = 587

# Define PDF file path (full path recommended)
$pdfAttachmentPath = "/Users/YabeshJasper/Powershell/AutomateEmail/YabeshJasper.pdf"

foreach ($recipient in $recipients) {
    $to = $recipient.Email
    $name = $recipient.Name

    # Compose email
    $subject = "Hello $name!"
    $body = "Dear $name,`n`nThis is a personalized message sent via PowerShell.`nHave a great day!"

    $mailMessage = New-Object system.net.mail.mailmessage
    $mailMessage.from = $gmailUser
    $mailMessage.To.add($to)
    $mailMessage.Subject = $subject
    $mailMessage.Body = $body
    $mailMessage.IsBodyHtml = $false

    # 🔗 Add PDF Attachment
    $attachment = New-Object System.Net.Mail.Attachment($pdfAttachmentPath)
    $mailMessage.Attachments.Add($attachment)

    $smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $smtp.EnableSsl = $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential($gmailUser, $gmailPassword);

    try {
        $smtp.Send($mailMessage)
        Write-Host "Email sent to $name <$to>"
    } catch {
        Write-Host "Failed to send to $to. Error: $_"
    }
}
