Import-Module ExchangeOnlineManagement

$certBytes = [Convert]::FromBase64String($env:EXO_CERT_BASE64)
$certPath = "$env:RUNNER_TEMP\cert.pfx"
[IO.File]::WriteAllBytes($certPath, $certBytes)
$securePassword = ConvertTo-SecureString $env:EXO_CERT_PASSWORD -AsPlainText -Force

Connect-ExchangeOnline `
  -AppId $env:EXO_APP_ID `
  -Organization $env:EXO_TENANT_ID `
  -CertificateFilePath $certPath `
  -CertificatePassword $securePassword

$template = Get-Content "./template/signature.html" -Raw
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

foreach ($mailbox in $mailboxes) {

    $user = Get-User $mailbox.Identity

    $displayName = $user.DisplayName
    $jobTitle = $user.Title
    $businessPhone = $user.Phone
    $mobilePhone = $user.MobilePhone
    $email = $mailbox.PrimarySmtpAddress

    $jobBlock = ""
    if ($jobTitle) {
        $jobBlock = "<span style='color:#062f3e; font-family:Arial; font-size:12px;'>$jobTitle</span><br>"
    }

    $businessBlock = ""
    if ($businessPhone) {
        $businessBlock = "<span style='color:#0086BE; font-weight:525;'>P:</span> $businessPhone<br>"
    }

    $mobileBlock = ""
    if ($mobilePhone) {
        $mobileBlock = "<span style='color:#0086BE; font-weight:525;'>M:</span> $mobilePhone<br>"
    }

    $emailBlock = ""
    if ($email) {
        $emailBlock = "<span style='color:#0086BE; font-weight:525;'>E:</span> $email"
    }

    $signature = $template `
        -replace "%%DisplayName%%", $displayName `
        -replace "%%JOBTITLE_BLOCK%%", $jobBlock `
        -replace "%%BUSINESSPHONE_BLOCK%%", $businessBlock `
        -replace "%%MOBILEPHONE_BLOCK%%", $mobileBlock `
        -replace "%%EMAIL_BLOCK%%", $emailBlock

    Write-Host "Updating $email"

    Set-MailboxMessageConfiguration `
        -Identity $email `
        -SignatureHtml $signature `
        -AutoAddSignature $true `
        -AutoAddSignatureOnReply $true `
        -SignatureText $null
}

Get-MailboxMessageConfiguration -Identity fportes@summitcare.com | 
Select AutoAddSignature, AutoAddSignatureOnReply

Disconnect-ExchangeOnline -Confirm:$false
