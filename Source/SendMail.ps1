<#---------------------------------------------------------------------------
   Sending Email in PowerShell

   Dominik Deak <deak.software@gmail.com>, DEAK Software

   This sample code demonstrates how to take advantage of .NET libraries in
   PowerShell to facilitate email communications without using an actual
   email client.

   Prerequisites:

   Before using the demo code, the following tools and libraries are
   required:

   * PowerShell 6: <https://github.com/PowerShell/PowerShell/releases>

   Instructions:

   Supply the outgoing email server configuration below, including the
   credentials needed for authentication.

   Supporting Resources:

   * Article: <https://deaksoftware.com.au/articles/sending_email_in_powershell/>
   * GitHub: <https://github.com/DEAKSoftware/Sending-Email-in-PowerShell/>

   Legal and Copyright:

   Released under the MIT License. Copyright 2018, DEAK Software.

   Permission is hereby granted, free of charge, to any person obtaining a
   copy of this software and associated documentation files (the "Software"),
   to deal in the Software without restriction, including without limitation
   the rights to use, copy, modify, merge, publish, distribute, sublicense,
   and/or sell copies of the Software, and to permit persons to whom the
   Software is furnished to do so, subject to the following conditions:

   The above copyright notice and this permission notice shall be included
   in all copies or substantial portions of the Software.

   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
   OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
   THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
   FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
   DEALINGS IN THE SOFTWARE.
  ---------------------------------------------------------------------------#>

Write-Host "Sending Email in PowerShell - Dominik Deak <deak.software@gmail.com>, DEAK Software" -ForegroundColor Yellow

<#---------------------------------------------------------------------------
   Configuration data.
  ---------------------------------------------------------------------------#>
# Outgoing email configuration used for sending messages
$outgoingUsername           = ""
$outgoingPassword           = ""
$outgoingServer             = ""
$outgoingPortSMTP           = 587   # Normally 25 (not secure), 465 (SSL), or 587 (TLS)
$outgoingEnableSSL          = $true
$outgoingFromAddress        = $outgoingUsername
$outgoingReplyToAddressList = @()
$outgoingToAddressList      = @()
$outgoingCCAddressList      = @()
$outgoingBCCAddressList     = @()
$outgoingAttachmentURLList  = @( "..\Attachments\File-01.pdf", "..\Attachments\File-02.svg" )

<#---------------------------------------------------------------------------
   Construct an SMTP client for the specified host name and credentials.
  ---------------------------------------------------------------------------#>
function makeSMTPClient
   {
   Param
      (
      [string] $server,
      [int] $port,
      [bool] $enableSSL,
      [string] $username,
      [string] $password
      )

   $smtpClient = New-Object Net.Mail.SmtpClient( $server, $port )
   $smtpClient.enableSSL = $enableSSL

   $smtpClient.credentials = New-Object System.Net.NetworkCredential( $username, $password )

   return $smtpClient;
   }

<#---------------------------------------------------------------------------
   Make a message.
  ---------------------------------------------------------------------------#>
function makeMessage
   {
   Param
      (
      [string] $fromAddress,
      [string[]] $replyToAddressList,
      [string[]] $toAddressList,
      [string[]] $ccAddressList,
      [string[]] $bccAddressList,
      [string[]] $attachmentURLList,
      [string] $subject,
      [string] $body,
      [bool] $isHTML
      )

   $message = New-Object Net.Mail.MailMessage

   $message.from = $fromAddress

   foreach ( $replyToAddress in $replyToAddressList )
      {
      $message.replyToList.add( $replyToAddress )
      }

   foreach ( $toAddress in $toAddressList )
      {
      $message.to.add( $toAddress )
      }

   foreach ( $ccAddress in $ccAddressList )
      {
      $message.cc.add( $ccAddress )
      }

   foreach ( $bccAddress in $bccAddressList )
      {
      $message.bcc.add( $bccAddress )
      }

   foreach ( $attachmentURL in $attachmentURLList )
      {
      # Resolve relative attachment paths to absolute paths
      $attachmentURL = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath( $attachmentURL )

      $message.attachments.add( $attachmentURL )
      }

   $message.subject = $subject
   $message.body = $body
   $message.isBodyHTML = $isHTML

   return $message
   }

<#---------------------------------------------------------------------------
   Run the demo.
  ---------------------------------------------------------------------------#>
Write-Host "Connecting to SMTP server: $outgoingServer`:$outgoingPortSMTP"

$smtpClient = makeSMTPClient `
   $outgoingServer $outgoingPortSMTP $outgoingEnableSSL `
   $outgoingUsername $outgoingPassword

Remove-Variable -Name outgoingPassword

$outgoingMessage = makeMessage `
   $outgoingFromAddress $outgoingReplyToAddressList `
   $outgoingToAddressList $outgoingCCAddressList `
   $outgoingBCCAddressList $outgoingAttachmentURLList `
   "Test Subject 1" "This is test message 1." $false

try {
   Write-Host "Sending message 1..."

   $smtpClient.send( $outgoingMessage )

   Write-Host "Sending message 2..."

   $smtpClient.send(
      $outgoingFromAddress, $outgoingToAddressList[0],
      "Test Subject 2", "This is test message 2." )
   }

catch { Write-Error "Caught SMTP client exception:`n`t$PSItem" }

Write-Host "Disconnecting from SMTP server: $outgoingServer`:$outgoingPortSMTP"

$smtpClient.dispose()

Remove-Variable -Name smtpClient
