# Using SMTP in PowerShell

Dominik Deák


## 1 Introduction

This repository contains sample code demonstrating how to take advantage of .NET libraries in PowerShell to facilitate email communications without using an actual email client.

**Warning:** This is a demo only. These scripts are not intended to be used directly in a production environment.


## 2 Instructions

### 2.1 Prerequisites

Before using the demo code, the following tools and libraries are required. You may skip this section if these prerequisites are already installed.

* PowerShell 6: <https://github.com/PowerShell/PowerShell/releases>

### 2.2 Configuration

1. Create a new email address dedicated for experimentation purposes. This way you don't need to worry about accidentally deleting important messages.

2. Edit the PowerShell script file `Source\SendMail.ps1` and supply the outgoing SMTP server configuration, including the credentials needed for authentication:

	```powershell
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
	```


## 3 Supporting Resources

* [Using SMTP in PowerShell](https://deaksoftware.com.au/articles/using_smtp_in_powershell/) - Main article


## 4 Legal and Copyright

Released under the [MIT License](./license.md).

Copyright 2018, DEAK Software