# Gmail for PowerShell

A PowerShell module for managing your Gmail, with all the tools you'll need. Search, 
read and send emails, archive, mark as read/unread, delete emails, 
and manage labels.

**This module is still under development.**

## Table of contents

* [Install](#install)
* [Features](#features)
* [Get help](#get-help)
* [Commands](#commands)
	* [New-GmailSession](#new-gmailsession)
	* [Remove-GmailSession](#remove-gmailsession)
	* [Invoke-GmailSession](#invoke-gmailsession)
	* [Get-GmailSession](#get-gmailsession)
	* [Clear-GmailSession](#clear-gmailsession)
	* [Get-Mailbox](#get-mailbox)
	* [Get-Message](#get-message)
	* [Update-Message](#update-message)
	* [Receive-Message](#receive-message)
	* [Move-Message](#move-message)
	* [Remove-Message](#remove-message)
	* [Measure-Message](#measure-message)
	* [Get-Conversation](#get-conversation)
	* [Save-Attachment](#save-attachment)
	* [Get-Label](#get-label)
	* [New-Label](#new-label)
	* [Set-Label](#set-label)
	* [Remove-Label](#remove-label)
* [Roadmap](#roadmap)
* [History](#history)
* [Author](#author)
* [Third Party Libraries](#third-party-libraries)
* [Contributing](#contributing)
* [License](#license)

## Install

If you have [PsGet](http://psget.net/) installed you can simply execute:

```powershell
Install-Module Gmail.ps
```

Or install it manually:

    git clone https://github.com/nikoblag/Gmail.ps.git
    cd Gmail.ps
    .\install.ps1

## Features

* Read emails
* Search emails
* (Update) emails: label, archive, delete, mark as read/unread/spam, star
* Manage labels
* Move between labels/mailboxes
* Automatic authentication, using the Windows Credential Manager

## Get help

* List of all available commands

    ```powershell
	Get-Command -Module Gmail.ps
    ```

* Help for a specific command.

    ```powershell
	Get-Help <command>
    ```

## Commands

***For more detailed information about a command use the help.***

### New-GmailSession

Opens a connection to a Gmail account using the specified credentials and creates a new session. If a generic credential is 
created using the Windows Credential Manager (address: `Gmail.ps:default`), a session is automatically created using the 
stored credentials each time the cmdlet is executed without a `-Credential` parameter.

```powershell
New-GmailSession [[-Credential] <PSCredential>] [<CommonParameters>]
```

#### Parameters

Name          | Pipeline input | Default
---           | ---            | ---
`-Credential` | No             | `Get-StoredCredential Gmail.ps:default` or `Get-Credential`

#### Examples

1. Authenticating a Gmail session using the stored credential in the `Gmail.ps:default` entry. 
   If there is no credential stored a prompt for username and password will be displayed.

    ```powershell
	$gmail = New-GmailSession
	# play with your gmail...
    ```

### Remove-GmailSession

Closes the connection to Gmail and destroys the session.

```powershell
Remove-GmailSession [-Session] <ImapClient> [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input
---        | ---
`-Session` | ByValue, ByPropertyName

#### Examples

1. Closing an already opened connection to a Gmail account:

    ```powershell
	$gmail | Remove-GmailSession
    ```

### Invoke-GmailSession

Creates a new Gmail session and passes it to a script block. Once the block is executed, the session is automatically closed.

```powershell
Invoke-GmailSession [[-Credential] <PSCredential>] [-ScriptBlock] <ScriptBlock> [<CommonParameters>]
```

#### Parameters

Name           | Pipeline input | Default
---            | ---            | ---
`-Credential`  | No             | `Get-StoredCredential Gmail.ps:default` or `Get-Credential`
`-ScriptBlock` | No             |

#### Examples

1. Creates a Gmail session, returns the number of messages in the Inbox and then closes the session.
   The automatically created session can be accessed inside the script block via the `$args` variable.

    ```powershell
	Invoke-GmailSession -ScriptBlock {
	    $args | Count-Message
	}
    ```

2. Creates a Gmail session, returns all the labels used in that account and then closes the session.
   The automatically created session can be accessed inside the script block via the `$gmail` variable.

    ```powershell
	Invoke-GmailSession -ScriptBlock {
	    param($gmail)
	    $gmail | Get-Label
	}
    ```

### Get-GmailSession

Returns a list of all opened Gmail sessions.

```powershell
Get-GmailSession
```

### Clear-GmailSession

Closes all opened Gmail sessions.

```powershell
Clear-GmailSession
```

### Get-Mailbox

Returns the `Inbox` if no parameters are specified, an existing Label or one of the default 
Gmail folders (`All Mail`, `Starred`, `Drafts`, `Important`, `Sent Mail`, `Spam`).

> **Alias:** `Select-Mailbox`

```powershell
Get-Mailbox -Session <ImapClient> [[-Name] <String>] [<CommonParameters>]

Get-Mailbox -Session <ImapClient> [-Label <String>] [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input          | Default (List of possible values)
---        | ---                     | ---
`-Session` | ByValue, ByPropertyName |
`-Name`    | ByPropertyName          | `Inbox` (`All Mail`, `Starred`, `Drafts`, `Important`, `Sent Mail`, `Spam`)
`-Label`   | ByPropertyName          |

#### Examples

1. Get the unread messages in the inbox:

    ```powershell
	$inbox = $gmail | Get-Mailbox
	$inbox | Get-Message -Unread
    ```

2. Get the messages marked as Important by Gmail:

    ```powershell
	$gmail | Get-Mailbox "Important" | Get-Message
    ```

### Get-Message

Returns a (filtered) list of the messages inside a selected mailbox (see [`Get-Mailbox`](#get-mailbox)).
The returned messages will have their body and attachments downloaded only if the `-Prefetch` parameter is specified. 

Every listed message has a set of flags indicating the message's status and properties.

Flag | Meaning
---  | ---
`u`  | Is unread
`f`  | Is fetched
`i`  | Is important
`s`  | Is starred
`a`  | Has attachment

Any flag may be unset. An unset flag is the equivalent of "is not" and is represented as a `-` character.
`--i-a` means the message is not Unread, is not Fetched, is Important, is not Starred and has atleast one attachment.

Supports automatic name completion for the existing labels.

> **Alias:** `Filter-Message`

```powershell
Get-Message [-Session] <ImapClient> 
			[[-From] <String>] [[-To] <String>] 
			[[-On] <DateTime>] [[-After] <DateTime>] [[-Before] <DateTime>] 
			[[-Cc] <String>] [[-Bcc] <String>] 
			[[-Subject] <String>] [[-Text] <String>] [[-Body] <String>] 
			[[-Label] <String[]>] [[-FileName] <String>] [[-Category] <String>] 
			[-Unread ] [-Read ] [-Starred ] [-Unstarred ] [-HasAttachment ] 
			[-Answered ] [-Draft ] [-Undraft ] [-Prefetch ] [<CommonParameters>]
```

#### Parameters

Name             | Pipeline input          | Default (List of possible values)
---              | ---                     | ---
`-Session`       | ByValue, ByPropertyName |
`-From`          | No                      | 
`-To`            | No                      |
`-On`            | No                      |
`-After`         | No                      |
`-Before`        | No                      |
`-Cc`            | No                      |
`-Bcc`           | No                      |
`-Subject`       | No                      |
`-Text`          | No                      |
`-Body`          | No                      |
`-Label`         | No                      |
`-FileName`      | No                      |
`-Category`      | No                      | *none* (`Primary`, `Personal`, `Social`, `Promotions`, `Updates`, `Forums`)
`-Unread`        | No                      |
`-Read`          | No                      |
`-Starred`       | No                      |
`-Unstarred`     | No                      |
`-HasAttachment` | No                      |
`-Answered`      | No                      |
`-Draft`         | No                      |
`-Undraft`       | No                      |
`-Prefetch`      | No                      |

#### Examples

1. Get the unread messages in the inbox:

    ```powershell
	$inbox = $gmail | Get-Mailbox
	$inbox | Get-Message -Unread
    ```

2. Get the messages marked as Important by Gmail:

    ```powershell
	$gmail | Get-Mailbox "Important" | Get-Message
    ```

3. Filter with some criteria:

    ```powershell
	$inbox | Get-Message -After "2011-06-01" -Before "2012-01-01"
	$inbox | Get-Message -On "2011-06-01"
	$inbox | Get-Message -From "x@gmail.com"
	$inbox | Get-Message -To "y@gmail.com"
    ```

4. Combine flags and options:

    ```powershell
	$inbox | Get-Message -Unread -From "myboss@gmail.com"
    ```

### Update-Message

Archives, marks as spam, as read/undead and adds/removes a star from a given message.

```powershell
Update-Message -Session <ImapClient> -Message <MailMessage> [-Read ] [-Star ] [-Archive ] [-Spam ] [<CommonParameters>]

Update-Message -Session <ImapClient> -Message <MailMessage> [-Read ] [-Unstar ] [-Archive ] [-Spam ] [<CommonParameters>]

Update-Message -Session <ImapClient> -Message <MailMessage> [-Unread ] [-Star ] [-Archive ] [-Spam ] [<CommonParameters>]

Update-Message -Session <ImapClient> -Message <MailMessage> [-Unread ] [-Unstar ] [-Archive ] [-Spam ] [<CommonParameters>]
```

#### Parameters

Name        | Pipeline input
---         | ---
`-Session`  | ByValue, ByPropertyName
`-Message`  | ByValue
`-Read `    | No
`-Unread `  | No
`-Archive ` | No
`-Star `    | No
`-Unstar `  | No
`-Spam `    | No

#### Examples

1. Each message can be manipulated using block style. Remember that every message in a conversation/thread will come as a separate message.

    ```powershell
	$messages = $inbox | Get-Message -Unread | Select-Object -Last 10
	foreach ($msg in $messages) {
	    $msg | Update-Message -Read # you can use -Unread, -Spam, -Star, -Unstar, -Archive too
	}
    ```

### Receive-Message

Fetches the whole message from the server (including the body and the attachments).

```powershell
Receive-Message [-Session] <ImapClient> [-Message] <MailMessage> [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input
---        | ---
`-Session` | ByValue, ByPropertyName
`-Message` | ByValue

#### Examples

1. To read the actual body of a message you have to first fetch it from the Gmail servers:

    ```powershell
	$msg = $inbox | Get-Message -From "x@gmail.com" | Receive-Message
	$msg.Body # returns the body of the message
    ```

### Move-Message

Moves a message to a different mailbox or label. 

Supports automatic name completion for the existing labels.

```powershell
Move-Message -Session <ImapClient> -Message <MailMessage> [-Mailbox] <String> [<CommonParameters>]

Move-Message -Session <ImapClient> -Message <MailMessage> -Label <String> [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input
---        | ---
`-Session` | ByValue, ByPropertyName
`-Message` | ByValue
`-Mailbox` | No
`-Label`   | No

#### Examples

1. Move the message to the `All Mail` mailbox:

    ```powershell
    $msg | Move-Message "All Mail"
    ```

2. Move the message to the `Test` label:

    ```powershell
	$msg | Move-Message -Label "Test"
    ```

### Remove-Message

Sends a message to the Gmail's `Trash` folder.

```powershell
Remove-Message [-Session] <ImapClient> [-Message] <MailMessage> [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input
---        | ---
`-Session` | ByValue, ByPropertyName
`-Message` | ByValue

#### Examples

1. Delete all emails from X:

    ```powershell
	$inbox | Get-Message -From "x@gmail.com" | Remove-Message
    ```

### Measure-Message

Returns the number of messages in a mailbox (supports labels too).

> **Alias:** `Count-Message`

```powershell
Measure-Message [-Session] <ImapClient> [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input
---        | ---
`-Session` | ByValue, ByPropertyName

#### Examples

1. Count the messages in the inbox:

    ```powershell
	$inbox | Measure-Message
    ```

2. Count the important messages:

    ```powershell
	$gmail | Get-Mailbox "Important" | Measure-Message
    ```

3. Note that `Measure-Message` will return the number of all messages in the selected mailbox, not the number of the returned messages (if any). To count the returned messages, use `Measure-Object`. For example if we have 2 unread and 98 read messages in the `Important` mailbox:

    ```powershell
    # returns 100, the number of messages in `Important`
	$gmail | Get-Mailbox "Important" | Get-Message -Unread | Measure-Message

	# returns 2, the number of unread messages in `Important`
	$gmail | Get-Mailbox "Important" | Get-Message -Unread | Measure-Object
    ```

### Get-Conversation

Returns a list of messages that are part of a conversation.

> **Alias:** `Get-Thread`

```powershell
Get-Conversation [-Session] <ImapClient> [-Message] <MailMessage> [-Prefetch] [<CommonParameters>]
```

#### Parameters

Name        | Pipeline input
---         | ---
`-Session`  | ByValue, ByPropertyName
`-Message`  | ByValue
`-Prefetch` | No

#### Examples

1. Search the Inbox based on the message returned by [`Get-Message`](#get-message), 
   and return all messages that are part of that conversaton and are in the Inbox:

    ```powershell
	$gmail | Get-Mailbox "Inbox" | Get-Message -From "z@gmail.com" | Get-Conversaion
    ```

2. Search "All Mail" based on the message returned by [`Get-Message`](#get-message), 
   and return all messages that are part of that conversaton:

    ```powershell
	$gmail | Get-Mailbox "All Mail" | Get-Message -From "z@gmail.com" | Get-Conversaion
    ```

### Save-Attachment

Downloads the attachments of a message to a local folder.

```powershell
Save-Attachment [-Message <MailMessage>] [-Path] <String[]> [-PassThru ] [<CommonParameters>]

Save-Attachment [-Message <MailMessage>] -LiteralPath <String[]> [-PassThru ] [<CommonParameters>]
```

#### Parameters

Name           | Pipeline input
---            | ---
`-Message`     | ByValue
`-Path`        | No
`-LiteralPath` | No
`-PassThru`    | No

#### Examples

1. Save all attachments in the "Important" label to a local folder. 
   Note that without the `-Prefetch` parameter, no attachments will be downloaded:

    ```powershell
	$gmail | Get-Mailbox -Label "Important" | Get-Message -Prefetch | Save-Attachment $folder
    ```

2. Save just the first attachment from the newest unread email:

    ```powershell
	$msg = $inbox | Get-Message -Unread -HasAttachment | Select-Object -Last 1
    $fetchedMsg = $msg | Receive-Message # or use -Prefetch on Get-Message above
    $fetchedMsg.Attachments[0].Save($location)
    ```

### Get-Label

Returns the labels applied to a message or all labels that exist.

```powershell
Get-Label -Session <ImapClient> [-Message <MailMessage>] [[-Like] <String>] [-All ] [<CommonParameters>]
```

#### Parameters

Name (Alias)      | Pipeline input
---               | ---
`-Session`        | ByValue, ByPropertyName
`-Message`        | ByValue
`-Like` (`-Name`) | No
`-All`            | No

#### Examples

1. Get all labels applied to a message:

    ```powershell
	$msg | Get-Label
    ```

2. Get a list of the defined labels:

    ```powershell
	$gmail | Get-Label
    ```

3. Check if a label exists:

    ```powershell
	$gmail | Get-Label -Name "SomeLabel" # returns null if the label doesn't exist
	```

### New-Label

Creates a new label.

```powershell
New-Label [-Name] <String[]> -Session <ImapClient> [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input
---        | ---
`-Name`    | No
`-Session` | ByValue, ByPropertyName

### Set-Label

Applies a label to a message.

Supports automatic name completion for the existing labels.

> **Alias:** `Add-Label`

```powershell
Set-Label -Session <ImapClient> -Message <MailMessage> [-Name] <String[]> [-Force ] [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input
---        | ---
`-Session` | ByValue, ByPropertyName
`-Message` | ByValue
`-Name`    | No
`-Force`   | No

#### Examples

1. Apply a single or multiple labels:

	```powershell
	$msg | Set-Label "Important"
	$msg | Set-Label "Important","Banking"
	```

2. The example above will raise error if one of the specified labels doesn't exist. To avoid that, label creation can be forced:

    ```powershell
	$msg | Set-Label "Important","Banking" -Force
    ```

### Remove-Label

Removes a label from a message or deletes the label from the account.

Supports automatic name completion for the existing labels.

```powershell
Remove-Label [-Name] <String[]> -Session <ImapClient> [-Message <MailMessage>] [<CommonParameters>]
```

#### Parameters

Name       | Pipeline input
---        | ---
`-Name`    | No
`-Session` | ByValue, ByPropertyName
`-Message` | ByValue

## Roadmap

* Write tests
* Send mail via Google's SMTP servers
* Backup/restore all messages and labels

## History

Check [Release](https://github.com/nikoblag/Gmail.ps/releases) list.

## Author

* Nikolay Blagoev [https://github.com/nikoblag]

## Third Party Libraries

* [AE.Net.Mail](https://github.com/andyedinborough/aenetmail) library - Copyright (c) 2013 [Andy Edinborough](https://github.com/andyedinborough)
* [Get-StoredCredential](https://gist.github.com/toburger/2947424) cmdlet - Copyright (c) 2012 [Tobias Burger](https://github.com/toburger)

## Contributing

1. Fork it.
2. Create a branch (`git checkout -b my_feature`)
3. Commit your changes (`git commit -am "Added Feature"`)
4. Push to the branch (`git push origin my_feature`)
5. Open a [Pull Request](https://github.com/nikoblag/Gmail.ps/compare/)
6. Enjoy an ice cream and wait

## License

[MIT License](https://github.com/nikoblag/Gmail.ps/blob/master/LICENSE)
