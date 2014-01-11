# Gmail for PowerShell

A PowerShell module for managing your Gmail, with all the tools you'll need. Search, 
read and send emails, archive, mark as read/unread, delete emails, 
and manage labels.

__This module is still under development.__

## Installation

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

## Usage:

### Authenticating Gmail sessions

To authenticate a Gmail session, use `New-GmailSession` and provide your username and password. 
If you want to be automatically logged in to your account, create a generic credential using the Windows Credential Manager: 
go to 'Control Panel\User Accounts and Family Safety\Credential Manager', click 'Add a generic credential', then type your
Gmail username and password, and use `Gmail.ps:default` as address. 

```powershell
PS> $gmail = New-GmailSession
PS> # play with your gmail...
PS> $gmail | Remove-GmailSession
```

If you use `Invoke-GmailSession` and pass a block, the session will be passed into the block, 
and will be logged out after it's executed. Use `$args` to access the session or parameterize the block: 

```powershell
PS> Invoke-GmailSession -ScriptBlock {
>>      param($gmail) # to use $gmail instead of $args
>>      $gmail | Get-Label
>>  }
```

You can also check which accounts are logged in at any time using `Get-GmailSession` or close all of them with `Clear-GmailSession`.

### Gathering emails
    
Get the messages in the inbox:

```powershell
PS> $inbox = $gmail | Get-Mailbox
PS> $inbox | Get-Message -Unread
```

Get the messages marked as Important by Gmail:

```powershell
PS> $gmail | Get-Mailbox "Important" | Get-Message
```

With `Get-Mailbox` you can access the `"All Mail"`, `"Starred"`, `"Drafts"`, `"Important"`, `"Sent Mail"` and `"Spam"` folders

Filter with some criteria:

```powershell
PS> $inbox | Get-Message -After "2011-06-01" -Before "2012-01-01"
PS> $inbox | Get-Message -On "2011-06-01"
PS> $inbox | Get-Message -From "x@gmail.com"
PS> $inbox | Get-Message -To "y@gmail.com"
```

Combine flags and options:

```powershell
PS> $inbox | Get-Message -Unread -From "myboss@gmail.com"
```

Browsing labeled emails is similar to working with the inbox.

```powershell
PS> $gmail | Get-Mailbox -Label "Important"
```

You can count the messages too:

```powershell
PS> $inbox | Measure-Message
PS> $inbox | Get-Message -Unread | Measure-Message
```
    
Also you can manipulate each message using block style. Remember that every message in a conversation/thread will come as a separate message.

```powershell
PS> $messages = $inbox | Get-Message -Unread | Select-Object -Last 10
PS> foreach ($msg in $messages) {
>>     $msg | Update-Message -Read # you can use -Unread, -Spam, -Star, -Unstar, -Archive too
>> }
```
    
### Working with emails!

Delete all emails from X:

```powershell
PS> $inbox | Get-Message -From "x@gmail.com" | Remove-Message
```

Save all attachments in the "Important" label to a local folder. 
Note that without the `-Prefetch` parameter, no attachments will be downloaded from the server:

```powershell
PS> $gmail | Get-Mailbox -Label "Important" | Get-Message -Prefetch | Save-Attachment $folder
```

Save just the first attachment from the newest unread email:

```powershell
PS> $msg = $inbox | Get-Message -Unread -HasAttachment | Select-Object -Last 1
PS> $fetchedMsg = $msg | Receive-Message # or use -Prefetch on Get-Message above
PS> $fetchedMsg.Attachments[0].Save($location)
```

Get all labels applied to a message:

```powershell
PS> $msg | Get-Label
```

Add a label to a message (or remove it):

```powershell
PS> $msg | Set-Label "Important"
PS> $msg | Remove-Label "Important"
```

You can apply multiple lables:

```powershell
PS> $msg | Set-Label "Important","Banking"
```

The example above will raise error when you don't have one of the specified labels. You can avoid this using:

```powershell
PS> $msg | Set-Label "Important","Banking" -Force # If one of the labels does't exist, it will be automatically created now
```

You can also move message to a label/mailbox:

```powershell
PS> $msg | Move-Message -Label "Test"
PS> $msg | Move-Message "All Mail"
```

### Managing labels

With the Gmail module you can also manage your labels. You can get list of defined labels:

```powershell
PS> $gmail | Get-Label
```

Create new label:

```powershell
PS> $gmail | New-Label -Name "MyLabel"
```

Remove labels:

```powershell
PS> $gmail | Remove-Label -Name "MyLabel"
```

Or check if given label exists:

```powershell
PS> $gmail | Get-Label -Name "SomeLabel" # returns null if the label doesn't exist
```

Cmdlets that can take a label name as a parameter value (`Get-Message`, `Move-Message`, `Set-Label` and `Remove-Label`) support automatic completion:

```powershell
PS> $gmail | Remove-Label <tab> # use <tab> key to cycle through all existing labels
```

## Roadmap
* Write tests
* Send mail via Google's SMTP servers
* Backup/restore all messages and labels

## Contributing

1. Fork it.
2. Create a branch (`git checkout -b my_feature`)
3. Commit your changes (`git commit -am "Added Feature"`)
4. Push to the branch (`git push origin my_feature`)
5. Open a [Pull Request](https://github.com/nikoblag/Gmail.ps/compare/)
6. Enjoy an ice cream and wait

## Author

* Nikolay Blagoev [https://github.com/nikoblag]

## Copyright

* Copyright (c) 2013 Nikolay Blagoev
* Copyright (c) 2013 Andy Edinborough - AE.Net.Mail library
* Copyright (c) 2012 Tobias Burger - Get-StoredCredential cmdlet

See LICENSE for details.
