# Gmail for PowerShell

A PowerShell module for managing your Gmail, with all the tools you'll need. Search, 
read and send emails, archive, mark as read/unread, delete emails, 
and manage labels.

__This library is still under development.__

## Installation

You can install it easily using chocolatey:

    cinst posh-gmail
    
Or install it manually:

    git clone https://github.com/nikoblag/gmail.ps.git
    cd gmail.ps
    .\install.ps1

## Features

* Manage labels

## Usage:

### Authenticating gmail sessions

This will you automatically log in to your account. 

```powershell
PS> $gmail = New-GmailSession
PS> # play with your gmail...
PS> $gmail | Remove-GmailSession
```

If you use `Enter-GmailSession` and pass a block, the session will be passed into the block, 
and will be logged out after the block is executed.

```powershell
PS> Enter-GmailSession -Credential (Get-Credential) -Script {
PS>     # play with your gmail...
PS> }
```

You can also check which accounts are logged in at any time:

```powershell
PS> Get-GmailSession
```

### Gathering emails
    
Get the messages in the inbox:

```powershell
PS> $gmail | Get-Inbox
PS> $gmail | Get-Inbox | Filter-Message -Unread
```

Filter with some criteria:

```powershell
PS> $gmail | Get-Inbox | Filter-Message -After "2011-06-01" -Before "2012-01-01"
PS> $gmail | Get-Inbox | Filter-Message -On "2011-06-01"
PS> $gmail | Get-Inbox | Filter-Message -From "x@gmail.com"
PS> $gmail | Get-Inbox | Filter-Message -To "y@gmail.com"
```

Combine flags and options:

```powershell
PS> $gmail | Get-Inbox | Filter-Message -Unread -From "myboss@gmail.com"
```

Browsing labeled emails is similar to working with the inbox.

```powershell
PS> $gmail | Get-Mailbox -Label "Important"
```

You can count the messages too:

```powershell
PS> $gmail | Get-Inbox | Filter-Message -Unread | Count-Message
PS> $gmail | Count-Message
```
    
Also you can manipulate each message using block style. Remember that every message in a conversation/thread will come as a separate message.

```powershell
PS> $messages = $gmail | Get-Inbox | Filter-Message -Unread -Last 10
PS> foreach ($msg in $messages) {
PS>     $msg | Update-Message -Read # you can use -Unread, -Spam, -Star, -Unstar, -Archive too
PS> }
```
    
### Working with emails!

Delete emails from X:

```powershell
PS> $gmail | Get-Inbox | Filter-Message -From "x@gmail.com" | ForEach-Object { Remove-Message $_ }
```

Save all attachments in the "Important" label to a local folder:

```powershell
PS> $messages = $gmail | Get-Mailbox -Label "Important"
PS> foreach ($msg in $messages) {
PS>     if ($msg.HasAttachments) {
PS>         $msg.FetchAttachments($folder)
PS>     }
PS> }
```

Save just the first attachment from the newest unread email:

```powershell
PS> $msg = $gmail | Get-Inbox | Filter-Message -Unread -Last 1
PS> $msg.Fetch()
PS> $msg.Attachments[0].SaveTo($location)
```

Add a label to a message:

```powershell
$msg | Update-Message -Label "Important"
```

Example above will raise error when you don't have the `Important` label. You can avoid this using:

```powershell
$msg | Update-Message -Label "Important" -Force # The `Important` label will be automatically created now
```

You can apply multiple lables:

```powershell
$msg | Update-Message -Label "Important","Banking"
```

You can also move message to a label/mailbox:

```powershell
$msg | Move-Message -Label "Test"
```

### Managing labels

With the Gmail module you can also manage your labels. You can get list of defined labels:

```powershell
$gmail | Get-Label
```

Create new label:

```powershell
$gmail | New-Label -Name "MyLabel"
```

Remove labels:

```powershell
$gmail | Remove-Label -Name "MyLabel"
```

Or check if given label exists:

```powershell
$gmail | Get-Label -Name "SomeLabel" # returns null if the label doesn't exist
```

## Roadmap
* Search emails
* Read emails 
* Emails: label, archive, delete, mark as read/unread/spam, star
* Write tests
* Move between labels/mailboxes
* Send mail via Google's SMTP servers

## Contributing

1. Fork it.
2. Create a branch (`git checkout -b my_feature`)
3. Commit your changes (`git commit -am "Added Feature"`)
4. Push to the branch (`git push origin my_feature`)
5. Open a [Pull Request][1]
6. Enjoy an ice cream and wait

## Author

* Nikolay Blagoev [https://github.com/nikoblag]

## Copyright

* Copyright (c) 2013 Nikolay Blagoev

See LICENSE for details.
