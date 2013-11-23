# Gmail for PowerShell

A PowerShell module for managing your Gmail, with all the tools you'll need. Search, 
read and send emails, archive, mark as read/unread, delete emails, 
and manage labels.

__This library is still under development.__

## Installation

You can install it easily using chocolatey:

    cinst Gmail.ps

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
PS> $gmail | Get-Mailbox
PS> $gmail | Get-Mailbox | Filter-Message -Unread
```

Get the messages marked as Important by Gmail:

```powershell
PS> $gmail | Get-Mailbox "Important"
```

With `Get-Mailbox` you can access the `"All Mail"`, `"Starred`", `"Drafts"`, `"Important"`, `"Sent Mail"` and `"Spam"` folders

Filter with some criteria:

```powershell
PS> $gmail | Get-Mailbox | Filter-Message -After "2011-06-01" -Before "2012-01-01"
PS> $gmail | Get-Mailbox | Filter-Message -On "2011-06-01"
PS> $gmail | Get-Mailbox | Filter-Message -From "x@gmail.com"
PS> $gmail | Get-Mailbox | Filter-Message -To "y@gmail.com"
```

Combine flags and options:

```powershell
PS> $gmail | Get-Mailbox | Filter-Message -Unread -From "myboss@gmail.com"
```

Browsing labeled emails is similar to working with the inbox.

```powershell
PS> $gmail | Get-Mailbox -Label "Important"
```

You can count the messages too:

```powershell
PS> $gmail | Get-Mailbox | Filter-Message -Unread | Count-Message
PS> $gmail | Count-Message
```
    
Also you can manipulate each message using block style. Remember that every message in a conversation/thread will come as a separate message.

```powershell
PS> $messages = $gmail | Get-Mailbox | Filter-Message -Unread | Select-Object -Last 10
PS> foreach ($msg in $messages) {
>>     $msg | Update-Message -Read # you can use -Unread, -Spam, -Star, -Unstar, -Archive too
>> }
```
    
### Working with emails!

Delete emails from X:

```powershell
PS> $gmail | Get-Mailbox | Filter-Message -From "x@gmail.com" | ForEach-Object { Remove-Message $_ }
```

Save all attachments in the "Important" label to a local folder:

```powershell
PS> $messages = $gmail | Get-Mailbox -Label "Important" | Get-Message
PS> foreach ($msg in $messages) {
>>     if ($msg.HasAttachments) {
>>         $msg.FetchAttachments($folder)
>>     }
>> }
```

Save just the first attachment from the newest unread email:

```powershell
PS> $msg = $gmail | Get-Mailbox | Filter-Message -Unread | Select-Object -Last 1
PS> $msg.Fetch()
PS> $msg.Attachments[0].SaveTo($location)
```

Get all labels applied to a message:

```powershell
$msg | Get-Label
```

Add a label to a message (or remove it):

```powershell
$msg | Set-Label "Important"
$msg | Remove-Label "Important"
```

You can apply multiple lables:

```powershell
$msg | Set-Label "Important","Banking"
```

The example above will raise error when you don't have one of the specified labels. You can avoid this using:

```powershell
$msg | Set-Label "Important","Banking" -Force # If one of the labels does't exist, it will be automatically created now
```

You can also move message to a label/mailbox:

```powershell
$msg | Move-Message -Label "Test"
$msg | Move-Message "All Mail"
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
* Write tests
* Prettify the output
* Send mail via Google's SMTP servers

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
