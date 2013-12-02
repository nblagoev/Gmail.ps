
$GmailSessions = @();

function New-GmailSession {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential = $($cr = (Get-StoredCredential Gmail.ps:default); if ($cr -eq $null) {Get-Credential} else {$cr})
    )

    $ghost = "imap.gmail.com"
    $gport = 993
    $guser = $Credential.UserName
    $gpass = $Credential.GetNetworkCredential().Password
    
    $session = New-Object -TypeName AE.Net.Mail.ImapClient -ArgumentList $ghost,$guser,$gpass,Login,$gport,$true,$false
    $GmailSessions += $session
    $session

<#
.Synopsis
    Creates a Gmail session.
.Description
    Opens a connection to a Gmail account using the specified credentials and creates a new session. If a generic credential is 
    created using the Windows Credential Manager (address: 'Gmail.ps:default'), a session is automatically created using the 
    stored credentials each time the cmdlet is executed without a -Credential parameter
.Parameter Credential
    The credentials that will be used to connect to Gmail.
.Link
    Remove-GmailSession
.Link
    Invoke-GmailSession
.Link
    Get-GmailSession
.Link
    Clear-GmailSession
#>
}

function Remove-GmailSession {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    $Session.Disconnect()

<#
.Synopsis
    Removes a Gmail session.
.Description
    Closes the connection to Gmail and destroys the session.
.Parameter Credential
    The credentials that will be used to connect to Gmail.
.Link
    New-GmailSession
.Link
    Invoke-GmailSession
.Link
    Get-GmailSession
.Link
    Clear-GmailSession
#>
}

function Invoke-GmailSession {
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [ScriptBlock]$ScriptBlock,

        [Parameter(Position = 0, Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential = $($cr = (Get-StoredCredential Gmail.ps:default); if ($cr -eq $null) {Get-Credential} else {$cr})
    )

    $gmail = New-GmailSession -Credential $Credential
    & $ScriptBlock $gmail
    $gmail | Remove-GmailSession

<#
.Synopsis
    Invokes a block of code on a Gmail session.
.Description
    Creates new Gmail session and passes it to a script block. Once the block is executed, the session is automatically closed.
.Parameter ScriptBlock
    Script that is executed once a session is opened.
.Parameter Credential
    The credentials that will be used to connect to Gmail.
.Link
    New-GmailSession
.Link
    Remove-GmailSession
.Example
    PS> Invoke-GmailSession -ScriptBlock {
    PS>     $args | Count-Message
    PS> }

    Description
    -----------
    Creates a Gmail session, returns the number of messages in the Inbox and then closes the session.
    The automatically created session can be accessed inside the script block via the $args variable.
.Example
    PS> Invoke-GmailSession -ScriptBlock {
    PS>     param($gmail)
    PS>     $gmail | Get-Label
    PS> }

    Description
    -----------
    Creates a Gmail session, returns all the labels used in that account and then closes the session.
    The automatically created session can be accessed inside the script block via the $gmail variable.
#>
}

function Get-GmailSession {
    $GmailSessions

<#
.Synopsis
    Returns a list of all opened Gmail sessions.
.Description
    Returns a list of all opened Gmail sessions.
.Link
    New-GmailSession
.Link
    Clear-GmailSession
#>
}

function Clear-GmailSession {
    $GmailSessions | ForEach-Object -Process { $_ | Remove-GmailSession }
    $GmailSessions = @();

<#
.Synopsis
    Closes all opened Gmail sessions.
.Description
    Closes all opened Gmail sessions.
.Link
    New-GmailSession
.Link
    Get-GmailSession
#>
}

function Get-Mailbox {
    [CmdletBinding(DefaultParameterSetName = "DefaultFolder")]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,

        [Parameter(Position = 0, ValueFromPipelineByPropertyName = $true, ParameterSetName = "DefaultFolder")]
        [ValidateSet("All Mail", "Starred", "Drafts", "Important", "Sent Mail", "Spam", "Inbox")]
        [string]$Name = "Inbox",

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = "LabelFolder")]
        [string]$Label
    )

    if ($Label) {
        $mailbox = $Session.SelectMailbox($Label)
    } elseif ($Name -and ($Name -ne "Inbox")) {
        $mailbox = $Session.SelectMailbox("[Gmail]/" + $Name)
    } elseif ($Name -and ($Name -eq "Inbox")) {
        $mailbox = $Session.SelectMailbox("Inbox")
    }

    AddSessionTo $mailbox $Session

<#
.Synopsis
    Returns a mailbox.
.Description
    Returns the Inbox if no parameters are specified, an existing Label or one of the default 
    Gmail folders (All Mail, Starred, Drafts, Important, Sent Mail, Spam)
.Parameter Session
    The opened session that will be manipulated.
.Parameter Name
    The name of the default Gmail folder to be accessed.
.Parameter Label
    The name of an existing label to be accessed.
.Link
    Get-Message
.Link
    Measure-Message
#>
}

function Get-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,

        [switch]$Prefetch,
        [switch]$Unread,
        [switch]$Read,
        [switch]$Answered,
        [switch]$Draft,
        [switch]$Undraft,
        [switch]$Starred,
        [switch]$Unstarred,
        [switch]$HasAttachment,
        [DateTime]$On,
        [DateTime]$After,
        [DateTime]$Before,
        [string]$From,
        [string]$To,
        [string]$Cc,
        [string]$Bcc,
        [string]$Text,
        [string]$Body,
        [string]$Subject,
        [string]$Label,
        [string]$FileName,

        [ValidateSet("Primary", "Personal", "Social", "Promotions", "Updates", "Forums")]
        [string]$Category
    )

    $imap = @()
    $xgm = @()

    if ($Unread) {
        $imap += "UNSEEN"
    } elseif ($Read) {
        $imap += "SEEN"
    }

    if ($Answered) {
        $imap += "ANSWERED"
    }

    if ($Draft) {
        $imap += "DRAFT"
    } elseif ($Undraft) {
        $imap += "UNDRAFT"
    }

    if ($Starred) {
        $imap += "FLAGGED"
    } elseif ($Unstarred) {
        $imap += "UNFLAGGED"
    }

    if ($On) {
        $imap += 'ON "' + $(GetRFC2060Date $After) + '"'
    }

    if ($From) {
        $imap += 'FROM "' + $From + '"'
    }

    if ($To) {
        $imap += 'TO "' + $To + '"'
    }

    if ($After) {
        $imap += 'AFTER "' + $(GetRFC2060Date $After) + '"'
    }

    if ($Before) {
        $imap += 'BEFORE "' + $(GetRFC2060Date $Before) + '"'
    }

    if ($Cc) {
        $imap += 'CC "' + $Cc + '"'
    }

    if ($Bcc) {
        $imap += 'BCC "' + $Bcc + '"'
    }

    if ($Text) {
        $imap += 'TEXT "' + $Text + '"'
    }

    if ($Body) {
        $imap += 'BODY "' + $Body + '"'
    }

    if ($Subject) {
        $imap += 'SUBJECT "' + $Subject + '"'
    }
    
    if ($Label) {
        $xgm += 'label:' + $Label
    }

    if ($HasAttachment) {
        $xgm += 'has:attachment'
    }

    if ($FileName) {
        $xgm += 'filename:' + $FileName
    }

    if ($Category) {
        $xgm += 'category:' + $Category
    }

    if ($imap.Length -gt 0) {
        $criteria = ($imap -join ') (')
    }

    if ($xgm.Length -gt 0) {
        $gmcr = 'X-GM-RAW "' + ($xgm -join ' ') + '"'
        if ($imap.Length -gt 0) {
            $criteria = $criteria + ' (' + $gmcr + ')'
        } else {
            $criteria = $gmcr
        }
    }

    $result = $Session.Search('(' + $criteria + ')');
    $i = 1
    foreach ($item in $result)
    {
        $msg = $Session.GetMessage($item, !$Prefetch, $false)
        AddSessionTo $msg $Session
        Write-Progress -Activity "Gathering messages" -Status "Progress: $($i)/$($result.Count)" -PercentComplete ($i / $result.Count * 100) -Id 90017
        $i += 1
    }

<#
.Synopsis
    Returns a list of messages.
.Description
    Returns a (filtered) list of the messages inside a selected mailbox (using Get-Mailbox).
.Parameter Session
    The opened session that will be manipulated.
.Parameter Prefetch
    If specified, fetches the message's body and attachments; otherwise only the headers are downloaded from the server
.Parameter Unread
    Forces only unread messages to be returned
.Parameter Read
    Forces only read messages to be returned
.Parameter Answered
    Forces only messages that has been answered to, to be returned
.Parameter Draft
    If set, only drafts will be returned
.Parameter Undraft
    If set, only non-draft messages will be returned
.Parameter Starred
    Indicates only starred mesages to be returned
.Parameter Unstarred
    Indicates only mesages that are not marked with Star to be returned
.Parameter On
    Filters the messages based on an exact date of receiving 
.Parameter After
    Returns only messages received after a given date
.Parameter Before
    Returns only messages received before a given date
.Parameter From
    Filters the messages based on the sender's name and email address
.Parameter To
    Filters the messages based on the recipient's name and email address
.Parameter Cc
    Filters the messages based on the Cc recipient's name and email address
.Parameter Bcc
    Filters the messages based on the Bcc recipient's name and email address
.Parameter Text
    A text to search the entire message for
.Parameter Body
    A substring to search the message's body for
.Parameter Subject
    A substring to search the message's subject for
.Parameter Label
    Returns only messages having a particular label applied
.Parameter HasAttachment
    Returns only messages with attachments
.Parameter FileName
    Returns only messages having attachments with a given name
.Parameter Category
    Returns only messages within a particular category
.Link
    Get-Message
.Link
    Update-Message
.Link
    Remove-Message
.Link
    Measure-Message
.Link
    Reveive-Message
#>
}

function Remove-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message
    )
    
    process {
        $Session.DeleteMessage($Message)
    }

<#
.Synopsis
    Deletes a message.
.Description
    Sends a message to the Gmail's Trash folder.
.Parameter Session
    The opened session that will be manipulated.
.Parameter Message
    The message that will be deleted.
.Link
    Get-Message
.Link
    Update-Message
#>
}

function Update-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message,

        [Parameter(ParameterSetName = "Seen")]
        [switch]$Read,

        [Parameter(ParameterSetName = "Unseen")]
        [switch]$Unread,

        [Parameter(ParameterSetName = "Unseen")]
        [Parameter(ParameterSetName = "Seen")]
        [Parameter(ParameterSetName = "Flagged")]
        [Parameter(ParameterSetName = "Unflagged")]
        [switch]$Archive,

        [Parameter(ParameterSetName = "Flagged")]
        [switch]$Star,

        [Parameter(ParameterSetName = "Unflagged")]
        [switch]$Unstar,
        
        [Parameter(ParameterSetName = "Unseen")]
        [Parameter(ParameterSetName = "Seen")]
        [Parameter(ParameterSetName = "Flagged")]
        [Parameter(ParameterSetName = "Unflagged")]
        [switch]$Spam
    )
    
    process {
        if ($Archive) {
            $Session.MoveMessage($Message.Uid, "[Gmail]/All Mail")
        }

        if ($Spam) {
            $Session.MoveMessage($Message.Uid, "[Gmail]/Spam")
        }

        $replace = $false
        $changed = $false

        if ($Read) {
            $flags = $flags -bor [AE.Net.Mail.Flags]::Seen
            $changed = $true
        } elseif ($Unread) {
            $flags = $Message.Flags
            $flags = $flags -bxor [AE.Net.Mail.Flags]::Seen
            $changed = $true
            $replace = $true
        } 

        if ($Star) {
            $flags = $flags -bor [AE.Net.Mail.Flags]::Flagged
            $changed = $true
        } elseif ($Unstar) {
            $flags = $Message.Flags
            $flags = $flags -bxor [AE.Net.Mail.Flags]::Flagged
            $changed = $true
            $replace = $true
        }

        if ($changed) {
            if (-not $replace) {
                $Session.AddFlags([AE.Net.Mail.Flags]$flags, @($Message))
            } else {
                $Session.SetFlags([AE.Net.Mail.Flags]$flags, @($Message))
            }
        }
    }

<#
.Synopsis
    Flags a message.
.Description
    Archives, marks as spam, as read/undead or adds/removes a star from a given message.
.Parameter Session
    The opened session that will be manipulated.
.Parameter Message
    The message that will be updated.
.Parameter Read
    Marks a message as read
.Parameter Unread
    Marks a message as undead
.Parameter Star
    Flags a message with a Star
.Parameter Unstar
    Removes the star from a message
.Parameter Archive
    Archives a message
.Parameter Spam
    Forces a message to be marked as spam
.Link
    Get-Message
.Link
    Remove-Message
#>
}

function Receive-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message
    )

    process {
        $Session.GetMessage($Message.Uid, $false)
    }

<#
.Synopsis
    Fetches a message.
.Description
    Fetches the whole message from the server (including the body and the attachments).
.Parameter Session
    The opened session that will be manipulated.
.Parameter Message
    The message that will be fetched.
.Link
    Get-Message
#>
}

function Move-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message,

        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true, ParameterSetName = "DefaultFolder")]
        [ValidateSet("Inbox", "All Mail", "Starred", "Drafts", "Important", "Sent Mail", "Spam")]
        [string]$Mailbox,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "LabelFolder")]
        [string]$Label
    )

    if ($Label) {
        $Session.MoveMessage($Message.Uid, $Label)
    } elseif ($Mailbox ) {
        $Session.MoveMessage($Message.Uid, $Mailbox)
    }

<#
.Synopsis
    Moves a message.
.Description
    Moves a message to a different mailbox or label
.Parameter Session
    The opened session that will be manipulated.
.Parameter Message
    The message that will be fetched.
.Parameter Mailbox
    The name of a mailbox the message will be moved to.
.Parameter Label
    The name of a label the message will be moved to.
.Link
    Get-Message
#>
}

function Measure-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    $Session.GetMessageCount()

<#
.Synopsis
    Counts messages.
.Description
    Returns the number of messages in a mailbox (supports labels too).
.Parameter Session
    The opened session that will be manipulated.
.Link
    Get-Message
#>
}

function Save-Attachment {
    [CmdletBinding(DefaultParameterSetName = "Path")]
    param (
        [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message,

        [Parameter(Position = 0, ParameterSetName = "Path", Mandatory = $true)]
        [string[]] $Path,

        [Parameter(ParameterSetName = "LiteralPath", Mandatory = $true)]
        [string[]] $LiteralPath
    )

    process {
        foreach ($a in $Message.Attachments)
        {
            $p = Convert-Path ($Path + $LiteralPath)
            $loc = Join-Path $p $Message.Uid
            $a.Save((Join-Path $loc $a.Filename))
        }
    }

<#
.Synopsis
    Downloads the attachments of a message.
.Description
    Downloads the attachments of a message to a local forlder.
.Parameter Message
    The message whose attachments will be downloaded
.Parameter Path
    Specifies a path to the directory where the attachments will be saved
.Parameter LiteralPath
    Specifies a path to the directory where the attachments will be saved. The value of the LiteralPath parameter is used 
    exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters, enclose it in 
    single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any characters as escape sequences.
.Link
    Get-Message
#>
}

function Get-Label {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,

        [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message,
        
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias("Name")]
        [string]$Like = "",
        
        [Parameter()]
        [switch]$All
    )

    process {
        if ($Message) {
            $Message.Labels
        } else {
            if ($All) {
                $Session.ListMailboxes($Like, "*")
            } else {
                $Session.ListMailboxes($Like, "*") | Where-Object { $_.Name -notmatch "\[Gmail\]" -and $_.Name -ne "INBOX" }
            }
        }
    }

<#
.Synopsis
    Returns the labels applied to a message or all labels that exist.
.Description
    Returns the labels applied to a message or all labels that exist.
.Parameter Session
    The opened session that will be used to fetch all existing labels.
.Parameter Message
    The message whose labels will be returned
.Link
    New-Label
.Link
    Set-Label
.Link
    Remove-Label
#>
}

function New-Label {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string[]]$Name,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    foreach ($item in $Name)
    {
        $Session.CreateMailbox($item)
    }

<#
.Synopsis
    Creates a label.
.Description
    Creates a new label.
.Parameter Session
    The opened session that will be manipulated.
.Parameter Name
    The of the label that will be created.
.Link
    Get-Label
.Link
    Set-Label
.Link
    Remove-Label
#>
}

function Remove-Label {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string[]]$Name,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session, 

        [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message
    )

    foreach ($item in $Name)
    {
        if ($Message) {
            $Session.RemoveLabels($Name, @($Message))
        } else {
            $Session.DeleteMailbox($item)
        }
    }

<#
.Synopsis
    Removes a label from a message or deletes the label from the account.
.Description
    Removes a label from a message or deletes the label from the account.
.Parameter Session
    The opened session that will be manipulated.
.Parameter Name
    The of the label that will be removed.
.Parameter Message
    The message from which the label will be removed; if not specified the label will be deleted from the account
.Link
    Get-Label
.Link
    Set-Label
.Link
    New-Label
#>
}

function Set-Label {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message,
        
        [Parameter(Position = 0, Mandatory = $true)]
        [string[]]$Name,

        [Parameter()]
        [switch]$Force
    )

    process {
        $labels = $Session | Get-Label | ForEach-Object { $_.Name };
        
        foreach ($label in $Name)
        {
            if (!$labels.Contains($label)) {
                if ($Force) {
                    $Session | New-Label $label | Out-Null
                } else {
                    Write-Error "The label '$label' doesn't exist! Use the -Force parameter to create and apply it"
                    $er = $true
                }
            }
        }

        if (!$er) {
            $Session.AddLabels($Name, @($Message))
        }
    }

<#
.Synopsis
   Adds a label to a message.
.Description
   Adds a label to a message.
.Parameter Session
    The opened session that will be manipulated.
.Parameter Message
    The message to which the label will be applied.
.Parameter Name
    The name of the label that will be apllied.
.Parameter Force
    Forces the creation of the label if it doesn't exist. An error will be thrown if the 
    label doesn't exist and the command is executed without the -Force parameter
.Link
    Get-Label
.Link
    New-Label
.Link
    Remove-Label
#>
}

if (Test-Path Function:\TabExpansion) {
    Rename-Item Function:\TabExpansion TabExpansionBackup
}

# Revert the old tabexpnasion when Gmail.ps is unloaded
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
    Write-Verbose "Revert tab expansion back"
    Remove-Item Function:\TabExpansion -ErrorAction SilentlyContinue
    if (Test-Path Function:\TabExpansionBackup) {
        Rename-Item Function:\TabExpansionBackup Function:\TabExpansion
    }
}

function global:TabExpansion($line, $lastWord) {
    $lastBlock = ($line -split ';')[-1].TrimStart()
    $matched = $lastBlock -match "^\`$(?<svar>(?:\w|_)+)\s*\|\s*(?<rest>.*)$"
    $svar = "Variable:\$($Matches['svar'])"

    if ($matched -and $Matches['svar'] -and (Test-Path $svar) -and ((Get-Item $svar).Value.ToString() -eq "AE.Net.Mail.ImapClient")) {
        switch -regex ($Matches['rest']) {
            # Execute Gmail.ps tab completion for all related commands
            "^$(Get-LabelCmdPattern)(.*)-(Name|Label)\s?(.*)$" { Get-LabelsForSession $svar $lastWord }
            "^($(Get-AliasPattern Remove-Label)|$(Get-AliasPattern Set-Label))(.*)$" { Get-LabelsForSession $svar $lastWord }

            # Fall back on existing tab expansion
            default { DefaultTabExpansion $line $lastWord }
        }
    } else {
        DefaultTabExpansion $line $lastWord
    }
}

function DefaultTabExpansion($line, $lastWord) {
    if (Test-Path Function:\TabExpansionBackup) { TabExpansionBackup $line $lastWord }
}

function Get-LabelCmdPattern {
    $cmdlets = @("Get-Message", "Move-Message", "Remove-Label", "Set-Label")
    $pattern = @()
    
    foreach ($cmd in $cmdlets) {
        $pattern += Get-AliasPattern
    }

    "($($pattern -join '|'))"
}

function Get-AliasPattern($cmd) {
    @(Get-Alias -Definition $cmd | Select-Object -ExpandProperty Name) -join '|'
}

function Get-LabelsForSession($session, $filter) {
    (Get-Item $session).Value | Get-Label -Like $filter | foreach { $_.Name }
}

function GetRFC2060Date([DateTime]$date) {
    $date.ToString("dd-MMM-yyyy hh:mm:ss zz", [CultureInfo]::GetCultureInfo("en-US"))
}

function AddSessionTo($item, [AE.Net.Mail.ImapClient]$session) {
    $item | Add-Member -MemberType NoteProperty -Name Session -Value $session -PassThru
}

function Get-StoredCredential
{
    [CmdletBinding()]
    [OutputType([PSCredential])]
    Param
    (
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias("Address", "Location", "TargetName")]
        [string]$Name
    )

    End
    {
        $nCredPtr= New-Object IntPtr

        $success = [ADVAPI32.Util]::CredRead($Name,1,0,[ref] $nCredPtr)

        if ($success) {
            $critCred = New-Object ADVAPI32.Util+CriticalCredentialHandle $nCredPtr
            $cred = $critCred.GetCredential()
            $username = $cred.UserName
            $securePassword = $cred.CredentialBlob | ConvertTo-SecureString -AsPlainText -Force
            $cred = $null
            Write-Output (New-Object System.Management.Automation.PSCredential $username, $securePassword)
        } else {
            Write-Verbose "No credentials where found in Windows Credential Manager for TargetName: $Name"
        }
    }

    Begin
    {
        $sig = @"

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct NativeCredential
{
    public UInt32 Flags;
    public CRED_TYPE Type;
    public IntPtr TargetName;
    public IntPtr Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public UInt32 CredentialBlobSize;
    public IntPtr CredentialBlob;
    public UInt32 Persist;
    public UInt32 AttributeCount;
    public IntPtr Attributes;
    public IntPtr TargetAlias;
    public IntPtr UserName;

    internal static NativeCredential GetNativeCredential(Credential cred)
    {
        NativeCredential ncred = new NativeCredential();
        ncred.AttributeCount = 0;
        ncred.Attributes = IntPtr.Zero;
        ncred.Comment = IntPtr.Zero;
        ncred.TargetAlias = IntPtr.Zero;
        ncred.Type = CRED_TYPE.GENERIC;
        ncred.Persist = (UInt32)1;
        ncred.CredentialBlobSize = (UInt32)cred.CredentialBlobSize;
        ncred.TargetName = Marshal.StringToCoTaskMemUni(cred.TargetName);
        ncred.CredentialBlob = Marshal.StringToCoTaskMemUni(cred.CredentialBlob);
        ncred.UserName = Marshal.StringToCoTaskMemUni(System.Environment.UserName);
        return ncred;
    }
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct Credential
{
    public UInt32 Flags;
    public CRED_TYPE Type;
    public string TargetName;
    public string Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public UInt32 CredentialBlobSize;
    public string CredentialBlob;
    public UInt32 Persist;
    public UInt32 AttributeCount;
    public IntPtr Attributes;
    public string TargetAlias;
    public string UserName;
}

public enum CRED_TYPE : uint
    {
        GENERIC = 1,
        DOMAIN_PASSWORD = 2,
        DOMAIN_CERTIFICATE = 3,
        DOMAIN_VISIBLE_PASSWORD = 4,
        GENERIC_CERTIFICATE = 5,
        DOMAIN_EXTENDED = 6,
        MAXIMUM = 7,      // Maximum supported cred type
        MAXIMUM_EX = (MAXIMUM + 1000),  // Allow new applications to run on old OSes
    }

public class CriticalCredentialHandle : Microsoft.Win32.SafeHandles.CriticalHandleZeroOrMinusOneIsInvalid
{
    public CriticalCredentialHandle(IntPtr preexistingHandle)
    {
        SetHandle(preexistingHandle);
    }

    public Credential GetCredential()
    {
        if (!IsInvalid)
        {
            NativeCredential ncred = (NativeCredential)Marshal.PtrToStructure(handle,
                  typeof(NativeCredential));
            Credential cred = new Credential();
            cred.CredentialBlobSize = ncred.CredentialBlobSize;
            cred.CredentialBlob = Marshal.PtrToStringUni(ncred.CredentialBlob,
                  (int)ncred.CredentialBlobSize / 2);
            cred.UserName = Marshal.PtrToStringUni(ncred.UserName);
            cred.TargetName = Marshal.PtrToStringUni(ncred.TargetName);
            cred.TargetAlias = Marshal.PtrToStringUni(ncred.TargetAlias);
            cred.Type = ncred.Type;
            cred.Flags = ncred.Flags;
            cred.Persist = ncred.Persist;
            return cred;
        }
        else
        {
            throw new InvalidOperationException("Invalid CriticalHandle!");
        }
    }

    override protected bool ReleaseHandle()
    {
        if (!IsInvalid)
        {
            CredFree(handle);
            SetHandleAsInvalid();
            return true;
        }
        return false;
    }
}

[DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool CredRead(string target, CRED_TYPE type, int reservedFlag, out IntPtr CredentialPtr);

[DllImport("Advapi32.dll", EntryPoint = "CredFree", SetLastError = true)]
public static extern bool CredFree([In] IntPtr cred);


"@
        try
        {
            Add-Type -MemberDefinition $sig -Namespace "ADVAPI32" -Name 'Util' -ErrorAction Stop
        }
        catch
        {
            Write-Error -Message "Could not load custom type. $($_.Exception.Message)"
        }
    }
}

New-Alias -Name Select-Mailbox -Value Get-Mailbox
New-Alias -Name Filter-Message -Value Get-Message
New-Alias -Name Count-Message -Value Measure-Message
New-Alias -Name Add-Label -Value Set-Label

Export-ModuleMember -Alias * -Function New-GmailSession, Remove-GmailSession, Invoke-GmailSession, 
                                       Get-GmailSession, Clear-GmailSession, Get-Mailbox, Get-Message, 
                                       Measure-Message, Remove-Message, Update-Message, Move-Message, 
                                       Get-Label, New-Label, Remove-Label, Set-Label, Receive-Message, 
                                       Save-Attachment
