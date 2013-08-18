
function New-GmailSession {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential = (Get-Credential)
    )

    $ghost = "imap.gmail.com"
    $gport = 993
    $guser = "test.dummy.nb@gmail.com" # $Credential.UserName
    $gpass = "dummyaccount" #TODO: Parse it from the $Credential

    New-Object -TypeName AE.Net.Mail.ImapClient -ArgumentList $ghost,$guser,$gpass,Login,$gport,$true,$false
}

function Remove-GmailSession {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    $Session.Disconnect()
}

function Get-Inbox {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    Get-Mailbox -Session $Session -Name "Inbox"
}

function Get-Mailbox {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,

        [Parameter(Position = 1, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias("Name")]
        [string]$Label = ""
    )

    $mailbox = $Session.SelectMailbox($Label)
    AddSessionTo $mailbox $Session
}

function Filter-Message {
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
        [string]$Query,
        [int]$Last
    )

    $ar = @()

    if ($Unread) {
        $ar += "UNSEEN"
    } elseif ($Read) {
        $ar += "SEEN"
    }

    if ($Answered) {
        $ar += "ANSWERED"
    }

    if ($Draft) {
        $ar += "DRAFT"
    } elseif ($Undraft) {
        $ar += "UNDRAFT"
    }

    if ($Starred) {
        $ar += "FLAGGED"
    } elseif ($Unstarred) {
        $ar += "UNFLAGGED"
    }

    if ($On) {
        $ar += 'ON "' + $(GetRFC2060Date $After) + '"'
    }

    if ($From) {
        $ar += 'FROM "' + $From + '"'
    }

    if ($To) {
        $ar += 'TO "' + $To + '"'
    }

    if ($After) {
        $ar += 'AFTER "' + $(GetRFC2060Date $After) + '"'
    }

    if ($Before) {
        $ar += 'BEFORE "' + $(GetRFC2060Date $Before) + '"'
    }

    if ($Cc) {
        $ar += 'CC "' + $Cc + '"'
    }

    if ($Bcc) {
        $ar += 'BCC "' + $Bcc + '"'
    }

    if ($Text) {
        $ar += 'TEXT "' + $Text + '"'
    }

    if ($Body) {
        $ar += 'BODY "' + $Body + '"'
    }

    if ($Label) {
        $ar += 'LABEL "' + $Label + '"'
    }

    if ($Query) {
        $ar += 'QUERY "' + $Query + '"'
    }

    if ($Subject) {
        $ar += 'SUBJECT "' + $Subject + '"'
    }

    if ($Last) {
        Write-Warning "The -Last parameter is not yet implemented; will be ignored"
    }

    $criteria = '(' + ($ar -join ') (') + ')'
    $result = $Session.Search($criteria);

    foreach ($item in $result)
    {
        $msg = $Session.GetMessage($item, !$Prefetch, $false)
        AddSessionTo $msg $Session
    }

}

function GetRFC2060Date([DateTime]$date) {
    $date.ToString("dd-MMM-yyyy hh:mm:ss zz", [CultureInfo]::GetCultureInfo("en-US"))
}

function AddSessionTo($item, [AE.Net.Mail.ImapClient]$session) {
    $item | Add-Member -MemberType NoteProperty -Name Session -Value $session -PassThru
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
}

function Update-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message,

        [Parameter()]
        [switch]$Read,
        [Parameter()]
        [switch]$Unread,
        [Parameter()]
        [switch]$Archive,
        [Parameter()]
        [switch]$Star,
        [Parameter()]
        [switch]$Unstar,
        [Parameter()]
        [switch]$Spam,
        [Parameter(ParameterSetName = "LabelSet")]
        [string[]]$Label,
        [Parameter(ParameterSetName = "LabelSet")]
        [switch]$Force
    )
    
    process {
        if ($Unread -or $Unstar) {
            Write-Warning "The -Unread and-Unstar parameters are not yet implemented; will be ignored"
        }

        if ($Archive) {
            $Session.MoveMessage($Message.Uid, "[Gmail]/All Mail")
        }

        if ($Spam) {
            $Session.MoveMessage($Message.Uid, "[Gmail]/Spam")
        }

        $flags = [AE.Net.Mail.Flags]::None

        if ($Read) {
            $flags -bor [AE.Net.Mail.Flags]::Seen
        }

        if ($Star) {
            $flags -bor [AE.Net.Mail.Flags]::Flagged
        }

        if ($flags -ne [AE.Net.Mail.Flags]::None) {
            $Session.AddFlags($flags, @($Message))
        }
    }
}

function Move-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [AE.Net.Mail.MailMessage]$Message,
        [Parameter(Mandatory = $true)]
        [string]$Mailbox
    )
    
    Write-Warning "This cmdlet may lead to unexpected results"
    $res = Read-Host -Prompt "Type 'y' to continue"
    if ($res -eq "y") {
        $Session.MoveMessage($Message.Uid, $Mailbox)
    }
}

function Count-Message {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    $Session.GetMessageCount()
}

function Get-Label {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,
        
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias("Name")]
        [string]$Like = "",
        
        [Parameter()]
        [switch]$All
    )

    if ($All) {
        $Session.ListMailboxes($Like, "*")
    } else {
        $Session.ListMailboxes($Like, "*") | Where-Object { $_.Name -notmatch "\[Gmail\]" -and $_.Name -ne "INBOX" }
    }
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
}

function Remove-Label {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string[]]$Name,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    foreach ($item in $Name)
    {
        $Session.DeleteMailbox($item)
    }
}

Export-ModuleMember -Function New-GmailSession, Remove-GmailSession, Get-Inbox, Get-Mailbox, 
                                Filter-Message, Count-Message, Remove-Message, Update-Message, 
                                Get-Label, New-Label, Remove-Label
