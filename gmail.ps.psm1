
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

function Get-Label {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session,
        
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
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
        [string]$Name,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    process {
        $Session.CreateMailbox($Name)
    }
}

function Remove-Label {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AE.Net.Mail.ImapClient]$Session
    )

    process {
        $Session.DeleteMailbox($Name)
    }
}

Export-ModuleMember -Function * -Alias * -Cmdlet *