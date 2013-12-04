

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")

if ($PSVersionTable["PSVersion"].Major -eq 2 -and !$isAdmin) {
    Write-Host "To install Gmail.ps on PowerShell 2.0, pease run this script as Administrator" -Foreground Red
    return
}

$ErrorCount = $Error.Count
$ModulePaths = @($Env:PSModulePath -split ';')

$ExpectedUserModulePath = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath WindowsPowerShell\Modules
$Destination = $ModulePaths | Where-Object { $_ -eq $ExpectedUserModulePath}

if (-not $Destination) {
    $Destination = $ModulePaths | Select-Object -Index 0
}

if (-not (Test-Path $Destination)) {
    New-Item $Destination -ItemType Directory -Force | Out-Null
} elseif (Test-Path (Join-Path $Destination "Gmail.ps\Gmail.ps.psm1")) {
    Write-Host "Gmail.ps is already installed" -Foreground Green
    return
}

try {
    git --version | Out-Null
    $hasGit = $true
} catch {
    $hasGit = $false
    $ErrorCount -= 1
}

$CurrentLocation = Get-Location
Push-Location $Destination

if (Test-Path (Join-Path $CurrentLocation "Gmail.ps.psm1")) {
    Copy-Item -Path $CurrentLocation -Destination ($Destination + "\Gmail.ps\") -Recurse
} elseif ($hasGit) {
    git clone https://github.com/nikoblag/Gmail.ps.git
} else {
    New-Item ($Destination + "\Gmail.ps\") -ItemType Directory -Force | Out-Null
    Write-Host "Downloading Gmail.ps from https://github.com/nikoblag/Gmail.ps"
    $rawMasterURL = "https://github.com/nikoblag/Gmail.ps/raw/master/"
    $files = @("Gmail.ps.psm1","Gmail.ps.psd1","AE.Net.Mail.dll","LICENSE","README.md","FormatData\MailMessage.Format.ps1xml","FormatData\Mailbox.Format.ps1xml","FormatData\Attachment.Format.ps1xml")
    $client = (New-Object Net.WebClient)
    $client.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

    foreach ($file in $files) {
        $client.DownloadFile($rawMasterURL + $file, $Destination + "\Gmail.ps\" + $file)
    }
}

Pop-Location

$executionPolicy  = (Get-ExecutionPolicy)
$executionRestricted = ($executionPolicy -eq "Restricted")

if ($executionRestricted) {
    Write-Warning @"
Your execution policy is $executionPolicy, this means you will not be able import or use any scripts including modules.
To fix this, change your execution policy to something like RemoteSigned.

        PS> Set-ExecutionPolicy RemoteSigned

For more information execute:
        
        PS> Get-Help about_execution_policies

"@
}

if (!$executionRestricted) {
    if ($PSVersionTable["PSVersion"].Major -gt 2) {
        Import-Module -Name $Destination\Gmail.ps
    } else {
        # TODO: Do this the correct way:
        # Using COMPLUS_ApplicationMigrationRuntimeActivationConfigPath env variable
        $config_text = @"
<?xml version="1.0"?>
<configuration>
    <startup useLegacyV2RuntimeActivationPolicy="true">
        <supportedRuntime version="v4.0"/>
        <supportedRuntime version="v2.0"/>
    </startup>
</configuration>
"@
        if (Test-Path $pshome\powershell.exe.config) {
            Copy-Item $pshome\powershell.exe.config $pshome\powershell.exe.config.old -Force
            Write-Host "Backup created at $pshome\powershell.exe.config.old"
        }

        if (Test-Path $pshome\powershell_ise.exe.config) {
            Copy-Item $pshome\powershell_ise.exe.config $pshome\powershell_ise.exe.config.old -Force
            Write-Host "Backup created at $pshome\powershell_ise.exe.config.old"
        }

        $config_text | Out-File $pshome\powershell.exe.config
        $config_text | Out-File $pshome\powershell_ise.exe.config
        Write-Host "Please restart PowerShell!"
    }
}

if ($ErrorCount -eq $Error.Count) {
    Write-Host "Gmail.ps is installed and ready to use" -Foreground Green
} else {
    Write-Host "Something went wrong. Gmail.ps may not work." -Foreground Red
}
