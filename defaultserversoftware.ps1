param (
    [Parameter(Mandatory = $true)]
    [string]$VOUCHER,

    [Parameter(Mandatory = $true)]
    [string]$CUSTOMTOKEN
)

# ============================================================
# Zorg dat C:\Temp bestaat
# ============================================================
$tempPath = "C:\Temp"
if (-not (Test-Path $tempPath)) {
    New-Item -ItemType Directory -Path $tempPath -Force | Out-Null
}

# ============================================================
# Logging setup
# ============================================================
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$transcriptPath = "$tempPath\setup_$timestamp.transcript.log"
Start-Transcript -Path $transcriptPath

function Write-Status {
    param([string]$status)
    $statusFile = "C:\WindowsAzure\Logs\Plugins\Microsoft.Compute.CustomScriptExtension\status.txt"
    $status | Out-File -FilePath $statusFile -Encoding utf8
}

function Test-AppInstalled {
    param([string]$NamePattern)

    $uninstallPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($path in $uninstallPaths) {
        $match = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like $NamePattern } |
            Select-Object -First 1

        if ($null -ne $match) {
            return $true
        }
    }

    return $false
}

Write-Host "[$(Get-Date)] Script gestart"

# ============================================================
# Software download locaties
# ============================================================
$withSecureUrl = "https://raw.githubusercontent.com/ahmetsagirtauw/Server2025software/main/ElementsAgentOfflineInstaller.msi"
$rapid7Url     = "https://raw.githubusercontent.com/ahmetsagirtauw/Server2025software/main/agentInstaller-x86_64.msi"

$withSecureDest = "$tempPath\ElementsAgentOfflineInstaller.msi"
$rapid7Dest     = "$tempPath\agentInstaller-x86_64.msi"
$withSecureMsiLog = "$tempPath\withsecure_$timestamp.msi.log"
$rapid7MsiLog     = "$tempPath\rapid7_$timestamp.msi.log"

Write-Host "[$(Get-Date)] Downloaden van MSI bestanden..."

$ProgressPreference = 'SilentlyContinue'
try {
    Invoke-WebRequest -Uri $withSecureUrl -OutFile $withSecureDest -UseBasicParsing
    Invoke-WebRequest -Uri $rapid7Url -OutFile $rapid7Dest -UseBasicParsing

    if ((-not (Test-Path $withSecureDest)) -or ((Get-Item $withSecureDest).Length -eq 0)) {
        throw "WithSecure MSI is niet gedownload of is leeg."
    }

    if ((-not (Test-Path $rapid7Dest)) -or ((Get-Item $rapid7Dest).Length -eq 0)) {
        throw "Rapid7 MSI is niet gedownload of is leeg."
    }
}
catch {
    Write-Host "Download mislukt: $_"
    Write-Status "FAILED: Download error"
    Stop-Transcript
    exit 3
}

# ============================================================
# Installatie WithSecure
# ============================================================
Write-Host "[$(Get-Date)] Installatie WithSecure agent gestart..."

if (Test-AppInstalled -NamePattern '*WithSecure*') {
    Write-Host "WithSecure lijkt al geinstalleerd. Installatie wordt overgeslagen."
}
else {
    $withSecure = Start-Process msiexec.exe -ArgumentList "/i `"$withSecureDest`" VOUCHER=$VOUCHER /quiet /norestart /L*v `"$withSecureMsiLog`"" -Wait -PassThru

    if ($withSecure.ExitCode -notin @(0, 3010, 1641)) {
        Write-Host "WithSecure installatie mislukt. Exitcode: $($withSecure.ExitCode). MSI log: $withSecureMsiLog"
        Write-Status "FAILED: WithSecure install error"
        Stop-Transcript
        exit 21
    }

    Write-Host "WithSecure installatie geslaagd. Exitcode: $($withSecure.ExitCode). MSI log: $withSecureMsiLog"
}

# ============================================================
# Installatie Rapid7
# ============================================================
Write-Host "[$(Get-Date)] Installatie Rapid7 agent gestart..."

if (Test-AppInstalled -NamePattern '*Rapid7*') {
    Write-Host "Rapid7 lijkt al geinstalleerd. Installatie wordt overgeslagen."
}
else {
    $rapid7 = Start-Process msiexec.exe -ArgumentList "/i `"$rapid7Dest`" CUSTOMTOKEN=$CUSTOMTOKEN /quiet /norestart /L*v `"$rapid7MsiLog`"" -Wait -PassThru

    if ($rapid7.ExitCode -notin @(0, 3010, 1641)) {
        Write-Host "Rapid7 installatie mislukt. Exitcode: $($rapid7.ExitCode). MSI log: $rapid7MsiLog"
        Write-Status "FAILED: Rapid7 install error"
        Stop-Transcript
        exit 22
    }

    Write-Host "Rapid7 installatie geslaagd. Exitcode: $($rapid7.ExitCode). MSI log: $rapid7MsiLog"
}

# ============================================================
# Afronding
# ============================================================
Write-Host "[$(Get-Date)] Alle software succesvol geïnstalleerd en reboot wordt ingepland..."

# Maak een geplande taak die over 30 seconden reboot
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command `"Restart-Computer -Force`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(30)
Register-ScheduledTask -TaskName "PostCSEReboot" -Action $action -Trigger $trigger -RunLevel Highest -Force

Write-Status "SUCCESS"

Stop-Transcript

exit 0
