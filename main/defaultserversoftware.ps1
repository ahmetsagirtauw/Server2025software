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
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = "$tempPath\setup_$timestamp.log"
Start-Transcript -Path $logPath

function Write-Log {
    param([string]$message)
    $message | Add-Content -Path $logPath
}

function Write-Status {
    param([string]$status)
    $statusFile = "C:\WindowsAzure\Logs\Plugins\Microsoft.Compute.CustomScriptExtension\status.txt"
    $status | Out-File -FilePath $statusFile -Encoding utf8
}

Write-Log "[$(Get-Date)] Script gestart"

# ============================================================
# Software download locaties
# ============================================================
$withSecureUrl = "https://raw.githubusercontent.com/bramlever-tauw/Server2025software/main/ElementsAgentOfflineInstaller.msi"
$rapid7Url     = "https://raw.githubusercontent.com/bramlever-tauw/Server2025software/main/agentInstaller-x86_64.msi"

$withSecureDest = "$tempPath\ElementsAgentOfflineInstaller.msi"
$rapid7Dest     = "$tempPath\agentInstaller-x86_64.msi"

Write-Log "[$(Get-Date)] Downloaden van MSI bestanden..."

try {
    Invoke-WebRequest -Uri $withSecureUrl -OutFile $withSecureDest -UseBasicParsing
    Invoke-WebRequest -Uri $rapid7Url -OutFile $rapid7Dest -UseBasicParsing
}
catch {
    Write-Log "Download mislukt: $_"
    Write-Status "FAILED: Download error"
    Stop-Transcript
    exit 3
}

# ============================================================
# Installatie WithSecure
# ============================================================
Write-Log "[$(Get-Date)] Installatie WithSecure agent gestart..."

$withSecure = Start-Process msiexec.exe -ArgumentList "/i `"$withSecureDest`" VOUCHER=$VOUCHER /quiet /norestart" -Wait -PassThru

if ($withSecure.ExitCode -ne 0) {
    Write-Log "WithSecure installatie mislukt. Exitcode: $($withSecure.ExitCode)"
    Write-Status "FAILED: WithSecure install error"
    Stop-Transcript
    exit 2
}

# ============================================================
# Installatie Rapid7
# ============================================================
Write-Log "[$(Get-Date)] Installatie Rapid7 agent gestart..."

$rapid7 = Start-Process msiexec.exe -ArgumentList "/i `"$rapid7Dest`" CUSTOMTOKEN=$CUSTOMTOKEN /quiet /norestart" -Wait -PassThru

if ($rapid7.ExitCode -ne 0) {
    Write-Log "Rapid7 installatie mislukt. Exitcode: $($rapid7.ExitCode)"
    Write-Status "FAILED: Rapid7 install error"
    Stop-Transcript
    exit 2
}

# ============================================================
# Afronding
# ============================================================
Write-Log "[$(Get-Date)] Alle software succesvol geïnstalleerd en reboot wordt ingepland..."

# Maak een geplande taak die over 30 seconden reboot
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command `"Restart-Computer -Force`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(30)
Register-ScheduledTask -TaskName "PostCSEReboot" -Action $action -Trigger $trigger -RunLevel Highest -Force

Write-Status "SUCCESS"

Stop-Transcript

exit 0
