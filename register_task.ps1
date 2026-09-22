<#
.SYNOPSIS
  Enregistre une tache planifiee Windows pour executer bal_to_csv_rovo.ps1 periodiquement.
.DESCRIPTION
  Usage :
    .\register_task.ps1                      # Toutes les 15 minutes, utilisateur courant
    .\register_task.ps1 -IntervalMinutes 30  # Toutes les 30 minutes
    .\register_task.ps1 -Unregister          # Supprime la tache planifiee
  La tache s'execute en arriere-plan (fenetre masquee) via powershell.exe -WindowStyle Hidden.
#>
param(
    [int]$IntervalMinutes = 15,
    [string]$TaskName = "bal_to_csv_rovo",
    [switch]$Unregister
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptPath = Join-Path $ScriptDir "bal_to_csv_rovo.ps1"

if (-not (Test-Path $ScriptPath)) {
    Write-Error "Script introuvable: $ScriptPath"
    exit 1
}

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Tache planifiee '$TaskName' supprimee (si elle existait)"
    exit 0
}

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -Silent"
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "BAL Outlook vers CSV Jira" -Force | Out-Null
Write-Host "Tache planifiee '$TaskName' creee: toutes les $IntervalMinutes minutes"
Write-Host "Script execute: $ScriptPath -Silent"
