<#
.SYNOPSIS
  Script pour traiter les BAL Enedis via Outlook COM et generer un CSV importable dans Jira.
.DESCRIPTION
  Les fonctions sont dans bal_to_csv_rovo.functions.ps1.
  La configuration (BAL, projets, priorites...) est dans config.json (valeurs par defaut si absent).
.PARAMETER DryRun
  N'ecrit ni CSV, ni categorie Outlook, ni trace d'import.
.PARAMETER Silent
  Reserve (compatibilite).
.PARAMETER Force
  Retraite tous les mails, meme deja importes.
#>
param([switch]$DryRun, [switch]$Silent, [switch]$Force)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) { $ScriptDir = $PWD.Path }
. (Join-Path $ScriptDir "bal_to_csv_rovo.functions.ps1")

# ========== CONFIGURATION ==========
$Config = Get-Config -Path (Join-Path $ScriptDir "config.json")
$JiraCategory       = $Config.JiraCategory
$DefaultAssignee    = $Config.DefaultAssignee
$DefaultPriority    = $Config.DefaultPriority
$DefaultPriorityBug = $Config.DefaultPriorityBug
$MaxFetchPerBal     = [int]$Config.MaxFetchPerBal
$MaxMailAgeDays     = [int]$Config.MaxMailAgeDays
$RetentionDays      = [int]$Config.RetentionDays
$Mailboxes          = Get-Mailboxes -Config $Config
$Projects           = Get-Projects -Config $Config

$LogDate = Get-Date -Format "yyyy-MM-dd"
$LogFile = Join-Path $ScriptDir "log_bal_to_csv_$LogDate.txt"
$ProcessedIdsFile = Join-Path $ScriptDir "processed_msgids.txt"
$LockFile = Join-Path $ScriptDir "bal_to_csv.lock"
$CsvFile = Join-Path $ScriptDir "bal_to_jira_$LogDate.csv"

# ========== MAIN ==========
if (Test-Path $LockFile) {
    $lockPid = Get-Content $LockFile -ErrorAction SilentlyContinue
    if ($lockPid -and (Get-Process -Id $lockPid -ErrorAction SilentlyContinue)) {
        Log-Msg ERROR "Script deja en cours (PID: $lockPid)"
        exit 1
    } else {
        Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
    }
}
$currentPid = [System.Diagnostics.Process]::GetCurrentProcess().Id
Set-Content $LockFile $currentPid -Force
Log-Msg INFO "Lock acquis (PID: $currentPid)"

$outlook = $null

try {
    Log-Msg INFO "========================================"
    Log-Msg INFO "DEMARRAGE bal_to_csv_rovo.ps1"
    Log-Msg INFO "PC: $env:COMPUTERNAME | User: $env:USERNAME"
    Log-Msg INFO "========================================"

    Remove-OldLogs -Directory $ScriptDir -Pattern "log_bal_to_csv_*.txt" -RetentionDays $RetentionDays

    $outlook = Connect-Outlook
    Ensure-JiraCategory -Outlook $outlook -JiraCategory $JiraCategory
    $processedIds = @{}
    if (-not $Force) {
        $processedIds = Import-TraceFile -Path $ProcessedIdsFile
        Log-Msg INFO "$($processedIds.Count) Message-ID deja traites"
    }

    $csvLines = [System.Collections.Generic.List[hashtable]]::new()
    $totalProcessed = 0

    foreach ($mb in $Mailboxes) {
        $balAddress = $mb.Address
        $projectKey = $mb.ProjectKey
        Log-Msg INFO "-------- BAL: $balAddress --------"

        $project = $Projects[$projectKey]

        $inbox = $null
        $namespace = $null
        try {
            $inbox = Get-OutlookInbox -Outlook $outlook -BalAddress $balAddress
            if (-not $inbox) { Log-Msg WARN "BAL non trouvee: $balAddress"; continue }

            Log-Msg INFO "Traitement BAL: $balAddress (chemin: ${inbox.FolderPath})"
            $mails = Get-AllMails -Inbox $inbox -MaxFetch $MaxFetchPerBal -MaxAgeDays $MaxMailAgeDays
            if ($mails.Count -eq 0) { Log-Msg INFO "Aucun mail a traiter"; continue }

            foreach ($mail in $mails) {
                try {
                    $headers = Get-MailHeaders -mail $mail
                    $subject = $headers["Subject"]
                    $from = $headers["From"]
                    $body = Get-MailBody -mail $mail

                    $mailId = Get-DeduplicationId -MessageId $headers["Message-ID"] -Subject $subject -From $from -Body $body

                    $isReply = ($headers["In-Reply-To"] -ne "" -or $headers["References"] -ne "")
                    if (-not $Force -and $processedIds.ContainsKey($mailId)) {
                        # Deja importe: on ignore, sauf si c'est une reponse a un mail importe
                        if (-not $isReply) { continue }
                        Log-Msg INFO "Reponse recue sur un mail deja importe: $subject"
                    }

                    $issueType = Classify-Mail -subject $subject -body $body
                    $priority = if ($issueType -eq "Anomalie") { $DefaultPriorityBug } else { $DefaultPriority }

                    $resume = ($subject -replace '^(re|tr|fwd|fw):\s*', '').Trim()
                    if ($resume.Length -gt 120) { $resume = $resume.Substring(0, 120) + "..." }

                    $mailLink = "outlook:" + $mail.EntryID
                    $description = $body
                    if ($description.Length -gt 3000) { $description = $description.Substring(0, 3000) + "..." }
                    $description = "$description`r`n`r`nLien vers le mail: [Ouvrir le mail dans Outlook|$mailLink]"

                    if (-not $DryRun) { Set-JiraFlag -Mail $mail -JiraCategory $JiraCategory }

                    $csvLines += @{
                        Projet = $project.Key
                        Type_de_ticket = $issueType
                        Statut = "Nouveau"
                        Resume = $resume
                        Description = $description
                        Priorite = $priority
                        Assigne = $DefaultAssignee
                        Rapporteur = $from
                        Date_de_reception = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                    }

                    $processedIds[$mailId] = $true
                    $totalProcessed++
                    Log-Msg DEBUG "Mail traite: $subject ($issueType)"

                } catch {
                    Log-Msg ERROR "Erreur mail: $_"
                }
            }
        } finally {
            Release-ComObject $inbox
        }
    }

    # Generation CSV - le fichier est TOUJOURS cree (au minimum l'en-tete),
    # meme si aucun mail n'a ete traite, afin de ne jamais laisser un CSV absent/vide.
    if (-not $DryRun) {
        $csvContent = ConvertTo-CsvContent -Rows $csvLines
        $csvDir = Split-Path -Parent $CsvFile
        if ($csvDir -and -not (Test-Path $csvDir)) { New-Item -ItemType Directory -Path $csvDir -Force | Out-Null }
        # UTF-8 avec BOM pour une lecture correcte des accents dans Excel FR
        [System.IO.File]::WriteAllText($CsvFile, $csvContent, (New-Object System.Text.UTF8Encoding($true)))
        Log-Msg INFO "CSV genere: $CsvFile ($($csvLines.Count) ligne(s) de donnees)"
    } else {
        Log-Msg INFO "DryRun: CSV non ecrit ($($csvLines.Count) ligne(s) pretes)"
    }

    Log-Msg INFO "========================================"
    Log-Msg INFO "FIN: $totalProcessed mails traites"
    Log-Msg INFO "========================================"

    # Sauvegarde des Message-ID traites pour le dedoublonnage ulterieur
    try {
        if (-not $DryRun -and $totalProcessed -gt 0) {
            Save-TraceFile -Path $ProcessedIdsFile -Trace $processedIds
            Log-Msg INFO "Trace d'import mise a jour: $ProcessedIdsFile ($($processedIds.Count) Message-ID)"
        }
    } catch {
        Log-Msg WARN "Impossible de sauvegarder les Message-ID traites: $_"
    }

} catch {
    Log-Msg ERROR "Erreur fatale: $_"
} finally {
    if (Test-Path $LockFile) { Remove-Item $LockFile -Force -ErrorAction SilentlyContinue }
    Release-ComObject $outlook
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
