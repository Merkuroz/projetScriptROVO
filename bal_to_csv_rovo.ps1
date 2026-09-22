<#
.SYNOPSIS
  Script pour traiter les BAL Enedis via Outlook COM et generer un CSV importable dans Jira,
  ou creer directement les tickets via l'API REST Jira (optionnel).
.DESCRIPTION
  Les fonctions sont dans bal_to_csv_rovo.functions.ps1.
  La configuration (BAL, projets, priorites...) est dans config.json (valeurs par defaut si absent).
  Les mails sont marques dans Outlook seulement APRES ecriture reussie du CSV / creation Jira :
  ainsi un echec d'ecriture ne fait perdre aucun mail.
.PARAMETER DryRun
  N'ecrit ni CSV, ni categorie Outlook, ni trace d'import, ni ticket Jira.
.PARAMETER Silent
  Supprime la sortie console (le fichier de log reste ecrit).
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
$TraceRetentionDays = [int]$Config.TraceRetentionDays
$CsvWriteRetries    = [int]$Config.CsvWriteRetries
$Mailboxes          = Get-Mailboxes -Config $Config
$Projects           = Get-Projects -Config $Config

$script:Silent = ($Silent.IsPresent -or [bool]$Config.Silent)
$script:LogFile = Join-Path $ScriptDir "log_bal_to_csv_$(Get-Date -Format 'yyyy-MM-dd').txt"
$ProcessedIdsFile = Join-Path $ScriptDir "processed_msgids.txt"
$LockFile = Join-Path $ScriptDir "bal_to_csv.lock"
$CsvFile = Join-Path $ScriptDir "bal_to_jira_$(Get-Date -Format 'yyyy-MM-dd').csv"

$exitCode = 0
$outlook = $null

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
        $processedIds = Import-TraceFile -Path $ProcessedIdsFile -RetentionDays $TraceRetentionDays
        Log-Msg INFO "$($processedIds.Count) Message-ID deja traites"
    }

    $csvLines = [System.Collections.Generic.List[hashtable]]::new()
    $pendingMails = [System.Collections.Generic.List[object]]::new()
    $totalProcessed = 0

    foreach ($mb in $Mailboxes) {
        $balAddress = $mb.Address
        $projectKey = $mb.ProjectKey
        Log-Msg INFO "-------- BAL: $balAddress --------"

        $project = $Projects[$projectKey]
        if (-not $project) {
            Log-Msg ERROR "Projet inconnu '$projectKey' pour la BAL $balAddress (verifier config.json)"
            $exitCode = 1
            continue
        }

        $inbox = $null
        try {
            $inbox = Get-OutlookInbox -Outlook $outlook -BalAddress $balAddress
            if (-not $inbox) { Log-Msg WARN "BAL non trouvee: $balAddress"; $exitCode = 1; continue }

            Log-Msg INFO "Traitement BAL: $balAddress (chemin: ${inbox.FolderPath})"
            # Les mails deja traites (categorises) sont exclus cote Outlook :
            # le quota MaxFetch n'est plus consomme par des mails deja importes.
            $mails = Get-AllMails -Inbox $inbox -MaxFetch $MaxFetchPerBal -MaxAgeDays $MaxMailAgeDays -ExcludeCategory $(if ($Force) { "" } else { $JiraCategory })
            if ($mails.Count -eq 0) { Log-Msg INFO "Aucun mail a traiter"; continue }

            foreach ($mail in $mails) {
                try {
                    $headers = Get-MailHeaders -mail $mail
                    $subject = $headers["Subject"]
                    $from = $headers["From"]
                    $body = Get-MailBody -mail $mail
                    $receivedTime = $mail.ReceivedTime
                    $receivedTxt = $receivedTime.ToString("yyyy-MM-dd HH:mm:ss")

                    $mailId = Get-DeduplicationId -MessageId $headers["Message-ID"] -Subject $subject -From $from -Body $body -ReceivedTime $receivedTxt

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

                    # Lien cliquable (wiki Jira) + Message-ID en clair pour retrouver
                    # le mail meme s'il est deplace (l'EntryID change avec le dossier).
                    $mailLink = "outlook:" + $mail.EntryID
                    $description = $body
                    if ($description.Length -gt 3000) { $description = $description.Substring(0, 3000) + "..." }
                    $description = "$description`r`n`r`nLien vers le mail: [Ouvrir le mail dans Outlook|$mailLink]"
                    if ($headers["Message-ID"] -ne "") {
                        $description = "$description`r`nMessage-ID: $($headers['Message-ID'])"
                    }

                    $csvLines += @{
                        Projet = $project.Key
                        Type_de_ticket = $issueType
                        Statut = "Nouveau"
                        Resume = $resume
                        Description = $description
                        Priorite = $priority
                        Assigne = $DefaultAssignee
                        Rapporteur = $from
                        Date_de_reception = $receivedTxt
                        Composant = $project.Composants
                        Etiquettes = $project.Etiquettes
                    }
                    # Marquage differe: la categorie Outlook n'est appliquee qu'apres
                    # ecriture reussie du CSV (ou creation Jira), voir plus bas.
                    $pendingMails.Add($mail)

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

    # ========== ECRITURE CSV ==========
    # Le fichier est TOUJOURS cree (au minimum l'en-tete), meme si aucun mail
    # n'a ete traite, afin de ne jamais laisser un CSV absent/vide.
    $csvWritten = $true
    if (-not $DryRun) {
        $csvContent = ConvertTo-CsvContent -Rows $csvLines
        $csvWritten = Write-CsvFile -Path $CsvFile -Content $csvContent -Retries $CsvWriteRetries
        if ($csvWritten) {
            Log-Msg INFO "CSV genere: $CsvFile ($($csvLines.Count) ligne(s) de donnees)"
        } else {
            $exitCode = 1
        }
    } else {
        Log-Msg INFO "DryRun: CSV non ecrit ($($csvLines.Count) ligne(s) pretes)"
    }

    # ========== IMPORT JIRA DIRECT (optionnel) ==========
    $jiraResult = $null
    if (-not $DryRun -and $csvWritten -and $Config.Jira.Enabled -and $csvLines.Count -gt 0) {
        $jiraResult = Invoke-JiraImport -Rows $csvLines -Config $Config
        if ($jiraResult.Failed.Count -gt 0) { $exitCode = 1 }
    } elseif ($Config.Jira.Enabled -and -not $DryRun) {
        Log-Msg INFO "Import Jira actif mais aucun mail a importer"
    }

    # ========== MARQUAGE OUTLOOK (apres ecriture reussie) ==========
    if (-not $DryRun -and $csvWritten) {
        foreach ($mail in $pendingMails) {
            Set-JiraFlag -Mail $mail -JiraCategory $JiraCategory
        }
        Log-Msg INFO "$($pendingMails.Count) mail(s) categorise(s) '$JiraCategory'"

        # Sauvegarde des Message-ID traites pour le dedoublonnage ulterieur
        try {
            if ($totalProcessed -gt 0) {
                Save-TraceFile -Path $ProcessedIdsFile -Trace $processedIds
                Log-Msg INFO "Trace d'import mise a jour: $ProcessedIdsFile ($($processedIds.Count) Message-ID)"
            }
        } catch {
            Log-Msg WARN "Impossible de sauvegarder les Message-ID traites: $_"
            $exitCode = 1
        }
    }

    $finMsg = "FIN: $totalProcessed mails traites"
    if ($jiraResult) { $finMsg += " | Jira: $($jiraResult.Success) crees, $($jiraResult.Failed.Count) echecs" }
    Log-Msg INFO "========================================"
    Log-Msg INFO $finMsg
    Log-Msg INFO "========================================"

} catch {
    Log-Msg ERROR "Erreur fatale: $_"
    $exitCode = 1
} finally {
    if (Test-Path $LockFile) { Remove-Item $LockFile -Force -ErrorAction SilentlyContinue }
    Release-ComObject $outlook
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    exit $exitCode
}
