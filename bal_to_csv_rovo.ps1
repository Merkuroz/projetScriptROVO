<#
.SYNOPSIS
  Script corrigé pour traiter les BAL Enedis via Outlook COM
#>
param([switch]$DryRun, [switch]$Silent)

# ========== CONFIGURATION ==========
$OutputDir = Join-Path $env:USERPROFILE "Documents"
$DefaultAssignee = "frederic.izard@enedis.fr"
$DefaultPriority = "Moyenne"
$DefaultPriorityBug = "Haute"
$MaxFetchPerBal = 50
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) { $ScriptDir = $PWD.Path }
$LogDate = Get-Date -Format "yyyy-MM-dd"
$LogFile = Join-Path $ScriptDir "log_bal_to_csv_$LogDate.txt"
$ProcessedIdsFile = Join-Path $ScriptDir "processed_msgids.txt"
$LockFile = Join-Path $ScriptDir "bal_to_csv.lock"
$CsvFile = Join-Path $OutputDir "bal_to_jira_$LogDate.csv"

$Mailboxes = @(
    @{ Address = "DFP-UOF-IPED-DIP"; Name = "DFP-UOF-IPED-DIP" },
    @{ Address = "DFP-ADMIN-FORM-MYHR"; Name = "DFP-ADMIN-FORM-MYHR" }
)
$Projects = @{
    "Y6S" = @{ Key = "Y6S"; Prefix = "[eCampus]"; Composants = ""; Etiquettes = "" }
    "MHR" = @{ Key = "MHR"; Prefix = "[MyHR]"; Composants = "Injections"; Etiquettes = "" }
}

# ========== FONCTIONS UTILITAIRES ==========
function Normalize-Name {
    param([string]$name)
    $name = $name.ToLower()
    # Conserver les tirets et underscores
    $name = $name -replace '[àâä]', 'a' -replace '[éèêë]', 'e' -replace '[îï]', 'i' -replace '[ôö]', 'o' -replace '[ùûü]', 'u' -replace 'ç', 'c'
    $name = $name -replace '[^a-z0-9\\-_]', ''  # Garde a-z, 0-9, -, _
    return $name
}

function Log-Msg {
    param([string]$Level="INFO", [string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "[$ts] $Level $Message"
    Write-Host $line -ForegroundColor $(if($Level -eq "ERROR"){"Red"} elseif($Level -eq "WARN"){"Yellow"} elseif($Level -eq "DEBUG"){"DarkGray"} else{"White"})
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# ========== FONCTIONS OUTLOOK (CORRIGEES) ==========
function Connect-Outlook {
    try {
        $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        Log-Msg INFO "Outlook deja ouvert"
    } catch {
        try {
            $outlook = New-Object -ComObject Outlook.Application
            Log-Msg INFO "Outlook lance"
        } catch {
            throw "Impossible de lancer Outlook"
        }
    }
    return $outlook
}

function Get-OutlookInbox {
    param($Outlook, [string]$BalAddress)
    $namespace = $Outlook.GetNamespace("MAPI")
    $normBal = Normalize-Name $BalAddress
    Log-Msg DEBUG "Recherche BAL: '$BalAddress' (normalise: '$normBal')"

    # Recherche directe dans les stores
    foreach ($store in $namespace.Stores) {
        $normStoreName = Normalize-Name $store.DisplayName
        if ($normStoreName -eq $normBal) {
            $root = $store.GetRootFolder()

            # Priorite absolue: Boite de reception
            foreach ($sub in $root.Folders) {
                $normSubName = Normalize-Name $sub.Name
                if ($normSubName -eq "boitedereception" -or $normSubName -eq "inbox") {
                    if ($sub.Items.Count -gt 0) {
                        try {
                            $firstItem = $sub.Items.GetFirst()
                            if ($firstItem -ne $null -and ($firstItem.Class -eq 43 -or $firstItem.MessageClass -eq "IPM.Note")) {
                                Log-Msg INFO "BAL TROUVEE: ${sub.FolderPath} ($($sub.Items.Count) items)"
                                return $sub
                            }
                        } catch {
                            Log-Msg WARN "Erreur verification dossier ${sub.FolderPath}: $_"
                        }
                    }
                }
            }

            # Si pas de Boite de reception, retourner la racine
            if ($root.Items.Count -gt 0) {
                Log-Msg INFO "BAL TROUVEE (racine): ${root.FolderPath} ($($root.Items.Count) items)"
                return $root
            }
        }
    }

    # Recherche recursive
    $bestFolder = $null
    foreach ($store in $namespace.Stores) {
        $root = $store.GetRootFolder()
        $stack = [System.Collections.Stack]::new()
        $stack.Push($root)

        while ($stack.Count -gt 0 -and $bestFolder -eq $null) {
            $folder = $stack.Pop()
            if ($folder.DefaultItemType -eq 1 -or $folder.DefaultItemType -eq 9) { continue }

            $normName = Normalize-Name $folder.Name
            if ($normName -eq $normBal -or (Normalize-Name $folder.FolderPath) -like "*$normBal*") {
                if ($folder.Items.Count -gt 0) {
                    try {
                        $firstItem = $folder.Items.GetFirst()
                        if ($firstItem -ne $null -and ($firstItem.Class -eq 43 -or $firstItem.MessageClass -eq "IPM.Note")) {
                            if ($normName -eq "boitedereception" -or $normName -eq "inbox") {
                                $bestFolder = $folder
                                break
                            } elseif ($bestFolder -eq $null) {
                                $bestFolder = $folder
                            }
                        }
                    } catch {
                        Log-Msg WARN "Erreur verification dossier ${folder.FolderPath}: $_"
                    }
                }
            }
            foreach ($sub in $folder.Folders) { $stack.Push($sub) }
        }
    }

    if ($bestFolder -ne $null) {
        Log-Msg INFO "BAL TROUVEE: ${bestFolder.FolderPath} ($($bestFolder.Items.Count) items)"
        return $bestFolder
    }

    Log-Msg ERROR "BAL NON TROUVEE: $BalAddress"
    return $null
}

function Get-AllMails {
    param($Inbox, [int]$MaxFetch)
    $items = $Inbox.Items
    try { $items.Sort("[ReceivedTime]", $true) } catch {}
    $mails = @()
    $count = 0
    foreach ($item in $items) {
        if ($count -ge $MaxFetch) { break }
        try {
            if ($item -ne $null -and $item.Class -eq 43) {
                $mails += $item
                $count++
            }
        } catch {}
    }
    Log-Msg INFO "$count mail(s) trouves"
    return $mails
}

# ========== PARSING EMAIL (SECURISE) ==========
function Get-MailHeaders {
    param($mail)
    $headers = @{
        "Message-ID" = ""
        "From" = ""
        "To" = ""
        "Cc" = ""
        "Subject" = ""
        "In-Reply-To" = ""
        "References" = ""
    }

    try {
        if ($mail -eq $null) { return $headers }

        # Proprietes de base
        if ($mail.SenderEmailAddress) { $headers["From"] = $mail.SenderEmailAddress }
        if ($mail.To) { $headers["To"] = $mail.To }
        if ($mail.CC) { $headers["Cc"] = $mail.CC }
        if ($mail.Subject) { $headers["Subject"] = $mail.Subject }

        # Headers Internet
        if ($mail.PropertyAccessor) {
            $pa = $mail.PropertyAccessor
            $PR_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
            $PR_IN_REPLY_TO = "http://schemas.microsoft.com/mapi/proptag/0x1042001E"
            $PR_REFERENCES = "http://schemas.microsoft.com/mapi/proptag/0x1039001E"

            try { if ($pa.GetProperty($PR_MESSAGE_ID)) { $headers["Message-ID"] = $pa.GetProperty($PR_MESSAGE_ID) } } catch {}
            try { if ($pa.GetProperty($PR_IN_REPLY_TO)) { $headers["In-Reply-To"] = $pa.GetProperty($PR_IN_REPLY_TO) } } catch {}
            try { if ($pa.GetProperty($PR_REFERENCES)) { $headers["References"] = $pa.GetProperty($PR_REFERENCES) } } catch {}
        }
    } catch {
        Log-Msg ERROR "Erreur lecture headers: $_"
    }
    return $headers
}

function Get-MailBody {
    param($mail)
    $body = ""
    try {
        if ($mail -ne $null) {
            if ($mail.HTMLBody) {
                $body = $mail.HTMLBody -replace "<[^>]+>", " " -replace "\s+", " "
            } elseif ($mail.Body) {
                $body = $mail.Body
            }
        }
    } catch {
        try { if ($mail.Body) { $body = $mail.Body } } catch {}
    }
    return $body.Trim()
}

# ========== CLASSIFICATION ==========
function Classify-Mail {
    param($subject, $body)
    $text = ($subject + " " + $body).ToLower()
    $scores = @{"Anomalie"=0; "User Story"=0; "Support"=0; "Tache"=0}

    # Mots-cles simplifies
    if ($text -match "anomalie|bug|incident|panne|erreur|ne fonctionne pas") { $scores["Anomalie"] += 3 }
    if ($text -match "besoin|demande|evolution|nouvelle") { $scores["User Story"] += 3 }
    if ($text -match "comment faire|besoin d'aide|assistance") { $scores["Support"] += 3 }
    if ($text -match "configuration|parametrage|administration") { $scores["Tache"] += 3 }

    $bestType = "User Story"
    $bestScore = 0
    foreach ($type in $scores.Keys) {
        if ($scores[$type] -gt $bestScore) {
            $bestScore = $scores[$type]
            $bestType = $type
        }
    }
    return $bestType
}

# ========== MAIN SCRIPT ==========
if (Test-Path $LockFile) {
    $pid = Get-Content $LockFile -ErrorAction SilentlyContinue
    if ($pid -and (Get-Process -Id $pid -ErrorAction SilentlyContinue)) {
        Log-Msg ERROR "Script deja en cours (PID: $pid)"
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

    $outlook = Connect-Outlook
    $processedIds = @{}
    if (Test-Path $ProcessedIdsFile) {
        $processedIds = @{}
        Get-Content $ProcessedIdsFile -Encoding UTF8 | ForEach-Object {
            $line = $_.Trim()
            if ($line -ne "") { $processedIds[$line] = $true }
        }
        Log-Msg INFO "$($processedIds.Count) Message-ID deja traites"
    }

    $csvLines = @()
    $totalProcessed = 0

    foreach ($mb in $Mailboxes) {
        $balAddress = $mb.Address
        Log-Msg INFO "-------- BAL: $balAddress --------"

        $projectKey = if ($balAddress -like "*DFP-UOF-IPED-DIP*") { "Y6S" } elseif ($balAddress -like "*DFP-ADMIN-FORM-MYHR*") { "MHR" } else { "Y6S" }
        $project = $Projects[$projectKey]

        $inbox = Get-OutlookInbox -Outlook $outlook -BalAddress $balAddress
        if (-not $inbox) { Log-Msg WARN "BAL non trouvee: $balAddress"; continue }

        Log-Msg INFO "Traitement BAL: $balAddress (chemin: ${inbox.FolderPath})"
        $mails = Get-AllMails -Inbox $inbox -MaxFetch $MaxFetchPerBal
        if ($mails.Count -eq 0) { Log-Msg INFO "Aucun mail a traiter"; continue }

        foreach ($mail in $mails) {
            try {
                $headers = Get-MailHeaders -mail $mail
                $subject = $headers["Subject"]
                $from = $headers["From"]
                $body = Get-MailBody -mail $mail

                $mailId = if ($headers["Message-ID"] -ne "") { $headers["Message-ID"] } else { $subject + $from + $body }

                if ($processedIds.ContainsKey($mailId)) { continue }

                $issueType = Classify-Mail -subject $subject -body $body
                $priority = if ($issueType -eq "Anomalie") { $DefaultPriorityBug } else { $DefaultPriority }

                $resume = ($subject -replace '^(re|tr|fwd|fw):\s*', '').Trim()
                if ($resume.Length -gt 120) { $resume = $resume.Substring(0, 120) + "..." }

                $description = $body
                if ($description.Length -gt 3000) { $description = $description.Substring(0, 3000) + "..." }

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
                Log-Msg DEBUG "Mail traite: $subject"

            } catch {
                Log-Msg ERROR "Erreur mail: $_"
            }
        }
    }

    # Generation CSV
    if ($csvLines.Count -gt 0) {
        $csvHeaders = @("Projet","Type de ticket","Statut","Resume","Description","Priorite","Assigne","Rapporteur","Date de reception")
        $csvContent = ($csvHeaders -join ";") + "`r`n"
        foreach ($line in $csvLines) {
            $row = @(
                $line.Projet,
                $line.Type_de_ticket,
                $line.Statut,
                ($line.Resume -replace '"','""'),
                ($line.Description -replace '"','""'),
                $line.Priorite,
                $line.Assigne,
                $line.Rapporteur,
                $line.Date_de_reception
            ) -join ";"
            $csvContent += $row + "`r`n"
        }
        $csvContent | Out-File -FilePath $CsvFile -Encoding UTF8
        Log-Msg INFO "CSV genere: $CsvFile ($($csvLines.Count) lignes)"
    } else {
        Log-Msg INFO "Aucun mail a exporter"
    }

    Log-Msg INFO "========================================"
    Log-Msg INFO "FIN: $totalProcessed mails traites"
    Log-Msg INFO "========================================"

    # Sauvegarde des Message-ID traites pour le dedoublonnage ulterieur
    try {
        $processedIds.Keys | Out-File -FilePath $ProcessedIdsFile -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {
        Log-Msg WARN "Impossible de sauvegarder les Message-ID traites: $_"
    }

} catch {
    Log-Msg ERROR "Erreur fatale: $_"
} finally {
    if (Test-Path $LockFile) { Remove-Item $LockFile -Force -ErrorAction SilentlyContinue }
    if ($outlook) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}