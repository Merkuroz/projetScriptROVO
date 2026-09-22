<#
.SYNOPSIS
  Fonctions du script bal_to_csv_rovo.ps1
  Ce fichier est charge (dot-source) par le script principal et par les tests Pester.
#>

# ========== NORMALISATION ==========
function Normalize-Name {
    param([string]$name)
    $name = $name.ToLower()
    # Conserver les tirets et underscores
    $name = $name -replace '[àâä]', 'a' -replace '[éèêë]', 'e' -replace '[îï]', 'i' -replace '[ôö]', 'o' -replace '[ùûü]', 'u' -replace 'ç', 'c'
    $name = $name -replace '[^a-z0-9\-_]', ''  # Garde a-z, 0-9, -, _
    return $name
}

# ========== JOURNALISATION ==========
function Log-Msg {
    param([string]$Level="INFO", [string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "[$ts] $Level $Message"
    if (-not $script:Silent) {
        Write-Host $line -ForegroundColor $(if($Level -eq "ERROR"){"Red"} elseif($Level -eq "WARN"){"Yellow"} elseif($Level -eq "DEBUG"){"DarkGray"} else{"White"})
    }
    try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Release-ComObject {
    param($Object)
    if ($Object) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null } catch {}
    }
}

# ========== CONFIGURATION ==========
function Get-Config {
    param([string]$Path)
    $config = [pscustomobject]@{
        JiraCategory            = "a traiter dans Jira"
        DefaultAssignee         = "frederic.izard@enedis.fr"
        DefaultPriority         = "Moyenne"
        DefaultPriorityBug      = "Haute"
        MaxFetchPerBal          = 50
        MaxMailAgeDays          = 0
        RetentionDays           = 90
        TraceRetentionDays      = 365
        CsvWriteRetries         = 5
        Silent                  = $false
        Jira                    = [pscustomobject]@{
            Enabled      = $false
            BaseUrl     = ""
            Email       = ""
            ApiTokenEnvVar = "JIRA_API_TOKEN"
        }
    }
    if (Test-Path $Path) {
        try {
            $json = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $json.PSObject.Properties) {
                if ($null -ne $p.Value) {
                    if ($p.Name -eq "Jira" -and $p.Value -is [pscustomobject]) {
                        foreach ($jp in $p.Value.PSObject.Properties) {
                            if ($null -ne $jp.Value) { $config.Jira | Add-Member -MemberType NoteProperty -Name $jp.Name -Value $jp.Value -Force }
                        }
                    } else {
                        $config | Add-Member -MemberType NoteProperty -Name $p.Name -Value $p.Value -Force
                    }
                }
            }
        } catch {
            Log-Msg ERROR "config.json illisible, valeurs par defaut utilisees: $_"
        }
    } else {
        Log-Msg WARN "config.json absent, valeurs par defaut utilisees: $Path"
    }
    return $config
}

function Get-Mailboxes {
    param($Config)
    $default = @(
        @{ Address = "DFP-UOF-IPED-DIP"; Name = "DFP-UOF-IPED-DIP"; ProjectKey = "Y6S" },
        @{ Address = "DFP-ADMIN-FORM-MYHR"; Name = "DFP-ADMIN-FORM-MYHR"; ProjectKey = "MHR" }
    )
    if ($Config.Mailboxes) {
        $result = @()
        foreach ($mb in @($Config.Mailboxes)) {
            $result += @{ Address = [string]$mb.Address; Name = [string]$mb.Name; ProjectKey = [string]$mb.ProjectKey }
        }
        return $result
    }
    return $default
}

function Get-Projects {
    param($Config)
    $default = @{
        "Y6S" = @{ Key = "Y6S"; Prefix = "[eCampus]"; Composants = ""; Etiquettes = "" }
        "MHR" = @{ Key = "MHR"; Prefix = "[MyHR]"; Composants = "Injections"; Etiquettes = "" }
    }
    if ($Config.Projects) {
        $result = @{}
        foreach ($p in $Config.Projects.PSObject.Properties) {
            $result[$p.Name] = @{
                Key        = [string]$p.Value.Key
                Prefix     = [string]$p.Value.Prefix
                Composants = [string]$p.Value.Composants
                Etiquettes = [string]$p.Value.Etiquettes
            }
        }
        return $result
    }
    return $default
}

# ========== OUTLOOK ==========
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
    param($Inbox, [int]$MaxFetch, [int]$MaxAgeDays = 0, [string]$ExcludeCategory = "")
    # Filtre cote Outlook (DASL): IPM.Note, non categorises, age optionnel.
    # Beaucoup plus rapide qu'un parcours COM complet, et les mails deja traites
    # (categorises) ne consomment plus le quota MaxFetch.
    $PR_MESSAGE_CLASS = '"http://schemas.microsoft.com/mapi/proptag/0x001A001E"'
    $filter = "@SQL=$PR_MESSAGE_CLASS = 'IPM.Note'"
    if ($ExcludeCategory -ne "") {
        $filter += " AND (NOT `"urn:schemas-microsoft-com:office:office#Keywords`" LIKE '%$ExcludeCategory%')"
    }
    if ($MaxAgeDays -gt 0) {
        $since = (Get-Date).ToUniversalTime().AddDays(-$MaxAgeDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filter += " AND `"urn:schemas:httpmail:date`" >= '$since'"
    }
    Log-Msg DEBUG "Filtre Outlook: $filter"

    $items = $Inbox.Items
    try {
        $items = $Inbox.Items.Restrict($filter)
    } catch {
        Log-Msg WARN "Restrict DASL impossible ($_), retour au filtre MessageClass seul"
        try { $items = $Inbox.Items.Restrict("[MessageClass] = 'IPM.Note'") } catch { Log-Msg WARN "Restrict impossible, parcours complet: $_" }
    }

    try { $items.Sort("[ReceivedTime]", $true) } catch {}

    $mails = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $items) {
        if ($mails.Count -ge $MaxFetch) { break }
        try {
            if ($item -ne $null -and $item.Class -eq 43) { $mails.Add($item) }
        } catch {}
    }
    Log-Msg INFO "$($mails.Count) mail(s) a traiter"
    return $mails
}

# ========== PARSING EMAIL ==========
function Get-MailHeaders {
    param($mail)
    $headers = @{
        "Message-ID"  = ""
        "From"        = ""
        "To"          = ""
        "Cc"          = ""
        "Subject"     = ""
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
            $PR_MESSAGE_ID  = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
            $PR_IN_REPLY_TO = "http://schemas.microsoft.com/mapi/proptag/0x1042001E"
            $PR_REFERENCES  = "http://schemas.microsoft.com/mapi/proptag/0x1039001E"

            try { $v = $pa.GetProperty($PR_MESSAGE_ID);  if ($v) { $headers["Message-ID"]  = $v } } catch {}
            try { $v = $pa.GetProperty($PR_IN_REPLY_TO); if ($v) { $headers["In-Reply-To"] = $v } } catch {}
            try { $v = $pa.GetProperty($PR_REFERENCES);  if ($v) { $headers["References"]  = $v } } catch {}
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
                # Supprime les balises puis decode les entites HTML (&eacute;, &#232;, &amp;, ...)
                $body = $mail.HTMLBody -replace "<[^>]+>", " "
                try { $body = [System.Net.WebUtility]::HtmlDecode($body) } catch {}
                $body = $body -replace "\s+", " "
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
    param([string]$subject, [string]$body)
    $text = ($subject + " " + $body).ToLower()
    # Ordre fixe pour un resultat deterministe en cas d'egalite de score
    $scores = [ordered]@{
        "Anomalie"    = 0
        "User Story" = 0
        "Support"    = 0
        "Tache"      = 0
    }

    # Mots-cles simplifies
    if ($text -match "anomalie|bug|incident|panne|erreur|ne fonctionne pas") { $scores["Anomalie"] += 3 }
    # "besoin d'aide" est reserve au Support, pas a User Story
    if ($text -match "besoin(?! d'aide)|demande|evolution|nouvelle") { $scores["User Story"] += 3 }
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

# ========== DEDUPLICATION ==========
function Get-DeduplicationId {
    param([string]$MessageId, [string]$Subject, [string]$From, [string]$Body, [string]$ReceivedTime = "")
    if ($MessageId) { return "mid:$MessageId" }
    # ReceivedTime distingue deux mails au contenu identique (reponses en chaine)
    $raw = "$Subject|$From|$ReceivedTime|$Body"
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        $hash = $sha.ComputeHash($bytes)
        return "sha256:" + ([BitConverter]::ToString($hash) -replace '-','')
    } finally {
        $sha.Dispose()
    }
}

# ========== TRACE D'IMPORT ==========
function Import-TraceFile {
    param([string]$Path, [int]$RetentionDays = 0)
    $trace = @{}
    $cutoff = ""
    if ($RetentionDays -gt 0) { $cutoff = (Get-Date).AddDays(-$RetentionDays).ToString("yyyy-MM-ddTHH:mm:ss") }
    if (Test-Path $Path) {
        foreach ($line in (Get-Content $Path -Encoding UTF8)) {
            $t = $line.Trim()
            if ($t -eq "") { continue }
            # Format horodate: "2024-01-01T08:00:00|<id>"; format historique: "<id>"
            $sep = $t.IndexOf("|")
            if ($sep -gt 0) {
                $stamp = $t.Substring(0, $sep)
                $id = $t.Substring($sep + 1)
                if ($cutoff -ne "" -and $stamp -lt $cutoff) { continue }
            } else {
                $id = $t
            }
            if (-not $trace.ContainsKey($id)) { $trace[$id] = $true }
        }
    }
    return $trace
}

function Save-TraceFile {
    param([string]$Path, $Trace)
    # Ecriture atomique: fichier temporaire puis remplacement, pour ne jamais
    # corrompre la trace en cas de crash au milieu de l'ecriture.
    $tmp = "$Path.tmp"
    $lines = foreach ($id in $Trace.Keys) { "$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')|$id" }
    try {
        [System.IO.File]::WriteAllLines($tmp, $lines, (New-Object System.Text.UTF8Encoding($true)))
        Move-Item -Path $tmp -Destination $Path -Force
    } catch {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw
    }
}

function Remove-OldLogs {
    param([string]$Directory, [string]$Pattern, [int]$RetentionDays)
    if ($RetentionDays -le 0) { return }
    try {
        $cutoff = (Get-Date).AddDays(-$RetentionDays)
        Get-ChildItem -Path $Directory -Filter $Pattern -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {
        Log-Msg WARN "Nettoyage des anciens logs impossible: $_"
    }
}

# ========== CATEGORIES OUTLOOK ==========
function Ensure-JiraCategory {
    param($Outlook, [string]$JiraCategory)
    try {
        $categories = $Outlook.GetNamespace("MAPI").Categories
        $existing = $null
        foreach ($cat in $categories) {
            if ($cat.Name -eq $JiraCategory) { $existing = $cat; break }
        }
        if ($existing -eq $null) {
            $usedColors = @{}
            foreach ($cat in $categories) { $usedColors[[int]$cat.Color] = $true }
            # Couleur souhaitee: Dark Maroon (rose fonce, valeur 25)
            $preferred = 25
            $palette = @(25, 24, 23, 22, 21, 20, 19, 18, 17, 16, 15, 14, 12)
            $chosen = $preferred
            if ($usedColors.ContainsKey($preferred)) {
                foreach ($c in $palette) {
                    if (-not $usedColors.ContainsKey($c)) { $chosen = $c; break }
                }
            }
            $newCat = $categories.Add($JiraCategory)
            $newCat.Color = $chosen
            $newCat.ShortcutKey = 0
            Log-Msg INFO "Categorie '$JiraCategory' creee (couleur: $chosen)"
        }
    } catch {
        Log-Msg WARN "Impossible de creer la categorie '$JiraCategory': $_"
    }
}

function Set-JiraFlag {
    param($Mail, [string]$JiraCategory)
    try {
        if ($Mail -ne $null) {
            $cats = $Mail.Categories
            if (-not $cats) { $cats = "" }
            if (("," + $cats + ",") -notlike "*" + $JiraCategory + "*") {
                if ($cats -ne "") { $newCats = $cats + "," + $JiraCategory } else { $newCats = $JiraCategory }
                $Mail.Categories = $newCats
                $Mail.Save()
            }
        }
    } catch {
        Log-Msg WARN "Impossible d'appliquer la categorie '$JiraCategory': $_"
    }
}

# ========== GENERATION CSV ==========
function ConvertTo-CsvContent {
    param($Rows)
    $csvHeaders = @("Projet","Type de ticket","Statut","Resume","Description","Priorite","Assigne","Rapporteur","Date de reception","Composant","Etiquettes")
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(($csvHeaders -join ";") + "`r`n")
    foreach ($line in $Rows) {
        $row = @(
            $line.Projet,
            $line.Type_de_ticket,
            $line.Statut,
            $line.Resume,
            $line.Description,
            $line.Priorite,
            $line.Assigne,
            $line.Rapporteur,
            $line.Date_de_reception,
            $line.Composant,
            $line.Etiquettes
        ) | ForEach-Object {
            $v = "$_"
            if ($null -eq $_) { $v = "" }
            '"' + ($v -replace '"','""') + '"'
        }
        [void]$sb.Append(($row -join ";") + "`r`n")
    }
    return $sb.ToString()
}

function Write-CsvFile {
    param([string]$Path, [string]$Content, [int]$Retries = 5)
    $csvDir = Split-Path -Parent $Path
    if ($csvDir -and -not (Test-Path $csvDir)) { New-Item -ItemType Directory -Path $csvDir -Force | Out-Null }
    $enc = New-Object System.Text.UTF8Encoding($true)
    # UTF-8 avec BOM pour une lecture correcte des accents dans Excel FR.
    # Nouvelles tentatives si le fichier est verrouille (CSV ouvert dans Excel).
    for ($i = 1; $i -le $Retries; $i++) {
        try {
            [System.IO.File]::WriteAllText($Path, $Content, $enc)
            return $true
        } catch {
            if ($i -lt $Retries) {
                Log-Msg WARN "Ecriture CSV bloquee (tentative $i/$Retries), nouvelle tentative dans 3s: $_"
                Start-Sleep -Seconds 3
            } else {
                Log-Msg ERROR "CSV verrouille, abandon apres $Retries tentatives (fermez le fichier s'il est ouvert): $_"
            }
        }
    }
    return $false
}

# ========== IMPORT JIRA (API REST, optionnel) ==========
function Invoke-JiraImport {
    param($Rows, $Config)
    $result = @{ Success = 0; Failed = New-Object System.Collections.Generic.List[string] }
    $jira = $Config.Jira
    if (-not $jira.Enabled) { return $result }

    $token = [System.Environment]::GetEnvironmentVariable($jira.ApiTokenEnvVar)
    if (-not $jira.BaseUrl -or -not $jira.Email -or -not $token) {
        Log-Msg ERROR "Import Jira active mais configuration incomplete (BaseUrl, Email, ou variable d'environnement '$($jira.ApiTokenEnvVar)' absente)"
        $result.Failed.Add("Configuration Jira incomplete")
        return $result
    }

    try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}
    $pair = "$($jira.Email):$token"
    $auth = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))
    $headers = @{ Authorization = $auth; "Content-Type" = "application/json" }
    $apiUrl = $jira.BaseUrl.TrimEnd('/') + "/rest/api/2/issue"

    foreach ($row in $Rows) {
        $fields = [ordered]@{
            project   = @{ key = $row.Projet }
            issuetype = @{ name = $row.Type_de_ticket }
            summary   = $row.Resume
            priority  = @{ name = $row.Priorite }
        }
        if ($row.Assigne) { $fields["assignee"] = @{ name = $row.Assigne } }
        if ($row.Composant) { $fields["components"] = @(@{ name = $row.Composant }) }
        if ($row.Etiquettes) {
            $labels = @($row.Etiquettes -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
            if ($labels.Count -gt 0) { $fields["labels"] = $labels }
        }
        # Le rendu wiki de la description est gere par Jira (le champ Description en API
        # est interprete comme du texte wiki par defaut sur Jira Server/Data Center).
        $bodyJson = @{ fields = $fields } | ConvertTo-Json -Depth 5
        try {
            $resp = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $bodyJson -ContentType "application/json; charset=utf-8" -ErrorAction Stop
            Log-Msg INFO "Ticket Jira cree: $($resp.key) - $($row.Resume)"
            $result.Success++
        } catch {
            $msg = "$($row.Resume): $_"
            Log-Msg ERROR "Echec creation ticket Jira: $msg"
            $result.Failed.Add($msg)
        }
    }
    return $result
}
