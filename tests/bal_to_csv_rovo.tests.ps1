<#
.SYNOPSIS
  Tests Pester pour les fonctions pures de bal_to_csv_rovo.functions.ps1
  Executer avec:  Invoke-Pester -Path tests/bal_to_csv_rovo.tests.ps1
#>

BeforeAll {
    # $PSScriptRoot pointe vers le dossier tests/ quand le fichier de tests s'execute
    $here = $PSScriptRoot
    $sut = Join-Path (Split-Path -Parent $here) "bal_to_csv_rovo.functions.ps1"
    # Le dot-sourcing requiert $LogFile defini (utilise par Log-Msg)
    $script:LogFile = Join-Path ([System.IO.Path]::GetTempPath()) "pester_bal_to_csv_log.txt"
    $script:Silent = $true
    . $sut
}

Describe "Normalize-Name" {
    It "normalise les accents" {
        Normalize-Name "Boîte de Réception" | Should -Be "boitedereception"
    }
    It "conserve tirets et underscores" {
        Normalize-Name "DFP-UOF-IPED-DIP_2" | Should -Be "dfp-uof-iped-dip_2"
    }
    It "supprime les caracteres non alphanumeriques" {
        Normalize-Name "BAL (Paris) // 2024" | Should -Be "balparis2024"
    }
    It "retourne une chaine vide pour null" {
        Normalize-Name $null | Should -Be ""
    }
}

Describe "Classify-Mail" {
    It "detecte une anomalie" {
        Classify-Mail -subject "Anomalie connexion" -body "" | Should -Be "Anomalie"
    }
    It "detecte un bug" {
        Classify-Mail -subject "Bug d'affichage" -body "" | Should -Be "Anomalie"
    }
    It "detecte une demande d'evolution" {
        Classify-Mail -subject "Besoin d'une evolution" -body "" | Should -Be "User Story"
    }
    It "detecte une demande d'assistance" {
        Classify-Mail -subject "Besoin d'aide" -body "comment faire pour..." | Should -Be "Support"
    }
    It "retourne User Story par defaut sans mot-cle" {
        Classify-Mail -subject "Bonjour" -body "Au revoir" | Should -Be "User Story"
    }
    It "est deterministe (10 executions identiques)" {
        $results = 1..10 | ForEach-Object { Classify-Mail -subject "Besoin d'aide sur l'administration" -body "" }
        $results | Select-Object -Unique | Should -HaveCount 1
    }
}

Describe "Get-DeduplicationId" {
    It "utilise le Message-ID quand il est present" {
        $id = Get-DeduplicationId -MessageId "<abc@domain.fr>" -Subject "S" -From "F" -Body "B"
        $id | Should -Be "mid:<abc@domain.fr>"
    }
    It "genere un hash stable et compact sans Message-ID" {
        $longBody = "Corps tres long..." * 50
        $id1 = Get-DeduplicationId -MessageId "" -Subject "Sujet" -From "a@b.fr" -Body $longBody
        $id2 = Get-DeduplicationId -MessageId "" -Subject "Sujet" -From "a@b.fr" -Body $longBody
        $id1 | Should -Be $id2
        $id1.StartsWith("sha256:") | Should -Be $true
        ($id1 -replace 'sha256:','').Length | Should -Be 64
    }
    It "distingue des mails differents" {
        $id1 = Get-DeduplicationId -MessageId "" -Subject "A" -From "x" -Body "1"
        $id2 = Get-DeduplicationId -MessageId "" -Subject "B" -From "x" -Body "1"
        $id1 | Should -Not -Be $id2
    }
    It "distingue deux mails identiques par ReceivedTime (chaine de reponses)" {
        $id1 = Get-DeduplicationId -MessageId "" -Subject "A" -From "x" -Body "1" -ReceivedTime "2024-01-01 10:00:00"
        $id2 = Get-DeduplicationId -MessageId "" -Subject "A" -From "x" -Body "1" -ReceivedTime "2024-01-02 11:00:00"
        $id1 | Should -Not -Be $id2
    }
}

Describe "Import-TraceFile / Save-TraceFile" {
    BeforeEach {
        $tracePath = Join-Path ([System.IO.Path]::GetTempPath()) "trace_$([guid]::NewGuid()).txt"
    }
    AfterEach {
        Remove-Item $tracePath -Force -ErrorAction SilentlyContinue
    }

    It "ecrit au format horodate et recharge les identifiants" {
        $trace = @{ "mid:<a@b>" = $true; "sha256:abcd" = $true }
        Save-TraceFile -Path $tracePath -Trace $trace
        Test-Path $tracePath | Should -Be $true
        $reloaded = Import-TraceFile -Path $tracePath
        $reloaded.ContainsKey("mid:<a@b>") | Should -Be $true
        $reloaded.ContainsKey("sha256:abcd") | Should -Be $true
        $reloaded.Count | Should -Be 2
    }
    It "ne laisse pas de fichier temporaire derriere" {
        Save-TraceFile -Path $tracePath -Trace @{ "x" = $true }
        Test-Path "$tracePath.tmp" | Should -Be $false
    }
    It "accepte le format historique (sans horodatage)" {
        Set-Content -Path $tracePath -Value "mid:<ancien@format>" -Encoding UTF8
        $reloaded = Import-TraceFile -Path $tracePath
        $reloaded.ContainsKey("mid:<ancien@format>") | Should -Be $true
    }
    It "prune les entrees plus vieilles que la retention" {
        $old = (Get-Date).AddDays(-400).ToString("yyyy-MM-ddTHH:mm:ss")
        $recent = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
        "$($old)|mid:<vieux@b>" | Set-Content -Path $tracePath -Encoding UTF8
        Add-Content -Path $tracePath -Value "$($recent)|mid:<recent@b>" -Encoding UTF8
        $reloaded = Import-TraceFile -Path $tracePath -RetentionDays 365
        $reloaded.ContainsKey("mid:<vieux@b>") | Should -Be $false
        $reloaded.ContainsKey("mid:<recent@b>") | Should -Be $true
    }
    It "retourne une trace vide si le fichier est absent" {
        $reloaded = Import-TraceFile -Path (Join-Path ([System.IO.Path]::GetTempPath()) "inexistant_$([guid]::NewGuid()).txt")
        $reloaded.Count | Should -Be 0
    }
}

Describe "ConvertTo-CsvContent" {
    It "genere l'en-tete attendu (avec Composant et Etiquettes)" {
        $csv = ConvertTo-CsvContent -Rows @()
        $firstLine = ($csv -split "`r`n")[0]
        $firstLine | Should -Be "Projet;Type de ticket;Statut;Resume;Description;Priorite;Assigne;Rapporteur;Date de reception;Composant;Etiquettes"
    }
    It "echappe les guillemets et separateurs" {
        $rows = @(
            @{ Projet="Y6S"; Type_de_ticket="Tache"; Statut="Nouveau"; Resume='Un "resume"; special'; Description="Ligne"; Priorite="Moyenne"; Assigne="a@b.fr"; Rapporteur="c@d.fr"; Date_de_reception="2024-01-01 00:00:00"; Composant=""; Etiquettes="" }
        )
        $csv = ConvertTo-CsvContent -Rows $rows
        $dataLine = ($csv -split "`r`n")[1]
        $dataLine | Should -Match '"Un ""resume""; special"'
    }
    It "separe chaque ligne par un retour chariot" {
        $rows = @(
            @{ Projet="A"; Type_de_ticket="T"; Statut="S"; Resume="R1"; Description="D"; Priorite="P"; Assigne="a"; Rapporteur="r"; Date_de_reception="2024-01-01 00:00:00"; Composant=""; Etiquettes="" },
            @{ Projet="B"; Type_de_ticket="T"; Statut="S"; Resume="R2"; Description="D"; Priorite="P"; Assigne="a"; Rapporteur="r"; Date_de_reception="2024-01-02 00:00:00"; Composant="C"; Etiquettes="e1,e2" }
        )
        $csv = ConvertTo-CsvContent -Rows $rows
        ($csv -split "`r`n").Count | Should -Be 4
        ($csv -split "`r`n")[2] | Should -Match '"C";"e1,e2"'
    }
}

Describe "Write-CsvFile" {
    BeforeEach {
        $csvPath = Join-Path ([System.IO.Path]::GetTempPath()) "csv_$([guid]::NewGuid()).csv"
    }
    AfterEach {
        Remove-Item $csvPath -Force -ErrorAction SilentlyContinue
    }

    It "ecrit le contenu en UTF-8 avec BOM" {
        $ok = Write-CsvFile -Path $csvPath -Content "a;b`r`n" -Retries 2
        $ok | Should -Be $true
        $bytes = [System.IO.File]::ReadAllBytes($csvPath)
        $bytes[0] | Should -Be 0xEF
        $bytes[1] | Should -Be 0xBB
        $bytes[2] | Should -Be 0xBF
    }
    It "echoue proprement quand le fichier est verrouille" -Skip:($env:OS -ne "Windows_NT") {
        # Verrou exclusif comme Excel qui garde le CSV ouvert (Windows uniquement:
        # sur Linux le partage de fichier n'empeche pas l'ecriture concurrente)
        $fs = [System.IO.File]::Open($csvPath, 'Create', 'Read', [System.IO.FileShare]::None)
        try {
            $ok = Write-CsvFile -Path $csvPath -Content "x" -Retries 1
            $ok | Should -Be $false
        } finally {
            $fs.Close()
        }
    }
    It "cree le dossier parent si absent" {
        $dir = Join-Path ([System.IO.Path]::GetTempPath()) "pester_dir_$([guid]::NewGuid())"
        $path = Join-Path $dir "out.csv"
        try {
            $ok = Write-CsvFile -Path $path -Content "a" -Retries 1
            $ok | Should -Be $true
            Test-Path $path | Should -Be $true
        } finally {
            Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe "Get-MailBody" {
    It "decode les entites HTML du corps" {
        $mail = @{ HTMLBody = "<p>Bonjour&eacute; ceci est un test &amp; fin</p>" }
        # Les objets COM n'existent pas en test : on passe par une hashtable
        # Get-MailBody lit $mail.HTMLBody, $mail.Body
        $body = Get-MailBody -mail $mail
        $body | Should -Be "Bonjouré ceci est un test & fin"
    }
    It "retombe sur Body si HTMLBody absent" {
        $mail = @{ Body = "Corps brut" }
        (Get-MailBody -mail $mail) | Should -Be "Corps brut"
    }
    It "retourne une chaine vide pour null" {
        (Get-MailBody -mail $null) | Should -Be ""
    }
}

Describe "Get-Config" {
    It "charge les valeurs par defaut si le fichier est absent" {
        $cfg = Get-Config -Path (Join-Path ([System.IO.Path]::GetTempPath()) "fichier_inexistent_$([guid]::NewGuid()).json")
        $cfg.MaxFetchPerBal | Should -Be 50
        $cfg.DefaultPriority | Should -Be "Moyenne"
        $cfg.TraceRetentionDays | Should -Be 365
        $cfg.Jira.Enabled | Should -Be $false
    }
    It "lit un config.json valide" {
        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "cfg_$([guid]::NewGuid()).json"
        Set-Content -Path $tmp -Value '{"MaxFetchPerBal": 10, "DefaultPriority": "Basse"}' -Encoding UTF8
        try {
            $cfg = Get-Config -Path $tmp
            $cfg.MaxFetchPerBal | Should -Be 10
            $cfg.DefaultPriority | Should -Be "Basse"
        } finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
    It "retourne les valeurs par defaut si le JSON est invalide" {
        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "cfg_$([guid]::NewGuid()).json"
        Set-Content -Path $tmp -Value '{invalide' -Encoding UTF8
        try {
            $cfg = Get-Config -Path $tmp
            $cfg.MaxFetchPerBal | Should -Be 50
        } finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
    It "fusionne la section Jira avec les defauts" {
        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "cfg_$([guid]::NewGuid()).json"
        Set-Content -Path $tmp -Value '{"Jira": {"Enabled": true, "BaseUrl": "https://jira.example.fr"}}' -Encoding UTF8
        try {
            $cfg = Get-Config -Path $tmp
            $cfg.Jira.Enabled | Should -Be $true
            $cfg.Jira.BaseUrl | Should -Be "https://jira.example.fr"
            $cfg.Jira.ApiTokenEnvVar | Should -Be "JIRA_API_TOKEN"
        } finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe "Get-Mailboxes / Get-Projects" {
    It "retourne les BAL par defaut avec leur ProjectKey" {
        $cfg = Get-Config -Path (Join-Path ([System.IO.Path]::GetTempPath()) "inexistant_$([guid]::NewGuid()).json")
        $mbs = Get-Mailboxes -Config $cfg
        $mbs.Count | Should -Be 2
        $mbs[0].ProjectKey | Should -Be "Y6S"
        $mbs[1].ProjectKey | Should -Be "MHR"
    }
    It "retourne les projets par defaut" {
        $cfg = Get-Config -Path (Join-Path ([System.IO.Path]::GetTempPath()) "inexistant_$([guid]::NewGuid()).json")
        $projects = Get-Projects -Config $cfg
        $projects["MHR"].Composants | Should -Be "Injections"
    }
    It "retourne les projets surcharges par config.json" {
        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "cfg_$([guid]::NewGuid()).json"
        Set-Content -Path $tmp -Value '{"Projects": {"ABC": {"Key": "ABC", "Composants": "MonComposant", "Etiquettes": "e1"}}}' -Encoding UTF8
        try {
            $cfg = Get-Config -Path $tmp
            $projects = Get-Projects -Config $cfg
            $projects["ABC"].Composants | Should -Be "MonComposant"
        } finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe "Invoke-JiraImport" {
    It "ne fait rien quand Jira est desactive" {
        $cfg = [pscustomobject]@{ Jira = [pscustomobject]@{ Enabled = $false } }
        $rows = @(@{ Projet = "Y6S"; Resume = "Test" })
        $result = Invoke-JiraImport -Rows $rows -Config $cfg
        $result.Success | Should -Be 0
        $result.Failed.Count | Should -Be 0
    }
    It "signale une configuration incomplete sans creer de ticket" {
        $cfg = [pscustomobject]@{ Jira = [pscustomobject]@{ Enabled = $true; BaseUrl = ""; Email = ""; ApiTokenEnvVar = "JIRA_API_TOKEN" } }
        $rows = @(@{ Projet = "Y6S"; Resume = "Test" })
        $result = Invoke-JiraImport -Rows $rows -Config $cfg
        $result.Success | Should -Be 0
        $result.Failed.Count | Should -Be 1
    }
}
