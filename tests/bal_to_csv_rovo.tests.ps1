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
    It "detecte un bug avec priorite associee" {
        $t = Classify-Mail -subject "Bug d'affichage" -body ""
        $t | Should -Be "Anomalie"
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
}

Describe "ConvertTo-CsvContent" {
    It "genere l'en-tete attendu" {
        $csv = ConvertTo-CsvContent -Rows @()
        $firstLine = ($csv -split "`r`n")[0]
        $firstLine | Should -Be "Projet;Type de ticket;Statut;Resume;Description;Priorite;Assigne;Rapporteur;Date de reception"
    }
    It "echappe les guillemets et separateurs" {
        $rows = @(
            @{ Projet="Y6S"; Type_de_ticket="Tache"; Statut="Nouveau"; Resume='Un "resume"; special'; Description="Ligne"; Priorite="Moyenne"; Assigne="a@b.fr"; Rapporteur="c@d.fr"; Date_de_reception="2024-01-01 00:00:00" }
        )
        $csv = ConvertTo-CsvContent -Rows $rows
        $dataLine = ($csv -split "`r`n")[1]
        $dataLine | Should -Match '"Un ""resume""; special"'
    }
    It "separe chaque ligne par un retour chariot" {
        $rows = @(
            @{ Projet="A"; Type_de_ticket="T"; Statut="S"; Resume="R1"; Description="D"; Priorite="P"; Assigne="a"; Rapporteur="r"; Date_de_reception="2024-01-01 00:00:00" },
            @{ Projet="B"; Type_de_ticket="T"; Statut="S"; Resume="R2"; Description="D"; Priorite="P"; Assigne="a"; Rapporteur="r"; Date_de_reception="2024-01-02 00:00:00" }
        )
        $csv = ConvertTo-CsvContent -Rows $rows
        ($csv -split "`r`n").Count | Should -Be 4
    }
}

Describe "Get-Config" {
    It "charge les valeurs par defaut si le fichier est absent" {
        $cfg = Get-Config -Path (Join-Path ([System.IO.Path]::GetTempPath()) "fichier_inexistent_$([guid]::NewGuid()).json")
        $cfg.MaxFetchPerBal | Should -Be 50
        $cfg.DefaultPriority | Should -Be "Moyenne"
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
}
