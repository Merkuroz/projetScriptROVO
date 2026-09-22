# projetScriptROVO

Script PowerShell qui parcourt les mails des BAL Outlook (via COM) et génère un fichier CSV importable dans Jira (import de masse CSV), avec création directe des tickets via l'API REST Jira en option.

## Utilisation

```powershell
.\bal_to_csv_rovo.ps1            # Traitement normal
.\bal_to_csv_rovo.ps1 -DryRun    # Simulation : aucun CSV, aucune catégorie, aucune trace écrite
.\bal_to_csv_rovo.ps1 -Force     # Retraite tous les mails, même déjà importés
.\bal_to_csv_rovo.ps1 -Silent    # Pas de sortie console (log fichier conservé)
```

Le script :
1. Se connecte à Outlook (COM) et localise chaque BAL (boîte de réception).
2. Filtre les mails **côté Outlook** (`Restrict` DASL : `IPM.Note` + non catégorisés + âge max optionnel) — les mails déjà traités ne consomment plus le quota, aucun mail n'est perdu dans le backlog.
3. Dédoublonne via le `Message-ID` (fallback : hash SHA-256 du sujet/expéditeur/date/corps), trace horodatée dans `processed_msgids.txt` avec rétention.
4. Classifie chaque mail (Anomalie / User Story / Support / Tâche) par mots-clés ; une anomalie reçoit la priorité « Haute ».
5. Écrit `bal_to_jira_<date>.csv` (UTF-8 BOM, `;`, guillemets échappés, colonnes Composant et Étiquettes) — toujours créé, même vide, avec au minimum l'en-tête. Réessaie plusieurs fois si le fichier est ouvert dans Excel.
6. Crée les tickets directement dans Jira via l'API REST si activé dans la config (voir plus bas).
7. Catégorise les mails dans Outlook (« a traiter dans Jira ») **seulement après écriture réussie** : un échec CSV/Jira ne fait perdre aucun mail.

Le script retourne un **code de sortie** (0 = OK, 1 = erreur), exploitable par le Planificateur de tâches.

## Lien vers le mail dans Jira

Chaque description se termine par un lien cliquable (format wiki Jira) qui ouvre directement le mail dans Outlook, plus le `Message-ID` en clair pour retrouver le mail même s'il est déplacé (l'EntryID change avec le dossier) :

```
Lien vers le mail: [Ouvrir le mail dans Outlook|outlook:<EntryID>]
Message-ID: <abc@domain.fr>
```

⚠️ Le protocole `outlook:` fonctionne avec **Outlook classique** (desktop). Il peut être ignoré par le « nouveau Outlook » / OWA.

## Configuration

Tous les paramètres sont dans `config.json` (valeurs par défaut intégrées au script si le fichier est absent ou invalide) :

| Clé | Rôle |
|---|---|
| `JiraCategory` | Catégorie Outlook appliquée aux mails traités (sert aussi de filtre d'exclusion) |
| `DefaultAssignee` / `DefaultPriority` / `DefaultPriorityBug` | Assigné et priorités (Anomalie → `DefaultPriorityBug`) |
| `MaxFetchPerBal` | Nombre max de mails traités par BAL et par exécution |
| `MaxMailAgeDays` | Âge max des mails (0 = pas de limite) — poussé dans le filtre Outlook |
| `RetentionDays` | Rétention des logs `log_bal_to_csv_*.txt` |
| `TraceRetentionDays` | Rétention des entrées de `processed_msgids.txt` |
| `CsvWriteRetries` | Tentatives d'écriture du CSV si verrouillé (Excel ouvert) |
| `Silent` | Mode silencieux par défaut (désactivable via le paramètre) |
| `Mailboxes` | BAL à traiter avec `Address`, `Name`, `ProjectKey` |
| `Projects` | Clés projets Jira, composants, étiquettes (colonnes du CSV) |
| `Jira.Enabled` | Active la création directe des tickets via API REST |
| `Jira.BaseUrl` / `Jira.Email` / `Jira.ApiTokenEnvVar` | Connexion Jira (le token API est lu dans la variable d'environnement, jamais dans le fichier) |

## Import Jira direct (optionnel)

Pour créer les tickets sans passer par l'import CSV :

```powershell
# 1. Creer un token API dans Jira (Compte > Securite > API tokens)
# 2. Definir la variable d'environnement (jamais dans config.json)
[Environment]::SetEnvironmentVariable("JIRA_API_TOKEN", "mon_token", "User")
# 3. Activer "Jira": {"Enabled": true, ...} dans config.json
```

Le token n'apparaît ni dans le dépôt, ni dans les logs.

## Tâche planifiée

`register_task.ps1` crée l'entrée Planificateur Windows (exécution en arrière-plan, `-Silent`) :

```powershell
.\register_task.ps1 -IntervalMinutes 15   # toutes les 15 minutes
.\register_task.ps1 -Unregister           # supprime la tâche
```

## Structure

- `bal_to_csv_rovo.ps1` — script principal (orchestration)
- `bal_to_csv_rovo.functions.ps1` — fonctions (dot-sourcé par le script principal et les tests)
- `config.json` — configuration
- `register_task.ps1` — enregistrement de la tâche planifiée Windows
- `tests/bal_to_csv_rovo.tests.ps1` — tests Pester (36 tests sans Outlook)
- `.github/workflows/tests.yml` — CI GitHub Actions (Pester sur `windows-latest`)

## Tests

```powershell
Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser
Invoke-Pester -Path tests/bal_to_csv_rovo.tests.ps1
```

La CI GitHub Actions exécute ces tests à chaque push/PR.

## Fichiers générés (ignorés par git)

- `bal_to_jira_<date>.csv` — CSV pour Jira
- `log_bal_to_csv_<date>.txt` — journal d'exécution
- `processed_msgids.txt` — trace de dédoublonnage (format `date|id`)
- `bal_to_csv.lock` — verrou d'exécution unique
