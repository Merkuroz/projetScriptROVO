# projetScriptROVO

Script PowerShell qui parcourt les mails des BAL Outlook (via COM) et génère un fichier CSV importable dans Jira (import de masse CSV).

## Utilisation

```powershell
.\bal_to_csv_rovo.ps1            # Traitement normal
.\bal_to_csv_rovo.ps1 -DryRun    # Simulation : aucun CSV, aucune catégorie, aucune trace écrite
.\bal_to_csv_rovo.ps1 -Force     # Retraite tous les mails, même déjà importés
```

Le script :
1. Se connecte à Outlook (COM) et localise chaque BAL (boîte de réception).
2. Filtre les mails côté Outlook (`Restrict` sur `IPM.Note`) — bien plus rapide qu'un parcours complet.
3. Dédoublonne via le `Message-ID` (fallback : hash SHA-256 du sujet/expéditeur/corps), trace dans `processed_msgids.txt`.
4. Classifie chaque mail (Anomalie / User Story / Support / Tâche) par mots-clés ; une anomalie reçoit la priorité « Haute ».
5. Écrit `bal_to_jira_<date>.csv` (UTF-8 BOM, `;`, guillemets échappés) — toujours créé, même vide, avec au minimum l'en-tête.
6. Catégorise les mails traités dans Outlook (« a traiter dans Jira ») pour traçabilité.

## Lien vers le mail dans Jira

Chaque description se termine par un lien cliquable (format wiki Jira) qui ouvre directement le mail dans Outlook :

```
Lien vers le mail: [Ouvrir le mail dans Outlook|outlook:<EntryID>]
```

⚠️ Le protocole `outlook:` fonctionne avec **Outlook classique** (desktop). Il peut être ignoré par le « nouveau Outlook » / OWA.

## Configuration

Tous les paramètres sont dans `config.json` (valeurs par défaut intégrées au script si le fichier est absent ou invalide) :

| Clé | Rôle |
|---|---|
| `JiraCategory` | Catégorie Outlook appliquée aux mails traités |
| `DefaultAssignee` / `DefaultPriority` / `DefaultPriorityBug` | Assigné et priorités (Anomalie → `DefaultPriorityBug`) |
| `MaxFetchPerBal` | Nombre max de mails traités par BAL et par exécution |
| `MaxMailAgeDays` | Âge max des mails (0 = pas de limite) — pousse le filtre dans Outlook via `Restrict` |
| `RetentionDays` | Rétention des logs `log_bal_to_csv_*.txt` (supprimés au-delà) |
| `Mailboxes` | BAL à traiter avec `Address`, `Name`, `ProjectKey` |
| `Projects` | Clés projets Jira, préfixes, composants, étiquettes |

## Structure

- `bal_to_csv_rovo.ps1` — script principal (orchestration)
- `bal_to_csv_rovo.functions.ps1` — fonctions (dot-sourcé par le script principal et les tests)
- `config.json` — configuration
- `tests/bal_to_csv_rovo.tests.ps1` — tests Pester des fonctions pures (sans Outlook)

## Tests

```powershell
Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser
Invoke-Pester -Path tests/bal_to_csv_rovo.tests.ps1
```

## Fichiers générés (ignorés par git)

- `bal_to_jira_<date>.csv` — CSV pour Jira
- `log_bal_to_csv_<date>.txt` — journal d'exécution
- `processed_msgids.txt` — trace de dédoublonnage
- `bal_to_csv.lock` — verrou d'exécution unique
