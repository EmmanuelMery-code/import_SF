# Import Salesforce Multi-Onglets

---

## English

### Purpose

Django web application to **import multi-sheet Excel data into Salesforce** with automatic ID replication across related sheets.

Designed for complex imports where multiple objects are linked (e.g. Accounts → Contacts → Opportunities). The tool manages import order, column mappings and propagation of Salesforce IDs (AccountId, ParentId, etc.) without manual intervention.

**Key features:**
- No Connected App required: Session ID connection
- Excel file with `_Config_` sheet to define order, mappings and pivot columns
- Explorer interface: tree view, graphic view of mappings, sheet management (add, rename, delete)
- Sequential import with automatic ID replication

### Installation

```bash
pip install -r requirements.txt
python manage.py migrate
python manage.py runserver
```

Open http://localhost:8000/

### Salesforce connection

1. Click "Salesforce Login"
2. Retrieve your Session ID via the bookmarklet (on a Salesforce page) or manual entry
3. Paste the instance URL and Session ID

*Note: Lightning sessions do not work with the API. Use a Visualforce page in preview mode.*

### Documentation

- **`document/Manuel_Utilisateur_EN.rtf`** — User manual (English)
- **`document/Manuel_Utilisateur_FR.rtf`** — Manuel utilisateur (français)

### Excel file structure

The `_Config_` sheet defines:

| Section | Role |
|---------|------|
| OrdreOnglets | Import order, target Salesforce object, ID column |
| Correspondances | Links between sheets (e.g. RefCompte → AccountId) |
| ColonnesPivot | Columns not imported, used for mapping |

### Project structure

```
├── document/              # Documentation (FR/EN manuals)
├── manage.py
├── sf_import_project/     # Django project
├── sf_import/             # Import application
├── excel_import_config.py # Excel reader + config
├── salesforce_importer.py # Salesforce import logic
└── create_example_excel.py
```

---

## Français

### Objectif

Application web Django permettant d'**importer des données Excel multi-onglets vers Salesforce** avec réplication automatique des IDs entre onglets liés.

Conçu pour les imports complexes où plusieurs objets sont reliés (ex : Comptes → Contacts → Opportunités). L'outil gère l'ordre d'import, les correspondances entre colonnes et la propagation des IDs Salesforce (AccountId, ParentId, etc.) sans intervention manuelle.

**Points clés :**
- Pas de Connected App requis : connexion par Session ID
- Fichier Excel avec onglet `_Config_` pour définir ordre, correspondances et colonnes pivot
- Interface Explorer : arborescence, vue graphique des correspondances, gestion des onglets (ajout, renommage, suppression)
- Import séquentiel avec réplication automatique des IDs

### Installation

```bash
pip install -r requirements.txt
python manage.py migrate
python manage.py runserver
```

Ouvrez http://localhost:8000/

### Connexion Salesforce

1. Cliquez sur « Connexion Salesforce »
2. Récupérez votre Session ID via le bookmarklet (sur une page Salesforce) ou saisie manuelle
3. Collez l'URL d'instance et le Session ID

*Note : Les sessions Lightning ne fonctionnent pas avec l'API. Utilisez une page Visualforce en preview.*

### Documentation

- **`document/Manuel_Utilisateur_EN.rtf`** — User manual (English)
- **`document/Manuel_Utilisateur_FR.rtf`** — Manuel utilisateur (français)

### Structure du fichier Excel

L'onglet `_Config_` définit :

| Section | Rôle |
|---------|------|
| OrdreOnglets | Ordre d'import, objet Salesforce cible, colonne ID |
| Correspondances | Liens entre onglets (ex : RefCompte → AccountId) |
| ColonnesPivot | Colonnes non importées, utilisées pour la correspondance |

### Structure du projet

```
├── document/              # Documentation (manuels FR/EN)
├── manage.py
├── sf_import_project/     # Projet Django
├── sf_import/             # Application d'import
├── excel_import_config.py # Lecture Excel + config
├── salesforce_importer.py # Logique d'import Salesforce
└── create_example_excel.py
```
