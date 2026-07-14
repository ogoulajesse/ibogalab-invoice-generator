# Iboga Lab - Générateur de Devis & Factures (Skill Multi-Outils)

Ce dépôt contient un **Skill d'IA** hautement portable pour concevoir, gérer et compiler des devis ("Devis") et factures ("Factures") professionnels pour **Iboga Lab** (https://www.ibogalab.tech/).

Le principe de fonctionnement repose sur le paradigme **Single Source of Truth** : vous modifiez un document au format simple **Markdown** (`.md`), et un script de synchronisation compile automatiquement ce fichier en **HTML**, **PDF** (via Chrome/Edge headless) et **Word (.docx)** (via `python-docx` natif). Tous les fichiers restent parfaitement reliés et mis à jour.

---

## 🚀 Fonctionnalités
- **Compilation Multi-Format** : Génère automatiquement HTML, PDF et DOCX à partir d'un seul fichier Markdown source.
- **Respect de la Charte Graphique** : Rendu premium aux couleurs d'Iboga Lab (Vert Forêt `#103824`, Vert Tendre `#669D69`) avec typographie Montserrat et Poppins.
- **Calculs Automatisés** : Calcule automatiquement les montants par ligne (Quantité × P.U.), le total HT, la TVA et le Total TTC.
- **Suivi Persistant des Numéros** : Conserve l'état des numéros de factures et devis dans un fichier local de configuration (`~/.ibogalab-invoice-settings.json`) pour éviter les doublons et incrémenter de façon transparente.
- **Mode Veille (Watch Mode)** : Surveille les modifications du fichier Markdown en temps réel et régénère instantanément les formats HTML, PDF et Word à chaque sauvegarde.

---

## 🛠️ Prérequis
1. **Python 3.8+** installé sur votre machine.
2. **Google Chrome** ou **Microsoft Edge** installé pour l'export PDF (détecté automatiquement).
3. Les packages Python `markdown`, `python-docx` et `pyyaml` (installés automatiquement au premier lancement si manquants).

---

## 📦 Instructions d'Installation par Outil

Le skill est conçu pour fonctionner de manière transparente avec tous les assistants de codage IA modernes :

### 1. Google Antigravity / Gravitique (Skill Natif)
Les compétences ("skills") d'Antigravity sont des dossiers contenant des instructions système et des scripts utilitaires.

#### Option A : Installation sans téléchargement (Ligne de commande unique)
Vous pouvez installer directement le skill sans cloner le dépôt grâce aux commandes suivantes :
- **Sur Windows (PowerShell)** :
  ```powershell
  iwr -useb https://raw.githubusercontent.com/ogoulajesse/ibogalab-invoice-generator/main/install.ps1 | iex
  ```
- **Sur macOS / Linux (Terminal)** :
  ```bash
  curl -fsSL https://raw.githubusercontent.com/ogoulajesse/ibogalab-invoice-generator/main/install.sh | bash
  ```

#### Option B : Installation locale (Après clonage du dépôt)
Si vous préférez cloner le projet manuellement et l'installer :
- **Sur Windows (PowerShell)** :
  ```powershell
  Set-ExecutionPolicy Bypass -Scope Process -Force
  .\install.ps1
  ```
- **Sur macOS / Linux (Terminal)** :
  ```bash
  chmod +x install.sh
  ./install.sh
  ```
*Cela copiera ou téléchargera le dossier `skills/ibogalab-invoice-generator` dans votre répertoire global de configuration `~/.gemini/config/skills/`.*

### 2. CodeX / Cursor
Cursor utilise un fichier `.cursorrules` à la racine pour guider le comportement du LLM.
- **Installation** : Copiez le fichier `.cursorrules` fourni à la racine de votre projet.
- **Usage** : Lorsque vous demandez à Cursor de rédiger un devis, il appliquera les modèles YAML et vous guidera pour lancer `sync.py` ou démarrer le script en arrière-plan.

### 3. GitHub Copilot
Copilot utilise les fichiers d'instructions du dépôt `.github/copilot-instructions.md` pour personnaliser ses réponses.
- **Installation** : Copiez le dossier `.github/` dans votre projet.
- **Usage** : Copilot comprendra automatiquement le flux de génération de factures et respectera la charte graphique et la numérotation d'Iboga Lab lors de l'aide rédactionnelle.

### 4. ClaudeCode (VS Code / CLI)
ClaudeCode exploite les règles d'instructions personnalisées et s'intègre facilement avec les configurations système globales ou locales.
- **Installation** : En installant le skill dans `~/.gemini/config/skills/`, ClaudeCode en prendra connaissance lors de l'analyse du projet. Pour VS Code, vous pouvez utiliser le fichier de configuration de tâches VS Code (`.vscode/tasks.json`) fourni ci-dessous :
  ```json
  {
    "version": "2.0.0",
    "tasks": [
      {
        "label": "Iboga Lab Sync Document",
        "type": "shell",
        "command": "python",
        "args": [
          "skills/ibogalab-invoice-generator/scripts/sync.py",
          "${file}"
        ],
        "group": "build",
        "problemMatcher": []
      }
    ]
  }
  ```

---

## 🔄 Gestion des Mises à jour du Skill

Lorsque le skill est mis à jour sur GitHub (correctifs, améliorations des styles CSS, nouvelles fonctionnalités) :

### Si vous avez installé le skill via la ligne de commande unique (Option A) :
Il vous suffit de **relancer la même ligne de commande**. Le script d'installation détectera le répertoire existant, écrasera les anciens fichiers système par les dernières versions en provenance directe de GitHub, et conservera votre fichier de numérotation local intact (`~/.ibogalab-invoice-settings.json`).

### Si vous avez cloné le dépôt localement (Option B) :
1. Rendez-vous dans le dossier où vous avez cloné le dépôt.
2. Exécutez `git pull` pour récupérer les dernières modifications de GitHub.
3. Relancez le script d'installation local (`.\install.ps1` ou `./install.sh`) pour écraser les fichiers installés dans le répertoire Gemini par les nouveaux fichiers mis à jour.


---

## 📝 Guide d'Utilisation

### Étape 1 : Créer le fichier Markdown
Copiez le modèle depuis `skills/ibogalab-invoice-generator/examples/devis_template.md` (ou `facture_template.md`) vers votre dossier de projet.

Exemple de contenu (`mon_devis.md`) :
```yaml
---
type: DEVIS
number: auto           # Le script incrémentera le compteur et remplacera 'auto'
date: "2026-07-14"
validity: "30 jours"
client:
  name: "M. Nazaire"
  company: "GabSource"
  address: "Libreville, Gabon"
project:
  name: "MVP B2B"
  description: "React/Supabase"
auto_calculate: true
tax_rate: 0.0          # 0% TVA (Exonéré)
terms:
  - "Acompte de 40% à la commande."
  - "Solde de 60% à la livraison."
---

# Devis

| Désignation & Détails | Unité | Qté | P.U. HT |
| :--- | :---: | :---: | :---: |
| **Architecture & Design UI/UX**<br>Maquettes Figma, parcours utilisateurs, schéma de données. | forfait | 1 | 650000 |
| **Développement Frontend**<br>Landing page et tableaux de bord. | forfait | 1 | 1800000 |
```

### Étape 2 : Lancer la compilation ou la surveillance
Exécutez le script pour générer les fichiers HTML, PDF et DOCX :

* **Rendu unique** :
  ```bash
  python skills/ibogalab-invoice-generator/scripts/sync.py mon_devis.md
  ```
  *Le fichier `mon_devis.md` sera modifié pour remplacer `number: auto` par le numéro généré (ex: `IBGL-2026-D01`), et les fichiers `mon_devis.html`, `mon_devis.pdf` et `mon_devis.docx` seront créés dans le même dossier.*

* **Rendu en tâche de fond (Watch mode)** :
  ```bash
  python skills/ibogalab-invoice-generator/scripts/sync.py --watch mon_devis.md
  ```
  *Le script surveille `mon_devis.md`. À chaque sauvegarde, il recalcule les totaux et met à jour l'HTML, le PDF et le Word en moins d'une seconde.*

---

## 🗃️ Gestion de la numérotation (Settings JSON)
Les compteurs de séquence sont stockés dans `~/.ibogalab-invoice-settings.json` :
```json
{
  "counters": {
    "IBGL-2026-D": 2,
    "IBGL-2026-F": 1
  }
}
```
Vous pouvez ajuster directement ce fichier pour modifier la valeur du prochain numéro de facture ou de devis généré.
- Si vous définissez un préfixe personnalisé dans votre front-matter (ex: `prefix: "IBGL-2026-GS"`), le script l'enregistrera automatiquement comme un nouveau compteur et gérera sa propre séquence de numérotation.
