---
name: ibogalab-invoice-generator
description: Allows generating and synchronizing quotes (devis) and invoices (factures) for Iboga Lab in Markdown, HTML, PDF, and DOCX formats. Manages document sequence numbering automatically.
---

# Iboga Lab Invoice & Quote Generator Skill

This skill allows the agent to draft, edit, and compile official business documents (Devis and Factures) for **Iboga Lab** while maintaining strict visual branding (Forest Green `#103824`, Sage Green `#669D69`) and layout structures.

## 1. Document Structure

All invoices and quotes are authored in Markdown with a YAML front-matter header. 

### Structure du Front-matter :
```yaml
---
type: DEVIS            # ou FACTURE
number: auto           # Le script remplacera 'auto' par le numéro suivant (ex: IBGL-2026-D01)
date: "2026-07-14"
validity: "30 jours"   # Uniquement pour DEVIS
due_date: "2026-08-14" # Uniquement pour FACTURE

client:
  name: "Nom du Client"
  company: "Nom de l'Entreprise"
  address: "Adresse, Ville, Pays"

project:
  name: "Nom du Projet"
  description: "Description de l'architecture ou techno"
  duration: "12 semaines" # Optionnel

auto_calculate: true   # Remplir automatiquement le Montant HT et les totaux
currency: "FCFA"       # Devise (FCFA par défaut)
tax_rate: 0.0          # Taux de TVA (ex: 0.18 pour 18%, 0.0 pour exonéré)

# Coûts récurrents (optionnels)
recurring_costs:
  - designation: "Hébergement Cloud"
    price: "60 000 FCFA"
    period: "mois"

# Conditions contractuelles
terms:
  - "Condition 1..."
  - "Condition 2..."
---
```

### Table des prestations :
Rédigez les colonnes `Désignation & Détails`, `Unité`, `Qté`, et `P.U. HT`. Le script calculera le `Montant HT` de chaque ligne et insérera la colonne correspondante.
```markdown
# Devis

| Désignation & Détails | Unité | Qté | P.U. HT |
| :--- | :---: | :---: | :---: |
| **Prestation 1**<br>Description détaillée de la prestation. | forfait | 1 | 1500000 |
```

## 2. Compilation et Synchronisation

Après avoir rédigé ou modifié le fichier Markdown, exécutez le script de synchronisation Python. Il va automatiquement mettre à jour le Markdown (si `number: auto` est présent), générer le fichier HTML, compiler le PDF à l'aide de Google Chrome ou Microsoft Edge en mode headless, et créer le fichier Word DOCX.

### Commande simple :
```bash
python skills/ibogalab-invoice-generator/scripts/sync.py path/to/document.md
```

### Mode Veille (Watch) :
Pour synchroniser automatiquement les fichiers HTML, PDF et DOCX à chaque fois que vous sauvegardez le fichier Markdown :
```bash
python skills/ibogalab-invoice-generator/scripts/sync.py --watch path/to/document.md
```

## 3. Gestion de l'état des numéros
Les numéros de devis et factures sont conservés localement dans `~/.ibogalab-invoice-settings.json`.
Le script se charge d'incrémenter l'index approprié de façon persistante et d'écrire la valeur directement dans le front-matter du fichier Markdown source lors de sa première compilation pour éviter les doublons.

## 4. Organisation des Dossiers (Livrables Locaux)
Pour préserver la propreté du dépôt Git et éviter toute fuite de données confidentielles :
1. **Dossier de stockage** : Tous les fichiers Markdown source, ainsi que leurs versions HTML, PDF et DOCX générées, doivent être impérativement regroupés dans un dossier nommé **`devis-factures/`** à la racine de votre espace de travail.
2. **Création automatique** : Lors de la première génération de devis ou de facture, l'agent ou l'utilisateur doit créer ce dossier `devis-factures/` s'il n'existe pas déjà.
3. **Sécurité Git** : Ce dossier est configuré pour être exclu du dépôt distant (via le fichier `.gitignore`), assurant que les données commerciales restent en local sur la machine de l'utilisateur.

