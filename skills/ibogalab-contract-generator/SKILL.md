---
name: Iboga Lab - Contract Generator
description: Génère des contrats professionnels pour Iboga Lab à partir de Markdown en HTML, PDF et DOCX.
---

# Iboga Lab - Générateur de Contrats

Ce skill permet de générer des contrats professionnels aux formats HTML, PDF, et DOCX à partir d'un fichier Markdown source.
Il utilise l'identité visuelle de la marque Iboga Lab (couleurs, logo).

## Utilisation

1.  Demandez à l'utilisateur de fournir les éléments du contrat (parties, objet, durée, rémunération, articles).
2.  Générez le fichier Markdown du contrat dans le dossier `devis-factures/` en utilisant le format de front-matter ci-dessous.
3.  Exécutez le script `skills/ibogalab-contract-generator/scripts/sync_contract.py <chemin_vers_le_fichier.md>`.

## Format du Fichier Markdown

```yaml
---
reference: "CC-001/08/2026"
date: "31 Août 2026"
title: "Contrat de Prestation de Services de Consultation"
subtitle: "Accompagnement Multi-Canal"
location: "Gamba"
client:
  name: "M. GUENNOLE MAKELA NZAOU"
  company: "MNG & SERVICES"
  role: "Directeur Général — MNG & SERVICES"
  address: "BP 1320 Gamba, Quartier Plaine 3, République Gabonaise"
  nif: "202502009970-T"
  rccm: "RG-POG 01-2025-A10-00606"
  phone: "077 37 11 50 / 062 67 99 20"
  email: "mngestservices@gmail.com"
consultant:
  name: "M. Jesse Ogoula"
  company: "Iboga Lab"
  role: "Consultant Indépendant — Expert Produit & IA"
  address: "Port-Gentil, Ogooué-Maritime, République Gabonaise"
  phone: "+241 066 19 57 86"
  email: "adirignoogoula@gmail.com"
---

Ci-après désignés individuellement « la Partie » et collectivement « les Parties ».

# Article 1 — Objet du contrat
Le présent contrat a pour objet...

# Article 2 — Nature des prestations
...
```
