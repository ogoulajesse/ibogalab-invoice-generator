---
# Métadonnées du document Iboga Lab
type: DEVIS
number: auto           # auto = généré et incrémenté automatiquement par le script
date: "2026-07-14"
validity: "30 jours"

# Informations Client
client:
  name: "M. Nazaire TOUGOUAVANA MOUKAMBI PANGOU"
  company: "GabSource"
  address: "Libreville, Gabon"

# Détails du Projet
project:
  name: "MVP Marketplace B2B (Option 1)"
  description: "Architecture Supabase / Indie Hacker"
  duration: "12 semaines"

# Calculs de prix
auto_calculate: true   # true = calcule automatiquement le total HT et TTC
currency: "FCFA"       # Devise (ex: FCFA, EUR, USD)
tax_rate: 0.0          # Taux de TVA (0.0 = Exonéré ou inclus, 0.18 = 18%)

# Coûts Récurrents Optionnels (S'affichent sous le tableau principal)
recurring_costs:
  - designation: "Infrastructure Cloud (Supabase/Vercel)"
    price: "61 000 FCFA"
    period: "mois"
  - designation: "Maintenance & Support technique"
    price: "250 000 FCFA"
    period: "mois"

# Conditions de Paiement
terms:
  - "Acompte de 40% exigible à la signature du devis pour validation de la commande."
  - "Un second versement de 30% à la validation des maquettes / mi-parcours."
  - "Solde de 30% à la livraison et mise en ligne (avant remise des codes sources)."
  - "Les frais d'infrastructure mensuels et de maintenance débutent 30 jours après la mise en production."
---

# Devis

| Désignation & Détails | Unité | Qté | P.U. HT |
| :--- | :---: | :---: | :---: |
| **Architecture & Design UI/UX**<br>Maquettes Figma, parcours utilisateurs, schéma de données. | forfait | 1 | 650000 |
| **Développement Frontend**<br>Landing page, tableaux de bord (Entreprise/Admin), formulaires d'onboarding (React/Tailwind). | forfait | 1 | 1800000 |
| **Développement Backend (Supabase)**<br>Authentification, rôles utilisateurs, gestion base de données, stockage sécurisé. | forfait | 1 | 1200000 |
| **Intégrations & Matching**<br>Algorithme de matching, paiements (Stripe + Mobile Money), emails transactionnels. | forfait | 1 | 1000000 |
| **Tests & Déploiement**<br>Tests unitaires, assurance qualité (QA) et mise en production. | forfait | 1 | 800000 |
