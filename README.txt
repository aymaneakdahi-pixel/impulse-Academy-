====================================================
  IMPULSE ACADEMY – Générateur d'attestations
  Version 1.0 – Outil interne
====================================================

PRÉREQUIS
---------
  - Python 3.8+
  - LibreOffice (pour conversion PDF)

INSTALLATION (première fois uniquement)
----------------------------------------
  pip install flask python-pptx openpyxl

LANCEMENT
---------
  python3 app.py

  Puis ouvrir dans le navigateur :
  http://localhost:5000

FONCTIONNALITÉS
---------------
  ✓ Formulaire de session (type formation, fonction, dates)
  ✓ Import Excel / CSV (colonnes : Nom_Prenom, RPPS, Fonction)
  ✓ Ajout manuel de participants
  ✓ Modification / suppression des participants
  ✓ Règles RPPS automatiques :
      - Infirmier(e) → ligne RPPS supprimée
      - Médecin → RPPS inclus
  ✓ Génération PDF + PPTX par participant
  ✓ Export ZIP téléchargeable

STRUCTURE
---------
  app.py                 ← Backend Flask
  modele_original.pptx   ← Modèle PowerPoint (source de vérité)
  templates/index.html   ← Interface web
  generated/             ← Attestations générées (auto-créé)
  uploads/               ← Fichiers Excel temporaires (auto-créé)

MODÈLE PPTX
-----------
  Le modèle peut être remplacé via l'interface.
  Variables détectées automatiquement :
    - "Valérie SAULZE"    → nom du participant
    - "10106538399"       → numéro RPPS
    - "18/12/2025"        → date de délivrance
    - "15 au 18/12/2025"  → dates de formation
    - "Médecin du travail"→ fonction

====================================================
