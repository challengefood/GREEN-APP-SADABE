# GREEN'APP — SADABE

**GREEN'APP** est un prototype open source pour centraliser la planification mensuelle, l’analyse automatique des documents d’activités, l’affectation des responsables, les demandes de budget et les demandes de congé de SADABE.

## Fonctions principales

- Connexion par compte utilisateur.
- Création de compte par chaque membre, avec validation par l’administrateur.
- Tableau de bord filtré par mois, projet, partenaire/bailleur, personne et statut.
- Planification du mois avec activités à réaliser, urgence, retard, responsables, membres, équipes et tâches détaillées.
- Import intelligent Excel, CSV et Word : extraction automatique du titre, description, dates, projet, partenaire, responsable, budget, urgence et points à compléter.
- Ajout manuel d’activités non présentes dans les documents.
- Gestion des projets : SOS Lemurs, Darwin Initiatives, Seacology, Rainforest Trust, plus ajout libre.
- Gestion des partenaires/bailleurs : TGBS (MBG), MfM, UWE, Regen, UNI, ENS, plus ajout libre.
- Gestion des membres SADABE et de leurs postes.
- Gestion des équipes responsables.
- Demande de budget avec canevas financier Excel rempli automatiquement et calculs automatiques.
- Demande de congé avec canevas Word prêt à transmettre aux RH.
- Export Excel du planning filtré.
- Sauvegarde de la base SQLite.

## Installation locale

1. Dézipper le dossier `GREEN_APP`.
2. Ouvrir un terminal dans le dossier.
3. Installer les dépendances :

```bash
pip install -r requirements.txt
```

4. Lancer l’application :

```bash
streamlit run app.py
```

## Connexion initiale

Compte administrateur par défaut :

```text
Email : admin@sadabe.org
Mot de passe : admin123
```

Après installation, il est recommandé de modifier le compte administrateur.

## Déploiement sur GitHub + Streamlit Cloud

1. Créer un dépôt GitHub public ou privé.
2. Envoyer tout le contenu du dossier `GREEN_APP` dans le dépôt.
3. Aller sur Streamlit Community Cloud.
4. Créer une nouvelle application.
5. Choisir :

```text
Repository : votre dépôt GitHub
Branch : main
Main file path : app.py
```

6. Cliquer sur **Deploy**.

## Utilisation recommandée

### 1. Préparer les comptes

- L’administrateur se connecte.
- Les membres créent leurs comptes.
- L’administrateur valide les comptes dans **Administration**.

### 2. Ajouter les bases

- Aller dans **Équipe & responsables** pour ajouter les membres SADABE et leurs postes.
- Aller dans **Projets & partenaires** pour compléter les projets, partenaires et bailleurs.

### 3. Importer les documents

- Aller dans **Ajouter / Importer activités**.
- Importer un fichier Excel, CSV ou Word.
- GREEN'APP analyse automatiquement le document et signale les champs manquants.
- Vérifier les activités détectées, puis enregistrer.

### 4. Planifier le mois

- Aller dans **Planification du mois**.
- Filtrer par mois.
- Ajouter des activités non intégrées dans les documents.
- Affecter un responsable principal, des membres, des équipes et des tâches détaillées.

### 5. Générer les documents administratifs

- Aller dans **Demande de budget** pour générer le canevas financier Excel.
- Aller dans **Demande de congé** pour générer un document Word prêt à envoyer aux RH.

## Modèles inclus

- `templates/rapport_financier_canevas.xlsx` : canevas financier SADABE fourni, utilisé pour produire les demandes de budget.
- `templates/document_word_fourni_reference.docx` : document Word fourni comme référence. Le fichier fourni semble être un PV de passation, donc GREEN'APP génère un vrai canevas de demande de congé directement dans l’application.
- `sample_data/modele_import_activites.csv` : modèle simple pour tester l’import.

## Important pour un usage réel multi-utilisateur

La version actuelle utilise SQLite, très pratique pour un prototype et un usage local. Pour un vrai outil central utilisé par plusieurs personnes sur le long terme, il est conseillé de connecter GREEN'APP à une base externe comme PostgreSQL ou Supabase.

## Licence

MIT — open source.
