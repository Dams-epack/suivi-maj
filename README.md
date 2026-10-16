# Suivi de préparation des MAJ

Outil mono-fichier (HTML + JavaScript, sans build) pour suivre la préparation des mises à jour de modules clients ePackPro.

## Contenu
- `index.html` : toute l'application (suivi, MAJ à venir, recherche, générateur XML, modèles de mails).

## Fonctions
- **Créer un suivi de MAJ** : une ligne par module et par région, avec 2 étapes (fichier d'import, dépôt dans le SharePoint) et un chronomètre qui s'arrête quand tout est coché.
- **Dates** : date de la demande + délai de prévenance (J-7 par défaut, J-5, ou dates personnalisées) ; les dates de MAJ siège (J-3) et de déploiement (J) sont calculées en jours ouvrés (lundi au vendredi).
- **Générateur XML** : import de l'Excel du client et génération de l'XML selon le module (Mise en température, Service, Production, Étiquetage, etc.), enregistré en ZIP.
- **MAJ à venir** : deux blocs (MAJ siège, MAJ nationale), boutons ON/OFF d'envoi des mails, génération des mails types à copier.
- **Validation siège** : message Teams « NOM DU GROUPEMENT - Maj. validée par le siège », envoyé via un flux Power Automate (URL à coller dans l'outil, jamais dans le dépôt).

## Utilisation
Conçu pour tourner comme Artifact claude.ai (stockage partagé via la base de l'Artifact). Ouvert hors de claude.ai, il retombe sur le stockage local du navigateur (`localStorage`), donc les données ne sont pas partagées entre postes.

## Dépendances
- SheetJS (`xlsx` 0.18.5) chargé depuis cdnjs pour lire les fichiers Excel.

## Sécurité
Aucune clé ni URL de flux n'est dans le code : l'URL Power Automate est saisie dans l'outil et conservée dans le navigateur.
