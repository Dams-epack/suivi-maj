# Suivi de préparation des MAJ

Outil mono-fichier (HTML + JavaScript, sans build) pour suivre la préparation des mises à jour de modules clients ePackPro.

## Contenu
- `index.html` : toute l'application (suivi, MAJ à venir, recherche, générateur XML, modèles de mails).

## Fonctions
- **Créer un suivi de MAJ** : une ligne par module et par région, avec 2 étapes (fichier d'import, dépôt dans le SharePoint) et un chronomètre qui s'arrête quand tout est coché.
- **Dates** : date de la demande + délai de prévenance (J-7 par défaut, J-5, ou dates personnalisées) ; les dates de MAJ siège (J-3) et de déploiement (J) sont calculées en jours ouvrés (lundi au vendredi).
- **Générateur XML** : import de l'Excel du client et génération de l'XML selon le module (Mise en température, Service, Production, Étiquetage, etc.), un onglet par région. Une DLC vide bloque la génération de l'onglet concerné (plus de -1 par défaut) : une pastille rouge s'affiche sur la région et un mail au siège est proposé avec la liste des produits. Les onglets sans erreur cochent « Fichier d'import » dès la génération. Chaque bloc de langue est réductible (réduit par défaut si aucune erreur) et le détail des valeurs nettoyées est consultable.
- **MAJ à venir** : deux blocs (MAJ siège, MAJ nationale), boutons ON/OFF d'envoi des mails, génération des mails types à copier.
- **Mails** : chaque bouton ON/OFF ouvre le mail correspondant (01 prise en charge 5j/7j, 02 MAJ siège, 03 début de déploiement, 04a/04b fin de déploiement, 05 seconde passe).
- **Teams** : Validation siège et Début de déploiement préparent un message (liens ressources et rapports demandés en popup, un rapport par langue), envoyé seulement au clic sur « Envoyer dans Teams » via un flux Power Automate (URL saisie dans la roue crantée, jamais dans le dépôt).
- **Fin de déploiement** : rappel du rapport d'intervention final, question « seconde passe ? » et switch « 2nde passe ».
- **Liens de la MAJ** : les liens saisis sont affichés et modifiables sur la page de la MAJ.
- **ZIP** : nommé « NOM DU GROUPEMENT - Nom du fichier XML.zip », avec un dossier par région (FR, BE-FR, BE-NL, IT, REUNION, ANTILLES, NC ; FR+LUX uniquement pour O TACOS).
- **XML générés** : chaque ZIP enregistré est historisé sur la fiche (date, région, module, fichier, auteur) dans un tableau triable, réduit à l'ouverture de la fiche ; le ZIP est conservé dans la base partagée pour être retéléchargé.
- **Analyse du mail client** : détection du groupement, du site ou concept, des modules et des régions ; pop-ups pour chaque point flou et mail de retour proposé au client ; contrôle du délai J+7 ouvrés avec nouveau rétroplanning ; ajout de modules après coup (« + Ajouter un module »).
- **Accueil** : indicateurs, bloc « Prochaines échéances » (prochaine étape de chaque MAJ, retards en rouge), outils et graphiques.
- **MAJ archivées** : tableau triable par enseigne, langue (codes région) et date.
- **Calendrier Outlook** : le bouton « Enregistrer et envoyer dans Outlook » à côté des dates demande confirmation puis envoie deux événements en journée entière (MAJ siège, Déploiement MAJ Nationale ou Internationale, avec codes pays) à un second flux Power Automate. L'URL se saisit dans la roue crantée (Paramètres), jamais dans le dépôt. Un avertissement rappelle de supprimer à la main les événements précédents (risque de doublons).
- **Groupements** : recherche intelligente (tolérance aux fautes) et codes pays ajoutés au nom de la MAJ si plusieurs régions.
- **Enregistrer** (bas de la page de la MAJ) : enregistre la fiche dans la base commune et revient à « MAJ à venir ». Il n'envoie rien vers Outlook.

## Utilisation
Conçu pour tourner comme Artifact claude.ai (stockage partagé via la base de l'Artifact). Ouvert hors de claude.ai, il retombe sur le stockage local du navigateur (`localStorage`), donc les données ne sont pas partagées entre postes.

## Dépendances
- SheetJS (`xlsx` 0.18.5) chargé depuis cdnjs pour lire les fichiers Excel.

## Sécurité
Aucune clé ni URL de flux n'est dans le code : les URL Power Automate (Teams et calendrier) sont saisies dans l'outil et conservée dans le navigateur.
