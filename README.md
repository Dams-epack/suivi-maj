# Suivi maj

Page web autonome (un seul fichier, `index.html`) pour suivre la préparation des MAJ de modules demandées par les clients : liste des modules et des régions extraite du mail, suivi des étapes (import, configurateur, moulinette), progression, chronomètre et archivage.

## Utilisation

Ouvrir `index.html` dans un navigateur, ou la publier avec GitHub Pages (Settings > Pages > branche `main`, dossier racine).

## Hors de claude.ai

La page a été conçue dans claude.ai. Ailleurs :

- le suivi est enregistré dans le `localStorage` du navigateur (par poste, non partagé) ;
- l'analyse du mail passe par une lecture locale, sans l'aide de Claude ;
- les boutons de téléchargement direct peuvent proposer une solution de secours à la place.

## Données clients

Le `.gitignore` exclut les mails, XML, exports et sauvegardes pour éviter de pousser des données clients. Si le dépôt est public, relire `index.html` avant de pousser : la page cite « Boulangerie Ange » dans une règle de génération des XML.
