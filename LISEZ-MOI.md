# Suivi MAJ : guide rapide

Deux outils mono-fichier, sans installation :

- **Suivi de préparation des MAJ (SGC)** : `index.html`, publié comme page claude.ai.
- **Outils XML ePack (Support)** : `xml.html`, mêmes briques, réservé au Support (Excel → XML, Compléter un XML, XML → Excel). Un Mode d'emploi avec vidéos est intégré (bouton livre).

Les détails techniques et la liste complète des fonctions sont dans `README.md`.

## Suivi des MAJ

1. **Créer** : colle le mail du client ; l'outil propose groupement, modules, régions et dates (délai J-7, J-5 ou dates personnalisées).
2. **Préparer** : une ligne par module et par région ; coche « Fichier d'import », « Configurateur (XML) » et « Moulinette ». Le chrono s'arrête quand tout est coché.
3. **Excel → XML** : bouton sur chaque ligne ; l'outil lit l'Excel du client et génère l'XML.
4. **Mails** : les boutons de « MAJ à venir » ouvrent les mails types (prise en charge, MAJ siège, début, fin, seconde passe).
5. **Suivre** : l'accueil affiche « À traiter aujourd'hui » (retards, mails à envoyer, MAJ bloquées, MAJ à archiver).
6. **Archiver** : une fois la fin de déploiement envoyée, « Valider ». Une MAJ archivée est en **lecture seule** ; « Désarchiver » la rend modifiable.
7. **Statistiques** : « Statistiques détaillées » (carte « Temps de préparation » de l'accueil) donne les temps moyens par personne, par type et par module.

## Outils XML (Support)

- **Excel → XML** : dépose le fichier, choisis le ou les types, vérifie les colonnes (« Ajuster les colonnes »), complète les valeurs manquantes puis génère. Les erreurs bloquantes (Id en double, durée négative, minimum supérieur au maximum) sont expliquées ligne par ligne.
- **Compléter un XML** et **XML → Excel** : voir le Mode d'emploi et les vidéos.
- Tout reste dans le navigateur : rien n'est envoyé.

## Mettre à jour (mainteneur)

```
python3 build_support.py index.html xml.html
```
puis publier `index.html` et `xml.html`, et commit. Penser à incrémenter `APP_VERSION` (et `SB`).
