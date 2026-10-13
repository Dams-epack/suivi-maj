# Suivi MAJ – version autonome (GitHub Pages + fichier SharePoint)

Cette version n'a plus besoin de Claude. La page est hébergée sur GitHub Pages et les données partagées
sont dans **un seul fichier JSON** placé dans un dossier SharePoint synchronisé avec OneDrive.

Deux dossiers :
- `site-github/` : contient `index.html`, la page. **C'est la seule chose à mettre sur GitHub.**
- `donnees/` : contient `suivi-maj-donnees.json`, les données de départ (2 MAJ archivées + 135 groupes).
  **Ne jamais le mettre sur GitHub** : il va dans SharePoint.

## 1. Publier la page sur GitHub Pages (une seule fois)

1. Sur github.com, crée un dépôt, par exemple `suivi-maj` (onglet « New repository »).
   - GitHub Pages gratuit exige un dépôt **public**. La page ne contient aucune donnée client (seulement le catalogue
     des modules et les régions), mais n'importe qui avec le lien peut l'ouvrir : elle apparaît vide sans le fichier de données.
2. « Add file » > « Upload files » : dépose `index.html` (celui du dossier `site-github`) > « Commit changes ».
3. « Settings » > « Pages » > Source : « Deploy from a branch », branche `main`, dossier `/ (root)` > « Save ».
4. Après 1 à 2 minutes, l'adresse est : `https://TON-COMPTE.github.io/suivi-maj/`. C'est le lien à donner aux collègues.

Pour mettre à jour la page plus tard : remplace `index.html` dans le dépôt (« Upload files », même nom).

## 2. Mettre le fichier de données dans SharePoint

1. Dépose `suivi-maj-donnees.json` dans un dossier SharePoint de l'équipe (par exemple
   `SGC > 04 RESSOURCES INTERNES > 00. Skills Claude AI > Suivi des majs`).
2. Chaque collègue synchronise ce dossier avec OneDrive (bouton « Synchroniser » dans SharePoint), puis clic droit
   sur le dossier > « Toujours conserver sur cet appareil » (le fichier doit être réellement présent sur le poste).

## 3. Se connecter au fichier (chaque collègue, une fois)

1. Ouvrir la page **avec Microsoft Edge ou Google Chrome** (Firefox et Safari ne conviennent pas).
2. Dans la barre du haut : « Ouvrir le fichier de données » > choisir `suivi-maj-donnees.json` dans le dossier synchronisé.
3. Facultatif : saisir son prénom dans la barre (il apparaît comme « préparé par »).

Le navigateur retient le fichier. À une prochaine visite, il peut demander de confirmer l'autorisation :
cliquer sur « Reconnecter » dans la barre (dans Chrome, choisir « Autoriser à chaque visite » évite de le refaire).

## À savoir

- **Synchronisation** : la page relit le fichier toutes les 4 secondes. Mais OneDrive met quelques secondes à quelques
  minutes à propager le fichier entre les postes, donc ce n'est pas du temps réel.
- **Écritures simultanées** : chaque enregistrement relit le fichier avant d'écrire. Si deux personnes enregistrent au même instant,
  OneDrive peut créer une copie « conflit » du fichier : à ce moment-là, garder la plus récente et supprimer l'autre.
  Si deux personnes cochent des étapes de la **même ligne** en même temps, la dernière écriture gagne.
- **Sauvegarde** : le bouton « Sauvegarder (JSON) » de la page d'accueil reste disponible. SharePoint garde aussi l'historique
  des versions du fichier (« Historique des versions »), ce qui permet de revenir en arrière.
- **Téléchargement du XML** : le fichier XML est téléchargé directement (plus de zip).
- **Lecture du mail** : sans Claude, l'analyse du mail est locale (moins souple). Le « Suivi personnalisé » reste le plus fiable.
- **Données** : le fichier JSON contient tes MAJ et les groupes. Il reste dans ton SharePoint, jamais sur GitHub.
