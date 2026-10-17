# L'ODDysée d'Hugo, N°1 automne 2026

Version en ligne de la newsletter EDD et ODD du Lycée français Victor Hugo de Sofia.

## Contenu du dossier

- `index.html` la page complète, animations et quiz compris, sans aucune installation
- `assets/img` les photos optimisées pour le web
- `assets/L_ODDysee_d_Hugo_n1_automne_2026.pdf` la version PDF A4, téléchargeable depuis le site
- `assets/L_ODDysee_d_Hugo_n1_livret_A3.pdf` le livret à imprimer en A3, téléchargeable depuis le site

## Mettre le site en ligne avec GitHub Pages

1. Sur github.com, créez un dépôt public, par exemple `oddysee-hugo`, ou ouvrez votre dépôt `laclassedegreg.github.io`.
2. Cliquez sur **Add file** puis **Upload files** et glissez tout le contenu de ce dossier, en gardant le dossier `assets`.
3. Validez avec **Commit changes**.
4. Dans un nouveau dépôt, ouvrez **Settings** puis **Pages**, choisissez la branche `main` et le dossier `/ (root)`, puis enregistrez.
5. Après une à deux minutes, le site est visible à l'adresse `https://VOTRE-NOM.github.io/oddysee-hugo/`.

Ce numéro est publié dans le dépôt `laclassedegreg.github.io`, dossier `oddysee-hugo-ed1`, à l'adresse `https://laclassedegreg.github.io/oddysee-hugo-ed1/`. Pour l'édition 2, créez un dossier `oddysee-hugo-ed2` à côté, l'édition 1 reste ainsi en ligne.

## Versions en trois langues

- français `index.html`, à l'adresse `https://laclassedegreg.github.io/oddysee-hugo-ed1/`
- anglais `en/index.html`, à l'adresse `https://laclassedegreg.github.io/oddysee-hugo-ed1/en/`
- bulgare `bg/index.html`, à l'adresse `https://laclassedegreg.github.io/oddysee-hugo-ed1/bg/`

Un sélecteur FR, EN, BG en haut de chaque page permet de passer d'une langue à l'autre. Les trois pages partagent les mêmes photos du dossier `assets/img`, et chaque langue a son PDF A4 et son livret A3 dans `assets`.

Important, déposez toujours le dossier décompressé et jamais le fichier zip, GitHub Pages ne sait pas ouvrir une archive.

## Préparer le numéro suivant

Chaque article est un bloc `<article class="card">` facile à repérer dans `index.html`. Pour un nouveau numéro, dupliquez le dossier, remplacez les textes et les photos, puis mettez à jour la date dans la barre du haut et les dates du bandeau « À vos agendas » (attributs `data-start` et `data-end`), qui alimentent le compte à rebours.

## Imprimer le livret A3

Imprimez `L_ODDysee_d_Hugo_n1_livret_A3.pdf` sur du papier A3, en recto verso avec l'option **retourner sur le bord court**, à taille réelle. Empilez les trois feuilles avec la feuille des pages 1 et 12 à l'extérieur et celle des pages 6 et 7 au centre, puis pliez la pile au milieu pour obtenir le journal A4.
