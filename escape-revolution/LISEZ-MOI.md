# Les Archives brûlent à l'aube, mise en ligne et suivi des élèves

Grégory Cron, LaClasseDeGreg, Lycée français Victor Hugo de Sofia

## Ce que contient le dossier

- `index.html`, le jeu complet (leçons, images, PDF et espace professeur inclus)
- `musique/`, les neuf musiques de fond, chargées une par une pendant la partie
- `apps-script-suivi.gs`, le petit programme qui enregistre les résultats dans une feuille Google
- ce fichier

## Étape 1, créer la feuille de suivi (5 minutes, une seule fois)

1. Ouvrir https://sheets.new avec le compte Google du lycée et nommer la feuille « Suivi escape game Révolution ».
2. Menu Extensions, puis Apps Script. Effacer le code présent, coller tout le contenu de `apps-script-suivi.gs`, puis enregistrer.
3. Bouton Déployer, puis Nouveau déploiement. Choisir le type Application Web. Exécuter en tant que Moi. Qui a accès, Tout le monde. Cliquer sur Déployer et accepter les autorisations demandées par Google.
4. Copier l'URL de l'application Web, elle se termine par `/exec`.

## Étape 2, relier le jeu à la feuille

Ouvrir `index.html` dans un éditeur de texte, chercher la ligne `const SCRIPT_URL = "";` et coller l'URL entre les guillemets. Enregistrer. Tu peux aussi m'envoyer l'URL, je te renvoie le fichier prêt.

## Étape 3, mettre en ligne

Déposer `index.html` et le dossier `musique` dans un dossier `escape-revolution` du dépôt laclassedegreg.github.io. Le jeu sera accessible à l'adresse https://laclassedegreg.github.io/escape-revolution/

## Étape 4, tester avant la classe

Jouer une partie avec un nom d'élève, puis cliquer sur « Espace professeur » en bas de page et saisir le mot de passe. La ligne doit apparaître dans le tableau et dans la feuille Google.

## Bon à savoir

- Le mot de passe se trouve uniquement dans le script Google, jamais dans la page. Pour le changer, modifier la première ligne du script, puis Déployer, Gérer les déploiements, Modifier, Nouvelle version.
- Chaque partie crée une ligne, mise à jour à chaque registre sauvé. Le tableau retient la meilleure note parmi les parties terminées.
- Pour remettre le suivi à zéro, supprimer les lignes de l'onglet « Suivi » (garder la ligne d'en-tête).
- Les élèves sont affichés par prénom et initiale du nom, pour ne pas publier leurs noms complets sur un site ouvert.
- Musiques issues de Pixabay, libres de droits pour un usage pédagogique.
