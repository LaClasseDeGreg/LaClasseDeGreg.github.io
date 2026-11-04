# Chronopolis, parcours de révision de l'urbanisation du monde (4e)

Grégory Cron, LaClasseDeGreg, Lycée français Victor Hugo de Sofia, 2026-2027

Adresse du jeu https://laclassedegreg.github.io/chronopolis-urbanisation/
Espace professeur, bouton en bas de page, ou https://laclassedegreg.github.io/chronopolis-urbanisation/#prof

## Ce que contient le dossier

- `index.html`, le jeu complet (scénario, défis, corrections, bilan, salle d'arcade, espace professeur)
- `img/`, les photographies (Unsplash) et le logo du lycée
- `cours/Cours_4e_Urbanisation_LVH.pdf`, le cours que les élèves téléchargent depuis le jeu
- ce fichier

## Relier le suivi des élèves (une seule fois, environ 5 minutes)

1. Ouvrir la feuille Google « Suivi Chronopolis 4e » (dossier Drive Séquence 2 GEO urbanisation), ou en créer une avec https://sheets.new
2. Menu Extensions, puis Apps Script. Effacer le code présent, coller tout le contenu du fichier « Code Apps Script Chronopolis 4e » rangé dans le même dossier Drive, enregistrer. Ce code n'est volontairement pas publié sur le site, car il contient le mot de passe de l'espace professeur.
3. Choisir la fonction `initialiser` dans la liste, cliquer sur Exécuter et accepter les autorisations Google.
4. Bouton Déployer, puis Nouveau déploiement. Type Application Web. Exécuter en tant que Moi. Qui a accès, Tout le monde. Déployer.
5. Copier l'adresse qui se termine par `/exec` et l'envoyer à Claude, qui la colle dans le jeu et vérifie l'enregistrement. On peut aussi la coller soi-même dans `index.html`, sur la ligne `const SCRIPT_URL = "";`

## Bon à savoir

- Le mot de passe de l'espace professeur se trouve uniquement dans le script Google, sur sa première ligne (`MOT_DE_PASSE`). Pour le changer, modifier cette ligne, puis Déployer, Gérer les déploiements, Modifier, Nouvelle version.
- Chaque partie crée une ligne, mise à jour à chaque page. Le tableau de bord retient pour chaque élève la meilleure partie terminée.
- Le tableau de bord montre les élèves qui n'ont pas encore joué, les notions les moins maîtrisées, les questions les plus ratées et les paragraphes rédigés. Export CSV possible.
- Les élèves sont affichés par prénom et initiale du nom. Les visiteurs choisissent un pseudo.
- Une partie interrompue reprend sur le même appareil. Le bouton Lecture facile agrandit le texte et change la police, le bouton Lire à voix haute lit les dialogues.
- Photographies Unsplash (licence Unsplash), personnages dessinés pour LaClasseDeGreg, cartes d'après Natural Earth.
