/**
 * Suivi de l'escape game « Les Archives brûlent à l'aube »
 * Grégory Cron, LaClasseDeGreg, Lycée français Victor Hugo de Sofia
 *
 * À coller dans Extensions > Apps Script d'une feuille Google Sheets,
 * puis à déployer comme application Web (voir LISEZ-MOI).
 * Le mot de passe reste ici, côté Google. Il n'apparaît jamais dans la page du jeu.
 */
const MOT_DE_PASSE = '2702';
const ONGLET = 'Suivi';
const COLONNES = ['session', 'premiere_activite', 'derniere_activite', 'classe', 'eleve', 'temoin',
  'registres', 'sceaux_intacts', 'bonnes', 'repondues', 'total', 'pourcentage', 'note20',
  'temps_min', 'erreurs', 'termine', 'fouche'];

function feuille_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ONGLET);
  if (!sh) {
    sh = ss.insertSheet(ONGLET);
    sh.appendRow(COLONNES);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLONNES.length).setFontWeight('bold');
  }
  return sh;
}

// Réception des résultats envoyés par le jeu (une ligne par partie, mise à jour à chaque registre)
function doPost(e) {
  const verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.session || !d.eleve || !d.classe) return texte_('ignoré');
    const sh = feuille_();
    const maintenant = new Date();
    const valeurs = sh.getDataRange().getValues();
    let ligne = -1;
    for (let i = 1; i < valeurs.length; i++) {
      if (valeurs[i][0] === d.session) { ligne = i + 1; break; }
    }
    const premiere = ligne > 0 ? valeurs[ligne - 1][1] : maintenant;
    const rangee = COLONNES.map(function (c) {
      if (c === 'premiere_activite') return premiere;
      if (c === 'derniere_activite') return maintenant;
      return d[c] === undefined ? '' : String(d[c]).slice(0, 80);
    });
    if (ligne > 0) sh.getRange(ligne, 1, 1, COLONNES.length).setValues([rangee]);
    else sh.appendRow(rangee);
    return texte_('ok');
  } finally {
    verrou.releaseLock();
  }
}

// Lecture par l'espace professeur, uniquement avec le bon mot de passe
function doGet(e) {
  if (!e || !e.parameter || e.parameter.pwd !== MOT_DE_PASSE) return json_({ ok: false });
  const valeurs = feuille_().getDataRange().getValues();
  const tete = valeurs.shift();
  const lignes = valeurs.map(function (v) {
    const o = {};
    tete.forEach(function (h, i) { o[h] = v[i] instanceof Date ? v[i].toISOString() : v[i]; });
    return o;
  });
  return json_({ ok: true, rows: lignes });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function texte_(s) {
  return ContentService.createTextOutput(s);
}
