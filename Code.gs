/**
 * RÉPARTITION CRÉDIT SCOOPS — Baobab Burkina
 * ------------------------------------------
 * Web app (HTML Service) reproduisant la structure de REPARTITION_CREDIT_SCOOPS.xlsx.
 *
 * Logique financière validée sur Yiwo Yo de Tio et Doundolezi :
 *   Ratio        = Capital accordé / Σ besoins
 *   Capital(i)   = Besoin(i) × Ratio
 *   FraisTenue   = barème[moisAgio saisi par l'agent]
 *   TotalReel    = Σ échéances réelles + FraisTenue
 *   DépôtTotal   = Σ Capital(i) × TauxDépôt
 *   CoutTotal    = TotalReel − DépôtTotal − Montant Approuvé
 *   Intérêt(i)   = Capital(i) × CoutTotal / Montant Approuvé
 *
 * NOTE ARCHITECTURE : aucune plage nommée (Named Range) n'est utilisée.
 * Toutes les dépendances inter-cellules reposent sur des références absolues
 * directes ($O$2, $P$14, etc.) pré-calculées avant écriture du classeur.
 * Cela évite les erreurs "La plage nommée X n'existe pas" liées aux limites
 * de l'API setNamedRange dans Apps Script.
 *
 * NOTE EXPORT : DriveApp.getAs() ne sait pas convertir un Google Sheet en
 * .xlsx. L'export passe par l'URL d'export de Google Sheets (UrlFetchApp),
 * après SpreadsheetApp.flush() pour que toutes les écritures soient incluses.
 */

var INPUT_FONT_COLOR  = '#1155CC';
var FORMULA_FONT_COLOR = '#000000';
var HEADER_FILL  = '#0F3D3E';
var HEADER_FONT  = '#FFFFFF';
var SECTION_FILL = '#E7F2F1';
var WARN_FILL    = '#F4C7C3';
var NUM_FMT = '#,##0';
var CODE_VERSION = 'v2-sans-plages-nommees-2026-10-08';
var PCT_FMT = '0%';

/* ───── Entrée web app ───── */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Répartition Crédit SCOOPS — Baobab')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ───── Point d'entrée appelé depuis le client ───── */

function generateWorkbook(payload) {
  try {
    var result = generateWorkbookImpl_(payload);
    result.codeVersion = CODE_VERSION;
    return result;
  } catch (e) {
    // Le préfixe de version prouve quel code serveur a produit l'erreur.
    throw new Error('[' + CODE_VERSION + '] ' + (e && e.message ? e.message : e));
  }
}

/** Renvoie la version du code serveur réellement exécuté. */
function getCodeVersion() {
  return CODE_VERSION;
}

/**
 * Test à lancer depuis l'éditeur (menu « Exécuter » → testGeneration).
 * Crée un classeur de test et affiche son URL dans le journal d'exécution.
 */
function testGeneration() {
  var res = generateWorkbook({
    keepGoogleSheet: true,
    scoops: [{
      nom: 'SCOOPS FONG NINTA', localite: 'Test',
      capitalAccorde: 1000000, tauxDepot: 0.10, dureeMois: 6, moisAgio: 6,
      membres: [{ nom: 'Membre A', besoin: 600000 }, { nom: 'Membre B', besoin: 600000 }],
      echeances: [{ date: '2026-01-15', montant: 600000 }, { date: '2026-02-15', montant: 600000 }]
    }]
  });
  Logger.log('Version : ' + res.codeVersion);
  Logger.log('Classeur : ' + res.sheetUrl);
  Logger.log(res.base64 ? 'Export .xlsx : OK' : 'Export .xlsx : ÉCHEC — ' + res.exportError);
}

function generateWorkbookImpl_(payload) {
  if (!payload || !payload.scoops || payload.scoops.length === 0) {
    throw new Error('Aucun SCOOP à traiter.');
  }

  var ss;
  var reuseExisting = payload.existingSpreadsheetId &&
    String(payload.existingSpreadsheetId).trim() !== '';
  if (reuseExisting) {
    ss = SpreadsheetApp.openById(String(payload.existingSpreadsheetId).trim());
  } else {
    ss = SpreadsheetApp.create('REPARTITION_CREDIT_SCOOPS_' + formatDateStamp_());
    // Aligne le fuseau du classeur sur celui du script (dates d'échéances).
    ss.setSpreadsheetTimeZone(scriptTimeZone_());
  }

  var defaultSheet = reuseExisting ? null : ss.getSheets()[0];
  var summary = [];

  payload.scoops.forEach(function(scoop, idx) {
    var sheet;
    if (defaultSheet && idx === 0) {
      sheet = defaultSheet;
      sheet.clear();
      sheet.clearConditionalFormatRules();
    } else {
      sheet = ss.insertSheet();
    }
    summary.push(buildScoopSheet_(ss, sheet, scoop || {}, idx));
  });

  // Indispensable : sans flush, l'export peut ne pas contenir les dernières écritures.
  SpreadsheetApp.flush();

  var fileId   = ss.getId();
  var filename = ss.getName() + '.xlsx';
  var base64 = null, exportError = null;
  try {
    base64 = Utilities.base64Encode(exportAsXlsx_(fileId).getBytes());
  } catch (e) {
    // L'export peut être bloqué (autorisation script.external_request absente
    // ou refusée par l'administrateur Workspace). On ne perd pas le travail :
    // le classeur Google Sheets est conservé et son lien est renvoyé.
    exportError = (e && e.message) ? e.message : String(e);
  }

  var keep = payload.keepGoogleSheet !== false || base64 === null;
  if (!keep && !reuseExisting) DriveApp.getFileById(fileId).setTrashed(true);

  return {
    base64: base64,
    filename: filename,
    sheetUrl: (keep || reuseExisting) ? ss.getUrl() : null,
    exportError: exportError,
    summary: summary
  };
}

/* ───── Export .xlsx ───── */

function exportAsXlsx_(fileId) {
  var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Échec de l\'export .xlsx (HTTP ' + resp.getResponseCode() + ').');
  }
  return resp.getBlob();
}

/* ───── Construction d'une feuille SCOOP ───── */

function buildScoopSheet_(ss, sheet, scoop, sheetIdx) {
  /* ── Nommage de la feuille ── */
  var sheetName = String(scoop.nom || ('SCOOP ' + (sheetIdx + 1))).trim().substring(0, 95);
  if (!sheetName) sheetName = 'SCOOP ' + (sheetIdx + 1);
  sheet.setName(uniqueSheetName_(ss, sheet, sheetName));

  /* ── Données sources (normalisées : les champs HTML arrivent en texte) ── */
  var membres   = scoop.membres  || [];
  var echeances = scoop.echeances || [];
  var baremeList = (scoop.bareme && scoop.bareme.length)
    ? scoop.bareme
    : [{ mois:5, montant:5900 },  { mois:6, montant:7080 },
       { mois:7, montant:8260 },  { mois:8, montant:9440 },
       { mois:9, montant:10620 }, { mois:10, montant:11800 },
       { mois:11, montant:12980 },{ mois:12, montant:14160 }];  // = defaultBareme() côté client

  var capitalAccorde = toNumber_(scoop.capitalAccorde, 0);
  var tauxDepot      = toRate_(scoop.tauxDepot, 0.1);
  var dureeMois      = toNumber_(scoop.dureeMois, 0) || 6;
  var moisAgio       = toNumber_(scoop.moisAgio, 0) || 6;

  /* ── Colonnes du panneau paramètres (droit) ── */
  var LABEL = 14, VAL = 15, VAL2 = 16;  // N, O, P
  var cL = columnToLetter_(LABEL);       // 'N'
  var cV = columnToLetter_(VAL);         // 'O'
  var cV2 = columnToLetter_(VAL2);       // 'P'

  /* ── Compteurs de lignes ── */
  var n  = membres.length   || 1;   // lignes données (≥1)
  var ne = echeances.length || 1;   // lignes échéancier (≥1)
  var nb = baremeList.length;        // lignes barème

  /* ─────────────────────────────────────────────────────────────
   * Adresses absolues pré-calculées du panneau paramètres
   * (col O = VAL sauf totalEcheances et fraisRetenu qui sont en col P)
   *
   * Disposition verticale du panneau (col N/O/P, démarrage ligne 1) :
   *  1  : Header PARAMÈTRES
   *  2  : Capital accordé              ← $O$2
   *  3  : Somme besoins               ← $O$3
   *  4  : Ratio                       ← $O$4
   *  5  : Taux dépôt                  ← $O$5
   *  6  : Durée                       ← $O$6
   *  7  : Mois d'agio                 ← $O$7
   *  8  : (blanc)
   *  9  : Header ÉCHÉANCIER
   * 10  : En-têtes colonnes
   * 11..10+ne : lignes échéances
   * 11+ne     : Total échéances       ← $P$(11+ne)
   * 12+ne     : (blanc)
   * 13+ne     : Header BARÈME
   * 14+ne     : En-têtes colonnes
   * 15+ne..14+ne+nb : lignes barème
   * 15+ne+nb  : (blanc)
   * 16+ne+nb  : Frais retenu          ← $P$(16+ne+nb)
   * 17+ne+nb  : (blanc)
   * 18+ne+nb  : Total réel            ← $O$(18+ne+nb)
   * 19+ne+nb  : Montant approuvé      ← $O$(19+ne+nb)
   * 20+ne+nb  : Dépôt total           ← $O$(20+ne+nb)
   * 21+ne+nb  : Coût total crédit     ← $O$(21+ne+nb)
   * 22+ne+nb  : Écart vérification
   * ───────────────────────────────────────────────────────────── */
  var A = {                                  // A = "Addresses"
    capitalAccorde:  '$'+cV+'$2',
    sommeBesoins:    '$'+cV+'$3',
    ratio:           '$'+cV+'$4',
    tauxDepot:       '$'+cV+'$5',
    duree:           '$'+cV+'$6',
    moisAgio:        '$'+cV+'$7',
    totalEcheances:  '$'+cV2+'$'+(11+ne),
    fraisRetenu:     '$'+cV2+'$'+(16+ne+nb),
    totalReel:       '$'+cV+'$'+(18+ne+nb),
    montantApprouve: '$'+cV+'$'+(19+ne+nb),
    depotTotal:      '$'+cV+'$'+(20+ne+nb),
    coutTotal:       '$'+cV+'$'+(21+ne+nb)
  };

  /* Plage barème pour VLOOKUP / MATCH */
  var baremeStart = 15 + ne;
  var baremeEnd   = 14 + ne + nb;
  var fraisRange  = '$'+cL+'$'+baremeStart+':$'+cV+'$'+baremeEnd;
  var moisRange   = '$'+cL+'$'+baremeStart+':$'+cL+'$'+baremeEnd;

  /* ── Section données : disposition (cols A–L) ── */
  var headerRow   = 4;
  var firstDataRow = headerRow + 1;             // 5
  var lastDataRow  = firstDataRow + n - 1;
  var totalRow     = lastDataRow + 2;

  /* ═══════════════════════════════════════════
   *  ÉCRITURE DE LA FEUILLE
   * ═══════════════════════════════════════════ */

  /* — En-tête principale — */
  sheet.getRange(1, 1, 1, 12).merge()
    .setValue('SCOOP : '+(scoop.nom||'')+'   —   Localité : '+(scoop.localite||''))
    .setFontWeight('bold').setFontSize(13)
    .setBackground(HEADER_FILL).setFontColor(HEADER_FONT)
    .setHorizontalAlignment('left');
  sheet.getRange(2, 1, 1, 12).merge()
    .setValue("ATTENTION : cet exercice n'est valable que si l'échéancier réel du crédit (tableau d'amortissement) est disponible et exact.")
    .setFontStyle('italic').setFontSize(9).setFontColor('#555555');

  /* — En-têtes colonnes données — */
  var colHeaders = [
    'N°', 'NOM ET PRENOMS', 'Besoin exprimé', 'Capital retenu',
    'Intérêt + Agios', 'Durée (mois)', 'Intérêt Mensuel', 'Intérêt Total Crédit',
    'Taux Dépôt', 'Dépôt',
    'Montant à Rembourser (capital+intérêt+dépôt)',
    'Montant à Rembourser (intérêt déjà remboursé : capital+dépôt)'
  ];
  sheet.getRange(headerRow, 1, 1, colHeaders.length)
    .setValues([colHeaders]).setFontWeight('bold')
    .setBackground(HEADER_FILL).setFontColor(HEADER_FONT)
    .setWrap(true).setVerticalAlignment('middle');

  /* — Lignes membres (écriture groupée : 2 appels au lieu de 12 × n) — */
  if (membres.length > 0) {
    var inputs = [], formulas = [];
    for (var i = 0; i < membres.length; i++) {
      var r = firstDataRow + i;
      var m = membres[i] || {};
      inputs.push([i + 1, m.nom || '', toNumber_(m.besoin, 0)]);
      formulas.push([
        '=C'+r+'*'+A.ratio,                                              // D Capital
        '=IFERROR(D'+r+'*'+A.coutTotal+'/'+A.montantApprouve+',0)',      // E Intérêt
        '='+A.duree,                                                     // F Durée
        '=IFERROR(E'+r+'/F'+r+',0)',                                     // G Int. mensuel
        '=G'+r+'*F'+r,                                                   // H Int. total
        '='+A.tauxDepot,                                                 // I Taux dépôt
        '=D'+r+'*I'+r,                                                   // J Dépôt
        '=D'+r+'+E'+r+'+J'+r,                                            // K Cap+Int+Dép
        '=D'+r+'+J'+r                                                    // L Cap+Dép
      ]);
    }
    sheet.getRange(firstDataRow, 1, n, 3).setValues(inputs);
    sheet.getRange(firstDataRow, 3, n, 1).setFontColor(INPUT_FONT_COLOR);
    sheet.getRange(firstDataRow, 4, n, 9).setFormulas(formulas)
      .setFontColor(FORMULA_FONT_COLOR);
    sheet.getRange(firstDataRow, 3, n, 10).setNumberFormat(NUM_FMT);
    sheet.getRange(firstDataRow, 6, n, 1).setNumberFormat('0');
    sheet.getRange(firstDataRow, 9, n, 1).setNumberFormat(PCT_FMT);
  } else {
    sheet.getRange(firstDataRow, 2).setValue('(aucun membre saisi)');
    sheet.getRange(firstDataRow, 3, 1, 10).setValue(0);
  }

  /* — Ligne totaux — */
  sheet.getRange(totalRow, 2).setValue('TOTAL GÉNÉRAL').setFontWeight('bold');
  [3, 4, 5, 7, 8, 10, 11, 12].forEach(function(col) {
    var cLetter = columnToLetter_(col);
    setFormula_(sheet.getRange(totalRow, col),
      '=SUM('+cLetter+firstDataRow+':'+cLetter+lastDataRow+')', NUM_FMT);
  });
  sheet.getRange(totalRow, 1, 1, 12).setBorder(true, false, true, false, false, false);
  sheet.getRange(totalRow, 2, 1, 11).setFontWeight('bold');

  /* ═══════════════════════════════════════════
   *  PANNEAU PARAMÈTRES (droite : cols N / O / P)
   * ═══════════════════════════════════════════ */
  var row = 1;

  /* Header PARAMÈTRES */
  sheet.getRange(row, LABEL, 1, 3).merge()
    .setValue('PARAMÈTRES DU SCOOP').setFontWeight('bold').setBackground(SECTION_FILL);
  row++;  // 2

  /* Capital accordé */
  sheet.getRange(row, LABEL).setValue('Capital accordé au SCOOP (FCFA)');
  setInput_(sheet.getRange(row, VAL), capitalAccorde, NUM_FMT);
  row++;  // 3

  /* Somme besoins */
  sheet.getRange(row, LABEL).setValue('Somme des besoins exprimés (FCFA)');
  setFormula_(sheet.getRange(row, VAL),
    '=SUM(C'+firstDataRow+':C'+lastDataRow+')', NUM_FMT);
  row++;  // 4

  /* Ratio */
  sheet.getRange(row, LABEL).setValue('Ratio de prorata (Accordé / Besoins)');
  setFormula_(sheet.getRange(row, VAL),
    '=IFERROR('+A.capitalAccorde+'/'+A.sommeBesoins+',1)', '0.000');
  row++;  // 5

  /* Taux dépôt */
  sheet.getRange(row, LABEL).setValue('Taux Dépôt');
  setInput_(sheet.getRange(row, VAL), tauxDepot, PCT_FMT);
  row++;  // 6

  /* Durée */
  sheet.getRange(row, LABEL).setValue('Durée du crédit (mois)');
  setInput_(sheet.getRange(row, VAL), dureeMois, '0');
  row++;  // 7

  /* Mois d'agio */
  sheet.getRange(row, LABEL).setValue("Mois d'agio retenu (jugement de l'agent)");
  setInput_(sheet.getRange(row, VAL), moisAgio, '0');
  row++;  // 8 → blanc

  row++;  // 9 → ÉCHÉANCIER header

  sheet.getRange(row, LABEL, 1, 3).merge()
    .setValue('ÉCHÉANCIER RÉEL DU CRÉDIT').setFontWeight('bold').setBackground(SECTION_FILL);
  row++;  // 10 → en-têtes cols

  sheet.getRange(row, LABEL).setValue('N°').setFontWeight('bold');
  sheet.getRange(row, VAL).setValue('Date').setFontWeight('bold');
  sheet.getRange(row, VAL2).setValue('Montant (FCFA)').setFontWeight('bold');
  row++;  // 11 → première ligne échéance

  var echStart = row;   // = 11
  if (echeances.length === 0) {
    sheet.getRange(row, LABEL).setValue(1);
    setInput_(sheet.getRange(row, VAL2), 0, NUM_FMT);
    row++;
  } else {
    echeances.forEach(function(e, i) {
      e = e || {};
      sheet.getRange(row, LABEL).setValue(i + 1);
      var dateCell = sheet.getRange(row, VAL);
      var parsedDate = parseDateInput_(e.date);
      if (parsedDate) {
        // Formule DATE() : insensible au décalage de fuseau script ↔ classeur.
        dateCell.setFormula('=DATE(' + parsedDate.y + ',' + parsedDate.m + ',' + parsedDate.d + ')')
          .setNumberFormat('dd/mm/yyyy').setFontColor(INPUT_FONT_COLOR);
      } else {
        setInput_(dateCell, e.date || '', null);
      }
      setInput_(sheet.getRange(row, VAL2), toNumber_(e.montant, 0), NUM_FMT);
      row++;
    });
  }
  var echEnd = row - 1;  // = 10+ne

  /* Total échéances  (= ligne 11+ne, col P) */
  sheet.getRange(row, LABEL, 1, 2).merge()
    .setValue('Total échéances').setFontWeight('bold');
  setFormula_(sheet.getRange(row, VAL2),
    '=SUM('+cV2+echStart+':'+cV2+echEnd+')', NUM_FMT);
  row += 2;  // +2 → 13+ne

  /* Header BARÈME */
  sheet.getRange(row, LABEL, 1, 3).merge()
    .setValue('BARÈME FRAIS DE TENUE DE COMPTE (éditable)')
    .setFontWeight('bold').setBackground(SECTION_FILL);
  row++;  // 14+ne

  sheet.getRange(row, LABEL).setValue('Mois').setFontWeight('bold');
  sheet.getRange(row, VAL).setValue('Montant (FCFA)').setFontWeight('bold');
  row++;  // 15+ne → première ligne barème

  baremeList.forEach(function(b) {
    b = b || {};
    setInput_(sheet.getRange(row, LABEL), toNumber_(b.mois, 0),    '0');
    setInput_(sheet.getRange(row, VAL),   toNumber_(b.montant, 0), NUM_FMT);
    row++;
  });
  row++;  // 15+ne+nb → blanc, puis row = 16+ne+nb

  /* Frais de tenue de compte retenu  (ligne 16+ne+nb, col P) */
  sheet.getRange(row, LABEL, 1, 2).merge()
    .setValue('Frais de tenue de compte retenu');
  setFormula_(sheet.getRange(row, VAL2),
    '=IFERROR(VLOOKUP('+A.moisAgio+','+fraisRange+',2,FALSE),0)', NUM_FMT);
  /* Avertissement si mois absent du barème (col Q) */
  var warnCell = sheet.getRange(row, VAL2 + 1);
  setFormula_(warnCell,
    '=IF(ISNA(MATCH('+A.moisAgio+','+moisRange+',0)),"⚠ mois absent du barème","")', null);
  warnCell.setFontColor('#B00020').setFontStyle('italic');
  row += 2;  // 18+ne+nb

  /* Total réel */
  sheet.getRange(row, LABEL)
    .setValue('Total à rembourser réel (Échéances + Frais de tenue)');
  setFormula_(sheet.getRange(row, VAL),
    '='+A.totalEcheances+'+'+A.fraisRetenu, NUM_FMT);
  row++;  // 19+ne+nb

  /* Montant approuvé */
  sheet.getRange(row, LABEL).setValue('Montant Approuvé (= Σ Capital retenu)');
  setFormula_(sheet.getRange(row, VAL), '=D'+totalRow, NUM_FMT);
  row++;  // 20+ne+nb

  /* Dépôt total */
  sheet.getRange(row, LABEL).setValue('Dépôt total (= Σ Dépôt)');
  setFormula_(sheet.getRange(row, VAL), '=J'+totalRow, NUM_FMT);
  row++;  // 21+ne+nb

  /* Coût total du crédit */
  sheet.getRange(row, LABEL).setValue('COÛT TOTAL DU CRÉDIT').setFontWeight('bold');
  var coutTotalCell = sheet.getRange(row, VAL);
  setFormula_(coutTotalCell,
    '='+A.totalReel+'-'+A.depotTotal+'-'+A.montantApprouve, NUM_FMT);
  coutTotalCell.setFontWeight('bold');
  row++;  // 22+ne+nb

  /* Écart de vérification */
  sheet.getRange(row, LABEL)
    .setValue('Écart de vérification (Capital accordé − Montant Approuvé)');
  var ecartCell = sheet.getRange(row, VAL);
  setFormula_(ecartCell, '='+A.capitalAccorde+'-'+A.montantApprouve, NUM_FMT);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ABS($'+cV+'$'+row+')>1')
    .setBackground(WARN_FILL)
    .setRanges([ecartCell])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);

  /* — Mise en forme finale — */
  sheet.setColumnWidth(2, 170);
  for (var cc = 3; cc <= 12; cc++) sheet.setColumnWidth(cc, 95);
  sheet.setColumnWidth(LABEL, 230);
  sheet.setColumnWidth(VAL,   110);
  sheet.setColumnWidth(VAL2,  110);
  sheet.setFrozenRows(headerRow);

  return {
    sheetName: sheet.getName(),
    nbMembres: membres.length,
    capitalAccorde: capitalAccorde
  };
}

/* ───── Utilitaires d'écriture ───── */

function setInput_(range, value, fmt) {
  range.setValue(value).setFontColor(INPUT_FONT_COLOR);
  if (fmt) range.setNumberFormat(fmt);
}

function setFormula_(range, formula, fmt) {
  range.setFormula(formula).setFontColor(FORMULA_FONT_COLOR);
  if (fmt) range.setNumberFormat(fmt);
}

/* ───── Utilitaires généraux ───── */

/**
 * Convertit une saisie (nombre ou texte "1 250 000", "12,5") en nombre.
 * Sans cette conversion, un montant saisi avec des espaces est écrit comme
 * texte et ignoré silencieusement par SUM().
 */
function toNumber_(v, dflt) {
  if (typeof v === 'number') return isFinite(v) ? v : dflt;
  if (v === null || v === undefined) return dflt;
  var s = String(v).replace(/[\s  ]/g, '').replace(',', '.');
  if (s === '') return dflt;
  var x = Number(s);
  return isFinite(x) ? x : dflt;
}

/** Taux : accepte 0.1, "0,1", "10" ou "10%" → 0.1 */
function toRate_(v, dflt) {
  if (v === null || v === undefined || String(v).trim() === '') return dflt;
  var s = String(v).replace('%', '');
  var x = toNumber_(s, dflt);
  return x > 1 ? x / 100 : x;
}

/** Nom unique, en ignorant la feuille en cours de renommage elle-même. */
function uniqueSheetName_(ss, sheet, baseName) {
  var name = baseName;
  var i = 2;
  var existing;
  while ((existing = ss.getSheetByName(name)) &&
         existing.getSheetId() !== sheet.getSheetId()) {
    name = (baseName + ' (' + i + ')').substring(0, 99);
    i++;
  }
  return name;
}

function columnToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - mod) / 26);
  }
  return letter;
}

/** Accepte "yyyy-mm-dd" (input type=date) ou "dd/mm/yyyy". Renvoie {y,m,d} ou null. */
function parseDateInput_(str) {
  if (!str) return null;
  var s = String(str).trim();
  var y, mo, d, m;
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s))) {
    y = +m[1]; mo = +m[2]; d = +m[3];
  } else if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s))) {
    d = +m[1]; mo = +m[2]; y = +m[3];
  } else {
    return null;
  }
  var test = new Date(y, mo - 1, d);
  if (test.getFullYear() !== y || test.getMonth() !== mo - 1 || test.getDate() !== d) {
    return null;  // ex. 31/02/2026
  }
  return { y: y, m: mo, d: d };
}

function scriptTimeZone_() {
  return Session.getScriptTimeZone() || 'Africa/Ouagadougou';
}

function formatDateStamp_() {
  return Utilities.formatDate(new Date(), scriptTimeZone_(), 'yyyy-MM-dd_HHmm');
}
