/**
 * Party annuel DFuse 2026 — réception des RSVP
 * --------------------------------------------
 * - Enregistre chaque réponse dans l'onglet « RSVP » de la Google Sheet.
 * - Si le même courriel répond à nouveau, la ligne est mise à jour (pas de doublon).
 * - Envoie un courriel de confirmation à l'invité.
 * - Envoie (optionnel) un avis à l'organisateur.
 * - L'onglet « Résumé » calcule les totaux (personnes, boissons, restrictions).
 *
 * Installation : voir README.md (section « Google Sheet + Apps Script »).
 */

const CONFIG = {
  NOTIFY_EMAIL: '',            // ex. 'party@dfuse.ca' — laisse vide pour ne pas recevoir d'avis
  SENDER_NAME: 'DFuse',
  REPLY_TO: '',                // adresse où les invités peuvent répondre (vide = compte Google qui déploie)
  SHEET_NAME: 'RSVP',
  SUMMARY_NAME: 'Résumé',
  EVENT: {
    title: 'Party annuel DFuse 2026',
    dateText: 'Vendredi 11 décembre 2026, portes à 18h',
    address: '34, rue Blainville Ouest, Sainte-Thérèse, QC',
    dress: 'Casual propre',
    mapsUrl: 'https://www.google.com/maps/search/?api=1&query=34+rue+Blainville+Ouest,+Sainte-Th%C3%A9r%C3%A8se,+QC',
    gcalUrl: 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=Party+annuel+DFuse+2026&dates=20261211T180000/20261211T230000&ctz=America/Toronto&location=34+rue+Blainville+Ouest,+Sainte-Th%C3%A9r%C3%A8se,+QC'
  }
};

const HEADERS = [
  'Reçu le', 'Mis à jour le', 'Présence', 'Prénom', 'Nom', 'Courriel', 'Téléphone',
  'Accompagnateurs', 'Noms accompagnateurs', 'Total personnes', 'Boissons', 'Restrictions', 'Message'
];

/** À exécuter une fois après avoir collé le script (menu Exécuter > setup). */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getRsvpSheet_();
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#121212').setFontColor('#ffffff');
  sh.setColumnWidths(1, HEADERS.length, 150);

  // Recalcule le résumé automatiquement à chaque modification de la feuille
  // (ajout, modification ou suppression de lignes).
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshSummary').forSpreadsheet(ss).onChange().create();

  refreshSummary();
  // Demande les autorisations (Sheets + courriel) d'un coup.
  MailApp.getRemainingDailyQuota();
}

/** Menu « Party DFuse » dans la feuille. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Party DFuse')
    .addItem('Actualiser le résumé', 'refreshSummary')
    .addToUi();
}

/**
 * Calcule les totaux dans le script et écrit des valeurs (pas de formules) :
 * fonctionne peu importe la langue de la feuille.
 */
function refreshSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getRsvpSheet_();
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
  const rowsData = data.filter(function (r) { return String(r[5]).trim() !== ''; });
  const yes = rowsData.filter(function (r) { return r[2] === 'Oui'; });
  const sumOf = function (list, col) { return list.reduce(function (a, r) { return a + (Number(r[col]) || 0); }, 0); };
  const drink = function (name) { return yes.filter(function (r) { return String(r[10]).indexOf(name) !== -1; }).length; };

  const rows = [
    ['Indicateur', 'Valeur'],
    ['Réponses reçues', rowsData.length],
    ['Invités qui viennent', yes.length],
    ['Accompagnateurs', sumOf(yes, 7)],
    ['TOTAL PERSONNES ATTENDUES', sumOf(yes, 9)],
    ['Absents', rowsData.filter(function (r) { return r[2] === 'Non'; }).length],
    ['', ''],
    ['Boissons (nb de réponses « Oui » qui ont coché)', ''],
    ['Vin', drink('Vin')],
    ['Bulles', drink('Bulles')],
    ['Bière', drink('Bière')],
    ['Cocktails', drink('Cocktails')],
    ['Sans alcool', drink('Sans alcool')],
    ['', ''],
    ['Réponses avec restrictions alimentaires', yes.filter(function (r) { return String(r[11]).trim() !== ''; }).length],
    ['', ''],
    ['Dernière mise à jour', new Date()]
  ];

  const sum = ss.getSheetByName(CONFIG.SUMMARY_NAME) || ss.insertSheet(CONFIG.SUMMARY_NAME);
  sum.clear();
  sum.getRange(1, 1, rows.length, 2).setValues(rows);
  sum.getRange('A1:B1').setFontWeight('bold').setBackground('#121212').setFontColor('#ffffff');
  sum.getRange('A5:B5').setFontWeight('bold').setBackground('#c8a46a');
  sum.getRange('A8').setFontWeight('bold');
  sum.getRange('A17:B17').setFontColor('#888888');
  sum.setColumnWidth(1, 360);
  sum.setColumnWidth(2, 160);
}

function doGet() {
  return json_({ ok: true, service: 'RSVP DFuse 2026' });
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.site) return json_({ ok: true }); // pot de miel anti-spam

  const email = String(p.courriel || '').trim().toLowerCase();
  const prenom = clean_(p.prenom), nom = clean_(p.nom);
  if (!prenom || !nom || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
    return json_({ ok: false, error: 'Champs requis manquants' });
  }
  const yes = p.presence === 'Oui';
  const guests = yes ? Math.max(0, Math.min(3, parseInt(p.accompagnateurs, 10) || 0)) : 0;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let updated = false;
  try {
    const sh = getRsvpSheet_();
    const now = new Date();
    const row = [
      now, now, yes ? 'Oui' : 'Non', prenom, nom, email, clean_(p.telephone),
      guests, yes ? clean_(p.noms_accompagnateurs) : '', yes ? 1 + guests : 0,
      yes ? clean_(p.boissons) : '', yes ? clean_(p.restrictions) : '', clean_(p.message)
    ];
    const last = sh.getLastRow();
    let found = -1;
    if (last > 1) {
      const emails = sh.getRange(2, 6, last - 1, 1).getValues();
      for (let i = 0; i < emails.length; i++) {
        if (String(emails[i][0]).toLowerCase() === email) { found = i + 2; break; }
      }
    }
    if (found > 0) {
      row[0] = sh.getRange(found, 1).getValue(); // garde la date de première réponse
      sh.getRange(found, 1, 1, row.length).setValues([row]);
      updated = true;
    } else {
      sh.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }

  try { refreshSummary(); } catch (err) { console.error('Résumé', err); }

  try { sendGuestEmail_(email, prenom, yes, guests, p, updated); } catch (err) { console.error('Courriel invité', err); }
  try { sendOrganizerEmail_(prenom, nom, email, yes, guests, p, updated); } catch (err) { console.error('Avis organisateur', err); }

  return json_({ ok: true, updated: updated });
}

/* ---------- Courriels ---------- */

function sendGuestEmail_(to, prenom, yes, guests, p, updated) {
  const ev = CONFIG.EVENT;
  const subject = yes
    ? `${updated ? 'Mise à jour : ' : ''}C'est confirmé! On se voit le 11 décembre 🥂`
    : 'Merci pour ta réponse — Party DFuse 2026';

  const detailsRows = yes ? `
      ${tr_('Quand', ev.dateText)}
      ${tr_('Où', `<a href="${ev.mapsUrl}" style="color:#c8a46a">${ev.address}</a>`)}
      ${tr_('Code vestimentaire', ev.dress)}
      ${tr_('Ta réservation', `${1 + guests} personne${guests ? 's' : ''}${p.noms_accompagnateurs ? ' (' + esc_(p.noms_accompagnateurs) + ')' : ''}`)}
      ${p.restrictions ? tr_('Restrictions notées', esc_(p.restrictions)) : ''}` : '';

  const body = yes ? `
      <p style="font-size:16px;line-height:1.6;margin:0 0 16px">Salut ${esc_(prenom)},</p>
      <p style="font-size:16px;line-height:1.6;margin:0 0 20px">${updated ? 'Ta réponse a bien été mise à jour.' : 'Ta présence est confirmée!'} On a vraiment hâte de trinquer avec toi pour notre 3<sup>e</sup> party annuel.</p>
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 0 24px">${detailsRows}</table>
      <p style="font-size:15px;line-height:1.6;margin:0 0 8px"><strong>Au menu :</strong> tapas et bouchées du restaurant Tartaro, bar complet avec notre sommelier invité de Fiola Import, expo numérique 2026, notre podcast annuel et une salle VIP pour une dégustation signature.</p>
      <p style="margin:24px 0"><a href="${ev.gcalUrl}" style="background:#c8a46a;color:#0b0b0b;text-decoration:none;font-weight:700;padding:12px 22px;border-radius:999px;display:inline-block">Ajouter à mon agenda</a></p>
      <p style="font-size:14px;line-height:1.6;color:#666;margin:0">Un changement? Remplis à nouveau le formulaire avec le même courriel, ou réponds simplement à ce message.</p>`
    : `
      <p style="font-size:16px;line-height:1.6;margin:0 0 16px">Salut ${esc_(prenom)},</p>
      <p style="font-size:16px;line-height:1.6;margin:0 0 16px">Merci de nous avoir avisés! Tu vas nous manquer cette année. Si jamais tes plans changent, tu peux remplir le formulaire à nouveau avec le même courriel.</p>
      <p style="font-size:16px;line-height:1.6;margin:0">Merci d'avoir fait partie de notre belle année 2026 💙</p>`;

  const html = `
  <div style="background:#f3f4f6;padding:24px 12px;font-family:Montserrat,Arial,sans-serif;color:#121212">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden">
      <tr><td style="background:#121212;padding:28px 28px 24px">
        <div style="color:#c8a46a;font-size:12px;letter-spacing:.18em;text-transform:uppercase;font-weight:700">Party annuel · 3e édition</div>
        <div style="color:#ffffff;font-size:34px;line-height:1.05;font-weight:800;margin-top:8px;font-family:'Bebas Neue',Impact,Arial,sans-serif;letter-spacing:.02em">DFUSE 2026 🥂</div>
      </td></tr>
      <tr><td style="padding:28px">${body}</td></tr>
      <tr><td style="padding:18px 28px;border-top:1px solid #eee;font-size:12px;color:#888">DFuse — Studio de création publicitaire · <a href="https://www.dfuse.ca" style="color:#888">dfuse.ca</a></td></tr>
    </table>
  </div>`;

  const opts = { to: to, subject: subject, htmlBody: html, name: CONFIG.SENDER_NAME };
  if (CONFIG.REPLY_TO) opts.replyTo = CONFIG.REPLY_TO;
  MailApp.sendEmail(opts);
}

function sendOrganizerEmail_(prenom, nom, email, yes, guests, p, updated) {
  if (!CONFIG.NOTIFY_EMAIL) return;
  const subject = `[RSVP${updated ? ' maj' : ''}] ${prenom} ${nom} — ${yes ? 'OUI (' + (1 + guests) + ' pers.)' : 'NON'}`;
  const lines = [
    `Présence : ${yes ? 'Oui' : 'Non'}`,
    `Nom : ${prenom} ${nom}`,
    `Courriel : ${email}`,
    `Téléphone : ${clean_(p.telephone)}`,
    yes ? `Accompagnateurs : ${guests}${p.noms_accompagnateurs ? ' (' + clean_(p.noms_accompagnateurs) + ')' : ''}` : '',
    yes && p.boissons ? `Boissons : ${clean_(p.boissons)}` : '',
    yes && p.restrictions ? `Restrictions : ${clean_(p.restrictions)}` : '',
    p.message ? `Message : ${clean_(p.message)}` : ''
  ].filter(Boolean);
  lines.push('', `Feuille : ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
  MailApp.sendEmail({ to: CONFIG.NOTIFY_EMAIL, subject: subject, body: lines.join('\n'), name: 'RSVP DFuse' });
}

/* ---------- Utilitaires ---------- */

function getRsvpSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME, 0);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}
function tr_(k, v) {
  return `<tr><td style="padding:10px 0;border-bottom:1px solid #eee;font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#888;width:42%;vertical-align:top">${k}</td><td style="padding:10px 0;border-bottom:1px solid #eee;font-size:15px;font-weight:600">${v}</td></tr>`;
}
function clean_(v) {
  // Force le texte quand la valeur commence par un caractère de formule (ex. « +1 514… »).
  const s = String(v || '').trim().slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function esc_(v) {
  return String(v || '').replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
