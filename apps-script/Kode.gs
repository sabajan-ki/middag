/**
 * Middagsplanen – felles lagring i Google Regneark.
 *
 * Siden (middag.html) sender hele planen hit som JSON. Scriptet lagrer den i arket «_data»
 * og skriver en lesbar kopi av ukeplanen i arket «Ukeplan».
 *
 * Oppsett (én gang):
 *   1. Kjør funksjonen «oppsett» (godkjenn tilgang). Loggen viser regnearket og nøkkelen.
 *   2. Distribuer → Ny distribusjon → Nettapp. Kjør som: meg. Hvem har tilgang: Alle.
 *   3. Lim inn nettapp-adressen og nøkkelen under Oppsett i middagsplanen.
 */

const CHUNK = 45000; // en celle tåler 50 000 tegn; planen deles over flere celler

function oppsett() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('Middagsplanen – felles lagring');
    id = ss.getId();
    props.setProperty('SHEET_ID', id);
  }
  let key = props.getProperty('KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('KEY', key);
  }
  Logger.log('Regneark: https://docs.google.com/spreadsheets/d/' + id);
  Logger.log('Nøkkel: ' + key);
}

function doGet(e) {
  return svar(behandle(e.parameter || {}));
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return svar({ error: 'Ugyldig forespørsel' }); }
  return svar(behandle(body));
}

function behandle(p) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KEY')) return { error: 'Kjør oppsett-funksjonen først' };
  if (p.k !== props.getProperty('KEY')) return { error: 'Feil nøkkel' };

  if (p.action === 'get') return { state: les() };

  if (p.action === 'set') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const cur = les();
      // Har den andre lagret siden denne telefonen sist hentet, sendes den nyeste planen tilbake
      if (cur && p.base !== undefined && p.base !== null && cur.v !== p.base) return { conflict: true, state: cur };
      const state = p.state || {};
      state.v = (cur && cur.v ? cur.v : 0) + 1;
      skriv(state);
      if (p.readable && p.readable.ukeplan) speil(p.readable.ukeplan);
      return { ok: true, v: state.v };
    } finally {
      lock.releaseLock();
    }
  }
  return { error: 'Ukjent handling' };
}

function ark(navn) {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  return ss.getSheetByName(navn) || ss.insertSheet(navn);
}

function les() {
  const sh = ark('_data');
  const n = sh.getLastRow();
  if (n < 1) return null;
  // hver celle starter med «~» så Regneark aldri tolker innholdet som formel eller tall
  const txt = sh.getRange(1, 1, n, 1).getValues().map(r => String(r[0]).replace(/^~/, '')).join('');
  if (!txt) return null;
  try { return JSON.parse(txt); } catch (err) { return null; }
}

function skriv(state) {
  const sh = ark('_data');
  const txt = JSON.stringify(state);
  const deler = [];
  for (let i = 0; i < txt.length; i += CHUNK) deler.push(['~' + txt.slice(i, i + CHUNK)]);
  sh.clear();
  sh.getRange(1, 1, deler.length, 1).setNumberFormat('@').setValues(deler);
}

function speil(rader) {
  const sh = ark('Ukeplan');
  sh.clear();
  sh.getRange(1, 1, 1, 3).setValues([['Uke', 'Dag', 'Middag']]).setFontWeight('bold');
  if (rader.length) sh.getRange(2, 1, rader.length, 3).setValues(rader);
  sh.autoResizeColumns(1, 3);
}

function svar(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
