/**
 * RSVP backend — Google Apps Script web app.
 * Reads/writes the "Lista de Invitados" spreadsheet.
 *
 * SETUP (one time, ~5 min):
 *  1. Open the guest-list spreadsheet → Extensions → Apps Script.
 *  2. Paste this whole file into Code.gs (replace the default content).
 *  3. Set GUESTS_SHEET below to the exact name of the tab that holds the list.
 *  4. Label each household in the "Familia" column (e.g. "Murcia-Espitia").
 *     People sharing a Familia label get ONE shared code; blank = own code.
 *  5. Run assignCodes() once (Run ▶ button; authorize when asked).
 *     It adds the "Codigo" column if missing and fills empty codes.
 *  6. Deploy → New deployment → type: Web app →
 *     Execute as: Me · Who has access: Anyone → Deploy.
 *  7. Copy the web app URL into RSVP_ENDPOINT in index.html.
 *
 * Re-running assignCodes() is safe: existing codes are never overwritten.
 * Each RSVP submission appends rows to a "Respuestas RSVP" tab (audit log)
 * and updates the "Va o no" column in the guest list.
 */

const CONFIG = {
  GUESTS_SHEET: 'Final',
  RESPONSES_SHEET: 'Respuestas RSVP',
  COL_NAME: 'Nombre',
  COL_FAMILY: 'Grupo',      // big group (Lopez, Mila, …) — informational only
  COL_ATTENDING: 'Va o no',
  COL_HOUSEHOLD: 'Familia', // household/invitation unit — one shared code
  COL_CODE: 'Codigo',
};

// ---------- helpers ----------

function sheet_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.GUESTS_SHEET);
  if (!s) throw new Error('Sheet not found: ' + CONFIG.GUESTS_SHEET + ' — fix CONFIG.GUESTS_SHEET');
  return s;
}

// Finds the header row (the row containing "Nombre") and returns
// {row, cols: {name, family, attending, household, code}} with 1-based indexes (0 = missing).
function headers_(sheet) {
  const scan = sheet.getRange(1, 1, Math.min(5, sheet.getLastRow()), sheet.getLastColumn()).getValues();
  for (let r = 0; r < scan.length; r++) {
    const idx = {};
    scan[r].forEach((v, c) => { idx[String(v).trim().toLowerCase()] = c + 1; });
    if (idx[CONFIG.COL_NAME.toLowerCase()]) {
      return {
        row: r + 1,
        cols: {
          name: idx[CONFIG.COL_NAME.toLowerCase()] || 0,
          family: idx[CONFIG.COL_FAMILY.toLowerCase()] || 0,
          attending: idx[CONFIG.COL_ATTENDING.toLowerCase()] || 0,
          household: idx[CONFIG.COL_HOUSEHOLD.toLowerCase()] || 0,
          code: idx[CONFIG.COL_CODE.toLowerCase()] || 0,
        },
      };
    }
  }
  throw new Error('Header row with "' + CONFIG.COL_NAME + '" not found in first 5 rows');
}

const CODE_ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; // no 0/O/1/I/L

function newCode_(taken) {
  while (true) {
    let c = '';
    for (let i = 0; i < 6; i++) c += CODE_ALPHABET[Math.floor(Math.random() * CODE_ALPHABET.length)];
    if (!taken.has(c)) { taken.add(c); return c; }
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- run once from the editor ----------

function assignCodes() {
  const sheet = sheet_();
  let h = headers_(sheet);
  // Add missing Hogar / Codigo header cells at the end of the header row.
  let lastCol = sheet.getLastColumn();
  if (!h.cols.household) sheet.getRange(h.row, ++lastCol).setValue(CONFIG.COL_HOUSEHOLD);
  if (!h.cols.code) sheet.getRange(h.row, ++lastCol).setValue(CONFIG.COL_CODE);
  h = headers_(sheet);

  const nRows = sheet.getLastRow() - h.row;
  if (nRows < 1) return;
  const names = sheet.getRange(h.row + 1, h.cols.name, nRows, 1).getValues();
  const households = sheet.getRange(h.row + 1, h.cols.household, nRows, 1).getValues();
  const codes = sheet.getRange(h.row + 1, h.cols.code, nRows, 1).getValues();

  const taken = new Set(codes.map(r => String(r[0]).trim()).filter(Boolean));
  const byHousehold = {}; // existing or new code per household label
  codes.forEach((r, i) => {
    const hh = String(households[i][0]).trim();
    if (hh && String(r[0]).trim()) byHousehold[hh] = String(r[0]).trim();
  });

  let assigned = 0;
  codes.forEach((r, i) => {
    if (String(r[0]).trim() || !String(names[i][0]).trim()) return; // has code / empty row
    const hh = String(households[i][0]).trim();
    if (hh) {
      if (!byHousehold[hh]) byHousehold[hh] = newCode_(taken);
      r[0] = byHousehold[hh];
    } else {
      r[0] = newCode_(taken);
    }
    assigned++;
  });
  sheet.getRange(h.row + 1, h.cols.code, nRows, 1).setValues(codes);
  Logger.log('Assigned ' + assigned + ' codes');
}

// ---------- web app ----------

function guestsForCode_(code) {
  const sheet = sheet_();
  const h = headers_(sheet);
  if (!h.cols.code) return [];
  const nRows = sheet.getLastRow() - h.row;
  if (nRows < 1) return [];
  const data = sheet.getRange(h.row + 1, 1, nRows, sheet.getLastColumn()).getValues();
  const out = [];
  data.forEach((row, i) => {
    if (String(row[h.cols.code - 1]).trim().toUpperCase() === code) {
      out.push({
        row: h.row + 1 + i,
        name: String(row[h.cols.name - 1]).trim(),
        family: h.cols.family ? String(row[h.cols.family - 1]).trim() : '',
      });
    }
  });
  return out;
}

function doGet(e) {
  if (e.parameter.ping) return json_({ ok: true, ping: 'pong' });
  const code = String(e.parameter.code || '').trim().toUpperCase();
  if (!code) return json_({ ok: false, error: 'missing_code' });
  const guests = guestsForCode_(code);
  if (!guests.length) return json_({ ok: false, error: 'unknown_code' });
  return json_({ ok: true, code: code, guests: guests.map(g => ({ name: g.name, family: g.family })) });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  const code = String(body.code || '').trim().toUpperCase();
  const guests = guestsForCode_(code);
  if (!guests.length) return json_({ ok: false, error: 'unknown_code' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let resp = ss.getSheetByName(CONFIG.RESPONSES_SHEET);
  if (!resp) {
    resp = ss.insertSheet(CONFIG.RESPONSES_SHEET);
    resp.appendRow(['Timestamp', 'Codigo', 'Nombre', 'Asiste', 'Menú', 'Restricciones',
      'Niños (edades)', 'Cuidado niños', 'Acompañantes', 'Contacto', 'Gmail álbum', 'Notas', 'Idioma']);
  }

  const sheet = sheet_();
  const h = headers_(sheet);
  const now = new Date();
  (body.responses || []).forEach(r => {
    const guest = guests.find(g => g.name === r.name);
    if (!guest) return; // only accept names that belong to this code
    resp.appendRow([now, code, guest.name, r.attending ? 'Sí' : 'No', r.meal || '',
      body.dietary || '', body.kids || '', body.childcare ? 'Sí' : 'No',
      String(body.buddy || ''), body.contact || '', body.album || '', body.notes || '', body.lang || '']);
    if (h.cols.attending) {
      sheet.getRange(guest.row, h.cols.attending).setValue(r.attending ? 'Sí' : 'No');
    }
  });
  try { sendConfirmation_(body, guests); } catch (err) { /* email failure must not break the RSVP */ }
  return json_({ ok: true });
}

// ---------- confirmation email ----------

const SITE_URL = 'https://juan43ramirez.github.io/pao_y_juank/';
const MAPS_URL = 'https://maps.google.com/?q=Cielo+Alto+Centro+de+Eventos+Medellin';

function isEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim()); }

function icsInvite_() {
  // Mass starts 17:00 America/Bogota (UTC-5) = 22:00Z; party ends midnight = 05:00Z next day.
  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  return ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Pao y Juank//Boda//ES', 'BEGIN:VEVENT',
    'UID:boda-pao-juank-20261213', 'DTSTAMP:' + stamp,
    'DTSTART:20261213T220000Z', 'DTEND:20261214T050000Z',
    'SUMMARY:Boda Paola & Juan Camilo',
    'LOCATION:Centro de Eventos Cielo Alto\\, Carrera 9E #16A Sur-104\\, Medellín',
    'DESCRIPTION:' + SITE_URL, 'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
}

function sendConfirmation_(body, guests) {
  const to = isEmail_(body.contact) ? String(body.contact).trim()
           : isEmail_(body.album) ? String(body.album).trim() : '';
  if (!to) return;
  const es = String(body.lang || 'es') !== 'en';
  const t = (a, b) => es ? a : b;

  const names = (body.responses || [])
    .filter(r => guests.some(g => g.name === r.name))
    .map(r => '  • ' + r.name + ' — ' + (r.attending ? t('Sí', 'Yes') : 'No'));

  const lines = [
    t('¡Recibimos su confirmación! Esto fue lo que nos llegó:',
      'We got your RSVP! Here is what you sent us:'), '', names.join('\n')];
  if (body.dietary) lines.push('', t('Restricciones: ', 'Dietary: ') + body.dietary);
  if (body.kids) lines.push(t('Niños: ', 'Children: ') + body.kids +
    (body.childcare ? t(' (usarán el cuidado infantil)', ' (will use the childcare)') : ''));
  lines.push('',
    t('Si algo cambia, ingresen su código (' + body.code + ') de nuevo en la página y reenvíen el formulario, o escríbannos al WhatsApp +57 300 4122523.',
      'If anything changes, enter your code (' + body.code + ') again on the site and resubmit, or message us on WhatsApp +57 300 4122523.'), '',
    '🗓 ' + t('Adjuntamos la invitación para su calendario.', 'Calendar invite attached.'),
    '📍 Centro de Eventos Cielo Alto — ' + MAPS_URL);
  if (isEmail_(body.album)) lines.push('📸 ' +
    t('Les llegará la invitación al álbum compartido antes de la boda.',
      'Your invite to the shared photo album will arrive before the wedding.'));
  lines.push('', t('Con cariño,', 'With love,'), 'Paola & Juan Camilo', SITE_URL);

  MailApp.sendEmail(to,
    t('¡Recibimos su confirmación! · Paola & Juan Camilo — 13 Dic 2026',
      'RSVP received! · Paola & Juan Camilo — Dec 13, 2026'),
    lines.join('\n'),
    { name: 'Paola & Juan Camilo',
      attachments: [Utilities.newBlob(icsInvite_(), 'text/calendar', 'boda-paola-juan.ics')] });
}
