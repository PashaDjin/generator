/**
 * ═══════════════════════════════════════════════════════════════
 * ФАЙЛ 2: ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
 * ═══════════════════════════════════════════════════════════════
 * 
 * Содержит все утилиты и хелперы:
 * • Нормализация адресов и телефонов
 * • Работа с листами и фильтрами
 * • Работа со строками в ОБЪЕКТЫ
 * • Работа с Google Drive (папки, доступы)
 * • Замена плейсхолдеров в документах
 */

// ═══════════════════════════════════════════════════════════════
// НОРМАЛИЗАЦИЯ
// ═══════════════════════════════════════════════════════════════

function normAddr_(s) {
  return String(s || '').trim().replace(/\s+/g, ' ');
}

function normalizeRuPhoneStrict_(input) {
  const digits = String(input || '').replace(/\D+/g, '');
  if (digits.length === 10 && digits.startsWith('9')) return '7' + digits;
  if (digits.length === 11 && digits.startsWith('8')) return '7' + digits.substring(1);
  if (digits.length === 11 && digits.startsWith('7')) return digits;
  throw new Error('Телефон должен быть вида 79XXXXXXXXX (ввод 9/8/7 допустим).');
}

function normalizeRuPhoneSoft_(input) {
  try { return normalizeRuPhoneStrict_(input); } catch(_) { return ''; }
}

function titleCaseRu_(s) {
  return s.split(' ').map(w => {
    if (!w) return w;
    if (/^(ул\.|пр-т|д\.|кв\.|к\.|стр\.)$/i.test(w)) return w.toLowerCase();
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  }).join(' ');
}

function normalizeAddressRu_(input) {
  let s = String(input || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();

  s = s
    .replace(/\b(улица|ул\.?)\b/g, 'ул.')
    .replace(/\b(проспект|пр\-?т|просп\.)\b/g, 'пр-т')
    .replace(/\b(дом|д\.?)\b/g, 'д.')
    .replace(/\b(квартира|кв\.?)\b/g, 'кв.')
    .replace(/\b(корпус|к\.?)\b/g, 'к.')
    .replace(/\b(строение|стр\.?)\b/g, 'стр.');

  s = s.replace(/\s*,\s*/g, ', ').replace(/\s+/g, ' ').trim();

  s = s
    .replace(/\s+(д\.\s*\d+)/g, ', $1')
    .replace(/\s+(кв\.\s*\d+)/g, ', $1')
    .replace(/\s+(к\.\s*\d+)/g, ', $1')
    .replace(/\s+(стр\.\s*\d+)/g, ', $1');

  s = titleCaseRu_(s);

  s = s
    .replace(/\bУл\.\b/g, 'ул.')
    .replace(/\bПр\-Т\b/gi, 'пр-т')
    .replace(/\bД\.\b/g, 'д.')
    .replace(/\bКв\.\b/g, 'кв.')
    .replace(/\bК\.\b/g, 'к.')
    .replace(/\bСтр\.\b/g, 'стр.');

  return s;
}


// ═══════════════════════════════════════════════════════════════
// РАБОТА С ЛИСТАМИ
// ═══════════════════════════════════════════════════════════════

function removeFilterSafe_(sh) {
  try { const f = sh.getFilter && sh.getFilter(); if (f) f.remove(); } catch(_) {}
}

function ensureSheet_(book, name) {
  return book.getSheetByName(name) || book.insertSheet(name);
}

function ensureObjectsSheet_(book) {
  return ensureSheet_(book, OBJECTS_SHEET_NAME);
}

function findFirstEmptyRow_(sheet, col, startRow) {
  const last = Math.max(sheet.getLastRow(), startRow - 1);
  const height = Math.max(1, last - startRow + 2);
  const vals = sheet.getRange(startRow, col, height, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i][0];
    if (v === '' || v === null || (typeof v === 'string' && v.trim() === '')) {
      return startRow + i;
    }
  }
  sheet.insertRowsAfter(last, 1);
  return last + 1;
}

function setIfProvided_(sheet, row, colIndex1, value) {
  if (value === null || value === undefined) return;
  sheet.getRange(row, colIndex1).setValue(value);
}


// ═══════════════════════════════════════════════════════════════
// ПОИСК СТРОК В ОБЪЕКТЫ
// ═══════════════════════════════════════════════════════════════

function findRowByUid_(sheet, uid) {
  const u = String(uid || '').trim();
  if (!u) return 0;
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const rng = sheet.getRange(2, COL.A_UID, last - 1, 1).getValues();
  for (let i = 0; i < rng.length; i++) {
    if (String(rng[i][0] || '').trim() === u) return 2 + i;
  }
  return 0;
}

function findRowsByPhone_(sheet, phoneNorm) {
  const phone = String(phoneNorm || '').trim();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const colVals = sheet.getRange(2, COL.B_PHONE, last - 1, 1).getValues();
  const rows = [];
  for (let i = 0; i < colVals.length; i++) {
    const v = String(colVals[i][0] || '').trim();
    if (v === phone) rows.push(2 + i);
  }
  return rows;
}

function findRowByPhoneAddr_(sheet, phoneNorm, addrNormRaw) {
  const phone = String(phoneNorm || '').trim();
  const addrNorm = normAddr_(addrNormRaw || '');
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const phones = sheet.getRange(2, COL.B_PHONE, last - 1, 1).getValues();
  const addrs  = sheet.getRange(2, COL.C_ADDR,  last - 1, 1).getValues();
  for (let i = 0; i < phones.length; i++) {
    const p = String(phones[i][0] || '').trim();
    const a = normAddr_(addrs[i][0] || '');
    if (p === phone && a === addrNorm) return 2 + i;
  }
  return 0;
}

function findRowByPhoneAndNormalizedAddress_(sheet, phoneNorm, addressNorm) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const phones = sheet.getRange(2, COL.B_PHONE, last - 1, 1).getValues();
  const addrs  = sheet.getRange(2, COL.C_ADDR,  last - 1, 1).getValues();
  for (let i = 0; i < phones.length; i++) {
    const p = String(phones[i][0] || '').trim();
    const a = normAddr_(addrs[i][0] || '');
    if (p === phoneNorm && a === addressNorm) return 2 + i;
  }
  return 0;
}


// ═══════════════════════════════════════════════════════════════
// СОЗДАНИЕ/ОБНОВЛЕНИЕ СТРОК В ОБЪЕКТЫ
// ═══════════════════════════════════════════════════════════════

function createEmptyObjectRowWithUid_(sh, uid, { client, phoneNorm, addrNorm }) {
  if (!uid) throw new Error('UID пуст при создании строки объекта.');
  const row = findFirstEmptyRow_(sh, 1, 2);
  const vals = new Array(MAX_COL_OBJECTS).fill('');
  vals[COL.A_UID-1]    = uid;
  vals[COL.A_CLIENT-1] = client || '';
  vals[COL.B_PHONE-1]  = phoneNorm || '';
  vals[COL.C_ADDR-1]   = addrNorm || '';
  sh.getRange(row, 1, 1, MAX_COL_OBJECTS).setValues([vals]);
  return row;
}

function writeSmetaObjectWithUid_(smetaSpreadsheetId, payload) {
  const { uid, client, phoneNorm, address, measurer, company } = payload;
  const smetaSs = SpreadsheetApp.openById(smetaSpreadsheetId);
  const sh = smetaSs.getSheetByName(SHEET_SMETA_OBJECT);
  if (!sh) throw new Error('В шаблоне нет листа "🏠 Характеристики объекта".');

  sh.getRange('B1').setValue(uid);
  sh.getRange('B2').setValue(client);
  sh.getRange('B3').setValue(phoneNorm);
  sh.getRange('B5').setValue(address);
  sh.getRange('B6').setValue(measurer);
  sh.getRange('B7').setValue(company);
}

function upsertContractObjectByUid_({
  sheet, uid, client, phoneNorm, addressNorm, manager, legalEntity, sumTotal, contractUrl, folderUrl, mail
}) {
  removeFilterSafe_(sheet);
  let row = findRowByUid_(sheet, uid);
  if (!row) {
    row = createEmptyObjectRowWithUid_(sheet, uid, {
      client: client || '',
      phoneNorm: phoneNorm || '',
      addrNorm: addressNorm || ''
    });
  }
  sheet.getRange(row, COL.D_STATUS).setValue(STATUS_CONTRACT);
  sheet.getRange(row, COL.C_ADDR).setValue(addressNorm);
  if (client)  sheet.getRange(row, COL.A_CLIENT).setValue(client);
  if (manager) sheet.getRange(row, COL.F_MANAGER).setValue(manager);
  if (legalEntity) sheet.getRange(row, COL.P_LEGAL).setValue(legalEntity);
  if (mail) sheet.getRange(row, COL.R_MAIL).setValue(mail);
  if (sumTotal !== null && sumTotal !== undefined && sumTotal !== '') {
    sheet.getRange(row, COL.I_SUM).setValue(sumTotal);
  }
  if (contractUrl) sheet.getRange(row, COL.N_CONTRACT_LINK).setValue(contractUrl);
  if (folderUrl)   sheet.getRange(row, COL.O_FOLDER).setValue(folderUrl);
  return row;
}


// ═══════════════════════════════════════════════════════════════
// РАБОТА С GOOGLE DRIVE
// ═══════════════════════════════════════════════════════════════

function findOrCreateClientFolderByUID_(managerRootFolderId, uid, addressNorm, fullname) {
  const parent = DriveApp.getFolderById(managerRootFolderId);
  const prefix = String(uid).trim() + ' - ';
  const addr   = String(addressNorm || '').trim();
  const it = parent.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    const name = String(f.getName() || '').trim();
    if (!name.startsWith(prefix)) continue;
    const rest = name.substring(prefix.length);
    const parts = rest.split(' - ');
    if (parts.length < 2) continue;
    if (normAddr_(parts[0]) === addr) return f;
  }
  return parent.createFolder(`${uid} - ${addr} - ${fullname}`);
}

function findClientFolderByRawAddressAndRename_(rootFolderId, inits, rawAddress) {
  const parent = DriveApp.getFolderById(String(rootFolderId).trim());
  const suffix = ' - ' + String(rawAddress || '').trim();
  const exactRaw = String(rawAddress || '').trim();
  let candidate = null;

  const it = parent.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    const name = String(f.getName() || '').trim();
    if (name.endsWith(suffix) || name === exactRaw) { candidate = f; break; }
  }

  const desiredName = `${inits} - ${exactRaw}`;
  if (candidate) {
    if (candidate.getName() !== desiredName) candidate.setName(desiredName);
    return candidate;
  }
  return parent.createFolder(desiredName);
}

function addEditors_(file, emails) {
  (emails || []).filter(Boolean).forEach(e => {
    try { file.addEditor(e); } 
    catch (err) { throw new Error(`Не выдал доступ ${e}: ${err}`); }
  });
}


// ═══════════════════════════════════════════════════════════════
// РАБОТА С ДОКУМЕНТАМИ (ЗАМЕНА ПЛЕЙСХОЛДЕРОВ)
// ═══════════════════════════════════════════════════════════════

function parsePrice(s) {
  const t = String(s || '').replace(/\s+/g, '').replace(',', '.');
  const n = parseFloat(t);
  return isFinite(n) ? n : 0;
}

function escapeRegex(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function findAndHighlight(section, pattern, color) {
  let r = section.findText(pattern);
  while (r) {
    const el = r.getElement().asText();
    el.setBackgroundColor(r.getStartOffset(), r.getEndOffsetInclusive(), color);
    r = section.findText(pattern, r);
  }
}

function replaceWithOptionalHighlight(docId, map, missingSet) {
  const doc = DocumentApp.openById(docId);
  const parts = [];
  const body = doc.getBody();
  const header = doc.getHeader();
  const footer = doc.getFooter();
  if (body) parts.push(body);
  if (header) parts.push(header);
  if (footer) parts.push(footer);

  Object.keys(map || {}).forEach(k => {
    const v = map[k];
    if (v === undefined || v === null || String(v).trim() === '') return;
    const pat = escapeRegex(k);
    parts.forEach(p => p.replaceText(pat, String(v)));
  });

  if (missingSet && missingSet.size) {
    const color = '#fff59d';
    missingSet.forEach(k => {
      const pat = escapeRegex(k);
      parts.forEach(p => findAndHighlight(p, pat, color));
    });
  }

  doc.saveAndClose();
}

function mapCodeToLegal_(code, fallback) {
  const c = String(code || '').toUpperCase();
  if (fallback) return fallback;
  const map = {
    'D': 'ООО «Компания»',
    'K': 'ИП Фамилия Имя Отчество',
  };
  return map[c] || '';
}


// ═══════════════════════════════════════════════════════════════
// ПРОЧИЕ УТИЛИТЫ
// ═══════════════════════════════════════════════════════════════

function val(v){ 
  if(v===null||v===undefined) return ''; 
  return (typeof v==='string')? v.trim(): String(v).trim(); 
}

function requireFilled(v, label){ 
  if(v===null||v===undefined||String(v).trim()==='') 
    throw new Error('Заполни '+label+'.'); 
}

function toDDMMYYYY_(v) {
  const d = toDateSafe_(v);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy') : '';
}

function toDateSafe_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}

function addDays_(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

function parseDriveIdFromUrl_(url) {
  const u = String(url || '').trim();
  const m = u.match(/\/d\/([A-Za-z0-9_-]{20,})\//) ||
            u.match(/[?&]id=([A-Za-z0-9_-]{20,})\b/) ||
            u.match(/file\/d\/([A-Za-z0-9_-]{20,})\b/);
  return m ? m[1] : '';
}
