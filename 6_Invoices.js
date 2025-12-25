/************************************
 *   ГЕНЕРАТОР — РАБОТА СО СЧЕТАМИ (T-BANK API)
 ************************************/

function getTokensForBrand_(brandUpper) {
  const up = String(brandUpper || '').toUpperCase();
  const props = PropertiesService.getScriptProperties();
  const brProp = (props.getProperty('TBANK_BRANDMAR') || '').trim();
  const smProp = (props.getProperty('TBANK_STROYMAT') || '').trim();

  const apiBrandmar = brProp;
  const apiStroymat = smProp;

  if (up.includes('ООО «БРЕНДМАР»')) {
    if (!apiBrandmar) throw new Error('Токен TBANK_BRANDMAR не задан в свойствах скрипта.');
    if (typeof ACCOUNT_BRANDMAR === 'undefined' || !ACCOUNT_BRANDMAR) {
      throw new Error('ACCOUNT_BRANDMAR не задан.');
    }
    return { apiToken: apiBrandmar, accountNumber: ACCOUNT_BRANDMAR };
  }

  if (up.includes('ООО «СТРОЙМАТ»')) {
    if (!apiStroymat) throw new Error('Токен TBANK_STROYMAT не задан в свойствах скрипта.');
    if (typeof ACCOUNT_STROYMAT === 'undefined' || !ACCOUNT_STROYMAT) {
      throw new Error('ACCOUNT_STROYMAT не задан.');
    }
    return { apiToken: apiStroymat, accountNumber: ACCOUNT_STROYMAT };
  }

  throw new Error('Компания в G1 не распознана. Ожидалось: Брендмар / Строймат.');
}

function collectInvoiceData_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INVOICES);
  if (!sh) throw new Error(`Лист "${SHEET_INVOICES}" не найден.`);

  const E1v = sh.getRange('E1').getValue();
  if (!(E1v instanceof Date)) throw new Error('В E1 должна быть дата (тип Date).');

  const dd = Utilities.formatDate(new Date(E1v), Session.getScriptTimeZone(), 'dd');
  const mm = Utilities.formatDate(new Date(E1v), Session.getScriptTimeZone(), 'MM');

  const brandCell = String(sh.getRange('G1').getDisplayValue()).trim();
  const { apiToken, accountNumber } = getTokensForBrand_(brandCell);

  const payerName = String(sh.getRange('G2').getDisplayValue()).trim();
  if (!payerName) throw new Error('Заполни ФИО плательщика (G2).');

  const email = String(sh.getRange('G3').getDisplayValue()).trim();
  if (!email || !email.includes('@')) throw new Error('Заполни корректный email (G3).');

  const phoneRaw = String(sh.getRange('G4').getDisplayValue()).trim();
  if (!phoneRaw) throw new Error('Заполни телефон (G4).');
  const phoneNorm = normalizeRuPhoneStrict_(phoneRaw);

  const adresRaw = String(sh.getRange('G5').getDisplayValue()).trim();
  if (!adresRaw) throw new Error('Заполни адрес объекта (G5).');
  const addressNorm = normAddr_(adresRaw);

  const aktId = String(sh.getRange('G6').getDisplayValue()).trim();
  if (!aktId) throw new Error('Заполни номер акта или тип ("Материалы"/"Дизайн") (G6).');

  const last4 = (phoneNorm || '').slice(-4);
  if (last4.length !== 4) throw new Error('Не удалось получить последние 4 цифры телефона.');
  const invoiceNumber = String(Number(`${dd}${mm}${last4}`));

  const invoiceDate = new Date(E1v);
  const dueDate     = addDays_(invoiceDate, 3);
  const invoiceDateISO = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const dueDateISO     = Utilities.formatDate(dueDate,    Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const items = readItemsFromActSheet_(sh);
  if (!items.length) throw new Error('Нет позиций в таблице A2:D.');

  return {
    brand: brandCell,
    apiToken,
    accountNumber,
    invoiceNumber,
    invoiceDateISO,
    dueDateISO,
    payerName,
    email,
    phoneNorm,
    items,
    address: addressNorm,   // уже нормализованный
    aktId
  };
}

function readItemsFromActSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rng = sh.getRange(2, 1, lastRow - 1, 4).getValues(); // A..D
  const out = [];
  for (let i = 0; i < rng.length; i++) {
    const [name, unit, amountRaw, priceRaw] = rng[i];
    if (!name) break;
    const amount = Number(amountRaw);
    const price  = Number(priceRaw);
    if (!isFinite(amount) || amount <= 0) throw new Error(`Неверное количество в строке ${i+2}.`);
    if (!isFinite(price)  || price  < 0)  throw new Error(`Неверная цена в строке ${i+2}.`);
    out.push({
      name: String(name).trim(),
      unit: String(unit || 'Шт').trim(),
      amount,
      price,
      vat: '0'
    });
  }
  return out;
}

/** Проверка структуры листа ОБЪЕКТЫ — чтобы телефон был в C(3), адрес в D(4) */
function assertObjectsLayout_() {
  const sh = ensureObjectsSheet_(SpreadsheetApp.getActive());
  if (!sh) throw new Error('Не найден лист ОБЪЕКТЫ.');

  if (COL.B_PHONE !== 3 || COL.C_ADDR !== 4) {
    throw new Error('Неверная карта COL: телефон должен быть в C(3), адрес — в D(4).');
  }

  const headers = sh.getRange(1, 1, 1, 5).getDisplayValues()[0].map(v => String(v).toLowerCase());
  const telHdr = headers[2] || '';
  const addrHdr = headers[3] || '';
  if (!/тел|phone/.test(telHdr) || !/адрес|address/.test(addrHdr)) {
    Logger.log('⚠️ Проверь заголовки: ожидались "Телефон" в C и "Адрес" в D.');
  }
}


function resolveUidByPhoneAddress_(phoneNorm, adresRaw) {
  assertObjectsLayout_();
  const ss = SpreadsheetApp.getActive();
  const sh = ensureObjectsSheet_(ss);

  const phone = String(phoneNorm || '').trim();
  if (!/^\d{11}$/.test(phone)) throw new Error('resolveUidByPhoneAddress_: phoneNorm должен быть 11 цифр (79XXXXXXXXX).');

  const addressNorm = normAddr_(adresRaw);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Лист ОБЪЕКТЫ пуст: сначала создай смету (UID).');

  const phones = sh.getRange(2, COL.B_PHONE, last - 1, 1).getValues(); // C
  const addrs  = sh.getRange(2, COL.C_ADDR,  last - 1, 1).getValues(); // D

  let rowHit = 0;
  for (let i = 0; i < phones.length; i++) {
    const p = String(phones[i][0] || '').trim();
    const a = normAddr_(addrs[i][0] || '');
    if (p === phone && a === addressNorm) {
      if (rowHit) throw new Error('Найдено несколько объектов с одинаковыми телефоном+адресом. Удали дубликаты.');
      rowHit = 2 + i;
    }
  }
  if (!rowHit) throw new Error('Не найден объект по телефону C и адресу D. Проверь нормализацию и точность данных.');

  const uid = String(sh.getRange(rowHit, COL.A_UID).getValue() || '').trim();
  if (!uid) throw new Error('В найденной строке ОБЪЕКТЫ пуст UID (A). Заполни UID.');
  return uid;
}

function appendOrUpsertLocalInvoiceRegistry_(row) {
  const ss  = SpreadsheetApp.getActive();
  const reg = ensureSheet_(ss, REGISTRY_SHEET_NAME);
  removeFilterSafe_(reg);

  // Заголовок при пустом листе
  if (reg.getLastRow() === 0) {
    reg.getRange(1, 1, 1, REG_COL.LINK).setValues([[
      'uid','createdAt','brand','invoiceNumber','invoiceId','payer','email','phone',
      'total','invoiceDate','status','lastChecked','initiatorEmail','aktId','address','link'
    ]]);
  }

  // Поиск по invoiceId (идемпотентность)
  const last = reg.getLastRow();
  let targetRow = 0;
  if (last >= 2) {
    const ids = reg.getRange(2, REG_COL.INV_ID, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() === String(row.invoiceId || '').trim() && row.invoiceId) {
        targetRow = 2 + i; break;
      }
    }
  }
  if (!targetRow) targetRow = reg.getLastRow() + 1;

  const invoiceDatePretty = toDDMMYYYY_(row.invoiceDateISO);

  reg.getRange(targetRow, 1, 1, REG_COL.LINK).setValues([[
    row.uid || '',
    row.createdAt || '',
    row.brand || '',
    row.invoiceNumber || '',
    row.invoiceId || '',
    row.payer || '',
    row.email || '',
    row.phone || '',
    Number(row.total) || 0,
    invoiceDatePretty || '',
    row.status || '',
    row.lastChecked || '',
    row.initiatorEmail || '',
    row.aktId || '',
    row.address || '',
    row.link || ''
  ]]);
}

function appendOrUpsertTransitInvoiceRegistry_(row) {
  const book = SpreadsheetApp.openById(TRANSIT_FILE_ID);
  const sh   = ensureSheet_(book, REGISTRY_SHEET_NAME);
  removeFilterSafe_(sh);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, REG_COL.LINK).setValues([[
      'uid','createdAt','brand','invoiceNumber','invoiceId','payer','email','phone',
      'total','invoiceDate','status','lastChecked','initiatorEmail','aktId','address','link'
    ]]);
  }

  // upsert по invoiceId (как локально)
  const last = sh.getLastRow();
  let targetRow = 0;
  if (last >= 2) {
    const ids = sh.getRange(2, REG_COL.INV_ID, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() === String(row.invoiceId || '').trim() && row.invoiceId) {
        targetRow = 2 + i; break;
      }
    }
  }
  if (!targetRow) targetRow = sh.getLastRow() + 1;

  const invoiceDatePretty = toDDMMYYYY_(row.invoiceDateISO);

  sh.getRange(targetRow, 1, 1, REG_COL.LINK).setValues([[
    row.uid || '',
    row.createdAt || '',
    row.brand || '',
    row.invoiceNumber || '',
    row.invoiceId || '',
    row.payer || '',
    row.email || '',
    row.phone || '',
    Number(row.total) || 0,
    invoiceDatePretty || '',
    row.status || '',
    row.lastChecked || '',
    row.initiatorEmail || '',
    row.aktId || '',
    row.address || '',
    row.link || ''
  ]]);
}

function upsertTransitInvoiceById_({ invoiceId, status, lastChecked, link }) {
  const book = SpreadsheetApp.openById(TRANSIT_FILE_ID);
  const sh = book.getSheetByName(REGISTRY_SHEET_NAME);
  if (!sh) return;

  const last = sh.getLastRow();
  if (last < 2) return;

  const ids = sh.getRange(2, REG_COL.INV_ID, last - 1, 1).getValues();
  const needle = String(invoiceId || '').trim();
  if (!needle) return;

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === needle) {
      const row = 2 + i;
      if (status)      sh.getRange(row, REG_COL.STATUS).setValue(status);
      if (lastChecked) sh.getRange(row, REG_COL.LAST_CHECKED).setValue(lastChecked);
      if (link)        sh.getRange(row, REG_COL.LINK).setValue(link);
      return;
    }
  }
}

function sendInvoice() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Не удалось получить блокировку.');

  // === Автонормализация телефона в G4 ===
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_INVOICES);
  if (!sh) throw new Error(`Лист "${SHEET_INVOICES}" не найден.`);

  const rawPhone = String(sh.getRange('G4').getDisplayValue()).trim();
  if (rawPhone) {
    try {
      const normalized = normalizeRuPhoneStrict_(rawPhone);
      if (normalized !== rawPhone.replace(/\D+/g, '')) {
        sh.getRange('G4').setValue(normalized);
        SpreadsheetApp.flush();
      }
    } catch (e) {
      throw new Error('Телефон (G4) не удалось нормализовать. Введи, например, 89991112233 или +7 999 1112233.');
    }
  } else {
    throw new Error('Поле G4 (телефон) пустое.');
  }

  try {
    // 1) Собираем данные строго как раньше (для API)
    const d = collectInvoiceData_(); // brand, apiToken, accountNumber, invoiceNumber, invoiceDateISO, dueDateISO, payerName, email, phoneNorm, items, address(норм), aktId

    // 2) UID и сумма для РЕЕСТРОВ — из H1/H2 (ВАЖНО!)
    const uidFromH1 = String(sh.getRange('H1').getDisplayValue()).trim();
    if (!uidFromH1) throw new Error('В H1 должен быть UID (используется для реестров).');
    const totalFromH2Raw = String(sh.getRange('H2').getDisplayValue()).trim();
    const totalFromH2 = Number(totalFromH2Raw.replace(/\s+/g,'').replace(',','.'));
    const hasOverride = isFinite(totalFromH2) && totalFromH2 > 0;

    // 3) Готовим правильный payload под API Т-Банка (как в старой рабочей версии)
    const payload = {
      invoiceNumber: d.invoiceNumber,
      dueDate:       d.dueDateISO,
      invoiceDate:   d.invoiceDateISO,
      accountNumber: d.accountNumber,
      payer: { name: d.payerName },
      items: d.items,
      contacts: [{ email: d.email }, { email: 'invoice@remontvspb.ru' }],
      contactPhone: '+' + d.phoneNorm
    };

    const resp = UrlFetchApp.fetch(API_URL_CREATE, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: {
        'Accept': 'application/json',
        'Authorization': 'Bearer ' + d.apiToken,
        'X-Request-Id': Utilities.getUuid()
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (code < 200 || code >= 300) {
      Logger.log({ code, headers: resp.getAllHeaders(), body });
      throw new Error(`Ошибка API (${code}): ${body}`);
    }

    const data = JSON.parse(body || '{}');
    const invoiceId = data.invoiceId || data.id || '';
    const link      = data.invoiceUrl || data.url || '';
    if (!invoiceId) throw new Error('API не вернул invoiceId.');

    const totalApi = Number(d.items.reduce((s, it) => s + Number(it.amount) * Number(it.price), 0));
    const createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    const initiatorEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();

    // 4) Запись в ЛОКАЛЬНЫЙ реестр (A:UID сдвинуло всё на +1)
    appendOrUpsertLocalInvoiceRegistry_({
      uid: uidFromH1,
      createdAt,
      brand: d.brand,
      invoiceNumber: d.invoiceNumber,
      invoiceId,
      payer: d.payerName,
      email: d.email,
      phone: '+' + d.phoneNorm,
      total: hasOverride ? totalFromH2 : totalApi,
      invoiceDateISO: d.invoiceDateISO,
      status: 'Отправлено',
      lastChecked: '',
      link,
      initiatorEmail,
      aktId: d.aktId,
      address: d.address
    });

    // 5) Запись в ТРАНЗИТ
    appendOrUpsertTransitInvoiceRegistry_({
      uid: uidFromH1,
      createdAt,
      brand: d.brand,
      invoiceNumber: d.invoiceNumber,
      invoiceId,
      payer: d.payerName,
      email: d.email,
      phone: '+' + d.phoneNorm,
      total: hasOverride ? totalFromH2 : totalApi,
      invoiceDateISO: d.invoiceDateISO,
      status: 'Отправлено',
      lastChecked: '',
      link,
      initiatorEmail,
      aktId: d.aktId,
      address: d.address
    });

    SpreadsheetApp.getActive().toast(`Счёт выставлен. UID=${uidFromH1}`, 'OK', 6);

  } catch (err) {
    SpreadsheetApp.getActive().toast('Ошибка: ' + err.message, 'Ошибка', 10);
    throw err;
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

function checkInvoicesFromMenu() {
  const ss  = SpreadsheetApp.getActive();
  const reg = ss.getSheetByName(REGISTRY_SHEET_NAME);
  if (!reg || reg.getLastRow() < 2) { ss.toast('Реестр счетов пуст.', 'Инфо', 4); return; }

  const rows  = reg.getLastRow() - 1;
  const range = reg.getRange(2, 1, rows, REG_COL.LINK).getValues(); // A..LINK (UID уже учтён)
  let checked = 0, updated = 0;

  for (let i = 0; i < range.length; i++) {
    const r = range[i];
    const uid          = String(r[REG_COL.A_UID-1] || '').trim();
    const brand        = String(r[REG_COL.BRAND-1] || '').trim();
    const invoiceId    = String(r[REG_COL.INV_ID-1] || '').trim();
    const statusStored = String(r[REG_COL.STATUS-1] || '').trim();
    const linkStored   = String(r[REG_COL.LINK-1] || '').trim();
    const phoneRaw     = String(r[REG_COL.PHONE-1] || '').trim();
    const adresRaw   = String(r[REG_COL.ADDRESS-1] || '').trim();

    if (!invoiceId || invoiceId.startsWith('(нет')) continue;
    if (statusStored === '👻 ОПЛАЧЕНО! 💰') continue;

    const upBrand = brand.toUpperCase();
    const { apiToken } = getTokensForBrand_(upBrand);
    if (!apiToken) continue;

    const url = `${API_URL_INFO}/${encodeURIComponent(invoiceId)}/info`;
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Accept': 'application/json', 'Authorization': 'Bearer ' + apiToken, 'X-Request-Id': Utilities.getUuid() },
      muteHttpExceptions: true
    });
    checked++;

    if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
      let info = {}; try { info = JSON.parse(resp.getContentText()); } catch(_) {}
      const apiStatus = (info.status || '').toUpperCase();
      const newStatus = getStatusLabelFromApiStatus_(apiStatus); // см. ниже карту статусов
      const newLink   = info.invoiceUrl || info.url || linkStored || '';

      reg.getRange(i+2, REG_COL.STATUS).setValue(newStatus);
      reg.getRange(i+2, REG_COL.LAST_CHECKED).setValue(new Date());
      if (newLink && newLink !== linkStored) reg.getRange(i+2, REG_COL.LINK).setValue(newLink);
      if (newStatus !== statusStored) updated++;

  

      // Транзит апдейт по invoiceId
      upsertTransitInvoiceById_({
        invoiceId,
        status: newStatus,
        lastChecked: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy'),
        link: newLink
      });

    } else {
      const code = resp.getResponseCode();
      Logger.log({ code, body: resp.getContentText(), invoiceId });
      reg.getRange(i+2, REG_COL.LAST_CHECKED).setValue('ERR ' + code);
    }

    Utilities.sleep(250); // анти-лимиты
  }

  SpreadsheetApp.getActive().toast(`Проверено: ${checked}, обновлено: ${updated}.`, 'OK', 6);
}

// простая карта статусов API -> текст в реестре
function getStatusLabelFromApiStatus_(apiStatusUpper) {
  if (apiStatusUpper === 'EXECUTED') return '👻 ОПЛАЧЕНО! 💰';
  if (apiStatusUpper === 'SUBMITTED') return '😔 НЕ ОПЛАЧЕНО 😔';
  if (apiStatusUpper) return apiStatusUpper;
  return 'Отправлено';
}
