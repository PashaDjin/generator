/************************************
 *   ГЕНЕРАТОР — СОЗДАНИЕ ДОГОВОРА
 ************************************/

function createContract() {
  const ss   = SpreadsheetApp.getActive();
  const fl   = ss.getSheetByName(SHEET_CONTRACT);
  const tech = ss.getSheetByName(SHEET_TECH);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Блокировка не получена. Повтори.');

  try {
    if (!fl)   throw new Error('Лист "📜 Договор" не найден.');
    if (!tech) throw new Error('Лист "Тех" не найден.');

    // 1) Плейсхолдеры (C3:C30 — ключи, B3:B30 — значения)
    const keys = fl.getRange('C3:C30').getValues().map(r => String(r[0]).trim());
    const vals = fl.getRange('B3:B30').getValues().map(r => r[0]);
    const ps = {};
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i]; if (!k) continue;
      let v = vals[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM.yyyy');
      ps[k] = String(v ?? '').trim();
    }

    // 2) Шаблоны к разворачиванию (E3:E8)
    const docNames = fl.getRange('E3:E8').getValues().map(r=>String(r[0]).trim()).filter(Boolean);
    if (!docNames.length) throw new Error('В E3:E8 нет названий документов.');

    // 3) Индекс шаблонов из Тех (A:E): A=code, B=name, C=tid, D=cnt, E=autoAdd
   const lastRow = tech.getLastRow();
    const rows = tech.getRange('A2:E' + lastRow).getValues();
    const docIndex = {};
    rows.forEach((r,i)=>{
    const [code,name,tid,cnt/*,add*/] = r; if(!name) return;
    const nm = String(name).trim();
    docIndex[nm] = { code:String(code||'').trim(), tid:String(tid||'').trim(), cnt:Number(cnt)||0, row:i+2 };
    });


    

    // 5) Обязательные значения
    const manager  = (ps['{manager}']  || '').trim();
    const inits    = (ps['{inits}']    || '').trim();
    const fullname = (ps['{fullname}'] || '').trim();
    const adresRaw = (ps['{adres1}']   || '').trim();
 
    if (!manager || !fullname || !adresRaw) {
      throw new Error('Нужны {manager}, {fullname}, {adres1}.');
    }
    const phoneNorm = normalizeRuPhoneSoft_(ps['{phone}'] || '');
    const mail      = (ps['{mail}'] || '').trim();
 

    // 6) UID: либо уже пришёл в {number}, либо ищем по phone+address (норм.)
     let uid = (ps['{number}'] || '').trim();
     if (!uid) {
     throw new Error('Укажи UID в {uid} на листе "📜 Договор". Сначала создай смету, чтобы получить UID.');
     }
    const addressNorm = normAddr_(adresRaw); // остальное без изменений
    // {number} = UID (для шаблонов)
    ps['{number}'] = uid;

    // 7) Папка клиента: найти по second-part (сырой адрес) и ПЕРЕИМЕНОВАТЬ → "{inits} - {adres1}"
   const managerCell = tech.getRange('F2:F')
     .createTextFinder(manager)
      .matchEntireCell(true)
      .findNext();
    if (!managerCell) throw new Error('Менеджер не найден в Тех!F.');
    const managerRowIdx = managerCell.getRow();
    const managerFolderId = String(tech.getRange(managerRowIdx, 7).getValue() || '').trim(); // G
    if (!managerFolderId) throw new Error('В Тех!G пустой ID папки менеджера.');

    const clientFolder = findOrCreateClientFolderByUID_(managerFolderId, uid, addressNorm, fullname);
    const folderUrl = clientFolder.getUrl();

    // 8) Сумма и юрлицо
    let priceVal = (ps['{price}'] || '').trim();
    if (!priceVal) priceVal = fl.getRange('B17').getDisplayValue();
    const sumTotal = parsePrice(priceVal);
   const firstName = docNames[0];
    const firstMeta = docIndex[firstName] || {};
    const legalEntity = ps['{legal}'] ? ps['{legal}'] : mapCodeToLegal_(firstMeta.code, '');

    // 9) Создаём документы
    const outFG = []; // для вывода на форму [name, url]
    let lastContractUrl = ''; // чтобы в ОБЪЕКТЫ положить ссылку (N)
    for (const name of docNames) {
      const m = docIndex[name]; 
      if (!m) throw new Error(`Нет в индексе шаблонов: "${name}".`);
      if (!m.tid) throw new Error(`В Тех не указан templateId для "${name}".`);

      const title = `${uid} - ${name} - ${inits}`;
      const file  = DriveApp.getFileById(m.tid).makeCopy(title, clientFolder);

      // Подстановка плейсхолдеров (без {uid}, он не используется)
      replaceWithOptionalHighlight(file.getId(), ps, new Set());
      try { addEditors_(file, [ALWAYS_EDITOR]); } catch(_) {}

      const url = file.getUrl();
      outFG.push([title, url]);
      lastContractUrl = url;

      // ++ счётчик использования шаблона (оставляем фичу)
      const c = tech.getRange(m.row, 4);
      c.setValue((Number(c.getValue()) || 0) + 1);
    }

    // 10) Запись в ОБЪЕКТЫ и зеркало в ТРАНЗИТ (upsert по UID)
    const shObj = ensureObjectsSheet_(ss);
    upsertContractObjectByUid_({
      sheet: shObj, uid, client: fullname, phoneNorm,
      addressNorm, manager, legalEntity, sumTotal,
      contractUrl: lastContractUrl, folderUrl, mail
    });

    mirrorContractToTransitByUid_({
      uid, client: fullname, phoneNorm, addressNorm,
      manager, legalEntity, sumTotal, contractUrl: lastContractUrl
    });

  // 11) Вывод на форму: RICH-ссылки в F3:F… (напротив названий в E)
    // Чистим старые следы в F и G
    fl.getRange('F3:F8').clearContent();
    fl.getRange('G3:G8').clearContent();

    // outFG.push([title, url]);
    for (let i = 0; i < outFG.length; i++) {
      const [docName, docUrl] = outFG[i];
      if (!docName || !docUrl) continue;
     fl.getRange(3 + i, 6).setValue(docUrl); // F-колонка = голая ссылка
    }

    // Папка клиента — rich-ссылка в E9 (как согласовано ранее)
    if (folderUrl) {
     fl.getRange('E9').setValue(folderUrl);
    }

    SpreadsheetApp.getActive().toast('Договор(ы) созданы, ОБЪЕКТЫ обновлены и зеркалированы.', 'OK', 6);

  } catch (e) {
    SpreadsheetApp.getActive().toast('Ошибка: ' + e.message, 'Ошибка', 10);
    throw e;
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

/**
 * Зеркало договора в Транзит (upsert по UID)
 */
function mirrorContractToTransitByUid_({ uid, client, phoneNorm, addressNorm, manager, legalEntity, sumTotal, contractUrl }) {
  const bookT = SpreadsheetApp.openById(TRANSIT_FILE_ID);
  const shT   = ensureObjectsSheet_(bookT);
  removeFilterSafe_(shT);

  let rowT = findRowByUid_(shT, uid);
  if (!rowT) {
    rowT = createEmptyObjectRowWithUid_(shT, uid, {
      client, phoneNorm, addrNorm: addressNorm
    });
  }

  // Базовые поля
  shT.getRange(rowT, COL.D_STATUS).setValue(STATUS_CONTRACT);
  shT.getRange(rowT, COL.C_ADDR).setValue(addressNorm);
  if (client)      shT.getRange(rowT, COL.A_CLIENT).setValue(client);
  if (manager)     shT.getRange(rowT, COL.F_MANAGER).setValue(manager);
  if (legalEntity) shT.getRange(rowT, COL.P_LEGAL).setValue(legalEntity);
  if (sumTotal !== null && sumTotal !== undefined && sumTotal !== '') {
    shT.getRange(rowT, COL.I_SUM).setValue(sumTotal);
  }

  // Ссылка на договор
  shT.getRange(rowT, COL.N_CONTRACT_LINK).setValue(contractUrl);
}
