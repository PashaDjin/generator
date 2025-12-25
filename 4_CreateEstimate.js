/**
 * ИСПРАВЛЕННАЯ ВЕРСИЯ createEstimate() С ДЕДУПЛИКАЦИЕЙ
 * ======================================================
 * 
 * ИЗМЕНЕНИЯ:
 * • Добавлен вызов findOrCreateObjectUid_() ПЕРЕД nextUid_()
 * • При обнаружении похожего адреса показывается HTML-диалог
 * • При выборе "Переписать" обновляются поля B, C, D
 * • Синхронизация в ТРАНЗИТ работает аналогично
 * 
 * ВАЖНО: Этот файл ЗАМЕНЯЕТ старую функцию createEstimate()
 * в проекте ГЕНЕРАТОР.
 */

function createEstimate() {
  const ss   = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Блокировка не получена. Повтори.');

  try {
    const inputSh  = ss.getSheetByName(SHEET_EST_INPUT);  // "✅ СМЕТА"
    const techSh   = ss.getSheetByName(SHEET_TECH);       // "Тех"
    if (!inputSh || !techSh) throw new Error('Нужны листы: "✅ СМЕТА" и "Тех".');

    // Входные данные
    const company         = val(inputSh.getRange('C2').getValue()); // Юрлицо
    const manager         = val(inputSh.getRange('C3').getValue());
    const measurer        = val(inputSh.getRange('C4').getValue());
    const client          = val(inputSh.getRange('C5').getValue());
    const adresRaw        = val(inputSh.getRange('C6').getValue());
    const phoneRaw        = val(inputSh.getRange('C7').getValue());
    const managerId       = val(inputSh.getRange('D2').getValue());
    const managerFolderId = val(inputSh.getRange('D3').getValue());
    const measurerEmail   = val(inputSh.getRange('D4').getValue());
    const templateId      = val(techSh.getRange('L2').getValue());   // ID шаблона сметы

    // Валидация
    requireFilled(company, 'C2 (Юрлицо)');
    requireFilled(manager, 'C3 (Менеджер)');
    requireFilled(measurer, 'C4 (Замерщик)');
    requireFilled(client, 'C5 (Клиент)');
    requireFilled(adresRaw, 'C6 (Адрес)');
    requireFilled(phoneRaw, 'C7 (Телефон)');
    requireFilled(managerId, 'D2 (ID менеджера)');
    requireFilled(managerFolderId, 'D3 (ID папки менеджера)');
    requireFilled(measurerEmail, 'D4 (Email замерщика)');
    requireFilled(templateId, 'L2 (ID шаблона сметы)');

    // Нормализация телефона
    let phoneNorm;
    try {
      const normalized = normalizeRuPhoneStrict_(phoneRaw);
      // Если пользователь ввёл в "человеческом" виде — перезапишем ячейку нормализованным
      if (normalized !== String(phoneRaw).replace(/\D+/g, '').replace(/^8/, '7')) {
        inputSh.getRange('C7').setValue(normalized);
        SpreadsheetApp.flush();
      }
      phoneNorm = normalized;
    } catch (e) {
      throw new Error('Телефон (C7) не удалось нормализовать. Введи, например, 89991112233 или +7 999 1112233.');
    }
    
    const address = normAddr_(adresRaw);

    // === НОВОЕ: ПРОВЕРКА ДУБЛИКАТОВ ===
    const shObj = ensureObjectsSheet_(ss);
    removeFilterSafe_(shObj);
    
    const uidResult = findOrCreateObjectUid_({
      sheet: shObj,
      managerId: managerId,
      rawAddress: adresRaw,
      phoneNorm: phoneNorm,
      client: client,
      autoCreate: true
    });
    
    const uid = uidResult.uid;
    const action = uidResult.action;
    const existingRow = uidResult.existingRow;
    
    // Показываем предупреждение (если есть)
    if (uidResult.warning) {
      SpreadsheetApp.getActive().toast(uidResult.warning, 'Инфо', 4);
    }

    // Папка клиента
    const clientFolder = findOrCreateClientFolderByUID_(managerFolderId, uid, address, client);
    const folderUrl    = clientFolder.getUrl();

    // Копия шаблона сметы
    const smetaFile = DriveApp.getFileById(templateId).makeCopy(`${uid} - Смета - ${client}`, clientFolder);
    const smetaId   = smetaFile.getId();
    const smetaUrl  = smetaFile.getUrl();

    // В «🏠 Характеристики объекта»: B1 = UID
    writeSmetaObjectWithUid_(smetaId, { uid, client, phoneNorm, address, measurer, company });

    // Доступы
    addEditors_(smetaFile, [measurerEmail, ALWAYS_EDITOR]);

    // Ссылка на смету в форму
    inputSh.getRange('C10').setValue(smetaUrl);

    // === ЗАПИСЬ/ОБНОВЛЕНИЕ В ОБЪЕКТЫ ===
    let row;
    
    if (action === 'USE_OLD') {
      // Используем существующий объект — НЕ меняем данные
      row = existingRow;
      
    } else if (action === 'OVERWRITE') {
      // Переписываем данные существующего объекта
      row = existingRow;
      
      // ОБНОВЛЯЕМ 3 ПОЛЯ: B (имя), C (телефон), D (адрес)
      shObj.getRange(row, COL.A_CLIENT).setValue(client);     // B = имя
      shObj.getRange(row, COL.B_PHONE).setValue(phoneNorm);   // C = телефон
      shObj.getRange(row, COL.C_ADDR).setValue(address);      // D = адрес
      
    } else {
      // CREATE_NEW — создаём новую строку
      row = findRowByUid_(shObj, uid);
      if (!row) {
        row = createEmptyObjectRowWithUid_(shObj, uid, { client, phoneNorm, addrNorm: address });
      }
    }

    // Обновляем остальные поля (статус, менеджер, замерщик, ссылки)
    shObj.getRange(row, COL.D_STATUS).setValue('Делаем смету');
    if (manager)  shObj.getRange(row, COL.F_MANAGER).setValue(manager);
    if (company)  shObj.getRange(row, COL.P_LEGAL).setValue(company);
    if (measurer) shObj.getRange(row, COL.Q_MEASURER).setValue(measurer);

    // Ссылки
    shObj.getRange(row, COL.K_SMETA_LINK).setValue(smetaUrl);
    shObj.getRange(row, COL.O_FOLDER).setValue(folderUrl);
    setIfProvided_(shObj, row, COL.L_SMETA_ID, smetaId);
    setIfProvided_(shObj, row, COL.M_EST_ID,  uid);

    // === ЗЕРКАЛО В ТРАНЗИТ ===
    mirrorEstimateToTransitByUid_({
      uid, 
      client, 
      phoneNorm, 
      address, 
      status: 'Делаем смету',
      manager, 
      smetaUrl, 
      smetaId, 
      folderUrl, 
      legalEntity: company, 
      measurer,
      action: action  // Передаём action для правильной обработки в ТРАНЗИТ
    });

    SpreadsheetApp.getActive().toast(`Смета создана. UID=${uid}. ОБЪЕКТЫ обновлены и зеркалированы.`, 'OK', 7);

  } catch (err) {
    SpreadsheetApp.getActive().toast('Ошибка: ' + String(err.message || err), 'Ошибка', 10);
    throw err;
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}


/**
 * ОБНОВЛЁННАЯ ВЕРСИЯ mirrorEstimateToTransitByUid_()
 * ===================================================
 * 
 * Теперь поддерживает действие 'OVERWRITE' для обновления
 * полей B, C, D в ТРАНЗИТ.
 */
function mirrorEstimateToTransitByUid_({
  uid, client, phoneNorm, address, status,
  manager, smetaUrl, smetaId, folderUrl, legalEntity, measurer,
  action  // Новый параметр
}) {
  const bookT = SpreadsheetApp.openById(TRANSIT_FILE_ID);
  const shT   = ensureObjectsSheet_(bookT);
  removeFilterSafe_(shT);

  const addrNorm = normAddr_(address);
  
  // Найти/создать строку по UID
  let rowT = findRowByUid_(shT, uid);
  if (!rowT) {
    rowT = createEmptyObjectRowWithUid_(shT, uid, { client, phoneNorm, addrNorm });
  }

  // Если action = OVERWRITE — обновляем B, C, D
  if (action === 'OVERWRITE') {
    shT.getRange(rowT, COL.A_CLIENT).setValue(client);    // B = имя
    shT.getRange(rowT, COL.B_PHONE).setValue(phoneNorm);  // C = телефон
    shT.getRange(rowT, COL.C_ADDR).setValue(addrNorm);    // D = адрес
  } else {
    // Обычная логика — обновляем только адрес
    shT.getRange(rowT, COL.C_ADDR).setValue(addrNorm);
  }

  // Обновляем остальные поля
  shT.getRange(rowT, COL.D_STATUS).setValue(status);
  if (client)      shT.getRange(rowT, COL.A_CLIENT).setValue(client);
  if (manager)     shT.getRange(rowT, COL.F_MANAGER).setValue(manager);
  if (legalEntity) shT.getRange(rowT, COL.P_LEGAL).setValue(legalEntity);
  if (measurer)    shT.getRange(rowT, COL.Q_MEASURER).setValue(measurer);

  // Ссылки
  shT.getRange(rowT, COL.K_SMETA_LINK).setValue(smetaUrl);
  shT.getRange(rowT, COL.O_FOLDER).setValue(folderUrl);
  setIfProvided_(shT, rowT, COL.L_SMETA_ID, smetaId);
  setIfProvided_(shT, rowT, COL.M_EST_ID,  uid);
}
