/**
 * ═══════════════════════════════════════════════════════════════
 * ФАЙЛ 7: ИНТЕГРАЦИЯ С БИТРИКС24 (ВЕБХУКИ)
 * ═══════════════════════════════════════════════════════════════
 * 
 * Содержит:
 * • doPost() - обработчик входящих вебхуков от Битрикс24
 * • createEstimateFromBitrix() - создание сметы с данными от Битрикса
 * • sendResultToBitrix() - отправка результата обратно в Битрикс24
 * 
 * АРХИТЕКТУРА:
 * 1. Битрикс отправляет POST с минимальными данными клиента
 * 2. GAS создает смету, папку, выдает доступы
 * 3. GAS отправляет обратно в Битрикс: ссылки на смету/папку
 * 
 * ВАЖНО: ID сделки Битрикса используется как UID объекта!
 * 
 * ФОРМАТ ЗАПРОСА ОТ БИТРИКСА (МИНИМАЛЬНЫЙ):
 * {
 *   "bitrixDealId": "12345",      // ID сделки (используется как UID!)
 *   "client": "Сидоров Андрей",   // Имя клиента
 *   "address": "Москва, ул. Ленина, д. 10, кв. 5",
 *   "phone": "89991112233",       // Телефон
 *   "bitrixWebhookUrl": "https://your-domain.bitrix24.ru/rest/1/webhook_code/"
 * }
 * 
 * ОПЦИОНАЛЬНЫЕ ПОЛЯ (если нужно переопределить конфиг):
 * {
 *   "company": "БРЕНД-МАР",
 *   "manager": "Иванов Иван",
 *   "measurer": "Петров Петр",
 *   "managerFolderId": "1ABC...XYZ",
 *   "measurerEmail": "measurer@example.com",
 *   "templateId": "1DEF...GHI"
 * }
 */


// ═══════════════════════════════════════════════════════════════
// ОБРАБОТЧИК ВХОДЯЩИХ ВЕБХУКОВ
// ═══════════════════════════════════════════════════════════════

/**
 * doPost() - главная функция для обработки POST-запросов
 * 
 * Автоматически вызывается при HTTP POST на URL веб-приложения.
 * URL формата: https://script.google.com/macros/s/{DEPLOY_ID}/exec
 * 
 * @param {Object} e - событие с данными запроса
 * @return {ContentService.TextOutput} JSON-ответ
 */
function doPost(e) {
  try {
    // Парсим JSON из тела запроса
    const payload = JSON.parse(e.postData.contents);
    
    Logger.log('Получен запрос от Битрикс24: ' + JSON.stringify(payload));
    
    // Валидация ОБЯЗАТЕЛЬНЫХ полей (только минимум!)
    const required = ['bitrixDealId', 'client', 'address', 'phone'];
    
    for (const field of required) {
      if (!payload[field]) {
        throw new Error(`Отсутствует обязательное поле: ${field}`);
      }
    }
    
    // Создаём смету с данными от Битрикса
    const result = createEstimateFromBitrix(payload);
    
    // Если указан вебхук Битрикса - отправляем результат обратно
    if (payload.bitrixWebhookUrl && payload.bitrixDealId) {
      sendResultToBitrix(payload.bitrixWebhookUrl, payload.bitrixDealId, result);
    }
    
    // Возвращаем успешный ответ
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: result
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('ОШИБКА в doPost: ' + error.message);
    
    // Возвращаем ошибку
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message || String(error)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ═══════════════════════════════════════════════════════════════
// СОЗДАНИЕ СМЕТЫ С ДАННЫМИ ОТ БИТРИКСА
// ═══════════════════════════════════════════════════════════════

/**
 * createEstimateFromBitrix() - создание сметы с внешними данными
 * 
 * Аналог createEstimate(), но данные берутся не из Google Sheets,
 * а из параметров функции (от Битрикс24).
 * 
 * ВАЖНО: ID сделки Битрикса используется как UID объекта!
 * Генерация UID через nextUid_() больше НЕ используется.
 * 
 * @param {Object} data - данные от Битрикса
 * @return {Object} результат создания сметы
 */
function createEstimateFromBitrix(data) {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  
  if (!lock.tryLock(30000)) {
    throw new Error('Блокировка не получена. Повторите попытку.');
  }

  try {
    // === ОБЯЗАТЕЛЬНЫЕ ДАННЫЕ ===
    const bitrixDealId    = val(data.bitrixDealId);  // ID сделки = UID!
    const client          = val(data.client);
    const adresRaw        = val(data.address);
    const phoneRaw        = val(data.phone);

    // Валидация обязательных полей
    requireFilled(bitrixDealId, 'bitrixDealId (ID сделки)');
    requireFilled(client, 'client (Клиент)');
    requireFilled(adresRaw, 'address (Адрес)');
    requireFilled(phoneRaw, 'phone (Телефон)');

    // === ОПЦИОНАЛЬНЫЕ ДАННЫЕ (берём из конфига, если не указаны) ===
    const techSh = ss.getSheetByName(SHEET_TECH);
    if (!techSh) throw new Error('Лист "Тех" не найден для чтения конфига.');

    const company         = val(data.company || techSh.getRange('L3').getValue() || '');
    const manager         = val(data.manager || techSh.getRange('L4').getValue() || '');
    const measurer        = val(data.measurer || techSh.getRange('L5').getValue() || '');
    const managerFolderId = val(data.managerFolderId || techSh.getRange('L6').getValue());
    const measurerEmail   = val(data.measurerEmail || techSh.getRange('L7').getValue());
    const templateId      = val(data.templateId || techSh.getRange('L2').getValue());

    // Валидация технических полей (должны быть или в запросе, или в конфиге)
    requireFilled(managerFolderId, 'managerFolderId (ID папки менеджера) - укажи в Тех!L6');
    requireFilled(measurerEmail, 'measurerEmail (Email замерщика) - укажи в Тех!L7');
    requireFilled(templateId, 'templateId (ID шаблона сметы) - укажи в Тех!L2');

    // Нормализация телефона
    let phoneNorm;
    try {
      phoneNorm = normalizeRuPhoneStrict_(phoneRaw);
    } catch (e) {
      throw new Error('Телефон не удалось нормализовать. Формат: 89991112233 или +7 999 1112233.');
    }
    
    const address = normAddr_(adresRaw);

    // === UID = ID СДЕЛКИ БИТРИКСА ===
    const uid = `BITRIX-${bitrixDealId}`;  // Формат: BITRIX-12345
    
    Logger.log(`UID (из Битрикса): ${uid}`);

    // === ПРОВЕРКА ДУБЛИКАТОВ (по телефону и адресу, НЕ по UID) ===
    const shObj = ensureObjectsSheet_(ss);
    removeFilterSafe_(shObj);
    
    // Ищем существующий объект по телефону/адресу
    let row = findRowByPhoneAndNormalizedAddress_(shObj, phoneNorm, address);
    let action = 'CREATE_NEW';
    
    if (row) {
      // Найден объект с таким же телефоном и адресом
      const existingUid = String(shObj.getRange(row, COL.A_UID).getValue() || '').trim();
      
      if (existingUid && existingUid !== uid) {
        // Есть другой UID - обновляем на новый
        Logger.log(`Найден объект с UID=${existingUid}, обновляем на ${uid}`);
        action = 'OVERWRITE';
      } else if (existingUid === uid) {
        // Уже есть с тем же UID
        Logger.log(`Объект с UID=${uid} уже существует`);
        action = 'USE_OLD';
      }
    } else {
      // Ищем по UID (вдруг уже есть объект с таким ID сделки)
      row = findRowByUid_(shObj, uid);
      if (row) {
        action = 'USE_OLD';
      }
    }

    // === СОЗДАНИЕ ПАПКИ И СМЕТЫ ===
    const clientFolder = findOrCreateClientFolderByUID_(managerFolderId, uid, address, client);
    const folderUrl    = clientFolder.getUrl();

    const smetaFile = DriveApp.getFileById(templateId).makeCopy(`${uid} - Смета - ${client}`, clientFolder);
    const smetaId   = smetaFile.getId();
    const smetaUrl  = smetaFile.getUrl();

    // Заполнение шаблона сметы
    writeSmetaObjectWithUid_(smetaId, { uid, client, phoneNorm, address, measurer, company });

    // Доступы
    addEditors_(smetaFile, [measurerEmail, ALWAYS_EDITOR]);

    // === ЗАПИСЬ/ОБНОВЛЕНИЕ В ОБЪЕКТЫ ===
    if (!row) {
      // Создаём новую строку
      row = createEmptyObjectRowWithUid_(shObj, uid, { client, phoneNorm, addrNorm: address });
    } else if (action === 'OVERWRITE') {
      // Обновляем UID, имя, телефон, адрес
      shObj.getRange(row, COL.A_UID).setValue(uid);
      shObj.getRange(row, COL.A_CLIENT).setValue(client);
      shObj.getRange(row, COL.B_PHONE).setValue(phoneNorm);
      shObj.getRange(row, COL.C_ADDR).setValue(address);
    }

    // Обновляем поля
    shObj.getRange(row, COL.D_STATUS).setValue('Делаем смету');
    if (manager)  shObj.getRange(row, COL.F_MANAGER).setValue(manager);
    if (company)  shObj.getRange(row, COL.P_LEGAL).setValue(company);
    if (measurer) shObj.getRange(row, COL.Q_MEASURER).setValue(measurer);

    // Ссылки
    shObj.getRange(row, COL.K_SMETA_LINK).setValue(smetaUrl);
    shObj.getRange(row, COL.O_FOLDER).setValue(folderUrl);
    setIfProvided_(shObj, row, COL.L_SMETA_ID, smetaId);
    setIfProvided_(shObj, row, COL.M_EST_ID, uid);

    // === ЗЕРКАЛО В ТРАНЗИТ ===
    mirrorEstimateToTransitByUid_({
      uid, client, phoneNorm, address, 
      status: 'Делаем смету',
      manager, smetaUrl, smetaId, folderUrl, 
      legalEntity: company, measurer,
      action: action
    });

    Logger.log(`✅ Смета создана. UID=${uid}`);

    // Возвращаем результат для отправки в Битрикс
    return {
      uid: uid,
      bitrixDealId: bitrixDealId,
      smetaUrl: smetaUrl,
      smetaId: smetaId,
      folderUrl: folderUrl,
      action: action,
      status: 'Делаем смету',
      message: `Смета успешно создана. UID: ${uid}`
    };

  } catch (err) {
    Logger.log('ОШИБКА в createEstimateFromBitrix: ' + err.message);
    throw err;
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}


// ═══════════════════════════════════════════════════════════════
// ОТПРАВКА РЕЗУЛЬТАТА В БИТРИКС24
// ═══════════════════════════════════════════════════════════════

/**
 * sendResultToBitrix() - отправка данных обратно в Битрикс24
 * 
 * Обновляет поля сделки в Битрикс24 через исходящий REST API.
 * 
 * @param {string} webhookUrl - URL вебхука Битрикса (с токеном)
 * @param {string} dealId - ID сделки в Битриксе
 * @param {Object} result - результат создания сметы
 */
function sendResultToBitrix(webhookUrl, dealId, result) {
  if (!webhookUrl || !dealId) {
    Logger.log('⚠️ Не указан webhookUrl или dealId - пропускаем отправку в Битрикс');
    return;
  }

  try {
    // Формируем URL для обновления сделки
    // Пример: https://your-domain.bitrix24.ru/rest/1/webhook_code/crm.deal.update
    const apiUrl = webhookUrl.replace(/\/$/, '') + '/crm.deal.update';
    
    // Подготовка данных для Битрикса
    // UF_CRM_... - пользовательские поля (нужно создать в Битриксе)
    const payload = {
      id: dealId,
      fields: {
        'UF_CRM_SMETA_UID': result.uid,           // UID объекта (BITRIX-12345)
        'UF_CRM_SMETA_LINK': result.smetaUrl,     // Ссылка на смету
        'UF_CRM_SMETA_ID': result.smetaId,        // ID файла сметы
        'UF_CRM_FOLDER_LINK': result.folderUrl,   // Ссылка на папку клиента
        'COMMENTS': `Смета создана автоматически.\nUID: ${result.uid}\n${result.message}`
      }
    };

    // Отправка POST-запроса в Битрикс24
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`Ответ Битрикса (${responseCode}): ${responseText}`);

    if (responseCode !== 200) {
      throw new Error(`Битрикс вернул ошибку ${responseCode}: ${responseText}`);
    }

    const responseData = JSON.parse(responseText);
    
    if (!responseData.result) {
      throw new Error(`Битрикс не обновил сделку: ${responseText}`);
    }

    Logger.log(`✅ Данные успешно отправлены в Битрикс. Deal ID: ${dealId}`);

  } catch (error) {
    Logger.log(`❌ Ошибка отправки в Битрикс: ${error.message}`);
    // НЕ бросаем ошибку - смета уже создана, это не критично
  }
}


// ═══════════════════════════════════════════════════════════════
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ТЕСТИРОВАНИЯ
// ═══════════════════════════════════════════════════════════════

/**
 * testWebhook() - функция для тестирования вебхука вручную
 * 
 * Запускается через редактор Apps Script для проверки работы.
 * 
 * ВАЖНО: Перед запуском заполни в листе "Тех":
 * • L2 - ID шаблона сметы
 * • L3 - Юрлицо (БРЕНД-МАР / СТРОЙМАТ)
 * • L4 - Менеджер (имя)
 * • L5 - Замерщик (имя)
 * • L6 - ID папки менеджера в Google Drive
 * • L7 - Email замерщика
 */
function testWebhook() {
  const testData = {
    // === ОБЯЗАТЕЛЬНЫЕ ПОЛЯ (от Битрикса) ===
    bitrixDealId: '12345',                        // ID сделки в Битриксе
    client: 'Тестовый Клиент',                    // Имя клиента
    address: 'Москва, ул. Тестовая, д. 1, кв. 1', // Адрес
    phone: '89991112233',                         // Телефон
    
    // === ОПЦИОНАЛЬНЫЕ ПОЛЯ (можно указать или взять из конфига Тех!L3-L7) ===
    // company: 'БРЕНД-МАР',
    // manager: 'Иванов Иван',
    // measurer: 'Петров Петр',
    // managerFolderId: '1abc...xyz',
    // measurerEmail: 'test@example.com',
    // templateId: '1def...ghi',
    
    // === ДЛЯ ОТПРАВКИ РЕЗУЛЬТАТА ОБРАТНО ===
    bitrixWebhookUrl: 'https://your-domain.bitrix24.ru/rest/1/webhook_code/'
  };

  try {
    const result = createEstimateFromBitrix(testData);
    Logger.log('✅ ТЕСТ УСПЕШЕН:');
    Logger.log(JSON.stringify(result, null, 2));
  } catch (error) {
    Logger.log('❌ ОШИБКА ТЕСТА: ' + error.message);
  }
}


/**
 * getWebAppUrl() - получить URL веб-приложения для настройки вебхука
 * 
 * После деплоя веб-приложения запусти эту функцию, чтобы узнать URL.
 * Этот URL нужно указать в Битриксе как входящий вебхук.
 */
function getWebAppUrl() {
  const scriptId = ScriptApp.getScriptId();
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('ИНСТРУКЦИЯ ПО НАСТРОЙКЕ ВЕБХУКА:');
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('1. Перейди: https://script.google.com/home/projects/' + scriptId);
  Logger.log('2. Нажми "Развернуть" → "Новое развертывание"');
  Logger.log('3. Выбери тип: "Веб-приложение"');
  Logger.log('4. Настройки:');
  Logger.log('   - Описание: "Bitrix24 Webhook"');
  Logger.log('   - Запуск от имени: "Я"');
  Logger.log('   - У кого есть доступ: "Все"');
  Logger.log('5. Скопируй URL развертывания');
  Logger.log('6. Укажи этот URL в Битрикс24 как входящий вебхук');
  Logger.log('═══════════════════════════════════════════════');
}
