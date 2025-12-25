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
 * 1. Битрикс отправляет POST с данными клиента
 * 2. GAS создает смету, папку, выдает доступы
 * 3. GAS отправляет обратно в Битрикс: UID, ссылки на смету/папку
 * 
 * ФОРМАТ ЗАПРОСА ОТ БИТРИКСА:
 * {
 *   "company": "БРЕНД-МАР",
 *   "manager": "Иванов Иван",
 *   "measurer": "Петров Петр",
 *   "client": "Сидоров Андрей",
 *   "address": "Москва, ул. Ленина, д. 10, кв. 5",
 *   "phone": "89991112233",
 *   "managerId": "IV123",
 *   "managerFolderId": "1ABC...XYZ",
 *   "measurerEmail": "measurer@example.com",
 *   "templateId": "1DEF...GHI",
 *   "bitrixDealId": "12345",  // ID сделки в Битриксе
 *   "bitrixWebhookUrl": "https://your-domain.bitrix24.ru/rest/1/webhook_code/"
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
    
    // Валидация обязательных полей
    const required = ['company', 'manager', 'measurer', 'client', 'address', 'phone', 
                      'managerId', 'managerFolderId', 'measurerEmail', 'templateId'];
    
    for (const field of required) {
      if (!payload[field]) {
        throw new Error(`Отсутствует обязательное поле: ${field}`);
      }
    }
    
    // Создаём смету с данными от Битрикса
    const result = createEstimateFromBitrix(payload);
    
    // Если указан вебхук Битрикса - отправляем результат обратно
    if (payload.bitrixWebhookUrl) {
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
    // Извлекаем данные из параметров
    const company         = val(data.company);
    const manager         = val(data.manager);
    const measurer        = val(data.measurer);
    const client          = val(data.client);
    const adresRaw        = val(data.address);
    const phoneRaw        = val(data.phone);
    const managerId       = val(data.managerId);
    const managerFolderId = val(data.managerFolderId);
    const measurerEmail   = val(data.measurerEmail);
    const templateId      = val(data.templateId);
    const bitrixDealId    = val(data.bitrixDealId || '');

    // Валидация (уже проверено в doPost, но для надёжности)
    requireFilled(company, 'company (Юрлицо)');
    requireFilled(manager, 'manager (Менеджер)');
    requireFilled(measurer, 'measurer (Замерщик)');
    requireFilled(client, 'client (Клиент)');
    requireFilled(adresRaw, 'address (Адрес)');
    requireFilled(phoneRaw, 'phone (Телефон)');
    requireFilled(managerId, 'managerId (ID менеджера)');
    requireFilled(managerFolderId, 'managerFolderId (ID папки менеджера)');
    requireFilled(measurerEmail, 'measurerEmail (Email замерщика)');
    requireFilled(templateId, 'templateId (ID шаблона сметы)');

    // Нормализация телефона
    let phoneNorm;
    try {
      phoneNorm = normalizeRuPhoneStrict_(phoneRaw);
    } catch (e) {
      throw new Error('Телефон не удалось нормализовать. Формат: 89991112233 или +7 999 1112233.');
    }
    
    const address = normAddr_(adresRaw);

    // === ПРОВЕРКА ДУБЛИКАТОВ ===
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
    
    Logger.log(`UID: ${uid}, Action: ${action}, Row: ${existingRow}`);

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
    let row;
    
    if (action === 'USE_OLD') {
      row = existingRow;
    } else if (action === 'OVERWRITE') {
      row = existingRow;
      shObj.getRange(row, COL.A_CLIENT).setValue(client);
      shObj.getRange(row, COL.B_PHONE).setValue(phoneNorm);
      shObj.getRange(row, COL.C_ADDR).setValue(address);
    } else {
      row = findRowByUid_(shObj, uid);
      if (!row) {
        row = createEmptyObjectRowWithUid_(shObj, uid, { client, phoneNorm, addrNorm: address });
      }
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
    
    // Сохраняем ID сделки Битрикса (если есть свободная колонка)
    // Можно добавить новую колонку COL.S_BITRIX_DEAL_ID в конфиг
    if (bitrixDealId) {
      // shObj.getRange(row, COL.S_BITRIX_DEAL_ID).setValue(bitrixDealId);
    }

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
        'UF_CRM_SMETA_UID': result.uid,           // UID объекта
        'UF_CRM_SMETA_LINK': result.smetaUrl,     // Ссылка на смету
        'UF_CRM_SMETA_ID': result.smetaId,        // ID файла сметы
        'UF_CRM_FOLDER_LINK': result.folderUrl,   // Ссылка на папку клиента
        'STAGE_ID': 'C5:PREPARATION',             // Стадия "Делаем смету" (опционально)
        'COMMENTS': `Смета создана автоматически. UID: ${result.uid}\n${result.message}`
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
 */
function testWebhook() {
  const testData = {
    company: 'БРЕНД-МАР',
    manager: 'Иванов Иван Иванович',
    measurer: 'Петров Петр Петрович',
    client: 'Тестовый Клиент',
    address: 'Москва, ул. Тестовая, д. 1, кв. 1',
    phone: '89991112233',
    managerId: 'TEST123',
    managerFolderId: '1abc...xyz',  // Замени на реальный ID папки
    measurerEmail: 'test@example.com',
    templateId: '1def...ghi',       // Замени на реальный ID шаблона
    bitrixDealId: '12345',
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
