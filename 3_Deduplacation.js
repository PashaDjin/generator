/**
 * DEDUPLICATION.GS — СИСТЕМА ДЕДУПЛИКАЦИИ ОБЪЕКТОВ
 * ===================================================
 * 
 * Предотвращает создание дубликатов объектов при вводе
 * похожих адресов менеджерами.
 * 
 * ОСНОВНЫЕ ФУНКЦИИ:
 * • normalizeAddressAdvanced_() - нормализация адресов (ул., д., кв.)
 * • levenshtein_() - расстояние редактирования между строками
 * • addressSimilarity_() - процент сходства адресов (0.0-1.0)
 * • findSimilarObjectsByAddress_() - поиск похожих объектов
 * • findOrCreateObjectUid_() - главная логика (найти или создать UID)
 * • showDuplicateDialogHtml_() - HTML-диалог для менеджера
 * 
 * ВАЖНО:
 * • Все диалоги через HtmlService (НЕ UI.alert!)
 * • Информационные сообщения через toast()
 * • Пороги: ≥99% (точное), 85-99% (похожее), <85% (новый)
 */


/**
 * Расширенная нормализация адреса
 * ================================
 * 
 * Приводит адрес к единому формату:
 * - Москва ул Ленина 1 → Москва, ул. Ленина, д. 1
 * - СПб Невский проспект дом 50 → Санкт-Петербург, Невский пр-т, д. 50
 * 
 * @param {string} input - сырой адрес
 * @return {string} нормализованный адрес
 */
function normalizeAddressAdvanced_(input) {
  if (!input) return '';
  
  let s = String(input)
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();

  // Сокращения городов
  s = s
    .replace(/\b(спб|санкт-петербург|санкт петербург)\b/g, 'санкт-петербург')
    .replace(/\b(мск|москва)\b/g, 'москва');

  // Типы улиц/проспектов
  s = s
    .replace(/\b(улица|ул\.?)\b/g, 'ул.')
    .replace(/\b(проспект|пр-?т|просп\.?)\b/g, 'пр-т')
    .replace(/\b(переулок|пер\.?)\b/g, 'пер.')
    .replace(/\b(площадь|пл\.?)\b/g, 'пл.')
    .replace(/\b(бульвар|б-?р|бул\.?)\b/g, 'б-р');

  // Типы домов/квартир
  s = s
    .replace(/\b(дом|д\.?)\b/g, 'д.')
    .replace(/\b(квартира|кв\.?)\b/g, 'кв.')
    .replace(/\b(корпус|к\.?|корп\.?)\b/g, 'к.')
    .replace(/\b(строение|стр\.?)\b/g, 'стр.');

  // Вставить запятые перед номерами домов/квартир
  s = s
    .replace(/\s+(д\.\s*\d+)/g, ', $1')
    .replace(/\s+(кв\.\s*\d+)/g, ', $1')
    .replace(/\s+(к\.\s*\d+)/g, ', $1')
    .replace(/\s+(стр\.\s*\d+)/g, ', $1');

  // Убрать лишние запятые/пробелы
  s = s
    .replace(/\s*,\s*/g, ', ')
    .replace(/,+/g, ',')
    .replace(/\s+/g, ' ')
    .trim();

  // Удалить запятую в начале (если появилась)
  if (s.startsWith(', ')) s = s.substring(2);

  // Title case (первая буква заглавная)
  s = s.split(' ').map(w => {
    if (!w) return w;
    // Сокращения оставляем строчными
    if (/^(ул\.|пр-т|д\.|кв\.|к\.|стр\.|пер\.|пл\.|б-р)$/i.test(w)) {
      return w.toLowerCase();
    }
    // Первая буква заглавная
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  }).join(' ');

  // Корректировка сокращений после title case
  s = s
    .replace(/\bУл\./g, 'ул.')
    .replace(/\bПр-Т/gi, 'пр-т')
    .replace(/\bД\./g, 'д.')
    .replace(/\bКв\./g, 'кв.')
    .replace(/\bК\./g, 'к.')
    .replace(/\bСтр\./g, 'стр.')
    .replace(/\bПер\./g, 'пер.')
    .replace(/\bПл\./g, 'пл.')
    .replace(/\bБ-Р/gi, 'б-р');

  return s;
}


/**
 * Алгоритм Левенштейна (расстояние редактирования)
 * =================================================
 * 
 * Считает минимальное количество операций (вставка, удаление, замена)
 * для превращения одной строки в другую.
 * 
 * Примеры:
 * - levenshtein_("Ленина", "Ленина") → 0 (идентичны)
 * - levenshtein_("Ленина", "Ленена") → 1 (одна замена)
 * - levenshtein_("Ленина", "Пушкина") → 5 (разные строки)
 * 
 * @param {string} s1 - первая строка
 * @param {string} s2 - вторая строка
 * @return {number} расстояние редактирования
 */
function levenshtein_(s1, s2) {
  const a = String(s1).toLowerCase();
  const b = String(s2).toLowerCase();
  
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  const matrix = [];
  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // замена
          matrix[i][j - 1] + 1,     // вставка
          matrix[i - 1][j] + 1      // удаление
        );
      }
    }
  }

  return matrix[b.length][a.length];
}


/**
 * Вычисление процента сходства адресов
 * =====================================
 * 
 * Формула: 1.0 - (расстояние / макс. длина)
 * 
 * Примеры:
 * - addressSimilarity_("Москва, ул. Ленина, д. 1", "Москва, ул. Ленина, д. 1") → 1.0 (100%)
 * - addressSimilarity_("Москва, ул. Ленина, д. 1", "Москва, ул. Ленена, д. 1") → 0.96 (96%)
 * - addressSimilarity_("Москва, ул. Ленина", "Питер, ул. Ленина") → 0.7 (70%)
 * 
 * @param {string} addr1 - первый адрес
 * @param {string} addr2 - второй адрес
 * @return {number} сходство от 0.0 (разные) до 1.0 (идентичные)
 */
function addressSimilarity_(addr1, addr2) {
  const norm1 = normalizeAddressAdvanced_(addr1);
  const norm2 = normalizeAddressAdvanced_(addr2);
  
  if (norm1 === norm2) return 1.0;
  
  const maxLen = Math.max(norm1.length, norm2.length);
  if (maxLen === 0) return 1.0;
  
  const distance = levenshtein_(norm1, norm2);
  const similarity = 1.0 - (distance / maxLen);
  
  return Math.max(0, Math.min(1, similarity));
}


/**
 * Поиск похожих объектов по адресу
 * =================================
 * 
 * Ищет в листе ОБЪЕКТЫ все строки с адресами, похожими на целевой.
 * Возвращает отсортированный массив (самые похожие первыми).
 * 
 * @param {Sheet} sheet - лист ОБЪЕКТЫ
 * @param {string} targetAddress - адрес для поиска
 * @param {number} threshold - минимальный порог сходства (0.85 по умолчанию)
 * @return {Array<Object>} массив объектов: {row, uid, client, phone, address, similarity}
 */
function findSimilarObjectsByAddress_(sheet, targetAddress, threshold) {
  threshold = threshold || 0.85;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // A:D (UID, Client, Phone, Addr)
  const results = [];
  
  for (let i = 0; i < data.length; i++) {
    const uid = String(data[i][0] || '').trim(); // COL.A_UID = 1 (индекс 0)
    const client = String(data[i][1] || '').trim(); // COL.A_CLIENT = 2 (индекс 1)
    const phone = String(data[i][2] || '').trim(); // COL.B_PHONE = 3 (индекс 2)
    const address = String(data[i][3] || '').trim(); // COL.C_ADDR = 4 (индекс 3)
    
    if (!uid || !address) continue;
    
    const similarity = addressSimilarity_(targetAddress, address);
    
    if (similarity >= threshold) {
      results.push({
        row: 2 + i,
        uid: uid,
        client: client,
        phone: phone,
        address: address,
        similarity: similarity
      });
    }
  }
  
  // Сортировка по убыванию сходства
  results.sort((a, b) => b.similarity - a.similarity);
  
  return results;
}


/**
 * HTML-диалог для менеджера
 * ==========================
 * 
 * Показывает интерактивный HTML-диалог с информацией о дубликатах.
 * 
 * КНОПКИ:
 * - "Использовать старый" → вернёт существующий UID
 * - "Переписать" → обновит данные старого объекта новыми
 * - "Создать новый" → создаст новый объект (только при similarity < 0.99)
 * 
 * @param {Object} existingObj - существующий объект {uid, client, phone, address, similarity}
 * @param {Object} newData - новые данные {client, phone, address}
 * @return {string} 'USE_OLD' | 'OVERWRITE' | 'CREATE_NEW'
 */
function showDuplicateDialogHtml_(existingObj, newData) {
  const similarity = existingObj.similarity;
  const similarityPercent = Math.round(similarity * 100);
  
  // Определяем уровень предупреждения
  let warningLevel = '';
  let warningText = '';
  if (similarity >= 0.99) {
    warningLevel = 'exact';
    warningText = '⚠️ Найден такой же адрес';
  } else if (similarity >= 0.85) {
    warningLevel = 'similar';
    warningText = `⚠️ Найден ПОХОЖИЙ адрес (сходство: ${similarityPercent}%)`;
  }
  
  // Формируем HTML с кнопками
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Roboto', Arial, sans-serif;
            padding: 20px;
            margin: 0;
            background: #f5f5f5;
          }
          .container {
            background: white;
            border-radius: 8px;
            padding: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          }
          .warning {
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 12px 16px;
            margin-bottom: 20px;
            border-radius: 4px;
            font-weight: 500;
            color: #856404;
          }
          .warning.exact {
            background: #f8d7da;
            border-left-color: #dc3545;
            color: #721c24;
          }
          .section {
            margin-bottom: 20px;
            padding: 16px;
            background: #f8f9fa;
            border-radius: 4px;
          }
          .section-title {
            font-weight: 600;
            color: #495057;
            margin-bottom: 8px;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
          }
          .data-row {
            margin: 6px 0;
            font-size: 14px;
            line-height: 1.6;
          }
          .data-label {
            font-weight: 500;
            color: #6c757d;
            display: inline-block;
            width: 80px;
          }
          .data-value {
            color: #212529;
          }
          .highlight {
            background: #fff59d;
            padding: 2px 4px;
            border-radius: 2px;
          }
          .buttons {
            display: flex;
            gap: 12px;
            margin-top: 24px;
            flex-wrap: wrap;
          }
          button {
            flex: 1;
            min-width: 120px;
            padding: 12px 20px;
            border: none;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
            text-transform: uppercase;
            letter-spacing: 0.5px;
          }
          button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
          }
          .btn-primary {
            background: #007bff;
            color: white;
          }
          .btn-primary:hover {
            background: #0056b3;
          }
          .btn-warning {
            background: #ffc107;
            color: #212529;
          }
          .btn-warning:hover {
            background: #e0a800;
          }
          .btn-secondary {
            background: #6c757d;
            color: white;
          }
          .btn-secondary:hover {
            background: #545b62;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="warning ${warningLevel}">
            ${warningText}
          </div>
          
          <div class="section">
            <div class="section-title">📌 Существующий объект ${existingObj.uid}</div>
            <div class="data-row">
              <span class="data-label">Клиент:</span>
              <span class="data-value">${existingObj.client || '(не указан)'}</span>
            </div>
            <div class="data-row">
              <span class="data-label">Телефон:</span>
              <span class="data-value">${existingObj.phone || '(не указан)'}</span>
            </div>
            <div class="data-row">
              <span class="data-label">Адрес:</span>
              <span class="data-value">${existingObj.address}</span>
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">✨ Новые данные</div>
            <div class="data-row">
              <span class="data-label">Клиент:</span>
              <span class="data-value ${newData.client !== existingObj.client ? 'highlight' : ''}">${newData.client || '(не указан)'}</span>
            </div>
            <div class="data-row">
              <span class="data-label">Телефон:</span>
              <span class="data-value ${newData.phone !== existingObj.phone ? 'highlight' : ''}">${newData.phone || '(не указан)'}</span>
            </div>
            <div class="data-row">
              <span class="data-label">Адрес:</span>
              <span class="data-value ${newData.address !== existingObj.address ? 'highlight' : ''}">${newData.address}</span>
            </div>
          </div>
          
          <div class="buttons">
            <button class="btn-primary" onclick="returnValue('USE_OLD')">
              👍 Использовать старый
            </button>
            <button class="btn-warning" onclick="returnValue('OVERWRITE')">
              ✏️ Переписать данные
            </button>
            ${similarity < 0.99 ? '<button class="btn-secondary" onclick="returnValue(\'CREATE_NEW\')">➕ Создать новый</button>' : ''}
          </div>
        </div>
        
        <script>
          function returnValue(action) {
            google.script.host.close();
            google.script.run.processDialogResult_(action);
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(similarity >= 0.99 ? 400 : 450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Найден похожий объект');
  
  // ВАЖНО: Этот подход с google.script.host.close() требует доработки
  // Для синхронного возврата значения используем PropertiesService
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('DIALOG_ACTION', 'WAITING');
  
  // Ожидание ответа (максимум 5 минут)
  const startTime = Date.now();
  while (Date.now() - startTime < 300000) { // 5 минут
    Utilities.sleep(500);
    const action = userProps.getProperty('DIALOG_ACTION');
    if (action && action !== 'WAITING') {
      userProps.deleteProperty('DIALOG_ACTION');
      return action;
    }
  }
  
  // Таймаут — возвращаем USE_OLD по умолчанию
  return 'USE_OLD';
}

/**
 * Обработчик результата диалога (вызывается из HTML)
 */
function processDialogResult_(action) {
  PropertiesService.getUserProperties().setProperty('DIALOG_ACTION', action);
}


/**
 * ГЛАВНАЯ ФУНКЦИЯ: Найти или создать UID
 * ========================================
 * 
 * Ищет похожие объекты по адресу. Если находит:
 * - Показывает HTML-диалог менеджеру
 * - Возвращает результат: {uid, isNew, action, warning}
 * 
 * ПАРАМЕТРЫ:
 * @param {Object} params:
 *   - sheet: лист ОБЪЕКТЫ
 *   - managerId: ID менеджера (для генерации нового UID)
 *   - rawAddress: сырой адрес из формы
 *   - phoneNorm: нормализованный телефон
 *   - client: имя клиента
 *   - autoCreate: true = создавать UID автоматически, false = только искать
 * 
 * ВОЗВРАЩАЕТ:
 * {
 *   uid: "MGR-1",              // UID объекта
 *   isNew: false,              // true = создан новый, false = переиспользован
 *   action: "USE_OLD",         // "USE_OLD" | "OVERWRITE" | "CREATE_NEW"
 *   warning: "...",            // сообщение (если есть)
 *   existingRow: 5             // номер строки существующего объекта (для обновления)
 * }
 */
function findOrCreateObjectUid_(params) {
  const {
    sheet,
    managerId,
    rawAddress,
    phoneNorm,
    client,
    autoCreate = true
  } = params;
  
  if (!sheet || !rawAddress) {
    throw new Error('findOrCreateObjectUid_: обязательны sheet и rawAddress');
  }
  
  // Нормализация адреса
  const addressNorm = normalizeAddressAdvanced_(rawAddress);
  
  // Поиск похожих объектов
  const similar = findSimilarObjectsByAddress_(sheet, addressNorm, 0.85);
  
  // Если ничего не нашли — создаём новый
  if (similar.length === 0) {
    if (!autoCreate) {
      return { uid: null, isNew: true, action: 'CREATE_NEW' };
    }
    
    const newUid = nextUid_(managerId);
    return {
      uid: newUid,
      isNew: true,
      action: 'CREATE_NEW',
      warning: null
    };
  }
  
  // Нашли похожий объект — показываем диалог
  const best = similar[0];
  
  const action = showDuplicateDialogHtml_(best, {
    client: client || '',
    phone: phoneNorm || '',
    address: addressNorm
  });
  
  if (action === 'USE_OLD') {
    return {
      uid: best.uid,
      isNew: false,
      action: 'USE_OLD',
      existingRow: best.row,
      warning: `Используется существующий объект ${best.uid}`
    };
  }
  
  if (action === 'OVERWRITE') {
    return {
      uid: best.uid,
      isNew: false,
      action: 'OVERWRITE',
      existingRow: best.row,
      warning: `Данные объекта ${best.uid} будут обновлены`
    };
  }
  
  // CREATE_NEW — создаём новый объект
  if (!autoCreate) {
    return { uid: null, isNew: true, action: 'CREATE_NEW' };
  }
  
  const newUid = nextUid_(managerId);
  return {
    uid: newUid,
    isNew: true,
    action: 'CREATE_NEW',
    warning: `Создан новый объект ${newUid} (похожий существует: ${best.uid})`
  };
}
