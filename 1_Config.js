/**
 * ═══════════════════════════════════════════════════════════════
 * ФАЙЛ 1: КОНФИГУРАЦИЯ И КОНСТАНТЫ
 * ═══════════════════════════════════════════════════════════════
 * 
 * Содержит:
 * • Названия листов и ID файлов
 * • Константы колонок (COL, REG_COL)
 * • Генератор UID
 * • Меню OnOpen
 */

// ═══════════════════════════════════════════════════════════════
// НАЗВАНИЯ ЛИСТОВ И ID ФАЙЛОВ
// ═══════════════════════════════════════════════════════════════

const SHEET_EST_INPUT      = '✅ СМЕТА';
const SHEET_TECH           = 'Тех';
const SHEET_SMETA_OBJECT   = '🏠 Характеристики объекта';
const OBJECTS_SHEET_NAME   = 'ОБЪЕКТЫ';
const REGISTRY_SHEET_NAME  = 'РЕЕСТР СЧЕТОВ';
const SHEET_CONTRACT       = '📜 Договор';
const SHEET_INVOICES       = '💰 СЧЕТА';

const API_URL_CREATE = 'https://business.tbank.ru/openapi/api/v1/invoice/send';
const API_URL_INFO   = 'https://business.tbank.ru/openapi/api/v1/openapi/invoice';
const ACCOUNT_BRANDMAR = '40702810410000427172';
const ACCOUNT_STROYMAT = '40702810510001646580';

const TRANSIT_FILE_ID = '1u1VqcokQxDQ9MtuMYwPH_1FBrKXeIM4t95QwDs16M1k';
const ALWAYS_EDITOR   = 'ikris.0077@gmail.com';

// ═══════════════════════════════════════════════════════════════
// КОНФИГУРАЦИЯ ДЛЯ БИТРИКС24 (значения по умолчанию)
// ═══════════════════════════════════════════════════════════════
// Эти значения можно переопределить в листе "Тех":
// • L2 - ID шаблона сметы (обязательно!)
// • L3 - Юрлицо (БРЕНД-МАР / СТРОЙМАТ)
// • L4 - Менеджер по умолчанию
// • L5 - Замерщик по умолчанию
// • L6 - ID папки менеджера в Google Drive (обязательно!)
// • L7 - Email замерщика (обязательно!)
//
// Если эти ячейки пусты - используются константы ниже:

const DEFAULT_COMPANY         = 'БРЕНД-МАР';
const DEFAULT_MANAGER         = 'Менеджер';
const DEFAULT_MEASURER        = 'Замерщик';
const DEFAULT_MANAGER_FOLDER  = '';  // Укажи ID папки в Drive или заполни Тех!L6
const DEFAULT_MEASURER_EMAIL  = '';  // Укажи email или заполни Тех!L7
const DEFAULT_TEMPLATE_ID     = '';  // Укажи ID шаблона сметы или заполни Тех!L2


// ═══════════════════════════════════════════════════════════════
// МЕНЮ
// ═══════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ГЕНЕРАТОР')
    .addItem('✅ Создать смету', 'createEstimate')
    .addItem('📜 Сделать выбранный договор', 'createContract')
    .addItem('💰 СЧЕТ', 'sendInvoice')
    .addItem('🔎 Проверить статусы', 'checkInvoicesFromMenu')
    .addToUi();
}


// ═══════════════════════════════════════════════════════════════
// КОНСТАНТЫ КОЛОНОК
// ═══════════════════════════════════════════════════════════════

const MAX_COL_OBJECTS = 19;

const COL = {
  A_UID: 1,
  A_CLIENT: 2,
  B_PHONE: 3,
  C_ADDR: 4,
  D_STATUS: 5,
  E_DESIGN: 6,
  F_MANAGER: 7,
  G_PM: 8,
  H_NOTE: 9,
  I_SUM: 10,
  J_REST: 11,
  K_SMETA_LINK: 12,
  L_SMETA_ID: 13,
  M_EST_ID: 14,
  N_CONTRACT_LINK: 15,
  O_FOLDER: 16,
  P_LEGAL: 17,
  Q_MEASURER: 18,
  R_MAIL: 19
};

const REG_COL = {
  A_UID: 1,
  CREATED: 2,
  BRAND: 3,
  INV_NUMBER: 4,
  INV_ID: 5,
  PAYER: 6,
  EMAIL: 7,
  PHONE: 8,
  TOTAL: 9,
  INV_DATE: 10,
  STATUS: 11,
  LAST_CHECKED: 12,
  INITIATOR: 13,
  AKT_ID: 14,
  ADDRESS: 15,
  LINK: 16
};

const STATUS_CONTRACT = 'Подписание';


// ═══════════════════════════════════════════════════════════════
// ГЕНЕРАТОР UID
// ═══════════════════════════════════════════════════════════════

/**
 * Генератор UID формата "<managerId>-<counter>"
 * Берёт счётчик из Тех!V1, возвращает UID и увеличивает на 1
 */
function nextUid_(managerId) {
  const id = String(managerId || '').trim();
  if (!id) throw new Error('managerId пуст — не могу сгенерить UID.');

  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(SHEET_TECH);
  if (!tech) throw new Error('Лист "Тех" не найден для счётчика UID.');

  const cell = tech.getRange('V1');
  let counter = Number(cell.getValue());
  if (!isFinite(counter) || counter <= 0) counter = 1;

  const uid = `${id}-${counter}`;
  cell.setValue(counter + 1);

  return uid;
}
