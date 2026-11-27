/**
 * @OnlyCurrentDoc
 * ↑ Ограничивает доступ скрипта ТОЛЬКО этой таблицей.
 *   Без этой строки Google просит доступ ко всем таблицам аккаунта.
 */

// Мэтч наставников — приём результатов в Google-таблицу.
// Вставить в таблицу: Расширения → Apps Script, заменить весь код этим, сохранить,
// затем Начать развертывание → Новое развертывание → Веб-приложение
// (Запуск от имени: я; Доступ: все). Ссылку /exec прислать Claude.

const SHEET_NAME = 'Результаты';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sh.getLastRow() === 0) {
      sh.appendRow(d.header);
      sh.setFrozenRows(1);
      sh.setFrozenColumns(2);
      sh.getRange(1, 1, 1, d.header.length)
        .setFontWeight('bold').setBackground('#04234F').setFontColor('#FFFFFF').setWrap(true);
    }

    // Защита от формул в имени (=, +, -, @ в начале строки).
    const row = d.row.map(v => (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v);
    row[0] = new Date();
    sh.appendRow(row);

    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Открыть ссылку /exec в браузере — должно показать «Мэтч работает».
function doGet() {
  return ContentService.createTextOutput('Мэтч работает');
}
