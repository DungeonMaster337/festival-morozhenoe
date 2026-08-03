/**
 * Боевая общая база заявок фестиваля — Google Apps Script.
 * Установка: скрипт привязать к Google-таблице (Extensions → Apps Script),
 * вставить этот код → Deploy → New deployment → Web app →
 *   Execute as: Me, Who has access: Anyone → скопировать URL веб-приложения.
 * Затем в index.html submit слать fetch(URL, {method:'POST', body: JSON}).
 *
 * Первая строка листа (заголовки) создаётся автоматически при первой заявке.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();      // ponytail: global lock, ок для потока фестиваля
  lock.waitLock(20000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['№', 'Имя', 'Телефон', 'Email', 'Согласие реклама', 'Выпавший вкус', 'Дата/время']);
    }
    var d = JSON.parse(e.postData.contents || '{}');
    var n = sheet.getLastRow();                 // строка заголовков = 1, значит n = порядковый номер заявки
    sheet.appendRow([
      n,
      d.name || '',
      d.phone || '',
      d.email || '',
      d.ad ? 'да' : 'нет',
      d.flavor || '',                           // если добавим поле «выпавший вкус»
      new Date()
    ]);
    return json({ ok: true, n: n });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return json({ ok: true, ping: 'festival endpoint alive' }); }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
