/**
 * Боевая общая база заявок фестиваля — Google Apps Script.
 * Установка/обновление:
 *   Таблица → Расширения → Apps Script → вставить этот код → 💾 →
 *   Deploy → Manage deployments → ✏️ (Edit) → Version: New version → Deploy.
 *   (URL веб-приложения при этом НЕ меняется.)
 *
 * ВАЖНО при обновлении структуры: новая строка заголовков создаётся автоматически
 * ТОЛЬКО если лист пустой. Поэтому перед первым боевым запуском очисти лист
 * ПОЛНОСТЬЮ (удали все строки, включая старую шапку) — новая шапка встанет сама.
 *
 * Колонки: # | Билет | Имя | Телефон | Email | Согласие реклама | Выпавший вкус | Дата/время
 * «Билет» = уникальный код, который гость видит на экране (генерится на телефоне и шлётся сюда) —
 * экран и таблица показывают ОДИН код. «#» — порядковый номер строки для обзора.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();      // ponytail: global lock, ок для потока фестиваля
  lock.waitLock(20000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['#', 'Билет', 'Имя', 'Телефон', 'Email', 'Согласие реклама', 'Выпавший вкус', 'Дата/время']);
    }
    var d = JSON.parse(e.postData.contents || '{}');
    var n = sheet.getLastRow();                 // строка заголовков = 1 → n = порядковый номер заявки
    sheet.appendRow([
      n,
      d.ticket || '',                           // код билета, который видит гость
      d.name || '',
      d.phone || '',
      d.email || '',
      d.ad ? 'да' : 'нет',
      d.flavor || '',
      new Date()
    ]);
    return json({ ok: true, n: n, ticket: d.ticket || '' });
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
