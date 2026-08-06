/**
 * Боевая общая база заявок фестиваля — Google Apps Script.
 * Установка/обновление:
 *   Таблица → Расширения → Apps Script → вставить этот код (заменить старый) → 💾 →
 *   Deploy → Manage deployments → ✏️ (Edit) → Version: New version → Deploy.
 *   (URL веб-приложения при этом НЕ меняется.)
 *
 * ⚠️ ВАЖНО при обновлении структуры: новая строка заголовков создаётся автоматически
 * ТОЛЬКО если лист пустой. Меняется набор колонок → перед первым боевым запуском
 * ОЧИСТИ лист ПОЛНОСТЬЮ (удали все строки, включая старую шапку) — новая шапка встанет
 * сама, и нумерация участников пойдёт с 1.
 *
 * Колонки: # | Имя | Телефон | Email | Согласие реклама | Дата/время
 * «#» — порядковый номер участника (он же «билет»): его присваивает сервер по числу строк
 * и возвращает на телефон, где показывается как «№ 001». Экран и таблица = ОДИН номер.
 *
 * Отправка с телефона идёт через JSONP-GET (тег <script>): это обходит CORS и 302-редирект
 * Apps Script и не виснет на мобильных, в отличие от fetch. Поэтому запись делает doGet.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback;
  if (!p.name && !p.phone) {                 // просто пинг / проверка живости
    return reply(cb, { ok: true, ping: 'festival endpoint alive' });
  }
  var lock = LockService.getScriptLock();    // защита от гонок при одновременных отправках
  lock.waitLock(20000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['#', 'Имя', 'Телефон', 'Email', 'Согласие реклама', 'Дата/время']);
    }
    var n = sheet.getLastRow();              // строка заголовков = 1 → n = номер участника (1,2,3…)
    sheet.appendRow([
      n,
      p.name  || '',
      p.phone || '',
      p.email || '',
      p.ad === '1' ? 'да' : 'нет',
      new Date()
    ]);
    return reply(cb, { ok: true, n: n });
  } catch (err) {
    return reply(cb, { ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// JSONP-обёртка: если пришёл ?callback= — отдаём cb({...}) как JS, иначе обычный JSON.
function reply(cb, obj) {
  var out = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out)
    .setMimeType(ContentService.MimeType.JSON);
}
