// ================== НАСТРОЙКИ ==================
const MSK = { lat: 55.7558, lon: 37.6173 };
const AVG_SPEED_KMH = 55;
const HOURS_PER_DAY = 9;
const ROAD_FACTOR = 1.25;
const MIN_KM_PER_DAY = 150;
const MAX_AGE_DAYS = 14;

// Вебхук Битрикса в коде НЕ хранится: он лежит в свойствах скрипта.
// Задать: меню «🚚 Автовозы → Вебхук Битрикса» либо
// Настройки проекта → Свойства скрипта → ключ B24_WEBHOOK.
const B24_PROP = 'B24_WEBHOOK';

function b24() {
  const v = PropertiesService.getScriptProperties().getProperty(B24_PROP);
  if (!v) throw new Error(
    'Не задан вебхук Битрикса. Меню «🚚 Автовозы → Вебхук Битрикса» ' +
    'или Настройки проекта → Свойства скрипта → ключ ' + B24_PROP + '.');
  return String(v).trim().replace(/\/+$/, '') + '/';
}

function setB24Webhook() {
  const ui = SpreadsheetApp.getUi();
  const cur = PropertiesService.getScriptProperties().getProperty(B24_PROP);
  const NL = '\n\n';
  const hint = cur
    ? NL + 'Сейчас задан: ' + cur.replace(/rest\/\d+\/[^\/]+/, 'rest/…/…')
    : NL + 'Сейчас не задан.';
  const res = ui.prompt('Вебхук Битрикса',
    'Вставьте входящий вебхук целиком, например:\nhttps://ваш-портал.bitrix24.ru/rest/52/xxxxxxxx/' + hint,
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const url = String(res.getResponseText() || '').trim();
  if (!/^https:\/\/[^\/]+\/rest\/\d+\/[A-Za-z0-9]+\/?$/.test(url)) {
    ui.alert('Не похоже на вебхук. Ожидается вид https://портал.bitrix24.ru/rest/52/xxxxxxxx/');
    return;
  }
  PropertiesService.getScriptProperties().setProperty(B24_PROP, url);
  SpreadsheetApp.getActive().toast('Вебхук сохранён в свойствах скрипта', 'Готово', 4);
}

function checkB24() {
  const res = UrlFetchApp.fetch(b24() + 'profile.json', { muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  if (data.error) { SpreadsheetApp.getUi().alert('Битрикс отвечает ошибкой: ' + (data.error_description || data.error)); return; }
  SpreadsheetApp.getUi().alert('Связь есть. Пользователь вебхука: ' + (data.result && data.result.NAME || '') + ' ' + (data.result && data.result.LAST_NAME || ''));
}
const B24_CATEGORY  = 4;
const B24_TRUCK     = 'UF_CRM_1755854996126';
const B24_ARRIVED   = 'UF_CRM_1734620131';
const B24_TO_PLACE  = 'UF_CRM_1722393583543';
const B24_HIDE_TO   = ['Чита', 'Харгос-Москва'];
const B24_STAGES    = ['C4:UC_P1CA59', 'C4:UC_9HAG50', 'C4:UC_OWAJBE'];  // ранние: Транзит, Брокерский, Предв. пошлина
const B24_SHIP_DATE = 'UF_CRM_1776043454137';  // дата отправки
const B24_BATCH_DAYS = 2;   // окно партии, ± дней от самой свежей отправки
const B24_ARRIVING_LEFT = 2; // осталось столько машин в пути (или меньше) → статус "Прибывает", не на карте

const CLR = {
  headBg:'#1c3d5a', headText:'#ffffff', band1:'#ffffff', band2:'#f4f7fa', border:'#d5dde5',
  green:'#e8f5e9', greenT:'#1b5e20', yellow:'#fff8e1', yellowT:'#8d6e00',
  red:'#fdecea', redT:'#b71c1c', blue:'#e3f2fd', blueT:'#0d47a1', gray:'#eceff1', grayT:'#546e7a'
};

// ================== МЕНЮ ==================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚚 Автовозы')
    .addItem('Обработать все ссылки', 'processAllLinks')
    .addItem('Обновить панель и карту', 'refreshAll')
    .addItem('Синхронизировать из Битрикса', 'syncTrucksFromB24')
    .addSeparator()
    .addItem('Вебхук Битрикса', 'setB24Webhook')
    .addItem('Проверить связь с Битриксом', 'checkB24')
    .addSeparator()
    .addItem('Пересчитать сроки по истории', 'retrainEta')
    .addItem('Включить еженедельный пересчёт сроков', 'enableWeeklyEta')
    .addToUi();
}

function refreshAll() {
  rebuildExport(); rebuildDashboard();
  SpreadsheetApp.getActive().toast('Панель и карта пересобраны', 'Готово', 3);
}

// ================== ЧТЕНИЕ ЯЧЕЙКИ ==================
// Ссылка может лежать не в тексте, а в гиперссылке / «умном чипе» или в формуле
// HYPERLINK. getValue() в этих случаях вернёт подпись («Яндекс Карты»), а не URL.
function readLinkCell(sh, row, col) {
  const cell = sh.getRange(row, col);
  const out = [];

  try {
    const rt = cell.getRichTextValue();
    if (rt) {
      const direct = rt.getLinkUrl();
      if (direct) out.push(direct);
      rt.getRuns().forEach(function (run) {
        const u = run.getLinkUrl();
        if (u) out.push(u);
      });
    }
  } catch (err) { /* нет rich text — не страшно */ }

  const f = String(cell.getFormula() || '');
  const fm = f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
  if (fm) out.push(fm[1]);

  const txt = String(cell.getValue() || '').trim();
  if (txt) out.push(txt);

  return out;
}

// ================== ОБРАБОТКА РЕДАКТИРОВАНИЯ ==================
function handleEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== 'Журнал') return;

    const c1 = e.range.getColumn();
    const c2 = c1 + e.range.getNumColumns() - 1;
    const r1 = e.range.getRow();
    const rows = e.range.getNumRows();

    // Колонка C — вставили ссылку (в том числе пачкой на несколько строк)
    if (c1 <= 3 && c2 >= 3) {
      let touched = false;
      for (let i = 0; i < rows; i++) {
        if (fillCoordsForRow(sh, r1 + i)) touched = true;
      }
      if (touched) { rebuildExport(); rebuildDashboard(); }
      return;
    }
    if (c1 <= 7 && c2 >= 7) { rebuildExport(); return; }
    if (c1 <= 2 && c2 >= 2) { rebuildExport(); rebuildDashboard(); }
  } catch (err) {
    Logger.log('handleEdit: ' + err);
  }
}

// Разбирает строку Журнала: C -> lat/lon (D/E) + дата (A). true, если что-то записали.
function fillCoordsForRow(sh, row) {
  if (row < 2) return false;
  const candidates = readLinkCell(sh, row, 3);
  if (!candidates.length) return false;

  let c = null;
  for (let i = 0; i < candidates.length && !c; i++) c = resolveCoords(candidates[i]);

  if (!c) {
    sh.getRange(row, 4).setValue('ОШИБКА');
    sh.getRange(row, 5).setValue('');
    sh.getRange(row, 3).setNote('Не удалось распознать координаты.'
    + '\nПроверено: ' + candidates.join(' | ')
    + (LAST_EXPANDED ? '\nСсылка раскрылась в: ' + LAST_EXPANDED : ''));
    return true;
  }

  sh.getRange(row, 3).clearNote();
  sh.getRange(row, 1).setValue(new Date());
  sh.getRange(row, 4).setValue(c.lat);
  sh.getRange(row, 5).setValue(c.lon);
  return true;
}

// Меню: пройти по всему Журналу. Строки, где координаты уже есть, не трогаем.
function processAllLinks() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Журнал');
  const last = sh.getLastRow();
  let ok = 0, bad = 0, skipped = 0;

  for (let row = 2; row <= last; row++) {
    const lat = toNum(sh.getRange(row, 4).getValue());
    const lon = toNum(sh.getRange(row, 5).getValue());
    if (!isNaN(lat) && !isNaN(lon)) { skipped++; continue; }   // уже разобрано

    if (!fillCoordsForRow(sh, row)) continue;
    if (String(sh.getRange(row, 4).getValue()) === 'ОШИБКА') bad++; else ok++;
  }

  rebuildExport(); rebuildDashboard();
  SpreadsheetApp.getActive().toast(
    'Разобрано: ' + ok + ' · не распознано: ' + bad + ' · пропущено: ' + skipped,
    'Обработка ссылок', 6);
}

// ================== РАЗБОР ССЫЛКИ ==================
const SHORT_HOSTS = /(goo\.gl|ya\.ru|yandex\.ru\/maps\/-\/|go\.2gis\.com|link\.2gis\.com|clck\.ru|t\.co|bit\.ly|surl\.amap\.com|amap\.com\/p\/)/i;
const UA = 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1';

// куда в итоге раскрылась последняя ссылка — попадает в примечание при ошибке
var LAST_EXPANDED = '';

function safeDecode(s) {
  try { return decodeURIComponent(s); } catch (err) {
    try { return decodeURIComponent(String(s).replace(/%(?![0-9a-fA-F]{2})/g, '%25')); }
    catch (e2) { return String(s); }
  }
}

function fetchSafe(url, follow) {
  try {
    return UrlFetchApp.fetch(url, {
      followRedirects: !!follow,
      muteHttpExceptions: true,
      headers: { 'User-Agent': UA, 'Accept-Language': 'ru,en;q=0.8' }
    });
  } catch (err) { return null; }
}

// HTML-заглушки прячут настоящий адрес в meta refresh или просто в тексте страницы
function urlFromHtml(body) {
  const html = String(body || '').replace(/\\\//g, '/').replace(/&amp;/g, '&');
  const pats = [
    /<meta[^>]+http-equiv=["']?refresh["']?[^>]+url=([^"'>\s]+)/i,
    /(https?:\/\/[^"'\s<>]*\/maps\/(?:search|place|dir)\/[^"'\s<>]*)/i,
    /(https?:\/\/[^"'\s<>]*(?:ll=|whatshere|\/@)[^"'\s<>]*)/i
  ];
  for (let i = 0; i < pats.length; i++) {
    const m = html.match(pats[i]);
    if (m) return m[1];
  }
  return null;
}

function expandUrl(url) {
  for (let i = 0; i < 6; i++) {
    if (!SHORT_HOSTS.test(url)) break;

    const r = fetchSafe(url, false);
    if (!r) break;

    const h = r.getHeaders();
    const loc = h['Location'] || h['location'];
    if (loc) {
      url = loc.indexOf('http') === 0 ? loc : (url.split('/').slice(0, 3).join('/') + loc);
      continue;
    }

    // редиректа в заголовке нет — читаем тело (сначала уже полученное, потом с редиректами)
    let next = urlFromHtml(r.getContentText());
    if (!next) {
      const r2 = fetchSafe(url, true);
      if (r2) next = urlFromHtml(r2.getContentText());
    }
    if (!next) break;
    url = next;
  }
  LAST_EXPANDED = url;
  return url;
}

function pairFrom(m, order) {
  if (!m) return null;
  const a = parseFloat(m[1]), b = parseFloat(m[2]);
  if (isNaN(a) || isNaN(b)) return null;
  return order === 'lonlat' ? { lat: b, lon: a } : { lat: a, lon: b };
}

// Широта всегда |lat| <= 90. Маршрут Китай -> Москва: lat 35..75, lon 19..180.
function sane(c) {
  if (!c) return null;
  let lat = c.lat, lon = c.lon;
  if (Math.abs(lat) > 90 && Math.abs(lon) <= 90) { const t = lat; lat = lon; lon = t; }
  if (Math.abs(lat) > 90 || Math.abs(lon) > 180) return null;
  const inLat = function (x) { return x >= 35 && x <= 75; };
  const inLon = function (x) { return x >= 19 && x <= 180; };
  // обе величины формально валидны, но пара явно перевёрнута
  if (!inLat(lat) && inLat(lon) && inLon(lat)) { const t = lat; lat = lon; lon = t; }
  if (lat === 0 && lon === 0) return null;
  return { lat: lat, lon: lon };
}

// Разбор строки (URL или тела страницы).
// strict = true — только надёжные шаблоны, без «первой попавшейся пары чисел»:
// в HTML случайных дробных чисел полно.
function parseFromString(u, host, strict) {
  // Яндекс: порядок ДОЛГОТА, ШИРОТА
  if (/yandex\.|ya\.ru/i.test(host)) {
    const ry = sane(
      pairFrom(u.match(/whatshere\[point\]=(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/), 'lonlat') ||
      pairFrom(u.match(/[?&](?:ll|pt|rtext|from|to)=(-?\d+(?:\.\d+)?)[,~](-?\d+(?:\.\d+)?)/), 'lonlat') ||
      pairFrom(u.match(/[?&]text=(-?\d+(?:\.\d+)?)[,\s]+(-?\d+(?:\.\d+)?)/), 'latlon')
    );
    if (ry) return ry;
  }

  // 2ГИС: тоже долгота, широта
  if (/2gis\./i.test(host)) {
    const rg = sane(
      pairFrom(u.match(/[?&]m=(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/), 'lonlat') ||
      pairFrom(u.match(/\/(-?\d{1,3}\.\d+),(-?\d{1,3}\.\d+)(?:[\/?#]|$)/), 'lonlat')
    );
    if (rg) return rg;
  }

  // Google / Apple / Amap / прочие: широта, долгота
  const r = sane(
    pairFrom(u.match(/!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/), 'latlon') ||
    pairFrom(u.match(/\/@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/), 'latlon') ||
    // maps.app.goo.gl раскрывается в /maps/search/51.35,+110.39
    pairFrom(u.match(/\/maps\/(?:search|place|dir)\/(-?\d+(?:\.\d+)?)\s*,[\s+]*(-?\d+(?:\.\d+)?)/), 'latlon') ||
    pairFrom(u.match(/[?&](?:q|query|ll|sll|daddr|saddr|center|destination|coordinate|location)=(-?\d+(?:\.\d+)?),[\s+]*(-?\d+(?:\.\d+)?)/), 'latlon') ||
    pairFrom(u.match(/"latitude"\s*:\s*(-?\d+(?:\.\d+)?)\s*,\s*"longitude"\s*:\s*(-?\d+(?:\.\d+)?)/), 'latlon') ||
    pairFrom(u.match(/[?&]lat=(-?\d+(?:\.\d+)?)[\s\S]{0,80}?[?&](?:lon|lng|long)=(-?\d+(?:\.\d+)?)/), 'latlon')
  );
  if (r) return r;

  if (strict) return null;

  // последний шанс: первая пара дробных чисел где угодно в строке
  return sane(pairFrom(u.match(/(-?\d{1,3}\.\d{3,})\s*[,;][\s+]*(-?\d{1,3}\.\d{3,})/), 'latlon'));
}

function resolveCoords(input) {
  const s = String(input || '').trim();
  LAST_EXPANDED = '';
  if (!s) return null;

  // 1. Голые координаты: «55.7558, 37.6173», «55,7558 37,6173», «55.7558;37.6173»
  const bare = s.match(/^\(?\s*(-?\d{1,3}(?:[.,]\d+)?)\s*[,;\s]\s*(-?\d{1,3}(?:[.,]\d+)?)\s*\)?$/);
  if (bare) {
    const r0 = sane({ lat: parseFloat(bare[1].replace(',', '.')), lon: parseFloat(bare[2].replace(',', '.')) });
    if (r0) return r0;
  }

  // 2. Вытаскиваем URL из текста («Вот я: https://…»)
  const urlM = s.match(/https?:\/\/\S+/);
  let url = (urlM ? urlM[0] : s).replace(/[)\],.;]+$/, '');

  url = expandUrl(url);
  const u = safeDecode(url);
  const host = (u.match(/^https?:\/\/([^\/?#]+)/) || [])[1] || '';

  // 3. Разбираем сам адрес
  const r = parseFromString(u, host, false);
  if (r) return r;

  // 4. Не вышло — читаем страницу и ищем координаты в её содержимом
  if (/^https?:/i.test(url)) {
    const page = fetchSafe(url, true);
    if (page) {
      const body = safeDecode(String(page.getContentText()).slice(0, 400000))
        .replace(/\\\//g, '/').replace(/&amp;/g, '&');
      const rb = parseFromString(body, host, true);
      if (rb) return rb;
    }
  }

  return null;
}

function haversine(lat1, lon1, lat2, lon2) {
  const R = 6371, toRad = x => x * Math.PI / 180;
  const dLat = toRad(lat2 - lat1), dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat/2)**2 + Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*Math.sin(dLon/2)**2;
  return 2 * R * Math.asin(Math.sqrt(a));
}
function distanceToMoscow(lat, lon) { return Math.round(haversine(lat, lon, MSK.lat, MSK.lon) * ROAD_FACTOR); }
function etaDays(km) { if (km < 30) return 0; return Math.max(1, Math.ceil(km / (AVG_SPEED_KMH * HOURS_PER_DAY))); }

// ================== СРОКИ ПО ИСТОРИИ ==================
// Коридор «обычно … до»: сколько дней реально оставалось до СВХ с похожего
// расстояния — медиана и 80% по прошлым рейсам. Учится на Журнале и датах
// прибытия на СВХ из Битрикса, хранится в свойствах скрипта (ETA_MODEL),
// пересчитывается раз в неделю (retrainEta). Забайкальск — отдельно:
// там весь разброс — ожидание на границе.
const ETA_PROP = 'ETA_MODEL';
const ZAB = { lat: 49.6406, lon: 117.3260 };
const ZAB_RADIUS_KM = 40;
const ETA_KNOTS = [500, 1500, 2500, 3500, 4500, 5500, 6400];
const ETA_WINDOW_KM = 700;       // окно вокруг узла, чтобы хватало точек
const ETA_MIN_POINTS = 80;       // меньше — таблицу не трогаем
// стартовая таблица: 12.09.2026, 256 отметок за июль–сентябрь; [км, обычно, 80% за]
const ETA_SEED = {
  knots: [[500,1.4,2.3],[1500,1.8,2.5],[2500,2.9,3.7],[3500,4.6,5.6],[4500,5.6,6.6],[5500,8.4,9.4],[6400,9.3,9.7]],
  zab: [11.1, 14.4], n: 256, updated: '2026-09-12'
};
let ETA_CACHE = null;

function etaModel() {
  if (ETA_CACHE) return ETA_CACHE;
  const raw = PropertiesService.getScriptProperties().getProperty(ETA_PROP);
  if (raw) {
    try { const m = JSON.parse(raw); if (m && m.knots && m.knots.length) return (ETA_CACHE = m); } catch (e) {}
  }
  return (ETA_CACHE = ETA_SEED);
}

function isAtZab(lat, lon) { return haversine(lat, lon, ZAB.lat, ZAB.lon) <= ZAB_RADIUS_KM; }

// { lo, hi } в целых днях; 0/0 — уже в Москве
function etaRange(km, lat, lon, model) {
  if (km < 30) return { lo: 0, hi: 0 };
  const m = model || etaModel();
  let lo, hi;
  if (m.zab && lat !== undefined && isAtZab(lat, lon)) { lo = m.zab[0]; hi = m.zab[1]; }
  else {
    const t = m.knots;
    if (km <= t[0][0]) { const k = km / t[0][0]; lo = t[0][1] * k; hi = t[0][2] * k; }
    else if (km >= t[t.length - 1][0]) { lo = t[t.length - 1][1]; hi = t[t.length - 1][2]; }
    else for (let i = 1; i < t.length; i++) {
      if (km <= t[i][0]) {
        const f = (km - t[i - 1][0]) / (t[i][0] - t[i - 1][0]);
        lo = t[i - 1][1] + (t[i][1] - t[i - 1][1]) * f;
        hi = t[i - 1][2] + (t[i][2] - t[i - 1][2]) * f;
        break;
      }
    }
  }
  const L = Math.max(1, Math.round(lo)), H = Math.max(L, Math.ceil(hi));
  return { lo: L, hi: H };
}

function etaRangeText(r) { return r.hi === 0 ? 'прибыл' : (r.lo === r.hi ? String(r.hi) : r.lo + '–' + r.hi); }

function normKey(s) { return String(s).replace(/[\s\-–—]/g, '').toUpperCase(); }

// Чистая функция (без обращений к Google) — её же гоняем в Node для проверки.
// points: [{ t: мс, keys: [...], km, zab }], svh: { КЛЮЧ: [мс, ...] }
function fitEtaModel(points, svh) {
  const DAY = 864e5, lab = [];
  points.forEach(p => {
    if (p.km < 150) return;                                   // уже в Москве
    let best = null;
    p.keys.forEach(k => (svh[k] || []).forEach(d => {
      if (d >= p.t - DAY && d <= p.t + 40 * DAY && (best === null || d < best)) best = d;
    }));
    if (best !== null) lab.push({ km: p.km, zab: p.zab, left: Math.max(0, (best - p.t) / DAY) });
  });
  const q = (v, x) => {
    v = v.slice().sort((a, b) => a - b);
    const i = (v.length - 1) * x, lo = Math.floor(i), hi = Math.ceil(i);
    return v[lo] + (v[hi] - v[lo]) * (i - lo);
  };
  const road = lab.filter(r => !r.zab), knots = [];
  ETA_KNOTS.forEach(c => {
    const w = road.filter(r => Math.abs(r.km - c) <= ETA_WINDOW_KM).map(r => r.left);
    if (w.length >= 6) knots.push([c, q(w, .5), q(w, .8)]);
  });
  for (let i = 1; i < knots.length; i++) {                    // дальше — не быстрее
    knots[i][1] = Math.max(knots[i][1], knots[i - 1][1]);
    knots[i][2] = Math.max(knots[i][2], knots[i - 1][2]);
  }
  const z = lab.filter(r => r.zab).map(r => r.left);
  const r1 = v => Math.round(v * 10) / 10;
  return {
    knots: knots.map(k => [k[0], r1(k[1]), r1(k[2])]),
    zab: z.length >= 6 ? [r1(q(z, .5)), r1(q(z, .8))] : null,
    n: lab.length
  };
}

function sheetDateMs(v) {
  if (v instanceof Date) return v.getTime();
  const m = String(v || '').match(/(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)).getTime();
  const t = new Date(v).getTime();
  return isNaN(t) ? NaN : t;
}

// Только место ТО «Чита-Москва»: у просто «Читы» СВХ в Чите, это ~3 дня,
// а не ~11 до Москвы — такие сделки занизили бы сроки.
const ETA_TO_PLACE = 'Чита-Москва';

// даты прибытия на СВХ по номеру автовоза за последние полгода (только ETA_TO_PLACE)
function fetchSvhDatesByTruck() {
  const toMap = getToPlaceMap();                              // ID → название места ТО
  const placeIds = Object.keys(toMap).filter(id => toMap[id] === ETA_TO_PLACE);
  if (!placeIds.length) throw new Error('В справочнике мест ТО нет «' + ETA_TO_PLACE + '»');
  const since = Utilities.formatDate(new Date(Date.now() - 180 * 864e5), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const out = {};
  let start = 0;
  while (true) {
    const filter = { CATEGORY_ID: B24_CATEGORY };
    filter['>=' + B24_ARRIVED] = since;
    const res = UrlFetchApp.fetch(b24() + 'crm.deal.list.json', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ filter: filter, select: ['ID', B24_TRUCK, B24_ARRIVED, B24_TO_PLACE], start: start })
    });
    const d = JSON.parse(res.getContentText());
    if (d.error) throw new Error(d.error_description || d.error);
    (d.result || []).forEach(deal => {
      if (isEmpty(deal[B24_ARRIVED])) return;
      const tp = deal[B24_TO_PLACE];
      const code = String(Array.isArray(tp) ? tp[0] : tp).trim();
      if (placeIds.indexOf(code) === -1) return;
      // Дата СВХ приходит с временем прибытия («2026-07-20T04:22:00+03:00»).
      // Берём дату и время «как на часах», без пересчёта поясов — так же сравниваются
      // отметки Журнала. Так посчитана и стартовая таблица ETA_SEED.
      const dt = String(deal[B24_ARRIVED]).match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/);
      if (!dt) return;
      const ms = new Date(+dt[1], +dt[2] - 1, +dt[3], +(dt[4] || 0), +(dt[5] || 0)).getTime();
      const nums = deal[B24_TRUCK];
      (Array.isArray(nums) ? nums : [nums]).forEach(n => {
        if (isEmpty(n)) return;
        const k = normKey(n);
        (out[k] = out[k] || []).push(ms);
      });
    });
    if (d.next === undefined) break;
    start = d.next; Utilities.sleep(300);
  }
  return out;
}

// Меню и недельный триггер: пересчитать таблицу сроков по свежей истории
function retrainEta() {
  const points = [];
  SpreadsheetApp.getActive().getSheetByName('Журнал').getDataRange().getValues().slice(1).forEach(r => {
    const t = sheetDateMs(r[0]), lat = toNum(r[3]), lon = toNum(r[4]);
    if (!r[1] || isNaN(t) || isNaN(lat) || isNaN(lon)) return;
    const name = String(r[1]).trim();
    const keys = [normKey(name)].concat(
      name.split(/\s*[-–—]\s*/).filter(s => s.trim().length >= 5).map(normKey));
    points.push({ t: t, keys: keys, km: distanceToMoscow(lat, lon), zab: isAtZab(lat, lon) });
  });
  const m = fitEtaModel(points, fetchSvhDatesByTruck());
  const ok = m.n >= ETA_MIN_POINTS && m.knots.length >= 4;
  if (ok) {
    m.updated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    PropertiesService.getScriptProperties().setProperty(ETA_PROP, JSON.stringify(m));
    ETA_CACHE = m;
    rebuildExport(); rebuildDashboard();
  }
  Logger.log((ok ? 'Сроки обновлены: ' : 'Мало данных, таблица не тронута: ') + JSON.stringify(m));
  try {
    SpreadsheetApp.getActive().toast(ok
      ? 'Сроки пересчитаны по ' + m.n + ' отметкам'
      : 'Мало данных (' + m.n + ' отметок) — оставлена прежняя таблица сроков', 'Сроки', 6);
  } catch (e) { /* запуск по триггеру — тоста нет */ }
  return m;
}

function enableWeeklyEta() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'retrainEta')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('retrainEta').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  SpreadsheetApp.getActive().toast('Сроки будут пересчитываться каждый понедельник около 6 утра', 'Готово', 5);
}
function toNum(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return NaN;
  return parseFloat(String(v).replace(',', '.').trim());
}
function dayKey(ms) { const d = new Date(ms); const mm=('0'+(d.getMonth()+1)).slice(-2); const dd=('0'+d.getDate()).slice(-2); return d.getFullYear()+'-'+mm+'-'+dd; }

function styleHeader(sh, cols) {
  sh.getRange(1,1,1,cols).setBackground(CLR.headBg).setFontColor(CLR.headText)
    .setFontWeight('bold').setFontSize(11).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(1, 34); freezeRows(sh,1);
}

function autoFitAll(sh, cols) {
  for (var c = 1; c <= cols; c++) sh.autoResizeColumn(c);
}
function centerAll(sh, cols, rows) {
  if (rows > 0) sh.getRange(1,1,rows,cols).setHorizontalAlignment('center');
}

function styleBody(sh, startRow, numRows, cols) {
  if (numRows <= 0) return;
  sh.getRange(startRow,1,numRows,cols)
    .setBorder(true,true,true,true,true,true,CLR.border,SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setFontSize(10);
  for (let i = 0; i < numRows; i++)
    sh.getRange(startRow+i,1,1,cols).setBackground(i % 2 === 0 ? CLR.band1 : CLR.band2);
}

function getHistoryByTruck() {
  const byTruck = {};
  SpreadsheetApp.getActive().getSheetByName('Журнал').getDataRange().getValues().slice(1)
    .forEach(r => {
      const date=r[0],truck=r[1],note=r[6],prev=r[7];
      const lat=toNum(r[3]), lon=toNum(r[4]);
      if (!date || !truck || isNaN(lat) || isNaN(lon)) return;
      if (!looksLikeTruck(truck)) return;
      const key = String(truck).trim();
      if (!byTruck[key]) byTruck[key] = [];
      byTruck[key].push({ t:new Date(date).getTime(), lat, lon, date,
        note:String(note||'').trim(), prev:String(prev||'').trim() });
    });
  const result = {};
  Object.keys(byTruck).forEach(k => {
    const arr = byTruck[k].sort((a,b) => a.t - b.t);
    const perDay = {}; arr.forEach(p => { perDay[dayKey(p.t)] = p; });
    const days = Object.keys(perDay).sort().map(d => perDay[d]);
    const last = days[days.length-1];
    const prev = days.length >= 2 ? days[days.length-2] : null;
    let legKm = null, stuck = false;
    if (prev) {
      legKm = Math.round(haversine(prev.lat,prev.lon,last.lat,last.lon) * ROAD_FACTOR);
      const daysGap = Math.max((last.t - prev.t) / 864e5, 1);
      if (legKm / daysGap < MIN_KM_PER_DAY) stuck = true;
    }
    // трек текущего рейса: от последней отметки назад, пока нет разрыва больше
    // 10 дней и пока не упёрлись в Москву перед дальней точкой (конец прошлого рейса)
    const trip = [last];
    for (let i = days.length - 2; i >= 0; i--) {
      const p = days[i];
      if ((trip[0].t - p.t) / 864e5 > 10) break;
      if (distanceToMoscow(p.lat, p.lon) < 150 && distanceToMoscow(trip[0].lat, trip[0].lon) >= 150) break;
      trip.unshift(p);
    }
    const track = trip.map(p => [Math.round(p.lat * 1e5) / 1e5, Math.round(p.lon * 1e5) / 1e5, p.t]);
    result[k] = { last, note:last.note, prev:last.prev, legKm, stuck, track };
  });
  return result;
}

function refreshDropdown() {
  const ss = SpreadsheetApp.getActive();
  const jr = ss.getSheetByName('Журнал').getRange('B2:B1000');
  jr.clearDataValidations();
  const sh = ss.getSheetByName('Автовозы');
  if (!sh) return;
  const trucks = sh.getDataRange().getValues().slice(2)
    .filter(r => r[0] && String(r[1]).trim().toLowerCase() === 'да')
    .map(r => String(r[0]).trim());
  if (!trucks.length) return;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(trucks, true).setAllowInvalid(true).build();
  jr.setDataValidation(rule);
}

function styleJournal() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Журнал');
  if (!sh) return;
  sh.getRange(1,1,1,8).setValues([['Дата','Автовоз','Ссылка / координаты','lat','lon','Комментарий','Примечание','Прежний номер']]);
  styleHeader(sh, 8);
  sh.getRange('A2:A1000').setNumberFormat('dd.MM.yyyy HH:mm');
  sh.getRange('D2:E1000').setNumberFormat('0.000000');
  sh.getRange('A1:G1000').setHorizontalAlignment('center');
  sh.getRange('D1:E1000').setFontColor(CLR.grayT);
  autoFitAll(sh, 7);
}

function rebuildExport() {
  const now = Date.now();
  const hist = getHistoryByTruck();
  const active = {}, hidden = {};
  const av = SpreadsheetApp.getActive().getSheetByName('Автовозы');
  if (av) av.getDataRange().getValues().slice(2).forEach(r => {
    if (!r[0]) return;
    const t = String(r[0]).trim();
    if (String(r[1]).trim().toLowerCase() === 'да') active[t] = true;
    if (String(r[3]).trim().toLowerCase() === 'да') hidden[t] = true;
  });
  const out = Object.keys(hist)
    .filter(k => active[k])
    .filter(k => !hidden[k])
    .filter(k => (now - hist[k].last.t) / 864e5 <= MAX_AGE_DAYS)
    .sort()
    .map(k => {
      const h = hist[k], r = h.last, km = distanceToMoscow(r.lat, r.lon);
      const eta = etaRange(km, r.lat, r.lon);
      // ETA (кол. 6) — верхняя граница: старая карта продолжает работать как раньше
      return [r.lat, r.lon, k, r.date, km, eta.hi, r.note, h.legKm===null?'':h.legKm, h.stuck?1:0, h.prev||'',
              eta.lo, eta.hi, JSON.stringify(h.track || [])];
    });
  const sh = SpreadsheetApp.getActive().getSheetByName('Экспорт');
  sh.clear();
  sh.getRange(1,1,1,13).setValues([['lat','lon','Автовоз','Дата','До МСК','ETA','Коммент','Прошёл','Стоит','Прежний',
                                     'Срок от','Срок до','Трек']])
    .setBackground(CLR.gray).setFontColor(CLR.grayT).setFontWeight('bold').setFontSize(9);
  freezeRows(sh,1);
  if (out.length) sh.getRange(2,1,out.length,13).setValues(out);
  sh.getRange(1,1,Math.max(out.length+1,1),12).setHorizontalAlignment('center');
  autoFitAll(sh, 12);
  sh.setColumnWidth(13, 140);                    // трек — длинный JSON, не растягиваем
}

function rebuildDashboard() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Панель');
  if (!sh) sh = ss.insertSheet('Панель');
  const active = {}, hideMap = {}, arriving = {}, batchInfo = {};
  const av = ss.getSheetByName('Автовозы');
  if (av) av.getDataRange().getValues().slice(2).forEach(r => {
    if (!r[0]) return;
    const t = String(r[0]).trim();
    if (String(r[1]).trim().toLowerCase() === 'да') active[t] = true;
    if (String(r[3]).trim().toLowerCase() === 'да') hideMap[t] = true;
    if (String(r[5]).trim() === 'Прибывает') arriving[t] = true;
    batchInfo[t] = String(r[4] || '');
  });

  const now = Date.now();
  const hist = getHistoryByTruck();

  // на панель: ВСЕ активные автовозы, кроме скрытых по месту ТО.
  // Без локации → статус "Нет данных". "Прибывает" — виден, но не требует обновления.
  const keys = Object.keys(active).filter(k => {
    if (hideMap[k] && !arriving[k]) return false;        // скрыт по месту ТО (но "прибывает" оставляем)
    return true;
  });

  const rows = keys.map(k => {
    const h = hist[k];
    const fresh = h && (now - h.last.t) / 864e5 <= MAX_AGE_DAYS;

    // нет локации ИЛИ протухла
    if (!fresh) {
      const status = arriving[k] ? 'ПРИБЫВАЕТ' : 'НЕТ ДАННЫХ';
      return { truck:k, date:'', days:9999, km:'—', eta:'—', leg:'—',
               stuck:false, status:status, batch:batchInfo[k] };
    }

    const rec = h.last;
    const days = Math.floor((now - rec.t) / 864e5);
    const km = distanceToMoscow(rec.lat, rec.lon);
    let status;
    if (arriving[k]) status = 'ПРИБЫВАЕТ';
    else if (h.stuck) status = 'СТОИТ';
    else status = days<=2?'Свежие':days<=4?'Скоро обновить':'ОБНОВИТЬ';
    return { truck:k, date:rec.date, days, km, eta:etaRangeText(etaRange(km, rec.lat, rec.lon)),
      leg:h.legKm===null?'—':h.legKm, stuck:h.stuck, status:status, batch:batchInfo[k] };
  });

  // порядок: нет данных → стоят → обновить → скоро → свежие → прибывают
  rows.sort((a,b) => {
    const rank = x => x.status==='НЕТ ДАННЫХ'?0 : x.status==='СТОИТ'?1 : x.status==='ОБНОВИТЬ'?2
                     : x.status==='Скоро обновить'?3 : x.status==='ПРИБЫВАЕТ'?5 : 4;
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    return b.days - a.days;
  });
  sh.clear();
  const stuckN = rows.filter(r => r.stuck).length;
  const updN   = rows.filter(r => r.status === 'ОБНОВИТЬ').length;
  const arrN   = rows.filter(r => r.status === 'ПРИБЫВАЕТ').length;
  const noDataN= rows.filter(r => r.status === 'НЕТ ДАННЫХ').length;
  sh.getRange(1,1,1,8).merge()
    .setValue('  Всего: '+rows.length+'     ·     нет данных: '+noDataN+'     ·     обновить: '+updN+'     ·     стоят: '+stuckN+'     ·     прибывают: '+arrN)
    .setBackground(CLR.headBg).setFontColor('#ffffff').setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle');
  sh.setRowHeight(1, 38);
  sh.getRange(2,1,1,8).setValues([['Автовоз','Обновлено','Дней назад','Прошёл, км','До Москвы, км','Прибытие, дней','Партия','Статус']])
    .setBackground('#2e5a80').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(2, 30); freezeRows(sh,2);
  if (rows.length) {
    sh.getRange(3,7,rows.length,1).setNumberFormat('@');   // партия как текст
    sh.getRange(3,6,rows.length,1).setNumberFormat('@');   // срок «5–6» как текст, не дата
    sh.getRange(3,1,rows.length,8).setValues(rows.map(r => [r.truck,r.date,r.days===9999?'—':r.days,r.leg,r.km,r.eta,r.batch||'',r.status]));
    for (let i = 0; i < rows.length; i++) {
      const range = sh.getRange(i+3,1,1,8), st = sh.getRange(i+3,8,1,1);
      if (rows[i].status === 'НЕТ ДАННЫХ') { range.setBackground('#fbe9e7'); st.setFontColor('#bf360c'); }
      else if (rows[i].status === 'ПРИБЫВАЕТ') { range.setBackground('#e8f0fe'); st.setFontColor('#1a56b8'); }
      else if (rows[i].stuck) { range.setBackground(CLR.blue); st.setFontColor(CLR.blueT); }
      else if (rows[i].status === 'ОБНОВИТЬ') { range.setBackground(CLR.red); st.setFontColor(CLR.redT); }
      else if (rows[i].status === 'Скоро обновить') { range.setBackground(CLR.yellow); st.setFontColor(CLR.yellowT); }
      else { range.setBackground(CLR.green); st.setFontColor(CLR.greenT); }
    }
    sh.getRange(3,2,rows.length,1).setNumberFormat('dd.MM.yyyy HH:mm');
    sh.getRange(3,1,rows.length,8).setBorder(true,true,true,true,true,true,CLR.border,SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(3,1,rows.length,8).setHorizontalAlignment('center');
    sh.getRange(3,8,rows.length,1).setFontWeight('bold');
  }
  autoFitAll(sh, 8);
}

function isEmpty(v) {
  if (v === null || v === undefined) return true;
  if (Array.isArray(v)) return v.length === 0;
  return String(v).trim() === '';
}

// номер автовоза валиден, если это НЕ дата и НЕ чисто число
function looksLikeTruck(s) {
  s = String(s).trim();
  if (!s) return false;
  if (s.replace(/[-\s—]/g, '').length < 3) return false; // мусор: "-", "—", пусто, слишком короткий
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return false;
  if (/\d{1,2}\.\d{1,2}\.\d{4}/.test(s)) return false;
  if (/^\d+([.,]\d+)?$/.test(s)) return false;
  if (/(GMT|\bMon\b|\bTue\b|\bWed\b|\bThu\b|\bFri\b|\bSat\b|\bSun\b)/.test(s)) return false;
  return true;
}

function getToPlaceMap() {
  const res = UrlFetchApp.fetch(b24() + 'crm.deal.userfield.list.json', {
    method:'post', contentType:'application/json',
    payload: JSON.stringify({ filter: { FIELD_NAME: B24_TO_PLACE } }), muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  const f = (data.result || [])[0];
  const map = {};
  if (f && f.LIST) f.LIST.forEach(item => {
    let name = String(item.VALUE).replace(/[\u4e00-\u9fff].*$/, '').trim();
    if (!name) name = String(item.VALUE).trim();
    map[String(item.ID)] = name;
  });
  return map;
}

// Сколько дней назад по дате отправки брать сделки для списка «в пути».
// Текущие партии всех автовозов укладываются в 45 дней (проверено 12.09.2026
// на всей воронке: 45/60/90/120 дней дают тот же список, что и полная выгрузка),
// 120 — с запасом. Было 145 страниц по 50 сделок, стало ~31: запуск в разы короче.
const B24_SYNC_DAYS = 120;

// Закрепить строки, только если ещё не закреплены. Google иногда отказывает
// в этой косметической операции («You do not have permission to access the
// requested document») — из-за неё падал весь dailyRefresh. Теперь не падает.
function freezeRows(sheet, n) {
  try {
    if (sheet.getFrozenRows() !== n) sheet.setFrozenRows(n);
  } catch (e) {
    Logger.log('Не удалось закрепить строки на листе «' + sheet.getName() + '»: ' + e);
  }
}

// Запрос к Битриксу с повтором при временных сбоях (не JSON, 5xx, лимит запросов).
function b24Post(method, payload) {
  let lastErr = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const res = UrlFetchApp.fetch(b24() + method, {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      const data = JSON.parse(res.getContentText());
      if (data.error === 'QUERY_LIMIT_EXCEEDED') throw new Error('Битрикс: лимит запросов');
      if (data.error) throw new Error(data.error_description || data.error);   // настоящая ошибка — не повторяем
      return data;
    } catch (e) {
      lastErr = e;
      if (/Битрикс: лимит|JSON|Unexpected token|Address unavailable|timed out|DNS|50\d/.test(String(e))) {
        Utilities.sleep(2000 * attempt);
        continue;
      }
      throw e;
    }
  }
  throw lastErr;
}

function syncTrucksFromB24() {
  const toMap = getToPlaceMap();
  const earlySet = {}; B24_STAGES.forEach(x => earlySet[x] = true);

  // Сделки воронки с отправкой за последние B24_SYNC_DAYS дней: по автовозу копим
  // машины с датой отправки и стадией. Старые партии всё равно отсекаются ниже.
  const since = Utilities.formatDate(new Date(Date.now() - B24_SYNC_DAYS * 864e5), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const filter = { CATEGORY_ID: B24_CATEGORY };
  filter['>=' + B24_SHIP_DATE] = since;
  const deals = {};   // номер → [ { shipMs, stage, toPlace } ]
  let start = 0;
  while (true) {
    const data = b24Post('crm.deal.list.json', {
      filter: filter,
      select: [B24_TRUCK, B24_TO_PLACE, 'STAGE_ID', B24_SHIP_DATE, B24_ARRIVED], start: start
    });
    (data.result || []).forEach(deal => {
      const num = deal[B24_TRUCK];
      if (isEmpty(num)) return;
      const vals = Array.isArray(num) ? num : [num];
      const shipRaw = deal[B24_SHIP_DATE];
      const shipMs = isEmpty(shipRaw) ? null : new Date(shipRaw).getTime();
      const stage = String(deal['STAGE_ID'] || '');
      const svhDate = !isEmpty(deal[B24_ARRIVED]);   // дата прибытия на СВХ заполнена
      const tp = deal[B24_TO_PLACE];
      let place = '';
      if (!isEmpty(tp)) { const code = Array.isArray(tp) ? String(tp[0]).trim() : String(tp).trim(); place = toMap[code] || code; }
      vals.forEach(x => {
        const s = String(x).trim();
        if (!looksLikeTruck(s)) return;   // мусор (даты, числа) — мимо
        if (!deals[s]) deals[s] = [];
        deals[s].push({ shipMs, stage, place, svhDate });
      });
    });
    if (data.next === undefined) break;
    start = data.next; Utilities.sleep(300);
  }

  // По каждому автовозу выделяем ТЕКУЩУЮ партию по самой свежей дате отправки (± B24_BATCH_DAYS)
  const agg = {};
  Object.keys(deals).forEach(k => {
    const arr = deals[k];
    const shipDates = arr.map(d => d.shipMs).filter(x => x !== null);
    if (!shipDates.length) return;
    const latest = Math.max.apply(null, shipDates);
    const win = B24_BATCH_DAYS * 864e5;

    const batch = arr.filter(d => d.shipMs !== null && (latest - d.shipMs) <= win);
    if (!batch.length) return;

    let total = batch.length, arrived = 0, inTransitCnt = 0, place = '';
    batch.forEach(d => {
      // едет только если на ранней стадии И нет даты прибытия на СВХ
      const stillMoving = earlySet[d.stage] && !d.svhDate;
      if (stillMoving) inTransitCnt++;
      else arrived++;                          // дата СВХ есть ИЛИ поздняя стадия = прибыла
      if (!place && d.place) place = d.place;
    });

    const hasEarly = inTransitCnt > 0;                 // есть ли едущие
    const arriving = hasEarly && inTransitCnt <= B24_ARRIVING_LEFT;  // осталось мало → "прибывает"

    agg[k] = { total, arrived, inTransit: inTransitCnt, toPlace: place, hasEarly, arriving };
  });

  // в список берём всех, у кого есть едущие машины (и полноценно в пути, и "прибывает")
  const inTransit = {};
  Object.keys(agg).forEach(k => { if (agg[k].hasEarly) inTransit[k] = agg[k]; });
  const list = Object.keys(inTransit).sort();
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Автовозы');
  if (!sh) sh = ss.insertSheet('Автовозы');
  sh.clear();
  sh.getRange('A1:Z1000').clearDataValidations();

  // строка 1 — отметка времени синхронизации
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  sh.getRange(1,1,1,5).merge()
    .setValue('  Обновлено из Битрикса: ' + stamp + '     ·     автовозов в пути: ' + list.length)
    .setBackground('#2e5a80').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11).setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);

  // строка 2 — заголовки
  sh.getRange('E1:E1000').setNumberFormat('@');
  sh.getRange(2,1,1,6).setValues([['Автовоз','Активен','Место ТО','Скрыт с карты','На СВХ / в партии','Статус']])
    .setBackground(CLR.headBg).setFontColor(CLR.headText).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(2, 30);
  freezeRows(sh,2);

  if (list.length) {
    const values = list.map(t => {
      const a = inTransit[t];
      const hiddenByPlace = B24_HIDE_TO.indexOf(a.toPlace) !== -1;
      // "прибывает" — тоже скрываем с карты, но помечаем статусом
      const hidden = hiddenByPlace || a.arriving;
      const status = a.arriving ? 'Прибывает' : 'В пути';
      return [t, 'да', a.toPlace, hidden ? 'да' : '', a.arrived + '/' + a.total, status];
    });
    sh.getRange(3,5,values.length,1).setNumberFormat('@');  // партия как ТЕКСТ (иначе 8/9 → дата)
    sh.getRange(3,1,values.length,6).setValues(values);
    styleBody(sh, 3, values.length, 6);
    sh.getRange(3,1,values.length,6).setHorizontalAlignment('center');
    for (let i = 0; i < values.length; i++) {
      if (values[i][5] === 'Прибывает')
        sh.getRange(i+3,1,1,6).setBackground('#e8f0fe').setFontColor('#1a56b8');
      else if (values[i][3] === 'да')
        sh.getRange(i+3,1,1,6).setBackground('#f0f0f0').setFontColor('#9e9e9e');
    }
  }
  autoFitAll(sh, 6);
  refreshDropdown();
  Logger.log('В пути: ' + list.length + ', на нужных стадиях всего автовозов: ' + Object.keys(agg).length);
  return list.length;
}

function styleAllSheets() {
  styleJournal();
  syncTrucksFromB24();
  rebuildExport();
  rebuildDashboard();
  SpreadsheetApp.getActive().toast('Оформление применено', 'Готово', 3);
}

// Если синхронизация с Битриксом упала, карта и Панель всё равно пересобираются
// по текущему списку автовозов; ошибка пробрасывается в конце — видна в журнале.
function dailyRefresh() {
  let syncErr = null;
  try { syncTrucksFromB24(); } catch (e) { syncErr = e; Logger.log('syncTrucksFromB24: ' + e); }
  rebuildExport();
  rebuildDashboard();
  if (syncErr) throw syncErr;
}

function doGet() {
  const rows = SpreadsheetApp.getActive().getSheetByName('Экспорт').getDataRange().getValues().slice(1)
    .filter(r => !isNaN(toNum(r[0])) && !isNaN(toNum(r[1])));
  const num = v => (v === '' || v === undefined || v === null || isNaN(toNum(v))) ? null : toNum(v);
  const trackOf = v => { try { const a = JSON.parse(v); return Array.isArray(a) ? a : []; } catch (e) { return []; } };
  const data = rows.map(r => ({
    lat:toNum(r[0]), lon:toNum(r[1]), truck:String(r[2]), ts:new Date(r[3]).toISOString(),
    km:r[4], eta:r[5], note:String(r[6]||''), leg:r[7]===''?null:r[7], stuck:r[8]===1, prev:String(r[9]||''),
    etaLo:num(r[10]), etaHi:num(r[11]), track:trackOf(r[12])
  }));
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function findTruckField() {
  const res = UrlFetchApp.fetch(b24() + 'crm.deal.userfield.list.json', {
    method:'post', contentType:'application/json',
    payload: JSON.stringify({ order:{}, filter:{} }), muteHttpExceptions:true
  });
  const data = JSON.parse(res.getContentText());
  const out = [];
  (data.result || []).forEach(f => {
    let label = f.EDIT_FORM_LABEL;
    if (label && typeof label === 'object') label = label.ru || label.en || '';
    out.push(f.FIELD_NAME + '  →  ' + (label || ''));
  });
  Logger.log(out.join('\n'));
}
