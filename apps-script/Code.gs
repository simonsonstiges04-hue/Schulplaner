/**
 * Schulplaner – Google Apps Script (Backend) · Version 2
 * ------------------------------------------------------------
 * Was das Skript macht:
 *  1. Liest deine Google-Kalender (nur LESEN – es kann nichts eintragen oder löschen).
 *  2. Speichert eine Sicherung deiner privaten App-Daten (HÜ, Prüfungen) in diesem Skript.
 *     Die liegen nur in deinem Konto – nicht im Kalender, deine Eltern sehen davon nichts.
 *  3. Schickt Erinnerungen als Benachrichtigung aufs Handy (alle 5 Minuten wird geprüft)
 *     und optional jeden Abend eine Übersicht, was morgen fällig ist.
 *
 * Nach dem Einfügen einmal die Funktion "setup" ausführen und neu bereitstellen.
 */

const TZ = 'Europe/Vienna';
const CHUNK = 8000;            // Zeichen pro Speicherblock (Properties-Limit 9 KB)
const MAX_RANGE_DAYS = 200;    // maximal abfragbarer Zeitraum pro Anfrage

/* ============ Einrichtung ============ */

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) {
    props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 24));
  }
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkReminders').timeBased().everyMinutes(5).create();
  const n = CalendarApp.getAllCalendars().length;
  // einmal ins Internet, damit die Berechtigung für Benachrichtigungen abgefragt wird
  try { UrlFetchApp.fetch('https://www.google.com/generate_204', { muteHttpExceptions: true }); } catch (err) {}
  Logger.log('✅ Fertig. ' + n + ' Kalender gefunden.');
  Logger.log('🔑 Dein Token (in der App eintragen): ' + props.getProperty('TOKEN'));
}

/** Falls du den Token vergessen hast: diese Funktion ausführen und ins Protokoll schauen. */
function showToken() {
  Logger.log(PropertiesService.getScriptProperties().getProperty('TOKEN'));
}

/** Test-Benachrichtigung an alle Geräte, auf denen Benachrichtigungen aktiviert sind. */
function testPush() {
  Logger.log(JSON.stringify(notify_([{ id: 'test|' + Date.now(), title: '✅ Schulplaner-Test', body: 'Benachrichtigungen funktionieren.' }])));
}

/* ============ Web-Schnittstelle ============ */

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(body);
}

function handle_(p) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || p.token !== token) return json_({ ok: false, error: 'Falscher Token' });
  try {
    switch (p.action) {
      case 'ping':     return json_({ ok: true, now: nowLocal_(), v: 2 });
      case 'events':   return json_(Object.assign({ ok: true }, getEvents_(p.from, p.to)));
      case 'load':     return json_({ ok: true, data: loadBig_('DATA') });
      case 'save':
        saveData_(p.data);
        return json_({ ok: true, savedAt: Date.now() });
      case 'pushsub':  return json_(Object.assign({ ok: true }, savePushSub_(p)));
      case 'pushoff':  return json_(Object.assign({ ok: true }, removePushSub_(p.dev)));
      case 'pending':  return json_({ ok: true, items: pending_() });
      case 'testpush': return json_(Object.assign({ ok: true }, notify_([{ id: 'test|' + Date.now(), title: '✅ Schulplaner-Test', body: 'Benachrichtigungen funktionieren.' }])));
      default:         return json_({ ok: false, error: 'Unbekannte Aktion' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============ Kalender (nur lesen) ============ */

function getEvents_(fromStr, toStr) {
  const from = parseDay_(fromStr) || new Date();
  let to = parseDay_(toStr) || new Date(from.getTime() + 60 * 864e5);
  if ((to - from) / 864e5 > MAX_RANGE_DAYS) to = new Date(from.getTime() + MAX_RANGE_DAYS * 864e5);

  const calendars = [];
  const events = [];
  CalendarApp.getAllCalendars().forEach(c => {
    const id = c.getId();
    calendars.push({ id: id, name: c.getName(), color: c.getColor(), own: c.isOwnedByMe() });
    let list = [];
    try { list = c.getEvents(from, to); } catch (err) { return; }
    list.forEach(ev => {
      const allDay = ev.isAllDayEvent();
      const s = allDay ? ev.getAllDayStartDate() : ev.getStartTime();
      const en = allDay ? ev.getAllDayEndDate() : ev.getEndTime();
      events.push({
        id: ev.getId(), cal: id, title: ev.getTitle() || '(ohne Titel)',
        start: fmt_(s, allDay), end: fmt_(en, allDay), allDay: allDay, loc: ev.getLocation() || ''
      });
    });
  });
  return { from: fmt_(from, true), to: fmt_(to, true), calendars: calendars, events: events };
}

function parseDay_(s) {
  if (!s || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  const a = s.split('-').map(Number);
  return new Date(a[0], a[1] - 1, a[2]);
}
function fmt_(d, dayOnly) { return Utilities.formatDate(d, TZ, dayOnly ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm"); }
function nowLocal_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm"); }

/* ============ Speicher (große Werte in Blöcken) ============ */

function saveBig_(key, value) {
  const s = JSON.stringify(value);
  const props = PropertiesService.getScriptProperties();
  const oldN = Number(props.getProperty(key + '_N')) || 0;
  const n = Math.ceil(s.length / CHUNK);
  const obj = {}; obj[key + '_N'] = String(n);
  for (let i = 0; i < n; i++) obj[key + '_' + i] = s.substr(i * CHUNK, CHUNK);
  props.setProperties(obj, false);
  for (let i = n; i < oldN; i++) props.deleteProperty(key + '_' + i);
}
function loadBig_(key) {
  const props = PropertiesService.getScriptProperties().getProperties();
  const n = Number(props[key + '_N']) || 0;
  if (!n) return null;
  let s = '';
  for (let i = 0; i < n; i++) s += props[key + '_' + i] || '';
  try { return JSON.parse(s); } catch (err) { return null; }
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function saveData_(data) {
  if (!data || typeof data !== 'object') throw new Error('Keine Daten');
  if (JSON.stringify(data).length > 400000) throw new Error('Daten zu groß');
  withLock_(() => saveBig_('DATA', data));
}

/* ============ Benachrichtigungen (Web Push) ============
 * Die App schickt beim Aktivieren ihre Push-Adresse und vorab signierte Schlüssel (gültig ca. 30 Tage,
 * werden bei jedem Öffnen der App erneuert). Das Skript stupst damit das Handy an, die App holt sich
 * dann den Text der Erinnerung über "pending" ab und zeigt ihn als normale Android-Benachrichtigung. */

function savePushSub_(p) {
  if (!p.dev || !p.sub || !p.sub.endpoint || !p.pub || !p.jwts) throw new Error('Unvollständige Push-Daten');
  return withLock_(() => {
    const subs = loadBig_('PUSH') || {};
    subs[p.dev] = { endpoint: p.sub.endpoint, pub: p.pub, jwts: p.jwts, name: p.name || '', ts: Date.now() };
    saveBig_('PUSH', subs);
    return { devices: Object.keys(subs).length };
  });
}
function removePushSub_(dev) {
  return withLock_(() => {
    const subs = loadBig_('PUSH') || {};
    delete subs[dev]; saveBig_('PUSH', subs);
    return { devices: Object.keys(subs).length };
  });
}
function pending_() {
  const list = loadBig_('PENDING') || [];
  const limit = Date.now() - 24 * 3600e3;
  return list.filter(x => x.ts > limit);
}
/** Benachrichtigungen vormerken und alle Geräte anstupsen. */
function notify_(items) {
  if (!items.length) return { sent: 0 };
  const limit = Date.now() - 24 * 3600e3;
  const list = (loadBig_('PENDING') || []).filter(x => x.ts > limit);
  items.forEach(it => list.push(Object.assign({ ts: Date.now() }, it)));
  saveBig_('PENDING', list.slice(-40));
  return pushAll_();
}
function pushAll_() {
  const subs = loadBig_('PUSH') || {};
  const now = Math.floor(Date.now() / 1000);
  const result = { devices: 0, sent: 0, codes: [] };
  let changed = false;
  Object.keys(subs).forEach(dev => {
    result.devices++;
    const s = subs[dev];
    const tok = (s.jwts || []).filter(j => j.exp > now + 600 && j.exp < now + 86000).sort((a, b) => a.exp - b.exp)[0];
    if (!tok) { result.codes.push('kein gültiger Schlüssel – App öffnen'); return; }
    let code;
    try {
      const r = UrlFetchApp.fetch(s.endpoint, {
        method: 'post', muteHttpExceptions: true,
        headers: { Authorization: 'vapid t=' + tok.t + ', k=' + s.pub, TTL: '86400', Urgency: 'high' }
      });
      code = r.getResponseCode();
      if (code === 404 || code === 410) { delete subs[dev]; changed = true; }
      if (code >= 300) code += ' ' + r.getContentText().slice(0, 120);
    } catch (err) { code = String(err.message || err); }
    result.codes.push(code);
    if (String(code).charAt(0) === '2') result.sent++;
  });
  if (changed) saveBig_('PUSH', subs);
  return result;
}

/* ============ Erinnerungen (läuft automatisch alle 5 Min.) ============ */

function checkReminders() {
  const data = loadBig_('DATA');
  if (!data) return;
  const props = PropertiesService.getScriptProperties();
  const now = nowLocal_();
  const todayStr = now.slice(0, 10);
  let sent = {};
  try { sent = JSON.parse(props.getProperty('SENT') || '{}'); } catch (err) {}
  const subjects = {};
  (data.subjects || []).forEach(s => { subjects[s.id] = s.name; });
  const subj = id => subjects[id] ? subjects[id] + ': ' : '';
  const out = [];

  // 1) Einzelne Erinnerungen
  (data.items || []).forEach(it => {
    if (!it.remindAt || it.done) return;
    const key = it.id + '|' + it.remindAt;
    if (sent[key] || it.remindAt > now) return;
    if (it.remindAt < shiftDay_(todayStr, -1)) { sent[key] = todayStr; return; }
    const kind = it.type === 'event' ? '📌 ' : '📚 ';
    out.push({ id: key, title: kind + subj(it.subject) + it.title,
      body: (it.type === 'event' ? '' : 'Fällig ') + dayLabel_(it.due, todayStr) + (it.time ? ', ' + it.time + ' Uhr' : '') +
        (it.notes ? ' · ' + String(it.notes).split('\n')[0].slice(0, 80) : '') });
    sent[key] = todayStr;
  });

  // 2) Abend-Übersicht für morgen (zur eingestellten Stunde)
  const hour = data.settings && data.settings.digestHour;
  if (hour !== null && hour !== undefined && hour !== '' &&
      Number(now.slice(11, 13)) === Number(hour) && sent['digest'] !== todayStr) {
    const tomorrow = shiftDay_(todayStr, 1);
    const due = (data.items || []).filter(it => !it.done && it.type !== 'event' && it.due && it.due <= tomorrow);
    const exams = (data.exams || []).filter(x => x.date > todayStr && x.date <= shiftDay_(todayStr, 3));
    if (due.length || exams.length) {
      const lines = due.map(it => subj(it.subject) + it.title + (it.due < todayStr ? ' (überfällig)' : ''))
        .concat(exams.map(x => '📖 ' + x.title + ' ' + dayLabel_(x.date, todayStr)));
      out.push({ id: 'digest|' + todayStr,
        title: due.length ? '🗓️ Bis morgen: ' + due.length + ' HÜ offen' : '📖 Prüfung bald',
        body: lines.slice(0, 6).join(' · ') });
    }
    sent['digest'] = todayStr;
  }

  const limit = shiftDay_(todayStr, -14);
  Object.keys(sent).forEach(k => { if (sent[k] < limit) delete sent[k]; });
  props.setProperty('SENT', JSON.stringify(sent));
  if (out.length) withLock_(() => notify_(out));
}

function shiftDay_(dayStr, n) {
  const d = parseDay_(dayStr);
  d.setDate(d.getDate() + n);
  return fmt_(d, true);
}

function dayLabel_(dayStr, todayStr) {
  if (!dayStr) return '';
  if (dayStr === todayStr) return 'heute';
  if (dayStr === shiftDay_(todayStr, 1)) return 'morgen';
  const d = parseDay_(dayStr);
  const wd = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'][d.getDay()];
  return 'am ' + wd + ', ' + d.getDate() + '.' + (d.getMonth() + 1) + '.';
}
