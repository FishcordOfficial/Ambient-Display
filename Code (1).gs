/**
 * Ambient Display (settings live in Config.gs, which updates never overwrite) – Google Apps Script web app
 * Google services used: Calendar, Gmail, Drive (photo slideshow), Maps (geocoding),
 * UrlFetch (Open-Meteo weather + air quality), Cache.
 */
const VERSION = '1.0.0';

function doGet(e) {
  if (CONFIG.KEY && (!e || e.parameter.key !== CONFIG.KEY)) {
    return ContentService.createTextOutput('Not authorised');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ambient Display')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** One call for everything the screen needs. opts = {c: calendar, m: mail, a: air quality} */
function getDashboard(lat, lon, opts) {
  opts = opts || {};
  const out = { v: VERSION, a: null, aq: null, ag: null, unread: null, t: Date.now() };
  out.lock = !!CONFIG.LOCK_ABOUT; out.owner = CONFIG.OWNER || 'Michael'; out.by = CONFIG.MADE_BY || 'Claude (Anthropic)';
  try { out.a = fetchWeather_(lat, lon); } catch (e) { console.error('weather', e); }
  if (opts.a) { try { out.aq = fetchAir_(lat, lon); } catch (e) { console.error('air', e); } }
  if (opts.c) { try { out.ag = getAgenda_(); } catch (e) { console.error('calendar', e); } }
  if (opts.m) { try { out.unread = GmailApp.getInboxUnreadCount(); } catch (e) { console.error('mail', e); } }
  return out;
}

function fetchWeather_(lat, lon) {
  const url = 'https://api.open-meteo.com/v1/forecast?latitude=' + lat + '&longitude=' + lon +
    '&current=temperature_2m,apparent_temperature,relative_humidity_2m,weather_code,is_day,wind_speed_10m' +
    '&daily=weather_code,temperature_2m_max,temperature_2m_min,sunrise,sunset&hourly=temperature_2m,weather_code&forecast_hours=7&timezone=auto&forecast_days=5';
  return cachedJson_('w2:' + lat + ',' + lon, url);
}

function fetchAir_(lat, lon) {
  const url = 'https://air-quality-api.open-meteo.com/v1/air-quality?latitude=' + lat +
    '&longitude=' + lon + '&current=european_aqi';
  const j = cachedJson_('aq:' + lat + ',' + lon, url);
  return j && j.current ? j.current.european_aqi : null;
}

function cachedJson_(key, url) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
  const txt = res.getContentText();
  cache.put(key, txt, CONFIG.WEATHER_CACHE_SEC);
  return JSON.parse(txt);
}

/** Upcoming events from now until the end of tomorrow. */
function getAgenda_() {
  const cal = CONFIG.CALENDAR_ID
    ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID)
    : CalendarApp.getDefaultCalendar();
  const now = new Date();
  const end = new Date(now); end.setDate(end.getDate() + 2); end.setHours(0, 0, 0, 0);
  return cal.getEvents(now, end)
    .filter(ev => ev.getEndTime() > now)
    .slice(0, CONFIG.EVENT_LIMIT)
    .map(ev => ({ t: ev.getTitle(), s: ev.getStartTime().toISOString(), all: ev.isAllDayEvent() }));
}

/** City name -> {city, lat, lon} using the Google Maps service. */
function geocode(q) {
  const r = Maps.newGeocoder().geocode(q);
  if (r.status !== 'OK' || !r.results.length) return null;
  const p = r.results[0], c = p.address_components && p.address_components[0];
  return {
    city: c ? c.short_name : p.formatted_address.split(',')[0],
    lat: p.geometry.location.lat,
    lon: p.geometry.location.lng
  };
}

/** Drive photo slideshow: list image ids, then fetch each as a data URL. */
function listPhotos() {
  if (!CONFIG.PHOTO_FOLDER_ID) return [];
  const ids = [], files = DriveApp.getFolderById(CONFIG.PHOTO_FOLDER_ID).getFiles();
  while (files.hasNext() && ids.length < 100) {
    const f = files.next();
    if (f.getMimeType().indexOf('image/') === 0) ids.push(f.getId());
  }
  return ids;
}

function getPhoto(id) {
  const res = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + id + '&sz=w1920', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return null;
  const b = res.getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

/** Quick views for the Apps launcher. */
function getWeek() {
  const cal = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
  const now = new Date(), end = new Date(now.getTime() + 7 * 864e5);
  return cal.getEvents(now, end).filter(ev => ev.getEndTime() > now).slice(0, 12)
    .map(ev => ({ t: ev.getTitle(), s: ev.getStartTime().toISOString(), all: ev.isAllDayEvent(), loc: ev.getLocation() }));
}

function getInbox() {
  return GmailApp.getInboxThreads(0, 8).map(t => {
    const m = t.getMessages()[t.getMessageCount() - 1], f = m.getFrom();
    return { s: t.getFirstMessageSubject() || '(no subject)', f: f.replace(/<.*>/, '').replace(/"/g, '').trim() || f,
             d: m.getDate().toISOString(), u: t.isUnread() };
  });
}

/* ---------- Updater ----------
 * Reads version.json, Code and Index from GitHub (GITHUB_REPO) or the Drive folder UPDATE_FOLDER_ID,
 * then rewrites the Code and Index files of this project through the Apps Script API.
 * Config.gs is never touched. */
function newer_(a, b) {
  const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number);
  for (let i = 0; i < 3; i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); }
  return false;
}

function readUpdateFile_(name) {
  if (CONFIG.GITHUB_REPO) {
    // GitHub: version.json, Code.gs and Index.html in the repo (optionally inside GITHUB_PATH)
    const file = { 'Code.txt': 'Code.gs', 'Index.txt': 'Index.html' }[name] || name;
    const url = 'https://raw.githubusercontent.com/' + CONFIG.GITHUB_REPO + '/' + (CONFIG.GITHUB_BRANCH || 'main') + '/' +
      (CONFIG.GITHUB_PATH ? CONFIG.GITHUB_PATH + '/' : '') + file + '?cb=' + Date.now();
    const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN'); // only for private repos
    const r = UrlFetchApp.fetch(url, { headers: token ? { Authorization: 'token ' + token } : {}, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('Could not read ' + file + ' from GitHub (HTTP ' + r.getResponseCode() + ').');
    return r.getContentText();
  }
  const it = DriveApp.getFolderById(CONFIG.UPDATE_FOLDER_ID).getFilesByName(name);
  if (!it.hasNext()) throw new Error('Missing ' + name + ' in the update folder.');
  return it.next().getBlob().getDataAsString();
}

function checkUpdate() {
  const out = { current: VERSION, available: false, configured: !!(CONFIG.GITHUB_REPO || CONFIG.UPDATE_FOLDER_ID) };
  if (!out.configured) return out;
  const m = JSON.parse(readUpdateFile_('version.json'));
  out.latest = m.version; out.title = m.title || ''; out.notes = m.notes || [];
  out.available = newer_(m.version, VERSION);
  return out;
}

function applyUpdate() {
  const m = JSON.parse(readUpdateFile_('version.json'));
  if (!newer_(m.version, VERSION)) return { ok: true, live: true, version: VERSION, msg: 'Already up to date.' };
  const code = readUpdateFile_('Code.txt'), html = readUpdateFile_('Index.txt');
  const id = ScriptApp.getScriptId(), base = 'https://script.googleapis.com/v1/projects/' + id;
  const call = (method, path, body) => {
    const r = UrlFetchApp.fetch(base + path, {
      method: method, contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: body ? JSON.stringify(body) : undefined
    });
    const c = r.getResponseCode();
    if (c >= 300) {
      throw new Error(c === 403
        ? 'Turn on the Apps Script API at script.google.com/home/usersettings, then try again.'
        : 'Update failed (HTTP ' + c + ').');
    }
    return JSON.parse(r.getContentText());
  };
  const files = call('get', '/content').files.filter(f => f.name !== 'Code' && f.name !== 'Index');
  files.push({ name: 'Code', type: 'SERVER_JS', source: code }, { name: 'Index', type: 'HTML', source: html });
  call('put', '/content', { files: files });
  const v = call('post', '/versions', { description: 'Ambient Display ' + m.version });
  if (!CONFIG.DEPLOYMENT_ID) {
    return { ok: true, live: false, version: m.version,
      msg: 'Saved as version ' + v.versionNumber + '. Add DEPLOYMENT_ID in Config.gs so updates go live automatically.' };
  }
  call('put', '/deployments/' + CONFIG.DEPLOYMENT_ID, { deploymentConfig: {
    scriptId: id, versionNumber: v.versionNumber, manifestFileName: 'appsscript',
    description: 'Ambient Display ' + m.version } });
  return { ok: true, live: true, version: m.version, msg: 'Restart the display to finish.' };
}

/** Run once from the editor to check the GitHub connection (see the Execution log). */
function testGithub() {
  ['version.json', 'Code.txt', 'Index.txt'].forEach(n => {
    try { Logger.log(n + ': OK, ' + readUpdateFile_(n).length + ' characters'); }
    catch (e) { Logger.log(n + ': ' + e.message); }
  });
  Logger.log(JSON.stringify(checkUpdate()));
}
