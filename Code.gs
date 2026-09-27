/**
 * techo(手帳)v2.2 公開版(母艦=個人アカウント / 学校カレンダーはICSで読むだけ)
 * https://github.com/toitoitoi-lab/techo  (MIT License)
 *
 * 資産(コード・タスク・大石・ログ)は全部個人アカウント側。
 * 学校の予定は「限定公開ICS URL」を読み取り専用で表示するだけなので、
 * 異動・退職してもURLを差し替えれば手帳はそのまま使い続けられる。
 *
 * 初期設定(スクリプトプロパティ / 名前は正確にこの通り):
 * 1. GEMINI_API_KEY … GeminiのAPIキー
 * 2. ICS_URLS       … 外部カレンダーの一覧。1行に「ラベル|URL」、複数行可。例:
 *                       学校|https://calendar.google.com/calendar/ical/...basic.ics
 *                       部活|https://calendar.google.com/calendar/ical/...basic.ics
 *                     (学校カレンダー設定 >「カレンダーの統合」>
 *                      「iCal形式の限定公開URL」をコピー)
 * 3. PLACES_API_KEY … 会場検索用(任意)。Google Cloud Consoleで
 *                     Places API を有効化し、課金設定済みのキーを発行。
 *                     未設定でも他機能は動く(会場検索だけ使えない)。
 * 4. デプロイ > ウェブアプリ(自分として実行 / 自分のみアクセス)
 *    ※ 自分の予定・タスクが入るアプリなので、アクセスは必ず「自分のみ」にしてください。
 */

const PROP_SS_ID = 'PLANNER_SS_ID';
const SS_NAME = '手帳データ';

// ---------- エントリポイント ----------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('techo 手帳')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- スプレッドシート(DB) ----------

function getSS_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SS_ID);
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(SS_NAME);
    props.setProperty(PROP_SS_ID, ss.getId());
  }
  initSheets_(ss); // 既存シートは変更しない。無いシートだけ追加(既存データは消えない)
  return ss;
}

function initSheets_(ss) {
  const defs = {
    '役割': ['id', '役割名', '表示順'],
    '大石': ['id', '週キー', '役割id', '内容', '状態'],
    'タスク': ['id', '日付', '内容', '優先', '状態', '作成日時'],
    'ログ': ['日時', '種別', '内容'],
    '学校予定キュー': ['id', '日付', '開始', '終了', 'タイトル', '会場', '状態', '作成日時', '詳細'],
    '個人予定キュー': ['id', '日付', '開始', '終了', 'タイトル', '会場', '状態', '作成日時', '詳細'],
    '目標': ['種別', '内容', '更新日時'],
    '目標項目': ['id', '種別', '優先度', '内容', '締切', '状態', '順序', '作成日時'],
    '手書きメモ': ['日付', 'タイトル', 'DriveファイルID', 'URL', '更新日時']
  };
  Object.keys(defs).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(defs[name]);
    }
  });
  const first = ss.getSheets()[0];
  if (first.getName() === 'シート1' || first.getName() === 'Sheet1') ss.deleteSheet(first);

  // 役割の初期値(あとで画面から編集可)
  const roleSh = ss.getSheetByName('役割');
  if (roleSh.getLastRow() === 1) {
    [['r1', '授業', 1],
     ['r2', '校務分掌', 2],
     ['r3', '研修・学び', 3],
     ['r4', '家庭・自分', 4]].forEach(function (r) { roleSh.appendRow(r); });
  }

  const goalSh = ss.getSheetByName('目標');
  if (goalSh.getLastRow() === 1) {
    goalSh.appendRow(['使命', '', '']);
    goalSh.appendRow(['当面', '', '']);
    goalSh.appendRow(['長期', '', '']);
  }
}

function getGoals() {
  const rows = rows_('目標');
  const out = { 使命: '', 当面: '', 長期: '' };
  rows.forEach(function (r) { if (out.hasOwnProperty(r[0])) out[r[0]] = r[1] || ''; });
  return out;
}

function saveGoal(type, text) {
  const sh = sheet_('目標');
  const rows = rows_('目標');
  let found = false;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === type) {
      sh.getRange(i + 2, 2, 1, 2).setValues([[text, new Date()]]);
      found = true;
      break;
    }
  }
  if (!found) sh.appendRow([type, text, new Date()]);
  return getGoals();
}

// ---------- 手書きメモ(1日1ページ、Google Driveに保存) ----------

function getMemoFolder_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('MEMO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 作り直す */ }
  }
  const folder = DriveApp.createFolder('手帳_手書きメモ');
  props.setProperty('MEMO_FOLDER_ID', folder.getId());
  return folder;
}

function getHandwriteMemo(dateStr) {
  const rows = rows_('手書きメモ');
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === dateStr) {
      // 画像は公開リンクにせず、本人の権限で読み出して画面に渡す
      let dataUrl = '';
      if (rows[i][2]) {
        try {
          const blob = DriveApp.getFileById(rows[i][2]).getBlob();
          dataUrl = 'data:image/png;base64,' + Utilities.base64Encode(blob.getBytes());
        } catch (e) { /* 読めなければ空 */ }
      }
      return { date: rows[i][0], title: rows[i][1], url: dataUrl };
    }
  }
  return null;
}

function saveHandwriteMemo(dateStr, title, dataUrl) {
  const base64 = dataUrl.split(',')[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', 'memo_' + dateStr + '.png');
  const folder = getMemoFolder_();

  const rowsData = rows_('手書きメモ');
  let existingFileId = null, rowIndex = -1;
  for (let i = 0; i < rowsData.length; i++) {
    if (rowsData[i][0] === dateStr) { existingFileId = rowsData[i][2]; rowIndex = i; break; }
  }
  if (existingFileId) {
    try { DriveApp.getFileById(existingFileId).setTrashed(true); } catch (e) {}
  }
  const file = folder.createFile(blob);
  // 共有設定は変えない(自分だけが見られる状態のまま保存する)
  const url = file.getUrl();

  const sh = sheet_('手書きメモ');
  if (rowIndex >= 0) {
    sh.getRange(rowIndex + 2, 2, 1, 3).setValues([[title, file.getId(), url]]);
    sh.getRange(rowIndex + 2, 5).setValue(new Date());
  } else {
    sh.appendRow([dateStr, title, file.getId(), url, new Date()]);
  }
  return { date: dateStr, title: title, url: dataUrl };
}

function searchHandwriteMemos(query) {
  const rows = rows_('手書きメモ');
  const q = query.toLowerCase();
  return rows.filter(function (r) { return (r[1] || '').toLowerCase().indexOf(q) >= 0; })
    .map(function (r) { return { date: r[0], title: r[1] }; })
    .sort(function (a, b) { return b.date.localeCompare(a.date); });
}

// ---------- 月間見通し ----------

function getMonthOverview(dateStr) {
  const tz = Session.getScriptTimeZone();
  const base = new Date(dateStr + 'T00:00:00');
  const monthStart = new Date(base.getFullYear(), base.getMonth(), 1);
  const monthEnd = new Date(base.getFullYear(), base.getMonth() + 1, 1);
  const monthKey = Utilities.formatDate(monthStart, tz, 'yyyy-MM');

  const icsMonth = getIcsEvents_(monthStart, monthEnd);
  const calCounts = {};
  CalendarApp.getAllCalendars().forEach(function (cal) {
    try {
      cal.getEvents(monthStart, monthEnd).forEach(function (ev) {
        const d0 = new Date(ev.getStartTime()); d0.setHours(0,0,0,0);
        const key = Utilities.formatDate(d0, tz, 'yyyy-MM-dd');
        calCounts[key] = (calCounts[key] || 0) + 1;
      });
    } catch (e) {}
  });
  icsMonth.forEach(function (ev) {
    const d0 = new Date(ev.start); d0.setHours(0,0,0,0);
    const key = Utilities.formatDate(d0, tz, 'yyyy-MM-dd');
    calCounts[key] = (calCounts[key] || 0) + 1;
  });

  const days = [];
  const startWeekday = (monthStart.getDay() + 6) % 7; // 月曜=0
  for (let i = 0; i < startWeekday; i++) days.push(null);
  for (let d = new Date(monthStart); d < monthEnd; d.setDate(d.getDate() + 1)) {
    const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    days.push({ date: key, day: d.getDate(), count: calCounts[key] || 0 });
  }
  return { monthKey: monthKey, days: days };
}

// ---------- 目標項目(音声でまとめて話す → 長期/当面 × A/B/C に分類) ----------

function organizeGoals(text) {
  const prompt =
    'あなたは手帳の秘書です。以下は「いつまでに何をしたい」という願望・目標を' +
    'まとめて話した音声メモです。項目ごとに分解してください。\n' +
    '種別(type)は、数ヶ月〜数年がかりのものは"長期"、数週間〜今学期中で片付くものは"当面"。\n' +
    '優先度(priority)は、話し方の熱量・緊急度から A(最重要)/B(重要)/C(余裕があれば)を推測。\n' +
    '締切(deadline)は、話された内容から具体的な日付を推測してyyyy-MM-dd形式で出す' +
    '(例: "2027年3月まで"→"2027-03-31"、"今学期中"→学期末頃の妥当な日付)。' +
    '推測できなければ空文字。基準日は本日: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '。\n' +
    '出力はJSONのみ、前置きや```は禁止。形式:\n' +
    '{"items":[{"type":"長期|当面","priority":"A|B|C","text":"...","deadline":"..."}]}\n\n' +
    'メモ:\n' + text;

  const parsed = callGemini_(prompt);
  sheet_('ログ').appendRow([new Date(), '目標分類', text]);
  return parsed.items || [];
}

function addGoalItems(items) {
  const sh = sheet_('目標項目');
  const existing = rows_('目標項目');
  let maxOrder = 0;
  existing.forEach(function (r) { if (r[6] > maxOrder) maxOrder = r[6]; });
  items.forEach(function (it) {
    maxOrder++;
    sh.appendRow([uid_(), it.type || '当面', it.priority || 'B', it.text || '', it.deadline || '',
      '未', maxOrder, new Date()]);
  });
  return getGoalItems();
}

function getGoalItems() {
  const rows = rows_('目標項目');
  const order = { A: 0, B: 1, C: 2 };
  return rows.map(function (r) {
    return { id: r[0], type: r[1], priority: r[2], text: r[3], deadline: r[4], status: r[5], sortOrder: r[6] };
  }).sort(function (a, b) {
    if (a.type !== b.type) return a.type === '長期' ? -1 : 1;
    if (a.priority !== b.priority) return order[a.priority] - order[b.priority];
    return a.sortOrder - b.sortOrder;
  });
}

function setGoalItemMeta(id, type, priority) {
  const sh = sheet_('目標項目');
  const rows = rows_('目標項目');
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === id) { sh.getRange(i + 2, 2, 1, 2).setValues([[type, priority]]); break; }
  }
  return getGoalItems();
}

function setGoalItemStatus(id, status) {
  const sh = sheet_('目標項目');
  const rows = rows_('目標項目');
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === id) { sh.getRange(i + 2, 6).setValue(status); break; }
  }
  return getGoalItems();
}

function sheet_(name) { return getSS_().getSheetByName(name); }

function rows_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function uid_() { return Utilities.getUuid().slice(0, 8); }

// ---------- 今日ビュー ----------

function getToday(dateStr) {
  const tz = Session.getScriptTimeZone();
  const date = dateStr ? new Date(dateStr + 'T00:00:00') : new Date();
  const dayStart = new Date(date); dayStart.setHours(0, 0, 0, 0);
  const dayEnd = new Date(date); dayEnd.setHours(23, 59, 59, 999);
  const key = Utilities.formatDate(dayStart, tz, 'yyyy-MM-dd');

  // 予定: このアカウントに見えている全カレンダー + 個人ICS
  const events = [];
  CalendarApp.getAllCalendars().forEach(function (cal) {
    try {
      cal.getEvents(dayStart, dayEnd).forEach(function (ev) {
        events.push({
          title: ev.getTitle(),
          start: ev.isAllDayEvent() ? '終日' : Utilities.formatDate(ev.getStartTime(), tz, 'HH:mm'),
          end: ev.isAllDayEvent() ? '' : Utilities.formatDate(ev.getEndTime(), tz, 'HH:mm'),
          cal: cal.getName(),
          allDay: ev.isAllDayEvent()
        });
      });
    } catch (e) { /* 読めないカレンダーは無視 */ }
  });
  getIcsEvents_(dayStart, dayEnd).forEach(function (ev) {
    events.push({
      title: ev.title,
      start: ev.allDay ? '終日' : Utilities.formatDate(ev.start, tz, 'HH:mm'),
      end: ev.allDay ? '' : Utilities.formatDate(ev.end, tz, 'HH:mm'),
      cal: ev.cal,
      allDay: ev.allDay
    });
  });
  events.sort(function (a, b) {
    if (a.allDay !== b.allDay) return a.allDay ? -1 : 1;
    return a.start.localeCompare(b.start);
  });

  // タスク: その日 + 未完の持ち越し
  const order = { 'A': 0, 'B': 1, 'C': 2 };
  const tasks = rows_('タスク')
    .filter(function (r) {
      const d = Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM-dd');
      return d === key || (r[4] !== '完了' && d < key);
    })
    .map(function (r) {
      return {
        id: r[0],
        date: Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM-dd'),
        text: r[2], priority: r[3] || 'B', status: r[4] || '未',
        carried: Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM-dd') < key
      };
    });
  tasks.sort(function (a, b) {
    if ((a.status === '完了') !== (b.status === '完了')) return a.status === '完了' ? 1 : -1;
    return (order[a.priority] || 1) - (order[b.priority] || 1);
  });

  return { date: key, weekday: Utilities.formatDate(dayStart, tz, 'E'), events: events, tasks: tasks };
}

function addTask(dateStr, text, priority) {
  if (!text) return getToday(dateStr);
  sheet_('タスク').appendRow([uid_(), dateStr, text, priority || 'B', '未', new Date()]);
  return getToday(dateStr);
}

function setTaskStatus(id, status, dateStr) {
  const sh = sheet_('タスク');
  const data = rows_('タスク');
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === id) { sh.getRange(i + 2, 5).setValue(status); break; }
  }
  return getToday(dateStr);
}

// ---------- 週間コンパス ----------

function weekStart_(dateStr) {
  const d = dateStr ? new Date(dateStr + 'T00:00:00') : new Date();
  const day = d.getDay(); // 0=日
  const diff = (day === 0 ? -6 : 1 - day); // 月曜始まり
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getWeek(dateStr) {
  const tz = Session.getScriptTimeZone();
  const start = weekStart_(dateStr);
  const key = Utilities.formatDate(start, tz, 'yyyy-MM-dd');

  const roles = rows_('役割')
    .map(function (r) { return { id: r[0], name: r[1], order: r[2] }; })
    .sort(function (a, b) { return a.order - b.order; });

  const rocks = rows_('大石')
    .filter(function (r) { return r[1] === key; })
    .map(function (r) { return { id: r[0], roleId: r[2], text: r[3], status: r[4] || '未' }; });

  // 各曜日の予定件数と先頭2件(俯瞰用)。ICSは週分を一度だけ取得
  const weekEnd = new Date(start); weekEnd.setDate(start.getDate() + 7);
  const icsWeek = getIcsEvents_(start, weekEnd);
  const days = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(start); d.setDate(start.getDate() + i);
    const dEnd = new Date(d); dEnd.setHours(23, 59, 59, 999);
    let evs = [];
    CalendarApp.getAllCalendars().forEach(function (cal) {
      try {
        cal.getEvents(d, dEnd).forEach(function (ev) {
          evs.push({
            t: ev.isAllDayEvent() ? '終日' : Utilities.formatDate(ev.getStartTime(), tz, 'HH:mm'),
            title: ev.getTitle()
          });
        });
      } catch (e) {}
    });
    icsWeek.forEach(function (ev) {
      if (ev.start < dEnd && ev.end > d) {
        evs.push({
          t: ev.allDay ? '終日' : Utilities.formatDate(ev.start, tz, 'HH:mm'),
          title: ev.title
        });
      }
    });
    evs.sort(function (a, b) { return a.t.localeCompare(b.t); });
    days.push({
      date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
      label: Utilities.formatDate(d, tz, 'M/d(E)'),
      count: evs.length,
      items: evs
    });
  }

  return { weekKey: key, roles: roles, rocks: rocks, days: days };
}

function addBigRock(weekKey, roleId, text) {
  if (text) sheet_('大石').appendRow([uid_(), weekKey, roleId, text, '未']);
  return getWeek(weekKey);
}

function setBigRockStatus(id, status, weekKey) {
  const sh = sheet_('大石');
  const data = rows_('大石');
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === id) { sh.getRange(i + 2, 5).setValue(status); break; }
  }
  return getWeek(weekKey);
}

function addRole(name) {
  const roles = rows_('役割');
  sheet_('役割').appendRow([uid_(), name, roles.length + 1]);
  return getWeek(null);
}

// ---------- 外部カレンダー(ICS直接取得) ----------
// 限定公開ICS URLをUrlFetchAppで読む。共有・招待メール不要。
// 対応: 単発予定 / 繰り返し(DAILY・WEEKLY・簡易MONTHLY/YEARLY) / EXDATE
// キャッシュ10分。

function getIcsEvents_(rangeStart, rangeEnd) {
  const raw = PropertiesService.getScriptProperties().getProperty('ICS_URLS');
  if (!raw) return [];
  const cache = CacheService.getScriptCache();
  const all = [];
  raw.split(/[\n,]+/).forEach(function (entry, i) {
    entry = entry.trim();
    if (!entry) return;
    const parts = entry.split('|');
    const label = parts.length > 1 ? parts[0].trim() : '外部' + (i + 1);
    const url = (parts.length > 1 ? parts[1] : parts[0]).trim();
    let text = cache.get('ICS_' + i);
    if (!text) {
      try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) return;
        text = res.getContentText();
        if (text.length < 90000) cache.put('ICS_' + i, text, 600);
      } catch (e) { return; }
    }
    parseIcs_(text, rangeStart, rangeEnd).forEach(function (ev) {
      ev.cal = label;
      all.push(ev);
    });
  });
  all.sort(function (a, b) { return a.start - b.start; });
  return all;
}

function parseIcs_(text, rangeStart, rangeEnd) {
  // 折返し行(先頭が空白/タブ)を結合
  const lines = text.split(/\r?\n/);
  const unfolded = [];
  lines.forEach(function (ln) {
    if (/^[ \t]/.test(ln) && unfolded.length) unfolded[unfolded.length - 1] += ln.slice(1);
    else unfolded.push(ln);
  });

  const events = [];
  let cur = null;
  unfolded.forEach(function (ln) {
    if (ln === 'BEGIN:VEVENT') { cur = { exdates: [] }; return; }
    if (ln === 'END:VEVENT') {
      if (cur && cur.start) expandEvent_(cur, rangeStart, rangeEnd, events);
      cur = null; return;
    }
    if (!cur) return;
    const idx = ln.indexOf(':');
    if (idx < 0) return;
    const keyPart = ln.slice(0, idx);
    const val = ln.slice(idx + 1);
    const key = keyPart.split(';')[0];
    if (key === 'SUMMARY') cur.title = val.replace(/\\,/g, ',').replace(/\\n/g, ' ');
    else if (key === 'DTSTART') { cur.start = parseIcsDate_(val); cur.allDay = val.length === 8; }
    else if (key === 'DTEND') cur.end = parseIcsDate_(val);
    else if (key === 'RRULE') cur.rrule = val;
    else if (key === 'EXDATE') {
      val.split(',').forEach(function (v) { cur.exdates.push(parseIcsDate_(v).getTime()); });
    }
  });
  events.sort(function (a, b) { return a.start - b.start; });
  return events;
}

function parseIcsDate_(v) {
  // 20260719 / 20260719T083000 / 20260719T083000Z
  const y = +v.slice(0, 4), mo = +v.slice(4, 6) - 1, d = +v.slice(6, 8);
  if (v.length === 8) return new Date(y, mo, d);
  const h = +v.slice(9, 11), mi = +v.slice(11, 13), s = +v.slice(13, 15) || 0;
  if (v.slice(-1) === 'Z') return new Date(Date.UTC(y, mo, d, h, mi, s));
  return new Date(y, mo, d, h, mi, s); // TZID付きはJST前提
}

function expandEvent_(ev, rangeStart, rangeEnd, out) {
  const durMs = (ev.end ? ev.end - ev.start : (ev.allDay ? 86400000 : 3600000));
  function push(startDate) {
    if (ev.exdates.indexOf(startDate.getTime()) >= 0) return;
    const endDate = new Date(startDate.getTime() + durMs);
    if (startDate < rangeEnd && endDate > rangeStart) {
      out.push({ title: ev.title || '(無題)', start: startDate, end: endDate, allDay: !!ev.allDay });
    }
  }
  if (!ev.rrule) { push(ev.start); return; }

  const rule = {};
  ev.rrule.split(';').forEach(function (p) {
    const kv = p.split('='); rule[kv[0]] = kv[1];
  });
  const freq = rule.FREQ;
  const interval = +(rule.INTERVAL || 1);
  const until = rule.UNTIL ? parseIcsDate_(rule.UNTIL) : null;
  let count = rule.COUNT ? +rule.COUNT : null;
  const byday = rule.BYDAY ? rule.BYDAY.split(',') : null;
  const dayMap = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
  const hardStop = until && until < rangeEnd ? until : rangeEnd;

  let n = 0;
  if (freq === 'DAILY') {
    for (let d = new Date(ev.start); d <= hardStop && n < 1000; d.setDate(d.getDate() + interval)) {
      if (count !== null && ++n > count) break;
      push(new Date(d));
    }
  } else if (freq === 'WEEKLY') {
    const days = byday ? byday.map(function (b) { return dayMap[b.slice(-2)]; }) : [ev.start.getDay()];
    // 週の起点(日曜)から interval 週ごとに走査
    const anchor = new Date(ev.start); anchor.setHours(0, 0, 0, 0);
    anchor.setDate(anchor.getDate() - anchor.getDay());
    for (let w = new Date(anchor); w <= hardStop && n < 1000; w.setDate(w.getDate() + 7 * interval)) {
      for (let i = 0; i < days.length; i++) {
        const d = new Date(w); d.setDate(w.getDate() + days[i]);
        d.setHours(ev.start.getHours(), ev.start.getMinutes(), 0, 0);
        if (d < ev.start || d > hardStop) continue;
        if (count !== null && ++n > count) return;
        push(d);
      }
    }
  } else if (freq === 'MONTHLY') {
    for (let d = new Date(ev.start); d <= hardStop && n < 240; d.setMonth(d.getMonth() + interval)) {
      if (count !== null && ++n > count) break;
      push(new Date(d));
    }
  } else if (freq === 'YEARLY') {
    for (let d = new Date(ev.start); d <= hardStop && n < 50; d.setFullYear(d.getFullYear() + interval)) {
      if (count !== null && ++n > count) break;
      push(new Date(d));
    }
  } else {
    push(ev.start); // 未対応FREQは初回のみ表示
  }
}

// ---------- 会場検索(Google Places API) ----------
// スクリプトプロパティ PLACES_API_KEY が必要(要 Google Cloud 課金設定)。
// 電話番号での検索には非対応(Places APIの仕様上不可)。施設名・住所のみ。

function searchPlaces(query) {
  if (!query || query.trim().length < 2) return [];
  const key = PropertiesService.getScriptProperties().getProperty('PLACES_API_KEY');
  if (!key) throw new Error('スクリプトプロパティに PLACES_API_KEY がありません。');
  const url = 'https://maps.googleapis.com/maps/api/place/textsearch/json'
    + '?query=' + encodeURIComponent(query)
    + '&language=ja&region=jp&key=' + key;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Places APIエラー: ' + res.getResponseCode());
  const body = JSON.parse(res.getContentText());
  if (body.status === 'REQUEST_DENIED') throw new Error('Places APIが有効化されていないか、課金設定未完了です。');
  return (body.results || []).slice(0, 6).map(function (r) {
    return { name: r.name, address: r.formatted_address, lat: r.geometry.location.lat, lng: r.geometry.location.lng };
  });
}

// ---------- 予定登録(学校/個人カレンダー) ----------

function addEvent(calType, dateStr, startTime, endTime, title, location, detail, link) {
  if (!title) throw new Error('タイトルが空です。');
  const start = new Date(dateStr + 'T' + (startTime || '09:00') + ':00');
  const end = new Date(dateStr + 'T' + (endTime || startTime || '10:00') + ':00');
  if (end <= start) end.setTime(start.getTime() + 3600000);

  const descParts = [];
  if (detail) descParts.push(detail);
  if (link) descParts.push('リンク: ' + link);
  const description = descParts.join('\n\n');

  ensureMemoStub_(dateStr, title);

  if (calType === '個人') {
    CalendarApp.getDefaultCalendar().createEvent(title, start, end, { location: location || '', description: description });
    sheet_('個人予定キュー').appendRow([uid_(), dateStr, startTime || '09:00', endTime || startTime || '10:00',
      title, location || '', '登録済', new Date(), description]);
    sheet_('ログ').appendRow([new Date(), '予定登録(個人)', title + (location ? ' @ ' + location : '')]);
    return { ok: true, mode: '個人', message: '個人カレンダーに即時登録しました。' };
  } else {
    sheet_('学校予定キュー').appendRow([uid_(), dateStr, startTime || '09:00', endTime || startTime || '10:00',
      title, location || '', '未処理', new Date(), description]);
    return { ok: true, mode: '学校', message: '学校カレンダーへの登録を予約しました(橋渡しスクリプトが反映、数分〜十数分かかります)。' };
  }
}

// 予定登録時、その日の手書きメモが未作成ならタイトルだけ入れた空枠を用意する
function ensureMemoStub_(dateStr, title) {
  try {
    const rows = rows_('手書きメモ');
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === dateStr) return; // 既にその日のメモがあれば何もしない
    }
    sheet_('手書きメモ').appendRow([dateStr, title, '', '', new Date()]);
  } catch (e) { /* メモ枠作成の失敗で予定登録自体は止めない */ }
}

function addEventsBulk(items) {
  let personal = 0, school = 0;
  const errors = [];
  items.forEach(function (it) {
    try {
      addEvent(it.calType, it.date, it.start, it.end, it.title, it.location || '', it.detail || '', '');
      if (it.calType === '個人') personal++; else school++;
    } catch (e) { errors.push((it.title || '(無題)') + ': ' + e.message); }
  });
  let message = '個人' + personal + '件・学校' + school + '件を登録しました。';
  if (errors.length) message += '\n失敗' + errors.length + '件:\n' + errors.join('\n');
  return { personal: personal, school: school, failed: errors.length, errors: errors, message: message };
}


// ---------- テストデータの一括リセット(本運用前用) ----------
// スプレッドシート側のデータのみ消去。Googleカレンダー本体は対象外
// (そちらは手動で削除する必要がある)。

function resetTestData() {
  const targets = ['タスク', '大石', '学校予定キュー', '個人予定キュー', '目標項目', 'ログ', '手書きメモ'];
  let cleared = 0;
  targets.forEach(function (name) {
    const sh = sheet_(name);
    if (!sh) return;
    const last = sh.getLastRow();
    if (last > 1) { sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent(); cleared++; }
  });
  return 'クリアしました: ' + targets.join(' / ') + '(該当データがあった ' + cleared + ' シート)';
}
// 2026年6月以降、Google AI StudioはAPIキーを新形式(AQ.Ab...)で発行するようになった。
// 新形式は ?key=... ではなく X-goog-api-key ヘッダーでの認証が必要。
// 旧形式(AIza...)との両対応にしておく。

function callGemini_(prompt) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('スクリプトプロパティに GEMINI_API_KEY がありません。名前はこの通りに(カスタム名は不可)。');

  // モデル名は固定版(gemini-2.5-flash等)ではなく "-latest" エイリアスを使う。
  // Googleは頻繁にモデルを廃止するため、固定名だと数ヶ月おきにコード修正が必要になる。
  const isNewKey = key.indexOf('AQ.') === 0;
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent'
    + (isNewKey ? '' : '?key=' + key);
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0.2 }
    }),
    muteHttpExceptions: true
  };
  if (isNewKey) options.headers = { 'x-goog-api-key': key };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code === 429) throw new Error('Gemini APIのレート制限(429)。1分ほど待って再実行してください。');
  if (code !== 200) throw new Error('Gemini APIエラー: ' + code + ' / ' + res.getContentText().slice(0, 200));

  const body = JSON.parse(res.getContentText());
  const raw = body.candidates[0].content.parts[0].text.replace(/```json|```/g, '').trim();
  return JSON.parse(raw);
}

// ---------- Gemini(音声メモ → 複数予定に分解) ----------

function organizeEvents(text, baseDateStr) {
  const prompt =
    'あなたは手帳の秘書です。以下は複数の予定をまとめて話した音声メモです。\n' +
    '予定ごとに分解してください。同じ日に複数の予定がある場合(例: 午前は◯◯、9時から◯◯、' +
    '10時から◯◯、13時から◯◯…)は、日付が同じでも必ず別々のイベントとして分解すること。' +
    '1日にまとめず、時間帯ごとに分けるのが原則。\n' +
    '基準日: ' + baseDateStr + '(この曜日から「今週の水曜」等を計算)\n' +
    '時刻が明言されない予定は start に最も妥当な時間帯を推測して入れる(終日ならallDay:trueにしてstart/endは省略)。\n' +
    '出力はJSONのみ、前置きや```は禁止。形式:\n' +
    '{"events":[{"date":"yyyy-MM-dd","start":"HH:mm","end":"HH:mm","title":"...","allDay":false,' +
    '"venue":"場所が話されていれば(なければ空文字)","detail":"補足があれば(なければ空文字)"}]}\n\n' +
    'メモ:\n' + text;

  const parsed = callGemini_(prompt);
  sheet_('ログ').appendRow([new Date(), '予定振り分け', text]);
  const events = parsed.events || [];
  return attachDuplicateHints_(events);
}

// 既存の予定と日付・タイトルが似ていれば警告フラグを付ける(重複登録の事故防止)
function attachDuplicateHints_(events) {
  const tz = Session.getScriptTimeZone();
  const byDate = {};
  events.forEach(function (ev) { byDate[ev.date] = true; });

  const existingByDate = {};
  Object.keys(byDate).forEach(function (dateStr) {
    const dayStart = new Date(dateStr + 'T00:00:00');
    const dayEnd = new Date(dateStr + 'T23:59:59');
    const list = [];
    CalendarApp.getAllCalendars().forEach(function (cal) {
      try {
        cal.getEvents(dayStart, dayEnd).forEach(function (e) {
          list.push({ title: e.getTitle(), time: e.isAllDayEvent() ? '終日' : Utilities.formatDate(e.getStartTime(), tz, 'HH:mm') });
        });
      } catch (e) {}
    });
    getIcsEvents_(dayStart, dayEnd).forEach(function (e) {
      list.push({ title: e.title, time: e.allDay ? '終日' : Utilities.formatDate(e.start, tz, 'HH:mm') });
    });
    existingByDate[dateStr] = list;
  });

  function norm_(s) { return (s || '').replace(/\s+/g, '').toLowerCase(); }

  return events.map(function (ev) {
    const candidates = existingByDate[ev.date] || [];
    const evN = norm_(ev.title);
    let match = null;
    for (let i = 0; i < candidates.length; i++) {
      const cN = norm_(candidates[i].title);
      if (evN.length >= 3 && (cN.indexOf(evN) >= 0 || evN.indexOf(cN) >= 0)) { match = candidates[i]; break; }
    }
    ev.dupWarning = match ? (match.title + '(' + match.time + ')と似ています') : null;
    return ev;
  });
}

// ---------- Gemini(音声メモ → タスク整理) ----------

function organize(text, dateStr) {
  const prompt =
    'あなたは手帳の秘書です。以下の音声メモを、実行可能なタスクに分解してください。\n' +
    '基準日: ' + dateStr + '\n' +
    '優先度は A=最重要(今日必ず) B=重要 C=余裕があれば。\n' +
    '出力はJSONのみ。形式: {"tasks":[{"date":"yyyy-MM-dd","text":"...","priority":"A|B|C"}]}\n' +
    '日付の指定がなければ基準日。前置きや```は一切不要。\n\n' +
    'メモ:\n' + text;

  const parsed = callGemini_(prompt);
  sheet_('ログ').appendRow([new Date(), '音声整理', text]);
  return parsed.tasks || [];
}

function acceptTasks(tasks, dateStr) {
  const sh = sheet_('タスク');
  tasks.forEach(function (t) {
    sh.appendRow([uid_(), t.date || dateStr, t.text, t.priority || 'B', '未', new Date()]);
  });
  return getToday(dateStr);
}