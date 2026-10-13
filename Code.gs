/**
 * 消防指揮情境訓練｜AAR 接收器＋處置作為分析報表
 *
 * 安裝：在 Google 試算表中「擴充功能 → Apps Script」，以本檔全部取代原本內容並儲存，
 *      再到「部署 → 管理部署作業」編輯既有部署，版本選「新版本」後部署（網址不變）。
 * 報表：重新整理試算表後，上方選單會出現「訓練分析 → 更新分析報表」。
 */
const SHEET_RUNS = 'AAR紀錄';
const SHEET_LOG  = '時間軸';
const SHEET_SUM  = '分析總覽';
const SHEET_ANA  = '處置分析';
const SHEET_FB   = '意見回饋';
const SHEET_FBS  = '意見回饋彙整';
const LEGACY_VER = '舊版（未標示版本）';
const REPORT_VERSION = '最新';   // '最新' = 只分析最新版；'全部' = 不分版本；也可填指定版本，例如 '2026.10.13'
let REPORT_SCOPE_ = '全部', ALL_RUNS_ = [];
const REPORT_MODE = '';   // '' = 全部；改成 '考核' 或 '練習' 則只分析該模式
const FIXED_COLS = ['紀錄編號','上傳時間','姓名','單位','劇本','出勤人數','模式','結束原因','演練秒數','救出人數','通過檢核點','受困者結果','閃燃(秒)','MAYDAY','版本','支援單位','各隊人數'];
const NAME_RE = /帶隊官|隊員[甲乙丙]|增援隊員\d+|救助隊員\d+|支援隊員\d+|[AB]隊長|[AB]隊員\d+/g;

/* ================= 接收端 ================= */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad-json' }); }
  // 多人同時上傳時排隊等候（最久 60 秒）；等不到就回報忙碌，讓前端自動重試，不會當掉
  const lock = LockService.getScriptLock();
  try { lock.waitLock(60000); }
  catch (err) { return json_({ ok: false, busy: true, error: 'busy' }); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (d.type === 'feedback') {
      const fb = getSheet_(ss, SHEET_FB);
      if (fb.getLastRow() > 1 && d.fbId) {
        const hit = fb.getRange(2, 12, fb.getLastRow() - 1, 1).createTextFinder(String(d.fbId)).matchEntireCell(true).findNext();
        if (hit) return json_({ ok: true, dup: true });
      }
      if (fb.getLastRow() === 0) {
        fb.appendRow(['送出時間', '姓名', '單位', '情境', '真實度', '難度', '時間與人力', '不合理之處', '建議新增或修改', '關聯演練編號', '模式', '回饋編號']);
        fb.getRange(1, 1, 1, 12).setFontWeight('bold'); fb.setFrozenRows(1);
      }
      if (fb.getLastColumn() < 13) fb.getRange(1, 13).setValue('版本').setFontWeight('bold');
      fb.appendRow([new Date(), d.name, d.unit, d.scenario, d.real, d.diff, d.timing, d.bad, d.sug, d.runId, d.mode === 'exam' ? '考核' : d.mode === 'practice' ? '練習' : '', d.fbId, d.ver || '']);
      return json_({ ok: true });
    }
    const runs = getSheet_(ss, SHEET_RUNS);

    if (runs.getLastRow() > 1) {
      const hit = runs.getRange(2, 1, runs.getLastRow() - 1, 1)
        .createTextFinder(String(d.runId)).matchEntireCell(true).findNext();
      if (hit) return json_({ ok: true, dup: true });
    }

    const mark = p => p === true ? '✓ ' : p === false ? '✕ ' : '－ ';
    const row = {
      '紀錄編號': d.runId,
      '上傳時間': new Date(),
      '姓名': d.name,
      '單位': d.unit,
      '劇本': d.scenario,
      '出勤人數': d.crew,
      '模式': d.mode === 'exam' ? '考核' : '練習',
      '結束原因': d.reason,
      '演練秒數': d.duration,
      '救出人數': `${d.saved}/${d.total}`,
      '通過檢核點': `${d.passCount}/${d.checkCount}`,
      '受困者結果': (d.victims || []).map(v => `${v.name}：${v.status}`).join('；'),
      '閃燃(秒)': d.flashover ?? '',
      'MAYDAY': d.mayday ? `T+${d.mayday.t}s ${d.mayday.reason}${d.mayday.rit ? '（RIT救援）' : ''}` : ''
    };
    row['版本'] = d.ver || ''; row['支援單位'] = d.extraUnits || 0; row['各隊人數'] = d.unitSizes || '';
    (d.checks || []).forEach(c => row[c.item] = mark(c.pass) + c.detail);
    appendByHeader_(runs, row);

    const logSh = getSheet_(ss, SHEET_LOG);
    if (logSh.getLastRow() === 0) {
      logSh.appendRow(['紀錄編號', '姓名', 'T(秒)', '類型', '內容']);
      logSh.getRange(1, 1, 1, 5).setFontWeight('bold');
      logSh.setFrozenRows(1);
    }
    const rows = (d.log || []).map(l => [d.runId, d.name, l.t, l.type, l.msg]);
    if (rows.length) logSh.getRange(logSh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, msg: 'AAR 接收器運作中' });
}

function appendByHeader_(sh, row) {
  const lastCol = sh.getLastColumn();
  const headers = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  Object.keys(row).forEach(k => {
    if (!headers.includes(k)) {
      headers.push(k);
      sh.getRange(1, headers.length).setValue(k).setFontWeight('bold');
    }
  });
  sh.appendRow(headers.map(h => row[h] ?? ''));
  if (sh.getFrozenRows() === 0) sh.setFrozenRows(1);
}
function getSheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= 選單 ================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('訓練分析')
    .addItem('更新分析報表', 'buildReports')
    .addSeparator()
    .addItem('開啟每小時自動更新', 'enableAutoUpdate')
    .addItem('關閉自動更新', 'disableAutoUpdate')
    .addToUi();
}
function enableAutoUpdate() {
  disableAutoUpdate();
  ScriptApp.newTrigger('buildReports').timeBased().everyHours(1).create();
  SpreadsheetApp.getActive().toast('已開啟每小時自動更新報表', '訓練分析');
}
function disableAutoUpdate() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'buildReports').forEach(t => ScriptApp.deleteTrigger(t));
}

/* ================= 報表 ================= */
function buildReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let runs = readRuns_(ss);
  if (REPORT_MODE) runs = runs.filter(r => r.mode === REPORT_MODE);
  ALL_RUNS_ = runs;
  let scope = REPORT_VERSION;
  if (scope === '最新') {
    const vs = runs.map(r => r.ver).filter(v => v !== LEGACY_VER).sort();
    scope = vs.length ? vs[vs.length - 1] : '全部';
  }
  if (scope !== '全部') runs = runs.filter(r => r.ver === scope);
  REPORT_SCOPE_ = scope;
  const ids = new Set(runs.map(r => r.id));
  const logs = readLogs_(ss).filter(l => ids.has(l.id));
  buildSummary_(ss, runs);
  buildAnalysis_(ss, runs, logs);
  try { buildFeedback_(ss); } catch (e) {}
  try { ss.toast(`已更新，共 ${runs.length} 場演練`, '訓練分析'); } catch (e) {}
}

function readRuns_(ss) {
  const sh = ss.getSheetByName(SHEET_RUNS);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues(), h = v[0], ix = k => h.indexOf(k);
  const checkCols = h.map((k, i) => [String(k), i]).filter(([k]) => k && !FIXED_COLS.includes(k));
  const frac = s => { const m = String(s).match(/(\d+)\s*\/\s*(\d+)/); return m && +m[2] ? +m[1] / +m[2] : null; };
  return v.slice(1).filter(r => r[ix('紀錄編號')] !== '').map(r => {
    const checks = {};
    checkCols.forEach(([k, i]) => {
      const s = String(r[i] || '').trim(); if (!s) return;
      checks[k] = s.startsWith('✓') ? 'pass' : s.startsWith('✕') ? 'fail' : 'na';
    });
    const t = r[ix('上傳時間')];
    return {
      id: String(r[ix('紀錄編號')]), time: t instanceof Date ? t : new Date(t),
      name: String(r[ix('姓名')] || '未填'), unit: String(r[ix('單位')] || '未填'),
      sc: String(r[ix('劇本')] || '未知'), crew: r[ix('出勤人數')], mode: String(r[ix('模式')] || ''),
      ver: ix('版本') >= 0 && String(r[ix('版本')] || '').trim() ? String(r[ix('版本')]).trim() : LEGACY_VER,
      ext: ix('支援單位') >= 0 ? (+r[ix('支援單位')] || 0) : 0,
      dur: +r[ix('演練秒數')] || 0, savedRate: frac(r[ix('救出人數')]), passRate: frac(r[ix('通過檢核點')]),
      mayday: !!String(r[ix('MAYDAY')] || '').trim(), checks
    };
  });
}
function readLogs_(ss) {
  const sh = ss.getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .map(r => ({ id: String(r[0]), t: +r[2] || 0, type: String(r[3]), msg: String(r[4]) }));
}

/* ---------- 工具 ---------- */
const avg_ = a => { const b = a.filter(x => typeof x === 'number' && !isNaN(x)); return b.length ? b.reduce((s, x) => s + x, 0) / b.length : ''; };
const groupBy_ = (a, f) => a.reduce((m, x) => { const k = f(x); (m[k] = m[k] || []).push(x); return m; }, {});
const fmtDate_ = d => (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : '';
const secFmt_ = s => { if (s === '' || s == null) return ''; s = Math.round(s); return `T+${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`; };
const rateBg_ = r => r === '' || r == null ? '#ffffff' : r < 0.5 ? '#f8d0d0' : r < 0.75 ? '#fff0c2' : '#d4edda';

function resetSheet_(ss, name) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.clear();
  return sh;
}
function title_(sh, text, sub) {
  sh.getRange(1, 1).setValue(text).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1).setValue(sub).setFontColor('#666666');
}
/** 寫入表格；fmt = {欄序: 數字格式}；回傳位置資訊 */
function table_(sh, row, title, headers, rows, fmt) {
  sh.getRange(row, 1).setValue(title).setFontWeight('bold').setFontSize(12);
  const hr = row + 1;
  sh.getRange(hr, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#16202a').setFontColor('#ffffff');
  if (rows.length) {
    sh.getRange(hr + 1, 1, rows.length, headers.length).setValues(rows).setVerticalAlignment('middle');
    Object.entries(fmt || {}).forEach(([c, f]) => sh.getRange(hr + 1, 1 + +c, rows.length, 1).setNumberFormat(f));
  } else {
    sh.getRange(hr + 1, 1).setValue('（尚無資料）').setFontColor('#888888');
  }
  return { headerRow: hr, firstRow: hr + 1, n: rows.length, rows, end: hr + 1 + Math.max(rows.length, 1) };
}
function colorRates_(sh, t, cols) {
  if (!t.n) return;
  cols.forEach(c => sh.getRange(t.firstRow, 1 + c, t.n, 1).setBackgrounds(t.rows.map(r => [rateBg_(r[c])])));
}
/** 以表格第 1 欄為分類、valueCol 為數值畫圖；回傳圖表佔用的列數 */
function chart_(sh, t, valueCol, n, title, type, col) {
  if (!n) return 0;
  const h = type === Charts.ChartType.BAR ? Math.max(240, n * 28 + 80) : 300;
  const axis = { format: 'percent', viewWindow: { min: 0, max: 1 } };
  const b = sh.newChart().setChartType(type)
    .addRange(sh.getRange(t.headerRow, 1, n + 1, 1))
    .addRange(sh.getRange(t.headerRow, 1 + valueCol, n + 1, 1))
    .setNumHeaders(1)
    .setPosition(t.headerRow, col, 0, 0)
    .setOption('title', title)
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#c8102e'])
    .setOption('width', 560).setOption('height', h)
    .setOption(type === Charts.ChartType.BAR ? 'hAxis' : 'vAxis', axis);
  sh.insertChart(b.build());
  return Math.ceil(h / 21);
}

/* ---------- 分析總覽 ---------- */
function buildSummary_(ss, runs) {
  const sh = resetSheet_(ss, SHEET_SUM);
  title_(sh, '消防指揮情境訓練｜分析總覽', `更新時間 ${fmtDate_(new Date())}　資料範圍：${REPORT_MODE || '全部模式'}｜版本：${REPORT_SCOPE_}　共 ${runs.length} 場（改版前後難度不同，不混算；各版本對照見最下方）`);
  let row = 4;

  // 一、各情境總覽
  const bySc = groupBy_(runs, r => r.sc);
  const A = Object.keys(bySc).sort().map(k => {
    const g = bySc[k];
    return [k, g.length, new Set(g.map(r => r.name)).size, avg_(g.map(r => r.savedRate)), avg_(g.map(r => r.passRate)), g.filter(r => r.mayday).length, avg_(g.map(r => r.dur / 60))];
  });
  const nSc = A.length;
  if (runs.length) A.push(['合計', runs.length, new Set(runs.map(r => r.name)).size, avg_(runs.map(r => r.savedRate)), avg_(runs.map(r => r.passRate)), runs.filter(r => r.mayday).length, avg_(runs.map(r => r.dur / 60))]);
  let t = table_(sh, row, '一、各情境總覽', ['劇本', '場次', '參與人數', '平均救出率', '平均檢核通過率', 'MAYDAY 場次', '平均演練時間（分）'], A, { 3: '0%', 4: '0%', 6: '0.0' });
  colorRates_(sh, t, [3, 4]);
  if (t.n) sh.getRange(t.firstRow + t.n - 1, 1, 1, 7).setFontWeight('bold');
  const ch = chart_(sh, t, 4, nSc, '各情境平均檢核通過率', Charts.ChartType.COLUMN, 9);
  row = Math.max(t.end, t.headerRow + ch) + 2;

  // 二、出勤編組與模式比較
  const byCrew = groupBy_(runs, r => [r.sc, r.crew, r.ext, r.mode].join('|'));
  const B = Object.keys(byCrew).sort().map(k => {
    const g = byCrew[k], [sc, crew, ext, mode] = k.split('|');
    return [sc, (crew ? `本隊 ${crew} 人` : '') + (+ext ? `＋支援 ${ext} 隊` : ''), mode, g.length, avg_(g.map(r => r.savedRate)), avg_(g.map(r => r.passRate)), g.filter(r => r.mayday).length];
  });
  t = table_(sh, row, '二、出勤編組與模式比較', ['劇本', '出勤編組', '模式', '場次', '平均救出率', '平均檢核通過率', 'MAYDAY 場次'], B, { 4: '0%', 5: '0%' });
  colorRates_(sh, t, [4, 5]);
  row = t.end + 2;

  // 三、單位比較
  const byUnit = groupBy_(runs, r => r.unit);
  const C = Object.keys(byUnit).sort().map(k => {
    const g = byUnit[k];
    return [k, new Set(g.map(r => r.name)).size, g.length, avg_(g.map(r => r.passRate)), avg_(g.map(r => r.savedRate)), g.filter(r => r.mayday).length];
  }).sort((a, b) => (b[3] || 0) - (a[3] || 0));
  t = table_(sh, row, '三、單位比較', ['單位', '人數', '場次', '平均檢核通過率', '平均救出率', 'MAYDAY 場次'], C, { 3: '0%', 4: '0%' });
  colorRates_(sh, t, [3, 4]);
  row = t.end + 2;

  // 四、個人紀錄
  const byP = groupBy_(runs, r => r.name + '|' + r.unit);
  const D = Object.keys(byP).map(k => {
    const g = byP[k].slice().sort((a, b) => a.time - b.time), first = g[0], last = g[g.length - 1];
    const rates = g.map(r => r.passRate).filter(x => x != null);
    const prog = g.length > 1 && first.passRate != null && last.passRate != null ? last.passRate - first.passRate : '';
    return [first.name, first.unit, g.length, avg_(rates), rates.length ? Math.max(...rates) : '', last.passRate ?? '', prog, last.sc, fmtDate_(last.time)];
  }).sort((a, b) => b[2] - a[2]);
  t = table_(sh, row, '四、個人紀錄', ['姓名', '單位', '場次', '平均通過率', '最佳通過率', '最近一次通過率', '進步幅度（最近－首次）', '最近劇本', '最近演練'], D, { 3: '0%', 4: '0%', 5: '0%', 6: '+0%;-0%;0%' });
  colorRates_(sh, t, [3, 5]);

  row = t.end + 2;
  const byV = groupBy_(ALL_RUNS_, r => [r.ver, r.sc].join('|'));
  const VV = Object.keys(byV).sort().map(k => {
    const g = byV[k], [v, sc] = k.split('|');
    return [v, sc, g.length, avg_(g.map(r => r.savedRate)), avg_(g.map(r => r.passRate)), g.filter(r => r.mayday).length];
  });
  t = table_(sh, row, '五、版本對照（含舊版）', ['版本', '劇本', '場次', '平均救出率', '平均檢核通過率', 'MAYDAY 場次'], VV, { 3: '0%', 4: '0%' });
  colorRates_(sh, t, [3, 4]);

  sh.setColumnWidth(1, 300);
  for (let c = 2; c <= 9; c++) sh.setColumnWidth(c, 115);
  sh.setFrozenRows(2);
}

/* ---------- 意見回饋彙整 ---------- */
function buildFeedback_(ss) {
  const src = ss.getSheetByName(SHEET_FB), sh = resetSheet_(ss, SHEET_FBS);
  title_(sh, '消防指揮情境訓練｜意見回饋彙整', `更新時間 ${fmtDate_(new Date())}　資料來源：「${SHEET_FB}」工作表`);
  if (!src || src.getLastRow() < 2) { sh.getRange(4, 1).setValue('尚無意見回饋'); return; }
  const v = src.getDataRange().getValues(), h = v[0], ix = k => h.indexOf(k);
  const get = (r, k) => ix(k) >= 0 ? String(r[ix(k)] == null ? '' : r[ix(k)]) : '';
  const rows = v.slice(1).filter(r => r[ix('送出時間')] !== '').map(r => ({
    time: r[ix('送出時間')], name: get(r, '姓名'), unit: get(r, '單位'), sc: get(r, '情境') || '整體',
    real: get(r, '真實度'), diff: get(r, '難度'), timing: get(r, '時間與人力'),
    bad: get(r, '不合理之處'), sug: get(r, '建議新增或修改'), ver: get(r, '版本') || '舊版'
  }));
  const num = s => { const m = s.match(/^(\d)/); return m ? +m[1] : null; };
  const bySc = groupBy_(rows, r => r.sc);
  const A = Object.keys(bySc).sort().map(k => {
    const g = bySc[k], cnt = (f, xs) => g.filter(r => xs.includes(r[f])).length;
    return [k, g.length, avg_(g.map(r => num(r.real))), cnt('diff', ['太簡單', '偏簡單']), cnt('diff', ['剛好']), cnt('diff', ['偏難', '太難']), g.filter(r => r.timing && r.timing !== '大致合理').length];
  });
  const t = table_(sh, 4, '一、各情境回饋統計', ['情境', '回饋數', '真實度平均（1–5）', '偏簡單', '剛好', '偏難', '時間人力有疑義'], A, { 2: '0.0' });
  const B = rows.filter(r => r.bad || r.sug || (r.timing && r.timing !== '大致合理'))
    .sort((a, b) => new Date(b.time) - new Date(a.time))
    .map(r => [fmtDate_(r.time instanceof Date ? r.time : new Date(r.time)), r.sc, r.timing, r.bad, r.sug, r.name, r.unit, r.ver]);
  const t2 = table_(sh, t.end + 2, '二、文字意見與數值疑義（新到舊）', ['時間', '情境', '時間與人力', '不合理之處', '建議新增或修改', '姓名', '單位', '版本'], B, {});
  if (t2.n) sh.getRange(t2.firstRow, 4, t2.n, 2).setWrap(true);
  [150, 90, 110, 300, 300, 80, 100, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(2);
}

/* ---------- 處置分析（依情境） ---------- */
function actionKey_(l) {
  const m = l.msg;
  if (l.type === 'cmd') { const i = m.indexOf('：'); return i < 0 ? null : m.slice(i + 1).trim(); }
  if (l.type === 'radio') {
    if (m.startsWith('初期報告')) return '（指揮）初期報告';
    if (m.startsWith('帶隊官：請求')) return '（指揮）請求增援';
    if (m.startsWith('PAR')) return '（指揮）PAR 點名';
  }
  return null;
}
function buildAnalysis_(ss, runs, logs) {
  const sh = resetSheet_(ss, SHEET_ANA);
  title_(sh, '消防指揮情境訓練｜處置作為分析', `更新時間 ${fmtDate_(new Date())}　資料範圍：${REPORT_MODE || '全部模式'}｜版本：${REPORT_SCOPE_}　通過率紅色＜50%、黃色＜75%`);
  const logsBy = groupBy_(logs, l => l.id);
  const bySc = groupBy_(runs, r => r.sc);
  let row = 4;
  if (!runs.length) { sh.getRange(row, 1).setValue('（尚無資料）').setFontColor('#888888'); return; }

  Object.keys(bySc).sort().forEach(sc => {
    const g = bySc[sc];
    sh.getRange(row, 1).setValue(`■ ${sc}（${g.length} 場，平均通過率 ${Math.round((avg_(g.map(r => r.passRate)) || 0) * 100)}%）`)
      .setFontSize(14).setFontWeight('bold').setFontColor('#c8102e');
    row += 2;

    // 1. 檢核點通過率
    const it = {};
    g.forEach(r => Object.entries(r.checks).forEach(([k, st]) => {
      if (st === 'na') return;
      it[k] = it[k] || { p: 0, f: 0 }; it[k][st === 'pass' ? 'p' : 'f']++;
    }));
    const C = Object.entries(it).map(([k, o]) => [k, o.p + o.f, o.p, o.f, o.p / (o.p + o.f)]).sort((a, b) => a[4] - b[4]);
    const t1 = table_(sh, row, '1. 檢核點通過率（由低到高，越上面越需要加強）', ['檢核點', '適用場次', '通過', '未通過', '通過率'], C, { 4: '0%' });
    colorRates_(sh, t1, [4]);
    const ch = chart_(sh, t1, 4, C.length, `${sc}｜檢核點通過率`, Charts.ChartType.BAR, 8);
    row = Math.max(t1.end, t1.headerRow + ch) + 2;

    // 2. 處置作為首次執行時間
    const first = {};
    g.forEach(r => {
      const seen = {};
      (logsBy[r.id] || []).forEach(l => { const k = actionKey_(l); if (k && seen[k] == null) seen[k] = l.t; });
      Object.entries(seen).forEach(([k, t]) => (first[k] = first[k] || []).push(t));
    });
    const D = Object.entries(first)
      .map(([k, ts]) => [k, ts.length, ts.length / g.length, avg_(ts), Math.min(...ts), Math.max(...ts)])
      .sort((a, b) => a[3] - b[3])
      .map(r => [r[0], r[1], r[2], secFmt_(r[3]), secFmt_(r[4]), secFmt_(r[5])]);
    const t2 = table_(sh, row, '2. 處置作為首次執行時間（依平均時間排序＝典型處置順序；執行率低代表多數人沒做）', ['處置作為', '執行場次', '執行率', '平均首次時間', '最快', '最慢'], D, { 2: '0%' });
    colorRates_(sh, t2, [2]);
    row = t2.end + 2;

    // 3. 常見現場警訊
    const al = {};
    g.forEach(r => {
      const s = new Set();
      (logsBy[r.id] || []).forEach(l => { if (l.type === 'alert') s.add(l.msg.replace(NAME_RE, '○○')); });
      s.forEach(m => al[m] = (al[m] || 0) + 1);
    });
    const E = Object.entries(al).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([m, n]) => [m, n, n / g.length]);
    const t3 = table_(sh, row, '3. 常見現場警訊（每場只計一次，反映最常犯的失誤後果）', ['狀況', '發生場次', '發生率'], E, { 2: '0%' });
    row = t3.end + 3;
  });

  sh.setColumnWidth(1, 460);
  for (let c = 2; c <= 6; c++) sh.setColumnWidth(c, 105);
  sh.getRange(1, 1, row, 1).setWrap(true);
  sh.setFrozenRows(2);
}
