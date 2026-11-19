// JPX週次「信用取引現在高」(YYYYMMDD_mtcurrent.xlsx) の自動取得と履歴シート。
// ページ margin/04.html のリンクから最新週を辿る（GASから取得できることを2026-10-24に確認）。
// 東京市場の合計（売残・買残・前週比）を読み、'信用残推移' シートに申込み日ごとに1行ずつ貯める。
const JPX_MARGIN = {
  BASE: 'https://www.jpx.co.jp',
  PAGE: '/markets/statistics-equities/margin/04.html',
  SHEET: '信用残推移',
  HEADER: ['申込み日', '売残(億円)', '売残 前週比', '買残(億円)', '買残 前週比', '信用倍率(株数)', '信用倍率(金額)', '取得日'],
  WEEKS: 4,           // ページに載っている直近週のうち、取り込む上限
  STALE_DAYS: 14,     // 申込み日からこの日数を超えたら「古い」と警告
  TAB_COLOR: '#795548',
  SHEET2: '2市場信用倍率',
  HEADER2: ['日付', '売り残(百万円)', '買い残(百万円)', '信用倍率', '信用評価損益率'],
  TW_URL: 'https://www.traders.co.jp/margin_derivatives/margin_transition',  // 信用評価損益率の取得元（JPXには無い）
  RED_SELL_MAX: 800000,   // 売り残がこの値(百万円＝8000億円)以下なら赤字
};

// 変換済みグリッドから「信用取引残高合計」の東京市場の値を読む。
// 株数行: [見出し, '株数', 二市場(売,前週比,買,前週比), 東京(売,前週比,買,前週比)] / 直下の金額行（百万円）。
// 返り値 { date:'yyyy/MM/dd'|'', sellOku, sellChgOku, buyOku, buyChgOku, ratioShares, ratioVal } / 読めなければ null
function parseJpxMtcurrentGrid_(grid) {
  const num = v => {
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v == null ? '' : v).replace(/[,\s　]/g, '').replace(/▲/, '-'));
    return isFinite(n) ? n : NaN;
  };
  const filled = row => (row || []).filter(c => c !== '' && c != null);
  let date = '';
  for (let i = 0; i < Math.min(grid.length, 4) && !date; i++) {
    const c = filled(grid[i])[0];
    if (c instanceof Date) date = Utilities.formatDate(c, 'Asia/Tokyo', 'yyyy/MM/dd');
    else if (/^20\d{2}[\/-]\d{1,2}[\/-]\d{1,2}/.test(String(c || ''))) date = String(c).slice(0, 10).replace(/-/g, '/');
  }
  for (let i = 0; i < grid.length; i++) {
    const sh = filled(grid[i]);
    if (!/^信用取引残高合計/.test(String(sh[0] || ''))) continue;
    const val = filled(grid[i + 1]);
    const shs = sh.slice(2).map(num), mil = val.slice(1).map(num);
    if (shs.length < 8 || mil.length < 8 || !/金額|Val/i.test(String(val[0] || ''))) return null;
    if (![shs[4], shs[6], mil[4], mil[5], mil[6], mil[7]].every(isFinite) || shs[4] <= 0 || mil[4] <= 0) return null;
    const r2 = x => Math.round(x * 100) / 100;
    return {
      date: date,
      sellOku: Math.round(mil[4] / 100), sellChgOku: Math.round(mil[5] / 100),
      buyOku: Math.round(mil[6] / 100), buyChgOku: Math.round(mil[7] / 100),
      ratioShares: r2(shs[6] / shs[4]), ratioVal: r2(mil[6] / mil[4]),
      // 二市場（東京＋名古屋）合計・百万円。「2市場信用倍率」シート用（mil[0]=売, mil[2]=買）
      sell2M: mil[0], buy2M: mil[2], ratio2M: isFinite(mil[0]) && mil[0] > 0 ? r2(mil[2] / mil[0]) : NaN,
    };
  }
  return null;
}

// ページHTMLから _mtcurrent.xlsx のリンクを新しい順に返す（ファイル名の日付で並べ直す）。
function parseJpxMtcurrentLinks_(html) {
  const links = [];
  String(html || '').replace(/href="([^"]*?(20\d{6})_mtcurrent\.xlsx)"/g, (m, href, ymd) => { links.push({ href: href, ymd: ymd }); return m; });
  const seen = {};
  return links.filter(l => !seen[l.ymd] && (seen[l.ymd] = true)).sort((a, b) => b.ymd.localeCompare(a.ymd));
}

// xlsx を一時的にスプレッドシートへ変換して読み、一時ファイルは必ず消す。
function readJpxMtcurrentXlsx_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (res.getResponseCode() !== 200) throw new Error('JPX xlsx HTTP ' + res.getResponseCode());
  const blob = res.getBlob().setName('__tmp_jpx_mtcurrent.xlsx');
  let id = null;
  try {
    id = Drive.Files.create({ name: '__tmp_jpx_mtcurrent', mimeType: 'application/vnd.google-apps.spreadsheet' }, blob, { supportsAllDrives: true }).id;
    return parseJpxMtcurrentGrid_(SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues());
  } finally {
    if (id) { try { Drive.Files.remove(id); } catch (e) { /* 一時ファイルの掃除失敗は無視 */ } }
  }
}

function ensureJpxMarginSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(JPX_MARGIN.SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(JPX_MARGIN.SHEET);
  const H = JPX_MARGIN.HEADER;
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#37474F').setFontColor('#FFFFFF');
  sh.setColumnWidths(1, H.length, 110);
  sh.setTabColor(JPX_MARGIN.TAB_COLOR);
  sh.getRange(2, 2, 500, 5).setNumberFormat('#,##0');
  sh.getRange(2, 6, 500, 2).setNumberFormat('0.00');
  if (typeof trimSheet_ === 'function') trimSheet_(sh);
  try { if (typeof SheetOrder !== 'undefined' && SheetOrder.saved()) SheetOrder.lock(ss); } catch (e) { /* 固定記録が無ければ何もしない */ }
  return sh;
}

// 直近の週を取り込み、履歴シートに無い申込み日だけ追記する。返り値は最新週の値（無ければ null）。
function syncJpxMarginHistory_() {
  const page = UrlFetchApp.fetch(JPX_MARGIN.BASE + JPX_MARGIN.PAGE, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (page.getResponseCode() !== 200) throw new Error('JPXページ HTTP ' + page.getResponseCode());
  const links = parseJpxMtcurrentLinks_(page.getContentText()).slice(0, JPX_MARGIN.WEEKS);
  if (!links.length) throw new Error('JPXページに _mtcurrent.xlsx のリンクがありません');
  const sh = ensureJpxMarginSheet_();
  const have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().forEach(r => { have[r[0]] = true; });
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  let latest = null;
  links.slice().reverse().forEach(l => {   // 古い週から順に追記して、シートは日付昇順を保つ
    const ymd = l.ymd.slice(0, 4) + '/' + l.ymd.slice(4, 6) + '/' + l.ymd.slice(6, 8);
    const isNewest = l === links[0];
    if (have[ymd] && !isNewest) return;
    const r = readJpxMtcurrentXlsx_(JPX_MARGIN.BASE + l.href);
    if (!r) throw new Error('mtcurrent のパース失敗: ' + l.ymd);
    r.date = r.date || ymd;
    if (isNewest) latest = r;
    if (have[r.date]) return;
    sh.appendRow([r.date, r.sellOku, r.sellChgOku, r.buyOku, r.buyChgOku, r.ratioShares, r.ratioVal, today]);
    have[r.date] = true;
  });
  return latest;
}

// 「2市場信用倍率」シート：二市場合計の売り残・買い残・信用倍率を週1行で貯める。
// 売り残 8000億円(=800,000百万円)以下は赤字（条件付き書式。追記行にも効く）。
// 信用評価損益率は JPX の mtcurrent に無いので E列は手入力（空欄のまま）。
function ensureJpx2MarketSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(JPX_MARGIN.SHEET2);
  if (sh) return sh;
  sh = ss.insertSheet(JPX_MARGIN.SHEET2);
  const H = JPX_MARGIN.HEADER2;
  sh.getRange(1, 1, 1, H.length).setValues([H]);
  sh.setColumnWidths(1, H.length, 130);
  sh.setTabColor(JPX_MARGIN.TAB_COLOR);
  if (typeof trimSheet_ === 'function') trimSheet_(sh);   // 空白の行・列はカット（SheetSize.js）。以降の書き込みは ensureSheetSize_ を通す
  return sh;
}

// 表示ルール（毎回当て直す。行が増えても範囲がずれない）:
//  ・新しい順（日付の降順）・行列の固定なし・金額は3桁カンマ・ヘッダ＋行バンディング（styleSheet_）
//  ・売り残 8000億円(=800,000百万円)以下は赤字、信用倍率は紫、評価損益率は符号色（正=赤 #db1616 / 負=青 #003cbe）
function formatJpx2MarketSheet_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return;
  const r = n - 1;
  sh.getRange(2, 1, r, 5).sort({ column: 1, ascending: false });   // 範囲指定の sort（固定行に頼らない）
  if (typeof styleSheet_ === 'function') styleSheet_(sh, 5, '#37474F', '#FDF3D8');
  const body = sh.getRange(2, 1, r, 5);
  body.setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#000000');
  sh.getRange(2, 1, r, 1).setNumberFormat('yyyy/m/d');
  sh.getRange(2, 2, r, 2).setNumberFormat('#,##0');
  sh.getRange(2, 4, r, 2).setNumberFormat('0.00');
  sh.getRange(2, 4, r, 1).setFontColor('#7030A0');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($B2),$B2<=' + JPX_MARGIN.RED_SELL_MAX + ')')
      .setFontColor('#FF0000').setRanges([sh.getRange(2, 2, r, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor('#DB1616').setRanges([sh.getRange(2, 5, r, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
      .setFontColor('#003CBE').setRanges([sh.getRange(2, 5, r, 1)]).build(),
  ];
  sh.setConditionalFormatRules(rules);
}

// 直近週を取り込み、2市場信用倍率シートに無い日付だけ追記する（日付昇順）。追記した件数を返す。
function syncJpx2MarketSheet_() {
  const page = UrlFetchApp.fetch(JPX_MARGIN.BASE + JPX_MARGIN.PAGE, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (page.getResponseCode() !== 200) throw new Error('JPXページ HTTP ' + page.getResponseCode());
  const links = parseJpxMtcurrentLinks_(page.getContentText()).slice(0, JPX_MARGIN.WEEKS);
  if (!links.length) throw new Error('JPXページに _mtcurrent.xlsx のリンクがありません');
  const sh = ensureJpx2MarketSheet_();
  const have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().forEach(r => { have[r[0].replace(/\//g, '/')] = true; });
  let added = 0;
  links.slice().reverse().forEach(l => {
    const ymd = l.ymd.slice(0, 4) + '/' + l.ymd.slice(4, 6) + '/' + l.ymd.slice(6, 8);
    const key = ymd.replace(/\/0(\d)/g, '/$1');   // 表示は 2026/11/13 → yyyy/m/d の 2026/11/13、月日の0は落ちる
    if (have[key] || have[ymd]) return;
    const r = readJpxMtcurrentXlsx_(JPX_MARGIN.BASE + l.href);
    if (!r || !isFinite(r.sell2M)) throw new Error('二市場の値を読めません: ' + l.ymd);
    const d = new Date(Number(l.ymd.slice(0, 4)), Number(l.ymd.slice(4, 6)) - 1, Number(l.ymd.slice(6, 8)));
    if (sh.getLastRow() < 2) { if (typeof ensureSheetSize_ === 'function') ensureSheetSize_(sh, 2, 5); }
    else sh.insertRowBefore(2);                       // 新しい週を一番上へ（新しい順）
    sh.getRange(2, 1, 1, 5).setValues([[d, r.sell2M, r.buy2M, r.ratio2M, '']]);
    have[key] = true; added++;
  });
  return added;
}

// 週次ジョブ：JPXの新しい週が出ていたら2市場信用倍率シートに追記し、追記があった時だけメールする。
// 毎日1回走らせて（JPXの公開日は週によりずれる）、新データが無い日は何も送らない。
function jpxWeeklyJob() {
  const added = syncJpx2MarketSheet_();
  const pl = fillJpx2MarketPl_();   // 空きの評価損益率は毎回埋め直す
  if (pl.note) Logger.log(pl.note);
  const sh = SpreadsheetApp.getActive().getSheetByName(JPX_MARGIN.SHEET2);
  if (!sh || sh.getLastRow() < 2) return 0;
  formatJpx2MarketSheet_(sh);                       // 表示ルールを毎回当て直す（並び・書式・条件付き書式）
  const last = sh.getRange(2, 1, 1, 5).getValues()[0];   // 新しい順なので2行目が最新週
  const key = Utilities.formatDate(new Date(last[0]), 'Asia/Tokyo', 'yyyy/MM/dd');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('JPX2M_MAILED') === key) { Logger.log('JPX 2市場信用倍率: ' + key + ' は送信済み'); return 0; }
  // 更新時刻の実測：JPXは第2営業日16:00(JST)、評価損益率(トレーダーズ・ウェブ)は第3営業日まで。
  // 評価損益率が入るまで待って1通だけ送る。申込日から7日たっても入らなければ、欠けたまま送る。
  const ready = last[4] !== '' && last[4] != null;
  const overdue = (Date.now() - new Date(last[0]).getTime()) / 86400000 >= 7;
  if (!ready && !overdue) { Logger.log('JPX 2市場信用倍率: ' + key + ' は評価損益率待ち'); return added; }
  const date = Utilities.formatDate(new Date(last[0]), 'Asia/Tokyo', 'yyyy/M/d');
  const low = last[1] <= JPX_MARGIN.RED_SELL_MAX;
  const body = [
    date + ' 申込み週の信用取引現在高（二市場合計）を「' + JPX_MARGIN.SHEET2 + '」に追記しました。',
    '',
    '売り残: ' + Number(last[1]).toLocaleString('en-US') + ' 百万円' + (low ? '  ★8000億円以下（赤）' : ''),
    '買い残: ' + Number(last[2]).toLocaleString('en-US') + ' 百万円',
    '信用倍率: ' + last[3],
    '信用評価損益率: ' + (ready ? last[4] + '%' : '未取得（トレーダーズ・ウェブ未掲載のまま7日経過）'),
    pl.note ? '※' + pl.note : '',
    '',
    SpreadsheetApp.getActive().getUrl(),
  ].join('\n');
  sendMail_('【信用残】' + date + ' 売り残' + Math.round(last[1] / 100) + '億円・倍率' + last[3] + (ready ? '・評価損益率' + last[4] + '%' : '') + (low ? '（8000億円以下）' : ''), body);
  props.setProperty('JPX2M_MAILED', key);
  return added;
}

// 毎日19時に jpxWeeklyJob を走らせるトリガーを張り直す（メニューまたは clasp run から1回実行）。
function installJpxWeeklyTrigger() {
  clearTriggersFor_('jpxWeeklyJob');
  ScriptApp.newTrigger('jpxWeeklyJob').timeBased().everyDays(1).atHour(19).create();
}

// トレーダーズ・ウェブ「信用残の推移」のHTMLから 申込日 → {sellOku, buyOku, pl, ratio} を読む（純関数）。
// 列: 申込日, 売(株数,前週比,金額億円,前週比), 買(株数,前週比,金額億円,前週比), 評価損益率(%), 信用倍率
function parseTradersWebMargin_(html) {
  const out = {};
  const num = v => { const n = parseFloat(String(v).replace(/[,\s]/g, '')); return isFinite(n) ? n : NaN; };
  (String(html || '').match(/<tr[^>]*>[\s\S]*?<\/tr>/g) || []).forEach(tr => {
    const cells = (tr.match(/<td[^>]*>[\s\S]*?<\/td>/g) || []).map(c => c.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim());
    if (cells.length < 11 || !/^20\d\d\/\d\d\/\d\d$/.test(cells[0])) return;
    const v = { sellOku: num(cells[3]), buyOku: num(cells[7]), pl: num(cells[9]), ratio: num(cells[10]) };
    if ([v.sellOku, v.buyOku, v.pl, v.ratio].every(isFinite)) out[cells[0]] = v;
  });
  return out;
}

// 2市場信用倍率シートの空いている E列（信用評価損益率）を、トレーダーズ・ウェブの値で埋める。
// 利用規約が複製を禁じているため、取得は1回につき1ページ・評価損益率の列だけ。
// JPX側の売残・買残・倍率と突き合わせ、合わない行（列ずれ・別週）は埋めない。
// 返り値 { filled: 埋めた件数, note: 失敗・不一致の説明（無ければ ''） }
function fillJpx2MarketPl_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(JPX_MARGIN.SHEET2);
  if (!sh || sh.getLastRow() < 2) return { filled: 0, note: '' };
  let map;
  try {
    const res = UrlFetchApp.fetch(JPX_MARGIN.TW_URL, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (res.getResponseCode() !== 200) return { filled: 0, note: '信用評価損益率: トレーダーズ・ウェブ HTTP ' + res.getResponseCode() };
    map = parseTradersWebMargin_(res.getContentText());
  } catch (e) { return { filled: 0, note: '信用評価損益率: 取得失敗 ' + (e && e.message || e) }; }
  if (!Object.keys(map).length) return { filled: 0, note: '信用評価損益率: 表を読めませんでした（ページ形式が変わった可能性）' };
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  let filled = 0; const bad = [];
  rows.forEach((r, i) => {
    if (r[4] !== '' && r[4] != null) return;
    const key = Utilities.formatDate(new Date(r[0]), 'Asia/Tokyo', 'yyyy/MM/dd');
    const t = map[key];
    if (!t) return;                                   // まだ載っていない週は次回に回す
    const okSell = Math.abs(t.sellOku * 100 - Number(r[1])) <= 100;   // 億円丸め
    const okRatio = Math.abs(t.ratio - Number(r[3])) <= 0.02;
    if (!okSell || !okRatio) { bad.push(key); return; }
    sh.getRange(i + 2, 5).setValue(t.pl);
    filled++;
  });
  return { filled: filled, note: bad.length ? '信用評価損益率: JPXの値と合わず未入力 ' + bad.join(',') : '' };
}
