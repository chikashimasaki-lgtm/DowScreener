/**
 * 先読みの検査（freqtrade の lookahead-analysis と同じ考え方）。
 * 判断日 e より後の日足を切り捨てても、dfSetup_ の結果（確度・損切り・利確・スコア）が変わらないことを確かめる。
 * 変わるなら、バックテストは未来の値を見ていることになり、成績が実際より良く出る。
 *
 *   BARS_DIR=<日足> node tools/lookahead_check.js [--stocks 80] [--step 1]
 */
const fs = require('fs'), path = require('path');
const args = process.argv.slice(2), arg = (k, d) => { const i = args.indexOf(k); return i >= 0 ? Number(args[i + 1]) : d; };
const BARS_DIR = process.env.BARS_DIR; if (!BARS_DIR) { console.error('BARS_DIR を指定してください'); process.exit(1); }
const src = f => fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8');
const sb = { Logger: { log: () => {} }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) }, SpreadsheetApp: {}, UrlFetchApp: {},
  Utilities: { formatDate: d => d.toISOString().slice(0, 10), sleep: () => {} }, ScriptApp: {}, Session: {}, DriveApp: {}, MailApp: {}, GmailApp: {} };
const M = new Function(...Object.keys(sb), `${src('StockCode.js')}\n${src('Code.js')}\n${src('DowFib.js')}\nreturn { DF, dfSetup_, dfRawSwings_, parseYahooBars_ };`)(...Object.values(sb));
const files = fs.readdirSync(BARS_DIR).filter(f => /^\w+\.T\.json$/.test(f)).sort();
const N = arg('--stocks', 80), STEP = arg('--step', 1), TAIL = 300;   // 切り捨て後も十分な長さを残す
const pickEvery = Math.max(1, Math.floor(files.length / N));
const key = s => s ? [s.grade, s.stop, s.target, s.score].join('|') : 'なし';
let checked = 0, setups = 0, diff = 0; const samples = [];
for (let k = 0; k < files.length; k += pickEvery) {
  const bars = M.parseYahooBars_({ getResponseCode: () => 200, getContentText: () => fs.readFileSync(path.join(BARS_DIR, files[k]), 'utf8') });
  if (bars.length < 600) continue;
  const rawP = M.dfRawSwings_(bars, M.DF.PRIMARY_W), rawM = M.dfRawSwings_(bars, M.DF.MINOR_W);
  for (let e = bars.length - TAIL; e < bars.length - 1; e += STEP) {
    const full = M.dfSetup_(bars, e, rawP, rawM);
    const cut = bars.slice(0, e + 1);
    const trunc = M.dfSetup_(cut, e, M.dfRawSwings_(cut, M.DF.PRIMARY_W), M.dfRawSwings_(cut, M.DF.MINOR_W));
    checked++; if (full || trunc) setups++;
    if (key(full) !== key(trunc)) { diff++; if (samples.length < 5) samples.push(files[k] + ' e=' + e + ' 全体=' + key(full) + ' 切り捨て=' + key(trunc)); }
  }
}
console.log(`検査 ${checked} 日（うち候補あり ${setups}）/ 食い違い ${diff}`);
samples.forEach(s => console.log('  ' + s));
process.exit(diff ? 1 : 0);
