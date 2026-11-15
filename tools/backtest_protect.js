/**
 * 保護ルール（tools/protect.js）のバックテスト。候補は backtest_dowfib.js が DUMP で書き出したものを使う。
 *
 *   BARS_DIR=<日足> DUMP=cands.json node tools/backtest_dowfib.js
 *   node tools/backtest_protect.js cands.json [--grid | --cost]
 *
 * 日付順に1日ずつ進め、その日までに決済が済んだ取引だけを pgState_ に渡す（先読みしない）。
 * 保有中の銘柄は再び買わない。1日 TOP_N 件まで。尺度は R（損益 ÷ 1回のリスク額。本番は1回3万円で一定）。
 * R は台帳と同じ尺度で、手数料・スリッページを含まない。--cost は往復コストを引いた R を並べる
 * （コストの R 換算 = コスト率 ÷ 損切り幅の率。損切り幅が狭い取引ほど重くなる）。
 */
const fs = require('fs'), path = require('path');
const { pgState_ } = new Function(fs.readFileSync(path.join(__dirname, 'protect.js'), 'utf8') + '\nreturn { pgState_ };')();
const D = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const idxOf = {}; D.days.forEach((d, i) => { idxOf[d] = i; });
const byDay = {}; D.cands.forEach(c => (byDay[c.d] = byDay[c.d] || []).push(c));
Object.keys(byDay).forEach(d => byDay[d].sort((a, b) => b.score - a.score));
const PERIODS = [['A 〜2021', d => d < '2022-01-01'], ['B 2022〜2025-09', d => d >= '2022-01-01' && d < '2025-10-01'], ['C 2025-10〜', d => d >= '2025-10-01'], ['全期間', () => true]];

function run(cfg) {
  const trades = [], openUntil = {};
  let skippedDays = 0, skippedCool = 0;
  D.days.forEach((d, i) => {
    const cs = byDay[d]; if (!cs) return;
    const recent = trades.filter(t => t.exitIdx <= i && t.exitIdx > i - 260);
    const st = pgState_(recent, i, cfg);
    if (st.pause) { skippedDays++; return; }
    let n = 0;
    for (const c of cs) {
      if (n >= D.topN) break;
      if (openUntil[c.code] > i) continue;
      if (st.locked[c.code]) { skippedCool++; continue; }
      n++;
      if (!c.t) continue;
      const exitIdx = i + 1 + c.t.d;
      openUntil[c.code] = exitIdx;
      trades.push({ d, code: c.code, exitIdx, R: c.t.R - costR_(c.t, cfg.COST || 0), why: c.t.why });
    }
  });
  return { trades, skippedDays, skippedCool };
}
// 往復コスト（率）を R に直す。損切り幅の率 = (r + 書き出し時のコスト) / R。R が 0 に近い取引は全体の中央値で代用する。
const DUMP_COST = 0.002;
const riskFrac_ = t => Math.abs(t.R) > 0.05 ? (t.r + DUMP_COST) / t.R : NaN;
const MED_RISK = (() => { const a = D.cands.filter(c => c.t).map(c => riskFrac_(c.t)).filter(x => x > 0).sort((x, y) => x - y); return a[a.length >> 1]; })();
function costR_(t, cost) { if (!cost) return 0; const f = riskFrac_(t); return cost / (f > 0 ? f : MED_RISK); }
function stats(trades) {
  const s = trades.slice().sort((a, b) => a.exitIdx - b.exitIdx);
  let cum = 0, peak = 0, dd = 0, gw = 0, gl = 0, w = 0;
  s.forEach(t => { cum += t.R; if (cum > peak) peak = cum; if (peak - cum > dd) dd = peak - cum; if (t.R > 0) { gw += t.R; w++; } else gl -= t.R; });
  return { n: s.length, sum: cum, mean: s.length ? cum / s.length : 0, pf: gl ? gw / gl : NaN, dd, win: s.length ? w / s.length : 0 };
}
function line(label, cfg) {
  const r = run(cfg);
  const cols = PERIODS.map(([, f]) => { const s = stats(r.trades.filter(t => f(t.d))); return `n=${String(s.n).padStart(4)} 計${s.sum.toFixed(0).padStart(5)}R 平均${s.mean.toFixed(3)} PF${s.pf.toFixed(2)} DD${s.dd.toFixed(0).padStart(3)}R`; });
  console.log(label.padEnd(30) + cols.join(' | ') + `  見送り日=${r.skippedDays} cool=${r.skippedCool}`);
}
const Z = { GUARD_STOPS: 0, GUARD_LOOKBACK: 0, GUARD_PAUSE: 0, COOLDOWN_DAYS: 0, DD_LIMIT_R: 0, DD_LOOKBACK: 0, DD_PAUSE: 0 };
console.log(''.padEnd(30) + PERIODS.map(p => p[0].padEnd(40)).join(' | '));
line('保護なし', Z);
if (process.argv.includes('--cost')) {
  console.log('損切り幅の中央値 ' + (MED_RISK * 100).toFixed(2) + '%');
  for (const c of [0.001, 0.002, 0.004, 0.006, 0.01]) line('往復コスト ' + (c * 100).toFixed(1) + '%', Object.assign({}, Z, { COST: c }));
} else if (process.argv.includes('--grid')) {
  for (const cd of [5, 10, 20, 40]) line('cooldown ' + cd, Object.assign({}, Z, { COOLDOWN_DAYS: cd }));
  for (const [n, lb, p] of [[4, 5, 3], [4, 5, 5], [6, 5, 5], [6, 10, 5], [8, 10, 5], [8, 10, 10], [10, 10, 10], [12, 20, 10]]) line(`guard ${n}回/${lb}日→${p}日`, Object.assign({}, Z, { GUARD_STOPS: n, GUARD_LOOKBACK: lb, GUARD_PAUSE: p }));
  for (const [r, lb, p] of [[5, 20, 5], [8, 20, 5], [8, 20, 10], [10, 40, 10], [15, 40, 10], [15, 60, 20]]) line(`dd ${r}R/${lb}日→${p}日`, Object.assign({}, Z, { DD_LIMIT_R: r, DD_LOOKBACK: lb, DD_PAUSE: p }));
} else {
  const PG = new Function(fs.readFileSync(path.join(__dirname, 'protect.js'), 'utf8') + '\nreturn PG;')();
  line('現行の PG', PG);
}
