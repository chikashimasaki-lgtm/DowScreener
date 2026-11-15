/**
 * 保護ルール（新規の買いを一時的に見送る条件）。純関数のみ。2026-11-15 の検証では効果が誤差の範囲だったので本番（src）には入れていない。検証用。
 *
 * 公開の自動売買ソフト freqtrade の Protections（StoplossGuard / CooldownPeriod / MaxDrawdown）と同じ考え方:
 *   - 連続損切りガード: 直近 GUARD_LOOKBACK 営業日に損切りが GUARD_STOPS 回以上 → 最後の損切りから GUARD_PAUSE 営業日は新規を見送る
 *   - クールダウン: 損切りになった銘柄は COOLDOWN_DAYS 営業日のあいだ再び買わない
 *   - ドローダウン停止: 直近 DD_LOOKBACK 営業日に決済した取引の累計R の、山からの落ち込みが DD_LIMIT_R 以上 → 最後の決済から DD_PAUSE 営業日は見送る
 * 値が 0 の項目は無効。日付はすべて「営業日の通し番号」（整数）で渡す。判断日 idx までに決済が済んだ取引だけを見る（先読みしない）。
 *
 * closed: [{ code, exitIdx, R, why }]  why = 'stop' | 'target' | 'time'。R = 損益 ÷ 1回のリスク額
 */
const PG = {
  GUARD_STOPS: 0, GUARD_LOOKBACK: 0, GUARD_PAUSE: 0,
  COOLDOWN_DAYS: 0,
  DD_LIMIT_R: 0, DD_LOOKBACK: 0, DD_PAUSE: 0,
};

function pgState_(closed, idx, cfg) {
  const c = cfg || PG;
  const done = (closed || []).filter(t => t && t.exitIdx <= idx);
  const reasons = [];
  let pauseUntil = -1;
  if (c.GUARD_STOPS > 0) {
    const stops = done.filter(t => t.why === 'stop' && t.R < 0 && t.exitIdx > idx - c.GUARD_LOOKBACK);
    if (stops.length >= c.GUARD_STOPS) {
      const until = Math.max.apply(null, stops.map(t => t.exitIdx)) + c.GUARD_PAUSE;
      if (until > idx) { pauseUntil = Math.max(pauseUntil, until); reasons.push('直近' + c.GUARD_LOOKBACK + '営業日に損切り' + stops.length + '回'); }
    }
  }
  if (c.DD_LIMIT_R > 0) {
    const win = done.filter(t => t.exitIdx > idx - c.DD_LOOKBACK).sort((a, b) => a.exitIdx - b.exitIdx);
    let cum = 0, peak = 0, dd = 0;
    win.forEach(t => { cum += t.R; if (cum > peak) peak = cum; if (peak - cum > dd) dd = peak - cum; });
    if (win.length && dd >= c.DD_LIMIT_R) {
      const until = win[win.length - 1].exitIdx + c.DD_PAUSE;
      if (until > idx) { pauseUntil = Math.max(pauseUntil, until); reasons.push('直近' + c.DD_LOOKBACK + '営業日の落ち込み ' + (Math.round(dd * 10) / 10) + 'R'); }
    }
  }
  const locked = {};
  if (c.COOLDOWN_DAYS > 0) {
    done.forEach(t => {
      if (t.why !== 'stop') return;
      const until = t.exitIdx + c.COOLDOWN_DAYS;
      if (until > idx && !(locked[t.code] >= until)) locked[t.code] = until;
    });
  }
  return { pause: pauseUntil > idx, pauseUntil: pauseUntil, reasons: reasons, locked: locked };
}
