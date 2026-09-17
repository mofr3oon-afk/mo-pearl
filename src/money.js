// Round to the nearest piastre (two decimal places), rather than truncate.
// EPSILON scaled with magnitude protects common binary floating-point halfway cases.
export function roundMoney(value) {
  const n = Number(String(value ?? 0).replace(/,/g, '').trim() || 0);
  if (!Number.isFinite(n)) return NaN;
  const absolute = Math.abs(n);
  return Math.sign(n) * Math.round((absolute + Number.EPSILON * absolute) * 100) / 100;
}
export function moneyCents(value) { return Math.round(roundMoney(value) * 100); }
