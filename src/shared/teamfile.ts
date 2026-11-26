// Team file (version 1): read and total it. Used by both apps.
export function normTeamCode(v: unknown): string {
  return String(v || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
}
const nonNeg = (v: unknown): number => {
  const n = Number(v);
  return Number.isFinite(n) && n >= 0 ? n : 0;
};
export function importParse(txt: string) {
  txt = String(txt || '').trim();
  const a = txt.indexOf('{'), b = txt.lastIndexOf('}');
  if (a < 0 || b < a) throw new Error('No data found');
  const d = JSON.parse(txt.slice(a, b + 1));
  if (d.app !== 'polio-team' || !d.team) throw new Error('Not a team file');
  if (d.day !== undefined) {
    const dayNum = Number(d.day);
    if (Number.isFinite(dayNum) && dayNum > 0) d.day = dayNum;
  }
  return d;
}
export function importTotals(d: any) {
  const T: any = { A: 0, B: 0, O: 0, C: 0, mmpv: 0, mmpKids: 0, z: 0, lock: 0, houses: 0 };
  (Array.isArray(d.houses) ? d.houses : []).forEach((h: any) => {
    if (!h) return;
    T.houses++;
    if (h[3]) { T.lock++; return; }
    const a = nonNeg(h[0]);
    const b = Math.min(a, nonNeg(h[1]));
    const o = nonNeg(h[2]);
    const c = Math.min(a - b, nonNeg(h[5]));
    T.A += a; T.B += b; T.O += o; T.C += c;
    if (h[4]) { T.mmpv += b; T.mmpKids += a; }
    if (a === 0) T.z++;
  });
  T.D = nonNeg(d.outside);
  T.S = (Array.isArray(d.school) ? d.school : []).reduce((acc: number, s: any) => acc + nonNeg(s?.vac), 0);
  return T;
}
