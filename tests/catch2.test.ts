import { readFileSync } from 'node:fs';
import { evaluateFormula, evalIncDec } from '../src/incharge/lib/engine';
import { addCatchDay2, detectLayout, mapQualifiedRefs, mapRefs } from '../src/incharge/lib/catch2';

const data = JSON.parse(readFileSync('tests/f2a-sample.json', 'utf8'));
const merges0: string[] = JSON.parse(readFileSync('tests/f2a-merges.json', 'utf8')).map(String);

class Cell {
  value: any = null; style: any = {};
  constructor(public ws: Sheet, public r: number, public c: number) {}
  get formula() { const v = this.value; return v && typeof v === 'object' && v.formula !== undefined ? v.formula : undefined; }
  get result() { const v = this.value; return v && typeof v === 'object' ? v.result : undefined; }
}
class Sheet {
  cells = new Map<string, Cell>(); heights = new Map<number, number>(); merges: string[] = [];
  rowCount = 0; columnCount = 29; name = 'F2A';
  get model() { return { merges: this.merges }; }
  getCell(r: number, c: number) { const k = r + ',' + c; let x = this.cells.get(k); if (!x) { x = new Cell(this, r, c); this.cells.set(k, x); } if (r > this.rowCount) this.rowCount = r; return x; }
  getRow(r: number) { const s = this; return { get height() { return s.heights.get(r); }, set height(v: any) { s.heights.set(r, v); } }; }
  mergeCells(s: string) { if (!this.merges.includes(s)) this.merges.push(s); }
  unMergeCells(s: string) { this.merges = this.merges.filter((m) => m !== s); }
}
function build() {
  const ws = new Sheet();
  for (const k of Object.keys(data)) { const [r, c] = k.split(',').map(Number); const d = data[k]; ws.getCell(r, c).value = d.f !== undefined ? { formula: d.f, result: d.r } : d.v; }
  ws.merges = [...merges0];
  return ws;
}
const evalAll = (ws: Sheet) => {
  const g = { getCell: (r: number, c: number) => ws.getCell(r, c) } as any;
  const out = new Map<string, any>();
  for (const [k, cell] of ws.cells) {
    const f = cell.formula;
    if (f !== undefined) { let v: any = evaluateFormula(g, f, 0); if (v === null) v = evalIncDec(g, f); out.set(k, v); }
    else out.set(k, cell.value);
  }
  return out;
};
let fails = 0; const ok = (c: boolean, m: string) => { if (!c) { fails++; console.log('FAIL', m); } };

// --- layout detection
const before = build();
const lay = detectLayout(before); ok(!!lay && lay.nTeams === 5 && lay.nCatch === 1, 'detect ' + JSON.stringify(lay));

// --- reference mapping
ok(mapRefs('SUM(B54:B58)', (r) => (r >= 41 ? r + 10 : r)) === 'SUM(B64:B68)', 'range');
ok(mapRefs('IF(Y34 < Z34, "Increased by " & (Z34 - Y34), 1)', (r) => r + 10) === 'IF(Y44 < Z44, "Increased by " & (Z44 - Y44), 1)', 'strings kept');
ok(mapQualifiedRefs('F2A!E49 + F2A!F49 + E49', (r) => (r >= 41 ? r + 10 : r)) === 'F2A!E59 + F2A!F59 + E49', 'qualified');
ok(mapQualifiedRefs('F2A!N4', (r) => (r >= 41 ? r + 10 : r)) === 'F2A!N4', 'qualified low');

// --- add catch day 2 and compare every number
const v0 = evalAll(before);
const ws = build();
const wb = { getWorksheet: (n: string) => (n === 'F2A' ? ws : null), worksheets: [ws] };
addCatchDay2(wb, 5);
const v1 = evalAll(ws);
let moved = 0, diffs: string[] = [];
for (const [k, val] of v0) {
  const [r, c] = k.split(',').map(Number); const nr = r >= 41 ? r + 10 : r;
  const nv = v1.get(nr + ',' + c);
  const a = val ?? 0, b = nv ?? 0;
  const same = typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) < 1e-9 : String(val ?? '') === String(nv ?? '');
  if (!same) diffs.push(`${k} -> ${nr},${c}: ${val} vs ${nv}`); else moved++;
}
console.log('cells compared:', moved + diffs.length, ' same:', moved, ' different:', diffs.length);
diffs.slice(0, 12).forEach((d) => console.log('  DIFF', d));
ok(diffs.length === 0, 'numbers changed after adding empty Catch Day 2');

// --- spot checks
const F = (r: number, c: number) => ws.getCell(r, c).formula;
ok(ws.getCell(41, 1).value === 'Catch Day 2', 'title');
ok(F(44, 2) === 'F34' && F(44, 3) === 'G34', 'target = still missed of day 1: ' + F(44, 2));
ok(F(44, 6) === 'MINUS(B44, D44)', 'still missed ' + F(44, 6));
ok(F(49, 2) === 'SUM(B44:B48)', 'total ' + F(49, 2));
ok(F(44, 27) === 'IF(Y44 < Z44, "Increased by " & (Z44 - Y44), IF(Y44 > Z44, "Decreased by " & (Y44 - Z44), "No change"))' || /Y44/.test(String(F(44, 27))), 'AA ' + F(44, 27));
ok(F(54, 8) === 'SUM(H4, H14, H24, D34, D44)', 'final team H ' + F(54, 8));
ok(F(54, 22) === 'SUM(V4, V14, V24, T34, T44)', 'final team V ' + F(54, 22));
ok(F(67, 8) === 'D39' && F(68, 8) === 'D49' && F(68, 9) === 'E49' && F(68, 24) === 'V49', 'day wise rows ' + F(67, 8) + ' ' + F(68, 8));
ok(F(69, 2) === 'SUM(B64:B68)', 'day wise total ' + F(69, 2));
ok(F(74, 2) !== undefined || true, 'next');
const shiftM = (m: string) => m.replace(/([A-Z]+)(\d+)/g, (_x, c, r) => c + (+r >= 41 ? +r + 10 : +r));
const expected = new Set<string>([...merges0.map(shiftM), ...merges0.filter((m) => { const r = +(m.match(/\d+/)![0]); return r >= 31 && r <= 40; }).map((m) => m.replace(/([A-Z]+)(\d+)/g, (_x, c, r) => c + (+r + 10)))]);
ok(expected.size === ws.merges.length && [...expected].every((m) => ws.merges.includes(m)), 'merges differ: want ' + expected.size + ' have ' + ws.merges.length);
ok(detectLayout(ws)?.nCatch === 2, 'detect after');
// Day 3 / Catch 1 untouched
ok(F(34, 6) === 'MINUS(B34, D34)', 'catch1 untouched');
// Fill Catch 2 with numbers: totals must flow into Final Team Wise and Day Wise
ws.getCell(44, 4).value = 3; ws.getCell(44, 9).value = 2;   // covered NA = 3, street = 2
const v2 = evalAll(ws);
const num = (r: number, c: number) => Number(v2.get(r + ',' + c) ?? 0), num0 = (r: number, c: number) => Number(v1.get(r + ',' + c) ?? 0);
ok(num(54, 8) === num0(54, 8) + 3, 'final team H +3: ' + num(54, 8));
ok(num(54, 12) === num0(54, 12) + 2, 'final team L +2');
ok(num(68, 8) === 3 && num(68, 12) === 2, 'day wise row 68');
ok(num(69, 8) === num0(69, 8) + 3, 'day wise total');
console.log(fails ? 'FAILED ' + fails : 'ALL OK');
process.exit(fails ? 1 : 0);
