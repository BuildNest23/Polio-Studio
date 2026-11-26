// Add "Catch Day 2" to a file that has 1 catch-up day, and detect the layout (teams / catch days) from the file.
// Works on an ExcelJS workbook. Formulas are rewritten so everything keeps working.
import { LAST_COL, buildSectionMap, colLetter } from './core';

const REF = /(^|[^A-Za-z0-9_!.$])(\$?[A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g;
const QREF = /((?:'F2A'|F2A)!)(\$?[A-Z]{1,3}\$?)(\d+)(?::(\$?[A-Z]{1,3}\$?)(\d+))?/g;

/** Change the row number of every cell reference in a formula (text inside quotes is left alone). */
export function mapRefs(formula: string, rowFn: (row: number) => number): string {
  return String(formula).split(/("[^"]*")/).map((part, i) => (i % 2 ? part : part.replace(REF, (_m, pre, col, d, row) => pre + col + d + rowFn(parseInt(row, 10))))).join('');
}
/** Same, but only for references written like F2A!E49 (used for formulas in other sheets). */
export function mapQualifiedRefs(formula: string, rowFn: (row: number) => number): string {
  return String(formula).replace(QREF, (_m, sh, c1, r1, c2, r2) =>
    sh + c1 + rowFn(parseInt(r1, 10)) + (c2 ? ':' + c2 + rowFn(parseInt(r2, 10)) : ''));
}

const cloneVal = (v: any) => (v && typeof v === 'object' && !(v instanceof Date) ? { ...v } : v);
const cloneStyle = (s: any) => { try { return s ? JSON.parse(JSON.stringify(s)) : {}; } catch { return {}; } };

/** Read teams / catch days from the titles in column A. Returns null if the file does not look like an F2A file. */
export function detectLayout(sheet: any): { nTeams: number; nCatch: number } | null {
  const last = Math.max(sheet.rowCount || 0, 1);
  let nCatch = 0, hasFinal = false, firstTotal = 0;
  for (let r = 1; r <= last; r++) {
    const v = sheet.getCell(r, 1).value;
    const t = typeof v === 'string' ? v.trim() : v && typeof v === 'object' && typeof (v as any).result === 'string' ? String((v as any).result).trim() : '';
    if (!hasFinal && /^catch\s*day\s*\d+$/i.test(t)) nCatch++;
    else if (/^final report team wise$/i.test(t)) hasFinal = true;
    else if (!firstTotal && /^total$/i.test(t)) firstTotal = r;
  }
  const nTeams = firstTotal - 4;
  if (!hasFinal || !(nTeams >= 1 && nTeams <= 20) || nCatch > 2) return null;
  return { nTeams, nCatch };
}

export function addCatchDay2(wb: any, nTeams: number) {
  const sheet = wb.getWorksheet('F2A');
  const old = buildSectionMap(nTeams, 1), nw = buildSectionMap(nTeams, 2);
  const step = nTeams + 5;
  const sec = (m: any, k: string) => m.sections.find((s: any) => s.key === k);
  const catch1 = sec(old, 'catch1'), insertAt = sec(old, 'finalTeam').titleRow;
  const lo = catch1.titleRow, hi = catch1.titleRow + step - 1;
  const lastRow = Math.max(sheet.rowCount || 0, insertAt);
  const maxC = Math.max(sheet.columnCount || 0, LAST_COL);

  // 1. write shared formulas out one by one (so moving rows cannot break them)
  const normalize = (ws: any, test?: (f: string) => boolean) => {
    for (let r = 1; r <= (ws.rowCount || 0); r++) for (let c = 1; c <= Math.max(ws.columnCount || 0, 1); c++) {
      const cell = ws.getCell(r, c); const f = cell.formula;
      if (f && (!test || test(String(f)))) cell.value = { formula: String(f), result: cell.result };
    }
  };
  normalize(sheet);

  // 2. rows from the insert point down move by `step`: fix every formula that points there
  const shift = (row: number) => (row >= insertAt ? row + step : row);
  for (let r = 1; r <= lastRow; r++) for (let c = 1; c <= maxC; c++) {
    const cell = sheet.getCell(r, c); const f = cell.formula; if (!f) continue;
    const nf = mapRefs(String(f), shift);
    if (nf !== f) cell.value = { formula: nf, result: cell.result };
  }
  for (const ws of wb.worksheets) {
    if (ws === sheet) continue;
    for (let r = 1; r <= (ws.rowCount || 0); r++) for (let c = 1; c <= Math.max(ws.columnCount || 0, 1); c++) {
      const cell = ws.getCell(r, c); const f = cell.formula; if (!f || !/F2A/.test(String(f))) continue;
      const nf = mapQualifiedRefs(String(f), shift);
      if (nf !== f) cell.value = { formula: nf, result: cell.result };
    }
  }

  // 3. merged cells: remember, unmerge those that will move
  const merges: string[] = ((sheet.model && sheet.model.merges) || []).map(String);
  const parse = (s: string) => { const [a, b] = s.split(':'); const m1 = a.match(/^([A-Z]+)(\d+)$/)!, m2 = b.match(/^([A-Z]+)(\d+)$/)!; return { c1: m1[1], r1: +m1[2], c2: m2[1], r2: +m2[2] }; };
  const moving = merges.filter((s) => parse(s).r1 >= insertAt);
  const template = merges.filter((s) => { const p = parse(s); return p.r1 >= lo && p.r1 <= hi; });
  moving.forEach((s) => sheet.unMergeCells(s));

  // 4. move rows down (bottom first)
  for (let r = lastRow; r >= insertAt; r--) {
    for (let c = 1; c <= maxC; c++) {
      const src = sheet.getCell(r, c), dst = sheet.getCell(r + step, c);
      dst.value = cloneVal(src.value); dst.style = cloneStyle(src.style);
    }
    sheet.getRow(r + step).height = sheet.getRow(r).height;
  }
  moving.forEach((s) => { const p = parse(s); sheet.mergeCells(`${p.c1}${p.r1 + step}:${p.c2}${p.r2 + step}`); });

  // 5. build Catch Day 2 in the freed rows (same form as Catch Day 1)
  const c1 = catch1.teamRows, rowIn = (r: number) => r >= lo && r <= hi;
  for (let k = 0; k < step; k++) {
    const tr = lo + k, nr = insertAt + k;
    sheet.getRow(nr).height = sheet.getRow(tr).height;
    for (let c = 1; c <= maxC; c++) {
      const src = sheet.getCell(tr, c), dst = sheet.getCell(nr, c);
      dst.style = cloneStyle(src.style);
      const f = src.formula; const v = src.value;
      let out: any = null;
      if (tr === lo && c === 1) out = 'Catch Day 2';
      else if (f) {
        const ti = c1.indexOf(tr);
        if (ti >= 0 && (c === 2 || c === 3)) out = { formula: colLetter(c === 2 ? 6 : 7) + tr };   // target = still missed of Catch Day 1
        else out = { formula: mapRefs(String(f), (row) => (rowIn(row) ? row + step : row)) };
      } else if (typeof v === 'string') out = v;           // titles, headers, team codes
      dst.value = out;                                      // numbers (inputs) stay empty
    }
  }
  template.forEach((s) => { const p = parse(s); sheet.mergeCells(`${p.c1}${p.r1 + step}:${p.c2}${p.r2 + step}`); });

  // 6. Final Report Team Wise: add Catch Day 2 next to every Catch Day 1 value
  const ft = sec(nw, 'finalTeam');
  for (const r of ft.teamRows) for (let c = 1; c <= LAST_COL; c++) {
    const cell = sheet.getCell(r, c); const f = cell.formula; if (!f || !/SUM\(/.test(String(f))) continue;
    const nf = String(f).replace(REF, (m, pre, col, d, row) => (c1.includes(+row) ? `${pre}${col}${d}${row}, ${col}${d}${+row + step}` : m));
    if (nf !== f) cell.value = { formula: nf, result: cell.result };
  }
  // 7. Final Report Day Wise: the spare row (like the Catch Day 1 row)
  const fd = sec(nw, 'finalDay');
  if (fd.teamRows.length >= 5) {
    const from = fd.teamRows[3], to = fd.teamRows[4];
    const fromLbl = sheet.getCell(from, 1).value;
    const toLbl = sheet.getCell(to, 1);
    if (typeof fromLbl === 'string' && fromLbl.trim() && (toLbl.value === null || toLbl.value === undefined || String(toLbl.value).trim() === '')) {
      toLbl.value = 'Catch Day 2';
    }
    for (let c = 1; c <= maxC; c++) {
      const f = sheet.getCell(from, c).formula; if (!f) continue;
      sheet.getCell(to, c).value = { formula: mapRefs(String(f), (row) => (rowIn(row) ? row + step : row)) };
    }
  }
}
