// Everything that touches the Excel workbook (ExcelJS). Ported from the original Polio.html.
import { addCatchDay2 as addCatchDay2Core, detectLayout } from './catch2';
import {
  COLS, LAST_COL, MAX_COLS, COL_WIDTHS, CATCH_MAP, IFTXT, bodyLastCol, buildSectionMap, colLetter, colLetterToNum,
  groupsFor, sectionKind, sectionTitleText, todayISO, type Section,
} from './core';
import { cellDisplayValue, evalIncDec, evaluateFormula, formulaDisplay, getCellNumericValue } from './engine';
import { S, allSections, applySectionMap, emit, getMerge, idbDelete, idbGet, idbSet, isDirty, mainSections, nextSections, toast, uiConfirm, uiPrompt, uiTemplateForm } from './state';

async function XL(): Promise<any> { return (await import('exceljs')).default; }

/* ---------- small helpers ---------- */
const _THIN = { style: 'thin', color: { argb: 'FFDAD3C3' } };
const _BORDER = { top: _THIN, bottom: _THIN, left: _THIN, right: _THIN };
const _fill = (argb: string) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
export const getSheet = () => (S.workbook ? S.workbook.getWorksheet('F2A') : null);

function buildMergeMap(sheet: any) {
  const map: Record<string, any> = {};
  const merges = (sheet.model && sheet.model.merges) || [];
  for (const range of merges) {
    const parts = String(range).split(':');
    if (parts.length !== 2) continue;
    const m1 = parts[0].match(/^([A-Z]+)(\d+)$/i), m2 = parts[1].match(/^([A-Z]+)(\d+)$/i);
    if (!m1 || !m2) continue;
    const c1 = colLetterToNum(m1[1].toUpperCase()), r1 = parseInt(m1[2], 10);
    const c2 = colLetterToNum(m2[1].toUpperCase()), r2 = parseInt(m2[2], 10);
    const info = { r1, c1, r2, c2 };
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) map[r + ',' + c] = { ...info, isMaster: r === r1 && c === c1 };
  }
  return map;
}

/* ---------- campaign meta (start date) + file layout params ---------- */
const META_SHEET = '_F2A_Meta';
const metaKey = () => 'polioMeta:' + (S.fileName || '_default_');
const paramsKey = (n: string) => 'polioTemplateParams:' + n;
function lsGet(k: string) { try { return localStorage.getItem(k); } catch { return null; } }
function lsSet(k: string, v: string) { try { localStorage.setItem(k, v); } catch { /* ignore */ } }
function saveTemplateParams(name: string, p: any) { lsSet(paramsKey(name), JSON.stringify(p)); }
function loadTemplateParams(name: string) { try { const r = lsGet(paramsKey(name)); return r ? JSON.parse(r) : null; } catch { return null; } }
function readSheetMeta(): any {
  try {
    const s = S.workbook.getWorksheet(META_SHEET); if (!s) return null;
    const o: any = {}; s.eachRow((r: any) => { const k = r.getCell(1).value; if (k) o[k] = r.getCell(2).value; });
    return o;
  } catch { return null; }
}
function writeSheetMeta(obj: any) {
  try {
    let s = S.workbook.getWorksheet(META_SHEET);
    if (!s) { s = S.workbook.addWorksheet(META_SHEET); s.state = 'hidden'; }
    Object.keys(obj).forEach((k, i) => { s.getCell(i + 1, 1).value = k; s.getCell(i + 1, 2).value = obj[k]; });
  } catch { /* ignore */ }
}
export function setStartDate(v: string) {
  S.startDate = v;
  lsSet(metaKey(), JSON.stringify({ startDate: v }));
  if (S.workbook) { writeSheetMeta({ nTeams: S.params.nTeams, nCatch: S.params.nCatch, startDate: v }); scheduleAutosave(); }
  emit();
}

/* ---------- Area Incharge name (shown in the header, kept on this phone) ---------- */
export function setIncharge(v: string) {
  S.incharge = String(v || '').trim();
  try { localStorage.setItem('polioIncharge', S.incharge); } catch { /* ignore */ }
  emit();
}
async function askIncharge(step: string) {
  const v = await uiPrompt('Area Incharge name (shown at the top of the app). You can change it later in More.', S.incharge, { title: step, yesText: 'OK' });
  if (v !== null && v.trim() !== '') setIncharge(v);
}

/* ---------- dates ---------- */
function parseLocalDate(str: string) {
  const m = String(str || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return isNaN(d.getTime()) ? null : d;
}
export function sectionDateFull(key: string): string {
  const s = mainSections().find((x) => x.key === key);
  if (!s || s.dayOffset === undefined || !S.startDate) return '';
  const base = parseLocalDate(S.startDate); if (!base) return '';
  const d = new Date(base); d.setDate(d.getDate() + s.dayOffset);
  return d.toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
}

/* ---------- load / save ---------- */
export async function loadWorkbookFromBuffer(buffer: ArrayBuffer, name: string, quiet = false) {
  const ExcelJS = await XL();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheet = wb.getWorksheet('F2A');
  if (!sheet) { toast("No 'F2A' sheet found in this file.", 'error'); return false; }
  S.workbook = wb; S.originalBuffer = buffer; S.fileName = name || 'workbook.xlsx';
  S.edited = new Set(); S.bulk = false;
  S.mergeMap = buildMergeMap(sheet);
  try { wb.calcProperties.fullCalcOnLoad = true; } catch { /* ignore */ }

  const sm = readSheetMeta();
  // read teams / catch days from the file itself (so a wrong answer can never break the layout)
  const det = detectLayout(sheet);
  let params: any = det || (sm && sm.nTeams ? sm : loadTemplateParams(S.fileName) || loadTemplateParams(S.fileName.replace(/( - (Updated|Backup).*)?\.xlsx$/i, '') + '.xlsx'));
  if (det && !quiet && !S.incharge) await askIncharge('Welcome');
  if (!(params && params.nTeams)) {
    let n = parseInt((await uiPrompt('How many teams are in this file?', '5', { title: 'File layout', yesText: 'Next' })) || '', 10);
    if (!(n >= 1 && n <= 20)) n = 5;
    let k = parseInt((await uiPrompt('How many catch-up days (0, 1 or 2)?', '1', { title: 'File layout', yesText: 'OK' })) || '', 10);
    if (!(k >= 0 && k <= 2)) k = 1;
    await askIncharge('File layout');
    params = { nTeams: n, nCatch: k };
    saveTemplateParams(S.fileName, params);
  }
  applySectionMap(Number(params.nTeams), Number(params.nCatch) || 0);
  S.section = mainSections()[0]?.key || 'day1';
  S.nextSection = nextSections()[0]?.key || 'next1';
  let sd = ''; try { const r = lsGet(metaKey()); sd = r ? JSON.parse(r).startDate || '' : ''; } catch { /* ignore */ }
  S.startDate = sm && sm.startDate ? String(sm.startDate) : sd;
  writeSheetMeta({ nTeams: S.params.nTeams, nCatch: S.params.nCatch, startDate: S.startDate || '' });
  S.epoch++; emit();
  if (!quiet) toast('Loaded: ' + S.fileName, 'success');
  return true;
}
export async function openFile(file: File) {
  if (!/\.xlsx$/i.test(file.name)) { toast('Please choose an .xlsx file', 'error'); return; }
  try {
    toast('Opening ' + file.name + '…', 'info');
    const buf = await file.arrayBuffer();
    const ok = await loadWorkbookFromBuffer(buf, file.name);
    if (ok) { try { await idbSet('polioCurrentFile', { name: file.name, buffer: buf }); await idbSet('polioOriginalFile', { name: file.name, buffer: buf }); } catch { /* ignore */ } }
  } catch (err: any) { console.error(err); toast('Could not load file: ' + err.message, 'error'); }
}
export async function restoreSaved() {
  try {
    const saved = await idbGet('polioCurrentFile');
    if (saved && saved.buffer) {
      const ok = await loadWorkbookFromBuffer(saved.buffer, saved.name, true);
      if (ok && saved.dirty) {
        const o = await idbGet('polioOriginalFile');
        if (o && o.buffer && o.name === saved.name) S.originalBuffer = o.buffer;
        S.bulk = true; emit();
        toast('Restored your unsaved work', 'info');
      }
    }
  } catch (e) { console.warn('Could not restore', e); }
}
export async function discardChanges() {
  if (!S.originalBuffer) return;
  if (!(await uiConfirm('Discard all unsaved changes and reload the original file?', { title: 'Discard changes', yesText: 'Discard', danger: true }))) return;
  await loadWorkbookFromBuffer(S.originalBuffer, S.fileName, true);
  toast('Changes discarded.', 'info');
}
export async function changeFile() {
  if (!(await uiConfirm('Load a different file?\n\nUnsaved changes will be lost.', { title: 'Change file', yesText: 'Change', danger: true }))) return;
  S.workbook = null; S.originalBuffer = null; S.fileName = ''; S.mergeMap = null; S.edited = new Set(); S.bulk = false; S.startDate = ''; S.moreSub = '';
  applySectionMap(5, 1); S.section = 'day1'; S.nextSection = 'next1';
  try { await idbDelete('polioCurrentFile'); } catch { /* ignore */ }
  S.epoch++; emit();
}

let autosaveTimer: any;
export function scheduleAutosave() { clearTimeout(autosaveTimer); autosaveTimer = setTimeout(autosaveNow, 1500); }
export async function autosaveNow() {
  if (!S.workbook || !isDirty()) return;
  try { const buf = await serializeWorkbook(); await idbSet('polioCurrentFile', { name: S.fileName, buffer: buf, dirty: true }); } catch (e) { console.warn('Autosave failed', e); }
}
if (typeof document !== 'undefined') {
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') autosaveNow(); });
  window.addEventListener('pagehide', () => { autosaveNow(); });
}

/** Writes calculated results next to every formula, so the file also shows numbers in viewers that do not recalculate. */
function fillCachedResults() {
  const sheet = getSheet(); if (!sheet) return;
  sheet.eachRow({ includeEmpty: false }, (row: any) => row.eachCell({ includeEmpty: false }, (cell: any) => {
    const f0 = cell.formula; if (!f0) return;
    // MINUS(a,b) only exists in Google Sheets. Save it as a-b so real Excel also works (same answer).
    const mm = String(f0).match(/^MINUS\(\s*([A-Z]+\d+)\s*,\s*([A-Z]+\d+)\s*\)$/i);
    const f = mm ? mm[1].toUpperCase() + '-' + mm[2].toUpperCase() : f0;
    try {
      let res: any = evaluateFormula(sheet, f, 0);
      if (res === null || res === undefined || (typeof res === 'number' && !isFinite(res))) res = evalIncDec(sheet, f);
      if (res !== null && res !== undefined) cell.value = { formula: f, result: res };
      else if (mm) cell.value = { formula: f };
    } catch { if (mm) { try { cell.value = { formula: f }; } catch { /* ignore */ } } }
  }));
}
/** Make every F2A column wide enough for its header/text so nothing is cut off. Only widens, never narrows. */
function fitColumnWidths() {
  const sheet = getSheet(); if (!sheet) return;
  const widthOf = (c: number) => { const w = sheet.getColumn(c).width; return typeof w === 'number' && w > 0 ? w : 12.63; };
  const need: Record<number, number> = {};
  const mneed: { c1: number; c2: number; n: number }[] = [];
  const last = Math.max(sheet.rowCount || 0, 1);
  for (let r = 1; r <= last; r++) {
    for (let c = 1; c <= LAST_COL; c++) {
      const cell = sheet.getCell(r, c);
      let txt: any = cell.formula ? formulaDisplay(sheet, cell) : cellDisplayValue(cell);
      if (typeof txt !== 'string' || !txt.trim()) continue;
      const n = Math.max(...txt.split('\n').map((x: string) => x.length)) * 1.05 + 2.5;
      const m = getMerge(r, c);
      if (m && m.c2 > m.c1) { if (m.isMaster) mneed.push({ c1: m.c1, c2: m.c2, n }); }
      else need[c] = Math.max(need[c] || 0, n);
    }
  }
  const now: Record<number, number> = {};
  for (let c = 1; c <= LAST_COL; c++) now[c] = Math.max(widthOf(c), Math.min(need[c] || 0, 48));
  for (const m of mneed) {
    let tot = 0; for (let c = m.c1; c <= m.c2; c++) tot += now[c] ?? widthOf(c);
    if (tot < m.n) { const add = (m.n - tot) / (m.c2 - m.c1 + 1); for (let c = m.c1; c <= m.c2; c++) now[c] = (now[c] ?? widthOf(c)) + add; }
  }
  for (let c = 1; c <= LAST_COL; c++) { const w = Math.round(now[c] * 100) / 100; if (w > widthOf(c)) sheet.getColumn(c).width = w; }
}
export async function serializeWorkbook(): Promise<ArrayBuffer> {
  try { S.workbook.calcProperties.fullCalcOnLoad = true; } catch { /* ignore */ }
  fillCachedResults();
  try { fitColumnWidths(); } catch (e) { console.error(e); }
  return await S.workbook.xlsx.writeBuffer();
}
export async function downloadWorkbook(suffix?: string) {
  if (!S.workbook) { toast('No workbook loaded', 'error'); return; }
  try {
    toast('Preparing file…', 'info');
    const buf = await serializeWorkbook();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${S.fileName.replace(/\.xlsx$/i, '')} - ${suffix || 'Updated'}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    if (suffix) { toast('Backup downloaded.', 'success'); return; }
    S.originalBuffer = buf;
    try { await idbSet('polioCurrentFile', { name: S.fileName, buffer: buf, dirty: false }); await idbSet('polioOriginalFile', { name: S.fileName, buffer: buf }); } catch { /* ignore */ }
    S.edited = new Set(); S.bulk = false; S.epoch++; emit();
    toast('Downloaded. Open in Excel to see recalculated totals.', 'success');
  } catch (err: any) { console.error(err); toast('Could not generate file: ' + err.message, 'error'); }
}

/* ---------- editing ---------- */
/** Called on every keystroke in an input cell. Returns false when the cell cannot be edited. */
export function editCell(row: number, col: number, raw: string): boolean {
  const sheet = getSheet(); if (!sheet) return false;
  const cell = sheet.getCell(row, col);
  if (cell.formula) return false;
  if (getMerge(row, col)) { toast('Merged cells cannot be edited.', 'info'); return false; }
  if (raw === '') cell.value = null;
  else { const n = Number(raw); cell.value = !isNaN(n) && raw.trim() !== '' ? n : raw; }
  S.edited.add(row + ',' + col);
  emit(); scheduleAutosave();
  return true;
}

/** Team codes: the same code is written in every section (except Final Day Wise, which lists days). */
export function teamCodes(): string[] {
  const sheet = getSheet(); const first = mainSections()[0];
  if (!sheet || !first) return [];
  return first.teamRows.map((r) => String(cellDisplayValue(sheet.getCell(r, 1)) || '').trim());
}
export function setTeamCode(idx: number, raw: string): boolean {
  const sheet = getSheet(); if (!sheet) return false;
  const code = String(raw || '').trim();
  if (code) {
    const others = teamCodes();
    for (let i = 0; i < others.length; i++) if (i !== idx && others[i] && others[i].toLowerCase() === code.toLowerCase()) { toast('Another team already has that code.', 'error'); return false; }
  }
  for (const sec of allSections().filter((s) => s.key !== 'finalDay')) {
    const row = sec.teamRows[idx]; if (row === undefined) continue;
    const cell = sheet.getCell(row, 1); if (cell.formula) continue;
    cell.value = code || null; S.edited.add(row + ',1');
  }
  S.bulk = true; S.epoch++; emit(); scheduleAutosave();
  toast('Team code updated in all sections.', 'success');
  return true;
}

/* ---------- section table model (used by the screen and by print) ---------- */
export function detectLastUsedColumn(sheet: any, allRows: number[]) {
  let maxCol = 0;
  for (const r of allRows) for (let c = 1; c <= MAX_COLS; c++) {
    const v = sheet.getCell(r, c).value; if (v !== null && v !== undefined && v !== '') maxCol = Math.max(maxCol, c);
  }
  const merges = (sheet.model && sheet.model.merges) || [];
  for (const range of merges) {
    const parts = String(range).split(':'); if (parts.length !== 2) continue;
    const m1 = parts[0].match(/^([A-Z]+)(\d+)$/i), m2 = parts[1].match(/^([A-Z]+)(\d+)$/i); if (!m1 || !m2) continue;
    const sr = parseInt(m1[2], 10), er = parseInt(m2[2], 10), ec = colLetterToNum(m2[1].toUpperCase());
    if (allRows.some((r) => r >= sr && r <= er)) maxCol = Math.max(maxCol, ec);
  }
  return Math.min(maxCol || 1, MAX_COLS);
}
export function childCols(sheet: any, section: Section): number | null {
  const r = section.headerRows[0];
  for (let c = 1; c <= 45; c++) {
    let v = sheet.getCell(r, c).value;
    if (v && typeof v === 'object') v = v.richText ? v.richText.map((x: any) => x.text).join('') : v.result != null ? v.result : '';
    if (String(v == null ? '' : v).trim().toLowerCase() === 'children count') return c;
  }
  return null;
}

export interface MCell { tag: 'th' | 'td'; cls: string; colspan: number; rowspan: number; row: number; col: number; display: string | number; input?: boolean }
export interface MRow { kind: 'header' | 'team' | 'total'; idx: number; cells: MCell[] }

export function buildSectionModel(section: Section): MRow[] {
  const sheet = getSheet(); if (!sheet) return [];
  const allRows = [...section.headerRows, ...section.teamRows, ...(section.totalRow ? [section.totalRow] : [])];
  let lastCol = detectLastUsedColumn(sheet, allRows);
  if (/^(day[123]|catch\d+)$/.test(section.key)) { const cc = childCols(sheet, section); if (cc && cc <= lastCol) lastCol = cc - 1; }
  const skip: Record<string, boolean> = {};
  const maxR = Math.max(...allRows), minR = Math.min(...allRows);
  const rows: MRow[] = [];
  const mk = (kind: MRow['kind'], r: number, idx: number) => {
    const cells: MCell[] = [];
    let c = 1;
    while (c <= lastCol) {
      const key = r + ',' + c;
      if (skip[key]) { c++; continue; }
      const merge = getMerge(r, c);
      const cell = sheet.getCell(r, c);
      const v = cellDisplayValue(cell);
      let colspan = 1, rowspan = 1;
      if (merge) {
        if (!merge.isMaster) { skip[key] = true; c++; continue; }
        const effR2 = Math.min(merge.r2, maxR), effR1 = Math.max(merge.r1, minR);
        const effC2 = Math.min(merge.c2, lastCol), effC1 = Math.max(merge.c1, 1);
        colspan = effC2 - c + 1; rowspan = effR2 - r + 1;
        for (let cc = c + 1; cc <= effC2; cc++) skip[r + ',' + cc] = true;
        for (let rr = r + 1; rr <= effR2; rr++) for (let cc = effC1; cc <= effC2; cc++) skip[rr + ',' + cc] = true;
        void effR1;
      }
      const isFirst = c === 1, isFormula = !!cell.formula;
      const grayCol = /^next[123]$/.test(section.key) && c === 3;
      const base = { colspan, rowspan, row: r, col: c };
      if (kind === 'header') {
        cells.push({ ...base, tag: 'th', cls: (isFirst ? 'teamh ' : '') + (grayCol && colspan === 1 ? 'grayh' : ''), display: v });
      } else if (kind === 'team') {
        if (isFirst) cells.push({ ...base, tag: 'td', cls: 'team', display: v });
        else if (grayCol) cells.push({ ...base, tag: 'td', cls: 'gray', display: '—' });
        else if (isFormula) {
          const d = formulaDisplay(sheet, cell);
          const isText = typeof d === 'string' && !/^-?\d/.test(d);
          cells.push({ ...base, tag: 'td', cls: isText ? 'txt' : 'calc', display: d });
        } else if (merge && (colspan > 1 || rowspan > 1)) cells.push({ ...base, tag: 'td', cls: 'merged', display: v });
        else cells.push({ ...base, tag: 'td', cls: S.edited.has(r + ',' + c) ? 'changed' : '', display: cell.value === null || cell.value === undefined ? '' : v, input: true });
      } else {
        if (isFirst) cells.push({ ...base, tag: 'td', cls: 'team totalname', display: v });
        else if (isFormula) {
          const d = formulaDisplay(sheet, cell);
          const isText = typeof d === 'string' && !/^-?\d/.test(d);
          cells.push({ ...base, tag: 'td', cls: isText ? 'txt' : 'calc', display: d });
        } else cells.push({ ...base, tag: 'td', cls: '', display: v });
      }
      c += colspan;
    }
    rows.push({ kind, idx, cells });
  };
  section.headerRows.forEach((r, i) => mk('header', r, i));
  section.teamRows.forEach((r, i) => mk('team', r, i));
  if (section.totalRow) mk('total', section.totalRow, 0);
  return rows;
}

const esc = (s: any) => String(s ?? '').replace(/[&<>"']/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' } as any)[m]);
export function printSection(section: Section) {
  const rows = buildSectionModel(section);
  const body = rows.map((r) => '<tr' + (r.kind === 'total' ? ' class="tot"' : '') + '>' + r.cells.map((c) => {
    const a = (c.colspan > 1 ? ` colspan="${c.colspan}"` : '') + (c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : '');
    return `<${c.tag}${a}>${esc(c.display)}</${c.tag}>`;
  }).join('') + '</tr>').join('');
  const date = sectionDateFull(section.key);
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${esc(section.label)}</title><style>
  @page{size:A4 landscape;margin:8mm} body{font-family:Arial,Helvetica,sans-serif;color:#000;background:#fff}
  h1{font-size:14pt;text-align:center;margin:0 0 2mm} .d{text-align:center;font-size:11pt;font-weight:700;margin:0 0 4mm}
  table{width:100%;border-collapse:collapse} th,td{border:1px solid #000;padding:1.5mm 1mm;font-size:7pt;text-align:center;vertical-align:middle}
  th,tr.tot td{background:#e5e7eb;font-weight:800}</style></head><body>
  <h1>${esc(section.label)} — F2A Field Report</h1>${date ? `<div class="d">${esc(date)}</div>` : ''}<table>${body}</table></body></html>`;
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  iframe.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
  document.body.appendChild(iframe);
  const doc = iframe.contentDocument || iframe.contentWindow?.document;
  const win = iframe.contentWindow;
  if (!doc || !win) { iframe.remove(); toast('Could not open print view.', 'error'); return; }
  doc.open(); doc.write(html); doc.close();
  const cleanup = () => { if (iframe.isConnected) iframe.remove(); };
  win.addEventListener('afterprint', () => setTimeout(cleanup, 250), { once: true });
  setTimeout(() => {
    try { win.focus(); win.print(); } catch { cleanup(); }
    setTimeout(cleanup, 60000);
  }, 250);
}

/* ---------- Children tab data ---------- */
export function childDiff(p: number | null, n: number | null) {
  if (p === null || n === null) return { t: '—', c: '' };
  const d = Math.round((n - p) * 100) / 100;
  return d > 0 ? { t: 'Increased by ' + d, c: 'up' } : d < 0 ? { t: 'Decreased by ' + -d, c: 'down' } : { t: 'No change', c: '' };
}
function childVal(sheet: any, row: number, col: number) {
  const cell = sheet.getCell(row, col);
  if (cell.formula) { const v = evaluateFormula(sheet, cell.formula, 0); if (v !== null && isFinite(v)) return v; }
  const v = getCellNumericValue(sheet, row, col, 0);
  return v !== null && isFinite(v) ? v : null;
}
export function childrenData(key: string): any {
  const sheet = getSheet(); if (!sheet) return { error: 'No file loaded.' };
  const days = mainSections().filter((s) => /^day[123]$/.test(s.key));
  if (!days.length) return { error: 'Day sections not found.' };
  const cc = childCols(sheet, days[0]);
  if (!cc) return { error: "Could not find the 'Children Count' columns in this file." };
  const list = key === 'total' ? days : days.filter((d) => d.key === key);
  const rows: any[] = []; let tp: number | null = null, tn: number | null = null;
  for (let i = 0; i < days[0].teamRows.length; i++) {
    let p: number | null = null, w: number | null = null;
    list.forEach((d) => {
      const a = childVal(sheet, d.teamRows[i], cc), b = childVal(sheet, d.teamRows[i], cc + 1);
      if (a !== null) p = (p || 0) + a;
      if (b !== null) w = (w || 0) + b;
    });
    if (p !== null) tp = (tp || 0) + p;
    if (w !== null) tn = (tn || 0) + w;
    rows.push({ team: String(cellDisplayValue(sheet.getCell(days[0].teamRows[i], 1))), p, w });
  }
  return { days, list, rows, tp, tn };
}

/* ---------- new template ---------- */
function writeSectionTitle(sheet: any, s: Section) {
  const cell = sheet.getCell(s.titleRow, 1);
  cell.value = sectionTitleText(s); cell.font = { bold: true, size: 12 };
  sheet.mergeCells(s.titleRow, 1, s.titleRow, sectionKind(s) === 'next' ? 4 : 24);
}
function writeSectionHeaders(sheet: any, s: Section) {
  const [r1, r2] = s.headerRows;
  const put = (r: number, c: number, text: string) => {
    const cell = sheet.getCell(r, c); cell.value = text; cell.font = { bold: true, size: 9 };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; cell.fill = _fill('FFEFEBE2'); cell.border = _BORDER;
  };
  groupsFor(s).forEach(([label, a, b, subs]) => {
    put(r1, a, label);
    if (b > a) sheet.mergeCells(r1, a, r1, b); else if (!subs.length) sheet.mergeCells(r1, a, r2, a);
    subs.forEach((sub, i) => put(r2, a + i, sub));
  });
}
function writeSectionTeams(sheet: any, s: Section, codes: string[], all: Section[]) {
  const kind = sectionKind(s);
  const days = all.filter((x) => /^day[123]$/.test(x.key)), catches = all.filter((x) => /^catch\d+$/.test(x.key));
  const lastC = bodyLastCol(s);
  const style = (cell: any, argb?: string | null) => { cell.border = _BORDER; cell.alignment = { horizontal: 'center', vertical: 'middle' }; if (argb) cell.fill = _fill(argb); };
  const calc = (cell: any, f: string) => { cell.value = { formula: f }; cell.fill = _fill('FFF1F5F9'); };
  const calcText = (cell: any, f: string) => { cell.value = { formula: f }; cell.fill = _fill('FFFBF0D9'); };
  const L = colLetter;
  s.teamRows.forEach((row, i) => {
    const lab = sheet.getCell(row, COLS.TEAM);
    lab.value = kind === 'finalDay' ? ['Day 1', 'Day 2', 'Day 3'].concat(catches.map((_, j) => 'Catch Day ' + (j + 1)))[i] || null : codes[i] || 'T-' + (11 + i);
    lab.font = { bold: true, size: 9 }; lab.alignment = { horizontal: 'left', vertical: 'middle' }; lab.border = _BORDER;
    for (let c = 2; c <= lastC; c++) style(sheet.getCell(row, c));
    if (kind === 'day') {
      calc(sheet.getCell(row, 19), 'Q' + row + '-R' + row);
      calc(sheet.getCell(row, 25), 'D' + row);
      calc(sheet.getCell(row, 26), 'SUM(E' + row + ',F' + row + ')');
      calcText(sheet.getCell(row, 27), IFTXT('Y' + row, 'Z' + row));
      [25, 26, 27].forEach((c) => style(sheet.getCell(row, c), null));
    } else if (kind === 'catch') {
      const catchIdx = catches.findIndex((x) => x.key === s.key);
      if (catchIdx > 0) {
        const prevRow = catches[catchIdx - 1].teamRows[i];
        calc(sheet.getCell(row, 2), 'F' + prevRow);
        calc(sheet.getCell(row, 3), 'G' + prevRow);
      } else {
        const d = days.map((x) => x.teamRows[i]);
        if (d.length === 3) {
          calc(sheet.getCell(row, 2), 'SUM(F' + d[0] + ',F' + d[1] + ',F' + d[2] + ')-SUM(H' + d[0] + ',H' + d[1] + ',H' + d[2] + ')');
          calc(sheet.getCell(row, 3), 'SUM(G' + d[0] + ',G' + d[1] + ',G' + d[2] + ')-SUM(I' + d[0] + ',I' + d[1] + ',I' + d[2] + ')');
        }
      }
      calc(sheet.getCell(row, 6), 'B' + row + '-D' + row);
      calc(sheet.getCell(row, 7), 'C' + row + '-E' + row);
      calc(sheet.getCell(row, 15), 'M' + row + '-N' + row);
      if (i === 0 && s.key === 'catch1' && days.length === 3) {
        const t = days.map((x) => x.totalRow);
        [25, 26].forEach((c) => { style(sheet.getCell(row, c)); calc(sheet.getCell(row, c), 'SUM(' + L(c) + t[0] + ',' + L(c) + t[1] + ',' + L(c) + t[2] + ')'); });
        style(sheet.getCell(row, 27)); calcText(sheet.getCell(row, 27), IFTXT('Y' + row, 'Z' + row));
      }
    } else if (kind === 'finalTeam') {
      for (let c = 2; c <= 24; c++) {
        const parts = days.map((x) => L(c) + x.teamRows[i]);
        if (CATCH_MAP[c]) catches.forEach((cs) => parts.push(L(CATCH_MAP[c]) + cs.teamRows[i]));
        calc(sheet.getCell(row, c), 'SUM(' + parts.join(',') + ')');
      }
    } else if (kind === 'finalDay') {
      if (i < days.length) for (let c = 2; c <= 24; c++) calc(sheet.getCell(row, c), L(c) + days[i].totalRow);
      else if (catches[i - days.length]) {
        const cs = catches[i - days.length];
        const fm: Record<number, number> = Object.assign({}, CATCH_MAP, { 20: 16 });
        Object.keys(fm).forEach((c) => calc(sheet.getCell(row, Number(c)), L(fm[Number(c)]) + cs.totalRow));
      }
    } else if (kind === 'next') {
      const day = all.find((x) => x.key === 'day' + s.key.slice(-1));
      if (day) { calc(sheet.getCell(row, 2), 'C' + day.teamRows[i]); calc(sheet.getCell(row, 4), 'E' + day.teamRows[i] + '+F' + day.teamRows[i]); }
    }
  });
}
function writeSectionTotal(sheet: any, s: Section) {
  const row = s.totalRow, kind = sectionKind(s);
  const lab = sheet.getCell(row, COLS.TEAM);
  lab.value = 'Total'; lab.font = { bold: true, size: 9 }; lab.fill = _fill('FFDAD3C3'); lab.alignment = { horizontal: 'left', vertical: 'middle' }; lab.border = _BORDER;
  const first = s.teamRows[0], last = s.teamRows[s.teamRows.length - 1];
  const lastC = kind === 'day' ? 26 : bodyLastCol(s);
  for (let c = 2; c <= lastC; c++) {
    const cell = sheet.getCell(row, c);
    cell.font = { bold: true, size: 9 }; cell.fill = _fill('FFEFEBE2'); cell.alignment = { horizontal: 'center', vertical: 'middle' }; cell.border = _BORDER;
    cell.value = { formula: 'SUM(' + colLetter(c) + first + ':' + colLetter(c) + last + ')' };
  }
  if (kind === 'day') {
    const cell = sheet.getCell(row, 27);
    cell.font = { bold: true, size: 9 }; cell.fill = _fill('FFFBF0D9'); cell.alignment = { horizontal: 'center', vertical: 'middle' }; cell.border = _BORDER;
    cell.value = { formula: IFTXT('Y' + row, 'Z' + row) };
  }
}
export async function createBlankTemplate(nTeams: number, startCode: string, nCatch: number, startDate: string) {
  try {
    toast('Building template…', 'info');
    const ExcelJS = await XL();
    const wb = new ExcelJS.Workbook(); wb.creator = 'Polio Incharge'; wb.created = new Date();
    const sheet = wb.addWorksheet('F2A');
    const all = buildSectionMap(nTeams, nCatch).sections;
    const codes: string[] = [];
    for (let i = 0; i < nTeams; i++) { const m = startCode.match(/^(.*?)(\d+)$/); codes.push(m ? m[1] + (parseInt(m[2], 10) + i) : startCode + (i > 0 ? '-' + (i + 1) : '')); }
    all.forEach((s) => { writeSectionTitle(sheet, s); writeSectionHeaders(sheet, s); writeSectionTeams(sheet, s, codes, all); writeSectionTotal(sheet, s); });
    for (let c = 1; c <= LAST_COL; c++) sheet.getColumn(c).width = COL_WIDTHS[c] || 9;
    try {
      const meta = wb.addWorksheet(META_SHEET);
      meta.state = 'hidden';
      const metaObj: Record<string, any> = { nTeams, nCatch, startDate };
      Object.keys(metaObj).forEach((k, idx) => { meta.getCell(idx + 1, 1).value = k; meta.getCell(idx + 1, 2).value = metaObj[k]; });
    } catch { /* ignore */ }
    try { wb.calcProperties = { fullCalcOnLoad: true }; } catch { /* ignore */ }
    const buf = await wb.xlsx.writeBuffer();
    const d = new Date();
    const fname = `Polio Form 2A Template - Created ${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}.xlsx`;
    saveTemplateParams(fname, { nTeams, startCode, nCatch, startDate });
    if (await loadWorkbookFromBuffer(buf, fname, true)) {
      lsSet(metaKey(), JSON.stringify({ startDate }));
      try { await idbSet('polioCurrentFile', { name: fname, buffer: buf, dirty: false }); await idbSet('polioOriginalFile', { name: fname, buffer: buf }); } catch { /* ignore */ }
      toast('Template created. Start filling in your campaign data.', 'success');
    }
  } catch (err: any) { console.error(err); toast('Could not create template: ' + err.message, 'error'); }
}
export async function promptNewTemplate() {
  const res = await uiTemplateForm({ nTeams: 5, startCode: 'T-11', nCatch: 1, startDate: todayISO(), incharge: S.incharge });
  if (!res) return;
  if (res.incharge.trim()) setIncharge(res.incharge);
  const nTeams = Math.max(1, Math.min(20, Math.round(res.nTeams) || 5));
  const nCatch = Math.max(0, Math.min(2, Math.round(res.nCatch) || 0));
  const startCode = res.startCode.trim() || 'T-11';
  const startDate = /^\d{4}-\d{2}-\d{2}$/.test(res.startDate.trim()) ? res.startDate.trim() : todayISO();
  await createBlankTemplate(nTeams, startCode, nCatch, startDate);
}

/* ---------- start new campaign ---------- */
export async function startNewCampaign() {
  const sheet = getSheet(); if (!sheet) { toast('No workbook loaded.', 'error'); return; }
  const days = mainSections().filter((s) => /^day[123]$/.test(s.key)), catches = mainSections().filter((s) => /^catch\d+$/.test(s.key));
  const nexts = nextSections();
  if (nexts.length < 3 || days.length < 3) { toast('Template structure not recognized.', 'error'); return; }
  const pairs = nexts.slice(0, 3).map((n, i) => ({ next: n, target: days[i] }));
  const readCarry = (cell: any) => {
    const v = cell.value;
    if (cell.formula) { let n = evaluateFormula(sheet, cell.formula, 0); if (n === null || !isFinite(n)) n = typeof cell.result === 'number' ? cell.result : null; return n === null ? { ok: false, val: null } : { ok: true, val: n }; }
    if (v === null || v === undefined || v === '') return { ok: true, val: null };
    if (typeof v === 'number') return { ok: true, val: v };
    if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) return { ok: true, val: Number(v) };
    return { ok: false, val: null };
  };
  const captures: any[][] = [], problems: string[] = [];
  for (const p of pairs) {
    const batch: any[] = [];
    for (let i = 0; i < Math.min(p.next.teamRows.length, p.target.teamRows.length); i++) {
      const nr = p.next.teamRows[i];
      const t = readCarry(sheet.getCell(nr, COLS.TARGET)), c = readCarry(sheet.getCell(nr, COLS.CHILDREN_MP));
      if (!t.ok || !c.ok) problems.push(`${p.next.label}, team ${i + 1}`);
      batch.push({ targetRow: p.target.teamRows[i], target: t.val, children: c.val });
    }
    captures.push(batch);
  }
  if (problems.length) { await uiConfirm('Cannot start: these Next Campaign values are not plain numbers:\n\n' + problems.slice(0, 8).join('\n') + '\n\nNothing was changed.', { title: 'Cannot start', yesText: 'OK', noText: 'Close' }); return; }
  const lines = pairs.map((p) => `  • ${p.next.label} → ${p.target.label}`).join('\n');
  if (!(await uiConfirm('Start a new campaign?\n\n' + lines + '\n\nOnly the Target Micro Plan and Total Children Micro Plan are carried over. All other field data in Day 1/2/3 and Catch days is cleared.', { title: 'Start new campaign', yesText: 'Continue', danger: true }))) return;
  if (!(await uiConfirm('LAST WARNING\n\nThis permanently deletes all field data from the current campaign (Day 1–3 and Catch days). A backup file is downloaded first.', { title: 'Confirm deletion', yesText: 'Yes, start new campaign', danger: true }))) return;
  try {
    try { await downloadWorkbook('Backup before new campaign'); } catch { /* ignore */ }
    for (const sec of [...days, ...catches]) for (const r of sec.teamRows) for (let c = 1; c <= MAX_COLS; c++) {
      if (c === COLS.TEAM) continue;
      const cell = sheet.getCell(r, c);
      // Formula cells (Vaccine Returned, Still missed, Children Count...) keep their formula but must
      // lose the cached result from the old campaign, otherwise the old number is shown again.
      if (cell.formula) { cell.value = { formula: cell.formula }; continue; }
      const m = getMerge(r, c); if (m && !m.isMaster) continue;
      cell.value = null;
    }
    // Drop stale cached formula results across all sections (Day, Catch, Final Team, Final Day, Next).
    for (const sec of allSections()) {
      for (const r of [...sec.teamRows, ...(sec.totalRow ? [sec.totalRow] : [])]) {
        for (let c = 1; c <= MAX_COLS; c++) {
          const cell = sheet.getCell(r, c);
          if (cell.formula) cell.value = { formula: cell.formula };
        }
      }
    }
    for (const batch of captures) for (const it of batch) {
      if (it.target !== null && isFinite(it.target) && !sheet.getCell(it.targetRow, COLS.TARGET).formula) sheet.getCell(it.targetRow, COLS.TARGET).value = it.target;
      if (it.children !== null && isFinite(it.children) && !sheet.getCell(it.targetRow, COLS.CHILDREN_MP).formula) sheet.getCell(it.targetRow, COLS.CHILDREN_MP).value = it.children;
    }
    S.bulk = true; S.edited = new Set(); S.section = days[0].key; S.epoch++; emit(); scheduleAutosave();
    toast('New campaign started. Day 1/2/3 baselines set, old data cleared.', 'success');
  } catch (err: any) { console.error(err); toast('Failed to start new campaign: ' + err.message, 'error'); }
}

/* ---------- summary of the day (shown in the app only, never written to Excel) ---------- */
// Day:   Total Children Team (E) + Covered Missed (H, I) + School (J) + Guest (K) + Street (L)
// Catch: Covered Missed (D, E) + Guest Vaccinated (H) + Street (I)
const SUM_COLS: Record<string, number[]> = { day: [5, 8, 9, 10, 11, 12], catch: [4, 5, 8, 9] };
export const VACCINE_FACTOR = 1.11;
export function daySummary(section: any): { rows: { team: string; covered: number; vaccine: number }[]; total: { covered: number; vaccine: number } } | null {
  const sheet = getSheet(); if (!sheet) return null;
  const kind = /^day[123]$/.test(section.key) ? 'day' : /^catch\d+$/.test(section.key) ? 'catch' : '';
  if (!kind) return null;
  const cols = SUM_COLS[kind];
  const rows = section.teamRows.map((r: number) => {
    let covered = 0;
    for (const c of cols) { const v = getCellNumericValue(sheet, r, c); if (typeof v === 'number' && isFinite(v)) covered += v; }
    return { team: String(cellDisplayValue(sheet.getCell(r, 1)) || ''), covered, vaccine: covered * VACCINE_FACTOR };
  });
  const tc = rows.reduce((a: number, x: any) => a + x.covered, 0);
  return { rows, total: { covered: tc, vaccine: tc * VACCINE_FACTOR } };
}

/* ---------- add Catch Day 2 (button in the F2A tab) ---------- */
export async function addCatchDay2() {
  const sheet = getSheet(); if (!sheet) { toast('No workbook loaded.', 'error'); return; }
  if (Number(S.params.nCatch) !== 1) { toast('This file does not have exactly 1 catch-up day.', 'info'); return; }
  const ok = await uiConfirm('A second catch-up form will be added after Catch Day 1. Final Report and Next Campaign move down and all formulas are updated. Download a backup first if you want to be safe. Continue?', { title: 'Add Catch Day 2', yesText: 'Add', noText: 'Cancel' });
  if (!ok) return;
  toast('Adding Catch Day 2… please wait.', 'info');
  await new Promise((r) => setTimeout(r, 60)); // let the message appear before the heavy work starts
  try {
    addCatchDay2Core(S.workbook, Number(S.params.nTeams));
    applySectionMap(Number(S.params.nTeams), 2);
    S.mergeMap = buildMergeMap(sheet);
    saveTemplateParams(S.fileName, { nTeams: S.params.nTeams, nCatch: 2 });
    writeSheetMeta({ nTeams: S.params.nTeams, nCatch: 2, startDate: S.startDate || '' });
    S.section = 'catch2'; S.bulk = true; S.epoch++; emit(); scheduleAutosave();
    toast('Catch Day 2 added.', 'success');
  } catch (e: any) { console.error(e); toast('Could not add Catch Day 2: ' + (e && e.message ? e.message : e), 'error'); }
}
