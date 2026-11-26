import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { evaluateFormula, evalIncDec } from '../src/incharge/lib/engine';
import { addCatchDay2, detectLayout } from '../src/incharge/lib/catch2';
import { buildSectionMap } from '../src/incharge/lib/core';
import { importTotals, normTeamCode } from '../src/shared/teamfile';
import { imp, importAddData, importDoneKey, isImportDone } from '../src/incharge/lib/teamio';
import { addSchool, buildFile, initial, saveNext, setFound, setNum, setSchool, stillMissed, totals } from '../src/team/model';
import { createBlankTemplate } from '../src/incharge/lib/workbook';
import { S } from '../src/incharge/lib/state';

function makeSheet(cells: Record<string, any>) {
  return {
    getCell(r: number, c: number) {
      const v = cells[r + ',' + c];
      if (v === undefined) return { value: null, formula: undefined };
      if (typeof v === 'object' && v !== null && 'formula' in v) return { value: v, formula: v.formula, result: v.result };
      return { value: v, formula: undefined };
    },
  };
}

test('formula engine treats single blank cell as 0 in A+B and A-B, returns null when both blank', () => {
  const sheet = makeSheet({
    '4,5': 45,   // E4 = 45, F4 = blank
    '4,17': 100, // Q4 = 100, R4 = blank
  });
  assert.equal(evaluateFormula(sheet, 'E4+F4'), 45);
  assert.equal(evaluateFormula(sheet, 'Q4-R4'), 100);
  assert.equal(evaluateFormula(sheet, 'F4+G4'), null);
  assert.equal(evaluateFormula(sheet, 'R4-S4'), null);
});

test('formula engine supports $ absolute cell references', () => {
  const sheet = makeSheet({
    '4,2': 10,
    '5,2': 15,
    '4,25': 20,
    '4,26': 25,
  });
  assert.equal(evaluateFormula(sheet, 'SUM($B$4:$B$5)'), 25);
  assert.equal(evaluateFormula(sheet, '$B$5-$B$4'), 5);
  assert.equal(evaluateFormula(sheet, 'MINUS($B$5, $B$4)'), 5);
  assert.equal(
    evalIncDec(sheet, 'IF($Y$4<$Z$4,"Increased by "&($Z$4-$Y$4),IF($Y$4>$Z$4,"Decreased by "&($Y$4-$Z$4),"No change"))'),
    'Increased by 5',
  );
});

test('team model auto-clamps revisit count when house total or vaccinated changes', () => {
  let st = initial();
  st = setNum(st, 0, 't', 5);
  st = setNum(st, 0, 'v', 3); // 2 missed
  st = saveNext(st);
  st = setFound(st, 0, 2);    // 2 found on revisit
  assert.equal(stillMissed(st.houses[0]), 0);
  assert.equal(totals(st).still, 0);
  assert.equal(totals(st).totalVacc, 5);

  // Raise vaccinated to 5 -> revisit must drop to 0 so still is 0 (never -2) and totalVacc is 5 (never 7)
  st = setNum(st, 0, 'v', 5);
  assert.equal(st.houses[0].f, 0);
  assert.equal(totals(st).still, 0);
  assert.equal(totals(st).totalVacc, 5);

  // Lower total below vaccinated -> both v and f stay clamped
  st = setNum(st, 0, 'v', 2);
  st = setFound(st, 0, 3);
  st = setNum(st, 0, 't', 3);
  assert.equal(st.houses[0].v, 2);
  assert.equal(st.houses[0].f, 1);
  assert.equal(totals(st).still, 0);
});

test('team model sanitizes legacy corrupt state in totals and buildFile and clamps schools', () => {
  const corrupt = {
    cur: 0,
    houses: [{ t: 4, v: 4, g: 0, l: false, m: false, f: 3, done: true }],
    school: [{ name: 'Primary', u5: 10, vac: 15 }],
    outside: 0,
  };
  const T = totals(corrupt);
  assert.equal(T.C, 0);
  assert.equal(T.still, 0);
  assert.equal(T.S, 10);

  const exported = buildFile({ campaignId: 'C1', team: 'T-11', day: 1, area: '', uc: '', vaccinator: '', incharge: '' }, corrupt);
  assert.deepEqual(exported.houses[0], [4, 4, 0, 0, 0, 0]);
  assert.deepEqual(exported.school[0], { name: 'Primary', u5: 10, vac: 10 });

  let st = addSchool(initial());
  st = setSchool(st, 0, { u5: 5, vac: 9 });
  assert.equal(st.school[0].vac, 5);
  st = setSchool(st, 0, { u5: 2 });
  assert.equal(st.school[0].vac, 2);
});

test('teamfile normTeamCode and importTotals handle formatting and malformed values safely', () => {
  assert.equal(normTeamCode('T-12'), 'T12');
  assert.equal(normTeamCode(' t 12 '), 'T12');
  const T = importTotals({
    houses: [['5', '3', '1', 0, 1, '99'], [-4, 'abc', null, 0, 0, -2]],
    outside: '2',
    school: [{ name: 'A', u5: 5, vac: '4' }],
  });
  assert.equal(T.A, 5);
  assert.equal(T.B, 3);
  assert.equal(T.O, 1);
  assert.equal(T.C, 2); // clamped to A - B
  assert.equal(T.D, 2);
  assert.equal(T.S, 4);
});

test('importAddData clears Added to F2A flag when replacing an existing team/day file', () => {
  imp.files = [];
  imp.done = {};
  const payload1 = JSON.stringify({ app: 'polio-team', version: 1, campaignId: 'C1', team: 'T-11', day: 1, houses: [[3, 3, 0, 0, 0, 0]], school: [], outside: 0 });
  const payload2 = JSON.stringify({ app: 'polio-team', version: 1, campaignId: 'C1', team: 'T 11', day: 1, houses: [[4, 4, 0, 0, 0, 0]], school: [], outside: 1 });

  assert.equal(importAddData(payload1, 'f1.txt'), true);
  imp.done[importDoneKey({ team: 'T-11', day: 1 })] = Date.now();
  assert.equal(isImportDone({ team: 'T-11', day: 1 }), true);

  // Loading an updated file for same team (even with different spacing) replaces it & clears done
  assert.equal(importAddData(payload2, 'f2.txt'), true);
  assert.equal(imp.files.length, 1);
  assert.equal(imp.files[0].T.A, 4);
  assert.equal(isImportDone({ team: 'T-11', day: 1 }), false);
  assert.equal(isImportDone({ team: 'T 11', day: 1 }), false);
});

test('blank template with 2 catch-up days chains Catch Day 2 target to Catch Day 1 Still Missed and detects layout', async () => {
  for (const k of [0, 1, 2]) {
    await createBlankTemplate(5, 'T-11', k, '2026-11-25');
    const ws = S.workbook.getWorksheet('F2A');
    assert.deepEqual(detectLayout(ws), { nTeams: 5, nCatch: k });
  }
  const ws = S.workbook.getWorksheet('F2A');
  const map = buildSectionMap(5, 2);
  const c1 = map.sections.find((s) => s.key === 'catch1')!;
  const c2 = map.sections.find((s) => s.key === 'catch2')!;
  assert.equal(ws.getCell(c2.teamRows[0], 2).formula, 'F' + c1.teamRows[0]);
  assert.equal(ws.getCell(c2.teamRows[0], 3).formula, 'G' + c1.teamRows[0]);
});

test('addCatchDay2 populates Catch Day 2 label in Final Report Day Wise on blank 1-catch template and keeps detectLayout accurate', async () => {
  await createBlankTemplate(5, 'T-11', 1, '2026-11-25');
  const wb = S.workbook as ExcelJS.Workbook;
  addCatchDay2(wb, 5);
  const ws = wb.getWorksheet('F2A')!;
  const map = buildSectionMap(5, 2);
  const fd = map.sections.find((s) => s.key === 'finalDay')!;
  assert.equal(ws.getCell(fd.teamRows[4], 1).value, 'Catch Day 2');
  assert.deepEqual(detectLayout(ws), { nTeams: 5, nCatch: 2 });
});
