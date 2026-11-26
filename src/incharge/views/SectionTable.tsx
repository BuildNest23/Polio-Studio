import { useState, type ReactNode } from 'react';
import type { Dict } from '../lib/i18n';
import { S, emit } from '../lib/state';
import { buildSectionModel, daySummary, editCell, printSection, sectionDateFull } from '../lib/workbook';
import { PlainTextInput } from '../../shared/PlainTextInput';

// [font size, row height] for the A- / A+ buttons
const LEVELS: [number, number][] = [[11, 24], [12.5, 28], [14, 32], [16, 38], [18, 44]];

function CellInput({ row, col, init }: { row: number; col: number; init: string }) {
  const [v, setV] = useState(String(init));
  return (
    <PlainTextInput
      type="text" inputMode="decimal" aria-label={'R' + row + ' C' + col} value={v}
      style={{ width: Math.max(v.length, 2) + 1 + 'ch' }}
      onFocus={(e) => e.target.select()}
      onChange={(e) => { const raw = e.target.value; if (editCell(row, col, raw)) setV(raw); }}
      onBlur={() => { const t = v.trim(); if (t !== '' && !isNaN(Number(t))) setV(String(Number(t))); }}
    />
  );
}

interface Props {
  t: Dict;
  sections: any[];
  sectionKey: string;
  onSelectSection: (key: string) => void;
  readOnly?: boolean;
  top?: ReactNode;                 // card above the chips (file name etc.)
  action: { label: string; onClick: () => void; danger?: boolean };
  chipAction?: { label: string; onClick: () => void };   // secondary chip at end of section strip (e.g. + Catch Day 2)
}

/** Section chips + table + zoom/print buttons. Used by the F2A tab and by More > Next campaign. */
export function SectionTable({ t, sections, sectionKey, onSelectSection, readOnly = false, top, action, chipAction }: Props) {
  const [lvl, setLvl] = useState(() => { try { const n = parseInt(localStorage.getItem('pcCell3') || '1', 10); return n >= 0 && n < 5 ? n : 1; } catch { return 1; } });
  const zoom = (d: number) => { const n = Math.max(0, Math.min(4, lvl + d)); setLvl(n); try { localStorage.setItem('pcCell3', String(n)); } catch { /* ignore */ } };
  const section = sections.find((s) => s.key === sectionKey) || sections[0];
  if (!section) return null;
  const rows = buildSectionModel(section);
  const date = sectionDateFull(section.key);
  const [f, h] = LEVELS[lvl];
  const sum = readOnly ? null : daySummary(section);
  const fmt = (n: number) => String(Number(n.toFixed(1)));

  return (
    <>
      {top}
      <div className="chips">
        {sections.map((s) => (
          <button key={s.key} className={'chip' + (s.key === section.key ? ' on' : '')} onClick={() => onSelectSection(s.key)}>{s.label}</button>
        ))}
        {chipAction && (
          <button className="chip chip-action" onClick={chipAction.onClick}>{chipAction.label}</button>
        )}
      </div>
      <div className="sectitle">
        <b style={{ fontSize: 15 }}>{section.label}</b>
        {date ? <span dir="ltr" className="small muted">{date}</span> : null}
      </div>
      <div className="f2a-wrap" style={{ ['--fs' as any]: f + 'px', ['--rh' as any]: h + 'px' }}>
        <table className="f2a" key={section.key}>
          <thead>
            {rows.filter((r) => r.kind === 'header').map((r) => (
              <tr key={'h' + r.idx} className={r.idx === 0 ? 'h1r' : 'h2r'}>
                {r.cells.map((c) => <th key={c.col} className={c.cls} colSpan={c.colspan} rowSpan={c.rowspan}>{String(c.display)}</th>)}
              </tr>
            ))}
          </thead>
          <tbody>
            {rows.filter((r) => r.kind !== 'header').map((r) => (
              <tr key={r.kind + r.idx} className={r.kind === 'total' ? 'total' : ''}>
                {r.cells.map((c) => (
                  <td key={c.col + '-' + S.epoch} className={c.cls} colSpan={c.colspan} rowSpan={c.rowspan}>
                    {c.input && !readOnly ? <CellInput row={c.row} col={c.col} init={String(c.display)} /> : String(c.display)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {sum && (
        <div className="daysum">
          <b>{t.sumTitle}</b>
          <table className="blocktbl">
            <thead><tr><th style={{ textAlign: 'start' }}>{t.sumTeam}</th><th>{t.sumCovered}</th><th>{t.sumVaccine}</th></tr></thead>
            <tbody>
              {sum.rows.map((r, i) => <tr key={i}><td style={{ textAlign: 'start', fontWeight: 700 }}>{r.team}</td><td>{fmt(r.covered)}</td><td>{fmt(r.vaccine)}</td></tr>)}
              <tr className="sum"><td style={{ textAlign: 'start' }}>{t.total}</td><td>{fmt(sum.total.covered)}</td><td>{fmt(sum.total.vaccine)}</td></tr>
            </tbody>
          </table>
        </div>
      )}
      {!readOnly && (
        <div className="legend">
          <span><i style={{ background: 'var(--card)' }} />{t.legEdit}</span>
          <span><i style={{ background: 'var(--calc)' }} />{t.legCalc}</span>
          <span><i style={{ background: 'var(--chg)' }} />{t.legChg}</span>
        </div>
      )}
      <div className="f2a-actions">
        <button className="btn" style={{ width: 48, padding: 0 }} onClick={() => zoom(-1)} aria-label="Smaller cells">A−</button>
        <button className="btn" style={{ width: 48, padding: 0, fontSize: 17 }} onClick={() => zoom(1)} aria-label="Bigger cells">A+</button>
        <button className="btn" onClick={() => printSection(section)}>{t.print}</button>
        <button className={'btn grow ' + (action.danger ? 'danger' : 'primary')} onClick={action.onClick}>{action.label}</button>
      </div>
    </>
  );
}
