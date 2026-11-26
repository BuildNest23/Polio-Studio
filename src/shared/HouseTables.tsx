// House tables (30 houses per card) shown in BOTH apps. Columns: House, L/MMP mark, Total, Vaccinated, Guest, (Revisit).
import type { ReactElement } from 'react';

export interface HouseRow { n: number; t: number; v: number; g: number; f: number; l: boolean; m: boolean }
export interface HouseLabels { house: string; total: string; vacc: string; guest: string; revisit: string; sum: string; range: string; lock: string; mmp: string }

const safeNum = (v: unknown) => { const n = Number(v); return Number.isFinite(n) && n >= 0 ? n : 0; };

export function HouseTables({ rows, L, showRevisit = false }: { rows: HouseRow[]; L: HouseLabels; showRevisit?: boolean }) {
  const out: ReactElement[] = [];
  const size = 30;
  const last = rows.reduce((m, r) => Math.max(m, r.n), 0);
  for (let st = 1; st <= last; st += size) {
    const part = rows.filter((r) => r.n >= st && r.n < st + size);
    if (!part.length) continue;
    let a = 0, b = 0, g = 0, f = 0;
    part.forEach((r) => { if (!r.l) { a += safeNum(r.t); b += safeNum(r.v); g += safeNum(r.g); f += safeNum(r.f); } });
    out.push(
      <div className="card" key={st}>
        <b>{L.range} {st}–{st + size - 1}</b>
        <table className="rvtable">
          <thead><tr><th>{L.house}</th><th></th><th>{L.total}</th><th>{L.vacc}</th><th>{L.guest}</th>{showRevisit && <th>{L.revisit}</th>}</tr></thead>
          <tbody>
            {part.map((r) => {
              const t = safeNum(r.t), v = safeNum(r.v), gu = safeNum(r.g), fo = safeNum(r.f);
              return (
                <tr key={r.n}>
                  <td>{r.n}</td>
                  <td className="mk">{r.l ? L.lock : ''}{r.m ? (r.l ? ' ' : '') + L.mmp : ''}</td>
                  <td>{r.l ? '–' : t}</td><td>{r.l ? '–' : v}</td><td>{r.l ? '–' : gu}</td>
                  {showRevisit && <td>{r.l ? '–' : fo || ''}</td>}
                </tr>
              );
            })}
          </tbody>
          <tfoot><tr><th>{L.sum}</th><th></th><th>{a}</th><th>{b}</th><th>{g}</th>{showRevisit && <th>{f}</th>}</tr></tfoot>
        </table>
      </div>,
    );
  }
  return <>{out}</>;
}
