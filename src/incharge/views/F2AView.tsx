import type { Dict } from '../lib/i18n';
import { S, emit, isDirty, mainSections } from '../lib/state';
import { addCatchDay2, discardChanges, downloadWorkbook } from '../lib/workbook';
import { SectionTable } from './SectionTable';

export function F2AView({ t }: { t: Dict }) {
  const dirty = isDirty();
  const label = S.edited.size ? S.edited.size + ' ' + t.unsaved : S.bulk ? t.bulk : t.clean;
  const top = (
    <div className="card" style={{ flex: 'none', flexDirection: 'row', alignItems: 'center', padding: '8px 12px', borderRadius: 14 }}>
      <div className="grow" style={{ display: 'flex', flexDirection: 'column' }}>
        <b dir="ltr" style={{ fontSize: 13, textAlign: 'start', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{S.fileName}</b>
        <span className="small muted"><span className={'dot' + (dirty ? ' dirty' : '')} />{label}</span>
      </div>
      <button className="btn danger" style={{ minHeight: 40, padding: '0 14px', fontSize: 13 }} disabled={!dirty} onClick={discardChanges}>{t.discard}</button>
    </div>
  );
  return (
    <SectionTable
      t={t}
      sections={mainSections()}
      sectionKey={S.section}
      onSelectSection={(k) => { S.section = k; emit(); }}
      top={top}
      action={{ label: t.download, onClick: () => downloadWorkbook() }}
      chipAction={Number(S.params.nCatch) === 1 ? { label: t.addCatch2, onClick: addCatchDay2 } : undefined}
    />
  );
}
