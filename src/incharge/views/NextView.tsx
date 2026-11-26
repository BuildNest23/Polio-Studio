import type { Dict } from '../lib/i18n';
import { S, emit, nextSections } from '../lib/state';
import { startNewCampaign } from '../lib/workbook';
import { SectionTable } from './SectionTable';

/** More > Next campaign: Next Camp 1-3 to look at (not editable) + Start new campaign. */
export function NextView({ t }: { t: Dict }) {
  const next = nextSections();
  const top = (
    <div className="row" style={{ flex: 'none' }}>
      <button className="btn" onClick={() => { S.moreSub = ''; emit(); }}>‹ {t.back}</button>
      <h1 className="h1" style={{ margin: 0 }}>{t.nextCamp}</h1>
    </div>
  );
  return (
    <SectionTable
      t={t}
      sections={next}
      sectionKey={S.nextSection}
      onSelectSection={(k) => { S.nextSection = k; emit(); }}
      readOnly
      top={top}
      action={{ label: t.startNew, onClick: startNewCampaign, danger: true }}
    />
  );
}
