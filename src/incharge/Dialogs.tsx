import { useEffect, useRef, useState, type FormEvent } from 'react';
import { useLang } from '../shared/prefs';
import { PlainTextInput } from '../shared/PlainTextInput';
import { I18N } from './lib/i18n';
import { useDialog, useToasts, type TemplateFormData } from './lib/state';

function TemplateFormSheet({ def, resolve }: { def: TemplateFormData; resolve: (v: TemplateFormData | null) => void }) {
  const [lang] = useLang();
  const t = I18N[lang];
  const [teams, setTeams] = useState(String(def.nTeams));
  const [startCode, setStartCode] = useState(def.startCode);
  const [nCatch, setNCatch] = useState(def.nCatch);
  const [startDate, setStartDate] = useState(def.startDate);
  const [incharge, setIncharge] = useState(def.incharge);
  const [err, setErr] = useState('');

  const submit = (e: FormEvent) => {
    e.preventDefault();
    const n = parseInt(teams.trim(), 10);
    if (!Number.isInteger(n) || n < 1 || n > 20) { setErr(t.tplErrTeams); return; }
    if (!startCode.trim()) { setErr(t.tplErrCode); return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(startDate.trim())) { setErr(t.tplErrDate); return; }
    setErr('');
    resolve({ nTeams: n, startCode: startCode.trim(), nCatch, startDate: startDate.trim(), incharge: incharge.trim() });
  };

  return (
    <form className="sheet" role="dialog" aria-modal="true" aria-labelledby="tpl-dlg-title" onSubmit={submit}>
      <h3 id="tpl-dlg-title">{t.tplTitle}</h3>
      {err && <div className="notice red" role="alert">{err}</div>}
      <div className="tpl-grid">
        <label className="field">
          {t.tplTeams}
          <PlainTextInput className="input" inputMode="numeric" value={teams} onChange={(e) => { setTeams(e.target.value.replace(/[^\d]/g, '').slice(0, 2)); setErr(''); }} />
        </label>
        <label className="field">
          {t.tplStartCode}
          <PlainTextInput className="input" value={startCode} placeholder="T-11" onChange={(e) => { setStartCode(e.target.value); setErr(''); }} />
        </label>
      </div>
      <div className="field">
        <span>{t.tplCatch}</span>
        <div className="seg" role="group" aria-label={t.tplCatch}>
          {[0, 1, 2].map((k) => (
            <button key={k} type="button" className={'btn' + (nCatch === k ? ' on' : '')} aria-pressed={nCatch === k} onClick={() => setNCatch(k)}>
              {k}
            </button>
          ))}
        </div>
      </div>
      <label className="field">
        {t.startDate}
        <input className="input" type="date" value={startDate} onChange={(e) => { setStartDate(e.target.value); setErr(''); }} />
      </label>
      <label className="field">
        {t.tplIncharge}
        <PlainTextInput className="input" value={incharge} onChange={(e) => setIncharge(e.target.value)} />
      </label>
      <div className="row">
        <button type="button" className="btn grow" onClick={() => resolve(null)}>{t.cancel}</button>
        <button type="submit" className="btn grow primary">{t.create}</button>
      </div>
    </form>
  );
}

export function Dialogs() {
  const dlg = useDialog();
  const toasts = useToasts();
  const inp = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    if (!dlg) return;
    if (dlg.kind === 'prompt') {
      const id = setTimeout(() => { inp.current?.focus(); inp.current?.select(); }, 60);
      return () => clearTimeout(id);
    }
  }, [dlg]);

  useEffect(() => {
    if (!dlg) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        dlg.resolve(dlg.kind === 'confirm' ? false : null);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [dlg]);

  return (
    <>
      {dlg && (
        <div
          key={dlg.id}
          className={'overlay' + (dlg.kind === 'prompt' || dlg.kind === 'template' ? ' top' : '')}
          onClick={(e) => { if (e.target === e.currentTarget) dlg.resolve(dlg.kind === 'confirm' ? false : null); }}
        >
          {dlg.kind === 'template' && dlg.tplDef ? (
            <TemplateFormSheet def={dlg.tplDef} resolve={dlg.resolve} />
          ) : (
            <div className="sheet" role="dialog" aria-modal="true">
              <h3>{dlg.title}</h3>
              <p>{dlg.msg}</p>
              {dlg.kind === 'prompt' && (
                <PlainTextInput
                  key={dlg.id}
                  ref={inp}
                  className="input"
                  defaultValue={dlg.def}
                  onKeyDown={(e) => { if (e.key === 'Enter') dlg.resolve((e.target as HTMLTextAreaElement).value); }}
                />
              )}
              <div className="row">
                <button className="btn grow" onClick={() => dlg.resolve(dlg.kind === 'prompt' ? null : false)}>{dlg.no}</button>
                <button className={'btn grow ' + (dlg.danger ? 'danger' : 'primary')} onClick={() => dlg.resolve(dlg.kind === 'prompt' ? inp.current?.value ?? '' : true)}>{dlg.yes}</button>
              </div>
            </div>
          )}
        </div>
      )}
      <div className="toasts">{toasts.map((t) => <div key={t.id} className={'toast ' + t.type}>{t.msg}</div>)}</div>
    </>
  );
}
