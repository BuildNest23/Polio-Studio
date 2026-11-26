import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { clamp, MAX } from './model';
import { PlainTextInput } from '../shared/PlainTextInput';

/* ---- tiny confirm sheet + copy fallback sheet + toast (own copy, so this file stays single and small) ---- */
interface Ask { msg: string; yes: string; no: string; third?: string; resolve: (v: 'yes' | 'no' | 'third') => void }
interface CopyDlg { title: string; text: string; selectLabel: string; closeLabel: string; resolve: () => void }
let ask: Ask | null = null; const askSubs = new Set<() => void>();
const setAsk = (a: Ask | null) => { ask = a; askSubs.forEach((f) => f()); };
export function confirmSheet(msg: string, yes: string, no: string, third?: string): Promise<'yes' | 'no' | 'third'> {
  return new Promise((res) => setAsk({ msg, yes, no, third, resolve: (v) => { setAsk(null); res(v); } }));
}
let copyDlg: CopyDlg | null = null; const copySubs = new Set<() => void>();
const setCopyDlg = (c: CopyDlg | null) => { copyDlg = c; copySubs.forEach((f) => f()); };
export function copySheet(text: string, title: string, selectLabel: string, closeLabel: string): Promise<void> {
  return new Promise((res) => setCopyDlg({ title, text, selectLabel, closeLabel, resolve: () => { setCopyDlg(null); res(); } }));
}
let toasts: { id: number; msg: string; warn?: boolean }[] = []; let tid = 1; const toastSubs = new Set<() => void>();
export function toast(msg: string, warn = false) {
  const id = tid++; toasts = [{ id, msg, warn }]; /* one at a time: new replaces old */ toastSubs.forEach((f) => f());
  setTimeout(() => { toasts = toasts.filter((t) => t.id !== id); toastSubs.forEach((f) => f()); }, 3200);
}
export function Overlays() {
  const a = useSyncExternalStore((cb) => { askSubs.add(cb); return () => { askSubs.delete(cb); }; }, () => ask);
  const c = useSyncExternalStore((cb) => { copySubs.add(cb); return () => { copySubs.delete(cb); }; }, () => copyDlg);
  const ts = useSyncExternalStore((cb) => { toastSubs.add(cb); return () => { toastSubs.delete(cb); }; }, () => toasts);
  const copyRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    if (!a && !c) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      e.preventDefault();
      if (c) c.resolve();
      else if (a) a.resolve(a.third ? 'third' : 'no');
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [a, c]);

  useEffect(() => {
    if (c) {
      const id = setTimeout(() => { copyRef.current?.focus(); copyRef.current?.select(); }, 60);
      return () => clearTimeout(id);
    }
  }, [c]);

  return (
    <>
      {a && (
        <div className="overlay" onClick={(e) => { if (e.target === e.currentTarget) a.resolve(a.third ? 'third' : 'no'); }}>
          <div className="sheet" role="dialog" aria-modal="true">
            <p>{a.msg}</p>
            <div className="row">
              <button className="btn grow" onClick={() => a.resolve('no')}>{a.no}</button>
              {a.third && <button className="btn grow" onClick={() => a.resolve('third')}>{a.third}</button>}
              <button className="btn primary grow" onClick={() => a.resolve('yes')}>{a.yes}</button>
            </div>
          </div>
        </div>
      )}
      {c && (
        <div className="overlay" onClick={(e) => { if (e.target === e.currentTarget) c.resolve(); }}>
          <div className="sheet" role="dialog" aria-modal="true">
            <h3>{c.title}</h3>
            <textarea ref={copyRef} className="input" rows={5} readOnly value={c.text} onFocus={(e) => e.target.select()} />
            <div className="row">
              <button className="btn grow" onClick={() => { copyRef.current?.focus(); copyRef.current?.select(); }}>{c.selectLabel}</button>
              <button className="btn primary grow" onClick={() => c.resolve()}>{c.closeLabel}</button>
            </div>
          </div>
        </div>
      )}
      <div className="toasts">{ts.map((t) => <div key={t.id} className={'toast ' + (t.warn ? 'w' : t.id % 2 ? 'a' : 'b')}>{t.msg}</div>)}</div>
    </>
  );
}

/** Typed number: tap the number to type it. Empty becomes 0. */
export function NumInput({ value, onChange, label, max = MAX }: { value: number; onChange: (n: number) => void; label: string; max?: number }) {
  return (
    <PlainTextInput
      type="text" inputMode="numeric" aria-label={label} value={String(value)}
      onFocus={(e) => e.target.select()}
      onChange={(e) => { const n = parseInt(e.target.value.replace(/\D/g, ''), 10); onChange(clamp(isNaN(n) ? 0 : n, 0, max)); }}
    />
  );
}
/** [-] number [+]. editable=true lets the number be typed (Entry tab); false = buttons only (Missed tab). */
export function Stepper({ value, onChange, label, max = MAX, editable = true, small = false }: { value: number; onChange: (n: number) => void; label: string; max?: number; editable?: boolean; small?: boolean }) {
  return (
    <div className={'stepper' + (small ? ' small' : '')}>
      <button className="sbtn" aria-label="Minus" onClick={() => onChange(clamp(value - 1, 0, max))}>−</button>
      {editable ? <NumInput value={value} onChange={onChange} label={label} max={max} /> : <div className="stepnum" style={{ flex: small ? 'none' : 1, minWidth: 40 }}>{value}</div>}
      <button className="sbtn plus" aria-label="Plus" onClick={() => onChange(clamp(value + 1, 0, max))}>+</button>
    </div>
  );
}
export function useAutoFocus<T extends HTMLElement>() { const r = useRef<T>(null); useEffect(() => { r.current?.focus(); }, []); return r; }
export const useBool = (init = false) => useState(init);

/** Small Urdu text (only when Urdu is switched on). */
export function Ur({ on, children }: { on: boolean; children: string }) {
  return on ? <span className="urdu" lang="ur" dir="rtl">{children}</span> : null;
}
