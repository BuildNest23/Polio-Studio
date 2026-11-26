import { forwardRef, useCallback, useEffect, useRef, type ChangeEvent, type KeyboardEvent, type TextareaHTMLAttributes } from 'react';

/**
 * One-line text box built on <textarea>.
 *
 * Why: on Android (Chrome, PWA, Capacitor WebView) Google Autofill shows a key / card / location
 * toolbar above the keyboard on <input> fields, and autocomplete="off" does not reliably remove it.
 * It never shows on <textarea>, so every single-line text / phone / search field uses this instead.
 *
 * Behaves like <input>: one row, no wrapping, Enter never inserts a new line (it submits the
 * parent form, if any), pasted line breaks become spaces.
 * `type` is accepted so old call sites keep working: 'tel' -> inputMode="tel".
 * Do NOT use for number / date / month / time / checkbox / radio / file - a textarea cannot do those.
 */
export type PlainTextInputProps = Omit<TextareaHTMLAttributes<HTMLTextAreaElement>, 'rows' | 'wrap'> & {
  type?: 'text' | 'tel' | 'search' | string;
};

const LINE_BREAKS = /\r\n|\r|\n/g;

export const PlainTextInput = forwardRef<HTMLTextAreaElement, PlainTextInputProps>(function PlainTextInput(
  { type, className, onKeyDown, onChange, inputMode, ...rest },
  fwd,
) {
  const local = useRef<HTMLTextAreaElement | null>(null);
  const setRef = useCallback((el: HTMLTextAreaElement | null) => {
    local.current = el;
    if (typeof fwd === 'function') fwd(el);
    else if (fwd) (fwd as { current: HTMLTextAreaElement | null }).current = el;
  }, [fwd]);

  /* Android soft keyboards often send Enter as an input event (keyCode 229) with no usable keydown.
     Catch the line-break insertion before it happens and replay it as a normal Enter keydown, so the
     field's own onKeyDown and the form submit behave exactly as on a real <input>. */
  useEffect(() => {
    const el = local.current;
    if (!el) return;
    const onBefore = (ev: InputEvent) => {
      if (ev.inputType !== 'insertLineBreak' && ev.inputType !== 'insertParagraph') return;
      ev.preventDefault();
      el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true, cancelable: true }));
    };
    el.addEventListener('beforeinput', onBefore);
    return () => el.removeEventListener('beforeinput', onBefore);
  }, []);

  const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    onKeyDown?.(e);
    if (e.defaultPrevented) return;
    if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
      e.preventDefault();
      e.currentTarget.form?.requestSubmit();
    }
  };

  const handleChange = (e: ChangeEvent<HTMLTextAreaElement>) => {
    const el = e.currentTarget;
    if (LINE_BREAKS.test(el.value)) {
      LINE_BREAKS.lastIndex = 0;
      const pos = el.selectionStart ?? el.value.length;
      const before = el.value.slice(0, pos).replace(LINE_BREAKS, ' ').length;
      el.value = el.value.replace(LINE_BREAKS, ' ');
      try { el.setSelectionRange(before, before); } catch { /* ignore */ }
    }
    LINE_BREAKS.lastIndex = 0;
    onChange?.(e);
  };

  return (
    <textarea
      {...rest}
      ref={setRef}
      rows={1}
      wrap="off"
      autoComplete="off"
      autoCorrect="off"
      data-form-type="other"
      data-lpignore="true"
      inputMode={inputMode ?? (type === 'tel' ? 'tel' : undefined)}
      className={(className ? className + ' ' : '') + 'block resize-none overflow-x-auto overflow-y-hidden whitespace-nowrap'}
      onKeyDown={handleKeyDown}
      onChange={handleChange}
    />
  );
});

export default PlainTextInput;
