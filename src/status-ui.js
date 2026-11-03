// Gemeinsame Bausteine für Status-Anzeige und -Auswahl (Add-in und Board)
import { FIELD, colorOf } from './orders.js';

export const escHtml = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const fmtDate = (iso) => { if (!iso) return ''; const [y, m, d] = iso.split('-'); return `${d}.${m}.${y.slice(2)}`; };

export function statusChip(key, value, { full = false } = {}) {
  const f = FIELD[key];
  if (f?.type === 'date') {
    return value ? `<span class="st-date">${fmtDate(value)}</span>` : '<span class="st-empty">–</span>';
  }
  if (!value) return `<span class="st-chip st-none${full ? ' full' : ''}">–</span>`;
  const c = colorOf(key, value) || '#9aa1ab';
  return `<span class="st-chip${full ? ' full' : ''}" style="background:${c}" title="${escHtml(value)}">${escHtml(value)}</span>`;
}

let openPicker = null;
let openedAt = 0;
export function closePicker() { if (openPicker) { openPicker.remove(); openPicker = null; } }

// Öffnet eine Auswahl unter dem angeklickten Element. onPick(value) wird mit dem neuen Wert aufgerufen.
export function showPicker(anchor, key, current, onPick) {
  closePicker();
  const f = FIELD[key];
  const el = document.createElement('div');
  el.className = 'st-picker';
  el.setAttribute('role', 'dialog');
  el.innerHTML = `<div class="st-picker-title">${escHtml(f.label)}</div>` + (f.type === 'date'
    ? `<input type="date" class="st-date-input" value="${escHtml(current || '')}">
       <div class="st-picker-row"><button class="st-btn" data-clear>Leeren</button><button class="st-btn primary" data-ok>Übernehmen</button></div>`
    : f.options.map(([l, c]) => `<button class="st-opt${l === current ? ' sel' : ''}" data-v="${escHtml(l)}" style="background:${c}">${escHtml(l)}</button>`).join('') +
      `<button class="st-opt st-clear" data-v="">– leeren –</button>`);
  document.body.appendChild(el);

  const r = anchor.getBoundingClientRect();
  const w = Math.max(190, Math.min(240, window.innerWidth - 16));
  el.style.width = w + 'px';
  let left = Math.min(Math.max(8, r.left), window.innerWidth - w - 8);
  let top = r.bottom + 4;
  const h = el.offsetHeight;
  if (top + h > window.innerHeight - 8) top = Math.max(8, r.top - h - 4);
  el.style.left = left + 'px';
  el.style.top = top + 'px';
  openPicker = el;
  openedAt = Date.now();

  const done = (v) => { closePicker(); if (v !== current) onPick(v); };
  el.querySelectorAll('[data-v]').forEach((b) => b.addEventListener('click', () => done(b.dataset.v)));
  el.querySelector('[data-ok]')?.addEventListener('click', () => done(el.querySelector('.st-date-input').value));
  el.querySelector('[data-clear]')?.addEventListener('click', () => done(''));
  el.querySelector('.st-date-input')?.focus();
  setTimeout(() => {
    const away = (e) => { if (openPicker && !openPicker.contains(e.target)) { closePicker(); document.removeEventListener('mousedown', away, true); } };
    document.addEventListener('mousedown', away, true);
  }, 0);
}
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closePicker(); });
// Beim Scrollen schließen – aber nicht durch das Scrollen, das beim Öffnen selbst entsteht
window.addEventListener('scroll', (e) => { if (openPicker && !openPicker.contains(e.target) && Date.now() - openedAt > 400) closePicker(); }, true);
