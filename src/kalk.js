// Kalkulator (OEM-Kalkulation) – Verbindung zwischen Vertrieb und Kalkulationstool.
// Der Kalkulator liegt unter kalkulator/ (eigener Build der Browser-App mit dem geprüften Rechenkern „Excel V15“).
// Er wird aus der Vertriebskarte oder dem Board mit dem Deal geöffnet und übergibt sein Ergebnis je Staffel zurück.
// Demo: Übergabe über den Browser-Speicher (gleicher Ursprung). Live: Angebotspositionen am Elwis-Angebot.

export const LS_KALK = 'elwis-demo-kalkulationen-v1';

export function kalkFor(dealId) {
  try { return JSON.parse(localStorage.getItem(LS_KALK) || '{}')[dealId] || null; } catch { return null; }
}

export function kalkReset() {
  try { localStorage.removeItem(LS_KALK); } catch { /* ohne Speicher */ }
}

/** Adresse des Kalkulators; mit Deal wird der Kalkulator für diesen Deal geöffnet. */
export function kalkUrl(deal) {
  const base = new URL('kalkulator/index.html', location.href).toString();
  if (!deal) return base + '#/calc';
  const q = new URLSearchParams({ deal: deal.id, pid: deal.pid || '', projekt: deal.name || '', kunde: deal.company || '' });
  return `${base}#/calc?${q}`;
}

const eur = (v, d = 2) => new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR', minimumFractionDigits: d, maximumFractionDigits: d }).format(v);

/** Kurzfassung für Karte und Board, z. B. „3 Staffeln · VK 1,85–2,40 € · DB 24–31 %“ */
export function kalkSummary(k) {
  if (!k?.staffeln?.length) return '';
  const vk = k.staffeln.map((s) => s.vkStk);
  const db = k.staffeln.map((s) => s.dbPct);
  const range = (a, f) => { const lo = Math.min(...a), hi = Math.max(...a); return lo === hi ? f(lo) : `${f(lo)}–${f(hi)}`; };
  const n = k.staffeln.length;
  return `${n} ${n === 1 ? 'Staffel' : 'Staffeln'} · VK ${range(vk, (v) => eur(v).replace(/\s?€$/, ''))} € · DB ${range(db, (v) => Math.round(v))} %`;
}

export function kalkDate(k) {
  if (!k?.at) return '';
  return new Date(k.at).toLocaleString('de-DE', { dateStyle: 'short', timeStyle: 'short' });
}

/** Staffel-Tabelle für die Vertriebskarte */
export function kalkTable(k, esc) {
  if (!k?.staffeln?.length) return '';
  const int = (v) => new Intl.NumberFormat('de-DE').format(v);
  return `<table class="kalk-tbl"><thead><tr><th>Menge</th><th>SK/Stk</th><th>VK/Stk</th><th>DB</th></tr></thead><tbody>${
    k.staffeln.map((s) => `<tr title="Versand: ${esc(s.versand || '')}"><td>${int(s.qty)}</td><td>${eur(s.skStk, 3)}</td><td><b>${eur(s.vkStk, 3)}</b></td><td>${Math.round(s.dbPct)} %</td></tr>`).join('')
  }</tbody></table>`;
}
