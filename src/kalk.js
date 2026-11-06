// Kalkulator (OEM-Kalkulation) – Darstellung von Kalkulation, Versionen und Protokoll für Vertriebskarte und Board.
// Die Regeln (Anfordern, Freigeben, Zurückgeben, Versionen, Benachrichtigungen) stehen in freigabe.js.
// Der Kalkulator liegt unter kalkulator/ (Build der Browser-App mit dem geprüften Rechenkern „Excel V15“) und wird
// immer als eigenes Fenster geöffnet – für die schmale Outlook-Seitenleiste ist er zu breit.

import { kalkulation, aktuelleVersion, statusText } from './freigabe.js';

export { kalkulation, aktuelleVersion, statusText };

/** Adresse des Kalkulators; mit Deal wird er für diesen Deal geöffnet. */
export function kalkUrl(deal) {
  const base = new URL('kalkulator/index.html', location.href).toString();
  if (!deal) return base + '#/calc';
  const q = new URLSearchParams({ deal: deal.id, pid: deal.pid || '', projekt: deal.name || '', kunde: deal.company || '' });
  return `${base}#/calc?${q}`;
}

const eur = (v, d = 2) => new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR', minimumFractionDigits: d, maximumFractionDigits: d }).format(v);
const int = (v) => new Intl.NumberFormat('de-DE').format(v);
export const zeit = (iso) => (iso ? new Date(iso).toLocaleString('de-DE', { dateStyle: 'short', timeStyle: 'short' }) : '');

/** Kurzfassung, z. B. „3 Staffeln · VK 1,85–2,40 € · DB 24–31 %“ */
export function staffelSummary(staffeln) {
  if (!staffeln?.length) return '';
  const vk = staffeln.map((s) => s.vkStk); const db = staffeln.map((s) => s.dbPct);
  const range = (a, f) => { const lo = Math.min(...a), hi = Math.max(...a); return lo === hi ? f(lo) : `${f(lo)}–${f(hi)}`; };
  const n = staffeln.length;
  return `${n} ${n === 1 ? 'Staffel' : 'Staffeln'} · VK ${range(vk, (v) => eur(v).replace(/\s?€$/, ''))} € · DB ${range(db, (v) => Math.round(v))} %`;
}

/** Staffel-Tabelle */
export function staffelTable(staffeln, esc) {
  if (!staffeln?.length) return '';
  return `<table class="kalk-tbl"><thead><tr><th>Menge</th><th>SK/Stk</th><th>VK/Stk</th><th>DB</th></tr></thead><tbody>${
    staffeln.map((s) => `<tr title="Versand: ${esc(s.versand || '')}"><td>${int(s.qty)}</td><td>${eur(s.skStk, 3)}</td><td><b>${eur(s.vkStk, 3)}</b></td><td>${Math.round(s.dbPct)} %</td></tr>`).join('')
  }</tbody></table>`;
}

/** Eine Version als Zeile mit aufklappbaren Staffeln */
export function versionHtml(v, esc, offen = false) {
  const wer = v.art === 'selbst' ? `selbst freigegeben von ${v.freigegebenVon}` : `kalkuliert von ${v.kalkuliertVon}, freigegeben von ${v.freigegebenVon}`;
  return `<details class="kalk-ver"${offen ? ' open' : ''}><summary><b>Version ${v.nr}</b> · ${esc(zeit(v.at))} · ${esc(wer)}${v.preiseAngepasst ? ' · Preise angepasst' : ''}</summary>
    ${staffelTable(v.staffeln, esc)}
    <div class="kalk-elwis">${v.elwis?.geschrieben ? 'In Elwis' : 'Elwis (Demo: nicht geschrieben)'}: ${esc(v.elwis?.angebot || '')} – ${v.elwis?.positionen?.length || 0} Position(en)</div>
    ${v.korrekturen?.length ? `<div class="kalk-k" title="Korrigierte Fehler der Excel-Vorlage">Abweichend von der Excel-Vorlage (${esc(v.korrekturen.join(', '))})</div>` : ''}
  </details>`;
}

/** Protokoll (neueste zuerst) */
export function logHtml(k, esc) {
  if (!k?.log?.length) return '';
  return `<details class="kalk-log"><summary>Protokoll (${k.log.length})</summary><ul>${
    [...k.log].reverse().map((e) => `<li><span class="t">${esc(zeit(e.at))}</span> <b>${esc(e.wer)}</b>: ${esc(e.was)}${e.an ? ` → ${esc(e.an)}` : ''}${e.version ? ` (Version ${e.version})` : ''}${e.kommentar ? ` – „${esc(e.kommentar)}“` : ''}</li>`).join('')
  }</ul></details>`;
}
