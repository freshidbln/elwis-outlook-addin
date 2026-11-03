// Board-Übersicht „Auftragsabwicklung“ – ersetzt das Monday-Board.
// Datenquelle ist ausschließlich Elwis (Demo: Beispieldaten mit Browser-Speicher).
import { CONFIG } from './config.js';
import { provider } from './data.js';
import { STATUS_FIELDS, FIELD, PHASES, PHASE, RULES, computePhase, colorOf } from './orders.js';
import { statusChip, showPicker, escHtml as esc } from './status-ui.js';

if (typeof window.Office === 'undefined') window.Office = { onReady: (cb) => setTimeout(() => cb({ host: null }), 0) };

const $ = (id) => document.getElementById(id);
const money = (v, cur = 'EUR') => v == null || v === '' ? '' : new Intl.NumberFormat('de-DE', { style: 'currency', currency: cur, maximumFractionDigits: 0 }).format(v);

const INFO_COLS = [
  { key: 'pm', label: 'PM', w: 120 },
  { key: 'seller', label: 'Verkäufer', w: 120 },
  { key: 'contact', label: 'Kontakt', w: 130 },
  { key: 'company', label: 'Unternehmen', w: 180 },
  { key: 'sc', label: 'SC-Nr.', w: 80 },
  { key: 'amount', label: 'Auftragswert', w: 105, align: 'right' }
];
const STATUS_W = (f) => f.type === 'date' ? 104 : 138;

let orders = [];
const collapsed = new Set(['abruf']);
let flashId = null;

Office.onReady(() => setTimeout(init, 0));

async function init() {
  $('modeBadge').textContent = provider.isDemo ? 'Demo' : 'Live';
  $('resetBtn').hidden = !provider.isDemo;
  $('q').addEventListener('input', render);
  $('pmFilter').addEventListener('change', render);
  $('officeFilter').addEventListener('change', render);
  $('rulesBtn').addEventListener('click', toggleRules);
  $('rulesClose').addEventListener('click', toggleRules);
  $('resetBtn').addEventListener('click', async () => { provider.resetOrders(); await load(); toast('Demo zurückgesetzt.'); });
  window.addEventListener('storage', (e) => { if (e.key && e.key.startsWith('elwis-demo-orders')) load(); });
  await load();
  const pms = [...new Set(orders.map((o) => o.pm).filter(Boolean))].sort();
  $('pmFilter').innerHTML = '<option value="">Alle PMs</option>' + pms.map((p) => `<option>${esc(p)}</option>`).join('');
}

async function load() {
  orders = await provider.listOrders();
  render();
}

function filtered() {
  const q = $('q').value.trim().toLowerCase();
  const pm = $('pmFilter').value;
  const office = $('officeFilter').value;
  return orders.filter((o) =>
    (!pm || o.pm === pm) && (!office || o.office === office) &&
    (!q || [o.name, o.company, o.contact, o.sc, o.pid, o.number].join(' ').toLowerCase().includes(q)));
}

function render() {
  const list = filtered();
  const byPhase = Object.fromEntries(PHASES.map((p) => [p.key, []]));
  list.forEach((o) => byPhase[computePhase(o)].push(o));
  const totalW = 260 + INFO_COLS.reduce((s, c) => s + c.w, 0) + STATUS_FIELDS.reduce((s, f) => s + STATUS_W(f), 0) + 44;

  $('board').innerHTML = PHASES.map((ph) => {
    const rows = byPhase[ph.key];
    const isCol = collapsed.has(ph.key);
    return `<section class="grp" style="--gc:${ph.color}">
      <button class="grp-head" data-toggle="${ph.key}" aria-expanded="${!isCol}">
        <span class="caret">${isCol ? '▸' : '▾'}</span><span class="grp-title">${esc(ph.label)}</span>
        <span class="grp-count">${rows.length} ${rows.length === 1 ? 'Auftrag' : 'Aufträge'}</span>
      </button>
      ${isCol ? '' : `<div class="tbl-wrap"><table class="tbl" style="width:${totalW}px">
        <colgroup><col style="width:260px">${INFO_COLS.map((c) => `<col style="width:${c.w}px">`).join('')}${STATUS_FIELDS.map((f) => `<col style="width:${STATUS_W(f)}px">`).join('')}<col style="width:44px"></colgroup>
        <thead><tr>
          <th class="sticky">Auftrag</th>
          ${INFO_COLS.map((c) => `<th class="ro" title="Kommt aus Elwis">${esc(c.label)}</th>`).join('')}
          ${STATUS_FIELDS.map((f) => `<th title="${esc(f.label)}">${esc(f.short || f.label)}</th>`).join('')}
          <th></th>
        </tr></thead>
        <tbody>${rows.length ? rows.map(renderRow).join('') : `<tr><td class="sticky empty" colspan="${2 + INFO_COLS.length + STATUS_FIELDS.length}">Keine Aufträge in dieser Phase</td></tr>`}</tbody>
        ${rows.length ? renderFooter(rows) : ''}
      </table></div>`}
    </section>`;
  }).join('') + `<div class="b-foot">${list.length} von ${orders.length} Aufträgen · Gruppen werden aus den Status berechnet · ${provider.isDemo ? 'Demo-Daten, Änderungen nur in diesem Browser' : 'Daten live aus Elwis'}</div>`;

  $('board').querySelectorAll('[data-toggle]').forEach((b) => b.addEventListener('click', () => {
    const k = b.dataset.toggle; collapsed.has(k) ? collapsed.delete(k) : collapsed.add(k); render();
  }));
  $('board').querySelectorAll('[data-cell]').forEach((td) => td.addEventListener('click', () => {
    const [id, key] = td.dataset.cell.split('|');
    const o = orders.find((x) => x.id === id);
    showPicker(td, key, o[key] || '', (val) => change(id, key, val));
  }));
  $('board').querySelectorAll('[data-open]').forEach((b) => b.addEventListener('click', () => openInElwis(orders.find((x) => x.id === b.dataset.open))));
  if (flashId) {
    const tr = document.querySelector(`tr[data-id="${flashId}"]`);
    if (tr) { tr.classList.add('flash'); tr.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); }
    flashId = null;
  }
}

function renderRow(o) {
  return `<tr data-id="${o.id}">
    <td class="sticky name">
      <div class="n-main">${esc(o.name)}</div>
      <div class="n-sub">${o.pid ? esc(o.pid) + ' · ' : ''}${esc(o.office || '')}${o.okayToInvoice ? ' · <span class="okinv">Okay to Invoice</span>' : ''}</div>
    </td>
    ${INFO_COLS.map((c) => `<td class="ro${c.align === 'right' ? ' right' : ''}">${c.key === 'amount' ? money(o.amount, o.currency) : esc(o[c.key] || '')}</td>`).join('')}
    ${STATUS_FIELDS.map((f) => `<td class="st${f.type === 'date' ? ' dt' : ''}${f.type === 'date' && isOverdue(o, f.key) ? ' late' : ''}" data-cell="${o.id}|${f.key}" title="${esc(f.label)}">${statusChip(f.key, o[f.key], { full: f.type !== 'date' })}</td>`).join('')}
    <td class="ro"><button class="open-btn" data-open="${o.id}" title="Auftrag in Elwis öffnen">↗</button></td>
  </tr>`;
}

function isOverdue(o, key) {
  const v = o[key]; if (!v) return false;
  const t = new Date(); const today = `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
  if (v >= today) return false;
  if (key === 'korriDate') return ['warte auf Freigabe', 'Verzögert', 'Angefordert'].includes(o.korri);
  if (key === 'musterDate') return ['warte auf Freigabe', 'verzögert'].includes(o.muster);
  if (key === 'fertig' || key === 'ltLief') return o.prod !== 'Fertiggestellt';
  if (key === 'eta') return !['Angekommen'].includes(o.liefer) && o.lieferung !== 'Ausgeliefert';
  if (key === 'ltKunde') return o.lieferung !== 'Ausgeliefert' && o.liefer !== 'Angekommen' && o.lieferung !== 'Fertig zum Abrechnen';
  return false;
}

function renderFooter(rows) {
  const sums = {};
  rows.forEach((o) => { if (o.amount) sums[o.currency || 'EUR'] = (sums[o.currency || 'EUR'] || 0) + o.amount; });
  const sumTxt = Object.entries(sums).map(([c, v]) => money(v, c)).join('<br>');
  const bar = (f) => {
    if (f.type === 'date') return '';
    const counts = {}; rows.forEach((o) => { const v = o[f.key] || ''; counts[v] = (counts[v] || 0) + 1; });
    return `<div class="dist" title="${esc(Object.entries(counts).map(([k, n]) => `${k || 'leer'}: ${n}`).join(', '))}">` +
      Object.entries(counts).map(([v, n]) => `<span style="flex:${n};background:${v ? colorOf(f.key, v) : '#d9dee5'}"></span>`).join('') + '</div>';
  };
  return `<tfoot><tr>
    <td class="sticky"></td>
    ${INFO_COLS.map((c) => `<td class="${c.key === 'amount' ? 'right sum' : ''}">${c.key === 'amount' ? `${sumTxt}<div class="sum-l">Summe</div>` : ''}</td>`).join('')}
    ${STATUS_FIELDS.map((f) => `<td>${bar(f)}</td>`).join('')}<td></td>
  </tr></tfoot>`;
}

async function change(id, key, val) {
  const before = computePhase(orders.find((o) => o.id === id));
  const r = await provider.updateOrder(id, key, val);
  orders = orders.map((o) => (o.id === id ? r.order : o));
  if (computePhase(r.order) !== before) { flashId = id; collapsed.delete(computePhase(r.order)); }
  render();
  toast(r.applied.length ? `Gespeichert. ${r.applied.join(' · ')}` : 'Gespeichert.', true);
}

function openInElwis(o) {
  if (!o) return;
  const kind = 'order';
  const url = provider.isDemo
    ? `${CONFIG.elwisBaseUrl}/${CONFIG.listPages[o.office || 'NBG'][kind]}`
    : `${CONFIG.elwisBaseUrl}/${CONFIG.detailPages[o.office || 'NBG'][kind]}?artifactId=${encodeURIComponent(o.id)}`;
  if (provider.isDemo) toast('Demo: Es öffnet sich die Auftragsliste in Elwis. Live direkt der Auftrag.');
  window.open(url, '_blank', 'noopener');
}

function toggleRules() {
  const r = $('rules');
  if (r.hidden) {
    $('rulesBody').innerHTML = `
      <p class="muted">Diese Regeln ersetzen die 20 Monday-Automationen. Live laufen sie als Server-Abläufe in Elwis – egal, ob der Status in Elwis, im Outlook-Add-in oder in dieser Übersicht geändert wird.</p>
      <h4>Bei Statusänderung</h4>
      ${RULES.filter((x) => x.id.startsWith('R')).map((x) => `<div class="rule"><b>${x.id}</b><div><div>Wenn ${esc(x.when)}</div><div class="then">→ ${esc(x.then)}</div></div></div>`).join('')}
      <h4>Täglich (Termine)</h4>
      ${RULES.filter((x) => x.id.startsWith('T')).map((x) => `<div class="rule"><b>${x.id}</b><div><div>${esc(x.when.replace('Täglich: ', ''))}</div><div class="then">→ ${esc(x.then)}</div></div></div>`).join('')}
      <h4>Gruppen (Phasen)</h4>
      <p class="muted">Werden automatisch aus den Status berechnet – niemand muss Aufträge verschieben.</p>
      <h4>Felder am Elwis-Auftrag</h4>
      <table class="map"><thead><tr><th>Feld</th><th>In Elwis</th></tr></thead><tbody>
        ${STATUS_FIELDS.map((f) => `<tr><td>${esc(f.label)}</td><td>${f.elwis ? `vorhanden: ${esc(f.elwis)}` : '<span class="neu">neu anlegen</span>'}</td></tr>`).join('')}
        <tr><td>PM, Verkäufer, Kontakt, Unternehmen, SC-Nr., Auftragswert</td><td>vorhanden (werden nur angezeigt)</td></tr>
      </tbody></table>`;
  }
  r.hidden = !r.hidden;
}

let tt;
function toast(msg, ok = false) {
  const t = $('toast'); t.textContent = msg; t.classList.toggle('ok', ok); t.hidden = false;
  clearTimeout(tt); tt = setTimeout(() => { t.hidden = true; }, 5000);
}
