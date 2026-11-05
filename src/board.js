// Board-Übersicht – ersetzt die Monday-Boards „Anfrage & Vertrieb“ und „Auftragsabwicklung“.
// Datenquelle ist ausschließlich Elwis (Demo: Beispieldaten mit Browser-Speicher).
import { CONFIG } from './config.js';
import { provider } from './data.js';
import { STATUS_FIELDS, PHASES, RULES, computePhase } from './orders.js';
import { SALES_FIELDS, SALES_PHASES, SALES_RULES, SALES_COMMENT, computeSalesPhase, overdueDays } from './sales.js';
import { statusChip, showPicker, showTextEditor, colorOf, escHtml as esc, fmtDate } from './status-ui.js';
import { kalkFor, kalkUrl, kalkSummary, kalkDate, kalkReset } from './kalk.js';

if (typeof window.Office === 'undefined') window.Office = { onReady: (cb) => setTimeout(() => cb({ host: null }), 0) };

const $ = (id) => document.getElementById(id);
const money = (v, cur = 'EUR') => v == null || v === '' ? '' : new Intl.NumberFormat('de-DE', { style: 'currency', currency: cur, maximumFractionDigits: 0 }).format(v);
const todayIso = () => { const t = new Date(); return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`; };

function orderOverdue(o, key) {
  const v = o[key]; if (!v || v >= todayIso()) return false;
  if (key === 'korriDate') return ['warte auf Freigabe', 'Verzögert', 'Angefordert'].includes(o.korri);
  if (key === 'musterDate') return ['warte auf Freigabe', 'verzögert'].includes(o.muster);
  if (key === 'fertig' || key === 'ltLief') return o.prod !== 'Fertiggestellt';
  if (key === 'eta') return o.liefer !== 'Angekommen' && o.lieferung !== 'Ausgeliefert';
  if (key === 'ltKunde') return !['Ausgeliefert', 'Fertig zum Abrechnen'].includes(o.lieferung) && o.liefer !== 'Angekommen';
  return false;
}

const VIEWS = {
  sales: {
    tab: 'Anfrage & Vertrieb', sub: 'vor dem Auftrag', unit: ['Deal', 'Deals'],
    fields: SALES_FIELDS, phases: SALES_PHASES, phaseOf: computeSalesPhase, rules: SALES_RULES, textCol: { ...SALES_COMMENT, w: 260 },
    collapsed: ['verloren', 'uebergeben'], person: 'seller', personLabel: 'Alle Verkäufer',
    infoCols: [
      { key: 'seller', label: 'Verkäufer', w: 120 }, { key: 'contact', label: 'Kontakt', w: 130 },
      { key: 'company', label: 'Unternehmen', w: 180 }, { key: 'pid', label: 'PID', w: 112 },
      { key: 'amount', label: 'Wert', w: 100, align: 'right' }, { key: 'received', label: 'Anfrage vom', w: 96, date: true }
    ],
    list: () => provider.listDeals(),
    update: async (id, k, v) => { const r = await provider.updateDeal(id, k, v); return { item: r.deal, applied: r.applied, created: r.createdOrder }; },
    overdue: (d, key) => key === 'entscheidung' && overdueDays(d, todayIso()) > 0,
    badge: (d) => { const n = overdueDays(d, todayIso()); const k = kalkFor(d.id);
      return (n ? `<span class="od">überfällig ${n} T</span>` : (d.reactivated ? '<span class="od re">Wiedervorlage</span>' : ''))
        + (k ? `<button class="kalk-chip" data-kalk="${d.id}" title="${esc(kalkSummary(k) + ' · ' + kalkDate(k))}">kalkuliert</button>` : ''); },
    action: (d) => `<button class="open-btn" data-kalk="${d.id}" title="Kalkulation öffnen">∑</button>`,
    elwisKind: 'inquiry'
  },
  orders: {
    tab: 'Auftragsabwicklung', sub: 'nach Auftragseingang', unit: ['Auftrag', 'Aufträge'],
    fields: STATUS_FIELDS, phases: PHASES, phaseOf: computePhase, rules: RULES,
    collapsed: ['abruf'], person: 'pm', personLabel: 'Alle PMs',
    infoCols: [
      { key: 'pm', label: 'PM', w: 120 }, { key: 'seller', label: 'Verkäufer', w: 120 }, { key: 'contact', label: 'Kontakt', w: 130 },
      { key: 'company', label: 'Unternehmen', w: 180 }, { key: 'pid', label: 'PID', w: 112 }, { key: 'sc', label: 'SC-Nr.', w: 80 }, { key: 'amount', label: 'Auftragswert', w: 105, align: 'right' }
    ],
    list: () => provider.listOrders(),
    update: async (id, k, v) => { const r = await provider.updateOrder(id, k, v); return { item: r.order, applied: r.applied }; },
    overdue: orderOverdue,
    badge: (o) => (o.okayToInvoice ? '<span class="okinv">Okay to Invoice</span>' : '') + (o.handoverNote ? `<span class="od re" title="${esc(o.handoverNote)}">Übergabe-Notiz</span>` : ''),
    elwisKind: 'order'
  }
};

let viewKey = new URLSearchParams(location.search).get('view') || (() => { try { return localStorage.getItem('elwis-board-view'); } catch { return null; } })() || 'sales';
if (!VIEWS[viewKey]) viewKey = 'sales';
let items = [];
let collapsed = new Set(VIEWS[viewKey].collapsed);
let flashId = null;
const V = () => VIEWS[viewKey];
const statusW = (f) => (f.type === 'date' ? 104 : 140);

Office.onReady(() => setTimeout(init, 0));

async function init() {
  $('modeBadge').textContent = provider.isDemo ? 'Demo' : 'Live';
  $('resetBtn').hidden = !provider.isDemo;
  $('q').addEventListener('input', render);
  $('personFilter').addEventListener('change', render);
  $('officeFilter').addEventListener('change', render);
  $('rulesBtn').addEventListener('click', toggleRules);
  $('rulesClose').addEventListener('click', toggleRules);
  $('resetBtn').addEventListener('click', async () => { provider.resetOrders(); kalkReset(); await load(); toast('Demo zurückgesetzt.'); });
  $('kalkBtn').addEventListener('click', () => window.open(kalkUrl(null), '_blank'));
  document.querySelectorAll('[data-view]').forEach((b) => b.addEventListener('click', () => switchView(b.dataset.view)));
  window.addEventListener('storage', (e) => { if (e.key && e.key.startsWith('elwis-demo-')) load(); });
  await switchView(viewKey);
}

async function switchView(k) {
  viewKey = k;
  try { localStorage.setItem('elwis-board-view', k); } catch { /* egal */ }
  collapsed = new Set(V().collapsed);
  document.querySelectorAll('[data-view]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.view === k)));
  document.title = `${V().tab} – Elwis`;
  $('q').placeholder = k === 'sales' ? 'Deal, Kunde oder PID suchen' : 'Auftrag, Kunde, PID oder SC-Nr. suchen';
  $('rules').hidden = true;
  await load();
  const people = [...new Set(items.flatMap((o) => splitPeople(o[V().person])))].sort();
  $('personFilter').innerHTML = `<option value="">${V().personLabel}</option>` + people.map((p) => `<option>${esc(p)}</option>`).join('');
  render();
}

async function load() { items = await V().list(); render(); }
const splitPeople = (s) => String(s || '').split(/\s*,\s*/).filter(Boolean);

function filtered() {
  const q = $('q').value.trim().toLowerCase();
  const person = $('personFilter').value;
  const office = $('officeFilter').value;
  return items.filter((o) =>
    (!person || splitPeople(o[V().person]).includes(person)) && (!office || o.office === office) &&
    (!q || [o.name, o.company, o.contact, o.sc, o.pid, o.number, o.an, o.comment].join(' ').toLowerCase().includes(q)));
}

function render() {
  const v = V();
  const list = filtered();
  const byPhase = Object.fromEntries(v.phases.map((p) => [p.key, []]));
  list.forEach((o) => byPhase[v.phaseOf(o)].push(o));
  const tw = v.textCol ? v.textCol.w : 0;
  const aw = v.action ? 76 : 44; // Spalte für Aktionen (Vertrieb: Kalkulation + Elwis)
  const totalW = 260 + v.infoCols.reduce((s, c) => s + c.w, 0) + v.fields.reduce((s, f) => s + statusW(f), 0) + tw + aw;

  $('board').innerHTML = v.phases.map((ph) => {
    const rows = byPhase[ph.key];
    const isCol = collapsed.has(ph.key);
    return `<section class="grp" style="--gc:${ph.color}">
      <button class="grp-head" data-toggle="${ph.key}" aria-expanded="${!isCol}">
        <span class="caret">${isCol ? '▸' : '▾'}</span><span class="grp-title">${esc(ph.label)}</span>
        <span class="grp-count">${rows.length} ${rows.length === 1 ? v.unit[0] : v.unit[1]}${isCol ? ' · eingeklappt' : ''}</span>
      </button>
      ${isCol ? '' : `<div class="tbl-wrap"><table class="tbl" style="width:${totalW}px">
        <colgroup><col style="width:260px">${v.infoCols.map((c) => `<col style="width:${c.w}px">`).join('')}${v.fields.map((f) => `<col style="width:${statusW(f)}px">`).join('')}${tw ? `<col style="width:${tw}px">` : ''}<col style="width:${aw}px"></colgroup>
        <thead><tr>
          <th class="sticky">${viewKey === 'sales' ? 'Deal' : 'Auftrag'}</th>
          ${v.infoCols.map((c) => `<th class="ro" title="Kommt aus Elwis">${esc(c.label)}</th>`).join('')}
          ${v.fields.map((f) => `<th title="${esc(f.label)}">${esc(f.short || f.label)}</th>`).join('')}
          ${tw ? `<th title="${esc(v.textCol.label)}">${esc(v.textCol.label)}</th>` : ''}
          <th></th>
        </tr></thead>
        <tbody>${rows.length ? rows.map(renderRow).join('') : `<tr><td class="sticky empty" colspan="${2 + v.infoCols.length + v.fields.length + (tw ? 1 : 0)}">Keine Einträge in dieser Gruppe</td></tr>`}</tbody>
        ${rows.length ? renderFooter(rows) : ''}
      </table></div>`}
    </section>`;
  }).join('') + `<div class="b-foot">${list.length} von ${items.length} ${v.unit[1]} · Gruppen werden aus den Status berechnet · ${provider.isDemo ? 'Demo-Daten, Änderungen nur in diesem Browser' : 'Daten live aus Elwis'}</div>`;

  $('board').querySelectorAll('[data-toggle]').forEach((b) => b.addEventListener('click', () => {
    const k = b.dataset.toggle; collapsed.has(k) ? collapsed.delete(k) : collapsed.add(k); render();
  }));
  $('board').querySelectorAll('[data-cell]').forEach((td) => td.addEventListener('click', () => {
    const [id, key] = td.dataset.cell.split('|');
    const o = items.find((x) => x.id === id);
    showPicker(td, key, o[key] || '', (val) => change(id, key, val));
  }));
  $('board').querySelectorAll('[data-text]').forEach((td) => td.addEventListener('click', () => {
    const id = td.dataset.text; const o = items.find((x) => x.id === id); const k = V().textCol.key;
    showTextEditor(td, V().textCol.label, o[k] || '', (val) => { if (val !== (o[k] || '')) change(id, k, val); });
  }));
  $('board').querySelectorAll('[data-open]').forEach((b) => b.addEventListener('click', () => openInElwis(items.find((x) => x.id === b.dataset.open))));
  $('board').querySelectorAll('[data-kalk]').forEach((b) => b.addEventListener('click', (e) => { e.stopPropagation(); window.open(kalkUrl(items.find((x) => x.id === b.dataset.kalk)), '_blank'); }));
  if (flashId) {
    const tr = document.querySelector(`tr[data-id="${flashId}"]`);
    if (tr) { tr.classList.add('flash'); tr.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); }
    flashId = null;
  }
}

function renderRow(o) {
  const v = V();
  return `<tr data-id="${o.id}">
    <td class="sticky name">
      <div class="n-main">${esc(o.name)}</div>
      <div class="n-sub">${esc(o.office || '')} ${v.badge(o)}</div>
    </td>
    ${v.infoCols.map((c) => `<td class="ro${c.align === 'right' ? ' right' : ''}">${c.key === 'amount' ? money(o.amount, o.currency) : c.date ? fmtDate(o[c.key]) : (c.key === 'sc' && !o.sc && o.fromDeal) ? '<span class="neu-sc" title="SC-Nummer vergibt Elwis beim Anlegen der Sales Order">neu</span>' : esc(o[c.key] || '')}</td>`).join('')}
    ${v.fields.map((f) => `<td class="st${f.type === 'date' ? ' dt' : ''}${f.type === 'date' && v.overdue(o, f.key) ? ' late' : ''}" data-cell="${o.id}|${f.key}" title="${esc(f.label)}">${statusChip(f.key, o[f.key], { full: f.type !== 'date' })}</td>`).join('')}
    ${v.textCol ? `<td class="txt" data-text="${o.id}" title="${esc(o[v.textCol.key] || 'Klicken zum Schreiben')}">${o[v.textCol.key] ? esc(o[v.textCol.key]) : '<span class="txt-empty">+ Kommentar</span>'}</td>` : ''}
    <td class="ro act">${v.action ? v.action(o) : ''}<button class="open-btn" data-open="${o.id}" title="In Elwis öffnen">↗</button></td>
  </tr>`;
}

function renderFooter(rows) {
  const v = V();
  const sums = {};
  rows.forEach((o) => { if (o.amount) sums[o.currency || 'EUR'] = (sums[o.currency || 'EUR'] || 0) + o.amount; });
  const sumTxt = Object.entries(sums).map(([c, n]) => money(n, c)).join('<br>');
  const bar = (f) => {
    if (f.type === 'date') return '';
    const counts = {}; rows.forEach((o) => { const val = o[f.key] || ''; counts[val] = (counts[val] || 0) + 1; });
    return `<div class="dist" title="${esc(Object.entries(counts).map(([k, n]) => `${k || 'leer'}: ${n}`).join(', '))}">` +
      Object.entries(counts).map(([val, n]) => `<span style="flex:${n};background:${val ? colorOf(f.key, val) : '#d9dee5'}"></span>`).join('') + '</div>';
  };
  return `<tfoot><tr>
    <td class="sticky"></td>
    ${v.infoCols.map((c) => `<td class="${c.key === 'amount' ? 'right sum' : ''}">${c.key === 'amount' ? `${sumTxt}<div class="sum-l">Summe</div>` : ''}</td>`).join('')}
    ${v.fields.map((f) => `<td>${bar(f)}</td>`).join('')}${v.textCol ? '<td></td>' : ''}<td></td>
  </tr></tfoot>`;
}

async function change(id, key, val) {
  const v = V();
  const before = v.phaseOf(items.find((o) => o.id === id));
  let r;
  try { r = await v.update(id, key, val); } catch (err) { toast(err.message); return; }
  items = items.map((o) => (o.id === id ? r.item : o));
  const after = v.phaseOf(r.item);
  if (after !== before) { flashId = id; collapsed.delete(after); }
  render();
  let msg = r.applied.length ? `Gespeichert. ${r.applied.join(' · ')}` : 'Gespeichert.';
  if (r.created) msg += ' → In „Auftragsabwicklung“ wechseln, um den neuen Auftrag zu sehen.';
  toast(msg, true);
}

function openInElwis(o) {
  if (!o) return;
  const kind = V().elwisKind;
  const office = o.office || 'NBG';
  const url = provider.isDemo
    ? `${CONFIG.elwisBaseUrl}/${CONFIG.listPages[office][kind]}`
    : `${CONFIG.elwisBaseUrl}/${CONFIG.detailPages[office][kind]}?artifactId=${encodeURIComponent(o.projectId || o.id)}`;
  if (provider.isDemo) toast('Demo: Es öffnet sich die passende Liste in Elwis. Live direkt der Datensatz.');
  window.open(url, '_blank', 'noopener');
}

function toggleRules() {
  const r = $('rules');
  const v = V();
  if (r.hidden) {
    const change = v.rules.filter((x) => !x.id.startsWith('T'));
    const daily = v.rules.filter((x) => x.id.startsWith('T'));
    const fieldRows = viewKey === 'sales'
      ? SALES_FIELDS.map((f) => `<tr><td>${esc(f.label)}</td><td>${esc(f.elwis || '')}</td></tr>`).join('') +
        `<tr><td colspan="2" class="sub-h">Werte des Vertriebsstatus</td></tr>` +
        SALES_FIELDS[0].options.map(([l, , m]) => `<tr><td>${esc(l)}</td><td>${/^neu/.test(m) ? `<span class="neu">${esc(m)}</span>` : esc(m)}</td></tr>`).join('') +
        `<tr><td>${esc(SALES_COMMENT.label)}</td><td>${esc(SALES_COMMENT.elwis)}</td></tr>` +
        `<tr><td>PID – durchgehende Nummer von der Anfrage bis zur Rechnung</td><td>vorhanden (Project ID); muss bei jeder Anfrage automatisch vergeben werden</td></tr>` +
        `<tr><td>Verkäufer (auch mehrere), Kontakt, Unternehmen, Wert, Anfrage vom</td><td>vorhanden (werden nur angezeigt)</td></tr>`
      : STATUS_FIELDS.map((f) => `<tr><td>${esc(f.label)}</td><td>${f.elwis ? `vorhanden: ${esc(f.elwis)}` : '<span class="neu">neu anlegen</span>'}</td></tr>`).join('') +
        `<tr><td>PM, Verkäufer, Kontakt, Unternehmen, SC-Nr., Auftragswert</td><td>vorhanden (werden nur angezeigt)</td></tr>`;
    $('rulesBody').innerHTML = `
      <p class="muted">${viewKey === 'sales'
        ? 'Diese Regeln ersetzen die 22 Automationen des Monday-Boards „Anfrage & Vertrieb“.'
        : 'Diese Regeln ersetzen die 20 Automationen des Monday-Boards „Auftragsabwicklung“.'}
        Live laufen sie als Server-Abläufe in Elwis – egal, ob der Status in Elwis, im Outlook-Add-in oder in dieser Übersicht geändert wird.</p>
      ${viewKey === 'sales' ? '<p class="muted"><b>Wichtig:</b> „Überfällig“ ist hier ein Hinweis und kein Status. In Monday hat „Überfällig“ den eigentlichen Status überschrieben – dadurch ging verloren, ob ein Angebot beim Kunden lag oder noch in Verhandlung war.</p>' : ''}
      <h4>Bei Statusänderung</h4>
      ${change.map((x) => `<div class="rule"><b>${x.id}</b><div><div>Wenn ${esc(x.when)}</div><div class="then">→ ${esc(x.then)}</div></div></div>`).join('')}
      <h4>Täglich (Termine)</h4>
      ${daily.map((x) => `<div class="rule"><b>${x.id}</b><div><div>${esc(x.when.replace('Täglich: ', ''))}</div><div class="then">→ ${esc(x.then)}</div></div></div>`).join('')}
      <h4>Gruppen</h4>
      <p class="muted">Werden automatisch aus den Status berechnet – niemand muss Einträge verschieben.</p>
      <h4>Felder in Elwis</h4>
      <table class="map"><thead><tr><th>Feld</th><th>In Elwis</th></tr></thead><tbody>${fieldRows}</tbody></table>`;
  }
  r.hidden = !r.hidden;
}

let tt;
function toast(msg, ok = false) {
  const t = $('toast'); t.textContent = msg; t.classList.toggle('ok', ok); t.hidden = false;
  clearTimeout(tt); tt = setTimeout(() => { t.hidden = true; }, 6000);
}
