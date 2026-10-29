import { CONFIG } from './config.js';
import { provider } from './data.js';
import { status, KIND_LABEL } from './status.js';
import { demoEmails, STAFF } from './mock-data.js';

// Falls office.js nicht geladen werden konnte (z. B. lokale Vorschau ohne Internet)
if (typeof window.Office === 'undefined') {
  window.Office = { onReady: (cb) => setTimeout(() => cb({ host: null }), 0), HostType: { Outlook: 'Outlook' } };
}

// ---------- Hilfsfunktionen ----------
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const money = (v, cur = 'EUR') => v == null ? '–' : new Intl.NumberFormat('de-DE', { style: 'currency', currency: cur, maximumFractionDigits: 0 }).format(v);
const dateDe = (iso) => { if (!iso) return '–'; const [y, m, d] = iso.split('-'); return `${d}.${m}.${y}`; };
const daysFromToday = (iso) => { const t = new Date(); t.setHours(0, 0, 0, 0); const [y, m, d] = iso.split('-').map(Number); return Math.round((new Date(y, m - 1, d) - t) / 86400000); };
const relDue = (iso) => {
  if (!iso) return '';
  const n = daysFromToday(iso);
  if (n < 0) return `<span class="warn-text">überfällig seit ${-n} ${-n === 1 ? 'Tag' : 'Tagen'}</span>`;
  if (n === 0) return '<span class="soon-text">heute fällig</span>';
  if (n <= 3) return `<span class="soon-text">in ${n} ${n === 1 ? 'Tag' : 'Tagen'}</span>`;
  return `in ${n} Tagen`;
};
const badge = (kind, code) => { const s = status(kind, code); return `<span class="badge" style="background:${s.color}">${esc(s.label)}</span>`; };
const initials = (name) => name.split(/\s+/).filter(Boolean).slice(0, 2).map((p) => p[0]).join('').toUpperCase();

let state = { mail: null, ctx: null, tab: 'projects', inOutlook: false };

// ---------- Start ----------
Office.onReady((info) => {
  state.inOutlook = info.host === Office.HostType.Outlook;
  $('modeBadge').textContent = provider.isDemo ? 'Demo' : 'Live';
  $('modeBadge').classList.toggle('live', !provider.isDemo);
  $('footText').textContent = provider.isDemo
    ? 'Demo-Version · Beispieldaten, kein Zugriff auf Elwis'
    : `Elwis Live · nur lesend · v${CONFIG.version}`;

  bindUi();

  if (state.inOutlook) {
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => loadFromOutlook());
    } catch (e) { /* Pinning nicht verfügbar */ }
    loadFromOutlook();
  } else {
    setupSimulator();
  }
});

function bindUi() {
  $('refreshBtn').addEventListener('click', () => state.inOutlook ? loadFromOutlook() : loadMail(state.mail));
  $('sheetBack').addEventListener('click', closeSheet);
  $('sheetBackdrop').addEventListener('click', closeSheet);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeSheet(); });

  let t;
  $('search').addEventListener('input', (e) => {
    clearTimeout(t);
    const q = e.target.value;
    t = setTimeout(() => q.trim().length >= 2 ? runSearch(q) : renderHome(), 220);
  });
}

// ---------- Outlook lesen ----------
async function loadFromOutlook() {
  const item = Office.context.mailbox.item;
  if (!item) { renderEmpty('Keine E-Mail ausgewählt', 'Wähle eine E-Mail aus, um passende Elwis-Daten zu sehen.'); return; }
  try {
    let mail;
    if (typeof item.subject === 'string') {
      // Lesemodus
      mail = { mode: 'read', fromName: item.from?.displayName || '', fromEmail: item.from?.emailAddress || '', subject: item.subject || '', emails: [item.from?.emailAddress].filter(Boolean) };
    } else {
      // Verfassen: Empfänger und Betreff asynchron lesen
      const subject = await asyncGet(item.subject);
      const to = await asyncGet(item.to);
      const cc = item.cc ? await asyncGet(item.cc) : [];
      const all = [...(to || []), ...(cc || [])];
      mail = { mode: 'compose', fromName: all[0]?.displayName || '', fromEmail: all[0]?.emailAddress || '', subject: subject || '', emails: all.map((r) => r.emailAddress) };
    }
    loadMail(mail);
  } catch (e) {
    renderError(e);
  }
}

function asyncGet(prop) {
  return new Promise((resolve, reject) => prop.getAsync((r) => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)));
}

// ---------- Vorschau ohne Outlook ----------
function setupSimulator() {
  const sel = $('simSelect');
  sel.innerHTML = demoEmails.map((m, i) => `<option value="${i}">${esc(m.label)}</option>`).join('');
  sel.addEventListener('change', () => loadMail(toMail(demoEmails[sel.value])));
  $('simulator').hidden = false;
  loadMail(toMail(demoEmails[0]));
}
const toMail = (m) => ({ mode: 'read', fromName: m.fromName, fromEmail: m.fromEmail, subject: m.subject, emails: [m.fromEmail] });

// ---------- Laden und Darstellen ----------
async function loadMail(mail) {
  if (!mail) return;
  state.mail = mail;
  $('search').value = '';
  $('view').innerHTML = '<div class="skeleton"></div><div class="skeleton" style="height:110px"></div><div class="skeleton" style="height:150px"></div>';
  try {
    state.ctx = await provider.lookup({ emails: mail.emails, subject: mail.subject });
    state.tab = state.ctx.focus?.doc ? tabForKind(state.ctx.focus.doc.kind) : 'projects';
    renderHome();
  } catch (e) { renderError(e); }
}

function renderHome() {
  const { mail, ctx } = state;
  if (!ctx) return;
  const parts = [];

  parts.push(`<div class="mailctx">${mail.mode === 'compose' ? 'An:' : 'Von:'} <b>${esc(mail.fromName || mail.fromEmail || 'unbekannt')}</b> <span>${esc(mail.fromEmail || '')}</span></div>`);

  if (ctx.focus) parts.push(renderFocus(ctx.focus));

  if (!ctx.contact) {
    parts.push(renderUnknown(mail));
  } else {
    parts.push(renderContact(ctx.contact, ctx.company, ctx.colleagues));
    if (ctx.company) parts.push(renderCompany(ctx.company));
    parts.push(renderTabs(ctx));
  }

  parts.push(renderActions(ctx));
  $('view').innerHTML = parts.join('');
  wireView();
}

function renderFocus(focus) {
  if (!focus.project) {
    return `<div class="card focus"><div class="card-label">Im Betreff erkannt <span class="chip">${esc(focus.matchedBy)}</span></div>
      <div class="subtitle">Zu dieser Nummer wurde in Elwis kein Vorgang gefunden.</div></div>`;
  }
  const p = focus.project;
  const docLine = focus.doc ? `<div class="facts" style="margin-top:4px"><span>${KIND_LABEL[focus.doc.kind]} <b>${esc(focus.doc.number)}</b></span>${badge(focus.doc.kind, focus.doc.status)}${focus.doc.kind === 'invoice' ? badge('payment', focus.doc.payment) : ''}</div>` : '';
  const openTasks = (p.tasks || []).filter((t) => !t.done);
  const lastNote = (p.notes || [])[0];
  return `<div class="card focus">
    <div class="card-label">Im Betreff erkannt <span class="chip">${esc(focus.matchedBy)}</span></div>
    <div class="title-row"><div class="title">${esc(p.name)}</div><span class="chip office">${esc(p.office)}</span></div>
    <div class="subtitle">${esc(p.pid)} · ${esc(p.brand || '')}</div>
    ${docLine}
    ${renderSteps(p)}
    <div class="facts">
      <span>Projektleitung <b>${esc(p.projectLead)}</b></span>
      <span>Volumen <b>${money(p.amount, p.currency)}</b></span>
      ${nextDate(p)}
    </div>
    ${openTasks.length ? `<div class="mini"><div class="mini-title">Offene Aufgaben</div>${openTasks.map((t) => `<div class="task"><span>${esc(t.name)}</span><span class="who">${esc(t.assignedTo.split(' ')[0])} · ${relDue(t.deadline)}</span></div>`).join('')}</div>` : ''}
    ${lastNote ? `<div class="mini"><div class="mini-title">Letzte Notiz</div><div class="note"><div class="meta">${dateDe(lastNote.date)} · ${esc(lastNote.author)}</div>${esc(lastNote.text)}</div></div>` : ''}
    <div class="btn-row">
      <button class="btn primary" data-open-project="${p.id}">Projekt ansehen</button>
      <button class="btn" data-elwis="${p.kind}|${p.office}|${p.id}">In Elwis öffnen ↗</button>
    </div>
  </div>`;
}

function nextDate(p) {
  const order = (p.orders || [])[0];
  const openInv = (p.invoices || []).find((i) => i.openAmount > 0);
  if (openInv) return `<span>Rechnung ${esc(openInv.number)} ${relDue(openInv.due)}</span>`;
  if (order?.etd && daysFromToday(order.etd) >= 0) return `<span>ETD <b>${dateDe(order.etd)}</b></span>`;
  if (p.status === 1) return `<span>Deadline <b>${dateDe(p.deadline)}</b> · ${relDue(p.deadline)}</span>`;
  return '';
}

function renderSteps(p) {
  const steps = [
    ['Anfrage', [p], 'inquiry'],
    ['Muster', p.samples || [], 'sample'],
    ['Angebot', p.quotes || [], 'quote'],
    ['Auftrag', p.orders || [], 'order'],
    ['Versand', p.shipments || [], 'shipment'],
    ['Rechnung', p.invoices || [], 'invoice']
  ];
  const doneCodes = { inquiry: [2], sample: [2], quote: [4], order: [4, 5, 7], shipment: [2, 4], invoice: [6] };
  return `<div class="steps">${steps.map(([label, list, kind]) => {
    if (!list.length) return `<div class="step"><div class="dot"></div>${label}</div>`;
    const last = list[list.length - 1];
    const allDone = list.every((x) => doneCodes[kind].includes(x.status));
    return `<div class="step ${allDone ? 'done' : 'active'}" title="${esc(status(kind, last.status).label)}"><div class="dot"></div>${label}<span class="st">${esc(shortLabel(status(kind, last.status).label))}</span></div>`;
  }).join('')}</div>`;
}

const shortLabel = (l) => ({ 'Transferred to Quote': 'Transferred', 'Send to customer': 'Sent', 'Pending Prepayment': 'Prepayment', 'Partially Paid': 'Part. paid' }[l] || l.split('/')[0]);

function renderContact(c, company, colleagues = []) {
  return `<div class="card">
    <div class="card-label">Kontakt ${badge('contactType', c.type)}</div>
    <div class="contact-row">
      <div class="avatar">${esc(initials(c.name))}</div>
      <div class="contact-meta">
        <div class="name">${esc(c.name)}</div>
        <div class="subtitle">${esc(c.role || '')}${company ? ' · ' + esc(company.name) : ''}</div>
      </div>
      <button class="btn" data-elwis="contact|${c.office}|${c.id}" title="Kontakt in Elwis öffnen">↗</button>
    </div>
    <div class="lines">
      <div><span class="k">E-Mail</span><span class="v">${esc(c.email)}</span></div>
      <div><span class="k">Telefon</span><span class="v">${esc(c.phone || '–')}</span></div>
      ${colleagues.length ? `<div><span class="k">Weitere</span><span class="v">${colleagues.map((k) => esc(k.name)).join(', ')}</span></div>` : ''}
    </div>
  </div>`;
}

function renderCompany(co) {
  const usage = co.creditLimit ? Math.min(100, Math.round((co.balance / co.creditLimit) * 100)) : null;
  return `<div class="card">
    <div class="card-label">${co.kind === 'vendor' ? 'Lieferant' : 'Kunde'} ${badge('company', co.status)} <span class="chip office">${esc(co.office)}</span></div>
    <div class="title-row"><div class="title">${esc(co.name)}</div>
      <button class="btn" data-elwis="${co.kind}|${co.office}|${co.id}" title="In Elwis öffnen">↗</button></div>
    <div class="subtitle">${co.kind === 'vendor' ? 'Lieferanten-Nr.' : 'Kunden-Nr.'} ${esc(co.number)} · ${esc(co.city)}, ${esc(co.country)}</div>
    <div class="lines">
      <div><span class="k">Projektleitung</span><span class="v">${esc(co.projectLead)}</span></div>
      <div><span class="k">Team</span><span class="v">${esc(co.team.join(', '))}</span></div>
      <div><span class="k">Zahlungsbed.</span><span class="v">${esc(co.paymentTerms)}</span></div>
    </div>
    <div class="metrics">
      <div class="metric"><div class="k">Saldo</div><div class="v">${money(co.balance, co.currency)}</div></div>
      <div class="metric ${co.overdue > 0 ? 'bad' : ''}"><div class="k">Überfällig</div><div class="v">${money(co.overdue, co.currency)}</div></div>
      <div class="metric"><div class="k">Kreditlimit</div><div class="v">${co.creditLimit ? usage + ' %' : '–'}</div>
        ${co.creditLimit ? `<div class="bar ${usage > 70 ? 'warn' : ''}"><span style="width:${usage}%"></span></div>` : ''}</div>
    </div>
  </div>`;
}

const TABS = [
  ['projects', 'Projekte'], ['quotes', 'Angebote'], ['orders', 'Aufträge'], ['invoices', 'Rechnungen']
];
const tabForKind = (k) => ({ quote: 'quotes', order: 'orders', invoice: 'invoices' }[k] || 'projects');

function collect(ctx, key) {
  if (key === 'projects') return ctx.projects.map((p) => ({ ...p, _p: p }));
  return ctx.projects.flatMap((p) => (p[key] || []).map((x) => ({ ...x, _p: p })));
}

function renderTabs(ctx) {
  const counts = Object.fromEntries(TABS.map(([k]) => [k, collect(ctx, k).length]));
  const items = collect(ctx, state.tab).sort((a, b) => (b.date || b.received || '').localeCompare(a.date || a.received || ''));
  return `<div class="section-title">Vorgänge</div>
    <div class="tabs" role="tablist">${TABS.map(([k, l]) => `<button class="tab" role="tab" data-tab="${k}" aria-selected="${state.tab === k}">${l}<span class="n">${counts[k]}</span></button>`).join('')}</div>
    <div class="list">${items.length ? items.map(renderRow).join('') : '<div class="empty-list">Keine Einträge.</div>'}</div>`;
}

function renderRow(x) {
  const p = x._p;
  if (x.kind === 'inquiry') {
    return `<button class="row" data-open-project="${p.id}">
      <span class="r-title">${esc(p.name)}</span>
      <span class="r-right">${badge('inquiry', p.status)}<span>${money(p.amount, p.currency)}</span></span>
      <span class="r-sub">${esc(p.pid)} · ${p.status === 1 ? 'Deadline ' + dateDe(p.deadline) : esc(stageText(p))}</span>
    </button>`;
  }
  const extra = x.kind === 'invoice'
    ? (x.openAmount > 0 ? `offen ${money(x.openAmount, x.currency)} · fällig ${dateDe(x.due)}` : `bezahlt · ${dateDe(x.date)}`)
    : x.kind === 'quote' ? `gültig bis ${dateDe(x.validTo)}`
    : x.kind === 'order' ? (x.etd ? `ETD ${dateDe(x.etd)}` : dateDe(x.date)) : dateDe(x.date);
  return `<button class="row" data-open-doc="${p.id}|${x.id}">
    <span class="r-title">${esc(x.name)}</span>
    <span class="r-right">${badge(x.kind, x.status)}<span>${money(x.amount, x.currency)}</span></span>
    <span class="r-sub">${esc(x.number || '')} · ${esc(p.pid)} · ${extra}</span>
  </button>`;
}

function stageText(p) {
  if (p.invoices?.length) return 'Rechnung ' + status('invoice', p.invoices[p.invoices.length - 1].status).label;
  if (p.shipments?.length) return 'Versand ' + status('shipment', p.shipments[p.shipments.length - 1].status).label;
  if (p.orders?.length) return 'Auftrag ' + status('order', p.orders[p.orders.length - 1].status).label;
  if (p.quotes?.length) return 'Angebot ' + status('quote', p.quotes[p.quotes.length - 1].status).label;
  return '';
}

function renderUnknown(mail) {
  return `<div class="card"><div class="empty">
    <div class="big">Kein Kontakt in Elwis gefunden</div>
    <div>Zu <b>${esc(mail.fromEmail || 'dieser Adresse')}</b> gibt es keinen Kontakt.</div>
    <div class="btn-row" style="justify-content:center">
      <button class="btn" data-action="newcontact">Als Kontakt anlegen</button>
    </div>
    <div class="stage2">Anlegen kommt in Stufe 2.</div>
    ${provider.isDemo && state.inOutlook ? `<div class="mini" style="text-align:left"><div class="mini-title">Demo: Beispiel ansehen</div>
      ${demoEmails.slice(0, 4).map((m, i) => `<button class="btn block" style="margin-top:4px" data-demo="${i}">${esc(m.label)}</button>`).join('')}</div>` : ''}
  </div></div>`;
}

function renderActions(ctx) {
  return `<div class="section-title">Aus dieser E-Mail</div>
    <div class="actions">
      <button class="btn block" data-action="note">Als Notiz speichern</button>
      <button class="btn block" data-action="task">Aufgabe anlegen</button>
    </div>
    <div class="stage2">Vorschau auf Stufe 2 – in der Demo wird nichts gespeichert.</div>`;
}

function wireView() {
  const v = $('view');
  v.querySelectorAll('[data-tab]').forEach((b) => b.addEventListener('click', () => { state.tab = b.dataset.tab; renderHome(); }));
  v.querySelectorAll('[data-open-project]').forEach((b) => b.addEventListener('click', () => openProject(b.dataset.openProject)));
  v.querySelectorAll('[data-open-doc]').forEach((b) => b.addEventListener('click', () => { const [pid, did] = b.dataset.openDoc.split('|'); openDoc(pid, did); }));
  v.querySelectorAll('[data-elwis]').forEach((b) => b.addEventListener('click', (e) => { e.stopPropagation(); const [k, o, id] = b.dataset.elwis.split('|'); openInElwis(k, o, id); }));
  v.querySelectorAll('[data-action="note"]').forEach((b) => b.addEventListener('click', openNoteForm));
  v.querySelectorAll('[data-action="task"]').forEach((b) => b.addEventListener('click', openTaskForm));
  v.querySelectorAll('[data-action="newcontact"]').forEach((b) => b.addEventListener('click', () => toast('Stufe 2: Kontakt aus Absender anlegen – in der Demo nicht aktiv.')));
  v.querySelectorAll('[data-demo]').forEach((b) => b.addEventListener('click', () => loadMail(toMail(demoEmails[b.dataset.demo]))));
}

// ---------- Suche ----------
async function runSearch(q) {
  $('view').innerHTML = '<div class="spinner"></div>';
  const hits = await provider.search(q);
  $('view').innerHTML = `<div class="section-title">Suche „${esc(q)}“</div>
    <div class="list" style="border-top:1px solid var(--line);border-radius:var(--radius)">
    ${hits.length ? hits.map((h) => `<button class="row" data-hit='${esc(JSON.stringify(h))}'>
      <span class="r-title">${esc(h.title)}</span>
      <span class="r-right"><span class="chip">${esc(KIND_LABEL[h.kind] || h.kind)}</span></span>
      <span class="r-sub">${esc(h.subtitle)}</span></button>`).join('') : '<div class="empty-list">Nichts gefunden.</div>'}
    </div>`;
  $('view').querySelectorAll('[data-hit]').forEach((b) => b.addEventListener('click', async () => {
    const h = JSON.parse(b.dataset.hit);
    if (h.docId) return openDoc(h.projectId, h.docId);
    if (h.projectId) return openProject(h.projectId);
    if (h.email) { $('search').value = ''; state.mail = { mode: 'read', fromName: h.title, fromEmail: h.email, subject: '', emails: [h.email] }; state.ctx = await provider.lookup({ emails: [h.email], subject: '' }); state.tab = 'projects'; renderHome(); }
  }));
}

// ---------- Detail-Ansichten ----------
function openSheet(title, html) {
  $('sheetTitle').textContent = title;
  $('sheetBody').innerHTML = html;
  $('sheet').hidden = false;
  $('sheetBackdrop').hidden = false;
  $('sheetBody').scrollTop = 0;
}
function closeSheet() { $('sheet').hidden = true; $('sheetBackdrop').hidden = true; }

async function openProject(id) {
  const p = await provider.getProject(id);
  if (!p) return;
  const sec = (title, list, kind, fn) => list && list.length ? `<div class="section-title">${title}</div><div class="list" style="border-top:1px solid var(--line);border-radius:var(--radius)">${list.map(fn).join('')}</div>` : '';
  const docRow = (kind) => (x) => `<button class="row" data-open-doc="${p.id}|${x.id}">
      <span class="r-title">${esc(x.name)}</span>
      <span class="r-right">${badge(kind, x.status)}${x.amount ? `<span>${money(x.amount, x.currency)}</span>` : ''}</span>
      <span class="r-sub">${esc(x.number || '')}${x.number ? ' · ' : ''}${dateDe(x.date)}${kind === 'shipment' && x.eta ? ' · ETA ' + dateDe(x.eta) : ''}</span></button>`;
  openSheet(p.pid, `
    <div class="card">
      <div class="title-row"><div class="title">${esc(p.name)}</div>${badge('inquiry', p.status)}</div>
      <div class="subtitle">${esc(p.pid)} · ${esc(p.brand || '')} · ${esc(p.category || '')}</div>
      ${renderSteps(p)}
      <div class="kv" style="margin-top:10px">
        <span class="k">Projektleitung</span><span>${esc(p.projectLead)}</span>
        <span class="k">Team</span><span>${esc(p.team.join(', '))}</span>
        <span class="k">Anfrage erhalten</span><span>${dateDe(p.received)}</span>
        <span class="k">Deadline</span><span>${dateDe(p.deadline)}</span>
        <span class="k">Volumen</span><span>${money(p.amount, p.currency)}</span>
        <span class="k">Priorität</span><span>${badge('priority', p.priority)}</span>
        <span class="k">Closing Chance</span><span>${badge('closingChance', p.closingChance)}</span>
      </div>
      <div class="btn-row"><button class="btn primary" data-elwis="inquiry|${p.office}|${p.id}">In Elwis öffnen ↗</button></div>
    </div>
    ${sec('Muster', p.samples, 'sample', docRow('sample'))}
    ${sec('Angebote', p.quotes, 'quote', docRow('quote'))}
    ${sec('Aufträge', p.orders, 'order', docRow('order'))}
    ${sec('Versand', p.shipments, 'shipment', docRow('shipment'))}
    ${sec('Rechnungen', p.invoices, 'invoice', docRow('invoice'))}
    ${p.tasks?.length ? `<div class="section-title">Aufgaben</div><div class="card">${p.tasks.map((t) => `<div class="task"><span>${esc(t.name)}</span><span class="who">${esc(t.assignedTo)} · ${relDue(t.deadline)}</span></div>`).join('')}</div>` : ''}
    ${p.notes?.length ? `<div class="section-title">Notizen</div><div class="card">${p.notes.map((n) => `<div class="note"><div class="meta">${dateDe(n.date)} · ${esc(n.author)}</div>${esc(n.text)}</div>`).join('')}</div>` : ''}
  `);
  wireSheet();
}

async function openDoc(projectId, docId) {
  const p = await provider.getProject(projectId);
  const all = [...p.samples, ...p.quotes, ...p.orders, ...p.shipments, ...p.invoices];
  const x = all.find((d) => d.id === docId);
  if (!x) return;
  const rows = [['Projekt', `${esc(p.pid)} – ${esc(p.name)}`], ['Status', badge(x.kind, x.status)]];
  if (x.number) rows.push(['Nummer', esc(x.number)]);
  rows.push(['Datum', dateDe(x.date)]);
  if (x.amount != null) rows.push(['Betrag', money(x.amount, x.currency)]);
  if (x.kind === 'quote') rows.push(['Gültig bis', `${dateDe(x.validTo)} · ${relDue(x.validTo)}`]);
  if (x.kind === 'order') { rows.push(['Offener Betrag', money(x.openAmount, x.currency)]); if (x.etd) rows.push(['ETD / ETA', `${dateDe(x.etd)} / ${dateDe(x.eta)}`]); if (x.production) rows.push(['Produktion', esc(x.production)]); }
  if (x.kind === 'shipment') { rows.push(['Versandart', esc(x.mode)]); rows.push(['ETA', dateDe(x.eta)]); rows.push(['Tracking', esc(x.tracking)]); }
  if (x.kind === 'invoice') { rows.push(['Zahlung', badge('payment', x.payment)]); rows.push(['Rechnungsart', esc(x.type)]); rows.push(['Netto', money(x.net, x.currency)]); rows.push(['Offen', money(x.openAmount, x.currency)]); rows.push(['Fällig', `${dateDe(x.due)}${x.openAmount > 0 ? ' · ' + relDue(x.due) : ''}`]); }
  if (x.kind === 'sample' && x.deadline) rows.push(['Deadline', dateDe(x.deadline)]);
  openSheet(`${KIND_LABEL[x.kind]} ${x.number || ''}`.trim(), `
    <div class="card">
      <div class="title">${esc(x.name)}</div>
      <div class="kv" style="margin-top:10px">${rows.map(([k, v]) => `<span class="k">${k}</span><span>${v}</span>`).join('')}</div>
      <div class="btn-row">
        <button class="btn primary" data-elwis="${x.kind}|${p.office}|${x.id}">In Elwis öffnen ↗</button>
        <button class="btn" data-open-project="${p.id}">Zum Projekt</button>
      </div>
    </div>`);
  wireSheet();
}

function wireSheet() {
  const s = $('sheetBody');
  s.querySelectorAll('[data-open-doc]').forEach((b) => b.addEventListener('click', () => { const [pid, did] = b.dataset.openDoc.split('|'); openDoc(pid, did); }));
  s.querySelectorAll('[data-open-project]').forEach((b) => b.addEventListener('click', () => openProject(b.dataset.openProject)));
  s.querySelectorAll('[data-elwis]').forEach((b) => b.addEventListener('click', () => { const [k, o, id] = b.dataset.elwis.split('|'); openInElwis(k, o, id); }));
}

// ---------- Stufe 2: Notiz / Aufgabe (Demo) ----------
function projectOptions() {
  const list = state.ctx?.projects || [];
  const focusId = state.ctx?.focus?.project?.id;
  if (!list.length) return '<option value="">– kein Projekt –</option>';
  return list.map((p) => `<option value="${p.id}" ${p.id === focusId ? 'selected' : ''}>${esc(p.pid)} – ${esc(p.name)}</option>`).join('');
}

function openNoteForm() {
  const m = state.mail || {};
  openSheet('Als Notiz speichern', `
    <div class="form">
      <div class="hint">Demo: Es wird nichts in Elwis gespeichert. In der Live-Version entsteht eine Notiz (Notes NBG/HKG) am gewählten Projekt.</div>
      <label>Projekt<select id="nfProject">${projectOptions()}</select></label>
      <label>Titel<input type="text" id="nfTitle" value="${esc('E-Mail: ' + (m.subject || ''))}"></label>
      <label>Kommentar<textarea id="nfText" placeholder="Optional"></textarea></label>
      <label class="check"><input type="checkbox" id="nfBody" checked> E-Mail-Text übernehmen</label>
      <label class="check"><input type="checkbox" id="nfAtt"> Anhänge mitspeichern</label>
      <button class="btn primary block" id="nfSave">Notiz speichern</button>
    </div>`);
  $('nfSave').addEventListener('click', async (e) => {
    e.target.disabled = true; e.target.textContent = 'Speichere …';
    await provider.saveNote({ projectId: $('nfProject').value, title: $('nfTitle').value, text: $('nfText').value, includeBody: $('nfBody').checked, includeAttachments: $('nfAtt').checked });
    closeSheet();
    toast(provider.isDemo ? 'Demo: Notiz würde jetzt in Elwis gespeichert. Es wurde nichts verändert.' : 'Notiz in Elwis gespeichert.', true);
  });
}

function openTaskForm() {
  const m = state.mail || {};
  const due = new Date(); due.setDate(due.getDate() + 2);
  const dueIso = `${due.getFullYear()}-${String(due.getMonth() + 1).padStart(2, '0')}-${String(due.getDate()).padStart(2, '0')}`;
  openSheet('Aufgabe anlegen', `
    <div class="form">
      <div class="hint">Demo: Es wird nichts in Elwis gespeichert. In der Live-Version entsteht eine Aufgabe (Task) am gewählten Projekt.</div>
      <label>Projekt<select id="tfProject">${projectOptions()}</select></label>
      <label>Aufgabe<input type="text" id="tfName" value="${esc('Rückmeldung an ' + (m.fromName || m.fromEmail || '') + ': ' + (m.subject || ''))}"></label>
      <label>Zuständig<select id="tfWho">${Object.values(STAFF).map((s) => `<option>${esc(s)}</option>`).join('')}</select></label>
      <label>Deadline<input type="date" id="tfDue" value="${dueIso}"></label>
      <button class="btn primary block" id="tfSave">Aufgabe anlegen</button>
    </div>`);
  $('tfSave').addEventListener('click', async (e) => {
    e.target.disabled = true; e.target.textContent = 'Lege an …';
    await provider.createTask({ projectId: $('tfProject').value, name: $('tfName').value, assignedTo: $('tfWho').value, deadline: $('tfDue').value });
    closeSheet();
    toast(provider.isDemo ? 'Demo: Aufgabe würde jetzt in Elwis angelegt. Es wurde nichts verändert.' : 'Aufgabe in Elwis angelegt.', true);
  });
}

// ---------- Elwis öffnen ----------
function openInElwis(kind, office, id) {
  const base = CONFIG.elwisBaseUrl;
  let url;
  if (provider.isDemo) {
    url = `${base}/${CONFIG.listPages[office]?.[kind] || ''}`;
    toast('Demo: Es öffnet sich die passende Liste in Elwis. In der Live-Version direkt der Datensatz.');
  } else {
    url = `${base}/${CONFIG.detailPages[office]?.[kind]}?artifactId=${encodeURIComponent(id)}`;
  }
  try {
    if (state.inOutlook && Office.context.ui.openBrowserWindow) Office.context.ui.openBrowserWindow(url);
    else window.open(url, '_blank', 'noopener');
  } catch { window.open(url, '_blank', 'noopener'); }
}

// ---------- Meldungen ----------
let toastTimer;
function toast(msg, ok = false) {
  const t = $('toast');
  t.textContent = msg; t.classList.toggle('ok', ok); t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 4200);
}
function renderEmpty(title, text) { $('view').innerHTML = `<div class="empty"><div class="big">${esc(title)}</div>${esc(text)}</div>`; }
function renderError(e) {
  if (e && e.code === 401) return renderLogin();
  renderEmpty('Das hat nicht geklappt', (e && e.message) || 'Unbekannter Fehler');
}

function renderLogin(msg = '') {
  $('view').innerHTML = `<div class="card"><div class="form">
    <div class="title">Bei Elwis anmelden</div>
    <div class="subtitle">Einmal pro Sitzung mit deinem eigenen Elwis-Zugang. Das Passwort wird nicht gespeichert.</div>
    ${msg ? `<div class="hint">${esc(msg)}</div>` : ''}
    <label>E-Mail<input type="text" id="lgEmail" autocomplete="username" value="${esc(Office.context?.mailbox?.userProfile?.emailAddress || '')}"></label>
    <label>Passwort<input type="password" id="lgPass" autocomplete="current-password" style="font:inherit;padding:7px 8px;border:1px solid var(--line);border-radius:6px"></label>
    <button class="btn primary block" id="lgGo">Anmelden</button>
  </div></div>`;
  $('lgGo').addEventListener('click', async (ev) => {
    ev.target.disabled = true;
    try { await provider.login($('lgEmail').value.trim(), $('lgPass').value); state.inOutlook ? loadFromOutlook() : loadMail(state.mail); }
    catch (err) { renderLogin(err.message); }
  });
}
