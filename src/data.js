// Datenschicht. Die Oberfläche spricht nur mit diesem Interface:
//   lookup({ emails, subject })  -> Kontext zur E-Mail
//   search(query)                -> Trefferliste
//   getProject(id)               -> Projekt mit allen Vorgängen
//   saveNote(payload), createTask(payload)   (Stufe 2)
//   listOrders(), getOrderSync(id), updateOrder(id, key, value)   (Auftragsstatus / Board)
//   listDeals(), getDealSync(projectId), updateDeal(id, key, value), ordersForProject(id)   (Vertrieb)
// MockProvider liefert Beispieldaten, ApiProvider spricht mit dem Backend.

import { CONFIG } from './config.js';
import { companies, contacts, projects, FOLDER_TEMPLATE, ORDER_PM, extraOrders, SALES_PM, extraDeals } from './mock-data.js';
import { applyChange, applyDateRules, todayIso } from './orders.js';
import { parseRefs } from './refs.js';
import { applySalesChange, applySalesDateRules } from './sales.js';

// ---------- Demo-Speicher für Auftragsstatus ----------
// Live liegen die Werte an der Elwis Sales Order. In der Demo werden Änderungen im
// Browser gespeichert, damit Add-in und Board-Fenster denselben Stand zeigen.
const LS_KEY = 'elwis-demo-orders-v1';
const LS_CREATED = 'elwis-demo-orders-created-v1';
const LS_DEALS = 'elwis-demo-deals-v1';
const lsGet = (k, def) => { try { return JSON.parse(localStorage.getItem(k) || JSON.stringify(def)); } catch { return def; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* ohne Speicher */ } };
function loadPatches() { return lsGet(LS_KEY, {}); }
function savePatches(p) { lsSet(LS_KEY, p); }
let memPatches = loadPatches();

function baseOrders() {
  const out = [];
  projects.forEach((p) => (p.orders || []).forEach((o) => {
    const company = companies.find((c) => c.id === p.companyId);
    const contact = contacts.find((c) => c.id === p.contactId);
    out.push({ ...o, ...(ORDER_PM[o.id] || {}), projectId: p.id, pid: p.pid, office: p.office, company: company?.name || '',
      contact: contact?.name || '', pm: p.projectLead, seller: (p.team || [])[1] || p.projectLead, amount: o.amount, currency: o.currency });
  }));
  extraOrders.forEach((o) => out.push({ ...o }));
  // In der Demo per Übergabe aus dem Vertrieb entstandene Aufträge
  lsGet(LS_CREATED, []).forEach((o) => out.push({ ...o }));
  return out;
}
function mergedOrder(base) {
  memPatches = loadPatches();
  const withPatch = { ...base, ...(memPatches[base.id] || {}) };
  return applyDateRules(withPatch, todayIso()).order;
}
export const demoOrders = {
  all: () => baseOrders().map(mergedOrder),
  get: (id) => { const b = baseOrders().find((o) => o.id === id); return b ? mergedOrder(b) : null; },
  forProject: (pid) => baseOrders().filter((o) => o.projectId === pid).map(mergedOrder),
  update(id, key, value) {
    const cur = this.get(id);
    const res = applyChange(cur, key, value);
    memPatches = loadPatches();
    memPatches[id] = { ...(memPatches[id] || {}), ...res.patch };
    savePatches(memPatches);
    return { order: this.get(id), applied: res.applied };
  },
  create(deal) {
    const list = lsGet(LS_CREATED, []);
    const id = 'n' + Date.now().toString(36);
    const o = { id, kind: 'order', projectId: deal.projectId || null, pid: deal.pid, office: deal.office || 'NBG', name: deal.name,
      company: deal.company, contact: deal.contact, pm: deal.pm || deal.seller, seller: deal.seller, amount: deal.amount, currency: deal.currency || 'EUR',
      sc: '', number: '', date: todayIso(), ab: 'in Arbeit', usd: 'Nicht notwendig', fromDeal: deal.id, handoverNote: deal.comment || '' };
    list.push(o); lsSet(LS_CREATED, list);
    return o;
  },
  reset() { memPatches = {}; savePatches({}); lsSet(LS_CREATED, []); }
};

// ---------- Demo-Speicher für Vertriebsstatus ----------
function baseDeals() {
  const out = projects.map((p) => {
    const company = companies.find((c) => c.id === p.companyId);
    const contact = contacts.find((c) => c.id === p.contactId);
    return { id: 'deal-' + p.id, kind: 'deal', projectId: p.id, pid: p.pid, office: p.office, name: p.name, company: company?.name || '',
      contact: contact?.name || '', seller: (p.team || [])[0] || p.projectLead, pm: p.projectLead, amount: p.amount, currency: p.currency,
      received: p.received, ...(SALES_PM[p.id] || {}) };
  });
  extraDeals.forEach((d) => out.push({ ...d }));
  return out;
}
function mergedDeal(base) {
  const patches = lsGet(LS_DEALS, {});
  return applySalesDateRules({ ...base, ...(patches[base.id] || {}) }, todayIso());
}
export const demoDeals = {
  all: () => baseDeals().map(mergedDeal),
  get: (id) => { const b = baseDeals().find((d) => d.id === id); return b ? mergedDeal(b) : null; },
  forProject: (pid) => { const b = baseDeals().find((d) => d.projectId === pid); return b ? mergedDeal(b) : null; },
  update(id, key, value) {
    const cur = this.get(id);
    const res = applySalesChange(cur, key, value, todayIso());
    const patches = lsGet(LS_DEALS, {});
    patches[id] = { ...(patches[id] || {}), ...res.patch };
    lsSet(LS_DEALS, patches);
    let order = null;
    if (res.createOrder && !(cur.projectId && demoOrders.forProject(cur.projectId).length)) order = demoOrders.create(res.deal);
    return { deal: this.get(id), applied: res.applied, createdOrder: order };
  },
  reset() { lsSet(LS_DEALS, {}); }
};

function findDocByNumber(num) {
  for (const p of projects) {
    for (const list of [p.quotes, p.orders, p.invoices]) {
      const doc = list.find((x) => x.number === num);
      if (doc) return { project: p, doc };
    }
  }
  return null;
}

class MockProvider {
  get isDemo() { return true; }

  async lookup({ emails = [], subject = '' }) {
    await delay(250);
    const lower = emails.map((e) => (e || '').toLowerCase());
    const contact = contacts.find((c) => lower.includes(c.email.toLowerCase())) || null;
    const company = contact && contact.companyId ? companies.find((c) => c.id === contact.companyId) : null;

    // Projekt-, SC- oder Belegnummer im Betreff (Schreibweisen siehe refs.js)
    let focus = null;
    const refs = parseRefs(subject);
    for (const r of refs.pids) {
      const p = projects.find((x) => x.pid.split('-')[1] === r.num && (!r.office || x.office === r.office));
      if (p) { focus = { matchedBy: p.pid, project: p }; break; }
    }
    if (!focus) for (const sc of refs.scs) {
      const o = demoOrders.all().find((x) => x.sc === sc.replace(/^SC-/, ''));
      if (!o) continue;
      const p = o.projectId ? projects.find((x) => x.id === o.projectId) : null;
      focus = p ? { matchedBy: 'SC ' + sc, project: p, doc: p.orders.find((x) => x.id === o.id) } : { matchedBy: 'SC ' + sc, order: o };
      break;
    }
    if (!focus) for (const num of refs.docs) {
      const hit = findDocByNumber(num);
      if (hit) { focus = { matchedBy: num, project: hit.project, doc: hit.doc }; break; }
    }
    if (!focus && refs.pids.length) {
      const r = refs.pids[0];
      focus = { matchedBy: `PID ${r.num}${r.office ? '-' + r.office : ''}`, project: null };
    }

    const companyProjects = company ? projects.filter((p) => p.companyId === company.id) : [];
    const colleagues = company ? contacts.filter((c) => c.companyId === company.id && c.id !== contact.id) : [];
    return { contact, company, colleagues, projects: companyProjects, focus };
  }

  async search(query) {
    await delay(150);
    const q = query.trim().toLowerCase();
    if (q.length < 2) return [];
    const out = [];
    companies.forEach((c) => {
      if ((c.name + ' ' + c.number).toLowerCase().includes(q))
        out.push({ kind: c.kind, id: c.id, title: c.name, subtitle: `${c.kind === 'vendor' ? 'Lieferant' : 'Kunde'} ${c.number} · ${c.office}`, email: contacts.find((k) => k.companyId === c.id && k.primary)?.email });
    });
    contacts.forEach((c) => {
      if ((c.name + ' ' + c.email).toLowerCase().includes(q))
        out.push({ kind: 'contact', id: c.id, title: c.name, subtitle: c.email, email: c.email });
    });
    projects.forEach((p) => {
      if ((p.pid + ' ' + p.name).toLowerCase().includes(q))
        out.push({ kind: 'inquiry', id: p.id, title: p.name, subtitle: p.pid, projectId: p.id });
      [...p.quotes, ...p.orders, ...p.invoices].forEach((doc) => {
        if ((doc.number || '').includes(q) || (doc.name || '').toLowerCase().includes(q))
          out.push({ kind: doc.kind, id: doc.id, title: doc.name, subtitle: `${doc.number} · ${p.pid}`, projectId: p.id, docId: doc.id });
      });
    });
    return out.slice(0, 20);
  }

  async getProject(id) {
    await delay(100);
    return projects.find((p) => p.id === id) || null;
  }

  async getCompanyContext(companyId) {
    const company = companies.find((c) => c.id === companyId);
    const contact = contacts.find((c) => c.companyId === companyId && c.primary) || null;
    return {
      contact, company,
      colleagues: contacts.filter((c) => c.companyId === companyId && c !== contact),
      projects: projects.filter((p) => p.companyId === companyId), focus: null
    };
  }

  async listOrders() { await delay(120); return demoOrders.all(); }
  getOrderSync(id) { return demoOrders.get(id); }
  async updateOrder(id, key, value) { await delay(150); return demoOrders.update(id, key, value); }
  resetOrders() { demoOrders.reset(); demoDeals.reset(); }
  ordersForProject(pid) { return demoOrders.forProject(pid); }
  async listDeals() { await delay(120); return demoDeals.all(); }
  getDealSync(projectId) { return demoDeals.forProject(projectId); }
  async updateDeal(id, key, value) { await delay(150); return demoDeals.update(id, key, value); }

  async getFolders() { await delay(80); return FOLDER_TEMPLATE; }
  async saveFiles(payload) { await delay(700); return { ok: true, demo: true, payload }; }
  async saveNote(payload) { await delay(600); return { ok: true, demo: true, payload }; }
  async createTask(payload) { await delay(600); return { ok: true, demo: true, payload }; }
}

class ApiProvider {
  constructor(baseUrl) { this.baseUrl = baseUrl.replace(/\/$/, ''); }
  get isDemo() { return false; }

  async call(path, opts = {}) {
    const token = sessionStorage.getItem('elwis-addin-token');
    const res = await fetch(this.baseUrl + path, {
      ...opts,
      headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: 'Bearer ' + token } : {}), ...(opts.headers || {}) }
    });
    if (res.status === 401) throw Object.assign(new Error('Nicht angemeldet'), { code: 401 });
    if (!res.ok) { const j = await res.json().catch(() => ({})); throw Object.assign(new Error(j.error || `Fehler ${res.status}`), { code: res.status }); }
    return res.json();
  }

  async login(email, password) {
    const res = await fetch(this.baseUrl + '/api/session', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email, password }) });
    if (!res.ok) throw new Error(res.status === 403 ? 'Konto gesperrt' : 'Anmeldung fehlgeschlagen');
    const j = await res.json();
    sessionStorage.setItem('elwis-addin-token', j.token);
    return j;
  }

  async lookup({ emails, subject }) {
    const ctx = await this.call('/api/lookup', { method: 'POST', body: JSON.stringify({ emails, subject }) });
    this.cache = [...(ctx.projects || []), ...(ctx.focus?.project ? [ctx.focus.project] : [])];
    return ctx;
  }
  async search(q) { return []; } // Suche folgt im Backend
  async getProject(id) { return (this.cache || []).find((p) => p.id === id) || null; }
  getCompanyContext(id) { return this.call('/api/company/' + encodeURIComponent(id)); }
  getFolders(id) { return this.call('/api/project/' + encodeURIComponent(id) + '/folders'); }
  listOrders() { return this.call('/api/orders'); }
  getOrderSync(id) { for (const p of this.cache || []) { const o = (p.orders || []).find((x) => x.id === id); if (o) return o; } return null; }
  // Live: Backend schreibt das Feld in die Elwis Sales Order; die Regeln laufen als Elwis-Abläufe
  updateOrder(id, key, value) { return this.call('/api/orders/' + encodeURIComponent(id), { method: 'PATCH', body: JSON.stringify({ key, value }) }); }
  resetOrders() {}
  ordersForProject(pid) { const p = (this.cache || []).find((x) => x.id === pid); return p?.orders || []; }
  listDeals() { return this.call('/api/deals'); }
  getDealSync(projectId) { const p = (this.cache || []).find((x) => x.id === projectId); return p?.deal || null; }
  // Live: Vertriebsstatus wird auf Inquiry/Quote in Elwis abgebildet (siehe sales.js)
  updateDeal(id, key, value) { return this.call('/api/deals/' + encodeURIComponent(id), { method: 'PATCH', body: JSON.stringify({ key, value }) }); }
  saveFiles(payload) { return this.call('/api/files', { method: 'POST', body: JSON.stringify(payload) }); }
  saveNote(payload) { return this.call('/api/notes', { method: 'POST', body: JSON.stringify(payload) }); }
  createTask(payload) { return this.call('/api/tasks', { method: 'POST', body: JSON.stringify(payload) }); }
}

function delay(ms) { return new Promise((r) => setTimeout(r, ms)); }

export const provider = CONFIG.mode === 'live' && CONFIG.backendUrl
  ? new ApiProvider(CONFIG.backendUrl)
  : new MockProvider();
