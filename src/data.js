// Datenschicht. Die Oberfläche spricht nur mit diesem Interface:
//   lookup({ emails, subject })  -> Kontext zur E-Mail
//   search(query)                -> Trefferliste
//   getProject(id)               -> Projekt mit allen Vorgängen
//   saveNote(payload), createTask(payload)   (Stufe 2)
// MockProvider liefert Beispieldaten, ApiProvider spricht mit dem Backend.

import { CONFIG } from './config.js';
import { companies, contacts, projects } from './mock-data.js';

const DOC_NUMBER = /\b(20\d{8})\b/g;

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

    // Projektnummer oder Belegnummer im Betreff
    let focus = null;
    const pidMatch = subject.match(CONFIG.pidPattern);
    if (pidMatch) {
      const pid = `PID-${pidMatch[1]}-${pidMatch[2].toUpperCase()}`;
      const p = projects.find((x) => x.pid === pid);
      focus = p ? { matchedBy: pid, project: p } : { matchedBy: pid, project: null };
    } else {
      for (const m of subject.matchAll(DOC_NUMBER)) {
        const hit = findDocByNumber(m[1]);
        if (hit) { focus = { matchedBy: m[1], project: hit.project, doc: hit.doc }; break; }
      }
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
    if (!res.ok) throw new Error(`Fehler ${res.status}`);
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
  saveNote(payload) { return this.call('/api/notes', { method: 'POST', body: JSON.stringify(payload) }); }
  createTask(payload) { return this.call('/api/tasks', { method: 'POST', body: JSON.stringify(payload) }); }
}

function delay(ms) { return new Promise((r) => setTimeout(r, ms)); }

export const provider = CONFIG.mode === 'live' && CONFIG.backendUrl
  ? new ApiProvider(CONFIG.backendUrl)
  : new MockProvider();
