// Backend für „Elwis für Outlook“ – Stufe 1 (nur lesend).
// Start:  ELWIS_BASE_URL=https://freshid.elwis.cloud ALLOWED_ORIGIN=https://freshidbln.github.io node server.js
// Läuft mit Node 20+ ohne weitere Abhängigkeiten; lässt sich 1:1 in Azure App Service,
// einen Container oder (mit kleinem Adapter) Azure Functions übernehmen.
//
// Anmeldung: Der Benutzer meldet sich im Add-in mit seinem eigenen Elwis-Zugang an.
// Das Backend speichert NUR das Elwis-Sitzungs-Cookie im Arbeitsspeicher (kein Passwort)
// und gibt dem Add-in ein zufälliges Token zurück. Damit gelten in Elwis weiter die
// Rechte des jeweiligen Benutzers.

import http from 'node:http';
import crypto from 'node:crypto';
import { ElwisClient, httpError } from './elwis-client.js';
import { Model, MAP } from './model.js';
import { parseRefs } from '../src/refs.js';

const PORT = process.env.PORT || 8787;
const BASE = process.env.ELWIS_BASE_URL || 'https://freshid.elwis.cloud';
const ORIGIN = process.env.ALLOWED_ORIGIN || 'https://freshidbln.github.io';
const SESSION_HOURS = 8;

const sessions = new Map(); // token -> { client, model, expires, user }

function json(res, status, body) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': ORIGIN,
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Cache-Control': 'no-store'
  });
  res.end(JSON.stringify(body));
}

async function readBody(req) {
  let raw = '';
  for await (const chunk of req) { raw += chunk; if (raw.length > 100_000) throw httpError(413, 'Zu groß'); }
  return raw ? JSON.parse(raw) : {};
}

function session(req) {
  const token = (req.headers.authorization || '').replace(/^Bearer\s+/i, '');
  const s = sessions.get(token);
  if (!s || s.expires < Date.now()) { sessions.delete(token); throw httpError(401, 'Nicht angemeldet'); }
  return s;
}

// ---------- Fachlogik ----------
async function buildProject(model, office, inq) {
  const m = MAP[office];
  const p = model.normalizeDoc(office, 'inquiry', inq);
  p.pid = p.pid || p.name;
  p.samples = []; p.quotes = []; p.orders = []; p.shipments = []; p.invoices = []; p.tasks = []; p.notes = [];
  const plural = { quote: 'quotes', sample: 'samples', order: 'orders', shipment: 'shipments', invoice: 'invoices' };
  await Promise.all(Object.entries(m.links.children).map(async ([kind, linkName]) => {
    const docs = await model.linked([inq.id], [linkName], 'in');
    p[plural[kind]] = docs.map((d) => model.normalizeDoc(office, kind, d));
  }));
  return p;
}

async function lookup(s, { emails = [], subject = '' }) {
  const { model } = s;
  await model.load();
  const contacts = await model.findContactsByEmail(emails.filter(Boolean).slice(0, 5));
  const first = contacts[0] || null;

  let contact = null, company = null, projects = [];
  if (first) {
    const { office, raw } = first;
    contact = { id: raw.id, office, name: model.val(raw, 'Name'), email: model.val(raw, MAP[office].contact.email), phone: model.val(raw, 'Phone No.'), type: model.num(raw, 'Contact type') };
    const comps = await model.linked([raw.id], MAP[office].links.contactCompany, 'out');
    const co = comps[0];
    if (co) {
      const isVendor = model.typeName[co.artifactTypeId]?.startsWith('Vendor');
      company = { id: co.id, kind: isVendor ? 'vendor' : 'customer', office, name: model.val(co, 'Name'),
        number: model.val(co, MAP[office][isVendor ? 'vendor' : 'customer'].number), status: model.num(co, 'Company Status'),
        city: model.val(co, 'Sold to City'), country: '', projectLead: '', team: [], paymentTerms: model.val(co, 'Payment Terms Description')?.replace(/<[^>]+>/g, '') || '',
        currency: 'EUR', creditLimit: model.num(co, 'Credit Limit (EUR)'), balance: model.num(co, 'Balance (EUR)'), overdue: model.num(co, 'Overdue Payments') };
      contact.companyId = co.id;
      if (!isVendor) {
        const inqs = await model.linked([co.id], MAP[office].links.inquiryCustomer, 'in');
        projects = await Promise.all(inqs.slice(0, 15).map((i) => buildProject(model, office, i)));
      }
    }
  }

  // PID in allen gelebten Schreibweisen (siehe src/refs.js); ohne Büro-Kürzel erst NBG, dann HKG
  let focus = null;
  const refs = parseRefs(subject);
  for (const r of refs.pids) {
    for (const office of r.office ? [r.office] : ['NBG', 'HKG']) {
      const m = MAP[office].inquiry;
      const pid = `PID-${r.num}-${office}`;
      const hit = (await model.c.listArtifacts(model.byType(m.type, { [`attributes.${model.attrId[m.pid]}.value`]: { $eq: pid } }), { limit: 1 }))[0];
      if (hit) { focus = { matchedBy: pid, project: projects.find((p) => p.id === hit.id) || await buildProject(model, office, hit) }; break; }
    }
    if (focus) break;
  }
  // TODO Live: SC-Nummern (Sales Order NBG „SC Number“ / HKG „SC Number HKG“) und Belegnummern nachschlagen
  if (!focus && refs.pids.length) focus = { matchedBy: 'PID ' + refs.pids[0].num, project: null };
  return { contact, company, colleagues: [], projects, focus };
}

// ---------- Routen ----------
const server = http.createServer(async (req, res) => {
  try {
    if (req.method === 'OPTIONS') return json(res, 204, {});
    const url = new URL(req.url, 'http://x');

    if (req.method === 'POST' && url.pathname === '/api/session') {
      const { email, password } = await readBody(req);
      const client = new ElwisClient(BASE);
      const info = await client.login(email, password);
      const token = crypto.randomBytes(32).toString('base64url');
      sessions.set(token, { client, model: new Model(client), user: info.sessionUser?.email, expires: Date.now() + SESSION_HOURS * 3600e3 });
      return json(res, 200, { token, user: info.sessionUser?.email });
    }

    if (req.method === 'POST' && url.pathname === '/api/lookup') return json(res, 200, await lookup(session(req), await readBody(req)));

    if (req.method === 'GET' && url.pathname === '/api/health') return json(res, 200, { ok: true, readOnly: true });

    // Stufe 2 – bewusst noch nicht freigeschaltet
    if (['/api/notes', '/api/tasks'].includes(url.pathname)) return json(res, 501, { error: 'Schreiben ist in Stufe 1 deaktiviert' });

    return json(res, 404, { error: 'Nicht gefunden' });
  } catch (e) {
    console.error(e.message);
    return json(res, e.status || 500, { error: e.message });
  }
});

server.listen(PORT, () => console.log(`Elwis-Add-in-Backend auf Port ${PORT}, Ziel ${BASE}, nur lesend`));
