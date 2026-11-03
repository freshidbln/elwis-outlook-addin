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
    'Access-Control-Allow-Methods': 'GET, POST, PATCH, OPTIONS',
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
  const raw = {};
  await Promise.all(Object.entries(m.links.children).map(async ([kind, linkName]) => {
    raw[kind] = await model.linked([inq.id], [linkName], 'in');
    p[plural[kind]] = raw[kind].map((d) => model.normalizeDoc(office, kind, d));
  }));
  // Auftragsstatus-Karte und Vertriebsstatus-Karte im Add-in
  p.orders = raw.order.map((so) => model.toBoardOrder(office, so, { inquiry: inq }));
  p.deal = model.toDeal(office, inq, { quotes: raw.quote, orders: raw.order });
  return p;
}

// Zu einem Beleg (Angebot, Auftrag, Rechnung) das Projekt (Anfrage) finden
async function projectOfDoc(model, office, kind, doc) {
  const linkName = MAP[office].links.children[kind];
  const inq = (await model.linked([doc.id], [linkName], 'out'))[0];
  return inq ? buildProject(model, office, inq) : null;
}

// Belegnummern 20xxxxxxxx: Angebot, Auftrag oder Rechnung (NBG)
const DOC_FIELDS = [['quote', 'number'], ['order', 'number'], ['invoice', 'number']];

// Monday-Board-Ersatz: Listen für die Übersicht (nur lesend)
const ACTIVE_ORDER = (code) => ![6, 7].includes(code); // ohne Cancelled / Completed
async function listOrders(s) {
  const { model } = s; await model.load();
  const out = [];
  for (const office of ['NBG', 'HKG']) {
    const m = MAP[office];
    const sos = (await model.c.listArtifacts(model.byType(m.order.type), { limit: 300 })).filter((so) => ACTIVE_ORDER(model.num(so, m.order.status)));
    const inqMap = await model.linkedMap(sos.map((x) => x.id), [m.links.children.order], 'out');
    const inqIds = [...new Set([...inqMap.values()].flat().map((i) => i.id))];
    const custMap = await model.linkedMap(inqIds, m.links.inquiryCustomer, 'out');
    sos.forEach((so) => {
      const inquiry = (inqMap.get(so.id) || [])[0] || null;
      const company = inquiry ? (custMap.get(inquiry.id) || [])[0] : null;
      out.push(model.toBoardOrder(office, so, { inquiry, company }));
    });
  }
  return out;
}

async function listDeals(s) {
  const { model } = s; await model.load();
  const out = [];
  for (const office of ['NBG', 'HKG']) {
    const m = MAP[office];
    const inqs = await model.c.listArtifacts(model.byType(m.inquiry.type), { limit: 300 });
    const ids = inqs.map((i) => i.id);
    const [quotes, orders, cust] = await Promise.all([
      model.linkedMap(ids, [m.links.children.quote], 'in'),
      model.linkedMap(ids, [m.links.children.order], 'in'),
      model.linkedMap(ids, m.links.inquiryCustomer, 'out')
    ]);
    inqs.forEach((inq) => out.push(model.toDeal(office, inq, { quotes: quotes.get(inq.id) || [], orders: orders.get(inq.id) || [], company: (cust.get(inq.id) || [])[0] })));
  }
  return out;
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
  // SC-Nummern: 37xxxx/61xxxx (NBG „SC Number“), SC-xxxxxxxx (HKG „SC Number HKG“)
  if (!focus) for (const sc of refs.scs) {
    const hkg = /^SC-/.test(sc);
    const office = hkg ? 'HKG' : 'NBG';
    const so = (await model.findBy(MAP[office].order.type, MAP[office].order.sc, [hkg ? sc : sc.replace(/\D/g, ''), ...(hkg ? [sc.slice(3)] : [])], 1))[0];
    if (!so) continue;
    const project = await projectOfDoc(model, office, 'order', so);
    focus = project ? { matchedBy: 'SC ' + sc, project, doc: project.orders.find((o) => o.id === so.id) }
      : { matchedBy: 'SC ' + sc, order: model.toBoardOrder(office, so) };
    break;
  }
  // Belegnummern (Angebot / Auftrag / Rechnung)
  if (!focus) for (const num of refs.docs) {
    for (const [kind, key] of DOC_FIELDS) {
      const field = MAP.NBG[kind][key];
      const doc = field && (await model.findBy(MAP.NBG[kind].type, field, [num], 1))[0];
      if (!doc) continue;
      const project = await projectOfDoc(model, 'NBG', kind, doc);
      if (project) { focus = { matchedBy: num, project, doc: model.normalizeDoc('NBG', kind, doc) }; break; }
    }
    if (focus) break;
  }
  if (!focus && refs.pids.length) focus = { matchedBy: 'PID ' + refs.pids[0].num, project: null };
  if (!focus && refs.scs.length) focus = { matchedBy: 'SC ' + refs.scs[0], project: null };
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
    if (req.method === 'GET' && url.pathname === '/api/orders') return json(res, 200, await listOrders(session(req)));
    if (req.method === 'GET' && url.pathname === '/api/deals') return json(res, 200, await listDeals(session(req)));

    // Status ändern (Board / Add-in) – erst nach Freigabe durch fresh id und Softacus
    if (req.method === 'PATCH' && /^\/api\/(orders|deals)\//.test(url.pathname)) {
      return json(res, 501, { error: 'Status ändern ist noch nicht freigeschaltet (Stufe 1: nur lesen). Bitte vorerst in Elwis ändern.' });
    }

    // Stufe 2 – bewusst noch nicht freigeschaltet
    if (['/api/notes', '/api/tasks'].includes(url.pathname)) return json(res, 501, { error: 'Schreiben ist in Stufe 1 deaktiviert' });

    return json(res, 404, { error: 'Nicht gefunden' });
  } catch (e) {
    console.error(e.message);
    return json(res, e.status || 500, { error: e.message });
  }
});

server.listen(PORT, () => console.log(`Elwis-Add-in-Backend auf Port ${PORT}, Ziel ${BASE}, nur lesend`));
