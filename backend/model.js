// Übersetzt Elwis-Metadaten (Typen, Felder, Verknüpfungen) in die Form,
// die das Add-in erwartet. Alle Zuordnungen laufen über NAMEN, nicht über
// fest eingetragene IDs – die IDs werden beim Start aus Elwis gelesen.

import { oid, oids } from './elwis-client.js';

// Welche Elwis-Typen und Felder je Büro verwendet werden
export const MAP = {
  NBG: {
    contact: { type: 'Contact NBG', email: 'Email [NBG]', phone: 'Phone No.', ctype: 'Contact type' },
    customer: { type: 'Customer NBG', number: 'Customer Number' },
    vendor: { type: 'Vendor NBG', number: 'Vendor Number' },
    inquiry: { type: 'Inquiry NBG', folder: 'Project Folder', pid: 'Project Inquiry ID', status: 'Inquiry Status', amount: 'Inquiry Amount', deadline: 'Inquiry Deadline', received: 'Inquiry Received' },
    sample: { type: 'Sample NBG', status: 'Sample Status PM', date: 'Sample Order Date', deadline: 'Sample Deadline', amount: 'Quotation Amount' },
    quote: { type: 'Quote NBG', number: 'Quote Number', status: 'Quote Status', date: 'Quote Date', validTo: 'Quote Valid to Date', amount: 'Total' },
    order: { type: 'Sales Order NBG', number: 'Order Number', sc: 'SC Number', status: 'Order Status', date: 'Sales Order Date', amount: 'Total', openAmount: 'Open Amount', etd: 'ETD', eta: 'ETA' },
    shipment: { type: 'Shipment NBG', number: 'Shipment Number', status: 'Shipment Status PM', date: 'Shipment Date', eta: 'ETA' },
    invoice: { type: 'Invoice NBG', number: 'Invoice Number Nürnberg', status: 'Invoice Status', payment: 'Payment Status', date: 'Invoice Date', due: 'Invoice Due Date', amount: 'Total with VAT', net: 'Total', openAmount: 'Open Amount' },
    links: {
      contactCompany: ['Contact - Customer'],
      inquiryCustomer: ['Inquiry-Customer'],
      children: { quote: 'Quote - Inquiry', sample: 'Sample - Inquiry', order: 'Sales order - Inquiry', shipment: 'Shipment - Inquiry', invoice: 'Invoice - Inquiry' }
    }
  },
  HKG: {
    contact: { type: 'Contact HKG', email: 'Email', phone: 'Phone No.', ctype: 'Contact type' },
    customer: { type: 'Customer HKG', number: 'Customer Number HKG' },
    vendor: { type: 'Vendor HKG', number: 'Original Id' },
    inquiry: { type: 'Inquiry HKG', folder: 'Project Folder', pid: 'HKG PID', status: 'Inquiry Status', amount: 'Inquiry Amount', deadline: 'Inquiry Deadline', received: 'Inquiry Received' },
    sample: { type: 'Sample HKG', status: 'Sample Status PM', date: 'Sample Order Date', deadline: 'Sample Deadline', amount: 'Quotation Amount' },
    quote: { type: 'Quote HKG', status: 'Quote Status', date: 'Quote Date', validTo: 'Quote Valid to Date', amount: 'Total' },
    order: { type: 'Sales Order HKG', number: 'Order Number', sc: 'SC Number HKG', status: 'Order Status', date: 'Sales Order Date', amount: 'Total', openAmount: 'Open Amount', etd: 'ETD', eta: 'ETA' },
    shipment: { type: 'Shipment HKG', status: 'Shipment Status PM', date: 'Shipment Date', eta: 'ETA' },
    invoice: { type: 'Invoice HKG', status: 'Invoice Status', payment: 'Payment Status', date: 'Invoice Date', due: 'Invoice Due Date', amount: 'Total with VAT', net: 'Total', openAmount: 'Open Amount' },
    links: {
      contactCompany: ['HKG Contact - Customer'],
      inquiryCustomer: ['HKG Inquiry-Customer'],
      children: { quote: 'HKG Quote - Inquiry', sample: 'HKG Sample - Inquiry', order: 'HKG Sales order - Inquiry', shipment: 'HKG Shipment - Inquiry', invoice: 'HKG Invoice - Inquiry' }
    }
  }
};

// Vorhandene Elwis-Felder, die das Board „Auftragsabwicklung“ schon heute anzeigen kann
// (Schlüssel wie in src/orders.js). Die übrigen Status-Felder legt Softacus neu an.
export const ORDER_PM_FIELDS = {
  prod: { field: 'Production Stage', kind: 'label' }, fracht: { field: 'Mode of Shipment', kind: 'label' },
  fertig: { field: 'SC Ready Date', kind: 'date' }, eta: { field: 'ETA', kind: 'date' },
  ltLief: { field: 'Live PO Ready Date', kind: 'date' }, ltKunde: { field: 'Delivery Date', kind: 'date' }
};
// Vertriebsfelder (Schlüssel wie in src/sales.js); gesucht erst an der Anfrage, dann am Angebot
export const SALES_SOURCE = {
  chance: 'Closing Chance', prio: 'Priority', entscheidung: 'Closing Date', lostReason: 'Lost Reason'
};
const CHANCE = { 1: 'Close to an order', 2: 'Neutral', 3: 'Low chance' };

export class Model {
  constructor(client) { this.c = client; this.ready = null; }

  async load() {
    if (!this.ready) {
      this.ready = (async () => {
        const [types, attrs, links, dtypes] = await Promise.all([
          this.c.get('/api/tenant/artifact-type'), this.c.get('/api/tenant/attribute'), this.c.get('/api/tenant/link-type'),
          this.c.get('/api/tenant/data-type').catch(() => ({ data: [] }))
        ]);
        const live = (x) => x.data.filter((d) => !d.deleted);
        this.typeId = Object.fromEntries(live(types).map((t) => [t.name, t.id]));
        this.typeName = Object.fromEntries(live(types).map((t) => [t.id, t.name]));
        this.attrId = Object.fromEntries(live(attrs).map((a) => [a.name, a.id]));
        this.linkId = Object.fromEntries(live(links).map((l) => [l.name, l.id]));
        this.sys = { name: this.attrId['Name'], projectLead: this.attrId['Project Lead'], team: this.attrId['Team'] };
        // Auswahllisten: Wert -> Anzeigename (für Felder wie „Production Stage“)
        const dt = Object.fromEntries(live(dtypes).map((d) => [d.id, d]));
        this.enumOf = {};
        live(attrs).forEach((a) => {
          const d = dt[a.dataTypeId || a.dataType?.id || a.dataType];
          const vals = d && (d.values || d.options || d.listValues);
          if (Array.isArray(vals)) this.enumOf[a.name] = Object.fromEntries(vals.map((v) => [String(v.value ?? v.id ?? v.key), v.label ?? v.name ?? v.value]));
        });
      })();
    }
    return this.ready;
  }

  val(a, field) {
    const id = this.attrId[field];
    const v = id && a.attributes?.[id]?.value;
    return v === '' || v === undefined ? null : v;
  }
  num(a, field) { const v = this.val(a, field); const n = v == null ? null : parseFloat(v); return Number.isFinite(n) ? n : null; }
  date(a, field) { const v = this.val(a, field); return v ? String(v).slice(0, 10) : null; }
  label(a, field) { const v = this.val(a, field); if (v == null) return null; return this.enumOf?.[field]?.[String(v)] ?? String(v); }

  // Artefakte eines Typs, deren Feld einen der Werte hat (Zahl oder Text – je nach Elwis-Datentyp)
  async findBy(typeName, field, values, limit = 5) {
    await this.load();
    if (!this.typeId[typeName] || !this.attrId[field]) return [];
    const vs = values.flatMap((v) => (/^\d+$/.test(String(v)) ? [String(v), Number(v)] : [String(v)]));
    return this.c.listArtifacts(this.byType(typeName, { [`attributes.${this.attrId[field]}.value`]: { $in: vs } }), { limit });
  }

  // Wie linked(), aber für viele Artefakte auf einmal: Map(id -> [verknüpfte Artefakte])
  async linkedMap(artifactIds, linkNames, direction) {
    const out = new Map();
    const ltIds = linkNames.map((n) => this.linkId[n]).filter(Boolean);
    if (!ltIds.length || !artifactIds.length) return out;
    const side = direction === 'out' ? 'sourceArtifactId' : 'destinationArtifactId';
    const other = direction === 'out' ? 'destinationArtifactId' : 'sourceArtifactId';
    const links = [];
    for (let i = 0; i < artifactIds.length; i += 100) {
      links.push(...await this.c.listLinks({ deleted: { $eq: null }, linkTypeId: oids(ltIds), [side]: oids(artifactIds.slice(i, i + 100)) }, { limit: 2000 }));
    }
    const ids = [...new Set(links.map((l) => l[other]))];
    const arts = [];
    for (let i = 0; i < ids.length; i += 100) {
      const chunk = ids.slice(i, i + 100);
      arts.push(...await this.c.listArtifacts({ deleted: { $eq: null }, _id: oids(chunk) }, { limit: chunk.length }));
    }
    const byId = new Map(arts.map((a) => [a.id, a]));
    links.forEach((l) => { const a = byId.get(l[other]); if (!a) return; if (!out.has(l[side])) out.set(l[side], []); out.get(l[side]).push(a); });
    return out;
  }

  // Vertriebsstatus aus vorhandenen Elwis-Daten ableiten (Logik wie src/sales.js)
  toDeal(office, inq, { quotes = [], orders = [], company = null } = {}) {
    const m = MAP[office].inquiry;
    const q = [...quotes].sort((a, b) => String(b.created?.on || '').localeCompare(String(a.created?.on || '')))[0];
    const qs = q ? this.num(q, MAP[office].quote.status) : null;
    const is = this.num(inq, m.status);
    let vstatus = 'Offen';
    if (orders.length) vstatus = 'Übergeben';
    else if (qs === 5) vstatus = 'Verloren';
    else if (qs === 4) vstatus = 'Gewonnen';
    else if (qs === 3) vstatus = 'Raus an Kunden';
    else if (qs === 1 || qs === 2) vstatus = 'Angebot schreiben';
    else if (is === 3) vstatus = 'Verloren';
    const pick = (field, kind) => {
      for (const a of [inq, q].filter(Boolean)) {
        const v = kind === 'date' ? this.date(a, field) : kind === 'num' ? this.num(a, field) : this.label(a, field);
        if (v != null) return v;
      }
      return null;
    };
    const chance = pick(SALES_SOURCE.chance, 'num');
    const prio = pick(SALES_SOURCE.prio, 'num');
    return {
      id: 'deal-' + inq.id, kind: 'deal', projectId: inq.id, office, pid: this.val(inq, m.pid), name: this.val(inq, 'Name'),
      company: company ? this.val(company, 'Name') : '', contact: '', seller: this.val(inq, 'Project Lead') || '',
      an: q ? this.val(q, MAP[office].quote.number || 'Quote Number') : '', amount: (q && this.num(q, 'Total')) ?? this.num(inq, m.amount),
      currency: 'EUR', received: this.date(inq, m.received), vstatus,
      chance: chance ? CHANCE[chance] || String(chance) : '', prio: prio ? String(prio) : '',
      entscheidung: pick(SALES_SOURCE.entscheidung, 'date') || '', lostReason: pick(SALES_SOURCE.lostReason, 'label') || ''
    };
  }

  // Auftrag für das Board „Auftragsabwicklung“ (vorhandene Felder; neue Status-Felder bleiben leer)
  toBoardOrder(office, so, { inquiry = null, company = null } = {}) {
    const o = this.normalizeDoc(office, 'order', so);
    for (const [k, f] of Object.entries(ORDER_PM_FIELDS)) o[k] = (f.kind === 'date' ? this.date(so, f.field) : this.label(so, f.field)) || '';
    o.sc = o.sc != null ? String(o.sc) : '';
    o.projectId = inquiry?.id || null;
    o.pid = inquiry ? this.val(inquiry, MAP[office].inquiry.pid) : '';
    o.pm = this.val(so, 'Project Lead') || (inquiry && this.val(inquiry, 'Project Lead')) || '';
    o.seller = inquiry ? this.val(inquiry, 'Project Lead') || '' : '';
    o.company = company ? this.val(company, 'Name') : '';
    o.contact = '';
    return o;
  }

  byType(typeName, extra = {}) {
    return { deleted: { $eq: null }, artifactTypeId: { $in: [oid(this.typeId[typeName])] }, ...extra };
  }

  async findContactsByEmail(emails) {
    await this.load();
    const out = [];
    for (const office of ['NBG', 'HKG']) {
      const m = MAP[office].contact;
      for (const e of emails) {
        const rx = '^' + e.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$';
        const hits = await this.c.listArtifacts(this.byType(m.type, { [`attributes.${this.attrId[m.email]}.value`]: { $regex: rx, $options: 'i' } }), { limit: 5 });
        hits.forEach((h) => out.push({ office, raw: h }));
      }
    }
    return out;
  }

  // Ziel- bzw. Quell-Artefakte über Verknüpfungstypen
  async linked(artifactIds, linkNames, direction) {
    const ltIds = linkNames.map((n) => this.linkId[n]).filter(Boolean);
    if (!ltIds.length || !artifactIds.length) return [];
    const side = direction === 'out' ? 'sourceArtifactId' : 'destinationArtifactId';
    const other = direction === 'out' ? 'destinationArtifactId' : 'sourceArtifactId';
    const links = await this.c.listLinks({ deleted: { $eq: null }, linkTypeId: oids(ltIds), [side]: oids(artifactIds) });
    const ids = [...new Set(links.map((l) => l[other]))];
    if (!ids.length) return [];
    return this.c.listArtifacts({ deleted: { $eq: null }, _id: oids(ids) }, { limit: ids.length });
  }

  normalizeDoc(office, kind, a) {
    const m = MAP[office][kind];
    const doc = { id: a.id, kind, office, name: this.val(a, 'Name') };
    for (const [k, field] of Object.entries(m)) {
      if (k === 'type') continue;
      if (['amount', 'net', 'openAmount'].includes(k)) doc[k] = this.num(a, field);
      else if (['date', 'due', 'validTo', 'deadline', 'received', 'etd', 'eta'].includes(k)) doc[k] = this.date(a, field);
      else if (['status', 'payment'].includes(k)) doc[k] = this.num(a, field);
      else if (k === 'folder') { const v = this.val(a, field); const m = v && (typeof v === 'string' ? v : JSON.stringify(v)).match(/https?:\/\/[^"\s]+/); doc.folder = m ? m[0] : null; }
      else doc[k] = this.val(a, field);
    }
    doc.currency = 'EUR'; // TODO: Currency-Auswahlliste auswerten
    return doc;
  }
}
