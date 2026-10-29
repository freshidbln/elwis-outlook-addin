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
    inquiry: { type: 'Inquiry NBG', pid: 'Project Inquiry ID', status: 'Inquiry Status', amount: 'Inquiry Amount', deadline: 'Inquiry Deadline', received: 'Inquiry Received' },
    sample: { type: 'Sample NBG', status: 'Sample Status PM', date: 'Sample Order Date', deadline: 'Sample Deadline', amount: 'Quotation Amount' },
    quote: { type: 'Quote NBG', number: 'Quote Number', status: 'Quote Status', date: 'Quote Date', validTo: 'Quote Valid to Date', amount: 'Total' },
    order: { type: 'Sales Order NBG', number: 'Order Number', status: 'Order Status', date: 'Sales Order Date', amount: 'Total', openAmount: 'Open Amount', etd: 'ETD', eta: 'ETA' },
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
    inquiry: { type: 'Inquiry HKG', pid: 'HKG PID', status: 'Inquiry Status', amount: 'Inquiry Amount', deadline: 'Inquiry Deadline', received: 'Inquiry Received' },
    sample: { type: 'Sample HKG', status: 'Sample Status PM', date: 'Sample Order Date', deadline: 'Sample Deadline', amount: 'Quotation Amount' },
    quote: { type: 'Quote HKG', status: 'Quote Status', date: 'Quote Date', validTo: 'Quote Valid to Date', amount: 'Total' },
    order: { type: 'Sales Order HKG', number: 'Order Number', status: 'Order Status', date: 'Sales Order Date', amount: 'Total', openAmount: 'Open Amount', etd: 'ETD', eta: 'ETA' },
    shipment: { type: 'Shipment HKG', status: 'Shipment Status PM', date: 'Shipment Date', eta: 'ETA' },
    invoice: { type: 'Invoice HKG', status: 'Invoice Status', payment: 'Payment Status', date: 'Invoice Date', due: 'Invoice Due Date', amount: 'Total with VAT', net: 'Total', openAmount: 'Open Amount' },
    links: {
      contactCompany: ['HKG Contact - Customer'],
      inquiryCustomer: ['HKG Inquiry-Customer'],
      children: { quote: 'HKG Quote - Inquiry', sample: 'HKG Sample - Inquiry', order: 'HKG Sales order - Inquiry', shipment: 'HKG Shipment - Inquiry', invoice: 'HKG Invoice - Inquiry' }
    }
  }
};

export class Model {
  constructor(client) { this.c = client; this.ready = null; }

  async load() {
    if (!this.ready) {
      this.ready = (async () => {
        const [types, attrs, links] = await Promise.all([
          this.c.get('/api/tenant/artifact-type'), this.c.get('/api/tenant/attribute'), this.c.get('/api/tenant/link-type')
        ]);
        const live = (x) => x.data.filter((d) => !d.deleted);
        this.typeId = Object.fromEntries(live(types).map((t) => [t.name, t.id]));
        this.typeName = Object.fromEntries(live(types).map((t) => [t.id, t.name]));
        this.attrId = Object.fromEntries(live(attrs).map((a) => [a.name, a.id]));
        this.linkId = Object.fromEntries(live(links).map((l) => [l.name, l.id]));
        this.sys = { name: this.attrId['Name'], projectLead: this.attrId['Project Lead'], team: this.attrId['Team'] };
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
      else doc[k] = this.val(a, field);
    }
    doc.currency = 'EUR'; // TODO: Currency-Auswahlliste auswerten
    return doc;
  }
}
