// Erfundene Beispieldaten für die Demo. Aufbau entspricht den Elwis-Typen
// (Contact NBG, Customer NBG, Vendor HKG, Inquiry, Sample, Quote, Sales Order,
// Shipment, Invoice, Task, Notes). Keine echten Kunden-, Kontakt- oder Belegdaten.

const today = new Date();
const d = (offset) => {
  const x = new Date(today.getFullYear(), today.getMonth(), today.getDate() + offset);
  return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`;
};

const de = (iso) => iso.split('-').reverse().join('.');

export const STAFF = {
  anna: 'Anna Becker', tom: 'Tom Richter', jonas: 'Jonas Weber', mei: 'Mei Lin Chan'
};

export const companies = [
  { id: 'c1', kind: 'customer', office: 'NBG', name: 'Nordlicht Getränke GmbH', number: '11402', status: 2,
    city: 'Hamburg', country: 'Deutschland', projectLead: STAFF.anna, team: [STAFF.anna, STAFF.tom],
    paymentTerms: '30 Tage netto', currency: 'EUR', creditLimit: 50000, balance: 18940.20, overdue: 0, taxStatus: 'Taxable' },
  { id: 'c2', kind: 'customer', office: 'NBG', name: 'Alpenkräuter Spirituosen AG', number: '11388', status: 2,
    city: 'Innsbruck', country: 'Österreich', projectLead: STAFF.jonas, team: [STAFF.jonas, STAFF.tom],
    paymentTerms: '14 Tage 2 % Skonto, 30 Tage netto', currency: 'EUR', creditLimit: 75000, balance: 53210.00, overdue: 4280.00, taxStatus: 'Exempt EU' },
  { id: 'c3', kind: 'customer', office: 'NBG', name: 'Harbour Lane Retail Ltd', number: '11455', status: 3,
    city: 'London', country: 'Vereinigtes Königreich', projectLead: STAFF.anna, team: [STAFF.anna],
    paymentTerms: '50 % Anzahlung, Rest vor Versand', currency: 'GBP', creditLimit: 20000, balance: 0, overdue: 0, taxStatus: 'Exempt not EU' },
  { id: 'v1', kind: 'vendor', office: 'HKG', name: 'Shenzhen Brightpack Co., Ltd.', number: 'V-20417', status: 2,
    city: 'Shenzhen', country: 'China', projectLead: STAFF.mei, team: [STAFF.mei],
    paymentTerms: '30 % Anzahlung, 70 % vor Verschiffung', currency: 'USD', creditLimit: null, balance: 12400, overdue: 0, taxStatus: '–' }
];

export const contacts = [
  { id: 'k1', office: 'NBG', name: 'Lena Hartmann', email: 'lena.hartmann@nordlicht-getraenke.example', phone: '+49 40 555 1234', type: 2, companyId: 'c1', primary: true, role: 'Trade Marketing Managerin' },
  { id: 'k2', office: 'NBG', name: 'Felix Brandt', email: 'f.brandt@nordlicht-getraenke.example', phone: '+49 40 555 1240', type: 2, companyId: 'c1', primary: false, role: 'Einkauf' },
  { id: 'k3', office: 'NBG', name: 'Markus Vogel', email: 'markus.vogel@alpenkraeuter.example', phone: '+43 512 555 880', type: 2, companyId: 'c2', primary: true, role: 'Leitung Marketing' },
  { id: 'k4', office: 'NBG', name: 'Olivia Grant', email: 'olivia.grant@harbourlane.example', phone: '+44 20 5550 1188', type: 2, companyId: 'c3', primary: true, role: 'Visual Merchandising' },
  { id: 'k5', office: 'HKG', name: 'Kevin Liu', email: 'kevin.liu@brightpack.example', phone: '+86 755 5550 3321', type: 1, companyId: 'v1', primary: true, role: 'Sales Manager' },
  { id: 'k6', office: 'NBG', name: 'Sarah Klein', email: 'sarah@morgenrot-kaffee.example', phone: '+49 911 555 670', type: 3, companyId: null, primary: false, role: 'Inhaberin (Lead)' }
];

// Projekte = Anfragen (Inquiry) mit allen nachgelagerten Vorgängen
export const projects = [
  {
    id: 'p1', kind: 'inquiry', office: 'NBG', pid: 'PID-3201-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects/idNBGCurrent/Forms/AllItems.aspx', name: 'Sommer-Display Kühlschrankaufsteller',
    companyId: 'c1', contactId: 'k1', status: 2, received: d(-48), deadline: d(-35), amount: 24500, currency: 'EUR',
    brand: 'Nordlicht Pils', category: 'Metal', priority: 1, closingChance: 1, projectLead: STAFF.anna, team: [STAFF.anna, STAFF.tom],
    samples: [{ id: 's1', kind: 'sample', name: 'Sommer-Display – Muster 1', status: 2, date: d(-42), deadline: d(-30), amount: 380 }],
    quotes: [{ id: 'q1', kind: 'quote', number: '2026090112', name: 'Sommer-Display Kühlschrankaufsteller', status: 4, date: d(-28), validTo: d(2), amount: 24500, currency: 'EUR' }],
    orders: [{ id: 'o1', kind: 'order', number: '2026090045', name: 'Sommer-Display – 350 Stk.', status: 3, date: d(-21), amount: 24500, currency: 'EUR', openAmount: 14700, etd: d(5), eta: d(38), production: 'In Produktion' }],
    shipments: [{ id: 'sh1', kind: 'shipment', name: 'Seefracht Yantian → Hamburg', status: 1, date: d(5), eta: d(38), mode: 'Sea', tracking: 'wird nach Verladung ergänzt' }],
    invoices: [{ id: 'i1', kind: 'invoice', number: '2026100211', name: 'Anzahlung 40 % – Sommer-Display', status: 6, payment: 2, date: d(-20), due: d(-6), amount: 11662, net: 9800, currency: 'EUR', openAmount: 0, type: 'Partial' }],
    tasks: [{ id: 't1', kind: 'task', name: 'Druckfreigabe Etikett beim Kunden einholen', assignedTo: STAFF.anna, deadline: d(2), done: false }],
    notes: [
      { date: d(-3), author: STAFF.anna, text: 'Kunde wünscht Lieferung bis KW 50, Verladung bestätigt für ' + de(d(5)) + '.' },
      { date: d(-12), author: STAFF.tom, text: 'Anzahlung eingegangen, Auftrag freigegeben.' }
    ]
  },
  {
    id: 'p2', kind: 'inquiry', office: 'NBG', pid: 'PID-3215-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects/idNBGCurrent/Forms/AllItems.aspx', name: 'Weihnachts-Geschenkset 2026',
    companyId: 'c1', contactId: 'k2', status: 1, received: d(-6), deadline: d(3), amount: 18000, currency: 'EUR',
    brand: 'Nordlicht Winterbock', category: 'Paper', priority: 2, closingChance: 2, projectLead: STAFF.anna, team: [STAFF.anna, STAFF.tom],
    samples: [{ id: 's2', kind: 'sample', name: 'Geschenkbox mit Inlay – Muster', status: 1, date: d(-4), deadline: d(9), amount: 220 }],
    quotes: [], orders: [], shipments: [], invoices: [],
    tasks: [{ id: 't2', kind: 'task', name: 'Kalkulation an Kunden senden', assignedTo: STAFF.tom, deadline: d(1), done: false }],
    notes: [{ date: d(-5), author: STAFF.anna, text: 'Briefing erhalten: 2.000 Sets, Auslieferung Anfang Dezember.' }]
  },
  {
    id: 'p3', kind: 'inquiry', office: 'NBG', pid: 'PID-3188-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects/idNBGCurrent/Forms/AllItems.aspx', name: 'Thekenaufsteller Acryl',
    companyId: 'c2', contactId: 'k3', status: 2, received: d(-25), deadline: d(-14), amount: 8940, currency: 'EUR',
    brand: 'Alpenkräuter Original', category: 'Plastic', priority: 2, closingChance: 1, projectLead: STAFF.jonas, team: [STAFF.jonas],
    samples: [],
    quotes: [{ id: 'q3', kind: 'quote', number: '2026100087', name: 'Thekenaufsteller Acryl – 500 Stk.', status: 3, date: d(-10), validTo: d(12), amount: 8940, currency: 'EUR' }],
    orders: [], shipments: [], invoices: [],
    tasks: [{ id: 't3', kind: 'task', name: 'Nachfassen Angebot 2026100087', assignedTo: STAFF.jonas, deadline: d(0), done: false }],
    notes: []
  },
  {
    id: 'p4', kind: 'inquiry', office: 'NBG', pid: 'PID-3142-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects/idNBGCurrent/Forms/AllItems.aspx', name: 'Neon-Leuchtschild Bar-Edition',
    companyId: 'c2', contactId: 'k3', status: 2, received: d(-110), deadline: d(-95), amount: 21400, currency: 'EUR',
    brand: 'Alpenkräuter Original', category: 'Electronic', priority: 1, closingChance: 1, projectLead: STAFF.jonas, team: [STAFF.jonas, STAFF.tom],
    samples: [{ id: 's4', kind: 'sample', name: 'Leuchtschild – Muster', status: 2, date: d(-100), deadline: d(-88), amount: 450 }],
    quotes: [{ id: 'q4', kind: 'quote', number: '2026070054', name: 'Neon-Leuchtschild Bar-Edition', status: 4, date: d(-85), validTo: d(-55), amount: 21400, currency: 'EUR' }],
    orders: [{ id: 'o4', kind: 'order', number: '2026070031', name: 'Neon-Leuchtschild – 120 Stk.', status: 7, date: d(-80), amount: 21400, currency: 'EUR', openAmount: 4280 }],
    shipments: [{ id: 'sh4', kind: 'shipment', name: 'LKW Nürnberg → Innsbruck', status: 2, date: d(-40), eta: d(-38), mode: 'Truck', tracking: 'DE-TRK-88213' }],
    invoices: [
      { id: 'i4a', kind: 'invoice', number: '2026080144', name: 'Teilrechnung 80 %', status: 6, payment: 2, date: d(-60), due: d(-30), amount: 17120, net: 17120, currency: 'EUR', openAmount: 0, type: 'Partial' },
      { id: 'i4b', kind: 'invoice', number: '2026090318', name: 'Restrechnung 20 %', status: 4, payment: 1, date: d(-39), due: d(-9), amount: 4280, net: 4280, currency: 'EUR', openAmount: 4280, type: 'Remaining' }
    ],
    tasks: [],
    notes: [{ date: d(-2), author: STAFF.tom, text: 'Zahlungserinnerung für 2026090318 vorbereitet.' }]
  },
  {
    id: 'p5', kind: 'inquiry', office: 'NBG', pid: 'PID-3230-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects/idNBGCurrent/Forms/AllItems.aspx', name: 'Pop-up Store Kit UK',
    companyId: 'c3', contactId: 'k4', status: 1, received: d(-2), deadline: d(12), amount: 12000, currency: 'GBP',
    brand: 'Harbour Lane', category: 'Wood', priority: 3, closingChance: 2, projectLead: STAFF.anna, team: [STAFF.anna],
    samples: [],
    quotes: [{ id: 'q5', kind: 'quote', number: '2026100131', name: 'Pop-up Store Kit – Entwurf', status: 1, date: d(-1), validTo: d(29), amount: 12000, currency: 'GBP' }],
    orders: [], shipments: [], invoices: [], tasks: [], notes: []
  },
  {
    id: 'p6', kind: 'inquiry', office: 'HKG', pid: 'PID-1874-HKG', folder: 'https://freshidproducts.sharepoint.com/sites/freshidProductsIntranetHK/projects', name: 'OEM Glass Bottle Gift Box',
    companyId: 'v1', contactId: 'k5', status: 2, received: d(-60), deadline: d(-50), amount: 31800, currency: 'USD',
    brand: 'Nordlicht Winterbock', category: 'Glass', priority: 1, closingChance: 1, projectLead: STAFF.mei, team: [STAFF.mei],
    samples: [{ id: 's6', kind: 'sample', name: 'Gift Box – Golden Sample', status: 2, date: d(-45), deadline: d(-35), amount: 0 }],
    quotes: [{ id: 'q6', kind: 'quote', number: '2026080210', name: 'OEM Glass Bottle Gift Box', status: 4, date: d(-40), validTo: d(-10), amount: 31800, currency: 'USD' }],
    orders: [{ id: 'o6', kind: 'order', number: '2026080021', name: 'Gift Box – 5.000 Stk.', status: 3, date: d(-33), amount: 31800, currency: 'USD', openAmount: 22260, etd: d(9), eta: d(40), production: 'Massenproduktion 60 %' }],
    shipments: [], invoices: [],
    tasks: [{ id: 't6', kind: 'task', name: 'Inspektionstermin (FRI) bestätigen', assignedTo: STAFF.mei, deadline: d(4), done: false }],
    notes: [{ date: d(-1), author: STAFF.mei, text: 'Produktion 60 % fertig, Inspektion geplant.' }]
  }
];

// Beispiel-E-Mails für die Vorschau außerhalb von Outlook
export const demoEmails = [
  { label: 'Lena Hartmann – Lieferdatum Sommer-Display', fromName: 'Lena Hartmann', fromEmail: 'lena.hartmann@nordlicht-getraenke.example',
    subject: 'AW: fresh id / PID-3201-NBG / Sommer-Display – Lieferdatum?', received: d(0) },
  { label: 'Felix Brandt – Schreibweise wie im Alltag (PID ohne Büro + SC-Nr.)', fromName: 'Felix Brandt', fromEmail: 'f.brandt@nordlicht-getraenke.example',
    subject: 'AW: Sommer-Display // 376901 // PID 3201', received: d(0) },
  { label: 'Markus Vogel – Angebot nachfassen (Vertrieb, überfällig)', fromName: 'Markus Vogel', fromEmail: 'markus.vogel@alpenkraeuter.example',
    subject: 'AW: Ihr Angebot 2026100087 – Thekenaufsteller Acryl', received: d(0) },
  { label: 'Felix Brandt – neue Anfrage läuft (Vertrieb)', fromName: 'Felix Brandt', fromEmail: 'f.brandt@nordlicht-getraenke.example',
    subject: 'Weihnachts-Geschenkset PID 3215 – gibt es schon Preise?', received: d(0) },
  { label: 'Markus Vogel – Zahlung Rechnung 2026090318', fromName: 'Markus Vogel', fromEmail: 'markus.vogel@alpenkraeuter.example',
    subject: 'Zahlung Rechnung 2026090318', received: d(0) },
  { label: 'Kevin Liu – Produktionsstand (HKG)', fromName: 'Kevin Liu', fromEmail: 'kevin.liu@brightpack.example',
    subject: 'RE: PID-1874-HKG production update', received: d(-1) },
  { label: 'Sarah Klein – neue Anfrage (Lead)', fromName: 'Sarah Klein', fromEmail: 'sarah@morgenrot-kaffee.example',
    subject: 'Anfrage Verkaufsständer für unsere Kaffeebar', received: d(-1) },
  { label: 'Unbekannter Absender', fromName: 'Info', fromEmail: 'info@unbekannt.example',
    subject: 'Ihre Anfrage', received: d(-2) }
];

// Standard-Ordnerstruktur je Projekt – Beispiel von fresh id (Projektordner 618778, „fid Group“)
// depth = Ebene im Baum, path = Pfad relativ zum Projektordner
export const FOLDER_TEMPLATE = [
  { key: 'abbildungen', name: '#1 Abbildungen Angebote', path: '#1 Abbildungen Angebote', depth: 0 },
  { key: 'angLief', name: '#2 Angebote von Lieferanten', path: '#2 Angebote von Lieferanten', depth: 0 },
  { key: 'angKunde', name: '#3 Angebote an Kunden', path: '#3 Angebote an Kunden', depth: 0 },
  { key: 'auftrag', name: '#4 Auftrag', path: '#4 Auftrag', depth: 0 },
  { key: 'kunde', name: 'Kunde', path: '#4 Auftrag/Kunde', depth: 1 },
  { key: 'kundeAB', name: 'Auftragsbestätigung', path: '#4 Auftrag/Kunde/Auftragsbestätigung', depth: 2 },
  { key: 'druckdaten', name: 'Druckdaten', path: '#4 Auftrag/Kunde/Druckdaten', depth: 2 },
  { key: 'freigabemuster', name: 'Freigabemuster', path: '#4 Auftrag/Kunde/Freigabemuster', depth: 2 },
  { key: 'lieferant', name: 'Lieferant', path: '#4 Auftrag/Lieferant', depth: 1 },
  { key: 'liefAB', name: 'Auftragsbestätigung', path: '#4 Auftrag/Lieferant/Auftragsbestätigung', depth: 2 },
  { key: 'bestellung', name: 'Bestellung', path: '#4 Auftrag/Lieferant/Bestellung', depth: 2 },
  { key: 'korri', name: 'Korrekturabzug', path: '#4 Auftrag/Lieferant/Korrekturabzug', depth: 2 },
  { key: 'lieferdok', name: 'Lieferdokumente', path: '#4 Auftrag/Lieferant/Lieferdokumente', depth: 2 },
  { key: 'rechnung', name: '#5 Rechnung', path: '#5 Rechnung', depth: 0 },
  { key: 'muster', name: 'Muster', path: 'Muster', depth: 0 },
  { key: 'vorlagen', name: 'Vorlagen', path: 'Vorlagen', depth: 0 }
];

// ---------- Projektmanagement-Status der Aufträge (Demo) ----------
// SC-Nummern sind erfunden. Live stehen diese Werte an der Elwis Sales Order.
export const ORDER_PM = {
  o1: { sc: '376901', ab: 'Raus an Kunden', usd: 'Nicht notwendig', korri: 'Freigegeben', korriDate: d(-25), marks: 'Freigegeben',
        muster: 'Freigegeben', musterDate: d(-22), prod: 'in Produktion', fertig: d(4), fracht: 'Seefracht', liefer: '', eta: d(38), ltLief: d(4), ltKunde: d(44), lieferung: '' },
  o4: { sc: '376788', ab: 'Raus an Kunden', usd: 'Nicht notwendig', korri: 'Freigegeben', korriDate: d(-70), marks: 'Nicht notwendig',
        muster: 'ohne Freigabemuster', prod: 'Fertiggestellt', fertig: d(-45), fracht: 'LKW', liefer: 'Angekommen', eta: d(-38), ltLief: d(-45), ltKunde: d(-38), lieferung: 'Fertig zum Abrechnen' },
  o6: { sc: '376865', ab: 'Raus an Kunden', usd: 'Erledigt', korri: 'Freigegeben', korriDate: d(-30), marks: 'Angefordert',
        muster: 'Freigegeben', musterDate: d(-20), prod: 'in Produktion', fertig: d(8), fracht: 'Seefracht', liefer: '', eta: d(40), ltLief: d(8), ltKunde: d(45), lieferung: '' }
};

// Weitere Aufträge, damit die Board-Übersicht realistisch gefüllt ist (alle erfunden)
const X = (o) => ({ kind: 'order', office: 'NBG', currency: 'EUR', usd: 'Nicht notwendig', lieferung: '', ...o });
export const extraOrders = [
  X({ id: 'x1', sc: '376931', name: 'Hoodies Teamevent', company: 'Velotec Fahrräder GmbH', contact: 'Jana Roth', pm: 'Anna Becker', seller: 'Jonas Weber', amount: 7400,
      ab: 'in Arbeit', korri: 'Angefordert' }),
  X({ id: 'x2', sc: '376934', name: 'Notizblöcke A5 mit Logo', company: 'Stadtwerke Elbtal', contact: 'Peter Lang', pm: 'Lea Wolf', seller: 'Lea Wolf', amount: 1850,
      ab: 'Offen' }),
  X({ id: 'x3', sc: '376936', name: 'Powerbank 10.000 mAh', company: 'Medivet Tiergesundheit GmbH', contact: 'Sabine Kurz', pm: 'Mara Hoffmann', seller: 'Mara Hoffmann', amount: 2790,
      ab: 'in Arbeit', usd: 'Offen', korri: 'Angefordert' }),
  X({ id: 'x4', sc: '376912', name: 'Kühltasche Sommeraktion', company: 'Bergquell Brauerei GmbH', contact: 'Tim Becker', pm: 'Anna Becker', seller: 'Tom Richter', amount: 11200,
      ab: 'Raus an Kunden', korri: 'warte auf Freigabe', korriDate: d(-2), marks: 'Nicht notwendig', muster: 'ohne Freigabemuster', prod: 'Freigabeprozess', ltKunde: d(40) }),
  X({ id: 'x5', sc: '376915', name: 'Lanyards Messe 2027', company: 'Nordwind Versicherung AG', contact: 'Clara Neumann', pm: 'Lea Wolf', seller: 'Jonas Weber', amount: 3100,
      ab: 'Raus an Kunden', korri: 'Freigegeben', korriDate: d(-6), muster: 'warte auf Freigabe', musterDate: d(3), prod: 'Freigabeprozess', fracht: 'DHL etc./Kurier' }),
  X({ id: 'x6', sc: '376918', name: 'Schlüsselbänder recycelt', company: 'Lumina Kosmetik GmbH', contact: 'Eva Martin', pm: 'Mara Hoffmann', seller: 'Mara Hoffmann', amount: 4650,
      ab: 'Raus an Kunden', korri: 'warte auf Freigabe', korriDate: d(2), marks: 'Angefordert', muster: 'mit Freigabemuster', prod: 'Freigabeprozess' }),
  X({ id: 'x7', sc: '376874', name: 'Regenschirme mit Logo', company: 'Nordwind Versicherung AG', contact: 'Clara Neumann', pm: 'Lea Wolf', seller: 'Lea Wolf', amount: 18900,
      ab: 'Raus an Kunden', korri: 'Freigegeben', korriDate: d(-20), marks: 'Freigegeben', muster: 'ohne Freigabemuster', prod: 'in Produktion', fertig: d(10), fracht: 'Seefracht', ltLief: d(10), ltKunde: d(55) }),
  X({ id: 'x8', sc: '376869', name: 'Glasflaschen mit Gravur', company: 'Bergquell Brauerei GmbH', contact: 'Tim Becker', pm: 'Anna Becker', seller: 'Tom Richter', amount: 26400,
      ab: 'Raus an Kunden', usd: 'Erledigt', korri: 'Freigegeben', korriDate: d(-28), marks: 'Freigegeben', muster: 'Freigegeben', musterDate: d(-18), prod: 'in Produktion', fertig: d(-3), fracht: 'Zugfracht', ltLief: d(-3), ltKunde: d(30) }),
  X({ id: 'x9', sc: '376841', name: 'Bierbank-Garnituren Branding', company: 'Bergquell Brauerei GmbH', contact: 'Tim Becker', pm: 'Anna Becker', seller: 'Tom Richter', amount: 38500,
      ab: 'Raus an Kunden', korri: 'Freigegeben', marks: 'Freigegeben', muster: 'ohne Freigabemuster', prod: 'Fertiggestellt', fertig: d(-14), fracht: 'Seefracht', liefer: 'Unterwegs', eta: d(12), ltLief: d(-14), ltKunde: d(18) }),
  X({ id: 'x10', sc: '376852', name: 'Sporttaschen Team', company: 'Velotec Fahrräder GmbH', contact: 'Jana Roth', pm: 'Mara Hoffmann', seller: 'Jonas Weber', amount: 9800,
      ab: 'Raus an Kunden', korri: 'Freigegeben', muster: 'ohne Freigabemuster', prod: 'Fertiggestellt', fertig: d(-9), fracht: 'LKW', liefer: 'Unterwegs', eta: d(-1), ltKunde: d(3) }),
  X({ id: 'x11', sc: '376829', name: 'Kugelschreiber Metall', company: 'Stadtwerke Elbtal', contact: 'Peter Lang', pm: 'Lea Wolf', seller: 'Lea Wolf', amount: 2300,
      ab: 'Raus an Kunden', korri: 'Freigegeben', muster: 'ohne Freigabemuster', prod: 'Fertiggestellt', fertig: d(-5), fracht: 'DHL etc./Kurier', liefer: 'Nachlauf/Auslieferung', ltLief: d(1), ltKunde: d(2) }),
  X({ id: 'x12', sc: '376815', name: 'Tassen Weihnachten', company: 'Medivet Tiergesundheit GmbH', contact: 'Sabine Kurz', pm: 'Mara Hoffmann', seller: 'Mara Hoffmann', amount: 5200,
      ab: 'Raus an Kunden', korri: 'Freigegeben', muster: 'ohne Freigabemuster', prod: 'Fertiggestellt', fertig: d(-20), fracht: 'DHL etc./Kurier', liefer: 'Angekommen', ltKunde: d(-8), lieferung: 'Ausgeliefert' }),
  X({ id: 'x13', sc: '375990', name: 'Abruf Namensschilder Q4', company: 'Nordwind Versicherung AG', contact: 'Clara Neumann', pm: 'Lea Wolf', seller: 'Lea Wolf', amount: 560, abruf: true,
      ab: 'Raus an Kunden', korri: 'Freigegeben', muster: 'ohne Freigabemuster' }),
  X({ id: 'x14', sc: '375512', name: 'Abruf Flyer Messe', company: 'Lumina Kosmetik GmbH', contact: 'Eva Martin', pm: 'Mara Hoffmann', seller: 'Mara Hoffmann', amount: 1250, abruf: true,
      ab: 'Raus an Kunden', muster: 'ohne Freigabemuster' })
];

// ---------- Vertriebsstatus (Demo) ----------
// Für die Beispielprojekte p1–p6 (Anfragen in Elwis) plus weitere erfundene Deals
export const SALES_PM = {
  p1: { vstatus: 'Übergeben', chance: 'Close to an order', prio: '1', entscheidung: d(-30), an: '2026090112' },
  p2: { vstatus: 'Angefragt (Lieferant)', chance: 'Neutral', prio: '2', entscheidung: d(3), an: '' },
  p3: { vstatus: 'Raus an Kunden', chance: 'Close to an order', prio: '2', entscheidung: d(-2), an: '2026100087' },
  p4: { vstatus: 'Übergeben', chance: 'Close to an order', prio: '1', entscheidung: d(-90), an: '2026070054' },
  p5: { vstatus: 'Angebot schreiben', chance: 'Neutral', prio: '3', entscheidung: d(12), an: '2026100131' },
  p6: { vstatus: 'Übergeben', chance: 'Close to an order', prio: '1', entscheidung: d(-45), an: '2026080210' }
};

const D = (o) => ({ kind: 'deal', office: 'NBG', currency: 'EUR', chance: 'Neutral', prio: '2', ...o });
export const extraDeals = [
  D({ id: 'd1', pid: 'PID-3402-NBG', an: '', name: 'Thermobecher Messe', company: 'Velotec Fahrräder GmbH', contact: 'Jana Roth', seller: 'Jonas Weber', amount: 6200, received: d(-1), vstatus: 'Offen', entscheidung: d(6) }),
  D({ id: 'd2', pid: 'PID-3399-NBG', an: '', name: 'Adventskalender Kunden', company: 'Medivet Tiergesundheit GmbH', contact: 'Sabine Kurz', seller: 'Mara Hoffmann', amount: 14500, received: d(-4), vstatus: 'Offen', entscheidung: d(-1) }),
  D({ id: 'd3', pid: 'PID-3391-NBG', an: '', name: 'Bambus-Lunchbox', company: 'Stadtwerke Elbtal', contact: 'Peter Lang', seller: 'Lea Wolf', amount: 8800, received: d(-9), vstatus: 'Angefragt (Lieferant)', entscheidung: d(4) }),
  D({ id: 'd4', pid: 'PID-3387-NBG', an: '2026100152', name: 'Notizbücher recycelt', company: 'Nordwind Versicherung AG', contact: 'Clara Neumann', seller: 'Lea Wolf', amount: 4300, received: d(-12), vstatus: 'Raus an Kunden', entscheidung: d(9), chance: 'Close to an order' }),
  D({ id: 'd5', pid: 'PID-3380-NBG', an: '2026100140', name: 'Kühlakkus mit Logo', company: 'Bergquell Brauerei GmbH', contact: 'Tim Becker', seller: 'Tom Richter', amount: 12900, received: d(-20), vstatus: 'Raus an Kunden', entscheidung: d(-6), chance: 'Neutral' }),
  D({ id: 'd6', pid: 'PID-3371-NBG', an: '2026100118', name: 'Sonnenschirme Biergarten', company: 'Bergquell Brauerei GmbH', contact: 'Tim Becker', seller: 'Anna Becker', amount: 41000, received: d(-27), vstatus: 'In Verhandlung', entscheidung: d(5), chance: 'Close to an order', prio: '1' }),
  D({ id: 'd7', pid: 'PID-3365-NBG', an: '2026100103', name: 'Kosmetiktaschen', company: 'Lumina Kosmetik GmbH', contact: 'Eva Martin', seller: 'Mara Hoffmann', amount: 7600, received: d(-31), vstatus: 'In Verhandlung', entscheidung: d(-3), chance: 'Neutral' }),
  D({ id: 'd8', pid: 'PID-3342-NBG', an: '2026090170', name: 'Fahrradklingeln Aktion', company: 'Velotec Fahrräder GmbH', contact: 'Jana Roth', seller: 'Jonas Weber', amount: 3500, received: d(-50), vstatus: 'On Hold', onHoldSince: d(-20), entscheidung: d(25), chance: 'Low chance' }),
  D({ id: 'd9', pid: 'PID-3330-NBG', an: '2026090155', name: 'Messestand-Give-aways 2027', company: 'Nordwind Versicherung AG', contact: 'Clara Neumann', seller: 'Lea Wolf', amount: 9900, received: d(-60), vstatus: 'On Hold', onHoldSince: d(-35), entscheidung: d(-1) }),
  D({ id: 'd10', pid: 'PID-3355-NBG', an: '2026100095', name: 'Kaffeebecher To-go', company: 'Stadtwerke Elbtal', contact: 'Peter Lang', seller: 'Lea Wolf', amount: 5100, received: d(-35), vstatus: 'Gewonnen', entscheidung: d(-4), chance: 'Close to an order', prio: '1' }),
  D({ id: 'd11', pid: 'PID-3318-NBG', an: '2026090131', name: 'Isolierflaschen Team', company: 'Velotec Fahrräder GmbH', contact: 'Jana Roth', seller: 'Jonas Weber', amount: 6600, received: d(-70), vstatus: 'Verloren', lostReason: 'Preis', entscheidung: d(-40), chance: 'Low chance' }),
  D({ id: 'd12', pid: 'PID-3301-NBG', an: '2026090112', name: 'Strandtücher Sommer', company: 'Lumina Kosmetik GmbH', contact: 'Eva Martin', seller: 'Mara Hoffmann', amount: 9200, received: d(-80), vstatus: 'Verloren', lostReason: 'Lieferzeit', entscheidung: d(-50) })
];
