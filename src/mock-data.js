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
    id: 'p1', kind: 'inquiry', office: 'NBG', pid: 'PID-3201-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects', name: 'Sommer-Display Kühlschrankaufsteller',
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
    id: 'p2', kind: 'inquiry', office: 'NBG', pid: 'PID-3215-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects', name: 'Weihnachts-Geschenkset 2026',
    companyId: 'c1', contactId: 'k2', status: 1, received: d(-6), deadline: d(3), amount: 18000, currency: 'EUR',
    brand: 'Nordlicht Winterbock', category: 'Paper', priority: 2, closingChance: 2, projectLead: STAFF.anna, team: [STAFF.anna, STAFF.tom],
    samples: [{ id: 's2', kind: 'sample', name: 'Geschenkbox mit Inlay – Muster', status: 1, date: d(-4), deadline: d(9), amount: 220 }],
    quotes: [], orders: [], shipments: [], invoices: [],
    tasks: [{ id: 't2', kind: 'task', name: 'Kalkulation an Kunden senden', assignedTo: STAFF.tom, deadline: d(1), done: false }],
    notes: [{ date: d(-5), author: STAFF.anna, text: 'Briefing erhalten: 2.000 Sets, Auslieferung Anfang Dezember.' }]
  },
  {
    id: 'p3', kind: 'inquiry', office: 'NBG', pid: 'PID-3188-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects', name: 'Thekenaufsteller Acryl',
    companyId: 'c2', contactId: 'k3', status: 2, received: d(-25), deadline: d(-14), amount: 8940, currency: 'EUR',
    brand: 'Alpenkräuter Original', category: 'Plastic', priority: 2, closingChance: 1, projectLead: STAFF.jonas, team: [STAFF.jonas],
    samples: [],
    quotes: [{ id: 'q3', kind: 'quote', number: '2026100087', name: 'Thekenaufsteller Acryl – 500 Stk.', status: 3, date: d(-10), validTo: d(12), amount: 8940, currency: 'EUR' }],
    orders: [], shipments: [], invoices: [],
    tasks: [{ id: 't3', kind: 'task', name: 'Nachfassen Angebot 2026100087', assignedTo: STAFF.jonas, deadline: d(0), done: false }],
    notes: []
  },
  {
    id: 'p4', kind: 'inquiry', office: 'NBG', pid: 'PID-3142-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects', name: 'Neon-Leuchtschild Bar-Edition',
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
    id: 'p5', kind: 'inquiry', office: 'NBG', pid: 'PID-3230-NBG', folder: 'https://freshidproducts.sharepoint.com/sites/Nuremberg/Projects', name: 'Pop-up Store Kit UK',
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
  { label: 'Markus Vogel – Zahlung Rechnung 2026090318', fromName: 'Markus Vogel', fromEmail: 'markus.vogel@alpenkraeuter.example',
    subject: 'Zahlung Rechnung 2026090318', received: d(0) },
  { label: 'Kevin Liu – Produktionsstand (HKG)', fromName: 'Kevin Liu', fromEmail: 'kevin.liu@brightpack.example',
    subject: 'RE: PID-1874-HKG production update', received: d(-1) },
  { label: 'Sarah Klein – neue Anfrage (Lead)', fromName: 'Sarah Klein', fromEmail: 'sarah@morgenrot-kaffee.example',
    subject: 'Anfrage Verkaufsständer für unsere Kaffeebar', received: d(-1) },
  { label: 'Unbekannter Absender', fromName: 'Info', fromEmail: 'info@unbekannt.example',
    subject: 'Ihre Anfrage', received: d(-2) }
];

// Vorschlag für eine Standard-Ordnerstruktur je Projekt (wird mit fresh id abgestimmt)
export const FOLDER_TEMPLATE = [
  { key: 'brief', name: '01 Anfrage & Briefing' },
  { key: 'design', name: '02 Design & Artwork' },
  { key: 'quote', name: '03 Kalkulation & Angebote' },
  { key: 'vendor', name: '04 Lieferanten' },
  { key: 'sample', name: '05 Muster & Freigaben' },
  { key: 'order', name: '06 Auftrag & Bestellung' },
  { key: 'qc', name: '07 Produktion & QC' },
  { key: 'shipping', name: '08 Versand & Zoll' },
  { key: 'invoice', name: '09 Rechnungen' },
  { key: 'correspondence', name: '10 Korrespondenz' }
];
