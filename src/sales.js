// Vertriebsphase (vor dem Auftrag) – ersetzt das Monday-Board „Anfrage & Vertrieb“.
// Elwis bildet den Vertrieb bereits weitgehend ab (Inquiry Status, Quote Status Won/Lost,
// Closing Chance, Closing Date, Priority, Lost Reason). Der „Vertriebsstatus“ fasst
// Anfrage- und Angebotsstatus zu einer Pipeline zusammen; nur wenige Werte sind neu.

const C = {
  slate: '#7b8794', orange: '#fdab3d', yellow: '#e8b000', green: '#00c875', dgreen: '#037f4c', purple: '#a25ddc',
  dark: '#4b5563', red: '#e2445c', blue: '#0086c0', lblue: '#579bfc', teal: '#2fb8b0', lgreen: '#7cb342'
};

export const SALES_FIELDS = [
  { key: 'vstatus', label: 'Vertriebsstatus', short: 'Status',
    elwis: 'Inquiry Status + Quote Status (zusammengefasst)',
    options: [
      ['Offen', C.slate, 'Inquiry Status = Open'],
      ['Angefragt (Lieferant)', C.orange, 'neu: Inquiry Status „Angefragt“'],
      ['Angebot schreiben', C.yellow, 'Quote Status = Draft'],
      ['Kontrolle (Preise)', C.teal, 'neu: Quote Status „In Kontrolle“ – Kalkulation liegt beim Freigeber, der die Preise macht'],
      ['Preise freigegeben', C.lgreen, 'neu: Quote Status „Preise freigegeben“ – Kalkulation freigegeben, Preise stehen im Angebot'],
      ['Raus an Kunden', C.green, 'Quote Status = Send to customer'],
      ['In Verhandlung', C.purple, 'neu: Quote Status „In Verhandlung“'],
      ['On Hold', C.dark, 'neu: Inquiry Status „On Hold“'],
      ['Gewonnen', C.dgreen, 'Quote Status = Won'],
      ['Verloren', C.red, 'Quote Status = Lost'],
      ['Übergeben', C.blue, 'Sales Order angelegt (Button „Create Sales Order from Quote“)']
    ].map(([l, c, m]) => [l, c, m]) },
  { key: 'chance', label: 'Closing Chance', short: 'Chance', elwis: 'Closing Chance',
    options: [['Close to an order', C.green], ['Neutral', C.orange], ['Low chance', C.red]] },
  { key: 'prio', label: 'Priorität', short: 'Prio', elwis: 'Priority',
    options: [['1', C.green], ['2', C.orange], ['3', C.red]] },
  { key: 'entscheidung', label: 'Entscheidung bis / Angebot bis', short: 'Entscheidung bis', type: 'date', elwis: 'Closing Date' },
  { key: 'lostReason', label: 'Verlustgrund', short: 'Verlustgrund', elwis: 'Lost Reason (Werte mit Elwis abgleichen)',
    options: [['Preis', C.red], ['Lieferzeit', C.orange], ['Wettbewerb', C.purple], ['Projekt gestoppt', C.dark], ['Sonstiges', C.slate]] }
];
export const SALES_FIELD = Object.fromEntries(SALES_FIELDS.map((f) => [f.key, f]));
// Freitext wie die Monday-Spalte „Kommentar / Übergabe“ (51 von 879 Deals nutzen sie)
export const SALES_COMMENT = { key: 'comment', label: 'Kommentar / Übergabe', elwis: 'Notiz am Projekt (vorhanden); geht bei der Übergabe als Übergabe-Notiz an den Auftrag' };

export const ACTIVE = ['Offen', 'Angefragt (Lieferant)', 'Angebot schreiben', 'Kontrolle (Preise)', 'Preise freigegeben', 'Raus an Kunden', 'In Verhandlung'];

export const SALES_PHASES = [
  { key: 'anfragen', label: 'Anfragen', short: 'Anfrage', color: '#579bfc' },
  { key: 'vertrieb', label: 'Angebot beim Kunden', short: 'Angebot', color: '#00c875' },
  { key: 'onhold', label: 'On Hold', short: 'On Hold', color: '#4b5563' },
  { key: 'gewonnen', label: 'Gewonnen / in Übergabe', short: 'Gewonnen', color: '#037f4c' },
  { key: 'verloren', label: 'Nicht realisiert', short: 'Verloren', color: '#e2445c' },
  { key: 'uebergeben', label: 'An Auftragsabwicklung übergeben', short: 'Übergeben', color: '#0086c0' }
];
export const SALES_PHASE = Object.fromEntries(SALES_PHASES.map((p) => [p.key, p]));

export function computeSalesPhase(d) {
  switch (d.vstatus) {
    case 'Raus an Kunden': case 'In Verhandlung': return 'vertrieb';
    case 'On Hold': return 'onhold';
    case 'Gewonnen': return 'gewonnen';
    case 'Verloren': return 'verloren';
    case 'Übergeben': return 'uebergeben';
    default: return 'anfragen';
  }
}

// Überfällig ist ein Hinweis, kein Status – der echte Status bleibt erhalten.
export function overdueDays(d, todayIso) {
  if (!d.entscheidung || !ACTIVE.includes(d.vstatus)) return 0;
  if (d.entscheidung >= todayIso) return 0;
  return Math.round((new Date(todayIso) - new Date(d.entscheidung)) / 86400000);
}

export const SALES_RULES = [
  { id: 'V1', when: 'Neue Anfrage angelegt', then: 'Vertriebsstatus = Offen, Verkäufer = anlegende Person' },
  { id: 'V2', when: 'Vertriebsstatus → „Raus an Kunden“ und „Entscheidung bis“ leer', then: '„Entscheidung bis“ = heute + 14 Tage (Nachfassen)' },
  { id: 'V3', when: 'Vertriebsstatus → „On Hold“', then: '„On Hold seit“ = heute; Wiedervorlage über „Entscheidung bis“ (leer → heute + 30 Tage); Info an Verkäufer' },
  { id: 'V4', when: 'Vertriebsstatus → „Verloren“', then: 'Verlustgrund wird abgefragt (Pflichtfeld)' },
  { id: 'V5', when: 'Vertriebsstatus → „Übergeben“', then: 'Auftrag (Sales Order) wird angelegt und erscheint in der Auftragsabwicklung' },
  { id: 'V6', when: 'Vertriebsstatus → „Kontrolle (Preise)“', then: 'Aufgabe „Preise machen“ an den Freigeber; fehlt die Kalkulation, zeigt die Vertriebskarte einen Hinweis' },
  { id: 'V7', when: 'Kalkulation „Zur Freigabe senden“ (Teammitglied + Kommentar)', then: 'Vertriebsstatus = Kontrolle (Preise), „Freigabe durch“ = Freigeber, Benachrichtigung an den Freigeber, Eintrag im Protokoll' },
  { id: 'V8', when: 'Kalkulation „Freigeben“ (Freigeber) oder „Selbst freigeben“ (ohne DB-Schwelle)', then: 'Neue Version (Datum, kalkuliert von, freigegeben von); Vertriebsstatus = Preise freigegeben; Preise als Positionen ins Elwis-Angebot zur PID; Benachrichtigung an den Anforderer' },
  { id: 'V9', when: 'Kalkulation „Zurückgeben“ (Kommentar Pflicht) oder Anfrage zurückziehen', then: 'Vertriebsstatus = Angebot schreiben, Benachrichtigung mit Kommentar, Eintrag im Protokoll' },
  { id: 'T1', when: 'Täglich: „Entscheidung bis“ überschritten, Status aktiv', then: 'Markierung „überfällig“ + Erinnerung an Verkäufer (Status bleibt erhalten)' },
  { id: 'T2', when: 'Täglich: Wiedervorlage erreicht, Status On Hold', then: 'Vertriebsstatus = Offen (zurück in „Anfragen“) + Info an Verkäufer' }
];

const addDays = (iso, n) => { const [y, m, d] = iso.split('-').map(Number); const x = new Date(y, m - 1, d + n); return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`; };

export function applySalesChange(deal, key, value, todayIso) {
  const d = { ...deal };
  const patch = {};
  const applied = [];
  let createOrder = false;
  const set = (k, v, rule) => {
    if (d[k] === v) return;
    d[k] = v; patch[k] = v;
    const shown = /^\d{4}-\d{2}-\d{2}$/.test(v) ? v.split('-').reverse().join('.') : v;
    if (rule) applied.push(`${rule}: ${SALES_FIELD[k]?.label || (k === 'onHoldSince' ? 'On Hold seit' : k)} → ${shown}`);
  };
  const before = computeSalesPhase(deal);
  set(key, value || '');
  if (key === 'vstatus') {
    if (value === 'Raus an Kunden' && !d.entscheidung) set('entscheidung', addDays(todayIso, 14), 'V2');
    if (value === 'On Hold') { set('onHoldSince', todayIso, 'V3'); if (!d.entscheidung || d.entscheidung < todayIso) set('entscheidung', addDays(todayIso, 30), 'V3'); }
    if (value === 'Verloren' && !d.lostReason) applied.push('V4: Bitte Verlustgrund wählen');
    if (value === 'Kontrolle (Preise)') applied.push('V6: Aufgabe „Preise machen“ an den Vorgesetzten');
    if (value === 'Übergeben' && deal.vstatus !== 'Übergeben') { createOrder = true; applied.push('V5: Auftrag angelegt → erscheint in der Auftragsabwicklung'); }
  }
  const after = computeSalesPhase(d);
  if (after !== before) applied.push(`Gruppe: ${SALES_PHASE[before].label} → ${SALES_PHASE[after].label}`);
  return { patch, applied, deal: d, createOrder };
}

export function applySalesDateRules(deal, todayIso) {
  const d = { ...deal };
  if (d.vstatus === 'On Hold' && d.entscheidung && d.entscheidung <= todayIso) { d.vstatus = 'Offen'; d.reactivated = true; }
  return d;
}
