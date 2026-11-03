// Auftrags-Status (Projektmanagement) – Felder, Phasen und Regeln.
// Entspricht dem bisherigen Monday-Board „Auftragsabwicklung“, aber mit Elwis als
// einziger Datenquelle. In der Live-Version liegen die Felder am Elwis-Auftrag
// (Sales Order) und die Regeln laufen als Server-Abläufe in Elwis.

const C = {
  green: '#00c875', dgreen: '#037f4c', orange: '#fdab3d', red: '#e2445c', blue: '#0086c0',
  lblue: '#579bfc', sky: '#5bc0eb', purple: '#a25ddc', violet: '#7e3b8a', teal: '#2fb8b0',
  brown: '#7f5347', grey: '#9aa1ab', pink: '#ff158a', yellow: '#ffcb00'
};

// elwis: vorhandenes Elwis-Feld am Auftrag, das genutzt wird | neu: Feld muss ergänzt werden
export const STATUS_FIELDS = [
  { key: 'ab', label: 'Auftragsbestätigung', short: 'AB', neu: true,
    options: [['in Arbeit', C.orange], ['Raus an Kunden', C.green], ['Offen', C.red]] },
  { key: 'usd', label: 'US Dollar', short: 'USD', neu: true,
    options: [['Nicht notwendig', C.grey], ['Offen', C.red], ['Erledigt', C.green]] },
  { key: 'korri', label: 'Korrekturabzug', short: 'Korri', neu: true,
    options: [['Angefordert', C.orange], ['warte auf Freigabe', C.blue], ['Freigegeben', C.green], ['Verzögert', C.red], ['ohne Werbeanbringung', C.purple]] },
  { key: 'korriDate', label: 'Freigabetermin Korri', short: 'Termin Korri', type: 'date', neu: true },
  { key: 'marks', label: 'Shippingmarks + Label', short: 'Shippingmarks', elwis: 'Shipping Marks Approved (Datum) – ergänzt um Status',
    options: [['Nicht notwendig', C.grey], ['Angefordert', C.orange], ['Freigegeben', C.green]] },
  { key: 'muster', label: 'Freigabemuster', short: 'Muster', neu: true,
    options: [['ohne Freigabemuster', C.dgreen], ['mit Freigabemuster', C.purple], ['in Produktion', C.orange], ['warte auf Freigabe', C.blue], ['Freigegeben', C.green], ['verzögert', C.red]] },
  { key: 'musterDate', label: 'Freigabetermin Muster', short: 'Termin Muster', type: 'date', neu: true },
  { key: 'prod', label: 'Produktion', short: 'Produktion', elwis: 'Production Stage',
    options: [['Freigabeprozess', C.blue], ['in Produktion', C.orange], ['verspätet', C.red], ['Fertiggestellt', C.green]] },
  { key: 'fertig', label: 'Fertigstellung (geplant)', short: 'Fertigstellung', type: 'date', elwis: 'SC Ready Date' },
  { key: 'fracht', label: 'Frachtweg', short: 'Frachtweg', elwis: 'Mode of Shipment',
    options: [['DHL etc./Kurier', C.blue], ['Seefracht', C.teal], ['Luftfracht', C.lblue], ['Zugfracht', C.brown], ['LKW', C.purple]] },
  { key: 'liefer', label: 'Lieferstatus', short: 'Lieferstatus', neu: true,
    options: [['Unterwegs', C.blue], ['Nachlauf/Auslieferung', C.sky], ['Verspätet', C.red], ['Angekommen', C.green]] },
  { key: 'eta', label: 'ETA', short: 'ETA', type: 'date', elwis: 'ETA' },
  { key: 'ltLief', label: 'Liefertermin Lieferant', short: 'LT Lieferant', type: 'date', elwis: 'Live PO Ready Date' },
  { key: 'ltKunde', label: 'Liefertermin Kunde', short: 'LT Kunde', type: 'date', elwis: 'Delivery Date' },
  { key: 'lieferung', label: 'Lieferung', short: 'Lieferung', neu: true,
    options: [['Ausgeliefert', C.green], ['Fertig zum Abrechnen', C.violet]] }
];
export const FIELD = Object.fromEntries(STATUS_FIELDS.map((f) => [f.key, f]));

export function colorOf(key, value) {
  const f = FIELD[key];
  const o = f?.options?.find(([l]) => l === value);
  return o ? o[1] : null;
}

// Phasen = Gruppen im Board. Werden aus den Status berechnet, nicht von Hand verschoben.
export const PHASES = [
  { key: 'auftrag', label: 'Auftrag', short: 'Auftrag', color: C.yellow },
  { key: 'freigabe', label: 'Im Freigabeprozess', short: 'Freigabeprozess', color: C.purple },
  { key: 'produktion', label: 'Freigegeben / in Produktion', short: 'In Produktion', color: C.pink },
  { key: 'fracht', label: 'Seefracht / Zugfracht / Luftfracht / LKW', short: 'Fracht', color: C.red },
  { key: 'auslieferung', label: 'In Auslieferung', short: 'Auslieferung', color: C.green },
  { key: 'abrechnung', label: 'Fertig zum Abrechnen', short: 'Abrechnen', color: C.lblue },
  { key: 'abruf', label: 'Abrufaufträge', short: 'Abruf', color: C.grey }
];
export const PHASE = Object.fromEntries(PHASES.map((p) => [p.key, p]));

export function computePhase(o) {
  if (o.lieferung === 'Fertig zum Abrechnen') return 'abrechnung';
  if (o.abruf) return 'abruf';
  if (['Nachlauf/Auslieferung', 'Angekommen'].includes(o.liefer) || o.lieferung === 'Ausgeliefert') return 'auslieferung';
  if (o.prod === 'Fertiggestellt') return o.fracht && o.fracht !== 'DHL etc./Kurier' ? 'fracht' : 'auslieferung';
  if (['in Produktion', 'verspätet'].includes(o.prod)) return 'produktion';
  if (o.prod === 'Freigabeprozess' || ['warte auf Freigabe', 'Verzögert'].includes(o.korri) ||
      ['mit Freigabemuster', 'warte auf Freigabe', 'in Produktion', 'verzögert'].includes(o.muster)) return 'freigabe';
  return 'auftrag';
}

// Regeln – zusammengefasst aus den 20 Monday-Automationen. Text dient auch als Vorgabe für Softacus.
export const RULES = [
  { id: 'R1', when: 'Korrekturabzug → „warte auf Freigabe“', then: 'Produktion = Freigabeprozess (Phase: Im Freigabeprozess)' },
  { id: 'R2', when: 'Korrekturabzug → „Freigegeben“ und Freigabemuster = ohne Freigabemuster', then: 'Produktion = in Produktion' },
  { id: 'R3', when: 'Korrekturabzug → „Freigegeben“ und Freigabemuster = mit Freigabemuster', then: 'Freigabemuster = in Produktion (Muster wird gefertigt)' },
  { id: 'R4', when: 'Freigabemuster → „Freigegeben“', then: 'Produktion = in Produktion' },
  { id: 'R5', when: 'Auftragsbestätigung → „Raus an Kunden“, Korrekturabzug = ohne Werbeanbringung, Freigabemuster = ohne Freigabemuster', then: 'Produktion = in Produktion' },
  { id: 'R6', when: 'Produktion → „Fertiggestellt“ (oder Frachtweg wird danach gesetzt)', then: 'Lieferstatus = Unterwegs (See/Zug/Luft/LKW) bzw. Nachlauf/Auslieferung (Kurier)' },
  { id: 'R7', when: 'Lieferung → „Fertig zum Abrechnen“', then: 'Elwis-Feld „Okay to Invoice“ = ja → bestehende Mail an Buchhaltung' },
  { id: 'T1', when: 'Täglich: Freigabetermin Korri überschritten und Korrekturabzug = warte auf Freigabe', then: 'Korrekturabzug = Verzögert' },
  { id: 'T2', when: 'Täglich: Freigabetermin Muster überschritten und Freigabemuster = warte auf Freigabe', then: 'Freigabemuster = verzögert' },
  { id: 'T3', when: 'Täglich: geplante Fertigstellung überschritten und Produktion = in Produktion', then: 'Produktion = verspätet' },
  { id: 'T4', when: 'Täglich: ETA überschritten und Lieferstatus = Unterwegs', then: 'Lieferstatus = Verspätet' },
  { id: 'T5', when: 'Täglich: Liefertermin Lieferant überschritten, Phase In Auslieferung, noch nicht angekommen', then: 'Lieferstatus = Verspätet' }
];

// Wendet eine Änderung an und führt die Folge-Regeln aus.
// Rückgabe: { patch, applied: [Text...] }
export function applyChange(order, key, value) {
  const o = { ...order };
  const patch = {};
  const applied = [];
  const set = (k, v, ruleId) => {
    if (o[k] === v) return;
    o[k] = v; patch[k] = v;
    if (ruleId) applied.push(`${ruleId}: ${FIELD[k]?.label || k} → ${v === true ? 'ja' : v}`);
  };
  const before = computePhase(order);
  set(key, value || '');

  if (key === 'korri' && value === 'warte auf Freigabe' && !['in Produktion', 'Fertiggestellt'].includes(o.prod)) set('prod', 'Freigabeprozess', 'R1');
  if (key === 'korri' && value === 'Freigegeben' && o.muster === 'ohne Freigabemuster') set('prod', 'in Produktion', 'R2');
  if (key === 'korri' && value === 'Freigegeben' && o.muster === 'mit Freigabemuster') set('muster', 'in Produktion', 'R3');
  if (key === 'muster' && value === 'Freigegeben' && o.prod !== 'Fertiggestellt') set('prod', 'in Produktion', 'R4');
  if (key === 'ab' && value === 'Raus an Kunden' && o.korri === 'ohne Werbeanbringung' && o.muster === 'ohne Freigabemuster') set('prod', 'in Produktion', 'R5');
  if ((key === 'prod' && value === 'Fertiggestellt') || (key === 'fracht' && o.prod === 'Fertiggestellt' && !o.liefer)) {
    if (o.fracht === 'DHL etc./Kurier') set('liefer', 'Nachlauf/Auslieferung', 'R6');
    else if (o.fracht) set('liefer', 'Unterwegs', 'R6');
  }
  if (key === 'lieferung' && value === 'Fertig zum Abrechnen') { o.okayToInvoice = true; patch.okayToInvoice = true; applied.push('R7: Elwis „Okay to Invoice“ = ja → Mail an Buchhaltung'); }

  const after = computePhase(o);
  if (after !== before) applied.push(`Phase: ${PHASE[before].label} → ${PHASE[after].label}`);
  return { patch, applied, order: o };
}

// Termin-Regeln (laufen live täglich als geplanter Elwis-Ablauf)
export function applyDateRules(order, todayIso) {
  const o = { ...order };
  const hits = [];
  const past = (d) => d && d < todayIso;
  if (past(o.korriDate) && o.korri === 'warte auf Freigabe') { o.korri = 'Verzögert'; hits.push('T1'); }
  if (past(o.musterDate) && o.muster === 'warte auf Freigabe') { o.muster = 'verzögert'; hits.push('T2'); }
  if (past(o.fertig) && o.prod === 'in Produktion') { o.prod = 'verspätet'; hits.push('T3'); }
  if (past(o.eta) && o.liefer === 'Unterwegs') { o.liefer = 'Verspätet'; hits.push('T4'); }
  if (past(o.ltLief) && computePhase(o) === 'auslieferung' && !['Angekommen'].includes(o.liefer) && o.lieferung !== 'Ausgeliefert') { o.liefer = 'Verspätet'; hits.push('T5'); }
  return { order: o, hits };
}

export const todayIso = () => {
  const t = new Date();
  return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
};
