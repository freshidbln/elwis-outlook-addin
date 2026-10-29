// Status-Werte und Farben wie in Elwis konfiguriert (Stand 29.10.2026).
export const STATUS = {
  inquiry: { 1: ['Open', '#d6eeff'], 2: ['Transferred to Quote', '#ccfff1'], 3: ['Cancelled', '#ffdede'] },
  quote: { 1: ['Draft', '#e0e0e0'], 2: ['Approved', '#d6eeff'], 3: ['Send to customer', '#ccfff1'], 4: ['Won', '#cdffcc'], 5: ['Lost', '#ffdede'] },
  order: { 8: ['Draft', '#cfcfcf'], 1: ['Open', '#cfcfcf'], 2: ['Pending Prepayment', '#cfe5ff'], 3: ['Released', '#eaccff'], 4: ['Shipped', '#e8fffd'], 5: ['Invoiced', '#cdffcc'], 6: ['Cancelled', '#ffdede'], 7: ['Completed/Closed', '#bfe8bd'] },
  invoice: { 1: ['Draft', '#e0e0e0'], 2: ['Planned', '#adffe9'], 3: ['Approved', '#cdffcc'], 4: ['Sent-Open', '#c2e0ff'], 5: ['Partially Paid', '#e3fff2'], 6: ['Paid', '#cdffcc'], 7: ['Cancelled', '#f2aaaa'], 8: ['Credited', '#c1abff'] },
  payment: { 1: ['Open', '#ffd6d6'], 3: ['Partially paid', '#ffecb3'], 2: ['Paid', '#d4f5cc'], 4: ['On Hold', '#e0e0e0'] },
  sample: { 1: ['Ongoing/In Production', '#d6eeff'], 2: ['Completed', '#cdffcc'], 3: ['Cancelled/Lost', '#ffdede'], 4: ['Inactive', '#e0e0e0'] },
  shipment: { 1: ['Ongoing/In transit', '#d6eeff'], 2: ['Completed', '#cdffcc'], 3: ['Cancelled', '#ffdede'], 4: ['Closed', '#e0e0e0'], 5: ['Inactive', '#e0e0e0'] },
  company: { 1: ['Draft', '#dbdbdb'], 2: ['Active', '#ccffcd'], 3: ['Passive', '#cce6ff'], 4: ['Blocked', '#fccccc'], 5: ['Inactive', '#dbdbdb'] },
  closingChance: { 1: ['Close to an order', '#c7ffcd'], 2: ['Neutral', '#ffe3bd'], 3: ['Low chance of an order', '#ffd9d9'] },
  contactType: { 2: ['Customer Contact', '#d1ffe9'], 1: ['Vendor Contact', '#d9eaff'], 3: ['Lead', '#fff3c4'] },
  priority: { 1: ['Prio 1', '#d9f7d4'], 2: ['Prio 2', '#ffecb3'], 3: ['Prio 3', '#ffd6d6'] }
};

export function status(kind, code) {
  const s = STATUS[kind] && STATUS[kind][code];
  return s ? { label: s[0], color: s[1] } : { label: code == null ? '–' : String(code), color: '#eeeeee' };
}

// Deutsche Anzeigenamen der Vorgangsarten
export const KIND_LABEL = {
  inquiry: 'Anfrage', sample: 'Muster', quote: 'Angebot', order: 'Auftrag',
  shipment: 'Versand', invoice: 'Rechnung', task: 'Aufgabe', customer: 'Kunde', vendor: 'Lieferant', contact: 'Kontakt'
};
