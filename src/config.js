// Zentrale Konfiguration des Add-ins.
// mode: 'demo' = erfundene Beispieldaten, kein Zugriff auf Elwis
//       'live' = Daten über das Backend (siehe /backend), nur lesend in Stufe 1
export const CONFIG = {
  mode: 'demo',
  version: '1.4.2',
  elwisBaseUrl: 'https://freshid.elwis.cloud',
  backendUrl: '', // z. B. https://elwis-addin-api.example.net – wird für die Live-Version gesetzt

  // Detailseiten in Elwis (Seiten-Alias), aufgerufen mit ?artifactId=<id>
  detailPages: {
    NBG: {
      inquiry: 'inquiry-page-main', sample: 'sample-page', quote: 'quote-page',
      order: 'order-page', shipment: 'shipment-page', invoice: 'invoice-page',
      customer: 'customer-page-2', vendor: 'supplier-page', contact: 'contact-page', task: 'taks-page'
    },
    HKG: {
      inquiry: 'inquiry-hkg-page', sample: 'sample-hkg-page', quote: 'quote-hkg-page',
      order: 'sales-order-nbg-page', shipment: 'shipment-hkg-page', invoice: 'invoice-hkg-page',
      customer: 'customer-hkg-page', vendor: 'vendor-hkg-page', contact: 'contact-hkg-page', task: 'task-hkg-page'
    }
  },

  // Listenseiten – in der Demo öffnet „In Elwis öffnen“ die passende Liste,
  // weil die Beispiel-Datensätze in Elwis nicht existieren.
  listPages: {
    NBG: {
      inquiry: 'all-inquiries-1', sample: 'erp-view---list-of-samples', quote: 'erp-view---list-of-quotes',
      order: 'erp-view---list-of-order', shipment: 'list-of-shipments', invoice: 'list-of-invoices',
      customer: 'all-customers-nbg', vendor: 'list-of-suppliers', contact: 'list-of-contacts', task: 'list-of-tasks'
    },
    HKG: {
      inquiry: 'list-of-inquiries-hkg', sample: 'untitled-gtr7zyh49unpy7n', quote: 'quotes-hkg',
      order: 'orders-hkg', shipment: 'untitled-wcifajf2cirsyrx', invoice: 'invoices-hkg',
      customer: 'list-of-customers-hkg', vendor: 'suppliers-hkg', contact: 'hkg-customers', task: 'list-of-tasks-hkg'
    }
  },

  // Erkennung von PID-, SC- und Belegnummern: siehe src/refs.js
};
