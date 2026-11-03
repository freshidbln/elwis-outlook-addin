// Tests mit Schreibweisen wie in echten Betreffzeilen (Inhalte verfremdet)
import { parseRefs } from '../src/refs.js';
import assert from 'node:assert/strict';
const cases = [
  ['AW: Regenjacke / PID 3337', { pids: [{ num: '3337', office: null }], scs: [] }],
  ['Re: PID-1850-HKG/ Inquiry oven mitts // 619157', { pids: [{ num: '1850', office: 'HKG' }], scs: ['619157'] }],
  ['RE: PID-1866// Inquiry bracelet // 619130', { pids: [{ num: '1866', office: null }], scs: ['619130'] }],
  ['AW: HK project#1551: Inquiry bag // 618982 // NBG PID 3274', { pids: [{ num: '1551', office: 'HKG' }, { num: '3274', office: 'NBG' }], scs: ['618982'] }],
  ['AW: Project#1426 / Inquiry Ice Mould // NBG #2704 / HK#1426 / SC-02001421 / PO-00066', { pids: [{ num: '1426', office: 'HKG' }, { num: '2704', office: 'NBG' }], scs: ['SC-02001421'] }],
  ['RE: Customer order // 376070n // PID 3329', { pids: [{ num: '3329', office: null }], scs: ['376070'] }],
  ['AW: 261029 >>> Auftragsbestätigung Nr. 308814 <<< Fresh id Products', { pids: [], scs: [] }],
  ['WG: Artwork SC2299296_1 / PO-00066 nun freigeben', { pids: [], scs: [] }],
  ['AW: fresh id / PID-3201-NBG / Sommer-Display', { pids: [{ num: '3201', office: 'NBG' }], scs: [] }],
  ['Zahlung Rechnung 2026090318', { pids: [], scs: [], docs: ['2026090318'] }],
  ['Anfrage Verkaufsständer für unsere Kaffeebar', { pids: [], scs: [] }]
];
let ok = 0;
for (const [subj, exp] of cases) {
  const r = parseRefs(subj);
  assert.deepEqual(r.pids, exp.pids, subj + ' pids ' + JSON.stringify(r.pids));
  assert.deepEqual(r.scs, exp.scs, subj + ' scs ' + JSON.stringify(r.scs));
  if (exp.docs) assert.deepEqual(r.docs, exp.docs, subj);
  ok++;
}
console.log(`${ok}/${cases.length} Betreff-Muster korrekt erkannt`);
