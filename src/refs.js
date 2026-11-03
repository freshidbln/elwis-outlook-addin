// Erkennt Projekt- und Auftragsnummern in Betreff (und optional Text) einer E-Mail.
// Abgeleitet aus den tatsächlich verwendeten Schreibweisen bei fresh id, z. B.:
//   „PID 3337“, „PID-1850-HKG“, „PID-1866//“, „NBG PID 3274“, „NBG #2704“, „HK#1426“,
//   „HK project#1551“, „Project#1426“, „// 619157 //“, „376070n“, „SC-02001421“,
//   Elwis-Belegnummern „2026090318“.
// Nicht als eigene Nummer gewertet: Auftragsbestätigungen/SC/PO von Lieferanten
// (z. B. „Auftragsbestätigung Nr. 308814“, „SC2299296_1“, „PO-00066“).

const OFFICE = (s) => (!s ? null : /^hk/i.test(s) ? 'HKG' : 'NBG');

export function parseRefs(text = '') {
  const t = ' ' + text + ' ';
  const pids = [];
  const addPid = (num, office, at) => {
    const n = String(Number(num));
    const ex = pids.find((p) => p.num === n);
    if (ex) { if (!ex.office && office) ex.office = office; ex.at = Math.min(ex.at, at); return; }
    pids.push({ num: n, office, at });
  };

  // „PID 3337“, „PID-1850-HKG“, „PID-1866//“, „PID: 3207 NBG“
  for (const m of t.matchAll(/\bPID[\s\-#:]*(\d{3,5})(?:\s*[-\/]?\s*(NBG|HKG|HK)\b)?/gi)) addPid(m[1], OFFICE(m[2]), m.index);
  // „NBG PID 3274“, „NBG #2704“, „HK#1426“, „HK project#1551“
  for (const m of t.matchAll(/\b(NBG|HKG|HK)\s*(?:PID|project|projekt)?\s*#?\s*[-:]?\s*(\d{3,5})\b/gi)) addPid(m[2], OFFICE(m[1]), m.index);
  // „Project#1426“ ohne Büro – wird in HK so geschrieben
  for (const m of t.matchAll(/\bproje[ck]t\s*#\s*(\d{3,5})\b/gi)) addPid(m[1], 'HKG', m.index);

  // SC-Nummern: 6-stellig, beginnen mit 37 (NBG) oder 61 (Berlin), optional mit Buchstabe („376070n“)
  const scs = [];
  for (const m of t.matchAll(/(?<![\d.,])((?:3[67]|6[01])\d{4})[a-z]?(?![\d.,])/gi)) {
    const before = t.slice(Math.max(0, m.index - 25), m.index);
    if (/(auftragsbest[äa]tigung|ab|order confirmation|rechnung|invoice)\s*(nr\.?|no\.?)?\s*$/i.test(before)) continue;
    if (!scs.includes(m[1])) scs.push(m[1]);
  }
  // HKG-SC „SC-02001421“
  for (const m of t.matchAll(/\bSC-(\d{8})\b/g)) if (!scs.includes('SC-' + m[1])) scs.push('SC-' + m[1]);

  // Elwis-Belegnummern (Angebot, Auftrag, Rechnung): 10-stellig, beginnen mit 20
  const docs = [...new Set([...t.matchAll(/(?<!\d)(20\d{8})(?!\d)/g)].map((m) => m[1]))];

  pids.sort((a, b) => a.at - b.at);
  return { pids: pids.map(({ num, office }) => ({ num, office })), scs, docs };
}
