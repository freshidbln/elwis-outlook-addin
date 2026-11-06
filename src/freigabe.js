// Kalkulations-Freigabe (Vertrieb) – eine Stelle für alle Regeln, genutzt von Add-in, Board und Kalkulator.
//
// Ablauf
//   Zur Freigabe senden  → offene Anfrage an ein Teammitglied, Deal-Status „Kontrolle (Preise)“, „Freigabe durch“ = Freigeber,
//                          Benachrichtigung an den Freigeber                                           (Regel V7)
//   Freigeben            → neue Version (unveränderlich), Status „Preise freigegeben“, Preise ins Elwis-Angebot,
//                          Benachrichtigung an den Anforderer                                          (Regel V8)
//   Selbst freigeben     → wie Freigeben, ohne Anfrage (keine DB-Schwelle – Entscheidung 06.11.2026)  (Regel V8)
//   Zurückgeben          → Anfrage geschlossen mit Kommentar, Status „Angebot schreiben“, Benachrichtigung   (Regel V9)
//   Anfrage zurückziehen → nur durch den Anforderer, Status „Angebot schreiben“
// Jeder Schritt landet im Protokoll des Deals (wer, wann, was, an wen, Kommentar, Version).
//
// Demo: Speicher im Browser (gleicher Ursprung für Add-in, Board und Kalkulator).
// Live: dieselben Schritte als Backend-Aufrufe; Version und Protokoll in einer eigenen Tabelle, Preise als Positionen
// am Elwis-Angebot zur PID, Benachrichtigung per E-Mail/Outlook, Anwender = angemeldetes Outlook-Postfach.

import { demoDeals } from './data.js';
import { projects } from './mock-data.js';

export const LS_FREIGABE = 'elwis-demo-kalkulationen-v2';
export const LS_NACHRICHTEN = 'elwis-demo-benachrichtigungen-v1';
export const LS_ICH = 'elwis-demo-ich';

/** Team (Demo). Live: Elwis-Benutzer mit Rolle Vertrieb/Geschäftsleitung. */
export const TEAM = [
  { name: 'Anna Becker', rolle: 'Vertrieb' },
  { name: 'Tom Richter', rolle: 'Vertrieb' },
  { name: 'Jonas Weber', rolle: 'Vertrieb' },
  { name: 'Mei Lin Chan', rolle: 'Einkauf HKG' },
  { name: 'Katrin Seidel', rolle: 'Vertriebsleitung' }
];

const get = (k, d) => { try { return JSON.parse(localStorage.getItem(k) || JSON.stringify(d)); } catch { return d; } };
const set = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } };
const jetzt = () => new Date().toISOString();
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

// ---------- Anwender ----------
export function ich() { try { return localStorage.getItem(LS_ICH) || TEAM[0].name; } catch { return TEAM[0].name; } }
export function setIch(name) { try { localStorage.setItem(LS_ICH, name); } catch { /* ohne Speicher */ } }

// ---------- Daten je Deal ----------
/** { dealId, pid, projekt, kunde, anfrage|null, versionen[], log[] } */
export function kalkulation(dealId) { return get(LS_FREIGABE, {})[dealId] || null; }
function speichern(k) { const all = get(LS_FREIGABE, {}); all[k.dealId] = k; set(LS_FREIGABE, all); }
function basis(deal) {
  return kalkulation(deal.id) || { dealId: deal.id, pid: deal.pid || '', projekt: deal.name || '', kunde: deal.company || '', anfrage: null, versionen: [], log: [] };
}
/** Deal mit aktuellem Status (Demo-Speicher) */
export const dealById = (id) => demoDeals.get(id);
export const aktuelleVersion = (k) => (k?.versionen?.length ? k.versionen[k.versionen.length - 1] : null);

/** Status in einem Satz – für Karte, Board und Kalkulator */
export function statusText(k) {
  if (!k) return '';
  if (k.anfrage) return `Freigabe angefordert bei ${k.anfrage.an}`;
  const v = aktuelleVersion(k);
  if (v) return `Version ${v.nr} freigegeben${v.art === 'selbst' ? ' (selbst)' : ''} von ${v.freigegebenVon}`;
  return '';
}

// ---------- Benachrichtigungen ----------
function melden(an, text, dealId, art) {
  const list = get(LS_NACHRICHTEN, []);
  list.unshift({ id: uid(), an, at: jetzt(), text, dealId, art, gelesen: false });
  set(LS_NACHRICHTEN, list.slice(0, 200));
}
export function nachrichten(name = ich()) { return get(LS_NACHRICHTEN, []).filter((n) => n.an === name); }
export function ungelesen(name = ich()) { return nachrichten(name).filter((n) => !n.gelesen).length; }
/** eine Benachrichtigung als gelesen markieren; ohne id: alle des aktuellen Anwenders */
export function gelesen(id) {
  const l = get(LS_NACHRICHTEN, []);
  l.forEach((n) => { if (id ? n.id === id : n.an === ich()) n.gelesen = true; });
  set(LS_NACHRICHTEN, l);
}

// ---------- Hilfen ----------
function protokoll(k, eintrag) { k.log.push({ at: jetzt(), wer: ich(), ...eintrag }); }
function dealStatus(deal, vstatus, freigeber) {
  // über die Vertriebsregeln (applySalesChange), damit Board und Karte denselben Stand zeigen
  if (deal.vstatus !== vstatus) demoDeals.update(deal.id, 'vstatus', vstatus);
  if ((demoDeals.get(deal.id)?.freigeber || '') !== (freigeber || '')) demoDeals.update(deal.id, 'freigeber', freigeber || '');
}
/** Was beim Freigeben ins Elwis-Angebot geschrieben würde (Demo: nur angezeigt) */
function elwisAngebot(deal, staffeln) {
  const p = projects.find((x) => x.id === deal.projectId);
  const q = (p?.quotes || []).slice(-1)[0];
  return {
    angebot: q ? `Angebot ${q.number}` : `neues Angebot zur ${deal.pid || 'PID'}`,
    positionen: staffeln.map((s) => ({ menge: s.qty, vkStk: s.vkStk, gesamt: s.vkGesamt })),
    geschrieben: false // Demo: Elwis wird nicht verändert. Live: true, sobald die Positionen gespeichert sind.
  };
}
const fehler = (msg) => { throw new Error(msg); };

// ---------- Schritte ----------
/** Kalkulation zur Freigabe senden. erg = { staffeln, snapshot, korrekturen } aus dem Kalkulator */
export function anfordern(deal, an, kommentar, erg) {
  const k = basis(deal);
  if (k.anfrage) fehler(`Es läuft schon eine Freigabe-Anfrage bei ${k.anfrage.an}.`);
  if (!an || an === ich()) fehler('Bitte ein anderes Teammitglied als Freigeber wählen – oder selbst freigeben.');
  if (!erg?.staffeln?.length) fehler('Keine rechenbare Staffel.');
  k.anfrage = { id: uid(), at: jetzt(), von: ich(), an, kommentar: kommentar || '', ...erg };
  protokoll(k, { was: 'Freigabe angefordert', an, kommentar: kommentar || '' });
  speichern(k);
  melden(an, `${ich()} bittet um Freigabe der Kalkulation „${deal.name}“ (${deal.pid || '–'})${kommentar ? `: „${kommentar}“` : ''}`, deal.id, 'anfrage');
  dealStatus(deal, 'Kontrolle (Preise)', an);
  return k;
}

/** Freigeben – durch den angefragten Freigeber oder selbst (ohne offene Anfrage). erg = aktuelle Zahlen aus dem Kalkulator (optional). */
export function freigeben(deal, erg, kommentar = '') {
  const k = basis(deal);
  const a = k.anfrage;
  if (a && a.an !== ich()) fehler(a.von === ich() ? `Die Freigabe liegt bei ${a.an}. Erst die Anfrage zurückziehen, um selbst freizugeben.` : `Nur ${a.an} kann diese Anfrage freigeben.`);
  const daten = erg?.staffeln?.length ? erg : a;
  if (!daten?.staffeln?.length) fehler('Keine rechenbare Staffel.');
  const angepasst = !!(a && erg?.staffeln && JSON.stringify(erg.staffeln) !== JSON.stringify(a.staffeln));
  const v = {
    nr: k.versionen.length + 1, at: jetzt(),
    kalkuliertVon: a ? a.von : ich(), freigegebenVon: ich(), art: a ? 'freigabe' : 'selbst',
    angefordertAm: a?.at || null, anfrageKommentar: a?.kommentar || '', preiseAngepasst: angepasst, kommentar,
    staffeln: daten.staffeln, korrekturen: daten.korrekturen || [], snapshot: daten.snapshot || null,
    elwis: elwisAngebot(deal, daten.staffeln)
  };
  k.versionen.push(v);
  k.anfrage = null;
  protokoll(k, { was: a ? 'Freigegeben' : 'Selbst freigegeben', version: v.nr, kommentar: [angepasst ? 'Preise angepasst' : '', kommentar].filter(Boolean).join(' · ') });
  speichern(k);
  if (a) melden(a.von, `${ich()} hat die Kalkulation „${deal.name}“ (${deal.pid || '–'}) freigegeben – Version ${v.nr}${angepasst ? ', Preise angepasst' : ''}.`, deal.id, 'freigegeben');
  dealStatus(deal, 'Preise freigegeben', '');
  return v;
}

/** Zurückgeben (nur Freigeber) – Kommentar ist Pflicht */
export function zurueckgeben(deal, kommentar) {
  const k = basis(deal);
  const a = k.anfrage;
  if (!a) fehler('Es gibt keine offene Freigabe-Anfrage.');
  if (a.an !== ich()) fehler(`Nur ${a.an} kann zurückgeben.`);
  if (!kommentar?.trim()) fehler('Bitte kurz begründen, was geändert werden soll.');
  k.anfrage = null;
  protokoll(k, { was: 'Zurückgegeben', an: a.von, kommentar });
  speichern(k);
  melden(a.von, `${ich()} hat die Kalkulation „${deal.name}“ (${deal.pid || '–'}) zurückgegeben: „${kommentar}“`, deal.id, 'zurueck');
  dealStatus(deal, 'Angebot schreiben', '');
}

/** Anfrage zurückziehen (nur Anforderer) */
export function zurueckziehen(deal) {
  const k = basis(deal);
  const a = k.anfrage;
  if (!a) fehler('Es gibt keine offene Freigabe-Anfrage.');
  if (a.von !== ich()) fehler(`Nur ${a.von} kann die Anfrage zurückziehen.`);
  k.anfrage = null;
  protokoll(k, { was: 'Anfrage zurückgezogen', an: a.an });
  speichern(k);
  melden(a.an, `${ich()} hat die Freigabe-Anfrage „${deal.name}“ (${deal.pid || '–'}) zurückgezogen.`, deal.id, 'info');
  dealStatus(deal, 'Angebot schreiben', '');
}

/** Welche Knöpfe hat der aktuelle Anwender bei diesem Deal? */
export function moeglich(dealId) {
  const k = kalkulation(dealId); const a = k?.anfrage; const me = ich();
  return {
    anfordern: !a, selbstFreigeben: !a,
    freigeben: !!a && a.an === me, zurueckgeben: !!a && a.an === me,
    zurueckziehen: !!a && a.von === me
  };
}

export function reset() { set(LS_FREIGABE, {}); set(LS_NACHRICHTEN, []); try { localStorage.removeItem('elwis-demo-kalkulationen-v1'); } catch { /* egal */ } }
