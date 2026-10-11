// Buchungsregeln. Jede Buchung schreibt alle zusammengehörigen Zeilen in EINEM Speichervorgang.
import { datumZuSerial } from './xlsx.js';

const zeilenwert = (w) => {
  if (w == null) return null;
  if (typeof w !== 'object') return w;
  return w.datum ?? w.zahl ?? w.text ?? (w.formel != null ? '=' + w.formel : null);
};

async function zeilenHash(werte, spalten) {
  const daten = new TextEncoder().encode(JSON.stringify(spalten.map((s) => zeilenwert(werte[s]))));
  const h = new Uint8Array(await crypto.subtle.digest('SHA-256', daten));
  return [...h.slice(0, 8)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

const datumSchluessel = (werte) => (typeof werte.Datum === 'number' ? werte.Datum : null);
const jetzt = () => new Date().toLocaleString('sv-SE').slice(0, 16); // YYYY-MM-DD HH:MM, lokale Zeit

async function systemZeile(mappe, uuidZeile, buchungsId, blatt, werte, spalten, datei = {}) {
  await mappe.zeileAnhaengen('_System', [
    { text: uuidZeile }, { text: buchungsId }, { text: blatt },
    { text: datei.id || '' }, { text: datei.hash || '' },
    { text: await zeilenHash(werte, spalten) }, { text: jetzt() }, { text: jetzt() }, { text: 'Gebucht' },
  ]);
}

/** Prüft, ob für das Datum schon Kilometer erfasst sind. */
export async function kmDubletten(mappe, datum) {
  const tab = await mappe.tabelle('Km Fahrrad');
  const s = datumZuSerial(datum);
  return tab.daten.filter((d) => d.werte.Datum === s);
}

/** Km Fahrrad + verknüpfte Auskommen-Zeile. */
export async function kmBuchen(mappe, { datum, km, satz }) {
  const bid = crypto.randomUUID();
  const km_ = await mappe.tabelle('Km Fahrrad');
  const aus = await mappe.tabelle('Auskommen');
  const T = km_.name, A = aus.name;
  const s = datumZuSerial(datum);

  const kmWerte = {
    Datum: { datum },
    'Gefahrene Kilometer': { zahl: km },
    'Satz (€/km)': { zahl: satz },
    'Betrag (€)': { formel: `${T}[[#This Row],[Gefahrene Kilometer]]*${T}[[#This Row],[Satz (€/km)]]` },
    ID: { text: bid },
  };
  const ausWerte = {
    Datum: { datum },
    Beschreibung: { text: 'Fahrrad-Kilometergeld' },
    Kategorie: { text: 'Fahrrad / Kilometergeld' },
    Zahlungsmethode: null,
    'Betrag (€)': { formel: `INDEX(${T}[Betrag (€)],MATCH(${A}[[#This Row],[ID]],${T}[ID],0))` },
    Notizen: null,
    ID: { text: bid },
  };
  const z1 = await mappe.zeileEinfuegen('Km Fahrrad', kmWerte, datumSchluessel, s);
  const z2 = await mappe.zeileEinfuegen('Auskommen', ausWerte, datumSchluessel, s);
  await systemZeile(mappe, crypto.randomUUID(), bid, 'Km Fahrrad', kmWerte, km_.spalten);
  await systemZeile(mappe, crypto.randomUUID(), bid, 'Auskommen', ausWerte, aus.spalten);
  return { buchungsId: bid, zeilen: { 'Km Fahrrad': z1, Auskommen: z2 }, betrag: Math.round(km * satz * 100) / 100 };
}
