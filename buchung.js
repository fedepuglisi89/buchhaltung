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
    'Betrag (€)': { formel: `ROUND(${T}[[#This Row],[Gefahrene Kilometer]]*${T}[[#This Row],[Satz (€/km)]],2)` },
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

// ---------- Dokumente ----------
export const TELEKOM_KATEGORIE = 'Telekommunikationsaufwand (mobile + wifi Wohnung) 50%';

const alsZelle = (spalte, wert) => {
  if (wert == null || wert === '') return null;
  if (/datum$|^Datum$/i.test(spalte)) return { datum: wert };
  if (typeof wert === 'number') return { zahl: wert };
  return { text: String(wert) };
};

/** Prüft Dubletten: gleiche Datei, gleiche Rechnungsnummer oder gleiches Datum + Betrag. */
export async function dokumentDubletten(mappe, typ, w, hash) {
  const treffer = [];
  const sys = await mappe.blattZeilen('_System');
  if (hash && sys.some((z) => z[4] === hash && z[8] !== 'Storniert')) treffer.push('Diese Datei wurde bereits gebucht.');
  const tab = await mappe.tabelle(typ);
  for (const d of tab.daten) {
    if (d.leer) continue;
    const nr = w.Rechnungsnummer;
    if (nr && d.werte.Rechnungsnummer != null && String(d.werte.Rechnungsnummer) === String(nr)) {
      treffer.push(`Rechnungsnummer ${nr} ist im Blatt „${typ}“ schon vorhanden.`);
    }
    const datumSpalte = typ === 'Einkommen' ? 'Rechnungsdatum' : 'Datum';
    const betragSpalte = { Einkommen: 'Zahlung (brutto)', Tmobile: 'Rechnungsbetrag (€)' }[typ] || 'Betrag (€)';
    if (w[datumSpalte] && typeof d.werte[datumSpalte] === 'number' && d.werte[datumSpalte] === datumZuSerial(w[datumSpalte])
        && typeof d.werte[betragSpalte] === 'number' && Math.abs(d.werte[betragSpalte] - (w[betragSpalte] ?? NaN)) < 0.005) {
      treffer.push(`Am selben Tag gibt es im Blatt „${typ}“ schon einen Eintrag mit demselben Betrag.`);
    }
  }
  return [...new Set(treffer)];
}

/** Bucht ein geprüftes Dokument. w = freigegebene Werte, datei = {id, hash}. */
export async function dokumentBuchen(mappe, typ, w, datei) {
  const bid = crypto.randomUUID();
  const tab = await mappe.tabelle(typ);
  const T = tab.name;
  const werte = {};
  for (const s of tab.spalten) werte[s] = alsZelle(s, w[s]);
  werte.ID = { text: bid };

  let link = null;
  if (typ === 'Tmobile') {
    const r = (c) => `${T}[[#This Row],[${c}]]`;
    werte['Absetzbarer Betrag (€)'] = { formel: `ROUND((${r('Rechnungsbetrag (€)')}-${r('davon privat (€)')}-${r('davon WLAN (€)')})*${r('Mobilfunk absetzbar (%)')}+${r('davon WLAN (€)')}*${r('WLAN absetzbar (%)')},2)` };
    link = { Beschreibung: w.Beschreibung || 'T-Mobile', Kategorie: TELEKOM_KATEGORIE, spalte: 'Absetzbarer Betrag (€)' };
  } else if (typ === 'Bewirtungkosten') {
    werte['Absetzbarer Betrag (€)'] = { formel: `ROUND(${T}[[#This Row],[Betrag (€)]]*${T}[[#This Row],[Absetzbar (%)]],2)` };
    link = { Beschreibung: w.Beschreibung, Kategorie: 'Bewirtungkosten', spalte: 'Absetzbarer Betrag (€)' };
  }

  const schluessel = typ === 'Einkommen'
    ? (x) => (typeof x.Zahlungsdatum === 'number' ? x.Zahlungsdatum : typeof x.Rechnungsdatum === 'number' ? x.Rechnungsdatum : null)
    : datumSchluessel;
  const neuDatum = typ === 'Einkommen' ? (w.Zahlungsdatum || w.Rechnungsdatum) : w.Datum;
  const s = neuDatum ? datumZuSerial(neuDatum) : Infinity;
  await mappe.zeileEinfuegen(typ, werte, schluessel, s);
  await systemZeile(mappe, crypto.randomUUID(), bid, typ, werte, tab.spalten, datei);

  if (link) {
    const aus = await mappe.tabelle('Auskommen');
    const ausWerte = {
      Datum: { datum: w.Datum },
      Beschreibung: alsZelle('Beschreibung', link.Beschreibung),
      Kategorie: { text: link.Kategorie },
      Zahlungsmethode: alsZelle('Zahlungsmethode', w.Zahlungsmethode),
      'Betrag (€)': { formel: `INDEX(${T}[${link.spalte}],MATCH(${aus.name}[[#This Row],[ID]],${T}[ID],0))` },
      Notizen: null,
      ID: { text: bid },
    };
    await mappe.zeileEinfuegen('Auskommen', ausWerte, datumSchluessel, s);
    await systemZeile(mappe, crypto.randomUUID(), bid, 'Auskommen', ausWerte, aus.spalten, datei);
  }
  return { buchungsId: bid };
}
