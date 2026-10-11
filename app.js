import { CONFIG, JAHRESORDNER } from './config.js';
import * as drive from './drive.js';
import { Arbeitsmappe, ExcelFehler } from './xlsx.js';
import { kmBuchen, kmDubletten } from './buchung.js';
import { jahresuebersicht } from './auswertung.js';

const $ = (sel) => document.querySelector(sel);
const zustand = {
  jahr: Number(localStorageLesen('steuerjahr')) || new Date().getFullYear(),
  hauptordner: null,
  jahresordner: null,
  unterordner: {},
  excel: null,
  einstellungen: { ...CONFIG.standardEinstellungen },
};
const eur = (n) => n.toLocaleString('de-AT', { style: 'currency', currency: 'EUR' });
const datumDE = (iso) => iso.split('-').reverse().join('.');
const heute = () => new Date().toLocaleDateString('sv-SE');

function localStorageLesen(k) { try { return localStorage.getItem(k); } catch { return null; } }
function localStorageSchreiben(k, v) { try { localStorage.setItem(k, v); } catch { /* egal */ } }

function meldung(text, art = 'info') {
  const el = $('#meldung');
  el.textContent = text;
  el.className = `meldung ${art}`;
  el.hidden = !text;
}

function protokoll(text) {
  const li = document.createElement('li');
  li.textContent = text;
  $('#einrichtung-liste').append(li);
}

function zeigeAnsicht(angemeldet) {
  $('#ansicht-start').hidden = angemeldet;
  $('#ansicht-dashboard').hidden = !angemeldet;
  $('#abmelden').hidden = !angemeldet;
}

function jahrAuswahlFuellen() {
  const sel = $('#steuerjahr');
  sel.innerHTML = '';
  if (!CONFIG.steuerjahre.includes(zustand.jahr)) zustand.jahr = CONFIG.steuerjahre[0];
  for (const j of CONFIG.steuerjahre) {
    const o = document.createElement('option');
    o.value = j; o.textContent = j; o.selected = j === zustand.jahr;
    sel.append(o);
  }
}

/** Legt Ordnerstruktur und Excel des Steuerjahres an, falls noch nicht vorhanden. Löscht nie etwas. */
async function einrichten() {
  $('#einrichtung-liste').innerHTML = '';
  $('#einrichtung').hidden = false;
  meldung('Google Drive wird geprüft …');

  const haupt = await drive.ordnerFindenOderAnlegen(CONFIG.hauptordner);
  zustand.hauptordner = haupt;
  protokoll(`${haupt.neu ? 'Angelegt' : 'Gefunden'}: ${CONFIG.hauptordner}/`);

  const einst = await drive.suche('Einstellungen.json', haupt.id);
  if (!einst) {
    const blob = new Blob([JSON.stringify(CONFIG.standardEinstellungen, null, 2)], { type: 'application/json' });
    await drive.dateiHochladen('Einstellungen.json', haupt.id, blob, 'application/json');
    protokoll('Angelegt: Einstellungen.json');
  } else {
    protokoll('Gefunden: Einstellungen.json');
    try {
      zustand.einstellungen = { ...CONFIG.standardEinstellungen, ...JSON.parse(await (await drive.dateiLesen(einst.id)).text()) };
    } catch { protokoll('Einstellungen.json unlesbar – Standardwerte werden verwendet.'); }
  }

  const jahr = await drive.ordnerFindenOderAnlegen(String(zustand.jahr), haupt.id);
  zustand.jahresordner = jahr;
  protokoll(`${jahr.neu ? 'Angelegt' : 'Gefunden'}: ${zustand.jahr}/`);

  zustand.unterordner = {};
  for (const pfad of JAHRESORDNER) {
    const eltern = pfad.length === 1 ? jahr.id : zustand.unterordner[pfad.slice(0, -1).join('/')].id;
    const o = await drive.ordnerFindenOderAnlegen(pfad.at(-1), eltern);
    zustand.unterordner[pfad.join('/')] = o;
    if (o.neu) protokoll(`Angelegt: ${zustand.jahr}/${pfad.join('/')}/`);
  }

  const name = `Buchhaltung_${zustand.jahr}.xlsx`;
  let excel = await drive.suche(name, jahr.id);
  if (!excel) {
    const vorlage = await fetch(`vorlage/${name}`);
    if (!vorlage.ok) throw new drive.DriveFehler(`Für ${zustand.jahr} gibt es noch keine Excel-Vorlage.`);
    excel = await drive.dateiHochladen(name, jahr.id, await vorlage.blob(), drive.XLSX_MIME);
    protokoll(`Angelegt: ${zustand.jahr}/${name}`);
  } else {
    protokoll(`Gefunden: ${zustand.jahr}/${name}`);
  }
  zustand.excel = await drive.metadaten(excel.id);
  zeigeExcel();
  await uebersichtAktualisieren();
  meldung(`Steuerjahr ${zustand.jahr} ist bereit.`, 'ok');
}

async function uebersichtAktualisieren(mappe) {
  try {
    mappe ??= await Arbeitsmappe.laden(await drive.dateiLesen(zustand.excel.id));
    const u = await jahresuebersicht(mappe);
    const [a, b, c] = document.querySelectorAll('.zahlen .wert');
    a.textContent = eur(u.einnahmen); b.textContent = eur(u.ausgaben); c.textContent = eur(u.differenz);
    $('#uebersicht-hinweis').textContent = u.unsicher
      ? `${u.unsicher} Zeile(n) in Auskommen konnten nicht berechnet werden – bitte in Excel prüfen.`
      : 'Berechnet aus der aktuellen Excel-Datei.';
  } catch {
    $('#uebersicht-hinweis').textContent = 'Übersicht konnte nicht berechnet werden.';
  }
}

/**
 * Liest die neueste Excel-Version, wendet die Änderung an und speichert nur,
 * wenn die Datei in der Zwischenzeit nicht verändert wurde. Sonst neu versuchen.
 */
async function sicherSpeichern(aenderung) {
  const id = zustand.excel.id;
  for (let versuch = 1; versuch <= 3; versuch++) {
    const vorher = await drive.metadaten(id);
    const mappe = await Arbeitsmappe.laden(await drive.dateiLesen(id));
    const ergebnis = await aenderung(mappe);
    const neu = await mappe.speichern();
    await tagesBackup(id);
    const jetzt = await drive.metadaten(id);
    if (jetzt.version !== vorher.version) continue; // inzwischen geändert → neueste Version neu lesen
    zustand.excel = await drive.dateiAktualisieren(id, neu, drive.XLSX_MIME);
    zeigeExcel();
    await uebersichtAktualisieren(await Arbeitsmappe.laden(neu));
    return ergebnis;
  }
  throw new drive.DriveFehler('Die Excel-Datei wurde gerade an anderer Stelle geändert. Nichts wurde gebucht – bitte erneut versuchen.');
}

/** Vor der ersten Buchung des Tages eine Kopie in _Backup ablegen. */
async function tagesBackup(id) {
  const ordner = zustand.unterordner['_Backup'];
  const name = `Buchhaltung_${zustand.jahr}_${heute()}.xlsx`;
  if (!(await drive.suche(name, ordner.id))) await drive.dateiKopieren(id, name, ordner.id);
}

// ---------- Kilometer eintragen ----------
let kmBestaetigt = false;

function kmWerte() {
  return {
    datum: $('#km-datum').value,
    km: Number(String($('#km-km').value).replace(',', '.')),
    satz: Number(String($('#km-satz').value).replace(',', '.')),
  };
}

function kmPruefen() {
  const w = kmWerte();
  const fehler = [];
  if (!w.datum) fehler.push('Bitte ein Datum wählen.');
  else if (!w.datum.startsWith(String(zustand.jahr))) fehler.push(`Das Datum liegt nicht im Steuerjahr ${zustand.jahr}.`);
  if (!(w.km > 0 && w.km <= 500)) fehler.push('Bitte gültige Kilometer eingeben (mehr als 0, höchstens 500).');
  if (!(w.satz > 0 && w.satz <= 5)) fehler.push('Bitte einen gültigen Satz pro km eingeben.');
  return { w, fehler };
}

function kmVorschau() {
  kmBestaetigt = false;
  $('#km-buchen').textContent = 'Freigeben und buchen';
  $('#km-warnung').hidden = true;
  const { w, fehler } = kmPruefen();
  const v = $('#km-vorschau');
  if (fehler.length) { v.innerHTML = ''; const p = document.createElement('p'); p.className = 'warnung'; p.textContent = fehler[0]; v.append(p); return; }
  const betrag = Math.round(w.km * w.satz * 100) / 100;
  v.innerHTML = `
    <div><span class="label">Km Fahrrad</span>
      <span>${datumDE(w.datum)} · ${w.km.toLocaleString('de-AT')} km × ${eur(w.satz)} = <strong>${eur(betrag)}</strong></span></div>
    <div><span class="label">Auskommen (automatisch verknüpft)</span>
      <span>${datumDE(w.datum)} · Fahrrad-Kilometergeld · Fahrrad / Kilometergeld · <strong>${eur(betrag)}</strong></span></div>`;
}

function kmOeffnen() {
  $('#km-formular').hidden = false;
  $('#km-datum').value = heute().startsWith(String(zustand.jahr)) ? heute() : `${zustand.jahr}-01-01`;
  $('#km-km').value = zustand.einstellungen.Arbeitsweg_km;
  $('#km-satz').value = zustand.einstellungen.FahrradKilometersatz;
  $('#km-arbeitstag').textContent = `Arbeitstag (${String(zustand.einstellungen.Arbeitsweg_km).replace('.', ',')} km)`;
  kmVorschau();
  $('#km-formular').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

async function kmFreigeben() {
  const { w, fehler } = kmPruefen();
  if (fehler.length) return meldung(fehler[0], 'fehler');
  $('#km-buchen').disabled = true;
  try {
    meldung('Buchung wird gespeichert …');
    const ergebnis = await sicherSpeichern(async (mappe) => {
      const dub = await kmDubletten(mappe, w.datum);
      if (dub.length && !kmBestaetigt) {
        const err = new Error('Dublette'); err.dubletten = dub; throw err;
      }
      return kmBuchen(mappe, w);
    });
    meldung(`Gebucht: ${datumDE(w.datum)} · ${w.km.toLocaleString('de-AT')} km · ${eur(ergebnis.betrag)} (Km Fahrrad + Auskommen).`, 'ok');
    $('#km-formular').hidden = true;
  } catch (e) {
    if (e.dubletten) {
      kmBestaetigt = true;
      const km = e.dubletten.map((d) => `${String(d.werte['Gefahrene Kilometer']).replace('.', ',')} km`).join(', ');
      $('#km-warnung').textContent = `Möglicherweise bereits erfasst: Am ${datumDE(w.datum)} gibt es schon ${e.dubletten.length} Eintrag/Einträge (${km}). Nur buchen, wenn das ein weiterer Weg war.`;
      $('#km-warnung').hidden = false;
      $('#km-buchen').textContent = 'Trotzdem buchen';
      meldung('');
    } else throw e;
  } finally {
    $('#km-buchen').disabled = false;
  }
}

function zeigeExcel() {
  const e = zustand.excel;
  $('#excel-name').textContent = e.name;
  $('#excel-geaendert').textContent = new Date(e.modifiedTime).toLocaleString('de-AT');
  $('#excel-oeffnen').href = e.webViewLink;
  $('#drive-oeffnen').href = `https://drive.google.com/drive/folders/${zustand.jahresordner.id}`;
}

async function sicher(fn) {
  try {
    await fn();
  } catch (e) {
    console.error(e?.name || 'Fehler'); // keine Inhalte protokollieren
    meldung(e instanceof drive.DriveFehler || e instanceof ExcelFehler ? e.message : 'Unerwarteter Fehler. Bitte erneut versuchen.', 'fehler');
  }
}

function init() {
  jahrAuswahlFuellen();
  zeigeAnsicht(false);
  drive.vorbereiten().catch((e) => meldung(e.message, 'fehler'));

  $('#anmelden').addEventListener('click', () => sicher(async () => {
    meldung('Anmeldung bei Google …');
    await drive.anmelden(true);
    zeigeAnsicht(true);
    await einrichten();
  }));

  $('#abmelden').addEventListener('click', () => {
    drive.abmelden();
    zeigeAnsicht(false);
    meldung('Abgemeldet.');
  });

  $('#steuerjahr').addEventListener('change', (ev) => sicher(async () => {
    zustand.jahr = Number(ev.target.value);
    localStorageSchreiben('steuerjahr', zustand.jahr);
    if (drive.angemeldet()) await einrichten();
  }));

  $('#km-oeffnen').addEventListener('click', kmOeffnen);
  for (const id of ['#km-datum', '#km-km', '#km-satz']) $(id).addEventListener('input', kmVorschau);
  $('#km-arbeitstag').addEventListener('click', () => { $('#km-km').value = zustand.einstellungen.Arbeitsweg_km; kmVorschau(); });
  $('#km-abbrechen').addEventListener('click', () => { $('#km-formular').hidden = true; meldung(''); });
  $('#km-buchen').addEventListener('click', () => sicher(kmFreigeben));

  for (const knopf of document.querySelectorAll('[data-bald]')) {
    knopf.addEventListener('click', () => meldung(`„${knopf.textContent.trim()}“ kommt in einem der nächsten Schritte.`));
  }
}

init();
