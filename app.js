import { CONFIG, JAHRESORDNER } from './config.js';
import * as drive from './drive.js';

const $ = (sel) => document.querySelector(sel);
const zustand = {
  jahr: Number(localStorageLesen('steuerjahr')) || new Date().getFullYear(),
  hauptordner: null,
  jahresordner: null,
  unterordner: {},
  excel: null,
};

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
  meldung(`Steuerjahr ${zustand.jahr} ist bereit.`, 'ok');
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
    meldung(e instanceof drive.DriveFehler ? e.message : 'Unerwarteter Fehler. Bitte erneut versuchen.', 'fehler');
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

  for (const knopf of document.querySelectorAll('[data-bald]')) {
    knopf.addEventListener('click', () => meldung(`„${knopf.textContent.trim()}“ kommt in einem der nächsten Schritte.`));
  }
}

init();
