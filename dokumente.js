// Upload → Analyse → Prüfung → Freigabe → Archivierung → Excel
import * as drive from './drive.js';
import { analysieren, dateiHash, pdfZeilen, pdfInfo, pdfSeitenAlsBilder, HOCH, MITTEL, NIEDRIG } from './analyse.js';
import { dokumentBuchen, dokumentDubletten } from './buchung.js';

const $ = (s) => document.querySelector(s);

const FELDER = {
  Einkommen: [
    ['Zahlungsdatum', 'datum'], ['Zahlungsart', 'auswahl', ['Überweisung', 'Bar']], ['Zahlung (brutto)', 'betrag'],
    ['Rechnungsnummer', 'text'], ['Rechnungsempfänger', 'text'], ['Mitarbeiter', 'text'], ['Rechnungsdatum', 'datum'],
    ['Rechnungsstatus', 'auswahl', ['Bezahlt', 'Offen']],
  ],
  Auskommen: [
    ['Datum', 'datum'], ['Beschreibung', 'text'], ['Kategorie', 'kategorie'],
    ['Zahlungsmethode', 'auswahl', ['Karte', 'Überweisung', 'Bankeinzug', 'Bar']], ['Betrag (€)', 'betrag'], ['Notizen', 'text'],
  ],
  Tmobile: [
    ['Datum', 'datum'], ['Beschreibung', 'text'], ['Rechnungsnummer', 'text'],
    ['Zahlungsmethode', 'auswahl', ['Bankeinzug', 'Karte', 'Überweisung']], ['Rechnungsbetrag (€)', 'betrag'],
    ['davon privat (€)', 'betrag'], ['davon WLAN (€)', 'betrag'], ['Mobilfunk absetzbar (%)', 'prozent'], ['WLAN absetzbar (%)', 'prozent'],
  ],
  Bewirtungkosten: [
    ['Datum', 'datum'], ['Beschreibung', 'text'], ['Kategorie', 'text'],
    ['Zahlungsmethode', 'auswahl', ['Karte', 'Bar', 'Überweisung']], ['Betrag (€)', 'betrag'], ['Absetzbar (%)', 'prozent'],
  ],
};
const ZIELORDNER = { Einkommen: 'Einkommen', Auskommen: 'Auskommen', Tmobile: 'Auskommen/Tmobile', Bewirtungkosten: 'Auskommen/Bewirtungkosten' };
const PFLICHT = {
  Einkommen: ['Zahlung (brutto)', 'Rechnungsnummer', 'Rechnungsempfänger', 'Rechnungsdatum', 'Rechnungsstatus', 'Zahlungsart'],
  Auskommen: ['Datum', 'Beschreibung', 'Kategorie', 'Betrag (€)'],
  Tmobile: ['Datum', 'Rechnungsbetrag (€)', 'davon privat (€)', 'davon WLAN (€)', 'Mobilfunk absetzbar (%)', 'WLAN absetzbar (%)'],
  Bewirtungkosten: ['Datum', 'Beschreibung', 'Betrag (€)', 'Absetzbar (%)'],
};

let ctx;            // { zustand, meldung, sicher, sicherSpeichern, eur, datumDE }
let aktuell = null; // { datei, blob, vorschlag, typ, werte, sicherheit, bestaetigt }

export function einrichtenDokumente(kontext) {
  ctx = kontext;
  $('#dok-oeffnen').addEventListener('click', () => $('#dok-datei').click());
  $('#dok-datei').addEventListener('change', (e) => ctx.sicher(() => hochladen([...e.target.files]).finally(() => { e.target.value = ''; })));
  $('#pruef-oeffnen').addEventListener('click', () => ctx.sicher(zurPruefungZeigen));
  $('#pruef-typ').addEventListener('change', (e) => typWechseln(e.target.value));
  $('#pruef-buchen').addEventListener('click', () => ctx.sicher(freigeben));
  $('#pruef-zurueck').addEventListener('click', () => { schliessen(); ctx.meldung('Zurückgestellt – das Dokument bleibt unter „Zur Prüfung“.', 'ok'); zaehlerAktualisieren(); });
  $('#pruef-ablehnen').addEventListener('click', () => ctx.sicher(ablehnen));
  $('#pruef-ablegen').addEventListener('click', () => ctx.sicher(nurAblegen));
}

// ---------- Upload ----------
async function hochladen(dateien) {
  if (!dateien.length) return;
  const eingang = ctx.zustand.unterordner['_Eingang'];
  let erste = null;
  for (const f of dateien) {
    ctx.meldung(`„${f.name}“ wird hochgeladen …`);
    const hash = await dateiHash(f);
    const vorhanden = await drive.sucheHash(hash);
    if (vorhanden.length) {
      ctx.meldung(`Möglicherweise bereits erfasst: „${f.name}“ wurde schon einmal hochgeladen (${vorhanden[0].name}). Nicht erneut hochgeladen.`, 'fehler');
      continue;
    }
    const istPdf = f.type === 'application/pdf' || /\.pdf$/i.test(f.name);
    const info = istPdf ? await pdfInfo(f).catch(() => null) : null;
    if (info && info.seiten > 1 && info.gescannt) {
      // Eingescanntes PDF mit mehreren Seiten: jede Seite = ein Beleg. Das Original bleibt unverändert in „Scans“.
      ctx.meldung(`„${f.name}“: ${info.seiten} Seiten erkannt – jede Seite wird als eigener Beleg vorbereitet …`);
      await drive.dateiHochladen(f.name, ctx.zustand.unterordner['Scans'].id, f, 'application/pdf', { hash, status: 'scan' });
      const bilder = await pdfSeitenAlsBilder(f);
      const basis = f.name.replace(/\.pdf$/i, '');
      for (let i = 0; i < bilder.length; i++) {
        const seitenHash = await dateiHash(bilder[i]);
        const neu = await drive.dateiHochladen(`${basis} – Seite ${i + 1} von ${bilder.length}.jpg`, eingang.id, bilder[i], 'image/jpeg',
          { hash: seitenHash, quelle: hash.slice(0, 32), seite: String(i + 1), status: 'eingang' });
        erste ??= { datei: neu, blob: bilder[i] };
      }
      ctx.meldung(`${bilder.length} Belege aus „${f.name}“ liegen unter „Zur Prüfung“. Das Original-PDF ist in ${ctx.zustand.jahr}/Scans/ gespeichert.`, 'ok');
      continue;
    }
    // Original sofort sichern – es geht nie verloren, auch wenn später etwas schiefgeht.
    const neu = await drive.dateiHochladen(f.name, eingang.id, f, f.type || 'application/octet-stream', { hash, status: 'eingang' });
    erste ??= { datei: neu, blob: f };
  }
  await zaehlerAktualisieren();
  if (erste) await pruefungOeffnen(erste.datei, erste.blob);
}

export async function zaehlerAktualisieren() {
  const eingang = ctx.zustand.unterordner['_Eingang'];
  if (!eingang) return;
  const n = (await drive.ordnerInhalt(eingang.id)).length;
  $('#pruef-oeffnen').textContent = n ? `Zur Prüfung (${n})` : 'Zur Prüfung';
}

async function zurPruefungZeigen() {
  const liste = await drive.ordnerInhalt(ctx.zustand.unterordner['_Eingang'].id);
  const ul = $('#pruef-liste-eintraege');
  ul.innerHTML = '';
  $('#pruef-liste').hidden = false;
  if (!liste.length) { ul.innerHTML = '<li class="hinweis">Keine offenen Dokumente.</li>'; return; }
  for (const d of liste) {
    const li = document.createElement('li');
    const b = document.createElement('button');
    b.textContent = d.name;
    b.addEventListener('click', () => ctx.sicher(async () => {
      ctx.meldung('Dokument wird geladen …');
      await pruefungOeffnen(d, await drive.dateiLesen(d.id));
    }));
    const zeit = document.createElement('span');
    zeit.className = 'hinweis';
    zeit.textContent = ` hochgeladen ${new Date(d.createdTime).toLocaleDateString('de-AT')}`;
    li.append(b, zeit);
    ul.append(li);
  }
  $('#pruef-liste').scrollIntoView({ behavior: 'smooth' });
}

// ---------- Prüfung ----------
async function pruefungOeffnen(datei, blob) {
  $('#pruef-liste').hidden = true;
  const typ = blob.type || datei.mimeType;
  const b = typ && typ !== 'application/octet-stream' ? blob : new Blob([blob], { type: datei.mimeType });
  const vorschlag = await analysieren(b, datei.name, ctx.zustand.einstellungen, (t) => ctx.meldung(t));
  if (vorschlag.fehler && !vorschlag.typ) { ctx.meldung(vorschlag.fehler, 'fehler'); }
  aktuell = { datei, blob: b, vorschlag, typ: vorschlag.typ || 'Auskommen', werte: {}, sicherheit: {}, geaendert: new Set(), bestaetigt: false };
  for (const [k, f] of Object.entries(vorschlag.felder || {})) { aktuell.werte[k] = f.wert; aktuell.sicherheit[k] = f; }
  await originalZeigen(b, datei.name);
  $('#pruef-typ').value = aktuell.typ;
  $('#pruef-typ-hinweis').textContent = vorschlag.typSicherheit === HOCH ? '' : 'Dokumenttyp nicht sicher erkannt – bitte prüfen.';
  felderZeigen();
  $('#pruef-ablegen').hidden = aktuell.typ !== 'Einkommen';
  $('#pruefung').hidden = false;
  $('#pruefung').scrollIntoView({ behavior: 'smooth', block: 'start' });
  ctx.meldung('Bitte die Daten prüfen und dann „Freigeben und buchen“.');
}

async function originalZeigen(blob, name) {
  const box = $('#pruef-original');
  box.innerHTML = '';
  if (box.dataset.url) URL.revokeObjectURL(box.dataset.url);
  const url = URL.createObjectURL(blob);
  box.dataset.url = url;
  const link = document.createElement('a');
  link.href = url; link.target = '_blank'; link.rel = 'noopener'; link.textContent = 'Original in neuem Fenster öffnen';
  if (blob.type === 'application/pdf' || /\.pdf$/i.test(name)) {
    try {
      const lib = await import('./pdf.min.js');
      lib.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.min.js', import.meta.url).href;
      const doc = await lib.getDocument({ data: new Uint8Array(await blob.arrayBuffer()), verbosity: 0 }).promise;
      const seite = await doc.getPage(1);
      const vp = seite.getViewport({ scale: 1.5 });
      const c = document.createElement('canvas');
      c.width = vp.width; c.height = vp.height;
      await seite.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      box.append(c);
    } catch { /* nur Link */ }
  } else {
    const img = document.createElement('img');
    img.src = url; img.alt = 'Originaldokument';
    box.append(img);
  }
  box.append(link);
}

function typWechseln(neu) {
  // Gemeinsame Felder übernehmen
  const w = aktuell.werte;
  const betrag = w['Betrag (€)'] ?? w['Rechnungsbetrag (€)'] ?? w['Zahlung (brutto)'];
  const datum = w.Datum ?? w.Rechnungsdatum;
  if (neu === 'Einkommen') { w.Rechnungsdatum ??= datum; w['Zahlung (brutto)'] ??= betrag; }
  else { w.Datum ??= datum; if (neu === 'Tmobile') w['Rechnungsbetrag (€)'] ??= betrag; else w['Betrag (€)'] ??= betrag; }
  const e = ctx.zustand.einstellungen;
  if (neu === 'Bewirtungkosten') { w['Absetzbar (%)'] ??= e.Bewirtungsanteil; w.Kategorie = w.Kategorie && w.Kategorie !== 'Bewirtungkosten' ? 'Bewirtungkosten' : 'Bewirtungkosten'; }
  if (neu === 'Tmobile') { w['Mobilfunk absetzbar (%)'] ??= e.Mobilfunkanteil; w['WLAN absetzbar (%)'] ??= e.Telekommunikationsanteil; w.Beschreibung ??= 'T-Mobile'; }
  if (neu === 'Auskommen' && w.Kategorie === 'Bewirtungkosten') w.Kategorie = null;
  if (neu === 'Einkommen') { w.Mitarbeiter ??= e.EigenerName; w.Zahlungsart ??= e.StandardZahlungsartEinkommen; w.Rechnungsstatus ??= 'Bezahlt'; }
  aktuell.typ = neu;
  aktuell.bestaetigt = false;
  felderZeigen();
  $('#pruef-ablegen').hidden = neu !== 'Einkommen';
}

function felderZeigen() {
  const box = $('#pruef-felder');
  box.innerHTML = '';
  for (const [name, art, optionen] of FELDER[aktuell.typ]) {
    const label = document.createElement('label');
    const kopf = document.createElement('span');
    kopf.className = 'feldkopf';
    kopf.textContent = name;
    const s = aktuell.sicherheit[name];
    const wert = aktuell.werte[name];
    const badge = document.createElement('span');
    const pflicht = PFLICHT[aktuell.typ].includes(name);
    if (!aktuell.geaendert.has(name)) {
      if ((wert == null || wert === '') && pflicht) { badge.className = 'badge niedrig'; badge.textContent = 'Bitte prüfen'; }
      else if (s?.sicherheit === NIEDRIG && wert != null && wert !== '') { badge.className = 'badge niedrig'; badge.textContent = 'Nicht eindeutig erkannt'; }
      else if (s?.sicherheit === MITTEL) { badge.className = 'badge mittel'; badge.textContent = 'Mittel'; }
    }
    if (badge.textContent) kopf.append(badge);
    let input;
    if (art === 'auswahl') {
      input = document.createElement('select');
      const leer = document.createElement('option'); leer.value = ''; leer.textContent = '– bitte wählen –'; input.append(leer);
      for (const o of new Set([...(optionen || []), ...(wert && !optionen.includes(wert) ? [wert] : [])])) {
        const op = document.createElement('option'); op.value = o; op.textContent = o; input.append(op);
      }
      input.value = wert ?? '';
    } else {
      input = document.createElement('input');
      if (art === 'datum') { input.type = 'date'; input.value = wert ?? ''; }
      else if (art === 'betrag') { input.type = 'text'; input.inputMode = 'decimal'; input.value = wert != null ? String(wert).replace('.', ',') : ''; }
      else if (art === 'prozent') { input.type = 'text'; input.inputMode = 'decimal'; input.value = wert != null ? String(Math.round(wert * 1000) / 10).replace('.', ',') : ''; }
      else { input.type = 'text'; input.value = wert ?? ''; }
      if (art === 'kategorie') input.setAttribute('list', 'kategorien');
      if (art === 'prozent') { const w = document.createElement('span'); w.className = 'einheit'; w.textContent = '%'; label.append(w); }
    }
    input.addEventListener('input', () => {
      aktuell.werte[name] = leseEingabe(input.value, art);
      aktuell.geaendert.add(name);
      aktuell.bestaetigt = false;
      badge.remove();
      vorschauZeigen();
    });
    if (s?.hinweis && !aktuell.geaendert.has(name)) { const h = document.createElement('small'); h.textContent = s.hinweis; label.append(kopf, input, h); }
    else label.append(kopf, input);
    box.append(label);
  }
  const hinweise = $('#pruef-hinweise');
  hinweise.innerHTML = '';
  for (const t of aktuell.vorschlag.hinweise || []) { const li = document.createElement('li'); li.textContent = t; hinweise.append(li); }
  if (aktuell.vorschlag.fehler) { const li = document.createElement('li'); li.textContent = aktuell.vorschlag.fehler; hinweise.append(li); }
  vorschauZeigen();
}

function leseEingabe(v, art) {
  if (v === '') return null;
  if (art === 'betrag') { const n = Number(v.replace(/\s/g, '').replace(/\.(?=\d{3})/g, '').replace(',', '.')); return Number.isFinite(n) ? n : null; }
  if (art === 'prozent') { const n = Number(v.replace(',', '.')); return Number.isFinite(n) ? n / 100 : null; }
  return v;
}

function berechnen() {
  const w = aktuell.werte;
  const r = (x) => Math.round(x * 100) / 100;
  if (aktuell.typ === 'Tmobile') {
    const wl = w['davon WLAN (€)'] ?? 0, p = w['davon privat (€)'] ?? 0;
    return r(((w['Rechnungsbetrag (€)'] ?? 0) - p - wl) * (w['Mobilfunk absetzbar (%)'] ?? 0) + wl * (w['WLAN absetzbar (%)'] ?? 0));
  }
  if (aktuell.typ === 'Bewirtungkosten') return r((w['Betrag (€)'] ?? 0) * (w['Absetzbar (%)'] ?? 0));
  return null;
}

function pruefen() {
  const w = aktuell.werte, typ = aktuell.typ, jahr = String(ctx.zustand.jahr);
  const fehler = [];
  for (const p of PFLICHT[typ]) if (w[p] == null || w[p] === '') fehler.push(`„${p}“ fehlt.`);
  if (typ === 'Einkommen' && w.Rechnungsstatus === 'Bezahlt' && !w.Zahlungsdatum) fehler.push('Bei „Bezahlt“ bitte das Zahlungsdatum eintragen (oder Status „Offen“ wählen).');
  const d = typ === 'Einkommen' ? (w.Zahlungsdatum || w.Rechnungsdatum) : w.Datum;
  if (d && !d.startsWith(jahr)) fehler.push(`Das Datum ${ctx.datumDE(d)} liegt nicht im Steuerjahr ${jahr}.`);
  for (const k of Object.keys(w)) if (/\(%\)$/.test(k) && w[k] != null && (w[k] < 0 || w[k] > 1)) fehler.push(`„${k}“ muss zwischen 0 % und 100 % liegen.`);
  return fehler;
}

function vorschauZeigen() {
  const v = $('#pruef-vorschau');
  v.innerHTML = '';
  const w = aktuell.werte, typ = aktuell.typ;
  const zeile = (titel, text) => {
    const d = document.createElement('div');
    const l = document.createElement('span'); l.className = 'label'; l.textContent = titel;
    const t = document.createElement('span'); t.textContent = text;
    d.append(l, t); v.append(d);
  };
  const dt = (x) => (x ? ctx.datumDE(x) : '–');
  const e = (x) => (x != null ? ctx.eur(x) : '–');
  if (typ === 'Einkommen') zeile('Einkommen', `${dt(w.Zahlungsdatum)} · ${w.Rechnungsempfänger || '–'} · Nr. ${w.Rechnungsnummer || '–'} · ${e(w['Zahlung (brutto)'])} · ${w.Rechnungsstatus || '–'}`);
  if (typ === 'Auskommen') zeile('Auskommen', `${dt(w.Datum)} · ${w.Beschreibung || '–'} · ${w.Kategorie || '–'} · ${e(w['Betrag (€)'])}`);
  if (typ === 'Tmobile' || typ === 'Bewirtungkosten') {
    const abs = berechnen();
    zeile(typ, typ === 'Tmobile'
      ? `${dt(w.Datum)} · Rechnung ${e(w['Rechnungsbetrag (€)'])} − privat ${e(w['davon privat (€)'])} · WLAN ${e(w['davon WLAN (€)'])} → absetzbar ${e(abs)}`
      : `${dt(w.Datum)} · ${w.Beschreibung || '–'} · ${e(w['Betrag (€)'])} × ${Math.round((w['Absetzbar (%)'] ?? 0) * 100)} % = ${e(abs)}`);
    zeile('Auskommen (automatisch verknüpft)', `${dt(w.Datum)} · ${typ === 'Tmobile' ? (w.Beschreibung || 'T-Mobile') : (w.Beschreibung || '–')} · ${typ === 'Tmobile' ? 'Telekommunikation' : 'Bewirtungkosten'} · ${e(abs)}`);
  }
  zeile('Ablage in Google Drive', `${ctx.zustand.jahr}/${ZIELORDNER[typ]}/`);
  const f = pruefen();
  $('#pruef-warnung').hidden = !f.length && !aktuell.bestaetigt;
  $('#pruef-warnung').textContent = f.join(' ');
  $('#pruef-buchen').textContent = 'Freigeben und buchen';
}

function zielName() {
  const w = aktuell.werte, typ = aktuell.typ;
  const datum = (typ === 'Einkommen' ? w.Rechnungsdatum : w.Datum) || 'ohne-Datum';
  const wer = (typ === 'Einkommen' ? `Rechnung_${w.Rechnungsnummer}` : (w.Beschreibung || typ)).replace(/[\\/:*?"<>|]+/g, '-').slice(0, 60);
  const betrag = w['Betrag (€)'] ?? w['Rechnungsbetrag (€)'] ?? w['Zahlung (brutto)'];
  const endung = (/\.[a-z0-9]{2,5}$/i.exec(aktuell.datei.name) || [''])[0];
  return `${datum}_${wer}${betrag != null ? '_' + betrag.toFixed(2).replace('.', ',') : ''}${endung}`;
}

async function freigeben() {
  const fehler = pruefen();
  if (fehler.length) { $('#pruef-warnung').hidden = false; $('#pruef-warnung').textContent = fehler.join(' '); return ctx.meldung('Bitte zuerst die markierten Angaben ergänzen.', 'fehler'); }
  const knopf = $('#pruef-buchen');
  knopf.disabled = true;
  try {
    ctx.meldung('Buchung wird gespeichert …');
    const hash = aktuell.datei.appProperties?.hash || await dateiHash(aktuell.blob);
    const w = { ...aktuell.werte };
    if (aktuell.typ === 'Einkommen' && w.Rechnungsnummer && /^\d{1,15}$/.test(String(w.Rechnungsnummer))) w.Rechnungsnummer = Number(w.Rechnungsnummer);
    await ctx.sicherSpeichern(async (mappe) => {
      const dub = await dokumentDubletten(mappe, aktuell.typ, w, hash);
      if (dub.length && !aktuell.bestaetigt) { const e = new Error('Dublette'); e.dubletten = dub; throw e; }
      return dokumentBuchen(mappe, aktuell.typ, w, { id: aktuell.datei.id, hash });
    });
    // Erst nach erfolgreicher Buchung archivieren (verschieben, keine Kopie)
    const ziel = ctx.zustand.unterordner[ZIELORDNER[aktuell.typ]];
    try {
      await drive.dateiVerschieben(aktuell.datei.id, ziel.id, ctx.zustand.unterordner['_Eingang'].id, zielName());
      ctx.meldung(`Gebucht und abgelegt in ${ctx.zustand.jahr}/${ZIELORDNER[aktuell.typ]}/.`, 'ok');
    } catch {
      ctx.meldung('Gebucht. Das Dokument konnte aber nicht verschoben werden – es liegt noch in _Eingang. Bitte später erneut „Zur Prüfung“ öffnen und ablehnen oder manuell verschieben.', 'fehler');
    }
    schliessen();
    await weiter();
  } catch (e) {
    if (!e.dubletten) throw e;
    aktuell.bestaetigt = true;
    $('#pruef-warnung').hidden = false;
    $('#pruef-warnung').textContent = `Möglicherweise bereits erfasst: ${e.dubletten.join(' ')} Nur buchen, wenn es wirklich ein weiterer Beleg ist.`;
    knopf.textContent = 'Trotzdem buchen';
    ctx.meldung('');
  } finally {
    knopf.disabled = false;
  }
}

/** Eigene Rechnung, deren Zahlung schon (z. B. über den Appointmed-Import) gebucht ist: nur archivieren. */
async function nurAblegen() {
  if (!aktuell) return;
  if (!confirm('Nur ablegen? Das Dokument wird NICHT gebucht, sondern nur im Ordner Einkommen archiviert. Sinnvoll, wenn die Zahlung schon über den Appointmed-Import erfasst ist.')) return;
  await drive.dateiVerschieben(aktuell.datei.id, ctx.zustand.unterordner['Einkommen'].id, ctx.zustand.unterordner['_Eingang'].id, zielName());
  schliessen();
  ctx.meldung(`Abgelegt in ${ctx.zustand.jahr}/Einkommen/ (ohne Buchung).`, 'ok');
  await weiter();
}

async function ablehnen() {
  if (!aktuell) return;
  if (!confirm('Dokument ablehnen? Es wird nicht gebucht und in den Ordner _Abgelehnt verschoben (nicht gelöscht).')) return;
  await drive.dateiVerschieben(aktuell.datei.id, ctx.zustand.unterordner['_Abgelehnt'].id, ctx.zustand.unterordner['_Eingang'].id);
  schliessen();
  ctx.meldung('Abgelehnt und nach _Abgelehnt verschoben.', 'ok');
  await weiter();
}

/** Nach einer Entscheidung: offene Belege gleich anzeigen. */
async function weiter() {
  await zaehlerAktualisieren();
  const rest = await drive.ordnerInhalt(ctx.zustand.unterordner['_Eingang'].id);
  if (rest.length) await zurPruefungZeigen();
}

function schliessen() {
  $('#pruefung').hidden = true;
  const box = $('#pruef-original');
  if (box.dataset.url) { URL.revokeObjectURL(box.dataset.url); delete box.dataset.url; }
  box.innerHTML = '';
  aktuell = null;
}

/** Kategorien für die Auswahlliste (aus dem versteckten Blatt _Parameter). */
export async function kategorienLaden(mappe) {
  try {
    const zeilen = await mappe.blattZeilen('_Parameter');
    const liste = zeilen.map((z) => z[4]).filter((x) => typeof x === 'string' && x.trim());
    const dl = $('#kategorien');
    dl.innerHTML = '';
    for (const k of liste) { const o = document.createElement('option'); o.value = k; dl.append(o); }
  } catch { /* Liste ist nur eine Hilfe */ }
}

export { pdfZeilen };
