// Import der Zahlungsliste aus Appointmed (CSV) → Blatt Einkommen.
import * as drive from './drive.js';
import { dateiHash, betragLesen, datumLesen } from './analyse.js';
import { dokumentBuchen } from './buchung.js';
import { datumZuSerial, serialZuDatum, Arbeitsmappe } from './xlsx.js';

const $ = (s) => document.querySelector(s);
const SPALTEN = ['Zahlungsdatum', 'Zahlungsart', 'Zahlung (brutto)', 'Rechnungsnummer', 'Rechnungsempfänger', 'Mitarbeiter', 'Rechnungsdatum', 'Rechnungsstatus'];
let ctx, stand = null; // { datei, zeilen: [{w, status, grund, gewaehlt}] }

export function einrichtenImport(kontext) {
  ctx = kontext;
  $('#imp-oeffnen').addEventListener('click', () => $('#imp-datei').click());
  $('#imp-datei').addEventListener('change', (e) => ctx.sicher(() => csvLaden(e.target.files[0]).finally(() => { e.target.value = ''; })));
  $('#imp-buchen').addEventListener('click', () => ctx.sicher(buchen));
  $('#imp-abbrechen').addEventListener('click', () => { $('#import').hidden = true; stand = null; ctx.meldung(''); });
}

function dekodieren(buf) {
  let b = new Uint8Array(buf);
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) b = b.subarray(3);
  try { return new TextDecoder('utf-8', { fatal: true }).decode(b); } catch { return new TextDecoder('windows-1252').decode(b); }
}

function csvTeilen(zeile, trenner) {
  const aus = []; let feld = '', inQ = false;
  for (let i = 0; i < zeile.length; i++) {
    const ch = zeile[i];
    if (inQ) { if (ch === '"' && zeile[i + 1] === '"') { feld += '"'; i++; } else if (ch === '"') inQ = false; else feld += ch; }
    else if (ch === '"') inQ = true;
    else if (ch === trenner) { aus.push(feld); feld = ''; }
    else feld += ch;
  }
  aus.push(feld);
  return aus.map((s) => s.trim());
}

export function csvLesen(text) {
  const zeilen = text.split(/\r?\n/).filter((z) => z.trim());
  if (!zeilen.length) return { fehler: 'Die Datei ist leer.' };
  const trenner = (zeilen[0].match(/;/g) || []).length >= (zeilen[0].match(/,/g) || []).length ? ';' : ',';
  return tabelleLesen(zeilen.map((z) => csvTeilen(z, trenner)));
}

/** Gemeinsamer Teil für CSV und Excel: erste Zeile = Spaltennamen. */
function tabelleLesen(roh) {
  const kopf = roh[0].map((k) => String(k ?? '').trim());
  const fehlt = SPALTEN.filter((s) => !kopf.includes(s));
  if (fehlt.length) return { fehler: `Das ist keine Appointmed-Zahlungsliste – es fehlen Spalten: ${fehlt.join(', ')}.` };
  const idx = Object.fromEntries(kopf.map((k, i) => [k, i]));
  const datum = (v) => (typeof v === 'number' ? serialZuDatum(v) : datumLesen(String(v ?? '')));
  const betrag = (v) => (typeof v === 'number' ? Math.round(v * 100) / 100 : betragLesen(v));
  const daten = roh.slice(1).filter((t) => t.some((v) => v != null && String(v).trim() !== '')).map((t) => {
    const roh_ = (k) => { const v = idx[k] != null ? t[idx[k]] : null; return v && typeof v === 'object' ? v.wert : v; };
    const g = (k) => String(roh_(k) ?? '').trim();
    return {
      Zahlungsdatum: datum(roh_('Zahlungsdatum')),
      Zahlungsart: g('Zahlungsart') || null,
      'Zahlung (brutto)': betrag(roh_('Zahlung (brutto)')),
      Rechnungsnummer: /^\d{1,15}$/.test(g('Rechnungsnummer')) ? Number(g('Rechnungsnummer')) : g('Rechnungsnummer') || null,
      Rechnungsempfänger: g('Rechnungsempfänger') || null,
      Mitarbeiter: g('Mitarbeiter') || null,
      Rechnungsdatum: datum(roh_('Rechnungsdatum')),
      Rechnungsstatus: g('Rechnungsstatus') || null,
      _summe: betrag(roh_('Rechnungssumme (brutto)')),
    };
  });
  return { daten };
}

/** Nur bezahlte Rechnungen mit Zahlungsdatum und Betrag im Steuerjahr werden übernommen – alles andere ist gesperrt. */
export function zeileBewerten(w, jahr) {
  const st = (w.Rechnungsstatus || '').trim();
  if (/storn/i.test(st)) return { status: 'aus', grund: 'Storniert' };
  if (!/^bezahlt$/i.test(st)) return { status: 'aus', grund: `Nicht bezahlt (${st || 'ohne Status'})` };
  if (!w.Zahlungsdatum) return { status: 'aus', grund: 'Ohne Zahlungsdatum' };
  if (!(w['Zahlung (brutto)'] > 0)) return { status: 'aus', grund: 'Ohne Zahlungsbetrag' };
  if (!w.Zahlungsdatum.startsWith(jahr)) return { status: 'aus', grund: `Zahlung im Jahr ${w.Zahlungsdatum.slice(0, 4)}` };
  return { status: 'neu', grund: '' };
}
export { tabelleLesen as _tabelleLesen, zeileBewerten as _status };

async function csvLaden(datei) {
  if (!datei) return;
  ctx.meldung('Zahlungsliste wird gelesen …');
  let ergebnis;
  if (/\.xlsx$/i.test(datei.name)) {
    const m = await Arbeitsmappe.laden(datei);
    const blatt = Object.keys(m.blaetter)[0];
    ergebnis = tabelleLesen(await m.blattZeilen(blatt, true));
  } else {
    ergebnis = csvLesen(dekodieren(await datei.arrayBuffer()));
  }
  const { daten, fehler } = ergebnis;
  if (fehler) return ctx.meldung(fehler, 'fehler');
  const mappe = await ctx.mappeLaden();
  const ein = await mappe.tabelle('Einkommen');
  const vorhanden = ein.daten.filter((d) => !d.leer).map((d) => d.werte);
  const jahr = String(ctx.zustand.jahr);
  const schl = (w) => `${w.Rechnungsnummer}|${typeof w.Zahlungsdatum === 'number' ? w.Zahlungsdatum : w.Zahlungsdatum ? datumZuSerial(w.Zahlungsdatum) : ''}|${w['Zahlung (brutto)']}`;
  const bekannt = new Set(vorhanden.map(schl));
  const nummern = new Set(vorhanden.map((w) => String(w.Rechnungsnummer)));
  const zeilen = daten.map((w) => {
    let { status, grund } = zeileBewerten(w, jahr);
    if (status === 'neu') w.Rechnungsstatus = 'Bezahlt';
    if (status !== 'aus' && bekannt.has(schl(w))) { status = 'aus'; grund = 'Bereits gebucht'; }
    else if (status !== 'aus' && nummern.has(String(w.Rechnungsnummer))) { status = 'pruefen'; grund = 'Rechnungsnummer schon vorhanden (weitere Teilzahlung?)'; }
    if (status === 'neu' && w._summe != null && w['Zahlung (brutto)'] != null && Math.abs(w._summe - w['Zahlung (brutto)']) > 0.005) {
      status = 'pruefen'; grund = `Teilzahlung: ${ctx.eur(w['Zahlung (brutto)'])} von ${ctx.eur(w._summe)}`;
    }
    if (w.Rechnungsdatum && w.Zahlungsdatum && w.Rechnungsdatum.slice(0, 4) !== w.Zahlungsdatum.slice(0, 4) && status === 'neu') grund = `Rechnung ${w.Rechnungsdatum.slice(0, 4)}, bezahlt ${w.Zahlungsdatum.slice(0, 4)} → zählt für ${w.Zahlungsdatum.slice(0, 4)}`;
    return { w, status, grund, gewaehlt: status === 'neu' };
  }).sort((a, b) => (a.w.Zahlungsdatum || '9').localeCompare(b.w.Zahlungsdatum || '9'));
  stand = { datei, zeilen };
  tabelleZeigen();
  $('#import').hidden = false;
  $('#import').scrollIntoView({ behavior: 'smooth', block: 'start' });
  ctx.meldung('Bitte die Liste prüfen. Nur angehakte Zeilen werden gebucht.');
}

function tabelleZeigen() {
  const tb = $('#imp-zeilen');
  tb.innerHTML = '';
  for (const z of stand.zeilen) {
    const tr = document.createElement('tr');
    if (z.status === 'aus') tr.className = 'aus';
    if (z.status === 'pruefen') tr.className = 'pruefen';
    const td0 = document.createElement('td');
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.checked = z.gewaehlt; cb.disabled = z.status === 'aus';
    cb.addEventListener('change', () => { z.gewaehlt = cb.checked; zusammenfassung(); });
    td0.append(cb); tr.append(td0);
    const w = z.w;
    for (const t of [w.Zahlungsdatum ? ctx.datumDE(w.Zahlungsdatum) : '–', w.Rechnungsempfänger, String(w.Rechnungsnummer ?? ''), w.Zahlungsart,
      w['Zahlung (brutto)'] != null ? ctx.eur(w['Zahlung (brutto)']) : '–', z.grund]) {
      const td = document.createElement('td'); td.textContent = t ?? ''; tr.append(td);
    }
    tb.append(tr);
  }
  zusammenfassung();
}

function zusammenfassung() {
  const z = stand.zeilen;
  const gew = z.filter((x) => x.gewaehlt);
  const summe = gew.reduce((s, x) => s + (x.w['Zahlung (brutto)'] || 0), 0);
  const zaehl = (f) => z.filter(f).length;
  $('#imp-info').textContent = `${z.length} Zeilen: ${zaehl((x) => x.status === 'neu')} neu · ${zaehl((x) => x.status === 'pruefen')} zu prüfen · ` +
    `${zaehl((x) => x.grund === 'Bereits gebucht')} bereits gebucht · ${zaehl((x) => x.grund === 'Storniert')} storniert · ` +
    `${zaehl((x) => /^Nicht bezahlt|^Ohne/.test(x.grund))} nicht bezahlt · ${zaehl((x) => /^Zahlung im Jahr/.test(x.grund))} anderes Jahr.`;
  $('#imp-buchen').textContent = gew.length ? `${gew.length} Zahlungen (${ctx.eur(summe)}) freigeben und buchen` : 'Nichts ausgewählt';
  $('#imp-buchen').disabled = !gew.length;
}

async function buchen() {
  const gew = stand.zeilen.filter((x) => x.gewaehlt);
  if (!gew.length) return;
  $('#imp-buchen').disabled = true;
  try {
    ctx.meldung(`${gew.length} Zahlungen werden gebucht …`);
    // CSV im Ordner Einkommen ablegen (einmalig, keine Kopie bei erneutem Import derselben Datei)
    const hash = await dateiHash(stand.datei);
    let csv = (await drive.sucheHash(hash))[0];
    if (!csv) {
      csv = await drive.dateiHochladen(`Appointmed_Zahlungen_${new Date().toLocaleDateString('sv-SE')}.csv`,
        ctx.zustand.unterordner['Einkommen'].id, stand.datei, 'text/csv', { hash, status: 'import' });
    }
    let gebucht = 0;
    await ctx.sicherSpeichern(async (mappe) => {
      gebucht = 0;
      const ein = await mappe.tabelle('Einkommen');
      const schon = new Set(ein.daten.filter((d) => !d.leer).map((d) => `${d.werte.Rechnungsnummer}|${d.werte.Zahlungsdatum}|${d.werte['Zahlung (brutto)']}`));
      for (const z of gew) {
        const w = { ...z.w }; delete w._summe;
        if (w.Rechnungsstatus !== 'Bezahlt' || !w.Zahlungsdatum || !(w['Zahlung (brutto)'] > 0)) continue; // nur Bezahltes
        const k = `${w.Rechnungsnummer}|${w.Zahlungsdatum ? datumZuSerial(w.Zahlungsdatum) : null}|${w['Zahlung (brutto)']}`;
        if (schon.has(k)) continue; // in der Zwischenzeit gebucht
        await dokumentBuchen(mappe, 'Einkommen', w, { id: csv.id, hash: '' });
        gebucht++;
      }
    });
    $('#import').hidden = true;
    stand = null;
    ctx.meldung(`${gebucht} Zahlungen ins Blatt Einkommen gebucht. Die Zahlungsliste liegt in ${ctx.zustand.jahr}/Einkommen/.`, 'ok');
  } finally {
    $('#imp-buchen').disabled = false;
  }
}
