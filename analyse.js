// Dokumentanalyse – alles läuft lokal im Browser. Es werden keine Dokumente an fremde Dienste geschickt.
// PDF: Text über pdf.js. Fotos: österreichischer Beleg-QR-Code (RKSV), Texterkennung nur als Ergänzung.

export const HOCH = 'Hoch', MITTEL = 'Mittel', NIEDRIG = 'Niedrig';
const feld = (wert, sicherheit, hinweis) => ({ wert: wert ?? null, sicherheit: wert == null || wert === '' ? NIEDRIG : sicherheit, hinweis });

export const betragLesen = (s) => {
  if (s == null) return null;
  const t = String(s).replace(/[€\s]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.');
  const n = Number(t);
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : null;
};
export const datumLesen = (s) => {
  let m = /(\d{1,2})\.(\d{1,2})\.(\d{2,4})/.exec(s || '');
  if (m) { const j = m[3].length === 2 ? '20' + m[3] : m[3]; return `${j}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
  m = /(\d{4})[-/](\d{2})[-/](\d{2})/.exec(s || '');
  return m ? `${m[1]}-${m[2]}-${m[3]}` : null;
};

// ---------- SHA-256 der Datei (für die Dublettenprüfung) ----------
export async function dateiHash(blob) {
  const h = new Uint8Array(await crypto.subtle.digest('SHA-256', await blob.arrayBuffer()));
  return [...h].map((b) => b.toString(16).padStart(2, '0')).join('');
}

// ---------- PDF ----------
let pdfjs = null;
async function pdfjsLaden() {
  if (!pdfjs) {
    pdfjs = await import('./pdf.min.js');
    pdfjs.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.min.js', import.meta.url).href;
  }
  return pdfjs;
}

/** Liefert die Textzeilen eines PDFs (Elemente einer Zeile durch " | " getrennt). */
export async function pdfZeilen(blob, maxSeiten = 4) {
  const lib = await pdfjsLaden();
  const doc = await lib.getDocument({ data: new Uint8Array(await blob.arrayBuffer()), verbosity: 0 }).promise;
  const zeilen = [];
  for (let p = 1; p <= Math.min(doc.numPages, maxSeiten); p++) {
    const tc = await (await doc.getPage(p)).getTextContent();
    const map = new Map();
    for (const it of tc.items) {
      if (!it.str || !it.str.trim()) continue;
      const y = Math.round(it.transform[5] / 2) * 2;
      if (!map.has(y)) map.set(y, []);
      map.get(y).push([it.transform[4], it.str.trim()]);
    }
    for (const [, teile] of [...map].sort((a, b) => b[0] - a[0])) {
      zeilen.push(teile.sort((a, b) => a[0] - b[0]).map((t) => t[1]).join(' | '));
    }
  }
  return zeilen;
}

const wertNach = (zeilen, label) => {
  for (const z of zeilen) {
    const teile = z.split(' | ');
    const i = teile.findIndex((t) => t.replace(/:$/, '').trim().toLowerCase() === label.toLowerCase());
    if (i >= 0 && teile[i + 1]) return teile[i + 1];
  }
  return null;
};
const nurZiffern = (s) => String(s || '').replace(/\D/g, '');

// ---------- Magenta / T-Mobile ----------
function magenta(zeilen, einst) {
  const text = zeilen.join('\n');
  if (!/magenta|T-Mobile Austria/i.test(text)) return null;
  const eigene = nurZiffern(einst.EigeneRufnummer);
  const datum = datumLesen(wertNach(zeilen, 'Rechnungsdatum'));
  const nummer = wertNach(zeilen, 'Rechnungsnummer');
  const gesamtZeile = zeilen.find((z) => /^Rechnungsbetrag \|/.test(z));
  const gesamt = betragLesen(gesamtZeile?.split(' | ').at(-1));
  let privat = 0, wlan = 0, eigen = 0, gefunden = 0;
  for (const z of zeilen) {
    const t = z.split(' | ');
    if (t.length < 5) continue;
    const brutto = betragLesen(t.at(-1));
    if (brutto == null) continue;
    if (/^0\d{3}[\d ]{6,}$/.test(t[0])) {
      gefunden++;
      if (eigene && nurZiffern(t[0]) === eigene) eigen += brutto; else privat += brutto;
    } else if (/^Internet$/i.test(t[0])) { gefunden++; wlan += brutto; }
  }
  const summeOk = gefunden && gesamt != null && Math.abs(eigen + privat + wlan - gesamt) < 0.05;
  const aufteilung = gefunden ? (summeOk && eigene ? HOCH : MITTEL) : NIEDRIG;
  return {
    typ: 'Tmobile',
    typSicherheit: HOCH,
    felder: {
      Datum: feld(datum, HOCH),
      Beschreibung: feld('T-Mobile', HOCH),
      Rechnungsnummer: feld(nummer, HOCH),
      Zahlungsmethode: feld(/SEPA|eingezogen/i.test(text) ? 'Bankeinzug' : null, MITTEL),
      'Rechnungsbetrag (€)': feld(gesamt, HOCH),
      'davon privat (€)': feld(gefunden ? Math.round(privat * 100) / 100 : null, aufteilung,
        eigene ? 'Alle Rufnummern außer deiner eigenen' : 'Eigene Rufnummer in den Einstellungen fehlt'),
      'davon WLAN (€)': feld(gefunden ? Math.round(wlan * 100) / 100 : null, aufteilung),
      'Mobilfunk absetzbar (%)': feld(einst.Mobilfunkanteil, HOCH),
      'WLAN absetzbar (%)': feld(einst.Telekommunikationsanteil, HOCH),
    },
    hinweise: summeOk || !gefunden ? [] : ['Die Summe der Verträge stimmt nicht mit dem Rechnungsbetrag überein – bitte prüfen.'],
  };
}

// ---------- Appointmed (eigene Rechnung → Einkommen) ----------
function appointmed(zeilen, einst) {
  const text = zeilen.join('\n');
  const name = einst.EigenerName || 'Federico Puglisi';
  if (!(new RegExp(`Abs\\.:\\s*${name.split(' ')[0]}`, 'i').test(text) || /FEDEPHYSIO/i.test(text)) || !/Rechnungsnummer/.test(text)) return null;
  const nummer = nurZiffern(wertNach(zeilen, 'Rechnungsnummer'));
  const rdatum = datumLesen(wertNach(zeilen, 'Rechnungsdatum'));
  const gesamt = betragLesen(wertNach(zeilen, 'Gesamtbetrag'));
  // Empfänger: erste Zeile nach der Absenderzeile
  const i = zeilen.findIndex((z) => /^Abs\.:/.test(z));
  let empf = null;
  for (const z of i >= 0 ? zeilen.slice(i + 1, i + 12) : []) {
    const t = z.split(' | ')[0].trim();
    if (/^Rechnung$/i.test(t)) break;
    if (/:$/.test(t) || /\d/.test(t) || !t) continue;
    empf = t; break;
  }
  const bezahlt = /BEZAHLT|dankend erhalten/i.test(text);
  return {
    typ: 'Einkommen',
    typSicherheit: HOCH,
    felder: {
      Zahlungsdatum: feld(null, NIEDRIG, 'Zahlungsdatum steht nicht auf der Rechnung'),
      Zahlungsart: feld(einst.StandardZahlungsartEinkommen || 'Überweisung', MITTEL, 'Vorschlag'),
      'Zahlung (brutto)': feld(gesamt, HOCH),
      Rechnungsnummer: feld(nummer ? Number(nummer) : null, HOCH),
      Rechnungsempfänger: feld(empf, empf ? MITTEL : NIEDRIG),
      Mitarbeiter: feld(name, HOCH),
      Rechnungsdatum: feld(rdatum, HOCH),
      Rechnungsstatus: feld(bezahlt ? 'Bezahlt' : 'Offen', bezahlt ? HOCH : MITTEL),
    },
    hinweise: [],
  };
}

// ---------- Allgemeine Rechnung (PDF) ----------
function allgemeinPdf(zeilen, einst) {
  const text = zeilen.join('\n');
  const datum = datumLesen(wertNach(zeilen, 'Rechnungsdatum') || wertNach(zeilen, 'Datum') || text);
  let betrag = null;
  for (const label of ['Gesamtbetrag', 'Rechnungsbetrag', 'Gesamtsumme', 'Summe', 'Total', 'Endbetrag', 'Zu zahlen']) {
    const z = zeilen.find((x) => new RegExp(`^${label}`, 'i').test(x));
    if (z) { betrag = betragLesen(z.split(' | ').at(-1)); if (betrag != null) break; }
  }
  const lieferant = lieferantErkennen(text, einst);
  return auskommenVorschlag({ datum, betrag, lieferant, text, sicherheitBetrag: betrag != null ? MITTEL : NIEDRIG, einst });
}

// ---------- Lieferanten und Kategorien ----------
export function lieferantErkennen(text, einst) {
  for (const r of einst.Lieferantenregeln || []) {
    if (new RegExp(r.muster, 'i').test(text)) return r;
  }
  return null;
}

function auskommenVorschlag({ datum, betrag, lieferant, text, sicherheitBetrag, einst, zahlart, ust }) {
  const essen = (ust && ust[10] > 0) || /Restaurant|Gastronomie|Kellner|Tisch \d|Trattoria|Pizzeria|Gasthaus|Café|Bar\b/i.test(text);
  const istBewirtung = lieferant?.blatt === 'Bewirtungkosten' || (!lieferant && essen);
  const hinweise = [];
  if (istBewirtung) hinweise.push('Als Bewirtung erkannt – nur buchen, wenn es ein Geschäftsessen war.');
  if (!lieferant) hinweise.push('Lieferant/Kategorie nicht sicher erkannt – bitte prüfen.');
  if (istBewirtung) {
    return {
      typ: 'Bewirtungkosten', typSicherheit: lieferant ? HOCH : MITTEL,
      felder: {
        Datum: feld(datum, datum ? HOCH : NIEDRIG),
        Beschreibung: feld(lieferant?.name || null, lieferant ? HOCH : NIEDRIG),
        Kategorie: feld('Bewirtungkosten', HOCH),
        Zahlungsmethode: feld(zahlart || null, zahlart ? MITTEL : NIEDRIG),
        'Betrag (€)': feld(betrag, sicherheitBetrag),
        'Absetzbar (%)': feld(einst.Bewirtungsanteil, HOCH),
      },
      hinweise,
    };
  }
  return {
    typ: 'Auskommen', typSicherheit: lieferant ? HOCH : NIEDRIG,
    felder: {
      Datum: feld(datum, datum ? HOCH : NIEDRIG),
      Beschreibung: feld(lieferant?.name || null, lieferant ? HOCH : NIEDRIG),
      Kategorie: feld(lieferant?.kategorie || null, lieferant?.kategorie ? MITTEL : NIEDRIG),
      Zahlungsmethode: feld(zahlart || null, zahlart ? MITTEL : NIEDRIG),
      'Betrag (€)': feld(betrag, sicherheitBetrag),
      Notizen: feld('', HOCH),
    },
    hinweise,
  };
}

// ---------- RKSV-Beleg-QR (Österreich) ----------
/** _R1-AT1_KassenID_Belegnr_2025-06-08T22:09:39_Normal_Ermäßigt1_Ermäßigt2_Null_Besonders_… */
export function rksvLesen(text) {
  const t = String(text || '').split('_');
  if (t.length < 10 || !/^R1-AT\d+$/.test(t[1])) return null;
  const n = (s) => betragLesen(s) ?? 0;
  const ust = { 20: n(t[5]), 10: n(t[6]), 13: n(t[7]), 0: n(t[8]), 19: n(t[9]) };
  const summe = Math.round(Object.values(ust).reduce((a, b) => a + b, 0) * 100) / 100;
  return { kasse: t[2], beleg: t[3], datum: t[4].slice(0, 10), zeit: t[4].slice(11, 16), ust, summe };
}

async function qrSuchen(bild) {
  const versuche = [];
  const { width: w, height: h } = bild;
  for (const [x, y, bw, bh] of [[0, 0, w, h], [0, h / 2, w, h / 2], [0, h / 3, w, (2 * h) / 3], [w / 2, h / 2, w / 2, h / 2], [0, h / 2, w / 2, h / 2]]) {
    for (const ziel of [1200, 800, 1800]) versuche.push({ x, y, bw, bh, ziel });
  }
  if ('BarcodeDetector' in window) {
    try {
      const det = new BarcodeDetector({ formats: ['qr_code'] });
      const r = await det.detect(bild);
      if (r[0]?.rawValue) return r[0].rawValue;
    } catch { /* weiter mit jsQR */ }
  }
  const jsQR = await jsqrLaden();
  if (!jsQR) return null;
  const c = document.createElement('canvas');
  const ctx = c.getContext('2d', { willReadFrequently: true });
  for (const v of versuche) {
    const s = Math.min(1, v.ziel / Math.max(v.bw, v.bh));
    c.width = Math.round(v.bw * s); c.height = Math.round(v.bh * s);
    ctx.drawImage(bild, v.x, v.y, v.bw, v.bh, 0, 0, c.width, c.height);
    const d = ctx.getImageData(0, 0, c.width, c.height);
    const r = jsQR(d.data, c.width, c.height, { inversionAttempts: 'attemptBoth' });
    if (r?.data) return r.data;
  }
  return null;
}

function skriptLaden(src) {
  return new Promise((ok, fehler) => {
    const s = document.createElement('script');
    s.src = src; s.async = true; s.crossOrigin = 'anonymous';
    s.onload = ok; s.onerror = () => fehler(new Error('Laden fehlgeschlagen'));
    document.head.append(s);
  });
}
async function jsqrLaden() {
  if (window.jsQR) return window.jsQR;
  try { await skriptLaden('https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js'); } catch { return null; }
  return window.jsQR || null;
}

let tessWorker = null;
async function texterkennung(bild, bereich) {
  if (!window.Tesseract) await skriptLaden('https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js');
  if (!tessWorker) tessWorker = await window.Tesseract.createWorker('deu');
  const c = document.createElement('canvas');
  const [x, y, w, h] = bereich || [0, 0, bild.width, bild.height];
  const s = Math.min(1, 1600 / Math.max(w, h));
  c.width = Math.round(w * s); c.height = Math.round(h * s);
  c.getContext('2d').drawImage(bild, x, y, w, h, 0, 0, c.width, c.height);
  const r = await tessWorker.recognize(c);
  return r.data.text || '';
}

async function bildAnalysieren(blob, einst, fortschritt) {
  let bild;
  try { bild = await createImageBitmap(blob); } catch {
    return { typ: null, fehler: 'Dieses Bildformat kann der Browser nicht öffnen. Bitte als JPG oder PDF hochladen (iPhone: Fotos werden beim Hochladen meist automatisch umgewandelt).' };
  }
  fortschritt?.('QR-Code auf dem Beleg wird gesucht …');
  const qr = rksvLesen(await qrSuchen(bild).catch(() => null));
  let text = '';
  fortschritt?.('Text wird erkannt (beim ersten Mal kann das etwas dauern) …');
  try {
    text = await texterkennung(bild, qr ? [0, 0, bild.width, Math.round(bild.height * 0.35)] : null);
  } catch { /* Texterkennung optional */ }
  const lieferant = lieferantErkennen(text, einst);
  const kopf = text.split('\n').map((z) => z.trim()).filter((z) => z.length > 2)[0] || null;
  let datum = qr?.datum || datumLesen(text);
  let betrag = qr?.summe ?? null, sicherheitBetrag = qr ? HOCH : NIEDRIG;
  if (!qr) {
    const m = /(?:Summe|Total|Endbetrag|Gesamt|Betrag)[^\d\n]{0,25}(\d{1,4}[.,]\d{2})/i.exec(text);
    betrag = m ? betragLesen(m[1]) : null;
  }
  const v = auskommenVorschlag({
    datum, betrag, lieferant: lieferant || (kopf ? null : null), text, sicherheitBetrag, einst,
    zahlart: /Mastercard|Visa|Karte|Debit|Bankomat|Kontaktlos/i.test(text) ? 'Karte' : /\bbar\b|Bargeld/i.test(text) ? 'Bar' : null,
    ust: qr?.ust,
  });
  if (!lieferant && kopf) { v.felder.Beschreibung = feld(kopf, NIEDRIG, 'Aus der ersten Zeile gelesen'); }
  if (qr) v.hinweise.unshift(`Datum und Betrag aus dem Beleg-QR-Code (Registrierkasse) – sehr zuverlässig.${qr.ust[10] ? ` Davon 10 % USt.-Anteil: ${qr.ust[10].toFixed(2).replace('.', ',')} €.` : ''}`);
  else v.hinweise.unshift('Kein Beleg-QR-Code lesbar – Datum und Betrag stammen aus der Texterkennung. Bitte genau prüfen.');
  v.rksv = qr;
  return v;
}

/** Hauptfunktion: analysiert eine Datei und liefert einen Buchungsvorschlag. */
export async function analysieren(blob, dateiname, einst, fortschritt) {
  const istPdf = blob.type === 'application/pdf' || /\.pdf$/i.test(dateiname);
  if (istPdf) {
    fortschritt?.('PDF wird gelesen …');
    let zeilen = [];
    try { zeilen = await pdfZeilen(blob); } catch { return { typ: null, fehler: 'Das PDF konnte nicht gelesen werden.' }; }
    if (!zeilen.length) return { typ: 'Auskommen', ...leer(einst), hinweise: ['Das PDF enthält keinen Text (eingescannt?). Bitte die Daten von Hand eintragen.'] };
    return magenta(zeilen, einst) || appointmed(zeilen, einst) || allgemeinPdf(zeilen, einst);
  }
  if (/^image\//.test(blob.type) || /\.(jpe?g|png|heic|heif|webp)$/i.test(dateiname)) return bildAnalysieren(blob, einst, fortschritt);
  return { typ: null, fehler: 'Nur PDF, JPG, JPEG oder PNG werden unterstützt.' };
}

function leer(einst) {
  return auskommenVorschlag({ datum: null, betrag: null, lieferant: null, text: '', sicherheitBetrag: NIEDRIG, einst });
}

export { magenta as _magenta, appointmed as _appointmed };
