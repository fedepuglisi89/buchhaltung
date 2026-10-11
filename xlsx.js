// Excel-Motor ohne externe Bibliotheken.
// Grundsatz "Excel gewinnt": Die Datei wird nie neu erzeugt. Nur die geänderten XML-Teile
// werden neu geschrieben; alle anderen Teile bleiben Byte für Byte unverändert.

export class ExcelFehler extends Error {}

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';

// ---------- ZIP ----------
const CRC_TABELLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABELLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
async function stromVerarbeiten(bytes, strom) {
  const s = new Blob([bytes]).stream().pipeThrough(strom);
  return new Uint8Array(await new Response(s).arrayBuffer());
}
const entpacken = (b) => stromVerarbeiten(b, new DecompressionStream('deflate-raw'));
const packen = (b) => stromVerarbeiten(b, new CompressionStream('deflate-raw'));

async function zipLesen(buffer) {
  const b = new Uint8Array(buffer);
  const dv = new DataView(buffer);
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 66000); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new ExcelFehler('Die Datei ist keine gültige Excel-Datei.');
  const anzahl = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const eintraege = [];
  const dec = new TextDecoder();
  for (let i = 0; i < anzahl; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new ExcelFehler('Excel-Datei beschädigt (Verzeichnis).');
    const methode = dv.getUint16(p + 10, true);
    const zeit = dv.getUint16(p + 12, true), datum = dv.getUint16(p + 14, true);
    const crc = dv.getUint32(p + 16, true);
    const gepackt = dv.getUint32(p + 20, true), groesse = dv.getUint32(p + 24, true);
    const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), kl = dv.getUint16(p + 32, true);
    const lokal = dv.getUint32(p + 42, true);
    const name = dec.decode(b.subarray(p + 46, p + 46 + nl));
    const lnl = dv.getUint16(lokal + 26, true), lxl = dv.getUint16(lokal + 28, true);
    const start = lokal + 30 + lnl + lxl;
    eintraege.push({ name, methode, zeit, datum, crc, gepackt, groesse, roh: b.subarray(start, start + gepackt) });
    p += 46 + nl + xl + kl;
  }
  return eintraege;
}

async function zipSchreiben(eintraege) {
  const enc = new TextEncoder();
  const teile = [], zentral = [];
  let offset = 0;
  for (const e of eintraege) {
    const name = enc.encode(e.name);
    const flags = /[^\x00-\x7f]/.test(e.name) ? 0x0800 : 0;
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, flags, true);
    lh.setUint16(8, e.methode, true); lh.setUint16(10, e.zeit, true); lh.setUint16(12, e.datum, true);
    lh.setUint32(14, e.crc, true); lh.setUint32(18, e.gepackt, true); lh.setUint32(22, e.groesse, true);
    lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
    teile.push(new Uint8Array(lh.buffer), name, e.roh);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
    ch.setUint16(8, flags, true); ch.setUint16(10, e.methode, true); ch.setUint16(12, e.zeit, true);
    ch.setUint16(14, e.datum, true); ch.setUint32(16, e.crc, true); ch.setUint32(20, e.gepackt, true);
    ch.setUint32(24, e.groesse, true); ch.setUint16(28, name.length, true); ch.setUint32(42, offset, true);
    zentral.push(new Uint8Array(ch.buffer), name);
    offset += 30 + name.length + e.roh.length;
  }
  const zGroesse = zentral.reduce((s, t) => s + t.length, 0);
  const ende = new DataView(new ArrayBuffer(22));
  ende.setUint32(0, 0x06054b50, true); ende.setUint16(8, eintraege.length, true);
  ende.setUint16(10, eintraege.length, true); ende.setUint32(12, zGroesse, true); ende.setUint32(16, offset, true);
  return new Blob([...teile, ...zentral, new Uint8Array(ende.buffer)],
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// ---------- Hilfen ----------
export function spalteZuIndex(buchst) {
  let n = 0;
  for (const ch of buchst) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
export function indexZuSpalte(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
const zelleTeilen = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? [m[1], Number(m[2])] : null; };
const bereichTeilen = (ref) => { const [a, b = a] = ref.split(':'); return [zelleTeilen(a), zelleTeilen(b)]; };

export const datumZuSerial = (iso) => {
  const [j, m, t] = iso.split('-').map(Number);
  return (Date.UTC(j, m - 1, t) - Date.UTC(1899, 11, 30)) / 864e5;
};
export const serialZuDatum = (n) => new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10);

const kinder = (el, name) => [...el.childNodes].filter((n) => n.nodeType === 1 && n.localName === name);
const kind = (el, name) => kinder(el, name)[0] || null;

/** Verschiebt Bereichsangaben (sqref) ab Zeile ab um 1 nach unten. */
function sqrefVerschieben(sqref, ab) {
  return sqref.split(/\s+/).map((teil) => {
    const [[c1, r1], [c2, r2]] = bereichTeilen(teil);
    const nr1 = r1 >= ab ? r1 + 1 : r1;
    const nr2 = r2 >= ab ? r2 + 1 : r2;
    return nr1 === nr2 && c1 === c2 && !teil.includes(':') ? `${c1}${nr1}` : `${c1}${nr1}:${c2}${nr2}`;
  }).join(' ');
}

// ---------- Arbeitsmappe ----------
export class Arbeitsmappe {
  static async laden(blob) {
    const m = new Arbeitsmappe();
    m.eintraege = await zipLesen(await blob.arrayBuffer());
    m.docs = new Map(); // geänderte/geparste XML-Teile
    await m.#strukturLesen();
    return m;
  }

  #eintrag(name) { return this.eintraege.find((e) => e.name === name); }

  async #text(name) {
    const e = this.#eintrag(name);
    if (!e) return null;
    const daten = e.methode === 0 ? e.roh : await entpacken(e.roh);
    return new TextDecoder().decode(daten);
  }

  async #doc(name) {
    if (!this.docs.has(name)) {
      const t = await this.#text(name);
      if (t == null) throw new ExcelFehler(`Teil fehlt: ${name}`);
      const d = new DOMParser().parseFromString(t, 'application/xml');
      if (d.getElementsByTagName('parsererror').length) throw new ExcelFehler(`XML fehlerhaft: ${name}`);
      this.docs.set(name, { doc: d, geaendert: false });
    }
    return this.docs.get(name);
  }

  async #rels(pfad) {
    const teile = pfad.split('/');
    const datei = teile.pop();
    const relsPfad = [...teile, '_rels', `${datei}.rels`].join('/');
    const t = await this.#text(relsPfad);
    if (!t) return {};
    const d = new DOMParser().parseFromString(t, 'application/xml');
    const basis = teile.join('/');
    const map = {};
    for (const r of d.getElementsByTagNameNS(NS_PKG, 'Relationship')) {
      let ziel = r.getAttribute('Target');
      ziel = ziel.startsWith('/') ? ziel.slice(1) : normalisieren(`${basis}/${ziel}`);
      map[r.getAttribute('Id')] = { ziel, typ: r.getAttribute('Type') };
    }
    return map;
  }

  async #strukturLesen() {
    const wb = (await this.#doc('xl/workbook.xml')).doc;
    const rels = await this.#rels('xl/workbook.xml');
    this.blaetter = {};
    for (const s of wb.getElementsByTagNameNS(NS, 'sheet')) {
      const rid = s.getAttributeNS(NS_R, 'id');
      this.blaetter[s.getAttribute('name')] = { pfad: rels[rid].ziel, tabellen: null };
    }
    const sst = await this.#text('xl/sharedStrings.xml');
    this.sharedStrings = [];
    if (sst) {
      const d = new DOMParser().parseFromString(sst, 'application/xml');
      for (const si of d.getElementsByTagNameNS(NS, 'si')) {
        this.sharedStrings.push([...si.getElementsByTagNameNS(NS, 't')].map((t) => t.textContent).join(''));
      }
    }
  }

  async #tabellenVon(blatt) {
    const info = this.blaetter[blatt];
    if (!info) throw new ExcelFehler(`Blatt „${blatt}“ nicht gefunden.`);
    if (!info.tabellen) {
      const rels = await this.#rels(info.pfad);
      info.tabellen = Object.values(rels).filter((r) => r.typ.endsWith('/table')).map((r) => r.ziel);
    }
    return info.tabellen;
  }

  /** Liest die (erste) Tabelle eines Blattes: Spalten, Datenzeilen, Summenzeile. */
  async tabelle(blatt) {
    const [tPfad] = await this.#tabellenVon(blatt);
    if (!tPfad) throw new ExcelFehler(`Im Blatt „${blatt}“ fehlt die Excel-Tabelle.`);
    const tDoc = (await this.#doc(tPfad)).doc.documentElement;
    const sDoc = (await this.#doc(this.blaetter[blatt].pfad)).doc;
    const [[c1, r1], [c2, r2]] = bereichTeilen(tDoc.getAttribute('ref'));
    const kopf = Number(tDoc.getAttribute('headerRowCount') ?? 1);
    const summe = Number(tDoc.getAttribute('totalsRowCount') ?? 0);
    const spalten = [...tDoc.getElementsByTagNameNS(NS, 'tableColumn')].map((c) => c.getAttribute('name'));
    const zeilen = this.#zeilenMap(sDoc);
    const daten = [];
    for (let r = r1 + kopf; r <= r2 - summe; r++) {
      const werte = {};
      const row = zeilen.get(r);
      spalten.forEach((name, i) => {
        const z = row && this.#zelle(row, `${indexZuSpalte(spalteZuIndex(c1) + i)}${r}`);
        werte[name] = z ? this.#wert(z) : null;
      });
      daten.push({ r, werte, leer: Object.values(werte).every((v) => v == null || v === '') });
    }
    return { blatt, tPfad, name: tDoc.getAttribute('displayName'), spalten, startSpalte: spalteZuIndex(c1),
      kopfZeile: r1, endZeile: r2, summenZeilen: summe, daten };
  }

  #zeilenMap(sDoc) {
    const sd = sDoc.getElementsByTagNameNS(NS, 'sheetData')[0];
    const m = new Map();
    for (const row of kinder(sd, 'row')) m.set(Number(row.getAttribute('r')), row);
    return m;
  }

  #zelle(row, ref) { return kinder(row, 'c').find((c) => c.getAttribute('r') === ref) || null; }

  #wert(c) {
    const t = c.getAttribute('t');
    const f = kind(c, 'f');
    const v = kind(c, 'v');
    if (f) return { formel: f.textContent, wert: v ? v.textContent : null };
    if (t === 'inlineStr') return [...c.getElementsByTagNameNS(NS, 't')].map((x) => x.textContent).join('');
    if (!v) return null;
    if (t === 's') return this.sharedStrings[Number(v.textContent)];
    if (t === 'str' || t === 'e') return v.textContent;
    if (t === 'b') return v.textContent === '1';
    return Number(v.textContent);
  }

  /**
   * Fügt eine Zeile chronologisch in eine Tabelle ein.
   * werte: { Spaltenname: {zahl}|{datum:'YYYY-MM-DD'}|{text}|{formel} }
   * sortiere: Funktion (zeilenWerte) => Sortierschlüssel (Zahl) oder null
   */
  async zeileEinfuegen(blatt, werte, sortiere, neuSchluessel) {
    const tab = await this.tabelle(blatt);
    const eintrag = await this.#doc(this.blaetter[blatt].pfad);
    const sDoc = eintrag.doc;
    const zeilen = this.#zeilenMap(sDoc);
    const ersteDaten = tab.kopfZeile + 1;
    const platzhalter = tab.daten.length === 1 && tab.daten[0].leer;

    let ziel;
    if (platzhalter) {
      ziel = ersteDaten;
    } else {
      // Nach der letzten Zeile mit Schlüssel <= neuem Schlüssel einfügen (stabil)
      ziel = ersteDaten;
      tab.daten.forEach((d) => {
        const k = sortiere(d.werte);
        if (d.leer || k == null || k <= neuSchluessel) ziel = d.r + 1;
      });
      this.#zeilenVerschieben(sDoc, zeilen, ziel, tab);
    }

    // Vorlage für Formate: Zeile darüber, sonst die erste Datenzeile (nach dem Verschieben)
    const zeilenNeu = this.#zeilenMap(sDoc);
    const vorlage = zeilenNeu.get(ziel === ersteDaten ? (platzhalter ? ziel : ziel + 1) : ziel - 1);
    const row = this.#zeileHolen(sDoc, zeilenNeu, ziel);
    tab.spalten.forEach((name, i) => {
      const spalte = indexZuSpalte(tab.startSpalte + i);
      const ref = `${spalte}${ziel}`;
      let c = this.#zelle(row, ref);
      const stil = vorlage && this.#zelle(vorlage, `${spalte}${vorlage.getAttribute('r')}`)?.getAttribute('s');
      if (!c) {
        c = sDoc.createElementNS(NS, 'c');
        c.setAttribute('r', ref);
        this.#zelleEinsortieren(row, c);
      }
      if (stil && !c.hasAttribute('s')) c.setAttribute('s', stil);
      this.#zelleSetzen(sDoc, c, werte[name]);
    });

    if (!platzhalter) {
      // Tabellenbereich vergrößern
      const tEintrag = await this.#doc(tab.tPfad);
      const tEl = tEintrag.doc.documentElement;
      const [[tc1, tr1], [tc2, tr2]] = bereichTeilen(tEl.getAttribute('ref'));
      tEl.setAttribute('ref', `${tc1}${tr1}:${tc2}${tr2 + 1}`);
      const af = kind(tEl, 'autoFilter');
      if (af) {
        const [[ac1, ar1], [ac2, ar2]] = bereichTeilen(af.getAttribute('ref'));
        af.setAttribute('ref', `${ac1}${ar1}:${ac2}${ar2 + 1}`);
      }
      tEintrag.geaendert = true;
    }
    this.#dimensionAktualisieren(sDoc);
    eintrag.geaendert = true;
    return ziel;
  }

  /** Alle Zeilen eines Blattes als Werte-Listen (ohne Kopfzeile). */
  async blattZeilen(blatt, mitKopf = false) {
    const sDoc = (await this.#doc(this.blaetter[blatt].pfad)).doc;
    const aus = [];
    for (const [r, row] of [...this.#zeilenMap(sDoc)].sort((a, b) => a[0] - b[0])) {
      if (r === 1 && !mitKopf) continue;
      const z = [];
      for (const c of kinder(row, 'c')) z[spalteZuIndex(zelleTeilen(c.getAttribute('r'))[0]) - 1] = this.#wert(c);
      aus.push(z);
    }
    return aus;
  }

  /** Hängt eine Zeile ans Ende eines Blattes ohne Tabelle (z. B. _System). */
  async zeileAnhaengen(blatt, werteListe) {
    const eintrag = await this.#doc(this.blaetter[blatt].pfad);
    const sDoc = eintrag.doc;
    const zeilen = this.#zeilenMap(sDoc);
    let letzte = 0;
    for (const [r, row] of zeilen) if (kinder(row, 'c').some((c) => kind(c, 'v') || kind(c, 'is') || kind(c, 'f'))) letzte = Math.max(letzte, r);
    const ziel = letzte + 1;
    const row = this.#zeileHolen(sDoc, zeilen, ziel);
    werteListe.forEach((w, i) => {
      const c = sDoc.createElementNS(NS, 'c');
      c.setAttribute('r', `${indexZuSpalte(i + 1)}${ziel}`);
      this.#zelleEinsortieren(row, c);
      this.#zelleSetzen(sDoc, c, w);
    });
    this.#dimensionAktualisieren(sDoc);
    eintrag.geaendert = true;
  }

  #zeilenVerschieben(sDoc, zeilen, ab, tab) {
    // Nur unterstützte Fälle: Zellbezüge in Formeln dürfen nicht von Zeilennummern abhängen.
    for (const [r, row] of zeilen) {
      if (r < ab) continue;
      for (const c of kinder(row, 'c')) {
        const f = kind(c, 'f');
        if (f && /(^|[^A-Za-z_\[])\$?[A-Z]{1,3}\$?\d+/.test(f.textContent.replace(/\[[^\]]*\]/g, ''))) {
          throw new ExcelFehler(`In „${tab.blatt}“ (Zelle ${c.getAttribute('r')}) steht eine Formel mit festen Zellbezügen. ` +
            'Damit nichts kaputtgeht, wurde nicht gebucht. Bitte melden.');
        }
      }
    }
    const sortiert = [...zeilen.entries()].filter(([r]) => r >= ab).sort((a, b) => b[0] - a[0]);
    for (const [r, row] of sortiert) {
      row.setAttribute('r', String(r + 1));
      for (const c of kinder(row, 'c')) {
        const [sp] = zelleTeilen(c.getAttribute('r'));
        c.setAttribute('r', `${sp}${r + 1}`);
      }
    }
    const ws = sDoc.documentElement;
    for (const el of [...ws.getElementsByTagNameNS(NS, 'dataValidation'), ...ws.getElementsByTagNameNS(NS, 'conditionalFormatting')]) {
      if (el.hasAttribute('sqref')) el.setAttribute('sqref', sqrefVerschieben(el.getAttribute('sqref'), ab));
    }
    for (const el of ws.getElementsByTagNameNS(NS, 'mergeCell')) {
      el.setAttribute('ref', sqrefVerschieben(el.getAttribute('ref'), ab));
    }
    for (const el of ws.getElementsByTagNameNS(NS, 'hyperlink')) {
      el.setAttribute('ref', sqrefVerschieben(el.getAttribute('ref'), ab));
    }
  }

  #zeileHolen(sDoc, zeilen, r) {
    if (zeilen.has(r)) return zeilen.get(r);
    const sd = sDoc.getElementsByTagNameNS(NS, 'sheetData')[0];
    const row = sDoc.createElementNS(NS, 'row');
    row.setAttribute('r', String(r));
    const danach = kinder(sd, 'row').find((x) => Number(x.getAttribute('r')) > r);
    sd.insertBefore(row, danach || null);
    zeilen.set(r, row);
    return row;
  }

  #zelleEinsortieren(row, c) {
    const idx = spalteZuIndex(zelleTeilen(c.getAttribute('r'))[0]);
    const danach = kinder(row, 'c').find((x) => spalteZuIndex(zelleTeilen(x.getAttribute('r'))[0]) > idx);
    row.insertBefore(c, danach || null);
    row.removeAttribute('spans');
  }

  #zelleSetzen(sDoc, c, w) {
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
    if (w == null || w === '' || (typeof w === 'object' && w.text === '')) return;
    if (typeof w !== 'object') w = typeof w === 'number' ? { zahl: w } : { text: String(w) };
    if ('formel' in w) {
      const f = sDoc.createElementNS(NS, 'f'); f.textContent = w.formel; c.append(f);
    } else if ('datum' in w) {
      if (!w.datum) return;
      const v = sDoc.createElementNS(NS, 'v'); v.textContent = String(datumZuSerial(w.datum)); c.append(v);
    } else if ('zahl' in w) {
      if (w.zahl == null || Number.isNaN(w.zahl)) return;
      const v = sDoc.createElementNS(NS, 'v'); v.textContent = String(w.zahl); c.append(v);
    } else {
      c.setAttribute('t', 'inlineStr');
      const is = sDoc.createElementNS(NS, 'is'); const t = sDoc.createElementNS(NS, 't');
      t.textContent = w.text;
      if (/^\s|\s$/.test(w.text)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      is.append(t); c.append(is);
    }
  }

  #dimensionAktualisieren(sDoc) {
    const dim = sDoc.getElementsByTagNameNS(NS, 'dimension')[0];
    if (!dim) return;
    let maxR = 1, maxC = 1;
    for (const c of sDoc.getElementsByTagNameNS(NS, 'c')) {
      const [sp, r] = zelleTeilen(c.getAttribute('r'));
      maxR = Math.max(maxR, r); maxC = Math.max(maxC, spalteZuIndex(sp));
    }
    dim.setAttribute('ref', `A1:${indexZuSpalte(maxC)}${maxR}`);
  }

  /** Erzeugt die neue Datei. Unveränderte Teile werden roh übernommen. */
  async speichern() {
    const ser = new XMLSerializer();
    const neu = [];
    for (const e of this.eintraege) {
      const d = this.docs.get(e.name);
      if (d && d.geaendert) {
        let xml = ser.serializeToString(d.doc);
        if (!xml.startsWith('<?xml')) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
        const roh = new TextEncoder().encode(xml);
        const gepackt = await packen(roh);
        neu.push({ ...e, methode: 8, crc: crc32(roh), groesse: roh.length, gepackt: gepackt.length, roh: gepackt });
      } else {
        neu.push(e);
      }
    }
    return zipSchreiben(neu);
  }
}

function normalisieren(pfad) {
  const aus = [];
  for (const t of pfad.split('/')) {
    if (t === '..') aus.pop(); else if (t && t !== '.') aus.push(t);
  }
  return aus.join('/');
}
