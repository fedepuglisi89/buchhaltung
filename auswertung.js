// Jahresübersicht direkt aus der Excel-Datei (die App rechnet die bekannten Formeln selbst nach).

const zahl = (v) => (typeof v === 'number' ? v : null);

/** Wert einer Zelle: Zahl direkt, Formel über bekannte Regeln, sonst gespeicherter Excel-Wert. */
function zeilenBetrag(blatt, w) {
  const spalte = { 'Km Fahrrad': 'Betrag (€)', Tmobile: 'Absetzbarer Betrag (€)', Bewirtungkosten: 'Absetzbarer Betrag (€)' }[blatt];
  const v = w[spalte];
  if (typeof v === 'number') return v;
  if (v && typeof v === 'object') {
    if (blatt === 'Km Fahrrad') return (zahl(w['Gefahrene Kilometer']) ?? 0) * (zahl(w['Satz (€/km)']) ?? 0);
    if (blatt === 'Bewirtungkosten') return (zahl(w['Betrag (€)']) ?? 0) * (zahl(w['Absetzbar (%)']) ?? 0);
    if (blatt === 'Tmobile') {
      const r = zahl(w['Rechnungsbetrag (€)']) ?? 0, p = zahl(w['davon privat (€)']) ?? 0, wl = zahl(w['davon WLAN (€)']) ?? 0;
      return (r - p - wl) * (zahl(w['Mobilfunk absetzbar (%)']) ?? 1) + wl * (zahl(w['WLAN absetzbar (%)']) ?? 0);
    }
    if (v.wert != null && !Number.isNaN(Number(v.wert))) return Number(v.wert);
  }
  return null;
}

export async function jahresuebersicht(mappe) {
  let unsicher = 0;
  const ein = await mappe.tabelle('Einkommen');
  let einnahmen = 0;
  for (const d of ein.daten) {
    const b = zahl(d.werte['Zahlung (brutto)']);
    if (b != null && d.werte.Rechnungsstatus !== 'Offen') einnahmen += b;
  }

  const quellen = {};
  for (const blatt of ['Km Fahrrad', 'Tmobile', 'Bewirtungkosten']) {
    const t = await mappe.tabelle(blatt).catch(() => null);
    if (!t) continue;
    for (const d of t.daten) if (d.werte.ID) quellen[d.werte.ID] = zeilenBetrag(blatt, d.werte);
  }
  const aus = await mappe.tabelle('Auskommen');
  let ausgaben = 0;
  for (const d of aus.daten) {
    const v = d.werte['Betrag (€)'];
    if (typeof v === 'number') ausgaben += v;
    else if (v && typeof v === 'object') {
      const q = d.werte.ID != null ? quellen[d.werte.ID] : null;
      if (q != null) ausgaben += q;
      else if (v.wert != null && !Number.isNaN(Number(v.wert))) ausgaben += Number(v.wert);
      else unsicher++;
    }
  }
  const r = (x) => Math.round(x * 100) / 100;
  return { einnahmen: r(einnahmen), ausgaben: r(ausgaben), differenz: r(einnahmen - ausgaben), unsicher };
}
