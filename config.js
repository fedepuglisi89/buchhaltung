// Öffentliche Konfiguration. Die Client-ID ist nicht geheim (sie ist in jeder Web-App sichtbar).
export const CONFIG = {
  clientId: '54618218179-ollvh065omtfogiurcf07u0eeeonjjcj.apps.googleusercontent.com',
  // Nur Dateien, die diese App selbst erstellt oder öffnet – kein Zugriff auf den restlichen Drive.
  scope: 'https://www.googleapis.com/auth/drive.file',
  hauptordner: 'Fedephysio Buchhaltung',
  steuerjahre: [2026, 2027],
  standardEinstellungen: {
    Telekommunikationsanteil: 0.5,
    Mobilfunkanteil: 1.0,
    Bewirtungsanteil: 0.7,
    FahrradKilometersatz: 0.5,
    Arbeitsweg_km: 9.3,
    Waehrung: 'EUR',
    EigenerName: 'Federico Puglisi',
    EigeneRufnummer: '0676 709 21 86',
    StandardZahlungsartEinkommen: 'Überweisung',
    // Erkennung von Lieferanten (aus den Buchungen 2025). Muster = regulärer Ausdruck, ohne Groß-/Kleinschreibung.
    Lieferantenregeln: [
      { muster: 'Petra Walk', name: 'Petra Walk', kategorie: 'Praxis Mieten', blatt: 'Auskommen' },
      { muster: 'Physio ?Austria', name: 'Physio Austria', kategorie: 'Physio Austria', blatt: 'Auskommen' },
      { muster: 'aruba', name: 'Aruba.com', kategorie: 'Website Domain Fedephysio', blatt: 'Auskommen' },
      { muster: 'freelancer', name: 'freelancer.com', kategorie: 'Website Fedephysio', blatt: 'Auskommen' },
      { muster: 'Wiener ?Linien', name: 'Wiener Linien', kategorie: 'Jahreskarte', blatt: 'Auskommen' },
      { muster: 'Ryanair|Iberia|Austrian Airlines|Wizz ?Air|easyJet|Eurowings|ITA Airways', name: 'Fluggesellschaft', kategorie: 'Betriebliche Reiseaufwendungen', blatt: 'Auskommen' },
      { muster: 'Media ?Markt', name: 'MediaMarkt', kategorie: 'Arbeitsgerät', blatt: 'Auskommen' },
      { muster: 'Apple', name: 'Apple', kategorie: 'Arbeitsgerät', blatt: 'Auskommen' },
      { muster: 'IKEA', name: 'Ikea', kategorie: 'Arbeitsgerät', blatt: 'Auskommen' },
      { muster: '\\bAction\\b', name: 'Action', kategorie: 'Arbeitsgerät', blatt: 'Auskommen' },
      { muster: 'North ?Face|VF Austria', name: 'The North Face', kategorie: 'Arbeitskleidung', blatt: 'Auskommen' },
      { muster: 'Nike', name: 'Nike', kategorie: 'Arbeitskleidung', blatt: 'Auskommen' },
      { muster: 'Birkenstock', name: 'Birkenstock', kategorie: 'Arbeitskleidung', blatt: 'Auskommen' },
      { muster: 'New Balance', name: 'New Balance', kategorie: 'Arbeitskleidung', blatt: 'Auskommen' },
      { muster: 'Apotheke', name: 'Apotheke', kategorie: 'Medikament', blatt: 'Auskommen' },
      { muster: 'Italfisio', name: 'Italfisio', kategorie: 'Arbeitcreme', blatt: 'Auskommen' },
      { muster: 'Mosquito', name: 'Mosquito', kategorie: 'Bewirtungkosten', blatt: 'Bewirtungkosten' },
      { muster: 'TRIXIE|Kiddo', name: "Trixie Kiddo's", kategorie: 'Bewirtungkosten', blatt: 'Bewirtungkosten' },
      { muster: "w[uü]e?d.?n Goass|wuede-goass", name: "Einkehr zur wüd'n Goass", kategorie: 'Bewirtungkosten', blatt: 'Bewirtungkosten' },
    ],
  },
};

// Unterordner je Steuerjahr (Pfad relativ zum Jahresordner)
export const JAHRESORDNER = [
  ['_Eingang'],
  ['Einkommen'],
  ['Auskommen'],
  ['Auskommen', 'Tmobile'],
  ['Auskommen', 'Bewirtungkosten'],
  ['_Abgelehnt'],
  ['Scans'],
  ['_Backup'],
];
