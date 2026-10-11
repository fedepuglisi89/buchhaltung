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
  ['_Backup'],
];
