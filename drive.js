// Google Drive: Anmeldung (OAuth über Google Identity Services) und Dateizugriff.
// Das Zugriffstoken liegt nur im Arbeitsspeicher, nie im Browser-Speicher.
import { CONFIG } from './config.js';

const API = 'https://www.googleapis.com/drive/v3';
const UPLOAD = 'https://www.googleapis.com/upload/drive/v3';
const ORDNER_MIME = 'application/vnd.google-apps.folder';
export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

let token = null;
let tokenClient = null;
let wartend = null;

export class DriveFehler extends Error {}

function warteAufGoogle() {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    (function pruefe() {
      if (window.google?.accounts?.oauth2) return resolve();
      if (Date.now() - start > 15000) return reject(new DriveFehler('Google-Anmeldung konnte nicht geladen werden. Bitte Internetverbindung prüfen.'));
      setTimeout(pruefe, 100);
    })();
  });
}

/** Bereitet die Anmeldung vor (beim Laden der Seite), damit der Klick das Fenster sofort öffnet. */
export async function vorbereiten() {
  await warteAufGoogle();
  if (!tokenClient) {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.clientId,
      scope: CONFIG.scope,
      callback: (antwort) => {
        if (!wartend) return;
        const { resolve, reject } = wartend;
        wartend = null;
        if (antwort.error || !google.accounts.oauth2.hasGrantedAllScopes(antwort, CONFIG.scope)) {
          token = null;
          reject(new DriveFehler('Anmeldung abgebrochen oder Berechtigung nicht erteilt.'));
        } else {
          token = { wert: antwort.access_token, ablauf: Date.now() + (antwort.expires_in - 60) * 1000 };
          resolve();
        }
      },
      error_callback: () => {
        if (!wartend) return;
        const { reject } = wartend;
        wartend = null;
        reject(new DriveFehler('Anmeldefenster wurde geschlossen oder blockiert.'));
      },
    });
  }
}

/** Meldet an. interaktiv=false versucht es still (ohne Fenster), falls bereits zugestimmt. */
export function anmelden(interaktiv = true) {
  if (!tokenClient) return vorbereiten().then(() => anmelden(interaktiv));
  return new Promise((resolve, reject) => {
    wartend = { resolve, reject };
    tokenClient.requestAccessToken({ prompt: interaktiv ? '' : 'none' });
  });
}

export function abmelden() {
  if (token) google.accounts.oauth2.revoke(token.wert, () => {});
  token = null;
}

export const angemeldet = () => !!token && Date.now() < token.ablauf;

async function anfrage(url, optionen = {}, versuch = 0) {
  if (!angemeldet()) await anmelden(false).catch(() => anmelden(true));
  const antwort = await fetch(url, {
    ...optionen,
    headers: { ...(optionen.headers || {}), Authorization: `Bearer ${token.wert}` },
  });
  if (antwort.status === 401 && versuch === 0) {
    token = null;
    return anfrage(url, optionen, 1);
  }
  if (!antwort.ok) {
    // Keine Inhalte aus der Antwort anzeigen (Datenschutz), nur den Status.
    throw new DriveFehler(`Google Drive hat die Anfrage abgelehnt (Fehler ${antwort.status}).`);
  }
  return antwort;
}

const esc = (s) => s.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
const FELDER = 'id,name,mimeType,webViewLink,modifiedTime,version,md5Checksum';

export async function suche(name, elternId, mime) {
  let q = `name = '${esc(name)}' and '${elternId}' in parents and trashed = false`;
  if (mime) q += ` and mimeType = '${mime}'`;
  const r = await anfrage(`${API}/files?q=${encodeURIComponent(q)}&fields=files(${FELDER})&spaces=drive`);
  const { files } = await r.json();
  return files[0] || null;
}

export async function ordnerFindenOderAnlegen(name, elternId = 'root') {
  const vorhanden = await suche(name, elternId, ORDNER_MIME);
  if (vorhanden) return { ...vorhanden, neu: false };
  const r = await anfrage(`${API}/files?fields=${FELDER}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, mimeType: ORDNER_MIME, parents: [elternId] }),
  });
  return { ...(await r.json()), neu: true };
}

/** Lädt eine neue Datei hoch (keine Umwandlung in Google-Format). */
export async function dateiHochladen(name, elternId, blob, mime) {
  const meta = { name, parents: [elternId], mimeType: mime };
  const form = new FormData();
  form.append('metadata', new Blob([JSON.stringify(meta)], { type: 'application/json' }));
  form.append('file', blob);
  const r = await anfrage(`${UPLOAD}/files?uploadType=multipart&fields=${FELDER}`, { method: 'POST', body: form });
  return r.json();
}

export async function dateiLesen(id) {
  const r = await anfrage(`${API}/files/${id}?alt=media`);
  return r.blob();
}

export async function metadaten(id) {
  const r = await anfrage(`${API}/files/${id}?fields=${FELDER}`);
  return r.json();
}
