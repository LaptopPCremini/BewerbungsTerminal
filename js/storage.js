// ==========================================================================
// STORAGE ENGINE (100 % Browser-Speicher: IndexedDB + LocalStorage)
// Ersetzt die lokalen Ordner /layouts und /data für GitHub Pages
// ==========================================================================

const DB_NAME = "BewerbungsTerminalDB";
const DB_VERSION = 1;
const STORE_LAYOUTS = "layouts";

const LS_STATE_KEY = "bewerbungsterminal_state_v1";
const LS_CONFIG_KEY = "bewerbungsterminal_config_v1";

export const DEFAULT_CV_TEXT = `Max Mustermann
Musterstraße 42, 10115 Berlin | +49 170 1234567 | max.mustermann@email.de

KURZPROFIL
Engagierter Softwareentwickler & Projektkoordinator mit über 4 Jahren Erfahrung in der Konzeption, Entwicklung und Optimierung moderner Web- und Geschäftsanwendungen. Starke analytische Fähigkeiten, hohe Kundenorientierung und strukturierte Arbeitsweise in agilen Teams.

BERUFSERFAHRUNG
04/2023 – heute | Senior Solution Developer | TechNova Solutions GmbH, Berlin
• Leitung und Umsetzung digitaler Kundenprojekte mit Python, TypeScript und modernen Cloud-Architekturen
• Automatisierung interner Geschäftsprozesse und Schnittstellen-Integration (Effizienzsteigerung um 35 %)
• Enge Abstimmung mit Stakeholdern, Anforderungsanalyse und technische Projektleitung im Scrum-Team

08/2020 – 03/2023 | Full-Stack Software Engineer | DigitalWerk AG, Hamburg
• Entwicklung und Wartung skalierbarer Webanwendungen und Daten-Dashboards
• Optimierung von SQL-Datenbanken, API-Design und Qualitätssicherung durch automatisierte Tests
• Betreuung von Unternehmenskunden sowie Durchführung technischer Workshops

AUSBILDUNG & STUDIUM
10/2016 – 07/2020 | B.Sc. Wirtschaftsinformatik | Technische Universität Berlin
• Schwerpunkt: Software Engineering, Datenanalyse & IT-Projektmanagement (Abschlussnote: 1,7)

KENNTNISSE & KOMPETENZEN
• Fachkompetenzen: Projektmanagement, Prozessoptimierung, Softwarearchitektur, Agile / Scrum
• IT & Tools: Python, JavaScript/TypeScript, React, SQL, Git, Docker, MS Office, Cloud (AWS/Azure)
• Sprachen: Deutsch (Muttersprache), Englisch (Verhandlungssicher C1)`;

function openDatabase() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_LAYOUTS)) {
        db.createObjectStore(STORE_LAYOUTS, { keyPath: "filename" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function getAllLayoutsFromDB() {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_LAYOUTS, "readonly");
    const store = tx.objectStore(STORE_LAYOUTS);
    const req = store.getAll();
    req.onsuccess = () => {
      const items = (req.result || []).sort((a, b) => a.filename.localeCompare(b.filename));
      resolve(items);
    };
    req.onerror = () => reject(req.error);
  });
}

export async function saveLayoutToDB(layoutObj) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_LAYOUTS, "readwrite");
    const store = tx.objectStore(STORE_LAYOUTS);
    const req = store.put(layoutObj);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

export async function deleteLayoutFromDB(filename) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_LAYOUTS, "readwrite");
    const store = tx.objectStore(STORE_LAYOUTS);
    const req = store.delete(filename);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

export function loadStateFromLS() {
  try {
    const raw = localStorage.getItem(LS_STATE_KEY);
    if (raw) {
      return JSON.parse(raw);
    }
  } catch (_) {}
  return {
    cv_text: DEFAULT_CV_TEXT,
    selected_cover_layout: "Muster_Anschreiben_Executive_Blau.docx",
    selected_cv_layout: "Muster_Lebenslauf_Executive_Blau.docx",
    jobs: [
      { id: "job_1", url: "", manual_text: "", title: "", company: "", status: "idle", result: null },
      { id: "job_2", url: "", manual_text: "", title: "", company: "", status: "idle", result: null },
      { id: "job_3", url: "", manual_text: "", title: "", company: "", status: "idle", result: null }
    ]
  };
}

export function saveStateToLS(state) {
  try {
    // Keine großen Blobs in LocalStorage serialisieren – nur Text & Metadaten
    const cleanJobs = (state.jobs || []).map((j) => ({
      id: j.id,
      url: j.url || "",
      manual_text: j.manual_text || "",
      title: j.title || "",
      company: j.company || "",
      status: j.status === "running" ? "idle" : (j.status || "idle"),
      result: null
    }));
    localStorage.setItem(
      LS_STATE_KEY,
      JSON.stringify({
        cv_text: state.cv_text || "",
        selected_cover_layout: state.selected_cover_layout || "",
        selected_cv_layout: state.selected_cv_layout || "",
        jobs: cleanJobs
      })
    );
  } catch (_) {}
}

export function loadConfigFromLS() {
  try {
    const raw = localStorage.getItem(LS_CONFIG_KEY);
    if (raw) {
      return JSON.parse(raw);
    }
  } catch (_) {}
  return {
    provider: "smart_local",
    api_key: "",
    model: "",
    ollama_url: "http://localhost:11434",
    use_cors_proxy: true
  };
}

export function saveConfigToLS(config) {
  try {
    localStorage.setItem(LS_CONFIG_KEY, JSON.stringify(config));
  } catch (_) {}
}
