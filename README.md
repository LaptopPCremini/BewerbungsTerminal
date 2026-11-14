# BewerbungsTerminal (GitHub Pages Edition)

Eine **100 % statische Web-Anwendung** zur automatisierten Erstellung individueller Bewerbungsschreiben und angepasster Lebensläufe auf Basis deiner eigenen **Muster-Word-Layouts (`.docx`)** – läuft komplett im Browser ohne Server oder Python-Backend.

## Funktionen

* **Drag & Drop für Word-Layouts (`.docx`):** Eigene Design-Vorlagen oben mittig ablegen, dauerhaft im Browser (**IndexedDB**) speichern und getrennt für Anschreiben und Lebenslauf auswählen.
* **Zwei Merge-Modi direkt im Browser (via JSZip):**
  * **Automatische Strukturerkennung** für beliebige `.docx`-Dokumente ohne Platzhalter (Schriftarten, Tabellen, Farben und Layout bleiben erhalten).
  * **Exakte Platzhalter** (`{{NAME}}`, `{{BETREFF}}`, `{{ANSCHREIBEN_TEXT}}`, `{{BERUFSERFAHRUNG}}`, `{{AUSBILDUNG}}`, `{{KENNTNISSE}}` u. a.).
* **Job-Links per Drag & Drop (links):** Beliebig viele Stellenanzeigen-Links (mit optionalem CORS-Proxy-Abruf) oder direkt kopierte Stellentexte in einzelne Felder ziehen.
* **Personalisierter Master-Lebenslauf (rechts):** Text-Editor mit automatischer Speicherung (**LocalStorage**) sowie Drag & Drop-Import aus `.docx` oder `.txt`.
* **Flexibles KI-Backend:** Funktioniert sofort lokal (**Smart-Lokal**) oder per direktem Browser-API-Aufruf für **Google Gemini**, **OpenAI**, **Anthropic Claude** oder **Ollama**.
* **Sofortiger Download:** Fertige Bewerbungsunterlagen einzeln als `.docx` oder gebündelt als `.zip` herunterladen.

## GitHub Pages aktivieren

1. Repository auf GitHub hochladen (`main` Branch).
2. In den Repository-Einstellungen unter **Settings → Pages** als Source **Deploy from a branch** (`main` / `/ (root)`) auswählen.
3. Nach wenigen Sekunden ist das Terminal unter `https://<dein-benutzername>.github.io/BewerbungsTerminal/` direkt im Browser nutzbar.
