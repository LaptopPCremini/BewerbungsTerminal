# BewerbungsTerminal (Portable Edition)

Eine portable deutsche Desktop-/Terminal-Anwendung zur automatisierten Erstellung individueller Bewerbungsschreiben und angepasster Lebensläufe auf Basis deiner eigenen **Muster-Word-Layouts (`.docx`)**.

## Funktionen

* **Drag & Drop für Word-Layouts (`.docx`):** Eigene Design-Vorlagen oben mittig ablegen, dauerhaft in `/layouts` speichern und getrennt für Anschreiben und Lebenslauf auswählen.
* **Zwei Merge-Modi gleichzeitig:**
  * **Automatische Struktur- & KI-Erkennung** für beliebige Word-Dokumente ohne Platzhalter (Schriftarten, Tabellen, Farben und Layout bleiben erhalten).
  * **Exakte Platzhalter** (`{{NAME}}`, `{{BETREFF}}`, `{{ANSCHREIBEN_TEXT}}`, `{{BERUFSERFAHRUNG}}`, `{{AUSBILDUNG}}`, `{{KENNTNISSE}}` u. a.).
* **Job-Links per Drag & Drop (links):** Beliebig viele Stellenanzeigen-Links (oder Stellentexte) in einzelne Felder ziehen. Jede Stelle wird separat verarbeitet.
* **Personalisierter Master-Lebenslauf (rechts):** Text-Editor mit Autosave sowie Drag & Drop-Import aus `.docx` oder `.txt`.
* **Flexibles KI-Backend:** Funktioniert sofort lokal (**Smart-Lokal**) oder optional mit **Google Gemini**, **OpenAI**, **Anthropic Claude** oder **Ollama**.
* **100 % Portabel:** Alle Vorlagen (`/layouts`), Daten (`/data`) und fertigen Word-Dokumente (`/output`) liegen direkt im Programmordner.

## Schnellstart (Windows)

1. Repository herunterladen / klonen:
   ```bash
   git clone https://github.com/LaptopPCremini/BewerbungsTerminal.git
   cd BewerbungsTerminal
   ```
2. Voraussetzung: **Python 3.10+** (keine externen Python-Pakete nötig – nutzt ausschließlich die Python-Standardbibliothek).
3. Programm starten per Doppelklick auf **`Start_BewerbungsTerminal.bat`** oder im Terminal:
   ```bash
   python app.py
   ```
   Die Oberfläche öffnet sich automatisch unter `http://127.0.0.1:8765`.
