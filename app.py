# -*- coding: utf-8 -*-
"""
BewerbungsTerminal - Portabler lokaler Server & Controller
Alle Daten, Muster-Word-Layouts (.docx), Job-Listen, Lebensläufe und generierten
Bewerbungsunterlagen werden direkt im selben Programmordner gespeichert.
"""

import os
import sys
import json
import base64
import re
import datetime
import webbrowser
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn

from docx_engine import (
    inspect_docx,
    merge_docx_template,
    auto_map_blocks_heuristic,
    create_sample_layouts,
)
from ai_engine import (
    scrape_job_url,
    generate_tailored_application,
    ai_map_docx_blocks,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LAYOUTS_DIR = os.path.join(BASE_DIR, "layouts")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DATA_DIR = os.path.join(BASE_DIR, "data")
UI_DIR = os.path.join(BASE_DIR, "ui")

STATE_FILE = os.path.join(DATA_DIR, "state.json")
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")


DEFAULT_CV_TEXT = """Max Mustermann
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
• Sprachen: Deutsch (Muttersprache), Englisch (Verhandlungssicher C1)"""


def ensure_directories():
    os.makedirs(LAYOUTS_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(UI_DIR, exist_ok=True)
    create_sample_layouts(LAYOUTS_DIR)


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "provider": "smart_local",
        "api_key": "",
        "model": "",
        "ollama_url": "http://localhost:11434"
    }


def save_config(cfg: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_state() -> dict:
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "cv_text": DEFAULT_CV_TEXT,
        "selected_cover_layout": "Muster_Anschreiben_Executive_Blau.docx",
        "selected_cv_layout": "Muster_Lebenslauf_Executive_Blau.docx",
        "jobs": [
            {
                "id": "job_1",
                "url": "",
                "manual_text": "",
                "title": "",
                "company": "",
                "status": "idle",
                "result": None
            },
            {
                "id": "job_2",
                "url": "",
                "manual_text": "",
                "title": "",
                "company": "",
                "status": "idle",
                "result": None
            },
            {
                "id": "job_3",
                "url": "",
                "manual_text": "",
                "title": "",
                "company": "",
                "status": "idle",
                "result": None
            }
        ]
    }


def save_state(state: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def list_layouts() -> list:
    items = []
    if not os.path.exists(LAYOUTS_DIR):
        return items
    for fname in sorted(os.listdir(LAYOUTS_DIR)):
        if fname.lower().endswith(".docx") and not fname.startswith("~$"):
            fpath = os.path.join(LAYOUTS_DIR, fname)
            try:
                info = inspect_docx(fpath)
                stat = os.stat(fpath)
                items.append({
                    "filename": fname,
                    "size_kb": round(stat.st_size / 1024, 1),
                    "modified": datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M"),
                    "placeholders": info["placeholders"],
                    "block_count": info["block_count"],
                    "guessed_type": info["guessed_type"],
                    "preview_text": info["preview_text"]
                })
            except Exception as e:
                items.append({
                    "filename": fname,
                    "size_kb": 0,
                    "modified": "",
                    "placeholders": [],
                    "block_count": 0,
                    "guessed_type": "unknown",
                    "preview_text": f"Fehler beim Lesen: {e}"
                })
    return items


def list_outputs() -> list:
    items = []
    if not os.path.exists(OUTPUT_DIR):
        return items
    for fname in sorted(os.listdir(OUTPUT_DIR), reverse=True):
        if fname.lower().endswith(".docx") and not fname.startswith("~$"):
            fpath = os.path.join(OUTPUT_DIR, fname)
            stat = os.stat(fpath)
            items.append({
                "filename": fname,
                "size_kb": round(stat.st_size / 1024, 1),
                "created": datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M:%S"),
                "url": f"/files/output/{urllib.parse.quote(fname)}"
            })
    return items


def safe_filename(text: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9ÄÖÜäöüß\-_]+", "_", (text or "").strip())
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    return cleaned[:45] or "Dokument"


def build_placeholder_map(gen_data: dict) -> dict:
    """
    Erzeugt eine umfassende Platzhalter-Map für verschiedenste Schreibweisen in .docx Vorlagen.
    """
    mapping = {
        "{{NAME}}": gen_data.get("name", ""),
        "{{VORNAME_NACHNAME}}": gen_data.get("name", ""),
        "{{KONTAKT}}": gen_data.get("kontakt", ""),
        "{{ADRESSE}}": gen_data.get("kontakt", ""),
        "{{EMPFAENGER}}": gen_data.get("empfaenger", ""),
        "{{UNTERNEHMEN}}": gen_data.get("unternehmen", ""),
        "{{FIRMA}}": gen_data.get("unternehmen", ""),
        "{{DATUM}}": gen_data.get("datum", ""),
        "{{ZIELPOSITION}}": gen_data.get("zielposition", ""),
        "{{POSITION}}": gen_data.get("zielposition", ""),
        "{{JOBTITEL}}": gen_data.get("zielposition", ""),
        "{{BETREFF}}": gen_data.get("betreff", ""),
        "{{ANREDE}}": gen_data.get("anrede", ""),
        "{{ANSCHREIBEN}}": gen_data.get("anschreiben_text", ""),
        "{{ANSCHREIBEN_TEXT}}": gen_data.get("anschreiben_text", ""),
        "{{TEXT}}": gen_data.get("anschreiben_text", ""),
        "{{KURZPROFIL}}": gen_data.get("kurzprofil", ""),
        "{{PROFIL}}": gen_data.get("kurzprofil", ""),
        "{{BERUFSERFAHRUNG}}": gen_data.get("berufserfahrung", ""),
        "{{ERFAHRUNG}}": gen_data.get("berufserfahrung", ""),
        "{{WERDEGANG}}": gen_data.get("berufserfahrung", ""),
        "{{AUSBILDUNG}}": gen_data.get("ausbildung", ""),
        "{{STUDIUM}}": gen_data.get("ausbildung", ""),
        "{{BILDUNGSWEG}}": gen_data.get("ausbildung", ""),
        "{{KENNTNISSE}}": gen_data.get("kenntnisse", ""),
        "{{SKILLS}}": gen_data.get("kenntnisse", ""),
        "{{KOMPETENZEN}}": gen_data.get("kenntnisse", ""),
        "{{LEBENSLAUF}}": gen_data.get("lebenslauf_komplett", ""),
    }
    # Auch kleingeschriebene Varianten unterstützen
    lower_additions = {}
    for k, v in mapping.items():
        lower_additions[k.lower()] = v
        lower_additions[k.title()] = v
    mapping.update(lower_additions)
    return mapping


def merge_with_selected_layout(template_filename: str, doc_type: str, gen_data: dict, config: dict, out_filename: str) -> dict:
    """
    Führt das generierte Anschreiben oder den Lebenslauf mit dem gewählten .docx Muster-Layout zusammen.
    Unterstützt:
    1. Feste Platzhalter ({{ANSCHREIBEN_TEXT}}, {{NAME}}, etc.)
    2. Automatische KI-Block-Erkennung + Heuristik, falls das .docx Layout keine Platzhalter hat!
    """
    template_path = os.path.join(LAYOUTS_DIR, template_filename)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Layout-Datei '{template_filename}' nicht im Ordner layouts/ gefunden.")

    info = inspect_docx(template_path)
    placeholders_map = build_placeholder_map(gen_data)
    block_replacements = {}

    # Wenn das gewählte Layout keine {{...}} Platzhalter besitzt, aktiviere KI-Block-Mapping + Heuristik
    if not info["placeholders"]:
        block_replacements = auto_map_blocks_heuristic(info["blocks"], doc_type, gen_data)
        ai_blocks = ai_map_docx_blocks(info["blocks"], doc_type, gen_data, config)
        if ai_blocks:
            block_replacements.update(ai_blocks)

    out_path = os.path.join(OUTPUT_DIR, out_filename)
    merge_res = merge_docx_template(
        template_path=template_path,
        output_path=out_path,
        placeholders_map=placeholders_map,
        block_replacements=block_replacements
    )
    return {
        "filename": out_filename,
        "url": f"/files/output/{urllib.parse.quote(out_filename)}",
        "replaced_placeholders": merge_res["replaced_placeholders"],
        "replaced_blocks": merge_res["replaced_blocks"],
        "mode": "Platzhalter + Design-Merge" if info["placeholders"] else "KI-Auto-Block-Ersetzung"
    }


class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True


class RequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        # Ruhiges Server-Logging
        pass

    def _send_json(self, data: dict, status: int = 200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, filepath: str, content_type: str = None, download_name: str = None):
        if not os.path.exists(filepath) or not os.path.isfile(filepath):
            self.send_error(404, "File not found")
            return
        with open(filepath, "rb") as f:
            data = f.read()
        self.send_response(200)
        if not content_type:
            if filepath.endswith(".html"):
                content_type = "text/html; charset=utf-8"
            elif filepath.endswith(".css"):
                content_type = "text/css; charset=utf-8"
            elif filepath.endswith(".js"):
                content_type = "application/javascript; charset=utf-8"
            elif filepath.endswith(".docx"):
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            else:
                content_type = "application/octet-stream"
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        if download_name:
            encoded_name = urllib.parse.quote(download_name)
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded_name}")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/" or path == "/index.html":
            self._send_file(os.path.join(UI_DIR, "index.html"))
        elif path in ("/styles.css", "/app.js"):
            self._send_file(os.path.join(UI_DIR, path.lstrip("/")))
        elif path == "/api/state":
            state = load_state()
            config = load_config()
            layouts = list_layouts()
            outputs = list_outputs()
            self._send_json({
                "ok": True,
                "base_dir": BASE_DIR,
                "layouts_dir": LAYOUTS_DIR,
                "output_dir": OUTPUT_DIR,
                "state": state,
                "config": config,
                "layouts": layouts,
                "outputs": outputs
            })
        elif path.startswith("/files/layouts/"):
            fname = urllib.parse.unquote(path[len("/files/layouts/"):])
            safe_name = os.path.basename(fname)
            self._send_file(os.path.join(LAYOUTS_DIR, safe_name), download_name=safe_name)
        elif path.startswith("/files/output/"):
            fname = urllib.parse.unquote(path[len("/files/output/"):])
            safe_name = os.path.basename(fname)
            self._send_file(os.path.join(OUTPUT_DIR, safe_name), download_name=safe_name)
        else:
            self.send_error(404, "Not found")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length) if length > 0 else b"{}"
        try:
            body = json.loads(raw_body.decode("utf-8"))
        except Exception:
            body = {}

        try:
            if path == "/api/save-state":
                if "state" in body:
                    save_state(body["state"])
                if "config" in body:
                    save_config(body["config"])
                self._send_json({"ok": True})

            elif path == "/api/upload-layout":
                files = body.get("files", [])
                saved = []
                for fobj in files:
                    raw_name = os.path.basename(fobj.get("name", "Layout.docx"))
                    if not raw_name.lower().endswith(".docx"):
                        continue
                    b64_data = fobj.get("data_base64", "")
                    if "," in b64_data:
                        b64_data = b64_data.split(",", 1)[1]
                    file_bytes = base64.b64decode(b64_data)
                    dest_path = os.path.join(LAYOUTS_DIR, raw_name)
                    with open(dest_path, "wb") as out_f:
                        out_f.write(file_bytes)
                    saved.append(raw_name)
                self._send_json({
                    "ok": True,
                    "saved": saved,
                    "layouts": list_layouts()
                })

            elif path == "/api/extract-docx-text":
                b64_data = body.get("data_base64", "")
                if "," in b64_data:
                    b64_data = b64_data.split(",", 1)[1]
                file_bytes = base64.b64decode(b64_data)
                temp_path = os.path.join(DATA_DIR, "_temp_cv_extract.docx")
                with open(temp_path, "wb") as tf:
                    tf.write(file_bytes)
                info = inspect_docx(temp_path)
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
                full_text = "\n".join(b["text"] for b in info["blocks"])
                self._send_json({"ok": True, "text": full_text})

            elif path == "/api/delete-layout":
                fname = os.path.basename(body.get("filename", ""))
                target = os.path.join(LAYOUTS_DIR, fname)
                if os.path.exists(target):
                    os.remove(target)
                self._send_json({"ok": True, "layouts": list_layouts()})

            elif path == "/api/scrape-job":
                url = body.get("url", "").strip()
                manual = body.get("manual_text", "").strip()
                res = scrape_job_url(url if url else manual)
                if manual and res.get("ok"):
                    res["text"] = (res.get("text", "") + "\n" + manual).strip()
                elif manual and not res.get("ok"):
                    res = {
                        "ok": True,
                        "url": url or "Manueller Text",
                        "title": manual.splitlines()[0][:80] if manual.splitlines() else "Stellenbeschreibung",
                        "company": "",
                        "text": manual
                    }
                self._send_json(res)

            elif path == "/api/process-job":
                job = body.get("job", {})
                cv_text = body.get("cv_text", "")
                cover_layout = body.get("cover_layout", "Muster_Anschreiben_Executive_Blau.docx")
                cv_layout = body.get("cv_layout", "Muster_Lebenslauf_Executive_Blau.docx")
                config = load_config()
                if "config" in body and isinstance(body["config"], dict):
                    config.update(body["config"])

                url = (job.get("url") or "").strip()
                manual = (job.get("manual_text") or "").strip()

                if not url and not manual:
                    self._send_json({"ok": False, "error": "Bitte Job-Link oder Stellenbeschreibung einfügen."}, status=400)
                    return

                # 1. Job-Daten abrufen
                scraped = scrape_job_url(url if url else manual)
                if manual:
                    if scraped.get("ok"):
                        scraped["text"] = (scraped.get("text", "") + "\n\n" + manual).strip()
                    else:
                        scraped = {
                            "ok": True,
                            "url": url or "Manueller Text",
                            "title": manual.splitlines()[0][:80] if manual.splitlines() else "Stellenanzeige",
                            "company": "",
                            "text": manual
                        }

                if not scraped.get("text"):
                    # Selbst wenn Webseite Text blockiert hat, nutze zumindest die URL/Titel als Hinweis
                    scraped["text"] = f"Stellenanzeige unter Link: {url}"

                # 2. Maßgeschneidertes Anschreiben & Lebenslauf generieren
                gen_data = generate_tailored_application(scraped, cv_text, config)

                # 3. Mit gewählten Muster-Word-Layouts (.docx) mergen
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                job_slug = safe_filename(gen_data.get("unternehmen") or gen_data.get("zielposition") or "Stelle")

                cover_out_name = f"Bewerbungsschreiben_{job_slug}_{timestamp}.docx"
                cv_out_name = f"Lebenslauf_{job_slug}_{timestamp}.docx"

                cover_merge = merge_with_selected_layout(
                    template_filename=cover_layout,
                    doc_type="cover_letter",
                    gen_data=gen_data,
                    config=config,
                    out_filename=cover_out_name
                )

                cv_merge = merge_with_selected_layout(
                    template_filename=cv_layout,
                    doc_type="cv",
                    gen_data=gen_data,
                    config=config,
                    out_filename=cv_out_name
                )

                self._send_json({
                    "ok": True,
                    "scraped": {
                        "title": scraped.get("title") or gen_data.get("zielposition"),
                        "company": scraped.get("company") or gen_data.get("unternehmen"),
                        "url": url,
                        "warning": scraped.get("error", "")
                    },
                    "generated": gen_data,
                    "files": {
                        "cover_letter": cover_merge,
                        "cv": cv_merge
                    },
                    "outputs": list_outputs()
                })

            elif path == "/api/remake-docx":
                # Erlaubt dem Nutzer, nach manueller Text-Nachbearbeitung im Terminal die .docx Dateien sofort neu zu mergen!
                gen_data = body.get("generated", {})
                cover_layout = body.get("cover_layout", "Muster_Anschreiben_Executive_Blau.docx")
                cv_layout = body.get("cv_layout", "Muster_Lebenslauf_Executive_Blau.docx")
                config = load_config()

                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                job_slug = safe_filename(gen_data.get("unternehmen") or gen_data.get("zielposition") or "Stelle")

                cover_out_name = f"Bewerbungsschreiben_{job_slug}_{timestamp}.docx"
                cv_out_name = f"Lebenslauf_{job_slug}_{timestamp}.docx"

                cover_merge = merge_with_selected_layout(cover_layout, "cover_letter", gen_data, config, cover_out_name)
                cv_merge = merge_with_selected_layout(cv_layout, "cv", gen_data, config, cv_out_name)

                self._send_json({
                    "ok": True,
                    "files": {
                        "cover_letter": cover_merge,
                        "cv": cv_merge
                    },
                    "outputs": list_outputs()
                })

            elif path == "/api/open-path":
                target_type = body.get("type", "output_dir")
                filename = os.path.basename(body.get("filename", ""))
                if target_type == "output_dir":
                    os.startfile(OUTPUT_DIR)
                elif target_type == "layouts_dir":
                    os.startfile(LAYOUTS_DIR)
                elif target_type == "base_dir":
                    os.startfile(BASE_DIR)
                elif target_type == "output_file" and filename:
                    fpath = os.path.join(OUTPUT_DIR, filename)
                    if os.path.exists(fpath):
                        os.startfile(fpath)
                elif target_type == "layout_file" and filename:
                    fpath = os.path.join(LAYOUTS_DIR, filename)
                    if os.path.exists(fpath):
                        os.startfile(fpath)
                self._send_json({"ok": True})

            else:
                self._send_json({"ok": False, "error": f"Unbekannter Endpunkt: {path}"}, status=404)

        except Exception as e:
            self._send_json({"ok": False, "error": str(e)}, status=500)


def run_server(port: int = 8765, open_browser: bool = True):
    ensure_directories()
    server = ThreadedHTTPServer(("127.0.0.1", port), RequestHandler)
    url = f"http://127.0.0.1:{port}"
    print("=" * 68)
    print("  BEWERBUNGS-TERMINAL (PORTABLE EDITION) GESTARTET")
    print("=" * 68)
    print(f"  Programm-Ordner : {BASE_DIR}")
    print(f"  Word-Layouts    : {LAYOUTS_DIR}")
    print(f"  Fertige Dateien : {OUTPUT_DIR}")
    print(f"  Terminal-URL    : {url}")
    print("=" * 68)
    print("  Druecke STRG+C in diesem Fenster, um das Programm zu beenden.")
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nBewerbungsTerminal wird beendet...")
        server.server_close()


if __name__ == "__main__":
    no_browser = "--no-browser" in sys.argv
    port_arg = 8765
    for arg in sys.argv[1:]:
        if arg.startswith("--port="):
            port_arg = int(arg.split("=", 1)[1])
    run_server(port=port_arg, open_browser=not no_browser)
