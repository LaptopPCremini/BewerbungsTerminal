# -*- coding: utf-8 -*-
"""
AI & Job-Scraping Engine für BewerbungsTerminal
Unterstützt:
- Intelligentes Scraping von Job-URLs (inkl. JSON-LD JobPosting Extraktion + HTML-Textbereinigung)
- Multi-Provider KI: Google Gemini, OpenAI, Anthropic Claude, Lokales Ollama
- Integrierter Smart-Generator (falls noch kein API-Key hinterlegt ist, funktioniert die App sofort offline!)
- Automatische KI-Block-Erkennung für beliebige .docx Muster-Vorlagen
"""

import json
import re
import datetime
import urllib.request
import urllib.error
import ssl
from html import unescape
from html.parser import HTMLParser


class _HTMLTextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.result = []
        self.skip_tags = {"script", "style", "noscript", "svg", "footer", "nav", "header", "form"}
        self.skip_depth = 0
        self.in_title = False
        self.title = ""

    def handle_starttag(self, tag, attrs):
        tag_l = tag.lower()
        if tag_l in self.skip_tags:
            self.skip_depth += 1
        elif tag_l == "title":
            self.in_title = True
        elif tag_l in ("p", "div", "br", "li", "h1", "h2", "h3", "h4", "tr", "section", "article"):
            self.result.append("\n")

    def handle_endtag(self, tag):
        tag_l = tag.lower()
        if tag_l in self.skip_tags and self.skip_depth > 0:
            self.skip_depth -= 1
        elif tag_l == "title":
            self.in_title = False
        elif tag_l in ("p", "div", "li", "h1", "h2", "h3", "h4", "section", "article"):
            self.result.append("\n")

    def handle_data(self, data):
        if self.in_title:
            self.title += data
        if self.skip_depth == 0:
            self.result.append(data)

    def get_text(self) -> str:
        raw = "".join(self.result)
        lines = [re.sub(r"[ \t]+", " ", line).strip() for line in raw.splitlines()]
        cleaned = "\n".join(line for line in lines if line)
        return cleaned


def scrape_job_url(url_or_text: str) -> dict:
    """
    Lädt eine Stellenanzeige von einer URL oder verarbeitet direkt eingefügten Text.
    Extrahiert falls vorhanden strukturierte JSON-LD JobPosting-Daten (Titel, Firma, Ort, Beschreibung).
    """
    raw = (url_or_text or "").strip()
    if not raw:
        return {"ok": False, "error": "Kein Link oder Text angegeben."}

    # Falls kein HTTP-Link, sondern direkt Text gedroppt/eingefügt wurde
    if not raw.lower().startswith(("http://", "https://")):
        first_line = raw.splitlines()[0][:80] if raw.splitlines() else "Manuelle Stellenbeschreibung"
        return {
            "ok": True,
            "url": "Manueller Text-Input",
            "title": first_line,
            "company": "",
            "text": raw
        }

    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7",
    }

    try:
        req = urllib.request.Request(raw, headers=headers)
        with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
            charset = resp.headers.get_content_charset() or "utf-8"
            html_bytes = resp.read()
            html_str = html_bytes.decode(charset, errors="replace")
    except Exception as e:
        return {
            "ok": False,
            "url": raw,
            "title": "Link konnte nicht direkt geladen werden",
            "company": "",
            "text": "",
            "error": f"Webseite blockiert automatischen Abruf oder ist nicht erreichbar ({str(e)}). Tipp: Kopiere den Stellentext kurz per Drag & Drop in das Zusatztext-Feld dieses Jobs."
        }

    # 1. Versuche JSON-LD JobPosting zu finden
    job_title = ""
    company_name = ""
    ld_description = ""

    ld_matches = re.findall(
        r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        html_str,
        flags=re.DOTALL | re.IGNORECASE
    )
    for ld_raw in ld_matches:
        try:
            data = json.loads(ld_raw.strip())
            items = data if isinstance(data, list) else [data]
            for item in items:
                if isinstance(item, dict) and "@graph" in item and isinstance(item["@graph"], list):
                    items.extend(item["@graph"])
            for item in items:
                if isinstance(item, dict) and "JobPosting" in str(item.get("@type", "")):
                    job_title = item.get("title", "") or job_title
                    org = item.get("hiringOrganization", {})
                    if isinstance(org, dict):
                        company_name = org.get("name", "") or company_name
                    desc_html = item.get("description", "")
                    if desc_html:
                        parser = _HTMLTextExtractor()
                        parser.feed(unescape(desc_html))
                        ld_description = parser.get_text()
        except Exception:
            pass

    # 2. Regulärer HTML-Text-Extraktor
    parser = _HTMLTextExtractor()
    try:
        parser.feed(html_str)
    except Exception:
        pass

    page_title = unescape(parser.title.strip()) if parser.title else ""
    body_text = parser.get_text()

    final_title = job_title or page_title or raw
    final_text = ld_description if len(ld_description) > 200 else body_text
    if len(final_text) > 12000:
        final_text = final_text[:12000]

    return {
        "ok": True,
        "url": raw,
        "title": final_title[:140],
        "company": company_name[:100],
        "text": final_text
    }


def _call_llm_json(prompt: str, system_prompt: str, config: dict) -> dict:
    """
    Ruft den konfigurierten KI-Provider auf und gibt das geparste JSON-Objekt zurück.
    Unterstützt: gemini, openai, anthropic, ollama.
    """
    provider = (config.get("provider") or "smart_local").lower().strip()
    api_key = (config.get("api_key") or "").strip()
    model = (config.get("model") or "").strip()
    ollama_url = (config.get("ollama_url") or "http://localhost:11434").rstrip("/")

    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    raw_response_text = ""

    if provider == "gemini" and api_key:
        model_name = model or "gemini-2.5-flash"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        payload = {
            "systemInstruction": {"parts": [{"text": system_prompt}]},
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.4,
                "responseMimeType": "application/json"
            }
        }
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
            res_data = json.loads(resp.read().decode("utf-8"))
            raw_response_text = res_data["candidates"][0]["content"]["parts"][0]["text"]

    elif provider == "openai" and api_key:
        model_name = model or "gpt-4o-mini"
        url = "https://api.openai.com/v1/chat/completions"
        payload = {
            "model": model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.4,
            "response_format": {"type": "json_object"}
        }
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
            res_data = json.loads(resp.read().decode("utf-8"))
            raw_response_text = res_data["choices"][0]["message"]["content"]

    elif provider == "anthropic" and api_key:
        model_name = model or "claude-3-5-sonnet-latest"
        url = "https://api.anthropic.com/v1/messages"
        payload = {
            "model": model_name,
            "max_tokens": 4096,
            "system": system_prompt + "\nAntworte AUSSCHLIESSLICH mit validem JSON ohne Markdown-Codeblöcke.",
            "messages": [
                {"role": "user", "content": prompt}
            ]
        }
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01"
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
            res_data = json.loads(resp.read().decode("utf-8"))
            raw_response_text = res_data["content"][0]["text"]

    elif provider == "ollama":
        model_name = model or "llama3.1"
        url = f"{ollama_url}/api/chat"
        payload = {
            "model": model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            "format": "json",
            "stream": False
        }
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=90, context=ctx) as resp:
            res_data = json.loads(resp.read().decode("utf-8"))
            raw_response_text = res_data["message"]["content"]
    else:
        raise ValueError("Kein externer KI-Provider aktiv oder API-Key fehlt.")

    # Bereinige eventuelle ```json ... ``` Blöcke
    cleaned = raw_response_text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    return json.loads(cleaned)


def _smart_local_generate(job_info: dict, cv_text: str) -> dict:
    """
    Intelligente lokale Generierungs-Engine (Offline / Sofort-Modus ohne API-Key).
    Analysiert den Lebenslauf-Text und die Stellenanzeige, extrahiert relevante Keywords,
    Name, Kontaktdaten, Berufserfahrung, Ausbildung und Skills und verfasst ein
    professionelles, auf die Stelle abgestimmtes deutsches Bewerbungsschreiben + Lebenslauf.
    """
    today_str = datetime.date.today().strftime("%d.%m.%Y")
    cv_lines = [line.strip() for line in (cv_text or "").splitlines() if line.strip()]

    # Name & Kontakt aus Lebenslauf extrahieren
    name = "Max Mustermann"
    if cv_lines:
        first = re.sub(r"^(lebenslauf|curriculum vitae|profil|name\s*:)\s*", "", cv_lines[0], flags=re.I).strip()
        if 2 <= len(first) <= 55:
            name = first

    email_match = re.search(r"[\w\.\-]+@[\w\.\-]+\.\w+", cv_text or "")
    phone_match = re.search(r"(\+?\d[\d\s/\-]{7,16}\d)", cv_text or "")
    contact_parts = []
    if email_match:
        contact_parts.append(email_match.group(0))
    if phone_match:
        contact_parts.append(phone_match.group(1).strip())
    for line in cv_lines[1:6]:
        if any(w in line.lower() for w in ["str.", "straße", "weg ", "platz", "berlin", "münchen", "hamburg", "köln", "frankfurt", "stuttgart", "düsseldorf", "wien", "zürich"]) or re.search(r"\b\d{5}\b", line):
            contact_parts.insert(0, line)
            break
    kontakt = "  |  ".join(dict.fromkeys(contact_parts)) if contact_parts else "Musterstraße 1, 10115 Berlin  |  kontakt@beispiel.de"

    # Sektionen aus dem Lebenslauf-Rohtext intelligent gliedern
    sections = {"kurzprofil": [], "berufserfahrung": [], "ausbildung": [], "kenntnisse": []}
    current_sec = "kurzprofil"
    for line in cv_lines[1:]:
        low = line.lower()
        if any(k in low for k in ["berufserfahrung", "werdegang", "praxis", "berufliche stationen", "erfahrung"]):
            current_sec = "berufserfahrung"
            continue
        elif any(k in low for k in ["ausbildung", "studium", "bildungsweg", "schule", "akademisch"]):
            current_sec = "ausbildung"
            continue
        elif any(k in low for k in ["kenntnisse", "skills", "fähigkeiten", "kompetenzen", "sprachen", "it-", "zertifikat"]):
            current_sec = "kenntnisse"
            continue
        sections[current_sec].append(line)

    # Jobtitel & Firma aus Job-Info bestimmen
    raw_title = (job_info.get("title") or "").strip()
    job_text = (job_info.get("text") or "").strip()
    company = (job_info.get("company") or "").strip()

    # Bereinige typische Seitentitel-Anhänge wie " - StepStone" oder " | Indeed.com"
    clean_title = re.split(r"\s+[\|\-–—]\s+", raw_title)[0].strip() if raw_title else ""
    if not clean_title or clean_title.lower().startswith("http") or "manuell" in clean_title.lower():
        # Versuche Jobtitel aus der ersten sinnvollen Zeile des Stellentextes zu lesen
        for jline in job_text.splitlines():
            jl = jline.strip()
            if 5 < len(jl) < 95 and not jl.lower().startswith("http"):
                clean_title = jl
                break
    if not clean_title:
        clean_title = "die ausgeschriebene Position"

    if not company:
        comp_match = re.search(r"(?:bei der|bei|unternehmen|firma|arbeitgeber)\s+([A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9&\.\-\s]{2,40}?(?:GmbH|AG|SE|KG|UG|Inc\.|Group|Solutions|Systems|Partner|Consulting))", job_text)
        if comp_match:
            company = comp_match.group(1).strip()

    empfaenger = f"{company}\nPersonalabteilung" if company else "Personalabteilung\n[Unternehmensname]"

    # Extrahiere wichtige Schlüsselbegriffe aus Stellentext + Lebenslauf für gezielten Bezug
    common_skills = [
        "Projektmanagement", "Python", "JavaScript", "TypeScript", "React", "Java", "C#", "SQL",
        "Cloud", "AWS", "Azure", "Docker", "Kubernetes", "SAP", "Marketing", "Vertrieb", "Sales",
        "Kundenbetreuung", "Controlling", "Finanzen", "Buchhaltung", "HR", "Personalwesen",
        "Design", "UI/UX", "Agil", "Scrum", "Datenanalyse", "KI", "Machine Learning", "IT-Support",
        "Qualitätsmanagement", "Logistik", "Einkauf", "Kommunikation", "Teamführung", "Strategie",
        "MS Office", "Excel", "PowerBI", "Prozessoptimierung", "Beratung", "Consulting", "Engineering"
    ]
    matched_keywords = []
    combined_source = (job_text + "\n" + cv_text).lower()
    for kw in common_skills:
        if kw.lower() in combined_source:
            matched_keywords.append(kw)
    if not matched_keywords:
        matched_keywords = ["strukturierte Arbeitsweise", "hohe Einsatzbereitschaft", "schnelle Auffassungsgabe", "Teamorientierung"]

    top_kw_str = ", ".join(matched_keywords[:5])

    beruf_text = "\n".join(sections["berufserfahrung"]).strip()
    ausb_text = "\n".join(sections["ausbildung"]).strip()
    skills_text = "\n".join(sections["kenntnisse"]).strip()
    profil_raw = "\n".join(sections["kurzprofil"]).strip()

    if not beruf_text:
        beruf_text = cv_text.strip() or "Mehrjährige praktische Erfahrung in relevanten Fachprojekten und erfolgreiche Umsetzung anspruchsvoller Aufgabenstellungen."
    if not ausb_text:
        ausb_text = "Fundierte fachliche Ausbildung sowie kontinuierliche Weiterbildung in aktuellen Methoden und Technologien."
    if not skills_text:
        skills_text = f"Kernkompetenzen: {top_kw_str}\nSprachen: Deutsch (fließend), Englisch (gute Kenntnisse)"

    kurzprofil = (
        f"Engagierte und ergebnisorientierte Fachkraft mit Fokus auf { clean_title }. "
        f"Fundierte Erfahrung in den Bereichen {top_kw_str} sowie ausgeprägte Stärke in der zielgerichteten Umsetzung anspruchsvoller Projekte."
    )
    if profil_raw and len(profil_raw) > 40 and not beruf_text.startswith(profil_raw[:30]):
        kurzprofil = f"{kurzprofil}\n{profil_raw[:350]}"

    firma_phrase = f"bei {company}" if company else "in Ihrem Unternehmen"
    betreff = f"Bewerbung als {clean_title}"
    anrede = "Sehr geehrte Damen und Herren,"

    anschreiben_text = (
        f"mit großem Interesse habe ich Ihre Stellenausschreibung für die Position als {clean_title} gelesen. "
        f"Da meine beruflichen Schwerpunkte und Qualifikationen hervorragend zu den von Ihnen gesuchten Anforderungen passen, "
        f"möchte ich meine Expertise gerne gewinnbringend {firma_phrase} einbringen.\n\n"
        f"In meinem bisherigen Werdegang konnte ich umfassende praktische Erfahrungen sammeln und meine Kompetenzen insbesondere "
        f"in den Bereichen {top_kw_str} erfolgreich unter Beweis stellen. Dabei zeichne ich mich durch eine analytische, "
        f"qualitätsbewusste und lösungsorientierte Arbeitsweise aus, mit der ich mich schnell und sicher in neue Themengebiete einarbeite.\n\n"
        f"Besonders reizt mich an der ausgeschriebenen Stelle als {clean_title} die Möglichkeit, verantwortungsvolle Aufgaben aktiv "
        f"mitzugestalten und gemeinsam im Team nachhaltige Ergebnisse zu erzielen. Die Verbindung aus fachlicher Tiefe, "
        f"Zuverlässigkeit und hoher Eigeninitiative bildet für mich die Grundlage einer erfolgreichen Zusammenarbeit.\n\n"
        f"Über die Gelegenheit, mich Ihnen in einem persönlichen Gespräch vorzustellen und Sie von meiner Motivation zu überzeugen, "
        f"freue ich mich sehr."
    )

    full_cv_text = (
        f"{name}\n{clean_title} | {kontakt}\n\n"
        f"KURZPROFIL\n{kurzprofil}\n\n"
        f"BERUFSERFAHRUNG\n{beruf_text}\n\n"
        f"AUSBILDUNG & STUDIUM\n{ausb_text}\n\n"
        f"KENNTNISSE & KOMPETENZEN\n{skills_text}"
    )

    return {
        "name": name,
        "kontakt": kontakt,
        "empfaenger": empfaenger,
        "datum": today_str,
        "zielposition": clean_title,
        "unternehmen": company or "Unternehmen",
        "betreff": betreff,
        "anrede": anrede,
        "anschreiben_text": anschreiben_text,
        "kurzprofil": kurzprofil,
        "berufserfahrung": beruf_text,
        "ausbildung": ausb_text,
        "kenntnisse": skills_text,
        "lebenslauf_komplett": full_cv_text,
        "used_provider": "Smart-Template-Engine (Lokal)"
    }


def generate_tailored_application(job_info: dict, cv_text: str, config: dict) -> dict:
    """
    Erzeugt für eine einzelne Stellenanzeige + den persönlichen Lebenslauf:
    1. Ein maßgeschneidertes, überzeugendes Bewerbungsschreiben auf Deutsch
    2. Einen auf die Stelle optimierten, strukturierten Lebenslauf auf Deutsch
    Nutzt den konfigurierten KI-Provider (Gemini, OpenAI, Claude, Ollama) oder
    fällt nahtlos auf die lokale Smart-Engine zurück.
    """
    today_str = datetime.date.today().strftime("%d.%m.%Y")
    provider = (config.get("provider") or "smart_local").lower().strip()
    api_key = (config.get("api_key") or "").strip()

    if provider != "smart_local" and (api_key or provider == "ollama"):
        system_prompt = (
            "Du bist ein erstklassiger deutscher Karriere-Coach und HR-Experte. "
            "Deine Aufgabe ist es, anhand einer Stellenanzeige und des bisherigen Lebenslaufs des Bewerbers "
            "ein individuell auf genau diese Stelle zugeschnittenes, überzeugendes deutsches Bewerbungsschreiben (Anschreiben) "
            "sowie einen passend auf die Stelle optimierten, professionell gegliederten Lebenslauf zu erstellen. "
            "Erfinde keine falschen Stationen hinzu, aber hebe genau die Erfahrungen, Erfolge und Kompetenzen aus dem Lebenslauf hervor, "
            "die für diese konkrete Stellenanzeige besonders relevant sind. "
            "Antworte AUSSCHLIESSLICH als valides JSON-Objekt mit exakt diesen Schlüsseln:\n"
            "{\n"
            '  "name": "Vor- und Nachname des Bewerbers",\n'
            '  "kontakt": "Adresse | Telefon | E-Mail | LinkedIn (kompakt in 1-2 Zeilen)",\n'
            '  "empfaenger": "Unternehmensname und Ansprechpartner/Abteilung für den Briefkopf",\n'
            '  "datum": "Heutiges Datum im Format TT.MM.JJJJ",\n'
            '  "zielposition": "Exakter Titel der Zielposition aus der Stellenanzeige",\n'
            '  "unternehmen": "Name des Unternehmens aus der Stellenanzeige",\n'
            '  "betreff": "Aussagekräftige Betreffzeile, z. B. Bewerbung als ...",\n'
            '  "anrede": "Passende Anrede, z. B. Sehr geehrte Damen und Herren, oder mit Namen",\n'
            '  "anschreiben_text": "Der komplette Haupttext des Anschreibens (3-4 Absätze, getrennt durch \\n\\n, OHNE Betreff, OHNE Anrede und OHNE Grußformel am Ende)",\n'
            '  "kurzprofil": "Auf die Stelle zugeschnittenes Kurzprofil (3-4 prägnante Sätze)",\n'
            '  "berufserfahrung": "Chronologisch strukturierte Berufserfahrung (relevanteste Aufgaben & Erfolge passend zur Stelle hervorgehoben, sauber mit Zeilenumbrüchen formatiert)",\n'
            '  "ausbildung": "Strukturierter Bildungsweg / Studium / Ausbildung",\n'
            '  "kenntnisse": "Auf die Stelle abgestimmte Fachkenntnisse, IT-Skills, Zertifikate und Sprachen"\n'
            "}"
        )

        user_prompt = (
            f"HEUTIGES DATUM: {today_str}\n\n"
            f"=== STELLENANZEIGE ===\n"
            f"URL: {job_info.get('url', '')}\n"
            f"Titel: {job_info.get('title', '')}\n"
            f"Unternehmen: {job_info.get('company', '')}\n"
            f"Inhalt der Stellenanzeige:\n{job_info.get('text', '')}\n\n"
            f"=== BISHERIGER LEBENSLAUF DES BEWERBERS ===\n"
            f"{cv_text}\n"
        )

        try:
            ai_res = _call_llm_json(user_prompt, system_prompt, config)
            ai_res.setdefault("datum", today_str)
            ai_res["lebenslauf_komplett"] = (
                f"{ai_res.get('name', '')}\n"
                f"{ai_res.get('zielposition', '')} | {ai_res.get('kontakt', '')}\n\n"
                f"KURZPROFIL\n{ai_res.get('kurzprofil', '')}\n\n"
                f"BERUFSERFAHRUNG\n{ai_res.get('berufserfahrung', '')}\n\n"
                f"AUSBILDUNG & STUDIUM\n{ai_res.get('ausbildung', '')}\n\n"
                f"KENNTNISSE & KOMPETENZEN\n{ai_res.get('kenntnisse', '')}"
            )
            ai_res["used_provider"] = f"KI ({provider.upper()})"
            return ai_res
        except Exception as e:
            fallback = _smart_local_generate(job_info, cv_text)
            fallback["used_provider"] = f"Smart-Template-Engine (Fallback wegen KI-Fehler: {str(e)[:80]})"
            return fallback

    return _smart_local_generate(job_info, cv_text)


def ai_map_docx_blocks(blocks: list, doc_type: str, generated_data: dict, config: dict) -> dict:
    """
    Falls ein .docx Muster-Layout KEINE {{PLATZHALTER}} verwendet, kann die KI
    anhand der extrahierten Textblöcke (`blocks`) entscheiden, welcher Block im Word-Dokument
    durch welchen neu generierten Text ersetzt werden soll, ohne Überschriften oder Design zu zerstören.
    """
    provider = (config.get("provider") or "smart_local").lower().strip()
    api_key = (config.get("api_key") or "").strip()

    if provider == "smart_local" or (not api_key and provider != "ollama"):
        return {}

    system_prompt = (
        "Du bist ein Experte für Word-Dokument-Strukturen (.docx). "
        "Dir wird eine Liste von Textblöcken (mit 'id' und 'text') aus einem Muster-Word-Layout übergeben, "
        "sowie die neuen Bewerbungsdaten eines Kandidaten. "
        "Deine Aufgabe: Identifiziere diejenigen Textblöcke im Muster-Layout, die Beispiel-/Blindtext des Vorbesitzers enthalten "
        "(z. B. alter Name, alte Kontaktdaten, alter Anschreibentext, alte Berufsstationen) und ordne ihnen den passenden neuen Text zu. "
        "WICHTIG: Lasse statische Sektionsüberschriften (wie 'BERUFSERFAHRUNG', 'AUSBILDUNG', 'KONTAKT', 'PROFIL', 'Mit freundlichen Grüßen') UNVERÄNDERT (nicht in das JSON aufnehmen)! "
        "Antworte AUSSCHLIESSLICH mit einem JSON-Objekt im Format: "
        '{"replacements": {"document.xml::p_0": "Neuer Text...", ...}}'
    )

    user_prompt = (
        f"DOKUMENT-TYP: {'Anschreiben' if doc_type == 'cover_letter' else 'Lebenslauf'}\n\n"
        f"EXTRAHIERTE TEXTBLÖCKE AUS DEM WORD-LAYOUT:\n"
        f"{json.dumps(blocks[:80], ensure_ascii=False, indent=2)}\n\n"
        f"NEUE BEWERBUNGSDATEN:\n"
        f"{json.dumps(generated_data, ensure_ascii=False, indent=2)}"
    )

    try:
        res = _call_llm_json(user_prompt, system_prompt, config)
        if isinstance(res, dict) and isinstance(res.get("replacements"), dict):
            return res["replacements"]
    except Exception:
        pass
    return {}
