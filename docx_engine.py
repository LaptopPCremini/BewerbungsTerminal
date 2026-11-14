# -*- coding: utf-8 -*-
"""
DOCX Engine für BewerbungsTerminal (100% Portable, Zero-Dependency XML Engine)
Unterstützt:
1. Beliebige .docx Muster-Layouts (Tabellen, Textboxen/Shapes, Kopf-/Fußzeilen, Absätze)
2. Platzhalter-Ersetzung (z. B. {{ANSCHREIBEN}}, {{NAME}}) auch über geteilte Word-Runs hinweg
3. Intelligente Block-Ersetzung (In-Place Text-Austausch unter vollständigem Erhalt von Schriftart, Größe, Farbe & Layout)
4. Generierung von designten Muster-Word-Layouts (.docx) zum direkten Start
"""

import os
import re
import zipfile
import shutil
import xml.etree.ElementTree as ET
from io import BytesIO

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("v", "urn:schemas-microsoft-com:vml")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def get_paragraph_text(p_elem: ET.Element) -> str:
    """Liest den gesamten sichtbaren Text eines <w:p> Elements inkl. Zeilenumbrüchen."""
    parts = []
    for elem in p_elem.iter():
        if elem.tag == _w("t") and elem.text:
            parts.append(elem.text)
        elif elem.tag == _w("br"):
            parts.append("\n")
        elif elem.tag == _w("tab"):
            parts.append("\t")
    return "".join(parts)


def set_paragraph_text_preserving_style(p_elem: ET.Element, new_text: str):
    """
    Ersetzt den Text eines <w:p>-Elements, behält aber <w:pPr> (Absatzformatierung)
    sowie das <w:rPr> (Schriftart, Farbe, Größe, Fett/Kursiv) des ersten Text-Runs bei.
    Unterstützt Zeilenumbrüche (\n) als saubere <w:br/> innerhalb des Absatzes.
    """
    runs = p_elem.findall(_w("r"))
    first_rpr = None
    for r in runs:
        rpr = r.find(_w("rPr"))
        if rpr is not None:
            first_rpr = ET.fromstring(ET.tostring(rpr))
            break

    # Entferne bestehende Runs und Hyperlinks mit Text, behalte aber pPr, Bookmarks oder Shapes/Drawings
    to_remove = []
    for child in list(p_elem):
        if child.tag in (_w("r"), _w("hyperlink"), _w("proofErr")):
            # Prüfen ob in diesem Run ein Drawing/Textbox steckt – falls ja, nicht löschen!
            has_drawing = (
                child.find(f".//{{{W_NS}}}drawing") is not None
                or child.find(".//{urn:schemas-microsoft-com:vml}shape") is not None
                or child.find(f".//{{{W_NS}}}pict") is not None
            )
            if not has_drawing:
                to_remove.append(child)

    for child in to_remove:
        p_elem.remove(child)

    # Neuen Run mit übernommenem Styling aufbauen
    new_run = ET.Element(_w("r"))
    if first_rpr is not None:
        new_run.append(first_rpr)

    lines = str(new_text).split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            br = ET.Element(_w("br"))
            new_run.append(br)
        t = ET.Element(_w("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = line
        new_run.append(t)

    p_elem.append(new_run)


def replace_placeholders_in_paragraph(p_elem: ET.Element, replacements: dict) -> bool:
    """
    Ersetzt Platzhalter wie {{NAME}} in einem Absatz, selbst wenn Word sie über mehrere <w:r> Runs
    zerrissen hat. Falls nur ein Teil des Absatzes ein Platzhalter ist, bleiben die anderen Runs
    oder der Gesamttext erhalten.
    """
    full_text = get_paragraph_text(p_elem)
    if not full_text:
        return False

    modified = False
    new_text = full_text
    for key, val in replacements.items():
        if key in new_text and val is not None:
            new_text = new_text.replace(key, str(val))
            modified = True

    if modified:
        set_paragraph_text_preserving_style(p_elem, new_text)
    return modified


def inspect_docx(docx_path: str) -> dict:
    """
    Analysiert eine .docx-Datei und extrahiert:
    - Alle Textblöcke mit eindeutiger Block-ID (z. B. "document.xml::p_0", "header1.xml::p_2")
    - Alle gefundenen {{PLATZHALTER}}
    - Dokumenttyp-Einschätzung (Anschreiben vs. Lebenslauf vs. Universal)
    """
    blocks = []
    placeholders = set()
    target_xml_pattern = re.compile(r"^word/(document|header\d+|footer\d+)\.xml$")

    with zipfile.ZipFile(docx_path, "r") as zf:
        for name in zf.namelist():
            if target_xml_pattern.match(name):
                short_name = name.split("/")[-1]
                xml_bytes = zf.read(name)
                root = ET.fromstring(xml_bytes)
                p_index = 0
                for p in root.iter(_w("p")):
                    text = get_paragraph_text(p).strip()
                    if text:
                        block_id = f"{short_name}::p_{p_index}"
                        blocks.append({
                            "id": block_id,
                            "text": text,
                            "part": short_name
                        })
                        for match in re.findall(r"\{\{[A-Za-z0-9_ÄÖÜäöüß\-]+\}\}", text):
                            placeholders.add(match)
                    p_index += 1

    full_joined = "\n".join(b["text"] for b in blocks).lower()
    guessed_type = "auto"
    if any(k in full_joined for k in ["lebenslauf", "curriculum vitae", "berufserfahrung", "ausbildung", "kenntnisse", "{{erfahrung}}"]):
        guessed_type = "cv"
    if any(k in full_joined for k in ["sehr geehrte", "bewerbung als", "bewerbung um", "mit freundlichen grüßen", "{{anschreiben}}"]):
        guessed_type = "cover_letter" if guessed_type != "cv" else "both"

    return {
        "blocks": blocks,
        "placeholders": sorted(list(placeholders)),
        "block_count": len(blocks),
        "guessed_type": guessed_type,
        "preview_text": "\n".join(b["text"] for b in blocks[:25])
    }


def merge_docx_template(
    template_path: str,
    output_path: str,
    placeholders_map: dict = None,
    block_replacements: dict = None
) -> dict:
    """
    Erzeugt eine neue .docx Datei basierend auf `template_path` und wendet sowohl:
    1. `placeholders_map` (z. B. {"{{ANSCHREIBEN}}": "...", "{{NAME}}": "..."})
    2. `block_replacements` (z. B. {"document.xml::p_12": "Neuer Text..."})
    an. Alle XML-Styles, Tabellen, Grafiken, Rahmen und Schriftarten bleiben 100% erhalten.
    """
    placeholders_map = placeholders_map or {}
    block_replacements = block_replacements or {}
    target_xml_pattern = re.compile(r"^word/(document|header\d+|footer\d+)\.xml$")

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    replaced_blocks_count = 0
    replaced_placeholders_count = 0

    with zipfile.ZipFile(template_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if target_xml_pattern.match(item.filename):
                    short_name = item.filename.split("/")[-1]
                    root = ET.fromstring(data)
                    p_index = 0
                    for p in root.iter(_w("p")):
                        block_id = f"{short_name}::p_{p_index}"
                        # 1. Direkte Block-ID Ersetzung (von KI oder direktem Mapping)
                        if block_id in block_replacements and block_replacements[block_id] is not None:
                            set_paragraph_text_preserving_style(p, str(block_replacements[block_id]))
                            replaced_blocks_count += 1
                        # 2. Platzhalter-Ersetzung (falls {{...}} enthalten)
                        if placeholders_map:
                            if replace_placeholders_in_paragraph(p, placeholders_map):
                                replaced_placeholders_count += 1
                        p_index += 1
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                zout.writestr(item, data)

    return {
        "output_path": output_path,
        "replaced_blocks": replaced_blocks_count,
        "replaced_placeholders": replaced_placeholders_count
    }


def auto_map_blocks_heuristic(blocks: list, doc_type: str, data: dict) -> dict:
    """
    Intelligente Heuristik-Zuordnung für .docx Muster-Layouts OHNE {{PLATZHALTER}}.
    Erkennt anhand der Struktur und Schlüsselwörter im bestehenden Word-Dokument
    Anschreiben-Elemente (Betreff, Anrede, Haupttext) bzw. Lebenslauf-Sektionen
    (Profil, Berufserfahrung, Ausbildung, Kenntnisse) und ersetzt die Textblöcke
    direkt unter Beibehaltung des Word-Stylings.
    """
    mapping = {}
    if not blocks:
        return mapping

    if doc_type == "cover_letter":
        anrede_idx = None
        gruss_idx = None
        betreff_idx = None

        for i, b in enumerate(blocks):
            txt_low = b["text"].lower().strip()
            if anrede_idx is None and (txt_low.startswith("sehr geehrt") or txt_low.startswith("liebe") or txt_low.startswith("guten tag")):
                anrede_idx = i
            elif betreff_idx is None and ("bewerbung" in txt_low or "stellenanzeige" in txt_low or "referenz" in txt_low) and len(b["text"]) < 140:
                betreff_idx = i
            elif "freundlichen grüßen" in txt_low or "freundlichen grüssen" in txt_low or "beste grüße" in txt_low:
                gruss_idx = i

        if betreff_idx is not None and data.get("betreff"):
            mapping[blocks[betreff_idx]["id"]] = data["betreff"]
        if anrede_idx is not None and data.get("anrede"):
            mapping[blocks[anrede_idx]["id"]] = data["anrede"]

        # Haupttext zwischen Anrede und Grußformel
        if anrede_idx is not None:
            end_idx = gruss_idx if gruss_idx is not None and gruss_idx > anrede_idx else len(blocks)
            body_blocks = blocks[anrede_idx + 1 : end_idx]
            paragraphs = [p.strip() for p in data.get("anschreiben_text", "").split("\n\n") if p.strip()]
            if body_blocks and paragraphs:
                if len(body_blocks) == 1:
                    mapping[body_blocks[0]["id"]] = "\n\n".join(paragraphs)
                else:
                    for j, b_blk in enumerate(body_blocks):
                        if j < len(paragraphs) - 1:
                            mapping[b_blk["id"]] = paragraphs[j]
                        elif j == len(paragraphs) - 1:
                            mapping[b_blk["id"]] = "\n\n".join(paragraphs[j:])
                        else:
                            mapping[b_blk["id"]] = ""
            if gruss_idx is not None and gruss_idx + 1 < len(blocks) and data.get("name"):
                mapping[blocks[gruss_idx + 1]["id"]] = data["name"]
        else:
            # Fallback: Längsten Textblock im Dokument durch Anschreiben-Text ersetzen
            longest = max(blocks, key=lambda x: len(x["text"]))
            mapping[longest["id"]] = data.get("anschreiben_text", "")

    elif doc_type == "cv":
        # Suche nach Überschriften und ersetze jeweils den darauffolgenden Inhaltsblock
        section_keywords = {
            "kurzprofil": ["kurzprofil", "profil", "über mich", "summary", "schwerpunkte"],
            "berufserfahrung": ["berufserfahrung", "beruflicher werdegang", "erfahrung", "praxis", "work experience", "werdegang"],
            "ausbildung": ["ausbildung", "studium", "bildungsweg", "education", "akademisch"],
            "kenntnisse": ["kenntnisse", "kompetenzen", "skills", "fähigkeiten", "sprachen", "zertifikate", "it-kenntnisse"]
        }
        used_indices = set()
        for i, b in enumerate(blocks):
            txt_low = b["text"].lower().strip()
            if len(txt_low) <= 45:
                for field_key, kw_list in section_keywords.items():
                    if any(kw == txt_low or txt_low.startswith(kw) for kw in kw_list):
                        if i + 1 < len(blocks) and (i + 1) not in used_indices:
                            used_indices.add(i)
                            used_indices.add(i + 1)
                            if data.get(field_key):
                                mapping[blocks[i + 1]["id"]] = data[field_key]
                        break

        # Erster kurzer Block oft Name
        if len(blocks[0]["text"]) < 50 and 0 not in used_indices and data.get("name"):
            mapping[blocks[0]["id"]] = data["name"]

    return mapping



def _build_minimal_docx(document_xml: str, styles_xml: str = None) -> bytes:
    """Erstellt ein valides, gestyltes .docx Archiv im Speicher."""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    if not styles_xml:
        styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="de-DE"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>"""

    bio = BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types.strip())
        zf.writestr("_rels/.rels", rels.strip())
        zf.writestr("word/_rels/document.xml.rels", doc_rels.strip())
        zf.writestr("word/styles.xml", styles_xml.strip())
        zf.writestr("word/document.xml", document_xml.strip())
    return bio.getvalue()


def create_sample_layouts(layouts_dir: str):
    """
    Erstellt beim ersten Start 4 hochwertige, designte Muster-Word-Layouts (.docx)
    im Ordner `layouts/`, damit sofort Vorlagen zum Testen bereitstehen.
    """
    os.makedirs(layouts_dir, exist_ok=True)

    # 1. Anschreiben - Modern Executive (Dunkelblau / Akzentlinie)
    anschreiben_modern_path = os.path.join(layouts_dir, "Muster_Anschreiben_Executive_Blau.docx")
    if not os.path.exists(anschreiben_modern_path):
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="18" w:space="6" w:color="1E3A8A"/></w:pBdr><w:spacing w:after="120"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="36"/></w:rPr><w:t>{{NAME}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="360"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="475569"/><w:sz w:val="19"/></w:rPr><w:t>{{KONTAKT}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="240"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="334155"/><w:sz w:val="21"/></w:rPr><w:t>{{EMPFAENGER}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="right"/><w:spacing w:after="240"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="64748B"/><w:sz w:val="20"/></w:rPr><w:t>{{DATUM}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="120" w:after="240"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>{{BETREFF}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="180"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="21"/></w:rPr><w:t>{{ANREDE}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="both"/><w:spacing w:line="280" w:lineRule="auto" w:after="280"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{ANSCHREIBEN_TEXT}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="160" w:after="120"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>Mit freundlichen Grüßen</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="240"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="22"/></w:rPr><w:t>{{NAME}}</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/>
    </w:sectPr>
  </w:body>
</w:document>"""
        with open(anschreiben_modern_path, "wb") as f:
            f.write(_build_minimal_docx(doc_xml))

    # 2. Anschreiben - Minimalist Anthrazit & Gold
    anschreiben_minimal_path = os.path.join(layouts_dir, "Muster_Anschreiben_Minimal_Anthrazit.docx")
    if not os.path.exists(anschreiben_minimal_path):
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="111827"/><w:sz w:val="38"/></w:rPr><w:t>{{NAME}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="8" w:color="B45309"/></w:pBdr><w:spacing w:after="320"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="6B7280"/><w:sz w:val="19"/></w:rPr><w:t>{{KONTAKT}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="200"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="374151"/><w:sz w:val="21"/></w:rPr><w:t>{{EMPFAENGER}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="right"/><w:spacing w:after="200"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="6B7280"/><w:sz w:val="20"/></w:rPr><w:t>{{DATUM}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="120" w:after="220"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="25"/></w:rPr><w:t>{{BETREFF}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="160"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="111827"/><w:sz w:val="22"/></w:rPr><w:t>{{ANREDE}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="both"/><w:spacing w:line="276" w:lineRule="auto" w:after="260"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="22"/></w:rPr><w:t>{{ANSCHREIBEN_TEXT}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="160" w:after="120"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="22"/></w:rPr><w:t>Mit freundlichen Grüßen</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="200"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="23"/></w:rPr><w:t>{{NAME}}</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/>
    </w:sectPr>
  </w:body>
</w:document>"""
        with open(anschreiben_minimal_path, "wb") as f:
            f.write(_build_minimal_docx(doc_xml))

    # 3. Lebenslauf - Modern Executive 2-Spalten-Stil (Blau)
    cv_modern_path = os.path.join(layouts_dir, "Muster_Lebenslauf_Executive_Blau.docx")
    if not os.path.exists(cv_modern_path):
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:shd w:val="clear" w:color="auto" w:fill="1E3A8A"/><w:spacing w:before="120" w:after="60"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="38"/></w:rPr><w:t>  {{NAME}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:shd w:val="clear" w:color="auto" w:fill="1E3A8A"/><w:spacing w:after="240"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="DBEAFE"/><w:sz w:val="22"/></w:rPr><w:t>  {{ZIELPOSITION}}  |  {{KONTAKT}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1E3A8A"/></w:pBdr><w:spacing w:before="180" w:after="100"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>KURZPROFIL</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="220" w:line="260" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{KURZPROFIL}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1E3A8A"/></w:pBdr><w:spacing w:before="180" w:after="100"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>BERUFSERFAHRUNG</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="220" w:line="260" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{BERUFSERFAHRUNG}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1E3A8A"/></w:pBdr><w:spacing w:before="180" w:after="100"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>AUSBILDUNG &amp; STUDIUM</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="220" w:line="260" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{AUSBILDUNG}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1E3A8A"/></w:pBdr><w:spacing w:before="180" w:after="100"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>KENNTNISSE &amp; KOMPETENZEN</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="180" w:line="260" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{KENNTNISSE}}</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1200" w:bottom="1134" w:left="1200"/>
    </w:sectPr>
  </w:body>
</w:document>"""
        with open(cv_modern_path, "wb") as f:
            f.write(_build_minimal_docx(doc_xml))

    # 4. Lebenslauf - Minimalist Anthrazit & Gold
    cv_minimal_path = os.path.join(layouts_dir, "Muster_Lebenslauf_Minimal_Anthrazit.docx")
    if not os.path.exists(cv_minimal_path):
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:after="40"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="111827"/><w:sz w:val="40"/></w:rPr><w:t>{{NAME}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>{{ZIELPOSITION}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="8" w:color="B45309"/></w:pBdr><w:spacing w:after="240"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="4B5563"/><w:sz w:val="19"/></w:rPr><w:t>{{KONTAKT}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="D1D5DB"/></w:pBdr><w:spacing w:before="140" w:after="80"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>PROFIL &amp; SCHWERPUNKTE</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="200" w:line="260" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr><w:t>{{KURZPROFIL}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="D1D5DB"/></w:pBdr><w:spacing w:before="140" w:after="80"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>BERUFLICHER WERDEGANG</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="200" w:line="260" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr><w:t>{{BERUFSERFAHRUNG}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="D1D5DB"/></w:pBdr><w:spacing w:before="140" w:after="80"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>BILDUNGSWEG</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="200" w:line="260" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr><w:t>{{AUSBILDUNG}}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="D1D5DB"/></w:pBdr><w:spacing w:before="140" w:after="80"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>SKILLS, SPRACHEN &amp; ZERTIFIKATE</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:after="160" w:line="260" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr><w:t>{{KENNTNISSE}}</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1200" w:bottom="1134" w:left="1200"/>
    </w:sectPr>
  </w:body>
</w:document>"""
        with open(cv_minimal_path, "wb") as f:
            f.write(_build_minimal_docx(doc_xml))

    # 5. Lebenslauf - 2-Spalten Sidebar Design (Dunkle linke Spalte + Hauptspalte rechts)
    cv_sidebar_path = os.path.join(layouts_dir, "Muster_Lebenslauf_Zweispaltig_Sidebar.docx")
    if not os.path.exists(cv_sidebar_path):
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="10000" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
          <w:insideH w:val="none"/><w:insideV w:val="none"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="3200"/>
        <w:gridCol w:w="6800"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="3200" w:type="dxa"/>
            <w:shd w:val="clear" w:color="auto" w:fill="0F172A"/>
            <w:tcMar>
              <w:top w:w="280" w:type="dxa"/><w:left w:w="240" w:type="dxa"/>
              <w:bottom w:w="280" w:type="dxa"/><w:right w:w="240" w:type="dxa"/>
            </w:tcMar>
          </w:tcPr>
          <w:p>
            <w:pPr><w:spacing w:after="80"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="38BDF8"/><w:sz w:val="30"/></w:rPr><w:t>{{NAME}}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:after="280"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="E2E8F0"/><w:sz w:val="20"/></w:rPr><w:t>{{ZIELPOSITION}}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="38BDF8"/><w:sz w:val="21"/></w:rPr><w:t>KONTAKT</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:after="260"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="CBD5E1"/><w:sz w:val="18"/></w:rPr><w:t>{{KONTAKT}}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="38BDF8"/><w:sz w:val="21"/></w:rPr><w:t>KOMPETENZEN &amp; SKILLS</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:after="200"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="CBD5E1"/><w:sz w:val="19"/></w:rPr><w:t>{{KENNTNISSE}}</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="6800" w:type="dxa"/>
            <w:tcMar>
              <w:top w:w="280" w:type="dxa"/><w:left w:w="320" w:type="dxa"/>
              <w:bottom w:w="280" w:type="dxa"/><w:right w:w="200" w:type="dxa"/>
            </w:tcMar>
          </w:tcPr>
          <w:p>
            <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="10" w:space="4" w:color="0284C7"/></w:pBdr><w:spacing w:after="100"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="24"/></w:rPr><w:t>PROFIL</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:after="220"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="334155"/><w:sz w:val="20"/></w:rPr><w:t>{{KURZPROFIL}}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="10" w:space="4" w:color="0284C7"/></w:pBdr><w:spacing w:before="140" w:after="100"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="24"/></w:rPr><w:t>BERUFSERFAHRUNG</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:after="220"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="20"/></w:rPr><w:t>{{BERUFSERFAHRUNG}}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="10" w:space="4" w:color="0284C7"/></w:pBdr><w:spacing w:before="140" w:after="100"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="24"/></w:rPr><w:t>AUSBILDUNG</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:after="180"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="20"/></w:rPr><w:t>{{AUSBILDUNG}}</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="850"/>
    </w:sectPr>
  </w:body>
</w:document>"""
        with open(cv_sidebar_path, "wb") as f:
            f.write(_build_minimal_docx(doc_xml))

