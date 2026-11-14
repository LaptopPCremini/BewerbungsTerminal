// ==========================================================================
// DOCX ENGINE (100 % Client-Side JavaScript mit JSZip)
// Portiert aus docx_engine.py:
// - Beliebige .docx Muster-Layouts (Tabellen, Textboxen, Kopf-/Fußzeilen, Absätze)
// - Platzhalter-Ersetzung ({{NAME}}, {{ANSCHREIBEN_TEXT}}, etc.) über geteilte Runs
// - Heuristische & KI-gestützte Block-Ersetzung unter vollem Styling-Erhalt
// - Erzeugung der 5 vorinstallierten Muster-Word-Layouts (.docx)
// ==========================================================================

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const VML_NS = "urn:schemas-microsoft-com:vml";

const TARGET_XML_REGEX = /^word\/(document|header\d+|footer\d+)\.xml$/;

function getParagraphText(pElem) {
  let out = "";
  const walk = (node) => {
    if (!node) return;
    if (node.nodeType === 1) {
      if (node.namespaceURI === W_NS) {
        if (node.localName === "t") {
          out += node.textContent || "";
          return;
        }
        if (node.localName === "br") {
          out += "\n";
          return;
        }
        if (node.localName === "tab") {
          out += "\t";
          return;
        }
      }
      for (let i = 0; i < node.childNodes.length; i++) {
        walk(node.childNodes[i]);
      }
    }
  };
  walk(pElem);
  return out;
}

function setParagraphTextPreservingStyle(xmlDoc, pElem, newText) {
  // Finde rPr des ersten Runs zur Übernahme von Schriftart, Farbe, Größe, Fett/Kursiv
  let firstRpr = null;
  const children = Array.from(pElem.childNodes);
  for (const child of children) {
    if (child.nodeType === 1 && child.namespaceURI === W_NS && child.localName === "r") {
      const rpr = Array.from(child.childNodes).find(
        (n) => n.nodeType === 1 && n.namespaceURI === W_NS && n.localName === "rPr"
      );
      if (rpr) {
        firstRpr = rpr.cloneNode(true);
        break;
      }
    }
  }

  // Entferne bestehende Text-Runs / Hyperlinks / proofErr, aber bewahre pPr, Bookmarks & Drawings/Shapes
  for (const child of children) {
    if (child.nodeType === 1 && child.namespaceURI === W_NS) {
      if (child.localName === "r" || child.localName === "hyperlink" || child.localName === "proofErr") {
        const hasDrawing =
          child.getElementsByTagNameNS(W_NS, "drawing").length > 0 ||
          child.getElementsByTagNameNS(VML_NS, "shape").length > 0 ||
          child.getElementsByTagNameNS(W_NS, "pict").length > 0;
        if (!hasDrawing) {
          pElem.removeChild(child);
        }
      }
    }
  }

  // Neuen Run mit übernommenem Styling erstellen
  const newRun = xmlDoc.createElementNS(W_NS, "w:r");
  if (firstRpr) {
    newRun.appendChild(firstRpr);
  }

  const lines = String(newText ?? "").split("\n");
  lines.forEach((line, idx) => {
    if (idx > 0) {
      const br = xmlDoc.createElementNS(W_NS, "w:br");
      newRun.appendChild(br);
    }
    const t = xmlDoc.createElementNS(W_NS, "w:t");
    t.setAttributeNS(XML_NS, "xml:space", "preserve");
    t.textContent = line;
    newRun.appendChild(t);
  });

  pElem.appendChild(newRun);
}

function replacePlaceholdersInParagraph(xmlDoc, pElem, replacements) {
  const fullText = getParagraphText(pElem);
  if (!fullText) return false;

  let modified = false;
  let newText = fullText;
  for (const [key, val] of Object.entries(replacements)) {
    if (val !== undefined && val !== null && newText.includes(key)) {
      newText = newText.split(key).join(String(val));
      modified = true;
    }
  }

  if (modified) {
    setParagraphTextPreservingStyle(xmlDoc, pElem, newText);
  }
  return modified;
}

export async function inspectDocx(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser();
  const blocks = [];
  const placeholders = new Set();

  const fileNames = Object.keys(zip.files).filter((name) => TARGET_XML_REGEX.test(name));
  for (const name of fileNames) {
    const shortName = name.split("/").pop();
    const xmlStr = await zip.files[name].async("string");
    const xmlDoc = parser.parseFromString(xmlStr, "application/xml");
    const paragraphs = xmlDoc.getElementsByTagNameNS(W_NS, "p");

    for (let i = 0; i < paragraphs.length; i++) {
      const text = getParagraphText(paragraphs[i]).trim();
      if (text) {
        blocks.push({
          id: `${shortName}::p_${i}`,
          text,
          part: shortName
        });
        const matches = text.match(/\{\{[A-Za-z0-9_ÄÖÜäöüß\-]+\}\}/g);
        if (matches) {
          matches.forEach((m) => placeholders.add(m));
        }
      }
    }
  }

  const fullJoined = blocks.map((b) => b.text).join("\n").toLowerCase();
  let guessedType = "auto";
  if (["lebenslauf", "curriculum vitae", "berufserfahrung", "ausbildung", "kenntnisse", "{{erfahrung}}"].some((k) => fullJoined.includes(k))) {
    guessedType = "cv";
  }
  if (["sehr geehrte", "bewerbung als", "bewerbung um", "mit freundlichen grüßen", "{{anschreiben}}"].some((k) => fullJoined.includes(k))) {
    guessedType = guessedType !== "cv" ? "cover_letter" : "both";
  }

  return {
    blocks,
    placeholders: Array.from(placeholders).sort(),
    blockCount: blocks.length,
    guessedType,
    previewText: blocks.slice(0, 25).map((b) => b.text).join("\n")
  };
}

export async function mergeDocxTemplate(arrayBuffer, placeholdersMap = {}, blockReplacements = {}) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  let replacedBlocks = 0;
  let replacedPlaceholders = 0;

  const fileNames = Object.keys(zip.files).filter((name) => TARGET_XML_REGEX.test(name));
  for (const name of fileNames) {
    const shortName = name.split("/").pop();
    const xmlStr = await zip.files[name].async("string");
    const xmlDoc = parser.parseFromString(xmlStr, "application/xml");
    const paragraphs = xmlDoc.getElementsByTagNameNS(W_NS, "p");

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const blockId = `${shortName}::p_${i}`;

      if (Object.prototype.hasOwnProperty.call(blockReplacements, blockId) && blockReplacements[blockId] !== null && blockReplacements[blockId] !== undefined) {
        setParagraphTextPreservingStyle(xmlDoc, p, String(blockReplacements[blockId]));
        replacedBlocks++;
      }

      if (placeholdersMap && Object.keys(placeholdersMap).length > 0) {
        if (replacePlaceholdersInParagraph(xmlDoc, p, placeholdersMap)) {
          replacedPlaceholders++;
        }
      }
    }

    const updatedXml = serializer.serializeToString(xmlDoc);
    zip.file(name, updatedXml);
  }

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });

  return {
    blob,
    replacedBlocks,
    replacedPlaceholders
  };
}

export function autoMapBlocksHeuristic(blocks, docType, data) {
  const mapping = {};
  if (!blocks || blocks.length === 0) return mapping;

  if (docType === "cover_letter") {
    let anredeIdx = null;
    let grussIdx = null;
    let betreffIdx = null;

    blocks.forEach((b, i) => {
      const txtLow = b.text.toLowerCase().trim();
      if (anredeIdx === null && (txtLow.startsWith("sehr geehrt") || txtLow.startsWith("liebe") || txtLow.startsWith("guten tag"))) {
        anredeIdx = i;
      } else if (betreffIdx === null && (txtLow.includes("bewerbung") || txtLow.includes("stellenanzeige") || txtLow.includes("referenz")) && b.text.length < 140) {
        betreffIdx = i;
      } else if (txtLow.includes("freundlichen grüßen") || txtLow.includes("freundlichen grüssen") || txtLow.includes("beste grüße")) {
        grussIdx = i;
      }
    });

    if (betreffIdx !== null && data.betreff) mapping[blocks[betreffIdx].id] = data.betreff;
    if (anredeIdx !== null && data.anrede) mapping[blocks[anredeIdx].id] = data.anrede;

    if (anredeIdx !== null) {
      const endIdx = grussIdx !== null && grussIdx > anredeIdx ? grussIdx : blocks.length;
      const bodyBlocks = blocks.slice(anredeIdx + 1, endIdx);
      const paragraphs = (data.anschreiben_text || "").split("\n\n").map((p) => p.trim()).filter(Boolean);

      if (bodyBlocks.length > 0 && paragraphs.length > 0) {
        if (bodyBlocks.length === 1) {
          mapping[bodyBlocks[0].id] = paragraphs.join("\n\n");
        } else {
          bodyBlocks.forEach((bBlk, j) => {
            if (j < paragraphs.length - 1) {
              mapping[bBlk.id] = paragraphs[j];
            } else if (j === paragraphs.length - 1) {
              mapping[bBlk.id] = paragraphs.slice(j).join("\n\n");
            } else {
              mapping[bBlk.id] = "";
            }
          });
        }
      }
      if (grussIdx !== null && grussIdx + 1 < blocks.length && data.name) {
        mapping[blocks[grussIdx + 1].id] = data.name;
      }
    } else {
      const longest = blocks.reduce((a, b) => (b.text.length > a.text.length ? b : a), blocks[0]);
      mapping[longest.id] = data.anschreiben_text || "";
    }
  } else if (docType === "cv") {
    const sectionKeywords = {
      kurzprofil: ["kurzprofil", "profil", "über mich", "summary", "schwerpunkte"],
      berufserfahrung: ["berufserfahrung", "beruflicher werdegang", "erfahrung", "praxis", "work experience", "werdegang"],
      ausbildung: ["ausbildung", "studium", "bildungsweg", "education", "akademisch"],
      kenntnisse: ["kenntnisse", "kompetenzen", "skills", "fähigkeiten", "sprachen", "zertifikate", "it-kenntnisse"]
    };
    const usedIndices = new Set();

    blocks.forEach((b, i) => {
      const txtLow = b.text.toLowerCase().trim();
      if (txtLow.length <= 45) {
        for (const [fieldKey, kwList] of Object.entries(sectionKeywords)) {
          if (kwList.some((kw) => kw === txtLow || txtLow.startsWith(kw))) {
            if (i + 1 < blocks.length && !usedIndices.has(i + 1)) {
              usedIndices.add(i);
              usedIndices.add(i + 1);
              if (data[fieldKey]) {
                mapping[blocks[i + 1].id] = data[fieldKey];
              }
            }
            break;
          }
        }
      }
    });

    if (blocks[0].text.length < 50 && !usedIndices.has(0) && data.name) {
      mapping[blocks[0].id] = data.name;
    }
  }

  return mapping;
}

export function buildPlaceholderMap(genData) {
  const base = {
    "{{NAME}}": genData.name || "",
    "{{VORNAME_NACHNAME}}": genData.name || "",
    "{{KONTAKT}}": genData.kontakt || "",
    "{{ADRESSE}}": genData.kontakt || "",
    "{{EMPFAENGER}}": genData.empfaenger || "",
    "{{UNTERNEHMEN}}": genData.unternehmen || "",
    "{{FIRMA}}": genData.unternehmen || "",
    "{{DATUM}}": genData.datum || "",
    "{{ZIELPOSITION}}": genData.zielposition || "",
    "{{POSITION}}": genData.zielposition || "",
    "{{JOBTITEL}}": genData.zielposition || "",
    "{{BETREFF}}": genData.betreff || "",
    "{{ANREDE}}": genData.anrede || "",
    "{{ANSCHREIBEN}}": genData.anschreiben_text || "",
    "{{ANSCHREIBEN_TEXT}}": genData.anschreiben_text || "",
    "{{TEXT}}": genData.anschreiben_text || "",
    "{{KURZPROFIL}}": genData.kurzprofil || "",
    "{{PROFIL}}": genData.kurzprofil || "",
    "{{BERUFSERFAHRUNG}}": genData.berufserfahrung || "",
    "{{ERFAHRUNG}}": genData.berufserfahrung || "",
    "{{WERDEGANG}}": genData.berufserfahrung || "",
    "{{AUSBILDUNG}}": genData.ausbildung || "",
    "{{STUDIUM}}": genData.ausbildung || "",
    "{{BILDUNGSWEG}}": genData.ausbildung || "",
    "{{KENNTNISSE}}": genData.kenntnisse || "",
    "{{SKILLS}}": genData.kenntnisse || "",
    "{{KOMPETENZEN}}": genData.kenntnisse || "",
    "{{LEBENSLAUF}}": genData.lebenslauf_komplett || ""
  };

  const full = { ...base };
  for (const [k, v] of Object.entries(base)) {
    full[k.toLowerCase()] = v;
  }
  return full;
}

// --- GENERIERUNG DER 5 MUSTER-WORD-LAYOUTS (.DOCX) IM BROWSER ---
async function buildMinimalDocxBuffer(documentXml) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</w:styles>`);
  zip.file("word/document.xml", documentXml.trim());
  return await zip.generateAsync({ type: "arraybuffer" });
}

export async function generateDefaultSampleLayouts() {
  const samples = [
    {
      filename: "Muster_Anschreiben_Executive_Blau.docx",
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="18" w:space="6" w:color="1E3A8A"/></w:pBdr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="36"/></w:rPr><w:t>{{NAME}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="360"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="475569"/><w:sz w:val="19"/></w:rPr><w:t>{{KONTAKT}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="334155"/><w:sz w:val="21"/></w:rPr><w:t>{{EMPFAENGER}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="right"/><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="64748B"/><w:sz w:val="20"/></w:rPr><w:t>{{DATUM}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:before="120" w:after="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>{{BETREFF}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="180"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="21"/></w:rPr><w:t>{{ANREDE}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="both"/><w:spacing w:line="280" w:lineRule="auto" w:after="280"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{ANSCHREIBEN_TEXT}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:before="160" w:after="120"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>Mit freundlichen Grüßen</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:before="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="22"/></w:rPr><w:t>{{NAME}}</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/></w:sectPr>
  </w:body>
</w:document>`
    },
    {
      filename: "Muster_Anschreiben_Minimal_Anthrazit.docx",
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="111827"/><w:sz w:val="38"/></w:rPr><w:t>{{NAME}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="8" w:color="B45309"/></w:pBdr><w:spacing w:after="320"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="6B7280"/><w:sz w:val="19"/></w:rPr><w:t>{{KONTAKT}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="200"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="374151"/><w:sz w:val="21"/></w:rPr><w:t>{{EMPFAENGER}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="right"/><w:spacing w:after="200"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="6B7280"/><w:sz w:val="20"/></w:rPr><w:t>{{DATUM}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:before="120" w:after="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="25"/></w:rPr><w:t>{{BETREFF}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="160"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="111827"/><w:sz w:val="22"/></w:rPr><w:t>{{ANREDE}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="both"/><w:spacing w:line="276" w:lineRule="auto" w:after="260"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="22"/></w:rPr><w:t>{{ANSCHREIBEN_TEXT}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:before="160" w:after="120"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="22"/></w:rPr><w:t>Mit freundlichen Grüßen</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:before="200"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="23"/></w:rPr><w:t>{{NAME}}</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/></w:sectPr>
  </w:body>
</w:document>`
    },
    {
      filename: "Muster_Lebenslauf_Executive_Blau.docx",
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="1E3A8A"/><w:spacing w:before="120" w:after="60"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="38"/></w:rPr><w:t>  {{NAME}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="1E3A8A"/><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="DBEAFE"/><w:sz w:val="22"/></w:rPr><w:t>  {{ZIELPOSITION}}  |  {{KONTAKT}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1E3A8A"/></w:pBdr><w:spacing w:before="180" w:after="100"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>KURZPROFIL</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="220" w:line="260" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{KURZPROFIL}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1E3A8A"/></w:pBdr><w:spacing w:before="180" w:after="100"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>BERUFSERFAHRUNG</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="220" w:line="260" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{BERUFSERFAHRUNG}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1E3A8A"/></w:pBdr><w:spacing w:before="180" w:after="100"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>AUSBILDUNG &amp; STUDIUM</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="220" w:line="260" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{AUSBILDUNG}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1E3A8A"/></w:pBdr><w:spacing w:before="180" w:after="100"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="24"/></w:rPr><w:t>KENNTNISSE &amp; KOMPETENZEN</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="180" w:line="260" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="21"/></w:rPr><w:t>{{KENNTNISSE}}</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1200" w:bottom="1134" w:left="1200"/></w:sectPr>
  </w:body>
</w:document>`
    },
    {
      filename: "Muster_Lebenslauf_Minimal_Anthrazit.docx",
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="111827"/><w:sz w:val="40"/></w:rPr><w:t>{{NAME}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>{{ZIELPOSITION}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="8" w:color="B45309"/></w:pBdr><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="4B5563"/><w:sz w:val="19"/></w:rPr><w:t>{{KONTAKT}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="D1D5DB"/></w:pBdr><w:spacing w:before="140" w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>PROFIL &amp; SCHWERPUNKTE</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="200" w:line="260" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr><w:t>{{KURZPROFIL}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="D1D5DB"/></w:pBdr><w:spacing w:before="140" w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>BERUFLICHER WERDEGANG</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="200" w:line="260" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr><w:t>{{BERUFSERFAHRUNG}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="D1D5DB"/></w:pBdr><w:spacing w:before="140" w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>BILDUNGSWEG</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="200" w:line="260" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr><w:t>{{AUSBILDUNG}}</w:t></w:r></w:p>
    <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="D1D5DB"/></w:pBdr><w:spacing w:before="140" w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="B45309"/><w:sz w:val="24"/></w:rPr><w:t>SKILLS, SPRACHEN &amp; ZERTIFIKATE</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="160" w:line="260" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr><w:t>{{KENNTNISSE}}</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1200" w:bottom="1134" w:left="1200"/></w:sectPr>
  </w:body>
</w:document>`
    },
    {
      filename: "Muster_Lebenslauf_Zweispaltig_Sidebar.docx",
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="10000" w:type="dxa"/>
        <w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="3200"/><w:gridCol w:w="6800"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="0F172A"/><w:tcMar><w:top w:w="280" w:type="dxa"/><w:left w:w="240" w:type="dxa"/><w:bottom w:w="280" w:type="dxa"/><w:right w:w="240" w:type="dxa"/></w:tcMar></w:tcPr>
          <w:p><w:pPr><w:spacing w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="38BDF8"/><w:sz w:val="30"/></w:rPr><w:t>{{NAME}}</w:t></w:r></w:p>
          <w:p><w:pPr><w:spacing w:after="280"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="E2E8F0"/><w:sz w:val="20"/></w:rPr><w:t>{{ZIELPOSITION}}</w:t></w:r></w:p>
          <w:p><w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="38BDF8"/><w:sz w:val="21"/></w:rPr><w:t>KONTAKT</w:t></w:r></w:p>
          <w:p><w:pPr><w:spacing w:after="260"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="CBD5E1"/><w:sz w:val="18"/></w:rPr><w:t>{{KONTAKT}}</w:t></w:r></w:p>
          <w:p><w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="38BDF8"/><w:sz w:val="21"/></w:rPr><w:t>KOMPETENZEN &amp; SKILLS</w:t></w:r></w:p>
          <w:p><w:pPr><w:spacing w:after="200"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="CBD5E1"/><w:sz w:val="19"/></w:rPr><w:t>{{KENNTNISSE}}</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="6800" w:type="dxa"/><w:tcMar><w:top w:w="280" w:type="dxa"/><w:left w:w="320" w:type="dxa"/><w:bottom w:w="280" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tcMar></w:tcPr>
          <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="10" w:space="4" w:color="0284C7"/></w:pBdr><w:spacing w:after="100"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="24"/></w:rPr><w:t>PROFIL</w:t></w:r></w:p>
          <w:p><w:pPr><w:spacing w:after="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="334155"/><w:sz w:val="20"/></w:rPr><w:t>{{KURZPROFIL}}</w:t></w:r></w:p>
          <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="10" w:space="4" w:color="0284C7"/></w:pBdr><w:spacing w:before="140" w:after="100"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="24"/></w:rPr><w:t>BERUFSERFAHRUNG</w:t></w:r></w:p>
          <w:p><w:pPr><w:spacing w:after="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="20"/></w:rPr><w:t>{{BERUFSERFAHRUNG}}</w:t></w:r></w:p>
          <w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="10" w:space="4" w:color="0284C7"/></w:pBdr><w:spacing w:before="140" w:after="100"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="24"/></w:rPr><w:t>AUSBILDUNG</w:t></w:r></w:p>
          <w:p><w:pPr><w:spacing w:after="180"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:color w:val="1E293B"/><w:sz w:val="20"/></w:rPr><w:t>{{AUSBILDUNG}}</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="850"/></w:sectPr>
  </w:body>
</w:document>`
    }
  ];

  const results = [];
  for (const s of samples) {
    const buffer = await buildMinimalDocxBuffer(s.xml);
    const info = await inspectDocx(buffer);
    results.push({
      filename: s.filename,
      buffer,
      sizeKb: Math.round((buffer.byteLength / 1024) * 10) / 10,
      modified: new Date().toLocaleString("de-DE"),
      placeholders: info.placeholders,
      blockCount: info.blockCount,
      guessedType: info.guessedType
    });
  }
  return results;
}
