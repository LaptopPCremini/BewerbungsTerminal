// ==========================================================================
// AI & JOB-SCRAPING ENGINE (100 % Client-Side Browser Fetch)
// Portiert aus ai_engine.py:
// - Job-URL Scraping mit optionalem CORS-Proxy + Fallback
// - JSON-LD JobPosting Extraktion + HTML-Textbereinigung
// - Smart-Lokal Generator (funktioniert sofort ohne API-Key)
// - Direktes Browser-Fetch für Gemini, OpenAI, Anthropic Claude & Ollama
// ==========================================================================

function extractTextFromHtml(htmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlString, "text/html");
  const title = (doc.querySelector("title")?.textContent || "").trim();

  // 1. JSON-LD JobPosting extrahieren
  let jobTitle = "";
  let companyName = "";
  let ldDescription = "";

  const ldScripts = doc.querySelectorAll('script[type="application/ld+json"]');
  ldScripts.forEach((sc) => {
    try {
      const data = JSON.parse(sc.textContent.trim());
      let items = Array.isArray(data) ? data : [data];
      items.forEach((item) => {
        if (item && Array.isArray(item["@graph"])) {
          items = items.concat(item["@graph"]);
        }
      });
      for (const item of items) {
        if (item && String(item["@type"] || "").includes("JobPosting")) {
          jobTitle = item.title || jobTitle;
          if (item.hiringOrganization && typeof item.hiringOrganization === "object") {
            companyName = item.hiringOrganization.name || companyName;
          }
          if (item.description) {
            const descDoc = parser.parseFromString(item.description, "text/html");
            ldDescription = (descDoc.body?.innerText || descDoc.body?.textContent || "").trim();
          }
        }
      }
    } catch (_) {}
  });

  // 2. HTML-Rauschen entfernen
  const removeSelectors = ["script", "style", "noscript", "svg", "footer", "nav", "header", "form"];
  removeSelectors.forEach((sel) => {
    doc.querySelectorAll(sel).forEach((el) => el.remove());
  });

  const rawBodyText = (doc.body?.innerText || doc.body?.textContent || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .filter(Boolean)
    .join("\n");

  const finalText = ldDescription.length > 200 ? ldDescription : rawBodyText;
  return {
    title: (jobTitle || title).slice(0, 140),
    company: companyName.slice(0, 100),
    text: finalText.slice(0, 12000)
  };
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 12000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, { ...options, signal: controller.signal });
    return res;
  } finally {
    clearTimeout(timer);
  }
}

export async function scrapeJobUrl(urlOrText, useCorsProxy = true) {
  const raw = (urlOrText || "").trim();
  if (!raw) {
    return { ok: false, error: "Kein Link oder Text angegeben." };
  }

  // Direkt eingefügter Stellentext statt URL
  if (!raw.toLowerCase().startsWith("http://") && !raw.toLowerCase().startsWith("https://")) {
    const firstLine = (raw.split(/\r?\n/)[0] || "Manuelle Stellenbeschreibung").slice(0, 80);
    return {
      ok: true,
      url: "Manueller Text-Input",
      title: firstLine,
      company: "",
      text: raw
    };
  }

  const candidates = [];
  if (useCorsProxy) {
    candidates.push(`https://api.allorigins.win/raw?url=${encodeURIComponent(raw)}`);
    candidates.push(`https://corsproxy.io/?${encodeURIComponent(raw)}`);
  }
  candidates.push(raw);

  let lastError = "";
  for (const targetUrl of candidates) {
    try {
      const resp = await fetchWithTimeout(targetUrl, { method: "GET" }, 12000);
      if (!resp.ok) {
        lastError = `HTTP ${resp.status}`;
        continue;
      }
      const htmlStr = await resp.text();
      if (htmlStr && htmlStr.length > 150) {
        const parsed = extractTextFromHtml(htmlStr);
        if (parsed.text && parsed.text.length > 60) {
          return {
            ok: true,
            url: raw,
            title: parsed.title || raw,
            company: parsed.company,
            text: parsed.text
          };
        }
      }
    } catch (err) {
      lastError = err.message || "CORS / Netzwerkfehler";
    }
  }

  return {
    ok: false,
    url: raw,
    title: "Link konnte nicht direkt geladen werden",
    company: "",
    text: "",
    error: `Webseite blockiert direkten Abruf (${lastError}). Tipp: Kopiere den Stellentext kurz über das 📝-Feld bei diesem Job.`
  };
}

async function callLlmJson(prompt, systemPrompt, config) {
  const provider = (config.provider || "smart_local").toLowerCase().trim();
  const apiKey = (config.api_key || "").trim();
  const model = (config.model || "").trim();
  const ollamaUrl = (config.ollama_url || "http://localhost:11434").replace(/\/+$/, "");

  let rawResponseText = "";

  if (provider === "gemini" && apiKey) {
    const modelName = model || "gemini-2.5-flash";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(modelName)}:generateContent?key=${encodeURIComponent(apiKey)}`;
    const resp = await fetchWithTimeout(
      url,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: systemPrompt }] },
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            temperature: 0.4,
            responseMimeType: "application/json"
          }
        })
      },
      45000
    );
    if (!resp.ok) throw new Error(`Gemini HTTP ${resp.status}`);
    const data = await resp.json();
    rawResponseText = data.candidates[0].content.parts[0].text;
  } else if (provider === "openai" && apiKey) {
    const modelName = model || "gpt-4o-mini";
    const resp = await fetchWithTimeout(
      "https://api.openai.com/v1/chat/completions",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`
        },
        body: JSON.stringify({
          model: modelName,
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: prompt }
          ],
          temperature: 0.4,
          response_format: { type: "json_object" }
        })
      },
      45000
    );
    if (!resp.ok) throw new Error(`OpenAI HTTP ${resp.status}`);
    const data = await resp.json();
    rawResponseText = data.choices[0].message.content;
  } else if (provider === "anthropic" && apiKey) {
    const modelName = model || "claude-3-5-sonnet-latest";
    const resp = await fetchWithTimeout(
      "https://api.anthropic.com/v1/messages",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true"
        },
        body: JSON.stringify({
          model: modelName,
          max_tokens: 4096,
          system: `${systemPrompt}\nAntworte AUSSCHLIESSLICH mit validem JSON ohne Markdown-Codeblöcke.`,
          messages: [{ role: "user", content: prompt }]
        })
      },
      45000
    );
    if (!resp.ok) throw new Error(`Claude HTTP ${resp.status}`);
    const data = await resp.json();
    rawResponseText = data.content[0].text;
  } else if (provider === "ollama") {
    const modelName = model || "llama3.1";
    const resp = await fetchWithTimeout(
      `${ollamaUrl}/api/chat`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: modelName,
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: prompt }
          ],
          format: "json",
          stream: false
        })
      },
      60000
    );
    if (!resp.ok) throw new Error(`Ollama HTTP ${resp.status}`);
    const data = await resp.json();
    rawResponseText = data.message.content;
  } else {
    throw new Error("Kein externer KI-Provider aktiv.");
  }

  const cleaned = rawResponseText
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "");
  return JSON.parse(cleaned);
}

function formatTodayDE() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

export function smartLocalGenerate(jobInfo, cvText) {
  const todayStr = formatTodayDE();
  const cvLines = (cvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  let name = "Max Mustermann";
  if (cvLines.length > 0) {
    const first = cvLines[0].replace(/^(lebenslauf|curriculum vitae|profil|name\s*:)\s*/i, "").trim();
    if (first.length >= 2 && first.length <= 55) name = first;
  }

  const emailMatch = (cvText || "").match(/[\w.\-]+@[\w.\-]+\.\w+/);
  const phoneMatch = (cvText || "").match(/(\+?\d[\d\s/\-]{7,16}\d)/);
  const contactParts = [];
  for (const line of cvLines.slice(1, 6)) {
    const low = line.toLowerCase();
    if (
      ["str.", "straße", "weg ", "platz", "berlin", "münchen", "hamburg", "köln", "frankfurt", "stuttgart", "düsseldorf", "wien", "zürich"].some((w) =>
        low.includes(w)
      ) ||
      /\b\d{5}\b/.test(line)
    ) {
      contactParts.push(line);
      break;
    }
  }
  if (emailMatch) contactParts.push(emailMatch[0]);
  if (phoneMatch) contactParts.push(phoneMatch[1].trim());
  const kontakt =
    contactParts.length > 0
      ? Array.from(new Set(contactParts)).join("  |  ")
      : "Musterstraße 1, 10115 Berlin  |  kontakt@beispiel.de";

  const sections = { kurzprofil: [], berufserfahrung: [], ausbildung: [], kenntnisse: [] };
  let currentSec = "kurzprofil";
  for (const line of cvLines.slice(1)) {
    const low = line.toLowerCase();
    if (["berufserfahrung", "werdegang", "praxis", "berufliche stationen", "erfahrung"].some((k) => low.includes(k))) {
      currentSec = "berufserfahrung";
      continue;
    }
    if (["ausbildung", "studium", "bildungsweg", "schule", "akademisch"].some((k) => low.includes(k))) {
      currentSec = "ausbildung";
      continue;
    }
    if (["kenntnisse", "skills", "fähigkeiten", "kompetenzen", "sprachen", "it-", "zertifikat"].some((k) => low.includes(k))) {
      currentSec = "kenntnisse";
      continue;
    }
    sections[currentSec].push(line);
  }

  const rawTitle = (jobInfo.title || "").trim();
  const jobText = (jobInfo.text || "").trim();
  let company = (jobInfo.company || "").trim();

  let cleanTitle = rawTitle ? rawTitle.split(/\s+[\|\-–—]\s+/)[0].trim() : "";
  if (!cleanTitle || cleanTitle.toLowerCase().startsWith("http") || cleanTitle.toLowerCase().includes("manuell")) {
    for (const jline of jobText.split(/\r?\n/)) {
      const jl = jline.trim();
      if (jl.length > 5 && jl.length < 95 && !jl.toLowerCase().startsWith("http")) {
        cleanTitle = jl;
        break;
      }
    }
  }
  if (!cleanTitle) cleanTitle = "die ausgeschriebene Position";

  if (!company) {
    const compMatch = jobText.match(
      /(?:bei der|bei|unternehmen|firma|arbeitgeber)\s+([A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9&.\-\s]{2,40}?(?:GmbH|AG|SE|KG|UG|Inc\.|Group|Solutions|Systems|Partner|Consulting))/
    );
    if (compMatch) company = compMatch[1].trim();
  }

  const empfaenger = company ? `${company}\nPersonalabteilung` : "Personalabteilung\n[Unternehmensname]";

  const commonSkills = [
    "Projektmanagement", "Python", "JavaScript", "TypeScript", "React", "Java", "C#", "SQL",
    "Cloud", "AWS", "Azure", "Docker", "Kubernetes", "SAP", "Marketing", "Vertrieb", "Sales",
    "Kundenbetreuung", "Controlling", "Finanzen", "Buchhaltung", "HR", "Personalwesen",
    "Design", "UI/UX", "Agil", "Scrum", "Datenanalyse", "KI", "Machine Learning", "IT-Support",
    "Qualitätsmanagement", "Logistik", "Einkauf", "Kommunikation", "Teamführung", "Strategie",
    "MS Office", "Excel", "PowerBI", "Prozessoptimierung", "Beratung", "Consulting", "Engineering"
  ];
  const combinedSource = `${jobText}\n${cvText}`.toLowerCase();
  let matchedKeywords = commonSkills.filter((kw) => combinedSource.includes(kw.toLowerCase()));
  if (matchedKeywords.length === 0) {
    matchedKeywords = ["strukturierte Arbeitsweise", "hohe Einsatzbereitschaft", "schnelle Auffassungsgabe", "Teamorientierung"];
  }
  const topKwStr = matchedKeywords.slice(0, 5).join(", ");

  const berufText =
    sections.berufserfahrung.join("\n").trim() ||
    (cvText || "").trim() ||
    "Mehrjährige praktische Erfahrung in relevanten Fachprojekten und erfolgreiche Umsetzung anspruchsvoller Aufgabenstellungen.";
  const ausbText =
    sections.ausbildung.join("\n").trim() ||
    "Fundierte fachliche Ausbildung sowie kontinuierliche Weiterbildung in aktuellen Methoden und Technologien.";
  const skillsText =
    sections.kenntnisse.join("\n").trim() ||
    `Kernkompetenzen: ${topKwStr}\nSprachen: Deutsch (fließend), Englisch (gute Kenntnisse)`;
  const profilRaw = sections.kurzprofil.join("\n").trim();

  let kurzprofil = `Engagierte und ergebnisorientierte Fachkraft mit Fokus auf ${cleanTitle}. Fundierte Erfahrung in den Bereichen ${topKwStr} sowie ausgeprägte Stärke in der zielgerichteten Umsetzung anspruchsvoller Projekte.`;
  if (profilRaw.length > 40 && !berufText.startsWith(profilRaw.slice(0, 30))) {
    kurzprofil = `${kurzprofil}\n${profilRaw.slice(0, 350)}`;
  }

  const firmaPhrase = company ? `bei ${company}` : "in Ihrem Unternehmen";
  const betreff = `Bewerbung als ${cleanTitle}`;
  const anrede = "Sehr geehrte Damen und Herren,";

  const anschreibenText =
    `mit großem Interesse habe ich Ihre Stellenausschreibung für die Position als ${cleanTitle} gelesen. ` +
    `Da meine beruflichen Schwerpunkte und Qualifikationen hervorragend zu den von Ihnen gesuchten Anforderungen passen, ` +
    `möchte ich meine Expertise gerne gewinnbringend ${firmaPhrase} einbringen.\n\n` +
    `In meinem bisherigen Werdegang konnte ich umfassende praktische Erfahrungen sammeln und meine Kompetenzen insbesondere ` +
    `in den Bereichen ${topKwStr} erfolgreich unter Beweis stellen. Dabei zeichne ich mich durch eine analytische, ` +
    `qualitätsbewusste und lösungsorientierte Arbeitsweise aus, mit der ich mich schnell und sicher in neue Themengebiete einarbeite.\n\n` +
    `Besonders reizt mich an der ausgeschriebenen Stelle als ${cleanTitle} die Möglichkeit, verantwortungsvolle Aufgaben aktiv ` +
    `mitzugestalten und gemeinsam im Team nachhaltige Ergebnisse zu erzielen. Die Verbindung aus fachlicher Tiefe, ` +
    `Zuverlässigkeit und hoher Eigeninitiative bildet für mich die Grundlage einer erfolgreichen Zusammenarbeit.\n\n` +
    `Über die Gelegenheit, mich Ihnen in einem persönlichen Gespräch vorzustellen und Sie von meiner Motivation zu überzeugen, ` +
    `freue ich mich sehr.`;

  const fullCvText =
    `${name}\n${cleanTitle} | ${kontakt}\n\n` +
    `KURZPROFIL\n${kurzprofil}\n\n` +
    `BERUFSERFAHRUNG\n${berufText}\n\n` +
    `AUSBILDUNG & STUDIUM\n${ausbText}\n\n` +
    `KENNTNISSE & KOMPETENZEN\n${skillsText}`;

  return {
    name,
    kontakt,
    empfaenger,
    datum: todayStr,
    zielposition: cleanTitle,
    unternehmen: company || "Unternehmen",
    betreff,
    anrede,
    anschreiben_text: anschreibenText,
    kurzprofil,
    berufserfahrung: berufText,
    ausbildung: ausbText,
    kenntnisse: skillsText,
    lebenslauf_komplett: fullCvText,
    used_provider: "Smart-Lokal"
  };
}

export async function generateTailoredApplication(jobInfo, cvText, config) {
  const todayStr = formatTodayDE();
  const provider = (config.provider || "smart_local").toLowerCase().trim();
  const apiKey = (config.api_key || "").trim();

  if (provider !== "smart_local" && (apiKey || provider === "ollama")) {
    const systemPrompt =
      "Du bist ein erstklassiger deutscher Karriere-Coach und HR-Experte. " +
      "Deine Aufgabe ist es, anhand einer Stellenanzeige und des bisherigen Lebenslaufs des Bewerbers " +
      "ein individuell auf genau diese Stelle zugeschnittenes, überzeugendes deutsches Bewerbungsschreiben (Anschreiben) " +
      "sowie einen passend auf die Stelle optimierten, professionell gegliederten Lebenslauf zu erstellen. " +
      "Erfinde keine falschen Stationen hinzu, aber hebe genau die Erfahrungen, Erfolge und Kompetenzen aus dem Lebenslauf hervor, " +
      "die für diese konkrete Stellenanzeige besonders relevant sind. " +
      "Antworte AUSSCHLIESSLICH als valides JSON-Objekt mit exakt diesen Schlüsseln:\n" +
      "{\n" +
      '  "name": "Vor- und Nachname des Bewerbers",\n' +
      '  "kontakt": "Adresse | Telefon | E-Mail | LinkedIn (kompakt in 1-2 Zeilen)",\n' +
      '  "empfaenger": "Unternehmensname und Ansprechpartner/Abteilung für den Briefkopf",\n' +
      '  "datum": "Heutiges Datum im Format TT.MM.JJJJ",\n' +
      '  "zielposition": "Exakter Titel der Zielposition aus der Stellenanzeige",\n' +
      '  "unternehmen": "Name des Unternehmens aus der Stellenanzeige",\n' +
      '  "betreff": "Aussagekräftige Betreffzeile, z. B. Bewerbung als ...",\n' +
      '  "anrede": "Passende Anrede, z. B. Sehr geehrte Damen und Herren, oder mit Namen",\n' +
      '  "anschreiben_text": "Der komplette Haupttext des Anschreibens (3-4 Absätze, getrennt durch \\n\\n, OHNE Betreff, OHNE Anrede und OHNE Grußformel am Ende)",\n' +
      '  "kurzprofil": "Auf die Stelle zugeschnittenes Kurzprofil (3-4 prägnante Sätze)",\n' +
      '  "berufserfahrung": "Chronologisch strukturierte Berufserfahrung (relevanteste Aufgaben & Erfolge passend zur Stelle hervorgehoben, sauber mit Zeilenumbrüchen formatiert)",\n' +
      '  "ausbildung": "Strukturierter Bildungsweg / Studium / Ausbildung",\n' +
      '  "kenntnisse": "Auf die Stelle abgestimmte Fachkenntnisse, IT-Skills, Zertifikate und Sprachen"\n' +
      "}";

    const userPrompt =
      `HEUTIGES DATUM: ${todayStr}\n\n` +
      `=== STELLENANZEIGE ===\n` +
      `URL: ${jobInfo.url || ""}\n` +
      `Titel: ${jobInfo.title || ""}\n` +
      `Unternehmen: ${jobInfo.company || ""}\n` +
      `Inhalt der Stellenanzeige:\n${jobInfo.text || ""}\n\n` +
      `=== BISHERIGER LEBENSLAUF DES BEWERBERS ===\n${cvText}\n`;

    try {
      const aiRes = await callLlmJson(userPrompt, systemPrompt, config);
      aiRes.datum = aiRes.datum || todayStr;
      aiRes.lebenslauf_komplett =
        `${aiRes.name || ""}\n${aiRes.zielposition || ""} | ${aiRes.kontakt || ""}\n\n` +
        `KURZPROFIL\n${aiRes.kurzprofil || ""}\n\n` +
        `BERUFSERFAHRUNG\n${aiRes.berufserfahrung || ""}\n\n` +
        `AUSBILDUNG & STUDIUM\n${aiRes.ausbildung || ""}\n\n` +
        `KENNTNISSE & KOMPETENZEN\n${aiRes.kenntnisse || ""}`;
      aiRes.used_provider = `KI (${provider.toUpperCase()})`;
      return aiRes;
    } catch (err) {
      const fallback = smartLocalGenerate(jobInfo, cvText);
      fallback.used_provider = `Smart-Lokal (Fallback: ${String(err.message).slice(0, 60)})`;
      return fallback;
    }
  }

  return smartLocalGenerate(jobInfo, cvText);
}

export async function aiMapDocxBlocks(blocks, docType, generatedData, config) {
  const provider = (config.provider || "smart_local").toLowerCase().trim();
  const apiKey = (config.api_key || "").trim();

  if (provider === "smart_local" || (!apiKey && provider !== "ollama")) {
    return {};
  }

  const systemPrompt =
    "Du bist ein Experte für Word-Dokument-Strukturen (.docx). " +
    "Dir wird eine Liste von Textblöcken (mit 'id' und 'text') aus einem Muster-Word-Layout übergeben, " +
    "sowie die neuen Bewerbungsdaten eines Kandidaten. " +
    "Identifiziere diejenigen Textblöcke im Muster-Layout, die Beispiel-/Blindtext des Vorbesitzers enthalten " +
    "und ordne ihnen den passenden neuen Text zu. " +
    "WICHTIG: Lasse statische Sektionsüberschriften UNVERÄNDERT! " +
    'Antworte AUSSCHLIESSLICH mit JSON: {"replacements": {"document.xml::p_0": "Neuer Text...", ...}}';

  const userPrompt =
    `DOKUMENT-TYP: ${docType === "cover_letter" ? "Anschreiben" : "Lebenslauf"}\n\n` +
    `TEXTBLÖCKE:\n${JSON.stringify(blocks.slice(0, 80), null, 2)}\n\n` +
    `NEUE DATEN:\n${JSON.stringify(generatedData, null, 2)}`;

  try {
    const res = await callLlmJson(userPrompt, systemPrompt, config);
    if (res && typeof res.replacements === "object") {
      return res.replacements;
    }
  } catch (_) {}
  return {};
}
