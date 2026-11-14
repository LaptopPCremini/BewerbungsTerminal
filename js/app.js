// ==========================================================================
// BEWERBUNGS-TERMINAL // 100 % CLIENT-SIDE GITHUB PAGES CONTROLLER
// ==========================================================================

import {
  getAllLayoutsFromDB,
  saveLayoutToDB,
  deleteLayoutFromDB,
  loadStateFromLS,
  saveStateToLS,
  loadConfigFromLS,
  saveConfigToLS
} from "./storage.js";

import {
  inspectDocx,
  mergeDocxTemplate,
  autoMapBlocksHeuristic,
  buildPlaceholderMap,
  generateDefaultSampleLayouts
} from "./docx_engine.js";

import {
  scrapeJobUrl,
  generateTailoredApplication,
  aiMapDocxBlocks
} from "./ai_engine.js";

const AppState = {
  cv_text: "",
  selected_cover_layout: "",
  selected_cv_layout: "",
  jobs: [],
  layouts: [], // { filename, buffer, sizeKb, modified, placeholders, blockCount, guessedType }
  config: {
    provider: "smart_local",
    api_key: "",
    model: "",
    ollama_url: "http://localhost:11434",
    use_cors_proxy: true
  }
};

let saveTimeout = null;

function logTerminal(msg, level = "info") {
  const box = document.getElementById("terminalLogOutput");
  if (!box) return;
  const now = new Date().toLocaleTimeString("de-DE", { hour12: false });
  const div = document.createElement("div");
  const cls =
    level === "ok" ? "log-ok" :
    level === "warn" ? "log-warn" :
    level === "err" ? "log-err" : "log-info";
  div.innerHTML = `<span class="log-time">[${now}]</span><span class="${cls}">${escapeHtml(msg)}</span>`;
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
}

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function safeFilename(text) {
  const cleaned = String(text || "")
    .trim()
    .replace(/[^A-Za-z0-9ÄÖÜäöüß\-_]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
  return cleaned.slice(0, 45) || "Dokument";
}

function triggerBlobDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

function scheduleAutoSave() {
  const ind = document.getElementById("cvSaveIndicator");
  if (ind) ind.textContent = "Speichere...";
  if (saveTimeout) clearTimeout(saveTimeout);
  saveTimeout = setTimeout(() => {
    saveStateToLS(AppState);
    if (ind) ind.textContent = "Gespeichert";
  }, 350);
}

async function initApp() {
  try {
    const savedState = loadStateFromLS();
    AppState.cv_text = savedState.cv_text || "";
    AppState.selected_cover_layout = savedState.selected_cover_layout || "";
    AppState.selected_cv_layout = savedState.selected_cv_layout || "";
    AppState.jobs = Array.isArray(savedState.jobs) && savedState.jobs.length > 0
      ? savedState.jobs
      : [
          { id: "job_1", url: "", manual_text: "", title: "", company: "", status: "idle", result: null },
          { id: "job_2", url: "", manual_text: "", title: "", company: "", status: "idle", result: null },
          { id: "job_3", url: "", manual_text: "", title: "", company: "", status: "idle", result: null }
        ];
    AppState.config = loadConfigFromLS();

    // Lade Layouts aus IndexedDB; falls leer, erzeuge 5 Muster-Layouts
    let layouts = await getAllLayoutsFromDB();
    if (layouts.length === 0) {
      logTerminal("Erster Start: Erstelle 5 Muster-Word-Layouts im Browser-Speicher...", "info");
      const defaults = await generateDefaultSampleLayouts();
      for (const item of defaults) {
        await saveLayoutToDB(item);
      }
      layouts = await getAllLayoutsFromDB();
    }
    AppState.layouts = layouts;

    document.getElementById("cvTextEditor").value = AppState.cv_text;
    renderLayoutsUI();
    renderJobSlots();
    renderResultsUI();
    updateAiStatusPill();
    logTerminal(`Terminal bereit (${AppState.layouts.length} Word-Layouts im Browser geladen).`, "ok");
  } catch (err) {
    logTerminal("Initialisierungsfehler: " + err.message, "err");
  }

  bindEvents();
}

function renderLayoutsUI() {
  const coverSel = document.getElementById("selectCoverLayout");
  const cvSel = document.getElementById("selectCvLayout");
  const badge = document.getElementById("layoutCountBadge");

  badge.textContent = `${AppState.layouts.length} Vorlage(n) im Browser`;

  if (AppState.layouts.length > 0) {
    const names = AppState.layouts.map((l) => l.filename);
    if (!names.includes(AppState.selected_cover_layout)) {
      const defCover = AppState.layouts.find((l) => l.filename.toLowerCase().includes("anschreiben")) || AppState.layouts[0];
      AppState.selected_cover_layout = defCover.filename;
    }
    if (!names.includes(AppState.selected_cv_layout)) {
      const defCv = AppState.layouts.find((l) => l.filename.toLowerCase().includes("lebenslauf")) || AppState.layouts[0];
      AppState.selected_cv_layout = defCv.filename;
    }
  }

  coverSel.innerHTML = AppState.layouts
    .map((l) => `<option value="${escapeHtml(l.filename)}" ${l.filename === AppState.selected_cover_layout ? "selected" : ""}>${escapeHtml(l.filename)}</option>`)
    .join("");

  cvSel.innerHTML = AppState.layouts
    .map((l) => `<option value="${escapeHtml(l.filename)}" ${l.filename === AppState.selected_cv_layout ? "selected" : ""}>${escapeHtml(l.filename)}</option>`)
    .join("");
}

function downloadStoredLayout(filename) {
  const layout = AppState.layouts.find((l) => l.filename === filename);
  if (!layout) return;
  const blob = new Blob([layout.buffer], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
  triggerBlobDownload(blob, layout.filename);
}

async function deleteLayoutByName(filename) {
  if (!filename) return;
  if (!confirm(`Vorlage "${filename}" wirklich aus dem Browser-Speicher löschen?`)) return;
  await deleteLayoutFromDB(filename);
  AppState.layouts = await getAllLayoutsFromDB();
  renderLayoutsUI();
  scheduleAutoSave();
  logTerminal(`Vorlage gelöscht: ${filename}`, "warn");
}

function renderJobSlots() {
  const container = document.getElementById("jobSlotsList");
  container.innerHTML = "";

  AppState.jobs.forEach((job, idx) => {
    const row = document.createElement("div");
    row.className = `job-row status-${job.status || "idle"}`;

    const showManual = Boolean(job.manual_text && job.manual_text.trim().length > 0) || Boolean(job._manualOpen);

    row.innerHTML = `
      <div class="job-row-main">
        <span class="job-num">#${idx + 1}</span>
        <span class="job-status-dot ${job.status || ""}"></span>
        <input
          type="text"
          class="job-url-input"
          placeholder="Job-Link hier ablegen (Drag & Drop) oder URL einfügen..."
          value="${escapeHtml(job.url || "")}"
        />
        <button class="icon-btn" data-action="toggle-manual" title="Zusatztext / Stellenbeschreibung einfügen">📝</button>
        <button class="icon-btn" data-action="run-single" title="Nur diese Stelle generieren">▶</button>
        <button class="icon-btn danger" data-action="remove" title="Feld entfernen">✕</button>
      </div>
      ${showManual ? `
        <textarea class="job-manual-area" placeholder="Optional: Stellenbeschreibung hier einfügen oder per Drag & Drop ablegen...">${escapeHtml(job.manual_text || "")}</textarea>
      ` : ""}
    `;

    row.querySelector('[data-action="toggle-manual"]').addEventListener("click", () => {
      AppState.jobs[idx]._manualOpen = !AppState.jobs[idx]._manualOpen;
      renderJobSlots();
    });
    row.querySelector('[data-action="run-single"]').addEventListener("click", () => processSingleJob(idx));
    row.querySelector('[data-action="remove"]').addEventListener("click", () => removeJobSlot(idx));

    const urlInput = row.querySelector(".job-url-input");
    urlInput.addEventListener("input", (e) => {
      const val = e.target.value;
      if (val.includes("\n")) {
        handleMultiLinksPaste(val, idx);
        return;
      }
      AppState.jobs[idx].url = val;
      scheduleAutoSave();
    });

    const manualArea = row.querySelector(".job-manual-area");
    if (manualArea) {
      manualArea.addEventListener("input", (e) => {
        AppState.jobs[idx].manual_text = e.target.value;
        scheduleAutoSave();
      });
    }

    row.addEventListener("dragover", (e) => {
      e.preventDefault();
      row.classList.add("drag-over");
    });
    row.addEventListener("dragleave", () => row.classList.remove("drag-over"));
    row.addEventListener("drop", (e) => {
      e.preventDefault();
      row.classList.remove("drag-over");
      const droppedUri = e.dataTransfer.getData("text/uri-list") || e.dataTransfer.getData("text/plain") || "";
      if (droppedUri.trim()) {
        handleDroppedTextOnJobSlot(droppedUri.trim(), idx);
      }
    });

    container.appendChild(row);
  });
}

function handleDroppedTextOnJobSlot(text, idx) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const allAreUrls = lines.length > 0 && lines.every((l) => l.startsWith("http://") || l.startsWith("https://"));

  if (allAreUrls && lines.length > 1) {
    handleMultiLinksPaste(text, idx);
  } else if (text.startsWith("http://") || text.startsWith("https://")) {
    AppState.jobs[idx].url = lines[0];
    logTerminal(`Link #${idx + 1}: ${lines[0]}`, "info");
    renderJobSlots();
    scheduleAutoSave();
  } else {
    AppState.jobs[idx].manual_text = text;
    AppState.jobs[idx]._manualOpen = true;
    logTerminal(`Stellentext in #${idx + 1} übernommen.`, "info");
    renderJobSlots();
    scheduleAutoSave();
  }
}

function handleMultiLinksPaste(rawText, startIdx = null) {
  const lines = rawText.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
  if (lines.length === 0) return;

  if (startIdx !== null && startIdx < AppState.jobs.length) {
    AppState.jobs[startIdx].url = lines[0];
    for (let i = 1; i < lines.length; i++) {
      AppState.jobs.push({
        id: "job_" + Date.now() + "_" + i,
        url: lines[i],
        manual_text: "",
        title: "",
        company: "",
        status: "idle",
        result: null
      });
    }
  } else {
    lines.forEach((line, i) => {
      const emptySlot = AppState.jobs.find((j) => !j.url && !j.manual_text);
      if (emptySlot) {
        emptySlot.url = line;
      } else {
        AppState.jobs.push({
          id: "job_" + Date.now() + "_" + i,
          url: line,
          manual_text: "",
          title: "",
          company: "",
          status: "idle",
          result: null
        });
      }
    });
  }

  renderJobSlots();
  scheduleAutoSave();
}

function removeJobSlot(idx) {
  if (AppState.jobs.length <= 1) {
    AppState.jobs[0] = { id: "job_1", url: "", manual_text: "", title: "", company: "", status: "idle", result: null };
  } else {
    AppState.jobs.splice(idx, 1);
  }
  renderJobSlots();
  renderResultsUI();
  scheduleAutoSave();
}

async function mergeWithSelectedLayoutClient(layoutFilename, docType, genData, outFilename) {
  const layoutObj = AppState.layouts.find((l) => l.filename === layoutFilename);
  if (!layoutObj) {
    throw new Error(`Layout '${layoutFilename}' nicht gefunden.`);
  }

  const info = await inspectDocx(layoutObj.buffer);
  const placeholdersMap = buildPlaceholderMap(genData);
  let blockReplacements = {};

  if (!info.placeholders || info.placeholders.length === 0) {
    blockReplacements = autoMapBlocksHeuristic(info.blocks, docType, genData);
    const aiBlocks = await aiMapDocxBlocks(info.blocks, docType, genData, AppState.config);
    if (aiBlocks && Object.keys(aiBlocks).length > 0) {
      blockReplacements = { ...blockReplacements, ...aiBlocks };
    }
  }

  const mergeRes = await mergeDocxTemplate(layoutObj.buffer, placeholdersMap, blockReplacements);
  return {
    filename: outFilename,
    blob: mergeRes.blob,
    replacedPlaceholders: mergeRes.replacedPlaceholders,
    replacedBlocks: mergeRes.replacedBlocks,
    mode: info.placeholders.length > 0 ? "Platzhalter + Design-Merge" : "KI-Auto-Block-Ersetzung"
  };
}

async function processSingleJob(idx) {
  const job = AppState.jobs[idx];
  if (!job) return;
  if (!job.url.trim() && !(job.manual_text || "").trim()) {
    document.getElementById("bridgeProgressText").textContent = `Feld #${idx + 1} ist leer`;
    return;
  }

  AppState.cv_text = document.getElementById("cvTextEditor").value;
  job.status = "running";
  renderJobSlots();

  const prog = document.getElementById("bridgeProgressText");
  prog.textContent = `Erstelle Stelle #${idx + 1}...`;
  logTerminal(`Verarbeite Stelle #${idx + 1}...`, "info");

  try {
    const url = (job.url || "").trim();
    const manual = (job.manual_text || "").trim();

    let scraped = await scrapeJobUrl(url || manual, AppState.config.use_cors_proxy !== false);
    if (manual) {
      if (scraped.ok) {
        scraped.text = `${scraped.text}\n\n${manual}`.trim();
      } else {
        scraped = {
          ok: true,
          url: url || "Manueller Text",
          title: (manual.split(/\r?\n/)[0] || "Stellenanzeige").slice(0, 80),
          company: "",
          text: manual
        };
      }
    }

    if (!scraped.text) {
      if (scraped.error) logTerminal(`Hinweis #${idx + 1}: ${scraped.error}`, "warn");
      scraped.text = `Stellenanzeige unter Link: ${url}`;
    }

    const genData = await generateTailoredApplication(scraped, AppState.cv_text, AppState.config);

    const now = new Date();
    const ts =
      `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_` +
      `${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}${String(now.getSeconds()).padStart(2, "0")}`;
    const jobSlug = safeFilename(genData.unternehmen || genData.zielposition || "Stelle");

    const coverOutName = `Bewerbungsschreiben_${jobSlug}_${ts}.docx`;
    const cvOutName = `Lebenslauf_${jobSlug}_${ts}.docx`;

    const coverMerge = await mergeWithSelectedLayoutClient(AppState.selected_cover_layout, "cover_letter", genData, coverOutName);
    const cvMerge = await mergeWithSelectedLayoutClient(AppState.selected_cv_layout, "cv", genData, cvOutName);

    job.status = "done";
    job.title = genData.zielposition || scraped.title || `Stelle #${idx + 1}`;
    job.company = genData.unternehmen || scraped.company || "";
    job.result = {
      generated: genData,
      files: {
        cover_letter: coverMerge,
        cv: cvMerge
      },
      scraped
    };

    prog.textContent = `✓ Stelle #${idx + 1} fertig`;
    logTerminal(`✓ Fertig: ${coverOutName} & ${cvOutName}`, "ok");
  } catch (err) {
    job.status = "error";
    prog.textContent = `Fehler bei #${idx + 1}`;
    logTerminal(`Fehler bei #${idx + 1}: ${err.message}`, "err");
  } finally {
    renderJobSlots();
    renderResultsUI();
    scheduleAutoSave();
  }
}

async function processAllJobs() {
  const filledIndices = AppState.jobs
    .map((j, idx) => ((j.url && j.url.trim()) || (j.manual_text && j.manual_text.trim()) ? idx : -1))
    .filter((idx) => idx !== -1);

  if (filledIndices.length === 0) {
    alert("Bitte füge links mindestens einen Job-Link (oder Stellentext) ein.");
    return;
  }

  for (let c = 0; c < filledIndices.length; c++) {
    const idx = filledIndices[c];
    document.getElementById("bridgeProgressText").textContent = `Verarbeite ${c + 1} / ${filledIndices.length}...`;
    await processSingleJob(idx);
  }
  document.getElementById("bridgeProgressText").textContent = `✓ Alle ${filledIndices.length} Stelle(n) fertig`;
}

function renderResultsUI() {
  const section = document.getElementById("resultsSection");
  const container = document.getElementById("resultsContainer");
  const completedJobs = AppState.jobs
    .map((job, idx) => ({ job, idx }))
    .filter((item) => item.job.result && item.job.result.generated);

  if (completedJobs.length === 0) {
    section.hidden = true;
    container.innerHTML = "";
    return;
  }

  section.hidden = false;
  container.innerHTML = "";

  completedJobs.forEach(({ job, idx }) => {
    const gen = job.result.generated;
    const files = job.result.files;
    const isExpanded = Boolean(job._editOpen);

    const card = document.createElement("div");
    card.className = "result-card";
    card.innerHTML = `
      <div class="result-top">
        <div>
          <span class="result-title">#${idx + 1} ${escapeHtml(gen.zielposition || "Bewerbung")}</span>
          <span class="result-sub">${escapeHtml(gen.unternehmen || "")} · ${escapeHtml(gen.used_provider || "")}</span>
        </div>
        <div class="result-actions">
          <button class="btn-download" data-dl="cover">📘 Anschreiben (.docx)</button>
          <button class="btn-download" data-dl="cv">📗 Lebenslauf (.docx)</button>
          <button class="btn-subtle btn-xs" data-toggle="edit">
            ${isExpanded ? "▲ Text ausblenden" : "✏️ Text ansehen / bearbeiten"}
          </button>
        </div>
      </div>
      ${isExpanded ? `
        <div class="result-edit-box">
          <div class="edit-col">
            <label>Anschreiben (${escapeHtml(gen.betreff || "")})</label>
            <textarea class="edit-textarea" id="editCover_${idx}">${escapeHtml(gen.anschreiben_text || "")}</textarea>
          </div>
          <div class="edit-col">
            <label>Angepasste Berufserfahrung (Lebenslauf)</label>
            <textarea class="edit-textarea" id="editCvExp_${idx}">${escapeHtml(gen.berufserfahrung || "")}</textarea>
            <div style="display:flex; justify-content:flex-end;">
              <button class="btn-subtle btn-xs" data-remerge="true">🔄 Änderungen mit Layout neu mergen</button>
            </div>
          </div>
        </div>
      ` : ""}
    `;

    card.querySelector('[data-dl="cover"]').addEventListener("click", () => {
      triggerBlobDownload(files.cover_letter.blob, files.cover_letter.filename);
    });
    card.querySelector('[data-dl="cv"]').addEventListener("click", () => {
      triggerBlobDownload(files.cv.blob, files.cv.filename);
    });
    card.querySelector('[data-toggle="edit"]').addEventListener("click", () => {
      AppState.jobs[idx]._editOpen = !AppState.jobs[idx]._editOpen;
      renderResultsUI();
    });

    const remergeBtn = card.querySelector('[data-remerge="true"]');
    if (remergeBtn) {
      remergeBtn.addEventListener("click", () => reMergeJobResult(idx));
    }

    container.appendChild(card);
  });
}

async function reMergeJobResult(idx) {
  const job = AppState.jobs[idx];
  if (!job || !job.result) return;

  const gen = job.result.generated;
  gen.anschreiben_text = document.getElementById(`editCover_${idx}`).value;
  gen.berufserfahrung = document.getElementById(`editCvExp_${idx}`).value;

  const now = new Date();
  const ts =
    `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_` +
    `${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}${String(now.getSeconds()).padStart(2, "0")}`;
  const jobSlug = safeFilename(gen.unternehmen || gen.zielposition || "Stelle");

  const coverOutName = `Bewerbungsschreiben_${jobSlug}_${ts}.docx`;
  const cvOutName = `Lebenslauf_${jobSlug}_${ts}.docx`;

  const coverMerge = await mergeWithSelectedLayoutClient(AppState.selected_cover_layout, "cover_letter", gen, coverOutName);
  const cvMerge = await mergeWithSelectedLayoutClient(AppState.selected_cv_layout, "cv", gen, cvOutName);

  job.result.files = { cover_letter: coverMerge, cv: cvMerge };
  renderResultsUI();
  document.getElementById("bridgeProgressText").textContent = `✓ Stelle #${idx + 1} neu gemerged`;
  logTerminal(`Neu gemerged: ${coverOutName} & ${cvOutName}`, "ok");
}

async function downloadAllAsZip() {
  const completedJobs = AppState.jobs.filter((j) => j.result && j.result.files);
  if (completedJobs.length === 0) {
    alert("Es wurden noch keine Bewerbungsunterlagen generiert.");
    return;
  }

  const zip = new JSZip();
  completedJobs.forEach((j) => {
    const f = j.result.files;
    zip.file(f.cover_letter.filename, f.cover_letter.blob);
    zip.file(f.cv.filename, f.cv.blob);
  });

  const zipBlob = await zip.generateAsync({ type: "blob" });
  triggerBlobDownload(zipBlob, `Bewerbungsunterlagen_${new Date().toISOString().slice(0, 10)}.zip`);
  logTerminal(`ZIP-Paket mit ${completedJobs.length * 2} Dokumenten heruntergeladen.`, "ok");
}

async function handleLayoutFilesUpload(fileList) {
  const docxFiles = Array.from(fileList).filter((f) => f.name.toLowerCase().endsWith(".docx"));
  if (docxFiles.length === 0) {
    alert("Bitte nur Word-Dateien (.docx) als Vorlage ablegen.");
    return;
  }

  const savedNames = [];
  for (const file of docxFiles) {
    const buffer = await file.arrayBuffer();
    const info = await inspectDocx(buffer);
    await saveLayoutToDB({
      filename: file.name,
      buffer,
      sizeKb: Math.round((buffer.byteLength / 1024) * 10) / 10,
      modified: new Date().toLocaleString("de-DE"),
      placeholders: info.placeholders,
      blockCount: info.blockCount,
      guessedType: info.guessedType
    });
    savedNames.push(file.name);
  }

  AppState.layouts = await getAllLayoutsFromDB();
  if (savedNames.length > 0) {
    const firstNew = savedNames[0];
    if (firstNew.toLowerCase().includes("lebenslauf") || firstNew.toLowerCase().includes("cv")) {
      AppState.selected_cv_layout = firstNew;
    } else {
      AppState.selected_cover_layout = firstNew;
    }
  }
  renderLayoutsUI();
  scheduleAutoSave();
  logTerminal(`${savedNames.length} Layout(s) im Browser gespeichert.`, "ok");
}

async function handleCvFileImport(file) {
  if (!file) return;
  const nameLow = file.name.toLowerCase();
  if (nameLow.endsWith(".txt")) {
    const text = await file.text();
    AppState.cv_text = text;
    document.getElementById("cvTextEditor").value = text;
    scheduleAutoSave();
    logTerminal(`Lebenslauf aus '${file.name}' geladen.`, "ok");
  } else if (nameLow.endsWith(".docx")) {
    const buffer = await file.arrayBuffer();
    const info = await inspectDocx(buffer);
    const fullText = info.blocks.map((b) => b.text).join("\n");
    AppState.cv_text = fullText;
    document.getElementById("cvTextEditor").value = fullText;
    scheduleAutoSave();
    logTerminal(`Text aus '${file.name}' extrahiert.`, "ok");
  }
}

function updateAiStatusPill() {
  const txt = document.getElementById("aiStatusText");
  const p = AppState.config.provider || "smart_local";
  const labels = {
    smart_local: "Smart-Lokal",
    gemini: "Google Gemini",
    openai: "OpenAI GPT",
    anthropic: "Anthropic Claude",
    ollama: "Ollama (Lokal)"
  };
  txt.textContent = labels[p] || "Aktiv";
}

function bindEvents() {
  const docxDrop = document.getElementById("docxDropzone");
  const hiddenLayoutInput = document.getElementById("hiddenLayoutFileInput");

  docxDrop.addEventListener("click", () => hiddenLayoutInput.click());
  hiddenLayoutInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) handleLayoutFilesUpload(e.target.files);
    e.target.value = "";
  });

  docxDrop.addEventListener("dragover", (e) => {
    e.preventDefault();
    docxDrop.classList.add("drag-over");
  });
  docxDrop.addEventListener("dragleave", () => docxDrop.classList.remove("drag-over"));
  docxDrop.addEventListener("drop", (e) => {
    e.preventDefault();
    docxDrop.classList.remove("drag-over");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleLayoutFilesUpload(e.dataTransfer.files);
    }
  });

  document.getElementById("selectCoverLayout").addEventListener("change", (e) => {
    AppState.selected_cover_layout = e.target.value;
    scheduleAutoSave();
  });
  document.getElementById("selectCvLayout").addEventListener("change", (e) => {
    AppState.selected_cv_layout = e.target.value;
    scheduleAutoSave();
  });

  document.getElementById("btnOpenSelectedCover").addEventListener("click", () => {
    if (AppState.selected_cover_layout) downloadStoredLayout(AppState.selected_cover_layout);
  });
  document.getElementById("btnOpenSelectedCv").addEventListener("click", () => {
    if (AppState.selected_cv_layout) downloadStoredLayout(AppState.selected_cv_layout);
  });
  document.getElementById("btnDeleteCoverLayout").addEventListener("click", () => {
    deleteLayoutByName(AppState.selected_cover_layout);
  });
  document.getElementById("btnDeleteCvLayout").addEventListener("click", () => {
    deleteLayoutByName(AppState.selected_cv_layout);
  });

  document.getElementById("btnAddJobSlot").addEventListener("click", () => {
    AppState.jobs.push({ id: "job_" + Date.now(), url: "", manual_text: "", title: "", company: "", status: "idle", result: null });
    renderJobSlots();
    scheduleAutoSave();
  });

  document.getElementById("btnClearEmptyJobs").addEventListener("click", () => {
    const filtered = AppState.jobs.filter((j) => (j.url && j.url.trim()) || (j.manual_text && j.manual_text.trim()));
    AppState.jobs = filtered.length > 0 ? filtered : [{ id: "job_1", url: "", manual_text: "", title: "", company: "", status: "idle", result: null }];
    renderJobSlots();
    scheduleAutoSave();
  });

  const multiDrop = document.getElementById("multiLinkDropzone");
  multiDrop.addEventListener("click", () => {
    AppState.jobs.push({ id: "job_" + Date.now(), url: "", manual_text: "", title: "", company: "", status: "idle", result: null });
    renderJobSlots();
  });
  multiDrop.addEventListener("dragover", (e) => {
    e.preventDefault();
    multiDrop.classList.add("drag-over");
  });
  multiDrop.addEventListener("dragleave", () => multiDrop.classList.remove("drag-over"));
  multiDrop.addEventListener("drop", (e) => {
    e.preventDefault();
    multiDrop.classList.remove("drag-over");
    const dropped = e.dataTransfer.getData("text/uri-list") || e.dataTransfer.getData("text/plain") || "";
    if (dropped.trim()) handleMultiLinksPaste(dropped.trim(), null);
  });

  document.getElementById("btnProcessAllJobs").addEventListener("click", processAllJobs);

  const cvEditor = document.getElementById("cvTextEditor");
  cvEditor.addEventListener("input", (e) => {
    AppState.cv_text = e.target.value;
    scheduleAutoSave();
  });

  const cvPanel = document.getElementById("cvDropPanel");
  cvPanel.addEventListener("dragover", (e) => {
    if (e.dataTransfer.types.includes("Files")) {
      e.preventDefault();
      cvPanel.classList.add("drag-over");
    }
  });
  cvPanel.addEventListener("dragleave", () => cvPanel.classList.remove("drag-over"));
  cvPanel.addEventListener("drop", (e) => {
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      e.preventDefault();
      cvPanel.classList.remove("drag-over");
      handleCvFileImport(e.dataTransfer.files[0]);
    }
  });

  const hiddenCvInput = document.getElementById("hiddenCvFileInput");
  document.getElementById("btnImportCvFile").addEventListener("click", () => hiddenCvInput.click());
  hiddenCvInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) handleCvFileImport(e.target.files[0]);
    e.target.value = "";
  });

  document.getElementById("btnDownloadAllZip").addEventListener("click", downloadAllAsZip);
  document.getElementById("btnDownloadAllZipBottom").addEventListener("click", downloadAllAsZip);

  document.getElementById("btnToggleLog").addEventListener("click", () => {
    const drawer = document.getElementById("logDrawer");
    drawer.hidden = !drawer.hidden;
  });
  document.getElementById("btnClearLog").addEventListener("click", () => {
    document.getElementById("terminalLogOutput").innerHTML = "";
  });

  const modal = document.getElementById("settingsModal");
  const openModal = () => {
    document.getElementById("cfgProvider").value = AppState.config.provider || "smart_local";
    document.getElementById("cfgApiKey").value = AppState.config.api_key || "";
    document.getElementById("cfgModel").value = AppState.config.model || "";
    document.getElementById("cfgOllamaUrl").value = AppState.config.ollama_url || "http://localhost:11434";
    document.getElementById("cfgCorsProxy").checked = AppState.config.use_cors_proxy !== false;
    modal.hidden = false;
  };
  document.getElementById("btnOpenSettings").addEventListener("click", openModal);
  document.getElementById("aiStatusPill").addEventListener("click", openModal);
  document.getElementById("btnCloseSettings").addEventListener("click", () => { modal.hidden = true; });

  document.getElementById("btnSaveConfig").addEventListener("click", () => {
    AppState.config = {
      provider: document.getElementById("cfgProvider").value,
      api_key: document.getElementById("cfgApiKey").value.trim(),
      model: document.getElementById("cfgModel").value.trim(),
      ollama_url: document.getElementById("cfgOllamaUrl").value.trim() || "http://localhost:11434",
      use_cors_proxy: document.getElementById("cfgCorsProxy").checked
    };
    saveConfigToLS(AppState.config);
    updateAiStatusPill();
    modal.hidden = true;
    logTerminal(`Einstellungen im Browser gespeichert (${AppState.config.provider}).`, "ok");
  });
}

window.addEventListener("DOMContentLoaded", initApp);
