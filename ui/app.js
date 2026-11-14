// ==========================================================================
// BEWERBUNGS-TERMINAL // CLEAN FRONTEND CONTROLLER
// ==========================================================================

const AppState = {
  cv_text: "",
  selected_cover_layout: "",
  selected_cv_layout: "",
  jobs: [],
  layouts: [],
  outputs: [],
  config: {
    provider: "smart_local",
    api_key: "",
    model: "",
    ollama_url: "http://localhost:11434"
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

async function apiPost(path, payload = {}) {
  const res = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  return await res.json();
}

function scheduleAutoSave() {
  const ind = document.getElementById("cvSaveIndicator");
  if (ind) ind.textContent = "Speichere...";
  if (saveTimeout) clearTimeout(saveTimeout);
  saveTimeout = setTimeout(async () => {
    try {
      await apiPost("/api/save-state", {
        state: {
          cv_text: AppState.cv_text,
          selected_cover_layout: AppState.selected_cover_layout,
          selected_cv_layout: AppState.selected_cv_layout,
          jobs: AppState.jobs
        }
      });
      if (ind) ind.textContent = "Gespeichert";
    } catch (e) {
      if (ind) ind.textContent = "Fehler";
    }
  }, 400);
}

async function initApp() {
  try {
    const res = await fetch("/api/state");
    const data = await res.json();
    if (data.ok) {
      AppState.cv_text = data.state.cv_text || "";
      AppState.selected_cover_layout = data.state.selected_cover_layout || "";
      AppState.selected_cv_layout = data.state.selected_cv_layout || "";
      AppState.jobs = Array.isArray(data.state.jobs) && data.state.jobs.length > 0
        ? data.state.jobs
        : [
            { id: "job_1", url: "", manual_text: "", title: "", company: "", status: "idle", result: null },
            { id: "job_2", url: "", manual_text: "", title: "", company: "", status: "idle", result: null },
            { id: "job_3", url: "", manual_text: "", title: "", company: "", status: "idle", result: null }
          ];
      AppState.layouts = data.layouts || [];
      AppState.outputs = data.outputs || [];
      AppState.config = data.config || AppState.config;

      document.getElementById("cvTextEditor").value = AppState.cv_text;
      renderLayoutsUI();
      renderJobSlots();
      renderResultsUI();
      updateAiStatusPill();
      logTerminal(`Terminal bereit (${AppState.layouts.length} Word-Layouts geladen).`, "ok");
    }
  } catch (err) {
    logTerminal("Initialisierungsfehler: " + err.message, "err");
  }
  bindEvents();
}

function renderLayoutsUI() {
  const coverSel = document.getElementById("selectCoverLayout");
  const cvSel = document.getElementById("selectCvLayout");
  const badge = document.getElementById("layoutCountBadge");

  badge.textContent = `${AppState.layouts.length} Vorlage(n) gespeichert`;

  if (AppState.layouts.length > 0) {
    const names = AppState.layouts.map(l => l.filename);
    if (!names.includes(AppState.selected_cover_layout)) {
      const defCover = AppState.layouts.find(l => l.filename.toLowerCase().includes("anschreiben")) || AppState.layouts[0];
      AppState.selected_cover_layout = defCover.filename;
    }
    if (!names.includes(AppState.selected_cv_layout)) {
      const defCv = AppState.layouts.find(l => l.filename.toLowerCase().includes("lebenslauf")) || AppState.layouts[0];
      AppState.selected_cv_layout = defCv.filename;
    }
  }

  coverSel.innerHTML = AppState.layouts
    .map(l => `<option value="${escapeHtml(l.filename)}" ${l.filename === AppState.selected_cover_layout ? "selected" : ""}>${escapeHtml(l.filename)}</option>`)
    .join("");

  cvSel.innerHTML = AppState.layouts
    .map(l => `<option value="${escapeHtml(l.filename)}" ${l.filename === AppState.selected_cv_layout ? "selected" : ""}>${escapeHtml(l.filename)}</option>`)
    .join("");
}

window.openFileInWord = async function(type, filename) {
  await apiPost("/api/open-path", { type, filename });
};

async function deleteLayoutByName(filename) {
  if (!filename) return;
  if (!confirm(`Vorlage "${filename}" wirklich löschen?`)) return;
  const res = await apiPost("/api/delete-layout", { filename });
  if (res.ok) {
    AppState.layouts = res.layouts;
    renderLayoutsUI();
    scheduleAutoSave();
    logTerminal(`Vorlage gelöscht: ${filename}`, "warn");
  }
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
        <button class="icon-btn" onclick="toggleManualText(${idx})" title="Zusatztext / Stellenbeschreibung einfügen">📝</button>
        <button class="icon-btn" onclick="processSingleJob(${idx})" title="Nur diese Stelle generieren">▶</button>
        <button class="icon-btn danger" onclick="removeJobSlot(${idx})" title="Feld entfernen">✕</button>
      </div>
      ${showManual ? `
        <textarea class="job-manual-area" placeholder="Optional: Stellenbeschreibung hier einfügen oder per Drag & Drop ablegen...">${escapeHtml(job.manual_text || "")}</textarea>
      ` : ""}
    `;

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
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const allAreUrls = lines.length > 0 && lines.every(l => l.startsWith("http://") || l.startsWith("https://"));

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
  const lines = rawText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
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
      const emptySlot = AppState.jobs.find(j => !j.url && !j.manual_text);
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

window.toggleManualText = function(idx) {
  AppState.jobs[idx]._manualOpen = !AppState.jobs[idx]._manualOpen;
  renderJobSlots();
};

window.removeJobSlot = function(idx) {
  if (AppState.jobs.length <= 1) {
    AppState.jobs[0] = { id: "job_1", url: "", manual_text: "", title: "", company: "", status: "idle", result: null };
  } else {
    AppState.jobs.splice(idx, 1);
  }
  renderJobSlots();
  renderResultsUI();
  scheduleAutoSave();
};

window.processSingleJob = async function(idx) {
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
    const res = await apiPost("/api/process-job", {
      job,
      cv_text: AppState.cv_text,
      cover_layout: AppState.selected_cover_layout,
      cv_layout: AppState.selected_cv_layout,
      config: AppState.config
    });

    if (res.ok) {
      job.status = "done";
      job.title = res.generated.zielposition || res.scraped.title || `Stelle #${idx + 1}`;
      job.company = res.generated.unternehmen || res.scraped.company || "";
      job.result = { generated: res.generated, files: res.files, scraped: res.scraped };
      AppState.outputs = res.outputs || AppState.outputs;
      prog.textContent = `✓ Stelle #${idx + 1} fertig`;
      logTerminal(`✓ Fertig: ${res.files.cover_letter.filename} & ${res.files.cv.filename}`, "ok");
    } else {
      job.status = "error";
      prog.textContent = `Fehler bei #${idx + 1}`;
      logTerminal(`Fehler bei #${idx + 1}: ${res.error}`, "err");
    }
  } catch (err) {
    job.status = "error";
    prog.textContent = `Fehler bei #${idx + 1}`;
    logTerminal(`Serverfehler: ${err.message}`, "err");
  } finally {
    renderJobSlots();
    renderResultsUI();
    scheduleAutoSave();
  }
};

async function processAllJobs() {
  const filledIndices = AppState.jobs
    .map((j, idx) => ((j.url && j.url.trim()) || (j.manual_text && j.manual_text.trim()) ? idx : -1))
    .filter(idx => idx !== -1);

  if (filledIndices.length === 0) {
    alert("Bitte füge links mindestens einen Job-Link (oder Stellentext) ein.");
    return;
  }

  for (let c = 0; c < filledIndices.length; c++) {
    const idx = filledIndices[c];
    document.getElementById("bridgeProgressText").textContent = `Verarbeite ${c + 1} / ${filledIndices.length}...`;
    await window.processSingleJob(idx);
  }
  document.getElementById("bridgeProgressText").textContent = `✓ Alle ${filledIndices.length} Stelle(n) fertig`;
}

function renderResultsUI() {
  const section = document.getElementById("resultsSection");
  const container = document.getElementById("resultsContainer");
  const completedJobs = AppState.jobs
    .map((job, idx) => ({ job, idx }))
    .filter(item => item.job.result && item.job.result.generated);

  if (completedJobs.length === 0) {
    section.hidden = true;
    container.innerHTML = "";
    return;
  }

  section.hidden = false;
  container.innerHTML = completedJobs.map(({ job, idx }) => {
    const gen = job.result.generated;
    const files = job.result.files;
    const isExpanded = Boolean(job._editOpen);
    return `
      <div class="result-card">
        <div class="result-top">
          <div>
            <span class="result-title">#${idx + 1} ${escapeHtml(gen.zielposition || "Bewerbung")}</span>
            <span class="result-sub">${escapeHtml(gen.unternehmen || "")}</span>
          </div>
          <div class="result-actions">
            <button class="btn-download" onclick="openFileInWord('output_file', '${escapeHtml(files.cover_letter.filename)}')">
              📘 Anschreiben (.docx)
            </button>
            <button class="btn-download" onclick="openFileInWord('output_file', '${escapeHtml(files.cv.filename)}')">
              📗 Lebenslauf (.docx)
            </button>
            <button class="btn-subtle btn-xs" onclick="toggleResultEdit(${idx})">
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
                <button class="btn-subtle btn-xs" onclick="reMergeJobResult(${idx})">🔄 Änderungen mit Layout neu mergen</button>
              </div>
            </div>
          </div>
        ` : ""}
      </div>
    `;
  }).join("");
}

window.toggleResultEdit = function(idx) {
  AppState.jobs[idx]._editOpen = !AppState.jobs[idx]._editOpen;
  renderResultsUI();
};

window.reMergeJobResult = async function(idx) {
  const job = AppState.jobs[idx];
  if (!job || !job.result) return;

  const gen = job.result.generated;
  gen.anschreiben_text = document.getElementById(`editCover_${idx}`).value;
  gen.berufserfahrung = document.getElementById(`editCvExp_${idx}`).value;

  const res = await apiPost("/api/remake-docx", {
    generated: gen,
    cover_layout: AppState.selected_cover_layout,
    cv_layout: AppState.selected_cv_layout
  });

  if (res.ok) {
    job.result.files = res.files;
    AppState.outputs = res.outputs || AppState.outputs;
    renderResultsUI();
    scheduleAutoSave();
    document.getElementById("bridgeProgressText").textContent = `✓ Stelle #${idx + 1} neu gemerged`;
    logTerminal(`Neu gemerged: ${res.files.cover_letter.filename}`, "ok");
  }
};

async function handleLayoutFilesUpload(fileList) {
  const docxFiles = Array.from(fileList).filter(f => f.name.toLowerCase().endsWith(".docx"));
  if (docxFiles.length === 0) {
    alert("Bitte nur Word-Dateien (.docx) als Vorlage ablegen.");
    return;
  }

  const payloadFiles = [];
  for (const file of docxFiles) {
    const b64 = await readFileAsBase64(file);
    payloadFiles.push({ name: file.name, data_base64: b64 });
  }

  const res = await apiPost("/api/upload-layout", { files: payloadFiles });
  if (res.ok) {
    AppState.layouts = res.layouts;
    if (res.saved.length > 0) {
      const firstNew = res.saved[0];
      if (firstNew.toLowerCase().includes("lebenslauf") || firstNew.toLowerCase().includes("cv")) {
        AppState.selected_cv_layout = firstNew;
      } else {
        AppState.selected_cover_layout = firstNew;
      }
    }
    renderLayoutsUI();
    scheduleAutoSave();
    logTerminal(`${res.saved.length} Layout(s) gespeichert.`, "ok");
  }
}

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = err => reject(err);
    reader.readAsDataURL(file);
  });
}

async function handleCvFileImport(file) {
  if (!file) return;
  const nameLow = file.name.toLowerCase();
  if (nameLow.endsWith(".txt")) {
    const text = await file.text();
    AppState.cv_text = text;
    document.getElementById("cvTextEditor").value = text;
    scheduleAutoSave();
  } else if (nameLow.endsWith(".docx")) {
    const b64 = await readFileAsBase64(file);
    const res = await apiPost("/api/extract-docx-text", { data_base64: b64 });
    if (res.ok) {
      AppState.cv_text = res.text;
      document.getElementById("cvTextEditor").value = res.text;
      scheduleAutoSave();
    }
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
    if (AppState.selected_cover_layout) openFileInWord("layout_file", AppState.selected_cover_layout);
  });
  document.getElementById("btnOpenSelectedCv").addEventListener("click", () => {
    if (AppState.selected_cv_layout) openFileInWord("layout_file", AppState.selected_cv_layout);
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
    const filtered = AppState.jobs.filter(j => (j.url && j.url.trim()) || (j.manual_text && j.manual_text.trim()));
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

  document.getElementById("btnOpenLayoutsDir").addEventListener("click", () => apiPost("/api/open-path", { type: "layouts_dir" }));
  document.getElementById("btnOpenOutputDir").addEventListener("click", () => apiPost("/api/open-path", { type: "output_dir" }));
  document.getElementById("btnOpenOutputDirBottom").addEventListener("click", () => apiPost("/api/open-path", { type: "output_dir" }));

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
    modal.hidden = false;
  };
  document.getElementById("btnOpenSettings").addEventListener("click", openModal);
  document.getElementById("aiStatusPill").addEventListener("click", openModal);
  document.getElementById("btnCloseSettings").addEventListener("click", () => { modal.hidden = true; });

  document.getElementById("btnSaveConfig").addEventListener("click", async () => {
    AppState.config = {
      provider: document.getElementById("cfgProvider").value,
      api_key: document.getElementById("cfgApiKey").value.trim(),
      model: document.getElementById("cfgModel").value.trim(),
      ollama_url: document.getElementById("cfgOllamaUrl").value.trim() || "http://localhost:11434"
    };
    await apiPost("/api/save-state", { config: AppState.config });
    updateAiStatusPill();
    modal.hidden = true;
    logTerminal(`Einstellungen gespeichert (${AppState.config.provider}).`, "ok");
  });
}

window.addEventListener("DOMContentLoaded", initApp);
