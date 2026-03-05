// DOM elements
const inputFile = document.getElementById("input-file");
const outputFile = document.getElementById("output-file");
const provider = document.getElementById("provider");
const cliRow = document.getElementById("cli-row");
const modelRow = document.getElementById("model-row");
const keyRow = document.getElementById("key-row");
const cliCommand = document.getElementById("cli-command");
const model = document.getElementById("model");
const apiKey = document.getElementById("api-key");
const zoteroDB = document.getElementById("zotero-db");
const font = document.getElementById("font");
const size = document.getElementById("size");
const bibHeading = document.getElementById("bib-heading");
const referenceDoc = document.getElementById("reference-doc");
const noCrossref = document.getElementById("no-crossref");
const dryRun = document.getElementById("dry-run");
const formatBtn = document.getElementById("format-btn");
const logOutput = document.getElementById("log-output");

// Resolve modal elements
const resolveModal = document.getElementById("resolve-modal");
const resolveHeader = document.getElementById("resolve-header");
const resolveCandidates = document.getElementById("resolve-candidates");
const resolveOk = document.getElementById("resolve-ok");

let pendingResolve = null;

// ---------- Provider visibility ----------

function updateProviderFields() {
  const isCli = provider.value === "cli";
  cliRow.classList.toggle("hidden", !isCli);
  modelRow.classList.toggle("hidden", isCli);
  keyRow.classList.toggle("hidden", isCli);
}

provider.addEventListener("change", updateProviderFields);
updateProviderFields();

// ---------- Auto-detect Zotero DB ----------

window.api.getDefaultZoteroDB().then((dbPath) => {
  if (dbPath) zoteroDB.value = dbPath;
});

// ---------- File browsers ----------

document.getElementById("browse-input").addEventListener("click", async () => {
  const path = await window.api.openFileDialog({
    filters: [
      { name: "Documents", extensions: ["docx", "md", "markdown", "txt"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
  if (path) inputFile.value = path;
});

document.getElementById("browse-output").addEventListener("click", async () => {
  const path = await window.api.saveFileDialog();
  if (path) outputFile.value = path;
});

document.getElementById("browse-zotero").addEventListener("click", async () => {
  const path = await window.api.openFileDialog({
    filters: [
      { name: "SQLite", extensions: ["sqlite"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
  if (path) zoteroDB.value = path;
});

document.getElementById("browse-refdoc").addEventListener("click", async () => {
  const path = await window.api.openFileDialog({
    filters: [
      { name: "Word Documents", extensions: ["docx"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
  if (path) referenceDoc.value = path;
});

// ---------- Format button ----------

formatBtn.addEventListener("click", async () => {
  const inputPath = inputFile.value.trim();
  if (!inputPath) {
    alert("Please select an input file.");
    return;
  }

  logOutput.textContent = "";
  formatBtn.disabled = true;
  formatBtn.textContent = "Processing...";

  const args = {
    input: inputPath,
    output: outputFile.value.trim() || null,
    provider: provider.value,
    model: model.value.trim() || null,
    api_key: apiKey.value.trim() || null,
    cli_command: cliCommand.value.trim() || null,
    zotero_db: zoteroDB.value.trim() || null,
    reference_doc: referenceDoc.value.trim() || null,
    font: font.value.trim() || "Calibri",
    size: parseInt(size.value) || 11,
    bib_heading: bibHeading.value.trim() || "References",
    no_crossref: noCrossref.checked,
    dry_run: dryRun.checked,
  };

  const result = await window.api.startProcessing(args);
  if (!result.ok) {
    appendLog(`ERROR: ${result.error}\n`);
    formatBtn.disabled = false;
    formatBtn.textContent = "Format Citations";
  }
});

// ---------- Backend messages ----------

function appendLog(text) {
  logOutput.textContent += text;
  logOutput.scrollTop = logOutput.scrollHeight;
}

window.api.onBackendMessage((msg) => {
  switch (msg.type) {
    case "log":
      appendLog(msg.text + "\n");
      break;

    case "resolve":
      showResolveDialog(msg);
      break;

    case "done":
      if (msg.success) {
        appendLog(`\n${msg.message}\n`);
      } else {
        appendLog(`\nERROR: ${msg.message}\n`);
      }
      formatBtn.disabled = false;
      formatBtn.textContent = "Format Citations";
      break;
  }
});

// ---------- Resolve dialog ----------

function showResolveDialog(msg) {
  pendingResolve = msg;
  resolveHeader.innerHTML = `<strong>"${escapeHtml(msg.citation_text)}"</strong> matched to:`;
  resolveCandidates.innerHTML = "";

  const groupName = `resolve-${msg.id}`;

  msg.candidates.forEach((c, i) => {
    const item = c.item;
    const score = c.score;
    const title = Array.isArray(item.title) ? item.title[0] : (item.title || "Unknown");
    const authors = item.author || [];
    const firstAuthor = authors.length > 0 ? (authors[0].family || "") : "";
    const dateParts = (item.issued || {})["date-parts"] || [[]];
    const year = dateParts[0] && dateParts[0][0] ? dateParts[0][0] : "";
    const doi = item.DOI || "";
    const ct = item["container-title"];
    const journal = Array.isArray(ct) ? (ct[0] || "") : (ct || "");

    let labelText = `${firstAuthor} (${year})`;
    if (journal) labelText += ` ${journal} -`;
    labelText += ` ${title.substring(0, 80)}`;
    if (doi) labelText += `  [DOI: ${doi}]`;
    labelText += `  (score: ${score})`;

    const div = document.createElement("div");
    div.className = "resolve-option";
    div.innerHTML = `
      <input type="radio" name="${groupName}" value="${i}" id="${groupName}-${i}" ${i === 0 ? "checked" : ""}>
      <label for="${groupName}-${i}">${escapeHtml(labelText)}</label>
    `;
    resolveCandidates.appendChild(div);
  });

  // Skip option
  const skipDiv = document.createElement("div");
  skipDiv.className = "resolve-option";
  skipDiv.innerHTML = `
    <input type="radio" name="${groupName}" value="skip" id="${groupName}-skip">
    <label for="${groupName}-skip">Skip this citation</label>
  `;
  resolveCandidates.appendChild(skipDiv);

  // Manual DOI/PMID option
  const manualDiv = document.createElement("div");
  manualDiv.className = "resolve-manual-row";
  manualDiv.innerHTML = `
    <input type="radio" name="${groupName}" value="manual" id="${groupName}-manual">
    <label for="${groupName}-manual">Enter DOI/PMID:</label>
    <input type="text" id="manual-input" placeholder="e.g. 10.1000/xyz or 12345678">
  `;
  resolveCandidates.appendChild(manualDiv);

  // Focus manual radio when typing
  const manualInput = document.getElementById("manual-input");
  manualInput.addEventListener("input", () => {
    document.getElementById(`${groupName}-manual`).checked = true;
  });

  resolveModal.classList.remove("hidden");
}

resolveOk.addEventListener("click", () => {
  if (!pendingResolve) return;

  const groupName = `resolve-${pendingResolve.id}`;
  const selected = document.querySelector(`input[name="${groupName}"]:checked`);
  let choice;

  if (!selected || selected.value === "skip") {
    choice = "skip";
  } else if (selected.value === "manual") {
    const val = document.getElementById("manual-input").value.trim();
    choice = val || "skip";
  } else {
    choice = parseInt(selected.value);
  }

  window.api.resolveResponse({
    type: "resolve_response",
    id: pendingResolve.id,
    choice,
  });

  resolveModal.classList.add("hidden");
  pendingResolve = null;
});

// ---------- Utility ----------

function escapeHtml(str) {
  const div = document.createElement("div");
  div.appendChild(document.createTextNode(str));
  return div.innerHTML;
}
