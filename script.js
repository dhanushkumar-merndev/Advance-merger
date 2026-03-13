function refreshIcons() {
  if (typeof lucide !== "undefined") lucide.createIcons();
}
let GOOGLE_SA_EMAIL = "";
let GOOGLE_PRIVATE_KEY = "";

async function loadEnv() {
  try {
    const resp = await fetch("./env.txt");
    if (!resp.ok) return;
    const text = await resp.text();
    // Use a more robust parsing for env.txt
    const lines = text.split(/\r?\n/);
    let currentKey = "";
    let currentValue = "";

    function flush() {
      if (!currentKey) return;
      let val = currentValue.trim();
      if (val.startsWith('"') && val.endsWith('"')) val = val.slice(1, -1);

      if (
        currentKey === "GOOGLE_CLIENT_EMAIL" ||
        currentKey === "GOOGLE_SA_EMAIL"
      ) {
        GOOGLE_SA_EMAIL = val;
        updateSAUI();
      } else if (currentKey === "GOOGLE_PRIVATE_KEY") {
        GOOGLE_PRIVATE_KEY = val.replace(/\\n/g, "\n");
      }
    }

    for (const line of lines) {
      const match = line.match(/^([A-Z0-9_]+)\s*=\s*(.*)$/i);
      if (match) {
        flush();
        currentKey = match[1];
        currentValue = match[2];
      } else if (currentKey) {
        currentValue += "\n" + line;
      }
    }
    flush();
  } catch (e) {
    console.warn("Could not load .env:", e);
  }
}

function updateSAUI() {
  const el = document.getElementById("saEmailDisplay");
  if (el) {
    el.textContent = GOOGLE_SA_EMAIL.substring(0, 15) + "...";
    el.title = GOOGLE_SA_EMAIL;
  }
}

async function getGoogleAccessToken() {
  if (!GOOGLE_SA_EMAIL || !GOOGLE_PRIVATE_KEY) {
    throw new Error(
      "Google credentials missing. Make sure env.txt is loaded correctly.",
    );
  }

  const now = Math.floor(Date.now() / 1000);
  const expiry = now + 3600;

  const header = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: GOOGLE_SA_EMAIL,
    scope:
      "https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/drive.readonly",
    aud: "https://oauth2.googleapis.com/token",
    exp: expiry,
    iat: now,
  };

  try {
    const sHeader = JSON.stringify(header);
    const sPayload = JSON.stringify(payload);
    // Use KEYUTIL to parse the PEM string correctly
    const rsaKey = KEYUTIL.getKey(GOOGLE_PRIVATE_KEY);
    const sJWT = KJUR.jws.JWS.sign("RS256", sHeader, sPayload, rsaKey);

    const tokenResp = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        assertion: sJWT,
      }),
    });

    const data = await tokenResp.json();
    if (data.error) {
      console.error("Token exchange failed:", data);
      throw new Error(data.error_description || data.error);
    }
    if (data.access_token)
      console.log("Google Auth: Token obtained successfully");
    return data.access_token;
  } catch (err) {
    console.error("Auth error:", err);
    throw new Error("Google Auth Failed: " + err.message);
  }
}

async function copyServiceAccountEmail() {
  try {
    await navigator.clipboard.writeText(GOOGLE_SA_EMAIL);
    toast("Service Email copied!", "success");
  } catch (err) {
    toast("Copy failed", "error");
  }
}

// ─── SHEETJS LOADER ────────────────────────────────────────────────────────
async function loadSheetJS() {
  if (typeof XLSX !== "undefined") return;
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src =
      "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js";
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

function parseFileHeaders(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        await loadSheetJS();
        const data = new Uint8Array(e.target.result);
        // Read full workbook so we can inspect all sheets
        const wb = XLSX.read(data, { type: "array" });

        // Collect headers from all sheets (union, preserving first-seen order)
        const seen = new Set();
        const headers = [];
        for (const name of wb.SheetNames) {
          const ws = wb.Sheets[name];
          const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
          const row = json && json[0] ? json[0] : [];
          for (const h of row) {
            const norm = String(h || "")
              .toLowerCase()
              .trim();
            if (norm && !seen.has(norm)) {
              seen.add(norm);
              headers.push(norm);
            }
          }
        }

        resolve({ headers: headers, fileName: file.name });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── STATE ────────────────────────────────────────────────────────────────
const STATE = {
  templates: JSON.parse(localStorage.getItem("dmp_templates") || "[]"),
  groups: JSON.parse(localStorage.getItem("dmp_groups") || "[]"),
  campaigns: JSON.parse(localStorage.getItem("dmp_campaigns") || "[]"),
  runs: JSON.parse(localStorage.getItem("dmp_runs") || "0"),
  currentMode: 1,
  editingTemplate: null,
  editingGroup: null,
  editingCampaign: null,
  folderHandle: null,
  discoveredCols: [],
  _openNodes: new Set(), // Persistent tree state
  _outputFiles: [],
  _previewShowAll: { template: false, group: false, campaign: false },
  _lastCmd: "",
  _lastOneShot: "",
  _lastInputs: [],
  _inputFiles: [],
  _outputFiles: [],
  _activeOutputName: null,
  _currentWorkbook: null,
  _currentRows: [],
  _pendingHandle: null,
  currentTemplateLinks: [],
  analyticsStartDate: new Date().toISOString().split("T")[0],
  analyticsEndDate: new Date().toISOString().split("T")[0],
  analyticsView: "today",
  duplicationModal: {
    data: [],
    activeSheet: null,
    hideBaseline: true,
    filename: "",
  },
  runs: JSON.parse(localStorage.getItem("dmp_runs") || "0"),
  lastPage: localStorage.getItem("dmp_last_page") || "home",
};

function save() {
  localStorage.setItem("dmp_templates", JSON.stringify(STATE.templates));
  localStorage.setItem("dmp_groups", JSON.stringify(STATE.groups));
  localStorage.setItem("dmp_campaigns", JSON.stringify(STATE.campaigns));
  localStorage.setItem("dmp_runs", JSON.stringify(STATE.runs));
}

// ─── FORMAT CODES ─────────────────────────────────────────────────────────
const FORMAT_CODES = [
  { code: "f", label: "UPPERCASE", cat: "text" },
  { code: "g", label: "lowercase", cat: "text" },
  { code: "h", label: "Title Case", cat: "text" },
  { code: "l", label: "Align Left", cat: "text" },
  { code: "r", label: "Align Right", cat: "text" },
  { code: "z", label: "Align Center", cat: "text" },
  { code: "d", label: "Last 10 Digits", cat: "phone" },
  { code: "e", label: "Add +91", cat: "phone" },
  { code: "i", label: "Integers Only", cat: "phone" },
  { code: "a", label: "Today's Date", cat: "time" },
  { code: "b", label: "Time HH:MM", cat: "time" },
  { code: "c", label: "Time HH:MM:SS", cat: "time" },
  { code: "n", label: "Convert to IST", cat: "time" },
  { code: "j", label: "Remove Dashes", cat: "clean" },
  { code: "u", label: "Remove _", cat: "clean" },
  { code: "x", label: "Remove Dots", cat: "clean" },
  { code: "p", label: "Date: 3 Mar 2026", cat: "time" },
  { code: "0", label: "Blank Column", cat: "special" },
  { code: "k", label: "Dict Lookup", cat: "special" },
  { code: "q", label: "Dict + Default", cat: "special" },
];
const CAT_LABELS = {
  text: '<i data-lucide="type" class="icon"></i> Text',
  phone: '<i data-lucide="phone" class="icon"></i> Phone',
  time: '<i data-lucide="calendar" class="icon"></i> Date/Time',
  clean: '<i data-lucide="sparkles" class="icon"></i> Clean',
  special: '<i data-lucide="zap" class="icon"></i> Special',
};
const FMT_DESCRIPTIONS = {
  f: "→ UPPERCASE",
  g: "→ lowercase",
  h: "→ Title Case",
  l: "left-align",
  r: "right-align",
  z: "center-align",
  d: "last 10 digits",
  e: "prepend +91",
  i: "integers only",
  a: "today's date",
  b: "HH:MM",
  c: "HH:MM:SS",
  n: "→ IST",
  j: "no dashes",
  u: "no underscores",
  x: "no dots",
  p: "D MMM YYYY (Time)",
  0: "blank",
  k: "dict lookup",
  q: "dict + default",
};

// ─── NAVIGATION ──────────────────────────────────────────────────────────
async function navigate(page) {
  const el = document.getElementById("page-" + page);
  if (!el) return;

  const order = [
    "home",
    "input",
    "run",
    "template",
    "sources",
    "campaigns",
    "reports",
    "preview",
    "viewer",
    "settings",
  ];
  const oldIdx = order.indexOf(STATE.lastPage || "home");
  const newIdx = order.indexOf(page);
  const direction = newIdx >= oldIdx ? "slide-up" : "slide-down";
  STATE.lastPage = page;
  localStorage.setItem("dmp_last_page", page);

  document.querySelectorAll(".page").forEach((p) => {
    p.classList.remove("active", "slide-up", "slide-down");
  });
  document
    .querySelectorAll(".nav-item")
    .forEach((b) => b.classList.remove("active"));

  // Trigger directional animation reflow
  el.classList.add("active", direction);
  void el.offsetWidth; // Force reflow

  const navBtn = document.querySelector(`[data-page="${page}"]`);
  if (navBtn) navBtn.classList.add("active");
  closeAllPanels();
  if (page === "home") refreshHome();
  if (page === "template") renderTemplateList();
  if (page === "sources") renderGroupList();
  if (page === "campaigns") renderCampaignList();
  if (page === "reports") refreshReportsFiles();
  if (page === "run") await refreshRunPage();
  if (page === "preview") {
    STATE.previewMode = "template";
    STATE._previewShowAll = {
      template: false,
      group: false,
      campaign: false,
    };
    document
      .querySelectorAll(".preview-tab")
      .forEach((t) =>
        t.classList.toggle("active", t.dataset.pm === "template"),
      );
    refreshPreview();
  }
  if (page === "input") refreshInputPage();
  if (page === "viewer") refreshViewerPage();
  refreshIcons();
}

// ─── PREVIEW ──────────────────────────────────────────────────────────────
function setPreviewMode(mode, btn) {
  STATE.previewMode = mode;
  STATE._previewShowAll[mode] = false;
  document
    .querySelectorAll(".preview-tab")
    .forEach((t) => t.classList.remove("active"));
  btn.classList.add("active");
  refreshPreview();
}

function generateSampleRows(columns) {
  const SAMPLES = {
    phone: ["+919876543210", "+918765432109", "+917654321098"],
    mobile: ["+919876543210", "+918765432109", "+917654321098"],
    email: ["test@example.com", "user@sample.org", "lead@domain.in"],
    name: ["Rahul Kumar", "Priya Sharma", "Amit Singh"],
    date: ["25-01-2025", "26-01-2025", "27-01-2025"],
    city: ["Mumbai", "Delhi", "Bangalore"],
    state: ["MH", "DL", "KA"],
    source: ["Facebook", "Google", "Instagram"],
  };
  const rows = [[], [], []];
  columns.forEach((col) => {
    const key = col.name.toLowerCase().replace(/[^a-z]/g, "");
    const samples = SAMPLES[key] || [
      `${col.name}_1`,
      `${col.name}_2`,
      `${col.name}_3`,
    ];
    for (let i = 0; i < 3; i++) rows[i].push(samples[i] || "—");
  });
  return rows;
}

// ─── TOAST ─────────────────────────────────────────────────────────────────
let toastTimer;
function toast(msg, type = "ok") {
  const el = document.getElementById("toast");

  const iconMap = {
    success: "check-circle",
    error: "x-circle",
    info: "info",
    ok: "check",
  };

  const icon = iconMap[type] || "check";

  el.innerHTML = `
    <div style="display:flex;align-items:center;gap:8px">
      <i data-lucide="${icon}" class="icon"></i>
      <span>${msg}</span>
    </div>
  `;

  el.className =
    "toast show" +
    (type === "error"
      ? " error"
      : type === "success"
        ? " success"
        : type === "info"
          ? " info"
          : "");

  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove("show"), 3200);

  refreshIcons();
}

// ─── DIALOG ────────────────────────────────────────────────────────────────
function confirmDialog(title, body, cb) {
  document.getElementById("dialogTitle").textContent = title;
  document.getElementById("dialogBody").textContent = body;
  document.getElementById("dialogOverlay").style.display = "flex";
  document.getElementById("dialogConfirmBtn").onclick = () => {
    closeDialog();
    cb();
  };
}
function closeDialog() {
  document.getElementById("dialogOverlay").style.display = "none";
}

// ─── PANELS ────────────────────────────────────────────────────────────────
function openPanel(id) {
  document.getElementById(id).classList.add("open");
  document.getElementById("overlay").classList.add("show");
  refreshIcons();
}
function closePanel(id) {
  document.getElementById(id).classList.remove("open");
  document.getElementById("overlay").classList.remove("show");
}
function closeAllPanels() {
  document
    .querySelectorAll(".panel")
    .forEach((p) => p.classList.remove("open"));
  document.getElementById("overlay").classList.remove("show");
}

// ─── FOLDER ─────────────────────────────────────────────────────────────────
const FOLDER_DB_NAME = "dmp_folder_db";
const FOLDER_STORE = "handles";
const FOLDER_KEY = "project_folder";

function openFolderDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(FOLDER_DB_NAME, 1);
    req.onupgradeneeded = (e) =>
      e.target.result.createObjectStore(FOLDER_STORE);
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = (e) => reject(e);
  });
}
async function saveFolderHandle(handle) {
  try {
    const db = await openFolderDB();
    const tx = db.transaction(FOLDER_STORE, "readwrite");
    tx.objectStore(FOLDER_STORE).put(handle, FOLDER_KEY);
    await new Promise((res, rej) => {
      tx.oncomplete = res;
      tx.onerror = rej;
    });
  } catch (e) {}
}
async function loadFolderHandle() {
  try {
    const db = await openFolderDB();
    return await new Promise((resolve) => {
      const tx = db.transaction(FOLDER_STORE, "readonly");
      const req = tx.objectStore(FOLDER_STORE).get(FOLDER_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch (e) {
    return null;
  }
}
async function applyFolderUI(dir) {
  STATE.folderHandle = dir;

  document.getElementById("folderDot").className = "dot green";
  const nameEl = document.getElementById("folderNameEl");
  nameEl.textContent = "Connected";
  nameEl.className = "folder-name-label connected";

  const statusRow = document.querySelector(".folder-status-row");
  if (statusRow) statusRow.classList.add("is-connected");

  document.getElementById("btnFolderLink").innerHTML =
    '<i data-lucide="check-circle" class="icon"></i> ' + dir.name;

  const chk = document.getElementById("chkFolderIcon");
  if (chk) chk.textContent = "✓";

  const openBtn = document.getElementById("btnOpenFolder");
  if (openBtn) openBtn.disabled = false;

  refreshIcons();

  // Sequential sync to avoid race conditions
  await syncTemplatesWithFolder();
  await syncGroupsWithFolder();
  await syncCampaignsWithFolder();
  await autoDetectInputFiles();

  // Refresh UI to show newly loaded data
  renderTemplateList();
  renderGroupList();
  renderCampaignList();
}

async function syncGroupsWithFolder() {
  if (!STATE.folderHandle) return;
  try {
    const subDir = await STATE.folderHandle.getDirectoryHandle(
      "source_groups",
      { create: true },
    );
    const groups = [];
    for await (const [name, handle] of subDir.entries()) {
      if (handle.kind === "file" && name.endsWith(".json")) {
        try {
          const file = await handle.getFile();
          const json = JSON.parse(await file.text());
          if (json.name && json.sources) groups.push(json);
        } catch (e) {}
      }
    }
    if (groups.length > 0) {
      groups.forEach((g) => {
        const idx = STATE.groups.findIndex((ex) => ex.name === g.name);
        if (idx !== -1) STATE.groups[idx] = g;
        else STATE.groups.push(g);
      });
      save();
      renderGroupList();
    }
  } catch (e) {
    console.warn("Group sync failed", e);
  }
}

async function syncCampaignsWithFolder() {
  if (!STATE.folderHandle) return;
  try {
    const subDir = await STATE.folderHandle.getDirectoryHandle("campaigns", {
      create: true,
    });
    const campaigns = [];
    for await (const [name, handle] of subDir.entries()) {
      if (handle.kind === "file" && name.endsWith(".json")) {
        try {
          const file = await handle.getFile();
          const json = JSON.parse(await file.text());
          if (json.name && json.groups) campaigns.push(json);
        } catch (e) {}
      }
    }
    if (campaigns.length > 0) {
      campaigns.forEach((c) => {
        const idx = STATE.campaigns.findIndex((ex) => ex.name === c.name);
        if (idx !== -1) STATE.campaigns[idx] = c;
        else STATE.campaigns.push(c);
      });
      save();
      renderCampaignList();
      refreshCampaignSelects();
    }
  } catch (e) {
    console.warn("Campaign sync failed", e);
  }
}

function refreshCampaignSelects() {
  const select = document.getElementById("runCampaignConfig");
  if (select) {
    const val = select.value;
    select.innerHTML =
      '<option value="">-- Choose Campaign --</option>' +
      STATE.campaigns
        .map(
          (c) =>
            `<option value="${esc(c.name)}" ${c.name === val ? "selected" : ""}>${esc(c.name)}</option>`,
        )
        .join("");
  }
}

async function syncTemplatesWithFolder() {
  if (!STATE.folderHandle) return;
  try {
    const tplDir = await STATE.folderHandle.getDirectoryHandle("templates", {
      create: true,
    });
    const templates = [];
    for await (const [name, handle] of tplDir.entries()) {
      if (handle.kind === "file" && name.endsWith(".json")) {
        try {
          const file = await handle.getFile();
          const json = JSON.parse(await file.text());
          if (json.columns) {
            templates.push(jsonFileToStateTpl(name.replace(".json", ""), json));
          }
        } catch (e) {
          console.warn("Failed to parse template:", name, e);
        }
      }
    }
    if (templates.length > 0) {
      templates.forEach((t) => {
        const idx = STATE.templates.findIndex((ex) => ex.name === t.name);
        if (idx !== -1) STATE.templates[idx] = t;
        else STATE.templates.push(t);
      });
      save();
      renderTemplateList();
    }
  } catch (e) {
    console.warn("Sync templates failed:", e);
  }
}

function jsonFileToStateTpl(name, json) {
  // Map JSON file format to STATE format
  const columns = json.columns.map((colArr) => {
    const colName = colArr[0];
    const tokens = colArr.slice(1).filter((t) => typeof t === "string");
    const dict =
      colArr.find((t) => typeof t === "object" && t !== null) || null;

    // Separate src from formats
    // If tokens[0] is "0", it's explicitly no source.
    // If tokens[0] is a known format code, it's NOT a source (src="0").
    let finalSrc = "0";
    let finalFmtArr = tokens;

    if (tokens.length > 0) {
      const first = tokens[0];
      if (first === "0") {
        finalSrc = "0";
        finalFmtArr = tokens.slice(1);
      } else {
        const isFmt = FORMAT_CODES.some((f) => f.code === first);
        if (isFmt) {
          finalSrc = "0";
          finalFmtArr = tokens;
        } else {
          finalSrc = first;
          finalFmtArr = tokens.slice(1);
        }
      }
    }

    return { name: colName, src: finalSrc, fmt: finalFmtArr.join(" "), dict };
  });

  return {
    name: json.name || name,
    columns: columns,
    dedupCols: json.unique_columns || [],
    sortCol: json.sort_column || "",
    sortOrder: json.sort_order || "asc",
    googleLinks: json.google_links || json.googleLinks || [],
    pages: json.pages || [],
  };
}

async function disconnectFolder() {
  if (!confirm("Disconnect this project folder?")) return;
  try {
    const db = await openFolderDB();
    const tx = db.transaction(FOLDER_STORE, "readwrite");
    tx.objectStore(FOLDER_STORE).delete(FOLDER_KEY);
    await new Promise((resolve, reject) => {
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });

    STATE.folderHandle = null;
    document.getElementById("folderDot").className = "dot red";
    const nameEl = document.getElementById("folderNameEl");
    nameEl.textContent = "No folder linked";
    nameEl.className = "folder-name-label disconnected";

    const statusRow = document.querySelector(".folder-status-row");
    if (statusRow) statusRow.classList.remove("is-connected");

    document.getElementById("btnFolderLink").innerHTML =
      '<i data-lucide="link" class="icon"></i> Link Project Folder';

    const openBtn = document.getElementById("btnOpenFolder");
    if (openBtn) openBtn.disabled = true;

    refreshIcons();
    toast("Folder disconnected", "success");
  } catch (e) {
    console.error(e);
    toast("Failed to disconnect", "error");
  }
}

async function restoreFolderOnLoad() {
  const handle = await loadFolderHandle();
  if (!handle) return;
  try {
    const perm = await handle.queryPermission({ mode: "readwrite" });
    if (perm === "granted") {
      applyFolderUI(handle);
      return;
    }
    if (perm === "prompt") {
      document.getElementById("folderNameEl").textContent =
        handle.name + " (tap to reconnect)";
      STATE._pendingHandle = handle;
      document.getElementById("btnFolderLink").textContent = "🔄 Reconnect";
    }
  } catch (e) {}
}

async function writeToLinkedFolder(subfolder, filename, jsonObj) {
  if (!STATE.folderHandle) return false;
  try {
    const subHandle = await STATE.folderHandle.getDirectoryHandle(subfolder, {
      create: true,
    });
    const fileHandle = await subHandle.getFileHandle(filename, {
      create: true,
    });
    const writable = await fileHandle.createWritable();
    await writable.write(JSON.stringify(jsonObj, null, 2));
    await writable.close();
    return true;
  } catch (e) {
    console.warn(e);
    return false;
  }
}

async function linkFolder() {
  if (!window.showDirectoryPicker) {
    toast("Use Chrome or Edge for folder linking", "error");
    return;
  }
  if (STATE._pendingHandle) {
    try {
      const perm = await STATE._pendingHandle.requestPermission({
        mode: "readwrite",
      });
      if (perm === "granted") {
        applyFolderUI(STATE._pendingHandle);
        STATE._pendingHandle = null;
        toast("Folder reconnected!", "success");
        return;
      }
    } catch (e) {}
    STATE._pendingHandle = null;
  }
  try {
    const dir = await window.showDirectoryPicker({ mode: "readwrite" });
    applyFolderUI(dir);
    await saveFolderHandle(dir);
    toast("Folder linked: " + dir.name, "success");
  } catch (e) {
    if (e.name !== "AbortError") toast("Could not link folder", "error");
  }
}

function showFolderPath() {
  const folderName = STATE.folderHandle ? STATE.folderHandle.name : null;
  if (!folderName) {
    toast("No folder linked yet", "error");
    return;
  }
  const existing = document.getElementById("folderPathPopup");
  if (existing) {
    existing.remove();
    return;
  }
  const popup = document.createElement("div");
  popup.id = "folderPathPopup";
  popup.style.cssText =
    "position:fixed;bottom:90px;left:8px;width:210px;background:white;border:1.5px solid var(--border);border-radius:12px;padding:14px;z-index:600;box-shadow:var(--shadow-lg)";
  const cdCmd = `cd ~/${folderName}`;
  popup.innerHTML = `<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px"><div style="font-size:12px;font-weight:700;color:var(--text2)"><i data-lucide="folder" class="icon" style="vertical-align:text-bottom;margin-right:4px"></i> ${esc(folderName)}</div><button onclick="document.getElementById('folderPathPopup').remove()" style="background:none;border:none;cursor:pointer;font-size:14px;color:var(--text3)">✕</button></div><div onclick="copyToClipboard('${cdCmd.replace(/'/g, "\\'")}');toast('Copied!','success');document.getElementById('folderPathPopup').remove()" style="background:#0f1b3d;border-radius:8px;padding:8px 10px;font-family:var(--mono);font-size:11px;color:#7ec8a0;cursor:pointer">${esc(cdCmd)}</div><div style="font-size:10px;color:var(--text3);margin-top:6px;text-align:center">Click to copy</div>`;
  document.body.appendChild(popup);
  setTimeout(() => {
    const p = document.getElementById("folderPathPopup");
    if (p) p.remove();
  }, 12000);
}

function toggleTreeNode(el, nodeId) {
  const node = el.closest(".tree-node");
  if (!node) return;
  const isOpen = node.classList.toggle("open");
  if (isOpen) STATE._openNodes.add(nodeId);
  else STATE._openNodes.delete(nodeId);
  refreshIcons();
}

async function downloadInputFile(folder, fname) {
  if (!STATE.folderHandle) return;
  try {
    const inputDir = await STATE.folderHandle.getDirectoryHandle("input");
    const targetDir =
      folder === "_root" ? inputDir : await inputDir.getDirectoryHandle(folder);
    const fileHandle = await targetDir.getFileHandle(fname);
    const file = await fileHandle.getFile();
    const url = URL.createObjectURL(file);
    const a = document.createElement("a");
    a.href = url;
    a.download = fname;
    a.click();
    URL.revokeObjectURL(url);
    toast("Download started");
  } catch (e) {
    toast("Download failed: " + e.message, "error");
  }
}

async function downloadOutputFile(fname) {
  if (!STATE.folderHandle) return;
  try {
    const outDir = await STATE.folderHandle.getDirectoryHandle("output");
    const fileHandle = await outDir.getFileHandle(fname);
    const file = await fileHandle.getFile();
    const url = URL.createObjectURL(file);
    const a = document.createElement("a");
    a.href = url;
    a.download = fname;
    a.click();
    URL.revokeObjectURL(url);
    toast("Download started");
  } catch (e) {
    toast("Download failed: " + e.message, "error");
  }
}

// ─── HOME ──────────────────────────────────────────────────────────────────
function refreshHome() {
  refreshAnalyticsDashboard();
  document.getElementById("statTemplates").textContent = STATE.templates.length;
  document.getElementById("statGroups").textContent = STATE.groups.length;
  document.getElementById("statCampaigns").textContent = STATE.campaigns.length;
  document.getElementById("statRuns").textContent = STATE.runs;
}

// ─── TEMPLATES ─────────────────────────────────────────────────────────────
function renderTemplateList(query = "") {
  const el = document.getElementById("templateList");
  const items = query
    ? STATE.templates.filter((t) =>
        t.name.toLowerCase().includes(query.toLowerCase()),
      )
    : STATE.templates;
  if (!items.length) {
    el.innerHTML = query
      ? '<div class="empty-state">No templates matching "' +
        esc(query) +
        '"</div>'
      : '<div class="empty-state"><div class="es-icon"><i data-lucide="clipboard-list"></i></div><strong>No templates yet</strong>Create your first template to define output column mappings.</div>';
    return;
  }
  el.innerHTML = items
    .map((t, idx) => {
      const i = STATE.templates.indexOf(t);
      return `
    <div class="list-card">
      <div class="lc-left"><div class="lc-icon blue"><i data-lucide="clipboard-list"></i></div><div><div class="lc-title">${esc(t.name)}</div><div class="lc-sub">${t.columns.length} columns · Dedup: ${t.dedupCols.length ? t.dedupCols.join(", ") : "none"} ${t.sortCol ? ` · Sort: ${t.sortCol} (${t.sortOrder === "desc" ? "Z-A" : "A-Z"})` : ""}</div></div></div>

      <div class="lc-actions">

        <button class="btn-icon" onclick="downloadTemplateJson(STATE.templates[${i}])" title="Download JSON"><i data-lucide="download" class="icon"></i></button>
        <button class="btn-icon" onclick="editTemplate(${i})" title="Edit"><i data-lucide="edit-3" class="icon"></i></button>
        <button class="btn-icon danger" onclick="deleteTemplate(${i})" title="Delete"><i data-lucide="trash-2" class="icon"></i></button>
      </div>
    </div>`;
    })
    .join("");
  refreshIcons();
}

function openCreateTemplate() {
  STATE.editingTemplate = null;
  STATE.discoveredCols = [];
  document.getElementById("templatePanelTitle").textContent = "New Template";
  document.getElementById("tplName").value = "";
  document.getElementById("tplColumns").innerHTML = "";
  document.getElementById("tplDedupCols").innerHTML = "";
  document.getElementById("tplSortCol").innerHTML =
    '<option value="">-- No Sorting --</option>';
  document.getElementById("tplSortOrder").value = "asc";
  document.getElementById("discoveredColsArea").innerHTML = "";

  document.getElementById("scanStatus").textContent =
    "Choose a method above to detect columns from your data";
  document.getElementById("colCount").textContent = "0";
  STATE.currentTemplateLinks = [];
  STATE._scannedInputFileNames = [];
  STATE.selectedTemplatePages = [];
  renderTemplateLinkList();
  addTplColumn();
  navigate("template");
  openPanel("templatePanel");
}

function editTemplate(i) {
  STATE.editingTemplate = i;
  STATE.discoveredCols = [];
  STATE.selectedTemplatePages = [];
  const t = STATE.templates[i];
  document.getElementById("templatePanelTitle").textContent = "Edit Template";
  document.getElementById("tplName").value = t.name;
  document.getElementById("tplColumns").innerHTML = "";
  document.getElementById("discoveredColsArea").innerHTML = "";
  document.getElementById("scanStatus").textContent =
    "Choose a method above to detect columns";
  t.columns.forEach((col) => addTplColumn(col));
  refreshDedupCols(t.dedupCols || []);
  refreshSortCols(t.sortCol || "");
  document.getElementById("tplSortOrder").value = t.sortOrder || "asc";
  STATE.currentTemplateLinks = t.googleLinks || [];
  STATE.selectedTemplatePages = t.pages || [];
  renderTemplateLinkList();
  updateColCount();
  openPanel("templatePanel");
}

function addTplColumn(existing = null) {
  const container = document.getElementById("tplColumns");
  const idx = container.children.length + 1;
  const div = document.createElement("div");
  div.className = "col-card";
  const activeCodes =
    existing && existing.fmt
      ? existing.fmt.trim().split(/\s+/).filter(Boolean)
      : [];
  const existingDict = existing && existing.dict ? existing.dict : null;
  div.innerHTML = `
    <div class="col-card-header"><span class="col-num">Column ${idx}</span><button class="btn-icon danger" onclick="removeColCard(this)">✕</button></div>
    <div class="col-grid">
      <div><label class="form-label">Output Column Name</label><input type="text" class="input tpl-col-name" placeholder="e.g. Phone" value="${existing ? esc(existing.name) : ""}" onchange="refreshDedupCols();updateColCount();" oninput="updateColCount()"/></div>
      <div><label class="form-label">Source Column(s)</label>
        <div class="file-multiselect-wrap">
          <button type="button" class="file-multiselect-btn" onclick="toggleFmsDropdown(this)"><span class="fms-placeholder">Pick from detected columns…</span><span class="arrow">▼</span></button>
          <div class="file-multiselect-dropdown"></div>
        </div>
        <div class="fms-selected-tags"></div>
        <input type="text" class="input tpl-col-src" placeholder="or type: phone OR [phone,mobile]" value="${existing ? esc(existing.src || "") : ""}" style="margin-top:6px" oninput="syncFmsTags(this)"/>
      </div>
    </div>
    <div class="fmt-section"><div class="fmt-section-label"> Format Options</div>${buildFmtGroups(activeCodes)}</div>
    <div class="col-preview">
      <span class="preview-label">Active formats:</span>
      <div class="fmt-chips active-chips">${activeCodes.map((c) => fmtChip(c)).join("") || '<span style="color:var(--text3);font-size:11px">none</span>'}</div>
      <input type="hidden" class="tpl-col-fmt" value="${activeCodes.join(" ")}"/>
    </div>
    <div class="dict-editor" style="display:none"></div>
  `;
  div
    .querySelectorAll(".fmt-btn")
    .forEach((btn) => btn.addEventListener("click", () => toggleFmt(btn, div)));
  initFmsDropdown(div, existing ? existing.src || "" : "");
  if (existingDict)
    setTimeout(
      () => restoreDictEditor(div, existingDict, activeCodes.includes("q")),
      50,
    );
  refreshDedupCols();
  refreshSortCols();
  updateColCount();
  container.appendChild(div);
}

function removeColCard(btn) {
  btn.closest(".col-card").remove();
  renumberCols();
  refreshDedupCols();
  updateColCount();
}
function renumberCols() {
  document.querySelectorAll(".col-card").forEach((c, i) => {
    const n = c.querySelector(".col-num");
    if (n) n.textContent = `Column ${i + 1}`;
  });
}
function updateColCount() {
  document.getElementById("colCount").textContent =
    document.querySelectorAll(".col-card").length;
}

function buildFmtGroups(activeCodes) {
  return ["text", "phone", "time", "clean", "special"]
    .map((cat) => {
      const codes = FORMAT_CODES.filter((f) => f.cat === cat);
      return `<div style="margin-bottom:8px"><div class="fmt-section-label" style="font-size:10px;margin-bottom:4px">${CAT_LABELS[cat]}</div><div class="fmt-btn-group">${codes.map((f) => `<button class="fmt-btn cat-${cat} ${activeCodes.includes(f.code) ? "active" : ""}" data-code="${f.code}">${f.label}</button>`).join("")}</div></div>`;
    })
    .join("");
}

function fmtChip(code) {
  const f = FORMAT_CODES.find((x) => x.code === code);
  return `<span class="fmt-chip">${f ? f.label : code}<span class="fmt-chip-x" data-code="${code}" onclick="removeFmtChip(this)">×</span></span>`;
}

function toggleFmt(btn, colCard) {
  const code = btn.dataset.code;
  const fmtInput = colCard.querySelector(".tpl-col-fmt");
  let codes = fmtInput.value
    ? fmtInput.value.trim().split(/\s+/).filter(Boolean)
    : [];
  if (btn.classList.contains("active")) {
    codes = codes.filter((c) => c !== code);
    btn.classList.remove("active");
    if (code === "k" || code === "q") {
      codes = codes.filter((c) => c !== "k" && c !== "q");
      colCard
        .querySelectorAll('.fmt-btn[data-code="k"],.fmt-btn[data-code="q"]')
        .forEach((b) => b.classList.remove("active"));
      hideDictEditor(colCard);
    }
  } else {
    if (code === "k" || code === "q") {
      codes = codes.filter((c) => c !== "k" && c !== "q");
      colCard
        .querySelectorAll('.fmt-btn[data-code="k"],.fmt-btn[data-code="q"]')
        .forEach((b) => b.classList.remove("active"));
      codes.push(code);
      btn.classList.add("active");
      showDictEditor(colCard, code === "q");
    } else {
      if (!codes.includes(code)) codes.push(code);
      btn.classList.add("active");
    }
  }
  fmtInput.value = codes.join(" ");
  updateChips(colCard, codes);
}

function removeFmtChip(x) {
  const code = x.dataset.code;
  const colCard = x.closest(".col-card");
  const fmtInput = colCard.querySelector(".tpl-col-fmt");
  let codes = fmtInput.value
    ? fmtInput.value.trim().split(/\s+/).filter(Boolean)
    : [];
  codes = codes.filter((c) => c !== code);
  fmtInput.value = codes.join(" ");
  const btn = colCard.querySelector(`.fmt-btn[data-code="${code}"]`);
  if (btn) btn.classList.remove("active");
  if (code === "k" || code === "q") hideDictEditor(colCard);
  updateChips(colCard, codes);
}

function updateChips(colCard, codes) {
  const el = colCard.querySelector(".active-chips");
  el.innerHTML = codes.length
    ? codes.map((c) => fmtChip(c)).join("")
    : '<span style="color:var(--text3);font-size:11px">none</span>';
  el.querySelectorAll(".fmt-chip-x").forEach(
    (x) => (x.onclick = () => removeFmtChip(x)),
  );
}

function refreshDedupCols(selected = null) {
  const cols = [...document.querySelectorAll(".tpl-col-name")]
    .map((i) => i.value.trim())
    .filter(Boolean);
  const current =
    selected ||
    [...document.querySelectorAll("#tplDedupCols input:checked")].map(
      (i) => i.value,
    );
  document.getElementById("tplDedupCols").innerHTML = cols.length
    ? `<div class="dedup-pills">${cols.map((c) => `<label class="dedup-pill"><input type="checkbox" value="${esc(c)}" ${current.includes(c) ? "checked" : ""}/> ${esc(c)}</label>`).join("")}</div>`
    : '<span style="font-size:12px;color:var(--text3)">Add output columns above first</span>';
  refreshSortCols();
}
function refreshSortCols(selectedCol = null) {
  const cols = [...document.querySelectorAll(".tpl-col-name")]
    .map((i) => i.value.trim())
    .filter(Boolean);
  const sel = document.getElementById("tplSortCol");
  if (!sel) return;
  const currentVal = selectedCol || sel.value;
  sel.innerHTML =
    '<option value="">-- No Sorting --</option>' +
    cols
      .map(
        (c) =>
          `<option value="${esc(c)}" ${c === currentVal ? "selected" : ""}>${esc(c)}</option>`,
      )
      .join("");
}

function saveTemplate() {
  const name = document.getElementById("tplName").value.trim();
  if (!name) {
    toast("Template name required", "error");
    return;
  }
  const exists = STATE.templates.some(
    (t, idx) =>
      t.name.toLowerCase() === name.toLowerCase() &&
      idx !== STATE.editingTemplate,
  );
  if (exists) {
    toast("A template with this name already exists", "error");
    return;
  }
  const cols = [...document.querySelectorAll(".col-card")]
    .map((el) => {
      const fmt = el.querySelector(".tpl-col-fmt").value.trim();
      const col = {
        name: el.querySelector(".tpl-col-name").value.trim(),
        src: el.querySelector(".tpl-col-src").value.trim(),
        fmt,
      };
      const dictEditor = el.querySelector(".dict-editor");
      if (dictEditor && dictEditor.style.display !== "none")
        col.dict = readDictEditor(el);
      return col;
    })
    .filter((c) => c.name);
  if (!cols.length) {
    toast("Add at least one column", "error");
    return;
  }
  const dedupCols = [
    ...document.querySelectorAll("#tplDedupCols input:checked"),
  ].map((i) => i.value);
  const sortCol = document.getElementById("tplSortCol").value;
  const sortOrder = document.getElementById("tplSortOrder").value;
  const tpl = {
    name,
    columns: cols,
    dedupCols,
    sortCol,
    sortOrder,
    googleLinks: STATE.currentTemplateLinks,
    pages: STATE.selectedTemplatePages || [],
  };
  if (STATE.editingTemplate !== null) {
    STATE.templates[STATE.editingTemplate] = tpl;
    toast("Template updated", "success");
  } else {
    STATE.templates.push(tpl);
    toast("Template saved", "success");
  }
  save();
  closePanel("templatePanel");
  renderTemplateList();

  const jsonObj = templateToJsonFile(tpl);
  const filename =
    name
      .replace(/[^\w\-\s]/g, "_")
      .replace(/\s+/g, " ")
      .trim() + ".json";
  if (STATE.folderHandle) {
    writeToLinkedFolder("templates", filename, jsonObj).then((ok) => {
      if (ok) toast(`Saved to templates/${filename}`, "success");
      else {
        if (window.confirm(`Download "${filename}" to place in templates/?`))
          downloadTemplateJson(tpl);
      }
    });
    // Move scanned input files into template-named subfolder
    if (
      STATE._scannedInputFileNames &&
      STATE._scannedInputFileNames.length > 0
    ) {
      moveInputFilesToTemplateFolder(name);
    }
  } else {
    setTimeout(() => {
      if (window.confirm(`Download "${filename}" for templates/ folder?`))
        downloadTemplateJson(tpl);
    }, 400);
  }
}

async function deleteTemplate(i) {
  const tpl = STATE.templates[i];
  confirmDialog("Delete Template", `Delete "${tpl.name}"?`, async () => {
    if (STATE.folderHandle) {
      try {
        const tplDir = await STATE.folderHandle.getDirectoryHandle("templates");
        const filename = tpl.name.replace(/[^\w\-]/g, "_") + ".json";
        await tplDir.removeEntry(filename);
      } catch (e) {
        console.warn("File delete failed:", e);
      }
    }
    STATE.templates.splice(i, 1);
    save();
    renderTemplateList();
    toast("Deleted");
  });
}
function templateToJson(tpl) {
  const columns = tpl.columns.map((col) => {
    const tokens = [];
    tokens.push(col.src || "0");
    if (col.fmt) tokens.push(...col.fmt.split(/\s+/).filter(Boolean));
    const rule = [col.name, ...tokens];
    if (col.dict && Object.keys(col.dict).length) rule.push(col.dict);
    return rule;
  });
  return {
    name: tpl.name,
    columns,
    unique_columns: tpl.dedupCols,
    sort_column: tpl.sortCol || "",
    sort_order: tpl.sortOrder || "asc",
    google_links: tpl.googleLinks || [],
    pages: tpl.pages || [],
  };
}

const templateToJsonFile = templateToJson;
function downloadTemplateJson(tpl) {
  const json = templateToJson(tpl);
  const blob = new Blob([JSON.stringify(json, null, 2)], {
    type: "application/json",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = tpl.name + ".json";
  a.click();
  toast(`Downloaded ${tpl.name}.json`, "success");
}

// ─── SOURCE GROUPS ─────────────────────────────────────────────────────────
function renderGroupList(query = "") {
  const el = document.getElementById("groupList");
  const items = query
    ? STATE.groups.filter((g) =>
        g.name.toLowerCase().includes(query.toLowerCase()),
      )
    : STATE.groups;
  if (!items.length) {
    el.innerHTML = query
      ? '<div class="empty-state">No groups matching "' + esc(query) + '"</div>'
      : '<div class="empty-state"><div class="es-icon"><i data-lucide="package"></i></div><strong>No source groups yet</strong></div>';
    return;
  }
  el.innerHTML = items
    .map((g, idx) => {
      const i = STATE.groups.indexOf(g);
      return `
    <div class="list-card">
      <div class="lc-left"><div class="lc-icon blue"><i data-lucide="package"></i></div><div><div class="lc-title">${esc(g.name)}</div><div style="display:flex;align-items:center;gap:8px;margin-top:4px"><div class="lc-sub">${g.sources.length} source(s)</div><div class="xlsx-stat-pill" style="font-size:10px">${esc(g.templateName || "No Template")}</div></div></div></div>
      <div class="lc-actions"><button class="btn-icon" onclick="editGroup(${i})"><i data-lucide="edit-3" class="icon"></i></button><button class="btn-icon danger" onclick="deleteGroup(${i})"><i data-lucide="trash-2" class="icon"></i></button></div>
    </div>`;
    })
    .join("");
  refreshIcons();
}

async function openCreateGroup() {
  STATE.editingGroup = null;
  document.getElementById("groupPanelTitle").textContent = "New Source Group";
  document.getElementById("grpName").value = "";
  document.getElementById("grpTemplate").value = ""; // Added
  document.getElementById("grpSources").innerHTML = "";
  populateGrpTemplateDropdown();
  await autoDetectInputFiles();
  addGrpSource();
  navigate("sources");
  openPanel("groupPanel");
  refreshIcons(); // Added
}

async function editGroup(i) {
  STATE.editingGroup = i;
  const g = STATE.groups[i];
  document.getElementById("groupPanelTitle").textContent = "Edit Source Group";
  document.getElementById("grpName").value = g.name;
  document.getElementById("grpTemplate").value = g.templateName || ""; // Modified
  const list = document.getElementById("grpSources"); // Changed from grpSrcList to grpSources
  list.innerHTML = g.sources
    .map(
      (s) =>
        `<div class="smart-src-row">${buildSrcRow(s, "grp-src-label", "grp-src-path")}</div>`,
    )
    .join("");
  list
    .querySelectorAll(".smart-src-row")
    .forEach((row) => initSrcRow(row, ".grp-src-path", ".grp-src-label"));
  openPanel("groupPanel");
  refreshIcons(); // Added
}

function populateGrpTemplateDropdown(selectedName = null) {
  const select = document.getElementById("grpTemplate");
  if (!select) return;
  if (!STATE.templates.length) {
    select.innerHTML =
      '<option value="" disabled ' +
      (!selectedName ? "selected" : "") +
      ">No templates found - create one first</option>";
    return;
  }
  select.innerHTML =
    '<option value="" disabled ' +
    (!selectedName ? "selected" : "") +
    ">Select a template...</option>" +
    STATE.templates
      .map(
        (t) =>
          `<option value="${esc(t.name)}" ${t.name === selectedName ? "selected" : ""}>${esc(t.name)}</option>`,
      )
      .join("");
}

function addGrpSource(existing = null) {
  const div = document.createElement("div");
  div.className = "smart-src-row";
  div.innerHTML = buildSrcRow(existing, "grp-src-label", "grp-src-path");
  document.getElementById("grpSources").appendChild(div);
  initSrcRow(div, ".grp-src-path", ".grp-src-label");
  refreshIcons();
}

function saveGroup() {
  const name = document.getElementById("grpName").value.trim();
  const templateName = document.getElementById("grpTemplate").value;
  if (!name) {
    toast("Group name required", "error");
    return;
  }
  const exists = STATE.groups.some(
    (g, idx) =>
      g.name.toLowerCase() === name.toLowerCase() && idx !== STATE.editingGroup,
  );
  if (exists) {
    toast("A source group with this name already exists", "error");
    return;
  }
  if (!templateName) {
    toast("Please select a template", "error");
    return;
  }

  const sources = [
    ...document.getElementById("grpSources").querySelectorAll(".smart-src-row"),
  ]
    .map((el) => ({
      label: el.querySelector(".grp-src-label")?.value.trim() || "",
      path: el.querySelector(".grp-src-path")?.value.trim() || "",
      pages: el.querySelector(".src-pages-input")?.value.trim() || "",
    }))
    .filter((s) => s.path);
  if (!sources.length) {
    toast("Add at least one source", "error");
    return;
  }
  const grp = {
    name,
    templateName,
    sources,
    created: new Date().toISOString(),
  };
  if (STATE.editingGroup !== null) {
    STATE.groups[STATE.editingGroup] = grp;
    toast("Group updated", "success");
  } else {
    STATE.groups.push(grp);
    toast("Group saved", "success");
  }
  save();
  closePanel("groupPanel");
  renderGroupList();
  if (STATE.folderHandle)
    writeToLinkedFolder(
      "source_groups",
      name.replace(/[^\w\-]/g, "_") + ".json",
      grp,
    );
}

async function deleteGroup(i) {
  const grp = STATE.groups[i];
  if (!grp) {
    toast("Group index error", "error");
    return;
  }
  confirmDialog("Delete Group", `Delete "${grp.name}"?`, async () => {
    if (STATE.folderHandle) {
      try {
        const subDir =
          await STATE.folderHandle.getDirectoryHandle("source_groups");
        await subDir.removeEntry(grp.name.replace(/[^\w\-]/g, "_") + ".json");
        toast("File deleted from disk");
      } catch (e) {
        console.warn("Disk delete failed:", e);
      }
    }
    STATE.groups.splice(i, 1);
    save();
    renderGroupList();
    toast("Group removed from list");
  });
}

// ─── CAMPAIGNS ─────────────────────────────────────────────────────────────
function renderCampaignList(query = "") {
  const el = document.getElementById("campaignList");
  const items = query
    ? STATE.campaigns.filter((c) =>
        c.name.toLowerCase().includes(query.toLowerCase()),
      )
    : STATE.campaigns;
  if (!items.length) {
    el.innerHTML = query
      ? '<div class="empty-state">No campaigns matching "' +
        esc(query) +
        '"</div>'
      : '<div class="empty-state"><div class="es-icon"><i data-lucide="target"></i></div><strong>No campaigns yet</strong>Create a campaign to bundle source groups into a single Excel with multiple tabs.</div>';
    return;
  }
  el.innerHTML = items
    .map((c, idx) => {
      const i = STATE.campaigns.indexOf(c);
      return `
    <div class="list-card">
      <div class="lc-left" style="min-width:0;flex:1">
        <div class="lc-icon purple"><i data-lucide="target"></i></div>
        <div style="min-width:0;flex:1">
          <div class="lc-title">${esc(c.name)}</div>
          <div class="lc-sub">${c.groups.length} group(s) · ${c.groups.map((g) => esc(g)).join(", ")}</div>
        </div>
      </div>
      <div class="lc-actions">
        <button class="btn-icon" onclick="editCampaign(${i})"><i data-lucide="edit-3" class="icon"></i></button>
        <button class="btn-icon danger" onclick="deleteCampaign(${i})"><i data-lucide="trash-2" class="icon"></i></button>
      </div>
    </div>`;
    })
    .join("");
  refreshIcons();
}

function openCreateCampaign() {
  STATE.editingCampaign = null;
  document.getElementById("campaignPanelTitle").textContent = "New Campaign";
  document.getElementById("campName").value = "";
  renderCampaignGroupSelector();
  openPanel("campaignPanel");
}

function renderCampaignGroupSelector(selectedGroupNames = []) {
  const el = document.getElementById("campGroupList");
  if (!STATE.groups.length) {
    el.innerHTML =
      '<div style="color:var(--text3);font-size:12px;text-align:center;padding:20px">No source groups available. Create one first!</div>';
    return;
  }
  el.innerHTML = STATE.groups
    .map(
      (g) => `
    <label class="camp-grp-check">
      <input type="checkbox" value="${esc(g.name)}" ${selectedGroupNames.includes(g.name) ? "checked" : ""}>
      <div class="camp-grp-name">${esc(g.name)}</div>
      <div class="camp-grp-meta">${g.sources.length} sources · ${esc(g.templateName || "No Template")}</div>
    </label>
  `,
    )
    .join("");
}

function saveCampaign() {
  const name = document.getElementById("campName").value.trim();
  if (!name) {
    toast("Name required", "error");
    return;
  }
  const exists = STATE.campaigns.some(
    (c, idx) =>
      c.name.toLowerCase() === name.toLowerCase() &&
      idx !== STATE.editingCampaign,
  );
  if (exists) {
    toast("A campaign with this name already exists", "error");
    return;
  }
  const selectedGroups = [
    ...document.querySelectorAll("#campGroupList input:checked"),
  ].map((i) => i.value);
  if (!selectedGroups.length) {
    toast("Select at least one group", "error");
    return;
  }
  const campaign = {
    name,
    groups: selectedGroups,
    created: new Date().toISOString(),
  };
  if (STATE.editingCampaign !== null) {
    STATE.campaigns[STATE.editingCampaign] = campaign;
    toast("Campaign updated", "success");
  } else {
    STATE.campaigns.push(campaign);
    toast("Campaign saved", "success");
  }
  save();
  closePanel("campaignPanel");
  renderCampaignList();
  // Save to folder if linked
  if (STATE.folderHandle)
    writeToLinkedFolder(
      "campaigns",
      name.replace(/[^\w\-]/g, "_") + ".json",
      campaign,
    );
}

function editCampaign(i) {
  STATE.editingCampaign = i;
  const c = STATE.campaigns[i];
  document.getElementById("campaignPanelTitle").textContent = "Edit Campaign";
  document.getElementById("campName").value = c.name;
  renderCampaignGroupSelector(c.groups);
  openPanel("campaignPanel");
}

async function deleteCampaign(i) {
  const camp = STATE.campaigns[i];
  if (!camp) {
    toast("Campaign index error", "error");
    return;
  }
  confirmDialog("Delete Campaign", `Delete "${camp.name}"?`, async () => {
    if (STATE.folderHandle) {
      try {
        const subDir = await STATE.folderHandle.getDirectoryHandle("campaigns");
        await subDir.removeEntry(camp.name.replace(/[^\w\-]/g, "_") + ".json");
        toast("File deleted from disk");
      } catch (e) {
        console.warn("Disk delete failed:", e);
      }
    }
    STATE.campaigns.splice(i, 1);
    save();
    renderCampaignList();
    toast("Campaign removed from list");
  });
}

// ─── CAMPAIGN EXCEL EXPORT ──────────────────────────────────────────────────
async function exportCampaignExcel(campIdx) {
  const campaign = STATE.campaigns[campIdx];
  if (!campaign) {
    toast("Campaign not found", "error");
    return;
  }

  // Load SheetJS
  await loadSheetJS();

  toast(`Building Excel for "${campaign.name}"...`, "info");

  // Create workbook
  const wb = XLSX.utils.book_new();
  let tabsAdded = 0;
  let errors = [];

  for (const groupName of campaign.groups) {
    // Find the source group
    const group = STATE.groups.find((g) => g.name === groupName);
    if (!group) {
      errors.push(`Group not found: ${groupName}`);
      continue;
    }

    // Find the template
    const template = STATE.templates.find((t) => t.name === group.templateName);
    if (!template) {
      errors.push(`Template not found for group: ${groupName}`);
      continue;
    }

    // Check if sources reference uploaded files (browser can't auto-load from path)
    // We'll try to load from folderHandle if available, otherwise create empty sheet with headers
    let sheetData = [template.columns.map((c) => c.name)]; // header row

    if (STATE.folderHandle) {
      // Try to read each source file from linked folder
      for (const src of group.sources) {
        const rows = await tryReadFileFromFolder(src.path, template);
        if (rows && rows.length > 0) sheetData.push(...rows);
      }
    }

    // If only header, add note row
    if (sheetData.length === 1) {
      sheetData.push(template.columns.map((c) => `[${c.src || "—"}]`));
    }

    // Sanitize sheet name (Excel max 31 chars, no special chars)
    const sheetName = groupName
      .replace(/[\\\/\?\*\[\]:]/g, "")
      .substring(0, 31);
    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    // Style header row (set column widths)
    const colWidths = template.columns.map((c) => ({
      wch: Math.max(c.name.length + 4, 14),
    }));
    ws["!cols"] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    applyCellStyles(ws, template);
    tabsAdded++;
  }

  if (tabsAdded === 0) {
    toast(
      "No tabs could be created. Check your groups and templates.",
      "error",
    );
    return;
  }

  // Write and download
  const filename =
    campaign.name.replace(/[^\w\-]/g, "_") +
    "_" +
    new Date().toISOString().slice(0, 10) +
    ".xlsx";
  XLSX.writeFile(wb, filename);
  toast(`Excel exported: ${filename} (${tabsAdded} tabs)`, "success");
  STATE.runs++;
  save();
  refreshHome();
}

function applyCellStyles(ws, template) {
  if (!ws || !template) return;
  const range = XLSX.utils.decode_range(ws["!ref"]);

  template.columns.forEach((col, colIdx) => {
    const fmt = (col.fmt || "").toLowerCase();
    let align = null;
    if (
      fmt.includes(" l ") ||
      fmt.startsWith("l ") ||
      fmt.endsWith(" l") ||
      fmt === "l"
    )
      align = "left";
    if (
      fmt.includes(" r ") ||
      fmt.startsWith("r ") ||
      fmt.endsWith(" r") ||
      fmt === "r"
    )
      align = "right";
    if (
      fmt.includes(" z ") ||
      fmt.startsWith("z ") ||
      fmt.endsWith(" z") ||
      fmt === "z"
    )
      align = "center";

    if (align) {
      for (let R = range.s.r; R <= range.e.r; ++R) {
        const addr = XLSX.utils.encode_cell({ r: R, c: colIdx });
        if (!ws[addr]) continue;
        if (!ws[addr].s) ws[addr].s = {};
        ws[addr].s.alignment = { horizontal: align, vertical: "center" };
      }
    }
  });
}

// ─── DATA FORMATTING ────────────────────────────────────────────────────────
function applyJSFormat(val, code, dict = null) {
  try {
    let d_input = null;
    if (val instanceof Date) d_input = val;
    else if (typeof val === "number" && val > 40000 && val < 60000)
      d_input = new Date((val - 25569) * 86400 * 1000);

    const s = String(val === undefined || val === null ? "" : val).trim();
    if (code === "a") {
      const now = new Date();
      const dd = String(now.getDate()).padStart(2, "0");
      const mm = String(now.getMonth() + 1).padStart(2, "0");
      const yyyy = now.getFullYear();
      return `${dd}-${mm}-${yyyy}`;
    }
    if (code === "b")
      return new Date().toLocaleTimeString("en-GB", {
        hour: "2-digit",
        minute: "2-digit",
      });
    if (code === "c")
      return new Date().toLocaleTimeString("en-GB", {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      });
    if (code === "d") return s.replace(/\D/g, "").slice(-10);
    if (code === "e") return "+91" + s;
    if (code === "f") return s.toUpperCase();
    if (code === "g") return s.toLowerCase();
    if (code === "h")
      return s
        .split(" ")
        .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
        .join(" ");
    if (code === "i") {
      const n = parseInt(s.replace(/\D/g, ""));
      return isNaN(n) ? "0" : n.toString();
    }
    if (code === "j") return s.replace(/-/g, "");
    if (code === "u") return s.replace(/_/g, "");
    if (code === "x") return s.replace(/\./g, "");
    if (code === "n") {
      if (!s && !d_input) return s;
      let d = d_input;
      if (!d || isNaN(d.getTime())) d = new Date(s);

      // Try DD-MM-YYYY HH:MM:SS or DD-MM-YYYY HH:MM
      if (isNaN(d.getTime())) {
        const m = s.match(
          /^(\d{2})[\/\-](\d{2})[\/\-](\d{4})(?:[T\s](\d{2}):(\d{2})(?::(\d{2}))?)?/,
        );
        if (m)
          d = new Date(
            `${m[3]}-${m[2]}-${m[1]}T${m[4] || "00"}:${m[5] || "00"}:${m[6] || "00"}`,
          );
      }
      if (isNaN(d.getTime())) return s;

      // Get IST time parts
      const ist = new Date(
        d.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }),
      );
      const dd = String(ist.getDate()).padStart(2, "0");
      const mm = String(ist.getMonth() + 1).padStart(2, "0");
      const yyyy = ist.getFullYear();
      const hh = String(ist.getHours()).padStart(2, "0");
      const min = String(ist.getMinutes()).padStart(2, "0");
      const sec = String(ist.getSeconds()).padStart(2, "0");

      // Detect input format and mirror it
      if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(s)) {
        // ISO with or without offset → ISO IST
        return `${yyyy}-${mm}-${dd}T${hh}:${min}:${sec}+05:30`;
      }
      if (/^\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}/.test(s)) {
        return `${yyyy}-${mm}-${dd} ${hh}:${min}:${sec}`;
      }
      if (/^\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}/.test(s)) {
        return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
      }
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        return `${yyyy}-${mm}-${dd}`;
      }
      if (/^\d{2}[\/\-]\d{2}[\/\-]\d{4}\s\d{2}:\d{2}:\d{2}/.test(s)) {
        return `${dd}-${mm}-${yyyy} ${hh}:${min}:${sec}`;
      }
      if (/^\d{2}[\/\-]\d{2}[\/\-]\d{4}\s\d{2}:\d{2}/.test(s)) {
        return `${dd}-${mm}-${yyyy} ${hh}:${min}`;
      }
      if (/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/.test(s)) {
        return `${dd}-${mm}-${yyyy}`;
      }
      // Fallback
      return `${dd}-${mm}-${yyyy} ${hh}:${min}`;
    }
    if (code === "p") {
      if (!s && !d_input) return s;
      let d = d_input;
      if (!d || isNaN(d.getTime())) d = new Date(s);

      if (isNaN(d.getTime())) {
        const m = s.match(
          /^(\d{2})[\/\-](\d{2})[\/\-](\d{4})(?:[T\s](\d{2}):(\d{2})(?::(\d{2}))?)?/,
        );
        if (m)
          d = new Date(
            `${m[3]}-${m[2]}-${m[1]}T${m[4] || "00"}:${m[5] || "00"}:${m[6] || "00"}`,
          );
      }
      if (isNaN(d.getTime())) return s;

      const opts = { day: "numeric", month: "short", year: "numeric" };
      // If original string had time, include time
      if (
        s.includes(":") ||
        (d_input && d_input.getHours() + d_input.getMinutes() > 0)
      ) {
        opts.hour = "2-digit";
        opts.minute = "2-digit";
        opts.hour12 = false;
      }
      return d.toLocaleDateString("en-GB", opts).replace(",", "");
    }
    // Dictionary mapping (k and q)
    if ((code === "k" || code === "q") && dict) {
      const useDefault = code === "q";
      const normVal = s.toLowerCase().replace(/[\s_\-]/g, "");
      const matches = [];
      for (const [k, v] of Object.entries(dict)) {
        if (k === "__default__") continue;
        const normKey = k.toLowerCase().replace(/[\s_\-]/g, "");
        if (normVal.includes(normKey)) {
          matches.push(v);
        }
      }
      if (matches.length > 0) return [...new Set(matches)].join(", ");
      if (useDefault && dict["__default__"]) return dict["__default__"];
      return useDefault ? val : "";
    }
  } catch (e) {
    return val;
  }
  return val;
}

async function tryReadFileFromFolder(srcPath, template) {
  if (!STATE.folderHandle) return null;
  try {
    // srcPath might be "input/filename.csv" or just "filename.csv"
    const parts = srcPath.replace(/\\/g, "/").split("/");
    let handle = STATE.folderHandle;
    for (let i = 0; i < parts.length - 1; i++) {
      if (parts[i])
        handle = await handle.getDirectoryHandle(parts[i]).catch(() => null);
      if (!handle) return null;
    }
    const fname = parts[parts.length - 1];
    if (!fname) return null;
    const fileHandle = await handle.getFileHandle(fname).catch(() => null);
    if (!fileHandle) return null;
    const file = await fileHandle.getFile();
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(ab), { type: "array" });
    if (!wb.SheetNames.length) return null;

    const combined = [];

    // Iterate all sheets and map rows according to template, append results
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (!json || json.length < 2) continue; // no data
      const srcHeaders = (json[0] || []).map((h) =>
        String(h).toLowerCase().trim(),
      );
      const rows = json.slice(1);

      // Map rows to template columns and push as arrays (aoa)
      for (const row of rows) {
        const mapped = template.columns.map((col) => {
          const srcCol = col.src || "";
          const tokens = col.fmt ? col.fmt.split(/\s+/).filter(Boolean) : [];
          let val = "";

          if (srcCol === "0") {
            val = "";
          } else {
            const srcs = srcCol
              .replace(/^\[|\]$/g, "")
              .split(",")
              .map((s) => s.trim().toLowerCase());
            for (const s of srcs) {
              const idx = srcHeaders.indexOf(s);
              if (idx !== -1 && row[idx] !== undefined) {
                val = String(row[idx]);
                break;
              }
            }
          }

          // Apply formatters
          for (const token of tokens) {
            if (["k", "q"].includes(token)) continue; // Dict lookup not fully supported in JS yet
            val = applyJSFormat(val, token, col.dict);
          }
          return val;
        });
        combined.push(mapped);
      }
    }

    return combined.length ? combined : null;
  } catch (e) {
    console.warn("tryReadFileFromFolder error:", e);
    return null;
  }
}

// ─── SMART SOURCE ROW ──────────────────────────────────────────────────────
function detectSrcType(path) {
  if (!path) return "empty";
  if (path.includes("docs.google.com/spreadsheets")) return "gsheet";
  if (path.startsWith("http://") || path.startsWith("https://")) return "url";
  if (/\.(csv|xlsx|xls|tsv|txt|ods)$/i.test(path)) return "file";
  return "path";
}
function srcTypeBadge(type) {
  const map = {
    gsheet: [
      '<i data-lucide="file-spreadsheet" class="icon"></i>',
      "Google Sheet (Verified)",
      "badge-gsheet",
    ],
    url: ['<i data-lucide="globe" class="icon"></i>', "URL", "badge-url"],
    file: [
      '<i data-lucide="file-text" class="icon"></i>',
      "File",
      "badge-file",
    ],
    path: ['<i data-lucide="folder" class="icon"></i>', "Path", "badge-path"],
    empty: ["", "", ""],
  };
  const [icon, label, cls] = map[type] || map.empty;
  if (!label) return "";
  return `<span class="src-type-badge ${cls}">${icon} ${label}</span>`;
}

async function autoDetectInputFiles() {
  if (!STATE.folderHandle) return [];
  try {
    const inputDir = await STATE.folderHandle
      .getDirectoryHandle("input")
      .catch(() => null);
    if (!inputDir) return [];
    const files = [];
    // Scan root files
    for await (const [name, handle] of inputDir.entries()) {
      if (handle.kind === "file" && /\.(csv|xlsx|xls|tsv|txt)$/i.test(name))
        files.push(name);
      // Also scan subfolders
      if (handle.kind === "directory") {
        try {
          for await (const [fname, fh] of handle.entries()) {
            if (fh.kind === "file" && /\.(csv|xlsx|xls|tsv|txt)$/i.test(fname))
              files.push(name + "/" + fname);
          }
        } catch (e) {}
      }
    }
    STATE._inputFiles = files;
    return files;
  } catch (e) {
    console.warn(e);
    return [];
  }
}

async function scanInputFiles() {
  if (!STATE.folderHandle) {
    toast("Link a project folder first", "error");
    return;
  }
  try {
    const inputDir = await STATE.folderHandle.getDirectoryHandle("input", {
      create: true,
    });
    const result = { _root: [] };
    for await (const [name, handle] of inputDir.entries()) {
      if (
        handle.kind === "file" &&
        /\.(csv|xlsx|xls|tsv|txt|ods)$/i.test(name)
      ) {
        const file = await handle.getFile();
        result._root.push({
          name,
          size: file.size,
          lastModified: file.lastModified,
          folder: "_root",
        });
      }
      if (handle.kind === "directory") {
        result[name] = [];
        try {
          for await (const [fname, fh] of handle.entries()) {
            if (
              fh.kind === "file" &&
              /\.(csv|xlsx|xls|tsv|txt|ods)$/i.test(fname)
            ) {
              const file = await fh.getFile();
              result[name].push({
                name: fname,
                size: file.size,
                lastModified: file.lastModified,
                folder: name,
              });
            }
          }
        } catch (e) {}
      }
    }
    STATE._allInputFiles = result;
    return result;
  } catch (e) {
    console.warn(e);
    return { _root: [] };
  }
}

async function deleteInputFile(folder, fname) {
  if (!STATE.folderHandle) return;
  confirmDialog("Delete File", `Delete "${fname}"?`, async () => {
    try {
      const inputDir = await STATE.folderHandle.getDirectoryHandle("input");
      if (folder === "_root") {
        await inputDir.removeEntry(fname);
      } else {
        const subDir = await inputDir.getDirectoryHandle(folder);
        await subDir.removeEntry(fname);
      }
      toast("File deleted", "success");
      if (document.getElementById("page-input").classList.contains("active")) {
        refreshInputPage();
      } else {
        refreshPreview();
      }
    } catch (e) {
      toast("Delete failed: " + e.message, "error");
    }
  });
}

async function deleteAllInFolder(folder) {
  const folderName = folder === "_root" ? "input/" : folder;
  confirmDialog(
    "Delete All",
    `Permanently delete all files in "${folderName}"?`,
    async () => {
      try {
        const inputDir = await STATE.folderHandle.getDirectoryHandle("input");
        const targetDir =
          folder === "_root"
            ? inputDir
            : await inputDir.getDirectoryHandle(folder);

        const entries = [];
        for await (const entry of targetDir.values()) {
          if (entry.kind === "file") entries.push(entry.name);
        }

        if (!entries.length) {
          toast("Folder is already empty");
          return;
        }

        for (const name of entries) {
          await targetDir.removeEntry(name);
        }

        toast(
          `Deleted ${entries.length} file(s) from ${folderName}`,
          "success",
        );
        if (
          document.getElementById("page-input").classList.contains("active")
        ) {
          refreshInputPage();
        } else {
          refreshPreview();
        }
      } catch (e) {
        toast("Delete all failed: " + e.message, "error");
      }
    },
  );
}

async function handleInputFolderDrop(event, folder) {
  const SUPPORTED = [".csv", ".xlsx", ".xls", ".tsv", ".txt", ".ods"];

  function getExt(filename) {
    const dotIdx = filename.lastIndexOf(".");
    if (dotIdx <= 0) return "";
    return filename.slice(dotIdx).toLowerCase();
  }

  // ✅ Use .files (reliable for OS drags) with .items as fallback
  let files = [];
  if (event.dataTransfer.files && event.dataTransfer.files.length > 0) {
    files = [...event.dataTransfer.files];
  } else if (event.dataTransfer.items) {
    for (const item of event.dataTransfer.items) {
      if (item.kind === "file") {
        const f = item.getAsFile();
        if (f) files.push(f);
      }
    }
  }

  if (!files.length) {
    toast(
      "No files detected — try dragging directly from your file explorer",
      "error",
    );
    return;
  }

  if (!STATE.folderHandle) {
    toast("Link a project folder first", "error");
    return;
  }

  try {
    const inputDir = await STATE.folderHandle.getDirectoryHandle("input");
    const targetDir =
      folder === "_root"
        ? inputDir
        : await inputDir.getDirectoryHandle(folder, { create: true });

    let uploaded = 0;
    const skipped = [];

    for (const file of files) {
      const ext = getExt(file.name);
      if (!ext) {
        skipped.push(`${file.name} (no extension)`);
        continue;
      }
      if (!SUPPORTED.includes(ext)) {
        skipped.push(`${file.name} (${ext} not supported)`);
        continue;
      }

      const fileHandle = await targetDir.getFileHandle(file.name, {
        create: true,
      });
      const writable = await fileHandle.createWritable();
      await writable.write(file);
      await writable.close();
      uploaded++;
    }

    if (uploaded > 0) {
      toast(
        `Uploaded ${uploaded} file(s) to ${folder === "_root" ? "input/" : folder}`,
        "success",
      );
      if (document.getElementById("page-input").classList.contains("active")) {
        refreshInputPage();
      } else {
        refreshPreview();
      }
    }
    if (skipped.length > 0) {
      const names = skipped.slice(0, 2).join(", ");
      const extra = skipped.length > 2 ? ` +${skipped.length - 2} more` : "";
      toast(
        `Skipped: ${names}${extra} — use CSV, XLSX, XLS, TSV, TXT, ODS`,
        "error",
      );
    }
    if (uploaded === 0 && skipped.length === 0) {
      toast(
        "No valid files found — drag files directly from File Explorer",
        "error",
      );
    }
  } catch (e) {
    toast("Upload failed: " + e.message, "error");
  }
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}

function getBaseName(path) {
  return path.split(/[\/\\]/).pop();
}

function buildSrcRow(existing, labelClass, pathClass) {
  const label = existing ? esc(existing.label || "") : "";
  const path = existing ? esc(existing.path || "") : "";
  const type = detectSrcType(existing ? existing.path || "" : "");
  const badge = existing && existing.path ? srcTypeBadge(type) : "";
  const detectedFiles = STATE._inputFiles || [];
  const dropdownHtml = detectedFiles.length
    ? `<select class="input src-path-dropdown" onchange="syncPathDropdown(this)" style="margin-top:4px;font-size:11px"><option value="">-- Or select from input/ folder --</option>${detectedFiles.map((f) => `<option value="input/${esc(f)}" ${path === "input/" + f ? "selected" : ""}>${esc(f)}</option>`).join("")}</select>`
    : "";
  return `
    <div id="srcBadge_${Math.random().toString(36).slice(2)}">${badge}</div>
    <div class="src-row-grid">
      <div><label class="form-label">Label</label><input type="text" class="input ${labelClass}" placeholder="e.g. Facebook Jan" value="${label}"/></div>
      <div class="src-path-wrap">
        <label class="form-label">Source URL / File Path</label>
        <input type="text" class="input ${pathClass} src-path-input" placeholder="Paste Google Sheet URL or /path/to/file.xlsx" value="${path}"/>
        ${dropdownHtml}
        <div class="src-path-actions">
          <button type="button" class="src-helper-btn" onclick="pasteFromClipboard(this)"><i data-lucide="clipboard" class="icon"></i> Paste URL</button>
          <button type="button" class="src-helper-btn" onclick="clearSrcRow(this)"><i data-lucide="x" class="icon"></i> Clear</button>
        </div>
      </div>
      <div style="align-self:flex-start;padding-top:22px"><button class="btn-icon danger" onclick="this.closest('.smart-src-row').remove()"><i data-lucide="trash-2" class="icon"></i></button></div>
    </div>
  `;
}

async function fetchAndRenderPages(btn) {
  const row = btn.closest(".smart-src-row");
  const pathInput = row.querySelector(".src-path-input");
  const container = row.querySelector(".src-pages-container");
  const path = pathInput.value.trim();

  if (!path) {
    toast("Please provide a Source URL or File Path first.", "error");
    return;
  }

  const origHtml = btn.innerHTML;
  btn.innerHTML = '<i data-lucide="loader-2" class="icon spin"></i>';
  btn.disabled = true;
  refreshIcons();

  try {
    let sheetNames = [];
    if (path.startsWith("http")) {
      sheetNames = await fetchExternalSheetNames(path);
    } else {
      sheetNames = await readLocalSheetNames(path);
    }

    if (!sheetNames || sheetNames.length === 0) {
      container.innerHTML =
        '<div style="color:var(--red);">No pages found. Validate path.</div>';
    } else {
      // Create checkboxes
      container.innerHTML = sheetNames
        .map(
          (name) => `
        <label style="display: flex; align-items: center; gap: 4px; margin-bottom: 2px;">
          <input type="checkbox" value="${escAttr(name)}" onchange="updateSrcPagesInput(this)" />
          ${esc(name)}
        </label>
      `,
        )
        .join("");
      // Reset the hidden input so initially it counts as "all" unless user checks some
      row.querySelector(".src-pages-input").value = "";
    }
  } catch (e) {
    container.innerHTML = `<div style="color:var(--red);">Error: ${esc(e.message)}</div>`;
  } finally {
    btn.innerHTML = origHtml;
    btn.disabled = false;
    refreshIcons();
  }
}

function updateSrcPagesInput(cb) {
  const row = cb.closest(".smart-src-row");
  const container = row.querySelector(".src-pages-container");
  const hiddenInput = row.querySelector(".src-pages-input");

  const checked = Array.from(
    container.querySelectorAll('input[type="checkbox"]:checked'),
  ).map((el) => el.value);
  hiddenInput.value = checked.join(",");
}

async function fetchExternalSheetNames(url) {
  let fetchUrl = url;
  let isGSheet = false;
  let token = null;

  if (url.includes("docs.google.com/spreadsheets")) {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) {
      const sheetId = match[1];
      fetchUrl =
        "https://www.googleapis.com/drive/v3/files/" +
        sheetId +
        "/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      isGSheet = true;
      try {
        token = await getGoogleAccessToken();
      } catch (e) {}
    }
  }

  await loadSheetJS();
  let resp;
  if (isGSheet && token) {
    resp = await fetch(fetchUrl, {
      headers: { Authorization: "Bearer " + token },
    });
  } else {
    resp = await fetch(fetchUrl, { cache: "no-store" });
  }
  if (!resp.ok) throw new Error("Fetch failed");
  const ab = await resp.arrayBuffer();
  if (ab.byteLength < 5) return [];

  let wb;
  try {
    wb = XLSX.read(new Uint8Array(ab), { type: "array" });
  } catch (parseErr) {
    const text = new TextDecoder().decode(ab);
    wb = XLSX.read(text, { type: "string" });
  }
  return wb.SheetNames;
}

async function readLocalSheetNames(path) {
  if (!STATE.folderHandle) throw new Error("Folder not linked");
  const parts = path.replace(/\\/g, "/").split("/");
  let target = STATE.folderHandle;
  for (let i = 0; i < parts.length - 1; i++) {
    if (parts[i]) {
      target = await target.getDirectoryHandle(parts[i]).catch(() => null);
      if (!target) throw new Error("Dir not found");
    }
  }
  const fileHandle = await target
    .getFileHandle(parts[parts.length - 1])
    .catch(() => null);
  if (!fileHandle) throw new Error("File not found");

  const file = await fileHandle.getFile();
  const ab = await file.arrayBuffer();
  await loadSheetJS();
  const wb = XLSX.read(new Uint8Array(ab), { type: "array" });
  return wb.SheetNames;
}

function syncPathDropdown(sel) {
  const wrap = sel.closest(".src-path-wrap");
  const input = wrap.querySelector(".src-path-input");
  const row = sel.closest(".smart-src-row");
  const ls = row.querySelector(".grp-src-label")
    ? ".grp-src-label"
    : ".ext-file-label";
  if (sel.value) {
    input.value = sel.value;
    updateSrcRowUI(row, input, ls);
  }
  // Keep dropdown open - don't reset selection, user closes manually
}
function initSrcRow(rowEl, pathSel, labelSel) {
  const pathInput = rowEl.querySelector(pathSel);
  if (!pathInput) return;
  pathInput.addEventListener("input", () =>
    updateSrcRowUI(rowEl, pathInput, labelSel),
  );
  if (pathInput.value) updateSrcRowUI(rowEl, pathInput, labelSel);
}
function updateSrcRowUI(rowEl, pathInput, labelSel) {
  const val = pathInput.value.trim();
  const type = detectSrcType(val);
  const badgeEl = rowEl.querySelector('[id^="srcBadge_"]');
  if (badgeEl) badgeEl.innerHTML = val ? srcTypeBadge(type) : "";
  const labelInput = rowEl.querySelector(labelSel);
  if (labelInput && !labelInput.value) {
    if (type === "gsheet") labelInput.value = "Google Sheet";
    else if (type === "file" || type === "path")
      labelInput.value = val
        .split(/[\/\\]/)
        .pop()
        .replace(/\.[^.]+$/, "");
  }
}
async function pasteFromClipboard(btn) {
  try {
    const text = await navigator.clipboard.readText();
    const pathInput = btn
      .closest(".src-path-wrap")
      .querySelector(".src-path-input");
    const row = btn.closest(".smart-src-row");
    const ls = row.querySelector(".grp-src-label")
      ? ".grp-src-label"
      : ".ext-file-label";
    pathInput.value = text.trim();
    updateSrcRowUI(row, pathInput, ls);
    toast("Pasted!");
  } catch (e) {
    toast("Allow clipboard or paste manually", "error");
  }
}
function clearSrcRow(btn) {
  const wrap = btn.closest(".src-path-wrap");
  wrap.querySelector(".src-path-input").value = "";
  const row = btn.closest(".smart-src-row");
  const b = row.querySelector('[id^="srcBadge_"]');
  if (b) b.innerHTML = "";
}

// ─── FOLDER SCAN ───────────────────────────────────────────────────────────
async function scanFolder() {
  if (!STATE.folderHandle) {
    toast("Link a project folder first", "error");
    return;
  }

  const scanBtn = document.getElementById("scanBtn");
  const statusEl = document.getElementById("scanStatus");
  if (!scanBtn) return;

  // Show loader
  scanBtn.innerHTML =
    '<i data-lucide="loader-2" class="icon spin"></i> Scanning...';
  scanBtn.disabled = true;
  refreshIcons();

  STATE._scannedInputFileNames = [];

  try {
    await loadSheetJS();

    const allCols = new Set();
    const filesSeen = [];
    const SUPPORTED = [".csv", ".xlsx", ".xls", ".tsv", ".txt"];
    const dirsToSearch = [STATE.folderHandle];
    let inputDirHandle = null;

    // Detect input folder
    for await (const [name, handle] of STATE.folderHandle.entries()) {
      if (handle.kind === "directory" && name === "input") {
        dirsToSearch.push(handle);
        inputDirHandle = handle;
      }
    }

    // Scan files
    for (const dir of dirsToSearch) {
      for await (const [fname, fh] of dir.entries()) {
        if (fh.kind !== "file") continue;

        const ext = fname.toLowerCase().slice(fname.lastIndexOf("."));
        if (!SUPPORTED.includes(ext)) continue;

        try {
          statusEl.textContent = `Reading ${fname}...`;

          const file = await fh.getFile();
          const { headers } = await parseFileHeaders(file);

          headers.forEach((h) => allCols.add(h));
          filesSeen.push({ name: fname, count: headers.length });

          if (dir === inputDirHandle) {
            STATE._scannedInputFileNames.push(fname);
          }
        } catch (e) {
          console.warn(`Error reading ${fname}`, e);
        }
      }
    }

    // Restore button properly (IMPORTANT FIX)
    scanBtn.innerHTML =
      '<i data-lucide="search" class="icon"></i> Scan Input Folder';
    scanBtn.disabled = false;
    refreshIcons();

    if (allCols.size === 0) {
      statusEl.textContent = "No headers detected";
      toast("No columns found", "error");
    } else {
      showDiscoveredCols([...allCols], filesSeen);

      statusEl.innerHTML = `<i data-lucide="check" class="icon"></i> ${allCols.size} column(s) from ${filesSeen.length} file(s)`;

      refreshIcons();
    }
  } catch (e) {
    console.error(e);

    scanBtn.innerHTML =
      '<i data-lucide="search" class="icon"></i> Scan Input Folder';
    scanBtn.disabled = false;
    refreshIcons();

    statusEl.textContent = "Scan failed";
  }
}

async function moveInputFilesToTemplateFolder(templateName) {
  if (!STATE.folderHandle) return;
  const filesToMove = STATE._scannedInputFileNames || [];
  if (!filesToMove.length) return;
  try {
    const inputDir = await STATE.folderHandle.getDirectoryHandle("input", {
      create: true,
    });
    const safeName = templateName.replace(/[^\w\-]/g, "_");
    const subDir = await inputDir.getDirectoryHandle(safeName, {
      create: true,
    });
    let moved = 0;
    for (const fname of filesToMove) {
      try {
        // Read the original file
        const srcHandle = await inputDir.getFileHandle(fname);
        const file = await srcHandle.getFile();
        const content = await file.arrayBuffer();
        // Write to subfolder
        const destHandle = await subDir.getFileHandle(fname, {
          create: true,
        });
        const writable = await destHandle.createWritable();
        await writable.write(content);
        await writable.close();
        // Delete original
        await inputDir.removeEntry(fname);
        moved++;
      } catch (e) {
        console.warn("Could not move file:", fname, e);
      }
    }
    if (moved > 0) {
      toast(`Moved ${moved} file(s) to input/${safeName}/`, "success");
    }
    STATE._scannedInputFileNames = [];
  } catch (e) {
    console.warn("moveInputFiles error:", e);
  }
}

async function readSampleFile(input) {
  const files = [...input.files];
  if (!files.length) return;
  const statusEl = document.getElementById("scanStatus");
  statusEl.textContent = `Reading ${files.length} file(s)...`;
  try {
    await loadSheetJS();
    const allCols = new Set();
    const filesSeen = [];
    for (const file of files) {
      try {
        const { headers, fileName } = await parseFileHeaders(file);
        headers.forEach((h) => allCols.add(h));
        filesSeen.push({ name: fileName, count: headers.length });

        // Auto-upload to input/ if project linked
        if (STATE.folderHandle) {
          try {
            const inputDir = await STATE.folderHandle.getDirectoryHandle(
              "input",
              { create: true },
            );
            const fileHandle = await inputDir.getFileHandle(fileName, {
              create: true,
            });
            const writable = await fileHandle.createWritable();
            await writable.write(file);
            await writable.close();
            if (!STATE._scannedInputFileNames.includes(fileName)) {
              STATE._scannedInputFileNames.push(fileName);
            }
          } catch (err) {
            console.warn("Auto-upload failed:", err);
          }
        }
      } catch (e) {
        toast(`Could not read ${file.name}`, "error");
      }
    }
    if (allCols.size > 0) {
      showDiscoveredCols([...allCols], filesSeen);
      statusEl.textContent = `✅ ${allCols.size} column(s) from ${filesSeen.length} file(s)`;
    } else statusEl.textContent = "No headers found";
  } catch (e) {
    statusEl.textContent = "Read failed";
  }
  input.value = "";
}

async function fetchGoogleSheetCols() {
  const urlInput = document.getElementById("gsheetUrlInput");
  const url = urlInput ? urlInput.value.trim() : "";
  if (!url) {
    toast("Paste a Google Sheets URL first", "error");
    return;
  }
  const statusEl = document.getElementById("scanStatus");
  statusEl.textContent = "Authenticating with Google...";

  try {
    await loadSheetJS();
    const token = await getGoogleAccessToken();
    statusEl.textContent = "Fetching protected Google Sheet...";

    let fetchUrl = url;
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) throw new Error("Invalid Google Sheets URL");
    const sheetId = match[1];
    fetchUrl = `https://www.googleapis.com/drive/v3/files/${sheetId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;

    const resp = await fetch(fetchUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!resp.ok) {
      if (resp.status === 403)
        throw new Error(
          "Permission Denied — Share the sheet with " + GOOGLE_SA_EMAIL,
        );
      throw new Error("Fetch failed: " + resp.statusText);
    }
    const ab = await resp.arrayBuffer();
    if (ab.byteLength < 5) throw new Error("Empty response");
    const wb = XLSX.read(new Uint8Array(ab), { type: "array" });

    // Collect headers from ALL sheets (union, preserving first-seen order)
    const seen = new Set();
    const headers = [];
    const filesSeen = [];
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const row = json && json[0] ? json[0] : [];
      let sheetColCount = 0;
      const pageCols = [];
      for (const h of row) {
        const norm = String(h || "")
          .toLowerCase()
          .trim();
        if (norm && !seen.has(norm)) {
          seen.add(norm);
          headers.push(norm);
          sheetColCount++;
        }
        if (norm) pageCols.push(norm);
      }
      if (row.length > 0)
        filesSeen.push({
          name: "Page: " + name,
          count: row.length,
          columns: pageCols,
        });
    }

    if (headers.length) {
      showDiscoveredCols(headers, filesSeen);
      statusEl.textContent = `✅ ${headers.length} column(s) from ${wb.SheetNames.length} page(s)`;
      if (!STATE.currentTemplateLinks.includes(url)) {
        STATE.currentTemplateLinks.push(url);
        renderTemplateLinkList();
      }
      urlInput.value = "";
    } else statusEl.textContent = "No headers found";
  } catch (e) {
    statusEl.textContent = "Fetch failed";
    toast(e.message || "Could not fetch", "error");
  }
}

function renderTemplateLinkList() {
  const container = document.getElementById("tplLinkList");
  if (!container) return;
  if (!STATE.currentTemplateLinks.length) {
    container.innerHTML = "";
    return;
  }
  container.innerHTML = STATE.currentTemplateLinks
    .map((url, idx) => {
      let short = url;
      try {
        const match = url.match(/\/d\/([^\/]+)/);
        if (match) short = "Sheet: " + match[1].substring(0, 8) + "...";
      } catch (e) {}
      return `
      <div style="display:flex;align-items:center;justify-content:space-between;background:white;border:1px solid var(--border);border-radius:6px;padding:4px 8px;font-size:10px">
        <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:140px;color:var(--text2)" title="${esc(url)}">${esc(short)}</span>
        <button class="btn-icon danger" style="padding:2px;width:18px;height:18px" onclick="removeTemplateLink(${idx})">✕</button>
      </div>
    `;
    })
    .join("");
  refreshIcons();
}

function removeTemplateLink(idx) {
  STATE.currentTemplateLinks.splice(idx, 1);
  renderTemplateLinkList();
  toast("Link removed");
}

function showDiscoveredCols(cols, filesSeen = []) {
  STATE.discoveredCols = cols;
  STATE.discoveredFilesSeen = filesSeen;
  // Initialize selectedTemplatePages if not set OR empty (to ensure explicit save)
  if (
    !STATE.selectedTemplatePages ||
    STATE.selectedTemplatePages.length === 0
  ) {
    STATE.selectedTemplatePages = filesSeen.map((f) => {
      let n = f.name;
      if (n.startsWith("Page: ")) n = n.substring(6);
      return n;
    });
  }
  document.querySelectorAll(".col-card").forEach((card) => {
    const src = card.querySelector(".tpl-col-src");
    refreshFmsDropdown(card, src ? src.value : "");
  });
  const area = document.getElementById("discoveredColsArea");

  // Compute visible columns based on selected pages
  let visibleCols = cols;
  if (
    STATE.selectedTemplatePages.length > 0 &&
    filesSeen.some((f) => f.columns)
  ) {
    const allowedSet = new Set();
    filesSeen.forEach((f) => {
      let cleanName = f.name;
      if (cleanName.startsWith("Page: ")) cleanName = cleanName.substring(6);
      if (STATE.selectedTemplatePages.includes(cleanName) && f.columns) {
        f.columns.forEach((c) => allowedSet.add(c));
      }
    });
    visibleCols = cols.filter((c) => allowedSet.has(c));
  }

  const pagesHtml = filesSeen.length
    ? `<div style="margin-bottom:10px">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
          <span style="font-size:11px;font-weight:600;color:var(--text2)">Select Pages to Include:</span>
          <span style="font-size:10px;color:var(--text3);font-style:italic">(uncheck pages you don't need)</span>
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:6px">${filesSeen
          .map((f) => {
            let cleanName = f.name;
            if (cleanName.startsWith("Page: "))
              cleanName = cleanName.substring(6);
            const isChecked = STATE.selectedTemplatePages.includes(cleanName);
            return `<label style="display:flex;align-items:center;gap:4px;font-size:11px;background:white;border:1px solid ${isChecked ? "#a7f3d0" : "#e5e7eb"};border-radius:10px;padding:4px 10px;color:${isChecked ? "var(--green)" : "var(--text3)"};cursor:pointer;user-select:none;transition:all .15s">
            <input type="checkbox" value="${escAttr(cleanName)}" ${isChecked ? "checked" : ""} onchange="toggleTemplatePage(this)" style="accent-color:var(--green)" />
            📄 ${esc(f.name)} (${f.count} cols)
          </label>`;
          })
          .join("")}</div>
       </div>`
    : "";
  area.innerHTML = `<div class="discovered-cols-wrap"><div class="discovered-header"><div class="discovered-title"><i data-lucide="check" class="icon"></i> ${visibleCols.length} Column(s) Found</div><button class="btn btn-ghost btn-sm" onclick="autoPopulateFromCols()">Auto-add all \u2192</button></div>${pagesHtml}<div class="discovered-hint">Click a column chip to add it as an output column</div><div class="folder-col-chips">${visibleCols.map((c) => `<div class="col-chip" onclick="addTplColumnFromName('${escAttr(c)}')" title="Add as output column">+ ${esc(c)}</div>`).join("")}</div></div>`;
}

function toggleTemplatePage(cb) {
  const allCheckboxes = document.querySelectorAll(
    '#discoveredColsArea input[type="checkbox"]',
  );
  const checked = Array.from(allCheckboxes)
    .filter((el) => el.checked)
    .map((el) => el.value);

  STATE.selectedTemplatePages = checked;
  // Re-render with updated page selection
  showDiscoveredCols(STATE.discoveredCols, STATE.discoveredFilesSeen || []);
  refreshIcons();
}

function autoPopulateFromCols() {
  STATE.discoveredCols.forEach((c) => addTplColumnFromName(c));
  toast(`Added ${STATE.discoveredCols.length} columns`, "success");
}
function addTplColumnFromName(name) {
  addTplColumn({ name, src: name, fmt: "" });
  toast(`Added: ${name}`);
}

// ─── OUTPUT VIEWER ─────────────────────────────────────────────────────────
function refreshViewerPage() {
  if (STATE.folderHandle) scanOutputFolder(false);
  else {
    document.getElementById("outputFileSelector").innerHTML =
      '<div style="font-size:12px;color:var(--text3);text-align:center;padding:16px">Link a project folder to view output files.</div>';
  }
  refreshIcons();
}

async function scanOutputFolder(showLoading = true) {
  if (!STATE.folderHandle) {
    if (!showLoading) return;
    toast("Link a project folder first", "error");
    return;
  }
  const sidebar = document.getElementById("outputFileSelector");
  if (showLoading)
    sidebar.innerHTML =
      '<div style="padding:12px;font-size:12px;color:var(--text3);text-align:center">⏳ Scanning...</div>';
  try {
    const outDir = await STATE.folderHandle.getDirectoryHandle("output", {
      create: true,
    });
    const files = [];
    for await (const [name, handle] of outDir.entries()) {
      if (
        handle.kind === "file" &&
        (name.endsWith(".xlsx") || name.endsWith(".csv"))
      ) {
        const file = await handle.getFile();
        files.push({
          name,
          size: file.size,
          lastModified: file.lastModified,
          handle,
          isBaseline: false,
        });
      }
    }

    // Also scan baselines
    try {
      const templatesDir =
        await STATE.folderHandle.getDirectoryHandle("templates");
      const baselinesDir = await templatesDir.getDirectoryHandle("baselines");
      for await (const [name, handle] of baselinesDir.entries()) {
        if (
          handle.kind === "file" &&
          (name.endsWith(".xlsx") || name.endsWith(".csv"))
        ) {
          const file = await handle.getFile();
          files.push({
            name,
            size: file.size,
            lastModified: file.lastModified,
            handle,
            isBaseline: true,
          });
        }
      }
    } catch (e) {}

    if (!files.length) {
      sidebar.innerHTML =
        '<div style="padding:20px 10px;font-size:12px;color:var(--text3);text-align:center;border:1.5px dashed var(--border);border-radius:10px">No output files found.</div>';
      return;
    }
    files.sort((a, b) => b.lastModified - a.lastModified);
    STATE._outputFiles = files;
    const searchHtml = `<div class="search-wrap" style="max-width:none;margin-bottom:12px">
      <i data-lucide="search" class="search-icon"></i>
      <input type="text" class="input search-input" placeholder="Search files..." value="${escAttr(STATE.viewerQuery)}" oninput="searchViewerFiles(this.value)">
    </div>`;
    sidebar.innerHTML = searchHtml + '<div id="viewerTree"></div>';
    renderHierarchicalFileSelector(files, STATE.viewerQuery);
  } catch (e) {
    sidebar.innerHTML =
      '<div style="padding:12px;font-size:12px;color:var(--red);text-align:center">⚠ Scan failed</div>';
    console.error(e);
  }
}

function searchViewerFiles(query) {
  STATE.viewerQuery = query;
  renderHierarchicalFileSelector(STATE._outputFiles, query);
}

function renderHierarchicalFileSelector(files, query = "") {
  const el = document.getElementById("viewerTree");
  if (!el) return;
  const filtered = query
    ? files.filter((f) => f.name.toLowerCase().includes(query.toLowerCase()))
    : files;

  // Separate categories as requested
  const tree = {
    "Template Baselines": {},
    "Campaign Baselines": {},
    Templates: {},
    Campaigns: {},
    Other: [],
  };

  filtered.forEach((f) => {
    const base = f.name.replace(/\.[^.]+$/, "");
    const tpl = STATE.templates.find((t) => {
      const sn = t.name.replace(/[^\w\-]/g, "_");
      return (
        base === sn ||
        base === t.name ||
        base.startsWith(sn + "_") ||
        base.startsWith(t.name + "_")
      );
    });

    // Improved categorization
    let cat = "";
    let groupKey = "";

    if (tpl) {
      cat = f.isBaseline ? "Template Baselines" : "Templates";
      groupKey = tpl.name;
    } else {
      const camp = STATE.campaigns.find((c) => {
        const sn = c.name.replace(/[^\w\-]/g, "_");
        return (
          base === sn ||
          base === c.name ||
          base.startsWith(sn + "_") ||
          base.startsWith(c.name + "_")
        );
      });
      if (camp) {
        cat = f.isBaseline ? "Campaign Baselines" : "Campaigns";
        groupKey = camp.name;
      }
    }

    if (cat && groupKey) {
      if (!tree[cat][groupKey]) tree[cat][groupKey] = [];
      tree[cat][groupKey].push(f);
    } else {
      tree["Other"].push(f);
    }
  });

  const renderFileRow = (f) => `
    <div class="tree-file-row ${STATE._activeOutputName === f.name ? "active" : ""}" onclick="loadSelectedOutput('${escAttr(f.name)}')">
      <div class="tree-file-info">
        <div style="flex:1;min-width:0">
          <div class="tree-file-name" title="${esc(f.name)}">${esc(f.name)}</div>
          <div style="font-size:9px;color:var(--text3);opacity:.7;margin-top:1px">${formatDateTime(f.lastModified)} • ${formatFileSize(f.size)}</div>
        </div>
      </div>
      <div class="tree-file-actions">
        <button class="btn-icon" onclick="event.stopPropagation();displayDuplicationHistory('${escAttr(f.name)}')" title="View Duplicates"><i data-lucide="copy" class="icon"></i></button>
        <button class="btn-icon" onclick="event.stopPropagation();downloadOutputFile('${escAttr(f.name)}')" title="Download"><i data-lucide="download" class="icon"></i></button>
        <button class="btn-icon danger" onclick="event.stopPropagation();deleteOutputFile('${escAttr(f.name)}')" title="Delete"><i data-lucide="trash-2" class="icon"></i></button>
      </div>
    </div>
  `;

  const renderGroupNode = (name, files, icon = "folder", idPrefix = "") => {
    const nodeId = idPrefix + name;
    const isOpen = STATE._openNodes.has(nodeId);
    return `
    <div class="tree-node ${isOpen ? "open" : ""}">
      <div class="tree-node-hdr" onclick="toggleTreeNode(this,'${escAttr(nodeId)}')">
        <span class="tree-node-icon"><i data-lucide="chevron-right" class="icon"></i></span>
        <i data-lucide="${icon}" class="icon" style="width:12px;height:12px;color:var(--blue)"></i>
        <span style="flex:1;font-size:12px">${esc(name)}</span>
        <span class="tree-file-meta">${files.length}</span>
      </div>
      <div class="tree-children">
        ${files.map((f) => renderFileRow(f)).join("")}
      </div>
    </div>
  `;
  };

  let html = '<div class="output-tree">';

  if (Object.keys(tree["Template Baselines"]).length > 0) {
    html += `<div class="tree-category-label">Template Baselines</div>`;
    Object.keys(tree["Template Baselines"])
      .sort()
      .forEach(
        (name) =>
          (html += renderGroupNode(
            name,
            tree["Template Baselines"][name],
            "layout",
            "tpl-base-",
          )),
      );
  }
  if (Object.keys(tree["Campaign Baselines"]).length > 0) {
    html += `<div class="tree-category-label">Campaign Baselines</div>`;
    Object.keys(tree["Campaign Baselines"])
      .sort()
      .forEach(
        (name) =>
          (html += renderGroupNode(
            name,
            tree["Campaign Baselines"][name],
            "target",
            "cmp-base-",
          )),
      );
  }
  if (Object.keys(tree["Templates"]).length > 0) {
    html += `<div class="tree-category-label">Templates</div>`;
    Object.keys(tree["Templates"])
      .sort()
      .forEach(
        (name) =>
          (html += renderGroupNode(
            name,
            tree["Templates"][name],
            "layout",
            "tpl-",
          )),
      );
  }
  if (Object.keys(tree["Campaigns"]).length > 0) {
    html += `<div class="tree-category-label">Campaigns</div>`;
    Object.keys(tree["Campaigns"])
      .sort()
      .forEach(
        (name) =>
          (html += renderGroupNode(
            name,
            tree["Campaigns"][name],
            "target",
            "cmp-",
          )),
      );
  }
  if (tree["Other"].length > 0) {
    html += `<div class="tree-category-label">Misc</div>`;
    html += renderGroupNode(
      "Uncategorized",
      tree["Other"],
      "file-question",
      "misc-",
    );
  }
  html += "</div>";
  el.innerHTML = html;
  refreshIcons();
}
function confirmSaveBaseline() {
  if (!STATE._activeOutputName || !STATE._currentWorkbook) {
    toast("No file loaded", "error");
    return;
  }
  const campName = STATE._activeOutputName
    .replace(/_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}\.xlsx$/, "")
    .replace(/\.xlsx$/, "");

  // Count today's new rows
  let todayRows = 0;
  STATE._currentWorkbook.SheetNames.forEach((n) => {
    todayRows += XLSX.utils.sheet_to_json(
      STATE._currentWorkbook.Sheets[n],
    ).length;
  });

  confirmDialog(
    "Save Baseline",
    `This will merge today's ${todayRows} new rows into the cumulative baseline for "${campName}". Tomorrow's incremental run will compare against the updated baseline. Continue?`,
    () => saveActiveCampaignAsBaseline(),
  );
}

async function saveActiveCampaignAsBaseline() {
  if (!STATE._activeOutputName || !STATE._currentWorkbook) {
    toast("No file loaded", "error");
    return;
  }
  if (!STATE.folderHandle) {
    toast("Link a project folder first", "error");
    return;
  }
  const campName = STATE._activeOutputName
    .replace(/_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}\.xlsx$/, "")
    .replace(/\.xlsx$/, "");
  const filename = getNormalizedFilename(campName);

  try {
    await loadSheetJS();
    const templatesDir = await STATE.folderHandle.getDirectoryHandle(
      "templates",
      { create: true },
    );
    const baselinesDir = await templatesDir.getDirectoryHandle("baselines", {
      create: true,
    });

    // Load existing baseline if it exists
    let existingSheets = {};
    try {
      const existingHandle = await baselinesDir.getFileHandle(filename);
      const existingFile = await existingHandle.getFile();
      const existingAb = await existingFile.arrayBuffer();
      const existingWb = XLSX.read(new Uint8Array(existingAb), {
        type: "array",
      });
      existingWb.SheetNames.forEach((name) => {
        existingSheets[name] = XLSX.utils.sheet_to_json(
          existingWb.Sheets[name],
        );
      });
    } catch (e) {
      // No existing baseline — first time save
    }

    // Merge: existing baseline rows + today's new rows = full cumulative
    const newWb = STATE._currentWorkbook;
    const mergedWb = XLSX.utils.book_new();
    let totalRowsCount = 0;

    const allSheets = new Set([
      ...Object.keys(existingSheets),
      ...newWb.SheetNames,
    ]);

    allSheets.forEach((sheetName) => {
      const oldRows = existingSheets[sheetName] || [];
      const newSheet = newWb.Sheets[sheetName];
      const newRows = newSheet
        ? XLSX.utils.sheet_to_json(newSheet, { defval: "" })
        : [];

      // DEDUPLICATION: Find unique columns from template
      let keys = [];
      const group = STATE.groups.find((g) => g.name === sheetName);
      let tplName = group
        ? group.templateName
        : STATE._activeTemplateName || sheetName;
      const tpl = STATE.templates.find((t) => t.name === tplName);

      if (tpl && tpl.dedupCols && tpl.dedupCols.length > 0) {
        keys = tpl.dedupCols;
      }

      const seen =
        keys.length > 0
          ? new Set(oldRows.map((r) => getLeadKey(r, keys)))
          : null;
      const filteredNew = newRows.filter((r) => {
        if (keys.length === 0 || !seen) return true;
        const key = getLeadKey(r, keys);
        if (!key || key.replace(/\|/g, "").trim() === "") return true;
        if (seen.has(key)) return false;
        return true;
      });

      const combined = [...oldRows, ...filteredNew];
      totalRowsCount += combined.length;
      const ws = XLSX.utils.json_to_sheet(combined);
      XLSX.utils.book_append_sheet(mergedWb, ws, sheetName.substring(0, 31));
    });

    // Save merged baseline back
    const fileHandle = await baselinesDir.getFileHandle(filename, {
      create: true,
    });
    const writable = await fileHandle.createWritable();
    const ab = XLSX.write(mergedWb, { bookType: "xlsx", type: "array" });
    await writable.write(ab);
    await writable.close();

    toast(
      `Baseline saved — ${totalRowsCount} total rows in templates/baselines/${filename}`,
      "success",
    );
  } catch (e) {
    console.error(e);
    toast("Failed: " + e.message, "error");
  }
}

async function deleteOutputFile(name) {
  if (!confirm(`Permanently delete "${name}"?`)) return;
  const f = STATE._outputFiles.find((file) => file.name === name);
  if (!f) return;
  try {
    let dir;
    if (f.isBaseline) {
      const tDir = await STATE.folderHandle.getDirectoryHandle("templates");
      dir = await tDir.getDirectoryHandle("baselines");
    } else {
      dir = await STATE.folderHandle.getDirectoryHandle("output");
    }
    await dir.removeEntry(name);
    toast("File deleted", "success");
    if (STATE._activeOutputName === name) {
      document.getElementById("outputViewerContent").innerHTML =
        '<div class="empty-state" style="margin-top:40px"><div class="es-icon"><i data-lucide="table"></i></div><strong>Select a file to view</strong></div>';
      STATE._activeOutputName = null;
      refreshIcons();
    }
    scanOutputFolder(false);
  } catch (e) {
    toast("Could not delete file", "error");
  }
}

async function loadSelectedOutput(name) {
  if (!name) return;
  const fileInfo = STATE._outputFiles.find((f) => f.name === name);
  if (!fileInfo) return;
  STATE._activeOutputName = name;

  // Detect template for this file
  const base = name.replace(/\.[^.]+$/, "");
  const tpl = STATE.templates.find((t) => {
    const sn = t.name.replace(/[^\w\-]/g, "_");
    return (
      base === sn ||
      base === t.name ||
      base.startsWith(sn + "_") ||
      base.startsWith(t.name + "_")
    );
  });
  STATE._activeTemplateName = tpl ? tpl.name : null;

  const file = await fileInfo.handle.getFile();
  displayOutputFile(file);
  document.querySelectorAll(".tree-file-row").forEach((row) => {
    row.classList.toggle(
      "active",
      row.querySelector(".tree-file-name")?.textContent === name,
    );
  });
  refreshViewerPage();
}

async function displayOutputFile(file) {
  const content = document.getElementById("outputViewerContent");
  content.innerHTML =
    '<div style="padding:40px;text-align:center;color:var(--text3)">⏳ Processing...</div>';
  try {
    await loadSheetJS();
    const ab = await file.arrayBuffer();
    const data = new Uint8Array(ab);
    const workbook = XLSX.read(data, { type: "array" });
    renderWorkbook(workbook, file.name);
  } catch (err) {
    console.error("Load error:", err);
    content.innerHTML =
      '<div style="padding:20px;color:var(--red)">Failed to load file: ' +
      err.message +
      "</div>";
    toast("Error reading file", "error");
  }
}

function renderWorkbook(wb, fileName) {
  const tabs = document.getElementById("sheetTabsBar");
  tabs.style.display = wb.SheetNames.length > 1 ? "flex" : "none";
  tabs.className = "sheet-tabs-container";
  tabs.innerHTML = wb.SheetNames.map((name, i) => {
    const ws = wb.Sheets[name];
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
    const rowCount = Math.max(0, range.e.r); // e.r is index of last row, so range.e.r is count excluding first header row if we assume 0-indexed
    // Actually, sheet_to_json(ws, {header:1}).length - 1 is safer if it's small, but range is faster.
    // Let's use range but be careful. e.r = 0 means 1 row (header). e.r = 5 means 6 rows (header + 5 data).
    return `<button class="sheet-tab ${i === 0 ? "active" : ""}" onclick="switchSheet(this,${i})">${esc(name)} <span class="tab-count">${rowCount}</span></button>`;
  }).join("");
  STATE._currentWorkbook = wb;
  switchSheet(tabs.firstChild, 0);
}

function switchSheet(btn, idx) {
  if (btn) {
    document
      .querySelectorAll(".sheet-tab")
      .forEach((t) => t.classList.remove("active"));
    btn.classList.add("active");
  }
  const wb = STATE._currentWorkbook;
  const sheetName = wb.SheetNames[idx];
  const ws = wb.Sheets[sheetName];

  const activeOutputFile = STATE._outputFiles.find(
    (f) => f.name === STATE._activeOutputName,
  );
  const isBaselineFile = activeOutputFile ? activeOutputFile.isBaseline : false;
  const content = document.getElementById("outputViewerContent");

  content.innerHTML = `
    <div>
      <div class="output-preview-header">
        <div class="output-preview-title">
          <i data-lucide="file-spreadsheet" class="icon"></i> 
          ${esc(sheetName)} 
          <span class="row-count-badge" id="activeSheetRowCount">0 rows</span>
          <span style="font-size:11px;font-weight:400;color:var(--text3);margin-left:8px">from ${esc(STATE._activeOutputName || "file")} ${isBaselineFile ? "(Baseline)" : ""}</span>
        </div>
        <div style="display:flex;gap:12px;align-items:center;">
          <button class="expand-btn" onclick="toggleExpandTable()"><i data-lucide="maximize" id="expandIcon" class="icon"></i> <span id="expandText">Expand</span></button>
          <button class="btn btn-ghost btn-sm" onclick="copyTableToClipboard('viewerTable')"><i data-lucide="copy" class="icon"></i> Copy</button>
          <button class="btn btn-ghost btn-sm" onclick="displayDuplicationHistory(STATE._activeOutputName)" style="border-color:var(--blue);color:var(--blue);"><i data-lucide="copy" class="icon"></i> Duplicates</button>
          <div style="position:relative;display:flex;align-items:center;background:var(--surface2);border:1.5px solid var(--blue);border-radius:8px;padding:2px 4px;">
            <i data-lucide="search" class="icon" style="position:absolute;left:8px;width:12px;height:12px;color:var(--blue);"></i>
            <input type="number" id="jumpToRowInput" class="input" placeholder="Search Row #" style="width:150px;padding-left:26px;font-size:11px;height:24px;background:transparent;border:none;color:var(--text1);" onkeyup="if(event.key==='Enter')jumpToRow(this.value)">
          </div>
          ${
            !isBaselineFile
              ? `
          <button class="btn btn-success btn-sm" onclick="confirmSaveBaseline()" style="display:inline-flex;align-items:center;gap:6px;padding:7px 14px;background:#10b981;color:white;border:none;border-radius:8px;font-family:var(--font);font-size:13px;font-weight:600;cursor:pointer;"><i data-lucide="save" class="icon"></i>Save Baseline</button>
          `
              : ""
          }
        </div>
      </div>

      <div class="output-table-wrap" id="outputTableWrap">
        <table id="viewerTable"><thead><tr id="tableHead"></tr></thead><tbody id="tableBody"></tbody></table>
        <div id="loadMoreContainer" style="padding:20px;text-align:center;display:none"><button class="btn btn-ghost" onclick="renderMoreRows()">Load More Rows</button></div>
      </div>
      <button class="close-expand" onclick="toggleExpandTable()"><i data-lucide="x" class="icon"></i></button>
    </div>
  `;
  refreshIcons();
  const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const headers = json[0] || [];
  const rows = json.slice(1);
  STATE._currentRows = rows;
  STATE._currentHeaders = headers;
  STATE._renderedRowCount = 0;

  const countBadge = document.getElementById("activeSheetRowCount");
  if (countBadge) countBadge.textContent = `${rows.length} rows`;

  let headHtml = `<th style="width:40px;text-align:center;color:var(--text3);background:var(--surface2);position:sticky;left:0;z-index:10;user-select:none;pointer-events:none;">#</th>`;
  headHtml += headers
    .map((h) => {
      const activeTpl = STATE.templates.find(
        (t) => t.name === STATE._activeTemplateName,
      );
      const tplCol = activeTpl?.columns?.find((c) => c.name === h);
      const alignClass = getSmartAlignmentClass(h, tplCol ? tplCol.fmt : "");
      return `<th class="${alignClass}">${esc(String(h))}</th>`;
    })
    .join("");
  document.getElementById("tableHead").innerHTML = headHtml;

  renderMoreRows();
}
function updateToggleIcon() {
  const sidebar = document.getElementById("viewerSidebar");
  const btn = document.querySelector(".sidebar-toggle-btn");
  if (!sidebar || !btn) return;

  const isCollapsed = sidebar.classList.contains("collapsed");

  btn.innerHTML = isCollapsed
    ? `<i data-lucide="chevron-right" class="icon"></i>` // collapsed → show >
    : `<i data-lucide="chevron-left" class="icon"></i>`; // open → show <

  refreshIcons();
}
function renderMoreRows() {
  const body = document.getElementById("tableBody");
  const rows = STATE._currentRows || [];
  const start = STATE._renderedRowCount;
  const end = Math.min(start + 100, rows.length);
  const chunk = rows.slice(start, end);

  const headers = STATE._currentHeaders || [];

  body.insertAdjacentHTML(
    "beforeend",
    chunk
      .map((row, i) => {
        const rowIdx = start + i + 1; // 1-based indexing for data rows
        const cellsHtml = headers
          .map((h, j) => {
            const rawVal = row[j] !== undefined ? row[j] : "";
            const formatted = formatTableValue(rawVal, String(h));
            const activeTpl = STATE.templates.find(
              (t) => t.name === STATE._activeTemplateName,
            );
            const tplCol = activeTpl?.columns?.find((c) => c.name === h);
            const fmt = tplCol ? tplCol.fmt : "";
            const alignClass = getSmartAlignmentClass(h, fmt);
            const isDate =
              String(h).toLowerCase().includes("time") ||
              String(h).toLowerCase().includes("date");
            const isNumeric =
              typeof rawVal === "number" ||
              (!isNaN(rawVal) && String(rawVal).trim() !== "");
            return `<td class="${alignClass} ${isDate ? "date-col" : ""} ${isNumeric ? "number-cell" : ""}">${esc(formatted)}</td>`;
          })
          .join("");
        return `<tr id="preview-row-${rowIdx}" onclick="this.querySelectorAll('td').forEach(t=>t.style.whiteSpace=t.style.whiteSpace==='normal'?'nowrap':'normal')">
          <td style="width:40px;text-align:center;color:var(--text3);background:var(--surface2);position:sticky;left:0;font-size:10px;font-weight:700;user-select:none;pointer-events:none;border-right:1px solid var(--border);">${rowIdx}</td>
          ${cellsHtml}
        </tr>`;
      })
      .join(""),
  );

  STATE._renderedRowCount = end;
  const lm = document.getElementById("loadMoreContainer");
  if (lm) lm.style.display = end < rows.length ? "block" : "none";
}

function formatTableValue(val, colName) {
  if (val === null || val === undefined) return "";
  if (
    val instanceof Date ||
    (typeof val === "number" &&
      val > 40000 &&
      val < 60000 &&
      String(colName).toLowerCase().includes("date"))
  ) {
    try {
      const d = new Date(val);
      if (!isNaN(d.getTime())) {
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const yyyy = d.getFullYear();
        const hh = String(d.getHours()).padStart(2, "0");
        const min = String(d.getMinutes()).padStart(2, "0");
        return `${dd}-${mm}-${yyyy} ${hh}:${min}`;
      }
    } catch (e) {}
  }
  return String(val);
}

function getSmartAlignmentClass(colName, fmt) {
  if (fmt) {
    const codes = fmt.trim().split(/\s+/);
    if (codes.includes("l")) return "align-left";
    if (codes.includes("r")) return "align-right";
    if (codes.includes("z")) return "align-center";
  }

  const name = String(colName).toLowerCase();

  // Smart Defaults based on common column names
  if (
    name.includes("id") ||
    name.includes("count") ||
    name.includes("price") ||
    name.includes("amount") ||
    name.includes("total") ||
    name.includes("qty") ||
    name.includes("quantity") ||
    name.includes("phone") ||
    name.includes("zip") ||
    name.includes("pin")
  ) {
    return "align-right";
  }

  if (
    name.includes("date") ||
    name.includes("time") ||
    name.includes("status") ||
    name.includes("tag") ||
    name.includes("category") ||
    name.includes("code")
  ) {
    return "align-center";
  }

  return "align-left";
}

function toggleViewerSidebar() {
  const sidebar = document.getElementById("viewerSidebar");
  if (!sidebar) return;

  sidebar.classList.toggle("collapsed");
  updateToggleIcon();
}

function toggleExpandTable() {
  const wrap = document.getElementById("outputTableWrap");
  if (!wrap) return;
  const icon = document.getElementById("expandIcon");
  const text = document.getElementById("expandText");
  const isExpanded = wrap.classList.toggle("expanded");
  document.body.classList.toggle("body-expanded", isExpanded);
  if (icon) {
    icon.setAttribute("data-lucide", isExpanded ? "minimize" : "maximize");
  }
  if (text) text.textContent = isExpanded ? "Restore" : "Expand";
  document.body.style.overflow = isExpanded ? "hidden" : "";
  refreshIcons();
}

function copyTableToClipboard() {
  const rows = STATE._currentRows || [];
  const headers = STATE._currentHeaders || [];

  if (!rows.length && !headers.length) {
    toast("No data to copy", "warning");
    return;
  }

  // Data only, no headers
  let text = "";

  // Format each row (handle undefined/null cells)
  for (const row of rows) {
    const line = headers
      .map((_, i) => {
        const val = row[i];
        return val === null || val === undefined ? "" : String(val);
      })
      .join("\t");
    text += line + "\n";
  }

  copyToClipboard(text);
  toast(`Copied all ${rows.length} rows (data only) to clipboard`, "success");
}

// ─── RUN PAGE ──────────────────────────────────────────────────────────────
async function refreshRunPage() {
  await syncTemplatesWithFolder(true);
  const sortedTpls = [...STATE.templates].sort((a, b) =>
    a.name.toLowerCase().localeCompare(b.name.toLowerCase()),
  );
  const currentTpl = document.getElementById("runTemplate").value;
  const tplOpts =
    '<option value="">-- Select Template --</option>' +
    sortedTpls
      .map(
        (t) =>
          `<option value="${esc(t.name)}" ${t.name === currentTpl ? "selected" : ""}>${esc(t.name)}</option>`,
      )
      .join("");
  document.getElementById("runTemplate").innerHTML = tplOpts;

  const sortedCamps = [...STATE.campaigns].sort((a, b) =>
    a.name.toLowerCase().localeCompare(b.name.toLowerCase()),
  );
  const currentCamp = document.getElementById("runCampaignConfig")?.value;
  const campOpts =
    '<option value="">-- Select Config --</option>' +
    sortedCamps
      .map(
        (c) =>
          `<option value="${esc(c.name)}" ${c.name === currentCamp ? "selected" : ""}>${esc(c.name)}</option>`,
      )
      .join("");
  if (document.getElementById("runCampaignConfig"))
    document.getElementById("runCampaignConfig").innerHTML = campOpts;

  // Scan input subfolders for auto-select
  if (STATE.folderHandle) await scanInputSubfolders();

  // Clean up listeners
  document.querySelectorAll('[name="runProcMode"]').forEach((r) => {
    r.onchange = () => {
      document
        .getElementById("runAdvancedOpts1")
        .classList.toggle(
          "hidden",
          document.querySelector('[name="runProcMode"]:checked').value !==
            "advanced",
        );
      toggleCustomDedupUI();
    };
  });

  // Add listeners for deduplication radio buttons
  document.querySelectorAll('[name="runDedup1"]').forEach((r) => {
    r.onchange = () => {
      toggleCustomDedupUI();
    };
  });

  if (STATE.folderHandle) {
    const chk = document.getElementById("chkFolderIcon");
    if (chk) chk.textContent = "✓";
  }
}

function toggleCustomDedupUI() {
  const isAdvanced =
    document.querySelector('[name="runProcMode"]:checked').value === "advanced";
  const isCustom =
    document.querySelector('[name="runDedup1"]:checked')?.value === "custom";
  const container = document.getElementById("runCustomDedupContainer");
  if (!container) return;

  if (isAdvanced && isCustom) {
    container.classList.remove("hidden");
    renderCustomDedupList();
  } else {
    container.classList.add("hidden");
  }
}

function renderCustomDedupList() {
  const tplName = document.getElementById("runTemplate").value;
  const tpl = STATE.templates.find((t) => t.name === tplName);
  const listEl = document.getElementById("runCustomDedupList");
  if (!listEl) return;

  if (!tpl) {
    listEl.innerHTML =
      '<div style="font-size:11px;color:var(--text3)">Select a template first</div>';
    return;
  }

  listEl.innerHTML = tpl.columns
    .map(
      (c) => `
    <label class="col-chip" style="cursor:pointer;user-select:none;font-size:11px;display:flex;align-items:center;gap:4px">
      <input type="checkbox" class="run-custom-dedup-check" value="${esc(c.name)}">
      ${esc(c.name)}
    </label>
  `,
    )
    .join("");
}

async function scanInputSubfolders() {
  if (!STATE.folderHandle) return;
  try {
    const inputDir = await STATE.folderHandle.getDirectoryHandle("input", {
      create: true,
    });
    const subfolders = [];
    for await (const [name, handle] of inputDir.entries()) {
      if (handle.kind === "directory") subfolders.push(name);
    }
    STATE._inputSubfolders = subfolders;
  } catch (e) {
    STATE._inputSubfolders = [];
  }
}

function showTemplateInfo(name) {
  const tpl = STATE.templates.find((t) => t.name === name);
  const el = document.getElementById("runTemplate1Info");
  if (!el) return;
  el.innerHTML = tpl
    ? `${tpl.columns.length} columns · Dedup: ${tpl.dedupCols.length ? tpl.dedupCols.join(", ") : "none"}`
    : "";

  const links = tpl ? tpl.googleLinks || tpl.google_links || [] : [];

  // Auto-populate External URLs section for Standard Merge
  const extContainer = document.getElementById("runExtFiles1");
  if (extContainer && links.length > 0) {
    extContainer.innerHTML = "";
    links.forEach((url) => {
      const div = document.createElement("div");
      div.className = "smart-src-row";
      div.style.marginBottom = "12px";
      div.innerHTML = buildSrcRow(
        { path: url, label: "Sheet from Template" },
        "ext-file-label",
        "ext-file-path",
      );
      extContainer.appendChild(div);
      initSrcRow(div, ".ext-file-path", ".ext-file-label");
    });
    // Switch radio to external if links found
    const extRadio = document.querySelector(
      'input[name="runSrc1"][value="external"]',
    );
    if (extRadio) {
      extRadio.checked = true;
      toggleExtFiles1();
    }
    refreshIcons();
    toast(`${links.length} sources pre-filled from template`);
  }

  // Auto-select matching input subfolder
  if (name && STATE._inputSubfolders && STATE._inputSubfolders.length > 0) {
    const tplSafeName = name.replace(/[^\w\-]/g, "_");
    const match = STATE._inputSubfolders.find(
      (f) =>
        f === name ||
        f === tplSafeName ||
        f.toLowerCase() === name.toLowerCase() ||
        f.toLowerCase() === tplSafeName.toLowerCase(),
    );
    if (match) {
      const srcRadio = document.querySelector(
        'input[name="runSrc1"][value="folder"]',
      );
      if (srcRadio) {
        srcRadio.checked = true;
        toggleExtFiles1();
      }
      toast(`Auto-selected input folder: ${match}`, "info");
    }
  }

  toggleCustomDedupUI();
}

function handleGrpTemplateChange(tplName) {
  const tpl = STATE.templates.find((t) => t.name === tplName);
  const container = document.getElementById("grpSources");
  if (!tpl || !container) return;
  const links = tpl.googleLinks || tpl.google_links || [];
  if (links.length > 0) {
    container.innerHTML = "";
    links.forEach((url) => {
      const div = document.createElement("div");
      div.className = "smart-src-row";
      div.style.marginBottom = "12px";
      div.innerHTML = buildSrcRow(
        { path: url, label: "Sheet from Template" },
        "grp-src-label",
        "grp-src-path",
      );
      container.appendChild(div);
      initSrcRow(div, ".grp-src-path", ".grp-src-label");
    });
    refreshIcons();
    toast(`${links.length} sources pre-filled from template`);
  }
}

function renderCustomDedupList() {
  const tplName = document.getElementById("runTemplate").value;
  const tpl = STATE.templates.find((t) => t.name === tplName);
  const listEl = document.getElementById("runCustomDedupList");
  if (!listEl) return;

  if (!tpl) {
    listEl.innerHTML =
      '<div style="font-size:11px;color:var(--text3)">Select a template first</div>';
    return;
  }

  listEl.innerHTML = tpl.columns
    .map(
      (c) => `
    <label class="col-chip" style="cursor:pointer;user-select:none;font-size:11px;display:flex;align-items:center;gap:4px">
      <input type="checkbox" class="run-custom-dedup-check" value="${esc(c.name)}">
      ${esc(c.name)}
    </label>
  `,
    )
    .join("");
}

function toggleExtFiles1() {
  const v = document.querySelector('[name="runSrc1"]:checked').value;
  document
    .getElementById("runSrc1External")
    .classList.toggle("hidden", v !== "external");
}
function toggleCampAction() {
  const v = document.querySelector('[name="runCampAction"]:checked').value;
  document
    .getElementById("campExistingPick")
    .classList.toggle("hidden", v !== "existing");
  document
    .getElementById("campNewInfo")
    .classList.toggle("hidden", v !== "new");
}
function selectMode(m) {
  STATE.currentMode = m;
  document
    .querySelectorAll(".mode-tab")
    .forEach((b) => b.classList.toggle("active", +b.dataset.mode === m));
  document
    .querySelectorAll(".mode-section")
    .forEach((s) => s.classList.remove("active"));
  document.getElementById("runMode" + m).classList.add("active");
}
function addExtFile(containerId) {
  const div = document.createElement("div");
  div.className = "smart-src-row";
  div.innerHTML = buildSrcRow(null, "ext-file-label", "ext-file-path");
  document.getElementById(containerId).appendChild(div);
  initSrcRow(div, ".ext-file-path", ".ext-file-label");
  refreshIcons();
}

function updateCampPreview() {
  const name = document.getElementById("runCampaignConfig").value;
  const cfg = STATE.campaigns.find((c) => c.name === name);
  const box = document.getElementById("campPreview");
  if (!cfg) {
    box.innerHTML = "";
    return;
  }
  box.innerHTML = `<div style="margin-top:10px;background:var(--surface2);border:1.5px solid var(--border);border-radius:10px;padding:12px"><div style="font-size:11px;color:var(--text3);margin-bottom:8px">${cfg.groups.length} group(s)</div>${cfg.groups.map((g) => `<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--border);font-size:12px"><i data-lucide="package" class="icon" style="opacity:.6"></i><span>${esc(g)}</span></div>`).join("")}</div>`;
  refreshIcons();
}

function generateCommand() {
  const m = STATE.currentMode;
  const script = "advanced_merger.py";
  let inputs = [];
  let badge = "";
  if (m === 1) {
    const tplName = document.getElementById("runTemplate").value;
    const tpl = STATE.templates.find((t) => t.name === tplName);
    if (!tplName || !tpl) {
      toast("Select a template first", "error");
      return;
    }
    const srcMode = document.querySelector('[name="runSrc1"]:checked').value;
    const procMode = document.querySelector(
      '[name="runProcMode"]:checked',
    ).value;
    badge = "OPTION 1 — USE TEMPLATE";
    inputs.push({ prompt: "Select option:", value: "1" });
    if (srcMode === "folder") {
      inputs.push({
        prompt: "Choose (1/2):",
        value: "1",
        note: "input/ folder",
      });
    } else {
      const extFiles = [
        ...document.querySelectorAll("#runExtFiles1 .ext-file-path"),
      ]
        .map((i) => i.value.trim())
        .filter(Boolean);
      if (!extFiles.length) {
        toast("Add at least one external file/URL", "error");
        return;
      }
      inputs.push({
        prompt: "Choose (1/2):",
        value: "2",
        note: "External files/URLs",
      });
      extFiles.forEach((f, i) =>
        inputs.push({ prompt: `File ${i + 1}:`, value: f }),
      );
      inputs.push({
        prompt: `File ${extFiles.length + 1} (ENTER to finish):`,
        value: "",
        note: "Press ENTER",
      });
    }
    inputs.push({
      prompt: "Select template:",
      value: String(getTemplateIndex(tplName)),
      note: `#${getTemplateIndex(tplName)} = "${tplName}"`,
    });

    if (procMode === "quick") {
      inputs.push({
        prompt: "Quick or Advanced (1/2):",
        value: "1",
        note: "Quick = auto dedup + auto filename",
      });
    } else {
      inputs.push({ prompt: "Quick or Advanced (1/2):", value: "2" });
    }

    const inc = document.getElementById("runIncrementalCheck1").checked;
    inputs.push({
      prompt: "Full or Incremental (1/2):",
      value: inc ? "2" : "1",
      note: inc ? "Incremental" : "Full",
    });

    if (procMode === "advanced") {
      const outName = document.getElementById("runOutputName1").value.trim();
      inputs.push({
        prompt: "Output file name:",
        value: outName || tplName,
      });

      const dedupType = document.querySelector(
        '[name="runDedup1"]:checked',
      ).value;
      if (dedupType === "saved") {
        inputs.push({ prompt: "Deduplication (1/2/3):", value: "1" });
      } else if (dedupType === "custom") {
        inputs.push({ prompt: "Deduplication (1/2/3):", value: "2" });
        const customCols = [
          ...document.querySelectorAll(".run-custom-dedup-check:checked"),
        ]
          .map((c) => c.value)
          .join(", ");
        inputs.push({ prompt: "Custom columns:", value: customCols });
      } else {
        inputs.push({ prompt: "Deduplication (1/2/3):", value: "3" });
      }
    }
  } else if (m === 4) {
    const campName = document.getElementById("runCampaignConfig").value;
    const cfg = STATE.campaigns.find((c) => c.name === campName);
    if (!campName || !cfg) {
      toast("Select a campaign config", "error");
      return;
    }
    badge = "OPTION 2 — CAMPAIGN MODE";
    inputs.push({ prompt: "Select option:", value: "2" });
    inputs.push({
      prompt: "Select campaign number:",
      value: String(getCampaignIndex(campName)),
    });
    const inc = document.getElementById("runIncrementalCheck").checked;
    if (inc) {
      inputs.push({
        prompt: "Run in Incremental Mode (1/2):",
        value: "1",
        note: "Yes",
      });
    } else {
      inputs.push({
        prompt: "Run in Incremental Mode (1/2):",
        value: "2",
        note: "No",
      });
    }
  }
  const oneShot = buildOneShotCmd(`python ${script}`, inputs);
  STATE._lastOneShot = oneShot;
  document.getElementById("cmdBox").innerHTML =
    `<div style="font-family:var(--mono);font-size:11px;color:#7ec8a0;word-break:break-all;line-height:1.6">${esc(oneShot)}</div>`;
  document.getElementById("cmdBadge").textContent = badge;
  document.getElementById("cmdActions").style.display = "flex";
  STATE.runs++;
  save();
  refreshHome();
}

function buildOneShotCmd(baseCmd, inputs) {
  const answers = inputs.map((s) => String(s.value));
  const isWin = /Win/i.test(navigator.platform || "");
  if (isWin) {
    const escaped = answers.map((a) => `'${a.replace(/'/g, "''")}'`).join(", ");
    return `powershell -Command "echo ${escaped}, '' | ${baseCmd}"`;
  } else {
    const escaped = answers.map((a) =>
      a === ""
        ? ""
        : a.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\$/g, "\\$"),
    );
    return `printf "${escaped.join("\\n")}\\n" | ${baseCmd}`;
  }
}

function getTemplateIndex(name) {
  const sortedNames = STATE.templates
    .map((t) => t.name)
    .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
  const i = sortedNames.indexOf(name);
  return i >= 0 ? i + 1 : "?";
}
function getCampaignIndex(name) {
  const sortedNames = STATE.campaigns
    .map((c) => c.name)
    .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
  const i = sortedNames.indexOf(name);
  return i >= 0 ? i + 1 : "?";
}
function copyOneShotCmd() {
  if (!STATE._lastOneShot) {
    toast("Generate a command first", "error");
    return;
  }
  copyToClipboard(STATE._lastOneShot);
  toast("One-shot command copied!", "success");
}

// ─── SETTINGS ──────────────────────────────────────────────────────────────
function exportAll() {
  const data = {
    templates: STATE.templates,
    groups: STATE.groups,
    campaigns: STATE.campaigns,
    exported: new Date().toISOString(),
  };
  const blob = new Blob([JSON.stringify(data, null, 2)], {
    type: "application/json",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `datamerge_config_${Date.now()}.json`;
  a.click();
  toast("Config exported", "success");
}
function importAll() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json";
  input.onchange = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        if (data.templates) STATE.templates = data.templates;
        if (data.groups) STATE.groups = data.groups;
        if (data.campaigns) STATE.campaigns = data.campaigns;
        save();
        toast("Config imported!", "success");
        refreshHome();
      } catch (err) {
        toast("Invalid config file", "error");
      }
    };
    reader.readAsText(file);
  };
  input.click();
}
function clearAll() {
  confirmDialog(
    "Clear All Data",
    "This will permanently remove all templates, groups, and campaigns.",
    () => {
      STATE.templates = [];
      STATE.groups = [];
      STATE.campaigns = [];
      STATE.runs = 0;
      save();
      refreshHome();
      renderTemplateList();
      renderGroupList();
      renderCampaignList();
      toast("All data cleared");
    },
  );
}

// ─── FILE MULTI-SELECT ─────────────────────────────────────────────────────
const FMS_STATE = new WeakMap();
function initFmsDropdown(colCard, existingSrc) {
  FMS_STATE.set(colCard, new Set());
  if (existingSrc) {
    const names = existingSrc
      .replace(/^\[|\]$/g, "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);
    names.forEach((n) => FMS_STATE.get(colCard).add(n));
  }
  refreshFmsDropdown(colCard, existingSrc);
}
function refreshFmsDropdown(colCard, currentSrc) {
  const dropdown = colCard.querySelector(".file-multiselect-dropdown");
  const btnLabel = colCard.querySelector(".fms-placeholder");
  if (!dropdown) return;
  const sel = FMS_STATE.get(colCard) || new Set();
  const cols = STATE.discoveredCols || [];
  if (!cols.length) {
    dropdown.innerHTML =
      '<div style="padding:12px;font-size:12px;color:var(--text3);text-align:center">Detect columns first</div>';
  } else {
    dropdown.innerHTML = `<div class="fms-file-header">📋 Detected Columns</div>${cols.map((c) => `<div class="fms-col-item ${sel.has(c) ? "selected" : ""}" data-col="${escAttr(c)}"><span class="fms-check">${sel.has(c) ? "✓" : ""}</span>${esc(c)}</div>`).join("")}`;
  }
  dropdown.querySelectorAll(".fms-col-item").forEach((item) => {
    item.addEventListener("click", (e) => {
      e.stopPropagation();
      const val = item.dataset.col;
      if (sel.has(val)) sel.delete(val);
      else sel.add(val);
      syncSrcFromFms(colCard, sel);
      refreshFmsDropdown(colCard, "");
    });
  });
  renderFmsTags(colCard, sel);
  if (btnLabel)
    btnLabel.textContent = sel.size
      ? `${sel.size} column${sel.size > 1 ? "s" : ""} selected`
      : "Pick from detected columns…";
}
function syncSrcFromFms(colCard, sel) {
  const srcInput = colCard.querySelector(".tpl-col-src");
  if (!srcInput) return;
  const arr = [...sel];
  if (arr.length === 0) srcInput.value = "";
  else if (arr.length === 1) srcInput.value = arr[0];
  else srcInput.value = "[" + arr.join(",") + "]";
}
function syncFmsTags(srcInput) {
  const colCard = srcInput.closest(".col-card");
  if (!colCard) return;
  const val = srcInput.value.trim();
  const names = val
    .replace(/^\[|\]$/g, "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  const sel = FMS_STATE.get(colCard) || new Set();
  sel.clear();
  names.forEach((n) => sel.add(n));
  renderFmsTags(colCard, sel);
  const btnLabel = colCard.querySelector(".fms-placeholder");
  if (btnLabel)
    btnLabel.textContent = sel.size
      ? `${sel.size} column${sel.size > 1 ? "s" : ""} selected`
      : "Pick from detected columns…";
  colCard.querySelectorAll(".fms-col-item").forEach((item) => {
    item.classList.toggle("selected", sel.has(item.dataset.col));
    const check = item.querySelector(".fms-check");
    if (check) check.textContent = sel.has(item.dataset.col) ? "✓" : "";
  });
}
function renderFmsTags(colCard, sel) {
  const tagsEl = colCard.querySelector(".fms-selected-tags");
  if (!tagsEl) return;
  tagsEl.innerHTML = [...sel]
    .map(
      (c) =>
        `<span class="fms-tag">${esc(c)}<span class="fms-tag-x" onclick="removeFmsTag(this,'${escAttr(c)}')">×</span></span>`,
    )
    .join("");
}
function removeFmsTag(el, colName) {
  const colCard = el.closest(".col-card");
  const sel = FMS_STATE.get(colCard) || new Set();
  sel.delete(colName);
  syncSrcFromFms(colCard, sel);
  refreshFmsDropdown(colCard, "");
}
function toggleFmsDropdown(btn) {
  const wrap = btn.closest(".file-multiselect-wrap");
  const dropdown = wrap.querySelector(".file-multiselect-dropdown");
  const isOpen = dropdown.classList.contains("open");
  document.querySelectorAll(".file-multiselect-dropdown.open").forEach((d) => {
    d.classList.remove("open");
    d.closest(".file-multiselect-wrap")
      .querySelector(".file-multiselect-btn")
      .classList.remove("open");
  });
  if (!isOpen) {
    dropdown.classList.add("open");
    btn.classList.add("open");
    const colCard = btn.closest(".col-card");
    const src = colCard.querySelector(".tpl-col-src");
    refreshFmsDropdown(colCard, src ? src.value : "");
  }
}
document.addEventListener("click", (e) => {
  if (!e.target.closest(".file-multiselect-wrap"))
    document
      .querySelectorAll(".file-multiselect-dropdown.open")
      .forEach((d) => {
        d.classList.remove("open");
        d.closest(".file-multiselect-wrap")
          .querySelector(".file-multiselect-btn")
          .classList.remove("open");
      });
});

// ─── DICT EDITOR ───────────────────────────────────────────────────────────
function showDictEditor(colCard, withDefault) {
  const editor = colCard.querySelector(".dict-editor");
  if (!editor) return;
  editor.style.display = "block";
  renderDictEditor(editor, withDefault, {});
}
function hideDictEditor(colCard) {
  const editor = colCard.querySelector(".dict-editor");
  if (editor) {
    editor.style.display = "none";
    editor.innerHTML = "";
  }
}
function restoreDictEditor(colCard, dictData, withDefault) {
  const editor = colCard.querySelector(".dict-editor");
  if (!editor) return;
  editor.style.display = "block";
  renderDictEditor(editor, withDefault, dictData);
}
function renderDictEditor(editor, withDefault, existingData) {
  const defaultVal = existingData["__default__"] || "";
  const pairs = Object.entries(existingData).filter(
    ([k]) => k !== "__default__",
  );
  editor.innerHTML = `<div class="dict-editor-header"><div class="dict-editor-title"><i data-lucide="book" class="icon"></i> ${withDefault ? "Dict + Default" : "Dict Lookup"}</div><button type="button" class="btn btn-sm btn-danger" onclick="clearDictEditor(this.closest('.dict-editor'))"><i data-lucide="trash-2" class="icon"></i> Clear</button></div><div class="dict-bulk-paste-wrap"><div class="dict-bulk-label"><span><i data-lucide="zap" class="icon"></i> Bulk Paste — <code style="background:rgba(217,119,6,.1);padding:1px 5px;border-radius:3px;font-size:10px">key = value</code></span><button type="button" class="dict-bulk-parse-btn" onclick="parseBulkPaste(this.closest('.dict-editor'))">→ Parse</button></div><textarea class="dict-bulk-textarea" placeholder="hoodi = Area 1&#10;whitefield = Area 2"></textarea><div class="dict-bulk-hint">Supports: key = value · key : value · tab-separated</div></div>${withDefault ? `<div class="dict-default-row"><span class="dict-default-label"><i data-lucide="zap" class="icon"></i> Default</span><input type="text" class="input dict-default-val" placeholder="e.g. Unknown" value="${esc(defaultVal)}" style="flex:1"/></div>` : ""}<div style="display:flex;gap:6px;margin-bottom:6px"><div style="flex:1;font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase">KEY</div><div style="flex:0 0 20px"></div><div style="flex:1;font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase">VALUE</div><div style="flex:0 0 24px"></div></div><div class="dict-kv-list">${pairs.length ? pairs.map(([k, v]) => dictKvRow(k, v)).join("") : dictKvRow("", "")}</div><button type="button" class="dict-add-btn" onclick="addDictRow(this)">+ Add Pair</button>`;
  refreshIcons();
}
function dictKvRow(key = "", val = "") {
  return `<div class="dict-kv-row"><input type="text" class="dict-key" placeholder="key" value="${esc(key)}"/><div class="dict-kv-sep">→</div><input type="text" class="dict-val" placeholder="value" value="${esc(val)}"/><div class="dict-kv-del"><button type="button" class="btn-icon danger" onclick="this.closest('.dict-kv-row').remove()">✕</button></div></div>`;
}
function addDictRow(btn) {
  const list = btn.previousElementSibling;
  const row = document.createElement("div");
  row.className = "dict-kv-row";
  row.innerHTML = dictKvRow("", "");
  list.appendChild(row);
  row.querySelector(".dict-key").focus();
}
function readDictEditor(colCard) {
  const editor = colCard.querySelector(".dict-editor");
  if (!editor || editor.style.display === "none") return null;
  const dict = {};
  const defaultEl = editor.querySelector(".dict-default-val");
  if (defaultEl && defaultEl.value.trim())
    dict["__default__"] = defaultEl.value.trim();
  editor.querySelectorAll(".dict-kv-row").forEach((row) => {
    const k = row.querySelector(".dict-key").value.trim().toLowerCase();
    const v = row.querySelector(".dict-val").value.trim();
    if (k) dict[k] = v;
  });
  return dict;
}
function clearDictEditor(editor) {
  if (!editor) return;
  const kvList = editor.querySelector(".dict-kv-list");
  if (!kvList) return;
  kvList.innerHTML = "";
  const row = document.createElement("div");
  row.className = "dict-kv-row";
  row.innerHTML = dictKvRow("", "");
  kvList.appendChild(row);
  const defEl = editor.querySelector(".dict-default-val");
  if (defEl) defEl.value = "";
  const ta = editor.querySelector(".dict-bulk-textarea");
  if (ta) ta.value = "";
  toast("Cleared");
}

// ─── PREVIEW ──────────────────────────────────────────────────────────────

async function refreshInputPage() {
  const query =
    document.getElementById("inputSearch")?.value?.toLowerCase() || "";
  const area = document.getElementById("inputArea");
  if (!area) return;

  if (!STATE.folderHandle) {
    area.innerHTML =
      '<div class="empty-state"><div class="es-icon"><i data-lucide="folder"></i></div><strong>No folder linked</strong>Link a project folder to view input files.</div>';
    refreshIcons();
    return;
  }

  area.innerHTML =
    '<div style="padding:40px;text-align:center;color:var(--text3)"><i data-lucide="loader-2" class="icon spin"></i> Scanning input folder...</div>';
  refreshIcons();

  const inputData = await scanInputFiles();
  const folders = Object.keys(inputData).sort((a, b) =>
    a === "_root" ? -1 : b === "_root" ? 1 : a.localeCompare(b),
  );
  let totalFiles = 0;
  folders.forEach((f) => (totalFiles += inputData[f].length));

  if (totalFiles === 0) {
    area.innerHTML = `
      <div class="empty-state" 
           ondragover="event.preventDefault();this.classList.add('drag-over')" 
           ondragleave="this.classList.remove('drag-over')" 
           ondrop="event.preventDefault();this.classList.remove('drag-over');handleInputFolderDrop(event,'_root')">
        <div class="es-icon"><i data-lucide="folder-open"></i></div>
        <strong>No input files</strong>
        Place CSV, XLSX, or other data files in your input/ folder.
      </div>`;
    refreshIcons();
    return;
  }

  let html = `<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
    <div><span class="preview-label-badge" style="background:var(--green-soft);color:var(--green)">Input Storage</span>
    <span style="font-size:12px;color:var(--text3);margin-left:8px">${totalFiles} file(s) available</span></div>
  </div>`;

  html += '<div class="output-tree">';
  folders.forEach((folder) => {
    let files = inputData[folder];
    if (!files.length) return;
    if (query)
      files = files.filter((f) => f.name.toLowerCase().includes(query));
    if (!files.length) return;

    const folderLabel = folder === "_root" ? "input/" : folder;
    const folderIcon = folder === "_root" ? "folder" : "folder-open";
    const isOpen = STATE._openNodes.has("in-" + folder);
    html += `<div class="tree-node ${isOpen ? "open" : ""}" data-drop-folder="${escAttr(folder)}" ondragover="event.preventDefault();this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="event.preventDefault();this.classList.remove('drag-over');handleInputFolderDrop(event,'${escAttr(folder)}')">
      <div class="tree-node-hdr" onclick="toggleTreeNode(this,'in-${escAttr(folder)}')">
        <span class="tree-node-icon"><i data-lucide="chevron-right" class="icon"></i></span>
        <i data-lucide="${folderIcon}" class="icon" style="width:14px;height:14px;color:var(--blue)"></i>
        <span style="flex:1;font-size:13px;font-weight:600">${esc(folderLabel)}</span>
        <div class="folder-actions-bar" onclick="event.stopPropagation()">
          <span class="tree-file-meta">${files.length}</span>
          <button class="btn-icon danger" onclick="deleteAllInFolder('${escAttr(folder)}')" title="Delete all files in ${esc(folderLabel)}"><i data-lucide="trash" class="icon"></i></button>
        </div>
      </div>
      <div class="tree-children">`;
    files.forEach((f) => {
      const basename = getBaseName(f.name);
      const truncName =
        basename.length > 45 ? basename.substring(0, 42) + "..." : basename;
      html += `<div class="tree-file-row" onclick="this.classList.toggle('active')">
        <div class="tree-file-info">
          <i data-lucide="file-spreadsheet" class="icon" style="opacity:.6"></i>
          <div style="flex:1;min-width:0">
            <div class="tree-file-name" title="${esc(f.name)}">${esc(truncName)}</div>
          </div>
        </div>
        <div style="display:flex;align-items:center;gap:6px">
          <div style="text-align:right;line-height:1.2">
            <div style="font-size:10px;color:var(--text3)">${formatFileSize(f.size)}</div>
            <div style="font-size:9px;color:var(--text3);opacity:.7">${formatDateTime(f.lastModified)}</div>
          </div>
          <div class="tree-file-actions">
            <button class="btn-icon" onclick="event.stopPropagation();downloadInputFile('${escAttr(folder)}','${escAttr(f.name)}')" title="Download"><i data-lucide="download" class="icon"></i></button>
            <button class="btn-icon danger" onclick="event.stopPropagation();deleteInputFile('${escAttr(folder)}','${escAttr(f.name)}')" title="Delete"><i data-lucide="trash-2" class="icon"></i></button>
          </div>
        </div>
      </div>`;
    });
    html += "</div></div>";
  });

  html += `<div class="input-drop-zone" ondragover="event.preventDefault();this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="event.preventDefault();this.classList.remove('drag-over');handleInputFolderDrop(event,'_root')">
    <div class="drop-icon"><i data-lucide="upload" class="icon"></i></div>
    Drag & drop files here to add to input/ root
  </div>`;
  html += "</div>";

  area.innerHTML = html;
  refreshIcons();
}

async function refreshPreview() {
  const mode = STATE.previewMode || "template";
  const query =
    document.getElementById("previewSearch")?.value?.toLowerCase() || "";
  const area = document.getElementById("previewArea");
  if (!area) return;

  let html = "";
  const showAll = STATE._previewShowAll[mode] || false;

  if (mode === "template") {
    const items = query
      ? STATE.templates.filter((t) => t.name.toLowerCase().includes(query))
      : STATE.templates;
    const limit = query || showAll ? items.length : 2;
    const displayed = items.slice(0, limit);

    if (!items.length) {
      html = '<div class="empty-state">No templates found</div>';
    } else {
      html = displayed
        .map((t, i) => {
          const tableId = `previewTable-tpl-${i}`;
          return `
        <div class="card" style="margin-bottom:20px">
          <div class="preview-label-badge">Output Schema Preview</div>
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
            <h2 style="font-size:18px;font-weight:700">${esc(t.name)}</h2>
            <div style="display:flex;gap:8px;align-items:center">
              <button class="btn btn-ghost btn-sm" onclick="copyTableToClipboard('${tableId}')"><i data-lucide="copy" class="icon"></i> Copy</button>
              <div class="xlsx-stat-pill">${t.columns.length} Columns</div>
            </div>
          </div>
          <div class="output-table-wrap" style="max-height:300px">
            <table id="${tableId}">
              <thead><tr>${t.columns
                .map((c) => {
                  const alignClass = getAlignmentClass(c.fmt);
                  return `<th class="${alignClass}">${esc(c.name)}</th>`;
                })
                .join("")}</tr></thead>
              <tbody>
                <tr>${t.columns
                  .map((c) => {
                    const val = formatTableValue(`${esc(c.name)}_1`, c.name);
                    const alignClass = getAlignmentClass(c.fmt);
                    const isDate =
                      c.name.toLowerCase().includes("time") ||
                      c.name.toLowerCase().includes("date");
                    return `<td class="${alignClass} ${isDate ? "date-col" : ""}">${esc(val)}</td>`;
                  })
                  .join("")}</tr>
                <tr>${t.columns
                  .map((c) => {
                    const val = formatTableValue(`${esc(c.name)}_2`, c.name);
                    const alignClass = getAlignmentClass(c.fmt);
                    const isDate =
                      c.name.toLowerCase().includes("time") ||
                      c.name.toLowerCase().includes("date");
                    return `<td class="${alignClass} ${isDate ? "date-col" : ""}"><span style="opacity:0.5">${esc(val)}</span></td>`;
                  })
                  .join("")}</tr>
              </tbody>
            </table>
          </div>
          <div style="margin-top:12px;display:flex;gap:8px">
            ${t.dedupCols.length ? `<span class="xlsx-stat-pill" style="background:var(--purple-soft);color:var(--purple);border-color:#ddd6fe">Dedup: ${t.dedupCols.join(", ")}</span>` : ""}
            ${t.sortCol ? `<span class="xlsx-stat-pill" style="background:var(--blue-soft);color:var(--blue)">Sort: ${t.sortCol} (${t.sortOrder === "desc" ? "Z-A" : "A-Z"})</span>` : ""}
          </div>
        </div>
      `;
        })
        .join("");
      if (items.length > limit) {
        html += `<div style="text-align:center;margin:16px 0"><button class="btn btn-ghost" onclick="STATE._previewShowAll.template=true;refreshPreview()">Show all ${items.length} templates →</button></div>`;
      }
    }
  } else if (mode === "group") {
    const items = query
      ? STATE.groups.filter((g) => g.name.toLowerCase().includes(query))
      : STATE.groups;
    const limit = query || showAll ? items.length : 2;
    const displayed = items.slice(0, limit);

    if (!items.length) {
      html = '<div class="empty-state">No source groups found</div>';
    } else {
      html = displayed
        .map(
          (g) => `
        <div class="card" style="margin-bottom:20px">
          <div class="preview-label-badge" style="background:var(--green-soft);color:var(--green)">Source Group Preview</div>
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
            <h2 style="font-size:16px;font-weight:700">${esc(g.name)}</h2>
            <span class="xlsx-stat-pill">${esc(g.templateName || "None")}</span>
          </div>
          <div style="display:flex;flex-direction:column;gap:3px">
            ${g.sources
              .map((s) => {
                const basename = getBaseName(s.path);
                const truncName =
                  basename.length > 50
                    ? basename.substring(0, 47) + "..."
                    : basename;
                return `<div class="compact-src-card" onclick="this.classList.toggle('expanded')">
                <i data-lucide="link" class="icon" style="opacity:0.5;flex-shrink:0"></i>
                <div style="flex:1;min-width:0">
                  <div class="csc-name" title="${esc(s.path)}">${esc(truncName)}</div>
                  <div class="csc-full">${esc(s.path)}${s.label ? " (" + esc(s.label) + ")" : ""}</div>
                </div>
              </div>`;
              })
              .join("")}
          </div>
        </div>
      `,
        )
        .join("");
      if (items.length > limit) {
        html += `<div style="text-align:center;margin:16px 0"><button class="btn btn-ghost" onclick="STATE._previewShowAll.group=true;refreshPreview()">Show all ${items.length} groups →</button></div>`;
      }
    }
  } else if (mode === "campaign") {
    const items = query
      ? STATE.campaigns.filter((c) => c.name.toLowerCase().includes(query))
      : STATE.campaigns;
    const limit = query || showAll ? items.length : 2;
    const displayed = items.slice(0, limit);

    if (!items.length) {
      html = '<div class="empty-state">No campaigns found</div>';
    } else {
      html = displayed
        .map(
          (c) => `
        <div class="card" style="margin-bottom:20px">
          <div class="preview-label-badge" style="background:var(--purple-soft);color:var(--purple)">Campaign Preview</div>
          <h2 style="font-size:18px;font-weight:700;margin-bottom:12px">${esc(c.name)}</h2>
          <div class="qa-grid" style="grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));gap:10px">
            ${c.groups
              .map((gname) => {
                const g = STATE.groups.find((gr) => gr.name === gname);
                return `
                <div style="padding:10px;background:var(--surface2);border-radius:8px;border:1.5px solid var(--border)">
                  <div style="font-weight:700;font-size:13px">${esc(gname)}</div>
                  <div style="font-size:11px;color:var(--text3)">${g ? g.sources.length : "?"} sources</div>
                </div>
              `;
              })
              .join("")}
          </div>
        </div>
      `,
        )
        .join("");
      if (items.length > limit) {
        html += `<div style="text-align:center;margin:16px 0"><button class="btn btn-ghost" onclick="STATE._previewShowAll.campaign=true;refreshPreview()">Show all ${items.length} campaigns →</button></div>`;
      }
    }
  }

  area.innerHTML = html;

  // Apply top-to-bottom animation
  area.classList.remove("slide-down-anim");
  void area.offsetWidth; // Trigger reflow
  area.classList.add("slide-down-anim");

  refreshIcons();
}
function parseBulkPaste(editor) {
  const textarea = editor.querySelector(".dict-bulk-textarea");
  if (!textarea) return;
  const raw = textarea.value.trim();
  if (!raw) {
    toast("Paste some key=value lines first", "error");
    return;
  }
  const lines = raw.split("\n");
  const pairs = [];
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    let key = "",
      val = "";
    const sepMatch = trimmed.match(/^([^=:\t]+?)\s*(?:=|:|\t)\s*(.+)$/);
    if (sepMatch) {
      key = sepMatch[1].trim();
      val = sepMatch[2].trim();
    }
    if (key && val && !key.startsWith("#"))
      pairs.push({ key: key.toLowerCase(), val });
  }
  if (!pairs.length) {
    toast("No valid pairs found", "error");
    return;
  }
  const kvList = editor.querySelector(".dict-kv-list");
  [...kvList.querySelectorAll(".dict-kv-row")].forEach((r) => {
    const k = r.querySelector(".dict-key").value.trim();
    const v = r.querySelector(".dict-val").value.trim();
    if (!k && !v) r.remove();
  });
  let added = 0,
    updated = 0;
  pairs.forEach(({ key, val }) => {
    let found = false;
    kvList.querySelectorAll(".dict-kv-row").forEach((row) => {
      if (row.querySelector(".dict-key").value.trim().toLowerCase() === key) {
        row.querySelector(".dict-val").value = val;
        found = true;
        updated++;
      }
    });
    if (!found) {
      const row = document.createElement("div");
      row.className = "dict-kv-row";
      row.innerHTML = dictKvRow(key, val);
      kvList.appendChild(row);
      added++;
    }
  });
  textarea.value = "";
  toast(
    added > 0
      ? `✅ ${added} added${updated > 0 ? `, ${updated} updated` : ""}`
      : `✅ ${updated} updated`,
    "success",
  );
}

// ─── UTILS ─────────────────────────────────────────────────────────────────
function esc(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
function escAttr(str) {
  return String(str || "")
    .replace(/'/g, "\\'")
    .replace(/\\/g, "\\\\");
}
function copyToClipboard(text) {
  const fallback = () => {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand("copy");
    } catch (e) {}
    document.body.removeChild(ta);
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).catch(fallback);
  } else {
    fallback();
  }
}

function formatDateTime(ms) {
  if (!ms) return "";
  const d = new Date(ms);
  const now = new Date();
  const options = {
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  };
  if (d.getFullYear() !== now.getFullYear()) options.year = "numeric";
  return d.toLocaleString("en-US", options);
}

function getNormalizedFilename(name) {
  if (!name) return "";
  // Strip extension if present, trim, replace spaces with underscores
  return (
    name
      .replace(/\.xlsx$/i, "")
      .trim()
      .replace(/\s+/g, "_") + ".xlsx"
  );
}

// ─── DEDUPLICATION UTILITIES ───
function getLeadKey(row, dedupCols) {
  const getVal = (r, k) => {
    if (r[k] !== undefined) return String(r[k]).trim().toLowerCase();
    // Search case-insensitive
    const alt = Object.keys(r).find(
      (ak) => ak.toLowerCase() === k.toLowerCase(),
    );
    return alt ? String(r[alt]).trim().toLowerCase() : "";
  };

  // Filter out ALL system metadata: __dmp_ prefixes and [System] brackets
  const isSystemKey = (k) =>
    k.startsWith("__dmp") || k.includes("[") || k.includes("]");

  if (dedupCols && dedupCols.length > 0) {
    return dedupCols.map((k) => getVal(row, k)).join("|");
  }

  // Fallback: Use all non-system columns, sorted for stability
  return Object.entries(row)
    .filter(([k]) => !isSystemKey(k))
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([, v]) =>
      String(v || "")
        .trim()
        .toLowerCase(),
    )
    .join("|");
}

async function mergeGroupData(group, template, incremental, targetName = "") {
  let allRows = [];
  let metrics = { loaded: 0, deduped: 0, baseline: 0 };
  let duplicatedRows = [];

  for (const src of group.sources) {
    const pagesToInclude =
      template.pages && template.pages.length > 0
        ? template.pages.join(",")
        : src.pages || "";

    let sourceData = []; // Array of {name, rows}
    if (src.path.startsWith("http")) {
      sourceData = await fetchExternalData(src.path, template, pagesToInclude);
    } else {
      sourceData = await readLocalData(src.path, template, pagesToInclude);
    }

    for (const page of sourceData) {
      let pageRows = page.rows.map((r, i) => ({
        ...r,
        __dmp_src: src.label || src.path,
        __dmp_page: page.name,
        __dmp_idx: i + 2,
      }));

      metrics.loaded += pageRows.length;

      // --- Sheet-Level Deduplication ---
      // 1. Incremental (Baseline) check for THIS sheet
      if (incremental) {
        const baselineData = await loadBaseline(
          group.name,
          targetName,
          template.name,
        );
        if (baselineData && baselineData.length > 0) {
          const dedupCols = group.dedupCols || template.dedupCols || [];
          const baselineMap = new Map();
          baselineData.forEach((r) => {
            if (r.__dmp_page === page.name || !r.__dmp_page) {
              const key = getLeadKey(r, dedupCols);
              if (key && !baselineMap.has(key))
                baselineMap.set(key, r.__dmp_idx);
            }
          });

          const nextRows = [];
          for (const r of pageRows) {
            const key = getLeadKey(r, dedupCols);
            if (key && baselineMap.has(key)) {
              metrics.baseline++;
              duplicatedRows.push({
                source: `${r.__dmp_src} (Sheet: ${page.name})`,
                row_index: r.__dmp_idx,
                primary_source: `Historical Baseline (Sheet: ${group.name})`,
                primary_index: baselineMap.get(key),
                row_data: r,
                key: key,
                type: "baseline",
                sheetName: group.name,
              });
            } else {
              nextRows.push(r);
            }
          }
          pageRows = nextRows;
        }
      }

      const dedupCols = group.dedupCols || template.dedupCols || [];
      if (dedupCols.length > 0) {
        const seenMap = new Map();
        const unique = [];
        for (const r of pageRows) {
          const key = getLeadKey(r, dedupCols);
          const isEmpty = !key || key.replace(/\|/g, "").trim() === "";
          if (isEmpty) {
            unique.push(r);
          } else if (!seenMap.has(key)) {
            seenMap.set(key, {
              src: r.__dmp_src,
              idx: r.__dmp_idx,
              page: r.__dmp_page,
            });
            unique.push(r);
          } else {
            metrics.deduped++;
            const primary = seenMap.get(key);
            duplicatedRows.push({
              source: `${r.__dmp_src} (Sheet: ${page.name})`,
              row_index: r.__dmp_idx,
              primary_source: `${primary.src} (Sheet: ${primary.page})`,
              primary_index: primary.idx,
              row_data: r,
              key: key,
              type: "internal",
              sheetName: group.name,
            });
          }
        }
        pageRows = unique;
      }

      allRows.push(...pageRows);
    }
  }

  // 2. Sorting
  if (template.sortCol) {
    const isDesc = template.sortOrder === "desc";
    allRows.sort((a, b) => {
      const vA = a[template.sortCol];
      const vB = b[template.sortCol];
      if (vA < vB) return isDesc ? 1 : -1;
      if (vA > vB) return isDesc ? -1 : 1;
      return 0;
    });
  }

  return { rows: allRows, metrics, duplicatedRows };
}
// fetchExternalData: unified implementation exists later in file

async function readLocalData(path, template, pagesToInclude = "") {
  if (!STATE.folderHandle) return [];
  try {
    const parts = path.replace(/\\/g, "/").split("/");
    let target = STATE.folderHandle;
    for (let i = 0; i < parts.length - 1; i++) {
      if (parts[i]) {
        target = await target.getDirectoryHandle(parts[i]).catch(() => null);
        if (!target) {
          console.warn(
            `[readLocalData] Dir not found: "${parts[i]}" in "${path}"`,
          );
          return [];
        }
      }
    }
    const fileHandle = await target
      .getFileHandle(parts[parts.length - 1])
      .catch(() => null);
    if (!fileHandle) {
      console.warn(`[readLocalData] File not found: "${path}"`);
      return [];
    }
    const file = await fileHandle.getFile();
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(ab), { type: "array" });

    const output = [];
    const allowedPages = pagesToInclude
      ? pagesToInclude
          .split(",")
          .map((p) => p.trim().toLowerCase())
          .filter(Boolean)
      : [];

    for (const name of wb.SheetNames) {
      if (
        allowedPages.length > 0 &&
        !allowedPages.includes(name.trim().toLowerCase())
      ) {
        continue;
      }
      const ws = wb.Sheets[name];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (!json || json.length < 2) {
        continue;
      }
      const mapped = mapSourceToJson(json, template);

      if (mapped && mapped.length) {
        output.push({ name, rows: mapped });
      }
    }

    return output;
  } catch (e) {
    return [];
  }
}
function mapSourceToJson(rawJson, template) {
  if (rawJson.length < 1) return [];
  const headers = rawJson[0]
    ? rawJson[0].map((h) =>
        String(h || "")
          .toLowerCase()
          .trim(),
      )
    : [];
  const rows = rawJson.slice(1);

  // Skip truly empty rows to ensure accurate count
  const dataRows = rows.filter(
    (r) =>
      r &&
      r.length > 0 &&
      r.some((c) => c !== null && c !== undefined && String(c).trim() !== ""),
  );

  return dataRows.map((row) => {
    const mapped = {};
    template.columns.forEach((col) => {
      let val = "";
      if (col.src === "0" || !col.src) {
        val = "";
      } else {
        const srcNames = col.src
          .replace(/^\[|\]$/g, "")
          .split(",")
          .map((s) => s.trim().toLowerCase());
        for (const s of srcNames) {
          const idx = headers.indexOf(s);
          if (idx !== -1 && row[idx] !== undefined) {
            val = row[idx];
            break;
          }
        }
      }

      const tokens = col.fmt ? col.fmt.split(/\s+/).filter(Boolean) : [];
      tokens.forEach((code) => {
        val = applyJSFormat(val, code, col.dict);
      });

      mapped[col.name] = val;
    });
    return mapped;
  });
}

async function loadBaseline(groupName, targetName = "") {
  if (!STATE.folderHandle) return [];
  try {
    const tDir = await STATE.folderHandle
      .getDirectoryHandle("templates")
      .catch(() => null);
    if (!tDir) return [];
    const bDir = await tDir.getDirectoryHandle("baselines").catch(() => null);
    if (!bDir) return [];

    // 1. Try TargetName.xlsx (Campaign or custom run baseline)
    if (targetName) {
      const sanitized = targetName.replace(/[^\w\-]/g, "_");
      const hTarget = await bDir
        .getFileHandle(`${sanitized}.xlsx`)
        .catch(() => null);
      if (hTarget) {
        const file = await hTarget.getFile();
        const ab = await file.arrayBuffer();
        const wb = XLSX.read(new Uint8Array(ab), { type: "array" });
        // Match group name or truncated sheet name (31 chars)
        const sn = wb.SheetNames.find(
          (n) =>
            n.toLowerCase() === groupName.toLowerCase() ||
            n.toLowerCase() === groupName.substring(0, 31).toLowerCase(),
        );
        if (sn) return XLSX.utils.sheet_to_json(wb.Sheets[sn]);
      }
    }

    // 2. Try GroupName.xlsx (Standard standalone baseline)
    const hGroup = await bDir
      .getFileHandle(`${groupName}.xlsx`)
      .catch(() => null);
    if (hGroup) {
      const file = await hGroup.getFile();
      const ab = await file.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(ab), { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return XLSX.utils.sheet_to_json(ws);
    }

    return [];
  } catch (e) {
    console.warn("Baseline load error:", e);
    return [];
  }
}

// ─── DIRECT MERGE EXECUTION ────────────────────────────────────────────────
function resetRunForm() {
  // Mode 1 (Template)
  const templateSelect = document.getElementById("runTemplate");
  if (templateSelect) templateSelect.value = "";
  const templateInfo = document.getElementById("runTemplate1Info");
  if (templateInfo) templateInfo.innerHTML = "";
  const extFiles = document.getElementById("runExtFiles1");
  if (extFiles) extFiles.innerHTML = "";
  const outName = document.getElementById("runOutputName1");
  if (outName) outName.value = "";
  const incCheck1 = document.getElementById("runIncrementalCheck1");
  if (incCheck1) incCheck1.checked = false;

  // Mode 4 (Campaign)
  const campSelect = document.getElementById("runCampaignConfig");
  if (campSelect) campSelect.value = "";
  const campPreview = document.getElementById("campPreview");
  if (campPreview) campPreview.innerHTML = "";
  const incCheck = document.getElementById("runIncrementalCheck");
  if (incCheck) incCheck.checked = false;

  // Reset Input Source to Folder (Default)
  const folderRadio = document.querySelector(
    'input[name="runSrc1"][value="folder"]',
  );
  if (folderRadio) {
    folderRadio.checked = true;
    if (typeof toggleExtFiles1 === "function") toggleExtFiles1();
  }

  // Reset Processing Mode to Quick (Default)
  const quickRadio = document.querySelector(
    'input[name="runProcMode"][value="quick"]',
  );
  if (quickRadio) {
    quickRadio.checked = true;
    const advOpts = document.getElementById("runAdvancedOpts1");
    if (advOpts) advOpts.classList.add("hidden");
  }

  refreshIcons();
}

async function runDirectMerge() {
  const mode = STATE.currentMode;
  const progressWrap = document.getElementById("mergeProgressWrap");
  const progressFill = document.getElementById("mergeProgressFill");
  const statusText = document.getElementById("mergeStatusText");
  const btn = document.getElementById("btnDirectExecute");

  if (!progressWrap || !progressFill || !statusText || !btn) return;

  // 1. Validate Config
  let config = { groups: [], incremental: false, name: "" };
  try {
    if (mode === 1) {
      const tname = document.getElementById("runTemplate").value;
      if (!tname) {
        toast("Select a template", "error");
        return;
      }
      const template = STATE.templates.find((t) => t.name === tname);

      let sources = [];
      const srcType = document.querySelector('[name="runSrc1"]:checked').value;
      if (srcType === "folder") {
        const subDir = tname.replace(/[^\w\-]/g, "_");
        if (STATE.folderHandle) {
          const inputDir = await STATE.folderHandle
            .getDirectoryHandle("input")
            .catch(() => null);
          if (inputDir) {
            const tplSubDir = await inputDir
              .getDirectoryHandle(subDir)
              .catch(() => null);
            if (tplSubDir) {
              for await (const [name, handle] of tplSubDir.entries()) {
                if (
                  handle.kind === "file" &&
                  /\.(csv|xlsx|xls|tsv|txt)$/i.test(name)
                ) {
                  sources.push({
                    path: `input/${subDir}/${name}`,
                    label: name,
                  });
                }
              }
            }
          }
        }
        if (sources.length === 0) {
          const allInput = await autoDetectInputFiles();
          const slug = tname.toLowerCase().replace(/[^\w]/g, "");
          sources = allInput
            .filter((f) => {
              const fn = f.toLowerCase();
              const hasTplName =
                fn.includes(slug) || fn.includes(tname.toLowerCase());
              const isOldOutput = /_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}/.test(fn);
              return hasTplName && !isOldOutput;
            })
            .map((f) => ({ path: `input/${f}`, label: f }));
        }
      } else {
        sources = [
          ...document
            .getElementById("runExtFiles1")
            .querySelectorAll(".smart-src-row"),
        ]
          .map((el) => ({
            label: el.querySelector(".ext-file-label")?.value.trim() || "",
            path: el.querySelector(".ext-file-path")?.value.trim() || "",
            pages: el.querySelector(".src-pages-input")?.value.trim() || "",
          }))
          .filter((s) => s.path);
      }

      if (sources.length === 0) {
        toast("No sources found to merge", "error");
        return;
      }

      const isAdvanced =
        document.querySelector('[name="runProcMode"]:checked').value ===
        "advanced";
      const outName = isAdvanced
        ? document.getElementById("runOutputName1").value.trim() || tname
        : tname;
      const incremental = document.getElementById(
        "runIncrementalCheck1",
      ).checked;

      let dedupCols = template.dedupCols || [];
      if (isAdvanced) {
        const dedupType = document.querySelector(
          '[name="runDedup1"]:checked',
        ).value;
        if (dedupType === "custom") {
          dedupCols = [
            ...document.querySelectorAll(".run-custom-dedup-check:checked"),
          ].map((i) => i.value);
        } else if (dedupType === "skip") {
          dedupCols = [];
        }
      }

      config = {
        name: outName,
        groups: [{ name: outName, sources, templateName: tname, dedupCols }],
        incremental,
      };
    } else if (mode === 4) {
      const campName = document.getElementById("runCampaignConfig").value;
      if (!campName) {
        toast("Select a campaign", "error");
        return;
      }
      const camp = STATE.campaigns.find((c) => c.name === campName);
      if (!camp) {
        toast("Campaign not found", "error");
        return;
      }
      const incremental = document.getElementById(
        "runIncrementalCheck",
      ).checked;

      config = {
        name: camp.name,
        groups: camp.groups
          .map((gname) => STATE.groups.find((g) => g.name === gname))
          .filter(Boolean),
        incremental,
      };
      if (config.groups.length === 0) {
        toast("No valid groups in campaign", "error");
        return;
      }
    }
  } catch (e) {
    toast("Config error: " + e.message, "error");
    return;
  }

  // 2. Start Execution
  btn.disabled = true;
  btn.innerHTML = '<i data-lucide="loader-2" class="icon spin"></i> Running...';
  progressWrap.style.display = "block";
  progressFill.style.width = "0%";
  statusText.textContent = "Initializing...";
  refreshIcons();

  try {
    await loadSheetJS();
    const wb = XLSX.utils.book_new();
    let totalLoaded = 0;
    let totalDeduped = 0;
    let totalBaseline = 0;
    let totalGroups = config.groups.length;
    let completedGroups = 0;

    for (const group of config.groups) {
      statusText.textContent = `Processing ${group.name}...`;
      const template = STATE.templates.find(
        (t) => t.name === group.templateName,
      );
      if (!template) {
        console.warn(`Template not found: ${group.templateName}`);
        continue;
      }

      const result = await mergeGroupData(
        group,
        template,
        config.incremental,
        config.name,
      );
      const rows = result.rows;
      group.rows = rows; // Store for analytics
      group.duplicatedRows = result.duplicatedRows; // Store for history

      totalLoaded += result.metrics.loaded;
      totalDeduped += result.metrics.deduped;
      totalBaseline += result.metrics.baseline;

      // Clean up internal metadata before Excel export
      const exportRows = rows.map((r) => {
        const clean = { ...r };
        if (clean.__dmp_page) {
          clean["[System] Source Sheet"] = clean.__dmp_page;
        }
        delete clean.__dmp_src;
        delete clean.__dmp_idx;
        delete clean.__dmp_page;
        return clean;
      });

      const ws = XLSX.utils.json_to_sheet(exportRows);
      const colWidths = template.columns.map((c) => ({
        wch: Math.max(c.name.length + 4, 15),
      }));
      ws["!cols"] = colWidths;
      applyCellStyles(ws, template);

      XLSX.utils.book_append_sheet(wb, ws, group.name.substring(0, 31));

      completedGroups++;
      progressFill.style.width = `${(completedGroups / totalGroups) * 100}%`;
    }

    if (completedGroups === 0) throw new Error("No data was processed");

    // 3. Save Output
    const ts = new Date()
      .toISOString()
      .replace(/[:.]/g, "-")
      .slice(0, 16)
      .replace("T", "_");
    const filename = `${config.name.replace(/[^\w\-]/g, "_")}_${ts}.xlsx`;

    let stats = `${totalLoaded} rows loaded`;
    if (totalDeduped > 0) stats += `, ${totalDeduped} duplicates filtered`;
    if (totalBaseline > 0) stats += `, ${totalBaseline} already in baseline`;

    if (STATE.folderHandle) {
      const outDir = await STATE.folderHandle.getDirectoryHandle("output", {
        create: true,
      });
      const fileHandle = await outDir.getFileHandle(filename, {
        create: true,
      });
      const writable = await fileHandle.createWritable();
      const ab = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      await writable.write(ab);
      await writable.close();

      await scanOutputFolder(false);
      // Auto-open newest file
      if (STATE._outputFiles && STATE._outputFiles.length > 0) {
        const latest = STATE._outputFiles[0];
        STATE._activeOutputName = latest.name; // Keep state in sync for Save Baseline
        displayOutputFile(await latest.handle.getFile());
        navigate("viewer");
      }

      toast(
        `Merged ${filename}. ${stats}\nFiles included: ${config.groups.flatMap((g) => g.sources.map((s) => s.label)).join(", ")}`,
        "success",
      );
    } else {
      XLSX.writeFile(wb, filename);
      toast(`Merge complete! ${stats}`, "success");
    }

    STATE.runs++;
    save();

    // Save Duplication History
    const allDupes = config.groups.flatMap((g) =>
      (g.duplicatedRows || []).map((d) => ({ ...d, sheetName: g.name })),
    );
    if (allDupes.length > 0) {
      const dupFilename = filename.replace(".xlsx", "_duplicates.json");
      writeToLinkedFolder("output", dupFilename, allDupes);
    }

    // Update Analytics
    updateDailyAnalytics(config.name, config.groups);

    refreshHome();
    resetRunForm();
    statusText.textContent = "";
    progressWrap.style.display = "none";
  } catch (err) {
    console.error(err);
    toast(`Merge failed: ${err.message}`, "error");
    statusText.textContent = "Error occurred";
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<i data-lucide="zap"></i> Direct Execute';
    refreshIcons();
  }
}

// (Duplicate mergeGroupData removed)

async function fetchExternalData(url, template, pagesToInclude = "") {
  let fetchUrl = url;
  let isGSheet = false;
  let token = null;

  if (url.includes("docs.google.com/spreadsheets")) {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) {
      const sheetId = match[1];
      fetchUrl = `https://www.googleapis.com/drive/v3/files/${sheetId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
      isGSheet = true;
      try {
        token = await getGoogleAccessToken();
      } catch (e) {}
    }
  }

  try {
    await loadSheetJS();
    let resp;
    if (isGSheet && token) {
      resp = await fetch(fetchUrl, {
        headers: { Authorization: `Bearer ${token}` },
      });
    } else {
      resp = await fetch(fetchUrl, { cache: "no-store" });
    }

    if (!resp.ok) {
      return [];
    }
    const ab = await resp.arrayBuffer();

    if (ab.byteLength < 5) return [];

    let wb;
    try {
      wb = XLSX.read(new Uint8Array(ab), { type: "array" });
    } catch (parseErr) {
      // Fallback: try reading as CSV text
      try {
        const text = new TextDecoder().decode(ab);

        wb = XLSX.read(text, { type: "string" });
      } catch (csvErr) {
        console.error(
          "[fetchExternalData] CSV fallback also failed:",
          csvErr.message,
        );
        return [];
      }
    }

    const output = [];
    const allowedPages = pagesToInclude
      ? pagesToInclude
          .split(",")
          .map((p) => p.trim().toLowerCase())
          .filter(Boolean)
      : [];

    for (const sheetName of wb.SheetNames) {
      if (
        allowedPages.length > 0 &&
        !allowedPages.includes(sheetName.trim().toLowerCase())
      ) {
        continue;
      }
      const ws = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1 });

      if (!json || json.length < 2) continue;
      const mapped = mapSourceToJson(json, template);

      if (mapped && mapped.length) {
        output.push({ name: sheetName, rows: mapped });
      }
    }

    return output;
  } catch (e) {
    return [];
  }
}
// (Duplicates removed)

async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    toast("Copied to clipboard", "success");
  } catch (err) {
    toast("Copy failed", "error");
  }
}

function toggleExpandOutputViewer() {
  const viewer = document.getElementById("page-viewer");
  if (!viewer) return;
  viewer.classList.toggle("viewer-expanded");
  const icon = document.querySelector("#expandBtn i");
  if (icon) {
    if (viewer.classList.contains("viewer-expanded")) {
      icon.setAttribute("data-lucide", "minimize-2");
    } else {
      icon.setAttribute("data-lucide", "maximize-2");
    }
    refreshIcons();
  }
}

async function addOutputToInput(filename, isBaseline) {
  if (!STATE.folderHandle) {
    toast("Link project folder first", "error");
    return;
  }
  try {
    const file = STATE._outputFiles.find((f) => f.name === filename);
    if (!file) {
      toast("File not found", "error");
      return;
    }

    const data = await file.handle.getFile();
    const inputDir = await STATE.folderHandle.getDirectoryHandle("input", {
      create: true,
    });
    const newFile = await inputDir.getFileHandle(filename, {
      create: true,
    });
    const writable = await newFile.createWritable();
    await writable.write(data);
    await writable.close();

    toast(`File "${filename}" added to input/`, "success");
    refreshInputPage();
  } catch (e) {
    toast("Add failed: " + e.message, "error");
  }
}

// ─── HELPERS ────────────────────────────────────────────────────────────────
async function listFolderFiles(subfolder) {
  if (!STATE.folderHandle) return [];
  try {
    const dirHandle = await STATE.folderHandle.getDirectoryHandle(subfolder);
    const files = [];
    for await (const entry of dirHandle.values()) {
      if (entry.kind === "file") {
        const file = await entry.getFile();
        files.push({ name: entry.name, size: file.size });
      }
    }
    return files;
  } catch (e) {
    console.error("Error listing folder:", e);
    return [];
  }
}

// ─── REPORTS ───────────────────────────────────────────────────────────────
async function refreshReportsFiles() {
  let files = [];
  if (!STATE.folderHandle) {
    try {
      const res = await fetch("/api/list-input");
      files = await res.json();
    } catch (e) {
      document.getElementById("reportsFileList").innerHTML =
        '<div class="empty-state" style="padding:40px"><strong>No folder linked</strong><p style="margin-bottom:15px;color:var(--text3)">Link your project folder or add files to input/ folder.</p><button class="btn btn-primary" onclick="linkFolder()"><i data-lucide="link" class="icon"></i> Connect Project Folder</button></div>';
      refreshIcons();
      return;
    }
  } else {
    files = await listFolderFiles("input");
  }
  STATE._inputFiles = files; // Sync for consistency
  const list = document.getElementById("reportsFileList");
  if (!files.length) {
    list.innerHTML =
      '<div class="empty-state">No files found in input/ folder.</div>';
    return;
  }

  list.innerHTML = files
    .map((f) => {
      const isSelected = STATE.reports.selectedFiles.includes(f.name);
      return `
      <div class="list-card ${isSelected ? "active" : ""}" onclick="toggleReportFile('${escAttr(f.name)}')">
        <div class="lc-left">
          <div class="lc-icon ${f.name.endsWith(".csv") ? "blue" : "green"}">
            <i data-lucide="file" class="icon"></i>
          </div>
          <div>
            <div class="lc-title">${esc(f.name)}</div>
            <div class="lc-sub">${(f.size / 1024).toFixed(1)} KB</div>
          </div>
        </div>
        <div class="lc-actions">
           <i data-lucide="${isSelected ? "check-circle" : "circle"}" class="icon ${isSelected ? "text-success" : ""}"></i>
        </div>
      </div>
    `;
    })
    .join("");

  // Add scan button if any files selected
  if (STATE.reports.selectedFiles.length > 0) {
    list.innerHTML += `
      <div style="margin-top:15px">
        <button class="btn btn-primary" style="width:100%" onclick="scanReportsCampaigns()">
          <i data-lucide="search" class="icon"></i> Scan Selected Files for Campaigns
        </button>
      </div>
    `;
  }
  refreshIcons();
}

function toggleReportFile(name) {
  const idx = STATE.reports.selectedFiles.indexOf(name);
  if (idx >= 0) {
    STATE.reports.selectedFiles.splice(idx, 1);
  } else {
    STATE.reports.selectedFiles.push(name);
  }
  refreshReportsFiles();
  // Hide step 2 if files changed
  document.getElementById("reportsStep2").style.display = "none";
}

async function scanReportsCampaigns() {
  if (STATE.reports.selectedFiles.length === 0) return;
  toast("Scanning files...", "info");

  const filesParam = STATE.reports.selectedFiles
    .map((f) => `input/${f}`)
    .join(",");
  try {
    const res = await fetch(
      `/api/run-report?files=${encodeURIComponent(filesParam)}&list=true`,
    );
    const data = await res.json();

    if (data.error) throw new Error(data.error);

    let campaigns = [];
    try {
      // Find the first [ and last ] to extract JSON array
      const start = data.output.indexOf("[");
      const end = data.output.lastIndexOf("]");
      if (start !== -1 && end !== -1) {
        campaigns = JSON.parse(data.output.substring(start, end + 1));
      } else {
        campaigns = JSON.parse(data.output);
      }
    } catch (e) {
      console.error("JSON Parse Error:", e, data.output);
      throw new Error("Could not parse campaign list from script output.");
    }

    STATE.reports.campaigns = campaigns;
    STATE.reports.ordering = campaigns.map((_, i) => i);

    document.getElementById("reportsStep2").style.display = "block";
    renderCampaignOrdering();
    document
      .getElementById("reportsStep2")
      .scrollIntoView({ behavior: "smooth" });
  } catch (err) {
    console.error(err);
    toast("Scan failed: " + err.message, "error");
  }
}

function renderCampaignOrdering() {
  const list = document.getElementById("campaignOrderList");
  const campaigns = STATE.reports.campaigns;
  const order = STATE.reports.ordering;

  list.innerHTML = order
    .map((idx, i) => {
      const name = campaigns[idx];
      return `
      <div class="campaign-order-item" draggable="true" ondragstart="handleCampDragStart(event, ${i})" ondragover="handleCampDragOver(event)" ondrop="handleCampDrop(event, ${i})">
        <div class="drag-handle"><i data-lucide="grip-vertical" class="icon"></i></div>
        <div class="campaign-index">${i + 1}</div>
        <div class="campaign-name">${esc(name)}</div>
      </div>
    `;
    })
    .join("");
  refreshIcons();
}

let _dragCampIdx = null;
function handleCampDragStart(e, i) {
  _dragCampIdx = i;
  e.dataTransfer.effectAllowed = "move";
}
function handleCampDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = "move";
}
function handleCampDrop(e, i) {
  e.preventDefault();
  if (_dragCampIdx === null || _dragCampIdx === i) return;

  const order = STATE.reports.ordering;
  const movedIdx = order.splice(_dragCampIdx, 1)[0];
  order.splice(i, 0, movedIdx);

  _dragCampIdx = null;
  renderCampaignOrdering();
}

async function generateFinalReport() {
  const filesParam = STATE.reports.selectedFiles
    .map((f) => `input/${f}`)
    .join(",");
  const orderParam = STATE.reports.ordering.map((i) => i + 1).join(","); // Python script uses 1-based indices

  toast("Generating report...", "info");
  document.getElementById("reportsOutput").style.display = "block";
  document.getElementById("reportsResultMsg").textContent =
    "⏳ Processing Python script...";
  document.getElementById("btnDownloadReport").style.display = "none";

  try {
    const res = await fetch(
      `/api/run-report?files=${encodeURIComponent(filesParam)}&order=${orderParam}`,
    );
    const data = await res.json();

    if (data.error) throw new Error(data.error);

    const outMatch = data.output.match(/SUCCESS: (.*)/);
    if (outMatch) {
      const outPath = outMatch[1];
      document.getElementById("reportsResultMsg").innerHTML =
        `<div class="text-success" style="font-weight:600">Report Generated Successfully!</div><div style="font-size:12px;margin-top:5px">Saved as: ${esc(outPath)}</div>`;

      const btn = document.getElementById("btnDownloadReport");
      btn.style.display = "inline-flex";
      btn.onclick = () => {
        const a = document.createElement("a");
        a.href = "/" + outPath;
        a.download = outPath;
        a.click();
      };
      toast("Report generated!", "success");
    } else {
      document.getElementById("reportsResultMsg").textContent = data.output;
    }
  } catch (err) {
    document.getElementById("reportsResultMsg").innerHTML =
      `<div class="text-danger">❌ Error: ${esc(err.message)}</div>`;
    toast("Generation failed", "error");
  }
}

// ─── ANALYTICS & DUPLICATION HISTORY ───
function updateDailyAnalytics(campaignName, groups) {
  const today = new Date().toISOString().split("T")[0];
  let analytics = JSON.parse(localStorage.getItem("dmp_analytics") || "{}");

  if (!analytics[today]) analytics[today] = {};
  if (!analytics[today][campaignName]) analytics[today][campaignName] = {};

  groups.forEach((g) => {
    analytics[today][campaignName][g.name] = g.rows ? g.rows.length : 0;
  });

  localStorage.setItem("dmp_analytics", JSON.stringify(analytics));
  saveAnalyticsToFile(analytics);
  refreshAnalyticsDashboard();
}

async function saveAnalyticsToFile(analytics) {
  if (!STATE.folderHandle) return;
  try {
    const lines = [
      "DataMerge Pro - Lead Analytics",
      "Generated: " + new Date().toLocaleString(),
      "",
    ];
    for (const [date, campaigns] of Object.entries(analytics)) {
      lines.push(`--- DATE: ${date} ---`);
      for (const [camp, groups] of Object.entries(campaigns)) {
        lines.push(`Campaign: ${camp}`);
        for (const [group, count] of Object.entries(groups)) {
          lines.push(`  - ${group}: ${count} leads`);
        }
      }
      lines.push("");
    }
    const fileHandle = await STATE.folderHandle.getFileHandle(
      "leads_analytics.txt",
      { create: true },
    );
    const writable = await fileHandle.createWritable();
    await writable.write(lines.join("\n"));
    await writable.close();
  } catch (e) {
    console.warn("Could not save analytics file:", e);
  }
}

function setAnalyticsView(view) {
  STATE.analyticsView = view;
  refreshAnalyticsDashboard();
}

function refreshAnalyticsDashboard() {
  const today = new Date().toISOString().split("T")[0];
  const analytics = JSON.parse(localStorage.getItem("dmp_analytics") || "{}");
  const dash = document.getElementById("analyticsDashboard");
  const dateEl = document.getElementById("analyticsDate");
  if (!dash || !dateEl) return;

  // Calculate Bounds
  const dates = Object.keys(analytics).sort();
  const minDate = dates.length > 0 ? dates[0] : today;
  const maxDate = dates.length > 0 ? dates[dates.length - 1] : today;

  const view = STATE.analyticsView || "today";
  dateEl.innerHTML = `
    <div style="display:flex; align-items:center; gap:12px; flex-wrap: wrap;">
      <div class="toggle-group" style="display:inline-flex; background:var(--surface2); padding:2px; border-radius:6px;">
        <button class="toggle-btn ${view === "today" ? "active" : ""}" onclick="setAnalyticsView('today')" style="padding:4px 12px; font-size:10px; border:none; border-radius:4px; cursor:pointer; background:${view === "today" ? "var(--blue)" : "transparent"}; color:${view === "today" ? "white" : "var(--text3)"}; transition:all 0.2s;">Range</button>
        <button class="toggle-btn ${view === "total" ? "active" : ""}" onclick="setAnalyticsView('total')" style="padding:4px 12px; font-size:10px; border:none; border-radius:4px; cursor:pointer; background:${view === "total" ? "var(--blue)" : "transparent"}; color:${view === "total" ? "white" : "var(--text3)"}; transition:all 0.2s;">Total</button>
      </div>

      ${
        view === "today"
          ? `
        <div style="display:flex; align-items:center; gap:8px;">
          <div style="position:relative; display:flex; align-items:center; gap:5px;">
             <span style="font-size:10px; color:var(--text3); font-weight:600;">FROM</span>
             <div style="position:relative; display:flex; align-items:center;">
               <i data-lucide="calendar" style="position:absolute; left:8px; width:10px; height:10px; color:var(--text3); pointer-events:none;"></i>
               <input type="date" value="${STATE.analyticsStartDate}" 
                      min="${minDate}" max="${maxDate}"
                      onchange="setAnalyticsRange(this.value, STATE.analyticsEndDate)"
                      style="background:var(--surface2); border:1px solid var(--border); border-radius:4px; padding:2px 6px 2px 24px; font-size:10px; color:var(--text1); cursor:pointer;">
             </div>
             <span style="font-size:10px; color:var(--text3); font-weight:600;">TO</span>
             <div style="position:relative; display:flex; align-items:center;">
               <i data-lucide="calendar" style="position:absolute; left:8px; width:10px; height:10px; color:var(--text3); pointer-events:none;"></i>
               <input type="date" value="${STATE.analyticsEndDate}" 
                      min="${minDate}" max="${maxDate}"
                      onchange="setAnalyticsRange(STATE.analyticsStartDate, this.value)"
                      style="background:var(--surface2); border:1px solid var(--border); border-radius:4px; padding:2px 6px 2px 24px; font-size:10px; color:var(--text1); cursor:pointer;">
             </div>
          </div>
          <button class="btn btn-ghost" onclick="resetRangeAnalytics()" style="padding:2px 8px; font-size:9px;" title="Reset incorrect counts for this range"><span style="color:var(--red);">Reset</span></button>
        </div>
      `
          : ""
      }
    </div>
  `;

  let displayData = {}; // camp -> group -> count
  if (view === "today") {
    // Range Aggregation
    const start = STATE.analyticsStartDate;
    const end = STATE.analyticsEndDate;
    Object.entries(analytics).forEach(([date, dayData]) => {
      if (date >= start && date <= end) {
        Object.entries(dayData).forEach(([camp, groups]) => {
          if (!displayData[camp]) displayData[camp] = {};
          Object.entries(groups).forEach(([group, count]) => {
            displayData[camp][group] = (displayData[camp][group] || 0) + count;
          });
        });
      }
    });
  } else {
    // Cumulative Sum
    Object.values(analytics).forEach((dayData) => {
      Object.entries(dayData).forEach(([camp, groups]) => {
        if (!displayData[camp]) displayData[camp] = {};
        Object.entries(groups).forEach(([group, count]) => {
          displayData[camp][group] = (displayData[camp][group] || 0) + count;
        });
      });
    });
  }

  if (Object.keys(displayData).length === 0) {
    dash.innerHTML = `<div class="empty-state" style="padding: 30px;"><p>No ${view === "today" ? "runs today" : "analytics data"} yet.</p></div>`;
    return;
  }

  let html = `<table class="input-table" style="width:100%; border-collapse: collapse; font-size:12px;">
    <thead>
      <tr style="background: var(--surface2); border-bottom: 1px solid var(--border);">
        <th style="padding:10px; text-align:left;">Campaign</th>
        <th style="padding:10px; text-align:left;">Source Group</th>
        <th style="padding:10px; text-align:right;">Leads ${view === "today" ? "Today" : "Total"}</th>
      </tr>
    </thead>
    <tbody>`;

  for (const [camp, groups] of Object.entries(displayData)) {
    for (const [group, count] of Object.entries(groups)) {
      html += `<tr style="border-bottom: 1px solid var(--border);">
        <td style="padding:10px;"><strong>${esc(camp)}</strong></td>
        <td style="padding:10px;">${esc(group)}</td>
        <td style="padding:10px; text-align:right; font-weight:700; color:var(--blue);">${count}</td>
      </tr>`;
    }
  }
  html += `</tbody></table>`;
  dash.innerHTML = html;
  refreshIcons();
}

function setAnalyticsRange(start, end) {
  STATE.analyticsStartDate = start;
  STATE.analyticsEndDate = end;
  refreshAnalyticsDashboard();
}

function resetRangeAnalytics() {
  const start = STATE.analyticsStartDate;
  const end = STATE.analyticsEndDate;
  if (!confirm(`Clear lead analytics counts from ${start} to ${end}?`)) return;

  let analytics = JSON.parse(localStorage.getItem("dmp_analytics") || "{}");
  Object.keys(analytics).forEach((date) => {
    if (date >= start && date <= end) {
      delete analytics[date];
    }
  });

  localStorage.setItem("dmp_analytics", JSON.stringify(analytics));
  saveAnalyticsToFile(analytics);
  toast(`Analytics for range ${start} to ${end} reset`, "success");
  refreshAnalyticsDashboard();
}

async function displayDuplicationHistory(filename) {
  if (!STATE.folderHandle) return;
  const dupFilename = filename.replace(".xlsx", "_duplicates.json");
  try {
    const outDir = await STATE.folderHandle.getDirectoryHandle("output");
    const fileHandle = await outDir.getFileHandle(dupFilename);
    const file = await fileHandle.getFile();
    const data = JSON.parse(await file.text());

    STATE.duplicationModal.data = data;
    STATE.duplicationModal.filename = filename;

    const content = document.getElementById("duplicationContent");
    const title = document.getElementById("duplicationTitle");
    title.textContent = "Duplication History: " + filename;

    if (data.length === 0) {
      content.innerHTML = `<div class="empty-state"><p>No duplicates found in this run.</p></div>`;
    } else {
      // Group by sheetName
      const groups = {};
      data.forEach((d) => {
        const s = d.sheetName || "Default";
        if (!groups[s]) groups[s] = [];
        groups[s].push(d);
      });

      const sheetNames = Object.keys(groups);
      STATE.duplicationModal.activeSheet = sheetNames[0];

      // Build Sticky Header (Tabs + Filter)
      let html = `<div style="position:sticky; top:0; background:var(--surface1); z-index:10; padding:10px 0; border-bottom:1px solid var(--border); margin-bottom:15px;">
        <div style="display:flex; align-items:center; gap:15px; margin-bottom:15px;">
          <div style="font-size:12px; color:var(--text2); font-weight:600;">
             Total Merge Duplicates: <span style="color:var(--blue); font-size:14px;">${data.length}</span>
          </div>
          <label class="history-filter-pill" style="margin-left:auto;">
            <input type="checkbox" onchange="toggleDuplicationBaseline(this.checked)" ${STATE.duplicationModal.hideBaseline ? "checked" : ""}> 
            Hide Baseline Matches
          </label>
        </div>
        <div class="tabs-scroll">
          ${sheetNames
            .map(
              (name) => `
            <button class="tab-btn ${name === STATE.duplicationModal.activeSheet ? "active" : ""}" 
                    id="dup-tab-${name.replace(/\s+/g, "-")}"
                    onclick="setDuplicationTab('${name}')">
              ${esc(name)}
            </button>
          `,
            )
            .join("")}
        </div>
      </div>
      <div id="duplicationTabContent"></div>`;

      content.innerHTML = html;
      renderDuplicationTab();
    }

    document.getElementById("duplicationOverlay").classList.add("show");
    document.getElementById("duplicationPanel").classList.add("open");
    refreshIcons();
  } catch (e) {
    console.error(e);
    toast("No duplication history found for this file", "info");
  }
}

function setDuplicationTab(name) {
  STATE.duplicationModal.activeSheet = name;
  const content = document.getElementById("duplicationContent");
  content
    .querySelectorAll(".tab-btn")
    .forEach((btn) => btn.classList.remove("active"));
  const activeBtn = document.getElementById(
    `dup-tab-${name.replace(/\s+/g, "-")}`,
  );
  if (activeBtn) activeBtn.classList.add("active");
  renderDuplicationTab();
}

function toggleDuplicationBaseline(checked) {
  STATE.duplicationModal.hideBaseline = checked;
  renderDuplicationTab();
}

function renderDuplicationTab() {
  const { data, activeSheet, hideBaseline } = STATE.duplicationModal;
  const container = document.getElementById("duplicationTabContent");
  if (!container || !activeSheet) return;

  let items = data.filter((d) => (d.sheetName || "Default") === activeSheet);
  const totalInSheet = items.length;

  if (hideBaseline) {
    items = items.filter((d) => d.type !== "baseline");
  }

  const baselineCount = items.filter((d) => d.type === "baseline").length; // Should be 0 if hidden
  const internalCount =
    totalInSheet -
    data.filter(
      (d) =>
        (d.sheetName || "Default") === activeSheet && d.type === "baseline",
    ).length;

  let html = `<div style="margin-bottom:12px; display:flex; gap:10px; font-size:11px;">
    <div style="padding:4px 8px; border-radius:4px; background:rgba(251, 191, 36, 0.1); color:#78350f; font-weight:600; border:1px solid rgba(251, 191, 36, 0.2);">
      ${internalCount} New Duplicates
    </div>
    <div style="padding:4px 8px; border-radius:4px; background:rgba(59, 130, 246, 0.1); color:#3b82f6; font-weight:600; border:1px solid rgba(59, 130, 246, 0.2);">
      ${totalInSheet - internalCount} Matches in History
    </div>
  </div>`;

  if (items.length === 0) {
    html += `<div class="empty-state" style="padding:40px;"><p>${hideBaseline ? "No new duplicates found in this sheet (showing 0 of " + totalInSheet + " total)" : "No duplicates found"}.</p></div>`;
  } else {
    html += `<table class="input-table" style="width:100%; border-collapse: collapse; font-size:11px; margin-bottom:24px;">
      <thead>
        <tr style="background: var(--surface2); border-bottom: 2px solid var(--border);">
          <th style="padding:10px; text-align:left; width:140px;">Removed Row</th>
          <th style="padding:10px; text-align:left; width:140px;">Original Match</th>
          <th style="padding:10px; text-align:left;">Duplicated Lead Details</th>
        </tr>
      </thead>
      <tbody>`;

    items.forEach((d) => {
      const rowData = d.row_data || {};
      const displayData = Object.entries(rowData)
        .filter(([k]) => !k.startsWith("__dmp"))
        .map(
          ([k, v]) =>
            `<div><span style="color:var(--text3); font-weight:600;">${esc(k)}:</span> ${esc(String(v))}</div>`,
        )
        .join("");

      html += `<tr style="border-bottom: 1px solid var(--border);">
        <td style="padding:10px; vertical-align:top; border-right:1px solid var(--border);">
          <div style="font-weight:700; color:var(--red);">Row ${d.row_index}</div>
          <div style="font-size:10px; color:var(--text3); margin-top:2px;">
            ${esc(d.source).replace(/\(Sheet: (.*?)\)/g, '<span style="color:var(--blue); font-weight:600;">(Sheet: $1)</span>')}
          </div>
          <div style="margin-top:8px;">
            <span style="font-size:9px; padding:2px 6px; border-radius:4px; font-weight:700; text-transform:uppercase; 
              ${d.type === "baseline" ? "background: #3b82f6; color: white;" : "background: #fbbf24; color: #78350f;"}">
              ${d.type === "baseline" ? "Matched in History" : "Duplicate in Sheet"}
            </span>
          </div>
        </td>
        <td style="padding:10px; vertical-align:top; border-right:1px solid var(--border);">
          <div style="font-weight:700; color:var(--green);">${d.type === "baseline" ? "Baseline Row" : "Existing Row"} ${d.primary_index}</div>
          <div style="font-size:10px; color:var(--text3); margin-top:2px;">
            ${esc(d.primary_source).replace(/\(Sheet: (.*?)\)/g, '<span style="color:var(--blue); font-weight:600;">(Sheet: $1)</span>')}
          </div>
        </td>
        <td style="padding:10px; vertical-align:top;">
          <div style="max-height:120px; overflow-y:auto; padding:5px; background:var(--surface2); border-radius:4px; line-height:1.4;">
            ${displayData}
          </div>
        </td>
      </tr>`;
    });
    html += `</tbody></table>`;
  }
  container.innerHTML = html;
}

function closeDuplicationModal() {
  document.getElementById("duplicationOverlay").classList.remove("show");
  document.getElementById("duplicationPanel").classList.remove("open");
}

// ─── ROW SEARCH ───
// ─── ROW SEARCH ───
// ─── ROW SEARCH ───
async function jumpToRow(rowNum) {
  const num = parseInt(rowNum);
  if (isNaN(num) || num < 1) {
    toast("Invalid row number", "error");
    return;
  }
  const rows = STATE._currentRows || [];
  const maxRow = rows.length;
  if (num > maxRow) {
    toast(`Row ${num} not found (Total data rows: ${maxRow})`, "info");
    return;
  }

  // Ensure row is rendered (Lazy loading)
  // We need to loop until the desired row number is within the rendered count
  let safety = 0;
  while (
    STATE._renderedRowCount < num &&
    STATE._renderedRowCount < rows.length &&
    safety < 100
  ) {
    renderMoreRows();
    safety++;
  }

  // Robust jump with retries
  const attemptJump = (attempts = 0) => {
    const el = document.getElementById(`preview-row-${num}`);
    if (el) {
      el.scrollIntoView({ behavior: "auto", block: "center" });

      // Visual feedback: Strong Green High-Visibility Flash
      // Use !important style to override any potential cell backgrounds
      const cells = el.querySelectorAll("td");
      cells.forEach((td) => {
        td.style.setProperty(
          "background-color",
          "rgba(16, 185, 129, 0.4)",
          "important",
        );
        td.style.setProperty("transition", "none", "important");
      });

      setTimeout(() => {
        cells.forEach((td) => {
          td.style.setProperty(
            "transition",
            "background-color 2s ease",
            "important",
          );
          td.style.setProperty("background-color", "", "");
        });
      }, 3000);

      const input = document.getElementById("jumpToRowInput");
      if (input) input.value = "";
    } else if (attempts < 5) {
      setTimeout(() => attemptJump(attempts + 1), 100);
    } else {
      toast(`Could not focus row ${num}`, "error");
    }
  };

  attemptJump();
}

// ─── BASELINE LOADING ───
async function loadBaseline(groupName, targetName, templateName = "") {
  if (!STATE.folderHandle) return [];

  try {
    const templatesDir = await STATE.folderHandle.getDirectoryHandle(
      "templates",
      { create: true },
    );
    const baselinesDir = await templatesDir.getDirectoryHandle("baselines", {
      create: true,
    });

    // Search sequence: 1. Normalized Campaign/Output, 2. Normalized Group, 3. Normalized Template
    const filenames = [
      getNormalizedFilename(targetName),
      getNormalizedFilename(groupName),
      getNormalizedFilename(templateName),
    ].filter(Boolean);

    let fileHandle = null;
    let foundName = "";

    // De-duplicate names to avoid redundant checks
    const uniqueFilenames = [...new Set(filenames)];

    for (const fname of uniqueFilenames) {
      try {
        fileHandle = await baselinesDir.getFileHandle(fname);
        foundName = fname;
        break;
      } catch (e) {
        /* continue search */
      }
    }

    if (!fileHandle) return [];

    const file = await fileHandle.getFile();
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: "array" });

    // Choose the best matching sheet
    // Priority: 1. Exact group name, 2. Exact template name, 3. First sheet
    let sheetName = wb.SheetNames.includes(groupName)
      ? groupName
      : wb.SheetNames.includes(templateName)
        ? templateName
        : wb.SheetNames[0];
    if (!sheetName) return [];

    const ws = wb.Sheets[sheetName];
    // Use defval: "" to ensure identical parsing with live data
    return XLSX.utils.sheet_to_json(ws, { defval: "" }).map((r, i) => {
      const mapped = { ...r, __dmp_idx: i + 2 };
      // Restore internal page metadata from system column if present
      const srcSheet = r["[System] Source Sheet"] || r["[System] Page"];
      if (srcSheet) mapped.__dmp_page = srcSheet;
      return mapped;
    });
  } catch (e) {
    console.warn("Baseline load error:", e);
    return [];
  }
}

// ─── INIT ───
document.addEventListener("DOMContentLoaded", () => {
  loadEnv(); // Load secrets from .env
  navigate(STATE.lastPage); // Restore last active page
  restoreFolderOnLoad();
  refreshIcons();

  // Heartbeat to keep local server alive
  setInterval(() => {
    fetch("/heartbeat").catch(() => {});
  }, 3000);
});
