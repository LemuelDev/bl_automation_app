const MODULES = [
  { id: "shipping_list", 
    title: "1. Shipping List", 
    auto: true, 
    ready: true, 
    accepts: "files",
    desc: "Select IV Files.", 
    hint: "Select IV documents/files from the server." 
  },
  { id: "copy_iv_base_si", 
    title: "2. Copy IV Base Create SI", 
    auto: true, 
    ready: true, 
    accepts: "folder",
    desc: "Select WIP BL/Separated IV in the downloads folder", 
    hint: "Reads and process the SEPARATED IV FOLDER." 
  },
  { id: "input_data_b1", 
    title: "3. Input Data (Base B1)", 
    auto: true, 
    ready: true, 
    ccepts: "files",
    desc: "Select all .xlsm files in the WIP BL/FOR RPA in the downloads folder.", 
    hint: "Read and process .xlsm files from FOR RPA FOLDER." 
  },
  { id: "so_templates", 
    title: "4. BL Creation (Single and Multi)", 
    auto: true, 
    ready: true, 
    accepts: "folder",
    desc: "Make sure to update the templates for the single and multi files before running this stage.",
    hint: "Select and process the FOR RPA FOLDER with subfolder SINGLE and MULTI." 
  },
  { id: "checker", 
    title: "5. Checker Macro",
    auto: true, 
    ready: true, 
    accepts: "folder",
    desc: "Collects shipper, BL no, units, weight and m3 from the finished BL files into the Checker result template.", 
    hint: "Select and process the WIP BL/BL/ALL in the downloads folder." 
  },
  { id: "remove_zeros", 
    title: "6. Removing Zero Macro",
    auto: true, 
    ready: true, 
    accepts: "folder",
    desc: "Replaces the values in A18 and A19 of each selected BL file with a blank space, so no zeros show.", 
    hint: "Selects and process all files in WIP BL/BL/ALL" 
  },
  { id: "move_copy_attach", 
    title: "7. Move or Copy Attach Macro",
    auto: true, 
    ready: true, 
    accepts: "folder",
    desc: "Copies the ATCH (2) sheet of each multi-unit file into its matching BL file in BL\\ALL.", 
    hint: "Select FOR RPA\\MULTI." 
  },
  { id: "all_bl", 
    title: "8. All BL (FM/BL Creation) Macro",
    auto: true, 
    ready: true, 
    accepts: "folder",
    desc: "Combines every sheet of every BL file in the selected folder into one new workbook, the BL output.", 
    hint: "Select and run the folder BL\\ALL" 
  },
];
const AUTO_STEPS = MODULES.filter((m) => m.auto);

const $ = (sel, root = document) => root.querySelector(sel);
const api = () => (window.pywebview && window.pywebview.api) || null;
const fileName = (p) => (p || "").split(/[\\/]/).pop();

function el(tag, cls, text) {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
}

let busy = false;
function setBusy(v) {
  busy = v;
  document.body.classList.toggle("is-busy", v);
}

async function pick(kind) {
  const a = api();
  if (!a) throw new Error("The app is still starting, try again in a moment.");
  if (kind === "folder") {
    const p = await a.select_folder();
    return p ? [p] : [];
  }
  return (await a.select_files()) || [];
}

function addPaths(list, paths, type) {
  paths.forEach((path) => {
    if (path && !list.some((i) => i.path === path)) list.push({ type, path });
  });
}

function renderChips(box, list, emptyText) {
  box.replaceChildren();
  if (!list.length) return box.append(el("p", "empty", emptyText));

  if (list.length > 3) {
    const chip = el("div", "chip");
    chip.append(el("span", "path", `${list.length} items selected (${fileName(list[0].path)} and ${list.length - 1} more)`));
    const x = el("button", "x", "\u00d7");
    x.title = "Clear all";
    x.onclick = () => { list.length = 0; renderChips(box, list, emptyText); box.dispatchEvent(new Event("change")); };
    chip.append(x);
    return box.append(chip);
  }
  list.forEach((item, i) => {
    const chip = el("div", "chip");
    chip.append(el("span", "kind", item.type === "folder" ? "Folder" : "File"));
    const p = el("span", "path", fileName(item.path));
    p.title = item.path;
    chip.append(p);
    const x = el("button", "x", "\u00d7");
    x.title = "Remove";
    x.onclick = () => { list.splice(i, 1); renderChips(box, list, emptyText); box.dispatchEvent(new Event("change")); };
    chip.append(x);
    box.append(chip);
  });
}

const views = {};
const staticViews = { create: "#view-create", stages: "#view-stages" };

function showView(id) {
  document.querySelectorAll(".view").forEach((v) => v.classList.remove("is-active"));
  (staticViews[id] ? $(staticViews[id]) : views[id]).classList.add("is-active");
  const navId = staticViews[id] ? id : "stages";
  document.querySelectorAll(".nav button").forEach((b) => b.classList.toggle("is-active", b.dataset.target === navId));
}

function buildNav() {
  const items = [
    { id: "create", name: "Create BL" },
    { id: "stages", name: "Create BL per Macro" },
  ];
  const nav = $("#nav");
  if (!nav) return;
  nav.replaceChildren();
  items.forEach((it) => {
    const b = el("button");
    b.dataset.target = it.id;
    b.textContent = it.name;
    b.onclick = () => showView(it.id);
    nav.append(b);
  });
}

function applyTheme(t, save) {
  document.documentElement.dataset.theme = t;
  if (save) { try { localStorage.setItem("bl-theme", t); } catch (e) {} }
}

function initTheme() {
  const toggleBtn = $("#theme-toggle");
  if (toggleBtn) {
    toggleBtn.onclick = () => {
      const current = document.documentElement.dataset.theme || "dark";
      const next = current === "dark" ? "light" : "dark";
      applyTheme(next, true);
    };
  }
  applyTheme(document.documentElement.dataset.theme || "dark", false);
}

function buildCards() {
  const grid = $("#stage-grid");
  if (!grid) return;
  grid.replaceChildren();
  MODULES.forEach((m, i) => {
    const card = el("button", "glass card");
    card.disabled = !m.ready;
    card.append(el("h4", "", m.title), el("p", "", m.desc));
    if (m.hint) card.append(el("span", "hint", m.hint));
    card.onclick = () => showView(m.id);
    grid.append(card);
  });
}

const stageInputs = {};

function logTo(id, text, kind) {
  const box = views[id] ? $('[data-f="log"]', views[id]) : null;
  if (!box) return;
  box.querySelector(".empty")?.remove();
  const line = el("p", kind || "");
  line.append(el("span", "ts", `[${new Date().toLocaleTimeString()}] `), document.createTextNode(text));
  box.append(line);
  box.scrollTop = box.scrollHeight;
}

function buildModuleViews() {
  const tpl = $("#module-template");
  if (!tpl) return;
  MODULES.forEach((m, i) => {
    if (!m.ready) return;
    const view = tpl.content.firstElementChild.cloneNode(true);
    views[m.id] = view;
    stageInputs[m.id] = [];

    $(".eyebrow", view).textContent = `Stage ${i + 1}`;
    $(".title", view).textContent = m.title;
    $(".desc", view).textContent = m.desc;

    const chips = $('[data-f="inputs"]', view);
    const redraw = () => renderChips(chips, stageInputs[m.id], "No files or folders added yet.");
    chips.addEventListener("change", () => {});
    redraw();
    $('[data-f="log"]', view).append(el("p", "empty", "Nothing run yet. Add your inputs and run the stage."));

    // Handle button visibility based on what the stage accepts
    const btnFiles = $('[data-action="files"]', view);
    const btnFolder = $('[data-action="folder"]', view);

    if (m.accepts === "files") {
      if (btnFolder) btnFolder.style.display = "none"; // Hide folder button
      if (btnFiles) {
        btnFiles.onclick = async () => {
          try { addPaths(stageInputs[m.id], await pick("files"), "file"); redraw(); }
          catch (e) { logTo(m.id, String(e.message || e), "err"); }
        };
      }
    } else if (m.accepts === "folder") {
      if (btnFiles) btnFiles.style.display = "none"; // Hide files button
      if (btnFolder) {
        btnFolder.onclick = async () => {
          try { addPaths(stageInputs[m.id], await pick("folder"), "folder"); redraw(); }
          catch (e) { logTo(m.id, String(e.message || e), "err"); }
        };
      }
    }
    
    const backBtn = $('[data-act="back"]', view);
    if (backBtn) backBtn.onclick = () => showView("stages");

    const runBtn = $('[data-act="run"]', view);
    if (runBtn) runBtn.onclick = () => runModule(m.id);
    
    const clearBtn = $('[data-act="clear-log"]', view);
    if (clearBtn) {
      clearBtn.onclick = () => {
        const logBox = $('[data-f="log"]', view);
        logBox.replaceChildren(el("p", "empty", "Log cleared."));
      };
    }

    $("#deck").append(view);
  });
}

async function runModule(id) {
  if (busy) return;
  const inputs = stageInputs[id];
  if (!inputs.length) return logTo(id, "Add at least one file or folder first.", "err");
  const a = api();
  if (!a) return logTo(id, "The app is still starting, try again in a moment.", "err");

  setBusy(true);
  logTo(id, "Running stage process...");
  try {
    const paths = inputs.map((i) => i.path);
    const one = paths.length === 1 ? paths[0] : paths;
    const r = await a.run_macro(id, { file_paths: one, folder_or_files: one, sheet_name: "BASE" });
    const ok = !!r && (r.ok === true || r.status === "success");
    logTo(id, (r && r.message) || "Finished.", ok ? "ok" : "err");
    ((r && r.errors) || []).forEach((e) => logTo(id, e, "err"));
    ((r && r.warnings) || []).forEach((w) => logTo(id, w, "warn"));
  } catch (e) {
    logTo(id, "Unexpected error: " + e, "err");
  } finally {
    setBusy(false);
  }
}

function confirmClear(rows, mode) {
  return new Promise((resolve) => {
    const modal = $("#modal");
    const verb = mode === "delete" ? "deleted" : "moved to the archive folder";
    $("#modal-text").textContent = `These old files will be ${verb} before the stages that write to them start.`;
    const ul = $("#modal-list");
    ul.replaceChildren();
    rows.forEach(([label, n]) => {
      const li = el("li");
      li.append(el("span", "", label), el("span", "muted", `${n} file${n === 1 ? "" : "s"}`));
      ul.append(li);
    });
    const done = (v) => { modal.hidden = true; resolve(v); };
    $("#modal-ok").onclick = () => done(true);
    $("#modal-cancel").onclick = () => done(false);
    modal.hidden = false;
    $("#modal-ok").focus();
  });
}

const createInputs = [];
let pollTimer = null;
let lastLogKey = "";

function buildCreatePanel() {
  const box = $("#create-inputs");
  const btn = $("#create-btn");
  if (!box || !btn) return;

  const redraw = () => {
    renderChips(box, createInputs, "No files added yet.");
    btn.disabled = createInputs.length === 0;
    $("#create-error").hidden = true;
  };
  box.addEventListener("change", redraw);
  redraw();

  document.querySelectorAll("[data-add]").forEach((b) => {
    b.onclick = async () => {
      try { addPaths(createInputs, await pick(b.dataset.add), b.dataset.add === "folder" ? "folder" : "file"); redraw(); }
      catch (e) { showCreateError(String(e.message || e)); }
    };
  });
  btn.onclick = startPipeline;
  $("#reset-btn").onclick = resetCreate;
  
  const clearCreateLog = $("#clear-create-log");
  if (clearCreateLog) {
    clearCreateLog.onclick = () => {
      $("#run-log").replaceChildren(el("p", "empty", "Log cleared."));
    };
  }
}

function showCreateError(text) {
  const n = $("#create-error");
  if (n) {
    n.textContent = text;
    n.hidden = false;
  }
}

async function startPipeline() {
  if (busy || !createInputs.length) return;
  const a = api();
  if (!a) return showCreateError("The app is still starting, try again in a moment.");

  try {
    const prev = await a.pipeline_preview();
    const rows = [
      ["Separated IV folder", prev.stage1_clears],
      ["For RPA folders", prev.stage2_clears],
      ["BL folders", prev.stage4_clears],
    ].filter((r) => r[1] > 0);
    if (rows.length && !(await confirmClear(rows, prev.mode))) return;

    const paths = createInputs.map((i) => i.path);
    const res = await a.start_pipeline({ file_paths: paths, sheet_name: "BASE" });
    if (!res.ok) {
      const detail = (res.problems || [res.message]).slice(0, 6).join("\n");
      return showCreateError(`Cannot start.\n${detail}`);
    }
    enterRun();
    poll();
  } catch (e) {
    showCreateError("Unexpected error: " + e);
  }
}

function enterRun() {
  setBusy(true);
  lastLogKey = "";
  $("#create-setup").hidden = true;
  $("#create-run").hidden = false;
  $("#summary").hidden = true;
  $("#reset-btn").hidden = true;
  $("#run-log").replaceChildren();
  const steps = $("#steps");
  steps.replaceChildren();
  AUTO_STEPS.forEach((m) => {
    const li = el("li", "step pending");
    li.dataset.id = m.id;
    const top = el("div", "top");
    top.append(el("span", "dot"), el("span", "", m.title));
    li.append(top, el("div", "msg", "Waiting"));
    steps.append(li);
  });
}

function resetCreate() {
  createInputs.length = 0;
  $("#create-run").hidden = true;
  $("#create-setup").hidden = false;
  $("#create-inputs").dispatchEvent(new Event("change"));
}

async function poll() {
  clearTimeout(pollTimer);
  let s;
  try { s = await api().pipeline_status(); }
  catch (e) { pollTimer = setTimeout(poll, 1500); return; }

  renderRun(s);
  if (s.done) {
    setBusy(false);
    renderSummary(s);
    $("#reset-btn").hidden = false;
  } else {
    pollTimer = setTimeout(poll, 800);
  }
}

function renderRun(s) {
  document.querySelectorAll("#steps .step").forEach((li, i) => {
    const st = s.stages[i];
    if (!st) return;
    li.className = `step ${st.status}`;
    const text = { pending: "Waiting", running: "Running\u2026", skipped: "Skipped" }[st.status] || st.message;
    $(".msg", li).textContent = text || "";
  });
  $("#run-msg").textContent = s.done ? "" : (s.message || "");
  
  // Look for this section inside your renderRun(s) function in app.js:
    const key = s.log.length + "|" + (s.log[s.log.length - 1] || "");
    if (key !== lastLogKey) {
      lastLogKey = key;
      const log = $("#run-log");
      if (log) {
        log.replaceChildren(...s.log.map((t) => {
          const p = el("p", "", t);
          const lower = t.toLowerCase();
          
          // >>> UPDATE YOUR COLOR CONDITIONS HERE <<<
          if (
            lower.includes("✓") || 
            lower.includes("completed") || 
            lower.includes("complete") ||    // Catches Stage 2
            lower.includes("finished") || 
            lower.includes("saved") || 
            lower.includes("updated") || 
            lower.includes("created") ||     // Catches Stage 4
            lower.includes("copied") ||      // Catches Stage 4
            lower.includes("added to")
          ) {
            p.style.color = "#48bb78"; // Soft Green
          } else if (
            lower.includes("✗") || 
            lower.includes("failed") || 
            lower.includes("error") || 
            lower.includes("stopped")
          ) {
            p.style.color = "#f56565"; // Soft Red
          }
          return p;
        }));
        log.scrollTop = log.scrollHeight;
      }
    }
}

let addList = (box, label, items, cls) => {
  if (!items || !items.length) return;
  const d = el("details", cls);
  d.open = cls === "list-err";
  d.append(el("summary", "", `${label} (${items.length})`));
  const ul = el("ul");
  items.forEach((t) => ul.append(el("li", "", t)));
  d.append(ul);
  box.append(d);
};

function renderSummary(s) {
  const box = $("#summary");
  if (!box) return;
  box.hidden = false;
  box.className = "summary " + (s.ok ? "is-ok" : "is-fail");
  box.replaceChildren(el("h4", "", s.ok ? "Run complete" : "Run stopped"), el("p", "muted", s.message));
  addList(box, "Errors", s.errors, "list-err");
  addList(box, "Warnings", s.warnings, "list-warn");
  if (s.archive_dir) box.append(el("p", "muted small", "Previous files were archived to " + s.archive_dir));
}

buildNav();
buildCards();
buildModuleViews();
buildCreatePanel();
initTheme();
showView("create");

window.addEventListener("pywebviewready", async () => {
  try {
    const s = await api().pipeline_status();
    if (s && s.running) { enterRun(); poll(); }
  } catch (e) {}
});


// pipeline helper para sa create BL in one click shesh
async function startPipeline() {
  if (busy || !createInputs.length) return;
  const a = api();
  if (!a) return showCreateError("The app is still starting, try again in a moment.");

  try {
    // 1. Fetch preview counts from bridge.py -> PipelineRunner
    const prev = await a.pipeline_preview();
    const rows = [
      ["Separated IV folder", prev.stage1_clears || 0],
      ["For RPA folders", prev.stage2_clears || 0],
      ["BL folders", prev.stage4_clears || 0],
    ].filter((r) => r[1] > 0);

    // If there are old files, trigger the confirmation modal
    if (rows.length && !(await confirmClear(rows, prev.mode))) return;

    // 2. Start the pipeline with the selected input files
    const paths = createInputs.map((i) => i.path);
    const res = await a.start_pipeline({ file_paths: paths, sheet_name: "BASE" });
    
    if (!res.ok) {
      const detail = (res.problems || [res.message]).slice(0, 6).join("\n");
      return showCreateError(`Cannot start.\n${detail}`);
    }

    // 3. Switch UI to the active run progress screen and start polling
    enterRun();
    poll();
  } catch (e) {
    showCreateError("Unexpected error: " + e);
  }
}

async function poll() {
  clearTimeout(pollTimer);
  let s;
  try { 
    // Calls bridge.py -> _pipeline.status()
    s = await api().pipeline_status(); 
  } catch (e) { 
    pollTimer = setTimeout(poll, 1500); 
    return; 
  }

  renderRun(s);

  if (s.done) {
    setBusy(false);
    renderSummary(s);
    const resetBtn = $("#reset-btn");
    if (resetBtn) resetBtn.hidden = false;
  } else {
    pollTimer = setTimeout(poll, 800); // Poll every 0.8 seconds while running
  }
}

window.addEventListener("pywebviewready", async () => {
  try {
    const s = await api().pipeline_status();
    if (s && s.running) { 
      enterRun(); 
      poll(); 
    }
  } catch (e) {}
});