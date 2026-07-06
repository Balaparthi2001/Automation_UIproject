/************ CONFIG ************/
const ADMIN_PASS = "admin@123"; // ← change this for your setup
const ADMIN_SESSION_KEY = "admin_ok_once";
const VISIBLE_LABELS = ["Label_1","Label_2","Label_3","Label_4"]; // fixed list

/************ Tabs ************/
const tabBtns = document.querySelectorAll(".tab-btn");
const panels = document.querySelectorAll(".tab-panel");

function activateTab(tabId) {
  tabBtns.forEach(b => {
    const isActive = b.dataset.tab === tabId;
    b.classList.toggle("active", isActive);
    b.setAttribute("aria-selected", isActive ? "true" : "false");
  });
  panels.forEach(p => p.classList.toggle("active", p.id === tabId));
  if (location.hash !== `#${tabId}`) history.replaceState(null, "", `#${tabId}`);
}

async function requireAdmin(promptTitle = 'Admin Access') {
  if (sessionStorage.getItem(ADMIN_SESSION_KEY) === "1") return true;
  const result = await Swal.fire({
    title: promptTitle,
    input: 'password',
    inputLabel: 'Enter admin password',
    inputPlaceholder: 'Password',
    inputAttributes: { autocapitalize: 'off', autocorrect: 'off' },
    confirmButtonText: 'Unlock',
    showCancelButton: true,
    allowOutsideClick: false,
    icon: 'question'
  });
  if (!result.isConfirmed) return false;
  const pass = result.value || "";
  if (pass === ADMIN_PASS) {
    sessionStorage.setItem(ADMIN_SESSION_KEY, "1");
    await Swal.fire({ icon: 'success', title: 'Access granted', timer: 800, showConfirmButton: false });
    return true;
  }
  await Swal.fire({ icon: 'error', title: 'Access denied', text: 'Incorrect password' });
  return false;
}

async function guardAndActivate(tabId) {
  if (tabId === "run") return activateTab(tabId);
  const ok = await requireAdmin('Admin Access');
  if (ok) activateTab(tabId);
}

tabBtns.forEach(btn => {
  btn.addEventListener("click", (e) => {
    e.preventDefault();
    const targetTab = btn.dataset.tab;
    guardAndActivate(targetTab);
  });
});

(function initTabsOnLoad() {
  const params = new URLSearchParams(location.search);
  const qsTab = params.get("tab");
  const hashTab = location.hash ? location.hash.slice(1) : null;
  const allowed = new Set(["upload", "run", "update"]);
  let initial = "run";
  if (qsTab && allowed.has(qsTab)) initial = qsTab;
  else if (hashTab && allowed.has(hashTab)) initial = hashTab;
  activateTab(initial);
})();

/************ FileDrop helpers ************/
function fmtSize(bytes) {
  if (bytes === 0) return "0 B";
  if (!bytes && bytes !== 0) return "";
  const units = ["B","KB","MB","GB"]; let i = 0, n = bytes;
  while (n >= 1024 && i < units.length - 1) { n/=1024; i++; }
  return `${n.toFixed(n < 10 && i > 0 ? 1 : 0)} ${units[i]}`;
}
function updateFileHint(zoneId, file) {
  const zone = document.getElementById(zoneId);
  if (!zone) return;
  const big = zone.querySelector(".big");
  const small = zone.querySelector(".small");
  if (file) {
    const sizeTxt = file.size ? ` • ${fmtSize(file.size)}` : "";
    if (big) big.textContent = file.name;
    if (small) small.textContent = `Selected${sizeTxt} — click to change`;
    zone.classList.add("has-file");
  } else {
    if (zone.id === "drop-upload") {
      if (big) big.textContent = "Drag & Drop";
      if (small) small.textContent = "or click to select .py";
    } else {
      if (big) big.textContent = "Drop file";
      if (small) small.textContent = "or click to select";
    }
    zone.classList.remove("has-file");
  }
}
function wireDrop(zoneId, inputId){
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  if (!zone || !input) return;

  zone.addEventListener("click", () => input.click());
  ["dragenter","dragover"].forEach(ev => zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add("drag"); }));
  ["dragleave","drop"].forEach(ev => zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove("drag"); }));
  zone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) { input.files = e.dataTransfer.files; updateFileHint(zoneId, file); }
  });
  input.addEventListener("change", () => {
    const file = input.files && input.files[0];
    updateFileHint(zoneId, file || null);
  });
  updateFileHint(zoneId, null);
}
wireDrop("drop-upload", "u_file");
wireDrop("drop-update", "update-file");

/************ Elements ************/
const labelHeader  = document.getElementById("label-header");
const btnBack      = document.getElementById("btn-back");
const labelNameEl  = document.getElementById("label-name");
const labelCountEl = document.getElementById("label-count");
const labelListDiv = document.getElementById("label-list");
const fileListDiv  = document.getElementById("file-list");

const terminal     = document.getElementById("terminal");
const runMsg       = document.getElementById("run-msg");
const uploadMsg    = document.getElementById("upload-msg");
const updateMsg    = document.getElementById("update-msg");
const updateTarget = document.getElementById("update-target");
const search       = document.getElementById("search");
const clearSearch  = document.getElementById("clear-search");
const btnStop      = document.getElementById("btn-stop");

const editorEl     = document.getElementById("editor");
const btnReloadBase= document.getElementById("btn-reload-base");
const btnRunBase   = document.getElementById("btn-run-base");
const btnRunEdited = document.getElementById("btn-run-edited");

let eventSource   = null;
let fileTree      = [];
let filesByLabel  = {};  // { Label_1: ["Label_1/x.py", ...], ... }
let currentLabel  = null; // null = label list view
let currentPath   = null; // relative path under automations/
let baseCodeCache = "";  // latest fetched base
let editorDirty   = false;

/************ API ************/
async function api(url, opts){
  const res = await fetch(url, opts);
  if(!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  const ct = res.headers.get("content-type") || "";
  return ct.includes("application/json") ? res.json() : res.text();
}

/************ Files ************/
async function listFiles(){
  const tree = await api("/files/tree");
  fileTree = tree;
  buildFilesByLabel();
  renderLabels();
  renderTargetsFlat();
}

function buildFilesByLabel(){
  filesByLabel = {};
  VISIBLE_LABELS.forEach(lbl => filesByLabel[lbl] = []);
  fileTree.forEach(group => {
    group.files.forEach(rel => {
      const top = (rel.split('/')[0] || "").trim();
      if (VISIBLE_LABELS.includes(top)) filesByLabel[top].push(rel);
    });
  });
  // sort each list
  Object.keys(filesByLabel).forEach(k => filesByLabel[k].sort((a,b)=>a.localeCompare(b)));
}

function renderLabels(){
  // Switch to label list view
  currentLabel = null;
  toggleLeftView(false); // show labels
  labelListDiv.innerHTML = "";
  search.disabled = true;
  search.placeholder = "Select a label to search…";
  (clearSearch) && (clearSearch.disabled = true);

  VISIBLE_LABELS.forEach(lbl => {
    const count = filesByLabel[lbl]?.length || 0;
    const card = document.createElement("div");
    card.className = "label-card";

    const name = document.createElement("div");
    name.className = "label-name";
    name.textContent = lbl;

    const badge = document.createElement("div");
    badge.className = "count-badge";
    badge.textContent = count;

    card.appendChild(name);
    card.appendChild(badge);
    card.addEventListener('click', () => renderFilesForLabel(lbl));
    labelListDiv.appendChild(card);
  });
}

function renderFilesForLabel(lbl){
  currentLabel = lbl;
  toggleLeftView(true); // show files
  const all = filesByLabel[lbl] || [];
  const q = (search?.value || "").toLowerCase().trim();
  const filtered = q ? all.filter(f => f.toLowerCase().includes(q)) : all.slice();

  labelNameEl.textContent = lbl;
  labelNameEl.style.fontFamily ="Georgia, 'Times New Roman', Times, serif"
  labelCountEl.textContent = "Total Script :" +all.length;

  search.disabled = false;
  search.placeholder = `Search in ${lbl}…`;
  (clearSearch) && (clearSearch.disabled = false);

  fileListDiv.innerHTML = "";
  filtered.forEach(relPath => {
    const card = document.createElement("div");
    card.className = "card";

    const left = document.createElement("div");
    left.className = "name";
    left.title = relPath;
    left.textContent = relPath.split('/').slice(1).join('/') || relPath; // show path inside label
    left.addEventListener('click', () => openInEditor(relPath));

    const right = document.createElement("div");
    right.className = "actions";

    const openBtn = document.createElement("button");
    openBtn.className = "chip";
    openBtn.textContent = "📝 Edit";
    openBtn.onclick = () => openInEditor(relPath);

    const runBtn = document.createElement("button");
    runBtn.className = "chip";
    runBtn.textContent = "▶ Run";
    runBtn.onclick = () => runScript(relPath, 'auto');

    const delBtn = document.createElement("button");
    delBtn.className = "chip danger";
    delBtn.textContent = "🗑 Delete";
    delBtn.onclick = () => deleteFile(relPath);

    // right.appendChild(openBtn);
    // right.appendChild(runBtn);
    // right.appendChild(delBtn);
    card.appendChild(left);
    card.appendChild(right);
    fileListDiv.appendChild(card);
  });
}

function toggleLeftView(inLabel){
  // header/back
  labelHeader.classList.toggle('hidden', !inLabel);
  // lists
  labelListDiv.classList.toggle('hidden', inLabel);
  fileListDiv.classList.toggle('hidden', !inLabel);
}

btnBack?.addEventListener('click', () => {
  search.value = "";
  renderLabels();
});

// Search behaves only within current label
search?.addEventListener("input", () => {
  if (currentLabel) renderFilesForLabel(currentLabel);
});
clearSearch?.addEventListener("click", () => {
  if (search) search.value = "";
  if (currentLabel) renderFilesForLabel(currentLabel);
});

function renderTargetsFlat(){
  updateTarget.innerHTML = "";
  const flat = [];
  Object.values(filesByLabel).forEach(arr => arr.forEach(f => flat.push(f)));
  flat.sort((a,b) => a.localeCompare(b));
  flat.forEach(f => {
    const opt = document.createElement("option");
    opt.value = f; opt.textContent = f; updateTarget.appendChild(opt);
  });
}

/************ Editor ************/
async function openInEditor(relPath){
  try{
    const data = await api(`/files/content?path=${encodeURIComponent(relPath)}`);
    if (!data.ok) throw new Error(data.error || 'read failed');
    currentPath = relPath;
    baseCodeCache = data.content || '';
    editorEl.value = baseCodeCache;
    editorDirty = false;
    updateRunButtons();
    runMsg.textContent = `Selected: ${relPath}`;
  }catch(e){
    Swal.fire({ icon:'error', title:'Open failed', text: e.message });
  }
}

function updateRunButtons(){
  const dirty = (editorEl.value ?? '') !== (baseCodeCache ?? '');
  editorDirty = dirty;

  if (dirty) {
    // Code edited → Run Edited only
    btnRunEdited.disabled = false;
    btnRunEdited.classList.add('primary');

    btnRunBase.disabled = true;
    btnRunBase.classList.remove('primary');
  } else {
    // Clean / base code → Run Base only
    btnRunBase.disabled = false;
    btnRunBase.classList.add('primary');

    btnRunEdited.disabled = true;
    btnRunEdited.classList.remove('primary');
  }
}
editorEl?.addEventListener('input', updateRunButtons);

btnReloadBase?.addEventListener('click', () => {
  if (baseCodeCache !== undefined) editorEl.value = baseCodeCache;
  updateRunButtons();
});

btnRunBase?.addEventListener('click', () => {
  if (!currentPath) return Swal.fire({ icon:'info', title:'Pick a file first' });
  runScript(currentPath, 'auto');
});

btnRunEdited?.addEventListener('click', async () => {
  if (!currentPath) return Swal.fire({ icon:'info', title:'Pick a file first' });
  const code = editorEl.value ?? '';
  if (!code.trim()) return Swal.fire({ icon:'warning', title:'Editor is empty' });
  try{
    const data = await api('/run/preview/start', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: currentPath, code })
    });
    if (!data.ok) throw new Error(data.error || 'start failed');
    runScript(data.temp, 'tmp', `Edited: ${currentPath}`);
  }catch(e){
    Swal.fire({ icon:'error', title:'Preview run failed', text: e.message });
  }
});

/************ Upload (only .py) ************/
document.getElementById("btn-upload")?.addEventListener("click", async () =>{
  try{
    uploadMsg.textContent = "";
    const f = document.getElementById("u_file").files[0];
    const name = document.getElementById("u_name").value.trim();
    const cat  = document.getElementById("u_cat").value.trim();

    if(!f || !name) {
      Swal.fire({ icon:'warning', title:'Select file & name' });
      return;
    }
    if(!name.endsWith(".py") && !f.name.endsWith(".py")) {
      Swal.fire({ icon:'error', title:'Only Python (.py) files are allowed' });
      return;
    }

    const fd = new FormData();
    fd.append("file", f);
    fd.append("name", name);
    if (cat) fd.append("category", cat);

    const data = await api("/upload", { method: "POST", body: fd });
    if (data.ok) {
      Swal.fire({ icon:'success', title:`Uploaded: ${data.name}`, timer: 1200, showConfirmButton:false });
    } else {
      Swal.fire({ icon:'error', title:'Upload failed', text: data.error || 'Unknown error' });
    }
    await listFiles();
  }catch(e){
    Swal.fire({ icon:'error', title:'Upload error', text: e.message });
  }
});

/************ Update (only .py & filename must match target) ************/
document.getElementById("btn-update")?.addEventListener("click", async () =>{
  try{
    updateMsg.textContent = "";
    const f = document.getElementById("update-file").files[0];
    const target = document.getElementById("update-target").value;

    if(!f || !target){
      Swal.fire({ icon:'warning', title:'Pick target & new file' });
      return;
    }
    if(!f.name.endsWith(".py")){
      Swal.fire({ icon:'error', title:'Only Python (.py) files are allowed' });
      return;
    }
    if (f.name !== target.split('/').pop()){
      Swal.fire({
        icon:'error',
        title:'Filename mismatch',
        html:`Selected file <b>${f.name}</b> must match target name <b>${target.split('/').pop()}</b>.`}
      );
      return;
    }

    const fd = new FormData();
    fd.append("file", f);
    fd.append("target", target);
    const data = await api("/update", { method:"POST", body: fd });
    if (data.ok) {
      Swal.fire({ icon:'success', title:`Updated: ${target}`, timer: 1200, showConfirmButton:false });
    } else {
      Swal.fire({ icon:'error', title:'Update failed', text: data.error || 'Unknown error' });
    }
    await listFiles();
  }catch(e){
    Swal.fire({ icon:'error', title:'Update error', text: e.message });
  }
});

/************ Delete (confirm → admin password → delete) ************/
async function deleteFile(relPath){
  const go = await Swal.fire({
    icon:'warning',
    title:`Delete ${relPath}?`,
    text:'This action cannot be undone.',
    showCancelButton:true,
    confirmButtonText:'Delete',
    cancelButtonText:'Cancel'
  });
  if (!go.isConfirmed) return;

  const ok = await requireAdmin('Admin password required to delete');
  if (!ok) return;

  try{
    const url = `/files?path=${encodeURIComponent(relPath)}`;
    const data = await api(url, { method:"DELETE" });
    if (data.ok) {
      Swal.fire({ icon:'success', title:`Deleted: ${relPath}`, timer: 1000, showConfirmButton:false });
      await listFiles();
      if (currentLabel) renderFilesForLabel(currentLabel);
    } else {
      Swal.fire({ icon:'error', title:'Delete failed', text: data.error || 'Unknown error' });
    }
  }catch(e){
    Swal.fire({ icon:'error', title:'Delete failed', text: e.message });
  }
}

/************ Run with SSE + Stop confirm ************/
function runScript(relPath, scope='auto', label=null){
  if (eventSource) { try { eventSource.close(); } catch {} eventSource = null; }
  terminal.textContent = "";

  const display = label ?? relPath;
  runMsg.textContent = `Running: ${display}`;
  if (btnStop) btnStop.disabled = false;

  // Simple banner (no 'py', so user is not confused)
  terminal.append(`> ${display}\n\n`);

  const url = `/run/stream?file=${encodeURIComponent(relPath)}&scope=${encodeURIComponent(scope)}`;
  eventSource = new EventSource(url);

  eventSource.onmessage = (ev) => {
    const line = ev.data;
    if (line.startsWith("__EXIT_CODE__:")) {
      const code = parseInt(line.split(":")[1], 10);
      runMsg.textContent = (code === 0) ? "✅ Completed successfully" : `❌ Exited (code ${code})`;
      if (btnStop) btnStop.disabled = true;
      eventSource.close(); eventSource = null;
      return;
    }
    terminal.append(line + "\n");     // append is typically faster than insertAdjacentText
    terminal.scrollTop = terminal.scrollHeight;
  };

  eventSource.onerror = () => {
    runMsg.textContent = "Stream error / stopped.";
    if (btnStop) btnStop.disabled = true;
    if (eventSource) { try { eventSource.close(); } catch {} eventSource = null; }
  };
}

btnStop?.addEventListener("click", async ()=>{
  if (!eventSource) {
    Swal.fire({ icon:'info', title:'No active task' });
    return;
  }
  const res = await Swal.fire({
    icon: 'warning',
    title: 'End the task?',
    text: 'Are you sure you want to stop the running script?',
    showCancelButton: true,
    confirmButtonText: 'Yes, stop',
    cancelButtonText: 'No'
  });
  if (!res.isConfirmed) return;

  try{
    if(eventSource){ eventSource.close(); eventSource = null; }
    const data = await api('/run/stop', { method: 'POST' });
    if (btnStop) btnStop.disabled = true;
    runMsg.textContent = "⛔ Stopped";
    Swal.fire({ icon: data.ok ? 'success' : 'error', title: data.ok ? 'Stopped' : 'Stop failed', text: data.message || '' });
  }catch(e){
    Swal.fire({ icon:'error', title:'Stop failed', text: e.message });
  }
});

/************ Init ************/
document.getElementById("refresh-targets")?.addEventListener("click", listFiles);
listFiles();