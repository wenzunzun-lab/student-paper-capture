const MOBILE_KEY = "answerCaptureMobile_v1";
const DESK_KEY = "wrongDeskState_v3";
const DB_NAME = "answerCapturePhotos";
const DB_STORE = "photos";

let data = loadData();
let pending = new Map();
let previews = new Map();
let activePage = 1;

function loadData() {
  try {
    const saved = JSON.parse(localStorage.getItem(MOBILE_KEY) || "null");
    if (saved?.roster && saved?.submissions) return saved;
  } catch (_) {}
  let roster = [];
  try {
    const desk = JSON.parse(localStorage.getItem(DESK_KEY) || "null");
    roster = (desk?.students || []).map(s => ({ id: String(s.id || ""), name: String(s.name || ""), className: String(s.className || "") })).filter(s => s.id && s.name);
  } catch (_) {}
  return { roster, submissions: [], examName: "", pageCount: 2 };
}

function persist() { localStorage.setItem(MOBILE_KEY, JSON.stringify(data)); }
function esc(v) { return String(v ?? "").replace(/[&<>"']/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" })[c]); }
function safeName(v) { return String(v || "").replace(/[\\/:*?"<>|]/g, "_").trim() || "未命名"; }
function makeId(prefix) { return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2,8)}`; }
function examName() { return document.getElementById("examName").value.trim(); }
function pageCount() { return Number(document.querySelector('input[name="pageCount"]:checked')?.value) || 2; }
function student() { return data.roster.find(s => s.id === document.getElementById("studentSelect").value) || null; }
function examRows() { const exam = examName(); return data.submissions.filter(row => row.examName === exam); }
function icons() { if (window.lucide) lucide.createIcons(); }
function toast(message) { const el = document.getElementById("toast"); el.textContent = message; el.classList.add("show"); clearTimeout(toast.timer); toast.timer = setTimeout(() => el.classList.remove("show"), 2500); }

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(DB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function dbAction(mode, action) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, mode);
    const request = action(tx.objectStore(DB_STORE));
    if (request) request.onsuccess = () => resolve(request.result || null);
    tx.oncomplete = () => { db.close(); if (!request) resolve(); };
    tx.onerror = () => { db.close(); reject(tx.error); };
  });
}
const putPhoto = (key, blob) => dbAction("readwrite", store => store.put(blob, key));
const getPhoto = key => dbAction("readonly", store => store.get(key));
const deletePhoto = key => dbAction("readwrite", store => store.delete(key));

function csvRows(text) {
  const rows = []; let row = [], value = "", quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"' && quoted && text[i + 1] === '"') { value += '"'; i++; }
    else if (c === '"') quoted = !quoted;
    else if (c === "," && !quoted) { row.push(value); value = ""; }
    else if (/\r|\n/.test(c) && !quoted) { if (c === "\r" && text[i + 1] === "\n") i++; row.push(value); if (row.some(x => String(x).trim())) rows.push(row); row = []; value = ""; }
    else value += c;
  }
  row.push(value); if (row.some(x => String(x).trim())) rows.push(row); return rows;
}

async function rosterRows(file) {
  const csv = /\.csv$/i.test(file.name);
  if (!window.XLSX) { if (csv) return csvRows(await file.text()); throw new Error("请联网后重试，或上传 CSV 花名册"); }
  const book = csv ? XLSX.read(await file.text(), { type:"string", raw:false }) : XLSX.read(await file.arrayBuffer(), { type:"array", raw:false });
  return XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], { header:1, raw:false, defval:"" });
}

function parseRoster(rows) {
  rows = rows.map(row => row.map(cell => String(cell ?? "").trim()));
  let hi = rows.findIndex(row => row.some(x => /^(学号|学生学号|考号|座号|编号)$/.test(x)) && row.some(x => /^(姓名|学生姓名|名字)$/.test(x)));
  if (hi < 0) hi = 0;
  const header = rows[hi].map(x => x.replace(/\s/g, ""));
  const col = patterns => header.findIndex(x => patterns.some(re => re.test(x)));
  let idCol = col([/^学号$/, /^学生学号$/, /^考号$/, /^座号$/, /^编号$/]);
  let nameCol = col([/^姓名$/, /^学生姓名$/, /^名字$/]);
  const classCol = col([/^班级$/, /^行政班$/, /^班$/]);
  if (idCol < 0 || nameCol < 0) { idCol = 0; nameCol = 1; }
  const found = new Map();
  rows.slice(hi + 1).forEach(row => {
    const id = String(row[idCol] || "").replace(/\.0$/, ""); const name = String(row[nameCol] || "");
    if (id && name && /\d/.test(id)) found.set(id, { id, name, className: classCol >= 0 ? String(row[classCol] || "") : "" });
  });
  return [...found.values()];
}

async function importRoster(file) {
  if (!file) return;
  try {
    const roster = parseRoster(await rosterRows(file));
    if (!roster.length) throw new Error("没有识别到学号和姓名");
    data.roster = roster; persist(); renderRoster(); toast(`已载入 ${roster.length} 名学生`);
  } catch (error) { toast(error.message || "花名册读取失败"); }
}

function renderRoster(preferred = "") {
  const select = document.getElementById("studentSelect");
  data.roster.sort((a,b) => a.id.localeCompare(b.id, "zh-CN", { numeric:true }));
  select.disabled = !data.roster.length;
  select.innerHTML = data.roster.length ? data.roster.map(s => `<option value="${esc(s.id)}">${esc(s.id)} · ${esc(s.name)}</option>`).join("") : '<option value="">请先上传花名册</option>';
  if (preferred && data.roster.some(s => s.id === preferred)) select.value = preferred;
  document.getElementById("nextStudentButton").disabled = !data.roster.length;
  render();
}

function renderPages() {
  const required = pageCount();
  document.querySelectorAll(".page-slot").forEach(slot => {
    const page = Number(slot.dataset.page); const image = slot.querySelector("img");
    slot.classList.toggle("is-disabled", page > required);
    slot.classList.toggle("has-image", pending.has(page));
    image.src = previews.get(page) || "";
  });
  document.getElementById("saveStudentButton").disabled = !examName() || !student() || [...Array(required)].some((_,i) => !pending.has(i + 1));
}

function renderQueue() {
  const rows = examRows().sort((a,b) => a.studentId.localeCompare(b.studentId, "zh-CN", { numeric:true }));
  const unshared = rows.filter(row => !row.exported);
  document.getElementById("completedCount").textContent = rows.length;
  document.getElementById("queueTitle").textContent = rows.length ? `${rows.length} 人 · ${rows.reduce((sum,row) => sum + row.pages.length, 0)} 页` : "尚未采集";
  document.getElementById("queueList").innerHTML = rows.length ? rows.map(row => `<div class="queue-item"><div><strong>${esc(row.studentId)} · ${esc(row.studentName)}</strong><small>${esc(row.examName)}${row.exported ? " · 已分享" : ""}</small></div><span class="page-badge">${row.pages.length}/${row.pageCount} 页</span><button class="queue-delete" data-delete="${esc(row.id)}" title="删除"><i data-lucide="trash-2"></i></button></div>`).join("") : '<div class="queue-empty">尚未采集答题卡</div>';
  document.getElementById("clearQueueButton").disabled = !rows.length;
  document.getElementById("shareButton").disabled = !unshared.length;
  document.querySelector("#shareButton span").textContent = `分享未导出（${unshared.length}人）`;
  document.querySelectorAll("[data-delete]").forEach(button => button.onclick = () => { if (confirm("确定删除这名学生的扫描件吗？")) removeSubmission(button.dataset.delete); });
}

function render() {
  const current = student();
  document.getElementById("rosterCount").textContent = data.roster.length;
  document.getElementById("studentName").textContent = current ? `${current.name}${current.className ? ` · ${current.className}` : ""}` : "请上传花名册";
  renderPages(); renderQueue(); icons();
}

function clearPending() {
  previews.forEach(url => URL.revokeObjectURL(url)); previews.clear(); pending.clear();
  document.getElementById("cameraInput").value = ""; renderPages();
}

function loadImage(file) {
  return new Promise((resolve,reject) => {
    const image = new Image(); const url = URL.createObjectURL(file);
    image.onload = () => { URL.revokeObjectURL(url); resolve(image); };
    image.onerror = () => { URL.revokeObjectURL(url); reject(new Error("照片无法读取")); };
    image.src = url;
  });
}

async function normalizePhoto(file) {
  const image = await loadImage(file); const scale = Math.min(1, 2600 / Math.max(image.naturalWidth, image.naturalHeight));
  const canvas = document.createElement("canvas"); canvas.width = Math.max(1, Math.round(image.naturalWidth * scale)); canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
  canvas.getContext("2d", { alpha:false }).drawImage(image, 0, 0, canvas.width, canvas.height);
  return new Promise((resolve,reject) => canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error("照片压缩失败")), "image/jpeg", .9));
}

async function acceptPhoto(file) {
  if (!file) return;
  try {
    const blob = await normalizePhoto(file);
    if (previews.has(activePage)) URL.revokeObjectURL(previews.get(activePage));
    pending.set(activePage, blob); previews.set(activePage, URL.createObjectURL(blob)); renderPages();
  } catch (error) { toast(error.message || "照片处理失败"); }
}

function openCamera(page) {
  if (!examName()) return toast("请先填写考试名称");
  if (!student()) return toast("请先选择学生学号");
  activePage = page; const input = document.getElementById("cameraInput"); input.value = ""; input.click();
}

async function removeSubmission(id, shouldRender = true) {
  const index = data.submissions.findIndex(row => row.id === id); if (index < 0) return;
  const [removed] = data.submissions.splice(index, 1);
  await Promise.all(removed.pages.map(page => deletePhoto(page.photoKey).catch(() => {})));
  persist(); if (shouldRender) render();
}

function nextStudent() {
  if (!data.roster.length) return;
  const select = document.getElementById("studentSelect"); const done = new Set(examRows().map(row => row.studentId));
  const index = data.roster.findIndex(s => s.id === select.value);
  const ordered = [...data.roster.slice(index + 1), ...data.roster.slice(0,index + 1)];
  select.value = (ordered.find(s => !done.has(s.id)) || ordered[0]).id; render();
}

async function saveStudent() {
  const current = student(); const exam = examName(); const required = pageCount();
  if (!current || !exam || [...Array(required)].some((_,i) => !pending.has(i + 1))) return;
  try {
    const old = data.submissions.find(row => row.examName === exam && row.studentId === current.id);
    if (old) await removeSubmission(old.id, false);
    const id = makeId("submission"); const pages = [];
    for (let page = 1; page <= required; page++) { const photoKey = `${id}_page_${page}`; await putPhoto(photoKey, pending.get(page)); pages.push({ page, photoKey }); }
    data.examName = exam; data.pageCount = required;
    data.submissions.push({ id, examName:exam, studentId:current.id, studentName:current.name, className:current.className || "", pageCount:required, pages, capturedAt:new Date().toISOString(), exported:false });
    persist(); clearPending(); nextStudent(); toast(`${current.id} ${current.name} 已保存`);
  } catch (_) { toast("保存失败，请检查手机剩余空间"); }
}

async function shareFiles(batch) {
  const files = []; const manifest = { format:"student-answer-capture", version:1, examName:examName(), exportedAt:new Date().toISOString(), submissions:[] };
  for (const row of batch) {
    const item = { studentId:row.studentId, studentName:row.studentName, className:row.className, pageCount:row.pageCount, pages:[] };
    for (const page of row.pages) {
      const blob = await getPhoto(page.photoKey); if (!blob) continue;
      const fileName = `${safeName(row.studentId)}_${safeName(row.studentName)}_第${page.page}页.jpg`;
      files.push(new File([blob], fileName, { type:"image/jpeg", lastModified:Date.now() })); item.pages.push({ page:page.page, fileName });
    }
    manifest.submissions.push(item);
  }
  files.push(new File([JSON.stringify(manifest,null,2)], `${safeName(examName())}_扫描清单.json`, { type:"application/json" }));
  return files;
}

function downloadFiles(files) {
  files.forEach((file,index) => setTimeout(() => { const link = document.createElement("a"); link.href = URL.createObjectURL(file); link.download = file.name; link.click(); setTimeout(() => URL.revokeObjectURL(link.href), 2000); }, index * 250));
}

async function sharePending() {
  const batch = examRows().filter(row => !row.exported); if (!batch.length) return;
  const button = document.getElementById("shareButton"); button.disabled = true;
  try {
    const files = await shareFiles(batch);
    if (navigator.share && (!navigator.canShare || navigator.canShare({ files }))) await navigator.share({ title:examName(), text:`${examName()} · ${batch.length} 名学生`, files });
    else { downloadFiles(files); toast("已开始下载扫描文件"); }
    batch.forEach(row => row.exported = true); persist(); render();
  } catch (error) { if (error?.name !== "AbortError") toast("分享失败，请重试"); }
  finally { button.disabled = false; }
}

async function clearQueue() {
  const rows = examRows(); if (!rows.length || !confirm(`确定删除“${examName()}”的 ${rows.length} 名学生扫描件吗？`)) return;
  for (const row of rows) await removeSubmission(row.id, false);
  persist(); render(); toast("本次扫描队列已清空");
}

function networkStatus() {
  const status = document.getElementById("networkStatus"); status.textContent = navigator.onLine ? "本机保存" : "离线可用"; status.classList.toggle("offline", !navigator.onLine);
}

document.getElementById("rosterButton").onclick = () => document.getElementById("rosterInput").click();
document.getElementById("rosterInput").onchange = event => importRoster(event.target.files[0]);
document.getElementById("studentSelect").onchange = () => { clearPending(); render(); };
document.getElementById("examName").oninput = event => { data.examName = event.target.value; persist(); render(); };
document.getElementById("nextStudentButton").onclick = () => { clearPending(); nextStudent(); };
document.querySelectorAll('input[name="pageCount"]').forEach(input => input.onchange = () => { data.pageCount = pageCount(); persist(); renderPages(); });
document.querySelectorAll(".page-slot").forEach(slot => slot.onclick = () => openCamera(Number(slot.dataset.page)));
document.getElementById("cameraInput").onchange = event => acceptPhoto(event.target.files[0]);
document.getElementById("saveStudentButton").onclick = saveStudent;
document.getElementById("shareButton").onclick = sharePending;
document.getElementById("clearQueueButton").onclick = clearQueue;
document.getElementById("resetExportButton").onclick = () => { examRows().forEach(row => row.exported = false); persist(); render(); toast("已重置分享状态"); };
window.addEventListener("online", networkStatus); window.addEventListener("offline", networkStatus);

document.getElementById("examName").value = data.examName || "";
document.querySelector(`input[name="pageCount"][value="${Number(data.pageCount) === 1 ? 1 : 2}"]`).checked = true;
renderRoster(); networkStatus();
if ("serviceWorker" in navigator) navigator.serviceWorker.register("mobile-sw.js").catch(() => {});
