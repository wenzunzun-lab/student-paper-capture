const MOBILE_KEY = "answerCaptureMobile_v1";
const DESK_KEY = "wrongDeskState_v3";
const DB_NAME = "answerCapturePhotos";
const DB_STORE = "photos";

let data = loadData();
let previews = new Map(); // photoKey -> objectURL
let activePage = 1;
const studentSearchCache = new Map();

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
function isComplete(row) { const n = row.pageCount || 2; return [...Array(n)].every((_, i) => row.pages.some(p => p.page === i + 1)); }
function icons() { if (window.lucide) lucide.createIcons(); }
function toast(message) { const el = document.getElementById("toast"); el.textContent = message; el.classList.add("show"); clearTimeout(toast.timer); toast.timer = setTimeout(() => el.classList.remove("show"), 2600); }

function normalizedSearch(value) { return String(value || "").toLowerCase().replace(/\s+/g, ""); }

function studentSearchKey(row) {
  const cached = studentSearchCache.get(row.id);
  if (cached?.name === row.name) return cached.key;
  let fullPinyin = ""; let initials = "";
  try {
    const syllables = window.pinyinPro?.pinyin(row.name, { toneType:"none", type:"array", nonZh:"consecutive" });
    if (Array.isArray(syllables)) {
      fullPinyin = syllables.join("");
      initials = syllables.map(part => String(part).charAt(0)).join("");
    }
  } catch (_) {}
  const key = [row.id, row.name, fullPinyin, initials].map(normalizedSearch).join("|");
  studentSearchCache.set(row.id, { name:row.name, key });
  return key;
}

function matchingStudents() {
  const query = normalizedSearch(document.getElementById("studentSearch").value);
  return query ? data.roster.filter(row => studentSearchKey(row).includes(query)) : data.roster;
}

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
    data.roster = roster; persist(); document.getElementById("studentSearch").value = ""; renderRoster(); toast(`已载入 ${roster.length} 名学生`);
  } catch (error) { toast(error.message || "花名册读取失败"); }
}

function renderRoster(preferred = "") {
  const search = document.getElementById("studentSearch");
  data.roster.sort((a,b) => a.id.localeCompare(b.id, "zh-CN", { numeric:true }));
  studentSearchCache.clear();
  search.disabled = !data.roster.length;
  renderStudentOptions(preferred);
  document.getElementById("nextStudentButton").disabled = !data.roster.length;
}

function renderStudentOptions(preferred = "") {
  const select = document.getElementById("studentSelect");
  const rows = matchingStudents();
  const hasRoster = data.roster.length > 0;
  select.disabled = !rows.length;
  select.innerHTML = rows.length ? rows.map(s => `<option value="${esc(s.id)}">${esc(s.id)} · ${esc(s.name)}</option>`).join("") : `<option value="">${hasRoster ? "没有匹配的学生" : "请先上传花名册"}</option>`;
  if (preferred && rows.some(s => s.id === preferred)) select.value = preferred;
  document.getElementById("studentMatchLabel").textContent = hasRoster ? `匹配学生（${rows.length}）` : "选择学生";
  render();
}

function renderPages() {
  const required = pageCount();
  const current = student();
  const row = current ? data.submissions.find(r => r.examName === examName() && r.studentId === current.id) : null;
  document.querySelectorAll(".page-slot").forEach(slot => {
    const page = Number(slot.dataset.page);
    const disabled = page > required;
    const p = row?.pages.find(x => x.page === page);
    const image = slot.querySelector("img");
    slot.classList.toggle("is-disabled", disabled);
    slot.classList.toggle("has-image", !!p);
    if (p) {
      if (previews.has(p.photoKey)) { image.src = previews.get(p.photoKey); }
      else {
        image.removeAttribute("src");
        getPhoto(p.photoKey).then(blob => { if (!blob) return; const url = URL.createObjectURL(blob); previews.set(p.photoKey, url); const img = slot.querySelector("img"); if (img && img.closest(".page-slot") === slot) img.src = url; }).catch(() => {});
      }
    } else { image.removeAttribute("src"); }
  });
  const hint = document.getElementById("autoSaveHint");
  if (hint) {
    const done = !!row && isComplete(row);
    hint.textContent = done ? `${row.studentName} 已拍齐 ${required} 页 · 点击下一名继续` : "拍照即自动保存，无需点确认";
    hint.classList.toggle("is-done", !!done);
  }
  const shareCurrent = document.getElementById("shareCurrentButton");
  if (shareCurrent) shareCurrent.disabled = !row || !row.pages.length;
}

function renderQueue() {
  const rows = examRows().sort((a,b) => a.studentId.localeCompare(b.studentId, "zh-CN", { numeric:true }));
  const doneCount = rows.filter(isComplete).length;
  const unshared = rows.filter(row => !row.exported);
  document.getElementById("completedCount").textContent = doneCount;
  document.getElementById("queueTitle").textContent = rows.length ? `${rows.length} 人 · ${rows.reduce((sum,row) => sum + row.pages.length, 0)} 页 · 完成 ${doneCount}` : "尚未采集";
  document.getElementById("queueList").innerHTML = rows.length ? rows.map(row => `<div class="queue-item${isComplete(row) ? " is-done" : ""}"><div><strong>${esc(row.studentId)} · ${esc(row.studentName)}</strong><small>${esc(row.examName)}${row.exported ? " · 已分享" : ""}${isComplete(row) ? " · 已完成" : " · 未拍完"}</small></div><span class="page-badge">${row.pages.length}/${row.pageCount} 页</span><button class="queue-share" data-share="${esc(row.id)}" title="单独分享这名学生（导入扫描王用）"><i data-lucide="share-2"></i></button><button class="queue-delete" data-delete="${esc(row.id)}" title="删除"><i data-lucide="trash-2"></i></button></div>`).join("") : '<div class="queue-empty">尚未采集答题卡</div>';
  document.getElementById("clearQueueButton").disabled = !rows.length;
  document.getElementById("shareButton").disabled = !unshared.length;
  document.querySelector("#shareButton span").textContent = `分享未导出（${unshared.length}人）`;
  document.querySelectorAll("[data-share]").forEach(button => button.onclick = () => shareStudent(button.dataset.share));
  document.querySelectorAll("[data-delete]").forEach(button => button.onclick = () => { if (confirm("确定删除这名学生的扫描件吗？")) removeSubmission(button.dataset.delete); });
}

function render() {
  const current = student();
  document.getElementById("rosterCount").textContent = data.roster.length;
  const searching = normalizedSearch(document.getElementById("studentSearch").value);
  document.getElementById("studentName").textContent = current ? `${current.name}${current.className ? ` · ${current.className}` : ""}` : data.roster.length && searching ? "未找到学生" : "请上传花名册";
  renderPages(); renderQueue(); icons();
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

// 拍照即存：照片一进来立刻保存到当前学生名下，无需任何确认
async function acceptPhoto(file) {
  if (!file) return;
  const current = student();
  if (!current) return toast("请先选择学生");
  const exam = examName();
  if (!exam) return toast("请先填写考试名称");
  try {
    const blob = await normalizePhoto(file);
    const required = pageCount();
    let row = data.submissions.find(r => r.examName === exam && r.studentId === current.id);
    if (!row) {
      row = { id: makeId("submission"), examName: exam, studentId: current.id, studentName: current.name, className: current.className || "", pageCount: required, pages: [], capturedAt: new Date().toISOString(), exported: false };
      data.submissions.push(row);
    }
    row.pageCount = required;
    row.studentName = current.name;
    row.className = current.className || "";
    const photoKey = `${row.id}_page_${activePage}`;
    await putPhoto(photoKey, blob);
    const existing = row.pages.find(p => p.page === activePage);
    const wasExisting = !!existing;
    if (existing) { if (previews.has(existing.photoKey)) URL.revokeObjectURL(previews.get(existing.photoKey)); previews.delete(existing.photoKey); existing.photoKey = photoKey; }
    else row.pages.push({ page: activePage, photoKey });
    data.examName = exam;
    persist();
    render();
    if (isComplete(row)) {
      if (!wasExisting) {
        toast(`${current.name} 已保存（${required} 页）· 自动进入下一名`);
        nextStudent();
      } else {
        toast(`${current.name} 第 ${activePage} 页已更新`);
      }
    } else {
      const have = new Set(row.pages.map(p => p.page)).size;
      toast(`第 ${activePage} 页已保存 · ${current.name} 还需 ${required - have} 页`);
    }
  } catch (_) { toast("保存失败，请检查手机剩余空间"); }
}

function openCamera(page) {
  if (!examName()) return toast("请先填写考试名称");
  if (!student()) return toast("请先选择学生学号");
  activePage = page; const input = document.getElementById("cameraInput"); input.value = ""; input.click();
}

async function removeSubmission(id, shouldRender = true) {
  const index = data.submissions.findIndex(row => row.id === id); if (index < 0) return;
  const [removed] = data.submissions.splice(index, 1);
  await Promise.all(removed.pages.map(page => { if (previews.has(page.photoKey)) URL.revokeObjectURL(previews.get(page.photoKey)); previews.delete(page.photoKey); return deletePhoto(page.photoKey).catch(() => {}); }));
  persist(); if (shouldRender) render();
}

function nextStudent() {
  if (!data.roster.length) return;
  const select = document.getElementById("studentSelect"); const done = new Set(examRows().filter(isComplete).map(row => row.studentId));
  const index = data.roster.findIndex(s => s.id === select.value);
  const ordered = [...data.roster.slice(index + 1), ...data.roster.slice(0,index + 1)];
  const next = ordered.find(s => !done.has(s.id)) || ordered[0];
  document.getElementById("studentSearch").value = ""; renderStudentOptions(next.id);
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

// 单独分享一名学生（导入扫描王等场景用）：不改变“未导出”标记
async function shareStudent(id) {
  const row = data.submissions.find(r => r.id === id); if (!row || !row.pages.length) return;
  try {
    const files = await shareFiles([row]);
    if (navigator.share && (!navigator.canShare || navigator.canShare({ files }))) await navigator.share({ title:`${row.studentName} 答题卡`, text:`${row.studentId} ${row.studentName} · ${row.pages.length} 页`, files });
    else { downloadFiles(files); toast("已开始下载该学生扫描文件"); }
  } catch (error) { if (error?.name !== "AbortError") toast("分享失败，请重试"); }
}

async function sharePending() {
  const batch = examRows().filter(row => !row.exported); if (!batch.length) return;
  const button = document.getElementById("shareButton"); button.disabled = true;
  try {
    const files = await shareFiles(batch);
    if (navigator.share && (!navigator.canShare || navigator.canShare({ files }))) await navigator.share({ title:examName(), text:`${examName()} · ${batch.length} 名学生`, files });
    else { downloadFiles(files); toast("已开始下载扫描文件"); }
    batch.forEach(row => row.exported = true); persist();
  } catch (error) { if (error?.name !== "AbortError") toast("分享失败，请重试"); }
  finally { render(); }
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
document.getElementById("studentSelect").onchange = () => { render(); };
document.getElementById("studentSearch").oninput = () => {
  const before = document.getElementById("studentSelect").value;
  renderStudentOptions(before);
};
document.getElementById("examName").oninput = event => { data.examName = event.target.value; persist(); render(); };
document.getElementById("nextStudentButton").onclick = () => { nextStudent(); };
document.querySelectorAll('input[name="pageCount"]').forEach(input => input.onchange = () => { data.pageCount = pageCount(); persist(); renderPages(); });
document.querySelectorAll(".page-slot").forEach(slot => slot.onclick = () => openCamera(Number(slot.dataset.page)));
document.getElementById("cameraInput").onchange = event => acceptPhoto(event.target.files[0]);
document.getElementById("shareButton").onclick = sharePending;
document.getElementById("shareCurrentButton").onclick = () => {
  const current = student(); if (!current) return toast("请先选择学生");
  const row = data.submissions.find(r => r.examName === examName() && r.studentId === current.id);
  if (!row) return toast("这名学生还没有照片");
  shareStudent(row.id);
};
document.getElementById("clearQueueButton").onclick = clearQueue;
document.getElementById("resetExportButton").onclick = () => { examRows().forEach(row => row.exported = false); persist(); render(); toast("已重置分享状态"); };
window.addEventListener("online", networkStatus); window.addEventListener("offline", networkStatus);

document.getElementById("examName").value = data.examName || "";
document.querySelector(`input[name="pageCount"][value="${Number(data.pageCount) === 1 ? 1 : 2}"]`).checked = true;
renderRoster(); networkStatus();
if ("serviceWorker" in navigator) navigator.serviceWorker.register("mobile-sw.js").catch(() => {});
