/* =========================
   CONFIGURAÇÃO PRINCIPAL
========================= */
const CONFIG = {
  spreadsheetId: "1fD4DMjR_5iRgTpKsS27McaV_tdbCTk50HKScpNyW7_U",
  sheets: [
    { key: "Consulte Escolas", label: "Consulte Escolas", enabled: true },
    { key: "Cadastros", label: "Cadastros", enabled: true },
    { key: "Escolas", label: "Escolas", enabled: true },
    { key: "Descrit Saeb e AprBr", label: "Descrit SAEB & Aprova Brasil", enabled: true },
    { key: "Descrit CAEd", label: "Descrit CAEd", enabled: true },
    { key: "Sugestões", label: "Sugestões", enabled: true },
    { key: "Resumo Escolas", label: "Resumo Escolas", enabled: true },
    { key: "R 01 Metas Saeb 25", label: "R01 Metas Saeb 2025", enabled: true }
  ],
  requestTimeoutMs: 30000,
  defaultRowsPerPage: 25,
  maxRowsPerPage: 1000,
  cacheTtlMs: 5 * 60 * 1000
};

/* =========================
   STATE & HELPERS
========================= */
const state = {
  route: "home",
  currentSheet: null,
  page: 1,
  rowsPerPage: Number(localStorage.getItem("rowsPerPage")) || CONFIG.defaultRowsPerPage,
  search: "",
  filterColumn: "",
  filterValue: "",
  sort: { col: null, dir: "asc" },
  lastUpdate: null,
  dataCache: new Map(), // sheetName -> { headers, rows, ts }
};

const qs = (sel) => document.querySelector(sel);
const qsa = (sel) => Array.from(document.querySelectorAll(sel));

/* =========================
   BOOTSTRAP
========================= */
document.addEventListener("DOMContentLoaded", () => {
  buildNav();
  buildHomeCards();

  // Navegação
  qsa(".nav-btn").forEach(b => b.addEventListener("click", () => go("home")));
  qs("#btn-back").addEventListener("click", () => go("home"));

  // Controles
  const rowsSelect = qs("#rows-select");
  if (rowsSelect) {
    rowsSelect.value = state.rowsPerPage;
    rowsSelect.addEventListener("change", (e) => {
      state.rowsPerPage = clamp(parseInt(e.target.value, 10) || CONFIG.defaultRowsPerPage, 1, CONFIG.maxRowsPerPage);
      localStorage.setItem("rowsPerPage", String(state.rowsPerPage));
      state.page = 1;
      renderCurrent();
    });
  }

  const searchInput = qs("#search-input");
  if (searchInput) {
    searchInput.addEventListener("input", debounce((e) => {
      state.search = (e.target.value || "").trim();
      state.page = 1;
      renderCurrent();
    }, 200));
  }

  const filterColumn = qs("#filter-column");
  if (filterColumn) {
    filterColumn.addEventListener("change", (e) => {
      state.filterColumn = e.target.value || "";
      state.page = 1;
      renderCurrent();
    });
  }

  const filterValue = qs("#filter-value");
  if (filterValue) {
    filterValue.addEventListener("input", debounce((e) => {
      state.filterValue = (e.target.value || "").trim();
      state.page = 1;
      renderCurrent();
    }, 200));
  }

  qs("#btn-prev").addEventListener("click", () => { state.page = Math.max(1, state.page - 1); renderCurrent(); });
  qs("#btn-next").addEventListener("click", () => { state.page = state.page + 1; renderCurrent(); });

  qs("#btn-export").addEventListener("click", exportCSV);

  // Roteamento Simples via hash
  window.addEventListener("hashchange", handleHashRoute);
  handleHashRoute();
});

/* =========================
   ROTEAMENTO
========================= */
function handleHashRoute(){
  const hash = decodeURIComponent(location.hash.replace(/^#/, "")) || "";
  if (!hash) return;
  const [r, sheet] = hash.split(":");
  if (r === "sheet" && sheet) openSheet(sheet);
}

function go(view){
  state.route = view;
  toggleView(view);
  if (view === "home") location.hash = "";
}

function toggleView(view){
  qsa(".view").forEach(v => v.classList.remove("active"));
  const target = qs(`#view-${view}`);
  if (target) target.classList.add("active");
}

/* =========================
   NAV & HOME
========================= */
function buildNav(){
  const nav = qs("#nav-sheets");
  if (!nav) return;
  nav.innerHTML = "";
  CONFIG.sheets.filter(s => s.enabled).forEach(s => {
    const btn = document.createElement("button");
    btn.className = "sheet-link";
    btn.textContent = s.label;
    btn.addEventListener("click", () => openSheet(s.key));
    nav.appendChild(btn);
  });
}

function buildHomeCards(){
  const wrap = qs("#sheet-cards");
  if (!wrap) return;
  wrap.innerHTML = "";
  CONFIG.sheets.filter(s => s.enabled).forEach(async (s) => {
    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <div class="card-head">
        <div class="card-title">${escapeHtml(s.label)}</div>
        <span class="badge-mini">${escapeHtml(shortName(s.key))}</span>
      </div>
      <div class="card-meta">
        <span id="meta-${cssId(s.key)}-rows"><i class="fas fa-spinner fa-spin"></i> Linhas…</span>
        <span id="meta-${cssId(s.key)}-cols"><i class="fas fa-spinner fa-spin"></i> Colunas…</span>
      </div>
      <div class="card-actions">
        <button class="btn" id="open-${cssId(s.key)}"><i class="fas fa-eye"></i> Ver dados</button>
      </div>
    `;
    wrap.appendChild(card);
    card.querySelector(`#open-${cssId(s.key)}`).addEventListener("click", () => openSheet(s.key));

    try{
      const { headers, rows } = await loadSheet(s.key);
      const r = qs(`#meta-${cssId(s.key)}-rows`);
      const c = qs(`#meta-${cssId(s.key)}-cols`);
      if (r) r.innerHTML = `<i class="fas fa-database"></i> ${rows.length} linhas`;
      if (c) c.innerHTML = `<i class="fas fa-table-columns"></i> ${headers.length} colunas`;
    }catch(e){
      const r = qs(`#meta-${cssId(s.key)}-rows`);
      const c = qs(`#meta-${cssId(s.key)}-cols`);
      if (r) r.textContent = "Erro ao ler";
      if (c) c.textContent = "";
    }
  });
}

/* =========================
   ABERTURA DE ABA
========================= */
async function openSheet(sheetName){
  go("data");
  location.hash = `sheet:${encodeURIComponent(sheetName)}`;

  const title = qs("#sheet-title span");
  if (title) title.textContent = sheetName;

  const head = qs("#table-head");
  const body = qs("#table-body");
  const spinner = qs("#table-spinner");
  if (head) head.innerHTML = "";
  if (body) body.innerHTML = "";
  if (spinner) spinner.style.display = "block";
  setStats({rows:0, cols:0, pages:1, visible:0});
  toast("Carregando dados…");

  try{
    const { headers, rows } = await loadSheet(sheetName);
    state.currentSheet = sheetName;
    state.page = 1;
    state.search = "";
    state.filterColumn = "";
    state.filterValue = "";
    state.sort = { col: null, dir: "asc" };
    populateControls(headers);
    renderCurrent();
  }catch(err){
    if (spinner) spinner.style.display = "none";
    toast(`Erro ao carregar: ${err.message}`, true);
  }
}

/* =========================
   BUSCA, FILTRO, ORDENAÇÃO
========================= */
function applyAllTransforms(rows, headers){
  let out = rows;

  if (state.filterColumn && state.filterValue) {
    const idx = headers.indexOf(state.filterColumn);
    if (idx >= 0) {
      const v = String(state.filterValue).toLowerCase();
      out = out.filter(r => String(r[idx] ?? "").toLowerCase() === v);
    }
  }

  if (state.search) {
    const term = state.search.toLowerCase();
    out = out.filter(r => r.some(cell => String(cell ?? "").toLowerCase().includes(term)));
  }

  if (state.sort.col !== null) {
    const idx = headers.indexOf(state.sort.col);
    const dir = state.sort.dir === "desc" ? -1 : 1;
    out = stableSort(out, (a,b) => compareSmart(a[idx], b[idx]) * dir);
  }

  return out;
}

/* =========================
   RENDERIZAÇÃO
========================= */
function renderCurrent(){
  const sp = qs("#table-spinner");
  if (sp) sp.style.display = "none";

  const cache = state.dataCache.get(state.currentSheet);
  if (!cache) return;

  const headers = cache.headers;
  const baseRows = cache.rows;

  const rows = applyAllTransforms(baseRows, headers);

  const total = rows.length;
  const pages = Math.max(1, Math.ceil(total / state.rowsPerPage));
  state.page = clamp(state.page, 1, pages);
  const start = (state.page - 1) * state.rowsPerPage;
  const end = Math.min(start + state.rowsPerPage, total);
  const pageRows = rows.slice(start, end);

  renderHeader(headers);
  renderBodyChunked(pageRows, headers);
  updatePagination(pages, total);
  setStats({rows: baseRows.length, cols: headers.length, pages, visible: total});
}

function renderHeader(headers){
  const thead = qs("#table-head");
  if (!thead) return;
  const tr = document.createElement("tr");
  tr.append(...headers.map(h => {
    const th = document.createElement("th");
    th.className = "th-sort";
    th.tabIndex = 0;
    th.innerHTML = `${escapeHtml(h)} <span class="sort-icon"><i class="fas fa-sort"></i></span>`;
    th.addEventListener("click", () => toggleSort(h));
    th.addEventListener("keypress", (e) => { if (e.key === "Enter") toggleSort(h); });
    return th;
  }));
  thead.innerHTML = "";
  thead.appendChild(tr);
}

function renderBodyChunked(rows, headers){
  const tbody = qs("#table-body");
  if (!tbody) return;
  tbody.innerHTML = "";

  const CHUNK = 200;
  let i = 0;
  function renderChunk(){
    let count = 0;
    while (i < rows.length && count < CHUNK) {
      const tr = document.createElement("tr");
      const r = rows[i];
      r.forEach((cell) => {
        const td = document.createElement("td");
        td.textContent = formatCell(cell);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
      i++; count++;
    }
    if (i < rows.length) {
      if ("requestIdleCallback" in window) {
        requestIdleCallback(renderChunk);
      } else {
        setTimeout(renderChunk, 0);
      }
    }
  }
  renderChunk();
}

function updatePagination(pages, total){
  const info = qs("#page-info");
  if (info) info.textContent = `Página ${state.page} de ${pages}`;
  const prev = qs("#btn-prev");
  const next = qs("#btn-next");
  if (prev) prev.disabled = state.page <= 1;
  if (next) next.disabled = state.page >= pages;
}

function setStats({rows, cols, pages, visible}){
  const sr = qs("#stat-rows");
  const sc = qs("#stat-cols");
  const sp = qs("#stat-pages");
  const sv = qs("#stat-visible");
  if (sr) sr.textContent = rows;
  if (sc) sc.textContent = cols;
  if (sp) sp.textContent = pages;
  if (sv) sv.textContent = visible;
  const stamp = new Date().toLocaleString();
  const upd = qs("#last-update");
  if (upd) upd.textContent = `Atualizado: ${stamp}`;
  state.lastUpdate = stamp;
}

function populateControls(headers){
  const colSel = qs("#filter-column");
  if (colSel) {
    colSel.innerHTML = `<option value="">(Todas)</option>` + headers.map(h => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`).join("");
  }
  const si = qs("#search-input"); if (si) si.value = "";
  const fv = qs("#filter-value"); if (fv) fv.value = "";
}

/* =========================
   EXPORTAÇÃO CSV
========================= */
function exportCSV(){
  const cache = state.dataCache.get(state.currentSheet);
  if (!cache) return;

  const headers = cache.headers;
  const rows = applyAllTransforms(cache.rows, headers);
  const lines = [
    headers.map(csvEscape).join(","),
    ...rows.map(r => r.map(csvEscape).join(","))
  ];
  const blob = new Blob([lines.join("\\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${slugify(state.currentSheet)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

/* =========================
   CARREGAMENTO (GViz JSON)
========================= */
async function loadSheet(sheetName){
  const cached = state.dataCache.get(sheetName);
  const now = Date.now();
  if (cached && (now - cached.ts < CONFIG.cacheTtlMs)) {
    return cached;
  }

  const url = `https://docs.google.com/spreadsheets/d/${CONFIG.spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), CONFIG.requestTimeoutMs);

  let text;
  try{
    const res = await fetch(url, { signal: controller.signal, credentials: "omit", cache: "no-store" });
    clearTimeout(timeout);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    text = await res.text();
  }catch(err){
    clearTimeout(timeout);
    throw new Error(`Falha de rede ou acesso (${err.message})`);
  }

  const json = parseGViz(text);
  if (!json || !json.table || !Array.isArray(json.table.cols) || !Array.isArray(json.table.rows)) {
    throw new Error("Formato inesperado da resposta");
  }

  const headers = json.table.cols.map(c => c && (c.label || c.id) || "").map(String);
  const rows = json.table.rows.map(r => normalizeRow(r, headers.length));

  const payload = { headers, rows, ts: now };
  state.dataCache.set(sheetName, payload);
  return payload;
}

function parseGViz(raw){
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  if (start === -1 || end === -1) return null;
  try { return JSON.parse(raw.slice(start, end + 1)); }
  catch { return null; }
}

/* =========================
   NORMALIZAÇÃO E TIPAGEM
========================= */
function normalizeRow(rowObj, headerLen){
  const out = new Array(headerLen).fill("");
  const cells = rowObj && Array.isArray(rowObj.c) ? rowObj.c : [];
  for (let i=0; i<headerLen; i++){
    const cell = cells[i];
    if (!cell) { out[i] = ""; continue; }
    let v = cell.v;
    if (cell.f && isLikelyDateFormat(cell.f, v)) {
      out[i] = formatAsDate(v);
      continue;
    }
    out[i] = v == null ? "" : v;
  }
  return out;
}

function isLikelyDateFormat(fmt, val) {
  return /\by\b|\bY\b|dd|mm|mmm|yyyy|\/+|-/.test(fmt) ||
         (typeof val === 'string' && /\d{2,4}[\/-]\d{1,2}[\/-]\d{1,2}/.test(val));
}

function formatAsDate(v) {
  try {
    const d = new Date(v);
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
    return String(v);
  } catch {
    return String(v);
  }
}

/* =========================
   SORT / COMPARE / HELPERS
========================= */
function toggleSort(col){
  if (state.sort.col === col) {
    state.sort.dir = state.sort.dir === "asc" ? "desc" : "asc";
  } else {
    state.sort.col = col;
    state.sort.dir = "asc";
  }
  state.page = 1;
  renderCurrent();
}

function stableSort(arr, cmp){
  const a = arr.map((v,i) => [v,i]);
  a.sort((x,y) => {
    const res = cmp(x[0], y[0]);
    return res !== 0 ? res : x[1] - y[1];
  });
  return a.map(x => x[0]);
}

function compareSmart(a, b){
  const na = toNumber(a), nb = toNumber(b);
  if (!isNaN(na) && !isNaN(nb)) return na - nb;
  const da = Date.parse(a), db = Date.parse(b);
  if (!isNaN(da) && !isNaN(db)) return da - db;
  return String(a ?? "").localeCompare(String(b ?? ""), "pt-BR", { sensitivity:"base" });
}

function toNumber(x){
  if (typeof x === "number") return x;
  if (typeof x !== "string") return NaN;
  const s = x.replace(/\\./g,"").replace(",",".").replace(/[^\\d.-]/g,"");
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}

function clamp(v, min, max){ return Math.max(min, Math.min(max, v)); }
function debounce(fn, ms=150){
  let t; return (...args)=>{ clearTimeout(t); t = setTimeout(()=>fn(...args), ms); };
}
function escapeHtml(s){
  return String(s).replace(/[&<>\"']/g, (m) => ({
    "&":"&amp;",
    "<":"&lt;",
    ">":"&gt;",
    "\"":"&quot;",
    "'":"&#39;"
  }[m]));
}
function slugify(s){
  return String(s)
    .normalize("NFD")
    .replace(/[\\u0300-\\u036f]/g,"")
    .replace(/[^a-z0-9]+/gi,"-")
    .replace(/(^-|-$)/g,"")
    .toLowerCase();
}
function csvEscape(v){
  const s = String(v ?? "");
  if (/[\",\\n]/.test(s)) return `\"${s.replace(/\"/g,'\"\"')}\"`;
  return s;
}
function cssId(s){ return slugify(s); }
function shortName(s){ return s.length > 22 ? s.slice(0,19) + "…" : s; }

function formatCell(v){
  if (v === true) return "SIM";
  if (v === false) return "NÃO";
  return String(v ?? "");
}

/* =========================
   UX
========================= */
let toastTimer = null;
function toast(msg, isError=false){
  const el = qs("#toast");
  if (!el) return;
  el.textContent = msg;
  el.style.background = isError ? "#991b1b" : "#0f172a";
  el.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=> el.classList.remove("show"), 2800);
}
