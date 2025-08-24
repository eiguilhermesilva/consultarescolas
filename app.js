/* SheetsDB — app.js
 * Robust loader for public Google Sheets on GitHub Pages
 * - Lists tabs (worksheets) via public v3 feed (no API key)
 * - Loads data via GViz Query with server-side pagination (limit/offset)
 * - Sorting, search (page-level), CSV export, caching and good errors
 */

const UI = {
  idInput: document.getElementById("sheet-id"),
  loadBtn: document.getElementById("load-btn"),
  resetBtn: document.getElementById("reset-btn"),
  rowsPerPageSel: document.getElementById("rows-per-page"),
  headerRowsSel: document.getElementById("header-rows"),
  cacheToggle: document.getElementById("cache-toggle"),

  error: document.getElementById("error"),
  tabs: document.getElementById("tabs"),
  tableContainer: document.getElementById("table-container"),
  tableLoader: document.getElementById("table-loader"),
  searchInput: document.getElementById("search-input"),
  exportCSV: document.getElementById("export-csv"),

  statTotal: document.getElementById("stat-total"),
  statCols: document.getElementById("stat-cols"),
  statSheet: document.getElementById("stat-sheet"),
  statUpdated: document.getElementById("stat-updated"),

  prevBtn: document.getElementById("prev"),
  nextBtn: document.getElementById("next"),
  pageInfo: document.getElementById("page-info"),
};

const state = {
  sheetId: "",
  worksheets: /** @type {Array<{title:string,gid:number}>} */ ([]),
  active: /** @type {{title:string,gid:number}|null} */ (null),

  rowsPerPage: 25,
  headerRows: 1,
  useCache: true,

  // pagination
  page: 1,
  totalRowsEstimate: 0,

  // data of current page
  headers: /** @type {string[]} */ ([]),
  pageRows: /** @type {Array<Record<string, any>>} */ ([]),

  // sorting
  sortKey: null,
  sortDir: "asc",

  // search
  searchTerm: "",
};

const CACHE_PREFIX = "sheetsdb:v2:";

/* -------------------------- Utilities --------------------------- */
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

function showError(msg) {
  UI.error.textContent = msg;
  UI.error.classList.remove("hidden");
}
function clearError() {
  UI.error.classList.add("hidden");
  UI.error.textContent = "";
}

function setLoading(loading) {
  UI.tableLoader.classList.toggle("hidden", !loading);
  UI.tableLoader.setAttribute("aria-hidden", loading ? "false" : "true");
}

function fmtDateTime(d = new Date()) {
  return new Intl.DateTimeFormat("pt-BR", {
    dateStyle: "short",
    timeStyle: "medium",
  }).format(d);
}

function debounce(fn, ms=250){
  let t; 
  return (...args)=>{ clearTimeout(t); t=setTimeout(()=>fn(...args), ms); };
}

function toCSV(headers, rows){
  const escape = (v) => {
    if (v == null) return "";
    const s = String(v);
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g,'""')}"`;
    return s;
  };
  const lines = [];
  lines.push(headers.map(escape).join(","));
  for (const r of rows){
    lines.push(headers.map(h => escape(r[h])).join(","));
  }
  return lines.join("\n");
}

function downloadFile(filename, content, mime="text/plain;charset=utf-8"){
  const blob = new Blob([content], {type: mime});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ----------------------- Sheets Fetchers ------------------------ */
/** Try to list worksheets (tabs) using the legacy public feed.
 * Works on publicly shared sheets.
 * Returns [{title, gid}]
 */
async function listWorksheets(sheetId){
  // Strategy A: v3 public feed (no key)
  const feedUrl = `https://spreadsheets.google.com/feeds/worksheets/${sheetId}/public/full?alt=json`;
  try{
    const res = await fetch(feedUrl, { mode: "cors" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    const entries = json?.feed?.entry ?? [];
    if (!entries.length) throw new Error("Sem abas públicas visíveis.");
    const tabs = [];
    for (const e of entries){
      const title = e?.title?.$t ?? "Sem título";
      // Try to get gid from links (alternate/edit link usually contains gid=)
      const links = e?.link ?? [];
      let gid = null;
      for (const l of links){
        const href = l?.href || "";
        const m = href.match(/[?&#]gid=(\d+)/);
        if (m) { gid = Number(m[1]); break; }
      }
      // Fallback: try "gs$gid"
      if (gid == null && e["gs$gid"]) gid = Number(e["gs$gid"]);
      // As último recurso, mapear 'od6' -> 0 (antigo padrão)
      if (gid == null && /\/od6$/.test(e?.id?.$t || "")) gid = 0;

      if (gid == null) continue; // não conseguimos usar sem gid
      tabs.push({ title, gid });
    }
    if (tabs.length) return tabs;
    throw new Error("Não foi possível extrair GIDs das abas.");
  } catch(err){
    // Strategy B: Single-sheet fallback (we will still load gid 0)
    console.warn("listWorksheets fallback:", err);
    return [{ title: "Planilha 1", gid: 0 }];
  }
}

/** Build GViz query URL */
function gvizUrl({ sheetId, gid, headerRows, query }){
  const params = new URLSearchParams({
    tqx: "out:json",
    gid: String(gid),
    headers: String(headerRows),
  });
  if (query) params.set("tq", query);
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?${params.toString()}`;
}

/** Parse GViz JSON text into object */
function parseGviz(text){
  // GViz returns: google.visualization.Query.setResponse({...});
  const json = JSON.parse(text.replace(/^[^{]+/, "").replace(/;?\s*$/, ""));
  if (json.status !== "ok") throw new Error(json?.errors?.[0]?.detailed_message || "Erro GViz");
  return json;
}

/** Convert GViz table to rows+headers */
function tableToRows(table){
  const cols = table.cols || [];
  const headers = cols.map((c, i) => c.label || c.id || `Col${i+1}`);
  const rows = [];
  for (const r of (table.rows || [])){
    const obj = {};
    (r.c || []).forEach((cell, idx) => {
      let v = cell?.v ?? "";
      // normalize dates
      if (cell?.f && /(^\d{1,2}\/\d{1,2}\/\d{2,4}$)/.test(cell.f)) v = cell.f;
      obj[headers[idx]] = v;
    });
    rows.push(obj);
  }
  return { headers, rows };
}

/** Count total rows using first column (A) */
async function countRows(sheetId, gid, headerRows){
  const q = "select count(A)";
  const url = gvizUrl({ sheetId, gid, headerRows, query: q });
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Falha ao contar registros (gid ${gid}).`);
  const json = parseGviz(await res.text());
  const cell = json?.table?.rows?.[0]?.c?.[0];
  const count = Number(cell?.v ?? 0);
  // Subtrair cabeçalho se for headerRows > 0? GViz já considera headerRows para labels.
  // A contagem retorna linhas de dados (desconsidera cabeçalhos). Ok.
  return count || 0;
}

/** Fetch a page (limit/offset) */
async function fetchPage(sheetId, gid, headerRows, limit, offset){
  const q = `select * limit ${limit} offset ${offset}`;
  const url = gvizUrl({ sheetId, gid, headerRows, query: q });
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Falha ao carregar página (gid ${gid}).`);
  const json = parseGviz(await res.text());
  return tableToRows(json.table);
}

/* --------------------------- Renderers -------------------------- */
function renderTabs(tabs){
  UI.tabs.innerHTML = "";
  for (const t of tabs){
    const btn = document.createElement("button");
    btn.className = "tab";
    btn.type = "button";
    btn.textContent = t.title;
    btn.dataset.gid = String(t.gid);
    btn.addEventListener("click", () => setActiveTab(t));
    UI.tabs.appendChild(btn);
  }
  updateActiveTabUI();
}

function updateActiveTabUI(){
  const gid = state.active?.gid;
  [...UI.tabs.children].forEach(el => {
    el.classList.toggle("active", Number(el.dataset.gid) === gid);
  });
  UI.statSheet.textContent = state.active?.title ?? "-";
}

function renderTable({ headers, rows }){
  const hasRows = rows.length > 0;
  const h = headers;

  // search (page-level)
  let filtered = rows;
  if (state.searchTerm){
    const term = state.searchTerm.toLowerCase();
    filtered = filtered.filter(r => h.some(k => String(r[k] ?? "").toLowerCase().includes(term)));
  }

  // sort
  if (state.sortKey){
    const k = state.sortKey, dir = state.sortDir;
    filtered = [...filtered].sort((a,b)=>{
      const va = a[k] ?? "", vb = b[k] ?? "";
      if (va === vb) return 0;
      return (va > vb ? 1 : -1) * (dir === "asc" ? 1 : -1);
    });
  }

  // build table
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");

  h.forEach(col => {
    const th = document.createElement("th");
    th.textContent = col;
    th.classList.add("sortable");
    const ind = document.createElement("span");
    ind.className = "sort-indicator";
    ind.textContent = (state.sortKey === col) ? (state.sortDir === "asc" ? "▲" : "▼") : "↕";
    th.appendChild(ind);
    th.addEventListener("click", ()=>{
      if (state.sortKey === col) {
        state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
      } else {
        state.sortKey = col;
        state.sortDir = "asc";
      }
      renderTable({ headers: h, rows: state.pageRows });
    });
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  const tbody = document.createElement("tbody");
  if (!hasRows){
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = h.length;
    td.textContent = "Nenhum dado nesta página.";
    tr.appendChild(td);
    tbody.appendChild(tr);
  } else {
    for (const r of filtered){
      const tr = document.createElement("tr");
      for (const col of h){
        const td = document.createElement("td");
        const v = r[col];
        td.textContent = v == null || v === "" ? "-" : String(v);
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
  }

  table.appendChild(thead);
  table.appendChild(tbody);
  UI.tableContainer.innerHTML = "";
  UI.tableContainer.appendChild(table);

  // stats
  UI.statCols.textContent = String(h.length);
  UI.statTotal.textContent = String(state.totalRowsEstimate);
  UI.statUpdated.textContent = fmtDateTime();

  // enable tools
  UI.searchInput.disabled = false;
  UI.exportCSV.disabled = filtered.length === 0;

  // pagination info
  const totalPages = Math.max(1, Math.ceil(state.totalRowsEstimate / state.rowsPerPage));
  UI.pageInfo.textContent = `Página ${state.page} de ${totalPages}`;
  UI.prevBtn.disabled = state.page <= 1;
  UI.nextBtn.disabled = state.page >= totalPages;
}

/* --------------------------- Cache ------------------------------ */
function cacheKeyPage(sheetId, gid, headerRows, page, limit){
  return `${CACHE_PREFIX}${sheetId}:${gid}:h${headerRows}:p${page}:l${limit}`;
}
function cacheSet(key, value){
  try{ localStorage.setItem(key, JSON.stringify({ t: Date.now(), v: value })); }catch{}
}
function cacheGet(key, maxAgeMs=5*60*1000){
  try{
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const { t, v } = JSON.parse(raw);
    if (Date.now() - t > maxAgeMs) return null;
    return v;
  }catch{ return null; }
}

/* ------------------------ Controller Flow ----------------------- */
async function setActiveTab(tab){
  if (!tab) return;
  state.active = tab;
  state.page = 1;
  state.sortKey = null; state.sortDir = "asc";
  state.searchTerm = "";
  UI.searchInput.value = "";
  updateActiveTabUI();
  await loadActivePage();
}

async function loadActivePage(){
  if (!state.active) return;

  const { gid, title } = state.active;
  const limit = state.rowsPerPage;
  const offset = (state.page - 1) * limit;

  try{
    clearError();
    setLoading(true);

    // total rows (estimate) — cache per gid
    const totalKey = `${CACHE_PREFIX}${state.sheetId}:${gid}:count:h${state.headerRows}`;
    let total = state.useCache ? cacheGet(totalKey, 10*60*1000) : null;
    if (total == null){
      total = await countRows(state.sheetId, gid, state.headerRows);
      cacheSet(totalKey, total);
    }
    state.totalRowsEstimate = Number(total) || 0;

    // page data — cache per page
    const pageKey = cacheKeyPage(state.sheetId, gid, state.headerRows, state.page, limit);
    let page = state.useCache ? cacheGet(pageKey) : null;
    if (page == null){
      page = await fetchPage(state.sheetId, gid, state.headerRows, limit, offset);
      cacheSet(pageKey, page);
    }

    state.headers = page.headers;
    state.pageRows = page.rows;

    renderTable({ headers: state.headers, rows: state.pageRows });
  } catch(err){
    console.error(err);
    showError(err?.message || "Erro ao carregar dados da aba.");
  } finally {
    setLoading(false);
  }
}

async function connect(sheetId){
  try{
    clearError();
    setLoading(true);

    // List tabs
    const tabs = await listWorksheets(sheetId);
    state.worksheets = tabs;
    state.sheetId = sheetId;

    renderTabs(tabs);
    if (tabs.length){
      await setActiveTab(tabs[0]);
    } else {
      showError("Nenhuma aba pública encontrada.");
    }
  } catch(err){
    console.error(err);
    showError("Não foi possível acessar a planilha. Verifique se está pública e se o ID está correto.");
  } finally {
    setLoading(false);
  }
}

/* --------------------------- Events ----------------------------- */
UI.loadBtn.addEventListener("click", async () => {
  const id = UI.idInput.value.trim();
  if (!id){
    showError("Informe um ID de planilha.");
    return;
  }
  await connect(id);
});

UI.resetBtn.addEventListener("click", ()=>{
  state.sheetId = "";
  state.worksheets = [];
  state.active = null;
  state.page = 1;
  state.headers = [];
  state.pageRows = [];
  state.totalRowsEstimate = 0;
  state.sortKey = null; state.sortDir = "asc";
  state.searchTerm = "";

  UI.idInput.value = "";
  UI.tabs.innerHTML = "";
  UI.tableContainer.innerHTML = "";
  UI.pageInfo.textContent = "Página 1 de 1";
  UI.prevBtn.disabled = true;
  UI.nextBtn.disabled = true;
  UI.searchInput.value = "";
  UI.searchInput.disabled = true;
  UI.exportCSV.disabled = true;

  UI.statTotal.textContent = "0";
  UI.statCols.textContent = "0";
  UI.statSheet.textContent = "-";
  UI.statUpdated.textContent = "-";

  clearError();
});

UI.rowsPerPageSel.addEventListener("change", async (e)=>{
  state.rowsPerPage = Number(e.target.value);
  state.page = 1;
  await loadActivePage();
});

UI.headerRowsSel.addEventListener("change", async (e)=>{
  state.headerRows = Number(e.target.value);
  // invalidate count cache for this gid
  if (state.active){
    const { gid } = state.active;
    const totalKey = `${CACHE_PREFIX}${state.sheetId}:${gid}:count:h${state.headerRows}`;
    localStorage.removeItem(totalKey);
  }
  state.page = 1;
  await loadActivePage();
});

UI.cacheToggle.addEventListener("change", (e)=>{
  state.useCache = e.target.checked;
});

UI.prevBtn.addEventListener("click", async ()=>{
  if (state.page > 1){
    state.page--;
    await loadActivePage();
  }
});
UI.nextBtn.addEventListener("click", async ()=>{
  const totalPages = Math.max(1, Math.ceil(state.totalRowsEstimate / state.rowsPerPage));
  if (state.page < totalPages){
    state.page++;
    await loadActivePage();
  }
});

UI.searchInput.addEventListener("input", debounce(()=>{
  state.searchTerm = UI.searchInput.value.trim();
  renderTable({ headers: state.headers, rows: state.pageRows });
}, 200));

UI.exportCSV.addEventListener("click", ()=>{
  // Exporta a página atual (após filtros/sort aplicados pela UI)
  // Para exportar tudo, seria necessário iterar offsets (potencialmente pesado)
  const headers = state.headers;
  const term = state.searchTerm.toLowerCase();
  let rows = state.pageRows;
  if (term){
    rows = rows.filter(r => headers.some(h => String(r[h] ?? "").toLowerCase().includes(term)));
  }
  // Aplica sort igual à UI
  if (state.sortKey){
    const k = state.sortKey, dir = state.sortDir;
    rows = [...rows].sort((a,b)=>{
      const va = a[k] ?? "", vb = b[k] ?? "";
      if (va === vb) return 0;
      return (va > vb ? 1 : -1) * (dir === "asc" ? 1 : -1);
    });
  }
  const csv = toCSV(headers, rows);
  downloadFile(`${(state.active?.title||"dados").replace(/\s+/g,'_')}_p${state.page}.csv`, csv, "text/csv;charset=utf-8");
});

/* UX: clique no exemplo para colar */
document.addEventListener("click", (e)=>{
  if (e.target instanceof HTMLElement && e.target.classList.contains("selectable")){
    navigator.clipboard?.writeText(e.target.textContent.trim()).catch(()=>{});
    UI.idInput.value = e.target.textContent.trim();
  }
});

/* Prefill with demo for convenience */
window.addEventListener("DOMContentLoaded", ()=>{
  UI.idInput.value = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
});
