/* app.js – Apresentação fixa de uma planilha específica (somente leitura)
 * Recursos:
 * - Lista abas públicas automaticamente
 * - Paginação servidor (limit/offset via GViz)
 * - Ordenação por coluna, busca (na página), exportar CSV (página e tudo)
 * - Cache leve em localStorage
 * - Configuração por aba: renomear/ocultar/ordenar colunas, sort padrão, somatórios, formatação
 */

const SHEET_ID = "1fD4DMjR_5iRgTpKsS27McaV_tdbCTk50HKScpNyW7_U"; // <- o seu ID FIXO

/* ====================== CONFIGURAÇÃO ============================
 * Personalize a exibição por aba (título da guia exatamente como no Sheets)
 * - columns.order: define a ordem das colunas
 * - columns.rename: renomeia cabeçalhos
 * - columns.hide: oculta colunas
 * - defaultSort: { key: "Nome da Coluna", dir: "asc"|"desc" }
 * - summarize: lista de colunas numéricas para totalizar (exibe no rodapé)
 * - formats: { "Coluna": "number"|"currency"|"percent"|"date" }
 */
const CONFIG = {
  global: {
    rowsPerPage: 50,     // padrão
    headerRows: 1,       // quantas linhas são cabeçalho
    currency: { style: "currency", currency: "BRL", minimumFractionDigits: 2 },
    number: { maximumFractionDigits: 2 },
    percent: { style: "percent", maximumFractionDigits: 2 },
    locale: "pt-BR",
    cacheMinutes: 10,
  },
  // Exemplo de configuração por aba (ajuste os nomes das abas depois de carregar uma vez e ver os títulos)
  sheets: {
    // "Vendas 2025": {
    //   columns: {
    //     order: ["Data", "Cliente", "Produto", "Quantidade", "Preço", "Total"],
    //     rename: { "Preço": "Preço (R$)" },
    //     hide: ["Observações"]
    //   },
    //   defaultSort: { key: "Data", dir: "desc" },
    //   summarize: ["Total"],
    //   formats: { "Data": "date", "Preço": "currency", "Total": "currency", "Quantidade": "number" }
    // }
  }
};

/* ====================== Seletores de UI ======================== */
const UI = {
  error: document.getElementById("error"),
  tabs: document.getElementById("tabs"),
  tableContainer: document.getElementById("table-container"),
  tableLoader: document.getElementById("table-loader"),
  tableFooter: document.getElementById("table-footer"),
  searchInput: document.getElementById("search-input"),
  exportCSV: document.getElementById("export-csv"),
  exportAll: document.getElementById("export-all"),
  statTotal: document.getElementById("stat-total"),
  statCols: document.getElementById("stat-cols"),
  statSheet: document.getElementById("stat-sheet"),
  statUpdated: document.getElementById("stat-updated"),
  prevBtn: document.getElementById("prev"),
  nextBtn: document.getElementById("next"),
  pageInfo: document.getElementById("page-info"),
};

/* ========================= Estado ============================== */
const state = {
  worksheets: /** @type {Array<{title:string,gid:number}>} */ ([]),
  active: /** @type {{title:string,gid:number}|null} */ (null),
  rowsPerPage: CONFIG.global.rowsPerPage,
  headerRows: CONFIG.global.headerRows,
  page: 1,
  totalRows: 0,
  headers: /** @type {string[]} */ ([]),
  pageRows: /** @type {Array<Record<string, any>>} */ ([]),
  sortKey: null,
  sortDir: "asc",
  searchTerm: "",
};

const CACHE_PREFIX = "sheet-fixed:v1:";

/* ========================= Utils =============================== */
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
  return new Intl.DateTimeFormat(CONFIG.global.locale, { dateStyle: "short", timeStyle: "medium" }).format(d);
}
function debounce(fn, ms=250){ let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), ms); }; }

function toCSV(headers, rows){
  const esc = (v)=> {
    if (v == null) return "";
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s;
  };
  const lines = [headers.map(esc).join(",")];
  for (const r of rows) lines.push(headers.map(h=>esc(r[h])).join(","));
  return lines.join("\n");
}
function downloadFile(filename, content, mime="text/plain;charset=utf-8"){
  const blob = new Blob([content], {type: mime});
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement("a"), { href: url, download: filename });
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

/* ==================== Formatação de valores ==================== */
function formatValue(col, value){
  const sheetCfg = CONFIG.sheets[state.active?.title] || {};
  const fmt = (sheetCfg.formats && sheetCfg.formats[col]) || null;
  if (value == null || value === "") return "";
  try{
    switch(fmt){
      case "currency": return new Intl.NumberFormat(CONFIG.global.locale, CONFIG.global.currency).format(Number(value));
      case "percent": return new Intl.NumberFormat(CONFIG.global.locale, CONFIG.global.percent).format(Number(value));
      case "number":  return new Intl.NumberFormat(CONFIG.global.locale, CONFIG.global.number).format(Number(value));
      case "date": {
        // Tenta tratar tanto número serial do Google quanto string
        if (typeof value === "number") {
          // Serial do Google (dias desde 1899-12-30)
          const epoch = new Date(Date.UTC(1899, 11, 30));
          const ms = value * 86400000;
          const d = new Date(epoch.getTime() + ms);
          return new Intl.DateTimeFormat(CONFIG.global.locale).format(d);
        }
        const d = new Date(value);
        return isNaN(+d) ? String(value) : new Intl.DateTimeFormat(CONFIG.global.locale).format(d);
      }
      default: return String(value);
    }
  } catch { return String(value); }
}

/* ==================== Acesso ao Google Sheets ================== */
async function listWorksheets(sheetId){
  const url = `https://spreadsheets.google.com/feeds/worksheets/${sheetId}/public/full?alt=json`;
  try{
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    const entries = json?.feed?.entry ?? [];
    const tabs = [];
    for (const e of entries){
      const title = e?.title?.$t ?? "Sem título";
      let gid = null;
      for (const l of (e?.link || [])){
        const m = (l?.href || "").match(/[?&#]gid=(\d+)/);
        if (m) { gid = Number(m[1]); break; }
      }
      if (gid == null && e["gs$gid"]) gid = Number(e["gs$gid"]);
      if (gid == null && /\/od6$/.test(e?.id?.$t || "")) gid = 0;
      if (gid != null) tabs.push({ title, gid });
    }
    if (!tabs.length) throw new Error("Nenhuma aba pública encontrada.");
    return tabs;
  }catch(e){
    console.warn("Falha ao listar abas, usando fallback gid=0", e);
    return [{ title: "Planilha 1", gid: 0 }];
  }
}

function gvizUrl({ sheetId, gid, headers, query }){
  const p = new URLSearchParams({ tqx: "out:json", gid: String(gid), headers: String(headers) });
  if (query) p.set("tq", query);
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?${p.toString()}`;
}
function parseGviz(text){
  const json = JSON.parse(text.replace(/^[^{]+/, "").replace(/;?\s*$/, ""));
  if (json.status !== "ok") throw new Error(json?.errors?.[0]?.detailed_message || "Erro GViz");
  return json;
}
function tableToRows(table){
  const cols = table.cols || [];
  const headers = cols.map((c, i) => c.label || c.id || `Col${i+1}`);
  const rows = [];
  for (const r of (table.rows || [])){
    const obj = {};
    (r.c || []).forEach((cell, idx) => {
      let v = cell?.v ?? "";
      // preserva strings formatadas quando existirem
      if (cell?.f && typeof cell.v === "number") v = cell.v; // manter número cru; formatamos depois
      obj[headers[idx]] = v;
    });
    rows.push(obj);
  }
  return { headers, rows };
}
async function countRows(sheetId, gid, headers){
  const url = gvizUrl({ sheetId, gid, headers, query: "select count(A)" });
  const res = await fetch(url);
  if (!res.ok) throw new Error("Falha ao contar registros.");
  const json = parseGviz(await res.text());
  const cell = json?.table?.rows?.[0]?.c?.[0];
  return Number(cell?.v ?? 0);
}
async function fetchPage(sheetId, gid, headers, limit, offset){
  const q = `select * limit ${limit} offset ${offset}`;
  const res = await fetch(gvizUrl({ sheetId, gid, headers, query: q }));
  if (!res.ok) throw new Error("Falha ao carregar página.");
  const json = parseGviz(await res.text());
  return tableToRows(json.table);
}

/* ========================== Cache ============================== */
function cacheKeyPage(sheetId, gid, headerRows, page, limit){
  return `${CACHE_PREFIX}${sheetId}:${gid}:h${headerRows}:p${page}:l${limit}`;
}
function cacheKeyCount(sheetId, gid, headerRows){
  return `${CACHE_PREFIX}${sheetId}:${gid}:count:h${headerRows}`;
}
function cacheGet(key, maxMinutes){
  try{
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const { t, v } = JSON.parse(raw);
    if ((Date.now() - t) > maxMinutes*60*1000) return null;
    return v;
  }catch{ return null; }
}
function cacheSet(key, value){
  try{ localStorage.setItem(key, JSON.stringify({ t: Date.now(), v: value })); }catch{}
}

/* ======================== Renderização ========================= */
function applyColumnConfig(headers, rows){
  const sheetCfg = CONFIG.sheets[state.active?.title] || {};
  const rename = sheetCfg.columns?.rename || {};
  const hide = new Set(sheetCfg.columns?.hide || []);
  let ordered = [...headers];

  if (Array.isArray(sheetCfg.columns?.order) && sheetCfg.columns.order.length){
    const orderSet = new Set(sheetCfg.columns.order);
    const rest = ordered.filter(h => !orderSet.has(h) && !hide.has(h));
    ordered = sheetCfg.columns.order.filter(h => !hide.has(h)).concat(rest);
  } else {
    ordered = ordered.filter(h => !hide.has(h));
  }

  const finalHeaders = ordered.map(h => rename[h] || h);

  // mapeia valores com possíveis renomes
  const mappedRows = rows.map(r=>{
    const out = {};
    for (const h of ordered){
      const newKey = rename[h] || h;
      out[newKey] = r[h];
    }
    return out;
  });

  return { headers: finalHeaders, rows: mappedRows };
}

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

function renderFooterSummaries(headers, visibleRows){
  const sheetCfg = CONFIG.sheets[state.active?.title] || {};
  const totalsCols = sheetCfg.summarize || [];
  if (!totalsCols.length){ UI.tableFooter.classList.add("hidden"); UI.tableFooter.innerHTML = ""; return; }

  const sums = {};
  for (const col of totalsCols) sums[col] = 0;

  for (const row of visibleRows){
    for (const col of totalsCols){
      const v = Number(row[col]);
      if (!isNaN(v)) sums[col] += v;
    }
  }

  UI.tableFooter.classList.remove("hidden");
  UI.tableFooter.innerHTML = Object.entries(sums).map(([k, v])=>{
    const fmt = (CONFIG.sheets[state.active?.title]?.formats||{})[k];
    const val = formatValue(k, fmt ? v : v); // formata se houver config
    return `<div>${k}: <strong>${val}</strong></div>`;
  }).join("");
}

function renderTable({ headers, rows }){
  // busca na página
  let filtered = rows;
  if (state.searchTerm){
    const term = state.searchTerm.toLowerCase();
    filtered = filtered.filter(r => headers.some(k => String(r[k] ?? "").toLowerCase().includes(term)));
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

  // constrói tabela
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");

  headers.forEach(col => {
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
      renderTable({ headers, rows: state.pageRows });
    });
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  const tbody = document.createElement("tbody");
  if (!filtered.length){
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = headers.length;
    td.textContent = "Nenhum dado nesta página.";
    tr.appendChild(td);
    tbody.appendChild(tr);
  } else {
    for (const r of filtered){
      const tr = document.createElement("tr");
      for (const col of headers){
        const td = document.createElement("td");
        td.textContent = formatValue(col, r[col]) || "-";
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
  UI.statCols.textContent = String(headers.length);
  UI.statTotal.textContent = String(state.totalRows);
  UI.statUpdated.textContent = fmtDateTime();

  // tools habilitados
  UI.searchInput.disabled = false;
  UI.exportCSV.disabled = filtered.length === 0;
  UI.exportAll.disabled = state.totalRows === 0;

  // rodapé de somatórios (sobre linhas visíveis pós-busca/sort)
  renderFooterSummaries(headers, filtered);

  // paginação
  const totalPages = Math.max(1, Math.ceil(state.totalRows / state.rowsPerPage));
  UI.pageInfo.textContent = `Página ${state.page} de ${totalPages}`;
  UI.prevBtn.disabled = state.page <= 1;
  UI.nextBtn.disabled = state.page >= totalPages;

  // atualiza hash compartilhável
  updateHash();
}

/* ======================== Navegação/Hash ======================== */
function updateHash(){
  const params = new URLSearchParams();
  if (state.active) params.set("gid", String(state.active.gid));
  if (state.page > 1) params.set("page", String(state.page));
  if (state.searchTerm) params.set("q", state.searchTerm);
  if (state.sortKey) { params.set("sort", state.sortKey); params.set("dir", state.sortDir); }
  history.replaceState(null, "", `#${params.toString()}`);
}
function restoreFromHash(){
  const h = location.hash.replace(/^#/, "");
  if (!h) return {};
  const p = new URLSearchParams(h);
  return {
    gid: p.get("gid") ? Number(p.get("gid")) : null,
    page: p.get("page") ? Number(p.get("page")) : 1,
    q: p.get("q") || "",
    sort: p.get("sort") || null,
    dir: p.get("dir") || "asc",
  };
}

/* ========================== Controle =========================== */
async function setActiveTab(tab){
  state.active = tab;
  // sort default por aba
  const def = (CONFIG.sheets[tab.title] || {}).defaultSort;
  state.sortKey = def?.key || null;
  state.sortDir = def?.dir || "asc";
  state.page = 1;
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

    // total
    const ckCount = cacheKeyCount(SHEET_ID, gid, state.headerRows);
    let total = cacheGet(ckCount, CONFIG.global.cacheMinutes);
    if (total == null){
      total = await countRows(SHEET_ID, gid, state.headerRows);
      cacheSet(ckCount, total);
    }
    state.totalRows = Number(total) || 0;

    // página
    const ckPage = cacheKeyPage(SHEET_ID, gid, state.headerRows, state.page, limit);
    let page = cacheGet(ckPage, CONFIG.global.cacheMinutes/2);
    if (page == null){
      page = await fetchPage(SHEET_ID, gid, state.headerRows, limit, offset);
      cacheSet(ckPage, page);
    }

    // aplica config de colunas
    const mapped = applyColumnConfig(page.headers, page.rows);
    state.headers = mapped.headers;
    state.pageRows = mapped.rows;

    renderTable({ headers: state.headers, rows: state.pageRows });
  } catch(err){
    console.error(err);
    showError(err?.message || "Erro ao carregar dados.");
  } finally {
    setLoading(false);
  }
}

async function connect(){
  try{
    clearError();
    setLoading(true);

    const tabs = await listWorksheets(SHEET_ID);
    state.worksheets = tabs;

    renderTabs(tabs);

    // restauro do hash (gid, page, q, sort)
    const saved = restoreFromHash();
    let initial = tabs[0];
    if (saved.gid){
      const found = tabs.find(t => t.gid === saved.gid);
      if (found) initial = found;
    }

    await setActiveTab(initial);

    if (saved.page && saved.page > 1){
      state.page = saved.page; await loadActivePage();
    }
    if (saved.q){
      state.searchTerm = saved.q; UI.searchInput.value = saved.q;
      renderTable({ headers: state.headers, rows: state.pageRows });
    }
    if (saved.sort){
      state.sortKey = saved.sort; state.sortDir = saved.dir === "desc" ? "desc" : "asc";
      renderTable({ headers: state.headers, rows: state.pageRows });
    }
  }catch(err){
    console.error(err);
    showError("Não foi possível acessar a planilha pública. Verifique as permissões de compartilhamento.");
  }finally{
    setLoading(false);
  }
}

/* =========================== Eventos =========================== */
UI.prevBtn.addEventListener("click", async ()=>{
  if (state.page > 1){ state.page--; await loadActivePage(); }
});
UI.nextBtn.addEventListener("click", async ()=>{
  const totalPages = Math.max(1, Math.ceil(state.totalRows / state.rowsPerPage));
  if (state.page < totalPages){ state.page++; await loadActivePage(); }
});
UI.searchInput.addEventListener("input", debounce(()=>{
  state.searchTerm = UI.searchInput.value.trim();
  renderTable({ headers: state.headers, rows: state.pageRows });
}, 200));

UI.exportCSV.addEventListener("click", ()=>{
  const headers = state.headers;
  const term = state.searchTerm.toLowerCase();
  let rows = state.pageRows;
  if (term){
    rows = rows.filter(r => headers.some(h => String(r[h] ?? "").toLowerCase().includes(term)));
  }
  // aplica ordenação atual
  if (state.sortKey){
    const k = state.sortKey, dir = state.sortDir;
    rows = [...rows].sort((a,b)=>{
      const va = a[k] ?? "", vb = b[k] ?? "";
      if (va === vb) return 0;
      return (va > vb ? 1 : -1) * (dir === "asc" ? 1 : -1);
    });
  }
  const csv = toCSV(headers, rows);
  const name = (state.active?.title||"dados").replace(/\s+/g,'_');
  downloadFile(`${name}_p${state.page}.csv`, csv, "text/csv;charset=utf-8");
});

UI.exportAll.addEventListener("click", async ()=>{
  if (!state.active) return;
  const { gid } = state.active;
  const limit = 500; // exporta em blocos maiores para reduzir requisições
  const total = state.totalRows;
  const pages = Math.max(1, Math.ceil(total / limit));

  const allRows = [];
  let headers = null;

  setLoading(true);
  try{
    for (let p=1; p<=pages; p++){
      const offset = (p-1)*limit;
      const chunk = await fetchPage(SHEET_ID, gid, state.headerRows, limit, offset);
      const mapped = applyColumnConfig(chunk.headers, chunk.rows);
      headers = headers || mapped.headers;
      allRows.push(...mapped.rows);
    }
    // aplica busca/ordenação atuais
    let rows = allRows;
    if (state.searchTerm){
      const term = state.searchTerm.toLowerCase();
      rows = rows.filter(r => headers.some(h => String(r[h] ?? "").toLowerCase().includes(term)));
    }
    if (state.sortKey){
      const k = state.sortKey, dir = state.sortDir;
      rows = [...rows].sort((a,b)=>{
        const va = a[k] ?? "", vb = b[k] ?? "";
        if (va === vb) return 0;
        return (va > vb ? 1 : -1) * (dir === "asc" ? 1 : -1);
      });
    }
    const csv = toCSV(headers, rows);
    const name = (state.active?.title||"dados").replace(/\s+/g,'_');
    downloadFile(`${name}_completo.csv`, csv, "text/csv;charset=utf-8");
  } catch(err){
    showError("Falha ao exportar tudo. Tente novamente.");
  } finally {
    setLoading(false);
  }
});

/* Inicialização */
window.addEventListener("DOMContentLoaded", connect);
