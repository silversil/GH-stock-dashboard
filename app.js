/* global XLSX */
const PAGE_SIZE = 100;
const MIN_MATCH_LENGTH = 4;

const state = { excel: null, central: null, rows: [], stores: [], brands: [], diagnostics: null, visible: PAGE_SIZE };
const ui = Object.fromEntries([
  "status", "diagnosticsPanel", "diagnosticsText", "excelInput", "csvInput", "excelFileState", "csvFileState", "storeFilter", "brandFilter",
  "matchFilter", "minPieces", "searchInput", "sortFilter", "resetButton", "resultsBody", "emptyState", "resultsTitle",
  "resultCount", "exportButton", "loadMoreButton", "skuMetric", "piecesMetric", "storesMetric", "updatedMetric", "fileMetric"
].map((id) => [id, document.getElementById(id)]));

function text(value) { return String(value ?? "").trim(); }
function qty(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const normalized = text(value).replace(/\s/g, "").replace(/\.(?=\d{3}(?:\D|$))/g, "").replace(",", ".");
  const result = Number(normalized);
  return Number.isFinite(result) ? result : 0;
}
function compactStore(value) { return text(value).split(" - ")[0] || text(value); }
function isPush(value) { return /(^|[^A-Z])PUSH(?:[^A-Z]|\d|$)/i.test(text(value)); }
function normalizeSku(value, removeOt = false) {
  let normalized = text(value).replace(/\u00a0/g, " ").toLocaleUpperCase("it").replace(/[\s.]/g, "");
  if (removeOt) normalized = normalized.replace(/OT-/g, "");
  return normalized;
}
function escapeHtml(value) {
  return text(value).replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char]));
}
function setStatus(kind, title, detail) {
  ui.status.className = `status ${kind}`;
  ui.status.querySelector("strong").textContent = title;
  ui.status.querySelector("p").textContent = detail;
}
function findColumn(headers, tests) {
  return headers.find((header) => tests.some((test) => test.test(text(header)))) || null;
}
function firstSheetRows(arrayBuffer) {
  if (!globalThis.XLSX) throw new Error("Libreria di lettura file non disponibile. Controlla la connessione internet e riprova.");
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("Il file non contiene fogli leggibili.");
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
  if (!rows.length) throw new Error("Il file non contiene dati.");
  return rows;
}

function parseExcel(arrayBuffer, fileName) {
  const raw = firstSheetRows(arrayBuffer);
  const headers = Object.keys(raw[0]);
  const storeCol = findColumn(headers, [/descr.*negozio/i, /^negozio$/i, /^store$/i]);
  const skuCol = findColumn(headers, [/^articolo$/i, /^sku$/i, /codice.*articolo/i]);
  const descriptionCol = findColumn(headers, [/descrizione.*articolo/i, /^descrizione$/i]);
  const sizeCol = findColumn(headers, [/taglia/i, /^size$/i]);
  const qtyCol = findColumn(headers, [/^qt.*stock/i, /giacenza/i, /quantit/i, /^qty$/i]);
  if (!storeCol || !skuCol || !qtyCol) throw new Error("Nell’Excel servono le colonne Negozio, Articolo/SKU e Quantità in stock.");

  const stores = new Set();
  const grouped = new Map();
  let ignoredPushRows = 0;
  for (const row of raw) {
    const store = text(row[storeCol]);
    const sku = text(row[skuCol]);
    if (!store || !sku) continue;
    if (isPush(store)) { ignoredPushRows += 1; continue; }
    const skuKey = normalizeSku(sku);
    if (!skuKey) continue;
    stores.add(store);
    const key = `${store}\u0000${skuKey}`;
    const current = grouped.get(key) || { store, sku, skuKey, description: "", sizes: new Set(), qty: 0 };
    current.qty += qty(row[qtyCol]);
    if (!current.description && descriptionCol) current.description = text(row[descriptionCol]);
    if (sizeCol && text(row[sizeCol])) current.sizes.add(text(row[sizeCol]));
    grouped.set(key, current);
  }
  return { fileName, stores: [...stores].sort((a, b) => a.localeCompare(b, "it")), rows: [...grouped.values()].filter((row) => row.qty > 0), ignoredPushRows };
}

function parseCentralCsv(arrayBuffer, fileName) {
  const raw = firstSheetRows(arrayBuffer);
  const headers = Object.keys(raw[0]);
  const skuCol = findColumn(headers, [/^sku$/i, /^articolo$/i, /codice.*articolo/i]);
  const brandCol = findColumn(headers, [/^brand$/i, /^marca$/i]);
  if (!skuCol) throw new Error("Nel CSV Magento serve una colonna SKU.");

  const skuMap = new Map();
  let ignoredEmpty = 0;
  let otRows = 0;
  for (const row of raw) {
    const originalSku = text(row[skuCol]);
    if (!originalSku) { ignoredEmpty += 1; continue; }
    if (/OT-/i.test(originalSku)) otRows += 1;
    const skuKey = normalizeSku(originalSku, true);
    if (skuKey.length < MIN_MATCH_LENGTH) { ignoredEmpty += 1; continue; }
    const current = skuMap.get(skuKey) || { originals: new Set(), brands: new Set() };
    current.originals.add(originalSku);
    if (brandCol && text(row[brandCol])) current.brands.add(text(row[brandCol]));
    skuMap.set(skuKey, current);
  }
  const byLength = new Map();
  for (const skuKey of skuMap.keys()) {
    if (!byLength.has(skuKey.length)) byLength.set(skuKey.length, new Set());
    byLength.get(skuKey.length).add(skuKey);
  }
  const brandConflicts = [...skuMap.values()].filter((item) => item.brands.size > 1).length;
  return { fileName, skuMap, byLength, lengths: [...byLength.keys()].sort((a, b) => b - a), sourceRows: raw.length, ignoredEmpty, otRows, brandConflicts };
}

function findCentralMatch(excelSkuKey) {
  const { skuMap, byLength, lengths } = state.central;
  if (skuMap.has(excelSkuKey)) return { key: excelSkuKey, method: "exact", ambiguous: [] };
  const withoutPrefix = excelSkuKey.slice(3);
  if (withoutPrefix && skuMap.has(withoutPrefix)) return { key: withoutPrefix, method: "prefix3", ambiguous: [] };

  const candidates = new Set();
  for (const length of lengths) {
    if (length > excelSkuKey.length) continue;
    const index = byLength.get(length);
    for (let start = 0; start <= excelSkuKey.length - length; start += 1) {
      const candidate = excelSkuKey.slice(start, start + length);
      if (index.has(candidate)) candidates.add(candidate);
    }
  }
  if (!candidates.size) return null;
  const matches = [...candidates].sort((a, b) => b.length - a.length || a.localeCompare(b));
  const mostSpecific = matches[0];
  const independentMatches = matches.filter((candidate) => !mostSpecific.includes(candidate));
  if (independentMatches.length) return { key: null, method: "ambiguous", ambiguous: [mostSpecific, ...independentMatches] };
  return { key: mostSpecific, method: "contains", ambiguous: [] };
}

function centralSkuLabel(item) {
  if (!item) return "";
  const originals = [...item.originals].sort((a, b) => {
    const otDifference = Number(/OT-/i.test(a)) - Number(/OT-/i.test(b));
    return otDifference || a.localeCompare(b, "it", { numeric: true });
  });
  return `${originals.slice(0, 2).join(" / ")}${originals.length > 2 ? ` / +${originals.length - 2}` : ""}`;
}

function reconcile() {
  state.rows = [];
  state.brands = [];
  state.visible = PAGE_SIZE;
  if (!state.excel || !state.central) {
    const missing = !state.excel && !state.central ? "Excel negozi e CSV Magento" : !state.excel ? "Excel negozi" : "CSV Magento";
    setStatus("info", `Carica ${missing}`, "Il confronto parte automaticamente quando entrambi i file sono presenti.");
    state.stores = state.excel?.stores || [];
    populateFilters();
    render();
    return;
  }

  const uniqueSkus = new Map();
  for (const row of state.excel.rows) if (!uniqueSkus.has(row.skuKey)) uniqueSkus.set(row.skuKey, row);
  const matchCache = new Map();
  const methods = { exact: 0, prefix3: 0, contains: 0 };
  const ambiguous = [];

  for (const row of uniqueSkus.values()) {
    const match = findCentralMatch(row.skuKey);
    matchCache.set(row.skuKey, match);
    if (!match) continue;
    if (!match.key) { ambiguous.push({ sku: row.sku, candidates: match.ambiguous }); continue; }
    methods[match.method] += 1;
  }

  state.rows = state.excel.rows
    .map((row) => {
      const match = matchCache.get(row.skuKey);
      const base = { ...row, sizes: [...row.sizes].sort((a, b) => a.localeCompare(b, "it", { numeric: true })) };
      if (!match) return { ...base, matchStatus: "missing", matchMethod: "", csvSku: "", brand: "" };
      if (!match.key) {
        const candidateItems = match.ambiguous.map((key) => state.central.skuMap.get(key));
        return {
          ...base,
          matchStatus: "ambiguous",
          matchMethod: "ambiguous",
          csvSku: candidateItems.map(centralSkuLabel).filter(Boolean).join(" / "),
          brand: ""
        };
      }
      const centralItem = state.central.skuMap.get(match.key);
      return {
        ...base,
        matchStatus: "matched",
        matchMethod: match.method,
        csvSku: centralSkuLabel(centralItem),
        brand: [...centralItem.brands].sort((a, b) => a.localeCompare(b, "it")).join(" / ")
      };
    });
  state.stores = state.excel.stores;
  state.brands = [...new Set(state.rows.filter((row) => row.matchStatus === "matched" && row.brand).map((row) => row.brand))].sort((a, b) => a.localeCompare(b, "it"));
  state.diagnostics = { methods, ambiguous, uniqueExcelSkus: uniqueSkus.size };
  populateFilters();

  const matched = methods.exact + methods.prefix3 + methods.contains;
  const absentSkus = new Set(state.rows.filter((row) => row.matchStatus === "missing").map((row) => row.skuKey)).size;
  const exactNote = methods.exact === 0 ? " (normale se l’Excel aggiunge sempre un prefisso)" : "";
  ui.diagnosticsText.textContent = `Totale riconosciuto: ${matched.toLocaleString("it-IT")} · ${methods.exact.toLocaleString("it-IT")} coincidenze esatte dopo normalizzazione${exactNote} · ${methods.prefix3.toLocaleString("it-IT")} con prefisso di 3 caratteri · ${methods.contains.toLocaleString("it-IT")} contenuti come varianti · ${absentSkus.toLocaleString("it-IT")} assenti · ${ambiguous.length.toLocaleString("it-IT")} ambigui · ${state.central.otRows.toLocaleString("it-IT")} righe CSV con OT- normalizzato · ${state.central.brandConflicts.toLocaleString("it-IT")} SKU CSV con Brand discordanti`;
  ui.diagnosticsPanel.hidden = false;
  if (ambiguous.length) {
    const examples = ambiguous.slice(0, 3).map((item) => item.sku).join(", ");
    setStatus("warning", `${ambiguous.length} SKU con match ambiguo`, `Separate per prudenza nella vista “Match ambigui”: ${examples}${ambiguous.length > 3 ? "…" : ""}`);
  } else {
    setStatus("success", "Confronto completato", `${matched.toLocaleString("it-IT")} SKU riconosciute nel centrale; nessun match ambiguo. Le righe PUSH dell’Excel sono ignorate.`);
  }
  const now = new Date();
  ui.updatedMetric.textContent = new Intl.DateTimeFormat("it-IT", { dateStyle: "short", timeStyle: "short" }).format(now);
  ui.fileMetric.textContent = `${state.excel.fileName} + ${state.central.fileName}`;
  render();
}

function filters() {
  return { match: ui.matchFilter.value, store: ui.storeFilter.value, brand: ui.brandFilter.value, min: Math.max(1, Number(ui.minPieces.value) || 1), search: ui.searchInput.value.trim().toLocaleLowerCase("it"), sort: ui.sortFilter.value };
}
function filteredRows() {
  const active = filters();
  const rows = state.rows.filter((row) =>
    (active.match === "all" || row.matchStatus === active.match) &&
    (active.store === "all" || row.store === active.store) &&
    (active.brand === "all" || row.brand === active.brand) &&
    row.qty >= active.min &&
    (!active.search || `${row.sku} ${row.csvSku} ${row.description} ${row.brand}`.toLocaleLowerCase("it").includes(active.search))
  );
  rows.sort((a, b) => {
    if (active.sort === "qty-asc") return a.qty - b.qty || a.sku.localeCompare(b.sku);
    if (active.sort === "sku-asc") return a.sku.localeCompare(b.sku, "it", { numeric: true });
    return b.qty - a.qty || a.sku.localeCompare(b.sku);
  });
  return rows;
}
function exportFilteredRows() {
  if (!globalThis.XLSX) {
    setStatus("error", "Esportazione non disponibile", "La libreria Excel non è stata caricata. Controlla la connessione e riprova.");
    return;
  }
  const rows = filteredRows();
  if (!rows.length) return;
  try {
    const data = rows.map((row) => ({
      "SKU Excel": row.sku,
      "SKU CSV": row.csvSku,
      "Brand CSV": row.brand,
      Descrizione: row.description,
      Negozio: row.store,
      Taglie: row.sizes.join(", "),
      Giacenza: row.qty
    }));
    const sheet = XLSX.utils.json_to_sheet(data);
    sheet["!autofilter"] = { ref: sheet["!ref"] };
    sheet["!cols"] = [
      { wch: 24 }, { wch: 24 }, { wch: 22 }, { wch: 48 }, { wch: 30 }, { wch: 22 }, { wch: 12 }
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Risultati");
    const stamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
    XLSX.writeFile(workbook, `stock-gap-${ui.matchFilter.value}-${stamp}.xlsx`, { compression: true });
  } catch (error) {
    setStatus("error", "Esportazione non riuscita", error?.message || "Il browser non è riuscito a creare il file Excel.");
  }
}
function render() {
  const rows = filteredRows();
  const visible = rows.slice(0, state.visible);
  const titles = { missing: "Disponibili nei negozi, assenti nel CSV Magento", matched: "SKU riconosciute: verifica il matching Excel ↔ CSV", ambiguous: "SKU con più possibili corrispondenze", all: "Tutte le SKU dei negozi" };
  ui.resultsTitle.textContent = titles[ui.matchFilter.value] || titles.missing;
  ui.skuMetric.textContent = new Set(rows.map((row) => row.skuKey)).size.toLocaleString("it-IT");
  ui.piecesMetric.textContent = rows.reduce((sum, row) => sum + row.qty, 0).toLocaleString("it-IT", { maximumFractionDigits: 2 });
  ui.storesMetric.textContent = state.stores.length.toLocaleString("it-IT");
  ui.resultCount.textContent = `${rows.length.toLocaleString("it-IT")} ${rows.length === 1 ? "risultato" : "risultati"}`;
  ui.resultsBody.innerHTML = visible.map((row) => `<tr>
    <td>${escapeHtml(row.sku)}</td>
    <td>${escapeHtml(row.csvSku || "—")}</td>
    <td>${escapeHtml(row.brand || "—")}</td>
    <td class="description" title="${escapeHtml(row.description)}">${escapeHtml(row.description || "—")}</td>
    <td class="store" title="${escapeHtml(row.store)}">${escapeHtml(compactStore(row.store))}</td>
    <td class="sizes">${escapeHtml(row.sizes.join(", ") || "—")}</td>
    <td class="number"><strong>${row.qty.toLocaleString("it-IT", { maximumFractionDigits: 2 })}</strong></td>
  </tr>`).join("");
  const ready = Boolean(state.excel && state.central);
  ui.exportButton.disabled = !ready || rows.length === 0;
  ui.emptyState.hidden = !ready || rows.length !== 0;
  ui.loadMoreButton.hidden = visible.length >= rows.length;
}
function populateFilters() {
  const currentStore = state.stores.includes(ui.storeFilter.value) ? ui.storeFilter.value : "all";
  ui.storeFilter.innerHTML = `<option value="all">Tutti i negozi</option>${state.stores.map((store) => `<option value="${escapeHtml(store)}">${escapeHtml(compactStore(store))}</option>`).join("")}`;
  ui.storeFilter.value = currentStore;
  populateBrandFilter();
}
function populateBrandFilter() {
  const relevantBrands = ui.matchFilter.value === "missing" || ui.matchFilter.value === "ambiguous" ? [] : state.brands;
  const currentBrand = relevantBrands.includes(ui.brandFilter.value) ? ui.brandFilter.value : "all";
  ui.brandFilter.innerHTML = `<option value="all">Tutti i brand</option>${relevantBrands.map((brand) => `<option value="${escapeHtml(brand)}">${escapeHtml(brand)}</option>`).join("")}`;
  ui.brandFilter.value = currentBrand;
  ui.brandFilter.disabled = relevantBrands.length === 0;
}

ui.excelInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  setStatus("loading", "Lettura Excel negozi…", file.name);
  try {
    state.excel = parseExcel(await file.arrayBuffer(), file.name);
    ui.excelFileState.textContent = file.name;
    reconcile();
  } catch (error) { setStatus("error", "Excel non valido", error.message); }
  event.target.value = "";
});
ui.csvInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  setStatus("loading", "Lettura CSV Magento…", file.name);
  try {
    state.central = parseCentralCsv(await file.arrayBuffer(), file.name);
    ui.csvFileState.textContent = file.name;
    reconcile();
  } catch (error) { setStatus("error", "CSV Magento non valido", error.message); }
  event.target.value = "";
});
ui.matchFilter.addEventListener("input", () => { state.visible = PAGE_SIZE; populateBrandFilter(); render(); });
[ui.storeFilter, ui.brandFilter, ui.minPieces, ui.searchInput, ui.sortFilter].forEach((control) => control.addEventListener("input", () => { state.visible = PAGE_SIZE; render(); }));
ui.resetButton.addEventListener("click", () => {
  ui.matchFilter.value = "missing"; ui.storeFilter.value = "all"; ui.brandFilter.value = "all"; ui.minPieces.value = "1"; ui.searchInput.value = ""; ui.sortFilter.value = "qty-desc"; state.visible = PAGE_SIZE; populateBrandFilter(); render();
});
ui.loadMoreButton.addEventListener("click", () => { state.visible += PAGE_SIZE; render(); });
ui.exportButton.addEventListener("click", exportFilteredRows);

populateFilters();
render();
