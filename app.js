/* global XLSX */
const PAGE_SIZE = 100;

const state = { rows: [], stores: [], pushFound: false, visible: PAGE_SIZE, fileName: "Nessun file", sourceBuffer: null };
const ui = Object.fromEntries([
  "status", "refreshButton", "fileInput", "storeFilter", "minPieces", "minOutput", "searchInput",
  "sortFilter", "resetButton", "resultsBody", "emptyState", "resultCount", "loadMoreButton",
  "skuMetric", "piecesMetric", "storesMetric", "updatedMetric", "fileMetric"
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
function canonicalSku(value) { return text(value).replace(/\u00a0/g, " ").replace(/\s+/g, " ").toLocaleUpperCase("it"); }
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

function parseWorkbook(arrayBuffer) {
  if (!globalThis.XLSX) throw new Error("Libreria Excel non disponibile. Controlla la connessione internet e riprova.");
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("Il file non contiene fogli leggibili.");
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
  if (!raw.length) throw new Error("Il primo foglio non contiene dati.");

  const headers = Object.keys(raw[0]);
  const storeCol = findColumn(headers, [/descr.*negozio/i, /^negozio$/i, /^store$/i]);
  const skuCol = findColumn(headers, [/^articolo$/i, /^sku$/i, /codice.*articolo/i]);
  const descriptionCol = findColumn(headers, [/descrizione.*articolo/i, /^descrizione$/i]);
  const sizeCol = findColumn(headers, [/taglia/i, /^size$/i]);
  const qtyCol = findColumn(headers, [/^qt.*stock/i, /giacenza/i, /quantit/i, /^qty$/i]);
  if (!storeCol || !skuCol || !qtyCol) throw new Error("Colonne richieste non trovate: servono Negozio, Articolo/SKU e Quantità in stock.");

  const detail = raw.filter((row) => text(row[skuCol]) && text(row[storeCol]));
  const pushBySku = new Map();
  const stores = new Set();
  const grouped = new Map();

  for (const row of detail) {
    const store = text(row[storeCol]);
    const sku = text(row[skuCol]);
    const skuKey = canonicalSku(sku);
    const amount = qty(row[qtyCol]);
    if (isPush(store)) {
      pushBySku.set(skuKey, (pushBySku.get(skuKey) || 0) + amount);
      continue;
    }
    stores.add(store);
    const key = `${store}\u0000${skuKey}`;
    const current = grouped.get(key) || { store, sku, skuKey, description: "", sizes: new Set(), qty: 0 };
    current.qty += amount;
    if (!current.description && descriptionCol) current.description = text(row[descriptionCol]);
    if (sizeCol && text(row[sizeCol])) current.sizes.add(text(row[sizeCol]));
    grouped.set(key, current);
  }

  state.pushFound = detail.some((row) => isPush(row[storeCol]));
  state.stores = [...stores].sort((a, b) => a.localeCompare(b, "it"));
  state.rows = [...grouped.values()]
    .filter((row) => row.qty > 0 && (pushBySku.get(row.skuKey) || 0) <= 0)
    .map((row) => ({ ...row, sizes: [...row.sizes].sort((a, b) => a.localeCompare(b, "it", { numeric: true })) }));
}

function filters() {
  return {
    store: ui.storeFilter.value,
    min: Number(ui.minPieces.value),
    search: ui.searchInput.value.trim().toLocaleLowerCase("it"),
    sort: ui.sortFilter.value,
  };
}

function filteredRows() {
  const active = filters();
  const rows = state.rows.filter((row) =>
    (active.store === "all" || row.store === active.store) && row.qty >= active.min &&
    (!active.search || `${row.sku} ${row.description}`.toLocaleLowerCase("it").includes(active.search))
  );
  rows.sort((a, b) => {
    if (active.sort === "qty-asc") return a.qty - b.qty || a.sku.localeCompare(b.sku);
    if (active.sort === "sku-asc") return a.sku.localeCompare(b.sku, "it", { numeric: true });
    return b.qty - a.qty || a.sku.localeCompare(b.sku);
  });
  return rows;
}

function render() {
  ui.minOutput.value = ui.minPieces.value;
  const rows = filteredRows();
  const visible = rows.slice(0, state.visible);
  const uniqueSkus = new Set(rows.map((row) => row.sku)).size;
  const pieces = rows.reduce((sum, row) => sum + row.qty, 0);
  ui.skuMetric.textContent = uniqueSkus.toLocaleString("it-IT");
  ui.piecesMetric.textContent = pieces.toLocaleString("it-IT", { maximumFractionDigits: 2 });
  ui.storesMetric.textContent = state.stores.length.toLocaleString("it-IT");
  ui.resultCount.textContent = `${rows.length.toLocaleString("it-IT")} ${rows.length === 1 ? "risultato" : "risultati"}`;
  ui.resultsBody.innerHTML = visible.map((row) => `<tr>
    <td>${escapeHtml(row.sku)}</td>
    <td class="description" title="${escapeHtml(row.description)}">${escapeHtml(row.description || "—")}</td>
    <td class="store" title="${escapeHtml(row.store)}">${escapeHtml(compactStore(row.store))}</td>
    <td class="sizes">${escapeHtml(row.sizes.join(", ") || "—")}</td>
    <td class="number"><strong>${row.qty.toLocaleString("it-IT", { maximumFractionDigits: 2 })}</strong></td>
  </tr>`).join("");
  ui.emptyState.hidden = !state.sourceBuffer || rows.length !== 0;
  ui.loadMoreButton.hidden = visible.length >= rows.length;
}

function populateStores() {
  ui.storeFilter.innerHTML = `<option value="all">Tutti i negozi</option>${state.stores.map((store) =>
    `<option value="${escapeHtml(store)}">${escapeHtml(compactStore(store))}</option>`).join("")}`;
}

async function loadBuffer(buffer, fileName) {
  parseWorkbook(buffer);
  state.sourceBuffer = buffer.slice(0);
  state.fileName = fileName;
  state.visible = PAGE_SIZE;
  populateStores();
  const now = new Date();
  ui.updatedMetric.textContent = new Intl.DateTimeFormat("it-IT", { dateStyle: "short", timeStyle: "short" }).format(now);
  ui.fileMetric.textContent = fileName;
  ui.refreshButton.disabled = false;
  if (state.pushFound) {
    setStatus("success", "Dati aggiornati", `${state.rows.length.toLocaleString("it-IT")} combinazioni negozio/SKU confrontate con PUSH.`);
  } else {
    setStatus("warning", "Magazzino PUSH non trovato", `Il file contiene ${state.stores.length} negozi ma nessuna riga PUSH. I risultati restano nascosti per evitare confronti errati.`);
    state.rows = [];
  }
  render();
}

ui.refreshButton.addEventListener("click", async () => {
  if (!state.sourceBuffer) return;
  setStatus("loading", "Rilettura in corso…", state.fileName);
  try { await loadBuffer(state.sourceBuffer, state.fileName); }
  catch (error) { setStatus("error", "Impossibile rileggere il file", error.message); }
});
ui.fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  setStatus("loading", "Lettura del nuovo file…", file.name);
  try { await loadBuffer(await file.arrayBuffer(), file.name); }
  catch (error) { setStatus("error", "File non valido", error.message); }
  event.target.value = "";
});
[ui.storeFilter, ui.minPieces, ui.searchInput, ui.sortFilter].forEach((control) => control.addEventListener("input", () => { state.visible = PAGE_SIZE; render(); }));
ui.resetButton.addEventListener("click", () => {
  ui.storeFilter.value = "all"; ui.minPieces.value = "1"; ui.searchInput.value = ""; ui.sortFilter.value = "qty-desc"; state.visible = PAGE_SIZE; render();
});
ui.loadMoreButton.addEventListener("click", () => { state.visible += PAGE_SIZE; render(); });

render();
