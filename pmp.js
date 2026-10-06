/* global XLSX, firstSheetMatrix, text, escapeHtml, buildPmpResult, pmpExportMatrix */
const pmpState = { magento: null, inventory: null, result: null, visible: 100, loading: 0, selectedStores: null };
const el = (id) => document.getElementById(id);
const numberLabel = (value) => value.toLocaleString("it-IT", { maximumFractionDigits: 6 });
function status(kind, title, detail) {
  el("status").className = `status ${kind}`;
  el("status").querySelector("strong").textContent = title;
  el("status").querySelector("p").textContent = detail;
}
function parseSource(matrix, kind) {
  const skuTest = kind === "inventory" ? /^(cod|sku|articolo|codice.*articolo)$/i : /^(sku|articolo|codice.*articolo)$/i;
  const priceTest = /^costo\s*uni\s*pm$/i;
  const headerIndex = matrix.findIndex((row) => row.some((v) => skuTest.test(text(v))) && (kind === "magento" || row.some((v) => priceTest.test(text(v)))));
  if (headerIndex < 0) throw new Error(kind === "magento" ? "Nel CSV serve la colonna SKU." : "Nell’Excel servono Articolo/SKU e Costo Uni PM.");
  const headers = matrix[headerIndex];
  if (kind === "magento" && headers.some((v) => text(v).toUpperCase() === "PMP NEGOZI")) throw new Error("Il CSV contiene già PMP NEGOZI. Carica il CSV originale senza questa colonna.");
  let rows = matrix.slice(headerIndex + 1);
  if (!rows.length) throw new Error("Il file non contiene righe dati.");
  if (rows.some((row) => row.length > headers.length)) throw new Error("Sono presenti righe con più colonne delle intestazioni: verifica il file.");
  const skuCol = headers.findIndex((v) => skuTest.test(text(v)));
  const priceCol = headers.findIndex((v) => priceTest.test(text(v)));
  let storeCol = headers.findIndex((v) => /descr.*negozio|^negozio$|^store$/i.test(text(v)));
  const descriptionCol = headers.findIndex((v) => /^descrizione$/i.test(text(v)));
  const groupedReport = kind === "inventory" && /^cod$/i.test(text(headers[skuCol])) && storeCol < 0 && descriptionCol >= 0 && headers.some((v) => /^esistenza$/i.test(text(v)));
  const sourceRowNumbers = [];
  let store = "";
  if (groupedReport) storeCol = headers.length;
  rows = rows.flatMap((row, index) => {
    if (row.every((value) => !text(value))) return [];
    if (groupedReport) {
      const section = text(row[skuCol]) && text(row[descriptionCol]) && row.every((value, col) => col === skuCol || col === descriptionCol || !text(value));
      if (section) { store = `${text(row[skuCol])} - ${text(row[descriptionCol])}`; return []; }
      if (!store) throw new Error("Articolo senza intestazione negozio nel report PMP.");
      row = Array.from({ length: headers.length }, (_, col) => row[col] ?? "").concat(store);
    }
    sourceRowNumbers.push(headerIndex + index + 2);
    return [row];
  });
  if (!rows.length) throw new Error("Il file non contiene righe articolo.");
  return { headers, rows, headerIndex, skuCol, priceCol, storeCol, sourceRowNumbers };
}
function populateStoreFilter() {
  pmpState.selectedStores = null;
  const inventory = pmpState.inventory;
  const stores = inventory && inventory.storeCol >= 0 ? [...new Set(inventory.rows.map((row) => text(row[inventory.storeCol])).filter((store) => store && !isPush(store)))].sort((a, b) => a.localeCompare(b, "it")) : [];
  el("storeFilter").innerHTML = '<option value="" selected>Tutti i negozi</option>' + stores.map((store) => `<option value="${escapeHtml(store)}">${escapeHtml(store)}</option>`).join("");
}
function reconcilePmp() {
  if (!pmpState.magento || !pmpState.inventory) return;
  pmpState.result = buildPmpResult(pmpState.magento, pmpState.inventory, pmpState.selectedStores);
  pmpState.visible = 100;
  const scope = pmpState.selectedStores === null ? "Tutte le righe Magento" : `Solo articoli dei ${pmpState.selectedStores.length} negozi selezionati e relativi figli; PMP calcolato solo sui negozi selezionati`;
  status(pmpState.result.issues.length ? "warning" : "success", "Elaborazione completata", `${scope}. ${pmpState.result.issues.length} righe negozi escluse nel report. Ogni configurabile segue i propri figli.`);
}
function renderPmp() {
  el("storeFilter").disabled = pmpState.loading > 0 || !pmpState.inventory || pmpState.inventory.storeCol < 0;
  el("csvInput").disabled = pmpState.loading > 0;
  el("excelInput").disabled = pmpState.loading > 0;
  const result = pmpState.result;
  const rows = result?.rows || [];
  const matched = rows.filter((r) => r.price !== null).length;
  el("totalMetric").textContent = rows.length.toLocaleString("it-IT");
  el("matchedMetric").textContent = matched.toLocaleString("it-IT");
  el("missingMetric").textContent = (rows.length - matched).toLocaleString("it-IT");
  el("issuesMetric").textContent = (result?.issues.length || 0).toLocaleString("it-IT");
  const search = el("searchInput").value.trim().toUpperCase();
  const view = el("matchFilter").value;
  const filtered = rows.filter((r) => (view === "all" || (view === "matched" ? r.price !== null : r.price === null)) && `${r.sku} ${r.skus.join(" ")}`.toUpperCase().includes(search));
  el("resultCount").textContent = `${filtered.length.toLocaleString("it-IT")} righe`;
  el("resultsBody").innerHTML = filtered.slice(0, pmpState.visible).map((r) => `<tr><td>${escapeHtml(r.sku)}</td><td>${escapeHtml(r.skus.join(" / ") || "—")}</td><td>${escapeHtml(r.stores.join(" / ") || "—")}</td><td class="number">${r.count}</td><td class="number"><strong>${r.price === null ? "—" : numberLabel(r.price)}</strong></td></tr>`).join("");
  el("loadMoreButton").hidden = filtered.length <= pmpState.visible;
  el("emptyState").hidden = filtered.length > 0;
  el("exportButton").disabled = !rows.length || pmpState.loading > 0;
  el("exportButton").textContent = pmpState.selectedStores === null ? "↓ Scarica Excel completo" : "↓ Scarica Excel negozi selezionati";
  el("reportButton").disabled = !result?.issues.length || pmpState.loading > 0;
}
for (const [id, kind, label] of [["csvInput", "magento", "csvFileState"], ["excelInput", "inventory", "excelFileState"]]) {
  el(id).addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (!file) return;
    pmpState.loading++;
    pmpState[kind] = null;
    pmpState.result = null;
    el(label).textContent = "Lettura…";
    renderPmp();
    status("loading", "Lettura file…", file.name);
    try {
      pmpState[kind] = parseSource(firstSheetMatrix(await file.arrayBuffer(), file.name, true), kind);
      if (kind === "inventory") populateStoreFilter();
      el(label).textContent = file.name;
      if (pmpState.magento && pmpState.inventory) {
        reconcilePmp();
      } else status("info", "Carica il secondo file", "Servono il CSV Magento e l’Excel negozi con Costo Uni PM.");
    } catch (error) {
      el(label).textContent = "File non valido";
      status("error", "Caricamento non riuscito", error.message);
    } finally { pmpState.loading--; event.target.value = ""; renderPmp(); }
  });
}
async function download(matrix, name, sheetName, priceCol = -1) {
  try {
    if (matrix.length > 1048576 || matrix[0].length > 16384) throw new Error("I dati superano i limiti di un foglio Excel.");
    if (priceCol >= 0) {
      pmpState.loading++;
      renderPmp();
      try {
        const book = createPmpWorkbook(matrix, sheetName, priceCol);
        const buffer = await book.xlsx.writeBuffer();
        const url = URL.createObjectURL(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
        const link = document.createElement("a");
        link.href = url;
        link.download = name;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 60000);
      } finally { pmpState.loading--; renderPmp(); }
      return;
    }
    const sheet = XLSX.utils.aoa_to_sheet(matrix);
    sheet["!autofilter"] = { ref: sheet["!ref"] };
    if (priceCol >= 0) for (let r = 1; r < matrix.length; r++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c: priceCol })];
      if (cell?.t === "n") cell.z = "0.00####";
      for (let c = 2; c <= 5 && c < priceCol; c++) {
        const numericCell = sheet[XLSX.utils.encode_cell({ r, c })];
        if (numericCell?.t === "n") numericCell.z = "0.00####";
      }
    }
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, sheetName);
    XLSX.writeFile(book, name, { compression: true });
  } catch (error) { status("error", "Esportazione non riuscita", error.message); }
}
el("exportButton").addEventListener("click", () => {
  if (!pmpState.result) return;
  download(pmpExportMatrix(pmpState.magento, pmpState.result), "magento-prezzi-acquisto.xlsx", "Magento", pmpState.magento.headers.length);
});
el("reportButton").addEventListener("click", () => {
  if (!pmpState.result) return;
  download([["Riga Excel", "SKU Excel", "Negozio", "PMP originale", "Motivo esclusione", "Candidati normalizzati"], ...pmpState.result.issues], "pmp-righe-escluse.xlsx", "Righe escluse");
});
for (const id of ["searchInput", "matchFilter"]) el(id).addEventListener("input", () => { pmpState.visible = 100; renderPmp(); });
el("storeFilter").addEventListener("change", () => {
  const selected = Array.from(el("storeFilter").selectedOptions, (option) => option.value);
  pmpState.selectedStores = selected.includes("") ? null : selected;
  reconcilePmp();
  renderPmp();
});
el("loadMoreButton").addEventListener("click", () => { pmpState.visible += 100; renderPmp(); });
renderPmp();
