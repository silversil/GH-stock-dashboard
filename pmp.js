/* global XLSX, firstSheetMatrix, text, escapeHtml, buildPmpResult, pmpExportMatrix */
const pmpState = { magento: null, inventory: null, result: null, visible: 100, loading: 0 };
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
function renderPmp() {
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
  el("exportButton").disabled = !result || pmpState.loading > 0;
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
      el(label).textContent = file.name;
      if (pmpState.magento && pmpState.inventory) {
        pmpState.result = buildPmpResult(pmpState.magento, pmpState.inventory);
        pmpState.visible = 100;
        status(pmpState.result.issues.length || pmpState.result.orphanSimpleCount ? "warning" : "success", "Elaborazione completata", `${pmpState.result.issues.length} righe negozi escluse e disponibili nel report; ${pmpState.result.push} righe PUSH ignorate. ${pmpState.result.orphanSimpleCount} semplici senza configurabile nel CSV: PMP vuoto. Il download include tutte le righe Magento, raggruppate per SKU da Z ad A, con ogni configurabile subito sotto i suoi figli.`);
      } else status("info", "Carica il secondo file", "Servono il CSV Magento e l’Excel negozi con Costo Uni PM.");
    } catch (error) {
      el(label).textContent = "File non valido";
      status("error", "Caricamento non riuscito", error.message);
    } finally { pmpState.loading--; event.target.value = ""; renderPmp(); }
  });
}
function download(matrix, name, sheetName, priceCol = -1) {
  try {
    if (matrix.length > 1048576 || matrix[0].length > 16384) throw new Error("I dati superano i limiti di un foglio Excel.");
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
el("loadMoreButton").addEventListener("click", () => { pmpState.visible += 100; renderPmp(); });
renderPmp();
