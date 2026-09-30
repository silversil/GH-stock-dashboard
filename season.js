/* global escapeHtml, seasonExportCsv */
const seasonState = { result: null, page: 0, loading: false };
const seasonEl = id => document.getElementById(id);
const seasonLabels = { matched: "Coincide", mismatch: "Da aggiornare", unresolved: "Da verificare" };
let seasonWorker;
function seasonStatus(kind, title, detail) {
  const element = seasonEl("status");
  element.className = `status ${kind}`;
  element.querySelector("strong").textContent = title;
  element.querySelector("p").textContent = detail;
}
function renderSeason() {
  const rows = seasonState.result?.rows || [];
  for (const key of Object.keys(seasonLabels)) seasonEl(`${key}Metric`).textContent = rows.filter(row => row.status === key).length.toLocaleString("it-IT");
  seasonEl("excludedMetric").textContent = (seasonState.result?.excluded || 0).toLocaleString("it-IT");
  const query = seasonEl("searchInput").value.trim().toUpperCase();
  const view = seasonEl("matchFilter").value;
  const filtered = rows.filter(row => (view === "all" || row.status === view) && `${row.sku} ${row.skus.join(" ")}`.toUpperCase().includes(query));
  const pages = Math.ceil(filtered.length / 100);
  seasonState.page = Math.min(seasonState.page, Math.max(0, pages - 1));
  seasonEl("resultsBody").innerHTML = filtered.slice(seasonState.page * 100, (seasonState.page + 1) * 100).map(row => `<tr><td>${escapeHtml(row.sku)}</td><td>${escapeHtml(row.skus.join(" / ") || "—")}${row.inherited ? '<span class="season-detail">Dal configurabile</span>' : ""}</td><td>${escapeHtml(row.old || "—")}</td><td><strong>${escapeHtml(row.expected || "—")}</strong></td><td><span class="season-badge ${row.status}">${seasonLabels[row.status]}</span>${row.reason ? `<span class="season-detail">${escapeHtml(row.reason)}</span>` : ""}</td></tr>`).join("");
  seasonEl("resultCount").textContent = `${filtered.length.toLocaleString("it-IT")} articoli`;
  seasonEl("pageLabel").textContent = `Pagina ${pages ? seasonState.page + 1 : 0} di ${pages.toLocaleString("it-IT")}`;
  seasonEl("previousPage").disabled = seasonState.page === 0;
  seasonEl("nextPage").disabled = seasonState.page + 1 >= pages;
  seasonEl("emptyState").hidden = filtered.length > 0;
  seasonEl("exportButton").disabled = seasonState.loading || !rows.some(row => row.status === "mismatch");
  seasonEl("csvInput").disabled = seasonEl("excelInput").disabled = seasonState.loading;
}
function getSeasonWorker() {
  if (seasonWorker) return seasonWorker;
  seasonWorker = new Worker("season-worker.js");
  seasonWorker.onmessage = ({ data }) => {
    seasonState.loading = false;
    const label = seasonEl(data.kind === "magento" ? "csvFileState" : "excelFileState");
    if (data.error) { label.textContent = "File non valido"; seasonStatus("error", "Caricamento non riuscito", data.error); }
    else {
      label.textContent = data.name;
      seasonState.result = data.result;
      if (data.result) {
        const unresolved = data.result.rows.filter(row => row.status === "unresolved").length;
        seasonStatus(unresolved ? "warning" : "success", "Confronto completato", `${data.result.rows.length.toLocaleString("it-IT")} articoli confrontati, ${data.result.excluded.toLocaleString("it-IT")} articoli OUTLET/CORE esclusi. ${unresolved.toLocaleString("it-IT")} articoli da verificare. Righe BestStore: ${data.result.unmatchedExcel.toLocaleString("it-IT")} senza match Magento, ${data.result.ambiguousExcel.toLocaleString("it-IT")} con match ambiguo, ${data.result.invalidExcel.toLocaleString("it-IT")} abbinate con stagione non valida.`);
      } else seasonStatus("info", "Carica il secondo file", "Servono il CSV Magento e l’Excel BestStore.");
    }
    renderSeason();
  };
  seasonWorker.onerror = () => {
    seasonState.loading = false;
    seasonState.result = null;
    seasonWorker.terminate(); seasonWorker = null;
    seasonEl("csvFileState").textContent = seasonEl("excelFileState").textContent = "Da ricaricare";
    seasonStatus("error", "Lettura non disponibile", "Apri la dashboard via HTTP e verifica che i file dell’app siano disponibili, poi ricarica entrambi i file.");
    renderSeason();
  };
  return seasonWorker;
}
for (const [id, kind, label] of [["csvInput", "magento", "csvFileState"], ["excelInput", "inventory", "excelFileState"]]) {
  seasonEl(id).addEventListener("change", async event => {
    const file = event.target.files[0];
    if (!file) return;
    seasonState.loading = true; seasonState.result = null; seasonState.page = 0;
    seasonEl(label).textContent = "Lettura…";
    seasonStatus("loading", "Lettura e confronto in corso…", "Per gli Excel di grandi dimensioni l’elaborazione può richiedere qualche istante.");
    renderSeason();
    try {
      const buffer = await file.arrayBuffer();
      getSeasonWorker().postMessage({ kind, name: file.name, buffer }, [buffer]);
    } catch (error) {
      seasonState.loading = false;
      // Reset the worker so a failed replacement cannot leave an old source active.
      seasonWorker?.terminate(); seasonWorker = null;
      seasonEl("csvFileState").textContent = seasonEl("excelFileState").textContent = "Da ricaricare";
      seasonStatus("error", "Caricamento non riuscito", `${error.message}. Ricarica entrambi i file.`); renderSeason();
    } finally { event.target.value = ""; }
  });
}
for (const id of ["searchInput", "matchFilter"]) seasonEl(id).addEventListener("input", () => { seasonState.page = 0; renderSeason(); });
seasonEl("previousPage").addEventListener("click", () => { seasonState.page--; renderSeason(); seasonEl("resultCount").scrollIntoView({ block: "start" }); });
seasonEl("nextPage").addEventListener("click", () => { seasonState.page++; renderSeason(); seasonEl("resultCount").scrollIntoView({ block: "start" }); });
seasonEl("exportButton").addEventListener("click", () => {
  if (!seasonState.result || seasonState.loading) return;
  const url = URL.createObjectURL(new Blob([seasonExportCsv(seasonState.result)], { type: "text/csv;charset=utf-8" }));
  const link = document.createElement("a"); link.href = url; link.download = "magento-season-drop-differenze.csv";
  document.body.appendChild(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
});
renderSeason();
