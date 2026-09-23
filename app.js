/* global XLSX */
const PAGE_SIZE = 100;
const MIN_MATCH_LENGTH = 4;
const MAX_EXCEL_DATA_ROWS = 1048575;
const CLOTHING_SIZE_ORDER = new Map([
  ["XXS", 0], ["XS", 1], ["S", 2], ["S/M", 2.5], ["M", 3], ["M/L", 3.5], ["L", 4],
  ["L/XL", 4.5], ["XL", 5], ["XXL", 6], ["3XL", 7], ["4XL", 8]
]);
const CLOTHING_SIZE_ALIASES = new Map([
  ["2XS", "XXS"], ["SM", "S"], ["MD", "M"], ["LG", "L"], ["LXL", "L/XL"],
  ["2XL", "XXL"], ["XXXL", "3XL"], ["XXXXL", "4XL"], ["UNI", "OS"],
  ["ONESIZE", "OS"], ["OSFA", "OS"]
]);

const state = { excel: null, central: null, catalog: null, rows: [], stores: [], brands: [], diagnostics: null, visible: PAGE_SIZE };
const ui = Object.fromEntries([
  "status", "diagnosticsPanel", "diagnosticsText", "excelInput", "csvInput", "catalogInput", "excelFileState", "csvFileState", "catalogFileState", "storeFilter", "brandFilter",
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
function catalogSkuKey(value) {
  return text(value).replace(/\u00a0/g, " ").toLocaleUpperCase("it");
}
function exportSize(value) {
  const size = text(value).replace(/^=/, "").replace(/\*+$/, "").trim();
  const halfSize = size.match(/^(\d+)\s*-\s*$/);
  return halfSize ? `${halfSize[1]}.5` : size;
}
function normalizeSizeKey(value) {
  const label = exportSize(value).toLocaleUpperCase("it").replace(/\s/g, "");
  return CLOTHING_SIZE_ALIASES.get(label) || label;
}
function normalizeBarcode(value) {
  return text(value).replace(/^=/, "").replace(/\s/g, "").replace(/\.0$/, "");
}
function compareSizes(left, right) {
  const describe = (value) => {
    const label = exportSize(value).toLocaleUpperCase("it").replace(/\s/g, "");
    const normalized = normalizeSizeKey(value);
    if (/^\d+(?:[.,]\d+)?$/.test(normalized)) return { category: 0, rank: Number(normalized.replace(",", ".")), label };
    if (CLOTHING_SIZE_ORDER.has(normalized)) return { category: 1, rank: CLOTHING_SIZE_ORDER.get(normalized), label };
    return { category: 2, rank: 0, label };
  };
  const a = describe(left);
  const b = describe(right);
  return a.category - b.category || a.rank - b.rank || a.label.localeCompare(b.label, "it", { numeric: true });
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
function firstSheetMatrix(arrayBuffer, fileName) {
  if (!globalThis.XLSX) throw new Error("Libreria di lettura file non disponibile. Controlla la connessione internet e riprova.");
  let workbook;
  if (/\.csv$/i.test(fileName)) {
    const bytes = new Uint8Array(arrayBuffer);
    let source;
    if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
      source = new TextDecoder("utf-8").decode(bytes);
    } else {
      try { source = new TextDecoder("utf-8", { fatal: true }).decode(bytes); }
      catch { source = new TextDecoder("windows-1252").decode(bytes); }
    }
    workbook = XLSX.read(source, { type: "string", raw: true, dense: true });
  } else {
    workbook = XLSX.read(arrayBuffer, { type: "array", raw: true, dense: true });
  }
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("Il file non contiene fogli leggibili.");
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true, blankrows: false });
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
  const barcodeCol = findColumn(headers, [/^bcr$/i, /^barcode$/i, /^ean(?:13)?$/i]);
  const qtyCol = findColumn(headers, [/^qt.*stock/i, /giacenza/i, /quantit/i, /^qty$/i]);
  if (!storeCol || !skuCol || !qtyCol) throw new Error("Nell’Excel servono le colonne Negozio, Articolo/SKU e Quantità in stock.");

  const stores = new Set();
  const grouped = new Map();
  const detailGrouped = new Map();
  let ignoredPushRows = 0;
  for (const row of raw) {
    const store = text(row[storeCol]);
    const sku = text(row[skuCol]);
    if (!store || !sku) continue;
    if (isPush(store)) { ignoredPushRows += 1; continue; }
    const skuKey = normalizeSku(sku);
    const catalogKey = catalogSkuKey(sku);
    if (!skuKey || !catalogKey) continue;
    stores.add(store);
    const key = `${store}\u0000${catalogKey}`;
    const current = grouped.get(key) || { store, sku, skuKey, catalogKey, description: "", sizes: new Set(), qty: 0 };
    current.qty += qty(row[qtyCol]);
    if (!current.description && descriptionCol) current.description = text(row[descriptionCol]);
    if (sizeCol && text(row[sizeCol])) current.sizes.add(text(row[sizeCol]));
    grouped.set(key, current);

    const size = sizeCol ? text(row[sizeCol]) : "";
    const barcode = barcodeCol ? text(row[barcodeCol]) : "";
    const detailKey = `${key}\u0000${size}\u0000${barcode}`;
    const detail = detailGrouped.get(detailKey) || { store, sku, skuKey, catalogKey, size, barcode, description: "", qty: 0 };
    detail.qty += qty(row[qtyCol]);
    if (!detail.description && descriptionCol) detail.description = text(row[descriptionCol]);
    detailGrouped.set(detailKey, detail);
  }
  return {
    fileName,
    stores: [...stores].sort((a, b) => a.localeCompare(b, "it")),
    rows: [...grouped.values()].filter((row) => row.qty > 0),
    detailRows: [...detailGrouped.values()],
    ignoredPushRows
  };
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

function parseCatalog(arrayBuffer, fileName) {
  const matrix = firstSheetMatrix(arrayBuffer, fileName);
  const headerIndex = matrix.findIndex((row) => {
    const headers = row.map((value) => text(value).toLocaleUpperCase("it"));
    return headers.includes("CODICE") && headers.includes("TG") && headers.includes("BARCODE");
  });
  if (headerIndex < 0) throw new Error("Nell’anagrafica servono le colonne CODICE, DESCRIPTION, TG e BARCODE.");
  const headers = matrix[headerIndex].map((value) => text(value).toLocaleUpperCase("it"));
  const codeCol = headers.indexOf("CODICE");
  const descriptionCol = headers.indexOf("DESCRIPTION");
  const sizeCol = headers.indexOf("TG");
  const barcodeCol = headers.indexOf("BARCODE");
  if ([codeCol, descriptionCol, sizeCol, barcodeCol].some((index) => index < 0)) {
    throw new Error("Nell’anagrafica servono le colonne CODICE, DESCRIPTION, TG e BARCODE.");
  }

  const skuMap = new Map();
  let sourceRows = 0;
  let ignoredMissingBarcode = 0;
  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex];
    const originalSku = text(row[codeCol]);
    if (!originalSku) continue;
    sourceRows += 1;
    const skuKey = catalogSkuKey(originalSku);
    const item = skuMap.get(skuKey) || { sku: originalSku, description: "", children: [], sourceChildCount: 0, sourceChildCountBySize: new Map() };
    const description = text(row[descriptionCol]);
    if (!item.description && description) item.description = description;
    const size = exportSize(row[sizeCol]);
    const sizeKey = normalizeSizeKey(row[sizeCol]);
    item.sourceChildCount += 1;
    item.sourceChildCountBySize.set(sizeKey, (item.sourceChildCountBySize.get(sizeKey) || 0) + 1);
    const barcode = normalizeBarcode(row[barcodeCol]);
    if (!barcode) {
      ignoredMissingBarcode += 1;
      skuMap.set(skuKey, item);
      continue;
    }
    item.children.push({
      size,
      sizeKey,
      barcode
    });
    skuMap.set(skuKey, item);
  }
  if (!skuMap.size) throw new Error("Nessuna SKU leggibile trovata nell’anagrafica.");
  const skusWithoutExportableChildren = [...skuMap.values()].filter((item) => item.children.length === 0).length;
  return { fileName, skuMap, sourceRows, ignoredMissingBarcode, skusWithoutExportableChildren };
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
    setStatus("info", `Carica ${missing}`, "Il confronto parte con questi due file; l’anagrafica è necessaria per scaricare l’Excel completo.");
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
  const catalogNote = state.catalog
    ? `Anagrafica pronta: ${state.catalog.skuMap.size.toLocaleString("it-IT")} SKU; ${state.catalog.ignoredMissingBarcode.toLocaleString("it-IT")} righe senza barcode escluse; ${state.catalog.skusWithoutExportableChildren.toLocaleString("it-IT")} SKU con sola riga padre.`
    : "Carica anche il file anagrafica per abilitare il download Excel.";
  if (ambiguous.length) {
    const examples = ambiguous.slice(0, 3).map((item) => item.sku).join(", ");
    setStatus("warning", `${ambiguous.length} SKU con match ambiguo`, `Separate per prudenza nella vista “Match ambigui”: ${examples}${ambiguous.length > 3 ? "…" : ""}. ${catalogNote}`);
  } else {
    setStatus("success", "Confronto completato", `${matched.toLocaleString("it-IT")} SKU riconosciute nel centrale; nessun match ambiguo. ${catalogNote}`);
  }
  const now = new Date();
  ui.updatedMetric.textContent = new Intl.DateTimeFormat("it-IT", { dateStyle: "short", timeStyle: "short" }).format(now);
  ui.fileMetric.textContent = [state.excel.fileName, state.central.fileName, state.catalog?.fileName].filter(Boolean).join(" + ");
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
  if (!state.catalog) {
    setStatus("error", "Anagrafica non caricata", "Carica il file anagrafica prima di scaricare l’Excel.");
    return;
  }
  const rows = filteredRows();
  if (!rows.length) return;
  try {
    const skuOrder = [...new Set(rows.map((row) => row.catalogKey))];
    const missing = skuOrder.filter((catalogKey) => !state.catalog.skuMap.has(catalogKey));
    if (missing.length) {
      const labels = missing.map((catalogKey) => rows.find((row) => row.catalogKey === catalogKey)?.sku || catalogKey);
      const visibleLabels = labels.slice(0, 25).join(", ");
      const report = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(report, XLSX.utils.aoa_to_sheet([["SKU non trovata"], ...labels.map((sku) => [sku])]), "SKU mancanti");
      XLSX.writeFile(report, "sku-non-trovate-anagrafica.xlsx", { compression: true });
      setStatus("error", `${missing.length.toLocaleString("it-IT")} SKU non trovate nell’anagrafica`, `${visibleLabels}${labels.length > 25 ? ` e altre ${labels.length - 25}` : ""}. Ho scaricato il report completo delle SKU mancanti.`);
      return;
    }

    const selected = new Set(rows.map((row) => `${row.store}\u0000${row.catalogKey}`));
    const inventoryBySelection = new Map();
    for (const detail of state.excel.detailRows) {
      const key = `${detail.store}\u0000${detail.catalogKey}`;
      if (!selected.has(key)) continue;
      if (!inventoryBySelection.has(key)) inventoryBySelection.set(key, []);
      inventoryBySelection.get(key).push(detail);
    }

    const parentsBySku = new Map();
    for (const row of rows) {
      if (!parentsBySku.has(row.catalogKey)) parentsBySku.set(row.catalogKey, []);
      parentsBySku.get(row.catalogKey).push(row);
    }
    const expectedRows = skuOrder.reduce((total, catalogKey) => {
      const childCount = state.catalog.skuMap.get(catalogKey).children.length;
      const storeCount = parentsBySku.get(catalogKey).length;
      return total + (childCount * storeCount) + 1;
    }, 0);
    if (expectedRows > MAX_EXCEL_DATA_ROWS) {
      setStatus("error", "Troppi risultati per un singolo Excel", `${expectedRows.toLocaleString("it-IT")} righe superano il limite di Excel. Restringi i filtri per negozio, brand, SKU o giacenza minima.`);
      return;
    }
    const data = [];
    for (const catalogKey of skuOrder) {
      const catalogItem = state.catalog.skuMap.get(catalogKey);
      const children = [...catalogItem.children].sort((a, b) =>
        compareSizes(a.size, b.size) ||
        a.barcode.localeCompare(b.barcode, "it", { numeric: true })
      );
      const parents = parentsBySku.get(catalogKey).sort((a, b) => a.store.localeCompare(b.store, "it"));
      const lookups = new Map();
      for (const parent of parents) {
        const inventory = inventoryBySelection.get(`${parent.store}\u0000${parent.catalogKey}`) || [];
        const byBarcode = new Map();
        const bySize = new Map();
        let unclassifiedQuantity = 0;
        for (const detail of inventory) {
          const barcode = normalizeBarcode(detail.barcode);
          const sizeKey = normalizeSizeKey(detail.size);
          if (sizeKey) bySize.set(sizeKey, (bySize.get(sizeKey) || 0) + detail.qty);
          else if (!barcode) unclassifiedQuantity += detail.qty;
          if (barcode) byBarcode.set(barcode, (byBarcode.get(barcode) || 0) + detail.qty);
        }
        lookups.set(parent.store, { byBarcode, bySize, unclassifiedQuantity });
      }

      for (const child of children) {
        for (const parent of parents) {
          const lookup = lookups.get(parent.store);
          let quantity = 0;
          if (child.sizeKey && catalogItem.sourceChildCountBySize.get(child.sizeKey) === 1) {
            quantity = lookup.bySize.get(child.sizeKey) || 0;
          } else if (child.barcode) {
            quantity = lookup.byBarcode.get(child.barcode) || 0;
          }
          if (catalogItem.sourceChildCount === 1) quantity += lookup.unclassifiedQuantity;
          data.push({
            "SKU esploso con i figli": parent.sku,
            taglia: child.size,
            barcode: child.barcode,
            giacenza: quantity,
            negozio: parent.store,
            descrizione: catalogItem.description || parent.description
          });
        }
      }
      const parentSku = parents[0]?.sku || catalogKey;
      data.push({ "SKU esploso con i figli": parentSku, taglia: "", barcode: "", giacenza: "", negozio: "", descrizione: "" });
    }
    const sheet = XLSX.utils.json_to_sheet(data);
    sheet["!autofilter"] = { ref: sheet["!ref"] };
    sheet["!cols"] = [
      { wch: 30 }, { wch: 14 }, { wch: 18 }, { wch: 12 }, { wch: 42 }, { wch: 52 }
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
  ui.exportButton.disabled = !ready || !state.catalog || rows.length === 0;
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
ui.catalogInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  ui.catalogFileState.textContent = "Lettura…";
  setStatus("loading", "Lettura anagrafica…", file.name);
  try {
    const catalog = parseCatalog(await file.arrayBuffer(), file.name);
    state.catalog = catalog;
    ui.catalogFileState.textContent = file.name;
    reconcile();
  } catch (error) {
    ui.catalogFileState.textContent = state.catalog?.fileName || "Da caricare";
    setStatus("error", "Anagrafica non valida", error.message);
    render();
  }
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
