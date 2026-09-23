/* global normalizeSku, findCentralMatch, text, isPush */
function parsePmp(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  let source = text(value).replace(/[\s€]/g, "");
  if (!source) return null;
  if (source.includes(",")) {
    if (!/^[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+),\d+$/.test(source)) return null;
    source = source.replace(/\./g, "").replace(",", ".");
  } else if (!/^[+-]?\d+(?:\.\d+)?$/.test(source)) return null;
  const number = Number(source);
  return Number.isFinite(number) ? number : null;
}

function buildPmpResult(magento, inventory) {
  const skuMap = new Map();
  const byLength = new Map();
  for (const row of magento.rows) {
    const key = normalizeSku(row[magento.skuCol], true);
    if (key.length < 4) continue;
    skuMap.set(key, true);
    if (!byLength.has(key.length)) byLength.set(key.length, new Set());
    byLength.get(key.length).add(key);
  }
  const central = { skuMap, byLength, lengths: [...byLength.keys()].sort((a, b) => b - a) };
  const totals = new Map();
  const cache = new Map();
  const issues = [];
  let push = 0;
  inventory.rows.forEach((row, index) => {
    const sku = text(row[inventory.skuCol]);
    const store = inventory.storeCol < 0 ? "" : text(row[inventory.storeCol]);
    if (isPush(store)) { push++; return; }
    const key = normalizeSku(sku);
    if (!cache.has(key)) cache.set(key, findCentralMatch(key, central));
    const match = cache.get(key);
    const price = parsePmp(row[inventory.priceCol]);
    let reason = !sku ? "SKU vuoto" : !match ? "Nessun match" : !match.key ? "Match ambiguo" : price === null ? "PMP mancante o non numerico" : "";
    if (reason) {
      issues.push([inventory.sourceRowNumbers?.[index] ?? index + inventory.headerIndex + 2, sku, store, row[inventory.priceCol] ?? "", reason, match?.ambiguous?.join(" / ") || ""]);
      return;
    }
    const item = totals.get(match.key) || { mean: 0, count: 0, skus: new Set(), stores: new Set() };
    item.count++;
    item.mean = item.mean * ((item.count - 1) / item.count) + price / item.count;
    item.skus.add(sku);
    if (store) item.stores.add(store);
    totals.set(match.key, item);
  });
  const rows = magento.rows.map((row) => {
    const item = totals.get(normalizeSku(row[magento.skuCol], true));
    return { source: row, sku: text(row[magento.skuCol]), price: item ? item.mean : null, count: item?.count || 0, stores: [...(item?.stores || [])], skus: [...(item?.skus || [])] };
  });
  return { rows, issues, push };
}

function pmpExportMatrix(magento, result) {
  return [[...magento.headers, "prezzo_acquisto"], ...result.rows.map((row) => [...Array.from({ length: magento.headers.length }, (_, i) => row.source[i] ?? ""), row.price ?? ""])];
}
