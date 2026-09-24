/* global normalizeSku, findCentralMatch, text, isPush */
function parsePmp(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  let source = text(value).replace(/[\s€]/g, "");
  if (!source) return null;
  if (/^[+-]?\d{1,3}(?:,\d{3})+\.\d+$/.test(source)) source = source.replace(/,/g, "");
  if (source.includes(",")) {
    if (!/^[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+),\d+$/.test(source)) return null;
    source = source.replace(/\./g, "").replace(",", ".");
  } else if (!/^[+-]?\d+(?:\.\d+)?$/.test(source)) return null;
  const number = Number(source);
  return Number.isFinite(number) ? number : null;
}

function productKind(value) {
  const label = text(value).toLowerCase();
  if (/^(configurable product|configurable|configurabile)$/.test(label)) return "configurable";
  if (/^(simple product|simple|semplice)$/.test(label)) return "simple";
  return "other";
}

function parentSkuKey(sku, parents) {
  const key = normalizeSku(sku, true);
  // Scan separator boundaries from right to left, preserving hyphens in parent SKUs.
  for (let end = key.lastIndexOf("-"); end > 0; end = key.lastIndexOf("-", end - 1)) {
    const candidate = key.slice(0, end);
    if (end < key.length - 1 && parents.has(candidate)) return candidate;
  }
  return null;
}

function buildPmpResult(magento, inventory) {
  const typeCol = magento.headers.findIndex((header) => /^product\s*type$/i.test(text(header)));
  const skuMap = new Map();
  const byLength = new Map();
  for (const row of magento.rows) {
    if (typeCol >= 0 && productKind(row[typeCol]) !== "configurable") continue;
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
    const kind = typeCol < 0 ? "configurable" : productKind(row[typeCol]);
    const key = kind === "simple" ? parentSkuKey(row[magento.skuCol], skuMap) : kind === "configurable" ? normalizeSku(row[magento.skuCol], true) : null;
    const item = totals.get(key);
    return { source: row, sku: text(row[magento.skuCol]), kind, parentKey: kind === "simple" ? key : null, price: item ? item.mean : null, count: item?.count || 0, stores: [...(item?.stores || [])], skus: [...(item?.skus || [])] };
  });
  if (typeCol >= 0) {
    const groupKey = (row) => row.parentKey || normalizeSku(row.sku, true);
    rows.sort((a, b) => groupKey(b).localeCompare(groupKey(a), "it") ||
      Number(a.kind === "configurable") - Number(b.kind === "configurable") ||
      b.sku.localeCompare(a.sku, "it"));
  }
  return { rows, issues, push, orphanSimpleCount: rows.filter((row) => row.kind === "simple" && !row.parentKey).length };
}

function pmpExportMatrix(magento, result) {
  return [[...magento.headers, "PMP NEGOZI"], ...result.rows.map((row) => [...Array.from({ length: magento.headers.length }, (_, i) => {
    const value = row.source[i] ?? "";
    return i >= 2 && i <= 5 ? parsePmp(value) ?? value : value;
  }), row.price ?? ""])];
}
