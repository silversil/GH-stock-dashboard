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

function pmpChildSize(row) {
  if (!row.parentKey) return "";
  // Use the original suffix: SKU normalization removes decimal points.
  for (let i = 0; i < row.sku.length; i++) {
    if (row.sku[i] === "-" && normalizeSku(row.sku.slice(0, i), true) === row.parentKey) return row.sku.slice(i + 1).trim();
  }
  return "";
}
function comparePmpSizes(left, right) {
  const describe = (value) => {
    let label = text(value).toUpperCase().replace(/\s/g, "");
    if (/^\d+-$/.test(label)) label = label.slice(0, -1) + ".5";
    if (/^\d+(?:[.,]\d+)?$/.test(label)) return [0, Number(label.replace(",", ".")), label];
    const aliases = { SM: "S", MD: "M", LG: "L", LXL: "L/XL" };
    label = aliases[label] || label;
    const base = { S: 0, "S/M": 0.5, M: 1, "M/L": 1.5, L: 2, "L/XL": 2.5 };
    if (Object.hasOwn(base, label)) return [1, base[label], label];
    const extended = label.match(/^(X+|\d+X)(S|L)$/);
    if (extended) {
      const count = /^\d/.test(extended[1]) ? Number(extended[1].slice(0, -1)) : extended[1].length;
      return [1, extended[2] === "S" ? -count : 2 + count, label];
    }
    return [2, 0, label];
  };
  const a = describe(left), b = describe(right);
  return a[0] - b[0] || a[1] - b[1] || a[2].localeCompare(b[2], "it", { numeric: true });
}

function buildPmpResult(magento, inventory, selectedStores = null) {
  const storeSelection = selectedStores === null ? null : new Set(selectedStores);
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
  const presentKeys = new Set();
  const matchedStores = new Map();
  const cache = new Map();
  const issues = [];
  let push = 0;
  inventory.rows.forEach((row, index) => {
    const sku = text(row[inventory.skuCol]);
    const store = inventory.storeCol < 0 ? "" : text(row[inventory.storeCol]);
    if (storeSelection && !storeSelection.has(store)) return;
    if (isPush(store)) { push++; return; }
    const key = normalizeSku(sku);
    if (!cache.has(key)) cache.set(key, findCentralMatch(key, central));
    const match = cache.get(key);
    if (match?.key) {
      presentKeys.add(match.key);
      if (!matchedStores.has(match.key)) matchedStores.set(match.key, new Set());
      if (store) matchedStores.get(match.key).add(store);
    }
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
  let rows = magento.rows.map((row) => {
    const kind = typeCol < 0 ? "configurable" : productKind(row[typeCol]);
    const key = kind === "simple" ? parentSkuKey(row[magento.skuCol], skuMap) : kind === "configurable" ? normalizeSku(row[magento.skuCol], true) : null;
    const item = totals.get(key);
    return { source: row, sku: text(row[magento.skuCol]), inSelectedStores: presentKeys.has(key), kind, parentKey: kind === "simple" ? key : null, price: item ? item.mean : null, count: item?.count || 0, stores: [...(matchedStores.get(key) || [])].sort((a, b) => a.localeCompare(b, "it")), skus: [...(item?.skus || [])] };
  });
  if (storeSelection) rows = rows.filter((row) => row.inSelectedStores);
  if (typeCol >= 0) {
    const groupKey = (row) => row.parentKey || normalizeSku(row.sku, true);
    rows.sort((a, b) => groupKey(b).localeCompare(groupKey(a), "it") ||
      Number(a.kind === "configurable") - Number(b.kind === "configurable") ||
      comparePmpSizes(pmpChildSize(a), pmpChildSize(b)) ||
      b.sku.localeCompare(a.sku, "it"));
  }
  return { rows, issues, push, orphanSimpleCount: rows.filter((row) => row.kind === "simple" && !row.parentKey).length };
}

function pmpExportMatrix(magento, result) {
  const brandCol = magento.headers.findIndex((header) => /^(brand|marca)$/i.test(text(header)));
  const columns = magento.headers.map((_, i) => i).filter((i) => i !== brandCol);
  return [[...columns.map((i) => magento.headers[i]), "PMP NEGOZI", "Brand", "Negozio"], ...result.rows.map((row) => [...columns.map((i) => {
    const value = row.source[i] ?? "";
    return i >= 2 && i <= 5 ? parsePmp(value) ?? value : value;
  }), row.price ?? "", brandCol < 0 ? "" : row.source[brandCol] ?? "", (row.stores || []).join(" / ")])];
}
