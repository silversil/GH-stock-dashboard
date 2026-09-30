/* global text, normalizeSku, findCentralMatch */
function bestStoreSeason(year, season) {
  const prefix = { I: "FW", A: "FW", E: "SS", S: "SS" }[text(season).toUpperCase()];
  const value = text(year);
  return prefix && /^(?:\d{2}|20\d{2})$/.test(value) ? prefix + value.slice(-2) : null;
}

function parseSeasonMagento(matrix) {
  const headers = matrix[0] || [];
  const skuCol = headers.findIndex(v => /^sku$/i.test(text(v)));
  const seasonCol = headers.findIndex(v => /^season[ _]*drop$/i.test(text(v)));
  if (skuCol < 0 || seasonCol < 0) throw new Error("Nel CSV servono le colonne SKU e Season Drop (o season_drop).");
  if (headers.some(value => /^SEASON BST$/i.test(text(value)))) throw new Error("Il CSV contiene già SEASON BST. Carica il CSV Magento originale.");
  const rows = matrix.slice(1).filter(row => text(row[skuCol]));
  if (!rows.length) throw new Error("Il CSV non contiene articoli.");
  if (rows.some(row => row.length > headers.length)) throw new Error("Il CSV contiene righe con troppe colonne.");
  return { headers: [...headers], rows: rows.map(row => ({ sku: text(row[skuCol]), old: String(row[seasonCol] ?? ""), source: Array.from({ length: headers.length }, (_, col) => row[col] ?? "") })) };
}

function seasonParent(key, keys) {
  for (let end = key.lastIndexOf("-"); end > 0; end = key.lastIndexOf("-", end - 1)) {
    const parent = key.slice(0, end);
    if (end < key.length - 1 && keys.has(parent)) return parent;
  }
  return null;
}

function buildSeasonResult(magento, inventory) {
  const skuMap = new Map(magento.rows.map(row => [normalizeSku(row.sku, true), true]));
  const byLength = new Map();
  for (const key of skuMap.keys()) {
    if (!byLength.has(key.length)) byLength.set(key.length, new Set());
    byLength.get(key.length).add(key);
  }
  const central = { skuMap, byLength, lengths: [...byLength.keys()].sort((a, b) => b - a) };
  const cache = new Map(), matches = new Map(), ambiguous = new Set();
  let unmatchedExcel = 0, invalidExcel = 0, ambiguousExcel = 0;
  for (const row of inventory) {
    const sku = text(row[1]);
    if (!sku) continue;
    const key = normalizeSku(sku);
    if (!cache.has(key)) cache.set(key, findCentralMatch(key, central));
    const match = cache.get(key);
    if (!match) { unmatchedExcel++; continue; }
    if (!match.key) { ambiguousExcel++; match.ambiguous.forEach(k => ambiguous.add(k)); continue; }
    const item = matches.get(match.key) || { seasons: new Set(), skus: new Set(), invalid: false };
    const season = bestStoreSeason(row[4], row[5]);
    if (season) item.seasons.add(season);
    else { item.invalid = true; invalidExcel++; }
    item.skus.add(sku);
    matches.set(match.key, item);
  }
  let excluded = 0;
  const rows = magento.rows.flatMap(source => {
    if (/^(?:OUTLET-(?:SS|FW)|CORE)$/i.test(text(source.old))) { excluded++; return []; }
    const key = normalizeSku(source.sku, true);
    const parent = seasonParent(key, skuMap);
    // A direct Excel match takes precedence. Otherwise inherit the nearest known parent.
    let target = key;
    while (!matches.has(target) && !ambiguous.has(target)) {
      const next = seasonParent(target, skuMap);
      if (!next) break;
      target = next;
    }
    const item = matches.get(target);
    let reason = "";
    if (ambiguous.has(target)) reason = "Match SKU ambiguo";
    else if (!item) reason = "SKU non trovato in BestStore";
    else if (item.invalid) reason = "Anno o stagione BestStore non validi";
    else if (item.seasons.size !== 1) reason = "Stagioni BestStore discordanti";
    const expected = reason ? "" : [...item.seasons][0];
    const comparableSeason = text(source.old).toUpperCase().replace(/^(?:MAIN-|PRE-)+/, "");
    const status = reason ? "unresolved" : expected === comparableSeason ? "matched" : "mismatch";
    return [{ ...source, expected, status, reason, skus: [...(item?.skus || [])], inherited: target !== key, parent }];
  });
  rows.sort((a, b) => b.sku.localeCompare(a.sku, "it"));
  return { headers: magento.headers, rows, excluded, unmatchedExcel, invalidExcel, ambiguousExcel };
}

function seasonExportCsv(result) {
  const quote = value => '"' + String(value).replace(/"/g, '""') + '"';
  const insertAt = result.headers.length - 1;
  const insertSeason = (row, value) => [...row.slice(0, insertAt), value, ...row.slice(insertAt)];
  return '\uFEFF' + [insertSeason(result.headers, "SEASON BST"), ...result.rows.filter(row => row.status === "mismatch").map(row => insertSeason(row.source, row.expected))]
    .map(row => row.map(quote).join(",")).join("\r\n") + "\r\n";
}
