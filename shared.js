/* Shared parsing and SKU matching for both dashboards. */
function text(value) { return String(value ?? "").trim(); }

function isPush(value) { return /(^|[^A-Z])PUSH(?:[^A-Z]|\d|$)/i.test(text(value)); }

function normalizeSku(value, removeOt = false) {
  let normalized = text(value).replace(/\u00a0/g, " ").toLocaleUpperCase("it").replace(/[\s.]/g, "");
  if (removeOt) normalized = normalized.replace(/OT-/g, "");
  return normalized;
}

function escapeHtml(value) {
  return text(value).replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char]));
}

function findColumn(headers, tests) {
  return headers.find((header) => tests.some((test) => test.test(text(header)))) || null;
}

function firstSheetMatrix(arrayBuffer, fileName, keepBlankRows = false) {
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
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true, blankrows: keepBlankRows });
  if (!rows.length) throw new Error("Il file non contiene dati.");
  return rows;
}

function findCentralMatch(excelSkuKey, central = state.central) {
  const { skuMap, byLength, lengths } = central;
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
