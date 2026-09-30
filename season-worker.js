/* Keep the large BestStore workbook and matching off the UI thread. */
importScripts("vendor/xlsx.full.min.js", "shared.js", "season-core.js");
const sources = { magento: null, inventory: null };
self.onmessage = event => {
  const { kind, buffer, name } = event.data;
  sources[kind] = null;
  try {
    if (kind === "magento") sources.magento = parseSeasonMagento(firstSheetMatrix(buffer, name, true));
    else {
      const book = XLSX.read(buffer, { type: "array", dense: true, sheets: 0 });
      const sheet = book.Sheets[book.SheetNames[0]];
      if (!sheet?.["!ref"]) throw new Error("Il primo foglio Excel è vuoto.");
      const end = XLSX.utils.decode_range(sheet["!ref"]).e.r;
      sources.inventory = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: false, range: { s: { r: 2, c: 0 }, e: { r: end, c: 5 } } });
      if (!sources.inventory.some(row => text(row[1]))) throw new Error("Nessuno SKU in colonna B dalla riga 3.");
    }
    const result = sources.magento && sources.inventory ? buildSeasonResult(sources.magento, sources.inventory) : null;
    self.postMessage({ kind, name, result });
  } catch (error) { sources[kind] = null; self.postMessage({ kind, error: error.message }); }
};
