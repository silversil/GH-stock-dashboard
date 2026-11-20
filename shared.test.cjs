const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const context = { state: { central: null } };
vm.createContext(context);
vm.runInContext(fs.readFileSync("shared.js", "utf8"), context);

function centralWith(...keys) {
  const skuMap = new Map(keys.map((key) => [key, { originals: new Set([key]), brands: new Set() }]));
  const byLength = new Map();
  for (const key of keys) {
    if (!byLength.has(key.length)) byLength.set(key.length, new Set());
    byLength.get(key.length).add(key);
  }
  return { skuMap, byLength, lengths: [...byLength.keys()].sort((a, b) => b - a) };
}

const bstSku = context.normalizeSku("NFC0A52ST . UBO");
const magentoSku = context.normalizeSku("NF0A52STUBO", true);

assert.equal(bstSku, "NFC0A52STUBO");
assert.deepEqual(
  JSON.parse(JSON.stringify(context.findCentralMatch(bstSku, centralWith(magentoSku)))),
  { key: "NF0A52STUBO", method: "brandAlias", ambiguous: [] }
);

assert.deepEqual(
  JSON.parse(JSON.stringify(context.findCentralMatch(bstSku, centralWith(bstSku, magentoSku)))),
  { key: "NFC0A52STUBO", method: "exact", ambiguous: [] }
);

assert.equal(context.findCentralMatch("NFCX1234", centralWith("NF0X1234")), null);

console.log("shared matching tests: ok");
