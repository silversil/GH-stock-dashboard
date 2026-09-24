const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
vm.runInThisContext(fs.readFileSync('shared.js', 'utf8'));
vm.runInThisContext(fs.readFileSync('pmp-core.js', 'utf8'));
assert.equal(parsePmp('1.234,56'), 1234.56);
assert.equal(parsePmp('€1,570.00'), 1570);
assert.equal(parsePmp('€3,258,222.71'), 3258222.71);
assert.equal(parsePmp('12.3456'), 12.3456);
assert.equal(parsePmp('0'), 0);
assert.equal(parsePmp(''), null);
assert.equal(parsePmp('errore'), null);
const magento = { headers: ['SKU', 'barcode', 'descrizione'], skuCol: 0, rows: [['OT-ABCD', '001234', '=test'], ['WXYZ', '00001', ''], ['NONE', '', ''], ['ABCD', '', 'duplicato'], ['', '', 'vuoto']] };
const inventory = { skuCol: 0, storeCol: 1, priceCol: 2, headerIndex: 0, rows: [['ABC.D', 'Roma', 10], ['XYZABCD', 'Milano', '20,00'], ['XXABCDZZ', 'Bari', 0], ['ABCD', 'PUSH', 999], ['ABCD', 'Roma', ''], ['ABCDWXYZ', 'Roma', 500], ['WXYZ', 'Roma', 7], ['UNKNOWN', 'Roma', 3]] };
const result = buildPmpResult(magento, inventory);
assert.deepEqual(result.rows.map(r => r.price), [10, 7, null, 10, null]);
assert.equal(result.rows[0].count, 3);
assert.equal(result.issues.length, 3);
assert.equal(result.push, 1);
assert.equal(result.issues[1][4], 'Match ambiguo');
const matrix = pmpExportMatrix(magento, result);
assert.equal(matrix.length, magento.rows.length + 1);
assert.deepEqual(matrix[0], [...magento.headers, 'PMP NEGOZI']);
for (let i = 0; i < magento.rows.length; i++) assert.deepEqual(matrix[i + 1].slice(0, -1), magento.rows[i]);
assert.equal(matrix[3][3], '');
// Original dashboard still delegates to the same matcher via its default state.
global.state = { central: { skuMap: new Map([['ABCD', true]]), byLength: new Map([[4, new Set(['ABCD'])]]), lengths: [4] } };
assert.equal(findCentralMatch('XYZABCD').method, 'prefix3');
assert.equal(findCentralMatch('ABCD').method, 'exact');
assert.equal(findCentralMatch('XXABCDZZ').method, 'contains');
console.log('PMP: parsing, matching, ambiguity, arithmetic mean, zero, PUSH, preservation and original matching passed.');
global.document = { getElementById: () => ({ value: 'all', addEventListener() {}, querySelector() { return {}; } }) };
vm.runInThisContext(fs.readFileSync('pmp.js', 'utf8'));
const grouped = parseSource([
  ['Report Semplificato - Solo Export Excel'], [],
  ['Cod', 'Descrizione', 'Esistenza', 'Costo Uni PM'],
  ['NS06', 'Negozio uno'], ['ABC.D', 'Articolo', 2, 10],
  ['OU04', 'Negozio due'], ['ABC.D', 'Articolo', 5, 20]
], 'inventory');
assert.equal(grouped.rows.length, 2);
assert.deepEqual(grouped.sourceRowNumbers, [5, 7]);
assert.equal(grouped.rows[1][grouped.storeCol], 'OU04 - Negozio due');
assert.equal(buildPmpResult(magento, grouped).rows[0].price, 15);
console.log('Grouped PMP report: store sections, source row numbers and cross-store mean passed.');
const typed = { headers: ['SKU', 'Product Type', 'Prezzo Acquisto'], skuCol: 0, rows: [
  ['ABCD', 'Configurable Product', '€9.00'],
  ['ABCD-L', 'Simple Product', '€8.00'],
  ['ABCD-RED', 'Configurable Product', '€7.00'],
  ['OT-ABCD-RED-37-5', 'Simple Product', '€6.00'],
  ['MISSING-S', 'Simple Product', '€5.00'],
  ['ABCDX-S', 'Simple Product', '€4.00']
] };
const typedInventory = { ...inventory, rows: [['XYZABCD', 'Roma', 10], ['XYZABCD', 'Bari', 20], ['XYZABCD-RED', 'Roma', 0]] };
const typedResult = buildPmpResult(typed, typedInventory);
const prices = new Map(typedResult.rows.map(row => [row.sku, row.price]));
assert.equal(prices.get('ABCD-L'), 15);
assert.equal(prices.get('ABCD'), 15);
assert.equal(prices.get('OT-ABCD-RED-37-5'), 0);
assert.equal(prices.get('MISSING-S'), null);
assert.equal(prices.get('ABCDX-S'), null);
assert.equal(typedResult.orphanSimpleCount, 2);
assert.deepEqual(typedResult.rows.map(row => row.kind), ['simple', 'simple', 'simple', 'configurable', 'simple', 'configurable']);
const typedExport = pmpExportMatrix(typed, typedResult);
assert.deepEqual(typedExport[0], [...typed.headers, 'PMP NEGOZI']);
for (const row of typedExport.slice(1)) { const source = typed.rows.find(source => source[0] === row[0]); assert.deepEqual(row.slice(0, -1), [source[0], source[1], parsePmp(source[2])]); }
console.log('Configurable inheritance, longest parent, orphan handling, zero price, original columns and type Z-A ordering passed.');

const money = pmpExportMatrix({headers:['SKU','Margine','C','D','E','F']}, {rows:[{source:['001','47.81%','€12.80','€1.234,56','','€0.00'],price:10}]});
assert.deepEqual(money[1], ['001','47.81%',12.8,1234.56,'',0,10]);
for (const parent of typedResult.rows.filter(r=>r.kind==='configurable')) { const children=typedResult.rows.filter(r=>r.parentKey===normalizeSku(parent.sku,true)); const end=typedResult.rows.indexOf(parent); assert.deepEqual(typedResult.rows.slice(end-children.length,end),children); }
