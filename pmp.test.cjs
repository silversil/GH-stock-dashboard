const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
vm.runInThisContext(fs.readFileSync('shared.js', 'utf8'));
vm.runInThisContext(fs.readFileSync('pmp-core.js', 'utf8'));
assert.equal(parsePmp('1.234,56'), 1234.56);
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
assert.deepEqual(matrix[0], [...magento.headers, 'prezzo_acquisto']);
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
