const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
vm.runInThisContext(fs.readFileSync('shared.js', 'utf8'));
vm.runInThisContext(fs.readFileSync('season-core.js', 'utf8'));
for (const letter of ['I', 'A', 'i']) assert.equal(bestStoreSeason(23, letter), 'FW23');
for (const letter of ['E', 'S', 's']) assert.equal(bestStoreSeason('2024', letter), 'SS24');
for (const [year, letter] of [['', 'I'], [23, 'X'], ['2x', 'E'], [null, 'A']]) assert.equal(bestStoreSeason(year, letter), null);
const buildSeasonTestResult = (rows, inventory) => buildSeasonResult(parseSeasonMagento([['SKU', 'Season Drop'], ...rows.map(row => [row.sku, row.old])]), inventory);
const excel = (sku, year, season) => ['', sku, '', '', year, season];
const coreExcluded = buildSeasonTestResult([
  { sku: 'COREITEM', old: 'CORE' }, { sku: 'COREITEM-M', old: ' core ' },
  { sku: 'COREITEM-L', old: 'SS22' }
], [excel('COREITEM', 23, 'I')]);
assert.equal(coreExcluded.excluded, 2);
assert.equal(coreExcluded.rows.length, 1);
assert.equal(coreExcluded.rows[0].expected, 'FW23');
assert.ok(!seasonExportCsv(coreExcluded).includes('"CORE"'));
const prefixed = buildSeasonTestResult([
  { sku: 'MAINITEM', old: 'MAIN-SS22' },
  { sku: 'PREITEM', old: 'PRE-FW23' },
  { sku: 'CHANGED', old: 'MAIN-SS21' }
], [excel('MAINITEM', 22, 'E'), excel('PREITEM', 23, 'I'), excel('CHANGED', 22, 'E')]);
assert.equal(prefixed.rows.find(row => row.sku === 'MAINITEM').status, 'matched');
assert.equal(prefixed.rows.find(row => row.sku === 'PREITEM').status, 'matched');
assert.equal(prefixed.rows.find(row => row.sku === 'CHANGED').status, 'mismatch');
assert.ok(seasonExportCsv(prefixed).includes('"CHANGED","SS22","MAIN-SS21"'));
const input = parseSeasonMagento([
  ['SKU', 'Season Drop'], ['OT-ABCD', 'FW22'], ['OT-ABCD-M', 'FW22'], ['OT-ABCD-S', 'FW23'],
  ['WXYZ', 'OUTLET-FW'], ['WXYZ-L', 'SS20'], ['CONFLICT', 'SS22'], ['INVALID', 'FW23'],
  ['MISSING', 'SS22'], ['0001', 'SS20'], ['QUOT"E', 'old,"value'], ['ABCD-RED', 'FW22'], ['ABCD-RED-L', 'FW22']
]);
const result = buildSeasonResult(input, [excel('PREA.B C D', 23, 'A'), excel('WXYZ', 24, 'S'),
  excel('CONFLICT', 22, 'E'), excel('CONFLICT', 23, 'I'), excel('INVALID', '', 'I'),
  excel('0001', 24, 'E'), excel('QUOT"E', 24, 'E'), excel('ABCD-RED', 25, 'I')]);
const get = sku => result.rows.find(row => row.sku === sku);
assert.equal(result.excluded, 1);
assert.equal(get('OT-ABCD').expected, 'FW23');
assert.equal(get('OT-ABCD-M').expected, 'FW23');
assert.equal(get('OT-ABCD-S').status, 'matched');
assert.equal(get('WXYZ-L').expected, 'SS24'); // excluded parent still supports non-outlet children
assert.equal(get('CONFLICT').reason, 'Stagioni BestStore discordanti');
assert.equal(get('INVALID').status, 'unresolved');
assert.equal(get('MISSING').status, 'unresolved');
assert.equal(get('ABCD-RED-L').expected, 'FW25'); // nearest parent
const exported = seasonExportCsv(result);
assert.ok(exported.startsWith('\uFEFF"SKU","SEASON BST","Season Drop"\r\n'));
assert.ok(exported.includes('"0001","SS24","SS20"'));
assert.ok(exported.includes('"QUOT""E","SS24","old,""value"'));
assert.ok(!exported.includes('OUTLET'));
assert.ok(!exported.includes('CONFLICT'));
assert.ok(exported.indexOf('"OT-ABCD-M"') < exported.indexOf('"OT-ABCD"'));
const ambiguousResult = buildSeasonTestResult([{sku:'AAAA',old:'SS20'},{sku:'BBBB',old:'SS20'}], [excel('XXAAAABBBB',23,'I')]);
assert.equal(ambiguousResult.ambiguousExcel, 1);
assert.ok(ambiguousResult.rows.every(row => row.status === 'unresolved'));
const duplicate = buildSeasonTestResult([{sku:'AAAA',old:'SS20'}], [excel('AAAA',23,'I'),excel('AAAA',23,'A')]);
assert.equal(duplicate.rows[0].expected, 'FW23');
assert.throws(() => parseSeasonMagento([['SKU','Wrong'],['ABC','SS23']]));
const wide = parseSeasonMagento([['SKU','Price','Season Drop','Brand'],['0001','€12.80','PRE-SS21','A, B'],['','€12.80','','']]);
const wideResult = buildSeasonResult(wide, [excel('0001',22,'E')]);
assert.equal(wide.rows.length,1);
assert.equal(seasonExportCsv(wideResult), '\uFEFF'+['"SKU","Price","Season Drop","SEASON BST","Brand"','"0001","€12.80","PRE-SS21","SS22","A, B"',''].join('\r\n'));
console.log('Season tests passed: conversions, shared matching, inheritance, outlets, conflicts, ambiguity, CSV quoting and Z–A ordering.');
if (process.argv[2]) {
  const fixture = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
  const actual = buildSeasonResult(parseSeasonMagento(fixture.magento), fixture.inventory);
  const counts = Object.fromEntries(['matched','mismatch','unresolved'].map(key => [key, actual.rows.filter(row => row.status === key).length]));
  assert.equal(actual.rows.length + actual.excluded, fixture.magento.slice(1).filter(row => row.some(v => text(v))).length);
  assert.ok(actual.rows.every(row => !/^(?:OUTLET-(?:SS|FW)|CORE)$/i.test(text(row.old))));
  assert.ok(actual.rows.filter(row => row.status === 'mismatch').every(row => /^(SS|FW)\d{2}$/.test(row.expected) && row.expected !== row.old));
  console.log(JSON.stringify({ ...counts, excluded: actual.excluded, unmatchedExcel:actual.unmatchedExcel, invalidExcel:actual.invalidExcel, ambiguousExcel:actual.ambiguousExcel, unresolvedReasons: actual.rows.filter(row => row.status === 'unresolved').reduce((counts,row) => ({...counts,[row.reason]:(counts[row.reason]||0)+1}),{}) }, null, 2));
  fs.writeFileSync('.site-scaffold/season-actual-summary.json', JSON.stringify(counts));
}
