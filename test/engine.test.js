'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { runConsolidation } = require('../src/engine');

function parseCsv(text) {
  const rows = []; let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (c === '"') quoted = false;
      else field += c;
    } else if (c === '"') quoted = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n') { row.push(field.replace(/\r$/, '')); rows.push(row); row = []; field = ''; }
    else field += c;
  }
  if (field.length || row.length) { row.push(field.replace(/\r$/, '')); rows.push(row); }
  return rows;
}
function makeOrder(name, variation, sku, id) {
  const item = `[1] Product Name:${name}; Variation Name:${variation}; Price: RM 299.00; Quantity: 1; SKU Reference No.:${sku};`;
  return [['Tracking Number', 'Order ID', 'Item Data'], ['TRACK123', id || 'ORDER-00001', item]];
}
const inventory = parseCsv(fs.readFileSync(path.join(__dirname, 'fixtures', 'inventory.csv'), 'utf8'));

// Fujifilm film-only SKU must map to Xbox/Gaming Acc and expose its Item Code.
const fujiName = 'Fujifilm Film Instax Mini 10S (Single Pack)';
const fuji = runConsolidation({ orderRows: makeOrder(fujiName, 'Single Pack', fujiName), inventoryRows: inventory });
const fujiRow = fuji.rows.find(r => r[1] && String(r[1]).toLowerCase().includes('fujifilm'));
assert.ok(fujiRow, 'Fujifilm row should be output');
assert.equal(fujiRow[6], fujiName, 'Matched SKU column should show Inventory Item Code');
assert.equal(fujiRow.length, 9, 'Output should have nine columns including the preserved spacer');
assert.equal(fujiRow[7], '', 'Required spacer column must remain blank');
assert.ok(fuji.summary.some(r => r[0] === 'Xbox Gaming Acc'), 'Fujifilm film-only should be categorized as Xbox Gaming Acc');

// Scarlet selected base game: parent title may mention DLC, but the variation + SKU select base game.
const pokemon = runConsolidation({
  orderRows: makeOrder('NSW Nintendo Switch Pokemon Scarlet/Violet The Hidden Treasure of Area Zero 宝可梦: 朱/紫 零之秘宝 中英文版', 'Scarlet [US]', 'Switch Pokemon Scarlet', 'ORDER-00002'),
  inventoryRows: inventory
});
const pokemonRow = pokemon.rows.find(r => r[1] && /Pokemon Scarlet/i.test(r[1]));
assert.ok(pokemonRow, 'Pokemon Scarlet row should be output');
assert.ok(/Switch Pokemon Scarlet \[Us\]/i.test(pokemonRow[6]), `Expected base-game matched SKU, got: ${pokemonRow[6]}`);
assert.ok(!/Hidden Treasure|Area Zero/i.test(pokemonRow[6]), 'Matched SKU must not be the DLC bundle');
assert.equal(pokemonRow[7], '', 'Required spacer column must remain blank');

// An unknown SKU must remain blank in Matched SKU instead of being fabricated.
const unmatched = runConsolidation({
  orderRows: makeOrder('Mystery Product XYZ 9000', 'Standard', 'NO-SUCH-SKU-000', 'ORDER-00003'),
  inventoryRows: inventory
});
const unmatchedRow = unmatched.rows.find(r => r[1] && /Mystery Product/i.test(r[1]));
assert.ok(unmatchedRow, 'Unmatched product should remain in output');
assert.equal(unmatchedRow[6], '', 'Unmatched Matched SKU cell must remain blank');

console.log('PASS: Fujifilm category + matched SKU + spacer column');
console.log('PASS: Pokémon Scarlet base-game selection despite DLC wording in parent title');
console.log('PASS: unmatched product leaves Matched SKU blank');
