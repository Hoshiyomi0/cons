/**
 * Shopee Order Consolidation
 * =========================================================================
 * Reads a raw Shopee export (either the classic order export or the newer
 * "Packing List" format — both auto-detected) and produces a Consolidated
 * sheet of clean product names grouped by business category, with no orange review highlighting or completion popup; blue same-order highlighting is retained.
 *
 * ARCHITECTURE
 *   1. PARSING          raw sheet rows -> line items
 *   2. NAME RESOLUTION  raw product text -> clean display name
 *                       (normalize, ordered rule table, generic fallback)
 *   3. CATEGORIZATION   clean name -> category
 *                       (explicit rules, then inference, then review flag)
 *   4. OUTPUT           sorted sheets with blue same-order highlighting only
 *
 */

// =========================================================================
// CONFIGURATION
// =========================================================================
var RAW_SHEET_NAME = 'Raw';
var SKU_MASTER_SHEET_NAME = 'SKU_MASTER';
var PRODUCT_DB_SHEET_NAME = 'PRODUCT_DB';
var OUTPUT_SHEET_NAME = 'Consolidated';
var SUMMARY_SHEET_NAME = 'Summary';
// Stores category corrections you make, so they're applied automatically on
// future runs. Two columns: a keyword/brand phrase, and the category it
// should map to. Written by learnFromCorrections(), read on every run.
var LEARNED_RULES_SHEET_NAME = 'Learned_Rules';

// =========================================================================
// MENU
// =========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Shopee Tools')
    .addItem('Consolidate Orders', 'consolidateOrders')
    .addItem('Learn From Corrections', 'learnFromCorrections')
    .addItem('Prepare Product Database Sheet', 'prepareProductDatabaseSheet')
    .addToUi();
}

// =========================================================================
// SKU UTILITIES
// =========================================================================
function normalizeSku(sku) {
  sku = String(sku || '').toUpperCase().trim();
  var m = sku.match(/\b\d{3}-\d{6}\b/);
  if (m) return m[0];
  m = sku.match(/\b[A-Z0-9]+(?:-[A-Z0-9]+)+\b/);
  if (m) return m[0];
  return sku.replace(/^#/, '').trim();
}

function buildSkuMap() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SKU_MASTER_SHEET_NAME);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var sku = normalizeSku(data[i][0]);
    var product = String(data[i][1] || '').trim();
    if (!sku || !product) continue;
    if (!map[sku]) {
      map[sku] = product;
      continue;
    }
    // Duplicate SKU already seen — don't just let the later row silently win.
    // Prefer whichever version actually includes the SKU code in its name
    // (the more complete/correct entry); if both or neither do, keep the
    // longer one as the more descriptive version.
    var existingHasCode = map[sku].toUpperCase().indexOf(sku) !== -1;
    var newHasCode = product.toUpperCase().indexOf(sku) !== -1;
    if (newHasCode && !existingHasCode) {
      map[sku] = product;
    } else if (newHasCode === existingHasCode && product.length > map[sku].length) {
      map[sku] = product;
    }
  }
  return map;
}


// =========================================================================
// INVENTORY PRODUCT DATABASE
// =========================================================================
// Import the inventory CSV into a sheet named PRODUCT_DB with these columns:
// A: SKU / Item Code, B: Product Name, C: Category (ID or Name), D: Tags.
// The database is indexed once per consolidation run, not once per order line.

function prepareProductDatabaseSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRODUCT_DB_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(PRODUCT_DB_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([['SKU / Item Code', 'Product Name', 'Category (ID or Name)', 'Tags']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 460);
    sheet.setColumnWidth(3, 240);
    sheet.setColumnWidth(4, 220);
  }
  ss.setActiveSheet(sheet);
  ss.toast('Import the inventory CSV into this tab, keeping the four source columns and header row.', 'Product database', 8);
}

function normalizeDbKey(value) {
  return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function normalizeDbName(value) {
  return String(value || '').toLowerCase()
    .replace(/[\u4e00-\u9fff]+/g, ' ')
    .replace(/\bNSW\b/gi, ' ')
    .replace(/\[[^\]]*\]/g, ' ')
    .replace(/\([^)]*years? warranty[^)]*\)/gi, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

function inventoryTokens(value) {
  var stop = /^(the|and|for|with|from|this|that|set|edition|version|new|official|original|product|item|bundle|pack|box|booster|mini|super|game|games|accessory|accessories|accessories|english|eng|chinese|chi|japanese|jap|asia|global|ps4|ps5|switch|switch2|nintendo|xbox|playstation|pc|steam|wireless|controller|card|cards|tcg)$/i;
  var words = normalizeDbName(value).split(/\s+/);
  var out = [];
  var seen = {};
  words.forEach(function (w) {
    // Normalize the listing/inventory variant "Celebrations" vs "Celebration".
    if (w === 'celebrations') w = 'celebration';
    if (w.length < 3 || stop.test(w) || seen[w]) return;
    seen[w] = true;
    out.push(w);
  });
  return out;
}

function extractInventoryCodes(value) {
  var text = String(value || '').toUpperCase();
  var found = {};
  // Numeric catalogue codes such as ZD Toys 1907-04 / 1909-50 must be
  // indexed as exact model identifiers too. Without this pattern, a broad
  // parent listing containing many Iron Man marks can incorrectly win on
  // shared title tokens (e.g. Mk30 instead of the selected Mark 4 / Mark 50).
  var patterns = [/\b\d{3,5}-\d{2,5}\b/g, /\b[A-Z]{1,8}-[A-Z0-9]{2,10}\b/g, /\b[A-Z]{1,6}[0-9]{2,10}[A-Z]{0,4}\b/g, /\b[0-9]{5,8}\b/g];
  patterns.forEach(function (re) {
    var m;
    while ((m = re.exec(text)) !== null) {
      if (!/\d/.test(m[0])) continue;
      var code = m[0].replace(/[^A-Z0-9]/g, '');
      if (/^(PS4|PS5|XBOX|SWITCH|2026|2027)$/.test(code)) continue;
      found[code] = true;
    }
  });
  return Object.keys(found);
}

function buildInventoryDbIndex() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRODUCT_DB_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, Math.min(4, sheet.getLastColumn())).getValues();
  var db = { available: true, records: [], sku: {}, name: {}, code: {}, token: {}, tokenFreq: {} };

  data.forEach(function (row, offset) {
    var sku = String(row[0] || '').trim();
    var name = String(row[1] || '').trim();
    var category = String(row[2] || '').trim();
    var tags = String(row[3] || '').trim();
    if (!sku && !name) return;

    var rec = { row: offset + 2, sku: sku, name: name, category: category, tags: tags,
      skuKey: normalizeDbKey(sku), nameKey: normalizeDbName(name), tokens: inventoryTokens(name + ' ' + sku + ' ' + tags), codes: extractInventoryCodes(sku + ' ' + name) };
    var index = db.records.length;
    db.records.push(rec);

    if (rec.skuKey) { if (!db.sku[rec.skuKey]) db.sku[rec.skuKey] = []; db.sku[rec.skuKey].push(index); }
    if (rec.nameKey) { if (!db.name[rec.nameKey]) db.name[rec.nameKey] = []; db.name[rec.nameKey].push(index); }
    rec.codes.forEach(function (code) { if (!db.code[code]) db.code[code] = []; db.code[code].push(index); });
    rec.tokens.forEach(function (token) { db.tokenFreq[token] = (db.tokenFreq[token] || 0) + 1; });
  });

  // Keep only reasonably selective tokens in the inverted index. Very common
  // words create huge candidate lists and do not help distinguish products.
  db.records.forEach(function (rec, index) {
    rec.tokens.forEach(function (token) {
      if (db.tokenFreq[token] > 1800) return;
      if (!db.token[token]) db.token[token] = [];
      db.token[token].push(index);
    });
  });
  return db;
}

function inventorySimilarity(query, rec) {
  var q = inventoryTokens(query);
  if (!q.length || !rec.tokens.length) return 0;
  var rset = {};
  rec.tokens.forEach(function (t) { rset[t] = true; });
  var overlap = 0;
  q.forEach(function (t) { if (rset[t]) overlap++; });
  if (!overlap) return 0;
  var coverage = overlap / Math.min(q.length, rec.tokens.length);
  var jaccard = overlap / (q.length + rec.tokens.length - overlap);
  return Math.min(1, coverage * 0.72 + jaccard * 0.28);
}

function inventoryVariationConflicts(item, rec) {
  var variation = String(item.variation || '').split(',').map(function (x) { return x.trim(); })
    .filter(function (x) { return x && !/^(FULL PAYMENT|BALANCE PAYMENT|DEPOSIT|PARTIAL PAYMENT|P\.?O\.?|READY STOCK|RANDOM COVER|NO ADD-?ON)$/i.test(x); })
    .join(' ');
  var recordText = normalizeDbName(rec.name + ' ' + rec.sku);
  // ZD Toys parent listings contain many Iron Man model numbers. The selected
  // variation (Mark 4, Mark 50, etc.) must agree with the inventory record.
  if (/\bZD\s*Toys\b/i.test(item.name || '')) {
    var selectedZdMark = (variation.match(/\bMark\s*(\d{1,2})\b/i) || variation.match(/\bMk\s*(\d{1,2})\b/i) || [])[1];
    var recordZdMark = ((rec.name + ' ' + rec.sku).match(/\bMk\s*(\d{1,2})\b/i) || [])[1];
    if (selectedZdMark && recordZdMark && Number(selectedZdMark) !== Number(recordZdMark)) return true;
  }
  var selectedHardwareModel = (variation.match(/\b(G29|G920|G923)\b/i) || [])[1];
  if (selectedHardwareModel && !new RegExp('\\b' + selectedHardwareModel + '\\b', 'i').test(rec.name + ' ' + rec.sku)) return true;
  // Beyblade X option codes are the selected model identity. A database
  // listing for CX-01/UX-01 must not match when the buyer selected CX-12/UX-03.
  var selectedBeybladeCode = (variation.match(/\b(?:BX|CX|UX)-\d{2,3}\b/i) || [])[0];
  var dbBeybladeCode = ((rec.name + ' ' + rec.sku).match(/\b(?:BX|CX|UX)-\d{2,3}\b/i) || [])[0];
  if (selectedBeybladeCode && dbBeybladeCode &&
      selectedBeybladeCode.toUpperCase() !== dbBeybladeCode.toUpperCase()) return true;
  // A bundle selection that includes a shifter must not match a wheel-only DB item.
  if (/\bSHIFTER\b/i.test(variation) && !/\bSHIFTER\b/i.test(recordText)) return true;
  var important = inventoryTokens(variation).filter(function (t) {
    return /^(collector|deluxe|limited|special|ultimate|gold|silver|black|white|red|blue|green|purple|grey|gray|pink|clear|transparent|orange|yellow|brown|beige|navy|teal|crimson|rose|anniversary|rathalos|dran|brave|buster)$/.test(t);
  });
  if (!important.length) return false;
  return important.some(function (token) { return recordText.split(/\s+/).indexOf(token) === -1; });
}

function isVideoGameLineItem(item) {
  var name = String(item && item.name || '').trim();
  var variation = String(item && item.variation || '');
  if (/^JB Branch Listing\b/i.test(name) && /\b(?:NSW|Nintendo Switch|Switch)\b/i.test(name)) return true;
  if (!/^(?:Used\s+)?(?:PS[345]|PS4\/PS5|PlayStation|NSW2?|Nintendo\s+Switch|Switch2?|Switch\s+2|Xbox)\b/i.test(name)) return false;
  if (/\b(?:keyboard|mouse|mousepad|headset|earbuds?|controller|joy-?con|charging\s+dock|charging\s+stand|console|accessor(?:y|ies)|protective\s+case|carrying\s+pouch|screen\s+protector|charger|dock|grip|thumb\s+grip|steering\s+wheel)\b/i.test(name)) return false;
  if (/^JB Branch Listing\b/i.test(name)) return true;
  return true;
}

function inventoryGamePlatform(item) {
  var name = String(item && item.name || '');
  var variation = String(item && item.variation || '');
  if (/^JB Branch Listing\b/i.test(name) && /\b(?:NSW|Nintendo Switch|Switch)\b/i.test(name)) return 'Switch';
  // A selected platform in the variation overrides a mixed-platform parent listing.
  if (/\bSwitch\s*1\b|\bSwitch1\b|\bSW1\b/i.test(variation)) return 'Switch';
  if (/\bSwitch\s*2\b|\bSwitch2\b|\bSW2\b/i.test(variation)) return 'Switch2';
  if (/\bPS5\b|PlayStation\s*5/i.test(variation)) return 'PS5';
  if (/\bPS4\b|PlayStation\s*4/i.test(variation)) return 'PS4';
  if (/^(?:Used\s+)?(?:NSW2|Switch2|Switch\s*2|Nintendo\s+Switch\s*2)\b/i.test(name)) return 'Switch2';
  if (/^(?:Used\s+)?(?:NSW|Nintendo\s+Switch|Switch)\b/i.test(name)) return 'Switch';
  var psMatch = name.match(/^(?:Used\s+)?PlayStation\s+PS([345])\b/i);
  if (psMatch) return 'PS' + psMatch[1];
  if (/^(?:Used\s+)?(?:PS5|PlayStation\s*5)\b/i.test(name)) return 'PS5';
  if (/^(?:Used\s+)?(?:PS4|PlayStation\s*4)\b/i.test(name)) return 'PS4';
  if (/\bXbox\b/i.test(name)) return 'Xbox';
  if (/\bPC\b/i.test(name)) return 'PC';
  return '';
}

function inventoryRecordGamePlatform(rec) {
  var category = String(rec && rec.category || '').trim();
  var name = String(rec && rec.name || '') + ' ' + String(rec && rec.sku || '');
  if (/^(?:Used\s+)?(?:Ps5\s*Games|Ps5 Games)|^Used Games \[Ps5\]/i.test(category) || /^Used\s+Ps5\b|^Ps5\b/i.test(name)) return 'PS5';
  if (/^(?:Used\s+)?(?:Ps4\s*Games|Ps4 Games)|^Used Games \[Ps4\]/i.test(category) || /^Used\s+Ps4\b|^Ps4\b/i.test(name)) return 'PS4';
  if (/^(?:Switch2\s*Games)|^Used Games \[Switch2\]/i.test(category) || /^Used\s+Switch2?\b|^Switch2?\b/i.test(name)) {
    return /^Switch2\b|^Used Switch2\b/i.test(name) || /Switch2/i.test(category) ? 'Switch2' : 'Switch';
  }
  if (/^Switch\s*Games|^Used Games \[Switch\]/i.test(category) || /^Used Switch\b|^Switch\b/i.test(name)) return 'Switch';
  if (/^Xbox\s*Series\s*Games/i.test(category) || /^Xbox\b/i.test(name)) return 'Xbox';
  if (/^PC\s*Game$/i.test(category) || /^PC\b/i.test(name)) return 'PC';
  return '';
}

function inventoryRecordIsUsedGame(rec) {
  return /^Used Games\b/i.test(String(rec && rec.category || '')) ||
    /^Used\s+(?:Ps[345]|Switch2?|Xbox)\b/i.test(String(rec && rec.name || ''));
}

function inventoryItemIsUsedGame(item) {
  return /^Used\s+(?:Ps[345]|Switch2?|Xbox)\b/i.test(String(item && item.name || '')) ||
    /\bPre[- ]?Owned\b|\bUsed\s+Game\b/i.test(String(item && item.name || '') + ' ' + String(item && item.variation || ''));
}

function inventoryGameRegion(value) {
  var text = String(value || '')
    .replace(/\bHell\s+Is\s+Us\b/ig, 'Hell Is')
    .replace(/\bAmong\s+Us\b/ig, 'Among');
  var bracketed = text.match(/\[\s*(R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\s*\]/i);
  if (bracketed) return bracketed[1].toUpperCase();
  // Also support variation formats such as "R3 - Chinese + English" or "US - English".
  var marked = text.match(/\b(R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\b(?=\s*(?:[-,]|JP\b|$))/i);
  if (marked) return marked[1].toUpperCase();
  // Region suffixes may be unbracketed in some raw exports.
  var trailing = text.match(/\b(R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\b\s*$/i);
  return trailing ? trailing[1].toUpperCase() : '';
}

function inventoryGameLanguage(value) {
  var text = String(value || '');
  // Cover descriptors are not language selections; use the product title/SKU instead.
  if (/\b(?:JAP|JAPANESE)\s*\/\s*ENG\b|\b(?:ENGLISH|CHI|CHINESE)\s+COVER\b/i.test(text)) return '';
  if (/CHI\s*\/\s*ENG|ENG\s*\/\s*CHI|CHINESE\s*(?:\+|AND|\/)\s*ENGLISH|中文\s*\+\s*ENGLISH|中文\s*\+\s*ENG\b/i.test(text)) return 'Chi/Eng';
  if (/\bENG?(?:LISH)?\b/i.test(text)) return 'Eng';
  if (/\bCHI(?:NESE)?\b|中文版/i.test(text)) return 'Chi';
  return '';
}

function inventoryRequestedGameLanguage(item) {
  var variation = String(item && item.variation || '');
  var sku = String(item && item.sku || '');
  if (/\bSW[12]\b/i.test(variation) && /\bCHI\s+COVER\b/i.test(sku)) {
    return inventoryGameLanguage(sku) || inventoryGameLanguage(item && item.name);
  }
  return inventoryGameLanguage(variation) || inventoryGameLanguage(item && item.name);
}

function stripGameRegionMarker(value) {
  return String(value || '')
    .replace(/\s+\[(?:R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\]\s*/ig, ' ')
    .replace(/\s+\b(?:R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\b\s*$/ig, ' ')
    .replace(/\s+/g, ' ').trim();
}

function inventoryGameNumberSignature(value) {
  var text = String(value || '')
    .replace(/[\u4e00-\u9fff]+/g, ' ')
    .replace(/\bPS[345]\b|\bSwitch\s*2?\b|\bNSW2?\b|\bXbox\b/ig, ' ')
    .replace(/\bR[0-9]\b|\b(?:US|EU|AS|ASIA|UK|JP|AU|AUS)\b/ig, ' ')
    .replace(/\b(?:Chi\/Eng|Eng\/Chi|English|Chinese|Japanese|Version|Edition)\b/ig, ' ');
  var nums = text.match(/\b\d{1,3}\b/g) || [];
  var seen = {};
  return nums.filter(function (n) {
    if (/^(?:1|2)$/.test(n) || /^20\d\d$/.test(n) || seen[n]) return false;
    seen[n] = true;
    return true;
  }).join(',');
}

function inventoryGameCoreTokens(value) {
  var text = String(value || '')
    .replace(/\bNintendo\s+Switch\s+2\b/ig, ' ')
    .replace(/\bSwitch\s*[12]\b/ig, ' ')
    .replace(/\b(?:Full Payment|Balance Payment|Deposit|P\.?O\.?|Game Only|Game \+ Goodie Bag)\b/ig, ' ')
    .replace(/\[[^\]]*\]/g, ' ')
    .replace(/\bPlayStation\s+PS[345]\b/ig, ' ')
    .replace(/\b(?:PS[345]|PS4\/PS5|PlayStation\s*[345]|NSW2?|Nintendo\s+Switch|Switch2?|Switch\s*2|Xbox|PC)\b/ig, ' ')
    .replace(/\b(?:Chi\/Eng|Eng\/Chi|Chinese\/English|English|Chinese|Japanese|Eng|Chi|JAP|Version|Edition|Standard|Deluxe|Complete|Limited|Collector'?s?|Definitive|Ultimate|Porsche|Track|Remastered)\b/ig, ' ')
    .replace(/DestroyAllHumans/ig, 'Destroy All Humans')
    .replace(/LoverPretend/ig, 'Lover Pretend')
    .replace(/MegaMan/ig, 'Mega Man')
    .replace(/Fighters\b/ig, 'Fighter')
    .replace(/\b23\s+2023\b/ig, '2023')
    .replace(/[\u4e00-\u9fff]+/g, ' ');
  var words = normalizeDbName(text).split(/\s+/).filter(function (w) {
    return w && !/^(the|and|for|from|of|a|an|to|in|on|with|game|games|version|edition|new|us|eu|as|asia|uk|jp|aus|r[0-9])$/i.test(w);
  });
  var out = [];
  words.forEach(function (w) { if (out.indexOf(w) === -1) out.push(w); });
  return out;
}

function inventoryGameVariantConflict(item, rec) {
  var variation = String(item && item.variation || '');
  var sku = String(item && item.sku || '');
  var itemName = String(item && item.name || '');
  var recText = normalizeDbName(String(rec && rec.name || '') + ' ' + String(rec && rec.sku || ''));

  var editionPattern = /\b(Standard|Deluxe|Limited|Complete|Remastered|Collector'?s?|Special|Signature|Anniversary|Definitive|Ultimate|Porsche|Track|Encore)\b/ig;
  var variationEditions = variation.match(editionPattern) || [];
  var skuEditions = sku.match(editionPattern) || [];
  var titleEditions = itemName.match(editionPattern) || [];
  if (/\bFighters?\s+Edition\b/i.test(itemName)) titleEditions.push('Fighter');
  var selectedEdition = variationEditions.length ? variationEditions[variationEditions.length - 1] :
    (skuEditions.length ? skuEditions[skuEditions.length - 1] : '');

  // If the parent listing offers several editions but the buyer selected none,
  // do not guess which inventory edition was purchased.
  if (!selectedEdition && titleEditions.length > 1) return true;
  if (!selectedEdition && titleEditions.length === 1) selectedEdition = titleEditions[0];

  if (selectedEdition) {
    var selectedKey = normalizeDbName(selectedEdition).replace(/s$/, '');
    if (selectedKey === 'standard') {
      if (/\b(?:deluxe|limited|complete|collector|special|signature|anniversary|definitive|ultimate|porsche|track|encore)\b/i.test(recText)) return true;
    } else {
      if (recText.indexOf(selectedKey) === -1) return true;
    }
  }

  // A database record carrying a special edition label cannot be used for a
  // plain-title listing unless that edition is actually present in the raw
  // product title, variation or SKU reference.
  var sourceEditionText = (variation + ' ' + sku + ' ' + itemName).toLowerCase().replace(/fighters\b/g, 'fighter');
  var recordEdition = recText.match(/\b(Deluxe|Limited|Complete|Remastered|Collector|Special|Signature|Anniversary|Definitive|Ultimate|Porsche|Track|Encore)\b/i);
  if (recordEdition && sourceEditionText.indexOf(normalizeDbName(recordEdition[1])) === -1) return true;

  var selectedPearl = /\bShining\s+Pearl\b/i.test(variation);
  var selectedDiamond = /\bBrilliant\s+Diamond\b/i.test(variation);
  if (selectedPearl && !/\bShining\s+Pearl\b/i.test(recText)) return true;
  if (selectedDiamond && !/\bBrilliant\s+Diamond\b/i.test(recText)) return true;
  if ((selectedPearl || selectedDiamond) && /\bDouble\s+Pack\b/i.test(recText)) return true;

  // DLC matching must follow the buyer's selected variation/SKU, not merely
  // the parent listing title. Some Shopee parent titles mention an expansion
  // even when the selected variation is the base game (e.g. Pokemon Scarlet).
  var selectedSourceText = variation + ' ' + sku;
  var explicitNoDlc = /\b(?:without\s+DLC|no\s+DLC|DLC\s+not\s+included|not\s+include(?:d)?\s+DLC|excluding\s+DLC|DLC\s+excluded)\b/i.test(selectedSourceText + ' ' + itemName) ||
    /(?:不包含|不含|沒有包含|未包含|無)\s*(?:任何)?\s*DLC/i.test(selectedSourceText + ' ' + itemName);
  var selectedBasePokemon = /\b(?:scarlet|violet)\b/i.test(variation) &&
    /\bPokemon\s+Scarlet\/Violet\b/i.test(itemName) &&
    !/\b(?:DLC\s+Only|Expansion\s+Pass|Season\s+Pass)\b/i.test(selectedSourceText);
  var itemHasDlc = !explicitNoDlc && !selectedBasePokemon &&
    /\b(?:Hidden Treasure of Area Zero|Expansion Pass|DLC|Season Pass)\b/i.test(selectedSourceText + ' ' + itemName);
  var recHasDlc = /\b(?:Hidden Treasure of Area Zero|Expansion Pass|DLC|Season Pass)\b/i.test(String(rec && rec.name || '') + ' ' + String(rec && rec.sku || ''));
  if (itemHasDlc !== recHasDlc) return true;

  var itemHasDoublePack = /\bDouble\s+Pack\b/i.test(itemName);
  var recHasDoublePack = /\bDouble\s+Pack\b/i.test(recText);
  if (itemHasDoublePack !== recHasDoublePack) return true;

  var itemHasCharacterPass = /\bCharacter\s+Pass\b/i.test(itemName + ' ' + variation + ' ' + sku);
  var recHasCharacterPass = /\bCharacter\s+Pass\b/i.test(recText);
  if (itemHasCharacterPass !== recHasCharacterPass) return true;

  var itemNums = inventoryGameNumberSignature(itemName);
  var recNums = inventoryGameNumberSignature(rec && rec.name);
  if (itemNums && recNums && itemNums !== recNums) return true;
  return false;
}

function inventoryGameRecordCompatible(item, rec) {
  var requestedPlatform = inventoryGamePlatform(item);
  var recordPlatform = inventoryRecordGamePlatform(rec);
  if (!requestedPlatform || !recordPlatform || requestedPlatform !== recordPlatform) return false;

  if (inventoryItemIsUsedGame(item) !== inventoryRecordIsUsedGame(rec)) return false;

  var requestedRegion = inventoryGameRegion(item && item.variation) ||
    inventoryGameRegion(item && item.sku) ||
    inventoryGameRegion(item && item.name);
  var recordRegion = inventoryGameRegion(String(rec && rec.name || '') + ' ' + String(rec && rec.sku || ''));
  if (requestedRegion && recordRegion && requestedRegion !== recordRegion) return false;

  var requestedLanguage = inventoryRequestedGameLanguage(item);
  var recordLanguage = inventoryGameLanguage(String(rec && rec.name || '') + ' ' + String(rec && rec.sku || ''));
  if (requestedLanguage && recordLanguage && requestedLanguage !== recordLanguage) return false;

  if (inventoryGameVariantConflict(item, rec)) return false;

  var queryCore = inventoryGameCoreTokens(String(item && item.name || '') + ' ' + String(item && item.variation || '') + ' ' + String(item && item.sku || ''));
  var recordCore = inventoryGameCoreTokens(rec && rec.name);
  if (!queryCore.length || !recordCore.length) return false;
  var overlap = recordCore.filter(function (t) { return queryCore.indexOf(t) !== -1; }).length;
  // A shortened selected option may be a unique prefix of the inventory title
  // (e.g. "Aaero2: Black Ra" -> "Aaero2: Black Razor Edition").
  var variationCore = inventoryGameCoreTokens(String(item && item.variation || '') + ' ' + String(item && item.sku || ''));
  var prefixHit = variationCore.length >= 2 && variationCore.every(function (t) {
    return recordCore.some(function (r) { return r.indexOf(t) === 0; });
  });
  // The inventory title must substantially describe the same game, not merely share a franchise.
  if (overlap / recordCore.length < 0.80 && !prefixHit) return false;

  return true;
}

function inventoryGameMatchBoost(item, rec) {
  if (!inventoryGameRecordCompatible(item, rec)) return 0;
  var queryCore = inventoryGameCoreTokens(String(item && item.name || '') + ' ' + String(item && item.variation || '') + ' ' + String(item && item.sku || ''));
  var recordCore = inventoryGameCoreTokens(rec && rec.name);
  var overlap = recordCore.filter(function (t) { return queryCore.indexOf(t) !== -1; }).length;
  var coverage = overlap / Math.max(1, recordCore.length);
  var selected = String(item && item.variation || '');
  var variantSelected = /\b(?:Definitive|Deluxe|Limited|Collector'?s?|Ultimate|Porsche|Track|Shining\s+Pearl|Brilliant\s+Diamond)\b/i.test(selected);
  var variationCore = inventoryGameCoreTokens(String(item && item.variation || '') + ' ' + String(item && item.sku || ''));
  var prefixHit = variationCore.length >= 2 && variationCore.every(function (t) {
    return recordCore.some(function (r) { return r.indexOf(t) === 0; });
  });
  if (coverage >= 0.90) return 0.97;
  // Allow a few harmless inventory descriptors (e.g. "Digital Code in Box")
  // when platform, region, language and game identity already agree.
  if (coverage >= 0.80) return 0.93;
  if ((variantSelected || prefixHit) && coverage >= 0.60) return 0.93;
  return 0;
}

function isVideoGameInventoryRecord(item, rec) {
  return isVideoGameLineItem(item);
}

function inventoryResolveUniqueGameRegion(item, rec, db) {
  var explicit = inventoryGameRegion(item && item.variation) ||
    inventoryGameRegion(item && item.sku) ||
    inventoryGameRegion(item && item.name);
  if (explicit) return explicit;

  var key = rec && rec.nameKey ? rec.nameKey : normalizeDbName(rec && rec.name || '');
  var indexes = db && db.name && db.name[key] ? db.name[key] : [];
  var regions = {};
  indexes.forEach(function (idx) {
    var candidate = db.records[idx];
    if (!inventoryGameRecordCompatible(item, candidate)) return;
    var region = inventoryGameRegion(String(candidate.name || '') + ' ' + String(candidate.sku || ''));
    if (region) regions[region] = true;
  });
  var found = Object.keys(regions);
  return found.length === 1 ? found[0] : '';
}

function inventoryRecordWithGameRegion(item, rec, db) {
  if (!rec || !isVideoGameLineItem(item)) return rec;
  var copy = Object.assign({}, rec);
  copy._matchedRegion = inventoryResolveUniqueGameRegion(item, rec, db);
  return copy;
}

function matchInventoryProduct(item, resolvedName, db) {
  if (!db || !db.available) return { status: 'missing-db', note: 'DB missing; existing rules used' };
  var gameLine = isVideoGameLineItem(item);
  var skuKey = normalizeDbKey(item.sku);
  var rawNameKey = normalizeDbName(item.name);
  var resolvedKey = normalizeDbName(resolvedName);

  function compatible(rec) {
    return !inventoryVariationConflicts(item, rec) && (!gameLine || inventoryGameRecordCompatible(item, rec));
  }

  if (skuKey && db.sku[skuKey]) {
    var exactSkuRows = db.sku[skuKey];
    if (exactSkuRows.length === 1 && compatible(db.records[exactSkuRows[0]])) {
      return { record: inventoryRecordWithGameRegion(item, db.records[exactSkuRows[0]], db), score: 1, method: gameLine ? 'exact SKU / compatible game title' : 'exact SKU / Item Code', status: 'matched' };
    }
    var skuRanked = exactSkuRows.map(function (idx) {
      var r = db.records[idx];
      var score = Math.max(inventorySimilarity(item.name + ' ' + item.variation, r), inventorySimilarity(resolvedName, r));
      if (gameLine && compatible(r)) score = Math.max(score, inventoryGameMatchBoost(item, r));
      return { record: r, score: score, conflict: !compatible(r) };
    }).filter(function (x) { return !x.conflict; }).sort(function (a, b) { return b.score - a.score; });
    if (skuRanked.length === 1 || (skuRanked.length > 1 && skuRanked[0].score - (skuRanked[1] ? skuRanked[1].score : 0) >= 0.08)) {
      return { record: inventoryRecordWithGameRegion(item, skuRanked[0].record, db), score: gameLine ? Math.max(0.90, skuRanked[0].score) : 0.99, method: 'exact SKU / Item Code + title check', status: 'matched' };
    }
    if (skuRanked.length > 1) return { status: 'ambiguous', note: 'REVIEW: duplicate database SKU; no automatic replacement', score: skuRanked[0].score };
  }

  var exactKeys = [rawNameKey, resolvedKey];
  for (var e = 0; e < exactKeys.length; e++) {
    if (exactKeys[e] && db.name[exactKeys[e]]) {
      var exactCandidates = db.name[exactKeys[e]].map(function (idx) { return db.records[idx]; })
        .filter(compatible);
      if (exactCandidates.length === 1) return { record: inventoryRecordWithGameRegion(item, exactCandidates[0], db), score: 0.99, method: 'exact product name', status: 'matched' };
    }
  }

  var codes = extractInventoryCodes(item.name + ' ' + item.variation + ' ' + item.sku + ' ' + resolvedName);
  var candidates = {};
  codes.forEach(function (code) { (db.code[code] || []).forEach(function (idx) { candidates[idx] = true; }); });

  var queryTexts = [item.name, resolvedName, item.name + ' ' + item.variation, item.name + ' ' + item.variation + ' ' + item.sku];
  if (gameLine) {
    queryTexts.push(item.variation + ' ' + item.sku);
    queryTexts.push(cleanMatchedGameName(item.name));
  }
  var queryTokens = {};
  queryTexts.forEach(function (q) { inventoryTokens(q).forEach(function (t) { queryTokens[t] = true; }); });
  Object.keys(queryTokens).forEach(function (token) {
    var list = db.token[token] || [];
    // Avoid candidate explosion from a single popular word.
    if (list.length > 700) return;
    list.forEach(function (idx) { candidates[idx] = true; });
  });
  if (gameLine) {
    inventoryGameCoreTokens(String(item.name || '') + ' ' + String(item.variation || '') + ' ' + String(item.sku || '')).forEach(function (token) {
      var list = db.token[token] || [];
      if (list.length > 700) return;
      list.forEach(function (idx) { candidates[idx] = true; });
    });
  }

  var ranked = Object.keys(candidates).map(function (key) {
    var rec = db.records[Number(key)];
    var isCompatible = compatible(rec);
    var score = 0;
    if (isCompatible) {
      queryTexts.forEach(function (q) { score = Math.max(score, inventorySimilarity(q, rec)); });
      var codeHit = codes.some(function (code) { return rec.codes.indexOf(code) !== -1; });
      var uniqueCodeHit = codes.some(function (code) {
        var hits = db.code[code] || [];
        return hits.length === 1 && hits[0] === Number(key);
      });
      if (uniqueCodeHit) score = Math.max(score, 0.97);
      if (gameLine) score = Math.max(score, inventoryGameMatchBoost(item, rec));
      return { record: rec, score: score, codeHit: codeHit, uniqueCodeHit: uniqueCodeHit };
    }
    return { record: rec, score: 0, codeHit: false, uniqueCodeHit: false };
  }).filter(function (x) { return x.score > 0; }).sort(function (a, b) { return b.score - a.score; });

  if (!ranked.length) {
    return { status: 'unmatched', note: gameLine ? 'DB: no compatible game/platform/edition match; existing rules used' : 'DB: no match; existing rules used' };
  }
  var best = ranked[0];
  var second = ranked[1] ? ranked[1].score : 0;

  // Video games use a strict 90% confidence threshold after platform,
  // used/new status, region, language and selected-edition compatibility.
  if (gameLine && best.score >= 0.90) {
    return { record: inventoryRecordWithGameRegion(item, best.record, db), score: best.score, method: 'video-game inventory name similarity + variant validation', status: 'matched' };
  }
  if (!gameLine && best.score >= 0.80 && (best.score - second >= 0.08 || best.uniqueCodeHit)) {
    return { record: best.record, score: best.score, method: best.uniqueCodeHit ? 'unique model code' : 'product-name similarity', status: 'matched' };
  }
  if (best.score >= 0.48) {
    return { status: 'ambiguous', note: 'REVIEW: possible inventory match (' + Math.round(best.score * 100) + '%); existing rules used', score: best.score };
  }
  return { status: 'unmatched', note: 'DB: no reliable match; existing rules used' };
}

function cleanInventoryCanonicalName(value) {
  return cleanReferenceName(String(value || '').trim())
    .replace(/\s*\[(?:\d+\s*Years?\s+)?Warranty\]\s*$/i, '')
    .replace(/\s*\(\d+\s*\/\s*[Ss]et\)\s*$/i, '')
    .replace(/\s+/g, ' ').trim();
}

function extractDbItemCode(record) {
  var sku = String(record && record.sku || '').trim();
  var name = String(record && record.name || '').trim();
  var sources = [sku, name];
  // Prefer explicit model/item codes in square brackets, e.g. [TNS-0122B],
  // [L1094], [NSX-249A]. Ignore region markers such as [R3].
  for (var i = 0; i < sources.length; i++) {
    var matches = sources[i].match(/\[([A-Z0-9]+(?:-[A-Z0-9]+)+|[A-Z]{1,5}\d{2,10}[A-Z]?)\]/ig) || [];
    for (var j = 0; j < matches.length; j++) {
      var code = matches[j].replace(/[\[\]]/g, '').trim();
      if (!/^(R[0-9]|REGION[0-9])$/i.test(code)) return code.toUpperCase();
    }
  }
  // SKU / Item Code may itself be a bare code.
  if (sku && /^[A-Z0-9]+(?:-[A-Z0-9]+)+$/i.test(sku)) return sku.toUpperCase();
  if (sku && /^[A-Z]{1,5}\d{2,10}[A-Z]?$/i.test(sku)) return sku.toUpperCase();
  return '';
}

function extractDbItemCodeOriginal(record) {
  var sources = [String(record && record.sku || '').trim(), String(record && record.name || '').trim()];
  for (var i = 0; i < sources.length; i++) {
    var paren = sources[i].match(/\(([A-Z0-9]{3,}(?:-[A-Z0-9]+){1,3})\)/i);
    if (paren && !/^(?:R[0-9]|REGION[0-9])$/i.test(paren[1])) return paren[1];
    var bracket = sources[i].match(/\[([A-Z0-9]{3,}(?:-[A-Z0-9]+){1,3})\]/i);
    if (bracket && !/^(?:R[0-9]|REGION[0-9])$/i.test(bracket[1])) return bracket[1];
    var codes = sources[i].match(/\b[A-Z0-9]{3,}(?:-[A-Z0-9]{2,}){1,3}\b/ig) || [];
    for (var j = 0; j < codes.length; j++) if (/\d/.test(codes[j])) return codes[j];
  }
  return '';
}

function inventoryPlatformLabel(text) {
  var t = String(text || '');
  if (/\bPSVR2?\b|\bPlayStation VR\b/i.test(t)) return 'PSVR2';
  if (/\b(?:Nintendo\s*)?Switch\s*2\b|\bSwitch2\b/i.test(t)) return 'Switch2';
  if (/\bSwitch\b|\bNintendo\b/i.test(t)) return 'Switch';
  if (/\bPS5\b|\bPlayStation\s*5\b|\bPs5\b/i.test(t)) return 'PS5';
  if (/\bPS4\b|\bPlayStation\s*4\b|\bPs4\b/i.test(t)) return 'PS4';
  if (/\bPS3\b|\bPlayStation\s*3\b|\bPs3\b/i.test(t)) return 'PS3';
  if (/\bXbox\b/i.test(t)) return 'Xbox';
  if (/\bPSVR2?\b|\bPlayStation VR\b/i.test(t)) return 'PS';
  return '';
}

function cleanMatchedGameName(value) {
  var text = String(value || '').replace(/[\u4e00-\u9fff]+/g, ' ');
  // Keep region markers, but discard bracketed internal item codes.
  text = text.replace(/\[\s*(?:R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\s*\]/gi, ' ');
  // Preserve meaningful game edition/cover descriptors such as [Ecto Edition].
  text = text.replace(/\[([^\]]*(?:Edition|Cover)[^\]]*)\]/ig, ' $1 ');
  text = text.replace(/\[[^\]]*\]/g, ' ');
  // Remove inventory item codes even when the source export has a malformed bracket.
  text = text.replace(/\b[A-Z]{2,8}-[A-Z0-9]{3,}\]?\b/gi, ' ');
  text = text.replace(/[\[\]]/g, ' ');
  text = text.replace(/\([^)]*[\u4e00-\u9fff][^)]*\)/g, ' ');
  text = text.replace(/[\u4e00-\u9fff]+/g, ' ');
  text = text.replace(/\bPs([345])\b/gi, function (_, n) { return 'PS' + n; });
  text = text.replace(/\bPlaystation\b/gi, 'PlayStation');
  text = text.replace(/^NSW\/NSW2\s+Nintendo Switch\/Switch\s*2\s*/i, 'Switch ');
  text = text.replace(/^NSW\/NSW2\s+Nintendo Switch 1\/2\s*/i, 'Switch ');
  text = text.replace(/^NSW2\s+Nintendo Switch 2\s*/i, 'Switch2 ');
  text = text.replace(/^Switch\s+2\s+1\/2\s*/i, 'Switch2 ');
  text = text.replace(/^Switch\s+2\s+/i, 'Switch2 ');
  text = text.replace(/^NSW\s+Nintendo Switch\s*/i, 'Switch ');
  text = text.replace(/^Switch\s+Nintendo Switch\s+/i, 'Switch ');
  text = text.replace(/^PlayStation\s+PS([345])\s*/i, 'PS$1 ');
  text = text.replace(/^PS([345])\s+/i, 'PS$1 ');
  text = text.replace(/\bChinese\s*\/\s*English\s+Version\b/ig, 'Chi/Eng');
  text = text.replace(/\bChi\s*\/\s*Eng\s+Version\b/ig, 'Chi/Eng');
  text = text.replace(/\bEnglish\s+Version\b/ig, 'Eng');
  text = text.replace(/\bGTA\s*6\s+Grand\s+Theft\s+Auto\s+VI\b/ig, 'Grand Theft Auto VI');
  text = text.replace(/\s*\+\s*(?=Chi\/Eng|Eng\b)/ig, ' ');
  text = text.replace(/\s+/g, ' ').trim();
  return text;
}

function cleanDbMultilineText(value) {
  return String(value || '')
    .replace(/<\s*br\s*\/?\s*>/gi, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s*\+\s*/g, ' + ')
    .replace(/\s+/g, ' ').trim();
}

function dedupeProductWording(value) {
  var text = String(value || '').replace(/\s+/g, ' ').trim();
  // Normalize known duplicate game-title aliases created by Shopee listing titles.
  text = text.replace(/\bLoverPretend\s+Lover\s+Pretend\b/ig, 'Lover Pretend');
  text = text.replace(/\bLover\s+Pretend\s+Lover\s+Pretend\b/ig, 'Lover Pretend');
  text = text.replace(/\bMegaMan\s+Mega\s+Man\b/ig, 'Mega Man');
  text = text.replace(/\bMegaman\s+Mega\s+Man\b/ig, 'Mega Man');
  text = text.replace(/\bMega\s+Man\s+Mega\s+Man\b/ig, 'Mega Man');
  text = text.replace(/\bSpider\s+Man\s+Spiderman\s+Miles\s+Morales\b/ig, 'Spider-Man: Miles Morales');
  text = text.replace(/\bSpider\s+Man\s+Spiderman\b/ig, 'Spider-Man');
  text = text.replace(/\b(Standard|Limited|Deluxe|Collector'?s?)\s+Edition\b(.*)\b\1\s+Edition\b/ig, '$1 Edition$2');
  text = text.replace(/\b(Chi\/Eng|Eng\/Chi|Eng)\s+\1\b/ig, '$1');
  text = text.replace(/\b(Chi|Eng)\s+\1\b/ig, '$1');
  text = text.replace(/\bEnglish\s+Eng\b/ig, 'Eng');
  text = text.replace(/\bChinese\/English\s+Version\b/ig, 'Chi/Eng');
  text = text.replace(/\bChi\/Eng\s+Chi\/Eng\b/ig, 'Chi/Eng');
  return text.replace(/\s+/g, ' ').trim();
}

function preferredCompactInventoryName(item, record, currentName) {
  var sourceText = String(item.name || '') + ' ' + String(record.name || '') + ' ' + String(record.sku || '');
  var categoryText = String(record.category || '');
  var itemCode = extractDbItemCode(record);

  // Logitech wheel: the selected model in the variation is authoritative.
  if (/\bLogitech\b/i.test(sourceText) && /\b(?:G29|G920|G923)\b/i.test(sourceText) && /RACING|STEERING WHEEL/i.test(sourceText)) {
    var selectedLogitechModel = (String(item.variation || '').match(/\b(G29|G920|G923)\b/i) || [])[1] || (String(item.name || '').match(/\b(G29|G920|G923)\b/i) || [])[1] || '';
    if (selectedLogitechModel) {
      var shifterLogitech = /SHIFTER/i.test(String(item.variation || '')) || (!String(item.variation || '').trim() && /SHIFTER/i.test(item.name + ' ' + record.name + ' ' + record.sku));
      return 'Logitech ' + selectedLogitechModel.toUpperCase() + ' Driving Force Racing Wheel + Pedals' + (shifterLogitech ? ' + Shifter' : '');
    }
  }

  // Anbernic orders may intentionally contain a console and a memory card in
  // one SKU reference. Keep both parts and flatten the exported HTML break.
  if (/ANBERNIC/i.test(sourceText) && /MEMORY CARD/i.test(String(item.sku || ''))) {
    var anbernicCombined = cleanDbMultilineText(item.sku);
    if (anbernicCombined) return anbernicCombined;
  }

  // Sony DualSense controller codes (CFI-ZCT2G etc.) are internal hardware
  // identifiers, not part of the customer-facing product name.
  if (/DUALSENSE/i.test(sourceText) && /CONTROLLER/i.test(sourceText)) {
    var controllerName = cleanInventoryCanonicalName(record.name || record.sku || currentName);
    controllerName = controllerName.replace(/^Ps5\s+Acc\s+Asia\s+/i, 'PS5 ');
    controllerName = controllerName.replace(/\s*-?\s*CFI-ZCT[0-9]G(?:\s+[A-Z0-9]+)?\b/ig, '');
    controllerName = controllerName.replace(/\s+/g, ' ').trim();
    if (controllerName) return controllerName;
  }

  // LEGO: always use the selected five-digit set number, not the first
  // number in a multi-option listing title.
  if (/^LEGO\b/i.test(sourceText) || /^LEGO\b/i.test(categoryText)) {
    var legoSetCode = (String(item.variation || '').match(/\b\d{5}\b/) || String(item.sku || '').match(/\b\d{5}\b/) || String(record.sku || '').match(/\b\d{5}\b/) || String(record.name || '').match(/\b\d{5}\b/) || String(item.name || '').match(/\b\d{5}\b/) || [])[0] || '';
    if (legoSetCode) return 'LEGO ' + legoSetCode;
  }

  // ZD Toys: use the exact inventory model identified by the numeric SKU
  // reference (e.g. 1907-04 -> Mk4 LED Version; 1909-50 -> Mk50). Keep a
  // selected variation descriptor when it adds detail absent from inventory.
  if (/\bZD\s*Toys\b/i.test(sourceText)) {
    var zdName = cleanDbMultilineText(record.name || currentName);
    var zdVariation = String(item.variation || '');
    var zdModel = (zdVariation.match(/\bMark\s*(\d{1,2})\b/i) || zdVariation.match(/\bMk\s*(\d{1,2})\b/i) || [])[1];
    var zdRecordModel = ((zdName.match(/\bMk\s*(\d{1,2})\b/i) || [])[1]);
    if (zdModel && zdRecordModel && Number(zdModel) !== Number(zdRecordModel)) {
      // A mismatched model is never allowed to inherit the database title.
      zdName = 'ZD Toys Iron Man Mk' + Number(zdModel);
    }
    var zdCode = (String(item.sku || '').match(/\b\d{3,5}-\d{2,5}\b/) || String(record.name || '').match(/\b\d{3,5}-\d{2,5}\b/) || [])[0];
    if (zdCode) zdName = zdName.replace(new RegExp('\\s*\\(' + zdCode.replace('-', '\\-') + '\\)\\s*$', 'i'), '');
    if (/\bLED\b/i.test(zdVariation) && !/\bLED\b/i.test(zdName)) zdName += ' LED Version';
    var energyCanon = (zdVariation.match(/\bEnergy\s+Canon\b/i) || [])[0];
    if (energyCanon && !/\bEnergy\s+Canon\b/i.test(zdName)) zdName += ' Energy Canon';
    if (zdCode) zdName += ' (' + zdCode + ')';
    return zdName.replace(/\s+/g, ' ').trim();
  }

  // Blokees: brand + numeric model code only.
  if (/BLOKEES/i.test(sourceText)) {
    var blokeesCode = extractInventoryCodes(currentName + ' ' + item.variation + ' ' + item.name + ' ' + record.sku + ' ' + record.name)
      .filter(function (c) { return /^\d{5}$/.test(c); })[0] || '';
    if (blokeesCode) return 'Blokees ' + blokeesCode;
  }

  // Only Bandai Model Kits may be shortened to brand + kit code.
  if (/^Bandai Model Kits$/i.test(categoryText)) {
    var bandaiKitCode = (sourceText.match(/\b(50\d{5})\b/) || [])[1] || '';
    if (bandaiKitCode) return 'Bandai ' + bandaiKitCode;
  }

  // Die-cast: model/brand + exact catalogue item code.
  // Normalize the catalogue code to uppercase for consistent output
  // (e.g. IN64-LBMIURA-01 rather than IN64-LBMiura-01).
  if (/^Die-Cast Vehicles(?:\s*-\s*|$)/i.test(categoryText)) {
    var diecastBrand = (categoryText.match(/^Die-Cast Vehicles\s*-\s*(.+)$/i) || [])[1] || '';
    var diecastCode = extractDbItemCodeOriginal(record);
    if (diecastBrand && diecastCode) return diecastBrand + ' ' + diecastCode.toUpperCase();
  }

  // Accessory naming: brand + platform + database item code.
  var accessoryBrand = '';
  if (/\bGeekShare\b/i.test(sourceText)) accessoryBrand = 'GeekShare';
  else if (/\bDobe\b/i.test(sourceText)) accessoryBrand = 'Dobe';
  else if (/\biiNE\b/i.test(sourceText)) accessoryBrand = 'iiNE';
  else if (/\bHori\b/i.test(sourceText)) accessoryBrand = 'Hori';
  else if (/\bSkull\s*&\s*Co\b/i.test(sourceText)) accessoryBrand = 'Skull & Co';
  else if (/\bPMW\b/i.test(sourceText)) accessoryBrand = 'PMW';
  if (accessoryBrand && itemCode && /accessor|peripheral|gadget/i.test(categoryText + ' ' + sourceText)) {
    var platform = inventoryPlatformLabel(sourceText + ' ' + item.variation);
    if (accessoryBrand === 'GeekShare' || accessoryBrand === 'PMW') return accessoryBrand + ' ' + itemCode;
    return accessoryBrand + (platform ? ' ' + platform : '') + ' ' + itemCode;
  }

  // 8BitDo: retain full product identity and append its database item code.
  if (/\b8BitDo\b/i.test(sourceText)) {
    var bitdoCode = (String(record.sku || '').match(/\(([A-Z0-9]{5,10})\)\s*$/i) || [])[1] || itemCode || (String(item.sku || '').match(/\b\d{5,10}\b/) || [])[0] || '';
    var bitdoName = cleanDbMultilineText(record.name || record.sku || currentName)
      .replace(/\b8bitdo\b/gi, '8BitDo')
      .replace(/\s*\([^)]*\)\s*$/, '').trim();
    if (bitdoName && bitdoCode && bitdoName.toUpperCase().indexOf(bitdoCode.toUpperCase()) === -1) bitdoName += ' (' + bitdoCode.toUpperCase() + ')';
    return bitdoName || currentName;
  }

  // Console games: use the database's full SKU / Item Code as the title,
  // but source language and region from BOTH the listing and variation.
  if (/\b(?:Ps[345]\s*Games|Switch\s*2?\s*Games|Xbox\s*Series\s*Games|PC\s*Game)\b/i.test(categoryText) ||
      /\b(?:PS[345]|Switch|Xbox)\b/i.test(sourceText) && /\b(?:Games|Game)\b/i.test(categoryText)) {
    // Use the inventory Product Name as the canonical game title.
    var gameSource = String(record.name || '').trim();
    if (!gameSource) gameSource = record.sku;
    var gameName = cleanMatchedGameName(gameSource);
    // Cover language is packaging metadata, not the game language. Preserve
    // it separately because the generic language cleanup would otherwise turn
    // '(CHI Cover)' into the broken '( Cover)' or move it before Chi/Eng.
    var gameCover = ((gameSource.match(/\((?:CHI|ENG)\s+Cover\)/i) || [])[0] || '');
    if (gameCover) gameName = gameName.replace(/\((?:CHI|ENG)\s+Cover\)/ig, ' ');
    var itemGameFields = [item.name, item.variation, item.sku].join(' ');
    // Do not inherit a region from an arbitrary database row when the order
    // itself did not specify one (e.g. several regional SKUs share one title).
    var gameRegion = inventoryGameRegion(item.variation) ||
      inventoryGameRegion(item.sku) ||
      inventoryGameRegion(item.name) ||
      record._matchedRegion || '';
    var gameLanguage = inventoryRequestedGameLanguage(item) ||
      inventoryGameLanguage(record.name) ||
      inventoryGameLanguage(record.sku);
    if (gameName) {
      gameName = stripGameRegionMarker(gameName)
        .replace(/\b(?:Chi\/Eng|Eng\/Chi|Eng|Chi)\b/ig, ' ')
        .replace(/\s+/g, ' ').trim();
      // Put the selected edition before language and region markers.
      var variationEditionSource = String(item.variation || '');
      var skuEditionSource = String(item.sku || '');
      var editionPattern = /\b(Standard|Limited|Deluxe|Complete|Collector'?s?|Definitive|Ultimate|Porsche|Track|Fighters?|Anniversary|Special|Signature|Encore)\b/i;
      var editionPhrasePattern = /\b(Standard|Limited|Deluxe|Complete|Collector'?s?|Definitive|Ultimate|Porsche|Track|Fighters?|Anniversary|Special|Signature|Encore)\s+Edition\b/i;
      var variationPhrase = (variationEditionSource.match(editionPhrasePattern) || [])[0] || '';
      var variationWord = (variationEditionSource.match(editionPattern) || [])[1] || '';
      var skuPhrase = (skuEditionSource.match(editionPhrasePattern) || [])[0] || '';
      var titleEdition = (String(item.name || '').match(editionPhrasePattern) || [])[0] || '';
      var editionPick = variationPhrase ||
        (variationWord ? variationWord.replace(/s$/i, '') + ' Edition' : (skuPhrase || titleEdition));
      if (editionPick) editionPick = editionPick.replace(/^Fighters\b/i, 'Fighter');
      if (editionPick && !new RegExp('\\b' + editionPick.replace(/\s+/g, '\\s+') + '\\b', 'i').test(gameName)) gameName += ' ' + editionPick;
      if (gameLanguage) gameName += ' ' + gameLanguage;
      if (gameRegion) gameName += ' ' + gameRegion;
      if (gameCover) gameName += ' ' + gameCover;
      return dedupeProductWording(gameName);
    }
  }

  var productName = cleanInventoryCanonicalName(record.name);
  var skuName = cleanInventoryCanonicalName(record.sku);
  if (/^Original Tamagotchi\b/i.test(item.name)) {
    productName = productName.replace(/^Tamagotchi Original\b/i, 'Original Tamagotchi').replace(/^Original Tamagotchi\s*-\s*/i, 'Original Tamagotchi ');
    skuName = skuName.replace(/^Tamagotchi Original\b/i, 'Original Tamagotchi').replace(/^Original Tamagotchi\s*-\s*/i, 'Original Tamagotchi ');
  }
  if (skuName && skuName.length > productName.length + 2) {
    var skuCodes = extractInventoryCodes(skuName);
    var nameCodes = extractInventoryCodes(productName);
    var hasAdditionalCode = skuCodes.some(function (c) { return nameCodes.indexOf(c) === -1; });
    if (hasAdditionalCode && inventorySimilarity(productName, { tokens: inventoryTokens(skuName) }) >= 0.55) productName = skuName;
  }
  return productName || currentName;
}

function isPlatformLanguageOnlyGameVariation(variation) {
  var text = String(variation || '')
    .replace(/\b(?:Full Payment|Balance Payment|Deposit|Partial Payment|P\.?O\.?|Ready Stock|Random Cover|No Add-?On|Game Only|Game \+ Goodie Bag|With Gift|No Gift)\b/ig, ' ')
    .replace(/\b(?:PS[345]|PlayStation\s*[345]|Switch\s*[12]?|Switch2?|Xbox|Nintendo)\b/ig, ' ')
    .replace(/\b(?:Chi\/Eng|Eng\/Chi|Chinese\/English|English|Chinese|Japanese|Eng|Chi|JAP|R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\b/ig, ' ')
    .replace(/中文版|中文|英文|日本語/ig, ' ')
    .replace(/[\[\],+\/\-]/g, ' ')
    .replace(/\s+/g, ' ').trim();
  return !text;
}

function cleanRawGameListingName(item) {
  var raw = String(item && item.name || '');
  var platform = inventoryGamePlatform(item);
  if (/Promotion Games/i.test(raw) && String(item && item.variation || '').trim()) {
    raw = String(item.variation || '');
  } else {
    raw = raw
      .replace(/[\u4e00-\u9fff]+/g, ' ')
      .replace(/\b(?:Chinese\s*\/\s*English|Chi\s*\/\s*Eng|Eng\s*\/\s*Chi|English|Chinese|Japanese|Eng|Chi)\s+Version\b/ig, ' ')
      .replace(/\bVersion\b/ig, ' ')
      .replace(/\([^)]*\)/g, ' ');
    if (/^NSW\/NSW2\s+Nintendo Switch\/Switch\s*2\s*/i.test(raw)) {
      raw = raw.replace(/^NSW\/NSW2\s+Nintendo Switch\/Switch\s*2\s*/i, (platform === 'Switch2' ? 'Switch2 ' : 'Switch '));
    } else if (/^NSW\/NSW2\s+Nintendo Switch 1\/2\s*/i.test(raw)) {
      raw = raw.replace(/^NSW\/NSW2\s+Nintendo Switch 1\/2\s*/i, (platform === 'Switch2' ? 'Switch2 ' : 'Switch '));
    } else if (/^NSW2\s+Nintendo Switch 2\s*/i.test(raw)) {
      raw = raw.replace(/^NSW2\s+Nintendo Switch 2\s*/i, 'Switch2 ');
    } else if (/^NSW\s+Nintendo Switch\s*/i.test(raw)) {
      raw = raw.replace(/^NSW\s+Nintendo Switch\s*/i, 'Switch ');
    } else if (/^NSW\s+Switch\s*/i.test(raw)) {
      raw = raw.replace(/^NSW\s+Switch\s*/i, 'Switch ');
    } else if (/^Nintendo Switch\/Switch\s*2\s*/i.test(raw)) {
      raw = raw.replace(/^Nintendo Switch\/Switch\s*2\s*/i, (platform === 'Switch2' ? 'Switch2 ' : 'Switch '));
    } else if (/^PlayStation\s+PS([345])\s*/i.test(raw)) {
      raw = raw.replace(/^PlayStation\s+PS([345])\s*/i, 'PS$1 ');
    } else if (/^Playstation\s+PS([345])\s*/i.test(raw)) {
      raw = raw.replace(/^Playstation\s+PS([345])\s*/i, 'PS$1 ');
    } else if (/^PS4\/PS5\s*/i.test(raw)) {
      raw = raw.replace(/^PS4\/PS5\s*/i, (platform || 'PS5') + ' ');
    }
  }
  raw = cleanMatchedGameName(raw);
  if (platform === 'Switch2') raw = raw.replace(/^Switch\b/i, 'Switch2');
  if (platform === 'Switch') raw = raw.replace(/^Switch2\b/i, 'Switch');
  return dedupeProductWording(raw);
}

function normalizeUnmatchedGameName(item, currentName) {
  var source = [item.name, item.variation, currentName].join(' ');
  if (!isVideoGameLineItem(item)) return currentName;

  var text = cleanMatchedGameName(currentName);
  // If variation only identifies platform/language/region (not a title option),
  // restore the full source title. This prevents "Final Fantasy X / X-2" from
  // being truncated to "X-2" by the generic slash-option resolver.
  var rawTitle = cleanRawGameListingName(item);
  var editionOptionPattern = /\b(Standard|Deluxe|Limited|Complete|Remastered|Collector'?s?|Special|Signature|Anniversary|Definitive|Ultimate|Porsche|Track|Encore)\b/ig;
  var titleEditionOptions = String(item.name || '').match(editionOptionPattern) || [];
  var selectedEdition = (String(item.variation || '') + ' ' + String(item.sku || '')).match(editionOptionPattern) || [];
  if (titleEditionOptions.length > 1 && !selectedEdition.length && rawTitle) {
    text = rawTitle;
  } else if (isPlatformLanguageOnlyGameVariation(item.variation) && rawTitle && rawTitle.length > text.length + 4) {
    text = rawTitle;
  }

  // Select a buyer-chosen edition when the listing explicitly exposes edition options.
  var variation = String(item.variation || '');
  if (/\bStandard\b/i.test(variation) && /Standard\s*\/\s*(?:Deluxe|Complete|Limited)\s+Edition/i.test(text)) {
    text = text.replace(/Standard\s*\/\s*(?:Deluxe|Complete|Limited)\s+Edition/i, 'Standard Edition');
  } else if (/\bComplete\b/i.test(variation) && /Standard\s*\/\s*Complete\s+Edition/i.test(text)) {
    text = text.replace(/Standard\s*\/\s*Complete\s+Edition/i, 'Complete Edition');
  } else if (/\bDeluxe\b/i.test(variation) && /Standard\s*\/\s*Deluxe\s+Edition/i.test(text)) {
    text = text.replace(/Standard\s*\/\s*Deluxe\s+Edition/i, 'Deluxe Edition');
  } else if (/\bLimited\b/i.test(variation) && /Standard\s*\/\s*Limited\s+Edition/i.test(text)) {
    text = text.replace(/Standard\s*\/\s*Limited\s+Edition/i, 'Limited Edition');
  }

  var fields = [item.variation, item.name].join(' ');
  var lang = inventoryRequestedGameLanguage(item);
  var region = inventoryGameRegion(item.variation) ||
    inventoryGameRegion(item.sku) ||
    inventoryGameRegion(item.name);

  // Strip duplicate language/region markers before appending the selected values.
  text = text.replace(/\b(?:Chi\/Eng|Eng\/Chi|Eng|English|Chinese|Chi|JAP)\b/ig, ' ')
    .replace(/\bVersion\b/ig, ' ')
    .replace(/\s+/g, ' ').trim();
  if (region) text = text.replace(new RegExp('\\b' + region + '\\b', 'ig'), ' ');
  text = stripGameRegionMarker(text).replace(/\s*[:\-]\s*$/, '').trim();
  if (lang) text += ' ' + lang;
  if (region) text += ' ' + region;
  return dedupeProductWording(text);
}

function isFujifilmFilmOnlyProduct(value) {
  var text = String(value || '');
  return /\bFujifilm\b/i.test(text) && /\bfilm\b/i.test(text) &&
    !/\b(?:camera|printer|case|album|photo\s*book|paper\s*frame|photo\s*frame|stand|bag|accessor(?:y|ies))\b/i.test(text);
}

function mapInventoryCategory(dbCategory, productName, tags) {
  var source = String(dbCategory || '').trim();
  var text = String(productName || '') + ' ' + String(tags || '');

  // Fujifilm film-only consumables belong to the existing Xbox/Gaming Acc group.
  // Cameras, printers, cases, albums and other Fujifilm hardware remain unchanged.
  if (isFujifilmFilmOnlyProduct(productName) || isFujifilmFilmOnlyProduct(text)) {
    return { category: 'Xbox Gaming Acc', basis: 'Fujifilm film-only category override' };
  }

  if (/\b8BitDo\b/i.test(text) || /^Accessories \[8bitdo\]$/i.test(source)) {
    return { category: 'Xbox Gaming Acc', basis: 'existing 8BitDo category rule (fixed)' };
  }

  // These product families are always Retail, regardless of platform words
  // in their titles or the broad inventory source category.
  if (/Used Games\s*\[|Used Ps[345] Acc|Used Console|PlayStation [45] Console|Nintendo Switch(?:\s*2)? Console|Switch\s*2? Console|Peripherals \[Aula\]|Peripherals \[Keychron\]/i.test(source) ||
      /DUALSENSE/i.test(text) && /CONTROLLER/i.test(text) ||
      /^(?:Aula|Keychron)\b/i.test(productName || '') || /\bRazer\b/i.test(productName || '') ||
      /^Peripherals \[Razer\]$/i.test(source)) {
    return { category: 'Retail', basis: 'inventory: retail product family override' };
  }

  if (/^TCG/i.test(source)) {
    if (/booster\s*pack|sleeved\s*booster|blister\s*pack|promo\s*card/i.test(text)) return { category: 'TCG BP', basis: 'inventory: ' + source };
    return { category: 'TCG', basis: 'inventory: ' + source };
  }
  if (/blind.?box|blindbox|blind pack/i.test(source)) return { category: 'Blind Box', basis: 'inventory: ' + source };
  if (/^LEGO/i.test(source)) return { category: 'LEGO', basis: 'inventory: ' + source };
  if (/^(?:Ps[345] Games|Used Games \[Ps[345]\]|Psvr Games|Ps[345] Accessories|PlayStation [45] Official Accessories)/i.test(source)) return { category: 'PS', basis: 'inventory: ' + source };
  if (/^(?:Switch ?2? Games|Used Games \[Switch|Switch2? Accessories|Switch Official Accessories|Used Accessories \[Switch|Switch Console|Switch2 Console|3DS Game)/i.test(source)) return { category: 'SWITCH', basis: 'inventory: ' + source };
  if (/^(?:Xbox Series Games|Xbox Series Official Accessories|Xbox Series Console|Used Console \[Xbox)/i.test(source)) return { category: 'Xbox Gaming Acc', basis: 'inventory: ' + source };
  if (/^PC Game$/i.test(source)) return { category: 'Xbox Gaming Acc', basis: 'inventory: PC Game mapped to Xbox Gaming Acc' };
  if (/^Gunpla Accessories/i.test(source)) return { category: 'Gundam ACC  Mr.Hobby', basis: 'inventory: ' + source };
  if (/^Fidget Toy$/i.test(source)) return { category: 'CLEARANCE', basis: 'inventory: ' + source };
  if (/^(?:Add On Products|Extended Warranty|Warranty|Shopee Adjustment|Store Use|Gunnar Credit Note Stocks|Gashapon \[Credit Note\])/i.test(source)) return { category: 'Payments, Add-ons & Shipping', basis: 'inventory: ' + source };

  // Accessory categories are platform-sensitive. Product identity wins over
  // the generic database category label.
  if (/playstation|\bps[345]\b|dualsense|psvr/i.test(text)) return { category: 'PS', basis: 'inventory platform: PlayStation' };
  if (/nintendo|\bswitch\b|joy.?con|\bnsw2\b/i.test(text)) return { category: 'SWITCH', basis: 'inventory platform: Nintendo' };
  if (/\bxbox\b/i.test(text)) return { category: 'Xbox Gaming Acc', basis: 'inventory platform: Xbox' };

  var direct = categorize(productName || '');
  if (direct !== DEFAULT_CATEGORY) return { category: direct, basis: 'inventory category refined by product-name rule' };

  if (/^(?:Peripherals|Gadget|Gadgets|Sony Audio|Stationary|Flask|Snacks|Power Bank|Tempered Glass|PC Build|Accessories|Thumbler|Gadget)/i.test(source)) return { category: 'Retail', basis: 'inventory: ' + source };
  if (/^(?:Bandai Model Kits|Banpresto|Good Smile|Bandai Tamashii|Premium Bandai|Die-Cast|Plush|Blocks|Merchandise|Electronic Pets|Candy Toys|Figurine|China Model Kits|China Action Figurine|Takaratomy|Toys|DIY|Amiibo|Keychain|Jada Toys|Beast Kingdom|Fanthful|Funism|Ensky|Hasbro|Spin Master|Manga|Fukubukuro|Gashapon|Prop Replica|Jigsaw|Cluedo|Sanrio|Palverse|Fuggler|Squishy|Candy Toys|Bandai |Bearbrick|Hot Toys|Ichiban Kuji|Joytoy|Kotobukiya|LDCX|MGL Toys|Mattel|MegaHouse|Nanoblock|NERF|Paladone|Playart|Playzone|Soap Studio|Square Enix|Threezero|Waboba|XM Studio|ZD Toys|Vinlaa|Si Meng Toys|Sunrisepop|Tamiya|Theory11|Top Joy|Jandoon|Moyu Culture|Harry Potter Wands|Minecraft Foam Toys|Kpop Official Card|Room One Cards)/i.test(source + ' ' + tags)) {
    if (/Kpop Official Card|Room One Cards|Theory11/i.test(source)) return { category: 'TCG', basis: 'inventory: ' + source };
    return { category: 'ToysVT031XH018', basis: 'inventory: ' + source };
  }

  if (/^LRG Gameboy/i.test(source)) return { category: 'SWITCH', basis: 'inventory: ' + source };
  if (/^(?:PC Game|Digital Codes|Imagine Games)/i.test(source)) return { category: 'Xbox Gaming Acc', basis: 'inventory: ' + source };
  if (/^(?:Ps3 Game|PlayStation [45] Console|Used Console \[PlayStation [45]\]|Used Console \[PlayStation)/i.test(source)) return { category: 'PS', basis: 'inventory: ' + source };
  if (/^(?:Switch2? Official Accessories|Used Console \[Switch\])/i.test(source)) return { category: 'SWITCH', basis: 'inventory: ' + source };
  if (/^(?:Retro Console|Steam Deck|Used Console \[Steam Deck\]|Used Console \[Rog Ally\]|PlayStation 4 Console)/i.test(source)) return { category: 'Retail', basis: 'inventory: ' + source };

  // Final taxonomy fallback: use the source category and tags to select one
  // of the existing output groups, but mark it for review rather than imply
  // that an unknown source category was mapped with high confidence.
  var taxonomy = source + ' ' + tags + ' ' + text;
  if (/\b(?:TCG|trading card|card sleeve|card game|booster|toploader)\b/i.test(taxonomy)) return { category: 'TCG', basis: 'REVIEW: broad card-category mapping from ' + source };
  if (/\b(?:toy|toys|figure|figurine|plush|die.?cast|model kit|blocks|merchandise|collectible|blindbox|blind box|gashapon|kuji|nendoroid|action figure)\b/i.test(taxonomy)) return { category: 'ToysVT031XH018', basis: 'REVIEW: broad toy-category mapping from ' + source };
  if (/\b(?:game|console)\b/i.test(taxonomy)) {
    if (/playstation|\bps[345]\b|psvr/i.test(taxonomy)) return { category: 'PS', basis: 'REVIEW: broad PlayStation mapping from ' + source };
    if (/nintendo|switch/i.test(taxonomy)) return { category: 'SWITCH', basis: 'REVIEW: broad Nintendo mapping from ' + source };
    return { category: 'Xbox Gaming Acc', basis: 'REVIEW: broad PC/Xbox game mapping from ' + source };
  }
  if (source) return { category: 'Retail', basis: 'REVIEW: fallback mapping; verify inventory category ' + source };
  return { category: DEFAULT_CATEGORY, basis: 'REVIEW: inventory category is blank' };
}



// =========================================================================
// SHARED TEXT HELPERS
// =========================================================================
function reduceOptionList(fullText, variation) {
  var vKey = (variation || "").split(",")[0].trim();
  if (!vKey) return fullText;
  // Only split on "/" that has real whitespace on both sides — avoids
  // corrupting compact notations like "1/6" or "1/64" scale ratios, which
  // have no surrounding spaces and aren't option separators.
  // A slash with whitespace on AT LEAST ONE side is an option separator.
  // Requiring space on both sides missed real lists written "A /B /C", which
  // then got truncated at the first slash by the fallback. Requiring space on
  // neither side is still excluded, so compact notation like "1/64" scale
  // ratios and "Chi/Eng" language tags are left intact.
  var SEP = /\s+\/\s*|\s*\/\s+/;
  if (!SEP.test(fullText)) return fullText;
  var segs = fullText.split(SEP);
  if (segs.length < 2) return fullText;

  function stripPlatform(s) {
    return s.replace(/^(PS4|PS5|NSW2?|SWITCH\s*2?|XBOX(\s+SERIES)?)\s+/i, "").trim();
  }
  var vKeyClean = stripPlatform(vKey);

  function normWords(s) {
    return s.toLowerCase().replace(/\[[^\]]*\]/g, " ").replace(/[^a-z0-9\s]/g, " ")
      .split(/\s+/).filter(function (w) { return w.length > 1 && !/^(the|and|for|of|a)$/.test(w); });
  }
  var vWords = normWords(vKeyClean);

  var matchIdx = -1, bestScore = 0, exact = false, implicit = false;
  for (var i = 0; i < segs.length; i++) {
    var seg = segs[i].trim();
    if (!seg) continue;
    if (seg.toLowerCase().indexOf(vKeyClean.toLowerCase()) !== -1 || vKeyClean.toLowerCase().indexOf(seg.toLowerCase()) !== -1) {
      matchIdx = i; exact = true; break;
    }
    var segWords = normWords(seg);
    var score = segWords.filter(function (w) { return vWords.indexOf(w) !== -1; }).length;
    if (score > bestScore) { bestScore = score; matchIdx = i; }
  }

  // No shared word anywhere — if the LAST segment has no bracketed code
  // (unlike the other listed options), treat it as the implicit "other"
  // pick: its own text (e.g. a bare character name) plus the actual
  // variation chosen are complementary, not interchangeable, so combine
  // them rather than letting one replace the other.
  if (matchIdx === -1) {
    var lastSeg = segs[segs.length - 1].trim();
    if (lastSeg && lastSeg.indexOf("[") === -1) { matchIdx = segs.length - 1; implicit = true; }
  }
  if (matchIdx === -1) return fullText;

  if (implicit) {
    var dashIdx = segs[0].lastIndexOf(" - ");
    var introPart = dashIdx !== -1 ? segs[0].slice(0, dashIdx).trim() : "";
    var connector = dashIdx !== -1 ? " - " : " ";
    return (introPart + connector + segs[matchIdx].trim() + " " + vKeyClean).replace(/\s+/g, " ").trim();
  }

  if (exact && matchIdx === 0) return segs[0].trim();

  var matched = segs[matchIdx].trim();
  var sortedVWords = vWords.slice().sort(function (a, b) { return b.length - a.length; });
  var cutPos = -1;
  for (var vw = 0; vw < sortedVWords.length; vw++) {
    var idx = matched.toLowerCase().indexOf(sortedVWords[vw]);
    if (idx !== -1) { cutPos = idx; break; }
  }
  var prefix;
  if (cutPos > 0) {
    prefix = matched.slice(0, cutPos).trim();
  } else if (matchIdx > 0) {
    var prefixMatch = segs[0].match(/^(.*?)(\S+)\s*$/);
    prefix = prefixMatch ? prefixMatch[1].trim() : "";
  } else {
    prefix = "";
  }
  return (prefix + " " + vKeyClean).replace(/\s+/g, " ").trim();
}

function stripTrailingCjkSection(s) {
  var idx = s.search(/[\u4e00-\u9fff]/);
  if (idx === -1) return s;
  var cut = s.lastIndexOf("/", idx);
  if (cut === -1) {
    cut = idx;
    var mNum = s.slice(0, idx).match(/(\S+)\s*$/);
    if (mNum && /^\d+$/.test(mNum[1])) cut = idx - mNum[1].length;
  }
  return s.slice(0, cut).replace(/\s+$/, "").trim();
}


// =========================================================================
// NAME RESOLUTION -- step 1: normalize raw text
// =========================================================================
// Applied to every product name before any rule runs: drops packaging
// add-ons entirely, strips the CJK translation block, and removes internal
// listing markers ([READY STOCK], "BB ", "Top Up (For Shipping)") that are
// not part of the product name and would otherwise block brand rules.
// Returns "" for items that should be dropped from the output completely.
function normalizeProductText(product) {
  // null means "drop this line item entirely". Distinct from returning an
  // empty string, which means "nothing left after stripping markers" — that
  // still flows on, because the fallback can recover the real name from the
  // variation field (e.g. a bare "Top Up (For Shipping)" placeholder row).
  if (/EXTRA PROTECTION PACKAGING|BUBBLE WRAP|EXTRA THICKNESS BOX/i.test(product)) {
    return null;
  }

  product = product
    // The CJK translation block always follows the language marker in this
    // business's listings — anchor there and drop everything after it if
    // CJK is present in the remainder. This catches cases per-run CJK
    // stripping can't: English words embedded inside the CJK phrase (e.g.
    // "黑旗Resynced"), and full-width punctuation (！：～・ｉ) that isn't
    // in the CJK Unified Ideographs range at all.
    .replace(/((?:CHI\/ENG|ENG\/CHI)|(?:\bENGLISH\b|\bCHINESE\b|\bJAPANESE\b|\bCHI\b|\bENG\b|\bJAP\b)\s*VERSION)[\s\S]*[\u4e00-\u9fff][\s\S]*$/i, "$1")
    .replace(/\[[^\]]*[\u4e00-\u9fff][^\]]*\]/gu, "")   // remove any [...] that contains CJK
    .replace(/\s*[\u4e00-\u9fff][\s\S]*$/gu, "")           // remove trailing CJK translation block and anything after it
    .replace(/[\u4e00-\u9fff]+\d*/gu, "")                    // remove remaining bare CJK runs (and a digit directly attached, e.g. "黑色行动7")
    .replace(/\s{2,}/g, " ")
    .trim();
  // Strip a leading bracket tag like [READY STOCK], [2026 PROMOTION], [LOCAL READY STOCK]
  // — these are internal listing annotations, not part of the product name.
  // [PO] is exempt: the Sunway Tiny rule below depends on it being present.
  product = product.replace(/^\[(?!PO\])[^\]]*\]\s*/i, "").trim();
  // Strip a leading "BB " internal blind-box listing marker — not part of
  // the brand/product name, was blocking brand-name rules from matching.
  product = product.replace(/^BB\s+(?=[A-Z])/, "").trim();
  // Strip a leading "Top Up (For Shipping...)" internal placeholder — the
  // real item info lives in variation, this is just a shipping-fee marker.
  // Matches every observed wording: "Top Up (For Shippings)",
  // "Top Up Balance Payment ONLY (For Shippings)", "Top Up For Shipping".
  // Anything between "Top Up" and the shipping phrase is payment boilerplate.
  product = product.replace(/^Top\s*Up\b[^()]*?\(?\s*For\s*Shippings?\s*\)?\s*-?\s*/i, "").trim();

  return product;
}

// =========================================================================
// NAME RESOLUTION -- step 2: ordered rule table
// =========================================================================
// Rules are tried IN ORDER; the first that returns a value wins. A rule that
// returns nothing falls through to the next — identical semantics to the long
// if/else chain this replaces, but each rule is now isolated and individually
// inspectable.
//
// ORDER IS LOAD-BEARING. Two traps, both found the hard way:
//   * Mini GT is deliberately checked BEFORE the SKU_MASTER lookup, because
//     the wanted output is the terse "Mini GT MGT..." code, not the long
//     catalogue description stored against that same SKU.
//   * Accessory brands whose titles mention a console (BJBT, GeekShare,
//     Skull & Co, DOBE, iiNE) must match BEFORE the generic platform game
//     rule near the end, or that rule treats them as games and silently
//     drops their variation and SKU.
function cleanAdminVariation(value) {
  return String(value || '').split(',').map(function (part) { return part.trim(); })
    .filter(function (part) {
      return part &&
        !/^(FULL PAYMENT|BALANCE PAYMENT|DEPOSIT|PARTIAL PAYMENT|P\.?O\.?|RANDOM COVER|NO ADD-?ON|READY STOCK|SWITCH\s*[12])(?:\b|$)/i.test(part) &&
        !/^\d+\s*YEARS?\s+WARRANTY$/i.test(part);
    })
    .join(' - ');
}

function cleanReferenceName(value) {
  return stripTrailingCjkSection(String(value || '').trim())
    .replace(/\s*\(\d+\s*\/\s*[Ss]et\)\s*$/i, '')
    .replace(/\s*\(([1-9])\)\s*$/, '')
    .replace(/\b8bitdo\b/gi, '8BitDo')
    .replace(/\bdobe\b/gi, 'Dobe')
    .replace(/\s+/g, ' ').trim();
}

function referenceLooksDescriptive(value, brandPattern) {
  var text = cleanReferenceName(value);
  return !!text && brandPattern.test(text) && text.split(/\s+/).length >= 4 &&
    !/^(?:[A-Z0-9]+[-A-Z0-9]*)$/i.test(text);
}

var PRODUCT_RULES = [
  { name: 'Beyblade bare code (Top Up row)', resolve: function (product, variation, sku, skuMap, ctx) {
    // A "Top Up (For Shipping)" row strips down to an empty product name,
    // leaving only a bare Beyblade X code (BX-/CX-/UX-) in the variation.
    // Deliberately narrow: it fires ONLY when the product name is empty and
    // the variation is nothing but one of those three code prefixes, so a
    // real listing that happens to mention a code is never caught by it.
    // The code can end up in EITHER place once the Top Up prefix is stripped:
    // in the product name ("...(For Shippings) BX-35" -> "BX-35"), or, when
    // the listing title was only the placeholder, in the variation. Accept
    // both, but stay narrow: the candidate must be nothing except the code,
    // so a real listing that merely mentions one is never caught.
    var codeBey = (product || '').trim() || (variation || '').split(',')[0].trim();
    if (/^(?:BX|CX|UX)-\d+$/i.test(codeBey)) {
      return codeBey.toUpperCase();
    }
  } },

  { name: 'Mini', resolve: function (product, variation, sku, skuMap, ctx) {
      // Mini GT is checked before SKU_MASTER: the terse "Mini GT MGT..." code
      // format is the wanted output for this brand, not the long descriptive
      // name that may be sitting in the product catalog for the same SKU.
      if (/MINI\s*GT|MINIGT/i.test(product)) {
        var code3 = (sku.match(/MGT[\d-]+[A-Z]*/i) || [])[0] || (product.match(/MGT[\d-]+[A-Z]*/i) || [])[0]
          || (sku.match(/KHMG[\d-]+/i) || [])[0] || (product.match(/KHMG[\d-]+/i) || [])[0] || "";
        if (!code3) {
          var genericCode = normalizeSku(sku);
          if (genericCode && /\d/.test(genericCode) && !/PAYMENT|DEPOSIT/i.test(genericCode)) code3 = genericCode;
        }
        code3 = code3.replace(/-$/i, "");
        return ("Mini GT " + code3).trim();
      }
  } },

  { name: 'rule', resolve: function (product, variation, sku, skuMap, ctx) {
      if (ctx.normalizedSku && skuMap[ctx.normalizedSku]) {
        return skuMap[ctx.normalizedSku];
      }
  } },

  { name: 'Fidget Toy', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/FIDGET TOY|PUSH POP SENSORY/i.test(product)) {
        return "Fidget Toy - " + variation.trim();
      }
  } },

  { name: 'Digimon Digivice', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/DIGIMON.*DIGIVICE|DIGIVICE.*DIGIMON/i.test(product)) {
        if (referenceLooksDescriptive(sku, /BANDAI|DIGIMON|DIGIVICE/i)) {
          return cleanReferenceName(sku);
        }
        var digimonBase = stripTrailingCjkSection(product).replace(/\s+/g, ' ').trim();
        var digimonVariation = cleanAdminVariation(variation);
        if (digimonVariation && digimonBase.toLowerCase().indexOf(digimonVariation.toLowerCase()) === -1) {
          digimonBase += ' - ' + digimonVariation;
        }
        return digimonBase;
      }
  } },

  { name: 'Bandai', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/BANDAI/i.test(product)) {
        // Digimon Color V-Pets have descriptive SKU references with the exact
        // version and colour. Keep that reference instead of reducing it to a
        // numeric stock code (e.g. 00072).
        if (/DIGIMON.*DIGIVICE|DIGIVICE.*DIGIMON/i.test(product) &&
            referenceLooksDescriptive(sku, /BANDAI|DIGIMON|DIGIVICE/i)) {
          return cleanReferenceName(sku);
        }
        // Model-kit shortening is applied only after the inventory database
        // confirms the category is exactly 'Bandai Model Kits'.
      }
  } },

  { name: 'Angry Shinn', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/ANGRY SHINN/i.test(product)) return product.split(" / ")[0].trim();
      if (/YOLOPARK/i.test(product)) {
        var baseYP = product.split(" / ")[0].trim();
        var vYP = variation.split(",")[0].trim();
        return (vYP && baseYP.toLowerCase().indexOf(vYP.toLowerCase()) === -1) ? baseYP + " " + vYP : baseYP;
      }
  } },

  { name: 'Zd Toys', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/ZD TOYS/i.test(product)) {
        var baseZD = product.split(" / ")[0].trim();
        var vZD = variation.split(",")[0].trim();
        if (vZD && baseZD.toLowerCase().indexOf(vZD.toLowerCase()) === -1) {
          return baseZD + " " + vZD;
        }
        return baseZD;
      }
  } },

  { name: 'Takaratomy', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/TAKARATOMY/i.test(product)) {
        var takaraBase = product.split(" / ")[0].replace(/\s+/g, " ").trim();
        // Beyblade X options carry the selected model code in the variation.
        // The variation is authoritative: replace a conflicting code already
        // present in the listing title instead of keeping the advertised code.
        if (/BEYBLADE\s+X/i.test(takaraBase)) {
          var beybladeCode = String(variation || '').match(/\b(?:BX|CX|UX)-\d{2,3}\b/i) ||
            String(sku || '').match(/\b(?:BX|CX|UX)-\d{2,3}\b/i);
          if (beybladeCode) {
            var selectedCode = beybladeCode[0].toUpperCase();
            if (/\b(?:BX|CX|UX)-\d{2,3}\b/i.test(takaraBase)) {
              takaraBase = takaraBase.replace(/\b(?:BX|CX|UX)-\d{2,3}\b/i, selectedCode);
            } else {
              takaraBase = takaraBase.replace(/(BEYBLADE\s+X)/i, '$1 ' + selectedCode);
            }
          }
        }
        return takaraBase;
      }

      if (/^BLOKEES/i.test(product)) {
        // Blokees output must retain the brand plus its model code
        // (e.g. "Blokees 75710"), never the bare number or full listing title.
        var code2 = (String(sku || '').match(/\b[A-Z]{1,6}-\d{2,8}[A-Z]?\b/i) || [])[0] ||
          (String(variation || '').match(/\b[A-Z]{1,6}-\d{2,8}[A-Z]?\b/i) || [])[0] ||
          (String(product || '').match(/\b[A-Z]{1,6}-\d{2,8}[A-Z]?\b/i) || [])[0] ||
          (String(sku || '').match(/\b7\d{4}\b/) || [])[0] ||
          (String(product || '').match(/\b7\d{4}\b/) || [])[0];
        if (code2) return "Blokees " + code2.toUpperCase();
      }
  } },

  { name: 'China Moe Town', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^CHINA MOE TOWN/i.test(product)) {
        return stripTrailingCjkSection(product).replace(/\s*\(\d+\s*\/\s*[Ss]et\)\s*$/, "").replace(/\s+\/\s*.*$/, "").trim();
      }
  } },

  { name: 'Poprace', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/POPRACE/i.test(product)) {
        var code4 = (sku.match(/PRE\d+|PR\d+/i) || [])[0] || (product.match(/PRE\d+|PR\d+/i) || [])[0];
        if (code4) return "PopRace " + code4;
      }
  } },

  { name: 'Robotime', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/ROBOTIME/i.test(product)) {
        var code5 = (sku.match(/[A-Z]{2,4}\d+[A-Z]?/i) || [])[0] || (product.match(/[A-Z]{2,4}\d+[A-Z]?/i) || [])[0];
        if (code5) return "Robotime " + code5;
      }
  } },

  { name: 'Lego', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^LEGO/i.test(product)) {
        var setNo = (variation.match(/\b\d{5}\b/) || sku.match(/\b\d{5}\b/) || product.match(/\b\d{5}\b/) || [])[0] || '';
        if (setNo) return "LEGO " + setNo;
      }
  } },

  { name: 'Motor Nuclear', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/MOTOR NUCLEAR/i.test(product)) {
        var codeMN = (product.match(/(MNP-[A-Z0-9-]+|MNZ-[A-Z0-9-]+|BV-[A-Z0-9-]+)/i) || [])[0];
        if (codeMN) return "Motor Nuclear " + codeMN;
        if (variation && variation.trim()) return "Motor Nuclear " + variation.split(",")[0].trim();
        return "Motor Nuclear";
      }
  } },

  { name: 'Cutebee', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^CUTEBEE/i.test(product)) {
        return product.split(" - ")[0].trim() + " - " + variation.split(",")[0].trim();
      }
  } },

  { name: 'Sunway Tiny', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/SUNWAY TINY/i.test(product)) {
        var mSW = product.match(/^\[PO\]\s*(.*?\))\s*\//i);
        var baseSW = mSW ? mSW[1] : product.replace(/^\[PO\]\s*/i, "");
        baseSW = baseSW.replace(/\s*\([^)]*\)\s*$/, "").trim();
        return (baseSW + " " + variation.split(",")[0].trim()).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Maxsoft Console', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/MAXSOFT CONSOLE/i.test(product)) {
        return ("Nintendo Switch OLED Maxsoft Console " + variation.split(",")[0].trim()).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Switch', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/SWITCH\s*2\s*CONSOLE/i.test(product) || (/NINTENDO SWITCH\s*2/i.test(product) && /CONSOLE/i.test(product))) {
        return ("Nintendo Switch 2 Console " + variation.replace(/,/g, " ").trim()).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Moshow', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^MOSHOW/i.test(product)) {
        return stripTrailingCjkSection(product).split(" / ")[0].trim();
      }
  } },

  { name: 'Trimui', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^TRIMUI/i.test(product)) {
        var trimuiBase = product.split(/\s+with\s+|\s*\(|\s*\|/i)[0].trim();
        return (trimuiBase + " " + variation.replace(/,/g, " ").replace(/\s+/g, " ").trim()).trim();
      }
  } },

  { name: 'R36H', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^R36H\b/i.test(product)) {
        return ("R36H Retro Handheld Gaming Console - " + variation.replace(/,/g, " ").replace(/\s+/g, " ").trim()).trim();
      }
  } },

  { name: 'Ayaneo', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^AYANEO/i.test(product)) {
        var ayaBase = product.split(/\s+with\s+|\s*\(|\s+-\s+/i)[0].trim();
        var vAya = variation.split(",")[0].trim();
        return (ayaBase + (vAya && ayaBase.toLowerCase().indexOf(vAya.toLowerCase()) === -1 ? " " + vAya : "")).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Spin Master', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^SPIN MASTER/i.test(product)) {
        return product.split(" / ")[0].trim();
      }
  } },

  { name: 'Danboard', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^DANBOARD/i.test(product)) {
        var vD = variation.trim();
        return /^DANBOARD/i.test(vD) ? vD : "Danboard " + vD;
      }
  } },

  { name: 'Vt', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^VT\b|GUNPLA ACCESSORIES/i.test(product)) {
        var codeVT = (String(sku || '').match(/\b[A-Z]{1,5}-[A-Z0-9]{2,}(?:-[A-Z0-9]+)*\b/i) || String(product || '').match(/\b[A-Z]{1,5}-[A-Z0-9]{2,}(?:-[A-Z0-9]+)*\b/i) || [])[0];
        if (codeVT) return 'VT ' + codeVT.toUpperCase();
        var vVT = variation.split(",")[0].trim();
        if (vVT && !/^(Full Payment|Balance Payment)/i.test(vVT)) return "VT " + vVT;
      }
  } },

  { name: 'Model Kits', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^MODEL KITS? ACCESSORIES/i.test(product)) {
        var codeMK = normalizeSku(sku) || (product.match(/\bWS-\d+\b/i) || [])[0] || "";
        return (codeMK + " " + variation.split(",")[0].trim()).trim();
      }
  } },

  { name: 'Majorette', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^MAJORETTE/i.test(product)) {
        return product.split(" - ")[0].trim() + " - " + variation.split(",")[0].trim();
      }
  } },

  { name: 'Jisulife', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^JISULIFE/i.test(product)) {
        return product.split(" - ")[0].trim() + " - " + variation.split(",")[0].trim();
      }
  } },

  { name: 'Funism', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^FUNISM/i.test(product)) {
        var vFun = variation.split(",")[0].trim();
        return vFun ? product.split(" - ")[0].trim() + " - " + vFun : "Funism";
      }
  } },

  { name: 'G923 Racing Wheel', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/\bG(?:29|920|923)\b/i.test(product) && /RACING|STEERING WHEEL/i.test(product)) {
        var vG923 = cleanAdminVariation(variation);
        var selectedWheel = (String(variation || '').match(/\b(G29|G920|G923)\b/i) || [])[1] || 'G923';
        selectedWheel = selectedWheel.toUpperCase();
        var hasShifter = /SHIFTER/i.test(vG923) || (!vG923 && /SHIFTER/i.test(product));
        return 'Logitech ' + selectedWheel + ' Driving Force Racing Wheel + Pedals' + (hasShifter ? ' + Shifter' : '');
      }
  } },

  { name: 'Moza', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^MOZA\b/i.test(product)) {
        var codeMoza = normalizeSku(sku);
        return codeMoza ? product.split(" ").slice(0, 2).join(" ") + " " + codeMoza : product.split(" - ")[0].trim();
      }
  } },

  { name: 'Omelet Gaming', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^OMELET GAMING/i.test(product)) {
        return product.split(" - ")[0].trim() + " " + variation.split(",")[0].trim();
      }
  } },

  { name: 'Lucky Fox', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^LUCKY FOX/i.test(product)) {
        var vLF = variation.split(",")[0].trim();
        return vLF ? product.split(" [")[0].trim() + " " + vLF : product.split(" [")[0].trim();
      }
  } },

  { name: 'Wekki Block', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^WEKKI BLOCK/i.test(product)) {
        return product.split(" - ")[0].trim() + " - " + variation.split(",")[0].trim();
      }
  } },

  { name: 'Gsc', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/GSC/i.test(product)) {
        return product.replace(/\/.*$/, "").replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'PMW', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^PMW\b/i.test(product)) {
        var pmwBase = stripTrailingCjkSection(product).replace(/\s+/g, ' ').trim();
        var pmwVariation = cleanAdminVariation(variation);
        if (pmwVariation && pmwBase.toLowerCase().indexOf(pmwVariation.toLowerCase()) === -1) {
          var pmwWords = pmwVariation.toLowerCase().match(/[a-z0-9]+/g) || [];
          var pmwMissing = pmwWords.filter(function (word) {
            return word.length > 1 && pmwBase.toLowerCase().indexOf(word) === -1;
          });
          if (pmwMissing.length) pmwBase += ' - ' + pmwVariation;
        }
        return pmwBase.replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Brook', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/BROOK/i.test(product)) {
        var color = variation.split(",")[0].trim();
        var codeBrk = (sku.match(/[A-Z0-9-]+/) || [])[0];
        var baseBrk = stripTrailingCjkSection(product).trim();
        return (baseBrk + (color ? " - " + color : "") + (codeBrk ? " (" + codeBrk + ")" : "")).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Microsoft Xbox', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/MICROSOFT XBOX/i.test(product) && /CONTROLLER/i.test(product)) {
        return "Microsoft Xbox Controller " + variation;
      }
  } },

  { name: 'Geekshare', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^GEEKSHARE/i.test(product)) {
        var codeGS = normalizeSku(sku);
        var vGS = cleanAdminVariation(variation);
        var titleGS = stripTrailingCjkSection(product).replace(/\s+/g, ' ').trim();
        // If the variation is present in the listing, retain the selected
        // listing title and append its model code. If it is absent from a
        // multi-option title, do not accidentally select the final option.
        // Try selecting the matching option from a slash-separated listing
        // BEFORE checking for a substring: the selected option may appear
        // inside a list that still contains several unselected alternatives.
        var reducedGS = vGS ? reduceOptionList(titleGS, vGS) : titleGS;
        if (vGS && reducedGS !== titleGS) {
          return 'GeekShare ' + vGS + (codeGS ? ' (' + codeGS + ')' : '');
        }
        if (vGS && titleGS.toLowerCase().indexOf(vGS.toLowerCase()) !== -1) {
          return (titleGS + (codeGS && titleGS.toUpperCase().indexOf(codeGS.toUpperCase()) === -1 ? ' (' + codeGS + ')' : '')).replace(/\s+/g, ' ').trim();
        }
        // A variation that shares meaningful words with the listing is a
        // more precise label for that same item (e.g. Hello Kitty Box Set).
        var wordsGS = vGS.toLowerCase().match(/[a-z0-9]+/g) || [];
        var overlapGS = wordsGS.filter(function (w) { return w.length > 2 && titleGS.toLowerCase().indexOf(w) !== -1; }).length;
        if (overlapGS >= 2) {
          return (titleGS + (codeGS && titleGS.toUpperCase().indexOf(codeGS.toUpperCase()) === -1 ? ' (' + codeGS + ')' : '')).replace(/\s+/g, ' ').trim();
        }
        if (vGS && (titleGS.match(/\//g) || []).length >= 2) {
          return ('GeekShare ' + vGS + (codeGS ? ' (' + codeGS + ')' : '')).replace(/\s+/g, ' ').trim();
        }
        if (vGS) titleGS += ' - ' + vGS;
        return (titleGS + (codeGS && titleGS.toUpperCase().indexOf(codeGS.toUpperCase()) === -1 ? ' (' + codeGS + ')' : '')).replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Bjbt', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^BJBT/i.test(product)) {
        var codeBJ = normalizeSku(sku);
        if (codeBJ) return "BJBT " + codeBJ;
      }
  } },

  { name: 'Skull & Co', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^SKULL & CO/i.test(product)) {
        if (referenceLooksDescriptive(sku, /SKULL\s*&\s*CO/i)) return cleanReferenceName(sku);
        var baseSK = product.replace(/\s*-\s*$/, '').split(' - ')[0].trim();
        var vSK = cleanAdminVariation(variation);
        var codeSK = normalizeSku(sku);
        var resultSK = baseSK;
        if (vSK && resultSK.toLowerCase().indexOf(vSK.toLowerCase()) === -1) resultSK += ' - ' + vSK;
        if (codeSK) resultSK += ' (' + codeSK + ')';
        return resultSK.replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Dobe', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/DOBE/i.test(product)) {
        var modelDobe = (String(sku || '').match(/\[([A-Z0-9]+-[A-Z0-9]+)\]/i) || String(variation || '').match(/\b([A-Z0-9]+-[A-Z0-9]+)\b/i) || String(product || '').match(/\[([A-Z0-9]+-[A-Z0-9]+)\]/i) || [])[1] || '';
        if (!modelDobe) modelDobe = normalizeSku(sku);
        var platformDobe = inventoryPlatformLabel(product + ' ' + variation);
        if (modelDobe) return 'Dobe' + (platformDobe ? ' ' + platformDobe : '') + ' ' + modelDobe.toUpperCase();
      }
  } },

  { name: 'Nintendo Switch Sports Bundle', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^Switch\s+Switch\s+Sports Bundle/i.test(product) ||
          (/^Switch\s+Sports Bundle/i.test(product) && /SPORTS BUNDLE/i.test(sku))) {
        if (referenceLooksDescriptive(sku, /NINTENDO SWITCH|SWITCH SPORTS/i)) return cleanReferenceName(sku);
        return stripTrailingCjkSection(product).replace(/^Switch\s+Switch\s+/i, 'Switch ')
          .replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Nintendo Switch Software', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^(?:NSW\s+)?Switch\s+(?!2\s+Console)/i.test(product) && /CHI\/ENG|ENG\/CHI|ANIMAL CROSSING|MOVING OUT|SUPER MARIO PARTY|TOMB RAIDER|\bUS\b|\bEU\b/i.test(product + ' ' + variation)) {
        var titleSW = stripTrailingCjkSection(product)
          .replace(/^NSW\s+/i, '')
          .replace(/^Switch\s+Switch\s+/i, 'Switch ')
          .replace(/^Switch\s+/i, 'Switch ')
          .replace(/\s+/g, ' ').trim();
        var combinedSW = product + ' ' + variation + ' ' + sku;
        if (/CHI\s*\/\s*ENG|ENG\s*\/\s*CHI|CHINESE\s*(?:\+|AND|\/)\s*ENGLISH|中文\s*\+\s*ENGLISH|中文.*ENGLISH/i.test(combinedSW) && !/CHI\/ENG/i.test(titleSW)) titleSW += ' Chi/Eng';
        else if (/\bENG(?:LISH)?\b/i.test(variation) && !/\bENG\b/i.test(titleSW)) titleSW += ' Eng';
        var regionSW = (combinedSW.match(/\b(R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\b/i) || [])[1];
        if (regionSW && !new RegExp('\\b' + regionSW + '\\b', 'i').test(titleSW)) titleSW += ' ' + regionSW.toUpperCase();
        var vSW = cleanAdminVariation(variation)
          .replace(/[\u4e00-\u9fff]+/g, ' ')
          .replace(/CHINESE\s*(?:\+|AND|\/)\s*ENGLISH/ig, ' ')
          .replace(/\+\s*ENGLISH\b/ig, ' ')
          .replace(/\b(?:R[0-9]|US|EU|AS|ASIA|UK|JP|AU|AUS)\b/ig, ' ')
          .replace(/\s+/g, ' ').replace(/^[-,\s]+|[-,\s]+$/g, '').trim();
        if (vSW && titleSW.toLowerCase().indexOf(vSW.toLowerCase()) === -1) {
          var tokensSW = vSW.toLowerCase().match(/[a-z0-9]+/g) || [];
          var meaningfulSW = tokensSW.filter(function (w) { return w.length > 1; });
          var presentSW = meaningfulSW.filter(function (w) { return titleSW.toLowerCase().indexOf(w) !== -1; }).length;
          if (meaningfulSW.length && presentSW < meaningfulSW.length) titleSW += ' - ' + vSW;
        }
        return titleSW.replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Hori', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^HORI\b/i.test(product)) {
        if (referenceLooksDescriptive(sku, /HORI/i)) return cleanReferenceName(sku);
        var codeH = (String(sku || '').match(/\[([A-Z0-9]+-[A-Z0-9]+)\]/i) || String(sku || '').match(/\b(NSX-[A-Z0-9]+)\b/i) || String(product || '').match(/\b(NSX-[A-Z0-9]+)\b/i) || [])[1] || '';
        var platformH = inventoryPlatformLabel(product + ' ' + variation);
        if (codeH) return 'Hori' + (platformH ? ' ' + platformH : '') + ' ' + codeH.toUpperCase();
        var vH = cleanAdminVariation(variation);
        return ('Hori ' + vH).replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Iine', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^IINE/i.test(product)) {
        var codeI = (String(sku || '').match(/\[([A-Z0-9]+-?[A-Z0-9]+)\]/i) || String(variation || '').match(/\b(L\d+)\b/i) || String(product || '').match(/\[(L\d+)\]/i) || [])[1] || normalizeSku(sku);
        var platformI = inventoryPlatformLabel(product + ' ' + variation + ' ' + sku);
        if (codeI) return 'iiNE' + (platformI ? ' ' + platformI : '') + ' ' + codeI.toUpperCase();
        return stripTrailingCjkSection(product).replace(/\bNSW\s+(?=Nintendo Switch\b)/i, '');
      }
  } },

{ name: 'Clearedge', resolve: function (product, variation, sku, skuMap, ctx) {
    if (/CLEAREDGE/i.test(product)) {

      // Prefer the full product description in SKU Reference No.
      // It is more descriptive than the selected variation.
      var skuName = String(sku || '').trim();

      if (/^CLEAREDGE\b/i.test(skuName) && skuName.length >= 20) {
        return stripTrailingCjkSection(skuName)
          .replace(/\s+/g, ' ')
          .trim();
      }

      // Existing fallback rules when SKU Reference No. is not descriptive.
      if (/COREBINDER/i.test(product)) {
        return "ClearEdge CoreBinder " + variation;
      }

      if (/ZIPBINDER/i.test(product)) {
        return "ClearEdge ZipBinder " + variation;
      }

      if (/TOPLOADER/i.test(product)) {
        var sizeM = (product + " " + variation).match(/(\d+)\s*\/?\s*pcs?/i);
        return "ClearEdge Premium Toploader (" + (sizeM ? sizeM[1] : "25") + "pcs)";
      }

      if (/TOP 100/i.test(product)) {
        return "ClearEdge Premium " + variation;
      }

      if (/BLACK \(50PCS\)/i.test(variation)) {
        return "ClearEdge Premium Black (50pcs)";
      }

      return "ClearEdge Premium " + variation;
    }
} },


  { name: 'PlayStation DualSense', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/DUALSENSE/i.test(product) && /PS5|PLAYSTATION 5/i.test(product)) {
        var baseDS = referenceLooksDescriptive(sku, /DUALSENSE/i) ? cleanReferenceName(sku) : stripTrailingCjkSection(product).replace(/\s+/g, ' ').trim();
        baseDS = baseDS.replace(/\s*-?\s*CFI-ZCT[0-9]G(?:\s+[A-Z0-9]+)?\b/ig, '').replace(/\s+/g, ' ').trim();
        if (/PRE-?OWNED/i.test(variation) && !/PRE-?OWNED/i.test(baseDS)) baseDS += ' (Pre-Owned)';
        return baseDS;
      }
  } },

  { name: 'Bitdo', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/8BITDO/i.test(product)) {
        if (/TRAVEL CASE|CONTROLLER CASE/i.test(product)) {
          var caseCodeBit = (String(sku || '').match(/\b[A-Z0-9]{5,10}\b/i) || String(product || '').match(/\b[A-Z0-9]{5,10}\b/i) || [])[0] || normalizeSku(sku) || '';
          var caseNameBit = stripTrailingCjkSection(product).replace(/\s+/g, ' ').trim();
          return caseNameBit + (caseCodeBit ? ' (' + caseCodeBit + ')' : '');
        }
        if (referenceLooksDescriptive(sku, /8BITDO/i)) return cleanReferenceName(sku);
        var vBit = cleanAdminVariation(variation);
        var modelBit = /PRO\s*3/i.test(vBit) ? 'Pro 3' : (/PRO\s*2/i.test(vBit) ? 'Pro 2' : '');
        if (modelBit) {
          var colorBit = vBit.replace(/^.*?PRO\s*\d+\s*-?\s*/i, '').trim();
          return ('8BitDo ' + modelBit + ' Wireless Controller' + (colorBit ? ' - ' + colorBit : '') + (normalizeSku(sku) ? ' (' + normalizeSku(sku) + ')' : '')).replace(/\s+/g, ' ').trim();
        }
        if (vBit) return ('8BitDo ' + vBit).replace(/\s+/g, ' ').trim();
        return stripTrailingCjkSection(product).replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Anbernic', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/ANBERNIC/i.test(product) && /MEMORY CARD/i.test(sku)) {
        return cleanDbMultilineText(sku);
      }
  } },

  { name: 'Arzopa', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/ARZOPA/i.test(product)) {
        if (/CARRYING BAG/i.test(product)) return 'Arzopa Acc Bag';
        var allAr = product + ' ' + variation + ' ' + sku;
        var modelAr = (allAr.match(/\bZ[0-9][A-Z0-9]{2,5}\b/i) || [])[0] || '';
        var colourAr = (String(variation || '') + ' ' + product + ' ' + sku).match(/\b(Space\s+Gray|Space\s+Grey|Gray|Grey|Black|White|Silver|Pink|Blue)\b/i);
        if (modelAr) return 'Arzopa ' + modelAr.toUpperCase() + (colourAr ? ' ' + colourAr[1].replace(/\s+/g, ' ') : '');
        var resultAr = stripTrailingCjkSection(product).replace(/\s+/g, ' ').trim().replace(/^ARZOPA/i, 'Arzopa');
        var vAr = cleanAdminVariation(variation);
        if (vAr && resultAr.toLowerCase().indexOf(vAr.toLowerCase()) === -1 && /\b(gray|grey|black|white|silver|pink|blue)\b/i.test(vAr)) resultAr += ' ' + vAr;
        return resultAr.replace(/\s+/g, ' ').trim();
      }
  } },
  { name: 'Corsair', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^CORSAIR\b/i.test(product)) {
        return product.split(" / ")[0].trim();
      }
  } },

  { name: 'Jsaux', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^JSAUX\b/i.test(product)) {
        return product.split(" - ")[0].trim() + " - " + variation.trim();
      }
  } },

  { name: 'Tamagotchi', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/TAMAGOTCHI/i.test(product)) {
        if (/PARADISE/i.test(product)) {
          return "Tamagotchi Paradise - " + variation.split(",")[0].trim();
        }
        if (/HUGMY/i.test(product)) {
          var seriesMatchT = product.match(/Hugmy\s+(.+?)\s*\(/i);
          var seriesNameT = seriesMatchT ? seriesMatchT[1].trim() : "";
          return ("Tamagotchi Hugmy " + seriesNameT + " " + variation.split(",")[0].trim()).replace(/\s+/g, " ").trim();
        }
        if (/ORIGINAL TAMAGOTCHI/i.test(product)) {
          var originalTamaName = ("Original Tamagotchi " + cleanAdminVariation(variation)).replace(/\s+/g, " ").trim();
          // Prefer the SKU Reference No. item identifier. References may be
          // descriptive titles ending in a numeric code (e.g. (90898)), or
          // compact catalog codes such as TMG90033 (Gen 1).
          var tamaCode = String(sku || '').match(/\b(TMG\d{4,6})\b/i) ||
            String(sku || '').match(/\((\d{5,6})\)\s*$/);
          if (tamaCode && originalTamaName.toUpperCase().indexOf(tamaCode[1].toUpperCase()) === -1) {
            originalTamaName += " (" + tamaCode[1].toUpperCase() + ")";
          }
          return originalTamaName;
        }
        if (/ORIGINAL COLOR COLLECTION/i.test(product)) {
          return "Tamagotchi Original Color Collection - " + variation.split(",")[0].trim();
        }
        return ("Tamagotchi " + variation.split(",")[0].trim()).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Earfun', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/EARFUN/i.test(product)) {
        if (/AIR PRO 4/i.test(product)) return "EarFun Air Pro 4+ " + variation;
        if (/UBOOM L/i.test(product)) return "EarFun UBOOM L";
        if (/UBOOM SLIM/i.test(product)) return "EarFun UBOOM Slim";
      }
  } },

  { name: 'Royal Kludge', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/ROYAL KLUDGE|^RK\b/i.test(product)) {
        var modelRK = "";
        var mRK = product.match(/\bR\d+\b/i);
        if (mRK) modelRK = mRK[0].toUpperCase();
        var colorRK = variation.replace(/,/g, " ").replace(/\s+/g, " ").trim();
        return ["Royal Kludge", modelRK, "Wired Keyboard -", colorRK].filter(Boolean).join(" ");
      }
  } },

  { name: 'R36 Ultra', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/R36 ULTRA|R36S/i.test(product)) {
        var variationClean = variation.replace(/^R36S\s*/i, "").replace(/^R36\s*/i, "").replace(/,/g, " ").replace(/\s+/g, " ").trim();
        if (/R36 ULTRA/i.test(product)) return "R36 Ultra Retro - " + variationClean;
        if (/R36S/i.test(product)) return "R36S Retro - " + variationClean;
      }
  } },

  { name: 'Miyoo', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^MIYOO/i.test(product)) {
        var miyooBase = product.split(/\s+with\s+|\s*\(/i)[0].trim();
        return (miyooBase + " " + variation.replace(/,/g, " ").replace(/\s+/g, " ").trim()).trim();
      }
  } },

  { name: 'Pokemon Tcg', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/POKEMON TCG/i.test(product)) {
        // Premium Deck Set: use the complete SKU Reference No.
        // when it contains the full product title and edition.
        if (/PREMIUM DECK SET/i.test(product)) {
          var deckSku = String(sku || '').trim();

          if (
            /^Pokemon/i.test(deckSku) &&
            /Premium Deck Set/i.test(deckSku) &&
            /\bM6A\b/i.test(deckSku)
          ) {
            return stripTrailingCjkSection(deckSku)
              .replace(/\s+/g, ' ')
              .trim();
          }

          // Fallback if the SKU Reference No. is not descriptive.
          var deckVariation = variation.split(',')[0].trim();
          return ("Pokemon TCG " + deckVariation)
            .replace(/\s+/g, ' ')
            .trim();
        }

        var srcP = sku + " " + variation + " " + product;
        var langP = "";
        if (/SIMPLIFIED CHINESE/i.test(srcP) || /TRADITIONAL CHINESE/i.test(srcP) || /\bCHINESE\b/i.test(srcP) ||
            /\bCHI\b/i.test(srcP) || /CSV\d+/i.test(srcP) || /CBB\d+/i.test(srcP) || /[\u4e00-\u9fff]/.test(srcP)) {
          langP = "Chi";
        } else if (/\[JAPAN/i.test(srcP) || /\bJAP\b/i.test(srcP)) {
          langP = "Jap";
        } else {
          langP = "Eng";
        }
        var codeP = "";
        var mP = srcP.match(/(SVP\d+|SWSH\d+|SV\d+(?:\.\d+)?[A-Z]?|CS\d+[A-Z]*|S\d+[A-Z]+|ME\d+(?:\.\d+)?|CSV\d+(?:\.\d+)?[A-Z]?|CBB\d+(?:\.\d+)?[A-Z]?|M\d+[A-Z]?)/i);
        if (mP) {
          codeP = mP[1].toUpperCase();
          if (/\bSLIM\b/i.test(sku) || /\bSLIM\b/i.test(variation)) codeP += " Slim";
        }
        // Shopee's Chinese option label and abbreviated SKU both identify the
        // 30th Anniversary booster pack, although neither contains a normal set code.
        if (/30THC/i.test(sku) || /30周年庆典/.test(variation)) {
          codeP = "30th Anniversary";
        }
        if (/TACTICS DECK SERIES/i.test(product)) {
          var baseTD = stripTrailingCjkSection(product);
          var pickTD = sku.trim() || variation.split(",")[0].trim();
          var dashIdxTD = baseTD.indexOf(" - ");
          var prefixTD = dashIdxTD !== -1 ? baseTD.slice(0, dashIdxTD).trim() : baseTD;
          if (pickTD) return (prefixTD + " - " + pickTD).replace(/\s+/g, " ").trim();
          return baseTD;
        }
        if (/SURPRISE BOX COLLECTION|SUPER PREMIUM COLLECTION/i.test(product)) {
          var pickSPC = variation.split(",")[0].trim();
          if (/SURPRISE/i.test(pickSPC) && !/COLLECTION/i.test(pickSPC)) pickSPC += " Collection";
          else if (/PREMIUM/i.test(pickSPC) && !/COLLECTION/i.test(pickSPC)) pickSPC += " Collection";
          return ("Pokemon TCG " + langP + " " + pickSPC).replace(/\s+/g, " ").trim();
        }
        if (/EMPTY MINI TIN/i.test(product)) return "Pokemon TCG " + langP + " Empty Mini Tins - " + variation.trim();
        if (/PROMO CARD/i.test(product)) return "Pokemon TCG Promo Card " + (codeP || variation.trim());
        if (/DICE MAKER/i.test(product)) return "Pokemon TCG " + langP + " Dice Maker Set " + codeP;
        if (/ELITE TRAINER BOX/i.test(product)) return "Pokemon TCG " + langP + " Elite Trainer Box " + codeP;
        if (/BOOSTER BUNDLE/i.test(product)) {
        // The SKU distinguishes the actual item: a sealed "Booster Bundle" vs
        // a "Bundle Empty Box" (the empty packaging, sold separately). Both sit
        // under the same listing, so trust the SKU over the listing title.
        var typeBB = /EMPTY BOX/i.test(sku + " " + variation) ? "Bundle Empty Box" : "Booster Bundle";
        return ("Pokemon TCG " + langP + " " + typeBB + " " + codeP).replace(/\s+/g, " ").trim();
      }
        if (/BOOSTER BOX/i.test(product)) return "Pokemon TCG " + langP + " Booster Box " + codeP;
        if (/BOOSTER PACK/i.test(product)) return "Pokemon TCG " + langP + " Booster Pack " + codeP;
      }
  } },

  { name: 'Palworld Tcg', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/PALWORLD TCG/i.test(product)) {
        var langPW = /\[ENGLISH\]/i.test(product) ? "Eng" : (/\[JAPAN\]/i.test(product) ? "Jap" : (/\[CHINESE\]/i.test(product) ? "Chi" : ""));
        if (/TRIAL DECK/i.test(product)) {
          var variantPW = /GREEN/i.test(variation) ? "Green・Purple" : (/RED/i.test(variation) ? "Red・Blue" : variation.split(",")[0].trim());
          return "PalWorld TCG " + langPW + " Trial Deck - Dawn of Palpagos " + variantPW;
        }
        if (/BOOSTER CASE/i.test(product)) return "PalWorld TCG " + langPW + " Booster Case - Dawn of Palpagos";
        if (/BOOSTER BOX/i.test(product)) return "PalWorld TCG " + langPW + " Booster Box - Dawn of Palpagos";
      }
  } },

  { name: 'Gundam Card Game', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/GUNDAM CARD GAME/i.test(product)) {
        var srcG = product + " " + variation + " " + sku;
        var langG = /\[JAPAN\]|\bJAP\b/i.test(srcG) ? "Jap" : (/\[CHINESE\]|\bCHI\b|[\u4e00-\u9fff]/i.test(srcG) ? "Chi" : (/\[ENGLISH\]|\bENG\b/i.test(srcG) ? "Eng" : "Jap"));
        var typeG = "";
        if (/PREMIUM SET/i.test(product)) typeG = "Premium Set";
        else if (/BOOSTER PACK/i.test(product)) typeG = "Booster Pack";
        else if (/BOOSTER BOX/i.test(product)) typeG = "Booster Box";
        else if (/STARTER DECK/i.test(product)) typeG = "Starter Deck";
        else typeG = "Starter Deck"; // unrecognized type text — safest default, matches original behavior
        return "Gundam Card Game TCG " + langG + " " + typeG + " - " + variation.trim();
      }
  } },

  { name: 'Hololive', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/HOLOLIVE/i.test(product) && /BOOSTER BOX/i.test(product)) {
        return "Hololive TCG Booster Box " + variation.trim();
      }
  } },

  { name: 'Hololive', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/HOLOLIVE/i.test(product) && /BOOSTER PACK/i.test(product)) {
        var srcHP = product + " " + variation + " " + sku;
        var langHP = /\[JAPAN\]|\bJAP\b/i.test(srcHP) ? "Jap" : (/\[CHINESE\]|\bCHI\b|[\u4e00-\u9fff]/i.test(srcHP) ? "Chi" : "Eng");
        return "Hololive TCG " + langHP + " " + variation.trim() + " Booster Pack";
      }
  } },

  { name: 'Card Case', resolve: function (product, variation, sku, skuMap, ctx) {
      // Digimon TCG (booster boxes/packs) — distinct from the Digivice toy line below.
      // Word order here is code-before-type (matches how these are actually labeled),
      // unlike RiftBound/One Piece which are type-before-code.
      if (!/CARD CASE/i.test(product) && (/DIGIMON TCG/i.test(product) || (/DIGIMON/i.test(product) && /BOOSTER BOX|BOOSTER PACK/i.test(product) && /\[JAPAN\]|\[ENGLISH\]|\[CHINESE\]/i.test(product)))) {
        var langD = /\[JAPAN\]/i.test(product) ? "Jap" : (/\[ENGLISH\]/i.test(product) ? "Eng" : (/\[CHINESE\]/i.test(product) ? "Chi" : ""));
        var typeD = /BOOSTER BOX/i.test(product) ? "Booster Box" : (/BOOSTER PACK/i.test(product) ? "Booster Pack" : "");
        var srcD = variation + " " + product + " " + sku;
        var codeD = "";
        var mDg = srcD.match(/(BT\d+|EX\d+|AD\d+|ST\d+|PB-\d+)/i);
        if (mDg) codeD = mDg[1].toUpperCase();
        return ["Digimon TCG", langD, codeD, typeD].filter(Boolean).join(" ");
      }
  } },

  { name: 'Digimon D-Ark Tamers Digivice', resolve: function (product, variation, sku, skuMap, ctx) {
      // Digimon collectible hardware — two distinct product lines, each keeps its
      // own prefix and just appends the chosen color/variant.
      if (/DIGIMON D-ARK TAMERS DIGIVICE/i.test(product)) {
        return "Digimon D-Ark Tamers Digivice 25th Color Evolution " + variation.split(",")[0].trim();
      }
  } },

  { name: 'Digimon Digivice Color Digital Monster', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/DIGIMON DIGIVICE COLOR DIGITAL MONSTER/i.test(product)) {
        return "Digimon Digivice Color Digital Monster " + variation.split(",")[0].trim();
      }
  } },

  { name: 'Qubecarz', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/QUBECARZ/i.test(product)) {
        return "QubeCarz " + normalizeSku(sku);
      }
  } },

  { name: 'Arkaizo', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/ARKAIZO/i.test(product)) {
        return "ARKaizo Almost Real " + normalizeSku(sku);
      }
  } },

  { name: 'Viture', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/VITURE/i.test(product) && /USB-?C/i.test(product)) {
        return "VITURE USB-C to Glasses " + variation.trim();
      }
  } },

  { name: 'Voicsky', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/VOICSKY/i.test(product)) {
        var baseV = product.split(" - ")[0].replace(/\([^)]*\)/g, "").trim();
        var vParts = variation.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
        return (baseV + " - " + vParts.join(" - ") + " " + normalizeSku(sku)).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Aula', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^AULA\b/i.test(product)) {
        if (/^AULA\b/i.test(sku) && sku.length > 20) return cleanReferenceName(sku);
        var aulaBase = stripTrailingCjkSection(product).replace(/\s+/g, ' ').trim();
        var aulaVariation = cleanAdminVariation(variation);
        if (aulaVariation) {
          var aulaParts = aulaVariation.split(' - ');
          aulaParts.forEach(function (part) {
            if (part && aulaBase.toLowerCase().indexOf(part.toLowerCase()) === -1) aulaBase += ' ' + part;
          });
        }
        var aulaCode = (String(sku || '').match(/\b[A-Z0-9]{5,}(?:-[A-Z0-9]+)+\b/i) || [])[0] || '';
        if (aulaCode && aulaBase.toUpperCase().indexOf(aulaCode.toUpperCase()) === -1) aulaBase += ' ' + aulaCode.toUpperCase();
        return aulaBase.replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Keychron', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^KEYCHRON/i.test(product)) {
        return (product.split(" - ")[0].replace(/\([^)]*\)/g, "").trim() + " - " + variation.split(",")[0].trim() + " " + normalizeSku(sku)).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Hobby', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^MR\s?HOBBY/i.test(product)) {
        return ("Mr Hobby " + variation.split(",")[0].trim()).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Toploader Top Loader Card', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^TOPLOADER TOP LOADER CARD/i.test(product)) {
        return ("Toploader Top Loader Card 35PT " + variation.split(",")[0].trim()).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Kayou', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/KAYOU/i.test(product)) {
        return product.replace(/\[.*?\]/g, "").replace(/\s+\/\s*.*$/, "").replace(/[\s\-–—]+$/, "").trim();
      }
  } },

  { name: 'Topps UEFA Mini Booster Tin', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/TOPPS TCG UEFA CHAMPIONS LEAGUE MATCH ATTAX/i.test(product) && /MINI BOOSTER TIN/i.test(product)) {
        var yearUEFA = (product.match(/\b\d{4}\/\d{2}\b/) || [])[0] || '';
        var pickUEFA = (variation || '').split(',')[0].trim();
        return ('Topps TCG UEFA Champions League Match Attax ' + yearUEFA + ' Mini Booster Tin - ' + pickUEFA)
          .replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Topps Premier League Super Tin', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/TOPPS TCG PREMIER LEAGUE PREMIUM/i.test(product) && /SUPER TIN/i.test(product)) {
        // Keep the selected tin number; Random Cover is a packaging attribute,
        // not a separate product variant.
        var pickPL = (variation || '').split(',')[0].trim();
        pickPL = pickPL.replace(/,?\s*Random Cover\s*$/i, '').trim();
        pickPL = pickPL.replace(/(Super Tin)#(\d+)/i, '$1 #$2');
        return ('Topps TCG Premier League Premium ' + pickPL).replace(/\s+/g, ' ').trim();
      }
  } },

  { name: 'Topps Tcg U', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/TOPPS TCG U[A-Z]{1,4} MATCH ATTAX/i.test(product)) {
        var leagueM = product.match(/TOPPS TCG (U[A-Z]{1,3}) MATCH ATTAX/i);
        var league = leagueM ? leagueM[1].toUpperCase() : "";
        var yearM = product.match(/\d{4}\/\d{2}/);
        var yr = yearM ? yearM[0] : "";
        // The SKU field is often more precisely formatted ("Mega Tin - Goliath")
        // than variation ("Mega Tin Goliath") — prefer it when it's at least as
        // descriptive, since it's already the specific choice, not a re-listing
        // of every option.
        var pickTA = (sku && sku.trim() && sku.trim().length >= variation.trim().length) ? sku.trim() : variation.trim();
        if (/MINI TIN/i.test(product)) {
          return ("Topps TCG " + league + " Match Attax " + yr + " Mini Tin - " + pickTA).replace(/\s+/g, " ").trim();
        }
        return ("Topps TCG " + league + " Match Attax " + yr + " - " + pickTA).replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Hori', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/^HORI\b/i.test(product)) {
        if (/FIGHTING COMMANDER OCTA/i.test(product)) {
          var colorH = (product.match(/\bBlack\b|\bWhite\b/i) || ["Black"])[0];
          return "Hori Fighting Commander OCTA " + colorH + " " + variation.trim();
        }
        if (/FIGHTING STICK MINI/i.test(product)) return "Hori Fighting Stick Mini " + variation.trim();
        var codeH = (variation.match(/[A-Z0-9]+-[A-Z0-9]+/i) || [])[0] || normalizeSku(sku);
        return "Hori " + codeH;
      }
  } },

  { name: 'Sanrio Characters Hair Dryer', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/SANRIO CHARACTERS HAIR DRYER/i.test(product)) {
        return "Sanrio Characters Hair Dryer Series - " + variation.trim();
      }
  } },

  { name: 'Kyosho', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/KYOSHO/i.test(product)) {
        return "Kyosho " + variation.split(",")[0].trim() + " " + normalizeSku(sku);
      }
  } },

  { name: 'Snail Shell Studio', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/SNAIL SHELL STUDIO/i.test(product)) {
        var baseMatchSS = product.match(/^(.*?)(Action Figure|Addtional)/i);
        var basePrefixSS = baseMatchSS ? baseMatchSS[1].trim() : stripTrailingCjkSection(product);
        if (baseMatchSS && /Action Figure/i.test(baseMatchSS[2])) {
          return basePrefixSS + " Action Figure";
        }
        var partKeySS = variation.split(",")[0].trim();
        var partNameSS = /LEG/i.test(partKeySS) ? "Leg" : (/WING/i.test(partKeySS) ? "Wings" : (/WAIST/i.test(partKeySS) ? "Waist" : partKeySS));
        return (basePrefixSS + " Addtional " + partNameSS + " Parts").replace(/\s+/g, " ").trim();
      }
  } },

  { name: 'Yu-Gi-Oh', resolve: function (product, variation, sku, skuMap, ctx) {
    if (/YU-?GI-?OH|YUGIOH/i.test(product)) {

      // Keep the product family wording for every Original Artwork Duel Set
      // selection (Set of 3, Yugi Muto, Seto Kaiba, Joey Wheeler).
      if (/OCG/i.test(product) && /ORIGINAL ARTWORK/i.test(product) && /DUEL SET SERIES/i.test(product)) {
        var duelSetBase = product
          .replace(/\s*\[JAPAN\]\s*/i, ' Jap ')
          .replace(/\s*\(\s*Yugi\s*\/\s*Kaiba\s*\/\s*Joey\s*\)\s*/i, ' ')
          .replace(/\s+/g, ' ').trim();
        var duelSetPick = variation.split(',')[0].trim();
        return (duelSetBase + (duelSetPick ? ' ' + duelSetPick : '')).replace(/\s+/g, ' ').trim();
      }

      var langYG = '';

      if (/JAPAN|\bJAP\b|\bJP\b/i.test(product + ' ' + sku)) {
        langYG = 'Jap';
      } else if (/CHINESE|\bCHI\b/i.test(product + ' ' + sku)) {
        langYG = 'Chi';
      } else if (/ENGLISH|\bENG\b/i.test(product + ' ' + sku)) {
        langYG = 'Eng';
      }

      var typeYG = '';

      if (/BOOSTER BOX/i.test(product)) {
        typeYG = 'Booster Box';
      } else if (/BOOSTER PACK/i.test(product)) {
        typeYG = 'Booster Pack';
      } else if (/STARTER DECK/i.test(product)) {
        typeYG = 'Starter Deck';
      } else if (/STRUCTURE DECK/i.test(product)) {
        typeYG = 'Structure Deck';
      } else if (/DECK/i.test(product)) {
        typeYG = 'Deck';
      }

      // The variation contains the actual set code and product identity.
      var pickYG = variation.split(',')[0].trim();

      if (/^(Full Payment|Balance Payment(?:\s+P\.?O\.?)?|Deposit)$/i.test(pickYG)) {
        pickYG = '';
      }

      return [
        'Yu-Gi-Oh',
        langYG,
        typeYG,
        pickYG
      ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
    }
} },

  { name: 'One Piece Tcg', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/ONE PIECE TCG/i.test(product)) {
        if (/OFFICIAL CARD SLEEVE/i.test(product)) {
          return ("One Piece TCG: Official Card Sleeve " + variation.trim()).replace(/\s+/g, " ").trim();
        }
        var langOP = "";
        if (/\[JAPAN\]/i.test(product)) langOP = "Jap";
        else if (/\[ENGLISH\]/i.test(product)) langOP = "Eng";
        else if (/\[CHINESE\]/i.test(product)) langOP = "Chi";
        var typeOP = "";
        if (/BOOSTER BOX/i.test(product)) typeOP = "Booster Box";
        else if (/BOOSTER PACK/i.test(product)) typeOP = "Booster Pack";
        else if (/STARTER DECK/i.test(product)) typeOP = "Starter Deck";
        var codeOP = "";
        var mOP = variation.match(/(OP\d+\+EB\d+|PRB\d+|EB-?\d+|OP-?\d+|ST\d+)/i);
        if (mOP) codeOP = mOP[1].toUpperCase().replace(/-/g, "");
        return ["One Piece TCG", langOP, typeOP, codeOP].filter(Boolean).join(" ");
      }
  } },

  { name: 'Riftbound', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/RIFTBOUND|HOLOLIVE TCG|FINAL FANTASY TCG/i.test(product)) {
        var srcR = sku + " " + variation + " " + product;
        var langR = "";
        if (/SIMPLIFIED CHINESE/i.test(srcR) || /TRADITIONAL CHINESE/i.test(srcR) || /\bCHINESE\b/i.test(srcR) ||
            /\bCHI\b/i.test(srcR) || /CSV\d+/i.test(srcR) || /CBB\d+/i.test(srcR) || /[\u4e00-\u9fff]/.test(srcR)) {
          langR = "Chi";
        } else if (/\[JAPAN/i.test(srcR) || /\bJAP\b/i.test(srcR)) {
          langR = "Jap";
        } else {
          langR = "Eng";
        }
        var vTypeR = variation.split(",")[0].trim();
        var typeR = "";
        if (/BOOSTER BOX/i.test(vTypeR)) typeR = "Booster Box";
        else if (/BOOSTER PACK/i.test(vTypeR)) typeR = "Booster Pack";
        else if (/VAULT BUNDLE/i.test(vTypeR)) typeR = "Vault Bundle";
        else if (/STARTER DECK/i.test(vTypeR)) typeR = "Starter Deck";
        else if (/CHAMPION DECK/i.test(vTypeR)) typeR = "Champion Deck";
        else if (/BOOSTER BOX/i.test(product)) typeR = "Booster Box";
        else if (/BOOSTER PACK/i.test(product)) typeR = "Booster Pack";
        else if (/STARTER DECK/i.test(product)) typeR = "Starter Deck";
        else if (/CHAMPION DECK/i.test(product)) typeR = "Champion Deck";
        var codeR = "";
        var mR = srcR.match(/(RB-[A-Z0-9\-]+|BT\d+|EX\d+|AD\d+|OP\d+|PRB\d+|ST\d+)/i);
        if (mR) codeR = mR[1].toUpperCase();
        var gameR = "";
        if (/DIGIMON/i.test(product)) gameR = "Digimon TCG";
        else if (/RIFTBOUND/i.test(product)) gameR = "League Of Legends RiftBound TCG";
        else if (/FINAL FANTASY/i.test(product)) gameR = "Final Fantasy TCG";
        else if (/HOLOLIVE/i.test(product)) gameR = "Hololive TCG";
        // The series/edition name (e.g. "Unleashed") when variation states one
        // that isn't itself just a restatement of the type or payment status.
        var seriesR = "";
        if (vTypeR && !/^(Full Payment|Balance Payment|Deposit|Booster Box|Booster Pack|Starter Deck|Champion Deck|Vault Bundle)$/i.test(vTypeR)) {
          seriesR = vTypeR;
        }
        return [gameR, langR, typeR, seriesR, codeR].filter(Boolean).join(" ");
      }
  } },

  { name: 'Top 100 Best Games', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/TOP 100 BEST GAMES/i.test(product)) {
        var gnT100 = variation.split(",")[0].trim();
        return ["Switch", gnT100].filter(Boolean).join(" ");
      }
  } },

  { name: 'Promotion Games', resolve: function (product, variation, sku, skuMap, ctx) {
      if (/PROMOTION GAMES/i.test(product)) {
        var gameNamePG = variation.split(",")[0].trim();
        var platformPG = "";
        if (/^PS5/i.test(gameNamePG)) { platformPG = "PS5"; gameNamePG = gameNamePG.replace(/^PS5\s*/i, ""); }
        else if (/^PS4/i.test(gameNamePG)) { platformPG = "PS4"; gameNamePG = gameNamePG.replace(/^PS4\s*/i, ""); }
        else if (/^(SWITCH 2|SW2)/i.test(gameNamePG)) { platformPG = "Switch 2"; gameNamePG = gameNamePG.replace(/^(SWITCH 2|SW2)\s*/i, ""); }
        else if (/^(SWITCH|SW)/i.test(gameNamePG)) { platformPG = "Switch"; gameNamePG = gameNamePG.replace(/^(SWITCH|SW)\s*/i, ""); }
        if (!platformPG) {
          if (/^PS5/i.test(product)) platformPG = "PS5";
          else if (/^PS4/i.test(product)) platformPG = "PS4";
          else if (/NSW2|SWITCH 2/i.test(product)) platformPG = "Switch 2";
          else if (/NSW|SWITCH/i.test(product)) platformPG = "Switch";
        }
        return [platformPG, gameNamePG].filter(Boolean).join(" ");
      }
  } },

  { name: 'Ps4', resolve: function (product, variation, sku, skuMap, ctx) {
      // ---- console/switch/PS game name-building rule ----
      if (/PS4|PS5|NSW|NSW2|NINTENDO SWITCH|^SWITCH\b|^XBOX\b/i.test(product)) {
        var platform = "", language = "", region = "", addon = "";

        if (/^PS5\/NSW\/NSW2/i.test(product)) {
          if (/Switch 2/i.test(variation)) platform = "Switch 2";
          else if (/Switch 1/i.test(variation)) platform = "Switch";
          else if (/PS5/i.test(variation)) platform = "PS5";
        } else if (/^NSW\/NSW2/i.test(product)) {
          platform = /Switch 2/i.test(variation) ? "Switch 2" : "Switch";
        } else if (/^PS4\/PS5/i.test(product)) {
          platform = /PS5/i.test(variation) ? "PS5" : "PS4";
        } else {
          if (/^PLAYSTATION\s*PS5/i.test(product)) platform = "PS5";
          else if (/^PLAYSTATION\s*PS4/i.test(product)) platform = "PS4";
          else if (/^PS5/i.test(product)) platform = "PS5";
          else if (/^PS4/i.test(product)) platform = "PS4";
          else if (/^NSW2/i.test(product)) platform = "Switch 2";
          else if (/^NSW/i.test(product)) platform = "Switch";
          else if (/^NINTENDO SWITCH 2\b/i.test(product)) platform = "Switch 2";
          else if (/^NINTENDO SWITCH\b/i.test(product)) platform = "Switch";
          else if (/^SWITCH 2\b/i.test(product)) platform = "Switch 2";
          else if (/^SWITCH\b/i.test(product)) platform = "Switch";
          else if (/^XBOX SERIES\b/i.test(product)) platform = "Xbox Series";
          else if (/^XBOX\b/i.test(product)) platform = "Xbox";
        }

        // A dual-platform listing ("NSW /Switch 2 Product...") with the actual
        // choice living in variation as a bare "SW2" token — not caught by the
        // prefix-based detection above since the product text itself doesn't
        // start with a Switch-2-specific prefix.
        if (/\bSW2\b/i.test(variation) && !/\bNSW2\b/i.test(variation)) platform = "Switch 2";

        var langSource = product + " " + variation;
        if (/中文.*(?:english|eng)\b/i.test(variation) || /(?:english|eng)\b.*中文/i.test(variation)) language = "Chi/Eng";
        else if (/中文版|中文/i.test(variation)) language = "Chi";
        else if (/CHI\/ENG|ENG\/CHI/i.test(variation)) language = "Chi/Eng";
        else if (/\[ENG\]/i.test(variation) || /\bENGLISH\b/i.test(variation) || /\bENG\b/i.test(variation)) language = "Eng";
        else if (/\bCHI\b/i.test(variation)) language = "Chi";
        else if (/CHI\/ENG|ENG\/CHI/i.test(product)) language = "Chi/Eng";
        else if (/ENGLISH|ENG/i.test(langSource)) language = "Eng";

        var regionMatch = variation.match(/\[(R1|R2|R3|R4|US|EU|AU|AS|SEA)\]/i)
          || variation.match(/(?:^|\s|-)(R1|R2|R3|R4)(?:\s|-|,|$)/i)
          || variation.match(/(?:^|\s|-)(US|EU|AU|AS|SEA)(?:\s|-|,|$)/i);
        if (regionMatch) region = regionMatch[1].toUpperCase();

        variation.split(",").map(function (v) { return v.trim(); }).forEach(function (v) {
          if (/^ADD[- ]?ON/i.test(v)) { addon = v; return; }
          if (!addon) {
            for (var ci = 0; ci < CJK_EDITION_MARKERS.length; ci++) {
              if (CJK_EDITION_MARKERS[ci][0].test(v)) { addon = CJK_EDITION_MARKERS[ci][1]; break; }
            }
          }
          // English-language game edition names are often supplied only in
          // Shopee's variation field (e.g. "Collector - English"). Preserve
          // the edition instead of returning only the base game title.
          if (!addon) {
            var editionMarkers = [
              [/\bCOLLECTOR(?:'S)?(?:\s+EDITION)?\b/i, "Collector's Edition"],
              [/\bDELUXE(?:\s+EDITION)?\b/i, "Deluxe Edition"],
              [/\bLIMITED(?:\s+EDITION)?\b/i, "Limited Edition"],
              [/\bSPECIAL(?:\s+EDITION)?\b/i, "Special Edition"],
              [/\bSTANDARD(?:\s+EDITION)?\b/i, "Standard Edition"],
              [/\bULTIMATE(?:\s+EDITION)?\b/i, "Ultimate Edition"],
              [/\bGOLD(?:\s+EDITION)?\b/i, "Gold Edition"],
              [/\bCOMPLETE(?:\s+EDITION)?\b/i, "Complete Edition"],
              [/\bDAY\s*ONE(?:\s+EDITION)?\b/i, "Day One Edition"],
              [/\bLAUNCH(?:\s+EDITION)?\b/i, "Launch Edition"],
              [/\bANNIVERSARY(?:\s+EDITION)?\b/i, "Anniversary Edition"]
            ];
            for (var ei = 0; ei < editionMarkers.length; ei++) {
              if (editionMarkers[ei][0].test(v)) { addon = editionMarkers[ei][1]; break; }
            }
          }
        });

        var gameName = product
          .replace(/^PS5\/NSW\/NSW2\s*/i, "")
          .replace(/^PS4\/PS5\s*/i, "")
          .replace(/\(.*/, "")
          .replace(/^NSW\/NSW2\s*/i, "")
          .replace(/^NSW2\s*/i, "")
          .replace(/^NSW\s*/i, "")
          .replace(/^NINTENDO SWITCH\s*\/\s*SWITCH 2\s*/i, "")
          .replace(/^NINTENDO SWITCH\/SWITCH 2\s*/i, "")
          .replace(/^NINTENDO SWITCH 2\s*/i, "")
          .replace(/^NINTENDO SWITCH\s*/i, "")
          .replace(/^SWITCH 2\s*/i, "")
          .replace(/^SWITCH\s*/i, "")
          .replace(/^XBOX SERIES\s*/i, "")
          .replace(/^XBOX\s*/i, "")
          .replace(/^PLAYSTATION\s*/i, "")
          .replace(/^PS4\s*/i, "")
          .replace(/^PS5\s*/i, "")
          .replace(/\bCHI\/ENG\s*VERSION\b/gi, "")
          .replace(/\bENG\/CHI\s*VERSION\b/gi, "")
          .replace(/\bENGLISH\s*VERSION\b/gi, "")
          .replace(/\bENG\s*VERSION\b/gi, "")
          .replace(/\bCHI\s*VERSION\b/gi, "")
          .replace(/\bCHI\/ENG\b/gi, "")
          .replace(/\bENG\/CHI\b/gi, "")
          .replace(/\b中文版\b/gi, "")
          .replace(/\b中英文版\b/gi, "")
          .replace(/[一-龥ぁ-んァ-ヶ].*$/g, "")
          .replace(/：/g, ":")
          .replace(/\s+/g, " ")
          .trim();

        gameName = gameName
          .replace(/\bJUST DANCE\s+(\d{2})\s+20\1\b/i, "Just Dance 20$1")
          .replace(/\bCHINESE\s*\/\s*CHI\/ENG\b/gi, "")
          .replace(/\bCHINESE\b/gi, "")
          .replace(/^2\s+/i, "")
          .replace(/^SWITCH 2\s*/i, "")
          .replace(/^NINTENDO SWITCH\s*/i, "")
          .replace(/^SWITCH 2\s*/i, "")
          .replace(/^2\s+/i, "")
          .replace(/\b0 0\b/g, "0")
          .replace(/\b2 2\b/g, "2")
          .replace(/\b3 3\b/g, "3")
          .replace(/\bII II\b/g, "II")
          .replace(/\bIII III\b/g, "III")
          .replace(/STANDARD\s*\|\s*ULTIMATE EDITION/gi, "Standard Edition")
          .replace(/STANDARD\s*\/\s*DELUXE EDITION/gi, "Standard Edition")
          .replace(/\s*\/\s*Booster Course Pack\b/gi, "")
          .replace(/([A-Za-z]+)\s*\/\s*([A-Za-z]+)\s+Edition\b/i, function (m, opt1, opt2) {
            var chosen = (variation && new RegExp(opt2, "i").test(variation)) ? opt2 : opt1;
            return chosen + " Edition";
          })
          .replace(/\s+\/\s*$/g, "")
          .replace(/\s+\/\s+(?=Chi\/Eng|Eng|Chi)/gi, " ")
          .replace(/\s+/g, " ")
          .trim();

        var fcMatch = gameName.match(/\bFC\s*(\d{2})\b/i);
        if (fcMatch) {
          gameName = "FC" + fcMatch[1];
        } else {
          fcMatch = gameName.match(/FOOTBALL CLUB\s*(\d{2})/i);
          if (fcMatch) gameName = "FC" + fcMatch[1];
        }

        if (!region) {
          var bareRegionMatch = gameName.match(/\b(R1|R2|R3)\b\s*$/i);
          if (bareRegionMatch) {
            region = bareRegionMatch[1].toUpperCase();
            gameName = gameName.slice(0, bareRegionMatch.index).trim();
          }
        }

        // A title like "Final Fantasy VII / 7 Remake Intergrade / Remake + Rebirth
        // Twin Pack" is a list of edition alternatives, not one compound title —
        // distinguishable from a legit compound title like "X/X-2" by the spaces
        // around the slash. Reduce it to the actual chosen edition.
        if (/\s\/\s/.test(gameName)) {
          var vGame = (variation || "").split(",")[0].replace(/\[.*?\]/g, "").trim();
          var isPlatformMarkerOnly = /^(SW2|NSW2?|SWITCH\s*2?)$/i.test(vGame);
          if (vGame && !isPlatformMarkerOnly && !/^(Full Payment|Balance Payment|Deposit|P\.?O\.?)\b/i.test(vGame)) {
            var reducedGame = reduceOptionList(gameName, vGame);
            gameName = reducedGame !== gameName ? reducedGame : (gameName + " " + vGame).replace(/\s+/g, " ").trim();
          }
          // Still has an unresolved multi-option bracket (nothing matched it) —
          // drop it rather than leave a mangled partial bracket in the output.
          gameName = gameName.replace(/\s*\[[^\]]*\/[^\]]*\]\s*/g, " ").replace(/\s+/g, " ").trim();
        }

        return [platform, gameName, language, region, addon].filter(Boolean).join(" ");
      }
  } }
];


// =========================================================================
// NAME RESOLUTION -- step 3: generic fallback
// =========================================================================
// Reached only when no rule above matched. Trims marketing fluff, reduces a
// "/"-separated option list down to the variant actually ordered, and appends
// distinguishing info from `variation` when the name lacks it.
function resolveGenericFallback(product, variation, sku, skuMap) {
  // Game & Watch has two distinct models; use the selected buyer option.
  if (/Game\s*&\s*Watch/i.test(product || '')) {
    if (/\bMario\b/i.test(variation || '')) return 'Switch Game & Watch Super Mario Bros';
    if (/\bZelda\b/i.test(variation || '')) return 'Switch Game & Watch The Legend of Zelda';
  }

  var finalFallback = stripTrailingCjkSection(product);
  var vFallback = (variation || "").split(",")[0].trim();
  var isBoilerplateVariation = /^(Full Payment|Balance Payment|Deposit|P\.?O\.?)\b/i.test(vFallback) || /\bRandom\b/i.test(vFallback);

  // Try reducing a "/"-separated option list first — it needs the full,
  // un-cut text to find and pick the matching option. Only fall back to
  // blunt fluff-trimming (below) if there was nothing to reduce.
  if (vFallback && !isBoilerplateVariation) {
    var reducedFirst = reduceOptionList(finalFallback, vFallback);
    if (reducedFirst !== finalFallback) return reducedFirst;
  }

  // Trim marketing fluff: a "with ..." clause, a parenthetical spec block,
  // or a " - spec, spec, spec" list (comma-separated, so a genuine
  // "Model - Color" pattern without commas is left alone).
  var fluffCut = finalFallback.split(/\s+with\s+|\s*\(|\s*\||\s*\[/i)[0];
  var dashParts = fluffCut.split(/\s+-\s+/);
  if (dashParts.length > 1 && dashParts[1].indexOf(",") !== -1) {
    fluffCut = dashParts[0];
  }
  if (fluffCut.trim().length >= 8) finalFallback = fluffCut.trim();

  if (vFallback && !isBoilerplateVariation) {
    var reduced = reduceOptionList(finalFallback, vFallback);
    if (reduced !== finalFallback) return reduced;
    if (finalFallback.toLowerCase().indexOf(vFallback.toLowerCase()) === -1) {
      // The resolved text is missing real distinguishing info that `variation`
      // supplies (a color, model, or character name) — append it. This covers
      // both "lists every option, picks none" and "generic title, no variant
      // at all" cases, without needing a dedicated rule per brand.
      return (finalFallback + " " + vFallback).replace(/\s+/g, " ").trim();
    }
  }
  // No usable variation to pick an option with (boilerplate or blank) — a
  // dangling "/ alternate" list is worse than just the base name, so drop
  // it rather than show every unresolved option.
  finalFallback = finalFallback.replace(/\s+\/\s*.*$/, "").trim();
  // Clean up a trailing dangling separator left over from a stripped CJK tail.
  finalFallback = finalFallback.replace(/[\s\-–—]+$/, "").trim();
  return finalFallback;
}

// =========================================================================
// NAME RESOLUTION -- entry point
// =========================================================================
// Turns one raw line item into its clean display name. Returns "" for items
// that should not appear in the output at all (packaging add-ons).


// =========================================================================
// TCG SKU REFERENCE NAME DETECTION
// =========================================================================
// Use SKU Reference No. as the output name ONLY when it looks like a
// complete TCG product title, rather than a set code or variation name.
//
// Examples accepted:
//   Pokemon Tcg Jp Booster Box M6 - Storm Emeralda
//   Pokemon Tcg Ascended Heroes Focused Fighters Premium Collection ME2.5
//   Palworld Tcg Eng Booster Box BP02 - Legends Awaken
//   Riftbound LOL Tcg Series 3 Unleashed Champion Sleeves - LeBlanc
//
// Examples rejected:
//   M6
//   CSV9.5C
//   ME2.5
//   Sleeves,LeBlanc
//   Full Payment
//   Charizard
//   Anime Athletes

function getFullTcgSkuReference(sku, product, variation) {
  var ref = String(sku || '').trim();
  var productName = String(product || '');
  var variationName = String(variation || '').split(',')[0].trim();

  if (!ref) return '';

  // Reject payment, deposit and option-only references.
  if (/FULL PAYMENT|BALANCE PAYMENT|DEPOSIT|PARTIAL PAYMENT|RANDOM COVER/i.test(ref)) {
    return '';
  }

  // A full product reference must identify a TCG/game franchise.
  var hasGameName =
    /\bTCG\b|\bPOKEMON\b|\bONE PIECE\b|\bRIFTBOUND\b|\bPALWORLD\b|\bGUNDAM\b|\bDIGIMON\b|\bHOLOLIVE\b|\bFINAL FANTASY\b|\bYUGIOH\b|\bYU-GI-OH\b|\bTOPPS\b/i.test(ref);

  if (!hasGameName) return '';

  // It must also identify a product type, not merely a character or option.
  // MODEL KIT is included because Yu-Gi-Oh collaboration merchandise can be
  // a model kit rather than a card-game product.
  var hasProductType =
    /\bBOOSTER BOX\b|\bBOOSTER PACK\b|\bBOOSTER CASE\b|\bSTARTER DECK\b|\bTRIAL DECK\b|\bCHAMPION DECK\b|\bPREMIUM COLLECTION\b|\bPREMIUM SET\b|\bCOLLECTOR\b|\bGIFT BOX\b|\bCARD SLEEVE\b|\bSLEEVES\b|\bMINI TIN\b|\bMINI BOOSTER TIN\b|\bSUPER TIN\b|\bDECK BOX\b|\bCARD SET\b|\bTACTICS DECK\b|\bSURPRISE BOX\b|\bPLAYMAT\b|\bMODEL KIT\b/i.test(ref);

  if (!hasProductType) return '';

  // A Yu-Gi-Oh SKU that merely repeats the generic listing title is not a
  // descriptive SKU. Let the dedicated Yu-Gi-Oh rule append its variation.
  function comparableText(s) {
    return String(s || '').toUpperCase().replace(/[^A-Z0-9]+/g, '');
  }
  if (/YU-?GI-?OH|YUGIOH/i.test(ref) &&
      comparableText(ref) === comparableText(productName)) {
    return '';
  }

  // Generic SKU family label only: let the variation supply the actual set.
  if (/^(?:YU-?GI-?OH|YUGIOH)\s+(?:JAPAN|JAP)\s+BOOSTER BOX$/i.test(ref)) {
    return '';
  }

  // If the variation is a real selection, the SKU must agree with it.
  // This prevents a complete-looking parent/listing SKU from overriding a
  // different option selected by the buyer. Payment/deposit/random labels
  // are not product selections and therefore do not participate.
  var variationIsAdministrative = /^(FULL PAYMENT|BALANCE PAYMENT|DEPOSIT|PARTIAL PAYMENT|P\.?O\.?|RANDOM COVER)(?:\b|$)/i.test(variationName);
  if (variationName && !variationIsAdministrative && !/[\u4e00-\u9fff]/.test(variationName)) {
    var variationTokens = variationName.toUpperCase().match(/[A-Z0-9]+/g) || [];
    var ignoredVariationTokens = /^(THE|AND|FOR|OF|JAPAN|JAP|ENGLISH|ENG|CHINESE|CHI|BOOSTER|PACK|BOX|TIN|MINI|SUPER|TCG)$/;
    variationTokens = variationTokens.map(function (t) { return t.replace(/[^A-Z0-9]/g, ''); })
      .filter(function (t) { return t.length >= 3 && !ignoredVariationTokens.test(t); });
    if (variationTokens.length) {
      var normalizedRefForMatch = ref.toUpperCase().replace(/[^A-Z0-9]/g, '');
      var variationMatchesSku = variationTokens.some(function (t) { return normalizedRefForMatch.indexOf(t) !== -1; });
      if (!variationMatchesSku) return '';
    }
  }

  // Reject short references that are just codes or option labels.
  var words = ref.split(/\s+/).filter(Boolean);
  if (words.length < 4) return '';

  // Remove trailing CJK text while retaining the readable product title.
  ref = stripTrailingCjkSection(ref);
  // Shopee sometimes appends an option inventory count, e.g. '(2)' or '(3)'.
  // It is not part of the product's model/name.
  ref = ref.replace(/\s*\(\d+\s*\/\s*[Ss]et\)\s*$/i, '');
  ref = ref.replace(/\s*\(\d+\)\s*$/, '');
  ref = ref.replace(/\s+/g, ' ').trim();

  return ref;
}


function cleanProduct(product, variation, sku, skuMap) {
  product = normalizeProductText(product);
  if (product === null) return ""; // packaging add-on: intentionally dropped

  // Game & Watch listing has two models; keep the actual selected game title.
  if (/Game\s*&\s*Watch/i.test(product)) {
    if (/\bMario\b/i.test(variation || '')) return 'Switch Game & Watch Super Mario Bros';
    if (/\bZelda\b/i.test(variation || '')) return 'Switch Game & Watch The Legend of Zelda';
  }

  // -------------------------------------------------------
  // TCG: prefer a complete SKU Reference No. when available
  // -------------------------------------------------------
  var isTcg =
    /\bTCG\b|POKEMON|ONE PIECE|RIFTBOUND|PALWORLD|GUNDAM CARD GAME|DIGIMON TCG|HOLOLIVE TCG|FINAL FANTASY TCG|YU-GI-OH|YUGIOH/i.test(product);

  if (isTcg) {
    var fullTcgSkuName = getFullTcgSkuReference(sku, product, variation);

    if (fullTcgSkuName) {
      return fullTcgSkuName;
    }
  }

  // -------------------------------------------------------
  // Existing name-resolution rules remain unchanged
  // -------------------------------------------------------
  var ctx = { normalizedSku: normalizeSku(sku) };

  for (var i = 0; i < PRODUCT_RULES.length; i++) {
    var resolved = PRODUCT_RULES[i].resolve(
      product,
      variation,
      sku,
      skuMap,
      ctx
    );

    if (resolved !== undefined && resolved !== null) return resolved;
  }

  return resolveGenericFallback(product, variation, sku, skuMap);
}

// =========================================================================
// REVIEW FLAGS
// =========================================================================
// ---------------------------------------------------------------------------
// Flags a row for visual review when its variation still contains CJK text
// that isn't one of the edition markers we already know how to translate.
// This is the safety net for the "特典版 silently dropped" class of bug —
// instead of relying on catching every future CJK term one by one, anything
// unrecognized gets highlighted so it's visible in the sheet, not lost.
// Add new terms here (both list and detection regex stay in sync automatically).
var CJK_EDITION_MARKERS = [
  [/特典版/, "Special Edition"],
  [/初回版|初回限定版/, "First Press Edition"],
  [/限定版/, "Limited Edition"],
  [/普通版|通常版/, "Standard Edition"],
  [/豪華版|豪华版/, "Deluxe Edition"],
  [/典藏版/, "Collector's Edition"]
];
var KNOWN_CJK_MARKERS = new RegExp(CJK_EDITION_MARKERS.map(function (m) { return m[0].source; }).join("|"), "g");
function hasUnexplainedCjk(variation) {
  var text = variation || '';
  if (!/[\u4e00-\u9fff]/.test(text)) return false;
  var remaining = text.replace(KNOWN_CJK_MARKERS, '');
  return /[\u4e00-\u9fff]/.test(remaining);
}

function wasResolved(product, sku, skuMap) {
  var normalizedSku = normalizeSku(sku);
  if (normalizedSku && skuMap[normalizedSku]) return true;
  if (/EXTRA PROTECTION PACKAGING|BUBBLE WRAP|EXTRA THICKNESS BOX/i.test(product)) return true; // intentionally dropped, not a gap

  var triggers = [
    /FIDGET TOY|PUSH POP SENSORY/i, /BANDAI/i, /ANGRY SHINN/i, /YOLOPARK/i, /ZD TOYS/i, /TAKARATOMY/i,
    /^BLOKEES/i, /MINI\s*GT|MINIGT/i, /POPRACE/i, /ROBOTIME/i, /^LEGO/i, /MOTOR NUCLEAR/i, /GSC/i,
    /MICROSOFT XBOX/i, /^GEEKSHARE/i, /^BJBT/i, /^SKULL & CO/i, /^HORI\b/i, /DUALSENSE/i, /^CHINA MOE TOWN/i, /DOBE/i, /^IINE/i, /CLEAREDGE/i, /8BITDO/i, /ARZOPA/i, /EARFUN/i, /BROOK/i, /^CUTEBEE/i,
    /SUNWAY TINY/i, /SWITCH\s*2\s*CONSOLE/i, /MAXSOFT CONSOLE/i, /^MOSHOW/i, /^TRIMUI/i, /^R36H\b/i, /^SPIN MASTER/i, /^DANBOARD/i, /OFFICIAL CARD SLEEVE/i,
    /^KEYCHRON/i, /^MR\s?HOBBY/i, /^TOPLOADER TOP LOADER CARD/i, /^XBOX\b/i,
    /^VT (ACCESSORIES|GUNDAM GUNPLA)|GUNPLA ACCESSORIES/i, /^MODEL KITS? ACCESSORIES/i, /^MAJORETTE/i, /^JISULIFE/i, /^FUNISM/i, /^MOZA\b/i, /^OMELET GAMING/i, /^LUCKY FOX/i, /^WEKKI BLOCK/i,
    /ROYAL KLUDGE|^RK\b/i, /R36 ULTRA|R36S/i, /^MIYOO/i, /POKEMON TCG/i, /GUNDAM CARD GAME/i, /HOLOLIVE/i,
    /DIGIMON/i, /QUBECARZ/i, /ARKAIZO/i, /VOICSKY/i, /KAYOU/i, /TOPPS TCG U[A-Z]{1,3} MATCH ATTAX/i,
    /^HORI\b/i, /SANRIO/i, /KYOSHO/i, /SNAIL SHELL STUDIO/i, /PALWORLD TCG/i, /ONE PIECE SINGLE/i,
    /^CORSAIR\b/i, /^JSAUX\b/i, /TAMAGOTCHI/i,
    /\bTCG\b.*(BOOSTER BOX|BOOSTER PACK|VALUE BOX|VALUE BLASTER BOX|SINGLE BOOSTER PACK)/i,
    /ONE PIECE TCG/i, /RIFTBOUND|HOLOLIVE TCG|FINAL FANTASY TCG/i, /TOP 100 BEST GAMES/i, /PROMOTION GAMES/i,
    /PS4|PS5|NSW|NSW2|NINTENDO SWITCH|^SWITCH\b/i, /VITURE.*USB-?C/i
  ];
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].test(product)) return true;
  }
  // last-resort CJK strip counts as "resolved" only if it actually changed the text
  return stripTrailingCjkSection(product) !== product;
}


// =========================================================================
// CATEGORIZATION -- explicit rules
// =========================================================================
var CATEGORY_RULES = [
  ['PS', /^(PS4\/PS5|PS4|PS5)\b|^DOBE TP|^DOBE PS\b|^Dobe Ps4\b|^Dobe Ps5\b|^Dobe Psvr\b|^Dobe 3M Braided Usb Type A-C Data Cable|^Used Ps5 Acc\b/i],
  ['SWITCH', /^Switch\b|^Switch2\b|^Switch 2\b|^Nintendo Switch\b(?!.*Console)|^Hori\b(?!.*Fighting)|^iiNE\b|^DOBE (SW|TNS)|^Dobe Switch\b|^PMW\b|^GeekShare\b|^G-NS|^Jsaux\b|^FADED\b|^BJBT\b|^Skull & Co\b/i],
  ['Xbox Gaming Acc', /^8BitDo\b(?!.*(Keyboard|Numpad|Mouse|Wrist Rest|Adapter))|^Hori\b.*Fighting|^PowerA\b|^Microsoft Xbox\b|^DOBE\b|^Xbox\b|^PC Steam\b/i],
  ['Retail', /^Arzopa\b|^Asus\b|^Aula\b|^EarFun\b|^Logitech\b|^G923\b|^Razer\b|^Royal Kludge\b|^Steelseries\b|^Corsair\b|^Voicsky\b|^Keychron\b|^Tecware\b|^8BitDo\b.*(Keyboard|Numpad|Mouse|Wrist Rest|Adapter)|^R36\b|^R36S\b|^R36H\b|^TRIMUI\b|^Miyoo\b|^VITURE\b|^Sandisk\b|^CM\b|^MSI\b|^Mchose\b|^Vention\b|^Nanoleaf\b|^Archgon\b|^Jabra\b|^Outin\b|^ADATA\b|^Techgo\b|^MelGeek\b|^Sony\b|^GoPro\b|^Fujifilm\b|^Casio\b|^APC\b|^AYANEO\b|^Tomtoc\b|^PlayStation VR\d*\b|^Divoom\b|^MOZA\b|^ANBERNIC\b|^NuPhy\b|^Syitren\b|^Nintendo Switch.*Console\b/i],
  ['Blind Box', /^(?!.*\bTCG\b).*\(\d+\s*\/\s*[Ss]et\)|Blind Box|^Kayou\b|^52Toys\b|^ToyZeroPlus\b|^Re-Ment\b|^Sembo\b|^Miffy\b|^Miniso\b|^Smiski\b|^San-Ei\b|^Duoduobao\b|^Hugkis\b|^Butterbear\b|^Gracshaw\b|^Jaki\b|^Leobog\b|^Pop Mart\b/i],
  ['LEGO', /^LEGO\b/i],
  ['ToysVT031XH018', /Tamagotchi\b|^BB-\d+\b|^Beyblade\b|^(?:BX|CX|UX)-\d+(?:\s+\*\d+)?$|^Bandai\b|^Blokees\b|^Cutebee\b|^Howling Star\b|^Jada Toys\b|^Mini\s?GT\b|^PopRace\b|^Robotime\b|^TakaraTomy\b|^ZOE\b|^Sanrio\b|^Kyosho\b|^QubeCarz\b|^ARKaizo\b|^Almost Real\b|^Moreart\b|^SNAA\b|^Snail Shell\b|^Digimon\b(?!.*TCG)|^China Moe Town\b|^Sunway Tiny\b|^MotorHelix\b|^Motor Nuclear\b|^UPF\b|^Jisulife\b|^Xcartay\b|^Xcartoys\b|^ShouCangJia\b|^Zhe Gao\b|^MW Model\b|^Tarmac\b|^Inno64\b|^American Diorama\b|^Bburago\b|^GCD\b|^Tomica\b|^ModelModel\b|^Moshow\b|^Ensky\b|^Threezero\b|^Loz\b|^Hemoxian\b|^Joytoy\b|^GSC\b(?!.*Moderoid)|^Sentinel-Toys\b|^ACG Art\b|^Zen Of Collectible\b|^MegaGear\b|^Crayon Shin-Chan\b|^Amiibo\b|^ZD Toys\b|^Angry Shinn\b|^Wekki Block\b|^Majorette\b|^Banpresto\b|^URDU\b|^GS Toys\b|^Kraken House\b|^Megahouse\b|Pokeball Collection/i],
  ['Gundam ACC  Mr.Hobby', /^(?=.*(?:Yu-?Gi-?Oh|Yugioh))(?=.*\bModel Kit\b).*$|^Dspiae\b|^God Hand\b|^VT\d+|^VT\b|^Model Kit|^Comet Tech\b|^Fun Club\b|^Infinite Dimension\b|^Mecha Ark\b|^MJZ\b|^Gunpla Accessories\b|^Mr Hobby\b|^GSC.*Moderoid|^Farmer Creation\b|^Super Soldier Project\b|^WS-\d+|^GBC Workshop\b|^Pigsy\b|^Wolf Remains\b|^QianHui\b|^HH Model\b|^Solomon\b|^Yiya-G Studio\b|^E-Model\b/i],
  ['CLEARANCE', /^Fidget Toy\b/i],
  ['TCG', /Cardfight|ClearEdge|Gundam Card Game|Digimon TCG.*Booster Box|Digimon TCG.*ST\d+|Digimon Tcg Card Case Set|RiftBound TCG.*Booster Box|One Piece TCG.*(Booster Box|Premium Card Collection)|PalWorld TCG.*(Booster Box|Booster Case)|Pokemon TCG.*(Booster Box|Booster Bundle|Elite Trainer Box|Dice Maker Set|Premium Deck Set)|Pokemon Tcg.*(Premium Collection|Surprise Box Collection|Ultra Premium Collection|Partner Celebration Set|Pin Badge Set)|Gift Box|Panini TCG|Playfully Herock TCG|Wustyle TCG|Topps TCG.*(Booster Box|Value Box|Value Blaster Box)|Trial Deck|Try Deck|Starter Deck|Starter Pack|Champion Deck|Vault Bundle|Yu-?Gi-?Oh|Yugioh|Union Arena TCG|Official Card Sleeve|Official Card Case|^Toploader\b|\bTCG\b.*Booster Box/i],
  ['TCG BP', /RiftBound TCG.*Booster Pack|One Piece (Single|TCG.*Booster Pack)|Pokemon TCG.*(Booster Pack|Promo Card|Empty Mini Tins|Tactics Deck Series|Blister Pack|First Partner)|Pokemon TCG \S+ Ex$|Hololive TCG.*Booster Pack|Digimon TCG.*Booster Pack|Topps TCG/i],
  ['Payments, Add-ons & Shipping', /^Add-On\b|Extra Protection|Bubble Wrap|Extra Thickness Box/i],
];
// Absolute last resort only — a brand-new item that matches none of the above
// (genuinely never seen before). Kept visible and separate rather than
// silently dropped, but this should stay empty in normal operation; if
// something lands here, it's a signal to add a proper category rule above.
var DEFAULT_CATEGORY = '⚠ New / Uncategorized — add a rule above';
var CATEGORY_ORDER = CATEGORY_RULES.map(function (r) { return r[0]; }).concat([DEFAULT_CATEGORY]);

function categorize(name) {
  if (isFujifilmFilmOnlyProduct(name)) return 'Xbox Gaming Acc';
  // 8BitDo products always use the Xbox Gaming Acc category, including keyboards,
  // mice, adapters and wrist rests that the old broad rule accidentally excluded.
  if (/\b8BitDo\b/i.test(String(name || '').trim())) return 'Xbox Gaming Acc';
  // Retail order families must not be captured by earlier platform rules.
  if (/^(?:Used Games\b|Used Ps[345] Acc\b|Used Console\b|Nintendo Switch(?:\s*2)? Console\b|Switch2? Console\b|PlayStation [45] Console\b|PS[45] Console\b)/i.test(name) ||
      /^(?:PS5|PS4)\b.*DUALSENSE.*CONTROLLER/i.test(name) ||
      /^(?:Aula|Keychron)\b/i.test(name) || /\bRazer\b/i.test(name)) return 'Retail';
  for (var i = 0; i < CATEGORY_RULES.length; i++) {
    if (CATEGORY_RULES[i][1].test(name)) return CATEGORY_RULES[i][0];
  }
  return DEFAULT_CATEGORY;
}

// =========================================================================
// CATEGORIZATION -- inference for unmatched items
// =========================================================================
// Important: these are GUESSES, not facts. Anything inferred here is written
// to the sheet with its basis shown and stays highlighted for review — it is
// never silently filed as if an explicit rule had matched it. A wrong guess
// that looks confident is worse than an honest "uncategorized" flag, because
// a misfiled item vanishes into a category you trust.

// Keyword signals, checked in order — first match wins. Ordered so the more
// specific/distinctive terms are tested before generic ones.
var INFERENCE_KEYWORDS = [
  ['TCG BP', /\bbooster pack\b|\bpromo card\b|\bsingle\b.*\bcard\b|\bmini tins?\b/i],
  ['TCG', /\bbooster box\b|\bbooster case\b|\belite trainer\b|\bstarter deck\b|\btrial deck\b|\bcard sleeve\b|\bdeck box\b|\btoploader\b|\bbinder\b|\bplaymat\b|\bcard case\b|\btcg\b/i],
  ['Gundam ACC  Mr.Hobby', /\bmodel kits?\b|\bgunpla\b|\bplastic model\b|\b(?:hg|rg|mg|pg|sd)\b\s*1\/\d+|\bnipper\b|\bcement\b|\bpanel liner\b|\btopcoat\b/i],
  ['Blind Box', /\bblind box\b|\bfull case\b|\(\d+\s*\/\s*set\)|\bmystery box\b|\bsurprise box\b/i],
  ['LEGO', /^lego\b/i],
  ['Retail', /\bkeyboard\b|\bmouse\b|\bmousepad\b|\bmouse pad\b|\bheadset\b|\bearbuds?\b|\bmonitor\b|\bconsole\b|\bssd\b|\brouter\b|\bwebcam\b|\bmicrophone\b|\bcharger\b|\bpower bank\b|\bcapture card\b/i],
  ['ToysVT031XH018', /\bfigures?\b|\bplush\b|\bdiecast\b|\b1\/64\b|\b1\/18\b|\b1\/12\b|\b1\/6\b|\bnendoroid\b|\bstatue\b|\bcollectible\b|\bdiorama\b/i],
  ['SWITCH', /\bswitch\b|\bjoy-?con\b|\bnintendo\b/i],
  ['PS', /\bps4\b|\bps5\b|\bplaystation\b|\bdualsense\b/i],
  ['Xbox Gaming Acc', /\bxbox\b|\bgamepad\b|\bcontroller\b/i]
];

// Reads the Learned_Rules sheet — corrections you've made previously, which
// take priority over keyword guessing since they came from you, not a guess.
function buildLearnedRules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEARNED_RULES_SHEET_NAME);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) { // row 0 is the header
    var phrase = String(data[i][0] || '').trim();
    var category = String(data[i][1] || '').trim();
    if (!phrase || !category) continue;
    if (CATEGORY_ORDER.indexOf(category) === -1) continue; // ignore unknown categories
    out.push([phrase, category]);
  }
  return out;
}

// Returns {category, basis} or null. `basis` explains WHY, so the sheet can
// show you what the guess was based on rather than asking you to trust it.
function inferCategory(name, learnedRules, orderMates) {
  var text = name || '';

  // 1. A correction you made previously — highest confidence, it's your call.
  for (var i = 0; i < learnedRules.length; i++) {
    if (text.toUpperCase().indexOf(learnedRules[i][0].toUpperCase()) !== -1) {
      return { category: learnedRules[i][1], basis: 'learned: "' + learnedRules[i][0] + '"' };
    }
  }

  // 2. Everything else in the same order agrees on one category.
  if (orderMates && orderMates.length) {
    var distinct = {};
    orderMates.forEach(function (c) { distinct[c] = true; });
    var keys = Object.keys(distinct);
    if (keys.length === 1 && keys[0] !== DEFAULT_CATEGORY) {
      return { category: keys[0], basis: 'same order as ' + keys[0] };
    }
  }

  // 3. Keyword signal in the product name.
  for (var k = 0; k < INFERENCE_KEYWORDS.length; k++) {
    if (INFERENCE_KEYWORDS[k][1].test(text)) {
      return { category: INFERENCE_KEYWORDS[k][0], basis: 'keyword match' };
    }
  }

  return null;
}

// =========================================================================
// PARSING
// =========================================================================
// ---------------------------------------------------------------------------
var ITEM_RE_SOURCE =
  '\\[\\d+\\]\\s*Product Name:(?<name>.*?);\\s*' +
  'Variation Name:(?<variation>.*?);\\s*' +
  'Price:\\s*RM\\s*(?<price>[\\d,]+\\.?\\d*);\\s*' +
  'Quantity:\\s*(?<qty>\\d+);\\s*' +
  'SKU Reference No\\.:\\s*(?<sku>.*?);' +
  '(?:\\s*Parent SKU Reference No\\.:\\s*(?<parent_sku>.*?);)?';

// Detects the "Packing List" export header (tracking_number, order_sn,
// product_info, ...) and returns a column-index map if found, else null.
// This lets parseOrders support both file formats without the user having
// to reshape anything — just paste whichever export they have.
function detectPackingListColumns(headerRow) {
  if (!headerRow) return null;
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var h = (headerRow[i] === null || headerRow[i] === undefined) ? '' : String(headerRow[i]).trim().toLowerCase();
    if (h === 'tracking_number') map.tracking = i;
    else if (h === 'order_sn') map.orderId = i;
    else if (h === 'product_info') map.productInfo = i;
  }
  if (map.tracking !== undefined && map.orderId !== undefined && map.productInfo !== undefined) return map;
  return null;
}


function parseOrders(rows) {
  var items = [];

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];

    // New RAW input structure:
    // Column A = Tracking Number
    // Column B = Order ID
    // Column C = Item Data

    var tracking = row[0] == null ? '' : String(row[0]).trim();
    var orderId = row[1] == null ? '' : String(row[1]).trim();
    var itemText = row[2] == null ? '' : String(row[2]);

    // Skip header row
    if (
      tracking.toLowerCase() === 'tracking number' ||
      orderId.toLowerCase() === 'order id'
    ) {
      continue;
    }

    // Skip empty or invalid rows
    if (!orderId || !itemText.trim()) continue;

    // Keep existing 5-character order code for display
    var orderCode = orderId.slice(-5).toUpperCase();

    var re = new RegExp(ITEM_RE_SOURCE, 'gs');
    var m;

    while ((m = re.exec(itemText)) !== null) {
      var g = m.groups;

      items.push({
        orderId: orderId,
        orderCode: orderCode,
        name: g.name.trim(),
        variation: g.variation.trim(),
        sku: g.sku.trim(),
        qty: parseInt(g.qty, 10),
        tracking: tracking
      });
    }
  }

  return items;
}


// =========================================================================
// PIPELINE
// =========================================================================
function buildConsolidated(rows, skuMap, learnedRules, inventoryDb) {
  skuMap = skuMap || {};
  learnedRules = learnedRules || [];
  var raw = parseOrders(rows);
  var items = [];
  raw.forEach(function (it) {
    var finalProduct = cleanProduct(it.name, it.variation, it.sku, skuMap);
    if (!finalProduct || finalProduct.trim() === '') return; // ignored (packaging add-on etc)

    // Inventory DB is authoritative only when the match is sufficiently
    // strong. Otherwise keep the established resolver output and make the
    // uncertainty visible in column H rather than guessing.
    var dbMatch = matchInventoryProduct(it, finalProduct, inventoryDb);
    it.dbMatchStatus = dbMatch.status || '';
    it.matchedDbSku = dbMatch.status === 'matched' && dbMatch.record ? String(dbMatch.record.sku || '').trim() : '';
    it.dbMatchNote = dbMatch.note || '';
    it.dbMatchBasis = dbMatch.method ? 'DB: ' + dbMatch.method + (dbMatch.score < 0.999 ? ' (' + Math.round(dbMatch.score * 100) + '%)' : '') : '';
    if (dbMatch.status === 'matched' && dbMatch.record) {
      finalProduct = preferredCompactInventoryName(it, dbMatch.record, finalProduct);
      var mapped = mapInventoryCategory(dbMatch.record.category, finalProduct, dbMatch.record.tags);
      it.dbCategory = mapped.category;
      it.dbCategoryBasis = mapped.basis;
      if (mapped.category !== DEFAULT_CATEGORY) {
        it.category = mapped.category;
      } else {
        it.category = categorize(finalProduct);
        it.dbMatchNote = mapped.basis;
      }
    } else {
      finalProduct = normalizeUnmatchedGameName(it, finalProduct);
      it.category = categorize(finalProduct);
    }

    if (it.qty > 1) finalProduct += ' *' + it.qty;
    if (it.category === undefined) it.category = categorize(finalProduct);
    it.finalProduct = dedupeProductWording(finalProduct.replace(/\s*\(\d+\s*\/\s*[Ss]et\)/g, "").replace(/\s+/g, " ").trim());
    it.needsCjkReview = hasUnexplainedCjk(it.variation);
    it.needsWordingReview = !wasResolved(it.name, it.sku, skuMap);
    items.push(it);
  });

  // Second pass: try to infer a category for anything the explicit rules
  // didn't match. Runs after the first pass so order-co-membership can see
  // the categories its order-mates already resolved to.
  var byOrder = {};
  items.forEach(function (it) {
    if (!byOrder[it.orderCode]) byOrder[it.orderCode] = [];
    if (it.category !== DEFAULT_CATEGORY) byOrder[it.orderCode].push(it.category);
  });
  // If any line in an order is Retail, the entire order stays together in
  // Retail (e.g. a used DualSense plus a game on the same retail order ID).
  var retailOrders = {};
  items.forEach(function (it) { if (it.category === 'Retail') retailOrders[it.orderId] = true; });
  items.forEach(function (it) {
    // Keep 8BitDo in its dedicated Xbox Gaming Acc group even when another
    // line in the same order triggers the Retail order-family rule.
    if (retailOrders[it.orderId] && !/\b8BitDo\b/i.test(it.finalProduct || it.name || '')) it.category = 'Retail';
  });

  items.forEach(function (it) {
    if (it.category !== DEFAULT_CATEGORY) return;
    var guess = inferCategory(it.finalProduct, learnedRules, byOrder[it.orderId]);
    if (guess) {
      it.category = guess.category;
      it.inferredBasis = guess.basis; // shown in the sheet so you can verify
    }
  });

  return items;
}

// Reads your manual corrections back out of the Consolidated sheet and saves
// them to Learned_Rules, so the same item categorizes correctly next run
// without you fixing it again. Run this AFTER correcting categories by hand.
function learnFromCorrections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No "' + OUTPUT_SHEET_NAME + '" sheet found.');
    return;
  }
  var data = sheet.getDataRange().getValues();
  var learnSheet = ss.getSheetByName(LEARNED_RULES_SHEET_NAME);
  if (!learnSheet) {
    learnSheet = ss.insertSheet(LEARNED_RULES_SHEET_NAME);
    learnSheet.getRange(1, 1, 1, 3).setValues([['Keyword / Brand Phrase', 'Category', 'Added']]);
    learnSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    learnSheet.setColumnWidth(1, 260);
    learnSheet.setColumnWidth(2, 200);
  }
  var existing = {};
  var lr = learnSheet.getDataRange().getValues();
  for (var i = 1; i < lr.length; i++) {
    var p = String(lr[i][0] || '').trim().toUpperCase();
    if (p) existing[p] = true;
  }

  // Column G holds the category a human typed in to correct a row.
  var added = [];
  var today = new Date();
  for (var r = 0; r < data.length; r++) {
    var productName = String(data[r][1] || '').trim();
    var correction = String(data[r][6] || '').trim();
    if (!productName || !correction) continue;
    if (CATEGORY_ORDER.indexOf(correction) === -1) continue;
    // Learn the brand token so ONE correction covers that brand's other
    // products too, not just this exact item. Uses the first word when it's
    // distinctive enough to be a brand name; falls back to two words when
    // the first is short or generic (e.g. "The", "Pokemon"), where matching
    // on it alone would wrongly grab unrelated items.
    var words = productName.split(/\s+/);
    var firstWord = (words[0] || '').replace(/[^\w&-]/g, '');
    var GENERIC_FIRST_WORDS = /^(the|new|a|for|top|switch|nintendo|ps4|ps5|xbox|pokemon|one|super|mini|model|blind)$/i;
    var phrase;
    if (firstWord.length >= 4 && !GENERIC_FIRST_WORDS.test(firstWord)) {
      phrase = firstWord;
    } else {
      phrase = words.slice(0, 2).join(' ').trim();
    }
    if (!phrase || existing[phrase.toUpperCase()]) continue;
    existing[phrase.toUpperCase()] = true;
    added.push([phrase, correction, today]);
  }

  if (!added.length) {
    SpreadsheetApp.getUi().alert('No new corrections found.\n\nTo teach the script: type the correct category into column G next to any wrongly-categorized row on the "' + OUTPUT_SHEET_NAME + '" sheet, then run this again.');
    return;
  }
  learnSheet.getRange(learnSheet.getLastRow() + 1, 1, added.length, 3).setValues(added);
  SpreadsheetApp.getUi().alert('Learned ' + added.length + ' new rule(s). They will apply on the next run of Consolidate Orders.');
}

function groupByCategory(items) {
  var buckets = {};
  CATEGORY_ORDER.forEach(function (c) { buckets[c] = []; });
  items.forEach(function (it) { buckets[it.category].push(it); });
  var out = [];
  CATEGORY_ORDER.forEach(function (c) {
    if (buckets[c].length) out.push([c, buckets[c]]);
  });
  return out;
}

// =========================================================================
// SORTING
// =========================================================================
// same order adjacent as one block (using each order's lowest SKU as the
// block's sort key) rather than interleaving them with other orders.
// Groups rows by Order Code and sorts groups by their lowest key (via
// keyFn) — so items from the same order always stay adjacent within a
// category, even though the primary sort is by some other field. Used for
// every category: Retail sorts by real SKU reference, everything else
// sorts by product name — but same-order items stick together either way.
// Two explicit passes, matching the described behavior exactly:
// Pass 1 — sort every row purely by keyFn (SKU for Retail, product name
// for everything else), with no regard for Order ID.
// Pass 2 — walk that sorted list; the first time an Order ID is seen,
// pull in every other row sharing that Order ID right there (in their
// own key order), then skip them when the walk reaches them later.
function sortKeepingOrdersTogether(rows, keyFn) {
  var sorted = rows.slice().sort(function (a, b) {
    var ka = keyFn(a), kb = keyFn(b);
    return ka < kb ? -1 : (ka > kb ? 1 : 0);
  });
  var placed = {};
  var result = [];
  sorted.forEach(function (it) {
    if (placed[it.orderCode]) return;
    var group = sorted.filter(function (x) { return x.orderCode === it.orderCode; });
    result = result.concat(group);
    placed[it.orderCode] = true;
  });
  return result;
}

function sortBySkuKeepingOrdersTogether(rows) {
  return sortKeepingOrdersTogether(rows, function (it) { return (it.sku || '').toUpperCase(); });
}

function sortByProductNameKeepingOrdersTogether(rows) {
  return sortKeepingOrdersTogether(rows, function (it) { return (it.finalProduct || '').toUpperCase(); });
}

// =========================================================================
// SHEET OUTPUT
// =========================================================================
function writeConsolidatedSheet(ss, grouped) {
  var sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(OUTPUT_SHEET_NAME);

  // Count how many line items each order has across the whole output, so
  // any order with more than one item can be highlighted — helpful now
  // that non-Retail categories sort by product name, meaning items from
  // the same order can end up far apart on the sheet.
  var orderCodeCounts = {};
  grouped.forEach(function (pair) {
    pair[1].forEach(function (it) {
      var c = (it.orderCode || '').toUpperCase();
      orderCodeCounts[c] = (orderCodeCounts[c] || 0) + 1;
    });
  });

  var rowsOut = [];
  var sameOrderRows = [];

  grouped.forEach(function (pair) {
    var category = pair[0];
    var rows = category === 'Retail' ? sortBySkuKeepingOrdersTogether(pair[1]) : sortByProductNameKeepingOrdersTogether(pair[1]);

    rows.forEach(function (it) {
      var notes;
      if (it.dbMatchStatus === 'matched') {
        notes = String(it.dbMatchNote || '').indexOf('REVIEW:') === 0 ? [it.dbMatchNote] : [];
      } else {
        notes = [it.dbMatchNote, it.dbCategoryBasis, it.inferredBasis, it.dbMatchBasis].filter(function (v, i, a) { return v && a.indexOf(v) === i; });
      }
      rowsOut.push([it.orderCode, it.finalProduct, it.qty, it.tracking, it.variation, it.sku, it.matchedDbSku || '', '', notes.join(' | ')]);
      if (orderCodeCounts[(it.orderCode || '').toUpperCase()] > 1) sameOrderRows.push(rowsOut.length);
    });

    rowsOut.push(['', '', '', '', '', '', '', '']);
  });

  sheet.getRange(1, 1, rowsOut.length, 9).setValues(rowsOut);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 50);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);

  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 36);
  sheet.setColumnWidth(9, 260);
  sheet.getRange(1, 1, rowsOut.length, 9).setFontFamily('Arial').setFontSize(10);

  // Light blue: this order has more than one line item, possibly scattered
  // elsewhere on the sheet since sorting is now by product name.
  sameOrderRows.forEach(function (rowNum) {
    sheet.getRange(rowNum, 1, 1, 9).setBackground('#cfe2f3');
  });

}


function writeSummarySheet(ss, grouped) {
  var sheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(SUMMARY_SHEET_NAME);

  var rowsOut = [['Category', 'Line Items', 'Total Qty']];
  var totalLines = 0, totalQty = 0;
  var allOrderCodes = {};
  grouped.forEach(function (pair) {
    var rows = pair[1];
    var qty = rows.reduce(function (s, r) { return s + r.qty; }, 0);
    rowsOut.push([pair[0], rows.length, qty]);
    totalLines += rows.length;
    totalQty += qty;
    rows.forEach(function (r) { allOrderCodes[r.orderCode] = true; });
  });
  rowsOut.push(['TOTAL', totalLines, totalQty]);
  rowsOut.push(['Total Orders (unique)', Object.keys(allOrderCodes).length, '']);

  sheet.getRange(1, 1, rowsOut.length, 3).setValues(rowsOut);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setFontColor('#ffffff').setBackground('#4472c4');
  sheet.getRange(rowsOut.length - 1, 1, 1, 3).setFontWeight('bold');
  sheet.getRange(rowsOut.length, 1, 1, 3).setFontWeight('bold');
}


// =========================================================================
// MAIN ENTRY POINT
// =========================================================================
function consolidateOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName(RAW_SHEET_NAME);
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert('No sheet named "' + RAW_SHEET_NAME + '" found.');
    return;
  }
  var skuMap = buildSkuMap();
  var inventoryDb = buildInventoryDbIndex();
  var data = rawSheet.getDataRange().getValues();
  var learnedRules = buildLearnedRules();
  var items = buildConsolidated(data, skuMap, learnedRules, inventoryDb);
  if (!items.length) {
    SpreadsheetApp.getUi().alert('No line items were parsed from "' + RAW_SHEET_NAME + '".');
    return;
  }
  var grouped = groupByCategory(items);
  writeConsolidatedSheet(ss, grouped);
  writeSummarySheet(ss, grouped);
  tuckAwaySkuMasterSheet(ss);
}

// Keeps SKU_MASTER out of the way — hidden, and pushed to the last tab
// position — so it doesn't clutter the tab bar or get clicked into by
// accident. Safe to call every run: hiding an already-hidden sheet and
// moving an already-last sheet are both no-ops.
function tuckAwaySkuMasterSheet(ss) {
  var sheet = ss.getSheetByName(SKU_MASTER_SHEET_NAME);
  if (!sheet) return;
  var lastIndex = ss.getSheets().length;
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(lastIndex);
  sheet.hideSheet();
}
