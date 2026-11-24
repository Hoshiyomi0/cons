const $ = (id) => document.getElementById(id);
let inventoryFile = null;
let orderFile = null;
let currentResult = null;
let toastTimer;

function toast(message) {
  const el = $('toast'); el.textContent = message; el.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove('show'), 2400);
}
function setMessage(id, message, kind = '') {
  const el = $(id); el.textContent = message; el.className = 'message' + (kind ? ' ' + kind : '');
}
function fileLabel(id, file) { $(id).textContent = file ? `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB` : 'No file selected'; }
function updateButtons() {
  $('uploadInventory').disabled = !inventoryFile;
  $('processOrders').disabled = !orderFile;
}
function bindFileInput(inputId, dropId, onFile) {
  const input = $(inputId), drop = $(dropId);
  input.addEventListener('change', () => onFile(input.files[0] || null));
  ['dragenter', 'dragover'].forEach(type => drop.addEventListener(type, e => { e.preventDefault(); drop.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach(type => drop.addEventListener(type, e => { e.preventDefault(); drop.classList.remove('dragover'); }));
  drop.addEventListener('drop', e => { const file = e.dataTransfer.files[0]; if (file) onFile(file); });
}
bindFileInput('inventoryFile', 'inventoryDrop', file => {
  inventoryFile = file; fileLabel('inventoryFilename', file); updateButtons(); setMessage('inventoryMessage', '');
});
bindFileInput('orderFile', 'orderDrop', file => {
  orderFile = file; fileLabel('orderFilename', file); updateButtons(); setMessage('processMessage', '');
});

async function postFile(url, file, messageId, buttonId) {
  const button = $(buttonId); button.disabled = true; const oldText = button.textContent;
  button.textContent = 'Working…'; setMessage(messageId, 'Uploading and processing file…');
  const data = new FormData(); data.append('file', file);
  try {
    const response = await fetch(url, { method: 'POST', body: data });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Request failed.');
    setMessage(messageId, result.inventoryRows !== undefined ? `${result.inventoryRows.toLocaleString()} Inventory rows saved.` : `${result.summary?.lineItems ?? 0} order lines processed.`, 'success');
    return result;
  } catch (error) {
    setMessage(messageId, error.message, 'error'); throw error;
  } finally { button.textContent = oldText; updateButtons(); }
}

$('uploadInventory').addEventListener('click', async () => {
  if (!inventoryFile) return;
  try {
    const result = await postFile('/api/inventory', inventoryFile, 'inventoryMessage', 'uploadInventory');
    $('dbPill').className = 'status-pill ready'; $('dbPill').innerHTML = '<span class="dot"></span><span>Inventory ready · ' + result.inventoryRows.toLocaleString() + ' products</span>';
    toast('Inventory Database updated');
  } catch (_) {}
});
$('processOrders').addEventListener('click', async () => {
  if (!orderFile) return;
  try {
    const result = await postFile('/api/process', orderFile, 'processMessage', 'processOrders');
    currentResult = result; renderResults(result); toast('Order processing complete');
  } catch (_) {}
});

function renderResults(result) {
  $('results').classList.remove('hidden');
  $('resultMeta').textContent = `${result.filename} · ${result.summary.lineItems} line items`;
  const s = result.summary;
  $('stats').innerHTML = [
    ['Line items', s.lineItems], ['Matched SKU', s.matched], ['Unmatched / review', s.unmatched], ['Categories', s.categories]
  ].map(([label, value]) => `<div class="stat"><strong>${Number(value || 0).toLocaleString()}</strong><span>${label}</span></div>`).join('');
  const thead = $('resultTable').querySelector('thead');
  const tbody = $('resultTable').querySelector('tbody');
  thead.innerHTML = '<tr>' + result.columns.map((c, i) => `<th class="${i === 7 ? 'blank-col' : ''}">${escapeHtml(c || '')}</th>`).join('') + '</tr>';
  tbody.innerHTML = result.rows.map((row, index) => {
    const isBlank = row.every(v => v === '' || v === null || v === undefined);
    return `<tr data-row="${index}" class="${isBlank ? 'separator' : ''}">` + row.map((cell, i) => {
      const cls = i === 7 ? 'blank-col' : (i === 6 ? (cell ? 'match' : 'unmatched') : '');
      return `<td class="${cls}">${escapeHtml(cell ?? '')}</td>`;
    }).join('') + '</tr>';
  }).join('');
  $('searchBox').value = '';
  $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}
function escapeHtml(value) { return String(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
$('searchBox').addEventListener('input', () => {
  const q = $('searchBox').value.trim().toLowerCase();
  [...$('resultTable').querySelectorAll('tbody tr')].forEach(tr => {
    tr.hidden = q && !tr.textContent.toLowerCase().includes(q);
  });
});
function rowsToTsv(rows) {
  const data = $('includeHeaders').checked ? [currentResult.columns, ...rows] : rows;
  return data.map(row => row.map(cell => String(cell ?? '').replace(/\t/g, ' ').replace(/\r?\n/g, ' ')).join('\t')).join('\n');
}
async function copyRows(matchedOnly) {
  if (!currentResult) return;
  let rows = currentResult.rows;
  if (matchedOnly) rows = rows.filter(row => row.some(v => v !== '' && v !== null) && Boolean(row[6]));
  else rows = rows.filter(row => {
    const tr = $('resultTable').querySelector(`tr[data-row="${currentResult.rows.indexOf(row)}"]`);
    return tr && !tr.hidden;
  });
  const text = rowsToTsv(rows);
  try { await navigator.clipboard.writeText(text); toast(`${rows.length} rows copied`); }
  catch (_) {
    const area = document.createElement('textarea'); area.value = text; document.body.appendChild(area); area.select();
    const ok = document.execCommand('copy'); area.remove(); toast(ok ? `${rows.length} rows copied` : 'Copy failed. Select the table and copy manually.');
  }
}
$('copyAll').addEventListener('click', () => copyRows(false));
$('copyMatched').addEventListener('click', () => copyRows(true));

async function refreshStatus() {
  try {
    const res = await fetch('/api/status'); const data = await res.json();
    if (data.inventoryReady) {
      $('dbPill').className = 'status-pill ready';
      $('dbPill').innerHTML = `<span class="dot"></span><span>Inventory ready · ${data.inventoryRows.toLocaleString()} products</span>`;
    }
  } catch (_) {}
}
refreshStatus();
