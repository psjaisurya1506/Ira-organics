import './style.css';
import * as XLSX from 'xlsx';
import { db, configured, ref, onValue, set, remove, get, runTransaction } from './firebase.js';

/* ---------- helpers ---------- */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const r2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;
const inr = (n) => '₹' + r2(n).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const pad = (n) => String(n).padStart(2, '0');
const localDate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; // local, not UTC
const localTime = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
const num = (v) => (Number.isFinite(parseFloat(v)) ? parseFloat(v) : 0);
const int = (v) => (Number.isFinite(parseInt(v, 10)) ? parseInt(v, 10) : 0);

const DEFAULT_CATEGORIES = ['DALL', 'MILLET', 'RICE', 'SNACKS'];
const BADGES = { RICE: 'badge-rice', MILLET: 'badge-millet', DALL: 'badge-dall', SNACKS: 'badge-snacks' };
const badge = (c) => `<span class="badge ${BADGES[String(c).toUpperCase()] || 'badge-default'}">${esc(c)}</span>`;

/* ---------- state ---------- */
let categories = [], inventory = [], salesLog = [];
const ready = { categories: false, inventory: false, sales: false };
let viewMode = 'STOCK', formMode = 'ADD', terminalGroup = 'All', cart = [];

/* ---------- cloud sync ---------- */
let unsubs = [];
function setStatus(text, color = '#fff') { const el = $('connStatus'); el.textContent = text; el.style.color = color; }
function stopCloudSync() {
  unsubs.forEach((u) => u()); unsubs = [];
  categories = []; inventory = []; salesLog = [];
  Object.keys(ready).forEach((k) => (ready[k] = false));
}
function initCloudSync() {
  stopCloudSync();
  const onErr = (e) => setStatus(/permission/i.test(e.message) ? '● Permission denied – check your Realtime Database rules' : `● Database error: ${e.message}`, '#ffcdd2');
  const watch = (path, handler) => unsubs.push(onValue(ref(db, path), (s) => { handler(s); render(); }, onErr));
  unsubs.push(onValue(ref(db, '.info/connected'), (s) =>
    setStatus(s.val() ? '● Connected' : '● Offline – changes will sync when you reconnect', s.val() ? '#c8e6c9' : '#ffe0b2'), onErr));
  watch('categories', (s) => { categories = s.exists() ? Object.values(s.val()) : []; ready.categories = true; });
  watch('inventory', (s) => { inventory = s.exists() ? Object.entries(s.val()).map(([id, v]) => ({ ...v, id })) : []; ready.inventory = true; });
  watch('salesLog', (s) => {
    salesLog = s.exists() ? Object.entries(s.val()).map(([id, v]) => ({ ...v, id })) : [];
    salesLog.sort((a, b) => `${b.date} ${b.time || ''}${b.id}`.localeCompare(`${a.date} ${a.time || ''}${a.id}`));
    ready.sales = true;
  });
}

// Seed default categories exactly once (so deleting them all later doesn't bring them back)
async function seedDefaults() {
  const seeded = await get(ref(db, 'meta/seeded'));
  if (seeded.exists()) return;
  const cats = await get(ref(db, 'categories'));
  if (!cats.exists()) await set(ref(db, 'categories'), DEFAULT_CATEGORIES);
  await set(ref(db, 'meta/seeded'), true);
}

/* ---------- rendering ---------- */
function render() {
  if (!Object.values(ready).every(Boolean)) return;
  refreshCategoryDropdowns();
  if (viewMode === 'SALES') { renderTerminalGroupButtons(); refreshTerminalProductDropdown(); }

  const day = $('globalDateFilter').value;
  const cat = $('filterCategory').value || 'All';
  $('filterStatusBadge').textContent = day ? `Active Range Scope: [${day}] Only` : 'Active Range Scope: [All-Time Metrics Summary]';

  const stockVal = inventory.reduce((t, i) => t + (i.cost || 0) * (i.qty || 0), 0);
  const sales = day ? salesLog.filter((s) => s.date === day) : salesLog;
  const profit = sales.reduce((t, s) => t + (s.price || 0) - (s.cost || 0), 0);
  const units = sales.reduce((t, s) => t + (s.qty || 1), 0);
  $('dispStockVal').textContent = inr(stockVal);
  $('dispNetProfit').textContent = inr(profit);
  $('dispNetProfit').className = `value ${profit >= 0 ? 'profit-positive' : 'profit-negative'}`;
  $('dispTotalSold').textContent = `${units} units`;

  const items = inventory.filter((i) => cat === 'All' || i.category === cat).sort((a, b) => a.name.localeCompare(b.name));
  $('inventoryBody').innerHTML = items.map((i) => `
    <tr>
      <td><strong>${esc(i.name)}</strong></td>
      <td>${badge(i.category)}</td>
      <td><span class="badge badge-default" style="background-color:#616161;">${esc(i.unit || 'Unit')}</span></td>
      <td>Cost: ${inr(i.cost)}<br><span style="color:var(--primary-green);font-weight:600;">Sale: ${inr(i.price)}</span></td>
      <td><div style="display:flex;align-items:center;gap:6px;">
        <button class="action-btn" data-act="adj" data-d="-1" data-id="${esc(i.id)}" style="background:#777;">-</button>
        <span style="font-weight:bold;min-width:25px;text-align:center;">${i.qty ?? 0}</span>
        <button class="action-btn" data-act="adj" data-d="1" data-id="${esc(i.id)}">+</button></div></td>
      <td>${esc(i.stockEntryDate || '—')}</td>
      <td><button class="action-btn" style="background-color:#0288d1;" data-act="edit" data-id="${esc(i.id)}">✏️</button></td>
      <td><button class="danger action-btn" data-act="del" data-id="${esc(i.id)}">✕</button></td>
    </tr>`).join('') || '<tr><td colspan="8" style="text-align:center;color:#999;">No products found.</td></tr>';

  $('salesLedgerBody').innerHTML = sales.map((s) => {
    const p = (s.price || 0) - (s.cost || 0);
    return `<tr>
      <td><small>${esc(s.date)} ${esc(s.time || '')}</small></td>
      <td><strong>${esc(s.name)}</strong> <small>(${s.qty || 1} ${esc(s.unit || 'Unit')})</small></td>
      <td>${badge(s.category)}</td>
      <td>${inr(s.price)}<br><small style="color:${p >= 0 ? 'green' : 'red'}">Profit: ${inr(p)}</small></td>
      <td><button class="danger action-btn" data-act="reverse" data-id="${esc(s.id)}">Reverse</button></td></tr>`;
  }).join('') || '<tr><td colspan="5" style="text-align:center;color:#999;">No transaction logs matched active range timeline.</td></tr>';

  $('categoryManagementList').innerHTML = categories.map((c) => `
    <div class="category-item">${badge(c)}
      <button type="button" class="danger action-btn" data-act="rmcat" data-cat="${esc(c)}">✕</button></div>`).join('');

  if (formMode === 'EDIT') populateEditProductPicker();
  updateStockExportPreview();
  updateSalesExportPreview();
}

function refreshCategoryDropdowns() {
  const opts = '<option value="All">All Groups</option>' + categories.map((c) => `<option value="${esc(c)}">${esc(c)}</option>`).join('');
  const keep = (el, html, fallback) => { const prev = el.value || fallback; el.innerHTML = html; el.value = [...el.options].some((o) => o.value === prev) ? prev : fallback; };
  keep($('prodCategory'), categories.map((c) => `<option value="${esc(c)}">${esc(c)}</option>`).join(''), categories[0] ?? '');
  ['filterCategory', 'stockExportCategoryFilter', 'salesExportCategoryFilter'].forEach((id) => keep($(id), opts, 'All'));
}

/* ---------- products ---------- */
$('productForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const id = $('editingProductId').value;
  const p = {
    name: $('prodName').value.trim(), category: $('prodCategory').value,
    unit: $('prodUnit').value.trim() || 'Unit',
    cost: r2(num($('prodCost').value)), price: r2(num($('prodPrice').value)),
    qty: Math.max(0, int($('prodQty').value)),
  };
  if (!p.name) return alert('Item name is required.');
  if (!p.category) return alert('Create a stock group first.');
  const dup = inventory.find((i) => i.id !== id && i.name.toLowerCase() === p.name.toLowerCase() && i.unit === p.unit && i.category === p.category);
  if (dup) return alert('A product with the same name, group and unit already exists.');
  if (p.price < p.cost && !confirm('Selling price is below cost price. Save anyway?')) return;

  if (formMode === 'EDIT' && id) {
    const existing = inventory.find((item) => item.id === id);
    const stockEntryDate = existing?.qty !== p.qty ? localDate(new Date()) : existing?.stockEntryDate;
    await set(ref(db, `inventory/${id}`), { ...p, id, ...(stockEntryDate ? { stockEntryDate } : {}) });
    switchFormMode('ADD');
  } else {
    const nid = `${Date.now()}`;
    await set(ref(db, `inventory/${nid}`), { ...p, id: nid, stockEntryDate: localDate(new Date()) });
    e.target.reset();
  }
});

function loadProductIntoForm() {
  const p = inventory.find((i) => i.id === $('editProductPicker').value);
  if (!p) return;
  $('editingProductId').value = p.id; $('prodName').value = p.name; $('prodCategory').value = p.category;
  $('prodUnit').value = p.unit || 'Unit'; $('prodCost').value = p.cost; $('prodPrice').value = p.price; $('prodQty').value = p.qty;
}
const cancelEdit = () => switchFormMode('ADD');

function populateEditProductPicker() {
  const picker = $('editProductPicker'), prev = picker.value;
  picker.innerHTML = '<option value="">-- Choose a product --</option>' +
    [...inventory].sort((a, b) => a.name.localeCompare(b.name)).map((p) => `<option value="${esc(p.id)}">${esc(p.name)}</option>`).join('');
  if (inventory.some((i) => i.id === prev)) picker.value = prev;
}

function switchFormMode(mode) {
  formMode = mode;
  const edit = mode === 'EDIT';
  $('btnSwitchAdd').style.opacity = edit ? '0.5' : '1';
  $('btnSwitchEdit').style.opacity = edit ? '1' : '0.5';
  $('formPanelTitle').textContent = edit ? '✏️ Edit Existing Stock Item' : '➕ Add New Stock Item';
  $('editPickerRow').style.display = edit ? 'block' : 'none';
  $('formSubmitBtn').textContent = edit ? 'Update Item Metrics' : 'Add Item';
  $('formSubmitBtn').style.backgroundColor = edit ? '#0288d1' : 'var(--primary-green)';
  if (edit) { populateEditProductPicker(); loadProductIntoForm(); }
  else { $('productForm').reset(); $('editingProductId').value = ''; }
}

// Atomic stock changes: safe with multiple users/devices
async function adjustStock(id, d) {
  let changed = false;
  const result = await runTransaction(ref(db, `inventory/${id}/qty`), (q) => {
    const quantity = q || 0;
    const nextQuantity = Math.max(0, quantity + d);
    changed = nextQuantity !== quantity;
    return nextQuantity;
  });
  if (result.committed && changed) await set(ref(db, `inventory/${id}/stockEntryDate`), localDate(new Date()));
}
const deleteProduct = (id) => confirm('Delete this product from the database?') && remove(ref(db, `inventory/${id}`));

/* ---------- categories ---------- */
async function addNewCategory() {
  const input = $('newCategoryInput'), v = input.value.trim().toUpperCase();
  if (!v) return alert('Please type a valid category name.');
  if (categories.includes(v)) return alert('Group already listed.');
  await set(ref(db, 'categories'), [...categories, v].sort());
  input.value = '';
}
async function removeCategory(name) {
  const used = inventory.filter((i) => i.category === name).length;
  if (used) return alert(`Cannot remove "${name}": ${used} product(s) still use it.`);
  if (confirm(`Remove category group "${name}"?`)) await set(ref(db, 'categories'), categories.filter((c) => c !== name));
}

/* ---------- sales terminal ---------- */
const inCart = (id) => cart.find((c) => c.productId === id)?.qty || 0;

function refreshTerminalProductDropdown() {
  const sel = $('salesTerminalProductSelect'), prev = sel.value;
  const list = inventory.filter((i) => terminalGroup === 'All' || i.category === terminalGroup).sort((a, b) => a.name.localeCompare(b.name));
  sel.innerHTML = list.map((i) => `<option value="${esc(i.id)}">${esc(i.name)} (${esc(i.unit || 'Unit')} | Stock: ${Math.max(0, i.qty - inCart(i.id))} | Price: ${inr(i.price)})</option>`).join('');
  if (list.some((i) => i.id === prev)) sel.value = prev;
  updateTerminalPriceDisplay();
}
function renderTerminalGroupButtons() {
  $('terminalGroupButtons').innerHTML = ['All', ...categories].map((c) =>
    `<button type="button" class="group-search-btn ${terminalGroup === c ? 'active-group' : ''}" data-act="group" data-cat="${esc(c)}">${c === 'All' ? 'ALL DISPLAYS' : esc(c)}</button>`).join('');
}
function updateTerminalPriceDisplay() {
  const item = inventory.find((i) => i.id === $('salesTerminalProductSelect').value);
  $('salesTerminalPriceDisplay').value = item ? inr(item.price * int($('salesTerminalQtyInput').value)) : '₹0.00';
}
function addItemToCart() {
  const item = inventory.find((i) => i.id === $('salesTerminalProductSelect').value);
  if (!item) return alert('Select a product first.');
  const qty = int($('salesTerminalQtyInput').value);
  if (qty <= 0) return alert('Enter a valid quantity.');
  if (inCart(item.id) + qty > item.qty) return alert(`Insufficient stock. Only ${item.qty - inCart(item.id)} more available.`);
  const line = cart.find((c) => c.productId === item.id);
  if (line) line.qty += qty;
  else cart.push({ productId: item.id, name: item.name, category: item.category, unit: item.unit || 'Unit', unitPrice: item.price, costPrice: item.cost || 0, qty });
  $('salesTerminalQtyInput').value = '1';
  renderCart(); refreshTerminalProductDropdown();
}
function renderCart() {
  $('cartTableBody').innerHTML = cart.map((c, i) => `
    <tr><td><strong>${esc(c.name)}</strong> ${badge(c.category)}</td>
    <td><span class="badge badge-default" style="background-color:#616161;">${esc(c.unit)}</span></td>
    <td>${inr(c.unitPrice)}</td><td>${c.qty} items</td><td><strong>${inr(c.unitPrice * c.qty)}</strong></td>
    <td><button type="button" class="danger action-btn" data-act="rmcart" data-i="${i}">✕</button></td></tr>`).join('') ||
    '<tr><td colspan="6" style="text-align:center;color:#999;">Cart is empty. Add matching ledger products above.</td></tr>';
  $('cartGrandTotalDisplay').textContent = inr(cart.reduce((t, c) => t + c.unitPrice * c.qty, 0));
}
async function executeTerminalCheckout() {
  try { await runCheckout(); }
  catch (err) { alert(`Checkout was interrupted (${err.message}). Check stock and the sales list before retrying.`); }
}
async function runCheckout() {
  if (!cart.length) return alert('Your cart is empty.');
  const now = new Date(), failed = [];
  for (const it of cart) {
    // Re-check stock atomically at the moment of sale
    const res = await runTransaction(ref(db, `inventory/${it.productId}/qty`), (q) => (q === null ? 0 : q >= it.qty ? q - it.qty : undefined));
    if (!res.committed) { failed.push(it.productId); continue; }
    const id = `${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    await set(ref(db, `salesLog/${id}`), {
      id, productId: it.productId, date: localDate(now), time: localTime(now), name: it.name, category: it.category,
      unit: it.unit, qty: it.qty, price: r2(it.unitPrice * it.qty), cost: r2(it.costPrice * it.qty),
    });
  }
  cart = cart.filter((c) => failed.includes(c.productId));
  renderCart(); refreshTerminalProductDropdown();
  alert(failed.length ? 'Some items could not be sold because stock changed. They remain in the cart.' : '⚡ Checkout completed successfully!');
}
const removeCartItem = (i) => { cart.splice(i, 1); renderCart(); refreshTerminalProductDropdown(); };
const setTerminalFilterGroup = (g) => { terminalGroup = g; renderTerminalGroupButtons(); refreshTerminalProductDropdown(); };

async function removeSale(id) {
  const s = salesLog.find((x) => x.id === id);
  if (!s || !confirm('Reverse this sale and restore stock?')) return;
  // Prefer the stored productId; fall back to name+category for sales logged before this version
  const inv = inventory.find((i) => i.id === s.productId) || inventory.find((i) => i.name === s.name && i.category === s.category);
  if (inv) await runTransaction(ref(db, `inventory/${inv.id}/qty`), (q) => (q || 0) + (s.qty || 1));
  await remove(ref(db, `salesLog/${id}`));
}

/* ---------- view switching ---------- */
function switchViewMode(mode) {
  viewMode = mode;
  $('btnModeStock').classList.toggle('active', mode === 'STOCK');
  $('btnModeSales').classList.toggle('active', mode === 'SALES');
  $('sectionStockManagement').style.display = mode === 'STOCK' ? 'grid' : 'none';
  $('sectionSalesManagement').style.display = mode === 'SALES' ? 'grid' : 'none';
  render();
}
function clearDateFilter() { $('globalDateFilter').value = ''; render(); }

/* ---------- Excel export ---------- */
function updateStockExportPreview() {
  const c = $('stockExportCategoryFilter').value || 'All';
  $('stockExportPreview').textContent = `${inventory.filter((i) => c === 'All' || i.category === c).length} rows ready for export.`;
}
const monthSales = () => {
  const m = $('monthlyFilterPicker').value, c = $('salesExportCategoryFilter').value || 'All';
  return { m, c, rows: salesLog.filter((s) => s.date?.startsWith(m) && (c === 'All' || s.category === c)) };
};
function updateSalesExportPreview() {
  const { m, rows } = monthSales();
  $('salesExportPreview').textContent = m ? `${rows.length} matching sales records found for [${m}].` : 'Select a month above';
}
function download(rows, sheet, file) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), sheet);
  XLSX.writeFile(wb, file);
}
function exportStockExcel() {
  const c = $('stockExportCategoryFilter').value || 'All';
  const data = inventory.filter((i) => c === 'All' || i.category === c);
  if (!data.length) return alert('No stock rows match the filter.');
  download(data.map((i) => ({
    'Product ID': i.id, 'Item Name': i.name, Category: i.category, Unit: i.unit || 'Unit', 'Cost Price (₹)': i.cost,
    'Selling Price (₹)': i.price, 'Qty In Stock': i.qty, 'Stock Entry Date': i.stockEntryDate || '',
    'Stock Value at Cost (₹)': r2(i.cost * i.qty),
  })), 'Inventory', `Ira_Organics_Stock_Report_${c}.xlsx`);
}
const salesRows = (list) => list.map((s) => ({
  Date: s.date, Time: s.time || '', Product: s.name, Category: s.category, Unit: s.unit || 'Unit', Quantity: s.qty || 1,
  'Revenue (₹)': s.price, 'Cost (₹)': s.cost, 'Profit (₹)': r2(s.price - s.cost),
}));
function exportMonthlySalesExcel() {
  const { m, c, rows } = monthSales();
  if (!m) return alert('Select a month first.');
  if (!rows.length) return alert('No sales records match the filters.');
  download(salesRows(rows), 'Sales Ledger', `Ira_Organics_Sales_${m}_${c}.xlsx`);
}
function exportAllSalesExcel() {
  if (!salesLog.length) return alert('No sales have been logged yet.');
  download(salesRows(salesLog), 'Sales Ledger', 'Ira_Organics_Sales_AllTime.xlsx');
}

/* ---------- events & boot ---------- */
// Dynamic table/button actions use data attributes (safe for names containing quotes)
document.addEventListener('click', (e) => {
  const b = e.target.closest('[data-act]');
  if (!b) return;
  const { act, id, cat, i, d } = b.dataset;
  ({
    adj: () => adjustStock(id, int(d)),
    edit: () => { switchFormMode('EDIT'); $('editProductPicker').value = id; loadProductIntoForm(); },
    del: () => deleteProduct(id), reverse: () => removeSale(id), rmcat: () => removeCategory(cat),
    rmcart: () => removeCartItem(int(i)), group: () => setTerminalFilterGroup(cat),
  })[act]?.();
});

// Handlers referenced by static markup in index.html
Object.assign(window, {
  renderAll: render, switchViewMode, clearDateFilter, switchFormMode, loadProductIntoForm, cancelEdit, addNewCategory,
  updateStockExportPreview, updateSalesExportPreview, exportStockExcel, exportMonthlySalesExcel, exportAllSalesExcel,
  updateTerminalPriceDisplay, addItemToCart, executeTerminalCheckout,
});

const now = new Date();
$('globalDateFilter').value = localDate(now);
$('monthlyFilterPicker').value = `${now.getFullYear()}-${pad(now.getMonth() + 1)}`;
if (!configured) {
  setStatus('● Not configured – add VITE_FIREBASE_DATABASE_URL to your .env file and restart', '#ffcdd2');
} else {
  initCloudSync();
  seedDefaults().catch((e) => setStatus(`● Could not set up defaults: ${e.message}`, '#ffcdd2'));
}
