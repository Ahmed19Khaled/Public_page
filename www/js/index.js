/* ════════════════════════════════════════
   GRANADA PRINT - CORRECTED / ID-BASED ENGINE
   - Works with your exact template IDs
   - Barcode centered in preview + print
   - Supports Cordova + Web Preview
   - No duplicate declarations
═══════════════════════════════════════════ */

document.addEventListener('deviceready', onDeviceReady, false);

function onDeviceReady() {
  console.log('Running cordova-' + cordova.platformId + '@' + cordova.version);
  const el = document.getElementById('deviceready');
  if (el) el.classList.add('ready');
}

/* ════════ CORE STATE ════════ */
const SK = 'gp6_inv';
const DK = 'gp6_draft';
const LK = 'gp6_logo';
const INV_COUNTER_KEY = 'gp_invoice_counter';

let logoData = '';
let editId = null;
let scanStream = null;
let scanReader = null;
let camFace = 'environment';
let selected = new Set();
let draftTmr = null;

const $ = id => document.getElementById(id);
const v = id => ($(id)?.value || '').trim();
const today = () => new Date().toISOString().slice(0, 10);
const esc = s => String(s || '').replace(/[&<>"']/g, c => ({
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;'
}[c]));
const money = n => new Intl.NumberFormat('ar-EG').format(Number(n) || 0) + ' ج.م';
const money2 = n => new Intl.NumberFormat('ar-EG').format(Number(n) || 0) + ' ج.م';
const niceDate = iso => {
  if (!iso) return '-';
  const d = new Date(iso);
  return isNaN(d)
    ? iso
    : new Intl.DateTimeFormat('ar-EG', { day: 'numeric', month: 'short', year: 'numeric' }).format(d);
};
const nd = niceDate;

const loadList = () => {
  try { return JSON.parse(localStorage.getItem(SK) || '[]'); }
  catch { return []; }
};
const saveList = l => localStorage.setItem(SK, JSON.stringify(l));

const txt = (id, t) => { const e = $(id); if (e) e.textContent = t; };
const html = (id, h) => { const e = $(id); if (e) e.innerHTML = h; };

function setTextAny(selectors, value) {
  for (const sel of selectors) {
    const el = document.querySelector(sel);
    if (el) { el.textContent = value; return; }
  }
}

/* ════════ LOGO ════════ */
function saveLogo() {
  if (logoData) localStorage.setItem(LK, logoData);
}

function renderLogoThumb(src) {
  const t = $('logoThumb');
  if (!t) return;
  t.innerHTML = src
    ? `<img src="${src}" style="width:100%;height:100%;object-fit:contain">`
    : `<i class="fa-solid fa-building"></i>`;
}

function loadSavedLogo() {
  try {
    const savedLogo = localStorage.getItem(LK);
    if (!savedLogo) return;
    logoData = savedLogo;
    renderLogoThumb(savedLogo);
    updateUI();
  } catch { }
}

function loadLogo(input) {
  if (!input.files || !input.files[0]) return;

  const reader = new FileReader();
  reader.onload = e => {
    logoData = e.target.result;
    localStorage.setItem(LK, logoData);
    renderLogoThumb(logoData);
    saveLogo();
    updateUI();
  };
  reader.readAsDataURL(input.files[0]);
}

function loadLogoStored() {
  loadSavedLogo();
}

/* ════════ DATA ════════ */
function getData() {
  return {
    id: editId || null,
    companyName: v('companyName') || 'Granada Print',
    companyPhone: v('companyPhone') || '',
    companyAddress: v('companyAddress') || '',
    invoiceNumber: v('invoiceNumber') || 'GP1170',

    senderName: v('senderName') || '',
    senderPhone: v('senderPhone') || '',
    senderCity: v('senderCity') || '',
    senderAddress: v('senderAddress') || '',

    receiverName: v('receiverName') || '',
    receiverPhone: v('receiverPhone') || '',
    receiverCity: v('receiverCity') || '',
    receiverAddress: v('receiverAddress') || '',

    content: v('content') || '',
    deliveryType: v('deliveryType') || 'توصيل',
    pieces: v('pieces') || '1',
    weight: v('weight') || '0.5',
    shippingCompany: v('shippingCompany') || '',
    codAmount: v('codAmount') || '0',
    shipDate: v('shipDate') || today(),
    notes: v('notes') || '',

    companyVisible: $('companyVisible')?.checked !== false,
    logo: logoData,
    printSize: $('printSize')?.value || 'thermal80',
    createdAt: new Date().toISOString()
  };
}

function setData(d) {
  editId = d.id || null;

  [
    'companyName', 'companyPhone', 'companyAddress', 'invoiceNumber',
    'senderName', 'senderPhone', 'senderCity', 'senderAddress',
    'receiverName', 'receiverPhone', 'receiverCity', 'receiverAddress',
    'content', 'pieces', 'weight', 'shippingCompany', 'codAmount', 'shipDate', 'notes'
  ].forEach(k => {
    if ($(k)) $(k).value = d[k] ?? '';
  });

  if ($('deliveryType')) $('deliveryType').value = d.deliveryType || 'توصيل';
  if ($('companyVisible')) $('companyVisible').checked = d.companyVisible !== false;
  if (d.printSize && $('printSize')) $('printSize').value = d.printSize;

  logoData = d.logo || logoData || '';
  renderLogoThumb(logoData);
  updateUI();
  saveDraft();
}

/* ════════ BARCODE RENDERING ════════ */
function makeBarcodeSVG(value, opts = {}) {
  const code = String(value || 'GP0000').trim();
  const h = opts.height ?? 30;
  const w = opts.width ?? 1.4;
  const dark = opts.dark ?? '#1a1a1a';

  const wrapStyle = `
    display:flex;
    justify-content:center;
    align-items:center;
    width:100%;
  `;

  if (typeof JsBarcode === 'undefined') {
    return `<div style="${wrapStyle};font-family:'IBM Plex Mono',monospace;font-size:11px;letter-spacing:1px;text-align:center">${esc(code)}</div>`;
  }

  try {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    JsBarcode(svg, code, {
      format: 'CODE128',
      displayValue: false,
      margin: 0,
      lineColor: dark,
      background: 'transparent',
      width: w,
      height: h
    });

    svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
    svg.style.display = 'block';
    svg.style.margin = '0 auto';
    svg.style.maxWidth = '100%';
    svg.style.width = opts.fullWidth ? '100%' : 'auto';
    svg.style.height = 'auto';

    return `<div class="barcode-box" style="${wrapStyle}">${svg.outerHTML}</div>`;
  } catch (e) {
    console.error('Barcode error:', e);
    return `<div style="${wrapStyle};font-family:'IBM Plex Mono',monospace;font-size:11px;letter-spacing:1px;text-align:center">${esc(code)}</div>`;
  }
}

function drawBC(id, val, h, w, dark) {
  const el = $(id);
  if (!el) return;

  el.innerHTML = '';
  const code = String(val || 'GP0000').trim();

  if (typeof JsBarcode === 'undefined') {
    el.outerHTML = `<div id="${id}" class="barcode-box" style="display:flex;justify-content:center;align-items:center;width:100%;font-family:'IBM Plex Mono',monospace;font-size:11px;letter-spacing:1px">${esc(code)}</div>`;
    return;
  }

  try {
    JsBarcode(el, code, {
      format: 'CODE128',
      displayValue: false,
      margin: 0,
      lineColor: dark ? '#c9a227' : '#1a1a1a',
      background: 'transparent',
      width: w || 1.2,
      height: h || 38
    });

    el.setAttribute('preserveAspectRatio', 'xMidYMid meet');
    el.style.display = 'block';
    el.style.margin = '0 auto';
    el.style.maxWidth = '100%';
    el.style.height = 'auto';
  } catch (e) {
    console.error('drawBC error:', e);
    el.innerHTML = '';
  }
}

function applyBarcodePreviewCentering() {
  const previewIds = ['bcT80', 'bcA4', 'bc200side', 'bc150', 'bc70'];

  previewIds.forEach(id => {
    const el = $(id);
    if (!el) return;

    el.style.display = 'block';
    el.style.margin = '0 auto';
    el.style.maxWidth = '100%';
    el.style.height = 'auto';

    const parent = el.parentElement;
    if (parent) {
      parent.style.display = 'flex';
      parent.style.justifyContent = 'center';
      parent.style.alignItems = 'center';
      parent.style.width = '100%';
    }
  });

  const wrappers = document.querySelectorAll('.barcode-box, .t80-bc-wrap, .a4-bc-wrap, .r150-bc, .r70-bc-wrap, .r200-bc-inner');
  wrappers.forEach(w => {
    w.style.display = 'flex';
    w.style.justifyContent = 'center';
    w.style.alignItems = 'center';
    w.style.width = '100%';
    w.style.textAlign = 'center';
  });
}

/* ════════ UPDATE UI ════════ */
function updateUI() {
  const d = getData();
  const sv = d.companyVisible !== false;
  const lHTML = logoData
    ? `<img src="${logoData}" style="width:100%;height:100%;object-fit:contain;border-radius:inherit">`
    : `<i class="fa-solid fa-print"></i>`;

  /* THERMAL 80mm */
  html('logoT80', lHTML);
  txt('coNameT80', d.companyName);
  txt('coPhoneT80', d.companyPhone);
  txt('coAddrT80', d.companyAddress);
  txt('invNumT80', d.invoiceNumber);
  txt('bcNumT80', d.invoiceNumber);
  txt('footNumT80', d.invoiceNumber);
  txt('coFooterT80', d.companyName);
  txt('snNameT80', d.senderName || '-');
  txt('snPhoneT80', d.senderPhone || '-');
  txt('snCityT80', d.senderCity || '-');
  txt('snAddrT80', d.senderAddress || '-');
  txt('rcNameT80', d.receiverName || '-');
  txt('rcPhoneT80', d.receiverPhone || '-');
  txt('rcCityT80', d.receiverCity || '-');
  txt('rcAddrT80', d.receiverAddress || '-');
  txt('contentT80', d.content || '-');
  txt('piecesT80', d.pieces || '1');
  txt('weightT80', (d.weight || '0') + ' كجم');
  txt('dateT80', niceDate(d.shipDate));
  txt('dlvTypeT80', d.deliveryType || '-');
  txt('shipCoT80', d.shippingCompany || '-');
  txt('codT80', money(d.codAmount));
  if ($('notesBoxT80')) $('notesBoxT80').style.display = d.notes ? 'block' : 'none';
  if (d.notes) txt('notesT80', d.notes);
  ['logoT80', 'coNameT80', 'coPhoneT80', 'coAddrT80', 'coFooterT80'].forEach(eid => {
    const e = $(eid);
    if (e) e.style.display = sv ? '' : 'none';
  });
  drawBC('bcT80', d.invoiceNumber, 25, 1.4);

  /* A4 */
  html('logoA4', lHTML);
  txt('coNameA4', d.companyName);
  const infoParts = [];
  if (d.companyAddress) infoParts.push(d.companyAddress);
  if (d.companyPhone) infoParts.push(d.companyPhone);
  if ($('coInfoA4')) $('coInfoA4').innerHTML = infoParts.join('<br>');
  ['coNameA4', 'coInfoA4', 'logoA4'].forEach(eid => {
    const e = $(eid);
    if (e) e.style.display = sv ? '' : 'none';
  });
  txt('invNumA4', d.invoiceNumber);
  txt('bcNumA4', d.invoiceNumber);
  txt('coFootA4', d.companyName);
  txt('snNameA4', d.senderName || '-');
  txt('snPhoneA4', d.senderPhone || '-');
  txt('snCityA4', d.senderCity || '-');
  txt('snAddrA4', d.senderAddress || '-');
  txt('rcNameA4', d.receiverName || '-');
  txt('rcPhoneA4', d.receiverPhone || '-');
  txt('rcCityA4', d.receiverCity || '-');
  txt('rcAddrA4', d.receiverAddress || '-');
  txt('contentA4', d.content || '-');
  txt('piecesA4', d.pieces || '1');
  txt('weightA4', (d.weight || '0') + ' كجم');
  txt('dateA4', niceDate(d.shipDate));
  txt('dlvA4', d.deliveryType || '-');
  txt('shipCoA4', d.shippingCompany || '-');
  txt('codA4', money(d.codAmount));
  if ($('notesBoxA4')) $('notesBoxA4').style.display = d.notes ? 'block' : 'none';
  if (d.notes) txt('notesA4', d.notes);
  drawBC('bcA4', d.invoiceNumber, 25, 2.0);

  /* 200×80 */
  html('logo200', lHTML);
  txt('coName200', d.companyName);
  txt('coPhone200', d.companyPhone);
  txt('invNum200', d.invoiceNumber);
  txt('snName200', d.senderName || '-');
  txt('snPhone200', d.senderPhone || '-');
  txt('snCity200', d.senderAddress || '-');
  txt('rcName200', d.receiverName || '-');
  txt('rcPhone200', d.receiverPhone || '-');
  txt('rcCity200', d.receiverAddress || '-');
  txt('dlv200', d.deliveryType || '-');
  txt('date200', niceDate(d.shipDate));
  txt('content200', d.content || '-');
  txt('pieces200', d.pieces || '1');
  txt('weight200', d.weight || '0');
  txt('shipCo200', d.shippingCompany || '-');
  txt('shipCoFin200', d.shippingCompany || '-');
  txt('cod200', money(d.codAmount));
  txt('coFoot200', (d.companyName || 'GRANADA PRINT').toUpperCase());
  txt('invFoot200', d.invoiceNumber);
  ['coName200', 'coPhone200', 'logo200'].forEach(eid => {
    const e = $(eid);
    if (e) e.style.display = sv ? '' : 'none';
  });
  drawBC('bc200side', d.invoiceNumber, 40, 1.0);

  /* 100×150 */
  html('logo150', lHTML);
  txt('coName150', d.companyName);
  txt('coPhone150', d.companyPhone);
  txt('invNum150', d.invoiceNumber);
  txt('dlv150', d.deliveryType || '-');
  txt('date150', niceDate(d.shipDate));
  txt('date150b', niceDate(d.shipDate));
  txt('snName150', d.senderName || '-');
  setTextAny(['#snPhone150 span', '#snPhone150Span'], d.senderPhone || '-');
  setTextAny(['#snCity150 span', '#snCity150Span'], d.senderCity || '-');
  txt('snAddr150', d.senderAddress || '-');
  txt('rcName150', d.receiverName || '-');
  setTextAny(['#rcPhone150 span', '#rcPhone150Span'], d.receiverPhone || '-');
  setTextAny(['#rcCity150 span', '#rcCity150Span'], d.receiverCity || '-');
  txt('rcAddr150', d.receiverAddress || '-');
  txt('content150', d.content || '-');
  txt('pieces150', d.pieces || '1');
  txt('weight150', d.weight || '0');
  txt('shipCo150', d.shippingCompany || '-');
  txt('cod150', money(d.codAmount));
  txt('bcNum150', d.invoiceNumber);
  txt('coFoot150', d.companyName);
  if ($('notesBox150')) $('notesBox150').style.display = d.notes ? 'flex' : 'none';
  if (d.notes) txt('notes150', d.notes);
  ['coName150', 'coPhone150', 'logo150'].forEach(eid => {
    const e = $(eid);
    if (e) e.style.display = sv ? '' : 'none';
  });
  drawBC('bc150', d.invoiceNumber, 20, 1.2);

  /* 100×70 */
  html('logo70', lHTML);
  txt('coName70', d.companyName);
  txt('coPhone70', d.companyPhone);
  txt('invNum70', d.invoiceNumber);
  txt('bcNum70', d.invoiceNumber);
  txt('footNum70', d.invoiceNumber);
  txt('snName70', d.senderName || '-');
  txt('snPhone70', d.senderPhone || '-');
  txt('snCity70', d.senderAddress || '-');
  txt('rcName70', d.receiverName || '-');
  txt('rcPhone70', d.receiverPhone || '-');
  txt('rcCity70', d.receiverAddress || '-');
  txt('shipCo70', d.shippingCompany || '-');
  txt('cod70', money(d.codAmount));
  txt('coFoot70', d.companyName);
  if ($('r70Head')) $('r70Head').style.display = sv ? 'flex' : 'none';
  drawBC('bc70', d.invoiceNumber, 18, 0.9);

  saveDraftD();
  renderHistory();
  applyBarcodePreviewCentering();
}

function saveDraft() {
  localStorage.setItem(DK, JSON.stringify(getData()));
}

function saveDraftD() {
  clearTimeout(draftTmr);
  draftTmr = setTimeout(saveDraft, 280);
}

/* ════════ SIZE SWITCH ════════ */
function changeSize() {
  const sz = $('printSize')?.value || 'thermal80';
  ['thermal80', 'a4', '200x80', '100x150', '100x70'].forEach(id => {
    const el = $('page-' + id);
    if (el) el.style.display = (id === sz) ? 'block' : 'none';
  });

  const names = {
    thermal80: 'حراري 80mm',
    a4: 'A4 كامل',
    '200x80': '200×80 mm',
    '100x150': '100×150 mm',
    '100x70': '100×70 mm'
  };
  if ($('sizeBadge')) $('sizeBadge').textContent = names[sz] || 'حراري 80mm';

  const zooms = { thermal80: .9, a4: .62, '200x80': .72, '100x150': 1.0, '100x70': 1.05 };
  setZoom(zooms[sz] || .9);
  updateUI();
}

function setZoom(s) {
  const pw = $('previewWrap');
  if (pw) pw.style.transform = `scale(${s})`;
  document.querySelectorAll('.zoom-btn').forEach(b => b.classList.remove('on'));
  if (s === .9) $('z90')?.classList.add('on');
}

/* ════════ INVOICE NUMBER ════════ */
function getNextInvoiceNumber() {
  let counter = parseInt(localStorage.getItem(INV_COUNTER_KEY) || '1170', 10);
  const invoiceNo = 'GP' + counter;
  localStorage.setItem(INV_COUNTER_KEY, String(counter + 1));
  return invoiceNo;
}

function syncCounterFromInvoice(invoiceNumber) {
  const m = String(invoiceNumber || '').match(/GP(\d+)/i);
  if (!m) return;
  const n = parseInt(m[1], 10);
  if (Number.isFinite(n)) {
    const cur = parseInt(localStorage.getItem(INV_COUNTER_KEY) || '1170', 10);
    if (n >= cur) localStorage.setItem(INV_COUNTER_KEY, String(n + 1));
  }
}

function genNum() {
  if ($('invoiceNumber')) $('invoiceNumber').value = getNextInvoiceNumber();
  updateUI();
}

/* ════════ CRUD ════════ */
function saveInvoice() {
  const d = getData();
  const list = loadList();

  const item = {
    ...d,
    id: d.id || crypto.randomUUID(),
    updatedAt: new Date().toISOString()
  };

  const idx = list.findIndex(x => x.id === item.id || x.invoiceNumber === item.invoiceNumber);
  if (idx >= 0) list[idx] = item;
  else list.unshift(item);

  saveList(list);
  editId = item.id;
  syncCounterFromInvoice(item.invoiceNumber);
  renderHistory();
  saveDraft();

  const btn = document.querySelector('.action-bar .btn-gold');
  if (btn) {
    const old = btn.innerHTML;
    btn.innerHTML = '<i class="fa-solid fa-check"></i> تم';
    btn.classList.replace('btn-gold', 'btn-success');
    setTimeout(() => {
      btn.innerHTML = old;
      btn.classList.replace('btn-success', 'btn-gold');
    }, 1400);
  }
}

function newInvoice() {
  if (!confirm('بدء فاتورة جديدة؟')) return;
  editId = null;
  setData({
    companyName: v('companyName') || 'Granada Print',
    companyPhone: v('companyPhone') || '',
    companyAddress: v('companyAddress') || '',
    invoiceNumber: getNextInvoiceNumber(),
    senderName: '',
    senderPhone: '',
    senderCity: '',
    senderAddress: '',
    receiverName: '',
    receiverPhone: '',
    receiverCity: '',
    receiverAddress: '',
    content: '',
    deliveryType: 'توصيل',
    pieces: '1',
    weight: '0.5',
    shippingCompany: '',
    codAmount: '0',
    shipDate: today(),
    notes: '',
    companyVisible: true,
    logo: logoData,
    printSize: $('printSize')?.value || 'thermal80'
  });
}

function openInvoice(id) {
  const item = loadList().find(x => x.id === id);
  if (!item) return;
  setData(item);
  editId = item.id;
  if (item.printSize && $('printSize')) {
    $('printSize').value = item.printSize;
    changeSize();
  }
  renderHistory();
}

function deleteInvoice(id) {
  if (!confirm('حذف هذه الفاتورة نهائيًا؟')) return;
  saveList(loadList().filter(x => x.id !== id));
  selected.delete(id);
  if (editId === id) editId = null;
  renderHistory();
}

function clearHistory() {
  if (!confirm('حذف كل الفواتير المحفوظة؟')) return;
  localStorage.removeItem(SK);
  selected.clear();
  renderHistory();
}

function exportInvoices() {
  const b = new Blob([JSON.stringify(loadList(), null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(b);
  a.download = 'granada_' + new Date().toISOString().slice(0, 10) + '.json';
  a.click();
  URL.revokeObjectURL(a.href);
}

function openImport() {
  $('importModal')?.classList.add('on');
  if ($('importArea')) $('importArea').value = '';
}

function closeImport() {
  $('importModal')?.classList.remove('on');
}

function doImport() {
  try {
    const raw = v('importArea');
    if (!raw) return;
    const imp = JSON.parse(raw);
    if (!Array.isArray(imp)) throw new Error('يجب أن تكون مصفوفة JSON');
    const cur = loadList();
    const map = new Map(cur.map(x => [x.invoiceNumber, x]));
    imp.forEach(it => {
      if (it && it.invoiceNumber) {
        map.set(it.invoiceNumber, { ...it, id: it.id || crypto.randomUUID() });
      }
    });
    saveList(Array.from(map.values()));
    closeImport();
    renderHistory();
    alert('تم الاستيراد بنجاح');
  } catch (e) {
    alert('خطأ: ' + e.message);
  }
}

/* ════════ SELECTION & BATCH PRINT ════════ */
function toggleSel(id, cb) {
  if (cb.checked) selected.add(id);
  else selected.delete(id);
  updateBatchCount();
  renderHistory();
}

function updateBatchCount() {
  const el = $('batchCount');
  if (el) el.textContent = selected.size;

  const btn = $('batchPrintBtn');
  if (btn) {
    btn.style.background = selected.size > 0
      ? 'linear-gradient(135deg,#3b82f6,#2563eb)'
      : 'linear-gradient(135deg,#94a3b8,#64748b)';
  }
}

function toggleSelectAll() {
  const list = loadList();
  const q = (v('histSearch') || '').toLowerCase();
  const fil = list.filter(x => !q || [x.invoiceNumber, x.senderName, x.receiverName, x.senderCity, x.receiverCity].join(' ').toLowerCase().includes(q));
  const all = fil.every(x => selected.has(x.id));
  fil.forEach(x => all ? selected.delete(x.id) : selected.add(x.id));
  renderHistory();
  updateBatchCount();
}

/* ════════ PRINT ENGINE ════════ */
function getCSSText() {
  return Array.from(document.querySelectorAll('style')).map(s => s.textContent || '').join('\n');
}

function buildInvoiceHTML(item, sz) {
  const show = item.companyVisible !== false;
  const lg = item.logo || logoData || '';
  const lHTML = lg
    ? `<img src="${esc(lg)}" style="width:100%;height:100%;object-fit:contain;border-radius:inherit">`
    : `<i class="fa-solid fa-print"></i>`;

  if (sz === 'thermal80') {
    return `<article class="receipt r-t80" style="page-break-after:always;page-break-inside:avoid">
<div class="t80-accent"></div>
<div class="t80-head" style="${show ? '' : 'display:none'}">
  <div class="t80-logo">${lHTML}</div>
  <div class="t80-co">
    <div class="t80-co-name">${esc(item.companyName)}</div>
    <div class="t80-co-phone">${esc(item.companyPhone)}</div>
    <div class="t80-co-addr">${esc(item.companyAddress || '')}</div>
  </div>
</div>
<div class="t80-inv-box">
  <div class="t80-inv-label"><i class="fa-solid fa-barcode"></i> رقم الفاتورة</div>
  <div class="t80-inv-num">${esc(item.invoiceNumber)}</div>
</div>
<div class="t80-bc-wrap">
  ${makeBarcodeSVG(item.invoiceNumber, { height: 25, width: 1.4, fullWidth: true })}
  <div class="t80-bc-num">${esc(item.invoiceNumber)}</div>
</div>
<div class="t80-sec-head">
  <div class="t80-sec-dot" style="background:#1d4ed8"></div>
  <span class="t80-sec-label"><i class="fa-solid fa-truck-fast" style="color:#1d4ed8;font-size:7px"></i> الراسل</span>
</div>
<div class="t80-parties">
  <div class="t80-party snd">
    <div class="t80-party-tag"><i class="fa-solid fa-truck-fast" style="font-size:6px"></i> معلومات الراسل</div>
    <div class="t80-kv"><span class="t80-k">الاسم</span><span class="t80-v">${esc(item.senderName || '-')}</span></div>
    <div class="t80-kv"><span class="t80-k">الهاتف</span><span class="t80-v mono">${esc(item.senderPhone || '-')}</span></div>
    <div class="t80-kv"><span class="t80-k">المدينة</span><span class="t80-v">${esc(item.senderCity || '-')}</span></div>
    <div class="t80-kv"><span class="t80-k">العنوان</span><span class="t80-v">${esc(item.senderAddress || '-')}</span></div>
  </div>
</div>
<div class="t80-sec-head">
  <div class="t80-sec-dot" style="background:#059669"></div>
  <span class="t80-sec-label"><i class="fa-solid fa-user-check" style="color:#059669;font-size:7px"></i> المستلم</span>
</div>
<div class="t80-parties">
  <div class="t80-party rcv">
    <div class="t80-party-tag"><i class="fa-solid fa-user-check" style="font-size:6px"></i> معلومات المستلم</div>
    <div class="t80-kv"><span class="t80-k">الاسم</span><span class="t80-v">${esc(item.receiverName || '-')}</span></div>
    <div class="t80-kv"><span class="t80-k">الهاتف</span><span class="t80-v mono">${esc(item.receiverPhone || '-')}</span></div>
    <div class="t80-kv"><span class="t80-k">المدينة</span><span class="t80-v">${esc(item.receiverCity || '-')}</span></div>
    <div class="t80-kv"><span class="t80-k">العنوان</span><span class="t80-v">${esc(item.receiverAddress || '-')}</span></div>
  </div>
</div>
<div class="t80-details" style="margin-top:2.5mm">
  <div class="t80-details-grid">
    <div class="t80-dc"><span class="t80-dc-icon" style="color:#f59e0b"><i class="fa-solid fa-box"></i></span><div class="t80-dc-label">المحتوى</div><div class="t80-dc-val" style="font-family:'Cairo',sans-serif;font-size:6.5px;word-break:break-all">${esc(item.content || '-')}</div></div>
    <div class="t80-dc"><span class="t80-dc-icon" style="color:#8b5cf6"><i class="fa-solid fa-cubes"></i></span><div class="t80-dc-label">القطع</div><div class="t80-dc-val">${esc(item.pieces || '1')}</div></div>
    <div class="t80-dc"><span class="t80-dc-icon" style="color:#0ea5e9"><i class="fa-solid fa-weight-hanging"></i></span><div class="t80-dc-label">الوزن</div><div class="t80-dc-val">${esc(item.weight || '0')} كجم</div></div>
    <div class="t80-dc"><span class="t80-dc-icon" style="color:#10b981"><i class="fa-solid fa-calendar-days"></i></span><div class="t80-dc-label">التاريخ</div><div class="t80-dc-val" style="font-family:'Cairo',sans-serif;font-size:6px">${nd(item.shipDate)}</div></div>
  </div>
</div>
<div class="t80-delivery">
  <span class="t80-delivery-label"><i class="fa-solid fa-truck"></i> نوع التسليم</span>
  <span class="t80-delivery-val">${esc(item.deliveryType || '-')}</span>
</div>
<div class="t80-fins">
  <div class="t80-fin ship"><div class="t80-fin-icon"><i class="fa-solid fa-building-columns" style="color:#1d4ed8"></i></div><div class="t80-fin-label">شركة الشحن</div><div class="t80-fin-val" style="font-family:'Cairo',sans-serif;font-size:9px">${esc(item.shippingCompany || '-')}</div></div>
  <div class="t80-fin cod"><div class="t80-fin-icon"><i class="fa-solid fa-sack-dollar" style="color:#ea580c"></i></div><div class="t80-fin-label">التحصيل COD</div><div class="t80-fin-val">${money2(item.codAmount)}</div></div>
</div>
${item.notes ? `<div class="t80-notes"><div class="t80-notes-t"><i class="fa-solid fa-note-sticky"></i> ملاحظات</div><div class="t80-notes-text">${esc(item.notes)}</div></div>` : ''}
<div class="t80-footer">
  <div class="t80-footer-left" style="${show ? '' : 'visibility:hidden'}"><i class="fa-solid fa-print" style="font-size:6px"></i> <span class="t80-footer-brand">${esc(item.companyName)}</span></div>
  <div class="t80-footer-num">${esc(item.invoiceNumber)}</div>
  <div style="font-size:5.5px;color:#94a3b8;font-family:'IBM Plex Mono',monospace">80MM</div>
</div>
<div class="t80-bottom"></div>
</article>`;
  }

  if (sz === 'a4') {
    return `<article class="receipt r-a4" style="page-break-after:always;page-break-inside:avoid">
<div class="a4-band"></div>
<div class="a4-head">
  <div class="a4-brand">
    <div class="a4-logo" style="${show ? '' : 'display:none'}">${lHTML}</div>
    <div>
      <div class="a4-co-name" style="${show ? '' : 'display:none'}">${esc(item.companyName)}</div>
      <div class="a4-co-info" style="${show ? '' : 'display:none'}">${[item.companyAddress, item.companyPhone].filter(Boolean).join('<br>')}</div>
    </div>
  </div>
  <div class="a4-inv">
    <div class="a4-inv-label"><i class="fa-solid fa-barcode"></i> رقم الفاتورة</div>
    <div class="a4-inv-num">${esc(item.invoiceNumber)}</div>
    <div class="a4-bc-wrap">${makeBarcodeSVG(item.invoiceNumber, { height: 30, width: 2.0, fullWidth: true })}</div>
  </div>
</div>
<div class="a4-body">
  <div class="a4-st"><i class="fa-solid fa-address-card"></i> بيانات الراسل والمستلم</div>
  <div class="a4-parties">
    <div class="a4-party snd">
      <div class="a4-party-tag"><i class="fa-solid fa-truck-fast"></i> الراسل</div>
      <div class="a4-pname">${esc(item.senderName || '-')}</div>
      <div class="a4-pd">${esc(item.senderPhone || '-')}</div>
      <div class="a4-pd">${esc(item.senderCity || '-')}</div>
      <div class="a4-pd">${esc(item.senderAddress || '-')}</div>
    </div>
    <div class="a4-party rcv">
      <div class="a4-party-tag"><i class="fa-solid fa-user-check"></i> المستلم</div>
      <div class="a4-pname">${esc(item.receiverName || '-')}</div>
      <div class="a4-pd">${esc(item.receiverPhone || '-')}</div>
      <div class="a4-pd">${esc(item.receiverCity || '-')}</div>
      <div class="a4-pd">${esc(item.receiverAddress || '-')}</div>
    </div>
  </div>
  <div class="a4-st" style="margin-top:10px"><i class="fa-solid fa-box-open"></i> تفاصيل الشحنة</div>
  <div class="a4-shipment">
    <div class="a4-sc"><div class="a4-sc-label"><i class="fa-solid fa-box"></i> المحتوى</div><div class="a4-sc-val">${esc(item.content || '-')}</div></div>
    <div class="a4-sc"><div class="a4-sc-label"><i class="fa-solid fa-cubes"></i> القطع</div><div class="a4-sc-val">${esc(item.pieces || '1')}</div></div>
    <div class="a4-sc"><div class="a4-sc-label"><i class="fa-solid fa-weight-hanging"></i> الوزن</div><div class="a4-sc-val">${esc(item.weight || '0')} كجم</div></div>
    <div class="a4-sc"><div class="a4-sc-label"><i class="fa-solid fa-calendar-days"></i> التاريخ</div><div class="a4-sc-val">${nd(item.shipDate)}</div></div>
  </div>
  <div class="a4-delivery-badge"><span><i class="fa-solid fa-truck"></i> نوع التسليم</span><span>${esc(item.deliveryType || '-')}</span></div>
  <div class="a4-st"><i class="fa-solid fa-coins"></i> المبالغ</div>
  <div class="a4-fins">
    <div class="a4-fin ship"><div class="a4-fin-label"><i class="fa-solid fa-truck-moving"></i> شركة الشحن</div><div class="a4-fin-val" style="font-family:'Cairo',sans-serif;font-size:16px">${esc(item.shippingCompany || '-')}</div></div>
    <div class="a4-fin cod"><div class="a4-fin-label"><i class="fa-solid fa-sack-dollar"></i> التحصيل COD</div><div class="a4-fin-val">${money2(item.codAmount)}</div></div>
  </div>
  ${item.notes ? `<div class="a4-notes-box"><div class="a4-notes-label"><i class="fa-solid fa-note-sticky"></i> ملاحظات</div><div class="a4-notes-text">${esc(item.notes)}</div></div>` : ''}
</div>
<div class="a4-foot">
  <span class="a4-foot-brand" style="${show ? '' : 'visibility:hidden'}">${esc(item.companyName)}</span>
  <span class="a4-foot-mono">A4 · GRANADA PRINT SYSTEM</span>
</div>
</article>`;
  }

  if (sz === '200x80') {
    return `<article class="receipt r-200x80" style="page-break-after:always;page-break-inside:avoid">
<div class="r200-band"></div>
<div class="r200-body">
  <aside class="r200-side">
    <div class="r200-side-logo">
      <div class="r200-side-logo-box" style="${show ? '' : 'display:none'}">${lHTML}</div>
      <div style="${show ? '' : 'display:none'}"><div class="r200-co-name">${esc(item.companyName)}</div><div class="r200-co-phone">${esc(item.companyPhone)}</div></div>
    </div>
    <div class="r200-inv">
      <div class="r200-inv-label"><i class="fa-solid fa-receipt"></i> رقم الفاتورة</div>
      <div class="r200-inv-num">${esc(item.invoiceNumber)}</div>
    </div>
    <div class="r200-bc-area"><div class="r200-bc-inner">${makeBarcodeSVG(item.invoiceNumber, { height: 30, width: 1.0, fullWidth: true })}</div></div>
  </aside>
  <section class="r200-main">
    <div class="r200-top">
      <div class="r200-parties">
        <div class="r200-party snd">
          <div class="r200-party-tag"><i class="fa-solid fa-truck-fast" style="font-size:5.5px"></i> الراسل</div>
          <div class="r200-party-name">${esc(item.senderName || '-')}</div>
          <div class="r200-party-phone">${esc(item.senderPhone || '-')}</div>
          <div class="r200-party-phone" style="color:#334155;font-family:'Cairo',sans-serif">${esc(item.senderAddress || '-')}</div>
        </div>
        <div class="r200-party rcv">
          <div class="r200-party-tag"><i class="fa-solid fa-user-check" style="font-size:5.5px"></i> المستلم</div>
          <div class="r200-party-name">${esc(item.receiverName || '-')}</div>
          <div class="r200-party-phone">${esc(item.receiverPhone || '-')}</div>
          <div class="r200-party-phone" style="color:#334155;font-family:'Cairo',sans-serif">${esc(item.receiverAddress || '-')}</div>
        </div>
      </div>
      <div class="r200-meta">
        <div class="r200-badge type"><i class="fa-solid fa-truck" style="font-size:5.5px"></i><span>${esc(item.deliveryType || '-')}</span></div>
        <div class="r200-badge date"><i class="fa-solid fa-calendar" style="font-size:5.5px"></i><span>${nd(item.shipDate)}</span></div>
      </div>
    </div>
    <div class="r200-stats">
      <div class="r200-stat"><div class="r200-stat-icon" style="color:#f59e0b"><i class="fa-solid fa-box"></i></div><div class="r200-stat-label">المحتوى</div><div class="r200-stat-val" style="font-family:'Cairo',sans-serif">${esc(item.content || '-')}</div></div>
      <div class="r200-stat"><div class="r200-stat-icon" style="color:#8b5cf6"><i class="fa-solid fa-cubes"></i></div><div class="r200-stat-label">القطع</div><div class="r200-stat-val">${esc(item.pieces || '1')}</div></div>
      <div class="r200-stat"><div class="r200-stat-icon" style="color:#0ea5e9"><i class="fa-solid fa-weight-hanging"></i></div><div class="r200-stat-label">الوزن</div><div class="r200-stat-val">${esc(item.weight || '0')}</div></div>
      <div class="r200-stat"><div class="r200-stat-icon" style="color:#10b981"><i class="fa-solid fa-building-columns"></i></div><div class="r200-stat-label">شركة الشحن</div><div class="r200-stat-val" style="font-family:'Cairo',sans-serif">${esc(item.shippingCompany || '-')}</div></div>
    </div>
    <div class="r200-fins">
      <div class="r200-fin ship"><div class="r200-fin-icon"><i class="fa-solid fa-truck-moving" style="color:#1d4ed8"></i></div><div><div class="r200-fin-label">شركة الشحن</div><div class="r200-fin-val" style="font-family:'Cairo',sans-serif">${esc(item.shippingCompany || '-')}</div></div></div>
      <div class="r200-fin cod"><div class="r200-fin-icon"><i class="fa-solid fa-sack-dollar" style="color:#ea580c"></i></div><div><div class="r200-fin-label">التحصيل COD</div><div class="r200-fin-val">${money2(item.codAmount)}</div></div></div>
    </div>
  </section>
</div>
<div class="r200-footer">
  <div class="r200-footer-txt"><span style="width:7px;height:7px;background:#c9a227;border-radius:50%;display:inline-block"></span><span>${esc((item.companyName || 'GRANADA PRINT').toUpperCase())}</span><span style="width:7px;height:7px;background:#3b82f6;border-radius:50%;display:inline-block"></span><span>200×80 MM</span></div>
  <span class="r200-footer-num">${esc(item.invoiceNumber)}</span>
</div>
</article>`;
  }

  if (sz === '100x150') {
    return `<article class="receipt r-100x150" style="page-break-after:always;page-break-inside:avoid">
<div class="r150-hero">
  <div class="r150-hero-top">
    <div class="r150-brand">
      <div class="r150-logo-box" style="${show ? '' : 'display:none'}">${lHTML}</div>
      <div style="${show ? '' : 'display:none'}"><div class="r150-co-name">${esc(item.companyName)}</div><div class="r150-co-phone">${esc(item.companyPhone)}</div></div>
    </div>
    <div class="r150-inv-chip">
      <div class="r150-inv-label"><i class="fa-solid fa-barcode" style="font-size:5px"></i> رقم الفاتورة</div>
      <div class="r150-inv-val">${esc(item.invoiceNumber)}</div>
    </div>
  </div>
  <div class="r150-badges">
    <div class="r150-badge dlv"><i class="fa-solid fa-truck" style="font-size:5px"></i><span>${esc(item.deliveryType || '-')}</span></div>
    <div class="r150-badge dt"><i class="fa-solid fa-calendar" style="font-size:5px"></i><span>${nd(item.shipDate)}</span></div>
  </div>
</div>
<div class="r150-route">
  <div class="r150-rt-label"><i class="fa-solid fa-route" style="font-size:5.5px"></i> مسار الشحنة</div>
  <div class="r150-flow">
    <div class="r150-party snd">
      <div class="r150-ptag"><i class="fa-solid fa-truck-fast" style="font-size:5px"></i> الراسل</div>
      <div class="r150-pname">${esc(item.senderName || '-')}</div>
      <div class="r150-pphone"><i class="fa-solid fa-phone" style="font-size:5px"></i><span>${esc(item.senderPhone || '-')}</span></div>
      <div class="r150-pcity"><i class="fa-solid fa-city" style="font-size:5px"></i><span>${esc(item.senderCity || '-')}</span></div>
      <div class="r150-paddr">${esc(item.senderAddress || '-')}</div>
    </div>
    <div class="r150-party rcv">
      <div class="r150-ptag"><i class="fa-solid fa-user-check" style="font-size:5px"></i> المستلم</div>
      <div class="r150-pname">${esc(item.receiverName || '-')}</div>
      <div class="r150-pphone"><i class="fa-solid fa-phone" style="font-size:5px"></i><span>${esc(item.receiverPhone || '-')}</span></div>
      <div class="r150-pcity"><i class="fa-solid fa-city" style="font-size:5px"></i><span>${esc(item.receiverCity || '-')}</span></div>
      <div class="r150-paddr">${esc(item.receiverAddress || '-')}</div>
    </div>
  </div>
</div>
<div class="r150-details">
  <div class="r150-dgrid">
    <div class="r150-dc"><span class="r150-dc-icon" style="color:#f59e0b"><i class="fa-solid fa-box"></i></span><div class="r150-dc-label">المحتوى</div><div class="r150-dc-val" style="font-family:'Cairo',sans-serif;font-size:6px">${esc(item.content || '-')}</div></div>
    <div class="r150-dc"><span class="r150-dc-icon" style="color:#8b5cf6"><i class="fa-solid fa-cubes"></i></span><div class="r150-dc-label">القطع</div><div class="r150-dc-val">${esc(item.pieces || '1')}</div></div>
    <div class="r150-dc"><span class="r150-dc-icon" style="color:#0ea5e9"><i class="fa-solid fa-weight-hanging"></i></span><div class="r150-dc-label">الوزن</div><div class="r150-dc-val">${esc(item.weight || '0')}</div></div>
    <div class="r150-dc"><span class="r150-dc-icon" style="color:#10b981"><i class="fa-solid fa-calendar-days"></i></span><div class="r150-dc-label">التاريخ</div><div class="r150-dc-val" style="font-size:5.5px;font-family:'Cairo',sans-serif">${nd(item.shipDate)}</div></div>
  </div>
</div>
<div class="r150-fin">
  <div class="r150-fgrid">
    <div class="r150-fbox ship"><div class="r150-ficon"><i class="fa-solid fa-building-columns"></i></div><div><div class="r150-flabel">شركة الشحن</div><div class="r150-fval" style="font-family:'Cairo',sans-serif;font-size:7.5px">${esc(item.shippingCompany || '-')}</div></div></div>
    <div class="r150-fbox cod"><div class="r150-ficon"><i class="fa-solid fa-sack-dollar"></i></div><div><div class="r150-flabel">التحصيل COD</div><div class="r150-fval">${money2(item.codAmount)}</div></div></div>
  </div>
</div>
${item.notes ? `<div class="r150-notes-wrap"><div class="r150-notes"><span class="r150-notes-icon"><i class="fa-solid fa-note-sticky"></i></span><div class="r150-notes-text">${esc(item.notes)}</div></div></div>` : ''}
<div class="r150-foot">
  <div class="r150-bc">${makeBarcodeSVG(item.invoiceNumber, { height: 20, width: 1.3, fullWidth: true })}</div>
  <div class="r150-foot-brand"><span class="r150-foot-brand-name">${esc(item.companyName)}</span><span style="font-size:5.5px;color:#94a3b8"><i class="fa-solid fa-print" style="font-size:5px"></i> 100×150MM</span></div>
</div>
</article>`;
  }

  if (sz === '100x70') {
    return `<article class="receipt r-100x70" style="page-break-after:always;page-break-inside:avoid">
<div class="r70-band"></div>
<div class="r70-head" style="${show ? '' : 'display:none'}">
  <div class="r70-logo">${lHTML}</div>
  <div style="flex:1;min-width:0;padding:0 3px">
    <div class="r70-co-name">${esc(item.companyName)}</div>
    <div class="r70-co-phone">${esc(item.companyPhone)}</div>
  </div>
  <div>
    <div class="r70-num-label"><i class="fa-solid fa-barcode"></i> رقم الفاتورة</div>
    <div class="r70-num-val">${esc(item.invoiceNumber)}</div>
  </div>
</div>
<div class="r70-body">
  <div class="r70-bc-wrap">${makeBarcodeSVG(item.invoiceNumber, { height: 18, width: 0.9, fullWidth: false })}</div>
  <div class="r70-parties">
    <div class="r70-party">
      <div class="r70-ptag snd"><i class="fa-solid fa-truck-fast" style="font-size:5.5px"></i> الراسل</div>
      <div class="r70-kv"><span class="r70-k">الاسم</span><span class="r70-v">${esc(item.senderName || '-')}</span></div>
      <div class="r70-kv"><span class="r70-k">الهاتف</span><span class="r70-v">${esc(item.senderPhone || '-')}</span></div>
      <div class="r70-kv"><span class="r70-k">العنوان</span><span class="r70-v">${esc(item.senderAddress || '-')}</span></div>
    </div>
    <div class="r70-party">
      <div class="r70-ptag rcv"><i class="fa-solid fa-user-check" style="font-size:5.5px"></i> المستلم</div>
      <div class="r70-kv"><span class="r70-k">الاسم</span><span class="r70-v">${esc(item.receiverName || '-')}</span></div>
      <div class="r70-kv"><span class="r70-k">الهاتف</span><span class="r70-v">${esc(item.receiverPhone || '-')}</span></div>
      <div class="r70-kv"><span class="r70-k">العنوان</span><span class="r70-v">${esc(item.receiverAddress || '-')}</span></div>
    </div>
  </div>
  <div class="r70-amounts">
    <div class="r70-abox ship"><div class="r70-alabel"><i class="fa-solid fa-truck-moving"></i> شركة الشحن</div><div class="r70-aval" style="font-family:'Cairo',sans-serif;font-size:7px">${esc(item.shippingCompany || '-')}</div></div>
    <div class="r70-abox cod"><div class="r70-alabel"><i class="fa-solid fa-sack-dollar"></i> COD</div><div class="r70-aval">${money2(item.codAmount)}</div></div>
  </div>
</div>
<div class="r70-foot">
  <span class="r70-foot-brand">${esc(item.companyName)}</span>
  <span class="r70-foot-num">${esc(item.invoiceNumber)}</span>
  <span style="font-family:'IBM Plex Mono',monospace;font-size:5px;color:#94a3b8">100×70MM</span>
</div>
</article>`;
  }

  return '';
}

function buildPrintDoc(items, sz) {
  const cssText = getCSSText();
  const sizeMap = {
    thermal80: '@page{size:80mm auto;margin:0}',
    a4: '@page{size:210mm 297mm;margin:0}',
    '200x80': '@page{size:200mm 82mm;margin:0}',
    '100x150': '@page{size:100mm 150mm;margin:0}',
    '100x70': '@page{size:100mm 70mm;margin:0}'
  };

  const invHTML = items.map(item => buildInvoiceHTML(item, sz)).join('');

  return `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>Granada Print</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700;800;900&family=IBM+Plex+Mono:wght@400;500;600;700&family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
${sizeMap[sz] || '@page{margin:0}'}
*{box-sizing:border-box;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}
html,body{margin:0;padding:0;background:#fff;font-family:'Cairo','Tajawal',sans-serif;direction:rtl}
${cssText}
body{background:#fff!important;overflow:visible!important}
.receipt{box-shadow:none!important;margin:0 auto!important;transform:none!important;position:static!important;page-break-after:always;page-break-inside:avoid}
.receipt:last-child{page-break-after:auto}
.barcode-box{display:flex;justify-content:center;align-items:center;width:100%}
.barcode-box svg{display:block;margin:0 auto;max-width:100%}
.t80-bc-wrap,.a4-bc-wrap,.r150-bc,.r70-bc-wrap,.r200-bc-inner{display:flex;justify-content:center;align-items:center;width:100%}
</style>
</head>
<body>
<div id="print-root">${invHTML}</div>
<script>
(function(){
  function drawAllBC(){
    document.querySelectorAll('svg.bc-print').forEach(function(svg){
      try{
        var code = svg.getAttribute('data-code');
        if(code && window.JsBarcode){
          JsBarcode(svg, code, {
            format:'CODE128',
            displayValue:false,
            margin:0,
            lineColor:'#1a1a1a',
            background:'transparent',
            width:1.4,
            height:42
          });
        }
      }catch(e){}
    });
  }
  function doPrint(){
    drawAllBC();
    setTimeout(function(){ window.print(); }, 350);
  }
  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', doPrint);
  }else{
    doPrint();
  }
})();
</script>
</body>
</html>`;
}

/* ════════ PRINT HELPERS ════════ */
function triggerPrint(htmlContent) {
  let iframe = document.getElementById('pf');
  if (iframe) iframe.remove();

  iframe = document.createElement('iframe');
  iframe.id = 'pf';
  Object.assign(iframe.style, {
    position: 'fixed',
    top: '-9999px',
    left: '-9999px',
    width: '1px',
    height: '1px',
    border: 'none',
    opacity: '0'
  });

  document.body.appendChild(iframe);
  const doc = iframe.contentDocument || iframe.contentWindow.document;
  doc.open();
  doc.write(htmlContent);
  doc.close();
}

function printWithPluginOrBrowser(htmlContent) {
  if (window.cordova && window.cordova.plugins && window.cordova.plugins.printer) {
    try {
      window.cordova.plugins.printer.print(
        htmlContent,
        { name: 'GranadaPrint', duplex: false, landscape: false },
        function (res) { console.log('PRINT RESULT', res); },
        function (err) {
          console.error('PRINT ERROR', err);
          triggerPrint(htmlContent);
        }
      );
      return;
    } catch (e) {
      console.error('PRINTER EXCEPTION', e);
    }
  }

  triggerPrint(htmlContent);
}

function printInvoice() {
  const d = getData();
  const sz = $('printSize')?.value || 'thermal80';
  const htmlDoc = buildPrintDoc([d], sz);
  printWithPluginOrBrowser(htmlDoc);
}

function quickPrint(id) {
  const item = loadList().find(x => x.id === id);
  if (!item) return;
  const sz = item.printSize || $('printSize')?.value || 'thermal80';
  const htmlDoc = buildPrintDoc([item], sz);
  printWithPluginOrBrowser(htmlDoc);
}

async function batchPrint() {
  if (selected.size === 0) {
    alert('يرجى تحديد فاتورة واحدة على الأقل');
    return;
  }

  const invoices = loadList().filter(x => selected.has(x.id));
  if (!invoices.length) return;

  const sz = $('printSize')?.value || 'thermal80';
  const htmlDoc = buildPrintDoc(invoices, sz);
  printWithPluginOrBrowser(htmlDoc);
}

/* ════════ HISTORY ════════ */
function renderHistory() {
  const list = loadList();
  const q = (v('histSearch') || '').toLowerCase();
  const fil = list.filter(x =>
    !q || [x.invoiceNumber, x.senderName, x.receiverName, x.senderCity, x.receiverCity, x.content]
      .join(' ')
      .toLowerCase()
      .includes(q)
  );

  if ($('histCount')) $('histCount').textContent = fil.length;

  const wrap = $('histList');
  if (!wrap) return;

  if (!fil.length) {
    wrap.innerHTML = `<div class="hist-empty"><i class="fa-solid fa-clipboard-list fa-2x" style="opacity:.25"></i><span style="font-size:12px">لا توجد فواتير محفوظة</span></div>`;
    updateBatchCount();
    return;
  }

  wrap.innerHTML = fil.map(item => `
<div class="h-item ${item.id === editId ? 'active' : ''} ${selected.has(item.id) ? 'selected' : ''}" onclick="openInvoice('${item.id}')">
  <input type="checkbox" class="h-check" ${selected.has(item.id) ? 'checked' : ''} onclick="event.stopPropagation();toggleSel('${item.id}',this)">
  <div class="h-num"><i class="fa-solid fa-receipt"></i> ${item.invoiceNumber}</div>
  <div class="h-route" style="padding-right:22px">${item.senderName || '-'} ← ${item.receiverName || '-'}</div>
  <div style="font-size:10px;color:#94a3b8;padding-right:22px"><i class="fa-solid fa-location-dot"></i> ${item.senderCity || '-'} ← ${item.receiverCity || '-'}</div>
  <div class="h-meta"><span class="h-cod"><i class="fa-solid fa-sack-dollar" style="font-size:9px"></i> ${money(item.codAmount)}</span><span class="h-date">${niceDate(item.shipDate)}</span></div>
  <div class="h-actions">
    <button class="btn btn-dark" onclick="event.stopPropagation();openInvoice('${item.id}')"><i class="fa-solid fa-eye"></i> فتح</button>
    <button class="btn btn-blue" onclick="event.stopPropagation();quickPrint('${item.id}')"><i class="fa-solid fa-print"></i></button>
    <button class="btn btn-danger" onclick="event.stopPropagation();deleteInvoice('${item.id}')"><i class="fa-solid fa-trash"></i></button>
  </div>
</div>`).join('');

  updateBatchCount();
}

/* ════════ SCANNER ════════ */
async function reqCamPerm() {
  return new Promise((res, rej) => {
    if (!(window.cordova && cordova.plugins && cordova.plugins.permissions)) return res(true);

    const P = cordova.plugins.permissions;
    P.checkPermission(P.CAMERA, st => {
      if (st.hasPermission) return res(true);
      P.requestPermission(P.CAMERA, r => r.hasPermission ? res(true) : rej(new Error('تم رفض الكاميرا')), e => rej(e));
    }, e => rej(e));
  });
}

async function openScanner() {
  $('scanModal')?.classList.add('on');
  if ($('scStatus')) $('scStatus').textContent = 'جارِ تشغيل الكاميرا...';

  try {
    await reqCamPerm();
    if (scanStream) {
      scanStream.getTracks().forEach(t => t.stop());
      scanStream = null;
    }

    scanStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: camFace, width: { ideal: 1280 }, height: { ideal: 720 } }
    });

    const vid = $('scVideo');
    if (vid) {
      vid.srcObject = scanStream;
      vid.setAttribute('playsinline', 'true');
      await vid.play();
    }

    if ($('scStatus')) $('scStatus').textContent = 'وجّه الكاميرا نحو الباركود...';
    startScan();
  } catch (e) {
    if ($('scStatus')) $('scStatus').textContent = 'خطأ: ' + e.message;
  }
}

function stopScanner() {
  if (scanReader && scanReader.reset) {
    try { scanReader.reset(); } catch { }
  }
  if (scanStream) {
    scanStream.getTracks().forEach(t => t.stop());
    scanStream = null;
  }
}

function closeScanner() {
  $('scanModal')?.classList.remove('on');
  stopScanner();
}

function switchCam() {
  camFace = camFace === 'environment' ? 'user' : 'environment';
  closeScanner();
  setTimeout(openScanner, 300);
}

function startScan() {
  if (!window.ZXing || !window.ZXing.BrowserMultiFormatReader) return;

  scanReader = new ZXing.BrowserMultiFormatReader();
  scanReader.decodeFromConstraints({ audio: false, video: { facingMode: camFace } }, $('scVideo'), result => {
    if (!result) return;

    const code = (result.text || '').trim();
    if ($('scStatus')) {
      $('scStatus').innerHTML = `<span style="color:#059669;font-weight:800">✓ </span><span style="font-family:'IBM Plex Mono',monospace">${code}</span>`;
    }

    const found = loadList().find(x => (x.invoiceNumber || '').trim() === code);
    if (found) {
      setTimeout(() => {
        closeScanner();
        openInvoice(found.id);
      }, 700);
    } else {
      if ($('scStatus')) $('scStatus').innerHTML += `<br><span style="color:#dc2626">لا توجد فاتورة مطابقة</span>`;
    }
  });
}

/* ════════ PANEL TOGGLE ════════ */
function toggleRight() {
  $('rightPanel')?.classList.toggle('open');
  $('rOverlay')?.classList.toggle('on');
}
function closeRight() {
  $('rightPanel')?.classList.remove('open');
  $('rOverlay')?.classList.remove('on');
}

/* ════════ INIT ════════ */
function init() {
  const sd = $('shipDate');
  if (sd && !sd.value) sd.value = today();

  loadSavedLogo();

  const draftRaw = localStorage.getItem(DK);
  const saved = loadList();

  if (draftRaw) {
    try {
      const d = JSON.parse(draftRaw);
      if (!d.logo && logoData) d.logo = logoData;
      setData(d);
      editId = d.id || null;
    } catch {
      updateUI();
    }
  } else if (saved.length) {
    setData(saved[0]);
  } else {
    updateUI();
  }

  changeSize();
  renderHistory();
  applyBarcodePreviewCentering();

  if (window.cordova && cordova.plugins && cordova.plugins.permissions) {
    reqCamPerm().catch(() => { });
  }
}

if (window.cordova) document.addEventListener('deviceready', init, false);
else if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();

window.addEventListener('beforeunload', () => { try { stopScanner(); } catch { } });

document.addEventListener('keydown', e => {
  if (e.ctrlKey && e.key.toLowerCase() === 's') { e.preventDefault(); saveInvoice(); }
  if (e.ctrlKey && e.key.toLowerCase() === 'p') { e.preventDefault(); printInvoice(); }
  if (e.ctrlKey && e.key.toLowerCase() === 'b') { e.preventDefault(); openScanner(); }
  if (e.ctrlKey && e.key.toLowerCase() === 'n') { e.preventDefault(); newInvoice(); }
});

/* expose for inline handlers */
window.loadLogo = loadLogo;
window.loadSavedLogo = loadSavedLogo;
window.printInvoice = printInvoice;
window.batchPrint = batchPrint;
window.quickPrint = quickPrint;
window.newInvoice = newInvoice;
window.saveInvoice = saveInvoice;
window.openScanner = openScanner;
window.closeScanner = closeScanner;
window.switchCam = switchCam;
window.toggleRight = toggleRight;
window.closeRight = closeRight;
window.openInvoice = openInvoice;
window.deleteInvoice = deleteInvoice;
window.clearHistory = clearHistory;
window.exportInvoices = exportInvoices;
window.openImport = openImport;
window.closeImport = closeImport;
window.doImport = doImport;
window.genNum = genNum;