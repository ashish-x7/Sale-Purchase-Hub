/* ═══════════════════════════════════════════════════════════════════════
   AJIO • MYNTRA • FLIPKART — Sale & Purchase Processor
   Frontend Logic — Slim Upload, Full Width, Universal Search
   ═══════════════════════════════════════════════════════════════════════ */

const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwNUZf-kOwrze7bJOa89vHU1J2atjsyCTt3lQ7E15KJptTsBJ3g4tUH0MqnhSt-LHhoCw/exec';

// ─── State ──────────────────────────────────────────────────────────────
const state = {
    files: { sale_details:null, sale_summary:null, purchase_details:null, purchase_summary:null },
    saleData: [], purchaseData: [],
    currentTab: 'sale',
    saleOffset: 0, purchaseOffset: 0,
    pageSize: 25,
    saleTotal: 0, purchaseTotal: 0,

    // Sheet Viewer
    selectedSheet: '',
    viewerHeaders: [], viewerRows: [], viewerAllRows: [],
    viewerStartRow: 3, viewerPageSize: 25, viewerTotalRows: 0,
};

// ─── DOM ────────────────────────────────────────────────────────────────
const processBtn = document.getElementById('process-btn');
const exportBtn = document.getElementById('export-btn');
const loaderSection = document.getElementById('loader-section');
const resultSection = document.getElementById('result-section');
const loaderText = document.getElementById('loader-text');
const progressFill = document.getElementById('progress-fill');

// ─── File Upload (Slim Strip) ───────────────────────────────────────────
const fileKeys = ['sale_details', 'sale_summary', 'purchase_details', 'purchase_summary'];

fileKeys.forEach(key => {
    const card = document.getElementById(`card-${key.replace('_', '-')}`);
    const input = document.getElementById(`file-${key.replace('_', '-')}`);
    const nameEl = document.getElementById(`name-${key.replace('_', '-')}`);

    card.addEventListener('click', (e) => {
        if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'LABEL') input.click();
    });

    input.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            state.files[key] = file;
            nameEl.textContent = file.name;
            card.classList.add('file-added');
            checkAllFilesReady();
        }
    });

    card.addEventListener('dragover', (e) => { e.preventDefault(); card.style.borderColor = '#6366f1'; });
    card.addEventListener('dragleave', () => { card.style.borderColor = ''; });
    card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.style.borderColor = '';
        const file = e.dataTransfer.files[0];
        if (file && (file.name.endsWith('.xls') || file.name.endsWith('.xlsx'))) {
            state.files[key] = file;
            nameEl.textContent = file.name;
            card.classList.add('file-added');
            input.files = e.dataTransfer.files;
            checkAllFilesReady();
        } else {
            showToast('❌', 'Please drop a .xls or .xlsx file');
        }
    });
});

function checkAllFilesReady() {
    processBtn.disabled = !fileKeys.every(k => state.files[k] !== null);
}

// ─── Rows Per Page ──────────────────────────────────────────────────────
document.getElementById('rows-per-page').addEventListener('change', (e) => {
    state.pageSize = parseInt(e.target.value);
    // Reload data with new page size
    reloadLocalData();
});

async function reloadLocalData() {
    try {
        const res = await fetch(`/preview-more?offset=0&limit=${state.pageSize}`);
        const data = await res.json();
        if (!data.error) {
            state.saleData = data.sale_data || [];
            state.purchaseData = data.purchase_data || [];
            state.saleOffset = state.saleData.length;
            state.purchaseOffset = state.purchaseData.length;
            renderTable(state.currentTab);
        }
    } catch(e) { console.error(e); }
}

// ─── Process Files ──────────────────────────────────────────────────────
processBtn.addEventListener('click', async () => {
    if (processBtn.disabled) return;
    loaderSection.classList.remove('hidden');
    resultSection.classList.add('hidden');
    processBtn.disabled = true;

    const steps = [
        {text:'Reading Sale Details...',pct:15}, {text:'Reading Sale Summary...',pct:30},
        {text:'Reading Purchase Details...',pct:45}, {text:'Reading Purchase Summary...',pct:60},
        {text:'Running VLOOKUP merge...',pct:75}, {text:'Calculating columns...',pct:88},
        {text:'Checking duplicates...',pct:95},
    ];
    let i = 0;
    const iv = setInterval(() => {
        if (i < steps.length) { loaderText.textContent = steps[i].text; progressFill.style.width = steps[i].pct+'%'; i++; }
    }, 600);

    const fd = new FormData();
    fileKeys.forEach(k => fd.append(k, state.files[k]));

    try {
        const res = await fetch('/upload', { method:'POST', body:fd });
        clearInterval(iv);
        const raw = await res.text();
        let data;
        try {
            data = JSON.parse(raw);
        } catch(parseErr) {
            throw new Error(raw ? `Server returned invalid response: ${raw.slice(0, 200)}` : 'Server returned an empty response');
        }
        if (!res.ok || data.error) throw new Error(data.error || 'Upload failed');

        progressFill.style.width = '100%';
        loaderText.textContent = 'Done! ✅';

        state.saleData = data.sale_preview || [];
        state.purchaseData = data.purchase_preview || [];
        state.saleOffset = state.saleData.length;
        state.purchaseOffset = state.purchaseData.length;
        state.saleTotal = data.sale_count;
        state.purchaseTotal = data.purchase_count;

        document.getElementById('sale-count').textContent = formatNumber(data.sale_count);
        document.getElementById('purchase-count').textContent = formatNumber(data.purchase_count);
        document.getElementById('total-count').textContent = formatNumber(data.sale_count + data.purchase_count);
        if (data.financial_year) document.getElementById('fy-badge').textContent = `FY ${data.financial_year}`;

        if (data.sale_duplicates > 0 || data.purchase_duplicates > 0)
            showToast('🔍', `Duplicates skipped: Sale=${data.sale_duplicates}, Purchase=${data.purchase_duplicates}`);

        setTimeout(() => {
            loaderSection.classList.add('hidden');
            resultSection.classList.remove('hidden');
            renderTable('sale');
            updateHeaderStatus(data.sale_count, data.purchase_count);
            showToast('✅', `Processed! Sale: ${data.sale_count}, Purchase: ${data.purchase_count}`);
        }, 400);
    } catch(err) {
        clearInterval(iv);
        loaderSection.classList.add('hidden');
        processBtn.disabled = false;
        showToast('❌', err.message);
    }
});

// ─── Tabs ───────────────────────────────────────────────────────────────
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.currentTab = btn.dataset.tab;
        renderTable(btn.dataset.tab);
    });
});

// ─── Table Columns ──────────────────────────────────────────────────────
const SALE_COLUMNS = [
    {key:'no',label:'No.'},{key:'invoice_no',label:'Invoice No'},{key:'type',label:'TYPE'},
    {key:'invoice_date',label:'Invoice Date'},{key:'warehouse_name',label:'Warehouse'},
    {key:'warehouse_code',label:'WH Code'},{key:'gst_no',label:'GST No'},
    {key:'order_id',label:'Order ID'},{key:'item_asin',label:'Item Asin'},
    {key:'item_sku',label:'Item SKU'},{key:'item_name',label:'Item Name'},
    {key:'hsn_number',label:'HSN'},{key:'quantity',label:'Qty'},
    {key:'item_cost',label:'Item Cost'},{key:'gross',label:'Gross'},
    {key:'igst',label:'IGST%'},{key:'cgst',label:'CGST%'},{key:'sgst',label:'SGST%'},
    {key:'igst_amt',label:'IGST Amt'},{key:'cgst_amt',label:'CGST Amt'},
    {key:'sgst_amt',label:'SGST Amt'},{key:'invoice',label:'Invoice'},
    {key:'reason',label:'Reason'},{key:'zoho_status',label:'Zoho'},
    {key:'invoice_id',label:'Invoice ID'},{key:'state_code',label:'State Code'},
    {key:'sale_unique_id',label:'SALE UNIQUE ID'},
    {key:'calc_qty',label:'Qty*'},{key:'calc_cost',label:'Cost*'},
    {key:'calc_gross',label:'Gross*'},{key:'calc_igst',label:'IGST*'},
    {key:'calc_cgst',label:'CGST*'},{key:'calc_sgst',label:'SGST*'},
    {key:'calc_igst_amt',label:'IGST A*'},{key:'calc_cgst_amt',label:'CGST A*'},
    {key:'calc_sgst_amt',label:'SGST A*'},{key:'calc_invoice',label:'Invoice*'},
    {key:'state_code_short',label:'STATE'},
];

const PURCHASE_COLUMNS = [
    {key:'no',label:'No.'},{key:'invoice_no',label:'Invoice No'},
    {key:'warehouse_name',label:'Warehouse'},{key:'warehouse_code',label:'WH Code'},
    {key:'gst_no',label:'GST No'},{key:'order_id',label:'Order ID'},
    {key:'item_asin',label:'Item Asin'},{key:'item_sku',label:'Item SKU'},
    {key:'item_name',label:'Item Name'},{key:'hsn_number',label:'HSN'},
    {key:'quantity',label:'Qty'},{key:'item_cost',label:'Item Cost'},
    {key:'gross',label:'Gross'},{key:'igst',label:'IGST%'},
    {key:'cgst',label:'CGST%'},{key:'sgst',label:'SGST%'},
    {key:'igst_amt',label:'IGST Amt'},{key:'cgst_amt',label:'CGST Amt'},
    {key:'sgst_amt',label:'SGST Amt'},{key:'invoice',label:'Invoice'},
    {key:'purchase_unique_id',label:'PURCHASE UNIQUE ID'},
    {key:'calc_qty',label:'Qty*'},{key:'calc_cost',label:'Cost*'},
    {key:'calc_gross',label:'Gross*'},{key:'calc_igst',label:'IGST*'},
    {key:'calc_cgst',label:'CGST*'},{key:'calc_sgst',label:'SGST*'},
    {key:'calc_igst_amt',label:'IGST A*'},{key:'calc_cgst_amt',label:'CGST A*'},
    {key:'calc_sgst_amt',label:'SGST A*'},{key:'calc_total_amt',label:'Total Amt'},
];

function renderTable(tab) {
    const thead = document.getElementById('table-head');
    const tbody = document.getElementById('table-body');
    const cols = tab === 'sale' ? SALE_COLUMNS : PURCHASE_COLUMNS;
    const data = tab === 'sale' ? state.saleData : state.purchaseData;
    const thClass = tab === 'sale' ? 'sale-th' : 'purchase-th';
    const total = tab === 'sale' ? state.saleTotal : state.purchaseTotal;

    thead.innerHTML = '<tr>' + cols.map(c => `<th class="${thClass}">${c.label}</th>`).join('') + '</tr>';
    tbody.innerHTML = data.map(row =>
        '<tr>' + cols.map(c => {
            let v = row[c.key]; if (v==null) v='';
            if (typeof v==='number' && !Number.isInteger(v)) v=v.toFixed(2);
            return `<td title="${esc(String(v))}">${esc(String(v))}</td>`;
        }).join('') + '</tr>'
    ).join('');

    document.getElementById('showing-info').textContent = `Showing ${data.length} of ${formatNumber(total)} rows`;
    const btn = document.getElementById('load-more-btn');
    btn.classList.toggle('hidden', data.length >= total);
}

// ─── Load More ──────────────────────────────────────────────────────────
document.getElementById('load-more-btn').addEventListener('click', async () => {
    const btn = document.getElementById('load-more-btn');
    btn.textContent = 'Loading...'; btn.disabled = true;
    const offset = state.currentTab === 'sale' ? state.saleOffset : state.purchaseOffset;

    try {
        const res = await fetch(`/preview-more?offset=${offset}&limit=${state.pageSize}`);
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        const nd = state.currentTab === 'sale' ? data.sale_data : data.purchase_data;
        if (state.currentTab === 'sale') { state.saleData = state.saleData.concat(nd); state.saleOffset = state.saleData.length; }
        else { state.purchaseData = state.purchaseData.concat(nd); state.purchaseOffset = state.purchaseData.length; }
        renderTable(state.currentTab);
    } catch(e) { showToast('❌', e.message); }
    btn.textContent = 'Load More Rows'; btn.disabled = false;
});

// ─── Export ─────────────────────────────────────────────────────────────
exportBtn.addEventListener('click', async () => {
    exportBtn.disabled = true;
    exportBtn.querySelector('span').textContent = 'Generating...';
    try {
        const res = await fetch('/export');
        if (!res.ok) { const e = await res.json(); throw new Error(e.error||'Export failed'); }
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'AJIO_MYNTRA_FLIPKART_Sale_Purchase.xlsx';
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast('✅', 'Excel downloaded!');
    } catch(e) { showToast('❌', e.message); }
    exportBtn.disabled = false;
    exportBtn.querySelector('span').textContent = 'Export .xlsx';
});

// ─── Clear Local Data ───────────────────────────────────────────────────
function clearAllData() { document.getElementById('modal-overlay').classList.remove('hidden'); }
function closeModal() { document.getElementById('modal-overlay').classList.add('hidden'); }

async function confirmClearData() {
    closeModal();
    showGlobalLoader('Clearing Data', 'Removing all processed records...');
    try {
        const res = await fetch('/clear', {method:'POST'});
        const data = await res.json();
        hideGlobalLoader();
        if (data.success) {
            state.saleData=[]; state.purchaseData=[];
            state.saleTotal=0; state.purchaseTotal=0;
            state.saleOffset=0; state.purchaseOffset=0;
            document.getElementById('sale-count').textContent = '0';
            document.getElementById('purchase-count').textContent = '0';
            document.getElementById('total-count').textContent = '0';
            document.getElementById('table-head').innerHTML = '';
            document.getElementById('table-body').innerHTML = '';
            document.getElementById('showing-info').textContent = 'Showing 0 rows';
            document.getElementById('load-more-btn').classList.add('hidden');
            resultSection.classList.add('hidden');
            document.getElementById('header-status').classList.add('hidden');
            showToast('✅', 'All data cleared!');
        } else { showToast('❌', data.error||'Clear failed'); }
    } catch(e) { hideGlobalLoader(); showToast('❌', e.message); }
}

// ─── Header Status ──────────────────────────────────────────────────────
function updateHeaderStatus(s, p) {
    const el = document.getElementById('header-status');
    const t = document.getElementById('header-status-text');
    if (s>0||p>0) { t.textContent=`${formatNumber(s)} sale • ${formatNumber(p)} purchase loaded`; el.classList.remove('hidden'); }
    else el.classList.add('hidden');
}

// ─── Auto Load ──────────────────────────────────────────────────────────
async function loadExistingData() {
    try {
        const sRes = await fetch('/status');
        const st = await sRes.json();
        if (st.has_data) {
            state.saleTotal = st.sale_count; state.purchaseTotal = st.purchase_count;
            document.getElementById('sale-count').textContent = formatNumber(st.sale_count);
            document.getElementById('purchase-count').textContent = formatNumber(st.purchase_count);
            document.getElementById('total-count').textContent = formatNumber(st.sale_count + st.purchase_count);
            if (st.financial_year) document.getElementById('fy-badge').textContent = `FY ${st.financial_year}`;

            const pRes = await fetch(`/preview-more?offset=0&limit=${state.pageSize}`);
            const pd = await pRes.json();
            if (!pd.error) {
                state.saleData = pd.sale_data||[]; state.purchaseData = pd.purchase_data||[];
                state.saleOffset = state.saleData.length; state.purchaseOffset = state.purchaseData.length;
                resultSection.classList.remove('hidden');
                renderTable('sale');
                updateHeaderStatus(st.sale_count, st.purchase_count);
            }
        }
    } catch(e) { console.log('No existing data:', e.message); }
}
loadExistingData();

// ═══════════════════════════════════════════════════════════════════════
// GOOGLE SHEETS — Hardcoded URL
// ═══════════════════════════════════════════════════════════════════════

function showGlobalLoader(t, s) {
    document.getElementById('global-loader-title').textContent = t;
    document.getElementById('global-loader-status').textContent = s;
    document.getElementById('global-loader').classList.remove('hidden');
}
function hideGlobalLoader() { document.getElementById('global-loader').classList.add('hidden'); }

// ─── Load Sheet Names ───────────────────────────────────────────────────
async function refreshSheetDropdown() {
    try {
        const res = await fetch(`${WEB_APP_URL}?action=getSheetNames`);
        const r = await res.json();
        if (r.success) {
            const sel = document.getElementById('sheet-select');
            const syncSel = document.getElementById('sync-sheet-select');
            const curV = sel?sel.value:'';
            const curSV = syncSel?syncSel.value:'';

            if(sel) sel.innerHTML = '<option value="">-- Choose sheet --</option>';
            if(syncSel) syncSel.innerHTML = '<option value="">-- Choose sheet --</option>';

            r.sheets.forEach(sh => {
                const txt = `${sh.name} (${sh.rows} rows)`;
                if(sel) { const o=document.createElement('option'); o.value=sh.name; o.textContent=txt; sel.appendChild(o); }
                if(syncSel) { const o=document.createElement('option'); o.value=sh.name; o.textContent=txt; syncSel.appendChild(o); }
            });

            // Render Google Sheets style tabs at the bottom
            const tabsBar = document.getElementById('sheet-tabs-bar');
            if (tabsBar) {
                tabsBar.innerHTML = '';
                r.sheets.forEach(sh => {
                    const tabEl = document.createElement('div');
                    tabEl.className = 'sheet-tab';
                    tabEl.dataset.name = sh.name;
                    tabEl.textContent = sh.name;
                    tabEl.addEventListener('click', () => {
                        setActiveSheetTab(sh.name);
                    });
                    tabsBar.appendChild(tabEl);
                });
            }

            const def = `AJIO & MYNTRA SALE-PURCHASE ${r.fy}`;
            const hasDef = r.sheets.some(s=>s.name===def);

            if(sel) {
                if(curV && r.sheets.some(s=>s.name===curV)) {
                    setActiveSheetTab(curV);
                } else if(hasDef) {
                    setActiveSheetTab(def);
                } else if(r.sheets.length > 0) {
                    setActiveSheetTab(r.sheets[0].name);
                } else {
                    clearViewerTable();
                }
            }
            if(syncSel) {
                if(curSV && r.sheets.some(s=>s.name===curSV)) syncSel.value=curSV;
                else if(hasDef) syncSel.value=def;
            }
        } else { showToast('❌', 'Sheets Error: '+r.error); }
    } catch(e) { showToast('❌', 'Cannot connect to Google Sheets'); console.error(e); }
}
refreshSheetDropdown();

function setActiveSheetTab(sheetName) {
    state.selectedSheet = sheetName;
    const sel = document.getElementById('sheet-select');
    if (sel) sel.value = sheetName;
    
    // Update active class on tab elements
    document.querySelectorAll('.sheet-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.name === sheetName);
    });
    
    document.getElementById('universal-search').value = ''; // Clear search on sheet change
    if (sheetName) loadSheetData(3);
    else clearViewerTable();
}

function onSheetSelectionChange() {
    const sel = document.getElementById('sheet-select');
    setActiveSheetTab(sel.value);
}

function onSheetRowsChange() {
    const v = parseInt(document.getElementById('sheet-rows-per-page').value);
    state.viewerPageSize = v;
    if (state.selectedSheet) loadSheetData(3);
}

// ─── Load Sheet Data ────────────────────────────────────────────────────
async function loadSheetData(startRow) {
    if (!state.selectedSheet) return;
    state.viewerStartRow = startRow;
    showGlobalLoader('Loading Sheet', `Fetching from ${state.selectedSheet}...`);

    try {
        const res = await fetch(`${WEB_APP_URL}?action=getSheetData&sheetName=${encodeURIComponent(state.selectedSheet)}&startRow=${startRow}&numRows=${state.viewerPageSize}`);
        const r = await res.json();
        hideGlobalLoader();

        if (r.success) {
            state.viewerHeaders = r.headers||[];
            state.viewerRows = r.data||[];
            state.viewerAllRows = r.data||[];
            state.viewerTotalRows = r.totalRows||0;
            renderViewerTable();
        } else {
            showToast('❌', 'Data Error: '+r.error);
            clearViewerTable();
        }
    } catch(e) { hideGlobalLoader(); showToast('❌', 'Network error'); clearViewerTable(); }
}

// ─── Render Viewer Table ────────────────────────────────────────────────
function renderViewerTable() {
    const thead = document.getElementById('viewer-thead');
    const tbody = document.getElementById('viewer-tbody');
    const pInfo = document.getElementById('page-info');
    const tInfo = document.getElementById('viewer-total-info');

    if (state.viewerHeaders.length === 0) {
        thead.innerHTML=''; tbody.innerHTML='<tr><td style="text-align:center;color:var(--text-muted);">Sheet is empty.</td></tr>';
        document.getElementById('prev-btn').disabled = true;
        document.getElementById('next-btn').disabled = true;
        return;
    }

    const hdr = state.viewerHeaders[1]||state.viewerHeaders[0]||[];
    thead.innerHTML = `<tr>${hdr.map(h=>`<th>${h}</th>`).join('')}</tr>`;

    if (state.viewerRows.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${hdr.length}" style="text-align:center;color:var(--text-muted);">No records.</td></tr>`;
    } else {
        tbody.innerHTML = state.viewerRows.map(row =>
            `<tr>${row.map(v => {
                let d = (v==null)?'':v;
                if(typeof d==='number' && !Number.isInteger(d)) d=d.toFixed(2);
                return `<td title="${esc(String(d))}">${esc(String(d))}</td>`;
            }).join('')}</tr>`
        ).join('');
    }

    const endRow = Math.min(state.viewerStartRow + state.viewerPageSize - 1, state.viewerTotalRows + 2);
    pInfo.textContent = `${state.viewerStartRow} - ${endRow}`;
    tInfo.textContent = `Total: ${formatNumber(state.viewerTotalRows)} rows`;

    document.getElementById('prev-btn').disabled = (state.viewerStartRow <= 3);
    document.getElementById('next-btn').disabled = (endRow >= state.viewerTotalRows + 2);
}

function changePage(dir) {
    let next = state.viewerStartRow + (dir * state.viewerPageSize);
    if (next < 3) next = 3;
    document.getElementById('universal-search').value = ''; // Clear search on page change
    loadSheetData(next);
}

function clearViewerTable() {
    document.getElementById('viewer-thead').innerHTML='';
    document.getElementById('viewer-tbody').innerHTML='';
    document.getElementById('page-info').textContent='0 - 0';
    document.getElementById('viewer-total-info').textContent='Select a sheet.';
    document.getElementById('prev-btn').disabled=true;
    document.getElementById('next-btn').disabled=true;
}

// ─── Universal Search ───────────────────────────────────────────────────
function handleUniversalSearch() {
    const q = document.getElementById('universal-search').value.trim().toLowerCase();
    const tbody = document.getElementById('viewer-tbody');
    const rows = tbody.querySelectorAll('tr');

    if (!q) {
        rows.forEach(r => { r.classList.remove('search-hidden','search-match'); });
        return;
    }

    let matchCount = 0;
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        let found = false;
        cells.forEach(cell => {
            if (cell.textContent.toLowerCase().includes(q)) found = true;
        });
        if (found) {
            row.classList.remove('search-hidden');
            row.classList.add('search-match');
            matchCount++;
        } else {
            row.classList.add('search-hidden');
            row.classList.remove('search-match');
        }
    });

    document.getElementById('viewer-total-info').textContent = `Search: ${matchCount} match${matchCount!==1?'es':''} found`;
}

// ─── Clear Sheet ────────────────────────────────────────────────────────
function triggerClearSheetWarning() {
    if (!state.selectedSheet) { showToast('⚠️', 'Select a sheet first!'); return; }
    document.getElementById('modal-sheet-name').textContent = state.selectedSheet;
    document.getElementById('clear-sheet-modal').classList.remove('hidden');
}
function closeClearSheetModal() { document.getElementById('clear-sheet-modal').classList.add('hidden'); }

async function confirmClearSheet() {
    closeClearSheetModal();
    showGlobalLoader('Clearing Sheet', `Wiping ${state.selectedSheet}...`);
    try {
        const res = await fetch(`${WEB_APP_URL}?action=clearSheetData&sheetName=${encodeURIComponent(state.selectedSheet)}`);
        const r = await res.json();
        hideGlobalLoader();
        if (r.success) { showToast('✅', `Cleared: ${state.selectedSheet}`); refreshSheetDropdown(); }
        else showToast('❌', 'Error: '+r.error);
    } catch(e) { hideGlobalLoader(); showToast('❌', 'Network error'); }
}

// ─── Sync Modal ─────────────────────────────────────────────────────────
const syncBtn = document.getElementById('sync-sheet-btn');
if(syncBtn) syncBtn.addEventListener('click', openSyncModal);

function openSyncModal() { document.getElementById('sync-modal-overlay').classList.remove('hidden'); refreshSyncDropdown(); }
function closeSyncModal() { document.getElementById('sync-modal-overlay').classList.add('hidden'); }

async function refreshSyncDropdown() {
    try {
        const res = await fetch(`${WEB_APP_URL}?action=getSheetNames`);
        const r = await res.json();
        if (r.success) {
            const sel = document.getElementById('sync-sheet-select');
            const cur = sel?sel.value:'';
            if(sel) sel.innerHTML='<option value="">-- Choose --</option>';
            r.sheets.forEach(sh => { if(sel){ const o=document.createElement('option'); o.value=sh.name; o.textContent=`${sh.name} (${sh.rows})`; sel.appendChild(o); }});
            const def = `AJIO & MYNTRA SALE-PURCHASE ${r.fy}`;
            if(sel) { if(cur&&r.sheets.some(s=>s.name===cur)) sel.value=cur; else if(r.sheets.some(s=>s.name===def)) sel.value=def; }
        }
    } catch(e) { console.error(e); }
}

async function startSyncToGoogleSheets() {
    const sheetName = document.getElementById('sync-sheet-select')?.value||'';
    const mode = document.querySelector('input[name="sync-mode"]:checked')?.value||'append';
    if (!sheetName) { showToast('⚠️', 'Select a target sheet'); return; }

    closeSyncModal();
    
    // Set up progress bar
    const progBar = document.getElementById('global-progress-bar');
    const progFill = document.getElementById('global-progress-fill');
    if (progBar) progBar.style.display = 'block';
    if (progFill) progFill.style.width = '0%';
    
    showGlobalLoader('Preparing Sync', 'Fetching and deduplicating data...');
    
    try {
        // Fetch deduplicated and formatted rows from local backend
        const syncRes = await fetch('/get-sync-data', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ webAppUrl: WEB_APP_URL, sheetName, mode })
        });
        const syncData = await syncRes.json();
        if (!syncRes.ok || syncData.error) throw new Error(syncData.error || 'Deduplication failed');
        
        const saleRows = syncData.saleRows || [];
        const purchaseRows = syncData.purchaseRows || [];
        
        const maxRows = Math.max(saleRows.length, purchaseRows.length);
        if (maxRows === 0) {
            hideGlobalLoader();
            if (progBar) progBar.style.display = 'none';
            // Show duplicate warning/info
            document.getElementById('success-sale-sent').textContent = '0';
            document.getElementById('success-purchase-sent').textContent = '0';
            document.getElementById('success-sale-skipped').textContent = formatNumber(syncData.skippedSale);
            document.getElementById('success-purchase-skipped').textContent = formatNumber(syncData.skippedPurchase);
            document.getElementById('success-skipped-container').style.display = 'block';
            document.getElementById('sync-success-modal').classList.remove('hidden');
            showToast('ℹ️', 'No new records to sync (all duplicates skipped)');
            refreshSheetDropdown();
            return;
        }
        
        // Chunk sizes and total calculations
        const chunkSize = 2000;
        const totalBatches = Math.ceil(maxRows / chunkSize);
        
        for (let i = 1; i <= totalBatches; i++) {
            const startIdx = (i - 1) * chunkSize;
            const endIdx = i * chunkSize;
            
            const saleChunk = saleRows.slice(startIdx, endIdx);
            const purchaseChunk = purchaseRows.slice(startIdx, endIdx);
            
            const percent = Math.round((i / totalBatches) * 100);
            showGlobalLoader('Syncing to Google Sheets', `Uploading Batch ${i} of ${totalBatches}... (${percent}%)`);
            if (progFill) progFill.style.width = percent + '%';
            
            // Check action type: Overwrite + First batch should clear sheet.
            const actionType = (mode === 'overwrite' && i === 1) ? 'writeData' : 'appendBatch';
            const payload = {
                action: actionType,
                sheetName: sheetName,
                saleRows: saleChunk,
                purchaseRows: purchaseChunk,
                isAppend: !(mode === 'overwrite' && i === 1),
                batchIndex: i,
                totalBatches: totalBatches
            };
            
            // Call Google Apps Script directly
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                body: JSON.stringify(payload)
            });
            const resJSON = await response.json();
            if (!resJSON.success) throw new Error(resJSON.error || `Batch ${i} failed`);
        }
        
        // Done! Hide loader and show success modal
        hideGlobalLoader();
        if (progBar) progBar.style.display = 'none';
        
        document.getElementById('success-sale-sent').textContent = formatNumber(syncData.toSyncSale);
        document.getElementById('success-purchase-sent').textContent = formatNumber(syncData.toSyncPurchase);
        document.getElementById('success-sale-skipped').textContent = formatNumber(syncData.skippedSale);
        document.getElementById('success-purchase-skipped').textContent = formatNumber(syncData.skippedPurchase);
        document.getElementById('success-skipped-container').style.display = (mode === 'append') ? 'block' : 'none';
        
        document.getElementById('sync-success-modal').classList.remove('hidden');
        showToast('✅', 'Sync completed successfully!');
        refreshSheetDropdown();
        
    } catch(e) {
        hideGlobalLoader();
        if (progBar) progBar.style.display = 'none';
        showToast('❌', 'Sync Error: ' + e.message);
        document.getElementById('sync-modal-overlay').classList.remove('hidden');
    }
}

function closeSuccessModal() {
    document.getElementById('sync-success-modal').classList.add('hidden');
}

// ─── Helpers ────────────────────────────────────────────────────────────
function formatNumber(n) { return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function esc(s) { const d=document.createElement('div'); d.textContent=s; return d.innerHTML; }

function showToast(icon, msg) {
    const t = document.getElementById('toast');
    document.getElementById('toast-icon').textContent = icon;
    document.getElementById('toast-message').textContent = msg;
    t.classList.remove('hidden');
    t.classList.toggle('error', icon==='❌');
    clearTimeout(window._tt);
    window._tt = setTimeout(() => t.classList.add('hidden'), 5000);
}
