// ==========================================
// KONFIGURASI
// ==========================================
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz-Hss2p0kFI8tns7yHO0UGKHMmU6Hvp27rI6quANmdJ4ZuBDwPP-tGqS_n1jDUzENk/exec'; // GANTI DENGAN URL ANDA
const SESSION_DURATION = 24 * 60 * 60 * 1000;

// ==========================================
// STATE GLOBAL
// ==========================================
let user = null;
let authToken = null;
let configData = [];
let allTransactions = [];
let targetsData = {};
let parsedExcelData = null;

// ==========================================
// UTILITIES
// ==========================================
function parseRupiah(str) {
    if (!str) return 0;
    if (typeof str === 'number') return str;
    let cleaned = String(str).replace(/\./g, '').replace(/,/g, '.');
    return parseFloat(cleaned) || 0;
}

function formatRupiah(num) {
    return Math.round(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function securePost(payload) {
    if (authToken) payload.token = authToken;
    return fetch(SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify(payload)
    }).then(r => r.json());
}

function showToast(msg) {
    const t = document.getElementById('toast');
    if (t) {
        t.innerText = msg;
        t.classList.add('show');
        setTimeout(() => t.classList.remove('show'), 3000);
    }
}

// ==========================================
// AUTHENTICATION
// ==========================================
function doLogin() {
    const pass = document.getElementById('login-pass').value.trim();
    const msg = document.getElementById('login-msg');
    if (!pass) return msg.innerText = "Masukkan password!";
    msg.innerText = "Memproses...";

    securePost({ action: 'login', password: pass })
    .then(res => {
        if (res.status === 'success') {
            user = res.user;
            authToken = res.token;
            localStorage.setItem('app_session', JSON.stringify({
                token: authToken, user: user, timestamp: Date.now()
            }));
            initApp();
        } else {
            msg.innerText = res.message;
        }
    })
    .catch(() => msg.innerText = "Koneksi Error");
}

function doLogout() {
    localStorage.removeItem('app_session');
    user = null;
    authToken = null;
    location.reload();
}

function checkSession() {
    const sessionStr = localStorage.getItem('app_session');
    const loginMsg = document.getElementById('login-msg');

    if (sessionStr) {
        try {
            const session = JSON.parse(sessionStr);
            const elapsed = Date.now() - session.timestamp;

            if (elapsed > SESSION_DURATION) {
                localStorage.removeItem('app_session');
                if (loginMsg) loginMsg.innerText = "Sesi login telah habis. Silakan login ulang.";
            } else {
                authToken = session.token;
                user = session.user;
                initApp();
            }
        } catch (e) {
            localStorage.removeItem('app_session');
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    checkSession();

    // Setup drag & drop untuk upload
    const dropzone = document.getElementById('dropzone');
    if (dropzone) {
        ['dragenter', 'dragover'].forEach(evt => {
            dropzone.addEventListener(evt, e => {
                e.preventDefault();
                dropzone.classList.add('dragover');
            });
        });
        ['dragleave', 'drop'].forEach(evt => {
            dropzone.addEventListener(evt, e => {
                e.preventDefault();
                dropzone.classList.remove('dragover');
            });
        });
        dropzone.addEventListener('drop', e => {
            const file = e.dataTransfer.files[0];
            if (file) {
                const input = document.getElementById('excel-upload');
                const dt = new DataTransfer();
                dt.items.add(file);
                input.files = dt.files;
                handleFileUpload({ target: input });
            }
        });
    }
});

// ==========================================
// INIT APP - MENU BERDASARKAN ROLE
// ==========================================
function initApp() {
    document.getElementById('login-view').style.display = 'none';
    document.getElementById('app-view').style.display = 'block';
    document.getElementById('u-name').innerText = user.nama_admin;
    document.getElementById('u-role').innerText = `${user.role} | ${user.kelompok || user.desa || ''}`;

    const navUpload = document.getElementById('nav-upload');
    const navConfig = document.getElementById('nav-config');
    const navTargets = document.getElementById('nav-targets');

    // Reset semua tab ke hidden dulu
    navUpload.style.display = 'none';
    navConfig.style.display = 'none';
    navTargets.style.display = 'none';

    if (user.role === 'Kelompok') {
        navUpload.style.display = 'block';
        document.getElementById('tpl-kelompok-name').innerText = user.kelompok || '-';
        switchTab('upload');
    } else if (user.role === 'Desa') {
        switchTab('data');
    } else if (user.role === 'Daerah') {
        navConfig.style.display = 'block';
        navTargets.style.display = 'block';
        document.getElementById('filter-desa-container').style.display = 'block';
        document.getElementById('filter-desa-rekap-container').style.display = 'block';
        switchTab('data');
    }

    fetchConfig();
}

// ==========================================
// DATA FETCHING
// ==========================================
function fetchConfig() {
    securePost({ action: 'get_config' })
    .then(res => {
        if (res.status !== 'success') return alert(res.message);

        configData = res.config.map(r => ({
            cat_code: r.cat_code, cat_title: r.cat_title,
            item_id: r.item_id, item_label: r.item_label,
            tipe: r.tipe, allowed_groups: r.allowed_groups || "ALL",
            is_split: r.is_split || false,
            split_kel: r.split_kel || 0, split_desa: r.split_desa || 0, split_daerah: r.split_daerah || 0
        }));

        targetsData = res.targets || {};
        fetchData();

        if (user.role === 'Daerah') renderConfigTable();
    });
}

function fetchData() {
    securePost({ action: 'fetch_data' })
    .then(res => {
        if (res.status === 'success') {
            allTransactions = res.data;
            updateDesaFilterOptions();
            renderDataTable();
            renderRecapTable();
        }
    });
}

// ==========================================
// UPLOAD EXCEL (KELOMPOK)
// ==========================================

// Step 1: Download Template
async function downloadTemplate() {
    if (!configData || configData.length === 0) {
        return showToast("Config belum dimuat. Coba refresh halaman.");
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Input Data');

    // Header style
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B82F6' } };
    const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    const border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
    };

    // Kolom A & B fixed
    sheet.getColumn(1).width = 14;
    sheet.getColumn(2).width = 30;

    // Header row 1
    const headers = ['Periode', 'Nama Warga'];
    configData.forEach(c => headers.push(c.item_id));

    const headerRow = sheet.addRow(headers);
    headerRow.eachCell(cell => {
        cell.font = headerFont;
        cell.fill = headerFill;
        cell.border = border;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    headerRow.height = 25;

    // Baris contoh (contoh format periode)
    const contohRow = ['2025-01', 'Contoh: Nama Warga (HAPUS baris ini)'];
    configData.forEach(() => contohRow.push(0));
    const r2 = sheet.addRow(contohRow);
    r2.eachCell(cell => {
        cell.font = { italic: true, color: { argb: 'FF999999' } };
        cell.border = border;
    });

    // Sheet referensi item
    const refSheet = workbook.addWorksheet('Referensi Item');
    const refHeader = refSheet.addRow(['Kategori', 'ID Item', 'Label Item', 'Tipe']);
    refHeader.eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = headerFill;
        cell.border = border;
    });
    configData.forEach(c => {
        const r = refSheet.addRow([c.cat_title, c.item_id, c.item_label, c.tipe]);
        r.eachCell(cell => cell.border = border);
    });
    refSheet.getColumn(1).width = 25;
    refSheet.getColumn(2).width = 18;
    refSheet.getColumn(3).width = 35;
    refSheet.getColumn(4).width = 12;

    // Freeze header
    sheet.views = [{ state: 'frozen', ySplit: 1 }];

    // Note di bawah
    const noteRow = sheet.addRow([]);
    sheet.addRow(['CATATAN:']);
    sheet.addRow(['- Format Periode: YYYY-MM (contoh: 2025-01)']);
    sheet.addRow(['- Isi kolom item sesuai nominal (tanpa titik/koma, contoh: 50000)']);
    sheet.addRow(['- Baris contoh di atas mohon dihapus sebelum upload']);
    sheet.addRow(['- Jangan mengubah nama header kolom']);

    // Generate & download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Template_Input_${(user.kelompok || 'Kelompok').replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 7)}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("✅ Template berhasil didownload!");
}

// Step 2: Handle File Upload
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showToast("Format file harus .xlsx atau .xls");
        return;
    }

    try {
        showToast("Membaca file...");
        const workbook = new ExcelJS.Workbook();
        const buffer = await file.arrayBuffer();
        await workbook.xlsx.load(buffer);

        const sheet = workbook.getWorksheet('Input Data') || workbook.worksheets[0];
        if (!sheet) throw new Error("Sheet tidak ditemukan");

        // Baca header
        const headers = [];
        const headerRow = sheet.getRow(1);
        headerRow.eachCell((cell, colNumber) => {
            headers[colNumber] = String(cell.value || '').trim();
        });

        // Validasi header dasar
        if (headers[1] !== 'Periode' || headers[2] !== 'Nama Warga') {
            throw new Error("Format header tidak sesuai template! Kolom A=Periode, B=Nama Warga");
        }

        // Kumpulkan item_id valid
        const validItemIds = configData.map(c => c.item_id);

        const rows = [];
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // skip header

            const periode = String(row.getCell(1).value || '').trim();
            const nama = String(row.getCell(2).value || '').trim();

            // Skip baris kosong atau contoh
            if (!periode || !nama) return;
            if (nama.toLowerCase().includes('contoh') || nama.toLowerCase().includes('hapus')) return;
            if (!/^\d{4}-\d{2}$/.test(periode)) {
                console.warn(`Baris ${rowNumber}: format periode tidak valid: ${periode}`);
                return;
            }

            const rowData = { periode, nama_warga: nama, values: {} };

            for (let col = 3; col <= headers.length; col++) {
                const itemId = headers[col];
                if (!itemId || !validItemIds.includes(itemId)) continue;

                const cellVal = row.getCell(col).value;
                let nominal = 0;
                if (typeof cellVal === 'number') nominal = cellVal;
                else if (typeof cellVal === 'string') nominal = parseRupiah(cellVal);
                else if (cellVal && typeof cellVal === 'object' && cellVal.result !== undefined) {
                    nominal = parseFloat(cellVal.result) || 0;
                }

                if (nominal > 0) rowData.values[itemId] = nominal;
            }

            // Hanya ambil baris yang minimal 1 nilai
            if (Object.keys(rowData.values).length > 0) {
                rows.push(rowData);
            }
        });

        if (rows.length === 0) {
            throw new Error("Tidak ada data valid yang ditemukan. Periksa format file Anda.");
        }

        parsedExcelData = rows;
        renderPreview(rows);
        showToast(`✅ ${rows.length} baris berhasil dibaca`);

    } catch (err) {
        console.error(err);
        showToast("❌ " + err.message);
        resetUpload();
    }
}

// Step 3: Preview
function renderPreview(rows) {
    const container = document.getElementById('preview-container');
    const allItemIds = [...new Set(rows.flatMap(r => Object.keys(r.values)))];

    let html = '<table style="min-width:auto; margin:0; font-size:0.75rem;">';
    html += '<thead style="position:sticky; top:0; background:#f1f5f9; z-index:10;"><tr>';
    html += '<th style="padding:8px;">Periode</th><th style="padding:8px;">Nama</th>';
    allItemIds.forEach(id => html += `<th style="padding:8px; font-size:0.7rem;">${id}</th>`);
    html += '<th style="padding:8px;">Total</th></tr></thead><tbody>';

    let grandTotal = 0;
    rows.forEach(r => {
        const total = Object.values(r.values).reduce((a, b) => a + b, 0);
        grandTotal += total;
        html += '<tr>';
        html += `<td style="padding:6px 8px;">${r.periode}</td>`;
        html += `<td style="padding:6px 8px;">${r.nama_warga}</td>`;
        allItemIds.forEach(id => {
            const val = r.values[id] || 0;
            html += `<td style="padding:6px 8px; text-align:right;">${val ? val.toLocaleString('id-ID') : '-'}</td>`;
        });
        html += `<td style="padding:6px 8px; text-align:right; font-weight:bold; color:var(--accent);">${total.toLocaleString('id-ID')}</td>`;
        html += '</tr>';
    });

    html += `<tr style="background:#f8fafc; font-weight:bold;">`;
    html += `<td colspan="${2 + allItemIds.length}" style="padding:8px; text-align:right;">TOTAL KESELURUHAN:</td>`;
    html += `<td style="padding:8px; text-align:right; color:var(--success);">${grandTotal.toLocaleString('id-ID')}</td>`;
    html += '</tr></tbody></table>';

    container.innerHTML = html;
    document.getElementById('preview-count').innerText = rows.length;
    document.getElementById('upload-preview').style.display = 'flex';
}

// Step 4: Confirm & Save
async function confirmUpload(event) {
    if (!parsedExcelData || parsedExcelData.length === 0) {
        return showToast("Tidak ada data untuk disimpan");
    }

    const btn = event.target;
    const originalHTML = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-spinner"></i> Menyimpan...';

    try {
        const res = await securePost({
            action: 'batch_create',
            data: parsedExcelData
        });

        if (res.status === 'success') {
            showToast(`✅ ${res.saved} data berhasil disimpan!`);
            resetUpload();
            fetchData();
            setTimeout(() => switchTab('data'), 800);
        } else {
            throw new Error(res.message);
        }
    } catch (err) {
        showToast("❌ " + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalHTML;
    }
}

function resetUpload() {
    parsedExcelData = null;
    const input = document.getElementById('excel-upload');
    if (input) input.value = '';
    document.getElementById('upload-preview').style.display = 'none';
}

// ==========================================
// DATA TABLE
// ==========================================
function renderDataTable() {
    const tbody = document.getElementById('table-body-data');
    const filterPeriode = document.getElementById('filter-data-periode').value;
    const filterDesa = document.getElementById('filter-data-desa').value;

    if (!tbody) return;
    tbody.innerHTML = '';
    let grandTotal = 0;

    const filtered = allTransactions.filter(t => {
        let match = true;
        if (user.role === 'Kelompok') match = t.kelompok === user.kelompok;
        else if (user.role === 'Desa') match = t.desa === user.desa;
        if (match && filterPeriode) match = (t.periode && String(t.periode).trim() === filterPeriode);
        if (match && user.role === 'Daerah' && filterDesa) match = (t.desa && String(t.desa).trim() === filterDesa);
        return match;
    });

    filtered.reverse().forEach(t => {
        let valObj = {};
        try { valObj = JSON.parse(t.values_json); } catch (e) { }
        const total = Object.values(valObj).reduce((a, b) => a + (parseFloat(b) || 0), 0);
        grandTotal += total;

        const displayDate = t.timestamp ? String(t.timestamp).split(' ')[0] : '-';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${displayDate}</td>
            <td>${t.periode}</td>
            <td>${t.nama_warga}</td>
            <td>${t.desa || t.kelompok || '-'}</td>
            <td style="font-weight:bold; color:var(--accent)">Rp ${total.toLocaleString('id-ID')}</td>
            <td>
                <button class="btn btn-sm btn-danger" onclick="deleteData('${t.id}')"><i class="ph ph-trash"></i></button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    const grandTotalEl = document.getElementById('grand-total-data');
    if (grandTotalEl) grandTotalEl.innerText = `Rp ${grandTotal.toLocaleString('id-ID')}`;
}

// ==========================================
// REKAP TABLE
// ==========================================
function renderRecapTable() {
    const tbody = document.getElementById('table-body-rekap');
    const filterPeriodeRaw = document.getElementById('filter-rekap-periode').value;
    const filterDesa = document.getElementById('filter-rekap-desa').value;

    if (!tbody) return;
    tbody.innerHTML = '';

    if (!filterPeriodeRaw) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:2rem; color:var(--text-light);">Silakan pilih periode terlebih dahulu.</td></tr>';
        const footer = document.getElementById('grand-total-rekap');
        if (footer) footer.innerText = "Rp 0";
        return;
    }

    const [selectedYear, selectedMonthStr] = filterPeriodeRaw.split('-');
    const selectedMonth = parseInt(selectedMonthStr);

    let targetIdentity;
    if (user.role === 'Kelompok') targetIdentity = user.kelompok;
    else if (user.role === 'Desa') targetIdentity = user.desa;
    else if (user.role === 'Daerah') targetIdentity = filterDesa || "Daerah";
    else targetIdentity = "";

    const currentTargets = targetsData[targetIdentity] || {};
    const realisasiBlnIni = {};
    const akumulasiYTD = {};
    let grandTotalBlnIni = 0;

    allTransactions.forEach(t => {
        const tPeriode = String(t.periode).trim();
        const [tYear, tMonthStr] = tPeriode.split('-');
        const tMonth = parseInt(tMonthStr);

        let matchUser = true;
        if (user.role === 'Kelompok') matchUser = t.kelompok === user.kelompok;
        else if (user.role === 'Desa') matchUser = t.desa === user.desa;
        if (matchUser && user.role === 'Daerah' && filterDesa) matchUser = t.desa === filterDesa;

        if (matchUser && tYear === selectedYear) {
            let valObj = {};
            try { valObj = JSON.parse(t.values_json); } catch (e) { }
            for (let [key, val] of Object.entries(valObj)) {
                const nominal = parseFloat(val) || 0;
                akumulasiYTD[key] = (akumulasiYTD[key] || 0) + nominal;
                if (tMonth === selectedMonth) {
                    realisasiBlnIni[key] = (realisasiBlnIni[key] || 0) + nominal;
                    grandTotalBlnIni += nominal;
                }
            }
        }
    });

    configData.forEach(conf => {
        const id = conf.item_id;
        const bln = realisasiBlnIni[id] || 0;
        const ytd = akumulasiYTD[id] || 0;
        const tipe = (conf.tipe || "Tahunan").trim();
        const targetVal = parseFloat(currentTargets[id]) || 0;

        let persen = 0;
        const labelTarget = tipe === "Bulanan" ? "Target/Bln" : "Target/Thn";

        if (tipe === "Bulanan") persen = targetVal > 0 ? Math.round((bln / targetVal) * 100) : 0;
        else persen = targetVal > 0 ? Math.round((ytd / targetVal) * 100) : 0;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <div style="font-weight:600; font-size:0.85rem;">${conf.item_label}</div>
                <div style="font-size:0.6rem; color:#fff; background:${tipe === 'Bulanan' ? '#8b5cf6' : '#3b82f6'}; padding:1px 6px; border-radius:10px; display:inline-block; margin-top:2px;">
                    ${tipe}
                </div>
            </td>
            <td style="text-align:right; font-weight:700; color:var(--accent);">
                Rp ${bln.toLocaleString('id-ID')}
            </td>
            <td style="text-align:right; font-size:0.8rem; color:var(--text-light);">
                ${tipe === 'Bulanan' ? '-' : 'Rp ' + ytd.toLocaleString('id-ID')}
            </td>
            <td style="text-align:right;">
                <div style="font-size:0.8rem; font-weight:600;">Rp ${targetVal.toLocaleString('id-ID')}</div>
                <div style="font-size:0.55rem; color:var(--text-light); text-transform:uppercase;">${labelTarget}</div>
            </td>
            <td>
                <div style="font-size:0.7rem; margin-bottom:3px; display:flex; justify-content:space-between;">
                    <span style="font-weight:bold">${persen}%</span>
                    <span style="font-size:0.6rem; opacity:0.7">${tipe === 'Tahunan' ? 'YTD' : 'REAL'}</span>
                </div>
                <div class="progress-container">
                    <div class="progress-bar" style="width: ${persen > 100 ? 100 : persen}%; background: ${persen >= 100 ? '#22c55e' : (tipe === 'Bulanan' ? '#a78bfa' : '#3b82f6')};"></div>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    const footer = document.getElementById('grand-total-rekap');
    if (footer) footer.innerText = `Rp ${grandTotalBlnIni.toLocaleString('id-ID')}`;
}

// ==========================================
// CONFIG TABLE
// ==========================================
function renderConfigTable() {
    const tbody = document.getElementById('table-body-config');
    if (!tbody) return;
    tbody.innerHTML = '';

    configData.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><b>${item.item_id}</b></td>
            <td>${item.cat_title}</td>
            <td>${item.item_label}</td>
            <td><span style="font-size:0.75rem; background:${item.tipe === 'Bulanan' ? '#a78bfa' : '#3b82f6'}; color:white; padding:2px 6px; border-radius:4px;">${item.tipe}</span></td>
            <td>${item.allowed_groups || "ALL"}</td>
            <td>
                <button class="btn btn-sm" onclick="editConfig('${item.item_id}')"><i class="ph ph-pencil-simple"></i></button>
                <button class="btn btn-sm btn-danger" onclick="deleteConfig('${item.item_id}')"><i class="ph ph-trash"></i></button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function openConfigModal(isEdit = false) {
    document.getElementById('config-modal').style.display = 'flex';
    document.getElementById('modal-title-config').innerText = isEdit ? "Edit Item Config" : "Tambah Item Config";

    if (!isEdit) {
        ['cfg-old-id', 'cfg-item-id', 'cfg-cat-code', 'cfg-cat-title', 'cfg-item-label', 'cfg-groups']
            .forEach(id => document.getElementById(id).value = '');
        document.getElementById('cfg-tipe').value = 'Bulanan';
        document.getElementById('cfg-is-split').checked = false;
        document.getElementById('split-inputs').style.display = 'none';
        ['cfg-split-kel', 'cfg-split-desa', 'cfg-split-dah'].forEach(id => document.getElementById(id).value = 0);
    }
}

function closeConfigModal() {
    document.getElementById('config-modal').style.display = 'none';
}

function editConfig(id) {
    const item = configData.find(x => x.item_id === id);
    if (!item) return alert("Data tidak ditemukan!");

    document.getElementById('cfg-old-id').value = item.item_id;
    document.getElementById('cfg-item-id').value = item.item_id;
    document.getElementById('cfg-cat-code').value = item.cat_code || '';
    document.getElementById('cfg-cat-title').value = item.cat_title || '';
    document.getElementById('cfg-item-label').value = item.item_label || '';
    document.getElementById('cfg-tipe').value = item.tipe || 'Bulanan';
    document.getElementById('cfg-groups').value = item.allowed_groups || '';

    const isSplit = item.is_split || false;
    document.getElementById('cfg-is-split').checked = isSplit;
    document.getElementById('split-inputs').style.display = isSplit ? 'grid' : 'none';
    document.getElementById('cfg-split-kel').value = item.split_kel || 0;
    document.getElementById('cfg-split-desa').value = item.split_desa || 0;
    document.getElementById('cfg-split-dah').value = item.split_daerah || 0;

    openConfigModal(true);
}

function saveConfig() {
    const oldId = document.getElementById('cfg-old-id').value.trim();
    const itemId = document.getElementById('cfg-item-id').value.trim();
    const catCode = document.getElementById('cfg-cat-code').value.trim();
    const catTitle = document.getElementById('cfg-cat-title').value.trim();
    const itemLabel = document.getElementById('cfg-item-label').value.trim();
    const tipe = document.getElementById('cfg-tipe').value;
    const groups = document.getElementById('cfg-groups').value.trim();

    const isSplit = document.getElementById('cfg-is-split').checked;
    const splitKel = parseFloat(document.getElementById('cfg-split-kel').value) || 0;
    const splitDesa = parseFloat(document.getElementById('cfg-split-desa').value) || 0;
    const splitDah = parseFloat(document.getElementById('cfg-split-dah').value) || 0;

    if (!itemId || !catCode || !catTitle || !itemLabel) {
        return showToast("Lengkapi data wajib!");
    }
    if (isSplit && (splitKel + splitDesa + splitDah !== 100)) {
        return showToast("Total persentase split harus 100%!");
    }

    const payload = {
        action: oldId ? 'update_config' : 'create_config',
        old_id: oldId, item_id: itemId,
        cat_code: catCode, cat_title: catTitle, item_label: itemLabel,
        tipe: tipe, allowed_groups: groups || "ALL",
        is_split: isSplit, split_kel: splitKel, split_desa: splitDesa, split_daerah: splitDah
    };

    securePost(payload)
    .then(res => {
        if (res.status === 'success') {
            showToast("✅ Config disimpan!");
            closeConfigModal();
            fetchConfig();
        } else {
            showToast("Error: " + res.message);
        }
    });
}

function deleteConfig(id) {
    if (!confirm("Hapus item ini?")) return;
    securePost({ action: 'delete_config', item_id: id }).then(res => {
        if (res.status === 'success') fetchConfig();
    });
}

document.getElementById('cfg-is-split').addEventListener('change', function () {
    document.getElementById('split-inputs').style.display = this.checked ? 'grid' : 'none';
});

// ==========================================
// TARGETS
// ==========================================
function renderTargetsTable() {
    const thead = document.getElementById('thead-targets');
    const tbody = document.getElementById('tbody-targets');
    if (!thead || !tbody) return;

    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (!targetsData || Object.keys(targetsData).length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" style="text-align:center; padding:2rem;">Belum ada data Target.</td></tr>';
        return;
    }

    const itemsList = configData.map(c => c.item_id);

    let headerHTML = '<tr><th style="position:sticky; left:0; background:#f8fafc; z-index:10;">Identitas</th>';
    itemsList.forEach(itemId => {
        const itemConf = configData.find(c => c.item_id === itemId);
        const label = itemConf ? itemConf.item_label : itemId;
        headerHTML += `<th style="min-width:120px; text-align:center; font-size:0.75rem;">${label}</th>`;
    });
    headerHTML += '</tr>';
    thead.innerHTML = headerHTML;

    Object.keys(targetsData).forEach(identity => {
        const tr = document.createElement('tr');
        let rowHTML = `<td style="font-weight:bold; position:sticky; left:0; background:white; z-index:5; border-right:2px solid #ddd;">${identity}</td>`;
        const dataRow = targetsData[identity];

        itemsList.forEach(itemId => {
            const val = dataRow[itemId] || 0;
            rowHTML += `
                <td style="text-align:center;">
                    <input type="number" class="target-input" value="${val}"
                           style="width:100px; text-align:center; padding:5px; border:1px solid #ddd; border-radius:4px;"
                           onblur="updateTarget(this, '${identity}', '${itemId}')">
                </td>
            `;
        });

        tr.innerHTML = rowHTML;
        tbody.appendChild(tr);
    });
}

function updateTarget(inputEl, identity, itemId) {
    const newValue = inputEl.value;
    inputEl.style.background = '#fef9c3';

    securePost({
        action: 'update_target',
        identity: identity,
        item_id: itemId,
        value: newValue
    })
    .then(res => {
        if (res.status === 'success') {
            inputEl.style.background = '#dcfce7';
            setTimeout(() => inputEl.style.background = 'white', 1000);
            if (targetsData[identity]) {
                targetsData[identity][itemId] = parseFloat(newValue) || 0;
            }
        } else {
            alert("Gagal: " + res.message);
            inputEl.style.background = '#fee2e2';
        }
    })
    .catch(() => {
        inputEl.style.background = '#fee2e2';
    });
}

// ==========================================
// ACTIONS
// ==========================================
function deleteData(id) {
    if (!confirm("Yakin ingin menghapus data ini?")) return;
    securePost({ action: 'delete', id: id })
    .then(res => {
        if (res.status === 'success') {
            showToast("✅ Data dihapus");
            fetchData();
        } else {
            showToast("Gagal: " + res.message);
        }
    });
}

// ==========================================
// HELPERS
// ==========================================
function switchTab(tabName) {
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.view-section').forEach(s => s.classList.remove('active'));

    const nav = document.getElementById(`nav-${tabName}`);
    const view = document.getElementById(`view-${tabName}`);
    if (nav) nav.classList.add('active');
    if (view) view.classList.add('active');

    if (tabName === 'config' && user.role === 'Daerah') renderConfigTable();
    if (tabName === 'targets' && user.role === 'Daerah') renderTargetsTable();
    if (tabName === 'rekap') renderRecapTable();
}

function updateDesaFilterOptions() {
    if (user.role !== 'Daerah') return;
    const desas = [...new Set(allTransactions.map(t => t.desa))].filter(Boolean);
    const selects = [document.getElementById('filter-data-desa'), document.getElementById('filter-rekap-desa')];
    selects.forEach(sel => {
        if (!sel) return;
        sel.innerHTML = '<option value="">Semua Desa</option>';
        desas.forEach(d => sel.innerHTML += `<option value="${d}">${d}</option>`);
    });
}

// ==========================================
// EXCEL EXPORT (Rekap)
// ==========================================
async function downloadExcel() {
    const filterPeriodeRaw = document.getElementById('filter-rekap-periode').value;
    const filterDesa = document.getElementById('filter-rekap-desa').value;

    if (!filterPeriodeRaw) return showToast("Pilih periode dulu!");

    const [year, month] = filterPeriodeRaw.split('-');
    const dateObj = new Date(year, month - 1);
    const namaPeriode = dateObj.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });

    let namaKelompok = (user.role === 'Kelompok' ? user.kelompok : ".........").toUpperCase();
    let namaDesaDisplay = (user.desa || ".........").toUpperCase();

    if (user.role === 'Daerah') {
        namaDesaDisplay = filterDesa ? filterDesa.toUpperCase() : "";
    }

    let allowedCategories = [];
    if (user.role === 'Daerah') allowedCategories = ['A', 'B', 'C', 'D'];
    else if (user.role === 'Desa') allowedCategories = ['A', 'B', 'C', 'D', 'E'];
    else if (user.role === 'Kelompok') allowedCategories = ['A', 'B', 'C', 'D', 'E', 'F'];

    let reportTitle = user.role === 'Kelompok' ? "REKAPITULASI HIBAH BULANAN KELOMPOK KE DESA" :
        (user.role === 'Desa' ? "REKAPITULASI HIBAH BULANAN DESA KE DAERAH" : "REKAPITULASI HIBAH BULANAN DAERAH");

    let reportSubtitle = "";
    if (user.role === 'Kelompok') reportSubtitle = `KELOMPOK ${namaKelompok} DESA ${namaDesaDisplay}`;
    else if (user.role === 'Desa' || (user.role === 'Daerah' && filterDesa)) reportSubtitle = `DESA ${namaDesaDisplay}`;

    let targetKey = user.role === 'Kelompok' ? user.kelompok : (user.role === 'Desa' ? user.desa : (filterDesa || "Daerah"));
    const currentTargets = targetsData[targetKey] || {};

    const categorizedData = {};
    configData.forEach(conf => {
        if (!allowedCategories.includes(conf.cat_code.toUpperCase())) return;
        if (!categorizedData[conf.cat_title]) categorizedData[conf.cat_title] = [];

        const realisasiBln = allTransactions.filter(t => {
            const tPeriode = String(t.periode).trim();
            let matchRole = true;
            if (user.role === 'Kelompok') matchRole = t.kelompok === user.kelompok;
            else if (user.role === 'Desa') matchRole = t.desa === user.desa;
            else if (user.role === 'Daerah' && filterDesa) matchRole = t.desa === filterDesa;
            return tPeriode === filterPeriodeRaw && matchRole;
        }).reduce((acc, t) => {
            let vals = {}; try { vals = JSON.parse(t.values_json); } catch (e) { }
            return acc + (parseFloat(vals[conf.item_id]) || 0);
        }, 0);

        categorizedData[conf.cat_title].push({
            cat_code: conf.cat_code,
            label: conf.item_label,
            target: parseFloat(currentTargets[conf.item_id]) || 0,
            realisasi: realisasiBln
        });
    });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Laporan');

    worksheet.mergeCells('A1:E1');
    worksheet.mergeCells('A2:E2');
    worksheet.mergeCells('A3:E3');

    worksheet.getCell('A1').value = reportTitle;
    worksheet.getCell('A1').font = { bold: true, size: 14 };
    worksheet.getCell('A1').alignment = { horizontal: 'center' };

    worksheet.getCell('A2').value = reportSubtitle;
    worksheet.getCell('A2').font = { bold: true, size: 11 };
    worksheet.getCell('A2').alignment = { horizontal: 'center' };

    worksheet.getCell('A3').value = "PERIODE " + namaPeriode;
    worksheet.getCell('A3').font = { bold: true, size: 11 };
    worksheet.getCell('A3').alignment = { horizontal: 'center' };

    const headerRow = worksheet.addRow(['NO', 'URAIAN', 'TARGET', 'REALISASI', 'KETERANGAN']);
    headerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.alignment = { horizontal: 'center' };
    });

    let counter = 1;
    let grandTotalRealisasi = 0;
    let totalHakKelompok = 0;
    let totalHakDesa = 0;

    for (const [catTitle, items] of Object.entries(categorizedData)) {
        const cat = configData.find(c => c.cat_title === catTitle);
        const catCode = cat ? cat.cat_code : '';

        const catRow = worksheet.addRow(['', catTitle.toUpperCase(), '', '', '']);
        catRow.getCell(2).font = { bold: true };

        items.forEach(item => {
            grandTotalRealisasi += item.realisasi;
            if (catCode === 'F') totalHakKelompok += item.realisasi;
            if (catCode === 'E') totalHakDesa += item.realisasi;

            const row = worksheet.addRow([counter++, item.label, item.target, item.realisasi, ""]);
            row.getCell(3).numFmt = '#,##0';
            row.getCell(4).numFmt = '#,##0';
        });
    }

    const footerRow = worksheet.addRow(['', 'TOTAL KESELURUHAN', '', grandTotalRealisasi, '']);
    footerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    });
    footerRow.getCell(4).numFmt = '#,##0';

    let totalYangDisetor = 0;
    let labelSetoran = "";

    if (user.role === 'Kelompok') {
        totalYangDisetor = grandTotalRealisasi - totalHakKelompok;
        labelSetoran = "TOTAL YANG DISETOR KE DESA";
    } else if (user.role === 'Desa') {
        totalYangDisetor = grandTotalRealisasi - totalHakDesa;
        labelSetoran = "TOTAL YANG DISETOR KE DAERAH";
    }

    if ((user.role === 'Kelompok' || user.role === 'Desa') && totalYangDisetor > 0) {
        const setoranRow = worksheet.addRow(['', labelSetoran, '', totalYangDisetor, '']);
        setoranRow.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });
        setoranRow.getCell(4).numFmt = '#,##0';
    }

    worksheet.getColumn(1).width = 5;
    worksheet.getColumn(2).width = 45;
    worksheet.getColumn(3).width = 18;
    worksheet.getColumn(4).width = 18;
    worksheet.getColumn(5).width = 20;

    let fileName = (user.role === 'Kelompok' ? `Laporan_Kelompok_${namaKelompok}` :
        (user.role === 'Desa' ? `Laporan_Desa_${namaDesaDisplay}` :
            `Laporan_Daerah_Keseluruhan`)) + `_${filterPeriodeRaw}.xlsx`;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    window.URL.revokeObjectURL(url);
    showToast("✅ File Excel berhasil didownload!");
}
