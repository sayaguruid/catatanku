// Ganti URL ini dengan URL Web App Google Script Anda
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyo65StO07OygmbXGwFzoE-FVCGB9u-VfC9S9IL1uv78XxeeZpRMrLhdMlOLJqXiY2H/exec'; 

let user = null;
let authToken = null;
let configData = [];
let allTransactions = [];
let isEditing = false;
let editId = null;
let targetsData = {}; 

// --- UTILS ---
function parseRupiah(str) {
    if (!str) return 0;
    let cleaned = String(str).replace(/\./g, '').replace(/,/g, '.');
    return parseFloat(cleaned) || 0;
}

function formatRupiah(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function formatInputOnKey(el) {
    let cursorPosition = el.selectionStart;
    let originalLength = el.value.length;
    let value = el.value.replace(/[^0-9]/g, '');
    el.value = value.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    let newLength = el.value.length;
    if (value !== "") {
        el.selectionStart = el.selectionEnd = cursorPosition + (newLength - originalLength);
    }
    updateLiveTotal();
}

function updateLiveTotal() {
    let total = 0;
    document.querySelectorAll('.money-input').forEach(inp => {
        total += parseRupiah(inp.value);
    });
    document.getElementById('live-total').innerText = "Rp " + formatRupiah(total);
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
    t.innerText = msg; t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 3000);
}

// --- AUTH ---
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
            localStorage.setItem('app_token', authToken); 
            initApp();
        } else {
            msg.innerText = res.message;
        }
    })
    .catch(e => msg.innerText = "Koneksi Error");
}

function initApp() {
    document.getElementById('login-view').style.display = 'none';
    document.getElementById('app-view').style.display = 'block';
    document.getElementById('u-name').innerText = user.nama_admin;
    document.getElementById('u-role').innerText = `${user.role} | ${user.kelompok || user.desa || ''}`;
    
    if(user.role === 'Daerah') {
        document.getElementById('filter-desa-container').style.display = 'block';
        document.getElementById('filter-desa-rekap-container').style.display = 'block';
        document.getElementById('nav-config').style.display = 'block';
        document.getElementById('nav-targets').style.display = 'block';
    }

    fetchConfig();
}

// --- DATA FETCHING ---
function fetchConfig() {
    securePost({ action: 'get_config' }) 
    .then(res => {
        if(res.status !== 'success') return alert(res.message);
        configData = res.config.map(r => ({
            cat_code: r.cat_code, cat_title: r.cat_title, item_id: r.item_id, item_label: r.item_label,
            tipe: r.tipe, allowed_groups: r.allowed_groups || "ALL",
            is_split: r.is_split || false, split_kel: r.split_kel || 0, split_desa: r.split_desa || 0, split_daerah: r.split_daerah || 0
        }));
        
        // Pastikan targetsData adalah objek yang valid
        targetsData = res.targets || {}; 
        buildForm();
        fetchData();
        
        if(user.role === 'Daerah') renderConfigTable();
    });
}

function fetchData() {
    securePost({ action: 'fetch_data' })
    .then(res => {
        if(res.status === 'success') {
            allTransactions = res.data;
            updateDesaFilterOptions();
            renderDataTable();
            renderRecapTable(); // Panggil render rekap saat data fetch
        }
    });
}

// --- UI BUILDERS ---
function buildForm() {
    const tabsContainer = document.getElementById('dynamic-tabs');
    const itemsContainer = document.getElementById('dynamic-items');
    tabsContainer.innerHTML = ''; itemsContainer.innerHTML = '';

    const categories = {};
    configData.forEach(item => {
        if (!categories[item.cat_code]) {
            categories[item.cat_code] = { title: item.cat_title, items: [] };
        }
        categories[item.cat_code].items.push(item);
    });

    Object.keys(categories).forEach((code, index) => {
        const cat = categories[code];
        const btn = document.createElement('button');
        btn.className = `cat-tab ${index === 0 ? 'active' : ''}`;
        btn.id = `tab-btn-${code}`;
        btn.innerText = `${code}. ${cat.title}`;
        btn.onclick = () => switchCategory(code);
        tabsContainer.appendChild(btn);

        const pane = document.createElement('div');
        pane.id = `pane-${code}`;
        pane.className = `cat-pane ${index === 0 ? 'active' : ''}`;
        cat.items.forEach(item => {
            const row = document.createElement('div');
            row.className = 'item-row';
            row.innerHTML = `
                <label style="font-size:0.85rem">${item.item_label}</label>
                <div class="currency-wrap">
                    <span>Rp</span>
                    <input type="text" inputmode="numeric" class="form-control money-input" data-id="${item.item_id}" placeholder="0" oninput="formatInputOnKey(this)">
                </div>
                <button class="btn btn-sm" style="background:#e2e8f0;" onclick="quickFill(this, 50000)">+50rb</button>
            `;
            pane.appendChild(row);
            
            if (item.is_split) {
                const inp = row.querySelector('.money-input'); 
                const previewDiv = document.createElement('div');
                previewDiv.className = 'split-preview';
                previewDiv.style.cssText = "font-size:0.7rem; color:#64748b; margin-top:4px; padding:4px; background:#f1f5f9; border-radius:4px;";
                previewDiv.innerHTML = '<i class="ph ph-calculator"></i> <b>Otomatis:</b> <span class="val-kel">0</span> (Ke F) | <span class="val-desa">0</span> (Ke E) | <span class="val-dah">0</span> (Tetap C)';
                
                inp.addEventListener('input', () => {
                    const total = parseRupiah(inp.value);
                    const pKel = (total * (item.split_kel || 0)) / 100;
                    const pDesa = (total * (item.split_desa || 0)) / 100;
                    const pDah = (total * (item.split_daerah || 0)) / 100;
                    
                    previewDiv.querySelector('.val-kel').innerText = formatRupiah(pKel);
                    previewDiv.querySelector('.val-desa').innerText = formatRupiah(pDesa);
                    previewDiv.querySelector('.val-dah').innerText = formatRupiah(pDah);
                });
                row.appendChild(previewDiv); 
            }
        });
        itemsContainer.appendChild(pane);
    });
    document.getElementById('input-periode').value = new Date().toISOString().slice(0, 7);
    updateLiveTotal();
}

function renderDataTable() {
    const tbody = document.getElementById('table-body-data');
    const filterPeriode = document.getElementById('filter-data-periode').value;
    const filterDesa = document.getElementById('filter-data-desa').value;
    
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
        try { valObj = JSON.parse(t.values_json); } catch(e){}
        const total = Object.values(valObj).reduce((a,b)=>a+(parseFloat(b)||0),0);
        grandTotal += total;

        const displayDate = t.timestamp ? String(t.timestamp).split(' ')[0] : '-';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${displayDate}</td>
            <td>${t.periode}</td>
            <td>${t.nama_warga}</td>
            <td>${t.desa || '-'}</td>
            <td style="font-weight:bold; color:var(--accent)">Rp ${total.toLocaleString('id-ID')}</td>
            <td>
                <button class="btn btn-sm" onclick="editData('${t.id}')"><i class="ph ph-pencil-simple"></i></button>
                <button class="btn btn-sm btn-danger" onclick="deleteData('${t.id}')"><i class="ph ph-trash"></i></button>
            </td>
        `;
        tbody.appendChild(tr);
    });
    document.getElementById('grand-total-data').innerText = `Rp ${grandTotal.toLocaleString('id-ID')}`;
}

// --- FIX: RENDER REKAP TABLE (HTML) ---
function renderRecapTable() {
    const tbody = document.getElementById('table-body-rekap');
    const filterPeriodeRaw = document.getElementById('filter-rekap-periode').value;
    const filterDesa = document.getElementById('filter-rekap-desa').value; 
    tbody.innerHTML = '';
    
    // Jika periode belum dipilih, tampilkan pesan
    if (!filterPeriodeRaw) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:2rem; color:var(--text-light);">Silakan pilih periode terlebih dahulu.</td></tr>';
        document.getElementById('grand-total-rekap').innerText = "Rp 0";
        return;
    }

    const [selectedYear, selectedMonthStr] = filterPeriodeRaw.split('-');
    const selectedMonth = parseInt(selectedMonthStr);
    
    // Tentukan identitas target berdasarkan user
    let targetIdentity;
    if (user.role === 'Kelompok') {
        targetIdentity = user.kelompok;
    } else if (user.role === 'Desa') {
        targetIdentity = user.desa;
    } else if (user.role === 'Daerah') {
        targetIdentity = filterDesa || "Daerah"; 
    } else {
        targetIdentity = "";
    }

    const currentTargets = targetsData[targetIdentity] || {}; 
    const realisasiBlnIni = {}; 
    const akumulasiYTD = {};           
    let grandTotalBlnIni = 0;

    // Hitung Realisasi
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
            try { valObj = JSON.parse(t.values_json); } catch(e){}
            for (let [key, val] of Object.entries(valObj)) {
                const nominal = parseFloat(val) || 0;
                // Akumulasi YTD
                akumulasiYTD[key] = (akumulasiYTD[key] || 0) + nominal;
                // Realisasi Bulan Ini
                if (tMonth === selectedMonth) {
                    realisasiBlnIni[key] = (realisasiBlnIni[key] || 0) + nominal;
                    grandTotalBlnIni += nominal;
                }
            }
        }
    });

    // Render Baris
    configData.forEach(conf => {
        const id = conf.item_id;
        const bln = realisasiBlnIni[id] || 0;
        const ytd = akumulasiYTD[id] || 0;
        const tipe = (conf.tipe || "Tahunan").trim();
        const targetVal = parseFloat(currentTargets[id]) || 0;
        
        let persen = 0;
        let labelTarget = tipe === "Bulanan" ? "Target/Bln" : "Target/Thn";

        if (tipe === "Bulanan") {
            persen = targetVal > 0 ? Math.round((bln / targetVal) * 100) : 0;
        } else {
            persen = targetVal > 0 ? Math.round((ytd / targetVal) * 100) : 0;
        }

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
                <div class="progress-container" style="background:#e2e8f0; height:6px; border-radius:10px; overflow:hidden;">
                    <div class="progress-bar" style="width: ${persen > 100 ? 100 : persen}%; height:100%; background: ${persen >= 100 ? '#22c55e' : (tipe === 'Bulanan' ? '#a78bfa' : '#3b82f6')}; transition:0.3s;"></div>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('grand-total-rekap').innerText = `Rp ${grandTotalBlnIni.toLocaleString('id-ID')}`;
}

// --- ACTION HANDLERS ---
function submitData() {
    const btn = document.getElementById('btn-submit');
    const periode = document.getElementById('input-periode').value;
    const nama = document.getElementById('input-nama').value;
    if(!periode || !nama) return showToast("Periode & Nama harus diisi!");

    const values = {};
    let hasVal = false;
    document.querySelectorAll('.money-input').forEach(inp => {
        const v = parseRupiah(inp.value); 
        if(v > 0) { values[inp.dataset.id] = v; hasVal = true; }
    });
    if(!hasVal) return showToast("Isi minimal satu nominal!");

    let processedValues = JSON.parse(JSON.stringify(values));

    if (user.role !== 'Daerah') {
        Object.keys(values).forEach(key => {
            const conf = configData.find(c => c.item_id === key);
            if (conf && conf.is_split) {
                const totalAmount = values[key];
                const pKel = (totalAmount * (conf.split_kel || 0)) / 100;
                const pDesa = (totalAmount * (conf.split_desa || 0)) / 100;
                const pDah = (totalAmount * (conf.split_daerah || 0)) / 100;

                const targetSpecificE = configData.find(c => c.item_id === key + "_e");
                const targetSpecificF = configData.find(c => c.item_id === key + "_f");

                processedValues[key] = pDah;

                let destF = targetSpecificF || configData.find(c => c.cat_code === 'F');
                if (destF && pKel > 0) {
                    processedValues[destF.item_id] = (processedValues[destF.item_id] || 0) + pKel;
                }

                let destE = targetSpecificE || configData.find(c => c.cat_code === 'E');
                if (destE && pDesa > 0) {
                    processedValues[destE.item_id] = (processedValues[destE.item_id] || 0) + pDesa;
                }
            }
        });
    }

    btn.disabled = true; btn.innerText = "Menyimpan...";
    const payload = {
        action: isEditing ? 'update' : 'create',
        id: editId, periode: periode, nama_warga: nama, values: processedValues 
    };

    securePost(payload)
    .then(res => {
        if(res.status === 'success') {
            showToast("Berhasil!");
            resetForm(); fetchData();
            if(isEditing) switchTab('data');
        } else {
            showToast("Error: " + res.message);
        }
        btn.disabled = false; btn.innerHTML = `<i class="ph ph-floppy-disk"></i> Simpan Data`;
    })
    .catch(() => { showToast("Error Koneksi"); btn.disabled = false; });
}

function deleteData(id) {
    if(!confirm("Yakin ingin menghapus data ini?")) return;
    securePost({ action: 'delete', id: id })
    .then(res => {
        if(res.status === 'success') {
            showToast("Data dihapus");
            fetchData(); 
        } else {
            showToast("Gagal: " + res.message);
        }
    });
}

// --- EXPORT EXCEL (Menggunakan ExcelJS) ---
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
    if(user.role === 'Kelompok') reportSubtitle = `KELOMPOK ${namaKelompok} DESA ${namaDesaDisplay}`;
    else if(user.role === 'Desa' || (user.role === 'Daerah' && filterDesa)) reportSubtitle = `DESA ${namaDesaDisplay}`;

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
            let vals = {}; try { vals = JSON.parse(t.values_json); } catch(e){}
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
    const worksheet = workbook.addWorksheet('Laporan', { pageSetup: { paperSize: 9, orientation: 'portrait' } });

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
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
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
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        });
        setoranRow.getCell(4).numFmt = '#,##0';
        setoranRow.getCell(4).alignment = { horizontal: 'center' };
    }

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber >= 5) {
            row.eachCell((cell) => {
                if (!cell.fill || cell.fill.fgColor.argb !== 'FF0F766E') {
                   cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                }
                if (cell.address.startsWith('A') || cell.address.startsWith('C') || cell.address.startsWith('D')) {
                    cell.alignment = { horizontal: 'center' };
                }
            });
        }
    });

    worksheet.getColumn(1).width = 5;
    worksheet.getColumn(2).width = 45;
    worksheet.getColumn(3).width = 18;
    worksheet.getColumn(4).width = 18;
    worksheet.getColumn(5).width = 20;

    let fileName = (user.role === 'Kelompok' ? `Laporan_Kelompok_${namaKelompok}` : (user.role === 'Desa' ? `Laporan_Desa_${namaDesaDisplay}` : `Laporan_Daerah_Keseluruhan`)) + `_${filterPeriodeRaw}.xlsx`;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    window.URL.revokeObjectURL(url);
    showToast("File Excel berhasil didownload!");
}

// --- TARGETS VIEW FIX ---
function renderTargetsTable() {
    const thead = document.getElementById('thead-targets');
    const tbody = document.getElementById('tbody-targets');
    thead.innerHTML = ''; 
    tbody.innerHTML = '';

    if (!targetsData || Object.keys(targetsData).length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" style="text-align:center; padding:2rem;">Belum ada data Target. Silakan tambahkan identity (Kelompok/Desa) terlebih dahulu di sheet atau melalui input data pertama kali.</td></tr>';
        return;
    }

    // Ambil daftar ID item dari Config untuk dijadikan kolom
    const itemsList = configData.map(c => c.item_id);

    // 1. Buat Header
    let headerHTML = '<tr><th style="position:sticky; left:0; background:#f8fafc; z-index:10;">Identitas (Kelompok/Desa)</th>';
    itemsList.forEach(itemId => {
        const itemConf = configData.find(c => c.item_id === itemId);
        const label = itemConf ? itemConf.item_label : itemId;
        headerHTML += `<th style="min-width:120px; text-align:center; font-size:0.75rem;">${label}</th>`;
    });
    headerHTML += '</tr>';
    thead.innerHTML = headerHTML;

    // 2. Buat Body (Baris per Identity)
    Object.keys(targetsData).forEach(identity => {
        const tr = document.createElement('tr');
        let rowHTML = `<td style="font-weight:bold; position:sticky; left:0; background:white; z-index:5; border-right:2px solid #ddd;">${identity}</td>`;
        
        const dataRow = targetsData[identity]; // Ini object: { "fr": 100000, "ckm": 50000 }

        itemsList.forEach(itemId => {
            const val = dataRow[itemId] || 0; // Ambil value, default 0 jika belum ada
            rowHTML += `
                <td style="text-align:center;">
                    <input type="number" 
                           class="target-input" 
                           value="${val}" 
                           style="width:100px; text-align:center; padding:5px; border:1px solid #ddd; border-radius:4px;"
                           onblur="updateTarget(this, '${identity}', '${itemId}')"
                           onchange="this.style.background='#dcfce7'; setTimeout(()=>this.style.background='white', 500)">
                </td>
            `;
        });

        tr.innerHTML = rowHTML;
        tbody.appendChild(tr);
    });
}

function updateTarget(inputEl, identity, itemId) {
    const newValue = inputEl.value;
    inputEl.style.background = '#fef9c3'; // Kuning saat proses

    securePost({
        action: 'update_target',
        identity: identity,
        item_id: itemId,
        value: newValue
    })
    .then(res => {
        if(res.status === 'success') {
            inputEl.style.background = '#dcfce7'; // Hijau sukses
            setTimeout(() => inputEl.style.background = 'white', 1000);
            // Update local data
            if(targetsData[identity]) {
                targetsData[identity][itemId] = parseFloat(newValue) || 0;
            }
        } else {
            alert("Gagal menyimpan: " + res.message);
            inputEl.style.background = '#fee2e2'; // Merah gagal
        }
    })
    .catch(e => {
        console.error(e);
        inputEl.style.background = '#fee2e2';
    });
}

// --- HELPERS LAINNYA ---
function switchTab(tabName) {
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.view-section').forEach(s => s.classList.remove('active'));
    document.getElementById(`nav-${tabName}`).classList.add('active');
    document.getElementById(`view-${tabName}`).classList.add('active');
    if(tabName === 'config' && user.role === 'Daerah') renderConfigTable();
    if(tabName === 'targets' && user.role === 'Daerah') renderTargetsTable();
}

function switchCategory(code) {
    document.querySelectorAll('.cat-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.cat-pane').forEach(p => p.classList.remove('active'));
    document.getElementById(`tab-btn-${code}`).classList.add('active');
    document.getElementById(`pane-${code}`).classList.add('active');
}

function quickFill(btn, amount) {
    const inp = btn.parentElement.querySelector('input');
    let newVal = parseRupiah(inp.value) + amount;
    inp.value = formatRupiah(newVal);
    updateLiveTotal();
}

function updateDesaFilterOptions() {
    if(user.role !== 'Daerah') return;
    const desas = [...new Set(allTransactions.map(t => t.desa))].filter(Boolean);
    const selects = [document.getElementById('filter-data-desa'), document.getElementById('filter-rekap-desa')];
    selects.forEach(sel => {
        sel.innerHTML = '<option value="">Semua Desa</option>';
        desas.forEach(d => sel.innerHTML += `<option value="${d}">${d}</option>`);
    });
}

function resetForm() {
    isEditing = false; editId = null;
    document.getElementById('form-title').innerText = "Input Baru";
    document.getElementById('cancel-edit').style.display = 'none';
    document.getElementById('input-nama').value = '';
    document.querySelectorAll('.money-input').forEach(i => i.value = '');
    updateLiveTotal();
}

function editData(id) {
    const t = allTransactions.find(x => x.id == id);
    if(!t) return;
    isEditing = true; editId = id;
    document.getElementById('form-title').innerText = "Edit Data";
    document.getElementById('cancel-edit').style.display = 'block';
    document.getElementById('input-periode').value = t.periode;
    document.getElementById('input-nama').value = t.nama_warga;
    document.querySelectorAll('.money-input').forEach(i => i.value = '');
    try {
        const vals = JSON.parse(t.values_json);
        for(let key in vals) {
            const input = document.querySelector(`.money-input[data-id="${key}"]`);
            if(input) input.value = formatRupiah(vals[key]);
        }
    } catch(e){}
    updateLiveTotal();
    switchTab('input');
}

// Config & Modal Helpers
function renderConfigTable() {
    const tbody = document.getElementById('table-body-config');
    tbody.innerHTML = '';
    configData.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><b>${item.item_id}</b></td>
            <td>${item.cat_title}</td>
            <td>${item.item_label}</td>
            <td><span style="font-size:0.75rem; background:${item.tipe==='Bulanan'?'#a78bfa':'#3b82f6'}; color:white; padding:2px 6px; border-radius:4px;">${item.tipe}</span></td>
            <td>${item.allowed_groups || "ALL"}</td>
            <td>
                <button class="btn btn-sm" onclick="editConfig('${item.item_id}')"><i class="ph ph-pencil-simple"></i></button>
                <button class="btn btn-sm btn-danger" onclick="deleteConfig('${item.item_id}')"><i class="ph ph-trash"></i></button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function saveConfig() {
    const payload = {
        action: document.getElementById('cfg-old-id').value ? 'update_config' : 'create_config',
        item_id: document.getElementById('cfg-item-id').value.trim(),
        cat_code: document.getElementById('cfg-cat-code').value.trim(),
        cat_title: document.getElementById('cfg-cat-title').value.trim(),
        item_label: document.getElementById('cfg-item-label').value.trim(),
        tipe: document.getElementById('cfg-tipe').value,
        allowed_groups: document.getElementById('cfg-groups').value.trim() || "ALL",
        is_split: document.getElementById('cfg-is-split').checked,
        split_kel: parseFloat(document.getElementById('cfg-split-kel').value) || 0,
        split_desa: parseFloat(document.getElementById('cfg-split-desa').value) || 0,
        split_daerah: parseFloat(document.getElementById('cfg-split-dah').value) || 0
    };
    
    securePost(payload).then(res => {
        if(res.status === 'success') {
            showToast("Config tersimpan");
            closeConfigModal(); fetchConfig(); renderConfigTable();
        } else { showToast("Error: "+res.message); }
    });
}

function deleteConfig(id) {
    if(!confirm("Hapus item ini?")) return;
    securePost({ action: 'delete_config', item_id: id }).then(res => {
        if(res.status === 'success') { fetchConfig(); renderConfigTable(); }
    });
}
