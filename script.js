// --- KONFIGURASI ---
// GANTI INI DENGAN URL WEB APP ANDA (HASIL DEPLOY TERBARU)
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyo65StO07OygmbXGwFzoE-FVCGB9u-VfC9S9IL1uv78XxeeZpRMrLhdMlOLJqXiY2H/exec; 

let authToken = null;
let user = null;
let configData = [];
let allTransactions = [];
let isEditing = false;
let editId = null;
let targetsData = {}; 

// --- HELPER KEAMANAN ---
function secureFetch(payload) {
    if (!authToken) {
        alert("Sesi berakhir. Silakan login ulang.");
        location.reload();
        return Promise.reject("No Auth Token");
    }
    return fetch(SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify({ ...payload, token: authToken })
    });
}

// --- CHECK SESSION ON LOAD ---
window.addEventListener('DOMContentLoaded', () => {
    if (!authToken) {
        document.getElementById('login-view').style.display = 'flex';
        document.getElementById('app-view').style.display = 'none';
    }
});

// --- HELPERS ---
function parseRupiah(str) { if (!str) return 0; return parseFloat(String(str).replace(/\./g, '').replace(/,/g, '.')) || 0; }
function formatRupiah(num) { return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, "."); }

function formatInputOnKey(el) {
    let cursorPosition = el.selectionStart;
    let originalLength = el.value.length;
    let value = el.value.replace(/[^0-9]/g, '');
    el.value = value.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    if (value !== "") el.selectionStart = el.selectionEnd = cursorPosition + (el.value.length - originalLength);
    updateLiveTotal();
}

function updateLiveTotal() {
    let total = 0;
    document.querySelectorAll('.money-input').forEach(inp => total += parseRupiah(inp.value));
    document.getElementById('live-total').textContent = "Rp " + formatRupiah(total);
}

// --- LOGIN ---
function doLogin() {
    const pass = document.getElementById('login-pass').value.trim();
    const msg = document.getElementById('login-msg');
    if (!pass) return msg.innerText = "Masukkan password!";
    msg.innerText = "Memproses...";
    
    fetch(SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'login', password: pass })
    })
    .then(r => r.json())
    .then(res => {
        if (res.status === 'success') {
            user = res.user;
            authToken = res.token;
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
    
    // XSS Safe
    document.getElementById('u-name').textContent = user.nama_admin;
    document.getElementById('u-role').textContent = `${user.role} | ${user.kelompok || user.desa || ''}`;
    
    if(user.role === 'Daerah') {
        document.getElementById('filter-desa-container').style.display = 'block';
        document.getElementById('filter-desa-rekap-container').style.display = 'block';
        document.getElementById('nav-config').style.display = 'block';
        document.getElementById('nav-targets').style.display = 'block';
    }
    fetchConfig();
}

// --- FETCH CONFIG ---
function fetchConfig() {
    secureFetch({ action: 'get_config', kelompok: user.kelompok }) 
    .then(r => r.json())
    .then(res => {
        if(res.status !== 'success') {
            if(res.message.includes("Sesi")) location.reload();
            return;
        }
        configData = res.config.map(r => ({
            cat_code: r.cat_code, cat_title: r.cat_title, item_id: r.item_id, item_label: r.item_label,
            tipe: r.tipe, allowed_groups: r.allowed_groups,
            is_split: r.is_split, split_kel: r.split_kel, split_desa: r.split_desa, split_daerah: r.split_daerah
        }));
        targetsData = res.targets || {}; 
        buildForm();
        fetchData();
        if(user.role === 'Daerah') renderConfigTable();
    });
}

// --- BUILD FORM ---
function buildForm() {
    const tabsContainer = document.getElementById('dynamic-tabs');
    const itemsContainer = document.getElementById('dynamic-items');
    tabsContainer.innerHTML = ''; itemsContainer.innerHTML = '';

    const categories = {};
    configData.forEach(item => {
        if (!categories[item.cat_code]) categories[item.cat_code] = { title: item.cat_title, items: [] };
        categories[item.cat_code].items.push(item);
    });

    Object.keys(categories).forEach((code, index) => {
        const cat = categories[code];
        const btn = document.createElement('button');
        btn.className = `cat-tab ${index === 0 ? 'active' : ''}`;
        btn.id = `tab-btn-${code}`;
        btn.textContent = `${code}. ${cat.title}`;
        btn.onclick = () => switchCategory(code);
        tabsContainer.appendChild(btn);

        const pane = document.createElement('div');
        pane.id = `pane-${code}`;
        pane.className = `cat-pane ${index === 0 ? 'active' : ''}`;
        cat.items.forEach(item => {
            const row = document.createElement('div');
            row.className = 'item-row';
            
            const label = document.createElement('label');
            label.style.fontSize = '0.85rem';
            label.textContent = item.item_label;
            row.appendChild(label);

            const divCurrency = document.createElement('div');
            divCurrency.className = 'currency-wrap';
            divCurrency.innerHTML = '<span>Rp</span>';
            
            const inp = document.createElement('input');
            inp.type = 'text'; inp.inputMode = 'numeric';
            inp.className = 'form-control money-input';
            inp.dataset.id = item.item_id;
            inp.placeholder = '0';
            inp.oninput = function() { formatInputOnKey(this) };
            divCurrency.appendChild(inp);
            row.appendChild(divCurrency);

            const btnFill = document.createElement('button');
            btnFill.className = 'btn btn-sm';
            btnFill.style.background = '#e2e8f0';
            btnFill.textContent = '+50rb';
            btnFill.onclick = function() { quickFill(this, 50000) };
            row.appendChild(btnFill);
            pane.appendChild(row);

            // Split Preview
            if (item.is_split) {
                const previewDiv = document.createElement('div');
                previewDiv.className = 'split-preview';
                previewDiv.style.cssText = "font-size:0.7rem; color:#64748b; margin-top:4px; padding:4px; background:#f1f5f9; border-radius:4px;";
                previewDiv.innerHTML = '<i class="ph ph-calculator"></i> <b>Otomatis:</b> <span class="val-kel">0</span> (Ke F) | <span class="val-desa">0</span> (Ke E) | <span class="val-dah">0</span> (Tetap C)';
                
                inp.addEventListener('input', () => {
                    const total = parseRupiah(inp.value);
                    previewDiv.querySelector('.val-kel').textContent = formatRupiah((total * item.split_kel)/100);
                    previewDiv.querySelector('.val-desa').textContent = formatRupiah((total * item.split_desa)/100);
                    previewDiv.querySelector('.val-dah').textContent = formatRupiah((total * item.split_daerah)/100);
                });
                row.appendChild(previewDiv);
            }
        });
        itemsContainer.appendChild(pane);
    });
    document.getElementById('input-periode').value = new Date().toISOString().slice(0, 7);
    updateLiveTotal();
}

function switchCategory(code) {
    document.querySelectorAll('.cat-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.cat-pane').forEach(p => p.classList.remove('active'));
    document.getElementById(`tab-btn-${code}`).classList.add('active');
    document.getElementById(`pane-${code}`).classList.add('active');
}

// --- FETCH DATA ---
function fetchData() {
    secureFetch({ action: 'fetch_data' })
    .then(r => r.json())
    .then(res => {
        if(res.status !== 'success') {
            if(res.message.includes("Sesi")) location.reload();
            return;
        }
        allTransactions = res.data;
        updateDesaFilterOptions();
        renderDataTable();
        renderRecapTable();
    });
}

function updateDesaFilterOptions() {
    if(user.role !== 'Daerah') return;
    const desas = [...new Set(allTransactions.map(t => t.desa))].filter(Boolean);
    const selects = [document.getElementById('filter-data-desa'), document.getElementById('filter-rekap-desa')];
    selects.forEach(sel => {
        sel.innerHTML = '<option value="">Semua Desa</option>';
        desas.forEach(d => {
            const opt = document.createElement('option');
            opt.value = d; opt.textContent = d;
            sel.appendChild(opt);
        });
    });
}

// --- RENDER TABLES (XSS SAFE) ---
function renderDataTable() {
    const tbody = document.getElementById('table-body-data');
    const filterPeriode = document.getElementById('filter-data-periode').value;
    const filterDesa = document.getElementById('filter-data-desa').value;
    tbody.innerHTML = ''; let grandTotal = 0;

    const filtered = allTransactions.filter(t => {
        let match = true;
        if (user.role === 'Kelompok') match = t.kelompok === user.kelompok;
        else if (user.role === 'Desa') match = t.desa === user.desa;
        if (match && filterPeriode) match = (t.periode && String(t.periode).trim() === filterPeriode);
        if (match && user.role === 'Daerah' && filterDesa) match = (t.desa && String(t.desa).trim() === filterDesa);
        return match;
    });

    filtered.reverse().forEach(t => {
        let valObj = {}; try { valObj = JSON.parse(t.values_json); } catch(e){}
        const total = Object.values(valObj).reduce((a,b)=>a+(parseFloat(b)||0),0);
        grandTotal += total;
        const displayDate = t.timestamp ? String(t.timestamp).split(' ')[0] : '-';
        
        const tr = document.createElement('tr');
        // XSS Safe Rows
        tr.insertCell().textContent = displayDate;
        tr.insertCell().textContent = t.periode;
        tr.insertCell().textContent = t.nama_warga; // AMAN
        tr.insertCell().textContent = t.desa || '-';
        
        const tdTotal = tr.insertCell();
        tdTotal.style.cssText = "font-weight:bold; color:var(--accent)";
        tdTotal.textContent = 'Rp ' + total.toLocaleString('id-ID');

        const tdAction = tr.insertCell();
        tdAction.innerHTML = `<button class="btn btn-sm" onclick="editData('${t.id}')"><i class="ph ph-pencil-simple"></i></button> <button class="btn btn-sm btn-danger" onclick="deleteData('${t.id}')"><i class="ph ph-trash"></i></button>`;
        tbody.appendChild(tr);
    });
    document.getElementById('grand-total-data').textContent = `Rp ${grandTotal.toLocaleString('id-ID')}`;
}

function renderRecapTable() {
    const tbody = document.getElementById('table-body-rekap');
    const filterPeriode = document.getElementById('filter-rekap-periode').value;
    const filterDesa = document.getElementById('filter-rekap-desa').value; 
    tbody.innerHTML = '';
    if (!filterPeriode) return;

    const [y, m] = filterPeriode.split('-'); const mNum = parseInt(m);
    let targetKey = user.role==='Kelompok'?user.kelompok:(user.role==='Desa'?user.desa:(filterDesa||"Daerah"));
    const currentTargets = targetsData[targetKey] || {}; 
    const realisasiBln = {}; const akumulasiYTD = {}; let grandTotalBln = 0;

    allTransactions.forEach(t => {
        const [ty, tm] = (t.periode||"").split('-');
        if(ty == y && (user.role === 'Daerah' ? true : (user.role==='Desa'?t.desa===user.desa:t.kelompok===user.kelompok))) {
            let vals={}; try{vals=JSON.parse(t.values_json);}catch(e){}
            for(let k in vals){
                let v=parseFloat(vals[k])||0;
                if(parseInt(tm)<=mNum) akumulasiYTD[k]=(akumulasiYTD[k]||0)+v;
                if(parseInt(tm)===mNum) { realisasiBln[k]=(realisasiBln[k]||0)+v; grandTotalBln+=v; }
            }
        }
    });

    configData.forEach(c => {
        const bln = realisasiBln[c.item_id]||0;
        const ytd = akumulasiYTD[c.item_id]||0;
        const tVal = parseFloat(currentTargets[c.item_id])||0;
        const isBulanan = c.tipe==='Bulanan';
        const persen = tVal>0 ? Math.round((isBulanan?bln:ytd)/tVal*100) : 0;

        const tr = document.createElement('tr');
        const td1 = tr.insertCell(); td1.innerHTML = `<div style="font-weight:600">${c.item_label}</div><span style="font-size:0.6rem; color:#fff; background:${isBulanan?'#8b5cf6':'#3b82f6'}; padding:1px 6px; border-radius:10px; display:inline-block; margin-top:2px;">${c.tipe}</span>`;
        const td2 = tr.insertCell(); td2.style.cssText="text-align:right; font-weight:700; color:var(--accent)"; td2.textContent = 'Rp ' + bln.toLocaleString('id-ID');
        const td3 = tr.insertCell(); td3.style.cssText="text-align:right; color:var(--text-light)"; td3.textContent = isBulanan?'-':'Rp '+ytd.toLocaleString('id-ID');
        const td4 = tr.insertCell(); td4.style.textAlign='right'; td4.innerHTML = `<div style="font-weight:600">Rp ${tVal.toLocaleString('id-ID')}</div><div style="font-size:0.55rem; color:var(--text-light)">${isBulanan?'Target/Bln':'Target/Thn'}</div>`;
        const td5 = tr.insertCell(); td5.innerHTML = `<div style="font-size:0.7rem; display:flex; justify-content:space-between"><b>${persen}%</b><span style="opacity:0.7">${isBulanan?'REAL':'YTD'}</span></div><div class="progress-container" style="background:#e2e8f0; height:6px; border-radius:10px;"><div class="progress-bar" style="width:${persen>100?100:persen}%; height:100%; background:${persen>=100?'#22c55e':(isBulanan?'#a78bfa':'#3b82f6')};"></div></div>`;
        tbody.appendChild(tr);
    });
    document.getElementById('grand-total-rekap').textContent = `Rp ${grandTotalBln.toLocaleString('id-ID')}`;
}

// --- ACTIONS ---
function submitData() {
    const btn = document.getElementById('btn-submit');
    const periode = document.getElementById('input-periode').value;
    const nama = document.getElementById('input-nama').value;
    if(!periode || !nama) return showToast("Periode & Nama wajib diisi!");

    const values = {};
    let hasVal = false;
    document.querySelectorAll('.money-input').forEach(inp => {
        const v = parseRupiah(inp.value); 
        if(v > 0) { values[inp.dataset.id] = v; hasVal = true; }
    });
    if(!hasVal) return showToast("Isi minimal satu nominal!");

    btn.disabled = true; btn.innerHTML = `<i class="ph ph-spinner ph-spin"></i> Menyimpan...`;
    
    // Kirim raw values, server akan hitung split
    secureFetch({ action: isEditing ? 'update' : 'create', id: editId, periode, nama_warga: nama, values })
    .then(r => r.json())
    .then(res => {
        if(res.status === 'success') { showToast("Berhasil!"); resetForm(); fetchData(); if(isEditing) switchTab('data'); }
        else showToast("Gagal: " + res.message);
    })
    .catch(e => showToast("Error Koneksi"))
    .finally(() => { btn.disabled = false; btn.innerHTML = `<i class="ph ph-floppy-disk"></i> Simpan Data`; });
}

function editData(id) {
    const t = allTransactions.find(x => x.id == id);
    if(!t) return;
    isEditing = true; editId = id;
    document.getElementById('form-title').textContent = "Edit Data";
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

function resetForm() {
    isEditing = false; editId = null;
    document.getElementById('form-title').textContent = "Input Baru";
    document.getElementById('cancel-edit').style.display = 'none';
    document.getElementById('input-nama').value = '';
    document.querySelectorAll('.money-input').forEach(i => i.value = '');
    updateLiveTotal();
}

function switchTab(tabName) {
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.view-section').forEach(s => s.classList.remove('active'));
    document.getElementById(`nav-${tabName}`).classList.add('active');
    document.getElementById(`view-${tabName}`).classList.add('active');
    if(tabName === 'config' && user.role === 'Daerah') renderConfigTable();
    if(tabName === 'targets' && user.role === 'Daerah') renderTargetsTable();
}

function quickFill(btn, amount) {
    const inp = btn.parentElement.querySelector('input');
    inp.value = formatRupiah(parseRupiah(inp.value) + amount);
    updateLiveTotal();
}

function showToast(msg) {
    const t = document.getElementById('toast');
    t.textContent = msg; t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 3000);
}

function deleteData(id) {
    if(!confirm("Yakin hapus?")) return;
    secureFetch({ action: 'delete', id: id })
    .then(r => r.json())
    .then(res => { if(res.status==='success') { showToast("Terhapus"); fetchData(); } else showToast(res.message); });
}

// --- CONFIG CRUD ---
function renderConfigTable() {
    const tbody = document.getElementById('table-body-config');
    tbody.innerHTML = '';
    configData.forEach(i => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td><b>${i.item_id}</b></td><td>${i.cat_title}</td><td>${i.item_label}</td><td><span style="font-size:0.75rem; background:${i.tipe==='Bulanan'?'#a78bfa':'#3b82f6'}; color:white; padding:2px 6px; border-radius:4px;">${i.tipe}</span></td><td>${i.allowed_groups}</td><td><button class="btn btn-sm" onclick="editConfig('${i.item_id}')"><i class="ph ph-pencil-simple"></i></button> <button class="btn btn-sm btn-danger" onclick="deleteConfig('${i.item_id}')"><i class="ph ph-trash"></i></button></td>`;
        tbody.appendChild(tr);
    });
}
function openConfigModal(isEdit=false) {
    document.getElementById('config-modal').style.display = 'flex';
    document.getElementById('modal-title-config').textContent = isEdit ? "Edit Item" : "Tambah Item";
    if(!isEdit) {
        document.getElementById('cfg-old-id').value = '';
        ['cfg-item-id','cfg-cat-code','cfg-cat-title','cfg-item-label','cfg-groups'].forEach(id=>document.getElementById(id).value='');
        document.getElementById('cfg-tipe').value='Bulanan';
        document.getElementById('cfg-is-split').checked = false;
        document.getElementById('split-inputs').style.display = 'none';
    }
}
function closeConfigModal() { document.getElementById('config-modal').style.display = 'none'; }
function editConfig(id) {
    const i = configData.find(x=>x.item_id===id);
    openConfigModal(true);
    document.getElementById('cfg-old-id').value = i.item_id;
    document.getElementById('cfg-item-id').value = i.item_id;
    document.getElementById('cfg-cat-code').value = i.cat_code;
    document.getElementById('cfg-cat-title').value = i.cat_title;
    document.getElementById('cfg-item-label').value = i.item_label;
    document.getElementById('cfg-tipe').value = i.tipe;
    document.getElementById('cfg-groups').value = i.allowed_groups;
    document.getElementById('cfg-is-split').checked = i.is_split;
    document.getElementById('split-inputs').style.display = i.is_split ? 'grid' : 'none';
    document.getElementById('cfg-split-kel').value = i.split_kel;
    document.getElementById('cfg-split-desa').value = i.split_desa;
    document.getElementById('cfg-split-dah').value = i.split_daerah;
}
function deleteConfig(id) { if(!confirm("Hapus item?")) return; secureFetch({action:'delete_config',item_id:id}).then(r=>r.json()).then(res=>{if(res.status==='success'){showToast(res.message);fetchConfig();renderConfigTable();}}); }
function saveConfig() {
    const d = {
        old_id: document.getElementById('cfg-old-id').value,
        item_id: document.getElementById('cfg-item-id').value.trim(),
        cat_code: document.getElementById('cfg-cat-code').value.trim(),
        cat_title: document.getElementById('cfg-cat-title').value.trim(),
        item_label: document.getElementById('cfg-item-label').value.trim(),
        tipe: document.getElementById('cfg-tipe').value,
        allowed_groups: document.getElementById('cfg-groups').value.trim(),
        is_split: document.getElementById('cfg-is-split').checked,
        split_kel: parseFloat(document.getElementById('cfg-split-kel').value)||0,
        split_desa: parseFloat(document.getElementById('cfg-split-desa').value)||0,
        split_daerah: parseFloat(document.getElementById('cfg-split-dah').value)||0
    };
    if(!d.item_id || !d.cat_code || !d.cat_title || !d.item_label) return showToast("Lengkapi data!");
    if(d.is_split && (d.split_kel+d.split_desa+d.split_daerah!==100)) return showToast("Total % split harus 100%!");
    secureFetch({ action: d.old_id?'update_config':'create_config', ...d })
    .then(r=>r.json()).then(res=>{if(res.status==='success'){showToast(res.message);closeConfigModal();fetchConfig();renderConfigTable();}});
}

// --- TARGETS ---
function renderTargetsTable() {
    const thead = document.getElementById('thead-targets'), tbody = document.getElementById('tbody-targets');
    thead.innerHTML=''; tbody.innerHTML='';
    if(!targetsData || Object.keys(targetsData).length===0) return tbody.innerHTML='<tr><td colspan="10" style="text-align:center; padding:2rem;">Belum ada Target</td></tr>';
    
    let h = '<tr><th style="position:sticky; left:0; background:#f8fafc; z-index:10;">Identitas</th>';
    let items = configData.map(c=>c.item_id);
    if(items.length===0 && Object.keys(targetsData)[0]) items = Object.keys(targetsData[Object.keys(targetsData)[0]]);
    
    items.forEach(id => { const c=configData.find(x=>x.item_id===id); h+=`<th style="min-width:120px; text-align:center;">${c?c.item_label:id}</th>`; });
    thead.innerHTML = h+'</tr>';

    Object.keys(targetsData).forEach(iden => {
        const tr = document.createElement('tr');
        let row = `<td style="font-weight:bold; position:sticky; left:0; background:white; z-index:5; border-right:2px solid #ddd;">${iden}</td>`;
        items.forEach(iid => row += `<td style="text-align:center;"><input type="number" class="target-input" value="${targetsData[iden][iid]||0}" style="width:100px; text-align:center; padding:5px; border:1px solid #ddd; border-radius:4px;" onblur="updateTarget(this,'${iden}','${iid}')" onchange="this.style.background='#dcfce7'; setTimeout(()=>this.style.background='white',500)"></td>`);
        tr.innerHTML = row; tbody.appendChild(tr);
    });
}
function updateTarget(el, iden, iid) {
    el.style.background='#fef9c3';
    secureFetch({action:'update_target', identity:iden, item_id:iid, value:el.value})
    .then(r=>r.json()).then(res=>{if(res.status==='success'){el.style.background='#dcfce7'; setTimeout(()=>el.style.background='white',1000); targetsData[iden][iid]=parseFloat(el.value);}});
}
document.getElementById('cfg-is-split').addEventListener('change',function(){document.getElementById('split-inputs').style.display=this.checked?'grid':'none';});

// --- EXCEL DOWNLOAD (Client Side Logic - Safe) ---
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
        if(filterDesa) {
            namaDesaDisplay = filterDesa.toUpperCase();
        } else {
            namaDesaDisplay = ""; 
        }
    }

    // --- KONFIGURASI KATEGORI ---
    let allowedCategories = [];
    if (user.role === 'Daerah') {
        allowedCategories = ['A', 'B', 'C', 'D', 'E', 'F'];
    } else if (user.role === 'Desa') {
        allowedCategories = ['A', 'B', 'C', 'D', 'E', 'F'];
    } else if (user.role === 'Kelompok') {
        allowedCategories = ['A', 'B', 'C', 'D', 'E', 'F'];
    }

    // --- LOGIKA JUDUL ---
    let reportTitle = "";
    let reportSubtitle = "";
    let reportPeriod = "PERIODE " + namaPeriode;

    if (user.role === 'Kelompok') {
        reportTitle = "REKAPITULASI HIBAH BULANAN KELOMPOK KE DESA";
        reportSubtitle = `KELOMPOK ${namaKelompok} DESA ${namaDesaDisplay}`;
    } else if (user.role === 'Desa') {
        reportTitle = "REKAPITULASI HIBAH BULANAN DESA KE DAERAH";
        reportSubtitle = `DESA ${namaDesaDisplay}`;
    } else if (user.role === 'Daerah') {
        if (filterDesa) {
            reportTitle = "REKAPITULASI HIBAH BULANAN DESA KE DAERAH";
            reportSubtitle = `DESA ${namaDesaDisplay}`;
        } else {
            reportTitle = "REKAPITULASI HIBAH BULANAN DAERAH";
            reportSubtitle = ""; 
        }
    }

    // --- LOGIKA TARGET ---
    let targetKey;
    if (user.role === 'Kelompok') {
        targetKey = user.kelompok;
    } else if (user.role === 'Desa') {
        targetKey = user.desa;
    } else if (user.role === 'Daerah') {
        targetKey = filterDesa || "Daerah";
    } else {
        targetKey = "";
    }
    const currentTargets = targetsData[targetKey] || {};
    
    // --- PENGHITUNGAN DATA ---
    const categorizedData = {};
    configData.forEach(conf => {
        if (allowedCategories.length > 0 && !allowedCategories.includes(conf.cat_code.toUpperCase())) {
            return; 
        }

        if (!categorizedData[conf.cat_title]) categorizedData[conf.cat_title] = [];
        
        const realisasiBln = allTransactions.filter(t => {
            const tPeriode = String(t.periode).trim();
            let matchRole = true;

            if (user.role === 'Kelompok') {
                matchRole = t.kelompok === user.kelompok;
            } else if (user.role === 'Desa') {
                matchRole = t.desa === user.desa;
            } else if (user.role === 'Daerah') {
                if (filterDesa) {
                    matchRole = t.desa === filterDesa;
                }
            }

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

    // --- MEMBUAT EXCEL ---
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Laporan', {
        pageSetup: { paperSize: 9, orientation: 'portrait' }
    });

    worksheet.mergeCells('A1:E1');
    worksheet.mergeCells('A2:E2');
    worksheet.mergeCells('A3:E3');
    
    const titleRow = worksheet.getCell('A1');
    titleRow.value = reportTitle;
    titleRow.font = { bold: true, size: 14 };
    titleRow.alignment = { horizontal: 'center' };

    const subtitleRow = worksheet.getCell('A2');
    subtitleRow.value = reportSubtitle;
    subtitleRow.font = { bold: true, size: 11 };
    subtitleRow.alignment = { horizontal: 'center' };

    const periodRow = worksheet.getCell('A3');
    periodRow.value = reportPeriod;
    periodRow.font = { bold: true, size: 11 };
    periodRow.alignment = { horizontal: 'center' };

    // Header Tabel
    const headerRow = worksheet.addRow(['NO', 'URAIAN', 'TARGET', 'REALISASI', 'KETERANGAN']);
    headerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        cell.alignment = { horizontal: 'center' };
    });

    // Isi Data
    let counter = 1;
    let totalTarget = 0;
    let grandTotalRealisasi = 0;

    // Hitung Hak Khusus untuk Footer
    let totalHakKelompok = 0; // Kode F
    let totalHakDesa = 0;     // Kode E

    for (const [catTitle, items] of Object.entries(categorizedData)) {
        const cat = configData.find(c => c.cat_title === catTitle);
        const catCode = cat ? cat.cat_code : '';

        const catRow = worksheet.addRow(['', catTitle.toUpperCase(), '', '', '']);
        catRow.getCell(2).font = { bold: true };
        
        items.forEach(item => {
            totalTarget += item.target;
            grandTotalRealisasi += item.realisasi;

            // Cek Kode untuk menghitung Hak
            if (catCode === 'F') totalHakKelompok += item.realisasi;
            if (catCode === 'E') totalHakDesa += item.realisasi;

            const row = worksheet.addRow([counter++, item.label, item.target, item.realisasi, ""]);
            row.getCell(3).numFmt = '#,##0';
            row.getCell(4).numFmt = '#,##0';
        });
    }
    
    // Footer Total (Grand Total)
    const footerRow = worksheet.addRow(['', 'TOTAL KESELURUHAN', totalTarget, grandTotalRealisasi, '']);
    footerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    });
    footerRow.getCell(3).numFmt = '#,##0';
    footerRow.getCell(4).numFmt = '#,##0';

    // --- BARIS TOTAL YANG DISETOR ---
    let totalYangDisetor = 0;
    let labelSetoran = "";

    if (user.role === 'Kelompok') {
        totalYangDisetor = grandTotalRealisasi - totalHakKelompok;
        labelSetoran = "TOTAL YANG DISETOR KE DESA";
    } else if (user.role === 'Desa') {
        totalYangDisetor = grandTotalRealisasi - totalHakDesa - totalHakKelompok;
        labelSetoran = "TOTAL YANG DISETOR KE DAERAH";
    }

    if ((user.role === 'Kelompok' || user.role === 'Desa') && totalYangDisetor > 0) {
        const setoranRow = worksheet.addRow(['', labelSetoran, '', totalYangDisetor, '']);
        setoranRow.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        });
        setoranRow.getCell(3).numFmt = '#,##0';
        setoranRow.getCell(4).numFmt = '#,##0';
        setoranRow.getCell(4).alignment = { horizontal: 'center' };
    }

    // Styling Border
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber >= 5) {
            row.eachCell((cell) => {
                if (!cell.fill || cell.fill.fgColor.argb !== 'FF0F766E') {
                   cell.border = {
                        top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
                    };
                }
                if (cell.address.startsWith('A') || cell.address.startsWith('C') || cell.address.startsWith('D')) {
                    cell.alignment = { horizontal: 'center' };
                }
            });
        }
    });

    // Lebar Kolom
    worksheet.getColumn(1).width = 5;
    worksheet.getColumn(2).width = 45;
    worksheet.getColumn(3).width = 18;
    worksheet.getColumn(4).width = 18;
    worksheet.getColumn(5).width = 20;

    let fileName = "";
    if (user.role === 'Kelompok') {
        fileName = `Laporan_Kelompok_${namaKelompok}_${filterPeriodeRaw}.xlsx`;
    } else if (user.role === 'Desa') {
        fileName = `Laporan_Desa_${namaDesaDisplay}_${filterPeriodeRaw}.xlsx`;
    } else if (user.role === 'Daerah') {
        if (filterDesa) {
            fileName = `Laporan_Desa_${filterDesa}_${filterPeriodeRaw}.xlsx`;
        } else {
            fileName = `Laporan_Daerah_Keseluruhan_${filterPeriodeRaw}.xlsx`;
        }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    window.URL.revokeObjectURL(url);
    showToast("File Excel berhasil didownload!");
}// --- LOGIKA DOWNLOAD EXCEL (UPDATE 1) ---
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
        if(filterDesa) {
            namaDesaDisplay = filterDesa.toUpperCase();
        } else {
            namaDesaDisplay = ""; 
        }
    }

    // --- KONFIGURASI KATEGORI (DIUPDATE SESUAI REQUEST) ---
    let allowedCategories = [];
    if (user.role === 'Daerah') {
        // Daerah hanya melihat A, B, C, D (Sumber Utama)
        // E dan F tidak ditampilkan karena ini laporan penerimaan daerah
        allowedCategories = ['A', 'B', 'C', 'D'];
    } else if (user.role === 'Desa') {
        // Desa melihat A, B, C, D, E
        // F (Hak Kelompok) tidak ditampilkan
        allowedCategories = ['A', 'B', 'C', 'D', 'E'];
    } else if (user.role === 'Kelompok') {
        // Kelompok melihat semuanya A, B, C, D, E, F
        allowedCategories = ['A', 'B', 'C', 'D', 'E', 'F'];
    }

    // --- LOGIKA JUDUL ---
    let reportTitle = "";
    let reportSubtitle = "";
    let reportPeriod = "PERIODE " + namaPeriode;

    if (user.role === 'Kelompok') {
        reportTitle = "REKAPITULASI HIBAH BULANAN KELOMPOK KE DESA";
        reportSubtitle = `KELOMPOK ${namaKelompok} DESA ${namaDesaDisplay}`;
    } else if (user.role === 'Desa') {
        reportTitle = "REKAPITULASI HIBAH BULANAN DESA KE DAERAH";
        reportSubtitle = `DESA ${namaDesaDisplay}`;
    } else if (user.role === 'Daerah') {
        if (filterDesa) {
            reportTitle = "REKAPITULASI HIBAH BULANAN DESA KE DAERAH";
            reportSubtitle = `DESA ${namaDesaDisplay}`;
        } else {
            reportTitle = "REKAPITULASI HIBAH BULANAN DAERAH";
            reportSubtitle = ""; 
        }
    }

    // --- LOGIKA TARGET ---
    let targetKey;
    if (user.role === 'Kelompok') {
        targetKey = user.kelompok;
    } else if (user.role === 'Desa') {
        targetKey = user.desa;
    } else if (user.role === 'Daerah') {
        targetKey = filterDesa || "Daerah";
    } else {
        targetKey = "";
    }
    const currentTargets = targetsData[targetKey] || {};
    
    // --- PENGHITUNGAN DATA ---
    const categorizedData = {};
    configData.forEach(conf => {
        // Filter baris berdasarkan allowedCategories
        if (allowedCategories.length > 0 && !allowedCategories.includes(conf.cat_code.toUpperCase())) {
            return; 
        }

        if (!categorizedData[conf.cat_title]) categorizedData[conf.cat_title] = [];
        
        const realisasiBln = allTransactions.filter(t => {
            const tPeriode = String(t.periode).trim();
            let matchRole = true;

            if (user.role === 'Kelompok') {
                matchRole = t.kelompok === user.kelompok;
            } else if (user.role === 'Desa') {
                matchRole = t.desa === user.desa;
            } else if (user.role === 'Daerah') {
                if (filterDesa) {
                    matchRole = t.desa === filterDesa;
                }
            }

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

    // --- MEMBUAT EXCEL ---
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Laporan', {
        pageSetup: { paperSize: 9, orientation: 'portrait' }
    });

    worksheet.mergeCells('A1:E1');
    worksheet.mergeCells('A2:E2');
    worksheet.mergeCells('A3:E3');
    
    const titleRow = worksheet.getCell('A1');
    titleRow.value = reportTitle;
    titleRow.font = { bold: true, size: 14 };
    titleRow.alignment = { horizontal: 'center' };

    const subtitleRow = worksheet.getCell('A2');
    subtitleRow.value = reportSubtitle;
    subtitleRow.font = { bold: true, size: 11 };
    subtitleRow.alignment = { horizontal: 'center' };

    const periodRow = worksheet.getCell('A3');
    periodRow.value = reportPeriod;
    periodRow.font = { bold: true, size: 11 };
    periodRow.alignment = { horizontal: 'center' };

    // Header Tabel
    const headerRow = worksheet.addRow(['NO', 'URAIAN', 'TARGET', 'REALISASI', 'KETERANGAN']);
    headerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        cell.alignment = { horizontal: 'center' };
    });

    // Isi Data
    let counter = 1;
    let totalTarget = 0;
    let grandTotalRealisasi = 0;

    // Hitung Hak Khusus untuk Footer
    // Karena F disembunyikan untuk Desa, totalHakKelompok akan tetap 0 untuk user Desa
    let totalHakKelompok = 0; // Kode F
    let totalHakDesa = 0;     // Kode E

    for (const [catTitle, items] of Object.entries(categorizedData)) {
        const cat = configData.find(c => c.cat_title === catTitle);
        const catCode = cat ? cat.cat_code : '';

        const catRow = worksheet.addRow(['', catTitle.toUpperCase(), '', '', '']);
        catRow.getCell(2).font = { bold: true };
        
        items.forEach(item => {
            totalTarget += item.target;
            grandTotalRealisasi += item.realisasi;

            // Cek Kode untuk menghitung Hak (Hanya berjalan jika kategori ada di allowedCategories)
            if (catCode === 'F') totalHakKelompok += item.realisasi;
            if (catCode === 'E') totalHakDesa += item.realisasi;

            const row = worksheet.addRow([counter++, item.label, item.target, item.realisasi, ""]);
            row.getCell(3).numFmt = '#,##0';
            row.getCell(4).numFmt = '#,##0';
        });
    }
    
    // Footer Total (Grand Total)
    const footerRow = worksheet.addRow(['', 'TOTAL KESELURUHAN', totalTarget, grandTotalRealisasi, '']);
    footerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    });
    footerRow.getCell(3).numFmt = '#,##0';
    footerRow.getCell(4).numFmt = '#,##0';

    // --- BARIS TOTAL YANG DISETOR ---
    let totalYangDisetor = 0;
    let labelSetoran = "";

    if (user.role === 'Kelompok') {
        // Kelompok: Total - Hak Kelompok (F)
        totalYangDisetor = grandTotalRealisasi - totalHakKelompok;
        labelSetoran = "TOTAL YANG DISETOR KE DESA";
    } else if (user.role === 'Desa') {
        // Desa: Total - Hak Desa (E)
        // Catatan: Karena F tidak ditampilkan (allowedCategories), maka totalHakKelompok = 0
        totalYangDisetor = grandTotalRealisasi - totalHakDesa;
        labelSetoran = "TOTAL YANG DISETOR KE DAERAH";
    }
    // Daerah tidak perlu baris "Disetor" karena Daerah adalah tujuan akhir

    if ((user.role === 'Kelompok' || user.role === 'Desa') && totalYangDisetor > 0) {
        const setoranRow = worksheet.addRow(['', labelSetoran, '', totalYangDisetor, '']);
        setoranRow.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        });
        setoranRow.getCell(3).numFmt = '#,##0';
        setoranRow.getCell(4).numFmt = '#,##0';
        setoranRow.getCell(4).alignment = { horizontal: 'center' };
    }

    // Styling Border
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber >= 5) {
            row.eachCell((cell) => {
                if (!cell.fill || cell.fill.fgColor.argb !== 'FF0F766E') {
                   cell.border = {
                        top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
                    };
                }
                if (cell.address.startsWith('A') || cell.address.startsWith('C') || cell.address.startsWith('D')) {
                    cell.alignment = { horizontal: 'center' };
                }
            });
        }
    });

    // Lebar Kolom
    worksheet.getColumn(1).width = 5;
    worksheet.getColumn(2).width = 45;
    worksheet.getColumn(3).width = 18;
    worksheet.getColumn(4).width = 18;
    worksheet.getColumn(5).width = 20;

    let fileName = "";
    if (user.role === 'Kelompok') {
        fileName = `Laporan_Kelompok_${namaKelompok}_${filterPeriodeRaw}.xlsx`;
    } else if (user.role === 'Desa') {
        fileName = `Laporan_Desa_${namaDesaDisplay}_${filterPeriodeRaw}.xlsx`;
    } else if (user.role === 'Daerah') {
        if (filterDesa) {
            fileName = `Laporan_Desa_${filterDesa}_${filterPeriodeRaw}.xlsx`;
        } else {
            fileName = `Laporan_Daerah_Keseluruhan_${filterPeriodeRaw}.xlsx`;
        }
    }

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
