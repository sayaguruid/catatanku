// --- CONFIG & GLOBALS ---
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxGSOsV1HazZ-_4krq_ewUdfVLVYYLTU9kdkzX-rAzpNGGyr-RPhSU2YKtqzHJeUeJK/exec'; 
const SESSION_TIMEOUT = 30 * 60 * 1000; // 30 Menit ( dalam milidetik )
const STORAGE_KEY_USER = 'jurnalku_user';
const STORAGE_KEY_TIME = 'jurnalku_last_activity';

let user = null;
let db = { kategori: [], transaksi: [], tabungan: [], summary_balances: [] };
let sessionTimer;

// --- SESSION MANAGEMENT ---

// Cek sesi saat halaman dimuat
window.onload = function() {
    checkSession();
    // Pasang listener untuk aktivitas user (mouse, keyboard, klik)
    // agar tidak logout saat sedang bekerja
    window.addEventListener('mousemove', resetTimer);
    window.addEventListener('keypress', resetTimer);
    window.addEventListener('click', resetTimer);
};

function checkSession() {
    const storedUser = localStorage.getItem(STORAGE_KEY_USER);
    const storedTime = localStorage.getItem(STORAGE_KEY_TIME);
    const now = Date.now();

    if (storedUser && storedTime) {
        // Cek apakah waktu habis
        if (now - parseInt(storedTime) > SESSION_TIMEOUT) {
            doLogout(true); // Logout pakai karena timeout
        } else {
            // Masih login, restore data
            user = JSON.parse(storedUser);
            document.getElementById('login-view').classList.add('hidden');
            document.getElementById('app-view').style.display = 'block';
            initAppAfterLogin();
            resetTimer(); // Mulai timer ulang
        }
    } else {
        // Belum ada sesi, tampilkan login
        document.getElementById('login-view').classList.remove('hidden');
        document.getElementById('app-view').style.display = 'none';
    }
}

function saveSession() {
    localStorage.setItem(STORAGE_KEY_USER, JSON.stringify(user));
    localStorage.setItem(STORAGE_KEY_TIME, Date.now().toString());
}

function clearSession() {
    localStorage.removeItem(STORAGE_KEY_USER);
    localStorage.removeItem(STORAGE_KEY_TIME);
    user = null;
}

function resetTimer() {
    // Jika user belum login, jangan update timer
    if (!user) return;
    
    localStorage.setItem(STORAGE_KEY_TIME, Date.now().toString());
}

// Fungsi Logout Manual (Tombol Keluar)
function doLogout() {
    if(confirm('Apakah Anda yakin ingin keluar?')) {
        performLogout();
    }
}

// Fungsi Logout Otomatis (Timeout)
function doLogout(isTimeout = false) {
    if (isTimeout) {
        showToast("Sesi Anda telah berakhir. Silakan login kembali.", "error");
    }
    performLogout();
}

function performLogout() {
    clearSession();
    document.getElementById('app-view').style.display = 'none';
    document.getElementById('login-view').classList.remove('hidden');
    document.getElementById('login-pass').value = '';
    document.getElementById('btn-login').innerText = 'MASUK';
    document.getElementById('login-msg').innerText = '';
}

// --- LOGIC APLIKASI ---

// Helpers
const formatRp = (num) => "Rp " + Number(num).toLocaleString('id-ID');
const parseRupiah = (str) => {
    if(!str) return 0;
    return parseInt(str.replace(/[^0-9]/g, '')) || 0;
};

function formatRupiahInput(el) {
    let val = el.value.replace(/[^0-9]/g, '');
    if (val) {
        el.value = "Rp " + parseInt(val).toLocaleString('id-ID');
    } else {
        el.value = "";
    }
}

function showToast(msg, type='success') {
    const el = document.getElementById('toast');
    el.className = type; 
    el.innerHTML = `<i class="ph ${type==='success'?'ph-check-circle':'ph-warning-circle'}"></i> ${msg}`;
    el.classList.add('show');
    setTimeout(() => el.classList.remove('show'), 3000);
};

function getUserAkun() {
    if (!user) return '';
    if (user.role === 'Daerah') return 'Daerah';
    if (user.role === 'Desa') return user.desa;
    if (user.role === 'Kelompok') return user.kelompok;
    return '';
}

// Inisialisasi setelah login sukses
function initAppAfterLogin() {
    const namaAkun = getUserAkun();
    document.getElementById('header-detail').innerText = `Kelompok: ${user.kelompok || '-'}`;
    document.getElementById('my-akun-display').innerText = namaAkun;

    if(user.role === 'Desa' || user.role === 'Daerah') {
        document.getElementById('monitor-section').classList.remove('hidden');
    } else {
        document.getElementById('monitor-section').classList.add('hidden');
    }

    showToast(`Selamat Datang, ${user.nama_admin}`);
    initMonthFilter();
    loadData();
}

// --- LOGIN ---
function doLogin() {
    const pass = document.getElementById('login-pass').value;
    const msg = document.getElementById('login-msg');
    const btn = document.getElementById('btn-login');
    
    if(!pass) return msg.innerText = "Password tidak boleh kosong";

    btn.innerHTML = 'Memproses...';
    msg.innerText = "";
    msg.style.color = "var(--text-light)";

    fetch(SCRIPT_URL, { method:'POST', body:JSON.stringify({action:'login', password:pass}) })
    .then(r=>r.json()).then(res => {
        if(res.status==='success'){
            user = res.user;
            
            // Simpan sesi
            saveSession();

            document.getElementById('login-view').classList.add('hidden');
            document.getElementById('app-view').style.display = 'block';
            
            initAppAfterLogin();
        } else {
            msg.innerText = "Login Gagal: " + (res.message || "Password Salah");
            msg.style.color = "var(--danger)";
            btn.innerHTML = 'MASUK';
        }
    }).catch(err => {
        msg.innerText = "Koneksi Error. Coba lagi.";
        msg.style.color = "var(--danger)";
        btn.innerHTML = 'MASUK';
    });
}

// --- FILTER BULAN ---
function initMonthFilter() {
    const sel = document.getElementById('filter-bulan');
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    let options = `<option value="all">Semua Periode</option>`;
    for(let i=0; i<12; i++) {
        let d = new Date(year, month - 1 - i, 1);
        let yyyy = d.getFullYear();
        let mm = String(d.getMonth() + 1).padStart(2, '0');
        let val = `${yyyy}-${mm}`;
        let label = `Bulan ${mm} ${yyyy}`;
        options += `<option value="${val}">${label}</option>`;
    }
    sel.innerHTML = options;
    sel.value = `${year}-${String(month).padStart(2, '0')}`;
}

function updateFilter() { renderAll(); }

// --- DATA LOADING ---
function loadData() {
    fetch(SCRIPT_URL, { method:'POST', body:JSON.stringify({
        action:'get_data',
        role: user.role,
        desa: user.desa,
        kelompok: user.kelompok
    })})
    .then(r=>r.json())
    .then(res => {
        db = res;
        renderAll();
    });
}

function getMyCategories() {
    const myAkun = getUserAkun();
    return db.kategori.filter(k => {
        return k.akun === 'all' || k.akun === myAkun;
    });
}

// --- RENDER MONITORING ---
function renderMonitoring() {
    const grid = document.getElementById('monitor-grid');
    grid.innerHTML = '';
    
    if (db.summary_balances && db.summary_balances.length > 0) {
        db.summary_balances.forEach(item => {
            grid.innerHTML += `
            <div class="card mon">
                <div class="card-label">${item.nama} <span style="font-weight:400; text-transform:none;">(${item.role})</span></div>
                <div class="card-value">${formatRp(item.saldo)}</div>
            </div>`;
        });
    } else {
        grid.innerHTML = `<div style="grid-column: 1/-1; text-align:center; color:var(--text-light); font-size:0.8rem;">Tidak ada data anak/anak bawahan.</div>`;
    }
}

function renderAll() {
    renderDashboard();
    renderMonitoring(); 
    renderDropdowns();
    renderTransaksiTable();
    renderTabunganTable();
    renderKategoriList();
    document.getElementById('tr-tanggal').valueAsDate = new Date();
    document.getElementById('tb-tanggal').valueAsDate = new Date();
}

function renderDashboard() {
    const filterVal = document.getElementById('filter-bulan').value;
    document.getElementById('lbl-periode').innerText = filterVal === 'all' ? 'Semua' : filterVal;

    // 1. Hitung Data Utama (Total Berjalan)
    let sMasukMain=0, sKeluarMain=0;
    db.transaksi.forEach(t => {
        const nom = parseFloat(t.nominal);
        if (t.tipe === 'masuk') sMasukMain += nom;
        else sKeluarMain += nom;
    });

    // --- TAMBAHAN: Hitung Saldo Tabungan Total ---
    let sMasukTabungan=0, sKeluarTabungan=0;
    db.tabungan.forEach(t => {
        const nom = parseFloat(t.nominal);
        if (t.tipe === 'masuk') sMasukTabungan += nom;
        else sKeluarTabungan += nom;
    });

    document.getElementById('val-saldo-main').innerText = formatRp(sMasukMain - sKeluarMain);
    document.getElementById('val-saldo-tabungan').innerText = formatRp(sMasukTabungan - sKeluarTabungan); // Update kartu tabungan
    document.getElementById('val-masuk-main').innerText = formatRp(sMasukMain);
    document.getElementById('val-keluar-main').innerText = formatRp(sKeluarMain);

    // 2. Hitung Data Periode
    let sMasukPer=0, sKeluarPer=0;
    const dataPeriode = db.transaksi.filter(t => {
        return filterVal === 'all' || t.tanggal.startsWith(filterVal);
    });

    dataPeriode.forEach(t => {
        const nom = parseFloat(t.nominal);
        if (t.tipe === 'masuk') sMasukPer += nom;
        else sKeluarPer += nom;
    });

    document.getElementById('val-saldo-per').innerText = formatRp(sMasukPer - sKeluarPer);
    document.getElementById('val-masuk-per').innerText = formatRp(sMasukPer);
    document.getElementById('val-keluar-per').innerText = formatRp(sKeluarPer);
}
function renderDropdowns() {
    const myCats = getMyCategories();
    const selKas = document.getElementById('tr-kategori');
    selKas.innerHTML = '<option value="">-- Pilih Kategori Kas --</option>';
    myCats.filter(k => k.tipe === 'Kas').forEach(k => {
        const label = k.akun === 'all' ? k.nama_kategori : `${k.nama_kategori}`;
        selKas.innerHTML += `<option value="${k.id}">${label}</option>`;
    });

    const selTab = document.getElementById('tb-kategori');
    selTab.innerHTML = '<option value="">-- Pilih Jenis Tabungan --</option>';
    myCats.filter(k => k.tipe === 'Tabungan').forEach(k => {
        const label = k.akun === 'all' ? k.nama_kategori : `${k.nama_kategori}`;
        selTab.innerHTML += `<option value="${k.id}">${label}</option>`;
    });
}

function simpanTransaksi() {
    const payload = {
        action: 'create_transaksi',
        tanggal: document.getElementById('tr-tanggal').value,
        tipe: document.getElementById('tr-tipe').value,
        kategori_id: document.getElementById('tr-kategori').value,
        nominal: parseRupiah(document.getElementById('tr-nominal').value),
        keterangan: document.getElementById('tr-ket').value,
        nama_admin: user.nama_admin,
        role: user.role,
        desa: user.desa,
        kelompok: user.kelompok
    };
    if(!payload.kategori_id || !payload.nominal) return alert("Lengkapi data");

    fetch(SCRIPT_URL, { method:'POST', body:JSON.stringify(payload) })
    .then(r=>r.json()).then(res => {
        if(res.status === 'success') { showToast('Transaksi Disimpan'); document.getElementById('tr-nominal').value=''; document.getElementById('tr-ket').value=''; loadData(); }
    });
}

function renderTransaksiTable() {
    const tbody = document.querySelector('#table-transaksi tbody');
    tbody.innerHTML = '';
    const filterVal = document.getElementById('filter-bulan').value;
    
    const data = db.transaksi
        .filter(t => filterVal === 'all' || t.tanggal.startsWith(filterVal))
        .sort((a,b) => new Date(b.tanggal) - new Date(a.tanggal));

    data.forEach(t => {
        const isMasuk = t.tipe === 'masuk';
        const kat = getMyCategories().find(k => k.id === t.kategori_id);
        tbody.innerHTML += `<tr>
            <td>${t.tanggal}</td>
            <td><span class="badge ${isMasuk?'badge-kas':'badge-keluar'}" style="${!isMasuk?'background:#fee2e2;color:#991b1b':''}">${isMasuk?'+':'-'}</span></td>
            <td>${kat ? kat.nama_kategori : '?'}</td>
            <td>${t.keterangan}</td>
            <td style="text-align:right; font-weight:bold; color:${isMasuk?'var(--success)':'var(--danger)'}">${formatRp(t.nominal)}</td>
            <td><button class="btn btn-sm btn-danger" onclick="hapusTrans('${t.id}')"><i class="ph ph-trash"></i></button></td>
        </tr>`;
    });
}

function hapusTrans(id) {
    if(!confirm('Hapus?')) return;
    fetch(SCRIPT_URL, { method:'POST', body:JSON.stringify({action:'delete_transaksi', id:id}) })
    .then(r=>r.json()).then(res => { if(res.status==='success') loadData(); });
}

// --- TABUNGAN ---
function simpanTabungan() {
    const payload = {
        action: 'create_tabungan',
        tanggal: document.getElementById('tb-tanggal').value,
        tipe: document.getElementById('tb-tipe').value,
        kategori_id: document.getElementById('tb-kategori').value,
        nominal: parseRupiah(document.getElementById('tb-nominal').value),
        keterangan: document.getElementById('tb-ket').value, // TAMBAHAN KETERANGAN
        nama_admin: user.nama_admin,
        desa: user.desa,
        kelompok: user.kelompok
    };
    if(!payload.kategori_id || !payload.nominal) return alert("Lengkapi data");

    fetch(SCRIPT_URL, { method:'POST', body:JSON.stringify(payload) })
    .then(r=>r.json()).then(res => { 
        if(res.status==='success') { 
            showToast('Tabungan Disimpan'); 
            document.getElementById('tb-nominal').value=''; 
            document.getElementById('tb-ket').value=''; // RESET KETERANGAN
            loadData(); 
        }
    });
}

function renderTabunganTable() {
    const tbody = document.getElementById('tbody-tabungan');
    tbody.innerHTML = '';
    const data = db.tabungan.sort((a,b) => new Date(b.tanggal) - new Date(a.tanggal));
    data.forEach(t => {
        const isMasuk = t.tipe === 'masuk';
        const kat = getMyCategories().find(k => k.id === t.kategori_id);
        // UPDATE TABEL UNTUK MENAMPILKAN KETERANGAN
        tbody.innerHTML += `<tr>
            <td>${t.tanggal}</td>
            <td><span class="badge ${isMasuk?'badge-kas':'badge-keluar'}" style="${!isMasuk?'background:#fee2e2;color:#991b1b':''}">${isMasuk?'Setor':'Tarik'}</span></td>
            <td><span class="badge badge-tab">${kat ? kat.nama_kategori : '-'}</span></td>
            <td>${t.keterangan || '-'}</td>
            <td style="text-align:right; font-weight:bold; color:${isMasuk?'var(--success)':'var(--danger)'}">${formatRp(t.nominal)}</td>
            <td><button class="btn btn-sm btn-danger" onclick="hapusTrans('${t.id}')"><i class="ph ph-trash"></i></button></td>
        </tr>`;
    });
}

// --- KATEGORI ---
function switchTab(tab) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    event.currentTarget.classList.add('active');
    document.querySelectorAll('.view-section').forEach(s => s.classList.add('hidden'));
    document.getElementById('view-'+tab).classList.remove('hidden');
}

function addKategori() {
    const nama = document.getElementById('kat-nama').value;
    const tipe = document.getElementById('kat-tipe').value;
    if(!nama) return alert('Isi nama kategori');
    const myAkun = getUserAkun();
    const payload = {
        action: 'create_kategori', nama_kategori: nama, tipe: tipe, akun: myAkun, nama_admin: user.nama_admin
    };
    fetch(SCRIPT_URL, { method:'POST', body:JSON.stringify(payload) })
    .then(r=>r.json()).then(res => {
        if(res.status==='success') {
            document.getElementById('kat-nama').value = '';
            showToast('Kategori Ditambahkan');
            loadData();
        }
    });
}

function renderKategoriList() {
    const div = document.getElementById('list-kategori');
    div.innerHTML = '';
    const myCats = getMyCategories();
    if(myCats.length === 0) {
        div.innerHTML = `<div style="grid-column: 1/-1; text-align:center; padding: 2rem; color:var(--text-light); background:#f1f5f9; border-radius:8px;">
            <i class="ph ph-tag" style="font-size:2rem; margin-bottom:0.5rem;"></i><br>
            Belum ada kategori.
        </div>`;
        return;
    }
    myCats.forEach(k => {
        const isGlobal = k.akun === 'all';
        div.innerHTML += `<div style="background:#f8fafc; border:1px solid var(--border); padding:10px; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
            <div>
                <div style="font-weight:600; font-size:0.9rem;">${k.nama_kategori}</div>
                <div style="font-size:0.75rem; color:var(--text-light);">${k.tipe} • ${isGlobal ? '<b>Global</b>' : 'Sendiri'}</div>
            </div>
            ${!isGlobal ? `<button class="btn btn-sm btn-danger" onclick="hapusKat('${k.id}')"><i class="ph ph-trash"></i></button>` : ''}
        </div>`;
    });
}

function hapusKat(id) {
    if(!confirm("Hapus?")) return;
    fetch(SCRIPT_URL, { method:'POST', body:JSON.stringify({ action:'delete_kategori', id:id}) })
    .then(r=>r.json()).then(res => { if(res.status==='success') loadData(); });
}

// --- DOWNLOAD EXCEL ---
async function downloadExcel(tipe) {
    const wb = new ExcelJS.Workbook();
    const namaAkun = getUserAkun();
    const filterVal = document.getElementById('filter-bulan').value;

    let judulSheet = '';
    if (tipe === 'buku') judulSheet = 'Buku Kas';
    if (tipe === 'laporan') judulSheet = 'Laporan Kas';
    if (tipe === 'laporan-tabungan') judulSheet = 'Laporan Tabungan';

    // ================= SALDO AWAL =================
    let saldoAwal = 0;
    if (tipe === 'laporan' && filterVal !== 'all') {
        db.transaksi.forEach(t => {
            if (t.tanggal < filterVal) {
                const n = Number(t.nominal);
                saldoAwal += (t.tipe === 'masuk') ? n : -n;
            }
        });
    }

    const ws = wb.addWorksheet(judulSheet);

    // ================= HELPER =================
    function styleTitle(row, colEnd, text) {
        ws.mergeCells(`A${row}:${colEnd}${row}`);
        const c = ws.getCell(`A${row}`);
        c.value = text;
        c.font = { bold: true, size: 14 };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
    }

    function styleSubTitle(row, colEnd, text) {
        ws.mergeCells(`A${row}:${colEnd}${row}`);
        const c = ws.getCell(`A${row}`);
        c.value = text;
        c.font = { italic: true, size: 10 };
        c.alignment = { horizontal: 'center' };
    }

    function styleBorder(row, colStart, colEnd) {
        for (let col = colStart; col <= colEnd; col++) {
            ws.getRow(row).getCell(col).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        }
    }

    // =====================================================
    // ================= LAPORAN KAS =======================
    // =====================================================
    if (tipe === 'laporan') {
        const dataScope = db.transaksi.filter(
            t => filterVal === 'all' || t.tanggal.startsWith(filterVal)
        );

        styleTitle(1, 'B', 'LAPORAN PENERIMAAN DAN PENGELUARAN KAS');
        styleSubTitle(2, 'B', `Kelompok : ${namaAkun}`);
        styleSubTitle(3, 'B', `Periode  : ${filterVal === 'all' ? 'Semua' : filterVal}`);

        let r = 5;

        ws.addRow(['SALDO AWAL', saldoAwal]);
        ws.getRow(r).font = { bold: true };
        ws.getCell(`B${r}`).numFmt = '"Rp "#,##0';
        ws.getCell(`B${r}`).alignment = { horizontal: 'right' };
        ws.getCell(`B${r}`).fill = {
            type:'pattern', pattern:'solid', fgColor:{argb:'FFE2E8F0'}
        };
        styleBorder(r,1,2);

        let totalMasuk = 0;
        let totalKeluar = 0;

        // ----- PENERIMAAN -----
        r += 2;
        ws.addRow(['PENERIMAAN', 'NOMINAL']);
        ws.getRow(r).font = { bold: true };
        ws.getCell(`A${r}`).fill = {
            type:'pattern', pattern:'solid', fgColor:{argb:'FF10B981'}
        };
        ws.getCell(`B${r}`).fill = {
            type:'pattern', pattern:'solid', fgColor:{argb:'FF10B981'}
        };
        ws.getRow(r).alignment = { horizontal: 'center' };
        styleBorder(r,1,2);

        const inc = {};
        dataScope.filter(t => t.tipe === 'masuk').forEach(t => {
            const kat = getMyCategories().find(k => k.id === t.kategori_id);
            const nama = kat ? kat.nama_kategori : 'Lain-lain';
            inc[nama] = (inc[nama] || 0) + Number(t.nominal);
        });

        for (const k in inc) {
            r++;
            ws.addRow([k, inc[k]]);
            ws.getCell(`B${r}`).numFmt = '"Rp "#,##0';
            totalMasuk += inc[k];
            styleBorder(r, 1, 2);
        }

        r++;
        ws.addRow(['JUMLAH PENERIMAAN', totalMasuk]);
        ws.getRow(r).font = { bold: true };

        ws.getCell(`A${r}`).fill = {
            type:'pattern', pattern:'solid', fgColor:{argb:'FFD1FAE5'}
        };
        ws.getCell(`B${r}`).fill = {
            type:'pattern', pattern:'solid', fgColor:{argb:'FFD1FAE5'}
        };

        ws.getCell(`B${r}`).numFmt = '"Rp "#,##0';
        ws.getCell(`B${r}`).alignment = { horizontal: 'right' };
        styleBorder(r,1,2);

        // ----- PENGELUARAN -----
        r += 2;
        ws.addRow(['PENGELUARAN', 'NOMINAL']);
        ws.getRow(r).font = { bold: true };
        ws.getRow(r).alignment = { horizontal: 'center' };
        styleBorder(r, 1, 2);

        const exp = {};
        dataScope.filter(t => t.tipe === 'keluar').forEach(t => {
            const kat = getMyCategories().find(k => k.id === t.kategori_id);
            const nama = kat ? kat.nama_kategori : 'Lain-lain';
            exp[nama] = (exp[nama] || 0) + Number(t.nominal);
        });

        for (const k in exp) {
            r++;
            ws.addRow([k, exp[k]]);
            ws.getCell(`B${r}`).numFmt = '"Rp "#,##0';
            totalKeluar += exp[k];
            styleBorder(r, 1, 2);
        }

        r++;
        ws.addRow(['JUMLAH PENGELUARAN', totalKeluar]);
        ws.getRow(r).font = { bold: true };

        ws.getCell(`A${r}`).fill = {
            type:'pattern', pattern:'solid', fgColor:{argb:'FFFEE2E2'}
        };
        ws.getCell(`B${r}`).fill = {
            type:'pattern', pattern:'solid', fgColor:{argb:'FFFEE2E2'}
        };

        ws.getCell(`B${r}`).numFmt = '"Rp "#,##0';
        ws.getCell(`B${r}`).alignment = { horizontal: 'right' };
        styleBorder(r,1,2);

        // ----- SALDO AKHIR -----
        ws.addRow(['SALDO AKHIR', saldoAwal + totalMasuk - totalKeluar]);
        ws.getRow(r).font = { bold: true };
        ws.getCell(`B${r}`).numFmt = '"Rp "#,##0';
        ws.getCell(`B${r}`).alignment = { horizontal: 'right' };
        ws.getCell(`B${r}`).fill = {
            type:'pattern', pattern:'solid', fgColor:{argb:'FFDBEAFE'}
        };
        styleBorder(r,1,2);

        ws.getColumn(1).width = 32;
        ws.getColumn(2).width = 20;
        ws.getColumn(2).alignment = { horizontal: 'right' };
    }

    // =====================================================
    // ================= BUKU KAS ==========================
    // =====================================================
    else if (tipe === 'buku') {
        const dataScope = db.transaksi
            .filter(t => filterVal === 'all' || t.tanggal.startsWith(filterVal))
            .sort((a,b) => new Date(a.tanggal) - new Date(b.tanggal));

        styleTitle(1, 'G', 'BUKU KAS UMUM');
        styleSubTitle(2, 'G', `Kelompok : ${namaAkun} | Periode : ${filterVal}`);

        ws.addRow(['NO','TANGGAL','URAIAN','KATEGORI','MASUK','KELUAR','SALDO']);
        const hr = ws.getRow(3);

        for (let c=1;c<=7;c++) {
            const cell = hr.getCell(c);
            cell.font = { bold:true, color:{argb:'FFFFFFFF'} };
            cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF1F2937'} };
            cell.alignment = { horizontal:'center' };
            cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
        }

        let r = 3;

        // ===== HITUNG SALDO AWAL =====
        let saldoAwal = 0;
        if (filterVal !== 'all') {
            db.transaksi.forEach(t => {
                if (t.tanggal < filterVal) {
                    const n = Number(t.nominal);
                    saldoAwal += (t.tipe === 'masuk') ? n : -n;
                }
            });
        }

        let saldo = saldoAwal;
        let no = 1;

        // ===== TAMPILKAN SALDO AWAL =====
        if (saldoAwal !== 0) {
            r++;
            ws.addRow(['','', 'Saldo Awal Periode','', '', '', saldoAwal]);
            ws.getCell(`G${r}`).numFmt = '"Rp "#,##0';
            ws.getRow(r).font = { italic:true };
            styleBorder(r,1,7);
        }

        // ===== DAFTAR TRANSAKSI =====
        dataScope.forEach(t => {
            r++;
            const n = Number(t.nominal);
            const masuk = t.tipe === 'masuk';
            saldo = masuk ? saldo + n : saldo - n;

            ws.addRow([
                no++,
                t.tanggal,
                t.keterangan,
                getMyCategories().find(k=>k.id===t.kategori_id)?.nama_kategori || '-',
                masuk ? n : '',
                !masuk ? n : '',
                saldo
            ]);

            ws.getCell(`E${r}`).numFmt = '"Rp "#,##0';
            ws.getCell(`F${r}`).numFmt = '"Rp "#,##0';
            ws.getCell(`G${r}`).numFmt = '"Rp "#,##0';
            styleBorder(r,1,7);
        });

        ws.columns = [
            { width:5 }, { width:11 }, { width:26 },
            { width:18 }, { width:14 }, { width:14 }, { width:16 }
        ];
    }

    // =====================================================
    // ================= TABUNGAN (PER KATEGORI) ============
    // =====================================================
    else if (tipe === 'laporan-tabungan') {
        styleTitle(1, 'E', 'LAPORAN TABUNGAN PER KATEGORI');
        styleSubTitle(2, 'E', `Kelompok : ${namaAkun} | Periode : ${filterVal === 'all' ? 'Semua' : filterVal}`);

        // 1. Kelompokkan Data berdasarkan Kategori
        const groupedData = {};
        db.tabungan.forEach(t => {
            const catId = t.kategori_id;
            if (!groupedData[catId]) groupedData[catId] = [];
            groupedData[catId].push(t);
        });

        let currentRow = 4; // Mulai menulis data dari baris ke-4

        // 2. Loop setiap kategori dan buat tabelnya
        Object.keys(groupedData).forEach(catId => {
            const kategori = getMyCategories().find(k => k.id === catId);
            const namaKategori = kategori ? kategori.nama_kategori : 'Tanpa Kategori';
            const transaksi = groupedData[catId].sort((a,b) => new Date(a.tanggal) - new Date(b.tanggal));

            // --- Header Kategori ---
            ws.addRow([`KATEGORI: ${namaKategori}`]);
            const headerRow = ws.getRow(currentRow);
            headerRow.font = { bold: true, size: 12, color: { argb:'FF7C3AED' } };
            headerRow.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF3E8FF'} };
            ws.mergeCells(`A${currentRow}:E${currentRow}`);
            currentRow++;

            // --- Header Tabel ---
            ws.addRow(['TANGGAL','KETERANGAN','MASUK (+)','KELUAR (-)','SALDO']);
            const tableHeader = ws.getRow(currentRow);
            for (let c=1;c<=5;c++) {
                const cell = tableHeader.getCell(c);
                cell.font = { bold:true, color:{argb:'FFFFFFFF'} };
                cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF7C3AED'} };
                cell.alignment = { horizontal:'center' };
                cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
            }
            currentRow++;

            // --- Isi Data Transaksi ---
            let saldoKategori = 0;
            transaksi.forEach(t => {
                const n = Number(t.nominal);
                const ket = t.keterangan || (kategori ? kategori.nama_kategori : '-');
                saldoKategori = t.tipe === 'masuk' ? saldoKategori + n : saldoKategori - n;

                ws.addRow([
                    t.tanggal,
                    ket,
                    t.tipe === 'masuk' ? n : '',
                    t.tipe === 'keluar' ? n : '',
                    saldoKategori
                ]);

                const dataRow = ws.getRow(currentRow);
                ws.getCell(`C${currentRow}`).numFmt = '"Rp "#,##0';
                ws.getCell(`D${currentRow}`).numFmt = '"Rp "#,##0';
                ws.getCell(`E${currentRow}`).numFmt = '"Rp "#,##0';
                
                // Border sel
                for (let c=1;c<=5;c++) {
                    dataRow.getCell(c).border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
                }
                currentRow++;
            });

            // --- Total Kategori ---
            ws.addRow(['', 'SALDO AKHIR ' + namaKategori.toUpperCase(), '', '', saldoKategori]);
            const totalRow = ws.getRow(currentRow);
            totalRow.font = { bold: true };
            ws.getCell(`E${currentRow}`).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFD1FAE5'} };
            ws.getCell(`E${currentRow}`).numFmt = '"Rp "#,##0';
            ws.getCell(`E${currentRow}`).alignment = { horizontal: 'right' };
            for (let c=1;c<=5;c++) {
                totalRow.getCell(c).border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
            }
            
            currentRow += 2; // Beri jarak antar kategori
        });

        ws.columns = [
            { width:12 }, { width:35 }, { width:14 }, { width:14 }, { width:18 }
        ];
    }
    // ================= SAVE FILE =================
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${judulSheet}_${namaAkun}_${filterVal !== 'all' ? filterVal : 'All'}_${new Date().toISOString().slice(0,10)}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
}