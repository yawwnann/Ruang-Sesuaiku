// Alamat server backend Flask kita
const API_URL = "http://127.0.0.1:5000/api/cek_kepatuhan";
const EXPORT_API_URL = "http://127.0.0.1:5000/api/export_word";

// Opsi Zona
const ZONA_OPTIONS = {
    "A": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "B": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)"},
    "C": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "D": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "E": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "F": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "G": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "H": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "I": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "J": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "K": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "L": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "M": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" },
    "N": { "PS": "Perlindungan Setempat (PS)", "RTH": "Ruang Terbuka Hijau (RTH)", "CB": "Cagar Budaya (CB)", "BJ": "Badan Jalan (BJ)", "R": "Perumahan (R)", "K": "Perdagangan dan Jasa (K)", "SPU": "Sarana Pelayanan Umum (SPU)", "HanKam": "Pertahanan dan Keamanan", "Trans": "Transportasi", "C": "Campuran (C)" }
};

// --- FUNGSI 1: updateZonaOptions ---
function updateZonaOptions() {
    const swp = document.getElementById('swp').value;
    const zonaSelect = document.getElementById('zona');
    zonaSelect.innerHTML = '<option value="">-- Pilih Zona --</option>'; 
    document.getElementById('subZonaContainer').style.display = 'none'; 
    document.getElementById('subZona').innerHTML = '<option value="">-- Pilih Sub-Zona --</option>'; 
    document.getElementById('subZona').removeAttribute('required'); 
    if (swp && ZONA_OPTIONS[swp]) {
        for (const [key, value] of Object.entries(ZONA_OPTIONS[swp])) {
            const option = document.createElement('option');
            option.value = key;
            option.textContent = value; 
            zonaSelect.appendChild(option);
        }
        zonaSelect.disabled = false; 
    } else {
        zonaSelect.innerHTML = '<option value="">-- Pilih SWP Dahulu --</option>';
        zonaSelect.disabled = true; 
    }
}

// --- FUNGSI 2: tampilkanSubZona ---
async function tampilkanSubZona() {
    const swp = document.getElementById('swp').value;
    const zona = document.getElementById('zona').value;
    const subZonaContainer = document.getElementById('subZonaContainer');
    const subZonaSelect = document.getElementById('subZona');
    subZonaSelect.innerHTML = '<option value="">-- Memuat... --</option>';
    subZonaContainer.style.display = 'none';
    subZonaSelect.removeAttribute('required');
    if (!swp || !zona) {
        subZonaSelect.innerHTML = '<option value="">-- Pilih Zona Dahulu --</option>';
        return;
    }
    try {
        const response = await fetch(`http://127.0.0.1:5000/api/get_subzonas?swp=${swp}&zona=${zona}`);
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Gagal mengambil data sub-zona');
        }
        const subzonaList = await response.json(); 
        if (subzonaList && Array.isArray(subzonaList) && subzonaList.length > 0) {
            subZonaSelect.innerHTML = '<option value="">-- Pilih Sub-Zona --</option>'; 
            const filteredList = subzonaList.length > 1 ? subzonaList.filter(sz => !sz.toLowerCase().includes('default')) : subzonaList;
            filteredList.forEach(namaSubZona => {
                const option = document.createElement('option');
                option.value = namaSubZona;
                option.textContent = namaSubZona;
                subZonaSelect.appendChild(option);
            });
            if (filteredList.length > 0){
                subZonaContainer.style.display = 'block';
                subZonaSelect.setAttribute('required', 'required'); 
            } else {
                subZonaContainer.style.display = 'none';
                subZonaSelect.removeAttribute('required');
                subZonaSelect.innerHTML = ''; 
            }
        } else {
             subZonaContainer.style.display = 'none';
             subZonaSelect.removeAttribute('required');
             subZonaSelect.innerHTML = '';
        }
    } catch (error) {
        console.error("Error fetching sub-zona:", error);
        alert(`Gagal mengambil daftar sub-zona: ${error.message}. Pastikan server app.py berjalan.`);
        subZonaSelect.innerHTML = '<option value="">-- Error --</option>';
    }
}

// --- FUNGSI 3: cekKepatuhan ---
async function cekKepatuhan(event) {
    console.log("Tombol 'Periksa Kepatuhan' diklik.");
    event.preventDefault(); 

    const submitButton = document.querySelector('#step-2 button[type="submit"]');
    if (submitButton) { 
        submitButton.textContent = 'MEMERIKSA...';
        submitButton.disabled = true; 
    }

    // --- VALIDASI ---
    const step1LokasiInputs = document.querySelectorAll('#step-1 input[required], #step-1 select[required]');
    const step2Inputs = document.querySelectorAll('#step-2 input[required], #step-2 select[required]');
    let isFormValid = true;
    let firstInvalidElement = null;

    [...step1LokasiInputs, ...step2Inputs].forEach(input => {
        if (!input.checkValidity()) {
            isFormValid = false;
            if (!firstInvalidElement) firstInvalidElement = input; 
        }
    });
    
    if (!isFormValid) {
        console.log("Validasi Form Gagal.");
        if (firstInvalidElement) {
             firstInvalidElement.reportValidity(); 
            if ([...step1LokasiInputs].includes(firstInvalidElement)) {
                alert("Harap lengkapi semua data yang wajib diisi pada Step 1 (Data Lokasi).");
            }
        }
        if (submitButton) {
            submitButton.textContent = 'Periksa Kepatuhan';
            submitButton.disabled = false;
        }
        return; 
    }
    console.log("Validasi Form Lolos.");

    // --- PENGUMPULAN DATA ---
    const formData = {
        swpPilihan: document.getElementById('swp').value,
        zonaPilihan: document.getElementById('zona').value,
        subZonaPilihan: document.getElementById('subZona').value,
        luasPersilPilihan: document.getElementById('luasPersil').value,
        luasPersilInput: document.getElementById('luasPersilInput').value,
        kdbDiusulkan: document.getElementById('kdbDiusulkan').value,
        luasBangunanDiusulkan: document.getElementById('luasBangunan').value,
        luasLantaiDiusulkan: document.getElementById('luasLantai').value,
        ketinggian: document.getElementById('ketinggian').value,
        kdhDiusulkan: document.getElementById('kdhDiusulkan').value,
        isSubZonaRequired: document.getElementById('subZonaContainer').style.display !== 'none' && document.getElementById('subZona').hasAttribute('required'),
        gsbDiusulkan: document.getElementById('gsbDiusulkan').value 
    };

    console.log("Data yang akan dikirim ke API Cek Kepatuhan:", formData);

    try {
        const response = await fetch(API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(formData),
        });

        const data = await response.json(); 
        
        if (!response.ok) {
            throw new Error(data.error || `Server Error: ${response.statusText}`);
        }

        console.log("API sukses, menampilkan hasil...");

        // ============================================
        // [UPDATE] SIMPAN KE RIWAYAT (DATA LENGKAP)
        // ============================================
        if (typeof saveToHistory === 'function') {
            saveToHistory({
                swp: formData.swpPilihan,
                zona: formData.zonaPilihan,
                luas: formData.luasPersilInput,
                kdb: formData.kdbDiusulkan,
                status: data.isPatuh ? 'Sesuai' : 'Tidak Sesuai',
                inputData: formData,  // Simpan data input user
                resultData: data      // Simpan hasil API
            });
        }
        // ============================================

        tampilkanHasil( 
            data.isPatuh,
            data.detailPemeriksaan, 
            data.kdbAktual,
            data.klbAktual,
            data.kdbMaks,
            data.klbMaks,
            data.kdhMin
        );

        if (typeof window.showStep3 === 'function') {
            window.showStep3();
        } else {
            document.getElementById('hasil').style.display = 'block';
        }

    } catch (error) { 
        console.error('Error saat menghubungi atau memproses API:', error);
        alert(`Terjadi Kesalahan: ${error.message}. Cek Console (F12) untuk detail.`);
        const hasilDiv = document.getElementById('hasil');
        hasilDiv.className = 'p-6 rounded-lg border-2 border-red-400 bg-red-50 text-red-700';
        hasilDiv.innerHTML = `<h2><i data-feather="alert-triangle" class="inline mr-2"></i>Error</h2><p>${error.message}</p><p>Pastikan server backend (app.py) berjalan dan tidak ada error di terminalnya.</p>`;
        if (typeof feather !== 'undefined') feather.replace(); 
        
        if (typeof window.showStep3 === 'function') {
             window.showStep3();
        } else {
             document.getElementById('hasil').style.display = 'block';
        }
    } finally { 
        if (submitButton) {
            submitButton.textContent = 'Periksa Kepatuhan';
            submitButton.disabled = false;
        }
    }
}

// FUNGSI 4: tampilkanHasil (REVISI: Perhitungan Kekurangan KDH Sesuai Catatan)
function tampilkanHasil(status, detailPemeriksaan, kdbAktual, klbAktual, kdbMaks, klbMaks, kdhMin) {
    const hasilDiv = document.getElementById('hasil');

    // --- 1. Header Status Utama ---
    let headerColor = status ? 'text-green-600' : 'text-red-600';
    let headerIcon = status ? 'check-circle' : 'alert-circle';
    let headerText = status ? 'PATUH' : 'TIDAK PATUH';
    let headerDesc = status 
        ? 'Selamat! Rencana bangunan Anda telah memenuhi seluruh aturan tata ruang.'
        : 'Mohon maaf, rencana bangunan Anda belum memenuhi beberapa aturan. Silakan perbaiki poin yang bertanda merah.';

    let htmlContent = `
        <div class="text-center mb-8">
            <div class="inline-flex items-center justify-center w-16 h-16 rounded-full ${status ? 'bg-green-100 text-green-600' : 'bg-red-100 text-red-600'} mb-4">
                <i data-feather="${headerIcon}" class="w-8 h-8"></i>
            </div>
            <h2 class="text-2xl font-bold ${headerColor} mb-2">${headerText}</h2>
            <p class="text-gray-600 text-sm max-w-lg mx-auto">${headerDesc}</p>
        </div>
    `;

    // --- 2. Rincian Variabel (Kartu Ringkasan) ---
    htmlContent += `<h3 class="text-lg font-bold text-gray-800 mb-4 uppercase tracking-wide text-sm">Ringkasan Pemeriksaan</h3>`;
    htmlContent += `<div class="space-y-3 mb-8">`;

    detailPemeriksaan.forEach(item => {
        const isPass = item.status === 'pass';
        const cardClass = isPass 
            ? "flex items-start p-4 rounded-xl border border-green-200 bg-green-50/50" 
            : "flex items-start p-4 rounded-xl border border-red-200 bg-red-50/50";
        const iconClass = isPass ? "text-green-600" : "text-red-600";
        const iconName = isPass ? "check" : "x";
        const textClass = isPass ? "text-green-900" : "text-red-900";

        htmlContent += `
        <div class="${cardClass} transition-all hover:shadow-sm">
            <div class="flex-shrink-0 mr-3 mt-0.5">
                <div class="w-5 h-5 rounded-full flex items-center justify-center ${isPass ? 'bg-green-200' : 'bg-red-200'}">
                    <i data-feather="${iconName}" class="w-3 h-3 ${iconClass} stroke-3"></i>
                </div>
            </div>
            <div class="${textClass} text-sm font-medium leading-relaxed">
                ${item.message}
            </div>
        </div>`;
    });
    htmlContent += `</div>`;

    // --- 3. Tabel Angka Detail (DENGAN LOGIKA KEKURANGAN KDH) ---
    
    // Ambil nilai input
    const luasPersil = parseFloat(document.getElementById('luasPersilInput').value);
    
    // Parsing Input
    const kdbInput = parseFloat(document.getElementById('kdbDiusulkan').value).toFixed(2);
    const luasBangunanInput = parseFloat(document.getElementById('luasBangunan').value).toFixed(2);
    const luasLantaiInput = parseFloat(document.getElementById('luasLantai').value).toFixed(2);
    const kdhInput = parseFloat(document.getElementById('kdhDiusulkan').value); // Raw number
    const ketinggianInput = parseFloat(document.getElementById('ketinggian').value).toFixed(2);

    // Hitung Batas (Limit)
    const maxLuasBangunan = (luasPersil * (kdbMaks / 100)).toFixed(2);
    const maxLuasLantai = (luasPersil * klbMaks).toFixed(2);

    // --- LOGIKA KHUSUS KDH (MENAMPILKAN KEKURANGAN) ---
    let kdhDisplay = "";
    // Cek apakah KDH Input kurang dari KDH Minimal
    if (kdhInput < kdhMin) {
        const selisihPersen = (kdhMin - kdhInput).toFixed(2); // Contoh: 3.2%
        const selisihLuas = (luasPersil * (selisihPersen / 100)).toFixed(2); // Contoh: 83.1 m2
        
        // Format Teks Merah: "6.8% kurang 3.2% atau sebesar 83.1 m²"
        kdhDisplay = `<span class="font-bold text-red-600">${kdhInput}%</span> <span class="text-red-500 text-xs block mt-1">kurang ${selisihPersen}% atau sebesar ${selisihLuas} m²</span>`;
    } else {
        // Jika patuh, tampilkan luas hijaunya saja
        const luasKdhUser = (luasPersil * (kdhInput / 100)).toFixed(2);
        kdhDisplay = `${kdhInput}% <span class="text-gray-500 font-normal">(${luasKdhUser} m²)</span>`;
    }
    // ----------------------------------------------------

    const createRow = (label, nilaiAnda, batas, statusItem) => {
        const isPass = statusItem.status === 'pass';
        const badgeClass = isPass 
            ? "bg-green-100 text-green-700 border border-green-200" 
            : "bg-red-100 text-red-700 border border-red-200";
        const statusText = isPass ? "PATUH" : "TIDAK PATUH";
        const icon = isPass ? "check" : "x";

        return `
            <tr class="hover:bg-gray-50 transition">
                <td class="px-4 py-3 font-medium text-gray-700 align-top">${label}</td>
                <td class="px-4 py-3 font-bold text-gray-900 align-top">${nilaiAnda}</td>
                <td class="px-4 py-3 text-gray-500 align-top">${batas}</td>
                <td class="px-4 py-3 align-top">
                    <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-bold ${badgeClass}">
                        <i data-feather="${icon}" class="w-3 h-3 mr-1"></i> ${statusText}
                    </span>
                </td>
            </tr>
        `;
    };

    htmlContent += `
        <div class="mt-4 pt-6 border-t border-gray-100">
            <button type="button" 
                onclick="const el = document.getElementById('tabelDetail'); el.classList.toggle('hidden'); const icon = document.getElementById('toggleIcon'); if(el.classList.contains('hidden')){icon.setAttribute('data-feather', 'chevron-right')}else{icon.setAttribute('data-feather', 'chevron-down')}; feather.replace();"
                class="flex items-center justify-between w-full text-left focus:outline-none group hover:bg-gray-50 p-2 rounded-lg transition-colors">
                
                <h3 class="text-lg font-bold text-gray-800 uppercase tracking-wide text-sm flex items-center group-hover:text-primary transition-colors">
                    <i data-feather="list" class="w-4 h-4 mr-2"></i> TABEL DETAIL TEKNIS
                </h3>
                
                <i id="toggleIcon" data-feather="chevron-right" class="w-5 h-5 text-gray-400 group-hover:text-primary transition-colors"></i>
            </button>
            
            <div id="tabelDetail" class="hidden mt-4 overflow-hidden rounded-lg border border-gray-200 shadow-sm transition-all duration-300">
                <table class="min-w-full bg-white text-sm text-left">
                    <thead class="bg-gray-50 text-xs text-gray-700 uppercase tracking-wider border-b border-gray-200">
                        <tr>
                            <th class="px-4 py-3 font-bold">Parameter</th>
                            <th class="px-4 py-3 font-bold">Nilai Anda</th>
                            <th class="px-4 py-3 font-bold">Batas Aturan</th>
                            <th class="px-4 py-3 font-bold">Status</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-gray-100">
                        ${createRow("KDB (Koefisien Dasar Bangunan)", `${kdbInput}%`, `Maks ${kdbMaks}%`, detailPemeriksaan[0])}
                        ${createRow("Luas Bangunan Dasar", `${luasBangunanInput} m²`, `Maks ${maxLuasBangunan} m²`, detailPemeriksaan[1])}
                        ${createRow("Total Luas Lantai (KLB)", `${luasLantaiInput} m²`, `Maks ${maxLuasLantai} m²`, detailPemeriksaan[2])}
                        
                        ${createRow("KDH (Koefisien Dasar Hijau)", kdhDisplay, `Min ${kdhMin}%`, detailPemeriksaan[3])}
                        
                        ${createRow("Ketinggian Bangunan", `${ketinggianInput} m`, `Maks 15 m`, detailPemeriksaan[4])}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    hasilDiv.innerHTML = htmlContent;
    hasilDiv.style.display = 'block';

    if (typeof feather !== 'undefined') {
        feather.replace();
    }
}
// --- FUNGSI 5: exportToWord ---
async function exportToWord(fileType) {
    const buttonId = `export${fileType.toUpperCase()}Button`; 
    const submitButton = document.getElementById(buttonId);
    if (!submitButton) {
        console.error(`Tombol dengan ID ${buttonId} tidak ditemukan!`);
        return;
    }

    // --- VALIDASI BARU ---
    const pemohonInputs = document.querySelectorAll('#data-pemohon-section input[required]');
    let isPemohonValid = true;
    let firstInvalidPemohon = null;

    pemohonInputs.forEach(input => {
        if (!input.checkValidity()) {
            isPemohonValid = false;
            if (!firstInvalidPemohon) firstInvalidPemohon = input;
        }
    });

    if (!isPemohonValid) {
        alert("Harap lengkapi semua Data Pemohon sebelum mencetak laporan.");
        if (firstInvalidPemohon) firstInvalidPemohon.reportValidity();
        return;
    }

    const originalText = submitButton.innerHTML; 
    submitButton.innerHTML = 'Membuat dokumen...'; 
    submitButton.disabled = true; 

    // Kumpulkan SEMUA data
    const formData = {
        namaPemohon: document.getElementById('namaPemohon').value,
        nikPemohon: document.getElementById('nikPemohon').value,
        alamatPemohon: document.getElementById('alamatPemohon').value,
        nomorTelepon: document.getElementById('nomorTelepon').value,
        emailPemohon: document.getElementById('emailPemohon').value,
        jenisKegiatan: document.getElementById('jenisKegiatan').value,
        swpPilihan: document.getElementById('swp').value,
        zonaPilihan: document.getElementById('zona').value,
        subZonaPilihan: document.getElementById('subZona').value,
        luasPersilPilihan: document.getElementById('luasPersil').value,
        luasPersilInput: document.getElementById('luasPersilInput').value,
        kdbDiusulkan: document.getElementById('kdbDiusulkan').value,
        luasBangunanDiusulkan: document.getElementById('luasBangunan').value,
        luasLantaiDiusulkan: document.getElementById('luasLantai').value,
        ketinggian: document.getElementById('ketinggian').value,
        kdhDiusulkan: document.getElementById('kdhDiusulkan').value,
        isSubZonaRequired: document.getElementById('subZonaContainer').style.display !== 'none' && document.getElementById('subZona').hasAttribute('required'),
        gsbDiusulkan: document.getElementById('gsbDiusulkan').value, 
        fileType: fileType 
    };

    try {
        const response = await fetch(EXPORT_API_URL, { 
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(formData),
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => null); 
            throw new Error(errorData?.error || `Gagal membuat dokumen: ${response.statusText}`);
        }

        const blob = await response.blob(); 
        const url = window.URL.createObjectURL(blob); 
        const a = document.createElement('a'); 
        a.href = url;
        a.download = `Rekomendasi_${fileType}_${formData.namaPemohon || 'Report'}.docx`; 
        document.body.appendChild(a); 
        a.click(); 
        window.URL.revokeObjectURL(url); 
        document.body.removeChild(a); 

        alert(`✅ Dokumen ${fileType} berhasil diunduh!`);

    } catch (error) { 
        console.error(`Error export ${fileType} Word:`, error);
        alert(`❌ Gagal membuat dokumen ${fileType}: ${error.message}`);
    } finally { 
        submitButton.innerHTML = originalText;
        submitButton.disabled = false;
        if (typeof feather !== 'undefined' && feather.replace) {
            feather.replace(); 
        }
    }
}

// === Event listeners ===
document.addEventListener('DOMContentLoaded', () => {
    if (typeof updateZonaOptions === 'function') {
        updateZonaOptions(); 
    }
    const swpEl = document.getElementById('swp');
    const zonaEl = document.getElementById('zona');
    if (swpEl) swpEl.addEventListener('change', updateZonaOptions);
    if (zonaEl) zonaEl.addEventListener('change', tampilkanSubZona);

    const form = document.getElementById('formKepatuhan');
    if(form) {
        form.addEventListener('submit', cekKepatuhan);
    } else {
        console.error("Form dengan ID 'formKepatuhan' tidak ditemukan!");
    }
    
    if (typeof feather !== 'undefined') {
        feather.replace();
    }
});