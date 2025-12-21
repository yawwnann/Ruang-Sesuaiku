// =======================================================================
// FILE: server.js (EXPRESS.JS SERVER - KONVERSI DARI app.py)
// =======================================================================

// --- INSTALL DEPENDENCIES TERLEBIH DAHULU ---
// npm install express cors dotenv
// npm install docx
// =======================================================

const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  VerticalAlign,
  BorderStyle,
  WidthType,
  convertInchesToTwip,
} = require("docx");

// Import Aturan Lengkap dari data.js
const ATURAN_LENGKAP_SWP = require("./data.js");
console.log("✓ Aturan berhasil dimuat dari data.js");

// Inisialisasi Express App
const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(express.json());
app.use(cors());

// CSP Header - Allow all content (permissive)
app.use((req, res, next) => {
  res.removeHeader("Content-Security-Policy");
  next();
});

// Load HTML files ke cache
const htmlCache = {};
const htmlFiles = [
  "index.html",
  "kepatuhan.html",
  "kalkulator.html",
  "regulasi.html",
  "login.html",
  "tentang.html",
  "bantuan.html",
];

for (const file of htmlFiles) {
  try {
    htmlCache[file] = fs.readFileSync(path.join(__dirname, file), "utf8");
  } catch (e) {
    console.warn(`⚠ Warning: Could not load ${file}`);
  }
}

// Middleware untuk HTML files dari cache
app.use((req, res, next) => {
  const pathname = req.path;
  const filename = pathname.substring(1); // Remove leading /

  if (htmlCache[filename]) {
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    return res.send(htmlCache[filename]);
  }

  next();
});

// Static files middleware
app.use(express.static(path.join(__dirname)));

// Konstanta
const ATURAN_TINGGI_MAKS = 15;

// Helper Functions
function getData(dataObj, key, defaultVal = "...") {
  return dataObj[key] || defaultVal;
}

// Format Tanggal Indonesia Lengkap
function formatTanggalIndoLengkap() {
  const hariList = [
    "Senin",
    "Selasa",
    "Rabu",
    "Kamis",
    "Jumat",
    "Sabtu",
    "Minggu",
  ];
  const bulanList = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ];

  const now = new Date();
  const hari = hariList[now.getDay() === 0 ? 6 : now.getDay() - 1];
  const tanggal = now.getDate();
  const bulan = bulanList[now.getMonth()];
  const tahun = now.getFullYear();

  return `${hari}, tanggal ${tanggal} bulan ${bulan} tahun ${tahun}`;
}

// Format Tanggal Short
function formatTanggalShort() {
  const bulanList = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ];
  const now = new Date();
  return `${now.getDate()} ${bulanList[now.getMonth()]} ${now.getFullYear()}`;
}

// Main Function - Proses Cek Kepatuhan
function prosesCekKepatuhan(data) {
  try {
    // Extract data
    const swpPilihan = data.swpPilihan;
    const zonaPilihan = data.zonaPilihan;
    const subZonaPilihan = data.subZonaPilihan;
    const luasPersilPilihan = data.luasPersilPilihan;

    const luasPersilInput = parseFloat(data.luasPersilInput);
    const kdbDiusulkan = parseFloat(data.kdbDiusulkan);
    const luasBangunanDiusulkan = parseFloat(data.luasBangunanDiusulkan);
    const luasLantaiDiusulkan = parseFloat(data.luasLantaiDiusulkan);
    const ketinggian = parseFloat(data.ketinggian);
    const kdhDiusulkan = parseFloat(data.kdhDiusulkan);

    // Get SWP Rules
    let swpRules =
      ATURAN_LENGKAP_SWP[swpPilihan] ||
      ATURAN_LENGKAP_SWP.DEFAULT_PLACEHOLDER ||
      {};
    let zonaRules = swpRules[zonaPilihan];

    if (!zonaRules) {
      zonaRules = (ATURAN_LENGKAP_SWP.DEFAULT_PLACEHOLDER || {})[zonaPilihan];
      if (!zonaRules) {
        return { error: `Zona ${zonaPilihan} tidak ditemukan.` };
      }
    }

    // Tentukan Kunci Aturan Final
    let keyAturanFinal = "";
    const isSubZonaRequired = data.isSubZonaRequired;

    if (isSubZonaRequired && zonaRules[subZonaPilihan]) {
      keyAturanFinal = subZonaPilihan;
    } else {
      keyAturanFinal = Object.keys(zonaRules)[0];
    }

    const aturanSpesifikDict = zonaRules[keyAturanFinal] || {};
    const ATURAN_SPESIFIK = aturanSpesifikDict[luasPersilPilihan];

    if (!ATURAN_SPESIFIK) {
      return {
        error: `Aturan tidak ditemukan untuk '${keyAturanFinal}' dengan luas persil '${luasPersilPilihan}'.`,
      };
    }

    const [KDB_MAKS, KLB_MAKS, KDH_MIN] = ATURAN_SPESIFIK;

    // Perhitungan
    const LBD_MAKS = luasPersilInput * (KDB_MAKS / 100);
    const TLL_MAKS = luasPersilInput * KLB_MAKS;
    const KDB_AKTUAL =
      luasPersilInput > 0 ? (luasBangunanDiusulkan / luasPersilInput) * 100 : 0;
    const KLB_AKTUAL =
      luasPersilInput > 0 ? luasLantaiDiusulkan / luasPersilInput : 0;

    // Pemeriksaan Detail
    const detailPemeriksaan = [];

    // [0] KDB
    const statusKdb = kdbDiusulkan > KDB_MAKS ? "fail" : "pass";
    detailPemeriksaan.push({
      status: statusKdb,
      message: `KDB (%) Diusulkan (${kdbDiusulkan.toFixed(1)}%) ${
        statusKdb === "pass" ? "memenuhi" : "melebihi"
      } batas Zona (${KDB_MAKS}%).`,
      label: "KDB",
      val: kdbDiusulkan,
      limit: KDB_MAKS,
    });

    // [1] Luas Bangunan Dasar
    const statusLbd = luasBangunanDiusulkan > LBD_MAKS ? "fail" : "pass";
    detailPemeriksaan.push({
      status: statusLbd,
      message: `Luas Bangunan Dasar (${luasBangunanDiusulkan.toFixed(1)} m²) ${
        statusLbd === "pass" ? "aman" : "terlalu besar"
      } (Maks: ${LBD_MAKS.toFixed(1)} m²).`,
      label: "LBD",
      val: luasBangunanDiusulkan,
      limit: LBD_MAKS,
    });

    // [2] Total Luas Lantai (KLB)
    const statusKlb = KLB_AKTUAL > KLB_MAKS ? "fail" : "pass";
    detailPemeriksaan.push({
      status: statusKlb,
      message: `Koefisien Lantai Bangunan (${KLB_AKTUAL.toFixed(2)}) ${
        statusKlb === "pass" ? "memenuhi" : "melebihi"
      } batas maksimum (${KLB_MAKS}).`,
      label: "KLB",
      val: KLB_AKTUAL,
      limit: KLB_MAKS,
    });

    // [3] KDH
    const statusKdh = kdhDiusulkan < KDH_MIN ? "fail" : "pass";
    detailPemeriksaan.push({
      status: statusKdh,
      message: `KDH Hasil Pengukuran (${kdhDiusulkan.toFixed(2)}%) ${
        statusKdh === "pass" ? "cukup" : "kurang"
      } dari batas minimum (${KDH_MIN}%).`,
      label: "KDH",
      val: kdhDiusulkan,
      limit: KDH_MIN,
    });

    // [4] Ketinggian
    const statusTb = ketinggian > ATURAN_TINGGI_MAKS ? "fail" : "pass";
    detailPemeriksaan.push({
      status: statusTb,
      message: `Ketinggian (${ketinggian} m) ${
        statusTb === "pass" ? "aman" : "melebihi"
      } batas (${ATURAN_TINGGI_MAKS} m).`,
      label: "Tinggi",
      val: ketinggian,
      limit: ATURAN_TINGGI_MAKS,
    });

    const isPatuh = !detailPemeriksaan.some((item) => item.status === "fail");

    return {
      isPatuh,
      detailPemeriksaan,
      kdbAktual: KDB_AKTUAL,
      klbAktual: KLB_AKTUAL,
      kdbMaks: KDB_MAKS,
      klbMaks: KLB_MAKS,
      kdhMin: KDH_MIN,
      aturanTinggi: ATURAN_TINGGI_MAKS,
    };
  } catch (err) {
    return { error: `Terjadi kesalahan di server: ${err.message}` };
  }
}

// Favicon route (prevent 404)
app.get("/favicon.ico", (req, res) => {
  res.setHeader("Content-Type", "image/x-icon");
  res.status(204).end(); // No content
});

// API Endpoints
app.post("/api/cek_kepatuhan", (req, res) => {
  const data = req.body;

  if (!ATURAN_LENGKAP_SWP || Object.keys(ATURAN_LENGKAP_SWP).length <= 1) {
    return res.status(500).json({
      error: "Database Aturan (ATURAN_LENGKAP_SWP) kosong di server.",
    });
  }

  const hasil = prosesCekKepatuhan(data);

  if (hasil.error) {
    return res.status(400).json(hasil);
  }

  res.json(hasil);
});

app.get("/api/get_subzonas", (req, res) => {
  const swpPilihan = req.query.swp;
  const zonaPilihan = req.query.zona;

  if (!swpPilihan || !zonaPilihan) {
    return res.status(400).json({ error: "SWP dan Zona diperlukan" });
  }

  if (!ATURAN_LENGKAP_SWP || Object.keys(ATURAN_LENGKAP_SWP).length <= 1) {
    return res.status(500).json({
      error: "Database Aturan (ATURAN_LENGKAP_SWP) kosong di server.",
    });
  }

  let swpRules =
    ATURAN_LENGKAP_SWP[swpPilihan] ||
    ATURAN_LENGKAP_SWP.DEFAULT_PLACEHOLDER ||
    {};
  let zonaRules = swpRules[zonaPilihan];

  if (!zonaRules) {
    zonaRules = (ATURAN_LENGKAP_SWP.DEFAULT_PLACEHOLDER || {})[zonaPilihan];
    if (!zonaRules) {
      return res.json([]);
    }
  }

  const subzonaList = Object.keys(zonaRules);
  res.json(subzonaList);
});

app.post("/api/get_rules", (req, res) => {
  try {
    const data = req.body;
    const swpPilihan = data.swpPilihan;
    const zonaPilihan = data.zonaPilihan;
    const subZonaPilihan = data.subZonaPilihan;
    const luasPersilPilihan = data.luasPersilPilihan;

    let swpRules =
      ATURAN_LENGKAP_SWP[swpPilihan] ||
      ATURAN_LENGKAP_SWP.DEFAULT_PLACEHOLDER ||
      {};
    let zonaRules = swpRules[zonaPilihan];

    if (!zonaRules) {
      zonaRules =
        (ATURAN_LENGKAP_SWP.DEFAULT_PLACEHOLDER || {})[zonaPilihan] || {};
    }

    let keyAturanFinal = subZonaPilihan;
    if (!keyAturanFinal || !zonaRules[keyAturanFinal]) {
      keyAturanFinal = Object.keys(zonaRules)[0];
    }

    const aturanSpesifikDict = zonaRules[keyAturanFinal] || {};
    const ATURAN_SPESIFIK = aturanSpesifikDict[luasPersilPilihan];

    if (!ATURAN_SPESIFIK) {
      return res.status(404).json({
        error: `Aturan tidak ditemukan untuk ${keyAturanFinal} di ${luasPersilPilihan}.`,
      });
    }

    const [KDB_MAKS, KLB_MAKS, KDH_MIN] = ATURAN_SPESIFIK;

    res.json({
      kdbMaks: KDB_MAKS,
      klbMaks: KLB_MAKS,
      kdhMin: KDH_MIN,
    });
  } catch (err) {
    res.status(500).json({
      error: `Terjadi kesalahan di server: ${err.message}`,
    });
  }
});

app.post("/api/export_word", async (req, res) => {
  try {
    const data = req.body;
    const fileType = data.fileType;

    const hasil = prosesCekKepatuhan(data);
    if (hasil.error) {
      return res.status(400).json(hasil);
    }

    // Buat dokumen Word berdasarkan tipe file
    let doc;
    let filename;

    if (fileType === "PMP") {
      doc = createPmpDoc(data, hasil);
      filename = `Rekomendasi_PMP_${getData(
        data,
        "namaPemohon",
        "Pemohon"
      )}.docx`;
    } else if (fileType === "IKKPR") {
      doc = createIkkprDoc(data, hasil);
      filename = `Rekomendasi_IKKPR_${getData(
        data,
        "namaPemohon",
        "Pemohon"
      )}.docx`;
    } else if (fileType === "BA") {
      doc = createBaDoc(data, hasil);
      filename = `BA_Pemeriksaan_${getData(
        data,
        "namaPemohon",
        "Pemohon"
      )}.docx`;
    } else {
      return res.status(400).json({ error: "Tipe file tidak valid" });
    }

    // Convert dokumen ke buffer
    const bytes = await Packer.toBuffer(doc);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(bytes);
  } catch (err) {
    console.error("Error creating word doc:", err);
    res.status(500).json({ error: `Gagal membuat dokumen: ${err.message}` });
  }
});

// Fungsi Membuat Dokumen PMP
function createPmpDoc(data, hasil) {
  const luasPersilPilihan = getData(data, "luasPersilPilihan");

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.79),
              right: convertInchesToTwip(0.79),
              bottom: convertInchesToTwip(0.79),
              left: convertInchesToTwip(0.79),
            },
          },
        },
        children: [
          // Header
          new Paragraph({
            text: "REKOMENDASI HASIL PENILAIAN\nPERNYATAAN MANDIRI PELAKU USAHA MIKRO DAN KECIL",
            alignment: AlignmentType.CENTER,
            bold: true,
            spacing: { after: 200 },
          }),

          // Intro
          new Paragraph({
            text: `\n"Pada hari ini ${formatTanggalIndoLengkap()}, telah dilaksanakan penilaian Pernyataan Mandiri Pelaku Usaha Mikro dan Kecil (UMK) terhadap: "`,
            spacing: { after: 200 },
          }),

          // Data Pemohon
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    children: [new Paragraph("Nama Pelaku Usaha")],
                  }),
                  new TableCell({
                    width: { size: 5, type: WidthType.PERCENTAGE },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    width: { size: 65, type: WidthType.PERCENTAGE },
                    children: [new Paragraph(getData(data, "namaPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Nomor Identitas")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "nikPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Alamat")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "alamatPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Nomor Telepon")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "nomorTelepon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Email")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "emailPemohon"))],
                  }),
                ],
              }),
            ],
          }),

          // Intro text
          new Paragraph({
            text: '\n"Berdasarkan hasil survei lapangan dan analisis, diperoleh hasil penilaian Pernyataan Mandiri Pelaku Usaha Mikro dan Kecil UMK sebagai berikut: "',
            spacing: { before: 200, after: 100 },
          }),

          new Paragraph({
            text: "Hasil Penilaian Pernyataan Mandiri Pelaku UMK",
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),

          // Tabel Hasil (6 Kolom)
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              // Header
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "No",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Muatan yang Dinilai",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Hasil Penilaian\n(Sesuai/Tidak Sesuai)",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Keterangan",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Dasar Aturan",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Rekomendasi",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                ],
              }),

              // Section: Pemeriksaan
              new TableRow({
                children: [
                  new TableCell({
                    columnSpan: 6,
                    children: [
                      new Paragraph({ text: "Pemeriksaan", bold: true }),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("1")] }),
                  new TableCell({
                    children: [new Paragraph("Lokasi Kegiatan")],
                  }),
                  new TableCell({ children: [new Paragraph("Sesuai")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "alamatPemohon"))],
                  }),
                  new TableCell({ children: [new Paragraph("-")] }),
                  new TableCell({
                    children: [new Paragraph("Tidak ada rekomendasi")],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("2")] }),
                  new TableCell({
                    children: [new Paragraph("Jenis Kegiatan")],
                  }),
                  new TableCell({ children: [new Paragraph("Sesuai")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "jenisKegiatan"))],
                  }),
                  new TableCell({ children: [new Paragraph("-")] }),
                  new TableCell({
                    children: [new Paragraph("Tidak ada rekomendasi")],
                  }),
                ],
              }),

              // Section: Pengukuran
              new TableRow({
                children: [
                  new TableCell({
                    columnSpan: 6,
                    children: [
                      new Paragraph({
                        text: "Pengukuran (opsional)",
                        bold: true,
                      }),
                    ],
                  }),
                ],
              }),

              // 1. KDB
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("1")] }),
                  new TableCell({
                    children: [new Paragraph("Koefisien Dasar Bangunan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[0].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `(${hasil.detailPemeriksaan[0].val.toFixed(2)}%)`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Berdasarkan Peraturan Walikota Yogyakarta Nomor 118 Tahun 2021 Tentang RDTR Luas Tanah/Persil ${luasPersilPilihan} m2 KDB <${hasil.detailPemeriksaan[0].limit}%`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[0].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penyesuaian"
                      ),
                    ],
                  }),
                ],
              }),

              // 2. KLB
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("2")] }),
                  new TableCell({
                    children: [new Paragraph("Koefisien Lantai Bangunan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[2].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `(${hasil.detailPemeriksaan[2].val.toFixed(2)})`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Berdasarkan Peraturan Walikota Yogyakarta Nomor 118 Tahun 2021 Tentang RDTR Luas Tanah/Persil ${luasPersilPilihan} m2 KLB <${hasil.detailPemeriksaan[2].limit}`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[2].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penyesuaian"
                      ),
                    ],
                  }),
                ],
              }),

              // 3. GSB
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("3")] }),
                  new TableCell({
                    children: [new Paragraph("Garis Sempadan Bangunan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        getData(data, "gsbDiusulkan", "0") === "0"
                          ? "Tidak Sesuai"
                          : "Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `(${getData(data, "gsbDiusulkan", "0")} m)`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [new Paragraph("Perwal No 118 Th 2021")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        getData(data, "gsbDiusulkan", "0") === "0"
                          ? "Perlu penyesuaian"
                          : "Tidak ada rekomendasi"
                      ),
                    ],
                  }),
                ],
              }),

              // 4. Jarak Bebas
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("4")] }),
                  new TableCell({
                    children: [new Paragraph("Jarak Bebas Bangunan")],
                  }),
                  new TableCell({ children: [new Paragraph("Sesuai")] }),
                  new TableCell({ children: [new Paragraph("U:0m S:0m")] }),
                  new TableCell({
                    children: [new Paragraph("Perwal No 118 Th 2021")],
                  }),
                  new TableCell({
                    children: [new Paragraph("Tidak ada rekomendasi")],
                  }),
                ],
              }),

              // 5. KDH
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("5")] }),
                  new TableCell({
                    children: [new Paragraph("Koefisien Dasar Hijau")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[3].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `(${hasil.detailPemeriksaan[3].val.toFixed(2)}%)`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Berdasarkan Peraturan Walikota Kota Yogyakarta No 118 Tahun 2021 Tentang Koefisien Dasar Hijau >${hasil.detailPemeriksaan[3].limit}%`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[3].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penambahan RTH"
                      ),
                    ],
                  }),
                ],
              }),

              // 6. Basement
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("6")] }),
                  new TableCell({
                    children: [new Paragraph("Koefisien Tapak Basement")],
                  }),
                  new TableCell({ children: [new Paragraph("-")] }),
                  new TableCell({ children: [new Paragraph("-")] }),
                  new TableCell({ children: [new Paragraph("-")] }),
                  new TableCell({
                    children: [new Paragraph("Tidak ada rekomendasi")],
                  }),
                ],
              }),

              // 7. Ketinggian
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("7")] }),
                  new TableCell({
                    children: [new Paragraph("Ketinggian Bangunan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[4].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(`(${hasil.detailPemeriksaan[4].val} m)`),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Perwal No 118 Th 2021 (Maks ${hasil.detailPemeriksaan[4].limit} m)`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[4].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penyesuaian"
                      ),
                    ],
                  }),
                ],
              }),
            ],
          }),

          // Kesimpulan
          new Paragraph({
            text: '\n"Berdasarkan hasil penilaian tersebut, diperoleh kesimpulan sebagai berikut:"',
            spacing: { before: 200, after: 100 },
          }),

          new Paragraph({
            children: [
              {
                text: "Kegiatan Pemanfaatan Ruang berdasarkan pernyataan mandiri yang dibuat oleh pelaku UMK dinyatakan ",
              },
              { text: hasil.isPatuh ? "SESUAI" : "TIDAK SESUAI", bold: true },
              {
                text: hasil.isPatuh
                  ? "."
                  : " dan perlu menyesuaikan kondisi di lapangan.",
              },
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            text: '"Demikian untuk dapat digunakan sebagai acuan dalam penetapan hasil penilaian PMP UMK, atas perhatiannya diucapkan terimakasih."',
            spacing: { after: 50 },
          }),

          new Paragraph({
            text: "CP : 085601174240 (Bidang Dalwas DPTR Kota Yogyakarta)",
            spacing: { after: 200 },
          }),

          // Tanda Tangan
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("")],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [
                      new Paragraph({
                        text: `Yogyakarta, ${formatTanggalShort()}`,
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "Kepala Bidang\nPengendalian dan Pengawasan\nDinas Pertanahan dan Tata Ruang\n( Kundha Niti Mandala Sarta Tata Sasana )",
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "\n\n",
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "(Arif Amrullah, M.T., M.Sc.)",
                        alignment: AlignmentType.CENTER,
                        bold: true,
                      }),
                      new Paragraph({
                        text: "NIP. 197510132005011007",
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  return doc;
}

// Fungsi Membuat Dokumen IKKPR
function createIkkprDoc(data, hasil) {
  const luasPersilPilihan = getData(data, "luasPersilPilihan");

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.79),
              right: convertInchesToTwip(0.79),
              bottom: convertInchesToTwip(0.79),
              left: convertInchesToTwip(0.79),
            },
          },
        },
        children: [
          // Header
          new Paragraph({
            text: "REKOMENDASI HASIL KEPATUHAN PELAKSANAAN KETENTUAN\nKESESUAIAN KEGIATAN PEMANFAATAN RUANG\nPERIODE PASCA PEMBANGUNAN",
            alignment: AlignmentType.CENTER,
            bold: true,
            spacing: { after: 300 },
          }),

          new Paragraph({
            text: "Nomor: .../.../.../2024",
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: `${formatTanggalIndoLengkap()}`,
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
          }),

          // Bio Table
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("Nama Pelaku Usaha")],
                  }),
                  new TableCell({
                    width: { size: 5, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    width: { size: 65, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(getData(data, "namaPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("Nomor Identitas")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(getData(data, "nikPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("Alamat")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(getData(data, "alamatPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("Nomor Telepon")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(getData(data, "nomorTelepon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("Email")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(getData(data, "emailPemohon"))],
                  }),
                ],
              }),
            ],
          }),

          new Paragraph({ text: "", spacing: { after: 200 } }),

          // Lokasi Table
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("Lokasi Kegiatan")],
                  }),
                  new TableCell({
                    width: { size: 5, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    width: { size: 65, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(getData(data, "lokasiKegiatan"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("Jenis Kegiatan")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph(getData(data, "jenisKegiatan"))],
                  }),
                ],
              }),
            ],
          }),

          new Paragraph({ text: "", spacing: { after: 200 } }),

          new Paragraph({
            text: "Berdasarkan hasil survei lapangan dan analisis yang dilakukan oleh Tim Teknis Dinas Pertanahan dan Tata Ruang Kota Yogyakarta pada hari ... tanggal ... bulan ... tahun ... Terhadap Kegiatan Pemanfaatan Ruang yang dilakukan oleh pelaku UMK diatas, dengan hasil sebagai berikut:",
            spacing: { after: 200 },
          }),

          // Main Table - 6 columns
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              // Header
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 5, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        text: "No",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 20, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        text: "Muatan yang Dinilai",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 15, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        text: "Hasil Penilaian",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 15, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        text: "Keterangan",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        text: "Dasar Aturan",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 15, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        text: "Rekomendasi",
                        bold: true,
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                ],
              }),

              // Section: Pemeriksaan
              new TableRow({
                children: [
                  new TableCell({
                    columnSpan: 6,
                    children: [
                      new Paragraph({ text: "Pemeriksaan", bold: true }),
                    ],
                  }),
                ],
              }),

              // 1. Lokasi Kegiatan
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("1")] }),
                  new TableCell({
                    children: [new Paragraph("Lokasi Kegiatan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[1].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(`(${getData(data, "lokasiKegiatan")})`),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph("Sesuai zona pada Perwal No 118 Th 2021"),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[1].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penyesuaian"
                      ),
                    ],
                  }),
                ],
              }),

              // 2. Jenis Kegiatan
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("2")] }),
                  new TableCell({
                    children: [new Paragraph("Jenis Kegiatan")],
                  }),
                  new TableCell({ children: [new Paragraph("Sesuai")] }),
                  new TableCell({
                    children: [
                      new Paragraph(`(${getData(data, "jenisKegiatan")})`),
                    ],
                  }),
                  new TableCell({
                    children: [new Paragraph("Perwal No 118 Th 2021")],
                  }),
                  new TableCell({
                    children: [new Paragraph("Tidak ada rekomendasi")],
                  }),
                ],
              }),

              // Section: Pengukuran (opsional)
              new TableRow({
                children: [
                  new TableCell({
                    columnSpan: 6,
                    children: [
                      new Paragraph({
                        text: "Pengukuran (opsional)",
                        bold: true,
                      }),
                    ],
                  }),
                ],
              }),

              // 1. KDB
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("1")] }),
                  new TableCell({
                    children: [new Paragraph("Koefisien Dasar Bangunan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[0].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `(${hasil.detailPemeriksaan[0].val.toFixed(2)}%)`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Berdasarkan Perwal Kota Yogyakarta No 118 Tahun 2021 Tentang RDTR Luas Tanah/Persil ${luasPersilPilihan} m2 KLB <${hasil.detailPemeriksaan[0].limit}%`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[0].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penyesuaian"
                      ),
                    ],
                  }),
                ],
              }),

              // 2. KLB
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("2")] }),
                  new TableCell({
                    children: [new Paragraph("Koefisien Lantai Bangunan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[2].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `(${hasil.detailPemeriksaan[2].val.toFixed(2)})`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Berdasarkan Perwal Kota Yogyakarta No 118 Tahun 2021 Tentang Koefisien Lantai Bangunan <${hasil.detailPemeriksaan[2].limit}`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[2].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penyesuaian"
                      ),
                    ],
                  }),
                ],
              }),

              // 3. GSB
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("3")] }),
                  new TableCell({
                    children: [new Paragraph("Garis Sempadan Bangunan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        getData(data, "gsbDiusulkan", "0") === "0"
                          ? "Tidak Sesuai"
                          : "Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `(${getData(data, "gsbDiusulkan", "0")} m)`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [new Paragraph("Perwal No 118 Th 2021")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        getData(data, "gsbDiusulkan", "0") === "0"
                          ? "Perlu penyesuaian"
                          : "Tidak ada rekomendasi"
                      ),
                    ],
                  }),
                ],
              }),

              // 4. Jarak Bebas
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("4")] }),
                  new TableCell({
                    children: [new Paragraph("Jarak Bebas Bangunan")],
                  }),
                  new TableCell({ children: [new Paragraph("Sesuai")] }),
                  new TableCell({ children: [new Paragraph("U:0m S:0m")] }),
                  new TableCell({
                    children: [new Paragraph("Perwal No 118 Th 2021")],
                  }),
                  new TableCell({
                    children: [new Paragraph("Tidak ada rekomendasi")],
                  }),
                ],
              }),

              // 5. KDH
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("5")] }),
                  new TableCell({
                    children: [new Paragraph("Koefisien Dasar Hijau")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[3].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `(${hasil.detailPemeriksaan[3].val.toFixed(2)}%)`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Berdasarkan Perwal Kota Yogyakarta No 118 Tahun 2021 Tentang Koefisien Dasar Hijau >${hasil.detailPemeriksaan[3].limit}%`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[3].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penambahan RTH"
                      ),
                    ],
                  }),
                ],
              }),

              // 6. Basement
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("6")] }),
                  new TableCell({
                    children: [new Paragraph("Koefisien Tapak Basement")],
                  }),
                  new TableCell({ children: [new Paragraph("-")] }),
                  new TableCell({ children: [new Paragraph("-")] }),
                  new TableCell({ children: [new Paragraph("-")] }),
                  new TableCell({
                    children: [new Paragraph("Tidak ada rekomendasi")],
                  }),
                ],
              }),

              // 7. Ketinggian
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("7")] }),
                  new TableCell({
                    children: [new Paragraph("Ketinggian Bangunan")],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[4].status === "pass"
                          ? "Sesuai"
                          : "Tidak Sesuai"
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(`(${hasil.detailPemeriksaan[4].val} m)`),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Perwal No 118 Th 2021 (Maks ${hasil.detailPemeriksaan[4].limit} m)`
                      ),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        hasil.detailPemeriksaan[4].status === "pass"
                          ? "Tidak ada rekomendasi"
                          : "Perlu penyesuaian"
                      ),
                    ],
                  }),
                ],
              }),
            ],
          }),

          // Kesimpulan
          new Paragraph({
            text: '\n"Berdasarkan hasil penilaian tersebut, diperoleh kesimpulan sebagai berikut:"',
            spacing: { before: 200, after: 100 },
          }),

          new Paragraph({
            children: [
              {
                text: "Kegiatan Pemanfaatan Ruang berdasarkan Konfirmasi KKPR dinyatakan ",
              },
              { text: hasil.isPatuh ? "SESUAI" : "TIDAK SESUAI", bold: true },
              {
                text: hasil.isPatuh
                  ? " dan perlu mempertahankan kondisi di lapangan."
                  : " dan perlu menyesuaikan kondisi di lapangan sesuai parameter yang Tidak Patuh.",
              },
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            text: '"Demikian untuk dapat digunakan sebagai acuan dalam penetapan hasil penilaian Konfirmasi KKPR, atas perhatiannya diucapkan terimakasih."',
            spacing: { after: 50 },
          }),

          new Paragraph({
            text: "CP : 085601174240 (Bidang Dalwas DPTR Kota Yogyakarta)",
            spacing: { after: 200 },
          }),

          // Tanda Tangan
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [new Paragraph("")],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [
                      new Paragraph({
                        text: `Yogyakarta, ${formatTanggalShort()}`,
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "Plt. Kepala Bidang\nPengendalian dan Pengawasan\nDinas Pertanahan dan Tata Ruang\n( Kundha Niti Mandala Sarta Tata Sasana )",
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "\n\n",
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "(FX. Wahyu Setyowati, M.T)",
                        alignment: AlignmentType.CENTER,
                        bold: true,
                      }),
                      new Paragraph({
                        text: "NIP. 197109041997032006",
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  return doc;
}

// Fungsi Membuat Dokumen BA (Berita Acara)
function createBaDoc(data, hasil) {
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.79),
              right: convertInchesToTwip(0.79),
              bottom: convertInchesToTwip(0.79),
              left: convertInchesToTwip(0.79),
            },
          },
        },
        children: [
          // Header
          new Paragraph({
            text: "HASIL PEMERIKSAAN DAN PENGUKURAN\nPERNYATAAN MANDIRI PELAKU USAHA MIKRO DAN KECIL",
            alignment: AlignmentType.CENTER,
            bold: true,
            spacing: { after: 200 },
          }),
          new Paragraph({ text: "", spacing: { after: 200 } }),

          // Tabel Bio
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    children: [new Paragraph("Nama Pelaku Usaha")],
                  }),
                  new TableCell({
                    width: { size: 5, type: WidthType.PERCENTAGE },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    width: { size: 65, type: WidthType.PERCENTAGE },
                    children: [new Paragraph(getData(data, "namaPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Nomor Identitas")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "nikPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Alamat")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "alamatPemohon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Nomor Telepon")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "nomorTelepon"))],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Email")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "emailPemohon"))],
                  }),
                ],
              }),
            ],
          }),

          // Section Pemeriksaan
          new Paragraph({
            text: "\nPemeriksaan",
            bold: true,
            spacing: { after: 100 },
          }),
          new Paragraph({
            text: "Lokasi Kegiatan Usaha",
            spacing: { before: 100, after: 100 },
            indent: { left: convertInchesToTwip(0.2) },
          }),

          // Tabel Lokasi
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    children: [new Paragraph("Alamat")],
                  }),
                  new TableCell({
                    width: { size: 5, type: WidthType.PERCENTAGE },
                    children: [new Paragraph(":")],
                  }),
                  new TableCell({
                    width: { size: 65, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph(
                        getData(
                          data,
                          "alamatLokasi",
                          getData(data, "alamatPemohon")
                        )
                      ),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Desa/Kelurahan")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(getData(data, "kelurahanLoc", "-")),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Kecamatan")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(getData(data, "kecamatanLoc", "-")),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Kabupaten/Kota")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph("Kota Yogyakarta")],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Provinsi")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph("Daerah Istimewa Yogyakarta")],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Koordinat Lokasi")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `Lintang: ${getData(
                          data,
                          "lintang",
                          "-"
                        )} , Bujur: ${getData(data, "bujur", "-")}`
                      ),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Jenis Kegiatan Usaha")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [new Paragraph(getData(data, "jenisKegiatan"))],
                  }),
                ],
              }),
            ],
          }),

          // Section Pengukuran
          new Paragraph({
            text: "\nPengukuran",
            bold: true,
            spacing: { after: 100, before: 200 },
          }),

          // Tabel Teknis
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              // KDB Header
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        text: "Koefisien Dasar Bangunan",
                        bold: true,
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 5, type: WidthType.PERCENTAGE },
                    children: [new Paragraph("")],
                  }),
                  new TableCell({
                    width: { size: 55, type: WidthType.PERCENTAGE },
                    children: [new Paragraph("")],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Panjang Bangunan lantai dasar")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `${getData(data, "panjangBangunan", "-")} m`
                      ),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Lebar Bangunan lantai dasar")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(`${getData(data, "lebarBangunan", "-")} m`),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Luas Lantai Dasar Bangunan")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `${getData(data, "luasBangunanDiusulkan", "0")} m2`
                      ),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Luas Lahan")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `${getData(data, "luasPersilInput", "0")} m2`
                      ),
                    ],
                  }),
                ],
              }),

              // KLB Header
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Koefisien Lantai Bangunan",
                        bold: true,
                      }),
                    ],
                  }),
                  new TableCell({ children: [new Paragraph("")] }),
                  new TableCell({ children: [new Paragraph("")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Panjang lantai bangunan")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `${getData(data, "panjangBangunan", "-")} m`
                      ),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Lebar lantai bangunan")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(`${getData(data, "lebarBangunan", "-")} m`),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Luas seluruh lantai bangunan")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `${getData(data, "luasLantaiDiusulkan", "0")} m2`
                      ),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Jumlah lantai bangunan")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `${getData(data, "jumlahLantai", "1")} Lantai`
                      ),
                    ],
                  }),
                ],
              }),

              // GSB Header
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Garis Sempadan Bangunan",
                        bold: true,
                      }),
                    ],
                  }),
                  new TableCell({ children: [new Paragraph("")] }),
                  new TableCell({ children: [new Paragraph("")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph("Jarak bangunan terdepan dengan pagar"),
                    ],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(`${getData(data, "gsbDiusulkan", "0")} m`),
                    ],
                  }),
                ],
              }),

              // Jarak Bebas Header
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Jarak Bebas Bangunan",
                        bold: true,
                      }),
                    ],
                  }),
                  new TableCell({ children: [new Paragraph("")] }),
                  new TableCell({ children: [new Paragraph("")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph("Jarak bangunan terbelakang (Timur)"),
                    ],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({ children: [new Paragraph("0 m")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Jarak pagar samping")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph("Utara : 0 m\nSelatan : 0 m\nBarat : 0 m"),
                    ],
                  }),
                ],
              }),

              // KDH Header
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ text: "KDH", bold: true })],
                  }),
                  new TableCell({ children: [new Paragraph("")] }),
                  new TableCell({ children: [new Paragraph("")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Panjang lahan vegetasi")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({ children: [new Paragraph("-")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Lebar lahan vegetasi")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({ children: [new Paragraph("-")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Luas lahan vegetasi")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(
                        `${(
                          (parseFloat(getData(data, "kdhDiusulkan", 0)) / 100) *
                          parseFloat(getData(data, "luasPersilInput", 0))
                        ).toFixed(2)} m2`
                      ),
                    ],
                  }),
                ],
              }),

              // Basement Header
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Koefisien Tapak Basement",
                        bold: true,
                      }),
                    ],
                  }),
                  new TableCell({ children: [new Paragraph("")] }),
                  new TableCell({ children: [new Paragraph("")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Lokasi Basement")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({ children: [new Paragraph("-")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Luas Basement")] }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({ children: [new Paragraph("- m2")] }),
                ],
              }),

              // Ketinggian
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("")] }),
                  new TableCell({ children: [new Paragraph("")] }),
                  new TableCell({ children: [new Paragraph("")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("Ketinggian Bangunan")],
                  }),
                  new TableCell({ children: [new Paragraph(":")] }),
                  new TableCell({
                    children: [
                      new Paragraph(`${getData(data, "ketinggian", "0")} m`),
                    ],
                  }),
                ],
              }),
            ],
          }),

          // Closing
          new Paragraph({
            text: "\nDemikian Berita Acara ini dibuat dalam rangkap secukupnya untuk dipergunakan sebagaimana mestinya.",
            spacing: { before: 300, after: 300 },
          }),

          // Tanda Tangan Table
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [
                      new Paragraph({
                        text: "\n\nPemegang KKPR/Wakilnya",
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "\n\n\n\n(..................................................)",
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: {
                      top: { style: BorderStyle.NONE },
                      bottom: { style: BorderStyle.NONE },
                      left: { style: BorderStyle.NONE },
                      right: { style: BorderStyle.NONE },
                    },
                    children: [
                      new Paragraph({
                        text: "Kepala Bidang\nPengendalian dan Pengawasan\nDinas Pertanahan dan Tata Ruang\n( Kundha Niti Mandala Sarta Tata Sasana )",
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "\n\n",
                        alignment: AlignmentType.CENTER,
                      }),
                      new Paragraph({
                        text: "(Arif Amrullah, M.T., M.Sc)",
                        alignment: AlignmentType.CENTER,
                        bold: true,
                      }),
                      new Paragraph({
                        text: "NIP. 197510132005011007",
                        alignment: AlignmentType.CENTER,
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  return doc;
}

// Route untuk root path - serve index.html
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

// Start Server
app.listen(PORT, () => {
  console.log(`\n✓ Server Express berjalan di http://localhost:${PORT}`);
  console.log("Tekan CTRL+C untuk menghentikan server.\n");
});

module.exports = app;
