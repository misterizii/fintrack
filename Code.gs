/**
 * =====================================================
 * FINANCIAL TRACKER AND PLANNER - BACKEND
 * =====================================================
 * PERBAIKAN BUG: "anggaran baru tidak muncul walau notifikasi sukses".
 *
 * AKAR MASALAH (bukan soal format Tanggal):
 * Sheet 'Anggaran' versi lama memakai 5 kolom —
 *   [Kategori Pengeluaran, Batas Anggaran, Periode, Tahun, ID]
 * sedangkan versi sekarang memakai 4 kolom —
 *   [Kategori Pengeluaran, Bulan, Tahun, Batas Anggaran]
 * initializeDatabase() hanya menulis header saat sheet BARU dibuat, jadi
 * spreadsheet lama tetap memakai tata letak 5 kolom. Akibatnya deteksi kolom
 * salah: tidak ada kolom bernama "Bulan", sehingga colBulan jatuh ke nilai
 * default 2 — sama dengan colBatas ("Batas Anggaran" juga di kolom 2).
 * Saat menyimpan, nilai bulan dan nominal batas ditulis ke kolom yang sama
 * dan saling menimpa. Baris tersimpan (karena itu notifikasi sukses), tetapi
 * nilai bulannya jadi nominal batas (mis. 1500000), sehingga filter
 * `Number(a['Bulan']) === bulan` tidak pernah cocok dan anggaran itu tidak
 * pernah tampil di daftar.
 *
 * PERBAIKAN:
 * 1. normalizeAnggaranSheet_() — memigrasikan sheet tata letak lama ke tata
 *    letak baru SEKALI, dengan memetakan data yang sudah ada (kolom Periode &
 *    ID dibuang, Batas dipindah ke kolom 4). Idempoten: kalau header sudah
 *    baku, tidak ada yang diubah. Dipanggil dari initializeDatabase().
 * 2. parseAnggaranHeaderIndices_() — hanya menganggap tata letak valid kalau
 *    KEEMPAT kolom terdeteksi dan menunjuk empat kolom BERBEDA. Ini yang
 *    mencegah colBulan dan colBatas kembali bertabrakan.
 * 3. getAnggaranHeaderIndices() — kalau tata letak tidak valid, sheet
 *    dinormalisasi lebih dulu lalu header dibaca ulang, sehingga operasi
 *    baca/tulis anggaran tidak pernah memakai kolom yang salah.
 * 4. updateAnggaran() sekarang memvalidasi bulan/tahun/batas seperti
 *    addAnggaran(), supaya nilai tidak valid tidak bisa tertulis ke sheet.
 * 5. toSafeNumber()/dateSerialToNumber() tetap dipertahankan sebagai
 *    pengaman tambahan bila sel sempat ter-format Tanggal.
 * 6. sheetToObjects() membuang baris kosong, dan getAnggaran() melewati
 *    baris tanpa data, supaya tidak muncul kartu kosong di UI.
 * =====================================================
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_TRANSAKSI = 'Transaksi';
const SHEET_DOMPET    = 'Dompet';
const SHEET_ANGGARAN  = 'Anggaran';
const SHEET_HUTANG    = 'Hutang_Piutang';
const SHEET_TABUNGAN  = 'Tabungan';
const WALLET_DEFAULT  = 'Kas / Tunai';

const ADJUSTMENT_CATEGORY = 'Penyesuaian Saldo';
const TABUNGAN_PREFIX     = 'Tabungan: ';

// Nama dompet virtual (dihitung otomatis, tidak disimpan di sheet Dompet).
const WALLET_SISA_BULAN_LALU = 'SISA BULAN LALU';

function isReservedWalletName(nama) {
  return String(nama || '').trim().toLowerCase() === WALLET_SISA_BULAN_LALU.toLowerCase();
}

function isInternalCategory(kategori) {
  const k = String(kategori || '');
  return k === ADJUSTMENT_CATEGORY || k.indexOf(TABUNGAN_PREFIX) === 0;
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'Sistem sedang sibuk memproses, coba lagi sebentar.' };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  initializeDatabase();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Financial Tracker & Planner')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initializeDatabase() {
  try {
    const sheetsConfig = {
      [SHEET_TRANSAKSI]: ['Tanggal', 'Jenis', 'Sumber/Dompet', 'Kategori', 'Nominal', 'Keterangan', 'ID Transaksi'],
      [SHEET_DOMPET]:    ['Nama Dompet', 'Saldo Awal', 'Saldo Aktual (Balancing)'],
      [SHEET_ANGGARAN]:  ['Kategori Pengeluaran', 'Bulan', 'Tahun', 'Batas Anggaran'],
      [SHEET_HUTANG]:    ['ID', 'Jenis', 'Pihak Terkait', 'Nominal', 'Status', 'Tanggal Tenggat'],
      [SHEET_TABUNGAN]:  ['ID', 'Tujuan Tabungan', 'Target', 'Saldo Terkumpul']
    };

    Object.keys(sheetsConfig).forEach(name => {
      let sheet = SS.getSheetByName(name);
      if (!sheet) {
        sheet = SS.insertSheet(name);
        const headers = sheetsConfig[name];
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#8B6F47')
          .setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
        sheet.autoResizeColumns(1, headers.length);
      }
    });

    const sheetDompet = SS.getSheetByName(SHEET_DOMPET);
    if (sheetDompet && sheetDompet.getLastRow() < 2) {
      sheetDompet.appendRow([WALLET_DEFAULT, 0, 0]);
    }

    // PERBAIKAN BUG UTAMA (anggaran baru tidak muncul walau notifikasi sukses):
    // normalisasi tata letak sheet Anggaran lebih dulu — sheet yang dibuat
    // versi lama masih memakai 5 kolom (Kategori, Batas, Periode, Tahun, ID)
    // sehingga deteksi kolom jadi kacau. Setelah itu baru pastikan kolom
    // Bulan/Tahun/Batas berformat angka biasa, bukan Tanggal.
    try {
      normalizeAnggaranSheet_();
    } catch (e) {
      console.error('normalizeAnggaranSheet_ gagal: ' + e.message);
    }

    const sheetAnggaran = SS.getSheetByName(SHEET_ANGGARAN);
    if (sheetAnggaran) {
      const { colBulan, colTahun, colBatas } = getAnggaranHeaderIndices(sheetAnggaran);
      const dataRows = Math.max(sheetAnggaran.getLastRow() - 1, 1);
      [colBulan, colTahun, colBatas].forEach(col => {
        sheetAnggaran.getRange(2, col, dataRows, 1).setNumberFormat('0');
      });
    }

    return true;
  } catch(e) {
    console.error('initializeDatabase gagal: ' + e.message);
    return false;
  }
}

function sheetToObjects(sheetName) {
  try {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const tz = Session.getScriptTimeZone();

    return data.slice(1).map((row, idx) => {
      const obj = { _row: idx + 2 };
      headers.forEach((h, i) => {
        let val = row[i];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
        }
        obj[h] = (val !== undefined && val !== null) ? val : '';
      });
      return obj;
    })
    // Buang baris yang benar-benar kosong supaya tidak muncul sebagai
    // kartu/transaksi hantu di frontend.
    .filter(obj => Object.keys(obj).some(k => k !== '_row' && String(obj[k]).trim() !== ''));
  } catch(e) {
    return [];
  }
}

function generateId(prefix) {
  return prefix + '-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
}

function parseSafeDate(dateVal) {
  if (!dateVal) return new Date();
  if (dateVal instanceof Date) return dateVal;
  const parts = String(dateVal).split('-');
  if (parts.length === 3) {
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
  }
  const d = new Date(dateVal);
  return isNaN(d.getTime()) ? new Date() : d;
}

// BARU: ambil "yyyy-MM" dari sebuah nilai tanggal secara aman.
// Tanggal yang keluar dari sheetToObjects() sudah berupa string 'yyyy-MM-dd',
// jadi cukup diambil 7 karakter pertama — tidak perlu dan tidak boleh dibungkus
// ulang ke objek Date (itulah sumber bug pergeseran bulan sebelumnya).
function getYearMonthKey(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  const s = String(dateVal || '');
  return s.length >= 7 ? s.substring(0, 7) : '';
}

// BARU — INTI PERBAIKAN BUG ANGGARAN HILANG:
// Mengembalikan angka asli sebuah nilai, bahkan jika sel itu sudah kadung
// ter-format ulang menjadi Tanggal oleh Google Sheets. Serial tanggal Google
// Sheets dihitung sejak 30 Desember 1899, jadi kita balikkan perhitungannya
// untuk memulihkan angka bulan/tahun/batas anggaran yang asli.
function dateSerialToNumber(dateVal) {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  const asUtc = new Date(Date.UTC(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate()));
  return Math.round((asUtc - epoch) / 86400000);
}

function toSafeNumber(val) {
  if (val instanceof Date) return dateSerialToNumber(val);
  if (typeof val === 'string') {
    // String tanggal 'yyyy-MM-dd' (mis. hasil konversi sheetToObjects di
    // tempat lain) juga bisa muncul di sini — parse manual lalu hitung
    // ulang serial-nya supaya tetap mendapat angka yang benar.
    const m = val.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      return dateSerialToNumber(new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
    }
  }
  const n = Number(val);
  return isNaN(n) ? 0 : n;
}

function getAllInitData() {
  try {
    return {
      dashboard: getDashboardData(),
      dompets: getDompets(),
      sumberDana: getAllSumberDana(),
      kategoriList: getKategoriList(),
      kategoriPemasukanList: getKategoriPemasukanList(),
      anggaran: getAnggaran(),
      transaksi: getTransaksi(),
      hutangPiutang: getHutangPiutang(),
      tabungan: getTabungan()
    };
  } catch(e) {
    return {
      dashboard: { totalSaldo: 0, totalPemasukan: 0, totalPengeluaran: 0, rincian: [] },
      dompets: [],
      sumberDana: [WALLET_DEFAULT],
      kategoriList: getKategoriList(),
      kategoriPemasukanList: getKategoriPemasukanList(),
      anggaran: [],
      transaksi: [],
      hutangPiutang: [],
      tabungan: []
    };
  }
}

// ================= TRANSAKSI =================
function _addTransaksiRaw(data) {
  const nominal = Number(data.nominal);
  if (!nominal || nominal <= 0) return { success: false, error: 'Nominal harus lebih dari 0' };
  const sumber = String(data.sumber || '').trim();
  if (!sumber) return { success: false, error: 'Pilih dompet terlebih dahulu' };

  const sheet = SS.getSheetByName(SHEET_TRANSAKSI);
  const id = generateId('TRX');
  const tglObj = parseSafeDate(data.tanggal);

  sheet.appendRow([
    tglObj,
    data.jenis,
    sumber,
    String(data.kategori || '').trim() || (data.jenis === 'Pemasukan' ? 'Pemasukan' : 'Lainnya'),
    nominal,
    data.keterangan || '',
    id
  ]);
  return { success: true, id: id };
}

function addTransaksi(data) {
  return withLock(() => _addTransaksiRaw(data));
}

function addPemasukanSplit(data) {
  return withLock(() => {
    const sheet = SS.getSheetByName(SHEET_TRANSAKSI);
    const ids = [];
    const tglObj = parseSafeDate(data.tanggal);
    const kategori = String(data.kategori || '').trim() || 'Pemasukan';

    (data.splits || []).forEach(s => {
      const nominal = Number(s.nominal);
      const sumber = String(s.sumber || '').trim();
      if (nominal > 0 && sumber) {
        const id = generateId('TRX');
        sheet.appendRow([tglObj, 'Pemasukan', sumber, kategori, nominal, data.keterangan || '', id]);
        ids.push(id);
      }
    });

    if (!ids.length) {
      return { success: false, error: 'Isi minimal satu alokasi dompet dan nominal yang valid (> 0).' };
    }
    return { success: true, ids: ids };
  });
}

function updateTransaksi(id, data) {
  return withLock(() => {
    const nominal = Number(data.nominal);
    if (!nominal || nominal <= 0) return { success: false, error: 'Nominal harus lebih dari 0' };
    const sumber = String(data.sumber || '').trim();
    if (!sumber) return { success: false, error: 'Pilih dompet terlebih dahulu' };

    const sheet = SS.getSheetByName(SHEET_TRANSAKSI);
    const values = sheet.getDataRange().getValues();
    const tglObj = parseSafeDate(data.tanggal);

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][6]) === String(id)) {
        sheet.getRange(i + 1, 1, 1, 6).setValues([[
          tglObj, data.jenis, sumber, String(data.kategori || '').trim(), nominal, data.keterangan || ''
        ]]);
        return { success: true };
      }
    }
    return { success: false, error: 'Transaksi tidak ditemukan' };
  });
}

function deleteTransaksi(id) {
  return withLock(() => {
    const sheet = SS.getSheetByName(SHEET_TRANSAKSI);
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][6]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Transaksi tidak ditemukan' };
  });
}

function getTransaksi(searchTerm) {
  let data = sheetToObjects(SHEET_TRANSAKSI);
  data.sort((a, b) => new Date(b['Tanggal']) - new Date(a['Tanggal']));
  if (searchTerm) {
    const term = String(searchTerm).toLowerCase();
    data = data.filter(t =>
      String(t['Sumber/Dompet']).toLowerCase().includes(term) ||
      String(t['Kategori']).toLowerCase().includes(term) ||
      String(t['Keterangan']).toLowerCase().includes(term)
    );
  }
  return data;
}

// ================= DOMPET =================
function getDompets() {
  return sheetToObjects(SHEET_DOMPET);
}

function addDompet(nama, saldoAwal) {
  return withLock(() => {
    nama = String(nama || '').trim();
    if (!nama) return { success: false, error: 'Nama dompet tidak boleh kosong' };
    if (isReservedWalletName(nama)) return { success: false, error: 'Nama "' + WALLET_SISA_BULAN_LALU + '" dipakai otomatis oleh sistem, pilih nama lain' };

    const dup =getDompets().some(d => String(d['Nama Dompet']).toLowerCase() === nama.toLowerCase());
    if (dup) return { success: false, error: 'Dompet dengan nama itu sudah ada' };

    const sheet = SS.getSheetByName(SHEET_DOMPET);
    sheet.appendRow([nama, Number(saldoAwal) || 0, Number(saldoAwal) || 0]);
    return { success: true };
  });
}

function updateDompet(originalNama, data) {
  return withLock(() => {
    const namaBaru = String(data.nama || '').trim();
    if (!namaBaru) return { success: false, error: 'Nama dompet tidak boleh kosong' };
    if (isReservedWalletName(namaBaru)) return { success: false, error: 'Nama "' + WALLET_SISA_BULAN_LALU + '" dipakai otomatis oleh sistem, pilih nama lain' };

    if (namaBaru.toLowerCase() !== originalNama.toLowerCase()) {
      const dup = getDompets().some(d =>
        d['Nama Dompet'] !== originalNama && String(d['Nama Dompet']).toLowerCase() === namaBaru.toLowerCase()
      );
      if (dup) return { success: false, error: 'Dompet dengan nama itu sudah ada' };
    }

    const sheet = SS.getSheetByName(SHEET_DOMPET);
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === originalNama) {
        sheet.getRange(i + 1, 1, 1, 2).setValues([[namaBaru, Number(data.saldoAwal) || 0]]);
        if (namaBaru !== originalNama) {
          renameSumberDiTransaksi(originalNama, namaBaru);
        }
        return { success: true };
      }
    }
    return { success: false, error: 'Dompet tidak ditemukan' };
  });
}

function deleteDompet(nama) {
  return withLock(() => {
    const punyaTransaksi = sheetToObjects(SHEET_TRANSAKSI).some(t => t['Sumber/Dompet'] === nama);
    if (punyaTransaksi) {
      return { success: false, error: 'Dompet ini masih memiliki riwayat transaksi. Ganti namanya saja jika ingin memensiunkan.' };
    }
    const sheet = SS.getSheetByName(SHEET_DOMPET);
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === nama) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Dompet tidak ditemukan' };
  });
}

function renameSumberDiTransaksi(namaLama, namaBaru) {
  const sheet = SS.getSheetByName(SHEET_TRANSAKSI);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][2] === namaLama) {
      sheet.getRange(i + 1, 3).setValue(namaBaru);
    }
  }
}

function getAllSumberDana() {
  const dompets = getDompets().map(d => String(d['Nama Dompet'] || '').trim()).filter(Boolean);
  const set = new Set(dompets);
  const trx = sheetToObjects(SHEET_TRANSAKSI);
  trx.forEach(t => {
    const s = String(t['Sumber/Dompet'] || '').trim();
    if (s) set.add(s);
  });
  if (set.size === 0) set.add(WALLET_DEFAULT);
  return Array.from(set);
}

// ================= ANGGARAN (DENGAN DETEKSI HEADER FLEKSIBEL) =================

// Header baku sheet Anggaran versi sekarang.
const ANGGARAN_HEADERS = ['Kategori Pengeluaran', 'Bulan', 'Tahun', 'Batas Anggaran'];

/**
 * PERBAIKAN BUG UTAMA — normalisasi tata letak sheet Anggaran.
 *
 * Versi lama aplikasi ini memakai 5 kolom:
 *   [Kategori Pengeluaran, Batas Anggaran, Periode, Tahun, ID]
 * Versi sekarang memakai 4 kolom:
 *   [Kategori Pengeluaran, Bulan, Tahun, Batas Anggaran]
 *
 * Karena initializeDatabase() hanya menulis header saat sheet BARU dibuat,
 * spreadsheet lama tetap memakai tata letak 5 kolom. Akibatnya deteksi kolom
 * jadi salah: kolom "Bulan" dan "Batas Anggaran" sama-sama menunjuk kolom 2,
 * sehingga nilai bulan tertimpa nominal batas. Anggaran tersimpan (notifikasi
 * sukses) tapi tidak pernah lolos filter bulan -> seolah "tidak berhasil".
 *
 * Fungsi ini mengubah sheet lama ke tata letak baru SEKALI, dengan memetakan
 * data yang sudah ada (kolom Periode dibuang, ID dibuang, Batas dipindah ke
 * kolom 4). Aman dipanggil berulang kali: kalau header sudah baku, langsung
 * keluar tanpa mengubah apa pun.
 */
function normalizeAnggaranSheet_() {
  const sheet = SS.getSheetByName(SHEET_ANGGARAN);
  if (!sheet) return;

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = readAnggaranHeaders_(sheet);

  // Kalau header sudah baku (keempat kolom terdeteksi & tidak tumpang tindih),
  // tidak ada yang perlu diubah.
  if (parseAnggaranHeaderIndices_(headers).valid) return;

  // Cari posisi tiap kolom pada tata letak lama.
  let colKat = 1, colBatas = 2, colTahun = 3;
  headers.forEach((h, idx) => {
    const name = h.toLowerCase();
    if (name.includes('kategori')) colKat = idx + 1;
    else if (name.includes('batas') || name.includes('nominal')) colBatas = idx + 1;
    else if (name.includes('tahun')) colTahun = idx + 1;
  });

  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
    : [];

  const bulanDefault = new Date().getMonth() + 1;
  const tahunDefault = new Date().getFullYear();

  // Susun ulang jadi [Kategori, Bulan, Tahun, Batas].
  //
  // PENTING: header lama bisa sudah berisi baris yang ditulis oleh kode BARU
  // (karena header tidak pernah ikut diperbarui). Baris seperti itu punya
  // pola: kolom 2 = bulan (1-12), kolom 3 = tahun (2000-2100), kolom 4 =
  // nominal batas. Baris lama asli punya pola: kolom 2 = nominal batas,
  // kolom 3 = periode teks, kolom 4 = tahun. Kedua pola dibedakan per baris
  // supaya tidak ada data yang tertukar atau rusak.
  const baru = rows.map(row => {
    const kategori = String(row[colKat - 1] || '').trim();
    if (!kategori) return null;

    const v2 = toSafeNumber(row[1]);
    const v3 = toSafeNumber(row[2]);
    const v4 = toSafeNumber(row[3]);

    const terlihatFormatBaru = v2 >= 1 && v2 <= 12 && v3 >= 2000 && v3 <= 2100;
    if (terlihatFormatBaru) {
      return [kategori, v2, v3, v4];
    }

    // Baris lama: tidak menyimpan bulan, jadi diisi bulan berjalan agar
    // anggaran lama tetap terlihat (bisa diedit manual bila perlu).
    const tahun = toSafeNumber(row[colTahun - 1]) || tahunDefault;
    const batas = toSafeNumber(row[colBatas - 1]);
    return [kategori, bulanDefault, tahun, batas];
  }).filter(Boolean);

  // Tulis ulang sheet: header baku + data hasil pemetaan.
  sheet.clear();
  sheet.getRange(1, 1, 1, ANGGARAN_HEADERS.length).setValues([ANGGARAN_HEADERS])
    .setFontWeight('bold').setBackground('#8B6F47').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  if (baru.length) {
    sheet.getRange(2, 1, baru.length, ANGGARAN_HEADERS.length).setValues(baru);
  }
  sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 3).setNumberFormat('0');
}

function readAnggaranHeaders_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim());
}

/**
 * Menerjemahkan baris header menjadi nomor kolom.
 *
 * `valid` hanya true kalau KEEMPAT kolom (Kategori, Bulan, Tahun, Batas)
 * ditemukan DAN menunjuk empat kolom yang berbeda. Inilah kunci perbaikan:
 * pada tata letak lama, "Batas Anggaran" ada di kolom 2 sementara tidak ada
 * kolom "Bulan" sama sekali, sehingga Bulan dan Batas sama-sama menunjuk
 * kolom 2 — nilai bulan lalu tertimpa nominal batas saat menyimpan, dan
 * anggaran yang tersimpan tidak pernah lolos filter bulan.
 */
function parseAnggaranHeaderIndices_(headers) {
  let colKat = 0, colBulan = 0, colTahun = 0, colBatas = 0;

  headers.forEach((h, idx) => {
    const name = String(h || '').trim().toLowerCase();
    if (!name) return;
    if (name.includes('kategori')) colKat = idx + 1;
    else if (name.includes('bulan')) colBulan = idx + 1;
    else if (name.includes('tahun')) colTahun = idx + 1;
    else if (name.includes('batas') || name.includes('nominal')) colBatas = idx + 1;
  });

  const unik = new Set([colKat, colBulan, colTahun, colBatas]);
  const valid = colKat > 0 && colBulan > 0 && colTahun > 0 && colBatas > 0 &&
    unik.size === 4 && !unik.has(0);

  return {
    colKat: colKat || 1,
    colBulan: colBulan || 2,
    colTahun: colTahun || 3,
    colBatas: colBatas || 4,
    valid: valid
  };
}

/**
 * Selalu mengembalikan nomor kolom yang benar. Kalau tata letak sheet masih
 * versi lama (tidak valid), sheet dinormalisasi lebih dulu lalu header dibaca
 * ulang — jadi operasi baca/tulis anggaran tidak pernah memakai kolom yang
 * salah, bahkan bila normalizeAnggaranSheet_() di doGet sempat gagal.
 */
function getAnggaranHeaderIndices(sheet) {
  let idx = parseAnggaranHeaderIndices_(readAnggaranHeaders_(sheet));
  if (!idx.valid) {
    try {
      normalizeAnggaranSheet_();
      idx = parseAnggaranHeaderIndices_(readAnggaranHeaders_(sheet));
    } catch (e) {
      console.error('normalisasi sheet Anggaran gagal: ' + e.message);
    }
  }
  return idx;
}

// DITULIS ULANG — PERBAIKAN BUG UTAMA:
// Sekarang membaca nilai MENTAH langsung dari sheet (bukan lewat
// sheetToObjects yang mengonversi Date jadi string secara lossy), dan
// memakai toSafeNumber() untuk kolom Bulan/Tahun/Batas. Ini memastikan
// anggaran yang baru saja disimpan SELALU terbaca dengan angka yang benar,
// bahkan jika kolomnya sempat ter-format ulang jadi Tanggal oleh Google
// Sheets.
function getAnggaran() {
  const sheet = SS.getSheetByName(SHEET_ANGGARAN);
  if (!sheet) return [];

  const { colKat, colBulan, colTahun, colBatas } = getAnggaranHeaderIndices(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const hasil = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const kategori = String(row[colKat - 1] || '').trim();
    if (!kategori) continue;

    const bulan = toSafeNumber(row[colBulan - 1]);
    const tahun = toSafeNumber(row[colTahun - 1]);
    const batas = toSafeNumber(row[colBatas - 1]);

    // Lewati baris yang tidak punya data anggaran sama sekali (mis. baris
    // kosong sisa penghapusan) supaya tidak muncul sebagai kartu kosong.
    if (!bulan && !tahun && !batas) continue;

    hasil.push({
      _row: i + 1,
      'Kategori Pengeluaran': kategori,
      'Bulan': bulan,
      'Tahun': tahun,
      'Batas Anggaran': batas
    });
  }
  return hasil;
}

function addAnggaran(kategori, bulan, tahun, batas) {
  return withLock(() => {
    kategori = String(kategori || '').trim();
    if (!kategori) return { success: false, error: 'Kategori tidak boleh kosong' };
    const b = Number(bulan);
    const t = Number(tahun);
    const limit = Number(batas);
    if (!b || b < 1 || b > 12) return { success: false, error: 'Bulan harus 1-12' };
    if (!t || t < 2000) return { success: false, error: 'Tahun tidak valid' };
    if (!limit || limit <= 0) return { success: false, error: 'Batas anggaran harus > 0' };

    const sheet = SS.getSheetByName(SHEET_ANGGARAN);
    const { colKat, colBulan, colTahun, colBatas } = getAnggaranHeaderIndices(sheet);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const rowKat = String(data[i][colKat - 1] || '').trim().toLowerCase();
      const rowBulan = toSafeNumber(data[i][colBulan - 1]);
      const rowTahun = toSafeNumber(data[i][colTahun - 1]);

      if (rowKat === kategori.toLowerCase() && rowBulan === b && rowTahun === t) {
        sheet.getRange(i + 1, colBatas).setNumberFormat('0').setValue(limit);
        return { success: true, updated: true };
      }
    }

    const maxCol = Math.max(colKat, colBulan, colTahun, colBatas, 4);
    const newRow = new Array(maxCol).fill('');
    newRow[colKat - 1] = kategori;
    newRow[colBulan - 1] = b;
    newRow[colTahun - 1] = t;
    newRow[colBatas - 1] = limit;

    sheet.appendRow(newRow);

    // PENTING — PERBAIKAN BUG UTAMA: paksa format sel Bulan/Tahun/Batas
    // menjadi angka biasa (bukan Tanggal) tepat setelah ditulis. Ini
    // mencegah Google Sheets otomatis mengubah format kolom menjadi
    // "Date", yang sebelumnya menyebabkan anggaran baru "hilang" dari
    // daftar walau notifikasi menyatakan berhasil.
    const newRowIndex = sheet.getLastRow();
    sheet.getRange(newRowIndex, colBulan).setNumberFormat('0');
    sheet.getRange(newRowIndex, colTahun).setNumberFormat('0');
    sheet.getRange(newRowIndex, colBatas).setNumberFormat('0');

    return { success: true };
  });
}

function updateAnggaran(oldKategori, oldBulan, oldTahun, data, rowNumber) {
  return withLock(() => {
    const kategoriBaru = String(data.kategori || '').trim();
    if (!kategoriBaru) return { success: false, error: 'Kategori tidak boleh kosong' };

    // Validasi sama seperti addAnggaran(). Tanpa ini, nilai kosong/tidak valid
    // bisa tertulis sebagai NaN ke sheet dan membuat anggaran kembali "tidak
    // terlihat" — gejala yang persis sama dengan bug yang sedang diperbaiki.
    const bulanBaru = Number(data.bulan);
    const tahunBaru = Number(data.tahun);
    const batasBaru = Number(data.batas);
    if (!bulanBaru || bulanBaru < 1 || bulanBaru > 12) return { success: false, error: 'Bulan harus 1-12' };
    if (!tahunBaru || tahunBaru < 2000) return { success: false, error: 'Tahun tidak valid' };
    if (!batasBaru || batasBaru <= 0) return { success: false, error: 'Batas anggaran harus > 0' };

    const sheet = SS.getSheetByName(SHEET_ANGGARAN);
    const { colKat, colBulan, colTahun, colBatas } = getAnggaranHeaderIndices(sheet);
    const values = sheet.getDataRange().getValues();

    // Helper lokal: menulis satu baris sekaligus memaksa format angka biasa
    // untuk kolom Bulan/Tahun/Batas — mencegah bug format-Tanggal yang sama
    // seperti pada addAnggaran().
    function writeRow(rowIdx) {
      sheet.getRange(rowIdx, colKat).setValue(kategoriBaru);
      sheet.getRange(rowIdx, colBulan).setNumberFormat('0').setValue(bulanBaru);
      sheet.getRange(rowIdx, colTahun).setNumberFormat('0').setValue(tahunBaru);
      sheet.getRange(rowIdx, colBatas).setNumberFormat('0').setValue(batasBaru);
    }

    const idx = { colKat, colBulan, colTahun };
    const targetRow = findAnggaranRow_(values, idx, oldKategori, oldBulan, oldTahun, rowNumber);
    if (!targetRow) return { success: false, error: 'Anggaran tidak ditemukan di spreadsheet (mungkin sudah diubah/dihapus). Muat ulang halaman lalu coba lagi.' };

    // DIPERBAIKI: tolak kalau hasil edit bentrok dengan anggaran lain yang
    // kategori + bulan + tahunnya sama. Sebelumnya ini membuat baris ganda,
    // sehingga total anggaran terhitung dua kali.
    const bentrok = findAnggaranRow_(values, idx, kategoriBaru, bulanBaru, tahunBaru, null);
    if (bentrok && bentrok !== targetRow) {
      return { success: false, error: 'Anggaran "' + kategoriBaru + '" untuk ' + bulanBaru + '/' + tahunBaru + ' sudah ada. Edit anggaran itu saja.' };
    }

    writeRow(targetRow);
    return { success: true };
  });
}

/**
 * DIPERBAIKI: mencari baris anggaran secara aman.
 * Sebelumnya update/hapus langsung memakai nomor baris dari browser tanpa
 * dicek (kalau baris sudah bergeser karena ada yang dihapus, baris LAIN yang
 * ikut terhapus/tertimpa), dan pencarian cadangannya hanya mencocokkan
 * kategori (sehingga bisa mengenai anggaran bulan lain dengan kategori sama).
 * Sekarang nomor baris hanya dipakai kalau isinya benar-benar cocok
 * (kategori + bulan + tahun); kalau tidak, dicari ulang berdasarkan ketiganya.
 * Mengembalikan nomor baris di sheet (>= 2), atau 0 bila tidak ditemukan.
 */
function findAnggaranRow_(values, idx, kategori, bulan, tahun, rowNumber) {
  const targetKat = String(kategori || '').trim().toLowerCase();
  const b = Number(bulan);
  const t = Number(tahun);
  const cocok = row =>
    String(row[idx.colKat - 1] || '').trim().toLowerCase() === targetKat &&
    toSafeNumber(row[idx.colBulan - 1]) === b &&
    toSafeNumber(row[idx.colTahun - 1]) === t;

  const r = Number(rowNumber);
  if (r >= 2 && r <= values.length && cocok(values[r - 1])) return r;

  for (let i = 1; i < values.length; i++) {
    if (cocok(values[i])) return i + 1;
  }
  return 0;
}

function deleteAnggaran(kategori, bulan, tahun, rowNumber) {
  return withLock(() => {
    const sheet = SS.getSheetByName(SHEET_ANGGARAN);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: false, error: 'Tidak ada data anggaran' };

    const { colKat, colBulan, colTahun } = getAnggaranHeaderIndices(sheet);
    const targetRow = findAnggaranRow_(values, { colKat, colBulan, colTahun }, kategori, bulan, tahun, rowNumber);
    if (!targetRow) return { success: false, error: 'Anggaran tidak ditemukan (mungkin sudah dihapus). Muat ulang halaman lalu coba lagi.' };

    sheet.deleteRow(targetRow);
    return { success: true };
  });
}

// Satu fungsi tunggal yang mengembalikan daftar anggaran BESERTA jumlah
// pemakaiannya untuk bulan/tahun tertentu, dihitung langsung dari data
// Transaksi asli di backend. Frontend TIDAK menghitung sendiri, jadi selalu
// sinkron dengan data transaksi yang sebenarnya, dan sekarang juga kebal
// terhadap masalah format Tanggal berkat getAnggaran() yang sudah diperbaiki.
function getAnggaranDenganPemakaian(bulan, tahun) {
  try {
    bulan = Number(bulan);
    tahun = Number(tahun);
    const targetYM = tahun + '-' + String(bulan).padStart(2, '0');

    const anggaranList = getAnggaran().filter(a => Number(a['Bulan']) === bulan && Number(a['Tahun']) === tahun);
    const transaksi = sheetToObjects(SHEET_TRANSAKSI);

    // Kelompokkan SEMUA transaksi pengeluaran bulan ini per kategori
    // (kategori dicocokkan trim + lowercase, sama seperti perilaku lama,
    // tapi sekarang hanya dilakukan SATU KALI di sini).
    const perKategori = {};
    let totalPengeluaranBulanIni = 0;

    transaksi.forEach(t => {
      if (t['Jenis'] !== 'Pengeluaran') return;
      if (isInternalCategory(t['Kategori'])) return;
      if (getYearMonthKey(t['Tanggal']) !== targetYM) return;

      const nominal = Number(t['Nominal'] || 0);
      totalPengeluaranBulanIni += nominal;

      const key = String(t['Kategori'] || '').trim().toLowerCase();
      if (!perKategori[key]) perKategori[key] = { total: 0, transaksi: [] };
      perKategori[key].total += nominal;
      perKategori[key].transaksi.push(t);
    });

    const hasil = anggaranList.map(a => {
      const kategori = String(a['Kategori Pengeluaran'] || '').trim();
      const key = kategori.toLowerCase();
      const grup = perKategori[key] || { total: 0, transaksi: [] };
      const trxSorted = grup.transaksi.slice().sort((x, y) => (y['Tanggal'] > x['Tanggal'] ? 1 : (y['Tanggal'] < x['Tanggal'] ? -1 : 0)));

      return {
        row: a._row,
        kategori: kategori,
        bulan: bulan,
        tahun: tahun,
        batas: Number(a['Batas Anggaran'] || 0),
        terpakai: grup.total,
        jumlahTransaksi: trxSorted.length,
        transaksi: trxSorted
      };
    });

    // Urutkan: kategori yang paling banyak terpakai (relatif ke batasnya) di atas,
    // supaya kategori yang mendekati/melebihi anggaran langsung terlihat.
    hasil.sort((a, b) => (b.terpakai / Math.max(1, b.batas)) - (a.terpakai / Math.max(1, a.batas)));

    const totalAnggaran = hasil.reduce((s, a) => s + a.batas, 0);
    const totalTerpakaiBerAnggaran = hasil.reduce((s, a) => s + a.terpakai, 0);

    return {
      bulan: bulan,
      tahun: tahun,
      anggaran: hasil,
      totalAnggaran: totalAnggaran,
      totalTerpakaiBerAnggaran: totalTerpakaiBerAnggaran,
      totalPengeluaranBulanIni: totalPengeluaranBulanIni
    };
  } catch(e) {
    return {
      bulan: Number(bulan) || 0,
      tahun: Number(tahun) || 0,
      anggaran: [],
      totalAnggaran: 0,
      totalTerpakaiBerAnggaran: 0,
      totalPengeluaranBulanIni: 0,
      error: e.message
    };
  }
}

function getKategoriList() {
  const set = new Set();
  const anggaran = getAnggaran();
  anggaran.forEach(a => {
    const k = String(a['Kategori Pengeluaran'] || '').trim();
    if (k) set.add(k);
  });

  const trx = sheetToObjects(SHEET_TRANSAKSI);
  trx.forEach(t => {
    if (t['Jenis'] === 'Pengeluaran' && !isInternalCategory(t['Kategori'])) {
      const k = String(t['Kategori'] || '').trim();
      if (k) set.add(k);
    }
  });

  if (set.size === 0) {
    ['Makanan & Minuman', 'Transportasi', 'Tagihan', 'Belanja', 'Lainnya'].forEach(k => set.add(k));
  }
  return Array.from(set);
}

function getKategoriPemasukanList() {
  return ['Gaji', 'Bonus & THR', 'Hasil Usaha', 'Investasi / Bunga', 'Freelance', 'Hadiah', 'Pemasukan Lainnya'];
}

// ================= HUTANG & PIUTANG =================
function addHutangPiutang(data) {
  return withLock(() => {
    const nominal = Number(data.nominal);
    if (!nominal || nominal <= 0) return { success: false, error: 'Nominal harus lebih dari 0' };
    const sheet = SS.getSheetByName(SHEET_HUTANG);
    const id = generateId('HP');
    sheet.appendRow([id, data.jenis, data.pihak, nominal, data.status || 'Belum Lunas', data.tenggat]);
    return { success: true, id: id };
  });
}

function updateHutangPiutangStatus(id, status) {
  return withLock(() => {
    const sheet = SS.getSheetByName(SHEET_HUTANG);
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 5).setValue(status);
        return { success: true };
      }
    }
    return { success: false, error: 'Data tidak ditemukan' };
  });
}

function getHutangPiutang() {
  return sheetToObjects(SHEET_HUTANG);
}

// ================= TABUNGAN =================
function addTabungan(tujuan, target) {
  return withLock(() => {
    tujuan = String(tujuan || '').trim();
    const targetNum = Number(target);
    if (!tujuan) return { success: false, error: 'Nama tujuan tidak boleh kosong' };
    if (!targetNum || targetNum <= 0) return { success: false, error: 'Target harus lebih dari 0' };

    const sheet = SS.getSheetByName(SHEET_TABUNGAN);
    const id = generateId('TAB');
    sheet.appendRow([id, tujuan, targetNum, 0]);
    return { success: true, id: id };
  });
}

function setorTabungan(id, nominal, sumber) {
  return withLock(() => {
    nominal = Number(nominal);
    if (!nominal || nominal <= 0) return { success: false, error: 'Nominal harus lebih dari 0' };

    const sheet = SS.getSheetByName(SHEET_TABUNGAN);
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        // DIPERBAIKI: catat transaksi pengeluarannya DULU. Sebelumnya saldo
        // tabungan langsung ditambah, lalu hasil _addTransaksiRaw diabaikan —
        // kalau pencatatan gagal (mis. dompet kosong), saldo tabungan tetap
        // bertambah padahal uang di dompet tidak berkurang.
        const trx = _addTransaksiRaw({
          tanggal: new Date(),
          jenis: 'Pengeluaran',
          sumber: sumber,
          kategori: TABUNGAN_PREFIX + values[i][1],
          nominal: nominal,
          keterangan: 'Setoran tabungan ke "' + values[i][1] + '"'
        });
        if (!trx.success) return trx;

        const newSaldo = Number(values[i][3]) + nominal;
        sheet.getRange(i + 1, 4).setValue(newSaldo);
        return { success: true, newSaldo: newSaldo };
      }
    }
    return { success: false, error: 'Target tabungan tidak ditemukan' };
  });
}

function getTabungan() {
  return sheetToObjects(SHEET_TABUNGAN);
}

// ================= DASHBOARD & KALENDER =================
/**
 * DIPERBAIKI: totalPemasukan & totalPengeluaran sekarang HANYA menghitung
 * transaksi bulan berjalan (sebelumnya menjumlahkan seluruh riwayat sejak awal).
 *
 * BARU: sisaBulanLalu = saldo semua dompet per akhir bulan lalu
 * (saldo awal + semua pemasukan - semua pengeluaran SEBELUM bulan berjalan).
 * Nilai ini ditampilkan sebagai dompet virtual "SISA BULAN LALU" di tab
 * Anggaran. Dompet ini HANYA tampilan (dihitung otomatis, tidak ditulis ke
 * sheet), jadi saldo dompet asli (BCA, Tunai, dst.) dan fitur Balancing tidak
 * berubah. Total saldo tetap = sisa bulan lalu + arus kas bulan berjalan.
 *
 * Catatan: total pemasukan/pengeluaran bulan berjalan tetap mengecualikan
 * kategori internal (Penyesuaian Saldo & setoran Tabungan), sama seperti
 * sebelumnya. Sedangkan angka per dompet di `rincian` memasukkan semuanya,
 * karena dipakai untuk menghitung saldo dompet yang sebenarnya.
 */
function getDashboardData() {
  try {
    const transaksi = sheetToObjects(SHEET_TRANSAKSI);
    const dompets = getDompets();
    const sumberList = getAllSumberDana();

    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const currentYM = Utilities.formatDate(now, tz, 'yyyy-MM');
    const bulanBerjalan = Number(currentYM.substring(5, 7));
    const tahunBerjalan = Number(currentYM.substring(0, 4));

    const nominalOf = t => Number(t['Nominal'] || 0);
    const isBulanIni = t => getYearMonthKey(t['Tanggal']) === currentYM;
    // Perbandingan string 'yyyy-MM' aman secara leksikografis.
    const isSebelumBulanIni = t => getYearMonthKey(t['Tanggal']) < currentYM;

    const rincian = sumberList.map(sumber => {
      const sLower = String(sumber || '').trim().toLowerCase();
      const milikDompet = transaksi.filter(t => String(t['Sumber/Dompet'] || '').trim().toLowerCase() === sLower);
      const masuk = milikDompet.filter(t => t['Jenis'] === 'Pemasukan');
      const keluar = milikDompet.filter(t => t['Jenis'] === 'Pengeluaran');

      // Total sepanjang waktu (dipakai untuk saldo bersih).
      const pemasukan = masuk.reduce((s, t) => s + nominalOf(t), 0);
      const pengeluaran = keluar.reduce((s, t) => s + nominalOf(t), 0);

      // Arus kas bulan berjalan.
      const pemasukanBulanIni = masuk.filter(isBulanIni).reduce((s, t) => s + nominalOf(t), 0);
      const pengeluaranBulanIni = keluar.filter(isBulanIni).reduce((s, t) => s + nominalOf(t), 0);

      const saldoAwal = Number((dompets.find(d => String(d['Nama Dompet'] || '').trim().toLowerCase() === sLower) || {})['Saldo Awal'] || 0);
      const saldoBersih = saldoAwal + pemasukan - pengeluaran;

      // Saldo dompet ini per akhir bulan lalu.
      const sisaBulanLalu = saldoAwal
        + masuk.filter(isSebelumBulanIni).reduce((s, t) => s + nominalOf(t), 0)
        - keluar.filter(isSebelumBulanIni).reduce((s, t) => s + nominalOf(t), 0);

      return {
        sumber: sumber,
        pemasukan: pemasukan,
        pengeluaran: pengeluaran,
        pemasukanBulanIni: pemasukanBulanIni,
        pengeluaranBulanIni: pengeluaranBulanIni,
        saldoAwal: saldoAwal,
        saldoBersih: saldoBersih,
        sisaBulanLalu: sisaBulanLalu
      };
    });

    const totalSaldo = rincian.reduce((s, r) => s + r.saldoBersih, 0);
    const sisaBulanLalu = rincian.reduce((s, r) => s + r.sisaBulanLalu, 0);

    const transaksiAsliBulanIni = transaksi.filter(t => !isInternalCategory(t['Kategori']) && isBulanIni(t));
    const totalPemasukan = transaksiAsliBulanIni.filter(t => t['Jenis'] === 'Pemasukan').reduce((s, t) => s + nominalOf(t), 0);
    const totalPengeluaran = transaksiAsliBulanIni.filter(t => t['Jenis'] === 'Pengeluaran').reduce((s, t) => s + nominalOf(t), 0);

    return {
      totalSaldo: totalSaldo,
      totalPemasukan: totalPemasukan,
      totalPengeluaran: totalPengeluaran,
      sisaBulanLalu: sisaBulanLalu,
      bulanBerjalan: bulanBerjalan,
      tahunBerjalan: tahunBerjalan,
      rincian: rincian
    };
  } catch(e) {
    return { totalSaldo: 0, totalPemasukan: 0, totalPengeluaran: 0, sisaBulanLalu: 0, rincian: [] };
  }
}

// DIPERBAIKI: filter bulan/tahun sekarang membandingkan string 'yyyy-MM'
// langsung, bukan membungkus ulang ke objek Date (sumber bug pergeseran bulan).
function getPieChartData(bulan, tahun) {
  try {
    const transaksi = sheetToObjects(SHEET_TRANSAKSI);
    const targetYM = (bulan && tahun) ? (tahun + '-' + String(bulan).padStart(2, '0')) : null;

    const filtered = transaksi.filter(t => {
      if (t['Jenis'] !== 'Pengeluaran') return false;
      if (isInternalCategory(t['Kategori'])) return false;
      if (!targetYM) return true;
      return getYearMonthKey(t['Tanggal']) === targetYM;
    });
    const map = {};
    filtered.forEach(t => {
      const k = t['Kategori'] || 'Lainnya';
      map[k] = (map[k] || 0) + Number(t['Nominal'] || 0);
    });
    return { labels: Object.keys(map), values: Object.values(map) };
  } catch(e) {
    return { labels: [], values: [] };
  }
}

/**
 * Mengembalikan object { filled, pengeluaran }
 * - filled: array tanggal (yyyy-MM-dd) yang punya transaksi apa pun
 * - pengeluaran: map tanggal (yyyy-MM-dd) -> total nominal pengeluaran
 *   (kategori internal seperti Penyesuaian Saldo & Setoran Tabungan dikecualikan)
 *
 * DIPERBAIKI: sebelumnya dihitung ulang lewat `new Date(t['Tanggal'])` yang
 * rawan salah zona waktu. Sekarang murni perbandingan string, jadi selalu
 * konsisten dengan tanggal yang tersimpan di sheet Transaksi.
 */
function getKalenderData(bulan, tahun) {
  try {
    const transaksi = sheetToObjects(SHEET_TRANSAKSI);
    const targetYM = tahun + '-' + String(bulan).padStart(2, '0');
    const filled = new Set();
    const pengeluaranPerTanggal = {};

    transaksi.forEach(t => {
      const tgl = String(t['Tanggal'] || '');
      if (tgl.substring(0, 7) !== targetYM) return;

      filled.add(tgl);
      if (t['Jenis'] === 'Pengeluaran' && !isInternalCategory(t['Kategori'])) {
        pengeluaranPerTanggal[tgl] = (pengeluaranPerTanggal[tgl] || 0) + Number(t['Nominal'] || 0);
      }
    });

    return { filled: Array.from(filled), pengeluaran: pengeluaranPerTanggal };
  } catch(e) {
    return { filled: [], pengeluaran: {} };
  }
}

function balancingSaldo(sumber, saldoAktualBaru) {
  return withLock(() => {
    const dashboard = getDashboardData();
    const rincian = dashboard.rincian.find(r => r.sumber === sumber);
    if (!rincian) return { success: false, error: 'Dompet tidak ditemukan' };

    const selisih = Number(saldoAktualBaru) - rincian.saldoBersih;

    if (Math.abs(selisih) > 0.01) {
      _addTransaksiRaw({
        tanggal: new Date(),
        jenis: selisih > 0 ? 'Pemasukan' : 'Pengeluaran',
        sumber: sumber,
        kategori: ADJUSTMENT_CATEGORY,
        nominal: Math.abs(selisih),
        keterangan: 'Adjustment balancing saldo'
      });
    }

    const sheet = SS.getSheetByName(SHEET_DOMPET);
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === sumber) {
        sheet.getRange(i + 1, 3).setValue(Number(saldoAktualBaru));
        break;
      }
    }

    return { success: true, selisih: selisih };
  });
}
