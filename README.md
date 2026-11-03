# FinTrack — Financial Tracker & Planner

Aplikasi pencatat keuangan pribadi berbasis **Google Apps Script** (GAS) dengan
**Google Sheets** sebagai database.

## Isi repo

| File di repo       | Nama file di editor GAS | Isi                                   |
|--------------------|-------------------------|---------------------------------------|
| `Code.gs`          | `Code.gs`               | Backend (baca/tulis spreadsheet)      |
| `Index.html`       | `Index`                 | Struktur halaman & semua modal        |
| `Stylesheet.html`  | `Stylesheet`            | CSS tema *earth tone*                 |
| `JavaScript.html`  | `JavaScript`            | Logika frontend                       |

> Nama file HTML di editor GAS **harus** `Index`, `Stylesheet`, dan `JavaScript`
> (tanpa `.html`, huruf besar/kecil persis), karena dipanggil lewat
> `include('Stylesheet')` dan `include('JavaScript')`.

## Cara deploy / update di Google Apps Script

1. Buka spreadsheet FinTrack → **Ekstensi → Apps Script**.
2. Untuk tiap file di tabel atas: buka file dengan nama yang sama di editor,
   **hapus semua isinya**, lalu tempel isi file dari repo ini.
   Kalau filenya belum ada, buat lewat tombol **+** (Script untuk `Code`,
   HTML untuk tiga lainnya).
3. Klik **Simpan** (ikon disket / `Ctrl+S`).
4. **Terapkan → Kelola penerapan** → pilih penerapan yang sudah ada → ikon
   pensil (**Edit**) → *Versi*: **Versi baru** → **Terapkan**.
   Dengan cara ini URL aplikasi web tetap sama.
5. Buka URL aplikasi web dan muat ulang halaman (kalau perlu `Ctrl+Shift+R`).

Sheet yang dibutuhkan (`Transaksi`, `Dompet`, `Anggaran`, `Hutang_Piutang`,
`Tabungan`) dibuat otomatis saat aplikasi pertama kali dibuka.

## Catatan fitur

- **Pemasukan / Pengeluaran** di kartu atas hanya menghitung **bulan berjalan**.
  Total Saldo tetap menghitung seluruh saldo.
- **SISA BULAN LALU** (tab Anggaran → Kelola Dompet) adalah dompet *virtual*:
  saldo semua dompet per akhir bulan lalu, dihitung otomatis. Tidak tersimpan
  di sheet dan tidak memindahkan uang, sehingga saldo dompet asli & fitur
  Balancing tidak terpengaruh. Nama ini tidak bisa dipakai untuk dompet biasa.
- Kartu dompet menampilkan: *sisa bulan lalu + masuk bulan ini − keluar bulan
  ini = saldo sekarang*.
- Hutang/Piutang bisa dicatat, diedit, dihapus, dan ditandai lunas.
