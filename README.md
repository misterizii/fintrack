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

- Kartu atas (**Saldo, Pemasukan, Pengeluaran, Ditabung**) hanya menghitung **bulan
  berjalan**: Saldo = Pemasukan − Pengeluaran − Ditabung (bulan ini). Saldo
  awal dompet dan bulan-bulan sebelumnya tidak dihitung. Setoran tabungan
  ditampilkan terpisah (bukan pengeluaran) tapi mengurangi saldo, karena
  uangnya sudah disisihkan. Penyesuaian Saldo (Balancing) tidak dihitung.
- Tab **Tabungan** punya ringkasan sendiri: ditabung bulan ini, total
  terkumpul, dan progres terhadap total target.
- Fitur **Balancing** tetap memakai saldo dompet sebenarnya (seluruh riwayat).
- Kartu dompet di tab Anggaran menampilkan *saldo awal + seluruh pemasukan −
  seluruh pengeluaran = saldo sekarang*.
- Hutang/Piutang bisa dicatat, diedit, dihapus, dan ditandai lunas.
