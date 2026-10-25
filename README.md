# IndimediaWifi

Portal publik untuk informasi status jaringan, login pelanggan, dan akses admin IndimediaWiFi.

## Stack
- React + TypeScript + Vite
- Firebase SDK tersedia untuk tahap integrasi berikutnya
- Lucide React untuk ikon antarmuka

## Development
1. Gunakan Node.js versi LTS.
2. Jalankan `npm install`.
3. Jalankan `npm run dev` untuk pengembangan.
4. Jalankan `npm run build` untuk memeriksa build produksi.

## Status implementasi
- Halaman publik dan tema responsif: prototipe UI.
- Login pelanggan dengan ID: form UI saja, belum membaca data pelanggan.
- Login admin: form UI saja, belum melakukan autentikasi.
- Status server: belum terhubung ke monitoring real-time; ditampilkan sebagai "Belum diketahui".
- Tagihan, pembayaran, upload QRIS, import/export, dan backup: belum diimplementasikan.

## Keamanan
- ID pelanggan saja bukan bukti kepemilikan yang kuat. Jangan menampilkan tagihan sebelum alur verifikasi akses dirancang.
- Jangan membuka pendaftaran admin publik.
- Jangan commit file rahasia atau service-account key.
- Sebelum memakai proyek Firebase yang sama dengan Amira Kasir Pro, audit dan uji Firestore Security Rules terlebih dahulu. Jangan mengubah rules produksi kasir untuk menyalakan UI ini.
- Jangan menganggap status server aktif tanpa data monitoring nyata.
