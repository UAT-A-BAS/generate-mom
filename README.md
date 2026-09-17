# MOM Generator

MOM Generator dapat diakses dengan link https://generate-mom.apps.ocpdevgra.dti.co.id/
Download `ExportMOMToDraft.bas` untuk dapat mengintegrasikan flow dengan Outlook.

Lebih detailnya ada pada file Panduan menggunakan MOM Generator.docx

## HTML offline lokal

Download `generate-mom-offline.html` dan buka langsung di browser. Satu file ini sudah memuat aplikasi dan Macro Outlook, tanpa server atau koneksi internet. Kolaborasi dinonaktifkan; status tetap Personal Draft / Offline. File ini dibuat dari `index.html` dan `ExportMOMToDraft.bas`, sehingga perubahan dilakukan pada sumbernya.

- Buat ulang setelah perubahan: `node tools/build-offline-html.mjs`
- Periksa kesesuaian artefak dengan sumber: `node tools/build-offline-html.mjs --check`
- Aktifkan pembaruan otomatis sebelum commit, sekali per clone: `sh tools/install-git-hooks.sh`
- Uji artefak: `node tests/offline-artifact.test.cjs`
- Buktikan runtime lokal tanpa jaringan: `node tools/verify-offline-runtime.mjs` (memerlukan Playwright dan Chromium yang telah terpasang).

Hook membuat ulang dan memasukkan artefak ke commit; `git push` mengunggahnya bersama perubahan lain ke GitHub. Hook membaca sumber di working tree, jadi stage perubahan sumber terkait sebelum commit. Artefak tidak diunggah otomatis hanya dengan mengedit file.

## Kolaborasi realtime

Kolaborasi aktif saat URL memuat parameter `?session=<id>`. Tombol Start Collab membuat sesi baru dan Copy Share Link membagikan tautannya.

Susunan layanan:

- `index.html` di-host pada Cloudflare Pages `https://generate-mom.pages.dev/`.
- `functions/api/collab/[sessionId].js` mem-proxy WebSocket, permintaan POST, dan GET ke Worker.
- `worker/index.mjs` menjalankan Durable Object `MomCollabSession` yang menyimpan draft kanonik per sesi.

Protokol sinkronisasi:

- Urutan pesan ditentukan server. Klien hanya mengirim `baseVersion` berisi urutan terakhir yang benar-benar dilihat, dan tidak pernah menebak nomor urut berikutnya.
- Setiap perubahan dikirim per-field (`ops`), bukan penggantian seluruh draft. Dua editor yang mengubah kolom berbeda tidak saling menimpa.
- Penulisan pada field yang sama mengikuti aturan last-write-wins memakai urutan server, sehingga kedua sisi berakhir pada nilai yang sama.
- Perubahan struktur (tambah/hapus/pindah baris) dikirim sebagai snapshot array terkecil yang berubah, bukan seluruh draft.
- Klien mendeteksi celah urutan, meminta `resync`, dan server membalas dengan delta operasi atau snapshot penuh.
- Heartbeat `ping`/`pong` memutus koneksi setengah terbuka agar editor tidak mengetik ke koneksi yang sudah mati.
- Saat tab ditutup, perubahan yang belum terkirim di-flush lewat `sendBeacon`/`fetch keepalive`.
- Perubahan yang ditolak server (misalnya baris sudah dihapus editor lain) muncul sebagai penanda `Perlu dicek` pada panel kolaborasi, bukan hilang diam-diam.

### Siklus hidup sesi dan efisiensi Worker

Worker memakai WebSocket Hibernation API (`state.acceptWebSocket`). Selama ruangan tidak ada aktivitas, Durable Object dikeluarkan dari memori dan penagihan durasi berhenti, sementara koneksi editor tetap terbuka; pesan berikutnya membangunkan object dan draft dimuat ulang dari storage. Identitas editor disimpan sebagai socket attachment agar tetap ada setelah hibernasi.

Ada dua lapis penangguhan:

- Klien: setelah 5 menit tanpa aktivitas, atau 60 detik saat tab disembunyikan, socket ditutup dan sesi masuk mode `Disconnected`. Kembali aktif akan menyambung ulang otomatis.
- Server: hibernasi otomatis saat tidak ada lalu lintas pesan, jadi biaya durasi tidak berjalan hanya karena ada editor yang membuka halaman.

Draft kanonik disimpan di `storage` pada setiap perubahan, sehingga sesi tetap utuh walau seluruh editor menutup tab dan dibuka kembali nanti. Tidak ada masa kedaluwarsa otomatis: sesi bertahan sampai dihapus manual.

### Deploy

```sh
# Worker kolaborasi
cd worker && npx wrangler deploy

# Situs Cloudflare Pages
node tools/build-pages-dist.mjs
npx wrangler pages deploy dist-pages --project-name=generate-mom --branch=main
```

Folder `dist-pages/` hanya berisi berkas publik (`index.html`, `generate-mom-offline.html`, macro Outlook, panduan, `_headers`, dan `functions/`). Jalankan `node tools/build-offline-html.mjs` terlebih dahulu agar artefak offline yang ikut terunggah sudah terbaru.

### Pengujian

```sh
for f in tests/*.cjs; do node "$f"; done
```

- `tests/collab-protocol-v2.test.cjs` menguji urutan otorisasi server, last-write-wins per field, penolakan baris hantu, resync, dan beacon.
- `tests/collab-concurrency.test.cjs` menguji antrean keluar: kegagalan kirim tidak boleh membuang perubahan.
- `tests/offline-artifact.test.cjs` memastikan artefak offline selalu identik dengan hasil build terbaru.
