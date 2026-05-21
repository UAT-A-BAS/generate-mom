# MOM Generator

MOM Generator dapat diakses dengan link https://generate-mom.pages.dev/

Download `ExportMOMToDraft.bas` untuk dapat mengintegrasikan flow dengan Outlook.

Lebih detailnya ada pada file Panduan menggunakan MOM Generator.

Single-page MOM Generator ini mendukung export XLSX, export Outlook HTML, local draft JSON, dan real-time collaboration via Cloudflare Pages + Worker + Durable Objects.

## Jalankan Lokal

Static-only:

```powershell
npx wrangler pages dev .
```

Collab lokal:

```powershell
cd worker
npx wrangler dev
```

Terminal kedua:

```powershell
npx wrangler pages dev . --do MOM_COLLAB_SESSIONS=MomCollabSession@generate-mom-collab-worker-dev-staging
```

Buka URL Pages lokal, klik `Start Collab`, lalu buka share link di browser lain.

## Deploy ke GitHub

Staging dipush ke branch `dev-staging` supaya production branch tidak berubah:

```powershell
git switch dev-staging
git add index.html functions worker wrangler.toml README.md tests
git commit -m "Add MOM real-time collaboration"
git push -u origin dev-staging
```

## Deploy Cloudflare Pages

Cloudflare Pages Git integration akan membuat Preview Deployment untuk branch non-production. Untuk `dev-staging`, URL biasanya:

- `https://dev-staging.<project>.pages.dev`
- `https://<hash>.<project>.pages.dev`

Jika preview branch belum aktif: Cloudflare dashboard > Workers & Pages > Pages project > Settings > Builds & deployments > Branch deployment controls > Preview branch > include `dev-staging`.

Docs:

- [Cloudflare Pages Preview deployments](https://developers.cloudflare.com/pages/configuration/preview-deployments/)
- [Cloudflare Pages Branch deployment controls](https://developers.cloudflare.com/pages/configuration/branch-build-controls/)

## Deploy Worker / Durable Objects

Deploy Worker staging terpisah:

```powershell
cd worker
npx wrangler deploy
```

Root `wrangler.toml` sudah bind Pages Function ke Worker staging:

```toml
[[durable_objects.bindings]]
name = "MOM_COLLAB_SESSIONS"
class_name = "MomCollabSession"
script_name = "generate-mom-collab-worker-dev-staging"
```

Untuk production nanti, buat Worker production terpisah dan ubah `script_name` atau binding dashboard agar tidak memakai storage staging.

Docs:

- [Cloudflare Pages Durable Object bindings](https://developers.cloudflare.com/pages/functions/bindings/#durable-objects)
- [Cloudflare Durable Objects WebSockets](https://developers.cloudflare.com/durable-objects/best-practices/websockets/)

## Test Collab 2 Browser

1. Browser A buka Pages preview branch `dev-staging`.
2. Isi `Nama Editor`.
3. Klik `Start Collab`.
4. URL berubah menjadi `?session=<sessionId>`.
5. Klik `Copy Share Link`.
6. Browser B buka link itu.
7. Edit field di Browser A; Browser B menerima update.
8. Add/remove/move row; browser lain menerima full draft terbaru.
9. Tutup salah satu browser; `Users: n` turun.
10. Matikan koneksi; status berubah `Disconnected`.

## Compatibility

- `Save Draft Data` tetap export JSON lokal.
- `Load Draft Data` tetap import JSON lokal.
- `Preview Table`, `Preview Result`, `Export XLSX`, dan `Export to Outlook` tetap memakai data form yang sama.
- `Clear All` hanya clear form lokal, tidak menghapus Durable Object session.
