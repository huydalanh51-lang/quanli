# Phan Mem Chu Chuyen Dat Dai

Web dong Node.js + Express cho bang chu chuyen dat dai.

## Cau truc thu muc

- `public/`: frontend HTML, CSS, JavaScript
- `server/`: backend Express
- `data/`: SQLite database
- `uploads/`: file tai len
- `exports/`: file xuat ra

## Yeu cau

- Node.js 24 tro len, vi backend dung SQLite tich hop san trong Node (`node:sqlite`).
- npm.

## Chay du an

Tao file `.env` tu mau:

```bash
copy .env.example .env
```

Mo `.env` va dien API key AI. Chi can dien mot trong hai: `OPENAI_API_KEY` hoac `GEMINI_API_KEY`; khong nhung key vao file frontend.

```text
OPENAI_API_KEY=
OPENAI_MODEL=gpt-4.1-mini
GEMINI_API_KEY=
GEMINI_MODEL=gemini-2.5-flash
GEMINI_FALLBACK_MODEL=gemini-2.0-flash-lite
LIBRARY_ADMIN_USER=admin
LIBRARY_ADMIN_PASSWORD=mat-khau-manh-cua-ban
LIBRARY_GUEST_USER=khach
LIBRARY_GUEST_PASSWORD=khach
LIBRARY_GUEST_USER_01=khach01
LIBRARY_GUEST_PASSWORD_01=mat-khau-khach-01
LIBRARY_GUEST_USER_02=khach02
LIBRARY_GUEST_PASSWORD_02=mat-khau-khach-02
WEBGIS_ADMIN_USER=admin
WEBGIS_ADMIN_PASSWORD=mat-khau-webgis-manh
APP_STORAGE_DIR=
SUPABASE_URL=
SUPABASE_SERVICE_ROLE_KEY=
SUPABASE_BUCKET=library-documents
PORT=3000
```

Neu dien ca hai key, server se uu tien Gemini. Co the kiem tra trang thai cau hinh tai `GET /api/ai/status`.

```bash
npm install
npm start
```

Sau do mo:

```text
http://127.0.0.1:3000
```

## API

Luu du lieu du an:

```http
PUT /api/projects/default
Content-Type: application/json

{
  "data": {
    "D7": "10,50"
  }
}
```

Doc lai du lieu du an:

```http
GET /api/projects/default
```

Kiem tra server:

```http
GET /api/health
```

Kiem tra cau hinh AI:

```http
GET /api/ai/status
```

Hoi tro ly AI:

```http
POST /api/ai
Content-Type: application/json

{
  "question": "Kiem tra tong hien trang va quy hoach",
  "context": {}
}
```

Trong giao dien, nut `Tro ly AI` co trong ca module Chu chuyen dat dai va WebGIS. Khi dung o WebGIS, AI se nhan them danh sach layer dang bat, so doi tuong da nap va thuoc tinh cua doi tuong dang chon theo cau hinh hien thi cua admin.

WebGIS da tich hop bang mau ky hieu hien trang su dung dat trich tu `Chucnangsudungdat.style` theo TT08. Khi import GeoJSON/shapefile da chuyen doi, phan mem tu nhan dien ma dat tu cac truong nhu `Loaidat`, `loai_dat`, `ma_dat`, `MaLoaiDat`, `ky_hieu`, `MDSD` va to mau doi tuong theo ma dat.

Database duoc tao tu dong tai `data/projects.sqlite`.

## Luu du lieu ben vung tren Render

Render co filesystem tam. Neu khong gan Persistent Disk, database SQLite va file PDF upload se mat khi deploy hoac restart.

De giu du lieu PDF va SQLite:

1. Vao service tren Render, mo trang `Disks`.
2. Them Persistent Disk, mount path nen dat la `/var/data`.
3. Vao `Environment`, them bien:

```text
APP_STORAGE_DIR=/var/data
```

4. Chon `Save, rebuild, and deploy`.

Khi co `APP_STORAGE_DIR`, phan mem se luu:

- SQLite: `/var/data/data/projects.sqlite`
- PDF bao ve: `/var/data/protected_uploads/pdf`
- Anh bia: `/var/data/protected_uploads/covers`
- Upload/xuat tam: `/var/data/uploads`, `/var/data/exports`

Luu y: du lieu da bi xoa tren filesystem tam cua Render thuong khong the khoi phuc neu truoc do chua gan Persistent Disk hoac chua co ban sao luu.

## Luu thu vien PDF bang Supabase Storage

Neu khong muon nang cap Render de dung Persistent Disk, co the dung Supabase Storage private bucket cho thu vien PDF.

1. Tao project tren Supabase.
2. Vao `Storage`, tao bucket private ten `library-documents`.
3. Vao `Project Settings` > `API`, lay `Project URL` va `service_role key`.
4. Tren Render, them bien moi truong:

```text
SUPABASE_URL=https://...supabase.co
SUPABASE_SERVICE_ROLE_KEY=...
SUPABASE_BUCKET=library-documents
```

Khi co cac bien nay, phan mem se luu metadata thu vien tai `_metadata/library_documents.json` trong bucket, PDF trong thu muc `pdf/`, va anh bia trong `covers/`. File khong nam trong `public/` va backend van yeu cau token ngan han khi doc PDF.

## Thu vien so PDF

- Trang chu co menu `Thu vien tai lieu`.
- File PDF duoc luu trong `protected_uploads/`, khong nam trong `public/`.
- Admin dang nhap bang `LIBRARY_ADMIN_USER` va `LIBRARY_ADMIN_PASSWORD`.
- Tai khoan khach dang nhap bang `LIBRARY_GUEST_USER` va `LIBRARY_GUEST_PASSWORD` de chi doc tai lieu. Neu khong cau hinh, backend dung mac dinh `khach` / `khach`.
- Co the tao nhieu tai khoan khach bang cac cap bien danh so: `LIBRARY_GUEST_USER_01` + `LIBRARY_GUEST_PASSWORD_01`, `LIBRARY_GUEST_USER_02` + `LIBRARY_GUEST_PASSWORD_02`,... Tren Render khong duoc tao trung key, nen moi tai khoan khach can dung mot so rieng.
- Upload tai lieu bang giao dien quan tri: ten tai lieu, tac gia/don vi, nam, mo ta, danh muc, PDF va anh bia.
- Trinh doc PDF dung PDF.js de render len canvas, khong hien nut tai xuong/in va co watermark tren trang doc.

Luu y bao mat: tren moi truong web khong the chong tai xuong/copy/chup man hinh 100%. Phan mem chi han che nguoi dung pho thong bang cach khong public file goc, dung token ngan han, khong tao text layer, chan chuot phai va cac phim tat pho bien.

## WebGIS quan ly du lieu dat dai

- Trang chu co menu `WebGis`.
- Ban dau dung Leaflet.js + GeoJSON, chay ngay trong frontend.
- Co ban do nen OpenStreetMap, anh ve tinh Esri va dia hinh OpenTopoMap.
- Co cac lop mau: ranh gioi hanh chinh, hien trang su dung dat, quy hoach su dung dat, giao thong, thuy he, thua dat, cong trinh cong cong.
- Co tra cuu theo ma thua, chu su dung, ma loai dat, dia danh va quy hoach.
- Co popup thong tin, bang thuoc tinh, highlight dong/doi tuong, do khoang cach, do dien tich, in ban do, chup anh ban do va hien toa do con tro.
- Co file du lieu mau tai `public/webgis/sample-land-data.geojson`.
- Admin WebGIS dang nhap bang `WEBGIS_ADMIN_USER` va `WEBGIS_ADMIN_PASSWORD`. Neu chua cau hinh rieng, backend tu dung lai `LIBRARY_ADMIN_USER` va `LIBRARY_ADMIN_PASSWORD`.
- Khi upload GeoJSON hoac sua thuoc tinh doi tuong, nguoi dung phai dang nhap admin WebGIS. Cac thao tac quan tri se tu dong luu qua API `/api/webgis/webgis-default`; neu da cau hinh `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`, `SUPABASE_BUCKET` thi du lieu duoc luu vao Supabase Storage tai `_webgis/webgis-default.json`. Neu chua cau hinh Supabase thi fallback ve SQLite, va van co ban luu tam trong trinh duyet.
- Moi layer WebGIS co metadata quan ly hien thi gom `is_public`, `default_visible`, `allow_user_toggle`, `opacity`, `sort_order` va `category`. Admin co the chinh cac truong nay trong khung `Quan tri du lieu`; thay doi duoc cap nhat truc tiep qua API `PATCH /api/webgis/:id/layers/:layerId`.
- Admin co the xoa layer bang API `DELETE /api/webgis/:id/layers/:layerId`. Khi xoa layer mac dinh cua phan mem, id layer duoc luu vao `deletedLayerIds` de layer khong tu hien lai sau khi tai lai trang.
- Admin co the chon `visible_fields` cho tung layer de quyet dinh thuoc tinh nao duoc hien trong popup, panel thong tin va bang thuoc tinh cua nguoi xem.
- Trang WebGIS chi hien layer `is_public=true`. Layer chi duoc tai GeoJSON khi duoc bat hoac khi admin dat `default_visible=true`, sau do frontend cache lai de bat/tat nhanh hon.
- Nguoi xem khong dang nhap van duoc xem/tra cuu ban do. Neu ho doi bat/tat layer hoac do trong suot, thay doi chi luu tam tren trinh duyet va khong ghi de du lieu server.

Huong nang cap: khi du lieu lon, nen chuyen GeoJSON sang backend Node.js + PostgreSQL/PostGIS, phan trang/loc theo bbox, hoac tao vector tile de ban do nhe hon. Voi nhieu diem, nen dung clustering hoac tile point layer.

Ghi chu: server uu tien chay bang Express sau khi `npm install`. Neu moi truong chua cai duoc `express`, server co fallback Node thuan de van kiem tra duoc frontend va API co ban, nhung ban nen chay `npm install` de dung dung backend Express.
