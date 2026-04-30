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

Mo `.env` va dien API key moi:

```text
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4.1-mini
GEMINI_API_KEY=
GEMINI_MODEL=gemini-2.5-flash
GEMINI_FALLBACK_MODEL=gemini-2.0-flash-lite
LIBRARY_ADMIN_USER=admin
LIBRARY_ADMIN_PASSWORD=mat-khau-manh-cua-ban
APP_STORAGE_DIR=
PORT=3000
```

Neu dung Gemini thay OpenAI, dien `GEMINI_API_KEY`. Khi co `GEMINI_API_KEY`, server se uu tien Gemini.

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

Hoi tro ly AI:

```http
POST /api/ai
Content-Type: application/json

{
  "question": "Kiem tra tong hien trang va quy hoach",
  "context": {}
}
```

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

## Thu vien so PDF

- Trang chu co menu `Thu vien tai lieu`.
- File PDF duoc luu trong `protected_uploads/`, khong nam trong `public/`.
- Admin dang nhap bang `LIBRARY_ADMIN_USER` va `LIBRARY_ADMIN_PASSWORD`.
- Upload tai lieu bang giao dien quan tri: ten tai lieu, tac gia/don vi, nam, mo ta, danh muc, PDF va anh bia.
- Trinh doc PDF dung PDF.js de render len canvas, khong hien nut tai xuong/in va co watermark tren trang doc.

Luu y bao mat: tren moi truong web khong the chong tai xuong/copy/chup man hinh 100%. Phan mem chi han che nguoi dung pho thong bang cach khong public file goc, dung token ngan han, khong tao text layer, chan chuot phai va cac phim tat pho bien.

Ghi chu: server uu tien chay bang Express sau khi `npm install`. Neu moi truong chua cai duoc `express`, server co fallback Node thuan de van kiem tra duoc frontend va API co ban, nhung ban nen chay `npm install` de dung dung backend Express.
