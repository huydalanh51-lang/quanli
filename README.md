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

Ghi chu: server uu tien chay bang Express sau khi `npm install`. Neu moi truong chua cai duoc `express`, server co fallback Node thuan de van kiem tra duoc frontend va API co ban, nhung ban nen chay `npm install` de dung dung backend Express.
