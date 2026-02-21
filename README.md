# ⚡ FileShift — Universal File Converter

Konverter file universal yang berjalan di **Netlify** (serverless).  
Konversi .MSG, .EML, .DOCX, .HTML, .CSV, .JSON, .TXT, .MD → PDF, DOCX, TXT, HTML.

## 🏗 Arsitektur

```
fileshift/
├── public/
│   └── index.html          ← Frontend (static)
├── netlify/
│   └── functions/
│       └── convert.js      ← Serverless backend (Netlify Function)
├── package.json             ← Dependencies
├── netlify.toml             ← Netlify config
└── README.md
```

**Frontend** → `public/index.html` (di-serve sebagai static file)  
**Backend**  → `netlify/functions/convert.js` (Netlify Function, dipanggil via `/api/convert`)

## 📦 Library yang Digunakan (Konversi Akurat)

| Library      | Fungsi                                      |
|-------------|---------------------------------------------|
| `msgreader` | Parsing file .MSG (Outlook) secara akurat   |
| `mailparser`| Parsing file .EML dengan MIME support penuh |
| `mammoth`   | Ekstrak teks & HTML dari .DOCX              |
| `pdfkit`    | Generate PDF berkualitas tinggi              |
| `docx`      | Generate file .DOCX yang valid              |
| `busboy`    | Parsing multipart form upload               |

## 🚀 Cara Deploy ke Netlify

### Opsi 1: Deploy via GitHub (Recommended)

1. **Push ke GitHub:**
   ```bash
   cd fileshift
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/USERNAME/fileshift.git
   git push -u origin main
   ```

2. **Hubungkan ke Netlify:**
   - Buka https://app.netlify.com
   - Klik **"Add new site"** → **"Import an existing project"**
   - Pilih repo GitHub kamu
   - Settings:
     - **Build command:** `npm install` 
     - **Publish directory:** `public`
     - **Functions directory:** `netlify/functions`
   - Klik **"Deploy site"**

3. **Selesai!** Website akan live di `https://your-site.netlify.app`

### Opsi 2: Deploy via Netlify CLI

```bash
# Install Netlify CLI
npm install -g netlify-cli

# Login
netlify login

# Masuk ke folder project
cd fileshift

# Install dependencies
npm install

# Deploy (preview dulu)
netlify deploy

# Deploy ke production
netlify deploy --prod
```

### Opsi 3: Drag & Drop

1. Jalankan `npm install` dulu di lokal untuk install dependencies
2. Buka https://app.netlify.com/drop
3. Drag folder `fileshift` ke halaman tersebut

> ⚠️ **Penting:** Opsi drag & drop mungkin tidak include `node_modules` untuk functions.  
> Gunakan Opsi 1 atau 2 untuk hasil terbaik.

## 🧪 Test Lokal

```bash
# Install dependencies
npm install

# Jalankan dev server (butuh Netlify CLI)
npx netlify dev

# Buka http://localhost:8888
```

## 📋 Format yang Didukung

### Input
| Format | Deskripsi |
|--------|-----------|
| `.msg`  | Microsoft Outlook Email |
| `.eml`  | Standard Email Format |
| `.docx` | Microsoft Word Document |
| `.html` | Web Page |
| `.csv`  | Comma-Separated Values |
| `.json` | JSON Data |
| `.xml`  | XML Document |
| `.txt`  | Plain Text |
| `.md`   | Markdown |
| `.rtf`  | Rich Text Format |

### Output
| Format | Library | Kualitas |
|--------|---------|----------|
| **PDF**  | pdfkit  | ✅ High — layout proper, unicode support |
| **DOCX** | docx    | ✅ High — styled paragraphs, proper format |
| **TXT**  | native  | ✅ Clean text extraction |
| **HTML** | native  | ✅ Styled HTML with email headers |

## ⚙️ Cara Kerja

1. User upload file via browser
2. File dikirim ke Netlify Function (`/api/convert`) via `FormData`
3. Function server-side:
   - Parse file menggunakan library yang sesuai (msgreader, mailparser, dll)
   - Ekstrak konten (subject, from, to, body untuk email)
   - Generate output (PDF via pdfkit, DOCX via docx library)
4. File hasil dikirim balik ke browser sebagai download

## 🔒 Privasi

- **Tidak ada database** — file tidak disimpan
- **Tidak ada storage** — file langsung diproses dan dikembalikan
- **Serverless** — function hanya jalan saat ada request
- File diproses di memory dan langsung dihapus setelah response dikirim

## 📝 Limits (Netlify Free Tier)

- **Function timeout:** 10 detik (26 detik di Pro)
- **Payload size:** ~6 MB (base64 encoded)
- **Invocations:** 125K/bulan (free tier)
- Untuk file besar (>4MB), pertimbangkan upgrade ke Netlify Pro

## 🛠 Troubleshooting

**Function not found (404):**
- Pastikan `netlify.toml` ada di root project
- Pastikan `npm install` sudah dijalankan

**File terlalu besar:**
- Netlify Functions punya limit ~6MB payload
- Kompres file sebelum upload, atau upgrade plan

**MSG parsing gagal:**
- Beberapa file .MSG yang sangat lama mungkin menggunakan format berbeda
- Function akan fallback ke basic text extraction
