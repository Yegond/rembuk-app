// Server web sederhana: form input -> laporan Word + PDF.
// Hanya memakai modul bawaan Node + paket "docx".
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFile } = require('child_process');
const { buatLaporan, BULAN } = require('./laporan');

const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
const PORT = Number(process.env.PORT || CFG.port || 3000);
// Di hosting (Render dkk) wajib dengar di semua alamat (0.0.0.0).
// Di komputer sendiri dipakai 127.0.0.1 supaya cocok dengan ngrok di Windows.
const HOST = process.env.HOST || (process.env.RENDER || process.env.PORT ? '0.0.0.0' : '127.0.0.1');
const OUT = path.join(__dirname, 'hasil');
fs.mkdirSync(OUT, { recursive: true });

// Kode akses dan email boleh ditimpa lewat Environment Variables (dipakai saat
// di-deploy ke hosting, supaya rahasia tidak perlu ditulis di config.json).
if (process.env.KODE_AKSES) CFG.kode_akses = process.env.KODE_AKSES;
CFG.email = CFG.email || {};
if (process.env.EMAIL_PENGIRIM) CFG.email.pengirim = process.env.EMAIL_PENGIRIM;
if (process.env.EMAIL_APP_PASSWORD) CFG.email.app_password = process.env.EMAIL_APP_PASSWORD;
if (process.env.EMAIL_TUJUAN) CFG.email.tujuan = process.env.EMAIL_TUJUAN.split(',').map(s => s.trim()).filter(Boolean);
if (process.env.EMAIL_AKTIF) CFG.email.aktif = process.env.EMAIL_AKTIF === 'true';
const EMAIL_SIAP = !!(CFG.email.aktif && CFG.email.pengirim && CFG.email.app_password && (CFG.email.tujuan || []).length);

// Disk di hosting gratis bersifat sementara (hilang saat aplikasi tidur/di-restart).
// Kalau email aktif, setiap laporan tetap terkirim lewat email walau filenya hilang dari disk.
let kirimEmail = null;
if (EMAIL_SIAP) {
  try {
    const nodemailer = require('nodemailer');
    const transport = nodemailer.createTransport({
      service: 'gmail',
      auth: { user: CFG.email.pengirim, pass: CFG.email.app_password }
    });
    kirimEmail = async (judul, docxPath, pdfPath) => {
      const lampiran = [{ filename: path.basename(docxPath), path: docxPath }];
      if (pdfPath) lampiran.push({ filename: path.basename(pdfPath), path: pdfPath });
      await transport.sendMail({
        from: CFG.email.pengirim, to: CFG.email.tujuan.join(','), subject: judul,
        text: 'Laporan otomatis terlampir. Email ini dikirim karena penyimpanan di server bersifat sementara.',
        attachments: lampiran
      });
    };
  } catch (e) {
    console.error('Modul nodemailer belum terpasang, email dimatikan:', e.message);
  }
}

const MAKS_BODY = 40 * 1024 * 1024;     // 40 MB per kiriman
const MAKS_FOTO_BYTES = 4 * 1024 * 1024; // 4 MB per foto (setelah dikecilkan di browser)

// ------------------------------ utilitas ------------------------------

const json = (res, code, obj) => {
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(obj));
};
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const sama = (a, b) => {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
};

// batasi tebak-tebakan kode akses: 10 kali salah / 10 menit / IP
const gagal = new Map();
function ip(req) { return (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim(); }
function terblokir(req) {
  const g = gagal.get(ip(req));
  return g && g.n >= 10 && Date.now() - g.t < 10 * 60 * 1000;
}
function kodeBenar(req, kode) {
  if (!CFG.kode_akses) return true;
  if (terblokir(req)) return false;
  if (sama(kode || '', CFG.kode_akses)) { gagal.delete(ip(req)); return true; }
  const g = gagal.get(ip(req)) || { n: 0, t: 0 };
  gagal.set(ip(req), { n: g.n + 1, t: Date.now() });
  return false;
}

function bacaBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on('data', c => {
      size += c.length;
      if (size > MAKS_BODY) { reject(new Error('Kiriman terlalu besar (maks 40 MB). Kurangi jumlah foto.')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

// ------------------------- normalisasi & validasi ---------------------

const bersih = (s, max = 300) => String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/\s+/g, ' ').trim().slice(0, max);
const kapital = s => s.charAt(0).toUpperCase() + s.slice(1);
const titleCase = s => s.split(' ').map(kapital).join(' ');
const KECIL = new Set(['dan', 'di', 'ke', 'dari', 'yang', 'untuk', 'pada', 'dengan', 'atau']);
const titleJab = s => s.split(' ').map((w, i) => (i > 0 && KECIL.has(w.toLowerCase())) ? w.toLowerCase() : kapital(w)).join(' ');
const jam = s => {
  const m = String(s || '').match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return String(m[1]).padStart(2, '0') + '.' + m[2];
};
const barisList = (t, maxN = 30) => String(t || '').split(/\r?\n/)
  .map(s => bersih(s.replace(/^\s*(\d+[.)]|[-•*])\s+/, ''), 300)).filter(Boolean).slice(0, maxN).map(kapital);

function ukuranGambar(buf) {
  if (buf.length > 24 && buf.slice(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])))
    return { type: 'png', w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
    let i = 2;
    while (i + 9 < buf.length) {
      if (buf[i] !== 0xff) { i++; continue; }
      const m = buf[i + 1];
      if (m === 0xff) { i++; continue; }
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc)
        return { type: 'jpg', h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
      i += 2 + buf.readUInt16BE(i + 2);
    }
  }
  return null;
}

function normalisasi(b) {
  const tg = String(b.tanggal || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!tg) return { error: 'Tanggal pelaksanaan belum diisi.' };
  const mulai = jam(b.mulai), selesai = jam(b.selesai);
  if (!mulai || !selesai) return { error: 'Waktu mulai dan selesai harus diisi.' };

  const tempat = titleCase(bersih(b.tempat, 120));
  if (!tempat) return { error: 'Tempat pelaksanaan belum diisi.' };

  const orang = (arr, maxN) => (Array.isArray(arr) ? arr : []).slice(0, maxN)
    .map(x => ({ nama: titleCase(bersih(x && x.nama, 120)), jab: titleJab(bersih(x && x.jab, 120)) }))
    .filter(x => x.nama);

  const pim = orang([b.pimpinan], 1)[0];
  if (!pim) return { error: 'Nama pimpinan rapat belum diisi.' };
  const narasumber = orang(b.narasumber, 20);
  if (!narasumber.length) return { error: 'Isi minimal satu narasumber.' };
  const peserta = orang(b.peserta, 300);
  if (!peserta.length) return { error: 'Isi minimal satu peserta.' };

  const unsur = (Array.isArray(b.unsur) ? b.unsur : []).map(x => bersih(x, 80)).filter(Boolean).slice(0, 30);
  const lain = bersih(b.unsurLain, 120);
  if (lain) unsur.push(lain);
  if (!unsur.length) return { error: 'Pilih minimal satu unsur yang hadir.' };

  const jadwal = (Array.isArray(b.jadwal) ? b.jadwal : []).slice(0, 40)
    .map(x => ({ waktu: bersih(x && x.waktu, 40), materi: kapital(bersih(x && x.materi, 200)), ns: bersih(x && x.ns, 150) }))
    .filter(x => x.materi);

  const paket = ['P1', 'P2', 'P3', 'P4', 'P5'].map(k => barisList(b.paket && b.paket[k], 20));

  const foto = [];
  for (const f of (Array.isArray(b.foto) ? b.foto : []).slice(0, CFG.maks_foto || 8)) {
    const m = String((f && f.data) || '').match(/^data:image\/(jpeg|png);base64,([A-Za-z0-9+/=]+)$/);
    if (!m) continue;
    const buf = Buffer.from(m[2], 'base64');
    if (buf.length > MAKS_FOTO_BYTES) return { error: 'Ada foto yang terlalu besar (maks 4 MB per foto).' };
    const uk = ukuranGambar(buf);
    if (!uk || !uk.w || !uk.h) continue;
    foto.push({ buf, type: uk.type, w: uk.w, h: uk.h });
  }

  return {
    data: {
      Y: +tg[1], M: +tg[2], D: +tg[3], mulai, selesai, tempat,
      nomor: bersih(b.nomor, 80), pimpinan: pim, unsur, narasumber, peserta, jadwal,
      materi: barisList(b.materi, 20), paket, foto, uangSaku: !!b.uangSaku
    }
  };
}

// ------------------------------ PDF ----------------------------------

let antrean = Promise.resolve();
function kePdf(docx) {
  const job = antrean.then(() => konversi(docx));
  antrean = job.catch(() => {});
  return job;
}
function konversi(docx) {
  return new Promise(resolve => {
    const dir = path.dirname(docx);
    const pdf = docx.replace(/\.docx$/i, '.pdf');
    const kandidat = [CFG.soffice_path, process.env.SOFFICE_PATH, 'soffice', 'libreoffice',
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
      '/Applications/LibreOffice.app/Contents/MacOS/soffice'].filter(Boolean);

    const wordCom = () => {
      if (process.platform !== 'win32') return resolve(null);
      const q = s => s.replace(/'/g, "''");
      const ps = `$w=New-Object -ComObject Word.Application;$w.Visible=$false;` +
        `$d=$w.Documents.Open('${q(docx)}');$d.SaveAs([ref]'${q(pdf)}',[ref]17);$d.Close();$w.Quit()`;
      execFile('powershell', ['-NoProfile', '-Command', ps], { timeout: 120000 },
        err => resolve(!err && fs.existsSync(pdf) ? pdf : null));
    };

    (function coba(i) {
      if (i >= kandidat.length) return wordCom();
      execFile(kandidat[i], ['--headless', '--convert-to', 'pdf', '--outdir', dir, docx], { timeout: 120000 },
        err => (!err && fs.existsSync(pdf)) ? resolve(pdf) : coba(i + 1));
    })(0);
  });
}

// ------------------------------ route --------------------------------

function kirimBerkas(res, file, tipe, unduh) {
  fs.readFile(file, (err, buf) => {
    if (err) { res.writeHead(404); return res.end('Tidak ditemukan'); }
    const h = { 'Content-Type': tipe, 'Content-Length': buf.length, 'Cache-Control': 'no-store' };
    if (unduh) h['Content-Disposition'] = `attachment; filename="${unduh}"`;
    res.writeHead(200, h); res.end(buf);
  });
}

const MIME = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.pdf': 'application/pdf'
};

async function buatDanSimpan(d) {
  const id = crypto.randomBytes(8).toString('hex');
  const dir = path.join(OUT, id);
  fs.mkdirSync(dir);
  const tglStr = `${String(d.D).padStart(2, '0')}-${BULAN[d.M - 1]}-${d.Y}`;
  const nama = `Laporan_${CFG.nama_kegiatan.replace(/\s+/g, '_')}_Desa_${CFG.desa}_${tglStr}`;
  const docxPath = path.join(dir, nama + '.docx');
  fs.writeFileSync(docxPath, await buatLaporan(d, CFG));
  const pdfPath = await kePdf(docxPath);
  fs.writeFileSync(path.join(dir, 'info.json'), JSON.stringify({
    dibuat: new Date().toISOString(), tanggal: `${tglStr}`, tempat: d.tempat, peserta: d.peserta.length
  }));

  let emailTerkirim = false;
  if (kirimEmail) {
    try { await kirimEmail(nama, docxPath, pdfPath); emailTerkirim = true; }
    catch (e) { console.error('Gagal kirim email:', e.message); }
  }

  return {
    id, emailTerkirim,
    docx: `/hasil/${id}/${encodeURIComponent(nama)}.docx`,
    pdf: pdfPath ? `/hasil/${id}/${encodeURIComponent(nama)}.pdf` : null
  };
}

function halamanRiwayat() {
  const ids = fs.readdirSync(OUT).filter(x => /^[a-f0-9]{16}$/.test(x));
  const baris = ids.map(id => {
    try {
      const info = JSON.parse(fs.readFileSync(path.join(OUT, id, 'info.json'), 'utf8'));
      const files = fs.readdirSync(path.join(OUT, id));
      const link = ext => { const f = files.find(x => x.endsWith(ext)); return f ? `<a href="/hasil/${id}/${encodeURIComponent(f)}">${ext.slice(1).toUpperCase()}</a>` : '-'; };
      return { t: info.dibuat, html: `<tr><td>${esc(new Date(info.dibuat).toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' }))}</td><td>${esc(info.tanggal)}</td><td>${esc(info.tempat)}</td><td>${info.peserta}</td><td>${link('.docx')} · ${link('.pdf')}</td></tr>` };
    } catch (e) { return null; }
  }).filter(Boolean).sort((a, b) => b.t.localeCompare(a.t));
  return `<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Riwayat Laporan</title>
<style>body{font-family:system-ui,sans-serif;max-width:900px;margin:24px auto;padding:0 16px}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccc;padding:8px;text-align:left;font-size:14px}th{background:#f0f0f0}a{color:#0a58ca}</style>
<h2>Riwayat Laporan ${esc(CFG.nama_kegiatan)} – Desa ${esc(CFG.desa)}</h2>
<table><tr><th>Dibuat</th><th>Tanggal kegiatan</th><th>Tempat</th><th>Peserta</th><th>Unduh</th></tr>${baris.map(b => b.html).join('') || '<tr><td colspan="5">Belum ada laporan.</td></tr>'}</table>`;
}

const server = http.createServer(async (req, res) => {
  try {
    const u = new URL(req.url, 'http://localhost');

    if (req.method === 'GET' && (u.pathname === '/' || u.pathname === '/index.html'))
      return kirimBerkas(res, path.join(__dirname, 'public', 'index.html'), 'text/html; charset=utf-8');

    if (req.method === 'GET' && u.pathname === '/api/info')
      return json(res, 200, {
        kegiatan: CFG.nama_kegiatan, desa: CFG.desa, butuhKode: !!CFG.kode_akses,
        uangSaku: (CFG.uang_saku_per_peserta || 0) > 0, maksFoto: CFG.maks_foto || 8
      });

    if (req.method === 'POST' && u.pathname === '/api/laporan') {
      let body;
      try { body = JSON.parse(await bacaBody(req)); } catch (e) { return json(res, 400, { ok: false, pesan: e.message.startsWith('Kiriman') ? e.message : 'Data tidak terbaca.' }); }
      if (!kodeBenar(req, body.kode))
        return json(res, 403, { ok: false, pesan: terblokir(req) ? 'Terlalu banyak percobaan. Coba lagi 10 menit lagi.' : 'Kode akses salah.' });
      const n = normalisasi(body);
      if (n.error) return json(res, 400, { ok: false, pesan: n.error });
      const hasil = await buatDanSimpan(n.data);
      return json(res, 200, { ok: true, ...hasil });
    }

    if (req.method === 'GET' && u.pathname === '/riwayat') {
      if (!kodeBenar(req, u.searchParams.get('kode'))) { res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' }); return res.end('Akses ditolak. Tambahkan ?kode=KODE_AKSES pada alamat.'); }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(halamanRiwayat());
    }

    const m = u.pathname.match(/^\/hasil\/([a-f0-9]{16})\/([^/]+)$/);
    if (req.method === 'GET' && m) {
      const nama = decodeURIComponent(m[2]);
      if (!/^[\w.\-]+$/.test(nama) || !/\.(docx|pdf)$/i.test(nama)) { res.writeHead(404); return res.end('Tidak ditemukan'); }
      return kirimBerkas(res, path.join(OUT, m[1], nama), MIME[path.extname(nama).toLowerCase()], nama);
    }

    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }); res.end('Tidak ditemukan');
  } catch (err) {
    console.error(err);
    if (!res.headersSent) json(res, 500, { ok: false, pesan: 'Terjadi kesalahan di server: ' + err.message });
  }
});

server.on('error', err => {
  if (err.code === 'EADDRINUSE') console.error(`\n  GAGAL: port ${PORT} sedang dipakai program lain. Tutup jendela lama atau ubah "port" di config.json.\n`);
  else console.error(err);
  process.exit(1);
});

server.listen(PORT, HOST, () => {
  console.log(`\n  Aplikasi laporan ${CFG.nama_kegiatan} Desa ${CFG.desa} berjalan di ${HOST}:${PORT}`);
  console.log(`  Email notifikasi     : ${EMAIL_SIAP ? 'aktif -> ' + CFG.email.tujuan.join(', ') : 'tidak aktif'}`);
  if (HOST === '0.0.0.0') {
    console.log(`  Berjalan sebagai layanan hosting (Render dkk). Buka lewat alamat publik yang diberikan hosting.`);
  } else {
    console.log(`  Buka di komputer ini : http://localhost:${PORT}`);
    console.log(`  Riwayat laporan      : http://localhost:${PORT}/riwayat?kode=${CFG.kode_akses || ''}`);
    console.log(`  Kode akses           : ${CFG.kode_akses || '(tidak dipakai)'}`);
    console.log(`\n  Untuk dibagikan ke orang lain, jalankan di terminal lain:  ngrok http 127.0.0.1:${PORT}\n`);
  }
});
