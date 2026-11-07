// Pembuat dokumen Word laporan Rembuk Stunting.
// Dipanggil oleh server.js:  const buf = await buatLaporan(data, cfg)
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
  WidthType, BorderStyle, ShadingType, LevelFormat, VerticalAlign, HeightRule, ImageRun
} = require('docx');

const FONT = 'Times New Roman';
const HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const PAKET = ['Kesehatan Ibu dan Anak', 'Sanitasi dan Air Bersih', 'Konseling Gizi Terpadu', 'Jaminan Perlindungan Sosial', 'PAUD'];

const pad2 = n => (n < 10 ? '0' + n : '' + n);
const gabung = a => a.length <= 1 ? (a[0] || '-') : a.slice(0, -1).join(', ') + ' dan ' + a[a.length - 1];
const rp = n => 'Rp. ' + String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',-';

async function buatLaporan(d, cfg) {
  const { Y, M, D } = d;
  const tgl = new Date(Y, M - 1, D);
  const hari = HARI[tgl.getDay()];
  const tglStr = `${pad2(D)} ${BULAN[M - 1]} ${Y}`;
  const KEG = cfg.nama_kegiatan, DESA = cfg.desa;
  const JAB_KADES = cfg.jabatan_kades;
  const W = 9638; // lebar isi A4 dengan margin 2 cm (DXA)
  const AL = AlignmentType.LEFT, C = AlignmentType.CENTER, R = AlignmentType.RIGHT, J = AlignmentType.JUSTIFY;

  const run = (text, o = {}) => new TextRun({ text, font: FONT, size: (o.size || 12) * 2, bold: !!o.bold, underline: o.underline ? {} : undefined });
  const P = (text, o = {}) => new Paragraph({
    alignment: o.align || J,
    spacing: { before: (o.before || 0) * 20, after: (o.after === undefined ? 6 : o.after) * 20, line: 276 },
    indent: o.indent ? { left: o.indent } : undefined,
    pageBreakBefore: !!o.pb,
    keepNext: !!o.keepNext,
    border: o.rule ? { bottom: { style: BorderStyle.SINGLE, size: 18, color: '000000', space: 4 } } : undefined,
    children: [run(text, o)]
  });

  const NB = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const NOB = { top: NB, bottom: NB, left: NB, right: NB };
  const LB = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  const LINES = { top: LB, bottom: LB, left: LB, right: LB };

  const cell = (text, w, o = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: o.borders || LINES,
    verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: o.align || AL, spacing: { before: 0, after: 0 },
      children: [run(text, { size: o.size || 11, bold: o.bold, underline: o.underline })]
    })]
  });

  const kv = rows => {
    const w = [2400, 240, W - 2640];
    return new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: w,
      rows: rows.map(([k, v]) => new TableRow({
        children: [cell(k, w[0], { borders: NOB, size: 12 }), cell(':', w[1], { borders: NOB, size: 12 }), cell(v, w[2], { borders: NOB, size: 12 })]
      }))
    });
  };

  const grid = (header, rows, widths, aligns) => new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], { bold: true, align: C, fill: 'D9D9D9' })) }),
      ...rows.map(row => new TableRow({
        height: { value: 520, rule: HeightRule.ATLEAST }, cantSplit: true,
        children: row.map((t, i) => cell(t, widths[i], { align: aligns[i], bold: row.bold }))
      }))
    ]
  });

  const NAMA_KOSONG = '(................................................)';
  const ttd = (kiri, kanan) => {
    const w = [W / 2, W / 2];
    const c = (t, ww, o = {}) => cell(t, ww, { borders: NOB, align: C, size: 12, ...o });
    return new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: w,
      rows: [
        new TableRow({ children: [c(kiri.jab, w[0]), c(kanan.jab, w[1])] }),
        new TableRow({ height: { value: 1100, rule: HeightRule.EXACT }, children: [c('', w[0]), c('', w[1])] }),
        new TableRow({ children: [c(kiri.nama, w[0]), c(kanan.nama, w[1])] })
      ]
    });
  };
  const ttdBlok = () => ttd(
    { jab: `Ketua BPD Desa ${DESA}`, nama: cfg.nama_ketua_bpd ? cfg.nama_ketua_bpd : NAMA_KOSONG },
    { jab: `${JAB_KADES} ${DESA}`, nama: cfg.nama_kades ? cfg.nama_kades : NAMA_KOSONG });

  const numItem = (ref, text) => new Paragraph({
    numbering: { reference: ref, level: 0 }, alignment: J,
    spacing: { before: 0, after: 40, line: 276 }, children: [run(text)]
  });

  const materi = d.materi.length ? d.materi : [
    'Pembukaan',
    'Penyampaian Hasil Pendataan RT 1.000 HPK oleh KPM sesuai Peta Sosial Dasar',
    `Pembahasan Masalah dan Gagasan terkait intervensi 5 Paket Layanan dalam Perencanaan Desa TA. ${Y}`,
    `Penyusunan Prioritas Usulan Pencegahan Stunting TA. ${Y}`,
    'Penutup'
  ];

  const kids = [];

  // ===== 1. BERITA ACARA
  kids.push(
    P(`PEMERINTAHAN KABUPATEN ${cfg.kabupaten.toUpperCase()}`, { align: C, bold: true, after: 0 }),
    P(`KECAMATAN ${cfg.kecamatan.toUpperCase()}`, { align: C, bold: true, after: 0 }),
    P(`DESA ${DESA.toUpperCase()}`, { align: C, bold: true, size: 14, after: 0 }),
    P(`Alamat : ${cfg.alamat}`, { align: C, size: 10, after: 10, rule: true }),
    P('BERITA ACARA', { align: C, bold: true, size: 13, underline: true, before: 8, after: d.nomor ? 0 : 12 })
  );
  if (d.nomor) kids.push(P('Nomor : ' + d.nomor, { align: C, after: 12 }));
  kids.push(
    P(`Dalam rangka pelaksanaan Konvergensi Pencegahan Stunting Tahun ${Y} di Desa ${DESA}, Kecamatan ${cfg.kecamatan}, Kabupaten ${cfg.kabupaten}, Provinsi ${cfg.provinsi}, maka pada hari ini :`),
    kv([['Hari dan Tanggal', `${hari}, ${tglStr}`], ['Waktu', `${d.mulai} s.d. ${d.selesai} WIB`], ['Tempat', d.tempat]]),
    P(`Telah dilaksanakan kegiatan ${KEG} dalam rangka pembahasan hasil Pendataan Rumah Tangga 1.000 Hari Pertama Kehidupan (HPK), Percepatan, Pencegahan dan Penanganan Stunting untuk tahun ${Y}, yang dihadiri oleh ${gabung(d.unsur)} serta unsur lain yang terkait sebagaimana tercantum dalam Daftar Hadir terlampir.`, { before: 8 }),
    P('Materi atau topik yang dibahas dalam Rapat Koordinasi ini serta yang bertindak selaku unsur pimpinan rapat dan narasumber adalah :'),
    P('A. Materi atau Topik', { bold: true, align: AL, after: 2, keepNext: true }),
    ...materi.map(m => numItem('materi', m)),
    P('B. Unsur Pimpinan Rapat dan Narasumber/Fasilitator', { bold: true, align: AL, before: 6, after: 2, keepNext: true }),
    numItem('unsur', `${d.pimpinan.nama}${d.pimpinan.jab ? ', ' + d.pimpinan.jab : ''} (Pimpinan Rapat)`),
    ...d.narasumber.map(n => numItem('unsur', `${n.nama}${n.jab ? ', ' + n.jab : ''} (Narasumber)`)),
    P('Demikian Berita Acara ini dibuat dan disahkan dengan penuh tanggung jawab agar dapat dipergunakan sebagaimana mestinya.', { before: 8 }),
    kv([['Disahkan di', `Desa ${DESA}`], ['Pada tanggal', tglStr]]),
    P('', { after: 4 }),
    ttdBlok()
  );

  // ===== 2. RISALAH
  kids.push(
    P(`RISALAH HASIL ${KEG.toUpperCase()}`, { align: C, bold: true, size: 13, underline: true, after: 12, pb: true }),
    P(`Pada hari ini ${hari}, tanggal ${pad2(D)}, bulan ${BULAN[M - 1]}, tahun ${Y} bertempat di ${d.tempat}, telah dilaksanakan kegiatan ${KEG} Desa ${DESA} dalam rangka pelaksanaan Konvergensi Pencegahan Stunting di Desa.`),
    P(`Adapun usulan hasil ${KEG} yang akan tertuang dalam Rencana Kerja Pemerintah Desa Tahun ${Y} sesuai dengan intervensi 5 Paket Layanan sebagai berikut :`)
  );
  PAKET.forEach((j, i) => {
    kids.push(new Paragraph({
      numbering: { reference: 'paket', level: 0 }, keepNext: true,
      spacing: { before: 80, after: 40 }, children: [run(j, { bold: true })]
    }));
    const items = d.paket[i] || [];
    if (!items.length) kids.push(P('-', { indent: 720, after: 2 }));
    items.forEach(it => kids.push(new Paragraph({
      numbering: { reference: 'dash', level: 0 }, alignment: J,
      spacing: { before: 0, after: 40, line: 276 }, children: [run(it)]
    })));
  });
  kids.push(
    P('Demikian Risalah ini dibuat untuk ditindaklanjuti sebagaimana rencana kerja tindak lanjut di Desa.', { before: 12 }),
    P(`Desa ${DESA}, ${tglStr}`, { align: R, after: 4, keepNext: true }),
    P('Mengetahui,', { align: C, bold: true, after: 0, keepNext: true }),
    P(`Peserta ${KEG} sebagaimana Daftar Hadir terlampir.`, { align: C, after: 6, keepNext: true }),
    ttdBlok()
  );

  // ===== 3. DAFTAR HADIR
  kids.push(
    P(`DAFTAR HADIR PESERTA ${KEG.toUpperCase()} DESA ${DESA.toUpperCase()} TAHUN ${Y}`, { align: C, bold: true, after: 10, pb: true }),
    kv([['Hari / Tanggal', `${hari}, ${tglStr}`], ['Tempat', d.tempat]]),
    P('', { after: 4 }),
    grid(['No', 'Nama', 'Jabatan', 'Tanda Tangan'],
      d.peserta.map((p, i) => [String(i + 1), p.nama, p.jab, '']),
      [560, 3300, 2878, 2900], [C, AL, AL, AL])
  );

  // ===== 4. JADWAL
  if (d.jadwal.length) kids.push(
    P(`JADWAL ${KEG.toUpperCase()} DESA ${DESA.toUpperCase()}`, { align: C, bold: true, after: 0, pb: true }),
    P(`${hari}, ${tglStr}`.toUpperCase(), { align: C, bold: true, after: 10 }),
    grid(['No', 'Waktu', 'Pemateri', 'Narasumber'],
      d.jadwal.map((j, i) => [String(i + 1), j.waktu, j.materi, j.ns]),
      [560, 1900, 4300, 2878], [C, AL, AL, AL])
  );

  // ===== 5. DOKUMENTASI
  const half = W / 2;
  kids.push(P(`DOKUMENTASI ${KEG.toUpperCase()} DESA ${DESA.toUpperCase()} TAHUN ${Y}`, { align: C, bold: true, after: 10, pb: true }));
  if (d.foto.length) {
    const rows = [];
    for (let i = 0; i < d.foto.length; i += 2) {
      const cells = [0, 1].map(k => {
        const f = d.foto[i + k];
        const kids2 = [];
        if (f) {
          let w = 290, h = Math.round(f.h * w / f.w);
          if (h > 260) { h = 260; w = Math.round(f.w * h / f.h); }
          kids2.push(new Paragraph({
            alignment: C, spacing: { before: 60, after: 60 },
            children: [new ImageRun({
              type: f.type, data: f.buf, transformation: { width: w, height: h },
              altText: { title: 'Dokumentasi', description: `Foto dokumentasi ${i + k + 1}`, name: `foto${i + k + 1}` }
            })]
          }));
        } else kids2.push(new Paragraph({ children: [run('')] }));
        return new TableCell({
          width: { size: half, type: WidthType.DXA }, borders: f ? LINES : NOB,
          verticalAlign: VerticalAlign.CENTER, children: kids2
        });
      });
      rows.push(new TableRow({ cantSplit: true, children: cells }));
    }
    kids.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [half, half], rows }));
  } else {
    const fotoCell = () => new TableCell({
      width: { size: half, type: WidthType.DXA }, borders: LINES, verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: C, children: [run('')] })]
    });
    kids.push(new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [half, half],
      rows: [0, 1, 2].map(() => new TableRow({ height: { value: 4000, rule: HeightRule.EXACT }, cantSplit: true, children: [fotoCell(), fotoCell()] }))
    }));
  }

  // ===== 6. TANDA TERIMA UANG SAKU
  const nominal = cfg.uang_saku_per_peserta || 0;
  if (d.uangSaku && nominal > 0 && d.peserta.length) {
    const rows = d.peserta.map((p, i) => [String(i + 1), p.nama, p.jab, rp(nominal), '']);
    const tot = ['', 'TOTAL', '', rp(nominal * d.peserta.length), '']; tot.bold = true;
    rows.push(tot);
    kids.push(
      P(`TANDA TERIMA UANG SAKU ${KEG.toUpperCase()} ${Y}`, { align: C, bold: true, after: 10, pb: true }),
      kv([['Hari / Tanggal', `${hari}, ${tglStr}`], ['Tempat', d.tempat]]),
      P('', { after: 4 }),
      grid(['No', 'Nama Peserta', 'Jabatan', 'Nominal', 'Tanda Tangan'], rows,
        [560, 2700, 2078, 1700, 2600], [C, AL, AL, R, AL])
    );
  }

  const numCfg = (ref, fmt, text) => ({
    reference: ref,
    levels: [{
      level: 0, format: fmt, text, alignment: AL,
      style: { paragraph: { indent: { left: ref === 'dash' ? 1080 : 720, hanging: 360 } }, run: { font: FONT, size: 24, bold: ref === 'paket' } }
    }]
  });

  const doc = new Document({
    creator: `Pemerintah Desa ${DESA}`,
    title: `Laporan ${KEG} Desa ${DESA} ${Y}`,
    styles: { default: { document: { run: { font: FONT, size: 24 } } } },
    numbering: {
      config: [
        numCfg('materi', LevelFormat.DECIMAL, '%1.'),
        numCfg('unsur', LevelFormat.DECIMAL, '%1.'),
        numCfg('paket', LevelFormat.DECIMAL, '%1.'),
        numCfg('dash', LevelFormat.BULLET, '-')
      ]
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
      children: kids
    }]
  });

  return Packer.toBuffer(doc);
}

module.exports = { buatLaporan, BULAN, HARI };
