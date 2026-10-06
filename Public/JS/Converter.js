// Konversi Excel -> employees.json (berjalan di browser, butuh SheetJS)
const PN_LENGTH = 8;   // PERNR angka dipadatkan ke 8 digit (5670 -> 00005670). Isi 0 untuk menonaktifkan

const el = (id) => document.getElementById(id);
const clean = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();
const escH = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// Pola tebakan kolom, dicoba berurutan (PERNR didahulukan dari NIP)
const GUESS = {
  colPn: [/^pernr$/i, /^pn$/i, /no\.?\s*pegawai|personal/i, /^(nip|nik|id)$/i],
  colName: [/^completename$/i, /nama|name/i],
  colPos: [/^jobdesc$/i, /jabatan|posisi|title/i],
  colUker: [/^uker$/i, /unit\s*kerja|cabang|kantor/i],
  colKat: [/^kategori$/i, /kategori|category/i],
};

let headers = [];
let rows = [];
let result = [];

function guessCol(id) {
  for (const re of GUESS[id]) {
    const i = headers.findIndex((h) => re.test(h));
    if (i >= 0) return i;
  }
  return -1;
}

function readFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const wb = XLSX.read(e.target.result, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    const headerIdx = aoa.findIndex((r) => r.filter((c) => clean(c)).length >= 2);
    if (headerIdx < 0) { alert('Tidak ada data terbaca di sheet pertama.'); return; }
    headers = aoa[headerIdx].map(clean);
    rows = aoa.slice(headerIdx + 1);
    fillMapping();
    build();
  };
  reader.readAsArrayBuffer(file);
}

function fillMapping() {
  const opts = '<option value="-1">(pilih kolom)</option>' + headers.map((h, i) => `<option value="${i}">${escH(h || 'Kolom ' + (i + 1))}</option>`).join('');
  Object.keys(GUESS).forEach((id) => {
    el(id).innerHTML = opts;
    el(id).value = String(guessCol(id));
    el(id).onchange = build;
  });
  el('mapPanel').hidden = false;
}

function build() {
  const idx = Object.fromEntries(Object.keys(GUESS).map((id) => [id, Number(el(id).value)]));
  const seen = new Set();
  let skipped = 0, dup = 0;
  result = [];
  if (!Object.values(idx).some((i) => i < 0)) {
    rows.forEach((r) => {
      let pn = clean(r[idx.colPn]);
      const name = clean(r[idx.colName]);
      if (!pn || !name) { if (r.some((c) => clean(c))) skipped++; return; }
      if (PN_LENGTH && /^\d+$/.test(pn) && pn.length < PN_LENGTH) pn = pn.padStart(PN_LENGTH, '0');
      if (seen.has(pn)) { dup++; return; }
      seen.add(pn);
      result.push({ pn, name, position: clean(r[idx.colPos]).toUpperCase(), uker: clean(r[idx.colUker]), kategori: clean(r[idx.colKat]).toUpperCase() });
    });
  }

  const kat = {};
  result.forEach((e) => { kat[e.kategori || '(KOSONG)'] = (kat[e.kategori || '(KOSONG)'] || 0) + 1; });
  const katText = Object.entries(kat).map(([k, n]) => `${escH(k)}: <strong>${n}</strong>`).join(' &nbsp;|&nbsp; ');
  el('summary').innerHTML = `<p><strong>${result.length}</strong> pegawai siap diunduh. Dilewati karena PN atau nama kosong: ${skipped}. PN ganda dilewati: ${dup}.</p><p>Per kategori: ${katText || '-'}</p>`;
  el('preview').innerHTML = '<thead><tr><th>PN</th><th>Nama</th><th>Jabatan</th><th>Unit kerja</th><th>Kategori</th></tr></thead><tbody>' +
    result.slice(0, 8).map((e) => `<tr><td>${escH(e.pn)}</td><td>${escH(e.name)}</td><td>${escH(e.position)}</td><td>${escH(e.uker)}</td><td>${escH(e.kategori)}</td></tr>`).join('') + '</tbody>';
  el('resultPanel').hidden = false;
  el('btnDownload').disabled = !result.length;
}

function download() {
  const blob = new Blob([JSON.stringify({ employees: result }, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'employees.json';
  a.click();
  URL.revokeObjectURL(a.href);
}

el('file').addEventListener('change', (e) => e.target.files[0] && readFile(e.target.files[0]));
el('btnDownload').addEventListener('click', download);