// Konversi Excel -> employees.json (berjalan di browser, butuh SheetJS)
const PN_LENGTH = 8;   // PERNR angka dipadatkan ke 8 digit (5670 -> 00005670). Isi 0 untuk menonaktifkan

const QUOTA2_KATEGORI = ['BOH/MANAJER/SVP/RM/ADK'];   // harus sama dengan QUOTA2_KATEGORI di App.js
const isQuota2 = (k) => QUOTA2_KATEGORI.includes(String(k).toUpperCase().replace(/\s+/g, ''));

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

// ===== Turunan data: daftar Uker dan daftar jabatan, semuanya dari Excel =====
function buildUker() {
  const count = {};
  result.forEach((e) => { if (e.uker) count[e.uker] = (count[e.uker] || 0) + 1; });
  const typeOf = (n) => (/^KCP\b/i.test(n) ? ['KCP', 'Kantor Cabang Pembantu', 1] : /^KC\b/i.test(n) ? ['KC', 'Kantor Cabang', 0] : ['KOP', 'Kantor Operasional', 2]);
  const list = Object.entries(count).map(([name, n]) => ({ name, n, t: typeOf(name) }))
    .sort((a, b) => a.t[2] - b.t[2] || a.name.localeCompare(b.name));
  const seq = {};
  return list.map((u) => {
    seq[u.t[0]] = (seq[u.t[0]] || 0) + 1;
    return { id: u.t[0] + String(seq[u.t[0]]).padStart(3, '0'), name: u.name, type: u.t[1], employee_count: u.n };
  });
}

function buildPositions() {
  const map = {};
  result.forEach((e) => {
    if (!e.position) return;
    const key = e.position + '|' + e.kategori;
    if (!map[key]) map[key] = { name: e.position, kategori: e.kategori, employee_count: 0 };
    map[key].employee_count++;
  });
  return Object.values(map)
    .sort((a, b) => Number(isQuota2(b.kategori)) - Number(isQuota2(a.kategori)) || a.name.localeCompare(b.name))
    .map((p, i) => ({ id: 'POS' + String(i + 1).padStart(3, '0'), name: p.name, kategori: p.kategori, max_per_week: isQuota2(p.kategori) ? 1 : 5, employee_count: p.employee_count }));
}

function buildCategories() {
  const map = {};
  result.forEach((e) => { const k = e.kategori || '(KOSONG)'; map[k] = (map[k] || 0) + 1; });
  return Object.entries(map).map(([name, n]) => ({ name, max_per_week: isQuota2(name) ? 1 : 5, employee_count: n }))
    .sort((a, b) => a.max_per_week - b.max_per_week);
}

const ukerJson = () => ({
  uker_list: buildUker(),
  seluruh_uker: true,
  notes: 'Dibuat otomatis dari Excel. Limit pengajuan cuti berlaku untuk SELURUH UKER, bukan per UKER',
});

const  sJson = () => ({
  categories: buildCategories(),
  positions: buildPositions(),
  rules: {
    allowed_weeks: [1, 2, 3],
    allowed_days: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
    allowed_days_id: [1, 2, 3, 4, 5],
    description: 'Cuti hanya bisa diambil di minggu 1, 2, atau 3, hari kerja Senin-Jumat',
  },
  holiday_rule: {
    description: 'Jika ada tanggal merah (national holiday atau cuti bersama) di dalam 1 minggu, maka seluruh minggu tersebut TIDAK boleh digunakan untuk pengajuan cuti',
    example: 'Jika Selasa ada tanggal merah, minggu yang mengandung Selasa tersebut tidak boleh digunakan',
  },
});

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
  el('summary').innerHTML = `<p><strong>${result.length}</strong> pegawai siap diunduh. Dilewati karena PN atau nama kosong: ${skipped}. PN ganda dilewati: ${dup}.</p><p>Per kategori: ${katText || '-'}</p><p>Unit kerja terbaca: <strong>${buildUker().length}</strong>. Jabatan terbaca: <strong>${buildPositions().length}</strong>.</p>`;
  el('preview').innerHTML = '<thead><tr><th>PN</th><th>Nama</th><th>Jabatan</th><th>Unit kerja</th><th>Kategori</th></tr></thead><tbody>' +
    result.slice(0, 8).map((e) => `<tr><td>${escH(e.pn)}</td><td>${escH(e.name)}</td><td>${escH(e.position)}</td><td>${escH(e.uker)}</td><td>${escH(e.kategori)}</td></tr>`).join('') + '</tbody>';
  el('resultPanel').hidden = false;
  ['btnDownload', 'btnUker', 'btnPos', 'btnAll'].forEach((id) => { el(id).disabled = !result.length; });
}

function downloadJson(filename, obj) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

el('file').addEventListener('change', (e) => e.target.files[0] && readFile(e.target.files[0]));
el('btnDownload').addEventListener('click', () => downloadJson('employees.json', { employees: result }));
el('btnUker').addEventListener('click', () => downloadJson('uker.json', ukerJson()));
el('btnPos').addEventListener('click', () => downloadJson('positions.json', positionsJson()));
el('btnAll').addEventListener('click', () => {
  downloadJson('employees.json', { employees: result });
  setTimeout(() => downloadJson('uker.json', ukerJson()), 400);
  setTimeout(() => downloadJson('positions.json', positionsJson()), 800);
});