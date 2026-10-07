// ===== Konfigurasi =====
const ADMIN_PAGE = window.ADMIN_PAGE === true;   // true hanya di halaman /pathadmin/
const DATA_PATH = window.DATA_PATH || 'Data/';   // folder Data berada di dalam Public
const YEAR = 2027;
// Kategori (kolom KATEGORI di Excel) yang dihitung kuota 2 orang per minggu (digabung). Selain ini kuota 5 orang.
const QUOTA2_KATEGORI = ['BOH/MANAJER/SVP/RM/ADK'];
const LIMIT = { management: 1, staff: 5 };   // nilai awal. Nilai sebenarnya dibaca dari tabel quota_limits di Supabase
const ALLOWED_WEEKS = [1, 2, 3];
const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MONTHS = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

let holidays = new Map();   // 'YYYY-MM-DD' -> deskripsi
let employees = [];       // dari Data/employees.json
let selectedEmp = null;   // pegawai yang dipilih dari pencarian
let leaves = [];

// ===== Helper tanggal =====
const pad = (n) => String(n).padStart(2, '0');
const iso = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parse = (s) => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const mondayOf = (d) => { const x = new Date(d); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return x; };
const weekOfMonth = (monday) => Math.ceil(monday.getDate() / 7);   // berdasarkan tanggal Senin
const fmtShort = (d) => `${pad(d.getDate())} ${MONTHS[d.getMonth()].slice(0, 3)}`;
const groupOfKategori = (k) => (QUOTA2_KATEGORI.includes(String(k ?? '').toUpperCase().replace(/\s+/g, '').trim()) ? 'management' : 'staff');
const currentGroup = () => (selectedEmp ? groupOfKategori(selectedEmp.kategori) : 'staff');
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const $ = (id) => document.getElementById(id);

// ===== Aturan cuti =====
function holidaysInWeek(monday) {
  const found = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(monday); d.setDate(d.getDate() + i);
    if (holidays.has(iso(d))) found.push(`${fmtShort(d)}: ${holidays.get(iso(d))}`);
  }
  return found;
}

function employeesInWeek(mondayIso, group) {
  const set = new Set();
  leaves.forEach((l) => {
    if ((l.quota_group || 'staff') !== group) return;
    if (iso(mondayOf(parse(l.leave_start))) === mondayIso) set.add(l.employee_id.toUpperCase());
  });
  return set;
}

function validate(f) {
  if (!f.start || !f.end) return { error: 'Tanggal mulai dan selesai wajib diisi.' };
  const start = parse(f.start), end = parse(f.end);
  if (start.getFullYear() !== YEAR || end.getFullYear() !== YEAR) return { error: `Cuti hanya untuk tahun ${YEAR}.` };
  if (end < start) return { error: 'Tanggal selesai tidak boleh sebelum tanggal mulai.' };
  if ([0, 6].includes(start.getDay()) || [0, 6].includes(end.getDay())) return { error: 'Tanggal mulai dan selesai harus hari kerja (Senin-Jumat).' };

  const monday = mondayOf(start);
  if (iso(monday) !== iso(mondayOf(end))) return { error: 'Satu pengajuan harus berada dalam satu minggu yang sama.' };

  const week = weekOfMonth(monday);
  if (!ALLOWED_WEEKS.includes(week)) return { error: 'Cuti hanya boleh di minggu ke-1, 2, atau 3.' };

  const hol = holidaysInWeek(monday);
  if (hol.length) return { error: `Minggu ini tidak bisa dipakai karena ada tanggal merah/cuti bersama (${hol.join('; ')}).` };

  const group = currentGroup();
  const id = f.employee_id.toUpperCase();
  const current = employeesInWeek(iso(monday), group);
  if (!current.has(id) && current.size >= LIMIT[group]) {
    return { error: `Kuota minggu ini penuh (${current.size}/${LIMIT[group]} orang untuk ${group === 'management' ? 'BOH/MANAJER/SVP/RM/ADK' : 'jabatan lainnya'}).` };
  }

  const dates = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) dates.push(iso(d));
  const clash = leaves.find((l) => l.employee_id.toUpperCase() === id && l.leave_start <= f.end && l.leave_end >= f.start);
  if (clash) return { error: `Pegawai ini sudah punya pengajuan cuti pada ${clash.leave_start} s/d ${clash.leave_end}.` };

  return { dates, week };
}

// ===== Date picker (Flatpickr) =====
let startPicker, endPicker;

// Status satu tanggal: bisa dipilih atau tidak, beserta alasannya
function dayStatus(d) {
  if (d.getFullYear() !== YEAR) return { ok: false, reason: 'off' };
  if ([0, 6].includes(d.getDay())) return { ok: false, reason: 'off' };
  const monday = mondayOf(d);
  if (!ALLOWED_WEEKS.includes(weekOfMonth(monday))) return { ok: false, reason: 'off' };
  if (holidaysInWeek(monday).length) return { ok: false, reason: 'holiday-week' };
  const group = currentGroup();
  const id = $('employee_id').value.trim().toUpperCase();
  const current = employeesInWeek(iso(monday), group);
  if (!current.has(id) && current.size >= LIMIT[group]) return { ok: false, reason: 'full' };
  return { ok: true };
}

function setupPickers() {
  const base = {
    locale: { firstDayOfWeek: 1 },
    dateFormat: 'Y-m-d',
    altInput: true,
    altFormat: 'd/m/Y',
    minDate: `${YEAR}-01-01`,
    maxDate: `${YEAR}-12-31`,
    defaultDate: `${YEAR}-01-01`,
    disable: [(d) => !dayStatus(d).ok],
    onDayCreate: (_s, _v, _i, el) => {
      const d = el.dateObj;
      const key = iso(d);
      const st = dayStatus(d);
      if (holidays.has(key)) {
        el.classList.add('is-holiday');
        el.title = holidays.get(key);
      } else if (st.reason === 'full') {
        el.classList.add('is-full');
        el.title = 'Kuota minggu ini penuh';
      } else if (st.reason === 'holiday-week') {
        el.title = 'Minggu ini berisi tanggal merah/cuti bersama';
      }
    },
  };

  endPicker = flatpickr('#end_date', { ...base });
  startPicker = flatpickr('#start_date', {
    ...base,
    onChange: ([d]) => {
      if (!d) {   // dikosongkan
        endPicker.set('minDate', `${YEAR}-01-01`);
        endPicker.set('maxDate', `${YEAR}-12-31`);
        return;
      }
      const fri = new Date(mondayOf(d)); fri.setDate(fri.getDate() + 4);
      endPicker.set('minDate', iso(d));      // tanggal selesai hanya dalam minggu yang sama
      endPicker.set('maxDate', iso(fri));
      endPicker.setDate(d, true);
    },
  });
  startPicker.clear(); endPicker.clear();
}

const redrawPickers = () => { startPicker?.redraw(); endPicker?.redraw(); };

// ===== Supabase (database bersama) =====
const db = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
let adminSession = null;

async function loadLeaves() {
  let all = [];
  for (let from = 0; ; from += 1000) {   // PostgREST membatasi 1000 baris per permintaan
    const { data, error } = await db.from('leaves').select('*').order('leave_start').order('id').range(from, from + 999);
    if (error) { showMsg('Gagal memuat data pengajuan: ' + error.message, 'err'); return; }
    all = all.concat(data);
    if (data.length < 1000) break;
  }
  leaves = all;
}

// Batas kuota dibaca dari database, agar kalender dan server selalu sama
async function loadLimits() {
  const { data } = await db.from('quota_limits').select('*');
  (data || []).forEach((r) => { LIMIT[r.quota_group] = r.max_per_week; });
}

// ===== Render =====
function renderWeeks() {
  const month = Number($('monthFilter').value);
  const rows = [];
  const d = new Date(YEAR, month, 1);
  while (d.getDay() !== 1) d.setDate(d.getDate() + 1);   // Senin pertama
  for (let m = new Date(d); m.getMonth() === month; m.setDate(m.getDate() + 7)) {
    const week = weekOfMonth(m);
    if (!ALLOWED_WEEKS.includes(week)) continue;
    const fri = new Date(m); fri.setDate(fri.getDate() + 4);
    const hol = holidaysInWeek(m);
    let status;
    if (hol.length) {
      status = `<td data-label="Status" class="wide" colspan="2"><span class="tag block">Tidak tersedia</span> ${esc(hol.join('; '))}</td>`;
    } else {
      const cell = (g) => {
        const used = employeesInWeek(iso(m), g).size;
        const full = used >= LIMIT[g];
        return `<td data-label="${g === 'management' ? 'BOH / Manajer / SVP / RM / ADK' : 'Jabatan lainnya'}"><span class="tag ${full ? 'full' : 'open'}">${used}/${LIMIT[g]}</span> ${full ? 'penuh' : 'sisa ' + (LIMIT[g] - used)}</td>`;
      };
      status = cell('management') + cell('staff');
    }
    rows.push(`<tr><td data-label="Minggu">Minggu ${week}</td><td data-label="Tanggal">${fmtShort(m)} - ${fmtShort(fri)}</td>${status}</tr>`);
  }
  $('weekTable').innerHTML = `<thead><tr><th>Minggu</th><th>Tanggal</th><th>BOH / Manajer / SVP / RM / ADK</th><th>Jabatan lainnya</th></tr></thead><tbody>${rows.join('')}</tbody>`;
}

function renderLeaves() {
  const admin = isAdmin();
  const sorted = [...leaves].sort((a, b) => a.leave_start.localeCompare(b.leave_start));
  const body = sorted.length ? sorted.map((l) => `
    <tr>
      <td data-label="ID">${esc(l.id)}</td><td data-label="Nama">${esc(l.employee_name)}</td><td data-label="Personal Number">${esc(l.employee_id)}</td>
      <td data-label="Jabatan">${esc(l.position)}</td><td data-label="Uker">${esc(l.uker)}</td>
      <td data-label="Tanggal Mulai">${esc(l.leave_start)}</td><td data-label="Tanggal Selesai">${esc(l.leave_end)}</td>
      <td data-label="Jumlah Hari">${l.leave_days} hari</td><td data-label="Minggu">${l.leave_week}</td>
      <td data-label="Catatan">${esc(l.notes)}</td>
      ${admin ? `<td data-label=""><button class="btn small" data-del="${esc(l.id)}">Hapus</button></td>` : ''}
    </tr>`).join('') : `<tr><td class="empty" colspan="10">Belum ada pengajuan cuti.</td></tr>`;
  $('leaveTable').innerHTML = `<thead><tr><th>ID</th><th>Nama</th><th>Personal Number</th><th>Jabatan</th><th>Uker</th><th>Tanggal Mulai</th><th>Tanggal Selesai</th><th>Jumlah Hari</th><th>Minggu</th><th>Catatan</th>${admin ? '<th></th>' : ''}</tr></thead><tbody>${body}</tbody>`;
}

const refresh = () => { renderWeeks(); renderLeaves(); redrawPickers(); };

function showMsg(text, type) {
  const el = $('formMsg');
  el.textContent = text;
  el.className = type;
}

// ===== Event =====
async function onSubmit(e) {
  e.preventDefault();
  if (!selectedEmp) return showMsg('Pilih nama pegawai dari hasil pencarian (database pegawai).', 'err');
  const f = {
    name: $('employee_name').value.trim(),
    employee_id: $('employee_id').value.trim(),
    position: $('position').value,
    uker: $('uker').value,
    start: $('start_date').value,
    end: $('end_date').value,
    notes: $('notes').value.trim(),
  };
  const res = validate(f);
  if (res.error) return showMsg(res.error, 'err');

  const btn = e.target.querySelector('button[type="submit"]');
  btn.disabled = true;
  const { error } = await db.rpc('submit_leave', {
    p_name: f.name, p_pn: f.employee_id, p_position: f.position, p_uker: f.uker,
    p_group: currentGroup(), p_start: f.start, p_end: f.end, p_notes: f.notes,
  });
  btn.disabled = false;
  if (error) { await loadLeaves(); refresh(); return showMsg(error.message, 'err'); }   // pesan dari server, mis. kuota penuh

  e.target.reset();
  selectedEmp = null;
  startPicker.clear(); endPicker.clear();
  await loadLeaves();
  showMsg(`Pengajuan berhasil disimpan (${res.dates.length} hari kerja).`, 'ok');
  refresh();
}

async function onTableClick(e) {
  if (!isAdmin()) return;
  const id = e.target.dataset.del;
  if (!id || !confirm('Hapus pengajuan ini?')) return;
  const { error } = await db.from('leaves').delete().eq('id', id);
  if (error) { alert('Gagal menghapus: ' + error.message); return; }
  await loadLeaves(); refresh();
}

// ===== Login admin (Supabase Auth) =====
const isAdmin = () => ADMIN_PAGE && !!adminSession;

async function onLogin(e) {
  e.preventDefault();
  const msg = $('loginMsg');
  const { data, error } = await db.auth.signInWithPassword({ email: $('loginUser').value.trim(), password: $('loginPass').value });
  if (error) { msg.textContent = 'Email atau password salah.'; return; }
  adminSession = data.session;
  $('loginForm').reset(); msg.textContent = '';
  $('loginDialog').close();
  applyAuth();
}

function applyAuth() {
  const admin = isAdmin();
  if (ADMIN_PAGE) {
    $('btnExport').hidden = !admin;
    $('btnLogin').hidden = admin;
    $('btnLogout').hidden = !admin;
    $('adminLabel').hidden = !admin;
  }
  renderLeaves();
}

// ===== Pencarian pegawai (autocomplete) =====
const norm = (t) => String(t ?? '').toLowerCase();

function setEmployeeFields(emp) {
  $('employee_id').value = emp ? emp.pn : '';
  $('position').value = emp ? emp.position : '';
  $('uker').value = emp ? emp.uker : '';
}

function selectEmployee(emp) {
  selectedEmp = emp;
  $('employee_name').value = emp.name;
  setEmployeeFields(emp);
  $('empList').hidden = true;
  $('employee_name').setAttribute('aria-expanded', 'false');
  startPicker.clear(); endPicker.clear();   // jabatan berubah, kuota bisa berbeda
  redrawPickers();
}

function showEmployeeList() {
  const q = norm($('employee_name').value).trim();
  const ul = $('empList');
  if (!q) { ul.hidden = true; return; }
  const found = [];
  for (let i = 0; i < employees.length && found.length < 8; i++) {
    if (norm(employees[i].name).includes(q) || String(employees[i].pn).includes(q)) found.push(i);
  }
  ul.innerHTML = found.length
    ? found.map((i) => `<li role="option" data-i="${i}"><strong>${esc(employees[i].name)}</strong><span>${esc(employees[i].pn)} · ${esc(employees[i].position)} · ${esc(employees[i].uker)}</span></li>`).join('')
    : '<li class="none">Tidak ditemukan di database pegawai</li>';
  ul.hidden = false;
  $('employee_name').setAttribute('aria-expanded', 'true');
}

function setupEmployeeSearch() {
  const input = $('employee_name');
  const ul = $('empList');
  input.addEventListener('input', () => {
    if (selectedEmp && input.value !== selectedEmp.name) {   // nama diubah: pilihan dibatalkan
      selectedEmp = null; setEmployeeFields(null); redrawPickers();
    }
    showEmployeeList();
  });
  input.addEventListener('focus', showEmployeeList);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') ul.hidden = true;
    if (e.key === 'Enter' && !ul.hidden) {
      const first = ul.querySelector('li[data-i]');
      if (first) { e.preventDefault(); selectEmployee(employees[Number(first.dataset.i)]); }
    }
  });
  ul.addEventListener('mousedown', (e) => {   // mousedown agar terjadi sebelum input kehilangan fokus
    const li = e.target.closest('li[data-i]');
    if (li) { e.preventDefault(); selectEmployee(employees[Number(li.dataset.i)]); }
  });
  document.addEventListener('click', (e) => { if (!e.target.closest('.combo')) ul.hidden = true; });
}

// ===== Start =====
async function init() {
  try {
    const [h, e] = await Promise.all(['holidays', 'employees'].map((n) => fetch(`${DATA_PATH}${n}.json`).then((r) => { if (!r.ok) throw new Error(n); return r.json(); })));
    [...h.nationalHolidays, ...h.jointHolidays].forEach((x) => holidays.set(x.date, x.description));
    employees = e.employees;
  } catch (err) {
    showMsg('Gagal memuat folder Data (holidays.json / employees.json). Jalankan lewat Live Server dari folder "Blockleave System" (bukan klik dua kali pada index.html).', 'err');
    return;
  }

  setupEmployeeSearch();
  $('monthFilter').innerHTML = MONTHS.map((m, i) => `<option value="${i}">${m} ${YEAR}</option>`).join('');

  if (ADMIN_PAGE) {
    const { data: sess } = await db.auth.getSession();
    adminSession = sess.session;
  }
  await loadLimits();
  await loadLeaves();
  $('leaveForm').addEventListener('submit', onSubmit);
  setupPickers();
  // Kuota bergantung pada jabatan dan ID pegawai, jadi kalender digambar ulang saat keduanya berubah
  $('position').addEventListener('change', redrawPickers);
  $('employee_id').addEventListener('input', redrawPickers);
  $('monthFilter').addEventListener('change', renderWeeks);
  $('leaveTable').addEventListener('click', onTableClick);
  if (ADMIN_PAGE) {
  $('btnExport').addEventListener('click', () => exportLeavesToExcel(leaves));
  $('btnLogin').addEventListener('click', () => $('loginDialog').showModal());
  $('loginCancel').addEventListener('click', () => $('loginDialog').close());
  $('btnLogout').addEventListener('click', async () => { await db.auth.signOut(); adminSession = null; applyAuth(); });
  $('loginForm').addEventListener('submit', onLogin);
  }
  window.addEventListener('focus', async () => { await loadLeaves(); refresh(); });   // segarkan kuota saat kembali ke tab
  refresh();
  applyAuth();
}

init();