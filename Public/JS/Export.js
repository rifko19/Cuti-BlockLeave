// Export data cuti ke Excel (hanya admin). Butuh SheetJS (index.html) dan helper iso dari App.js

const fmtId = (s) => s.split('-').reverse().join('/');   // 2027-01-11 -> 11/01/2027

function exportLeavesToExcel(leaves) {
  if (!isAdmin()) { alert('Hanya admin yang dapat mengunduh data.'); return; }
  if (!leaves.length) { alert('Belum ada data cuti untuk diunduh.'); return; }

  const sorted = [...leaves].sort((a, b) => a.leave_start.localeCompare(b.leave_start) || a.employee_name.localeCompare(b.employee_name));
  const rows = sorted.map((l, i) => ({
    'No': i + 1,
    'Nama Pegawai': l.employee_name,
    'Pn': l.employee_id,
    'Unit Kerja': l.uker,
    'Tanggal Mulai Cuti': fmtId(l.leave_start),
    'Tanggal Selesai Cuti': fmtId(l.leave_end),
    'Jumlah Hari': l.leave_days,
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [6, 28, 20, 22, 14].map((w) => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data Cuti');
  XLSX.writeFile(wb, `data-cuti-${iso(new Date())}.xlsx`);
}