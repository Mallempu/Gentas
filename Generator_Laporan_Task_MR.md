# Generator Laporan Task dari GitLab Merge Request

Blueprint project untuk menarik **Merge Request (MR)** dari GitLab per sprint/periode, lalu merangkainya menjadi **laporan task dalam format Excel standar SDI**. Dokumen ini berisi semua yang dibutuhkan untuk membangun sendiri: arsitektur, pemetaan data, format output, kode lengkap, cara menjalankan, dan ide pengembangan.

---

## 1. Tujuan

- Mengubah MR yang sudah dikerjakan menjadi **dokumentasi task** tanpa input manual.
- Cakupan **beberapa MR per sprint/periode** (bukan satu-satu).
- Output: **Excel** dengan kolom & gaya standar tim (bisa langsung diarsip/dilaporkan).

---

## 2. Fitur

- Tarik MR dari GitLab API dengan filter: state, target branch, rentang tanggal, milestone.
- **Paginasi otomatis** (semua MR terambil, bukan hanya 100 pertama).
- **Pengelompokan** baris parent: per milestone / target branch / bulan.
- **Pemetaan** field MR ke kolom task standar.
- **Styling** sesuai gaya SDI (header navy, parent, child, border, freeze pane).

---

## 3. Prasyarat

- Node.js (v18+ agar `fetch` tersedia bawaan).
- Paket: `exceljs` → `npm install exceljs`.
- **Personal Access Token** GitLab dengan scope `read_api`.

---

## 4. Konfigurasi

| Kunci | Keterangan | Contoh |
|---|---|---|
| `host` | Base URL GitLab, tanpa slash akhir | `https://gitlab.contoh.co.id` |
| `projectId` | ID proyek atau path ter-encode | `123` atau `grup%2Frepo` |
| `token` | Token GitLab (pakai env var) | `process.env.GITLAB_TOKEN` |
| `state` | `merged` / `opened` / `closed` / `all` | `merged` |
| `targetBranch` | Branch tujuan MR (kosong = semua) | `develop` |
| `updatedAfter` | MR diperbarui setelah tanggal (ISO) | `2026-09-01` |
| `updatedBefore` | MR diperbarui sebelum tanggal (ISO) | `2026-10-01` |
| `milestone` | Judul milestone (opsional) | `Sprint 2026-09` |
| `groupBy` | `milestone` / `target_branch` / `month` | `milestone` |
| `outFile` | Nama file keluaran | `Laporan_Task_MR.xlsx` |

> Praktik baik: simpan `token` di **environment variable**, jangan hardcode di file yang di-commit.

---

## 5. Sumber Data — Endpoint GitLab API

Base: `{host}/api/v4/projects/{projectId}`

| Kebutuhan | Endpoint |
|---|---|
| Daftar MR (dengan filter + paginasi) | `GET /merge_requests?state=&target_branch=&updated_after=&updated_before=&milestone=&per_page=100&page=N` |
| Commit per MR (opsional) | `GET /merge_requests/{iid}/commits` |
| File yang diubah (opsional) | `GET /merge_requests/{iid}/changes` |
| Komentar/diskusi (opsional) | `GET /merge_requests/{iid}/notes` |

Autentikasi: header `PRIVATE-TOKEN: {token}`.
Paginasi: cek header respons `x-next-page` untuk halaman berikutnya.

---

## 6. Pemetaan MR → Kolom Task Standar

Kolom standar: `Name, Assignee, Status, Priority, Created, Start Date, Due Date, Date Done, Completed At`

| Kolom | Sumber dari MR |
|---|---|
| Name | `title` + ` (!{iid})` |
| Assignee | `assignee.name` (fallback `author.name`) |
| Status | `state` dipetakan: merged→Done, opened→In Progress, closed→Cancelled |
| Priority | label pertama `labels[0]` (sesuaikan dgn konvensi label tim) |
| Created | `created_at` (YYYY-MM-DD) |
| Start Date | `created_at` (ganti bila ada sumber lain) |
| Due Date | `milestone.due_date` |
| Date Done | `merged_at` (YYYY-MM-DD) |
| Completed At | `merged_at` (YYYY-MM-DD HH:MM) |

> Sesuaikan `STATUS_MAP` dan sumber **Priority** dengan kosakata label/status yang dipakai tim SDI.

---

## 7. Format Excel (gaya standar SDI)

- Kolom & lebar: `[65, 18, 12, 10, 14, 14, 14, 14, 14]`.
- **Header** (baris 1): fill `#1F3864`, teks putih tebal.
- **Baris parent** (grup sprint/milestone): fill `#2E75B6`, teks putih tebal.
- **Baris child** (tiap MR): fill `#D6E4F0`, teks hitam.
- **Border** tipis `#BFBFBF` semua sel.
- **Freeze panes** di `A2` (header selalu terlihat).
- Struktur: 1 baris parent per grup, baris detail di bawahnya.

---

## 8. Kode Lengkap (`generate-mr-report.js`)

```javascript
/**
 * generate-mr-report.js
 * Tarik Merge Request dari GitLab lalu buat laporan task (Excel) format standar SDI.
 * Pakai: npm install exceljs  →  node generate-mr-report.js
 * Butuh token GitLab scope 'read_api'. Jangan hardcode token bila file di-commit.
 */

const ExcelJS = require('exceljs');

// ====== KONFIGURASI ======
const CFG = {
  host: process.env.GITLAB_HOST || 'https://gitlab.contoh.co.id',
  projectId: process.env.GITLAB_PROJECT_ID || '123',
  token: process.env.GITLAB_TOKEN || 'ISI_TOKEN_DI_ENV',
  state: 'merged',              // 'merged' | 'opened' | 'closed' | 'all'
  targetBranch: 'develop',      // '' untuk semua target
  updatedAfter: '2026-09-01',
  updatedBefore: '2026-10-01',
  milestone: '',                // judul milestone; '' = abaikan
  groupBy: 'milestone',         // 'milestone' | 'target_branch' | 'month'
  outFile: 'Laporan_Task_MR.xlsx',
};

const STATUS_MAP = { merged: 'Done', opened: 'In Progress', closed: 'Cancelled', locked: 'On Hold' };

// ====== AMBIL DATA (paginasi) ======
async function fetchAllMRs() {
  const base = `${CFG.host}/api/v4/projects/${encodeURIComponent(CFG.projectId)}/merge_requests`;
  const params = new URLSearchParams({ per_page: '100', order_by: 'updated_at', sort: 'asc' });
  if (CFG.state && CFG.state !== 'all') params.set('state', CFG.state);
  if (CFG.targetBranch) params.set('target_branch', CFG.targetBranch);
  if (CFG.updatedAfter) params.set('updated_after', `${CFG.updatedAfter}T00:00:00Z`);
  if (CFG.updatedBefore) params.set('updated_before', `${CFG.updatedBefore}T00:00:00Z`);
  if (CFG.milestone) params.set('milestone', CFG.milestone);

  let page = 1, all = [];
  while (true) {
    params.set('page', String(page));
    const res = await fetch(`${base}?${params.toString()}`, { headers: { 'PRIVATE-TOKEN': CFG.token } });
    if (!res.ok) throw new Error(`GitLab API ${res.status}: ${await res.text()}`);
    all = all.concat(await res.json());
    const next = res.headers.get('x-next-page');
    if (!next) break;
    page = Number(next);
    if (page > 50) break; // jaring pengaman
  }
  return all;
}

const d  = (iso) => (iso ? iso.slice(0, 10) : '');
const dt = (iso) => (iso ? iso.slice(0, 16).replace('T', ' ') : '');

function groupKey(mr) {
  if (CFG.groupBy === 'milestone') return mr.milestone?.title || 'Tanpa Milestone';
  if (CFG.groupBy === 'target_branch') return `target: ${mr.target_branch}`;
  if (CFG.groupBy === 'month') return (mr.merged_at || mr.updated_at || '').slice(0, 7);
  return 'Semua';
}

function toRow(mr) {
  return [
    `${mr.title} (!${mr.iid})`,
    mr.assignee?.name || mr.author?.name || '',
    STATUS_MAP[mr.state] || mr.state,
    (mr.labels && mr.labels[0]) || '',
    d(mr.created_at),
    d(mr.created_at),
    d(mr.milestone?.due_date),
    d(mr.merged_at),
    dt(mr.merged_at),
  ];
}

// ====== TULIS EXCEL ======
async function buildExcel(mrs) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Laporan Task MR');

  const cols = ['Name','Assignee','Status','Priority','Created','Start Date','Due Date','Date Done','Completed At'];
  const widths = [65,18,12,10,14,14,14,14,14];
  ws.columns = cols.map((c, i) => ({ header: c, width: widths[i] }));

  const B = { style:'thin', color:{argb:'FFBFBFBF'} };
  const border = { top:B, left:B, bottom:B, right:B };
  const fill = (argb) => ({ type:'pattern', pattern:'solid', fgColor:{argb} });

  const hr = ws.getRow(1);
  hr.eachCell((c) => { c.fill = fill('FF1F3864'); c.font = { name:'Arial', bold:true, color:{argb:'FFFFFFFF'} };
    c.alignment = { horizontal:'center', vertical:'middle', wrapText:true }; c.border = border; });
  hr.height = 20;

  const groups = {};
  for (const mr of mrs) { const k = groupKey(mr); (groups[k] ??= []).push(mr); }

  for (const [gname, list] of Object.entries(groups)) {
    const p = ws.addRow([gname, '', '', '', '', '', '', '', '']);
    p.eachCell((c) => { c.fill = fill('FF2E75B6'); c.font = { name:'Arial', bold:true, color:{argb:'FFFFFFFF'} };
      c.alignment = { horizontal:'left', vertical:'middle' }; c.border = border; });
    for (const mr of list) {
      const row = ws.addRow(toRow(mr));
      row.eachCell((c, i) => { c.fill = fill('FFD6E4F0'); c.font = { name:'Arial', color:{argb:'FF000000'} };
        c.alignment = { horizontal: (i<=2 ? 'left':'center'), vertical:'middle', wrapText: i===1 }; c.border = border; });
    }
  }

  ws.views = [{ state:'frozen', ySplit:1 }];
  await wb.xlsx.writeFile(CFG.outFile);
  console.log(`Selesai: ${CFG.outFile}  (${mrs.length} MR, ${Object.keys(groups).length} grup)`);
}

// ====== MAIN ======
(async () => {
  try {
    const mrs = await fetchAllMRs();
    if (!mrs.length) { console.log('Tidak ada MR pada filter ini.'); return; }
    await buildExcel(mrs);
  } catch (e) {
    console.error('Gagal membuat laporan:', e.message);
    process.exit(1);
  }
})();
```

---

## 9. Cara Menjalankan

```bash
# sekali setup
npm init -y
npm install exceljs

# set token (contoh Linux/Mac)
export GITLAB_TOKEN="glpat-xxxxxxxx"
export GITLAB_HOST="https://gitlab.contoh.co.id"
export GITLAB_PROJECT_ID="123"

# jalankan
node generate-mr-report.js
# -> menghasilkan Laporan_Task_MR.xlsx
```

---

## 10. Keamanan

- **Token** di environment variable, bukan di kode. Tambahkan `.env` ke `.gitignore`.
- Gunakan token **read-only** (`read_api`) — tidak perlu hak tulis.
- Hati-hati membagikan file Excel hasilnya bila berisi info internal (nama branch, judul task sensitif).

---

## 11. Ide Pengembangan Lanjutan

- **Detail commit** per MR → kolom tambahan atau sheet terpisah (`/merge_requests/{iid}/commits`).
- **Filter per author** (mis. laporan per anggota tim) → tambah parameter `author_username`.
- **Kaitan ke issue** → parse `description` untuk referensi `#issue`, tampilkan di kolom.
- **Otomasi terjadwal** → jalankan via cron tiap akhir sprint, kirim hasil ke email/Slack/Telegram.
- **Multi-proyek** → loop beberapa `projectId` jadi satu laporan gabungan.
- **Rekap ringkas** → sheet kedua berisi jumlah MR per status/author per sprint.
