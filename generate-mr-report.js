/**
 * generate-mr-report.js (CLI)
 * node generate-mr-report.js <nama-project> [updatedAfter] [updatedBefore]
 */

require('dotenv').config();
const fs = require('fs');
const { loadProjects, fetchAllMRs, buildExcelBuffer } = require('./report-engine');

const projects = loadProjects();
const projectKey = process.argv[2];

if (!projectKey) {
  console.log('Pilih project:\n');
  for (const [key, p] of Object.entries(projects)) {
    if (key.startsWith('_')) continue;
    if (p.group) {
      console.log(`  node generate-mr-report.js ${key}    → [GROUP] ${p.name} (${p.projects.join(' + ')})`);
    } else {
      console.log(`  node generate-mr-report.js ${key}    → ${p.name} (ID: ${p.projectId})`);
    }
  }
  process.exit(0);
}

const HOST = process.env.GITLAB_HOST;
const TOKEN = process.env.GITLAB_TOKEN;
if (!HOST || !TOKEN) { console.error('GITLAB_HOST dan GITLAB_TOKEN harus diisi di .env'); process.exit(1); }

const updatedAfter = process.argv[3] || '';
const updatedBefore = process.argv[4] || '';

(async () => {
  try {
    const { mrs, selected } = await fetchAllMRs(HOST, TOKEN, projectKey, updatedAfter, updatedBefore, console.log);
    if (!mrs.length) { console.log('Tidak ada MR pada filter ini.'); return; }
    const buf = await buildExcelBuffer(mrs, selected);
    const outFile = `Laporan_Task_${selected.name.replace(/\s+/g, '_')}.xlsx`;
    fs.writeFileSync(outFile, buf);
    console.log(`Selesai: ${outFile}  (${mrs.length} MR)`);
  } catch (e) {
    console.error('Gagal:', e.message);
    process.exit(1);
  }
})();
