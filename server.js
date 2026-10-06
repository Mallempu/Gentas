/**
 * server.js — Web UI untuk Generator Laporan Task MR
 * Jalankan: node server.js
 * Buka: http://localhost:3000
 */

require('dotenv').config();
const express = require('express');
const path = require('path');
const { loadProjects, fetchAllMRs, fetchMembers, buildExcelBuffer, getMRsAsTable } = require('./report-engine');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const HOST = process.env.GITLAB_HOST;
const TOKEN = process.env.GITLAB_TOKEN;

// GET /api/projects — daftar project
app.get('/api/projects', (req, res) => {
  const projects = loadProjects();
  const list = [];
  for (const [key, p] of Object.entries(projects)) {
    if (key.startsWith('_')) continue;
    list.push({
      key,
      name: p.name,
      projectId: p.projectId || null,
      group: !!p.group,
      projects: p.projects || [],
      targetBranch: p.targetBranch || '',
      groupBy: p.groupBy || 'milestone',
    });
  }
  res.json(list);
});

// GET /api/members?project=xxx — daftar member untuk dropdown assignee
app.get('/api/members', async (req, res) => {
  const projectKey = req.query.project;
  if (!HOST || !TOKEN) return res.status(500).json({ error: 'GITLAB_HOST / GITLAB_TOKEN belum diisi' });
  if (!projectKey) return res.json([]);

  try {
    const members = await fetchMembers(HOST, TOKEN, projectKey);
    res.json(members);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/preview — preview data MR (tampil di tabel)
app.post('/api/preview', async (req, res) => {
  const { projectKey, updatedAfter, updatedBefore, assignee } = req.body;
  if (!HOST || !TOKEN) return res.status(500).json({ error: 'GITLAB_HOST / GITLAB_TOKEN belum diisi di .env' });
  if (!projectKey) return res.status(400).json({ error: 'projectKey wajib diisi' });

  try {
    const logs = [];
    const { mrs, selected } = await fetchAllMRs(HOST, TOKEN, projectKey, updatedAfter || '', updatedBefore || '', assignee || '', (msg) => logs.push(msg));
    const rows = getMRsAsTable(mrs, selected);
    res.json({ total: mrs.length, logs, rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/download — generate & download Excel
app.post('/api/download', async (req, res) => {
  const { projectKey, updatedAfter, updatedBefore, assignee } = req.body;
  if (!HOST || !TOKEN) return res.status(500).json({ error: 'GITLAB_HOST / GITLAB_TOKEN belum diisi di .env' });
  if (!projectKey) return res.status(400).json({ error: 'projectKey wajib diisi' });

  try {
    const { mrs, selected } = await fetchAllMRs(HOST, TOKEN, projectKey, updatedAfter || '', updatedBefore || '', assignee || '', () => {});
    if (!mrs.length) return res.status(404).json({ error: 'Tidak ada MR pada filter ini' });

    const buf = await buildExcelBuffer(mrs, selected);
    const filename = `Laporan_Task_${selected.name.replace(/\s+/g, '_')}.xlsx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server berjalan di http://localhost:${PORT}`);
});
