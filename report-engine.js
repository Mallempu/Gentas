/**
 * report-engine.js
 * Core logic: fetch MR dari GitLab + generate Excel.
 * Dipakai oleh CLI (generate-mr-report.js) dan server (server.js).
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const STATUS_MAP = { merged: 'Done', opened: 'In Progress', closed: 'Cancelled', locked: 'On Hold' };

function loadProjects() {
  const file = path.join(__dirname, 'projects.json');
  return JSON.parse(fs.readFileSync(file, 'utf-8'));
}

// Ambil daftar member project (untuk dropdown assignee)
async function fetchProjectMembers(host, token, projectId) {
  const url = `${host}/api/v4/projects/${encodeURIComponent(projectId)}/members/all?per_page=100`;
  const res = await fetch(url, { headers: { 'PRIVATE-TOKEN': token } });
  if (!res.ok) return [];
  const data = await res.json();
  return data.map(m => ({ id: m.id, name: m.name, username: m.username }));
}

// Ambil daftar member untuk project/group
async function fetchMembers(host, token, projectKey) {
  const projects = loadProjects();
  const selected = projects[projectKey];
  if (!selected) return [];

  if (selected.group) {
    const seen = new Set();
    let all = [];
    for (const key of selected.projects) {
      const proj = projects[key];
      if (!proj) continue;
      const members = await fetchProjectMembers(host, token, proj.projectId);
      for (const m of members) {
        if (!seen.has(m.username)) { seen.add(m.username); all.push(m); }
      }
    }
    return all.sort((a, b) => a.name.localeCompare(b.name));
  } else {
    const members = await fetchProjectMembers(host, token, selected.projectId);
    return members.sort((a, b) => a.name.localeCompare(b.name));
  }
}

// Ambil daftar commit per MR
async function fetchMRCommits(host, token, projectId, mrIid) {
  const url = `${host}/api/v4/projects/${encodeURIComponent(projectId)}/merge_requests/${mrIid}/commits?per_page=100`;
  const res = await fetch(url, { headers: { 'PRIVATE-TOKEN': token } });
  if (!res.ok) return [];
  return res.json();
}

async function fetchMRsForProject(host, token, proj, updatedAfter, updatedBefore, assigneeUsername, onLog) {
  const base = `${host}/api/v4/projects/${encodeURIComponent(proj.projectId)}/merge_requests`;
  const params = new URLSearchParams({ per_page: '100', order_by: 'updated_at', sort: 'asc' });

  const state = proj.state || 'merged';
  if (state && state !== 'all') params.set('state', state);
  if (proj.targetBranch) params.set('target_branch', proj.targetBranch);
  if (updatedAfter) params.set('updated_after', `${updatedAfter}T00:00:00Z`);
  if (updatedBefore) params.set('updated_before', `${updatedBefore}T00:00:00Z`);
  if (proj.milestone) params.set('milestone', proj.milestone);
  if (assigneeUsername) params.set('author_username', assigneeUsername);

  let page = 1, all = [];
  while (true) {
    params.set('page', String(page));
    const res = await fetch(`${base}?${params.toString()}`, { headers: { 'PRIVATE-TOKEN': token } });
    if (!res.ok) throw new Error(`GitLab API ${res.status} (${proj.name}): ${await res.text()}`);
    const data = await res.json();
    for (const mr of data) {
      mr._repoName = proj.name;
      mr._projectId = proj.projectId;
    }
    all = all.concat(data);
    const next = res.headers.get('x-next-page');
    if (!next) break;
    page = Number(next);
    if (page > 50) break;
  }

  // Fetch commits per MR
  for (const mr of all) {
    const commits = await fetchMRCommits(host, token, mr._projectId, mr.iid);
    mr._commits = commits;
    mr._commitMessages = commits.map(c => c.title || c.message || '').join('\n');
  }

  if (onLog) onLog(`${proj.name} (ID: ${proj.projectId}) → ${all.length} MR`);
  return all;
}

async function fetchAllMRs(host, token, projectKey, updatedAfter, updatedBefore, assigneeUsername, onLog) {
  const projects = loadProjects();
  const selected = projects[projectKey];
  if (!selected) throw new Error(`Project "${projectKey}" tidak ditemukan`);

  if (selected.group) {
    let all = [];
    for (const key of selected.projects) {
      const proj = projects[key];
      if (!proj) { if (onLog) onLog(`Project "${key}" tidak ditemukan, skip`); continue; }
      all = all.concat(await fetchMRsForProject(host, token, proj, updatedAfter, updatedBefore, assigneeUsername, onLog));
    }
    return { mrs: all, selected, projects };
  } else {
    const mrs = await fetchMRsForProject(host, token, selected, updatedAfter, updatedBefore, assigneeUsername, onLog);
    return { mrs, selected, projects };
  }
}

const d  = (iso) => (iso ? iso.slice(0, 10) : '');
const dt = (iso) => (iso ? iso.slice(0, 16).replace('T', ' ') : '');

function groupKey(mr, groupBy) {
  if (groupBy === 'repo') return mr._repoName || 'Unknown';
  if (groupBy === 'milestone') return mr.milestone?.title || 'Tanpa Milestone';
  if (groupBy === 'target_branch') return `target: ${mr.target_branch}`;
  if (groupBy === 'month') return (mr.merged_at || mr.updated_at || '').slice(0, 7);
  return 'Semua';
}

const COLS = ['MR ID','Name','Assignee','Status','Priority','Created','Start Date','Due Date','Date Done','Completed At','Commits'];
const COL_WIDTHS = [10, 55, 18, 12, 10, 14, 14, 14, 14, 14, 50];

function toRow(mr, isGroup) {
  const repoPrefix = isGroup ? `[${mr._repoName}] ` : '';
  return [
    `!${mr.iid}`,
    `${repoPrefix}${mr.title}`,
    mr.assignee?.name || mr.author?.name || '',
    STATUS_MAP[mr.state] || mr.state,
    (mr.labels && mr.labels[0]) || '',
    d(mr.created_at),
    d(mr.created_at),
    d(mr.milestone?.due_date),
    d(mr.merged_at),
    dt(mr.merged_at),
    mr._commitMessages || '',
  ];
}

async function buildExcelBuffer(mrs, selected) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Laporan Task MR');

  ws.columns = COLS.map((c, i) => ({ header: c, width: COL_WIDTHS[i] }));

  const B = { style:'thin', color:{argb:'FFBFBFBF'} };
  const border = { top:B, left:B, bottom:B, right:B };
  const fill = (argb) => ({ type:'pattern', pattern:'solid', fgColor:{argb} });

  const hr = ws.getRow(1);
  hr.eachCell((c) => {
    c.fill = fill('FF1F3864');
    c.font = { name:'Arial', bold:true, color:{argb:'FFFFFFFF'} };
    c.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
    c.border = border;
  });
  hr.height = 20;

  const gBy = selected.groupBy || (selected.group ? 'repo' : 'milestone');
  const groups = {};
  for (const mr of mrs) { const k = groupKey(mr, gBy); (groups[k] ??= []).push(mr); }

  const colCount = COLS.length;
  const emptyFill = Array(colCount - 1).fill('');

  for (const [gname, list] of Object.entries(groups)) {
    const p = ws.addRow([gname, ...emptyFill]);
    p.eachCell((c) => {
      c.fill = fill('FF2E75B6');
      c.font = { name:'Arial', bold:true, color:{argb:'FFFFFFFF'} };
      c.alignment = { horizontal:'left', vertical:'middle' };
      c.border = border;
    });
    for (const mr of list) {
      const row = ws.addRow(toRow(mr, !!selected.group));
      row.eachCell((c, i) => {
        c.fill = fill('FFD6E4F0');
        c.font = { name:'Arial', color:{argb:'FF000000'} };
        c.alignment = { horizontal: (i <= 3 ? 'left' : 'center'), vertical:'middle', wrapText: (i === 2 || i === colCount) };
        c.border = border;
      });
    }
  }

  ws.views = [{ state:'frozen', ySplit:1 }];
  return wb.xlsx.writeBuffer();
}

function getMRsAsTable(mrs, selected) {
  const gBy = selected.groupBy || (selected.group ? 'repo' : 'milestone');
  const groups = {};
  for (const mr of mrs) { const k = groupKey(mr, gBy); (groups[k] ??= []).push(mr); }

  const rows = [];
  for (const [gname, list] of Object.entries(groups)) {
    rows.push({ type: 'group', name: gname, count: list.length });
    for (const mr of list) {
      const r = toRow(mr, !!selected.group);
      rows.push({
        type: 'item',
        mrId: r[0], name: r[1], assignee: r[2], status: r[3], priority: r[4],
        created: r[5], startDate: r[6], dueDate: r[7], dateDone: r[8], completedAt: r[9],
        commits: r[10],
      });
    }
  }
  return rows;
}

module.exports = { loadProjects, fetchAllMRs, fetchMembers, buildExcelBuffer, getMRsAsTable };
