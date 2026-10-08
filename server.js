try {
  // Node 20.6+ 내장 .env 로더. 파일이 없으면 조용히 넘어간다 (예: PORT만
  // env로 넘기고 다른 값은 다 기본값을 쓰는 배포 환경).
  process.loadEnvFile();
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}

// 서버(카카오클라우드)는 사내 프록시를 거쳐야만 외부(구글드라이브, 아지트)로 나갈 수
// 있는데, Node 20의 내장 fetch는 HTTPS_PROXY 같은 환경변수를 자동으로 따르지 않는다.
// 프록시 환경변수가 있을 때만 모든 fetch가 프록시를 쓰도록 전역 설정한다
// (Mac 로컬처럼 프록시 변수가 없으면 아무것도 바뀌지 않음).
if (process.env.HTTPS_PROXY || process.env.https_proxy || process.env.HTTP_PROXY || process.env.http_proxy) {
  const { setGlobalDispatcher, EnvHttpProxyAgent } = require('undici');
  setGlobalDispatcher(new EnvHttpProxyAgent());
}

const express = require('express');
const fs = require('fs/promises');
const path = require('path');
const { execFile } = require('child_process');
const { promisify } = require('util');
const { buildReportModel, renderReportText, deriveAnomalyRecords, formatMentionLine, applyAccountAliases } = require('./report-core.js');
const {
  DriveAuthError,
  getDriveFolderIdForMonth,
  downloadMonthFolderToTemp,
  cleanupTempDir,
} = require('./drive-helper.js');

const execFileAsync = promisify(execFile);

const app = express();
const PORT = Number(process.env.PORT) || 3000;
const dataFile = path.join(__dirname, 'data', 'monthly-stats.json');
const historyFile = path.join(__dirname, 'data', 'monthly-anomaly-history.json');
const aliasesFile = path.join(__dirname, 'data', 'account-aliases.json');
const clarificationFile = path.join(__dirname, 'data', 'clarification-requests.json');
const ANALYZE_CLI_PATH = path.join(__dirname, 'analyze-cli.js');

app.use(express.json({ limit: '1mb' }));
app.use(express.static(path.join(__dirname, 'public')));

async function readStats() {
  try {
    const contents = await fs.readFile(dataFile, 'utf8');
    const stats = JSON.parse(contents);
    return Array.isArray(stats) ? stats : [];
  } catch (error) {
    if (error.code === 'ENOENT') return [];
    throw error;
  }
}

async function writeStats(stats) {
  await fs.mkdir(path.dirname(dataFile), { recursive: true });
  await fs.writeFile(dataFile, JSON.stringify(stats, null, 2) + '\n', 'utf8');
}

function validateRecord(record) {
  return record && /^\d{4}-\d{2}$/.test(record.month) &&
    typeof record.system === 'string' && record.system.trim().length > 0 &&
    Number.isFinite(Number(record.downloadCount)) && record.anomalies && typeof record.anomalies === 'object';
}

app.get('/api/monthly-stats', async (req, res) => {
  try {
    res.json(await readStats());
  } catch (error) {
    res.status(500).json({ error: '월별 데이터를 읽지 못했습니다.' });
  }
});

app.post('/api/monthly-stats', async (req, res) => {
  const record = req.body;
  if (!validateRecord(record)) {
    return res.status(400).json({ error: '월, 시스템명, 다운로드 건수, 이상 항목이 필요합니다.' });
  }

  try {
    const stats = await readStats();
    const normalized = {
      month: record.month,
      system: record.system.trim(),
      downloadCount: Number(record.downloadCount),
      anomalies: Object.fromEntries(Object.entries(record.anomalies).map(([key, value]) => [key, Number(value) || 0])),
      savedAt: new Date().toISOString()
    };
    const index = stats.findIndex(item => item.month === normalized.month && item.system === normalized.system);
    const exists = index >= 0;
    if (exists && !record.confirmOverwrite) {
      return res.status(409).json({ exists: true, message: '같은 연월과 시스템의 데이터가 이미 있습니다.' });
    }
    if (exists) stats[index] = normalized;
    else stats.push(normalized);
    stats.sort((a, b) => a.month.localeCompare(b.month) || a.system.localeCompare(b.system));
    await writeStats(stats);
    res.status(exists ? 200 : 201).json({ saved: normalized, replaced: exists });
  } catch (error) {
    res.status(500).json({ error: '월별 데이터를 저장하지 못했습니다.' });
  }
});

app.delete('/api/monthly-stats/all', async (req, res) => {
  try {
    await writeStats([]);
    res.json({ deleted: true });
  } catch (error) {
    res.status(500).json({ error: '전체 기록을 삭제하지 못했습니다.' });
  }
});

app.delete('/api/monthly-stats', async (req, res) => {
  const { month, system } = req.body || {};
  if (!month || !system) {
    return res.status(400).json({ error: '월과 시스템명이 필요합니다.' });
  }

  try {
    const stats = await readStats();
    const index = stats.findIndex(item => item.month === month && item.system === system);
    if (index < 0) {
      return res.status(404).json({ error: '해당 기록을 찾을 수 없습니다.' });
    }
    stats.splice(index, 1);
    await writeStats(stats);
    res.json({ deleted: true });
  } catch (error) {
    res.status(500).json({ error: '기록을 삭제하지 못했습니다.' });
  }
});

app.get('/api/monthly-stats/export', async (req, res) => {
  try {
    const stats = await readStats();
    const headers = ['연월', '시스템', '다운로드 건수', '이상 항목별 발생 횟수', '저장 시각'];
    const rows = stats.map(item => [
      item.month,
      item.system,
      item.downloadCount,
      Object.entries(item.anomalies || {}).map(([key, value]) => `${key}:${value}`).join(', '),
      item.savedAt || ''
    ]);
    const csv = [headers, ...rows].map(row => row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(',')).join('\n');
    res.set({ 'Content-Type': 'text/csv; charset=utf-8', 'Content-Disposition': 'attachment; filename="monthly-stats.csv"' });
    res.send('\ufeff' + csv);
  } catch (error) {
    res.status(500).json({ error: '엑셀 내보내기에 실패했습니다.' });
  }
});

// ── 점검 이력 분석 (월간 점검 리포트 계산 시 [3-2]~[3-9] hasIssue 스냅샷을 누적) ──
async function readHistory() {
  try {
    const contents = await fs.readFile(historyFile, 'utf8');
    const records = JSON.parse(contents);
    return Array.isArray(records) ? records : [];
  } catch (error) {
    if (error.code === 'ENOENT') return [];
    throw error;
  }
}

async function writeHistory(records) {
  await fs.mkdir(path.dirname(historyFile), { recursive: true });
  await fs.writeFile(historyFile, JSON.stringify(records, null, 2) + '\n', 'utf8');
}

// (month, system) 단위로 upsert한다. 같은 월을 다시 조회/발행하면 최신 계산값으로
// 덮어써진다 (멱등) - 리포트를 여러 번 열어봐도 이력이 중복되지 않는다.
async function recordAnomalyHistory(month, systems) {
  const newRecords = deriveAnomalyRecords(month, systems).map((r) => ({ ...r, savedAt: new Date().toISOString() }));
  if (!newRecords.length) return;
  const history = await readHistory();
  newRecords.forEach((rec) => {
    const index = history.findIndex((h) => h.month === rec.month && h.system === rec.system);
    if (index >= 0) history[index] = rec;
    else history.push(rec);
  });
  history.sort((a, b) => a.month.localeCompare(b.month) || a.system.localeCompare(b.system));
  await writeHistory(history);
}

app.get('/api/anomaly-history', async (req, res) => {
  try {
    res.json(await readHistory());
  } catch (error) {
    res.status(500).json({ error: '점검 이력을 읽지 못했습니다.' });
  }
});

// ── 소명 요청 (1단계: 아지트 글 발행 / 2단계는 사람이 직접 댓글을 복사-붙여넣기) ──
async function readClarificationRequests() {
  try {
    const contents = await fs.readFile(clarificationFile, 'utf8');
    const records = JSON.parse(contents);
    return Array.isArray(records) ? records : [];
  } catch (error) {
    if (error.code === 'ENOENT') return [];
    throw error;
  }
}

async function writeClarificationRequests(records) {
  await fs.mkdir(path.dirname(clarificationFile), { recursive: true });
  await fs.writeFile(clarificationFile, JSON.stringify(records, null, 2) + '\n', 'utf8');
}

app.get('/api/clarification-requests', async (req, res) => {
  try {
    res.json(await readClarificationRequests());
  } catch (error) {
    res.status(500).json({ error: '소명 요청 이력을 읽지 못했습니다.' });
  }
});

// (month) 단위로 upsert한다 - 같은 달에 소명 요청을 다시 보내면 최신 글 URL/항목으로 덮어써진다.
app.post('/api/publish-clarification', async (req, res) => {
  const { month, mentions, text, items } = req.body || {};
  if (!isValidMonth(month)) {
    return res.status(400).json({ error: '월(YYYY-MM) 형식이 올바르지 않습니다.' });
  }
  if (!text || !String(text).trim()) {
    return res.status(400).json({ error: '발행할 내용이 비어 있습니다.' });
  }

  const webhookUrl = process.env.AGIT_ANALYSIS_WEBHOOK_URL;
  if (!webhookUrl) {
    return res.status(500).json({ error: '서버에 AGIT_ANALYSIS_WEBHOOK_URL이 설정되어 있지 않습니다 (.env 확인).' });
  }

  try {
    const mentionLine = formatMentionLine(Array.isArray(mentions) ? mentions : []);
    const fullText = mentionLine + '\n\n' + String(text);

    const webhookRes = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: fullText, task: { assignees: [] } }),
    });
    const webhookBody = await webhookRes.json().catch(() => null);
    if (!webhookBody || webhookBody.status !== 'ok') {
      return res.status(502).json({ error: '아지트 게시에 실패했습니다.', detail: webhookBody });
    }

    const records = await readClarificationRequests();
    const record = {
      month,
      wallMessageUrl: webhookBody.url,
      requestedItems: Array.isArray(items) ? items : [],
      mentions: Array.isArray(mentions) ? mentions : [],
      savedAt: new Date().toISOString(),
    };
    const index = records.findIndex((r) => r.month === month);
    if (index >= 0) records[index] = record;
    else records.push(record);
    records.sort((a, b) => a.month.localeCompare(b.month));
    await writeClarificationRequests(records);

    res.json({ url: webhookBody.url, id: webhookBody.id });
  } catch (error) {
    res.status(500).json({ error: `소명 요청 발행에 실패했습니다: ${error.message}` });
  }
});

// ── 월간 점검 리포트 (analyze-cli.js 재사용) ────────────────────────────────
function isValidMonth(month) {
  return typeof month === 'string' && /^\d{4}-\d{2}$/.test(month);
}

async function runAnalyzeCli(folder, month) {
  const { stdout } = await execFileAsync('node', [ANALYZE_CLI_PATH, folder, month], {
    maxBuffer: 1024 * 1024 * 50,
    timeout: 120000,
  });
  return JSON.parse(stdout);
}

// state.json에서 해당 월의 drive_folder_id를 찾아 구글드라이브 파일들을 임시
// 디렉토리로 내려받고, analyze-cli.js로 분석한 systems 배열을 반환한다.
// (agit-automation/analyze_and_post.py의 run_and_post()와 동일한 흐름: 드라이브
// 폴더 조회 -> 임시 다운로드 -> 분석 -> 임시 폴더 삭제.)
async function getSystemsForMonth(month) {
  const folderId = await getDriveFolderIdForMonth(month);
  if (!folderId) {
    const error = new Error(`${month} 기간의 드라이브 폴더 정보를 찾을 수 없습니다.`);
    error.code = 'NO_FOLDER';
    throw error;
  }

  const { tmpDir, fileCount } = await downloadMonthFolderToTemp(folderId);
  try {
    if (fileCount === 0) {
      const error = new Error('드라이브 폴더에 분석할 xlsx 파일이 없습니다.');
      error.code = 'NO_FILES';
      throw error;
    }
    const data = await runAnalyzeCli(tmpDir, month);
    return data.systems || [];
  } finally {
    await cleanupTempDir(tmpDir);
  }
}

// ── 계정명 -> LDAP 매핑 (data/account-aliases.json) ─────────────────────
async function readAccountAliases() {
  try {
    const contents = await fs.readFile(aliasesFile, 'utf8');
    const parsed = JSON.parse(contents);
    return (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) ? parsed : {};
  } catch (error) {
    if (error.code === 'ENOENT') return {};
    throw error;
  }
}

async function writeAccountAliases(aliases) {
  await fs.mkdir(path.dirname(aliasesFile), { recursive: true });
  await fs.writeFile(aliasesFile, JSON.stringify(aliases, null, 2) + '\n', 'utf8');
}

app.get('/api/account-aliases', async (req, res) => {
  try {
    res.json(await readAccountAliases());
  } catch (error) {
    res.status(500).json({ error: '계정 매핑을 읽지 못했습니다.' });
  }
});

app.post('/api/account-aliases', async (req, res) => {
  const { from, to } = req.body || {};
  if (!from || typeof from !== 'string' || !to || typeof to !== 'string' || !from.trim() || !to.trim()) {
    return res.status(400).json({ error: '로그 상 계정명(from)과 실제 LDAP(to)이 모두 필요합니다.' });
  }
  try {
    const aliases = await readAccountAliases();
    aliases[from.trim()] = to.trim();
    await writeAccountAliases(aliases);
    res.json(aliases);
  } catch (error) {
    res.status(500).json({ error: '계정 매핑을 저장하지 못했습니다.' });
  }
});

app.delete('/api/account-aliases', async (req, res) => {
  const { from } = req.body || {};
  if (!from || typeof from !== 'string') {
    return res.status(400).json({ error: '삭제할 계정명(from)이 필요합니다.' });
  }
  try {
    const aliases = await readAccountAliases();
    delete aliases[from];
    await writeAccountAliases(aliases);
    res.json(aliases);
  } catch (error) {
    res.status(500).json({ error: '계정 매핑을 삭제하지 못했습니다.' });
  }
});

// state.json에서 systems를 가져온 뒤 바로 계정명 매핑을 적용한다 - 이후
// buildReportModel/renderReportText/recordAnomalyHistory 전부 실제 LDAP
// 기준으로 일관되게 동작한다.
async function getAliasedSystemsForMonth(month) {
  const [systems, aliases] = await Promise.all([getSystemsForMonth(month), readAccountAliases()]);
  return applyAccountAliases(systems, aliases);
}

function respondWithReportError(res, error) {
  if (error.code === 'NO_FOLDER' || error.code === 'NO_FILES') {
    return res.status(404).json({ error: error.message });
  }
  if (error instanceof DriveAuthError) {
    return res.status(502).json({ error: error.message });
  }
  return res.status(500).json({ error: `리포트 생성에 실패했습니다: ${error.message}` });
}

app.get('/api/monthly-report', async (req, res) => {
  const month = req.query.month;
  if (!isValidMonth(month)) {
    return res.status(400).json({ error: '월(YYYY-MM) 형식이 올바르지 않습니다.' });
  }

  try {
    const systems = await getAliasedSystemsForMonth(month);
    const model = buildReportModel(month, systems);
    await recordAnomalyHistory(month, systems);
    res.json(model);
  } catch (error) {
    respondWithReportError(res, error);
  }
});

app.post('/api/publish-report', async (req, res) => {
  const { month, comments, improvement, wikiLink, cc, mentions } = req.body || {};
  if (!isValidMonth(month)) {
    return res.status(400).json({ error: '월(YYYY-MM) 형식이 올바르지 않습니다.' });
  }

  const webhookUrl = process.env.AGIT_ANALYSIS_WEBHOOK_URL;
  if (!webhookUrl) {
    return res.status(500).json({ error: '서버에 AGIT_ANALYSIS_WEBHOOK_URL이 설정되어 있지 않습니다 (.env 확인).' });
  }

  try {
    const systems = await getAliasedSystemsForMonth(month);
    await recordAnomalyHistory(month, systems);
    const text = renderReportText(month, systems, {
      comments: comments || {},
      improvement,
      wikiLink,
      cc,
      assignees: Array.isArray(mentions) ? mentions : undefined,
      realMention: true,
    });

    const webhookRes = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text, task: { assignees: [] } }),
    });
    const webhookBody = await webhookRes.json().catch(() => null);
    if (!webhookBody || webhookBody.status !== 'ok') {
      return res.status(502).json({ error: '아지트 게시에 실패했습니다.', detail: webhookBody });
    }
    res.json({ url: webhookBody.url, id: webhookBody.id });
  } catch (error) {
    if (error.code === 'NO_FOLDER' || error.code === 'NO_FILES' || error instanceof DriveAuthError) {
      return respondWithReportError(res, error);
    }
    res.status(500).json({ error: `게시에 실패했습니다: ${error.message}` });
  }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// 개인정보 접속기록이 보이는 페이지인데 로그인 기능이 없으므로, 기본값은 서버 자기
// 자신(127.0.0.1)에서만 열리게 한다. 외부에서는 터널(포트포워딩)로만 접속한다.
// 꼭 다른 주소로 열어야 할 때만 .env 에 HOST 를 지정한다.
const HOST = process.env.HOST || '127.0.0.1';
app.listen(PORT, HOST, () => {
  console.log(`Access log dashboard: http://${HOST}:${PORT}`);
});
