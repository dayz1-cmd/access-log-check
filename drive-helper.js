// drive-helper.js
// agit-automation/drive_utils.py와 완전히 동일한 구글 OAuth 자격증명
// (~/agit-automation/credentials.json, token.json)을 그대로 재사용해
// 구글드라이브 폴더의 xlsx 파일들을 임시 디렉토리로 내려받는다.
// google-auth-library만 써서 access token 갱신을 처리하고, 실제 Drive REST
// 호출은 Node 내장 fetch로 직접 한다 (googleapis 전체 SDK는 필요 없음).
'use strict';

const fs = require('fs/promises');
const path = require('path');
const os = require('os');
const { OAuth2Client } = require('google-auth-library');

// agit-automation 프로젝트 경로. 환경변수(AGIT_AUTOMATION_PATH)로 덮어쓸 수 있고,
// 없으면 기존 하드코딩 기본값(~/agit-automation)을 그대로 쓴다 (하위호환 유지).
const AGIT_AUTOMATION_DIR = process.env.AGIT_AUTOMATION_PATH
  || path.join(os.homedir(), 'agit-automation');
const CREDENTIALS_PATH = path.join(AGIT_AUTOMATION_DIR, 'credentials.json');
const TOKEN_PATH = path.join(AGIT_AUTOMATION_DIR, 'token.json');
const STATE_PATH = path.join(AGIT_AUTOMATION_DIR, 'state.json');

const DRIVE_API = 'https://www.googleapis.com/drive/v3';

class DriveAuthError extends Error {}

async function readJson(filePath, notFoundMessage) {
  try {
    return JSON.parse(await fs.readFile(filePath, 'utf8'));
  } catch (error) {
    if (error.code === 'ENOENT') throw new DriveAuthError(notFoundMessage);
    throw new DriveAuthError(`${filePath} 파일을 읽지 못했습니다: ${error.message}`);
  }
}

// credentials.json + token.json으로 OAuth2Client를 만들고, refresh_token으로
// access token을 한 번 미리 갱신해 이후 요청에서 바로 쓸 수 있게 한다.
async function getDriveAccessToken() {
  const credentialsFile = await readJson(
    CREDENTIALS_PATH,
    `구글 인증 파일을 찾을 수 없습니다: ${CREDENTIALS_PATH} (agit-automation 쪽 설정을 확인해주세요.)`
  );
  const installed = credentialsFile.installed || credentialsFile.web;
  if (!installed || !installed.client_id || !installed.client_secret) {
    throw new DriveAuthError(`${CREDENTIALS_PATH}의 형식이 예상과 다릅니다 (installed.client_id/client_secret 필요).`);
  }

  const tokenFile = await readJson(
    TOKEN_PATH,
    `구글 인증 토큰을 찾을 수 없습니다: ${TOKEN_PATH} (agit-automation에서 최초 1회 인증을 완료해야 합니다.)`
  );
  if (!tokenFile.refresh_token) {
    throw new DriveAuthError(`${TOKEN_PATH}에 refresh_token이 없습니다. agit-automation에서 재인증이 필요합니다.`);
  }

  const client = new OAuth2Client(installed.client_id, installed.client_secret);
  client.setCredentials({ refresh_token: tokenFile.refresh_token });

  try {
    const { token } = await client.getAccessToken();
    if (!token) throw new Error('access token을 발급받지 못했습니다.');
    return token;
  } catch (error) {
    throw new DriveAuthError(`구글 인증에 실패했습니다: ${error.message} (agit-automation의 token.json이 만료/폐기되었을 수 있습니다.)`);
  }
}

async function driveFetch(accessToken, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Drive API 요청 실패 (${res.status}): ${body.slice(0, 300)}`);
  }
  return res;
}

async function listFilesInFolder(accessToken, folderId) {
  const q = encodeURIComponent(`'${folderId}' in parents and trashed = false`);
  const url = `${DRIVE_API}/files?q=${q}&fields=${encodeURIComponent('files(id,name,createdTime)')}&pageSize=1000`;
  const res = await driveFetch(accessToken, url);
  const body = await res.json();
  return body.files || [];
}

async function downloadFile(accessToken, fileId, destPath) {
  const url = `${DRIVE_API}/files/${fileId}?alt=media`;
  const res = await driveFetch(accessToken, url);
  const buffer = Buffer.from(await res.arrayBuffer());
  await fs.writeFile(destPath, buffer);
}

// state.json에서 해당 월(YYYY-MM)의 drive_folder_id를 찾는다.
async function getDriveFolderIdForMonth(month) {
  const state = await readJson(
    STATE_PATH,
    `state.json을 찾을 수 없습니다: ${STATE_PATH}`
  );
  const entry = state[month];
  if (!entry || !entry.drive_folder_id) {
    return null;
  }
  return entry.drive_folder_id;
}

// 구글드라이브 폴더의 xlsx 파일들을 새 임시 디렉토리로 전부 내려받고, 그
// 디렉토리 경로를 반환한다. 호출한 쪽에서 다 쓴 뒤 cleanupTempDir()로 지워야 한다.
async function downloadMonthFolderToTemp(driveFolderId) {
  const accessToken = await getDriveAccessToken();
  const files = await listFilesInFolder(accessToken, driveFolderId);
  const xlsxFiles = files.filter((f) => f.name.toLowerCase().endsWith('.xlsx'));

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'access-log-report-'));
  try {
    for (const f of xlsxFiles) {
      await downloadFile(accessToken, f.id, path.join(tmpDir, f.name));
    }
  } catch (error) {
    await cleanupTempDir(tmpDir);
    throw error;
  }
  return { tmpDir, fileCount: xlsxFiles.length };
}

async function cleanupTempDir(tmpDir) {
  if (!tmpDir) return;
  await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
}

module.exports = {
  DriveAuthError,
  getDriveFolderIdForMonth,
  downloadMonthFolderToTemp,
  cleanupTempDir,
  STATE_PATH,
};
