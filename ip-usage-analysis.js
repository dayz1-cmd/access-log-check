#!/usr/bin/env node
// ip-usage-analysis.js
// 실제 접속기록(구글드라이브 원본 xlsx)에서 등장하는 모든 IP를 뽑아 /24, /16
// 단위로 빈도 집계하고, SYSTEM_AUTHORIZED_IPS(공식 인가 IP 목록)와 대조한다.
// "공식 목록에 없는데 자주 등장하는 대역"을 찾아내 목록 누락 여부를
// 사람이 검토할 수 있게 하는 게 목적 - 자동으로 목록을 고치지는 않는다.
//
// 사용법: node ip-usage-analysis.js [YYYY-MM ...]
//   인자 없이 실행하면 2026-07, 2026-08을 기본으로 시도한다.
//   각 달의 구글드라이브 폴더 참조(agit-automation state.json)가 이미
//   삭제된 경우(그 달 처리가 끝나 정리된 경우) 해당 달은 건너뛰고 이유를
//   리포트에 남긴다.
'use strict';

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const AnalysisCore = require(path.join(__dirname, 'public', 'analysis-core.js'));
const { getDriveFolderIdForMonth, downloadMonthFolderToTemp, cleanupTempDir } = require('./drive-helper.js');

const TOP_N = 30;
const OUTPUT_PATH = path.join(__dirname, 'data', 'ip-analysis-report.md');

function isAuthorized(ip) {
  return AnalysisCore.AUTHORIZED_IP_RULES.some(function(rule) { return AnalysisCore.matchesIPRule(ip, rule); });
}

// analyzeLog()와 동일한 방식(fmt.ip/fmt.user 컬럼 인덱스 + splitLine/getCol)으로
// 한 시트의 모든 행에서 (ip, user) 쌍을 뽑는다 - 헤더 줄이나 IP 컬럼이 없는
// 포맷은 isValidIPAddress()가 자연스럽게 걸러준다(따로 헤더 스킵 로직이 필요 없음).
function extractIpUserPairs(logText, formatKey) {
  const fmt = AnalysisCore.SYSTEM_FORMATS[formatKey] || AnalysisCore.SYSTEM_FORMATS.custom;
  const lines = logText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter(function(l) { return l.trim(); });
  const pairs = [];
  for (let i = 0; i < lines.length; i++) {
    const parts = AnalysisCore.splitLine(lines[i], fmt);
    const ip = AnalysisCore.getCol(parts, fmt.ip);
    if (!AnalysisCore.isValidIPAddress(ip)) continue;
    const user = AnalysisCore.getCol(parts, fmt.user) || 'unknown';
    pairs.push({ ip: ip, user: user });
  }
  return pairs;
}

async function collectPairsForMonth(month, log) {
  const folderId = await getDriveFolderIdForMonth(month).catch(function() { return null; });
  if (!folderId) {
    log.skippedMonths.push({ month: month, reason: '드라이브 폴더 참조 없음 (이미 처리 완료되어 정리되었을 수 있음)' });
    return [];
  }

  const { tmpDir } = await downloadMonthFolderToTemp(folderId);
  const pairs = [];
  try {
    const files = fs.readdirSync(tmpDir).filter(function(f) { return f.toLowerCase().endsWith('.xlsx') && !f.startsWith('~$'); });
    files.forEach(function(fileName) {
      const filePath = path.join(tmpDir, fileName);
      let workbook;
      try {
        workbook = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });
      } catch (error) {
        return;
      }
      workbook.SheetNames.forEach(function(sheetName) {
        const formatKey = AnalysisCore.resolveFormatKeyForSheet(sheetName, fileName) || AnalysisCore.getFormatKeyBySheetName(sheetName);
        if (!formatKey) return;
        const sheet = workbook.Sheets[sheetName];
        const logText = AnalysisCore.sheetToLogText(sheet);
        if (!logText.trim()) return;
        extractIpUserPairs(logText, formatKey).forEach(function(p) {
          pairs.push({ ip: p.ip, user: p.user, month: month, formatKey: formatKey });
        });
      });
    });
  } finally {
    await cleanupTempDir(tmpDir);
  }
  log.usedMonths.push(month);
  return pairs;
}

function bandKey24(ip) {
  const p = ip.split('.');
  return p[0] + '.' + p[1] + '.' + p[2] + '.0/24';
}

function bandKey16(ip) {
  const p = ip.split('.');
  return p[0] + '.' + p[1] + '.0.0/16';
}

function groupByBand(pairs, bandFn) {
  const groups = {};
  pairs.forEach(function(p) {
    const key = bandFn(p.ip);
    if (!groups[key]) groups[key] = { band: key, count: 0, authorizedCount: 0, users: new Set(), ips: new Set(), systemCounts: {} };
    const g = groups[key];
    g.count++;
    g.users.add(p.user);
    g.ips.add(p.ip);
    g.systemCounts[p.formatKey] = (g.systemCounts[p.formatKey] || 0) + 1;
    if (isAuthorized(p.ip)) g.authorizedCount++;
  });
  return Object.keys(groups).map(function(k) {
    const g = groups[k];
    const coverage = g.count ? g.authorizedCount / g.count : 0;
    let status;
    if (coverage >= 0.99) status = '확인됨';
    else if (coverage <= 0.01) status = '⚠️ 검토 필요 - 공식 목록에 없지만 빈번하게 사용됨';
    else status = '⚠️ 검토 필요 - 일부만 인가됨 (' + Math.round(coverage * 100) + '%)';
    // 이 대역에 가장 많이 기여한 시스템(포맷) 1~2개 - "이 대역이 특정 시스템
    // 하나에서만 나온 건지, 여러 사내 시스템에 걸쳐 나온 건지"를 바로 알 수
    // 있게 해서 검토 우선순위를 사람이 판단하기 쉽게 한다.
    const topSystems = Object.keys(g.systemCounts)
      .sort(function(a, b) { return g.systemCounts[b] - g.systemCounts[a]; })
      .slice(0, 2)
      .map(function(s) { return s + '(' + g.systemCounts[s] + ')'; })
      .join(', ');
    return {
      band: g.band, count: g.count, userCount: g.users.size,
      uniqueIpCount: g.ips.size, coveragePct: Math.round(coverage * 100), status: status,
      topSystems: topSystems,
    };
  }).sort(function(a, b) { return b.count - a.count; });
}

function summarizeRegisteredRules(pairs) {
  return AnalysisCore.AUTHORIZED_IP_RULES.map(function(rule) {
    let count = 0;
    const users = new Set();
    pairs.forEach(function(p) {
      if (AnalysisCore.matchesIPRule(p.ip, rule)) { count++; users.add(p.user); }
    });
    return { rule: rule, count: count, userCount: users.size };
  }).sort(function(a, b) { return b.count - a.count; });
}

function mdTable(headers, rows) {
  const lines = [];
  lines.push('| ' + headers.join(' | ') + ' |');
  lines.push('|' + headers.map(function() { return ' --- '; }).join('|') + '|');
  rows.forEach(function(r) { lines.push('| ' + r.join(' | ') + ' |'); });
  return lines.join('\n');
}

async function main() {
  const months = process.argv.slice(2).length ? process.argv.slice(2) : ['2026-07', '2026-08'];
  const log = { usedMonths: [], skippedMonths: [] };

  console.log('분석 대상 월(요청): ' + months.join(', '));
  let allPairs = [];
  for (const m of months) {
    console.log('- ' + m + ' 드라이브 폴더 확인 및 다운로드 중...');
    const beforeSkipCount = log.skippedMonths.length;
    const pairs = await collectPairsForMonth(m, log);
    if (log.skippedMonths.length > beforeSkipCount) {
      console.log('  -> 스킵: ' + log.skippedMonths[log.skippedMonths.length - 1].reason);
    } else {
      console.log('  -> ' + pairs.length + '건의 유효 IP 행 수집');
    }
    allPairs = allPairs.concat(pairs);
  }

  if (!allPairs.length) {
    console.log('\n수집된 데이터가 없습니다 (모든 달이 스킵됨). 리포트를 생성하지 않습니다.');
    console.log(JSON.stringify(log, null, 2));
    return;
  }

  const byBand24 = groupByBand(allPairs, bandKey24).slice(0, TOP_N);
  const byBand16 = groupByBand(allPairs, bandKey16).slice(0, TOP_N);
  const registeredSummary = summarizeRegisteredRules(allPairs);

  const generatedAt = new Date().toISOString();
  const lines = [];
  lines.push('# IP 사용 현황 분석 리포트');
  lines.push('');
  lines.push('- 생성 시각: ' + generatedAt);
  lines.push('- 분석 대상 월(요청): ' + months.join(', '));
  lines.push('- 실제 분석된 월: ' + (log.usedMonths.join(', ') || '없음'));
  if (log.skippedMonths.length) {
    lines.push('- 스킵된 월: ' + log.skippedMonths.map(function(s) { return s.month + ' (' + s.reason + ')'; }).join(', '));
  }
  lines.push('- 전체 유효 IP 행 수: ' + allPairs.length + '건');
  lines.push('');
  lines.push('## /24 단위 상위 ' + byBand24.length + '개 대역 (등장 건수 순)');
  lines.push('');
  lines.push(mdTable(
    ['순위', '대역', '건수', '계정 수', '고유 IP 수', '주요 시스템(건수)', '공식 목록 대조'],
    byBand24.map(function(g, i) { return [i + 1, g.band, g.count, g.userCount, g.uniqueIpCount, g.topSystems, g.status]; })
  ));
  lines.push('');
  lines.push('## /16 단위 상위 ' + byBand16.length + '개 대역 (등장 건수 순)');
  lines.push('');
  lines.push(mdTable(
    ['순위', '대역', '건수', '계정 수', '고유 IP 수', '주요 시스템(건수)', '공식 목록 대조'],
    byBand16.map(function(g, i) { return [i + 1, g.band, g.count, g.userCount, g.uniqueIpCount, g.topSystems, g.status]; })
  ));
  lines.push('');
  lines.push('## 공식 인가 IP 목록(SYSTEM_AUTHORIZED_IPS) 규칙별 실사용 빈도');
  lines.push('');
  lines.push('등록은 되어 있는데 실사용 빈도가 낮은 규칙은 "참고용"으로만 봐주세요 (틀렸다는 뜻이 아니라 그만큼 드물게 쓰인다는 의미).');
  lines.push('');
  lines.push(mdTable(
    ['규칙', '실제 매칭 건수', '계정 수', '비고'],
    registeredSummary.map(function(r) { return [r.rule, r.count, r.userCount, r.count === 0 ? '⚪ 이번 기간엔 미사용' : (r.count < 10 ? '참고용 (낮은 빈도)' : '')]; })
  ));
  lines.push('');
  lines.push('---');
  lines.push('이 리포트는 자동으로 SYSTEM_AUTHORIZED_IPS를 수정하지 않습니다. "⚠️ 검토 필요"로 표시된 대역은 누락된 정상 대역일 수도, 실제 비인가 접속일 수도 있으니 사람이 직접 판단해서 필요시 analysis-core.js의 AUTHORIZED_IP_RULES에 반영해주세요.');
  lines.push('');

  fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
  fs.writeFileSync(OUTPUT_PATH, lines.join('\n'), 'utf8');
  console.log('\n리포트 저장됨: ' + OUTPUT_PATH);

  console.log('\n=== /24 상위 15개 (콘솔 미리보기) ===');
  console.table(byBand24.slice(0, 15).map(function(g) {
    return { 대역: g.band, 건수: g.count, 계정수: g.userCount, 주요시스템: g.topSystems, 대조: g.status };
  }));

  console.log('\n=== /16 상위 15개 (콘솔 미리보기) ===');
  console.table(byBand16.slice(0, 15).map(function(g) {
    return { 대역: g.band, 건수: g.count, 계정수: g.userCount, 주요시스템: g.topSystems, 대조: g.status };
  }));
}

main().catch(function(error) {
  console.error('분석 중 오류가 발생했습니다:', error);
  process.exitCode = 1;
});
