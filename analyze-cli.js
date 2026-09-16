#!/usr/bin/env node
// 브라우저 없이 터미널에서 xlsx 접근 로그를 분석하는 CLI.
// 사용법: node analyze-cli.js <폴더경로> [YYYY-MM]
//   - <폴더경로>: xlsx 파일들이 들어있는 폴더 (하위 폴더는 스캔하지 않음)
//   - [YYYY-MM]: 결과 최상위 "month" 필드. 생략 시 파일명에서 YYYYMM 패턴을 찾아
//                가장 많이 등장하는 월을 사용하고, 그마저 없으면 오늘 날짜의 연월을 사용.
// 판정 로직은 public/analysis-core.js를 그대로 재사용하므로 브라우저 UI와 결과가 100% 동일하다.

'use strict';

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const AnalysisCore = require('./public/analysis-core.js');

// 브라우저의 "이번 달 기록으로 저장" 버튼이 쓰는 것과 동일한 파일.
// server.js의 readStats()/writeStats()와 같은 JSON 배열 구조를 그대로 따른다.
const MONTHLY_STATS_FILE = path.join(__dirname, 'data', 'monthly-stats.json');

function printUsageAndExit(message) {
  if (message) console.error(message);
  console.error('사용법: node analyze-cli.js <폴더경로> [YYYY-MM]');
  process.exit(1);
}

// "KC어드민_접속기록_202608.xlsx" → "KC어드민" (첫 '_' 앞부분. 없으면 파일명 전체)
// 시스템마다 파일명 규칙이 제각각이라 더 이상 systemName의 주 출처로 쓰지 않고,
// skipped(형식 미등록) 항목의 표시용 fallback으로만 사용한다.
function extractSystemName(fileName) {
  var base = fileName.replace(/\.[^/.]+$/, '');
  var idx = base.indexOf('_');
  return idx >= 0 ? base.slice(0, idx) : base;
}

// public/analysis-core.js의 SYSTEM_FORMATS에 등록된 모든 formatKey → 표시용
// 시스템명 매핑. 존(퍼블릭/공공, PI/CSAP 등)이 구분되는 시스템은 "이름(존)"으로,
// 아니면 "이름"으로 표기한다. SYSTEM_FORMATS에 새 formatKey를 추가했다면 여기도
// 같이 추가해야 하며, 빠뜨리면 resolveSystemName()이 콘솔에 경고를 출력한다.
var FORMAT_DISPLAY_NAMES = {
  'QueryPie-PI-UserLogin': 'QueryPie(PI)',
  'QueryPie-CSAP-UserLogin': 'QueryPie(CSAP)',
  'QueryPie-PI-UserMgmt': 'QueryPie(PI)',
  'QueryPie-PI-AdminActivity': 'QueryPie(PI)',
  'QueryPie-PI-Activation': 'QueryPie(PI)',
  'QueryPie-CSAP-UserMgmt': 'QueryPie(CSAP)',
  'QueryPie-CSAP-AdminActivity': 'QueryPie(CSAP)',
  'QueryPie-PI-AdminRoleHistory': 'QueryPie(PI)',
  'QueryPie-CSAP-AdminRoleHistory': 'QueryPie(CSAP)',
  'QueryPie-CSAP-Activation': 'QueryPie(CSAP)',
  'QueryPie-SpaceSeparated': 'QueryPie',
  'KCadmin-접속': 'KCadmin',
  'KCadmin-수행업무': 'KCadmin',
  'KCadmin-권한': 'KCadmin',
  'KCadmin-공공-접속': 'KCadmin(공공)',
  'KCadmin-공공-수행업무': 'KCadmin(공공)',
  'KCadmin-공공-권한': 'KCadmin(공공)',
  'KCadmin-퍼블릭-접속': 'KCadmin(퍼블릭)',
  'KCadmin-퍼블릭-수행업무': 'KCadmin(퍼블릭)',
  'KCadmin-퍼블릭-권한': 'KCadmin(퍼블릭)',
  'BTS-접근': 'BTS',
  'BTS-권한': 'BTS',
  'BTS-행위': 'BTS',
  'IAM-Access': 'IAM',
  'IAM-Action': 'IAM',
  'IAM-Permission': 'IAM',
  'IAM-공공-Access': 'IAM(공공)',
  'IAM-공공-Action': 'IAM(공공)',
  'IAM-공공-Permission': 'IAM(공공)',
  'IAM-퍼블릭-Access': 'IAM(퍼블릭)',
  'IAM-퍼블릭-Action': 'IAM(퍼블릭)',
  'IAM-퍼블릭-Permission': 'IAM(퍼블릭)',
  '빌링-접근': '빌링',
  '빌링-권한': '빌링',
  '빌링-행위': '빌링',
  'BTS-공공-접근': 'BTS(공공)',
  'BTS-공공-권한': 'BTS(공공)',
  'BTS-공공-행위': 'BTS(공공)',
  'BTS-퍼블릭-접근': 'BTS(퍼블릭)',
  'BTS-퍼블릭-권한': 'BTS(퍼블릭)',
  'BTS-퍼블릭-행위': 'BTS(퍼블릭)',
  '빌링-공공-접근': '빌링(공공)',
  '빌링-공공-권한': '빌링(공공)',
  '빌링-공공-행위': '빌링(공공)',
  '빌링-퍼블릭-접근': '빌링(퍼블릭)',
  '빌링-퍼블릭-권한': '빌링(퍼블릭)',
  '빌링-퍼블릭-행위': '빌링(퍼블릭)',
  '교육센터-액세스': '교육센터',
  '교육센터-액션': '교육센터',
  '교육센터-수료증': '교육센터',
  '교육센터-가입탈퇴': '교육센터',
  // 'iLaaS'와 'iLaaS-OM'은 같은 시스템의 존이 아니라 서로 다른 시스템(LaaS KEP / LaaS OM)
  // 이므로 하나로 뭉개지 않고 구분해서 표기한다.
  'iLaaS': 'iLaaS',
  'iLaaS-OM': 'iLaaS(OM)',
  '젠데스크': '젠데스크',
  '파트너센터-시스템접속': '파트너센터',
  '파트너센터-개인정보처리': '파트너센터',
  '파트너센터-권한': '파트너센터',
  // getFormatKeyBySheetName()은 'custom'을 자동 매칭 대상에서 제외하므로
  // 실제로는 도달하지 않지만, 매핑 누락 경고가 잘못 뜨지 않도록 등록해둔다.
  'custom': 'custom'
};

// formatKey 기준으로 표시용 systemName을 결정한다.
// formatKey가 없는 경우(형식 미등록 - 점검 제외)는 기존처럼 파일명 기반으로 표시한다.
function resolveSystemName(formatKey, fileName, sheetName) {
  if (!formatKey) {
    return extractSystemName(fileName);
  }
  if (Object.prototype.hasOwnProperty.call(FORMAT_DISPLAY_NAMES, formatKey)) {
    return FORMAT_DISPLAY_NAMES[formatKey];
  }
  console.error(
    '[WARN] formatKey "' + formatKey + '"에 대한 표시용 시스템명 매핑이 없습니다 ' +
    '(파일: ' + fileName + ', 시트: ' + sheetName + '). formatKey를 그대로 systemName으로 ' +
    '사용합니다. analyze-cli.js의 FORMAT_DISPLAY_NAMES에 매핑을 추가해주세요.'
  );
  return formatKey;
}

function extractMonthFromFileName(fileName) {
  var m = fileName.match(/(20\d{2})(0[1-9]|1[0-2])(?!\d)/);
  return m ? (m[1] + '-' + m[2]) : null;
}

// 젠데스크 감사로그처럼 한 파일에 "26.07"같은 "YY.MM" 형식의 시트가 여러 달치
// 함께 들어있는 경우, 분석 대상 월(targetMonth, "YYYY-MM")에 해당하는 시트
// 하나만 남기고 나머지 달 시트는 조용히 제외한다 (형식 미등록이 아니라 애초에
// 이번 분석 대상이 아니므로, 결과에도 남기지 않고 monthly-stats.json도 오염하지
// 않는다). "YY.MM" 패턴 시트가 아예 없는 파일은 그대로 둔다.
function filterMonthLabeledSheets(sheetNames, targetMonth) {
  var monthLabelPattern = /^\d{2}\.\d{2}$/;
  var hasMonthLabeledSheet = sheetNames.some(function(sn) { return monthLabelPattern.test(sn); });
  if (!hasMonthLabeledSheet) return sheetNames;

  var targetLabel = targetMonth.slice(2, 4) + '.' + targetMonth.slice(5, 7);
  return sheetNames.filter(function(sn) {
    return !monthLabelPattern.test(sn) || sn === targetLabel;
  });
}

function resolveMonth(explicitMonth, files) {
  if (explicitMonth) {
    if (!/^\d{4}-\d{2}$/.test(explicitMonth)) {
      printUsageAndExit('month 인자는 YYYY-MM 형식이어야 합니다: ' + explicitMonth);
    }
    return explicitMonth;
  }
  var counts = {};
  files.forEach(function(f) {
    var m = extractMonthFromFileName(f);
    if (m) counts[m] = (counts[m] || 0) + 1;
  });
  var best = null, bestCount = 0;
  Object.keys(counts).forEach(function(m) {
    if (counts[m] > bestCount) { best = m; bestCount = counts[m]; }
  });
  if (best) return best;
  var now = new Date();
  return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
}

// hasIssue는 실제 행위/데이터 기반 이상 탐지 항목에만 반응한다:
//   로그인 연속실패, 업무시간외 접근, 비인가 IP, 다수 IP, 다운로드 행위,
//   과다 조회, 계정/접속일시/IP/수행업무 값 누락(행 단위 실제 결측)
// 아래는 의도적으로 제외한다:
//   - result.logIssue: LOG_ISSUES(예: "로그인 실패 이력 로깅되지 않음")나 "필수항목
//     누락 있음" 같은, 로그 형식 자체의 구조적 한계를 알리는 문구라서 실제 데이터에
//     이상이 없어도 항상 true가 되어버림 (노이즈).
//   - result.logSchema: "로그인 성공/실패 여부", "다운로드 사유" 등 형식 정의에 아예
//     없는 컬럼을 알려주는 스키마 정보이며 result.logMissing(행 단위 결측)과는 별개.
//   - unauthorizedAccess / unauthorizedPermissions / retentionSummary: 이 CLI의
//     핵심 점검 항목(로그인연속실패·업무외시간·비인가IP·다수IP·다운로드·과다조회·
//     로그누락) 밖에 있는 부가 항목이라 hasIssue 판정에서 제외했다. 필요하면
//     unauthorizedAccess/unauthorizedPermissions/retention 필드로 별도 확인 가능.
function hasIssue(result) {
  return !!(
    result.failGroups.length ||
    result.ahSummary.length ||
    result.uaSummary.length ||
    result.multiSummary.length ||
    result.dlSummary.length ||
    result.overQueries.length ||
    (result.logMissing && result.logMissing.length)
  );
}

// ── 월별 다운로드 추이 자동 저장 ──────────────────────────────────────────
// 브라우저 "이번 달 기록으로 저장" 버튼과 동일한 스키마로 data/monthly-stats.json에
// 누적 저장하고, 직전 3개월 평균 대비 150% 초과 여부를 함께 계산한다.

function readMonthlyStats() {
  try {
    var parsed = JSON.parse(fs.readFileSync(MONTHLY_STATS_FILE, 'utf8'));
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function writeMonthlyStats(stats) {
  fs.mkdirSync(path.dirname(MONTHLY_STATS_FILE), { recursive: true });
  fs.writeFileSync(MONTHLY_STATS_FILE, JSON.stringify(stats, null, 2) + '\n', 'utf8');
}

// 브라우저의 resultAnomalyCounts()와 동일한 항목/집계 방식 (다운로드 건수 등
// monthly-stats.json에 이미 저장돼 있을 수 있는 값과 일관성을 유지하기 위함).
function anomalyCounts(result) {
  return {
    loginFailures: result.failGroups.reduce(function(t, g) { return t + g.length; }, 0),
    afterHours: result.ahSummary.reduce(function(t, s) { return t + s.count; }, 0),
    unauthorizedIPs: result.uaSummary.reduce(function(t, s) { return t + s.count; }, 0),
    multiIPs: result.multiSummary.length,
    downloads: result.dlSummary.reduce(function(t, s) { return t + s.count; }, 0),
    unauthorizedAccess: result.unauthorizedAccess.length,
    unauthorizedPermissions: result.unauthorizedPermissions.length,
    logMissing: result.logMissing.length,
    overQueries: result.overQueries.length,
    retention: result.retentionSummary.length
  };
}

// 시스템명(파일명의 날짜 부분을 뗀 것) + 시트명을 키로 써서 매달 동일하게 유지되도록 한다.
// (파일명 자체를 키로 쓰면 "KC어드민_접속기록_202608.xlsx"처럼 매달 파일명이 달라져서
//  직전 달 기록과 절대 매칭되지 않는다.)
function statsSystemKey(systemName, sheetName) {
  return systemName + ' - ' + sheetName;
}

function upsertMonthlyStat(stats, record) {
  var index = stats.findIndex(function(s) { return s.month === record.month && s.system === record.system; });
  if (index >= 0) stats[index] = record;
  else stats.push(record);
  stats.sort(function(a, b) { return a.month.localeCompare(b.month) || a.system.localeCompare(b.system); });
}

// 대상 월 기준 직전 3개월(연속) 데이터가 모두 있어야 평균을 계산한다 (브라우저
// averagePreviousThree()와 동일한 규칙).
function averagePreviousThreeMonths(stats, systemKey, targetMonth) {
  var parts = targetMonth.split('-');
  var year = Number(parts[0]);
  var monthNum = Number(parts[1]);
  var previous = [];
  for (var i = 1; i <= 3; i++) {
    var d = new Date(Date.UTC(year, monthNum - 1 - i, 1));
    var prevMonth = d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
    var item = stats.find(function(s) { return s.system === systemKey && s.month === prevMonth; });
    if (item) previous.push(item.downloadCount);
  }
  if (previous.length < 3) return null;
  return previous.reduce(function(t, c) { return t + c; }, 0) / 3;
}

function computeDownloadTrend(stats, systemKey, month, thisMonthCount) {
  var avg3m = averagePreviousThreeMonths(stats, systemKey, month);
  if (avg3m === null) {
    return {
      thisMonth: thisMonthCount,
      avg3m: null,
      percentage: null,
      exceeds150: false,
      insufficientData: true,
      label: '데이터 부족 - 비교 불가'
    };
  }
  var percentage = avg3m > 0 ? Math.round((thisMonthCount / avg3m) * 100) : null;
  var exceeds150 = avg3m > 0 ? thisMonthCount > avg3m * 1.5 : thisMonthCount > 0;
  var label = '정상 범위';
  if (exceeds150) {
    label = percentage !== null
      ? '다운로드 급증 - 3개월 평균 대비 ' + percentage + '% 초과'
      : '다운로드 급증 - 직전 3개월 평균 0건 대비 신규 발생';
  }
  return {
    thisMonth: thisMonthCount,
    avg3m: Math.round(avg3m * 10) / 10,
    percentage: percentage,
    exceeds150: exceeds150,
    insufficientData: false,
    label: label
  };
}

// analyzeLog()가 반환하는 원시 결과를 요구된 CLI 출력 구조로 변환한다.
function toCliRecord(result, meta) {
  return {
    systemName: meta.systemName,
    fileName: meta.fileName,
    sheetName: result.sheetName,
    formatKey: result.formatKey,
    skipped: !!result.skipped,
    skipReason: result.skipped ? result.logIssue : null,
    loginFailures: {
      count: result.failGroups.reduce(function(t, g) { return t + g.length; }, 0),
      groups: result.failGroups
    },
    afterHours: {
      count: result.ahSummary.reduce(function(t, s) { return t + s.count; }, 0),
      details: result.ahSummary
    },
    unauthorizedIP: {
      note: result.ipCheckNote || null,
      count: result.uaSummary.reduce(function(t, s) { return t + s.count; }, 0),
      details: result.uaSummary
    },
    multiIP: {
      count: result.multiSummary.length,
      details: result.multiSummary
    },
    download: {
      count: result.dlSummary.reduce(function(t, s) { return t + (s.rowCount !== undefined ? s.rowCount : s.count); }, 0),
      details: result.dlSummary
    },
    unauthorizedAccess: result.unauthorizedAccess,
    unauthorizedPermissions: result.unauthorizedPermissions,
    overQueries: result.overQueries,
    logMissing: {
      schema: result.logSchema,
      issues: result.logMissing
    },
    retention: {
      details: result.retentionSummary
    },
    logIssue: result.logIssue || null,
    // 형식 미등록/빈 시트로 건너뛴 시트는 "점검 불가"이지 "이상"이 아니므로
    // hasIssue를 true/false가 아닌 null로 표시해 이상탐지 대상에서 완전히 제외한다.
    hasIssue: result.skipped ? null : hasIssue(result)
  };
}

function main() {
  var folder = process.argv[2];
  var explicitMonth = process.argv[3];
  if (!folder) printUsageAndExit();

  var resolvedFolder = path.resolve(folder);
  var stat;
  try {
    stat = fs.statSync(resolvedFolder);
  } catch (error) {
    printUsageAndExit('폴더를 찾을 수 없습니다: ' + resolvedFolder);
  }
  if (!stat.isDirectory()) {
    printUsageAndExit('폴더 경로가 아닙니다: ' + resolvedFolder);
  }

  var files = fs.readdirSync(resolvedFolder)
    .filter(function(f) { return f.toLowerCase().endsWith('.xlsx') && !f.startsWith('~$'); })
    .sort();

  var month = resolveMonth(explicitMonth, files);
  var systems = [];
  var monthlyStats = readMonthlyStats();

  files.forEach(function(fileName) {
    var filePath = path.join(resolvedFolder, fileName);
    var workbook;
    try {
      workbook = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });
    } catch (error) {
      console.error('xlsx 파일을 읽지 못했습니다: ' + fileName + ' (' + error.message + ')');
      return;
    }
    var sheetNames = filterMonthLabeledSheets(workbook.SheetNames, month);
    var results = AnalysisCore.analyzeWorkbook(workbook, sheetNames, fileName);
    results.forEach(function(result) {
      // systemName은 이제 파일명이 아니라 이미 정확히 매칭된 formatKey를 기준으로
      // 결정한다 (파일명 규칙이 시스템마다 제각각이라 파일명 파싱으로는 안정적으로
      // 시스템을 구분할 수 없기 때문). 원본 파일명은 fileName 필드에 그대로 남는다.
      var systemName = resolveSystemName(result.formatKey, fileName, result.sheetName);
      var record = toCliRecord(result, { systemName: systemName, fileName: fileName });

      if (!result.skipped) {
        var systemKey = statsSystemKey(systemName, result.sheetName);
        var counts = anomalyCounts(result);
        upsertMonthlyStat(monthlyStats, {
          month: month,
          system: systemKey,
          downloadCount: counts.downloads,
          anomalies: counts,
          savedAt: new Date().toISOString()
        });
        record.downloadTrend = computeDownloadTrend(monthlyStats, systemKey, month, counts.downloads);
      }

      systems.push(record);
    });
  });

  writeMonthlyStats(monthlyStats);

  console.log(JSON.stringify({ month: month, systems: systems }, null, 2));
}

main();
