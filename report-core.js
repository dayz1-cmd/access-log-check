// report-core.js
// agit-automation/analyze_and_post.py의 9개 항목 리포트 생성 로직을 그대로 이식한
// Node 모듈. index.html의 "월간 점검 리포트" 탭(서버 API용)과, 필요하면 다른
// 서버 스크립트에서도 재사용할 수 있도록 analysis-core.js처럼 순수 함수로 둔다.
'use strict';

const path = require('path');
// 타임스탬프 표시 형식(엑셀 시리얼/각종 로그 포맷 -> "YYYY-MM-DD HH:mm:ss")을
// 브라우저 결과 테이블과 동일하게 맞추기 위해 analysis-core.js의 tsDisplay를 그대로 쓴다.
const AnalysisCore = require(path.join(__dirname, 'public', 'analysis-core.js'));
const tsDisplay = AnalysisCore.tsDisplay;

// public/analysis-core.js SYSTEM_FORMATS 중 rowCountCol(조회 건수 컬럼)이 정의된
// 포맷만 "과도한 조회(target_id/row_count)" 판정이 실제로 가능하다.
// (agit-automation/analyze_and_post.py의 ROW_COUNT_FORMATS와 반드시 동기화할 것)
const ROW_COUNT_FORMATS = new Set([
  'BTS-행위', 'BTS-공공-행위', 'BTS-퍼블릭-행위',
  'iLaaS-OM',
  '파트너센터-개인정보처리',
  '빌링-공공-행위', '빌링-퍼블릭-행위',
]);

// logMissing.schema의 present:false 항목을 사람이 읽을 문구로 매핑.
const LOG_MISSING_LABEL_TEXT = {
  '로그인 성공/실패 여부': '(로그인 실패 로그가 없음)',
  '다운로드 사유': '(다운로드 사유 로그가 없음)',
};
const ROW_FIELD_ORDER = ['계정', '접속일시', 'IP', '수행업무'];

// 기본 담당자 멘션(폴백용). 실제 값은 이제 호출하는 쪽(server.js)이
// options.assignees로 넘긴다 - UI의 "담당자 멘션" 입력창에서 직접 받는다.
const ASSIGNEE_MENTION = '@han.b';

// detail은 [3-2]/[3-3]/[3-4]/[3-5]/[3-6]에만 있다 (담당자가 실제로 어떤 계정이
// 이상행위였는지 보고 코멘트를 쓸 수 있도록). [3-1]/[3-7]/[3-8]/[3-9]는 요청된
// 범위가 아니라 건드리지 않는다.
const SECTION_DEFS = [
  { id: 1, title: '권한이 없는자의 접속행위', check: checkUnauthorizedAccess },
  { id: 2, title: '잦은 로그인 실패', check: checkLoginFailures, detail: detailLoginFailures },
  { id: 3, title: '업무시간 외 접속 행위', check: checkAfterHours, detail: detailAfterHours },
  { id: 4, title: '비인가 IP 접속 행위', check: checkUnauthorizedIp, detail: detailUnauthorizedIp },
  { id: 5, title: '1개의 계정에서 다수 IP 접속행위', check: checkMultiIp, detail: detailMultiIp },
  { id: 6, title: '개인정보 다운로드 여부', check: checkDownload, detail: detailDownload },
  { id: 7, title: '비인가된 자가 권한업무를 진행한 여부', check: checkUnauthorizedPermissions },
  { id: 8, title: '누락된 로그 여부', check: checkLogMissing },
];

// ── 아지트 마크다운 규칙: 볼드는 별표 1개, 인용구 박스는 각 줄 앞에 "> " ──────
function quoteBlock(text) {
  return text.split('\n').map(function(line) { return line ? '> ' + line : '>'; }).join('\n');
}

function groupByDisplayName(systems) {
  const groups = {};
  systems.forEach(function(s) {
    const name = s.systemName || s.fileName || '알수없음';
    if (!groups[name]) groups[name] = [];
    groups[name].push(s);
  });
  return groups;
}

function nonSkipped(entries) {
  return entries.filter(function(e) { return !e.skipped; });
}

function scopeLists(groups) {
  const kko = [], ilaas = [];
  Object.keys(groups).sort().forEach(function(name) {
    if (name.indexOf('LaaS') >= 0) ilaas.push(name);
    else kko.push(name);
  });
  return { kko: kko, ilaas: ilaas };
}

// ---------------------------------------------------------------------------
// [3-2]/[3-3]/[3-4]/[3-5]/[3-6] 세부 내역 - analyze-cli.js가 이미 넣어주는
// groups/details를 그대로 꺼내 사람이 읽을 행 목록으로 펼친다. 담당자가 어떤
// 계정이 이상행위였는지 보고 코멘트를 정확히 쓸 수 있게 하는 게 목적이라
// 여기서는 집계하지 않고 원본 단위(그룹/사용자별)를 그대로 반환한다.
// ---------------------------------------------------------------------------

function detailLoginFailures(entries) {
  const es = nonSkipped(entries);
  const rows = [];
  es.forEach(function(e) {
    ((e.loginFailures || {}).groups || []).forEach(function(g) {
      if (!g.length) return;
      rows.push({
        user: g[0].user,
        count: g.length,
        first: tsDisplay(g[0].ts),
        last: tsDisplay(g[g.length - 1].ts),
      });
    });
  });
  return rows;
}

function detailAfterHours(entries) {
  const es = nonSkipped(entries);
  const rows = [];
  es.forEach(function(e) {
    ((e.afterHours || {}).details || []).forEach(function(s) {
      if (s.dateGroups && s.dateGroups.length) {
        s.dateGroups.forEach(function(dg) {
          rows.push({ user: s.user, count: dg.count, first: tsDisplay(dg.first), last: tsDisplay(dg.last) });
        });
      } else {
        rows.push({ user: s.user, count: s.count, first: tsDisplay(s.first), last: tsDisplay(s.last) });
      }
    });
  });
  return rows;
}

function detailUnauthorizedIp(entries) {
  const es = nonSkipped(entries);
  const rows = [];
  es.forEach(function(e) {
    ((e.unauthorizedIP || {}).details || []).forEach(function(s) {
      rows.push({ user: s.user, ip: s.ip, count: s.count, first: tsDisplay(s.first), last: tsDisplay(s.last) });
    });
  });
  return rows;
}

function detailMultiIp(entries) {
  const es = nonSkipped(entries);
  const rows = [];
  es.forEach(function(e) {
    ((e.multiIP || {}).details || []).forEach(function(s) {
      rows.push({ user: s.user, ips: s.ips || [] });
    });
  });
  return rows;
}

function detailDownload(entries) {
  const es = nonSkipped(entries);
  const rows = [];
  es.forEach(function(e) {
    ((e.download || {}).details || []).forEach(function(s) {
      const count = s.rowCount !== undefined ? s.rowCount : s.count;
      rows.push({ user: s.user, count: count, first: tsDisplay(s.first), last: tsDisplay(s.last) });
    });
  });
  return rows;
}

// ---------------------------------------------------------------------------
// [3-1]~[3-9] 항목별 시스템 한 줄 판정. 모두 { base, needsComment, trend } 반환.
// base: 코멘트/추이 표시를 뺀 순수 판정 문구
// needsComment: true면 "[[코멘트 입력]]" 자리에 실제 입력값이 들어가야 함
// trend: base와 needsComment 뒤에 추가로 붙는 부가 설명(다운로드 추이 등), 없으면 ''
// ---------------------------------------------------------------------------

function checkUnauthorizedAccess(entries) {
  const es = nonSkipped(entries);
  const total = es.reduce(function(t, e) { return t + (e.unauthorizedAccess || []).length; }, 0);
  if (total === 0) return { base: '특이사항 없음', needsComment: false, trend: '' };
  return { base: total + '건 확인', needsComment: true, trend: '' };
}

function checkLoginFailures(entries) {
  const es = nonSkipped(entries);
  let groupsAll = [];
  es.forEach(function(e) { groupsAll = groupsAll.concat(((e.loginFailures || {}).groups) || []); });
  const total = groupsAll.reduce(function(t, g) { return t + g.length; }, 0);
  if (total === 0) return { base: '특이사항 없음', needsComment: false, trend: '' };
  const users = new Set(groupsAll.filter(function(g) { return g.length; }).map(function(g) { return g[0].user; }));
  return { base: users.size + '개 계정 총 ' + total + '회 오류', needsComment: true, trend: '' };
}

function checkAfterHours(entries) {
  const es = nonSkipped(entries);
  const users = new Set();
  es.forEach(function(e) {
    ((e.afterHours || {}).details || []).forEach(function(d) { users.add(d.user); });
  });
  if (users.size === 0) return { base: '특이사항 없음', needsComment: false, trend: '' };
  return { base: users.size + '명', needsComment: true, trend: '' };
}

function firstIpNote(es) {
  for (let i = 0; i < es.length; i++) {
    const note = (es[i].unauthorizedIP || {}).note;
    if (note) return note;
  }
  return null;
}

function checkUnauthorizedIp(entries) {
  const es = nonSkipped(entries);
  const note = firstIpNote(es);
  if (note) return { base: note, needsComment: false, trend: '' };
  const total = es.reduce(function(t, e) { return t + ((e.unauthorizedIP || {}).count || 0); }, 0);
  if (total === 0) return { base: '특이사항 없음', needsComment: false, trend: '' };
  return { base: total + '건 확인', needsComment: true, trend: '' };
}

function checkMultiIp(entries) {
  const es = nonSkipped(entries);
  // multiIP는 자체 note가 없지만 unauthorizedIP.note(ipCheckNote)와 같은
  // hasConfiguredAuthorizedIPs() 게이트를 공유하므로 재사용한다 (안 그러면
  // 인가 IP 미설정 시스템이 "0건 -> 특이사항 없음"으로 잘못 표시된다).
  const note = firstIpNote(es);
  if (note) return { base: note, needsComment: false, trend: '' };
  const total = es.reduce(function(t, e) { return t + ((e.multiIP || {}).count || 0); }, 0);
  if (total === 0) return { base: '특이사항 없음', needsComment: false, trend: '' };
  return { base: total + '건 확인', needsComment: true, trend: '' };
}

function checkDownload(entries) {
  const es = nonSkipped(entries);
  let total = 0;
  const users = new Set();
  es.forEach(function(e) {
    const d = e.download || {};
    total += d.count || 0;
    (d.details || []).forEach(function(det) { users.add(det.user); });
  });

  const trendNotes = [];
  es.forEach(function(e) {
    const trend = e.downloadTrend;
    if (trend && trend.exceeds150) {
      trendNotes.push('(이번 달 ' + trend.thisMonth + '건, 직전 3개월 평균 대비 ' + trend.percentage + '% - 급증 확인됨)');
    }
  });
  const trend = trendNotes.length ? ', ' + trendNotes.join(', ') : '';

  if (total === 0) return { base: '특이사항 없음', needsComment: false, trend: trend };
  return { base: users.size + '명 사용자 ' + total + '건 확인', needsComment: true, trend: trend };
}

function checkUnauthorizedPermissions(entries) {
  const es = nonSkipped(entries);
  const total = es.reduce(function(t, e) { return t + (e.unauthorizedPermissions || []).length; }, 0);
  if (total === 0) return { base: '특이사항 없음', needsComment: false, trend: '' };
  return { base: total + '건 확인', needsComment: true, trend: '' };
}

function checkLogMissing(entries) {
  const es = nonSkipped(entries);
  const notes = [];
  const seenLabels = new Set();
  es.forEach(function(e) {
    ((e.logMissing || {}).schema || []).forEach(function(item) {
      if (!item.present && !seenLabels.has(item.label)) {
        seenLabels.add(item.label);
        notes.push(LOG_MISSING_LABEL_TEXT[item.label] || ('(' + item.label + ' 없음)'));
      }
    });
  });

  let rowLines = [];
  es.forEach(function(e) {
    rowLines = rowLines.concat(((e.logMissing || {}).issues || []).filter(function(line) {
      return line.indexOf('행 ') === 0;
    }));
  });

  if (rowLines.length) {
    const fields = new Set();
    rowLines.forEach(function(line) {
      const idx = line.indexOf(':');
      let body = idx >= 0 ? line.slice(idx + 1).trim() : line;
      body = body.replace(/\s*누락$/, '');
      body.split(',').forEach(function(f) {
        f = f.trim();
        if (f) fields.add(f);
      });
    });
    const ordered = ROW_FIELD_ORDER.filter(function(f) { return fields.has(f); })
      .concat(Array.from(fields).filter(function(f) { return ROW_FIELD_ORDER.indexOf(f) < 0; }).sort());
    notes.push('(' + rowLines.length + '건 ' + ordered.join('/') + ' 누락)');
  }

  // check_log_missing()은 [[코멘트 입력]] 마커를 절대 붙이지 않는다 (요청된
  // 원본 템플릿 예시가 이미 구체적인 설명을 담고 있어서).
  if (!notes.length) return { base: '특이사항 없음', needsComment: false, trend: '' };
  return { base: notes.join(', '), needsComment: false, trend: '' };
}

function checkOverQueries(entries) {
  const es = nonSkipped(entries);
  const total = es.reduce(function(t, e) { return t + (e.overQueries || []).length; }, 0);
  if (total === 0) return { base: '특이사항 없음', needsComment: false, trend: '' };
  return { base: total + '건 확인', needsComment: true, trend: '' };
}

function overQueriesCapabilityMissing(allCheckableEntries) {
  return allCheckableEntries.some(function(e) { return !ROW_COUNT_FORMATS.has(e.formatKey); });
}

// [3-9] 안내 문구에 "일부 시스템"이라고만 뭉뚱그리지 않고 실제로 그 달 어떤
// 시스템이 target_id/row_count 필드가 없는지 이름을 나열하기 위해, 시스템
// (표시 이름) 단위로 ROW_COUNT_FORMATS 미보유 여부를 판정한다.
function systemsMissingRowCount(names, checkableGroups) {
  return names.filter(function(name) {
    return nonSkipped(checkableGroups[name]).some(function(e) { return !ROW_COUNT_FORMATS.has(e.formatKey); });
  });
}

function isCleanItem(it) {
  return !it.needsComment && !it.trend && (it.base === '특이사항 없음' || it.base.endsWith('점검 제외'));
}

// [3-2]~[3-6]처럼 계정별 세부 내역(detailFn)이 있는 항목은 코멘트 입력칸도
// "시스템×항목" 하나가 아니라 "시스템×항목×계정" 단위로 쪼갠다 - 한 항목에
// 여러 계정이 걸리면 서로 다른 사람의 소명이 코멘트 칸 하나를 두고 덮어쓰던
// 문제를 구조적으로 없앤다. 계정 판정은 index.html의 LDAP 그룹핑과 동일한
// 규칙(AnalysisCore의 extractLdap/isSystemAccount/isUuidAccount)을 그대로 쓴다.
function accountBucketKey(user) {
  if (!user || !String(user).trim()) return { key: 'unidentified', label: '(계정 미상)', kind: 'unidentified' };
  if (AnalysisCore.isSystemAccount(user)) {
    const base = AnalysisCore.systemAccountBase(user);
    return { key: 'sys:' + base, label: base, kind: 'system' };
  }
  if (AnalysisCore.isUuidAccount(user)) return { key: 'uuid', label: 'UUID 계정', kind: 'uuid' };
  const ldap = AnalysisCore.extractLdap(user);
  return { key: 'ldap:' + ldap, label: ldap, kind: 'human' };
}

function groupDetailsByAccount(details) {
  const order = [];
  const byKey = {};
  details.forEach(function(row) {
    const b = accountBucketKey(row.user);
    if (!byKey[b.key]) { byKey[b.key] = { key: b.key, label: b.label, kind: b.kind, rows: [] }; order.push(b.key); }
    byKey[b.key].rows.push(row);
  });
  // human 계정을 이름순으로 먼저, 그 다음 system/uuid/미상 순으로 정렬한다
  // (브라우저 쪽 표시 순서와 사람이 읽기 좋은 순서를 맞추기 위함).
  const kindRank = { human: 0, system: 1, uuid: 2, unidentified: 3 };
  order.sort(function(a, b) {
    const ba = byKey[a], bb = byKey[b];
    if (kindRank[ba.kind] !== kindRank[bb.kind]) return kindRank[ba.kind] - kindRank[bb.kind];
    return ba.label.localeCompare(bb.label);
  });
  return order.map(function(k) { return byKey[k]; });
}

function buildListSection(id, title, names, checkableGroups, checkFn, detailFn) {
  const items = names.map(function(name) {
    const r = checkFn(checkableGroups[name]);
    const item = { system: name, base: r.base, needsComment: r.needsComment, trend: r.trend, commentKey: '3-' + id + ':' + name };
    // 상세 내역은 실제로 코멘트가 필요한(=이상이 발견된) 행에만 계산해서 붙인다.
    if (detailFn && r.needsComment) {
      item.details = detailFn(checkableGroups[name]);
      item.accounts = groupDetailsByAccount(item.details).map(function(b) {
        return { key: b.key, label: b.label, kind: b.kind, commentKey: item.commentKey + '::' + b.key, rows: b.rows };
      });
    }
    return item;
  });
  const clean = items.length === 0 || items.every(isCleanItem);
  const emoji = clean ? '✅' : '❗';
  const status = clean ? '특이사항 없음' : '추가 검토 필요';
  return { id: id, title: title, kind: 'list', emoji: emoji, status: status, items: items };
}

// ---------------------------------------------------------------------------
// "점검 이력 분석" 탭(data/monthly-anomaly-history.json)용 - [3-2]~[3-9] 8개
// 항목의 시스템별 hasIssue(true/false/null)를 뽑아낸다. [3-1]은 요청 범위 밖.
// ---------------------------------------------------------------------------
const HISTORY_CATEGORY_KEYS = {
  2: 'loginFailures',
  3: 'afterHours',
  4: 'unauthorizedIP',
  5: 'multiIP',
  6: 'download',
  7: 'unauthorizedPermissions',
  8: 'logMissing',
  9: 'overQueries',
};

// true: 이상 발견(코멘트 필요 또는 [3-8]처럼 "특이사항 없음"이 아닌 구체적 사유가 있음)
// false: 특이사항 없음(정상 확인됨)
// null: 점검 자체가 불가(예: "인가 IP 목록 미설정 - 점검 제외", [3-9] 필드 미비)
function deriveItemHasIssue(item) {
  if (item.base.endsWith('점검 제외')) return null;
  if (!item.needsComment && !item.trend && item.base === '특이사항 없음') return false;
  return true;
}

function deriveAnomalyRecords(month, systems) {
  const report = computeReport(month, systems);
  const bySystem = {};

  report.sections.forEach(function(sec) {
    const key = HISTORY_CATEGORY_KEYS[sec.id];
    if (!key) return; // [3-1]은 이력 저장 대상이 아니다.

    if (sec.kind === 'paragraph') {
      // [3-9]가 "필드 미비로 판단 불가" 문단 모드일 때 - 리포트 전체에 걸친 판정이라
      // 이미 앞선 섹션들에서 채워진 시스템 전부에 대해 null로 기록한다.
      Object.keys(bySystem).forEach(function(sysName) { bySystem[sysName][key] = null; });
      return;
    }

    sec.items.forEach(function(it) {
      if (!bySystem[it.system]) bySystem[it.system] = { month: month, system: it.system };
      bySystem[it.system][key] = deriveItemHasIssue(it);
    });
  });

  return Object.keys(bySystem).map(function(k) { return bySystem[k]; });
}

// ---------------------------------------------------------------------------
// 계정명 -> 실제 LDAP 매핑 (data/account-aliases.json, server.js가 읽어서
// 넘겨준다). 일부 시스템 로그는 계정이 축약되거나 다르게 기록돼 있어서
// (예: 로그엔 "elvin", 실제 LDAP은 "elvin.h") 실제 LDAP과 안 맞는 경우가
// 있다 - computeReport()가 각 항목을 판정하기 전에, systems 원본에서
// 계정명이 나오는 자리를 전부 치환해서 이후 모든 처리(리포트/소명 요청
// 텍스트/이력)가 실제 LDAP 기준으로 일관되게 동작하게 한다.
function applyAccountAliases(systems, aliases) {
  if (!aliases || !Object.keys(aliases).length) return systems;
  function alias(user) {
    return Object.prototype.hasOwnProperty.call(aliases, user) ? aliases[user] : user;
  }
  function aliasRow(row) {
    return Object.assign({}, row, { user: alias(row.user) });
  }
  return systems.map(function(e) {
    const copy = Object.assign({}, e);
    if (copy.loginFailures && copy.loginFailures.groups) {
      copy.loginFailures = Object.assign({}, copy.loginFailures, {
        groups: copy.loginFailures.groups.map(function(g) { return g.map(aliasRow); }),
      });
    }
    ['afterHours', 'unauthorizedIP', 'multiIP', 'download'].forEach(function(key) {
      if (copy[key] && copy[key].details) {
        copy[key] = Object.assign({}, copy[key], { details: copy[key].details.map(aliasRow) });
      }
    });
    return copy;
  });
}

// ---------------------------------------------------------------------------
// 리포트 전체 구조 계산 (comments 주입 전 - 모델/텍스트 양쪽에서 재사용)
// ---------------------------------------------------------------------------
function computeReport(month, systems) {
  const parts = month.split('-');
  const year = parts[0], mon = parts[1];

  const groups = groupByDisplayName(systems);
  const checkableGroups = {};
  Object.keys(groups).forEach(function(name) {
    if (groups[name].some(function(e) { return !e.skipped; })) checkableGroups[name] = groups[name];
  });
  const names = Object.keys(checkableGroups).sort();

  const sections = SECTION_DEFS.map(function(def) {
    return buildListSection(def.id, def.title, names, checkableGroups, def.check, def.detail);
  });

  let allEntries = [];
  names.forEach(function(name) { allEntries = allEntries.concat(nonSkipped(checkableGroups[name])); });

  let section9;
  if (overQueriesCapabilityMissing(allEntries)) {
    const missingNames = systemsMissingRowCount(names, checkableGroups);
    section9 = {
      id: 9,
      title: '특정 정보주체 과도한 조회 여부',
      kind: 'paragraph',
      emoji: '❗',
      status: '추가 검토 필요',
      paragraphLines: [
        '다음 시스템은 개인정보 조회 대상(target_id) 또는 조회 건수(row_count) 필드가 없어',
        '특정 정보주체에 대한 과도한 조회 여부 판단이 불가함: ' + missingNames.join(', '),
      ],
      needsComment: true,
      commentKey: '3-9',
      commentPlaceholder: '조회 기준 및 검토 방법 관련 추가 코멘트',
    };
  } else {
    section9 = buildListSection(9, '특정 정보주체 과도한 조회 여부', names, checkableGroups, checkOverQueries);
  }
  sections.push(section9);

  const scope = scopeLists(groups);

  return {
    month: month, year: year, mon: mon,
    scope: {
      kko: scope.kko.length ? scope.kko.join(', ') : '해당 없음',
      ilaas: scope.ilaas.length ? scope.ilaas.join(', ') : '해당 없음',
    },
    sections: sections,
  };
}

const SECTION2_LINES = [
  '로그점검은 아래 점검 기준에 따라 점검이 진행됩니다.',
  '1️⃣ 접근이력',
  '    ☞ 권한이 없는자의 접속행위 : 취급자 목록과 접속기록 내 접속한 계정과 비교 확인 등',
  '    ☞ 비인가 IP 접속 행위 : VDI IP 대역 등 허용된 IP 외의 접속기록 확인',
  '    ☞ 다수 IP 접속 행위 : 하나의 계정에서 복수의 IP 접속기록 확인',
  '    ☞ 잦은 로그인 실패 : 20분 내 5회이상 연속 로그인 실패 이력 확인',
  '    ☞ 누락된 로그 여부 : 접근이력과 행위이력과의 접근 계정, 시간 등을 고려하여 검토',
  '    ☞ 업무시간 외 접속 행위 : 00시 ~ 06시 접속기록 확인',
  '2️⃣ 행위이력',
  '    ☞ 과도한 조회 여부 : 일정시간(4H) 또는 1개월 내 처리건수가 과도하게 많은 이력 확인',
  '    ☞ 다운로드 검토 : 개인정보 다운로드 이력 확인 및 사유 검토',
  '3️⃣ 권한이력',
  '    ☞ 비인가자의 권한업무 행위 : 부여된 권한 외 과도한 권한 처리 확인',
];

const TRAILER_KEYS = { improvement: 'improvement', wikiLink: 'wikiLink', cc: 'cc' };
const TRAILER_PLACEHOLDERS = {
  improvement: '로그 품질 개선 관련 내용, 시스템별 조치 예정일',
  wikiLink: '위키 링크',
  cc: '멘션할 담당팀',
};

// 인가 IP 대역표를 실제 로그 기반으로 보완하는 동안(대략 3개월) "개선 필요사항
// 및 추가 검토 예정" 섹션 맨 끝에 매달 자동으로 붙는 고정 안내문. 대역표
// 보완이 끝나거나 [3-4]/[3-5]/[3-6]이 정식 운영되기 시작하면 이 상수를
// 지우거나 내용을 갱신할 것 - renderReportText()에서만 참조하므로 여기
// 하나만 고치면 된다.
const IP_CHECK_ROLLOUT_NOTICE = [
  '*점검 항목 고도화 예정*',
  '',
  quoteBlock([
    '향후 약 3개월간 실제 로그의 IP 사용 패턴을 분석하여',
    '인가 IP 대역표를 보완하고, 이를 기반으로 [3-4] 비인가 IP',
    '접속 행위와 [3-5] 1개 계정에서 다수 IP 접속 행위 항목도',
    '정식으로 점검할 예정입니다.',
    '',
    '또한 [3-6] 개인정보 다운로드 여부는 단순 다운로드 이력',
    '모니터링에서 나아가, 직전 3개월 평균 대비 150%를 초과하는',
    '다운로드 건에 대해 소명을 요청하는 방식으로 운영할 예정입니다.',
  ].join('\n')),
].join('\n');

// UI 렌더링용 구조화 모델. 실제 서버 발행 텍스트와 100% 같은 판정 로직을 쓰되,
// 코멘트가 필요한 자리는 [[코멘트 입력]] 마커 대신 needsComment/commentKey로 표시한다.
function buildReportModel(month, systems) {
  const report = computeReport(month, systems);
  const sections = report.sections.map(function(sec) {
    if (sec.kind === 'paragraph') {
      return {
        id: sec.id, title: sec.title, kind: 'paragraph', emoji: sec.emoji, status: sec.status,
        text: sec.paragraphLines.join('\n'),
        needsComment: sec.needsComment, commentKey: sec.commentKey, commentPlaceholder: sec.commentPlaceholder,
      };
    }
    return {
      id: sec.id, title: sec.title, kind: 'list', emoji: sec.emoji, status: sec.status,
      items: sec.items.map(function(it) {
        return {
          system: it.system,
          text: it.base + (it.trend || ''),
          needsComment: it.needsComment,
          commentKey: it.commentKey,
          commentPlaceholder: it.needsComment ? '코멘트를 입력하세요' : null,
          details: it.details || null,
          accounts: it.accounts || null,
        };
      }),
    };
  });

  return {
    month: report.month, year: report.year, mon: report.mon,
    mention: ASSIGNEE_MENTION,
    scope: report.scope,
    section2Lines: SECTION2_LINES,
    sections: sections,
    trailer: {
      keys: TRAILER_KEYS,
      placeholders: TRAILER_PLACEHOLDERS,
    },
  };
}

function commentOrPlaceholder(comments, key) {
  const v = comments && comments[key];
  return (v && String(v).trim()) ? String(v).trim() : '[[코멘트 입력]]';
}

// 계정별로 쪼개진 코멘트를 "ldap1: 코멘트1 / ldap2: 코멘트2" 형태로 한 줄에
// 다시 이어붙인다 - 최종 발행 텍스트/위키 표는 계정별 코멘트를 각각 보여주되
// 항목 한 줄이라는 기존 리포트 형식은 그대로 유지한다.
function combineAccountComments(accounts, comments) {
  return accounts.map(function(a) {
    return a.label + ': ' + commentOrPlaceholder(comments, a.commentKey);
  }).join(' / ');
}

function renderListSectionText(sec, comments) {
  const body = sec.items.length
    ? sec.items.map(function(it) {
      let line;
      if (it.needsComment) {
        const commentPart = (it.accounts && it.accounts.length)
          ? combineAccountComments(it.accounts, comments)
          : commentOrPlaceholder(comments, it.commentKey);
        line = it.base + ' / ' + commentPart + (it.trend || '');
      } else {
        line = it.base + (it.trend || '');
      }
      return it.system + ' : ' + line;
    }).join('\n')
    : '해당 없음';
  const header = '*[3-' + sec.id + '] ' + sec.title + ' ' + sec.emoji + '* : ' + sec.status;
  return header + '\n\n' + quoteBlock(body);
}

function renderParagraphSectionText(sec, comments) {
  const commentText = comments && comments[sec.commentKey] && String(comments[sec.commentKey]).trim()
    ? String(comments[sec.commentKey]).trim()
    : '[[코멘트 입력: ' + sec.commentPlaceholder + ']]';
  const inner = sec.paragraphLines.concat([commentText]).join('\n');
  const header = '*[3-' + sec.id + '] ' + sec.title + ' ' + sec.emoji + '* : ' + sec.status;
  return header + '\n\n' + quoteBlock(inner);
}

// assignees(문자열 배열, "@" 없이 아이디만)를 "@id1 @id2 ..." 형태로 합친다.
// 비어있으면 기본 담당자(ASSIGNEE_MENTION)로 폴백한다.
function formatMentionLine(assignees) {
  const ids = (assignees || [])
    .map(function(a) { return String(a || '').trim().replace(/^@+/, ''); })
    .filter(Boolean);
  if (!ids.length) return ASSIGNEE_MENTION;
  return ids.map(function(id) { return '@' + id; }).join(' ');
}

// options: { comments: {commentKey: text}, improvement, wikiLink, cc, assignees, realMention }
// assignees: "@" 없는 아이디 문자열 배열(예: ['han.b', 'louis.l']). 생략하면 기본
// 담당자(@han.b)로 폴백한다. realMention=false면 assignees와 무관하게 항상
// 테스트 placeholder로 나간다 (실제 멘션 유출 방지).
function renderReportText(month, systems, options) {
  options = options || {};
  const comments = options.comments || {};
  const report = computeReport(month, systems);
  const year = report.year, mon = report.mon;
  const mentionLine = options.realMention ? formatMentionLine(options.assignees) : '[[테스트-실제멘션아님]]';

  const section1Inner = [
    '▶︎ *점검일시* : ' + year + '년 ' + mon + '월',
    '▶︎ *점검대상 시스템* : 개인정보처리시스템',
    ' ᄂ 카카오클라우드 : ' + report.scope.kko,
    'ㄴ iLaaS : ' + report.scope.ilaas,
    '▶︎ *로그 범위* : ' + year + '년 ' + mon + '월 전체 로그 (접근이력, 행위이력, 권한이력)',
  ].join('\n');

  const sectionTexts = report.sections.map(function(sec) {
    return sec.kind === 'paragraph' ? renderParagraphSectionText(sec, comments) : renderListSectionText(sec, comments);
  });

  const improvement = (options.improvement && String(options.improvement).trim())
    || '[[코멘트 입력: ' + TRAILER_PLACEHOLDERS.improvement + ']]';
  const wikiLink = (options.wikiLink && String(options.wikiLink).trim())
    || '[[코멘트 입력: ' + TRAILER_PLACEHOLDERS.wikiLink + ']]';
  const cc = (options.cc && String(options.cc).trim())
    || '[[코멘트 입력: ' + TRAILER_PLACEHOLDERS.cc + ']]';

  return [
    '#' + year + '년' + mon + '월 #개인정보처리시스템로그점검',
    mentionLine,
    '',
    '매월 정기적으로 진행하는 개인정보처리시스템 접속기록 검토 결과를 아래와 같이 공유드립니다.',
    '',
    '*[1] 점검 개요*',
    '',
    quoteBlock(section1Inner),
    '',
    '*[2] 점검 기준*',
    '',
    quoteBlock(SECTION2_LINES.join('\n')),
    '',
    '*[3] 점검 결과*',
    '',
    sectionTexts.join('\n\n'),
    '',
    '*개선 필요사항 및 추가 검토 예정*',
    '',
    '> ' + improvement,
    '',
    IP_CHECK_ROLLOUT_NOTICE,
    '',
    '*검토 세부 내용*',
    '- ' + wikiLink,
    '',
    'cc. ' + cc,
    '',
  ].join('\n');
}

module.exports = {
  buildReportModel: buildReportModel,
  renderReportText: renderReportText,
  formatMentionLine: formatMentionLine,
  deriveAnomalyRecords: deriveAnomalyRecords,
  quoteBlock: quoteBlock,
  applyAccountAliases: applyAccountAliases,
};
