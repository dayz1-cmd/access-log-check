// analysis-core.js
// 접근 로그 분석 로직 (브라우저 UI와 CLI가 공유하는 공용 모듈).
// 브라우저: <script src="analysis-core.js">가 XLSX(CDN)와 메인 <script> 사이에서 로드되어
//           window.AnalysisCore로 노출됨.
// Node:     require('./public/analysis-core.js')로 로드, xlsx 패키지를 사용.
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('xlsx'));
  } else {
    root.AnalysisCore = factory(root.XLSX);
  }
})(typeof self !== 'undefined' ? self : this, function (XLSX) {
var SYSTEM_FORMATS = {
  'QueryPie-PI-UserLogin':   { delim:'tab', ts:1, user:4, ip:7, action:[2,3], desc:'Action At, Action Type, Result, Username' },
  'QueryPie-CSAP-UserLogin': { delim:'tab', ts:1, user:4, ip:7, action:[2,3], desc:'Action At, Action Type, Result, Username' },
  'QueryPie-PI-UserMgmt':    { delim:'comma', ts:6, user:2, ip:-1, action:[5,1], desc:'UserAccountMgmtHistory' },
  'QueryPie-PI-AdminActivity':{ delim:'comma', ts:6, user:3, ip:5, action:[1,2], desc:'AdminActivityHistory' },
  'QueryPie-PI-Activation':  { delim:'comma', ts:7, user:2, ip:-1, action:[5,6], desc:'UserAccountActivationHistory' },
  'QueryPie-CSAP-UserMgmt':  { delim:'comma', ts:6, user:2, ip:-1, action:[5], desc:'UserAccountMgmtHistory' },
  'QueryPie-CSAP-AdminActivity':{ delim:'comma', ts:6, user:3, ip:5, action:[1,2], desc:'AdminActivityHistory' },
  'QueryPie-PI-AdminRoleHistory':  { delim:'comma', ts:1, user:3, ip:-1, action:[2,5], desc:'AdminRoleHistory' },
  'QueryPie-CSAP-AdminRoleHistory':{ delim:'comma', ts:1, user:3, ip:-1, action:[2,5], desc:'AdminRoleHistory' },
  // PI에는 있지만 CSAP에는 없던 Activation 전용 키. PI와 컬럼 구조가 동일해
  // 그대로 복사했다 (resolveFormatKeyForSheet()의 QueryPie 존 매칭이 참조한다).
  'QueryPie-CSAP-Activation':  { delim:'comma', ts:7, user:2, ip:-1, action:[5,6], desc:'UserAccountActivationHistory (CSAP)' },
  'QueryPie-SpaceSeparated': { delim:'space', ts:1, user:4, ip:6, action:[2,3], desc:'공백 구분 형식' },
  'KCadmin-접속':   { delim:'tab', ts:1, user:3, ip:4, action:[2], desc:'event_dt, event_tp, event_user_id, event_user_ip' },
  'KCadmin-수행업무':{ delim:'tab', ts:1, user:8, ip:7, action:[2,3,5], desc:'수행업무 로그' },
  'KCadmin-권한':   { delim:'tab', ts:1, user:4, ip:5, action:[2,6], desc:'권한 로그' },
  // KC-Admin 실제 export는 한 파일 안에 리전 접두사로 존이 섞여 있다: "KR2_*"는
  // 퍼블릭(민간), "kr-gov-central-1_*"/"kr-gov-central-2_*"는 공공(정부). 컬럼
  // 구조는 리전과 무관하게 동일하므로, 시스템명 구분용으로만 존별 키를 둔다.
  'KCadmin-공공-접속':   { delim:'tab', ts:1, user:3, ip:4, action:[2], desc:'[실제파일] 공공(kr-gov-central-*) - 컬럼 구조는 KCadmin-접속과 동일' },
  'KCadmin-공공-수행업무':{ delim:'tab', ts:1, user:8, ip:7, action:[2,3,5], desc:'[실제파일] 공공(kr-gov-central-*) - 컬럼 구조는 KCadmin-수행업무와 동일' },
  'KCadmin-공공-권한':   { delim:'tab', ts:1, user:4, ip:5, action:[2,6], desc:'[실제파일] 공공(kr-gov-central-*) - 컬럼 구조는 KCadmin-권한과 동일' },
  'KCadmin-퍼블릭-접속':   { delim:'tab', ts:1, user:3, ip:4, action:[2], desc:'[실제파일] 퍼블릭(KR2) - 컬럼 구조는 KCadmin-접속과 동일' },
  'KCadmin-퍼블릭-수행업무':{ delim:'tab', ts:1, user:8, ip:7, action:[2,3,5], desc:'[실제파일] 퍼블릭(KR2) - 컬럼 구조는 KCadmin-수행업무와 동일' },
  'KCadmin-퍼블릭-권한':   { delim:'tab', ts:1, user:4, ip:5, action:[2,6], desc:'[실제파일] 퍼블릭(KR2) - 컬럼 구조는 KCadmin-권한과 동일' },
  'BTS-접근':  { delim:'tab', ts:6, user:1, ip:2, action:[3], desc:'ldap_id, ip, action, created_at' },
  'BTS-권한':  { delim:'tab', ts:8, user:1, ip:2, action:[3,4,5], desc:'권한 로그' },
  'BTS-행위':  { delim:'tab', ts:13, user:1, ip:2, action:[3,4,5], resultCol:9, successCol:9, downloadReasonCol:8, rowCountCol:7, desc:'행위 로그' },
  'IAM-Access':    { delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'user_ldap_id, ip, menu_id, request_uri, status_code, action' },
  'IAM-Action':    { delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'개인정보 조회 로그 (user_ldap_id, ip, menu_id, request_uri, status_code, action)' },
  'IAM-Permission':{ delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'user_ldap_id, ip, menu_id, request_uri, status_code, action' },
  // IAM 실제 export도 한 파일 안에 시트명 접두사로 존이 섞여 있다: "IAM_gov_*"는
  // 공공, "IAM_public_*"는 퍼블릭. 컬럼 구조는 KCadmin과 마찬가지로 존/로그종류와
  // 무관하게 전부 동일하다.
  'IAM-공공-Access':    { delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'[실제파일] 공공(IAM_gov_Access) - 컬럼 구조는 IAM-Access와 동일' },
  'IAM-공공-Action':    { delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'[실제파일] 공공(IAM_gov_Personal Info) - 컬럼 구조는 IAM-Action과 동일' },
  'IAM-공공-Permission':{ delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'[실제파일] 공공(IAM_gov_Permission) - 컬럼 구조는 IAM-Permission과 동일' },
  'IAM-퍼블릭-Access':    { delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'[실제파일] 퍼블릭(IAM_public_Access) - 컬럼 구조는 IAM-Access와 동일' },
  'IAM-퍼블릭-Action':    { delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'[실제파일] 퍼블릭(IAM_public_Personal Info) - 컬럼 구조는 IAM-Action과 동일' },
  'IAM-퍼블릭-Permission':{ delim:'tab', ts:11, user:2, ip:3, action:[4,5,6,7], desc:'[실제파일] 퍼블릭(IAM_public_Permission) - 컬럼 구조는 IAM-Permission과 동일' },
  '빌링-접근':  { delim:'tab', ts:6, user:1, ip:2, action:[3], desc:'접근 로그' },
  '빌링-권한':  { delim:'tab', ts:8, user:1, ip:2, action:[3,4,6], desc:'권한 로그' },
  '빌링-행위':  { delim:'tab', ts:10, user:1, ip:2, action:[3,4], desc:'행위 로그' },
  // BTS/빌링 존(공공/퍼블릭)별 실제 export 컬럼 구조 매칭용.
  // "접근로그"/"권한로그"/"행위로그" 시트명 자체는 두 시스템·두 존이 전부 동일하게
  // 쓰므로(시트명만으로 구분 불가), resolveFormatKeyForSheet()가 파일명으로 아래
  // formatKey 중 하나를 먼저 골라준다.
  'BTS-공공-접근': { delim:'tab', ts:6, user:1, ip:2, action:[3], desc:'[실제파일] id,ldap_id,ip,action,vdi,uri,created_at (공공)' },
  'BTS-공공-권한': { delim:'tab', ts:8, user:1, ip:2, action:[3,4,5], desc:'[실제파일] id,ldap_id,ip,action,target_id,role_name,vdi,parameter,created_at (공공)' },
  'BTS-공공-행위': { delim:'tab', ts:13, user:1, ip:2, action:[3,4,5], resultCol:9, successCol:9, downloadReasonCol:8, rowCountCol:7, desc:'[실제파일] id,ldap_id,ip,menu,action,target_id,privacy_info,row_count,reason_for_download,success,error_message,vdi,parameter,created_at (공공)' },
  'BTS-퍼블릭-접근': { delim:'tab', ts:6, user:1, ip:2, action:[3], desc:'[실제파일] 공공과 컬럼 구조 동일 (퍼블릭)' },
  'BTS-퍼블릭-권한': { delim:'tab', ts:8, user:1, ip:2, action:[3,4,5], desc:'[실제파일] 공공과 컬럼 구조 동일 (퍼블릭)' },
  'BTS-퍼블릭-행위': { delim:'tab', ts:13, user:1, ip:2, action:[3,4,5], resultCol:9, successCol:9, downloadReasonCol:8, rowCountCol:7, desc:'[실제파일] 공공과 컬럼 구조 동일 (퍼블릭)' },
  '빌링-공공-접근': { delim:'tab', ts:6, user:1, ip:2, action:[3], desc:'[실제파일] id,ldap_id,ip,action,is_vdi,uri,created_at,memo (공공)' },
  '빌링-공공-권한': { delim:'tab', ts:8, user:1, ip:2, action:[3,4,6], desc:'[실제파일] id,ldap_id,ip,action,target_id,is_vdi,role_name,parameter,created_at (공공)' },
  '빌링-공공-행위': { delim:'tab', ts:10, user:1, ip:2, action:[3,4], rowCountCol:6, downloadReasonCol:9, desc:'[실제파일] id,ldap_id,ip,action,target_id,is_vdi,row_count,parameter,privacy_info,reason_for_download,created_at (공공, reason_for_download 컬럼 있음)' },
  '빌링-퍼블릭-접근': { delim:'tab', ts:6, user:1, ip:2, action:[3], desc:'[실제파일] 공공과 컬럼 구조 동일 (퍼블릭)' },
  '빌링-퍼블릭-권한': { delim:'tab', ts:8, user:1, ip:2, action:[3,4,6], desc:'[실제파일] 공공과 컬럼 구조 동일 (퍼블릭)' },
  // 퍼블릭 존은 reason_for_download 컬럼 자체가 없어서 공공보다 한 칸씩 당겨져
  // 있음(created_at이 10이 아니라 9) - 공공 정의를 그대로 재사용하면 안 됨.
  '빌링-퍼블릭-행위': { delim:'tab', ts:9, user:1, ip:2, action:[3,4], rowCountCol:6, desc:'[실제파일] id,ldap_id,ip,action,target_id,is_vdi,row_count,parameter,privacy_info,created_at (퍼블릭, reason_for_download 컬럼 없음)' },
  '교육센터-액세스':{ delim:'tab', ts:0, user:1, ip:'json:3:ip', action:[3], desc:'발생 시간, 고유 아이디, 이벤트, 내용' },
  '교육센터-액션': { delim:'tab', ts:0, user:1, ip:-1, action:[2], desc:'발생 시간, 아이디, 이벤트' },
  '교육센터-수료증':{ delim:'tab', ts:0, user:1, ip:-1, action:[2], desc:'발생 시간, 고유 아이디, 이벤트, 내용' },
  // "가입탈퇴" 시트: 컬럼 구조(발생 시간, 고유 아이디, 이벤트, 내용, 유저 에이전트)가
  // "액세스"와 동일해 같은 방식으로 매핑한다 (실제 파일 확인 완료).
  '교육센터-가입탈퇴':{ delim:'tab', ts:0, user:1, ip:'json:3:ip', action:[3], desc:'[실제파일] 발생 시간, 고유 아이디, 이벤트(회원가입/탈퇴), 내용' },
  'iLaaS':    { delim:'tab', ts:0, user:1, ip:2, action:[4,5,7], desc:'활동일시, LDAP 계정, IP' },
  'iLaaS-OM': { delim:'tab', ts:1, user:2, ip:3, action:[5,6,7], resultCol:8, rowCountCol:10, desc:'Id, 행위 실행시간, 사용자 명, IP, 회사ID, 액션 유형, 실행 위치, 요청 URI' },
  '젠데스크':  { delim:'tab', ts:0, user:2, ip:3, action:[4,5,6], desc:'시간, 작업자, IP 주소, 유형' },
  '파트너센터-시스템접속':  { delim:'tab', ts:6, user:0, ip:4, action:[1,2], desc:'시스템접속 로그' },
  '파트너센터-개인정보처리':{ delim:'tab', ts:8, user:7, ip:3, action:[1,2], rowCountCol:6, desc:'개인정보처리 로그' },
  '파트너센터-권한': { delim:'tab', ts:6, user:5, ip:3, action:[1,2], desc:'target_username, action, role, ip, parameter, created_by, created_at' },
  'custom': { delim:'auto', ts:0, user:2, ip:3, action:[4], desc:'직접 설정' }
};

var CHECK_CONFIG = {
  'QueryPie-PI-UserLogin':   {lf:true, ah:true, mip:true, dl:false},
  'QueryPie-CSAP-UserLogin': {lf:true, ah:true, mip:true, dl:false},
  'QueryPie-PI-UserMgmt':    {lf:false,ah:false,mip:false,dl:false},
  'QueryPie-PI-AdminActivity':{lf:false,ah:false,mip:false,dl:true},
  'QueryPie-PI-Activation':  {lf:false,ah:false,mip:false,dl:false},
  'QueryPie-CSAP-UserMgmt':  {lf:false,ah:false,mip:false,dl:false},
  'QueryPie-CSAP-AdminActivity':{lf:false,ah:false,mip:false,dl:true},
  'QueryPie-PI-AdminRoleHistory':  {lf:false,ah:false,mip:false,dl:false},
  'QueryPie-CSAP-AdminRoleHistory':{lf:false,ah:false,mip:false,dl:false},
  'QueryPie-CSAP-Activation':  {lf:false,ah:false,mip:false,dl:false},
  'QueryPie-SpaceSeparated': {lf:true, ah:true, mip:true, dl:false},
  'KCadmin-접속':   {lf:true, ah:true, mip:true, dl:true},
  'KCadmin-수행업무':{lf:false,ah:true, mip:false,dl:true},
  'KCadmin-권한':   {lf:false,ah:false,mip:false,dl:false},
  'KCadmin-공공-접속':   {lf:true, ah:true, mip:true, dl:true},
  'KCadmin-공공-수행업무':{lf:false,ah:true, mip:false,dl:true},
  'KCadmin-공공-권한':   {lf:false,ah:false,mip:false,dl:false},
  'KCadmin-퍼블릭-접속':   {lf:true, ah:true, mip:true, dl:true},
  'KCadmin-퍼블릭-수행업무':{lf:false,ah:true, mip:false,dl:true},
  'KCadmin-퍼블릭-권한':   {lf:false,ah:false,mip:false,dl:false},
  'BTS-접근':  {lf:true, ah:true, mip:true, dl:true},
  'BTS-권한':  {lf:false,ah:false,mip:false,dl:false},
  'BTS-행위':  {lf:false,ah:true, mip:false,dl:true},
  'IAM-Access':    {lf:true, ah:true, mip:true, dl:true},
  'IAM-Action':    {lf:false,ah:true, mip:false,dl:true},
  'IAM-Permission':{lf:false,ah:false,mip:false,dl:false},
  'IAM-공공-Access':    {lf:true, ah:true, mip:true, dl:true},
  'IAM-공공-Action':    {lf:false,ah:true, mip:false,dl:true},
  'IAM-공공-Permission':{lf:false,ah:false,mip:false,dl:false},
  'IAM-퍼블릭-Access':    {lf:true, ah:true, mip:true, dl:true},
  'IAM-퍼블릭-Action':    {lf:false,ah:true, mip:false,dl:true},
  'IAM-퍼블릭-Permission':{lf:false,ah:false,mip:false,dl:false},
  '빌링-접근':  {lf:true, ah:true, mip:true, dl:true},
  '빌링-권한':  {lf:false,ah:false,mip:false,dl:false},
  '빌링-행위':  {lf:false,ah:true, mip:false,dl:true},
  'BTS-공공-접근': {lf:true, ah:true, mip:true, dl:true},
  'BTS-공공-권한': {lf:false,ah:false,mip:false,dl:false},
  'BTS-공공-행위': {lf:false,ah:true, mip:false,dl:true},
  'BTS-퍼블릭-접근': {lf:true, ah:true, mip:true, dl:true},
  'BTS-퍼블릭-권한': {lf:false,ah:false,mip:false,dl:false},
  'BTS-퍼블릭-행위': {lf:false,ah:true, mip:false,dl:true},
  '빌링-공공-접근': {lf:true, ah:true, mip:true, dl:true},
  '빌링-공공-권한': {lf:false,ah:false,mip:false,dl:false},
  '빌링-공공-행위': {lf:false,ah:true, mip:false,dl:true},
  '빌링-퍼블릭-접근': {lf:true, ah:true, mip:true, dl:true},
  '빌링-퍼블릭-권한': {lf:false,ah:false,mip:false,dl:false},
  '빌링-퍼블릭-행위': {lf:false,ah:true, mip:false,dl:true},
  '교육센터-액세스':{lf:true, ah:true, mip:true, dl:true},
  '교육센터-액션': {lf:false,ah:true, mip:false,dl:true},
  '교육센터-수료증':{lf:false,ah:true, mip:false,dl:true},
  '교육센터-가입탈퇴':{lf:false,ah:true, mip:false,dl:true},
  'iLaaS':    {lf:true, ah:true, mip:true, dl:true},
  'iLaaS-OM': {lf:true, ah:true, mip:true, dl:true},
  '젠데스크':  {lf:true, ah:true, mip:true, dl:true},
  '파트너센터-시스템접속':  {lf:true, ah:true, mip:true, dl:true},
  '파트너센터-개인정보처리':{lf:false,ah:true, mip:false,dl:true},
  '파트너센터-권한': {lf:false,ah:false,mip:false,dl:false},
  'custom': {lf:true, ah:true, mip:true, dl:true}
};

// 전 시스템 공통 인가 IP/대역 목록 - 계정 취급자가 실제로 접속하는 사내망/
// VDI/VPN/SASE 출구 IP 전부. 단일 IP는 정확히 일치할 때만, CIDR(a.b.c.d/n)
// 표기는 matchesIPRule()의 비트마스크 비교로 해당 대역 포함 여부를 판단한다.
// 10.187.x는 172.x 사내망보다 훨씬 민감한 개인정보망/GOV망 전용 VDI라서
// /16이 아니라 /24로 좁게 잡았다 (필요시 이후 조정 가능).
var AUTHORIZED_IP_RULES = [
  '127.0.0.1',        // Local
  '203.246.171.161',  // 오피스(공인)
  '172.16.0.0/16',    // 전용단말(KR2)
  '172.25.0.0/16',    // 오피스(판교)
  '172.26.0.0/16',    // 내부(카카오)
  '172.27.0.0/16',    // 오피스(제주)
  '172.28.0.0/16',    // 내부(IDC)
  '172.29.0.0/16',    // 내부(VDI)
  '172.30.0.0/16',    // 내부(VPN)
  '172.31.0.0/16',    // 내부(VPN)
  '10.187.193.0/24',  // PI-VDI 개인정보망
  '10.187.196.0/24',  // PI-VDI 개인정보망
  '10.187.197.0/24',  // GOV-VDI GOV망
  '10.187.195.0/24',  // GOV-VDI GOV망
  '211.56.96.34',     // 오피스(공인4F)
  '211.56.96.80',     // 내부(VPN)
  '211.56.96.81',     // 내부(VPN)
  '211.56.96.82',     // 내부(VPN)
  '211.56.96.83',     // 내부(VPN)
  '211.56.96.84',     // 내부(VPN)
  // 오피스(121.65.239.x)/SASE(165.85.218.x)는 실사용 로그를 보면 단일 IP가
  // 아니라 대역(pool)이었다 (예: .17/.45가 함께 오피스, .54/.89가 함께
  // SASE로 관측됨) - /24로 등록해 그 대역 전체를 인가로 인정한다.
  '121.65.239.0/24',  // 오피스
  '165.85.218.0/24',  // SASE
  // ip-usage-analysis.js로 2026-08 실제 로그를 분석해 발견 - 정상 사용으로
  // 판단되어 추가 (사람이 직접 검토 후 확정).
  '144.125.248.0/24', // KCadmin(퍼블릭) 계정들이 실사용 중인 대역
  '110.249.0.0/16',   // 실사용 중인 대역 (출장/모바일망 등으로 판단)
  // 2026-09 IAM 어드민(퍼블릭/general) 로그에서 접속자 전원이 이 IP 하나로만
  // 찍힘 - 개인 PC IP가 아니라 퍼블릭 IAM 어드민 앞단 서버(게이트웨이) IP로
  // 판단되어 인가로 등록 (tyger.k/yenny.0k 비인가 IP 오탐 원인).
  '10.82.67.99',      // IAM 어드민(퍼블릭) 게이트웨이
];

// CHECK_CONFIG에 등록된 모든 포맷 키에 위 공통 목록을 그대로 적용한다 -
// 시스템별로 하드코딩해 나열하는 대신 CHECK_CONFIG 키 목록에서 파생시켜서,
// 새 포맷이 추가돼도 자동으로 같은 공통 목록이 적용되고 빠지는 시스템이
// 없게 한다. hasConfiguredAuthorizedIPs()가 hasOwnProperty + 배열 길이로
// "설정됨"을 판단하므로, 여기 등록되는 순간부터 해당 포맷은 더 이상
// "인가 IP 목록 미설정 - 점검 제외"로 스킵되지 않고 실제 비인가 IP 검토가
// 동작한다.
//
// 예외: iLaaS-OM은 계정 자체가 UUID(직원 LDAP이 아님)이고 접속 IP도 국내
// 다양한 통신사 대역에 걸쳐 있어 고객(엔드유저) 접속 로그로 판단된다 -
// "사내망 인가 IP" 개념 자체가 이 시스템에는 맞지 않으므로, 공통 목록을
// 적용하지 않고 기존처럼 빈 배열로 "미설정 - 점검 제외" 상태를 유지한다.
var SYSTEM_AUTHORIZED_IPS = Object.keys(CHECK_CONFIG).reduce(function(map, key) {
  map[key] = key === 'iLaaS-OM' ? [] : AUTHORIZED_IP_RULES;
  return map;
}, {});

// 계정명 앞에 붙는 시스템별 접두어 - 제거하면 실제 LDAP 계정이 남는다.
// 다른 시스템에서 새 접두어가 발견되면 여기에 계속 추가한다.
var LDAP_PREFIXES = ['csap.', 'csp.', 'pi.'];

// STATUS-BOARD-CLIENT처럼 대문자+숫자를 하이픈으로만 이어붙인 계정은 사람이
// 아니라 시스템/서비스 계정이다 - 소명 요청 대상에서 제외해야 한다.
// "STATUS-BOARD-CLIENT(chloe.cy@kakaocorp.com)"처럼 뒤에 담당자 이메일이 괄호로
// 붙는 변형도 실제 로그에서 발견돼, 판정 전에 뒤쪽 "(...)"는 떼고 검사한다.
function systemAccountBase(user) {
  return String(user || '').replace(/\s*\([^)]*\)\s*$/, '');
}

function isSystemAccount(user) {
  return /^[A-Z0-9]+(-[A-Z0-9]+)+$/.test(systemAccountBase(user));
}

// iLaaS(OM) 등 일부 시스템은 사람 계정 대신 UUID(8-4-4-4-12 16진수)를 그대로
// 로그에 남긴다 - 실명 매핑이 없어 소명 요청 대상으로 삼을 수 없다.
function isUuidAccount(user) {
  return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(String(user || ''));
}

// 계정명 앞의 등록된 접두어를 제거하고, 이메일 형식(로컬파트@도메인)이면
// '@' 앞의 로컬파트만 남겨서 LDAP 계정을 뽑아낸다. 같은 사람이 로그에 따라
// "davis.jeon"으로도 "davis.jeon@kakaocorp.com"으로도 남는 경우가 있어,
// 이 정규화를 거치면 둘 다 "davis.jeon"으로 합쳐진다. 접두어/이메일이 모두
// 없으면 계정명 그대로가 LDAP이다 (예: "elvin" -> "elvin").
function extractLdap(user) {
  var u = String(user || '');
  for (var i = 0; i < LDAP_PREFIXES.length; i++) {
    if (u.indexOf(LDAP_PREFIXES[i]) === 0) { u = u.slice(LDAP_PREFIXES[i].length); break; }
  }
  var atIdx = u.indexOf('@');
  if (atIdx >= 0) u = u.slice(0, atIdx);
  return u;
}

var OVER_QUERY_THRESHOLD = 1000;
var UNAUTHORIZED_ACCESS_KEYWORDS = ['unauthorized', 'forbidden', 'access denied', '접근 거부', '접근불가', '비인가 접속', '비인가접속'];
var UNAUTHORIZED_PERMISSION_KEYWORDS = ['privilege escalation', '권한 상승', '권한 오남용', '비인가 권한', '권한 없음', 'permission denied', 'role escalation'];

var LOG_ISSUES = {
  '빌링-접근':'로그인 실패 이력 로깅되지 않음',
  'BTS-접근':'로그인 실패 이력 로깅되지 않음',
  '젠데스크':'로그인 실패 이력 로깅되지 않음',
  '교육센터-액세스':'로그인 실패 이력 로깅되지 않음',
  'iLaaS':'로그인 이력 로깅되지 않음',
  'BTS-공공-접근':'로그인 실패 이력 로깅되지 않음',
  'BTS-퍼블릭-접근':'로그인 실패 이력 로깅되지 않음',
  '빌링-공공-접근':'로그인 실패 이력 로깅되지 않음',
  '빌링-퍼블릭-접근':'로그인 실패 이력 로깅되지 않음'
};

// ─── IP 분류 ───────────────────────────────────────────────
function ipToNumber(ip) {
  return ip.split('.').reduce(function(total, part) { return total * 256 + parseInt(part, 10); }, 0);
}

function matchesIPRule(ip, rule) {
  if (rule.indexOf('/') < 0) return rule.slice(-1) === '.' ? ip.indexOf(rule) === 0 : ip === rule;
  var parts = rule.split('/');
  var prefixLength = parseInt(parts[1], 10);
  if (!isValidIPAddress(parts[0]) || isNaN(prefixLength) || prefixLength < 0 || prefixLength > 32) return false;
  var mask = prefixLength === 0 ? 0 : (0xffffffff << (32 - prefixLength)) >>> 0;
  return (ipToNumber(ip) >>> 0 & mask) === (ipToNumber(parts[0]) >>> 0 & mask);
}

// 버그: 기존에는 SYSTEM_AUTHORIZED_IPS에 아예 등록조차 안 된 시스템(BTS/KCadmin/
// 빌링/젠데스크 등 대부분)을 "설정됨"으로 잘못 판단해, 아래 classifyIP()의 하드코딩된
// 소수의 사내망/재택 IP 목록과만 대조하는 바람에 실제 정상 트래픽(원격/VPN 등)이
// 전부 "비인가 IP"로 잘못 집계됐다(예: BTS(공공) 12706건). "설정됨"은 반드시
// SYSTEM_AUTHORIZED_IPS에 실제로 등록되어 있고(hasOwnProperty) 그 목록이 비어있지
// 않을 때만이어야 한다 - 등록 자체가 안 된 시스템은 iLaaS-OM(빈 배열로 등록)과
// 마찬가지로 "미설정 - 점검 제외"로 처리되는 게 맞는다.
function hasConfiguredAuthorizedIPs(systemKey) {
  return Object.prototype.hasOwnProperty.call(SYSTEM_AUTHORIZED_IPS, systemKey) && SYSTEM_AUTHORIZED_IPS[systemKey].length > 0;
}

function classifyIP(ip, systemKey) {
  if (!ip || ip === 'skip') return {ok:true, type:'Unknown'};
  if (Object.prototype.hasOwnProperty.call(SYSTEM_AUTHORIZED_IPS, systemKey)) {
    var authorizedIPs = SYSTEM_AUTHORIZED_IPS[systemKey];
    if (!authorizedIPs.length) return {ok:true, type:'미설정'};
    for (var i = 0; i < authorizedIPs.length; i++) {
      if (matchesIPRule(ip, authorizedIPs[i])) return {ok:true, type:'인가 IP'};
    }
    return {ok:false, type:'비인가 IP'};
  }
  if (ip === '127.0.0.1') return {ok:true, type:'Local'};
  if (ip === '203.246.171.161') return {ok:true, type:'오피스(공인)'};
  if (ip === '211.56.96.34') return {ok:true, type:'오피스(공인4F)'};
  if (ip === '165.85.218.54') return {ok:true, type:'SASE'};
  var last = ip.split('.')[3];
  if (ip.indexOf('211.56.96.8') === 0 && ['80','81','82','83','84'].indexOf(last) >= 0) return {ok:true, type:'VPN'};
  if (ip.indexOf('172.16.') === 0) return {ok:true, type:'전용단말(KR2)'};
  if (ip.indexOf('172.25.') === 0) return {ok:true, type:'오피스(판교)'};
  if (ip.indexOf('172.26.') === 0) return {ok:true, type:'내부(카카오)'};
  if (ip.indexOf('172.27.') === 0) return {ok:true, type:'오피스(제주)'};
  if (ip.indexOf('172.28.') === 0) return {ok:true, type:'내부(IDC)'};
  if (ip.indexOf('172.29.') === 0) return {ok:true, type:'내부(VDI)'};
  if (ip.indexOf('172.30.') === 0) return {ok:true, type:'내부(VPN)'};
  if (ip.indexOf('172.31.') === 0) return {ok:true, type:'내부(VPN)'};
  if (ip.indexOf('10.187.193.') === 0 || ip.indexOf('10.187.196.') === 0) return {ok:true, type:'PI-VDI(개인정보망)'};
  if (ip.indexOf('10.187.197.') === 0 || ip.indexOf('10.187.195.') === 0) return {ok:true, type:'GOV-VDI'};
  var homeIPs = {
    '211.56.96.61':'재택(pop.con)','122.40.170.196':'재택(kai.kkim)',
    '121.137.129.208':'재택(nandy.namu)','221.146.27.40':'재택(kali.s)',
    '58.234.91.247':'재택(irene.yond)','59.11.254.82':'재택(evan.ejin)',
    '218.50.176.224':'재택(Kevin.kep)','211.216.113.136':'재택(jake.pk)',
    '211.56.96.84':'재택(nandy.namu)','211.209.195.243':'재택(ty.jk)'
  };
  if (homeIPs[ip]) return {ok:true, type:homeIPs[ip]};
  return {ok:false, type:'비인가 IP'};
}

function containsKeyword(value, keywords) {
  var normalized = (value || '').toLowerCase();
  for (var i = 0; i < keywords.length; i++) {
    if (normalized.indexOf(keywords[i].toLowerCase()) >= 0) return true;
  }
  return false;
}

function isValidZendeskUser(user) {
  if (!user || /^[A-Z]+$/.test(user) || /^\d/.test(user) || /[가-힣]/.test(user)) return false;
  return true;
}

function parseRowCount(parts, fmt) {
  if (fmt.rowCountCol === undefined) return null;
  var value = parseInt(getCol(parts, fmt.rowCountCol), 10);
  return isNaN(value) ? null : value;
}

function hasColumnMapping(mapping) {
  return mapping !== undefined && mapping !== null && mapping !== -1 &&
    (typeof mapping === 'number' || (typeof mapping === 'string' && mapping.indexOf('json:') === 0));
}

function isValidIPAddress(ip) {
  if (!ip || !/^\d{1,3}(\.\d{1,3}){3}$/.test(ip)) return false;
  return ip.split('.').every(function(part) { return parseInt(part, 10) >= 0 && parseInt(part, 10) <= 255; });
}

// ─── 타임스탬프 파싱 ───────────────────────────────────────
function tsToMs(ts) {
  if (!ts) return null;
  ts = ts.trim();
  // 엑셀 시리얼 숫자
  var serial = parseFloat(ts);
  if (!isNaN(serial) && /^\d+\.?\d*$/.test(ts) && serial > 40000 && serial < 60000) {
    return new Date(1899, 11, 30).getTime() + serial * 86400000;
  }
  // "2026.07.31 18:12:37" (iLaaS-OM 형식)
  var mOM = ts.match(/(\d{4})\.(\d{2})\.(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (mOM) {
    return new Date(parseInt(mOM[1]), parseInt(mOM[2])-1, parseInt(mOM[3]), parseInt(mOM[4]), parseInt(mOM[5]), parseInt(mOM[6])).getTime();
  }
  // "2026. 6. 30. 016:13" or "2026. 6. 30. 16:13" (라스 형식)
  var m2 = ts.match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.\s*0?(\d{1,2}):(\d{2})/);
  if (m2) {
    return new Date(parseInt(m2[1]), parseInt(m2[2])-1, parseInt(m2[3]), parseInt(m2[4]), parseInt(m2[5]), 0).getTime();
  }
  // "2026. 3. 30 오후 1:22:37" or "2026. 3. 30. 오후 1:22:37"
  var m = ts.match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.?\s+(오전|오후)\s+(\d{1,2}):(\d{2}):(\d{2})/);
  if (m) {
    var h = parseInt(m[5]);
    if (m[4] === '오후' && h !== 12) h += 12;
    else if (m[4] === '오전' && h === 12) h = 0;
    return new Date(parseInt(m[1]), parseInt(m[2])-1, parseInt(m[3]), h, parseInt(m[6]), parseInt(m[7])).getTime();
  }
  // ISO UTC
  var iso = ts.match(/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})Z?$/);
  if (iso) {
    return Date.UTC(parseInt(iso[1]), parseInt(iso[2])-1, parseInt(iso[3]), parseInt(iso[4]), parseInt(iso[5]), parseInt(iso[6])) + 9*3600000;
  }
  // 일반 날짜
  var d = new Date(ts);
  return isNaN(d.getTime()) ? null : d.getTime();
}

function tsToKSTHour(ts) {
  var ms = tsToMs(ts);
  if (ms === null) return null;
  return new Date(ms).getHours();
}

function tsDisplay(ts) {
  if (!ts) return '';
  ts = ts.trim();
  // 한국어 오전/오후 형식 → 24시간으로 변환
  if (/오전|오후/.test(ts)) {
    var km = ts.match(/(\d{4}[\-.]\s*\d{1,2}[\-.]\s*\d{1,2}|\d{1,2}\s*\.\s*\d{1,2}).*?(오전|오후)\s*(\d{1,2}):(\d{2}):(\d{2})/);
    var ms2 = tsToMs(ts);
    if (ms2 !== null) {
      var d2 = new Date(ms2);
      return d2.getFullYear()+'-'+pad(d2.getMonth()+1)+'-'+pad(d2.getDate())+' '+pad(d2.getHours())+':'+pad(d2.getMinutes())+':'+pad(d2.getSeconds());
    }
    return ts;
  }
  var ms = tsToMs(ts);
  if (ms === null) return ts;
  var d = new Date(ms);
  return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+' '+pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());
}

function pad(n) { return String(n).padStart(2,'0'); }

// ─── 라인 파싱 ─────────────────────────────────────────────
function splitLine(line, fmt) {
  if (fmt.delim === 'comma') {
    var res = [], cur = '', inQ = false;
    for (var i = 0; i < line.length; i++) {
      var c = line[i];
      if (c === '"') inQ = !inQ;
      else if (c === ',' && !inQ) { res.push(cur.trim()); cur = ''; }
      else cur += c;
    }
    res.push(cur.trim());
    return res.map(function(v){ return (v[0]==='"'&&v[v.length-1]==='"') ? v.slice(1,-1) : v; });
  }
  if (fmt.delim === 'tab') {
    if (line.indexOf('\t') >= 0) return line.split('\t');
    // 탭이 없으면 붙여넣기 시 공백 치환된 것으로 간주 — 2개 이상 공백으로 분리
    return line.split(/  +/);
  }
  if (fmt.delim === 'space') return line.trim().split(/\s+/);
  // auto
  if (line.indexOf('\t') >= 0) return line.split('\t');
  if (line.indexOf(',') >= 0) return line.split(',');
  return line.trim().split(/\s+/);
}

function getCol(parts, idx) {
  if (idx < 0 || idx === -1) return '';
  if (typeof idx === 'string' && idx.indexOf('json:') === 0) {
    var sp = idx.split(':');
    var jsonStr = parts[parseInt(sp[1])] || '';
    var km = jsonStr.match(new RegExp('"'+sp[2]+'":"([^"]+)"'));
    return km ? km[1] : '';
  }
  return (parts[idx] || '').trim();
}

function getAction(parts, indices) {
  return indices.map(function(i){ return getCol(parts, i); }).filter(Boolean).join(' ');
}

// ─── 메인 분석 ─────────────────────────────────────────────
function analyzeLog(logText, systemName, fmtKey) {
  var fmt = SYSTEM_FORMATS[fmtKey] || SYSTEM_FORMATS['custom'];
  var cfg = CHECK_CONFIG[fmtKey] || CHECK_CONFIG['custom'];

  var lines = logText.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
  var rows = [];
  for (var i = 0; i < lines.length; i++) {
    var l = lines[i].replace(/[\r\n]+$/,''); // 줄끝 개행만 제거, 앞 탭은 유지
    if (l.trim()) rows.push(l);
  }

  // 헤더 스킵
  // 보통 1줄짜리 헤더지만, KC-Admin 실제 export처럼 "1. 시스템 접속과 관련한
  // 로그(...)" 같은 제목 줄이 실제 컬럼 헤더 위에 한 줄 더 있는 경우가 있다.
  // 첫 줄이 헤더 키워드와 매칭되지 않으면 둘째 줄도 확인해서, 매칭되면
  // 제목+헤더 두 줄을 함께 건너뛴다.
  function looksLikeHeader(line) {
    var fl = line.toLowerCase();
    return fl.indexOf('event_dt') >= 0 || fl.indexOf('username') >= 0 ||
        fl.indexOf('ldap_id') >= 0 || fl.indexOf('action at') >= 0 ||
        fl.indexOf('action_at') >= 0 || fl.indexOf('발생') >= 0 ||
        fl.indexOf('시간') >= 0 || fl.indexOf('no.') >= 0 ||
        fl.indexOf('user_id') >= 0 || fl.indexOf('created_at') >= 0 ||
        fl.indexOf('target_username') >= 0;
  }
  var start = 0;
  if (rows.length > 0) {
    if (looksLikeHeader(rows[0])) {
      start = 1;
    } else if (rows.length > 1 && looksLikeHeader(rows[1])) {
      start = 2;
    }
  }

  // IAM: 2026-09 export부터 menu_id/status_code 컬럼이 빠져 12칸 -> 10칸으로
  // 바뀌었다(이전엔 created_at이 12번째 칸, 이후엔 10번째 칸). 고정 칸 번호로
  // 읽으면 접속일시가 통째로 비고(=[3-8] "접속일시 누락" 오탐) 날짜도 안 나온다.
  // 그래서 IAM은 헤더 이름으로 칸 위치를 찾고, 못 찾을 때만 기존 고정값을 쓴다.
  var iamTargetCol = 8, iamParamCol = 10;
  if (fmtKey.indexOf('IAM') === 0 && start > 0) {
    var hdr = splitLine(rows[start - 1], fmt).map(function(h) { return h.trim().toLowerCase(); });
    var findCol = function(re) {
      for (var hi = 0; hi < hdr.length; hi++) { if (re.test(hdr[hi])) return hi; }
      return -1;
    };
    var cTs = findCol(/^created_at/), cUser = findCol(/^user_ldap_id$/), cIp = findCol(/^ip$/);
    if (cTs >= 0 && cUser >= 0 && cIp >= 0) {
      var actCols = [findCol(/^menu_id$/), findCol(/^request_uri$/), findCol(/^status_code$/), findCol(/^action$/)]
        .filter(function(c) { return c >= 0; });
      fmt = Object.assign({}, fmt, {ts: cTs, user: cUser, ip: cIp, action: actCols});
      var cTarget = findCol(/^target_id$/), cParam = findCol(/^parameter$/);
      if (cTarget >= 0) iamTargetCol = cTarget;
      if (cParam >= 0) iamParamCol = cParam;
    }
  }

  var loginFailures = [];
  var afterHours = [];
  var multiIPs = {};
  var downloads = [];
  var unauthIPs = [];
  var unauthorizedAccess = [];
  var unauthorizedPermissions = [];
  var overQueries = [];
  var parsedTimestamps = [];
  var missingFieldCounts = {account: 0, timestamp: 0, ip: 0, action: 0};
  var missingFieldRows = [];

  for (var i = start; i < rows.length; i++) {
    var line = rows[i];
    var parts = splitLine(line, fmt);

    var ts = getCol(parts, fmt.ts);
    var user = getCol(parts, fmt.user);
    var ip = getCol(parts, fmt.ip);
    var act = getAction(parts, fmt.action);

    var parsedTimestamp = tsToMs(ts);
    if (parsedTimestamp !== null) parsedTimestamps.push({ms: parsedTimestamp, ts: ts});

    // 젠데스크 username 정리
    if (fmtKey === '젠데스크' && user) {
      var um = user.match(/^([^\s(]+)/);
      if (um) user = um[1];
    }

    if (fmtKey === '젠데스크' && !isValidZendeskUser(user)) continue;

    if (!ts && !user && !act) continue;

    var missingFields = [];
    if (!user) { missingFieldCounts.account++; missingFields.push('계정'); }
    if (!ts || parsedTimestamp === null) { missingFieldCounts.timestamp++; missingFields.push('접속일시'); }
    if (!isValidIPAddress(ip)) { missingFieldCounts.ip++; missingFields.push('IP'); }
    if (!act) { missingFieldCounts.action++; missingFields.push('수행업무'); }
    if (missingFields.length) missingFieldRows.push({row: i + 1, fields: missingFields});

    // ── 로그인 연속 실패 ──
    if (cfg.lf) {
      var al = act.toLowerCase();
      if (al.indexOf('login') >= 0 && al.indexOf('fail') >= 0) {
        var u = user || '';
        // IAM: 실패 시 user_ldap_id가 비어있고 target_id(8) 또는 parameter(10)에 ldapId 있음
        if (fmtKey.indexOf('IAM') === 0) {
          var tid = getCol(parts, iamTargetCol);
          if (tid) {
            u = tid;
          } else {
            var param = getCol(parts, iamParamCol);
            var lm = param.match(/ldapId['"']?\s*:\s*['"']([^'"']+)/i);
            if (lm) u = lm[1];
          }
        }
        if (!u) u = 'unknown';
        loginFailures.push({ts: ts, user: u, ip: ip, act: act});
      }
    }

    // ── 업무시간 외 접근 ──
    if (cfg.ah) {
      var h = tsToKSTHour(ts);
      if (h !== null && h >= 0 && h < 6) {
        afterHours.push({ts: ts, user: user, ip: ip, act: act});
      }
    }

    // ── 다수 IP ──
    if (cfg.mip && hasConfiguredAuthorizedIPs(fmtKey) && user && ip && isValidIPAddress(ip) && classifyIP(ip, fmtKey).ok) {
      if (!multiIPs[user]) multiIPs[user] = {};
      multiIPs[user][ip] = true;
    }

    // ── 다운로드 ──
    if (cfg.dl) {
      var al2 = act.toLowerCase();
      if (al2.indexOf('download') >= 0 || al2.indexOf('export') >= 0 ||
          al2.indexOf('csv') >= 0 || al2.indexOf('excel') >= 0 ||
          al2.indexOf('엑셀') >= 0 || al2.indexOf('다운로드') >= 0 ||
          al2.indexOf('내보내기') >= 0) {
        var successOk = true;
        if (fmt.successCol !== undefined) {
          var sv = getCol(parts, fmt.successCol).toLowerCase();
          successOk = (sv === 'true' || sv === '1' || sv === 'success');
        }
        if (successOk) {
          var rc = fmt.rowCountCol !== undefined ? parseRowCount(parts, fmt) || 1 : 1;
          downloads.push({ts: ts, user: user, ip: ip, act: act, rowCount: rc});
        }
      }
    }

    if (containsKeyword(act, UNAUTHORIZED_ACCESS_KEYWORDS)) {
      unauthorizedAccess.push({ts: ts, user: user, act: act});
    }
    if (containsKeyword(act, UNAUTHORIZED_PERMISSION_KEYWORDS)) {
      unauthorizedPermissions.push({ts: ts, user: user, act: act});
    }

    var rowCount = parseRowCount(parts, fmt);
    if (rowCount !== null && rowCount >= OVER_QUERY_THRESHOLD) {
      overQueries.push({ts: ts, user: user, rowCount: rowCount, act: act});
    }

    // ── 비인가 IP ──
    if (hasConfiguredAuthorizedIPs(fmtKey) && isValidIPAddress(ip)) {
      var cls = classifyIP(ip, fmtKey);
      if (!cls.ok) unauthIPs.push({ts: ts, user: user, ip: ip, act: act});
    }
  }

  // ── 로그인 실패 그룹핑 ──
  var userFails = {};
  for (var i = 0; i < loginFailures.length; i++) {
    var f = loginFailures[i];
    if (!userFails[f.user]) userFails[f.user] = [];
    userFails[f.user].push(f);
  }
  var failGroups = [];
  var userKeys = Object.keys(userFails);
  for (var k = 0; k < userKeys.length; k++) {
    var list = userFails[userKeys[k]];
    list.sort(function(a,b){ return (tsToMs(a.ts)||0)-(tsToMs(b.ts)||0); });
    var grp = [list[0]];
    for (var j = 1; j < list.length; j++) {
      var t1 = tsToMs(grp[grp.length-1].ts);
      var t2 = tsToMs(list[j].ts);
      if (t1 !== null && t2 !== null && (t2-t1)/60000 <= 20) {
        grp.push(list[j]);
      } else {
        if (grp.length >= 5) failGroups.push(grp.slice());
        grp = [list[j]];
      }
    }
    if (grp.length >= 5) failGroups.push(grp.slice());
  }

  // ── 업무시간 외 그룹핑 (날짜별) ──
  var ahByUser = {};
  for (var i = 0; i < afterHours.length; i++) {
    var a = afterHours[i];
    if (!ahByUser[a.user]) ahByUser[a.user] = {};
    // 날짜 키 추출
    var ms = tsToMs(a.ts);
    var dateKey = ms ? new Date(ms).toLocaleDateString('ko-KR', {year:'numeric',month:'2-digit',day:'2-digit'}).replace(/\. /g,'-').replace(/\.$/,'') : a.ts.substring(0,10);
    if (!ahByUser[a.user][dateKey]) ahByUser[a.user][dateKey] = [];
    ahByUser[a.user][dateKey].push(a);
  }
  var ahSummary = [];
  var ahKeys = Object.keys(ahByUser);
  for (var k = 0; k < ahKeys.length; k++) {
    var user = ahKeys[k];
    var byDate = ahByUser[user];
    var dateKeys = Object.keys(byDate).sort();
    var totalCount = 0;
    var dateGroups = [];
    for (var d = 0; d < dateKeys.length; d++) {
      var items = byDate[dateKeys[d]];
      var tss = items.map(function(x){ return x.ts; }).sort();
      totalCount += items.length;
      dateGroups.push({date: dateKeys[d], count: items.length, first: tss[0], last: tss[tss.length-1]});
    }
    ahSummary.push({user: user, count: totalCount, dateGroups: dateGroups, first: dateGroups[0].first, last: dateGroups[dateGroups.length-1].last});
  }

  // ── 다수 IP 그룹핑 ──
  var multiSummary = [];
  var muKeys = Object.keys(multiIPs);
  for (var k = 0; k < muKeys.length; k++) {
    var ips = Object.keys(multiIPs[muKeys[k]]);
    if (ips.length >= 2) multiSummary.push({user: muKeys[k], ips: ips.sort()});
  }

  // ── 다운로드 그룹핑 ──
  var dlByUser = {};
  for (var i = 0; i < downloads.length; i++) {
    var d = downloads[i];
    if (!dlByUser[d.user]) dlByUser[d.user] = [];
    dlByUser[d.user].push(d);
  }
  var dlSummary = [];
  var dlKeys = Object.keys(dlByUser);
  for (var k = 0; k < dlKeys.length; k++) {
    var items = dlByUser[dlKeys[k]];
    var tss = items.map(function(x){ return x.ts; }).sort();
    var totalRows = items.reduce(function(acc, x){ return acc + (x.rowCount || 1); }, 0);
    dlSummary.push({user: dlKeys[k], count: items.length, rowCount: totalRows, first: tss[0], last: tss[tss.length-1]});
  }

  // ── 비인가 IP 그룹핑 ──
  var uaByKey = {};
  for (var i = 0; i < unauthIPs.length; i++) {
    var u = unauthIPs[i];
    var key = u.user + '|' + u.ip;
    if (!uaByKey[key]) uaByKey[key] = {user:u.user, ip:u.ip, items:[]};
    uaByKey[key].items.push(u);
  }
  var uaSummary = [];
  var uaKeys = Object.keys(uaByKey);
  for (var k = 0; k < uaKeys.length; k++) {
    var g = uaByKey[uaKeys[k]];
    var tss = g.items.map(function(x){ return x.ts; }).sort();
    uaSummary.push({user:g.user, ip:g.ip, count:g.items.length, first:tss[0], last:tss[tss.length-1]});
  }

  var autoLogIssues = [];
  var logSchema = [
    {label:'계정', present:hasColumnMapping(fmt.user)},
    {label:'접속일시', present:hasColumnMapping(fmt.ts)},
    {label:'IP', present:hasColumnMapping(fmt.ip)},
    {label:'수행업무', present:Array.isArray(fmt.action) && fmt.action.length > 0},
    {label:'로그인 성공/실패 여부', present:hasColumnMapping(fmt.resultCol)},
    {label:'다운로드 사유', present:hasColumnMapping(fmt.downloadReasonCol)}
  ];
  logSchema.forEach(function(item) {
    if (!item.present) item.message = item.label === '다운로드 사유' ?
      '다운로드 사유 누락 - 로그 양식 자체에 해당 컬럼 없음' :
      item.label + ' 누락 - 로그 양식 자체에 해당 컬럼 없음';
  });
  var missingFieldLabels = {account:'계정', timestamp:'접속일시', ip:'IP', action:'수행업무'};
  if (missingFieldRows.length) {
    Object.keys(missingFieldCounts).forEach(function(key) {
      autoLogIssues.push(missingFieldLabels[key] + ' 누락 ' + missingFieldCounts[key] + '건');
    });
    missingFieldRows.forEach(function(item) {
      autoLogIssues.push('행 ' + item.row + ': ' + item.fields.join(', ') + ' 누락');
    });
  }
  parsedTimestamps.sort(function(a, b){ return a.ms - b.ms; });
  var logIssueParts = [];
  if (LOG_ISSUES[fmtKey]) logIssueParts.push(LOG_ISSUES[fmtKey]);
  if (missingFieldRows.length) logIssueParts.push('필수항목 누락 있음');

  var retentionSummary = [];
  if (parsedTimestamps.length > 1) {
    var firstTimestamp = parsedTimestamps[0];
    var lastTimestamp = parsedTimestamps[parsedTimestamps.length - 1];
    var retentionDays = Math.floor((lastTimestamp.ms - firstTimestamp.ms) / (24 * 60 * 60 * 1000));
    if (retentionDays >= 365) {
      retentionSummary.push({first:firstTimestamp.ts, last:lastTimestamp.ts, days:retentionDays});
    }
  }

  return {
    system: systemName || fmtKey,
    failGroups: failGroups,
    ahSummary: ahSummary,
    multiSummary: multiSummary,
    dlSummary: dlSummary,
    uaSummary: uaSummary,
    unauthorizedAccess: unauthorizedAccess,
    unauthorizedPermissions: unauthorizedPermissions,
    overQueries: overQueries,
    ipCheckNote: hasConfiguredAuthorizedIPs(fmtKey) ? '' : '인가 IP 목록 미설정 - 점검 제외',
    logSchema: logSchema,
    logMissing: autoLogIssues,
    retentionSummary: retentionSummary,
    logIssue: logIssueParts.join(', ')
  };
}
function camelWords(str) {
  return str.replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .split(/[^a-zA-Z0-9가-힣]+/).filter(Boolean).map(function(w) { return w.toLowerCase(); });
}

function getFormatKeyBySheetName(sheetName) {
  // 정확한 매칭
  if (SYSTEM_FORMATS[sheetName]) return sheetName;

  // 키에 시트명이 포함되면 매칭
  var keys = Object.keys(SYSTEM_FORMATS);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf(sheetName) >= 0) return keys[i];
  }

  // 시트명에 키가 포함되면 매칭 (예: "UserLoginHistory" → "QueryPie-PI-UserLogin")
  var sheetWords = camelWords(sheetName.replace(/History$/, ''));
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    // 불필요한 키들 제외
    if (key.indexOf('custom') >= 0 || key.indexOf('공백') >= 0) continue;
    // "UserLogin" 같은 부분과 매칭
    var shortKey = key.replace(/QueryPie-\w+-/g, '').replace(/iLaaS.*-/g, '').replace(/-\w+/g, '');
    if (sheetName.indexOf(shortKey) >= 0 || shortKey.indexOf(sheetName.replace('History', '')) >= 0) {
      return key;
    }
    // 단어 단위 매칭 (예: "UserAccountMgmtHistory" → shortKey "UserMgmt"처럼
    // 시트명 안에 중간 단어가 끼어 있어 단순 substring으로는 못 잡는 경우)
    var shortWords = camelWords(shortKey);
    if (shortWords.length && shortWords.every(function(w) { return sheetWords.indexOf(w) >= 0; })) {
      return key;
    }
  }

  return null;
}

// QueryPie 시트명 -> formatKey 접미사. resolveFormatKeyForSheet()의 QueryPie
// 존(PI/CSAP) 매칭에서 사용한다.
var QUERYPIE_SHEET_SUFFIX = {
  'UserLoginHistory': 'UserLogin',
  'UserAccountMgmtHistory': 'UserMgmt',
  'AdminActivityHistory': 'AdminActivity',
  'UserAccountActivationHistory': 'Activation',
  'AdminRoleHistory': 'AdminRoleHistory'
};

// getFormatKeyBySheetName()의 일반 규칙만으로는 구분되지 않는 실제 파일들을 위한
// 보조 매칭. 파일명이 필요한(sheet name만으론 판별 불가한) 경우에만 쓰인다:
//  - BTS/빌링: 두 시스템이 "접근로그"/"권한로그"/"행위로그"라는 같은 시트명을
//    쓰고, 같은 시스템 안에서도 존(공공/퍼블릭)에 따라 시트명이 완전히 똑같다.
//  - LaaS OM/LaaS 내부 어드민: 시트명이 "Data"/"Sheet1"처럼 정보가 전혀 없다.
//  - 젠데스크: 한 파일에 "26.07"처럼 "YY.MM" 형식의 시트가 여러 달치 같이 있다.
// KC-Admin(리전 접두사 + 접미사)과 파트너센터(시트명에 공백 포함)는 파일명 힌트
// 없이 시트명 자체(접미사/공백 정규화)만으로 구분 가능해 여기서 함께 처리한다.
// 매칭되는 규칙이 없으면 null을 반환해 기존 getFormatKeyBySheetName()에 맡긴다.
function resolveFormatKeyForSheet(sheetName, fileName) {
  var fn = fileName || '';
  var sn = sheetName || '';

  // BTS / 빌링: 파일명으로 시스템+존을 먼저 정하고, 시트명으로 로그 종류를 정한다.
  var zonedBase = null;
  if (fn.indexOf('BTS') >= 0) {
    if (fn.indexOf('공공') >= 0) zonedBase = 'BTS-공공';
    else if (fn.indexOf('퍼블릭') >= 0) zonedBase = 'BTS-퍼블릭';
  } else if (fn.indexOf('빌링') >= 0) {
    if (fn.indexOf('공공') >= 0) zonedBase = '빌링-공공';
    else if (fn.indexOf('퍼블릭') >= 0) zonedBase = '빌링-퍼블릭';
  }
  if (zonedBase) {
    if (sn.indexOf('접근') >= 0 && sn.indexOf('로그') >= 0) return zonedBase + '-접근';
    if (sn.indexOf('권한') >= 0 && sn.indexOf('로그') >= 0) return zonedBase + '-권한';
    if (sn.indexOf('행위') >= 0 && sn.indexOf('로그') >= 0) return zonedBase + '-행위';
  }

  // LaaS OM: 시트명이 "Data"뿐이라 파일명으로만 판별 가능.
  if (fn.indexOf('LaaS') >= 0 && fn.indexOf('OM') >= 0 && sn === 'Data') {
    return 'iLaaS-OM';
  }
  // LaaS 내부 어드민: 시트명이 "Sheet1"뿐이라 파일명으로만 판별 가능.
  if (fn.indexOf('LaaS') >= 0 && fn.indexOf('내부') >= 0 && sn === 'Sheet1') {
    return 'iLaaS';
  }

  // QueryPie: PI/CSAP 두 존이 "UserLoginHistory" 등 완전히 같은 시트명을 쓰기
  // 때문에, 시트명만으로 매칭하면(getFormatKeyBySheetName) 항상 SYSTEM_FORMATS에
  // 먼저 등록된 PI 쪽으로만 잡혀서 CSAP 파일이 통째로 PI로 집계돼버린다.
  // 파일명에 "QueryPie"+존 키워드가 있으면 그걸로 먼저 존을 확정한다.
  if (fn.indexOf('QueryPie') >= 0) {
    var qpZone = fn.indexOf('CSAP') >= 0 ? 'CSAP' : (fn.indexOf('PI') >= 0 ? 'PI' : null);
    var qpSuffix = QUERYPIE_SHEET_SUFFIX[sn];
    if (qpZone && qpSuffix) {
      var qpKey = 'QueryPie-' + qpZone + '-' + qpSuffix;
      if (SYSTEM_FORMATS[qpKey]) return qpKey;
    }
  }

  // KC-Admin: 리전 접두사(KR2_, kr-gov-central-1_ 등)가 시트명 앞에 붙어있어
  // 접미사로만 판별 가능 (공백 유무 등은 정규화해서 비교). 리전 접두사 자체가
  // 존을 나타낸다: "kr-gov-*"는 공공, 그 외(KR2 등)는 퍼블릭.
  var normalizedSheet = sn.replace(/\s+/g, '');
  var kcSuffixMatch = normalizedSheet.match(/^(.+?)_(접속로그|수행업무로그|권한로그)$/);
  if (kcSuffixMatch) {
    var kcZone = /^kr-gov/i.test(kcSuffixMatch[1]) ? 'KCadmin-공공' : 'KCadmin-퍼블릭';
    var kcTypeMap = {'접속로그': '접속', '수행업무로그': '수행업무', '권한로그': '권한'};
    return kcZone + '-' + kcTypeMap[kcSuffixMatch[2]];
  }

  // IAM: 시트명 접두사가 존을 나타낸다 - "IAM_gov_*"는 공공, "IAM_public_*"는
  // 퍼블릭. 엑셀 시트명 31자 제한 때문에 뒤쪽 날짜 숫자가 파일마다 다르게
  // 잘려나가 있어(예: "IAM_gov_Personal Info Log_20260") 접두사/로그종류
  // 키워드만으로 판별한다.
  if (/^IAM_gov_/i.test(sn) || /^IAM_public_/i.test(sn)) {
    var iamZone = /^IAM_gov_/i.test(sn) ? 'IAM-공공' : 'IAM-퍼블릭';
    var iamType = sn.indexOf('Access') >= 0 ? 'Access'
      : sn.indexOf('Permission') >= 0 ? 'Permission'
      : (sn.indexOf('Personal') >= 0 || sn.indexOf('Info') >= 0) ? 'Action'
      : null;
    if (iamType) return iamZone + '-' + iamType;
  }

  // 파트너센터: 시트명에 공백이 있어("시스템 접속 로그") 일반 규칙에 안 걸림.
  // "권한로그"는 BTS/빌링과 겹치는 이름이라 파일명으로 한 번 더 확인한다.
  if (normalizedSheet === '시스템접속로그') return '파트너센터-시스템접속';
  if (normalizedSheet === '개인정보처리로그') return '파트너센터-개인정보처리';
  if (normalizedSheet === '권한로그' && fn.indexOf('파트너센터') >= 0) return '파트너센터-권한';

  // 젠데스크: 한 파일에 "26.07"처럼 "YY.MM" 형식의 시트가 여러 달치 들어있다.
  if (/^\d{2}\.\d{2}$/.test(sn) && fn.indexOf('젠데스크') >= 0) return '젠데스크';

  // 교육센터: "엑세스"는 등록된 "액세스"의 철자 변형(오타)이라 그대로는 안 걸린다.
  // "가입탈퇴"는 컬럼 구조가 "액세스"와 동일하지만 별도 로그 종류로 등록한다.
  // 둘 다 파일명으로 한 번 더 확인해 다른 시스템의 동명 시트와 겹치지 않게 한다.
  if (fn.indexOf('교육센터') >= 0) {
    if (normalizedSheet === '엑세스') return '교육센터-액세스';
    if (normalizedSheet === '가입탈퇴') return '교육센터-가입탈퇴';
  }

  return null;
}

function sheetToLogText(sheet) {
  var rows = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: '', raw: false});
  return rows.map(function(row) {
    return row.map(function(cell) {
      return String(cell).replace(/[\r\n\t]+/g, ' ');
    }).join('\t');
  }).join('\n');
}

function makeSkippedResult(label, reason, sheetName, formatKey) {
  return {
    system: label,
    sheetName: sheetName,
    formatKey: formatKey || null,
    skipped: true,
    failGroups: [], ahSummary: [], multiSummary: [], dlSummary: [], uaSummary: [],
    unauthorizedAccess: [], unauthorizedPermissions: [], overQueries: [],
    ipCheckNote: '', logSchema: [], logMissing: [], retentionSummary: [],
    logIssue: reason
  };
}

// 워크북의 모든 시트를 스캔해 시트명 → 로그 형식을 자동 매칭하고 각 시트를 독립적으로 분석한다.
// 순수 함수: DOM이나 전역 UI 상태를 건드리지 않으므로 브라우저/CLI 양쪽에서 그대로 재사용 가능.
function analyzeWorkbook(workbook, sheetNames, fileName) {
  var fileLabel = fileName.replace(/\.[^/.]+$/, '');
  var results = [];

  for (var i = 0; i < sheetNames.length; i++) {
    var sheetName = sheetNames[i];
    var label = fileLabel + ' - ' + sheetName;
    var formatKey = resolveFormatKeyForSheet(sheetName, fileName) || getFormatKeyBySheetName(sheetName);

    if (!formatKey) {
      results.push(makeSkippedResult(label, '형식 미등록 - 점검 제외', sheetName, null));
      continue;
    }

    var sheet = workbook.Sheets[sheetName];
    var logText = sheet ? sheetToLogText(sheet) : '';
    if (!logText.trim()) {
      results.push(makeSkippedResult(label, '빈 시트 - 점검 제외', sheetName, formatKey));
      continue;
    }

    var result = analyzeLog(logText, label, formatKey);
    result.sheetName = sheetName;
    result.formatKey = formatKey;
    results.push(result);
  }

  return results;
}

  return {
    SYSTEM_FORMATS: SYSTEM_FORMATS,
    SYSTEM_AUTHORIZED_IPS: SYSTEM_AUTHORIZED_IPS,
    AUTHORIZED_IP_RULES: AUTHORIZED_IP_RULES,
    CHECK_CONFIG: CHECK_CONFIG,
    LDAP_PREFIXES: LDAP_PREFIXES,
    isSystemAccount: isSystemAccount,
    systemAccountBase: systemAccountBase,
    isUuidAccount: isUuidAccount,
    extractLdap: extractLdap,
    OVER_QUERY_THRESHOLD: OVER_QUERY_THRESHOLD,
    UNAUTHORIZED_ACCESS_KEYWORDS: UNAUTHORIZED_ACCESS_KEYWORDS,
    UNAUTHORIZED_PERMISSION_KEYWORDS: UNAUTHORIZED_PERMISSION_KEYWORDS,
    LOG_ISSUES: LOG_ISSUES,
    ipToNumber: ipToNumber,
    matchesIPRule: matchesIPRule,
    hasConfiguredAuthorizedIPs: hasConfiguredAuthorizedIPs,
    classifyIP: classifyIP,
    containsKeyword: containsKeyword,
    isValidZendeskUser: isValidZendeskUser,
    parseRowCount: parseRowCount,
    hasColumnMapping: hasColumnMapping,
    isValidIPAddress: isValidIPAddress,
    tsToMs: tsToMs,
    tsToKSTHour: tsToKSTHour,
    tsDisplay: tsDisplay,
    pad: pad,
    splitLine: splitLine,
    getCol: getCol,
    getAction: getAction,
    analyzeLog: analyzeLog,
    camelWords: camelWords,
    getFormatKeyBySheetName: getFormatKeyBySheetName,
    resolveFormatKeyForSheet: resolveFormatKeyForSheet,
    sheetToLogText: sheetToLogText,
    makeSkippedResult: makeSkippedResult,
    analyzeWorkbook: analyzeWorkbook
  };
});
