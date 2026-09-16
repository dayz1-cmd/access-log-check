// iLaaS-OM 분석 테스트
const logData = `1	2026.07.31 18:12:37	john.kim	203.246.171.161	KAKAO	SELECT	/api/users	/query/users	S	-	0
2	2026.07.31 18:13:45	jane.lee	211.56.96.34	KAKAO	INSERT	/api/admin	/admin/update	S	-	1
3	2026.07.31 18:15:20	bob.park	172.16.0.10	KAKAO	DELETE	/api/data	/data/remove	F	-	0
4	2026.07.31 06:30:15	kim.chul	165.85.218.54	KAKAO	DOWNLOAD	/export	/report/export	S	-	5000
5	2026.07.31 02:45:00	lee.su	203.0.113.99	KAKAO	SELECT	/api/info	/query/info	S	-	100
6	2026.07.31 19:20:30	han.mi	172.25.1.50	KAKAO	UPDATE	/api/config	/config/update	S	-	0`;

const SYSTEM_FORMATS = {
  'iLaaS-OM': { delim:'tab', ts:1, user:2, ip:3, action:[5,6,7], resultCol:8, rowCountCol:10, desc:'Id, 행위 실행시간, 사용자 명, IP, 회사ID, 액션 유형, 실행 위치, 요청 URI' }
};

function hasColumnMapping(mapping) {
  return mapping !== undefined && mapping !== null && mapping !== -1 &&
    (typeof mapping === 'number' || (typeof mapping === 'string' && mapping.indexOf('json:') === 0));
}

function analyzeLogFormat(logText, fmtKey) {
  const fmt = SYSTEM_FORMATS[fmtKey];
  
  const logSchema = [
    {label:'계정', present:hasColumnMapping(fmt.user)},
    {label:'접속일시', present:hasColumnMapping(fmt.ts)},
    {label:'IP', present:hasColumnMapping(fmt.ip)},
    {label:'수행업무', present:Array.isArray(fmt.action) && fmt.action.length > 0},
    {label:'로그인 성공/실패 여부', present:hasColumnMapping(fmt.resultCol)},
    {label:'다운로드 사유', present:hasColumnMapping(fmt.downloadReasonCol)}
  ];
  
  logSchema.forEach(function(item) {
    if (!item.present) {
      item.message = item.label === '다운로드 사유' ?
        '다운로드 사유 누락 - 로그 양식 자체에 해당 컬럼 없음' :
        item.label + ' 누락 - 로그 양식 자체에 해당 컬럼 없음';
    }
  });
  
  return logSchema;
}

// 테스트 실행
console.log('=== iLaaS-OM 형식 분석 ===\n');
const schema = analyzeLogFormat(logData, 'iLaaS-OM');

console.log('로그 양식 구조 검사 결과:');
console.log('─'.repeat(50));
schema.forEach((item, index) => {
  const status = item.present ? '✓ 존재' : '✗ ' + item.message;
  console.log(`${index + 1}. ${item.label.padEnd(20)} : ${status}`);
});

console.log('\n상세 설정 정보:');
console.log('─'.repeat(50));
const fmt = SYSTEM_FORMATS['iLaaS-OM'];
console.log(`resultCol (로그인 성공/실패): ${fmt.resultCol !== undefined ? fmt.resultCol : '없음'}`);
console.log(`downloadReasonCol (다운로드 사유): ${fmt.downloadReasonCol !== undefined ? fmt.downloadReasonCol : '없음'}`);
console.log(`rowCountCol (행 개수): ${fmt.rowCountCol !== undefined ? fmt.rowCountCol : '없음'}`);

console.log('\n로그 샘플 데이터 파싱:');
console.log('─'.repeat(50));
const lines = logData.trim().split('\n');
const firstDataLine = lines[0].split('\t');
console.log(`첫 번째 행: ${lines[0]}`);
console.log(`  - 컬럼 0 (ID): ${firstDataLine[0]}`);
console.log(`  - 컬럼 1 (행위 실행시간): ${firstDataLine[1]}`);
console.log(`  - 컬럼 2 (사용자 명): ${firstDataLine[2]}`);
console.log(`  - 컬럼 3 (IP): ${firstDataLine[3]}`);
console.log(`  - 컬럼 8 (실행 결과): ${firstDataLine[8]}`);
console.log(`  - 컬럼 10 (행 개수): ${firstDataLine[10]}`);
