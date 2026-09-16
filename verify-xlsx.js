const XLSX = require('xlsx');

const filePath = '/Users/kakao_ent/Projects/access-log-check/QueryPie_Log_07.xlsx';
const workbook = XLSX.readFile(filePath);

console.log('=== QueryPie_Log_07.xlsx 시트 분석 ===\n');
console.log(`총 시트 개수: ${workbook.SheetNames.length}`);
console.log(`시트명: ${workbook.SheetNames.join(', ')}\n`);

workbook.SheetNames.forEach((sheetName) => {
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: '', raw: false});
  
  console.log(`📄 ${sheetName}`);
  console.log(`   - 헤더 행: 1행`);
  console.log(`   - 데이터 행: ${rows.length - 1}행`);
  console.log(`   - 총 행: ${rows.length}행`);
  
  if (sheetName === 'UserLoginHistory') {
    console.log(`   - 검증: 사용자 계정 개수: ${new Set(rows.slice(1).map(r => r[4])).size}`);
    console.log(`   - 첫 줄 데이터: ${rows[1].slice(0, 5).join(' | ')}`);
    console.log(`   - 마지막 줄 데이터: ${rows[rows.length-1].slice(0, 5).join(' | ')}`);
  }
  console.log();
});

console.log('✓ 모든 시트가 정상적으로 로드되었습니다.');
