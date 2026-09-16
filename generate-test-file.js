const XLSX = require('xlsx');

// UserLoginHistory 데이터 생성 (142행)
const userLoginData = [['Action At', 'Action Type', 'Result', 'Created By', 'Username', 'Organization', 'IP Address', 'Session ID']];
for (let i = 1; i <= 142; i++) {
  const time = new Date(2026, 6, 1 + Math.floor((i-1) / 20), 8 + Math.floor(Math.random() * 12), Math.floor(Math.random() * 60), Math.floor(Math.random() * 60));
  const formattedTime = time.toISOString().replace('T', ' ').replace('Z', '');
  userLoginData.push([
    formattedTime,
    'LOGIN',
    Math.random() > 0.1 ? 'Success' : 'Failure',
    'System',
    `user${i % 50}.kim`,
    'KAKAO',
    `203.246.171.${100 + (i % 50)}`,
    `session${i}`
  ]);
}

// UserAccountMgmtHistory 데이터
const userMgmtData = [['Account ID', 'Username', 'Department', 'Action', 'Created At', 'Action Type', 'Details']];
for (let i = 1; i <= 30; i++) {
  const time = new Date(2026, 6, 1 + Math.floor((i-1) / 5), 9, i * 2);
  const formattedTime = time.toISOString().replace('T', ' ').replace('Z', '');
  userMgmtData.push([
    i,
    `account${i}`,
    'Dev Team',
    'Create',
    formattedTime,
    'USER_CREATE',
    `Created new account: user${i}`
  ]);
}

// AdminActivityHistory 데이터
const adminActivityData = [['Activity ID', 'Admin Name', 'Action', 'Target', 'Timestamp', 'IP', 'Status', 'Details']];
for (let i = 1; i <= 40; i++) {
  const time = new Date(2026, 6, 2 + Math.floor((i-1) / 10), 10, i * 1.5);
  const formattedTime = time.toISOString().replace('T', ' ').replace('Z', '');
  adminActivityData.push([
    i,
    `admin${i % 5}.lee`,
    'MODIFY',
    `User${i}`,
    formattedTime,
    `211.56.96.${50 + i}`,
    'SUCCESS',
    `Modified admin permissions for user ${i}`
  ]);
}

// AdminRoleHistory 데이터
const adminRoleData = [['Role Change ID', 'Timestamp', 'Admin Name', 'User', 'Old Role', 'New Role', 'Reason']];
for (let i = 1; i <= 25; i++) {
  const time = new Date(2026, 6, 3 + Math.floor((i-1) / 8), 11, i * 2.4);
  const formattedTime = time.toISOString().replace('T', ' ').replace('Z', '');
  adminRoleData.push([
    i,
    formattedTime,
    `admin${i % 5}.park`,
    `user${i}.kim`,
    'User',
    i % 3 === 0 ? 'Manager' : 'Admin',
    'Role promotion due to team restructuring'
  ]);
}

// UserAccountActivationHistory 데이터
const activationData = [['Activation ID', 'Username', 'Created At', 'Activated At', 'Activated By', 'Status', 'Expiry Date', 'Notes']];
for (let i = 1; i <= 35; i++) {
  const createdTime = new Date(2026, 6, 1 + Math.floor((i-1) / 7), 7, i);
  const activatedTime = new Date(createdTime.getTime() + 3600000);
  const createdStr = createdTime.toISOString().replace('T', ' ').replace('Z', '');
  const activatedStr = activatedTime.toISOString().replace('T', ' ').replace('Z', '');
  const expiryDate = new Date(activatedTime.getTime() + 30 * 24 * 3600000).toISOString().split('T')[0];
  activationData.push([
    i,
    `newuser${i}.choi`,
    createdStr,
    activatedStr,
    `admin${i % 5}.kim`,
    'ACTIVE',
    expiryDate,
    'New employee onboarding'
  ]);
}

// 워크북 생성 및 시트 추가
const ws1 = XLSX.utils.aoa_to_sheet(userLoginData);
const ws2 = XLSX.utils.aoa_to_sheet(userMgmtData);
const ws3 = XLSX.utils.aoa_to_sheet(adminActivityData);
const ws4 = XLSX.utils.aoa_to_sheet(adminRoleData);
const ws5 = XLSX.utils.aoa_to_sheet(activationData);

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws1, 'UserLoginHistory');
XLSX.utils.book_append_sheet(wb, ws2, 'UserAccountMgmtHistory');
XLSX.utils.book_append_sheet(wb, ws3, 'AdminActivityHistory');
XLSX.utils.book_append_sheet(wb, ws4, 'AdminRoleHistory');
XLSX.utils.book_append_sheet(wb, ws5, 'UserAccountActivationHistory');

// 파일 저장
XLSX.writeFile(wb, '/Users/kakao_ent/Projects/access-log-check/QueryPie_Log_07.xlsx');
console.log('✓ QueryPie_Log_07.xlsx 생성 완료');
console.log('  - UserLoginHistory: 142행');
console.log('  - UserAccountMgmtHistory: 30행');
console.log('  - AdminActivityHistory: 40행');
console.log('  - AdminRoleHistory: 25행');
console.log('  - UserAccountActivationHistory: 35행');
