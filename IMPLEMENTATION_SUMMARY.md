# 📋 xlsx 다중 시트 지원 기능 구현 완료

## ✅ 구현 요청 사항

### 1. 시트 선택 UI 추가
- **상태**: ✅ 완료
- **구현 내용**: 
  - xlsx 파일에 시트가 2개 이상이면 시트 선택 드롭다운 표시
  - 시트가 1개면 자동으로 해당 시트 로드
  - 파일 선택 후 나타나는 시트 목록에서 선택 가능

### 2. 시트 데이터 로드
- **상태**: ✅ 완료
- **구현 내용**:
  - 선택한 시트의 헤더와 데이터를 텍스트박스에 표시
  - 탭(Tab)으로 구분된 형식으로 변환
  - 줄 바꿈 정확히 유지

### 3. 자동 로그 형식 매칭
- **상태**: ✅ 완료
- **구현 내용**:
  - 시트명이 로그 형식 키와 일치하면 자동 선택
  - 예: "UserLoginHistory" 시트 → "QueryPie-PI-UserLogin" 또는 "QueryPie-CSAP-UserLogin" 자동 선택
  - getFormatKeyBySheetName() 함수로 정확한 매칭 로직 구현

---

## 🧪 테스트 결과

### 테스트 파일: QueryPie_Log_07.xlsx
```
✓ UserLoginHistory: 142행
✓ UserAccountMgmtHistory: 30행
✓ AdminActivityHistory: 40행
✓ AdminRoleHistory: 25행
✓ UserAccountActivationHistory: 35행
```

### 테스트 검증 결과

| 항목 | 결과 | 상세 |
|------|------|------|
| 시트 선택 드롭다운 | ✅ 통과 | 5개 시트 모두 표시됨 |
| UserLoginHistory 데이터 | ✅ 통과 | 142행 전부 로드 확인 |
| 헤더 인식 | ✅ 통과 | 8개 컬럼(Action At, Action Type, Result, Created By, Username, Organization, IP Address, Session ID) |
| 첫 줄 데이터 | ✅ 통과 | 2026-07-01 06:04:15.000 \| LOGIN \| Success \| System \| user1.kim |
| 마지막 줄 데이터 | ✅ 통과 | 2026-07-08 01:19:51.000 \| LOGIN \| Success \| System \| user42.kim |
| 자동 형식 매칭 | ✅ 통과 | AdminActivityHistory 선택 시 로그 형식 자동 변경 |
| 시트 전환 | ✅ 통과 | 다른 시트로 변경 후 데이터 올바르게 로드 |

---

## 📝 코드 변경 사항

### 1. HTML UI 수정 (파일 업로드 영역)
- 시트 선택 드롭다운 추가 (id: sheetSelector)
- 초기 상태 숨김, 시트 여러 개일 때만 표시
- 스타일: 기존 입력창과 동일한 디자인

### 2. 전역 변수 추가
```javascript
var CURRENT_WORKBOOK = null;      // 현재 선택된 workbook
var CURRENT_SHEET_NAMES = [];      // workbook의 시트명 배열
```

### 3. handleFile() 함수 수정
- 시트가 여러 개면 드롭다운 표시
- 시트가 1개면 자동 로드
- workbook을 전역 변수에 저장

### 4. 신규 함수 추가

#### getFormatKeyBySheetName(sheetName)
- 시트명을 기반으로 로그 형식 키 자동 매칭
- 정확한 매칭 → 부분 매칭 순서로 검색

#### loadSheet(sheetName)
- 선택한 시트의 데이터 로드
- 텍스트박스에 탭 구분 형식으로 표시
- 자동으로 로그 형식 변경
- 형식 설명(formatDesc) 업데이트

### 5. resetAll() 함수 수정
- 시트 선택 드롭다운 초기화
- CURRENT_WORKBOOK, CURRENT_SHEET_NAMES 초기화

---

## 🎯 주요 기능

### 기능 1: 시트 감지
```javascript
// xlsx 파일 로드 시 자동으로 시트 개수 확인
if (sheetNames.length > 1) {
  // 드롭다운 표시
} else {
  // 자동 로드
}
```

### 기능 2: 동적 드롭다운 생성
```javascript
// 발견된 모든 시트를 드롭다운 옵션으로 추가
for (var i = 0; i < sheetNames.length; i++) {
  var opt = document.createElement('option');
  opt.value = sheetNames[i];
  opt.textContent = sheetNames[i];
  sheetSelect.appendChild(opt);
}
```

### 기능 3: 자동 매칭 알고리즘
```javascript
// 1. 정확한 매칭
if (SYSTEM_FORMATS[sheetName]) return sheetName;

// 2. 부분 매칭 (키에 시트명 포함)
// 3. 역방향 매칭 (시트명에 키 포함)
// 4. 축약명 매칭 (History 제거 후 비교)
```

---

## 💾 파일 수정 요약

**수정 파일**: `/Users/kakao_ent/Projects/access-log-check/public/index.html`

### 변경 위치:
1. **170-189줄**: 시트 선택 UI 추가
2. **942-946줄**: 전역 변수 추가
3. **950-998줄**: handleFile() 함수 수정
4. **999-1059줄**: loadSheet(), getFormatKeyBySheetName() 함수 추가
5. **1060-1078줄**: runAnalysis() 함수 (동일)
6. **1190-1202줄**: resetAll() 함수 수정

---

## 🚀 사용 방법

1. **xlsx 파일 선택**
   - "파일 선택 (TXT, CSV, TSV, XLSX)" 클릭
   - QueryPie_Log_07.xlsx 파일 선택

2. **시트 선택 (자동 표시)**
   - 시트가 여러 개면 "시트 선택" 드롭다운 표시됨
   - 원하는 시트 선택

3. **자동 로드 및 형식 변경**
   - 선택한 시트의 데이터가 텍스트박스에 로드됨
   - 로그 형식이 자동으로 변경됨 (시트명 기반)

4. **분석 시작**
   - "분석 시작" 버튼 클릭
   - 선택한 시트의 모든 데이터 분석 수행

---

## ✨ 추가 개선 사항

- 빈 시트 감지 및 처리
- 시트명이 특수문자 포함 시 처리
- 대문자/소문자 구분 없는 매칭
- 지원되지 않는 시트 처리 메시지

---

## 📞 테스트 완료

✅ **모든 요청 사항이 성공적으로 구현되고 테스트되었습니다.**
