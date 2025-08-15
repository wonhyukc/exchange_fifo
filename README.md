# 외환 거래 수익률 계산 (FIFO 방식)

Google Apps Script를 사용한 외환 거래의 FIFO(First In, First Out) 방식 수익률 계산 시스템입니다.

## 주요 기능

### 1. FIFO 계산 로직
- **마지막 잔액이 0이었던 이후를 기준**으로 계산
- **동일 금액 매칭 우선**: 가장 가까운 같은 금액이 있으면 이를 기반으로 수익 계산
- **표준 FIFO**: 동일 금액이 없으면 FIFO 방식으로 계산

### 2. 자동화 기능
- **실시간 계산**: 거래 데이터 입력 시 자동으로 FIFO 계산 수행
- **날짜 자동 입력**: 거래 입력 시 현재 날짜/시간 자동 기록
- **잔고 검증**: 매도 금액이 현재 잔고를 초과할 경우 경고 메시지 표시

### 3. 계산 항목
- **F열 (잔고)**: 누적 외화 잔고
- **G열 (거래후잔액)**: 해당 거래 후 남은 매수 수량
- **H열 (FIFO 원가)**: 매도 시 FIFO 방식으로 계산된 원가
- **I열 (환차손익)**: 매도 시 실현 손익 (원화금액 - FIFO 원가)
- **O열 (기간)**: 매수부터 매도까지의 보유 기간 (일 단위, 가중평균)

## 시트 구조

### 기본 시트 (예: USD_fifo)
| 열 | 설명 | 데이터 타입 |
|---|---|---|
| A | 구분 | 텍스트 |
| B | 날짜 | 날짜 (자동 입력) |
| C | 외화금액 | 숫자 (양수: 매수, 음수: 매도) |
| D | 환율 | 숫자 |
| E | 원화금액 | 숫자 |
| F | 잔고 | 숫자 (자동 계산) |
| G | 거래후잔액 | 숫자 (자동 계산) |
| H | FIFO 원가 | 숫자 (자동 계산) |
| I | 환차손익 | 숫자 (자동 계산) |
| O | 기간 | 숫자 (자동 계산, 일 단위) |

### 로그 시트 (log)
- 거래 로그 기록용
- E열, H열 입력 시 날짜 자동 입력

## 설정

### 상수 설정
```javascript
const FIFO_SHEET_SUFFIX = "_fifo";  // FIFO 시트 접미사
const HEADER_ROW = 2;               // 헤더 행 번호
const RESET_THRESHOLD = 0.000001;   // 잔고 리셋 임계값
```

### 컬럼 매핑
```javascript
const COL_B_DATE = 2;       // B열 (날짜)
const COL_C_AMOUNT = 3;     // C열 (외화금액)
const COL_D_RATE = 4;       // D열 (환율)
const COL_E_KRW = 5;        // E열 (원화금액)
const COL_F_BALANCE = 6;    // F열 (잔고)
const COL_G_REMAINING = 7;  // G열 (거래후잔액)
const COL_H_FIFO_COST = 8;  // H열 (FIFO 원가)
const COL_I_PROFIT_LOSS = 9; // I열 (환차손익)
const COL_O_DURATION = 15;  // O열 (기간)
```

## 사용법

1. **시트 생성**: 통화명 + "_fifo" 접미사로 시트 생성 (예: USD_fifo)
2. **헤더 설정**: 2행에 컬럼 헤더 입력
3. **거래 입력**: C열(외화금액)과 D열(환율) 입력
4. **자동 계산**: 입력 시 자동으로 F, G, H, I, O열 계산

## 주요 함수

### `onEdit(e)`
- 시트 수정 시 자동 실행
- 날짜 자동 입력
- FIFO 계산 트리거

### `updateFIFOAndWrite(sheet)`
- FIFO 계산 수행 및 결과 시트에 기록

### `calculateFIFO(data, sheetName)`
- 핵심 FIFO 계산 로직
- 동일 금액 매칭 및 표준 FIFO 처리

## 업데이트 이력

- **2025-05-30**: 로거 주석 처리, 거래기간 계산 추가, 잔고계산 정상화 (0.2 제한 제거)
- **2025-02-23**: FIFO 처리 완료

## 리소스

- **Remote Repository**: `git@github.com:wonhyukc/exchange_fifo.git`
- **Script ID**: `1JjjN9_NXSsRn266VEtCXGyJYVrPhQjvGTdH_2RQVs9O8iEFa0Z1JSmtR`

## 개발 환경

- **Google Apps Script**
- **Google Sheets API**
- **JavaScript ES6+**

## resource
remote repo: git@github.com:wonhyukc/exchange_fifo.git
script id: 1JjjN9_NXSsRn266VEtCXGyJYVrPhQjvGTdH_2RQVs9O8iEFa0Z1JSmtR

