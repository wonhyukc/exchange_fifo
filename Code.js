/**
 * @OnlyCurrentDoc
 */
/*
	20250530 loger.log 주석 처리
  20250530 거래기간 계산 추가
	20250530 잔고계산 정상. 0.2 제한 으로 인한 문제 제거
	20250223 FIFO 처리 완료
*/

// --- Configuration ---
const FIFO_SHEET_SUFFIX = "_fifo";
const HEADER_ROW = 2;
const COL_B_DATE = 2;       // B열 (날짜) - 추가
const COL_C_AMOUNT = 3;
const COL_D_RATE = 4;
const COL_E_KRW = 5;
const COL_F_BALANCE = 6;
const COL_G_REMAINING = 7;
const COL_H_FIFO_COST = 8;
const COL_I_PROFIT_LOSS = 9;
const COL_O_DURATION = 15;  // O열 (기간) - 추가
const RESET_THRESHOLD = 0.000001;
// --- End Configuration ---

/**
 * Convert column letter to column number
 */
function getColumnNumber(column) {
  return column.toUpperCase().charCodeAt(0) - 64;
}

/*
  col2Check 컬럼에 입력이 있으면, 우측으로 col2Put 만큼 이동하여 날짜 시각을 입력한다
*/
function putDate(selectedCell, col2Check, col2Put) {
  if (selectedCell.getColumn() === col2Check) {
    const dateTimeCell = selectedCell.offset(0, col2Put);
    if (dateTimeCell.isBlank()) {
      dateTimeCell.setValue(new Date());
    }
  }
}

/**
 * 시트가 수정될 때 자동으로 실행될 트리거 함수입니다.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e 이벤트 객체
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  const ui = SpreadsheetApp.getUi();

  if (sheetName === 'log') {
    putDate(range, getColumnNumber('E'), -1);
    putDate(range, getColumnNumber('H'), -1);
    return;
  }

  if (!sheetName.endsWith(FIFO_SHEET_SUFFIX) || editedRow <= HEADER_ROW) {
    return;
  }

  putDate(range, COL_C_AMOUNT, -1); // C열 입력 시 B열에 날짜 자동 입력

  if (editedCol === COL_C_AMOUNT && range.getValue() < 0) {
    const currentAmount = Number(range.getValue());
    let prevBalance = 0;
    if (editedRow -1 > HEADER_ROW) { // 직전 행이 헤더 아래인지 확인
        prevBalance = Number(sheet.getRange(editedRow - 1, COL_F_BALANCE).getValue() || 0);
    } else if (editedRow -1 === HEADER_ROW && editedRow > HEADER_ROW) { // 수정된 행이 헤더 바로 다음 행일 경우
        // 이 경우, 이전 잔고는 0으로 간주하거나, 시트의 가장 처음 상태에 따라 다를 수 있음
        // 여기서는 0으로 처리. 필요시 로직 수정.
        prevBalance = 0;
    }


    //logger.log(`onEdit for ${sheetName}: 현재 입력된 금액 (C열): ${currentAmount}, 직전 행 외화 잔고 (F열): ${prevBalance}`);

    if (Math.abs(currentAmount) > prevBalance + RESET_THRESHOLD) { // RESET_THRESHOLD를 고려하여 비교
      ui.alert(
        '경고: 매도 금액 초과',
        `(${sheetName}) 입력하신 매도 금액 (${Math.abs(currentAmount)})이 현재 외화 잔고 (${prevBalance.toFixed(6)})보다 많습니다.\n다시 확인해주세요.`,
        ui.ButtonSet.OK
      );
      // 초과 시 해당 셀 값 지우기 (선택 사항)
      // range.setValue(''); 
    }
  }

  if (editedCol === COL_C_AMOUNT || editedCol === COL_D_RATE) {
    try {
      updateFIFOAndWrite(sheet);
    } catch (error) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`FIFO 계산 중 오류 발생 (${sheetName}): ` + error.message, "오류", 30);
      //logger.log(`Error during FIFO update for ${sheetName}: ` + error);
    }
  }
}

/**
 * FIFO 계산을 수행하고 결과를 특정 시트에 쓰는 메인 함수입니다.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} currentSheet 대상 시트 객체
 */
function updateFIFOAndWrite(currentSheet) {
  const sheetName = currentSheet.getName();
  //logger.log(`updateFIFOAndWrite called for sheet: ${sheetName}`);

  const lastRow = currentSheet.getLastRow();
  if (lastRow <= HEADER_ROW) {
    //logger.log(`No data found below header row in sheet: ${sheetName}. Last row: ${lastRow}`);
    // 데이터가 없을 경우, 기존 계산된 값(F,G,H,I,O)을 지울 수 있습니다.
    const numColsToClear = COL_O_DURATION - COL_F_BALANCE + 1; // F열부터 O열까지의 컬럼 수
    if (currentSheet.getMaxRows() > HEADER_ROW) { // 시트에 행이 헤더 이상으로 존재할 때만 실행
        currentSheet.getRange(HEADER_ROW + 1, COL_F_BALANCE, currentSheet.getMaxRows() - HEADER_ROW, numColsToClear).clearContent();
    }
    return;
  }

  const dataRange = currentSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, currentSheet.getLastColumn());
  const data = dataRange.getValues();

  if (data.length === 0) {
    //logger.log(`No data values found in dataRange for sheet: ${sheetName}.`);
    return;
  }

  const results = calculateFIFO(data, sheetName);

  if (results && results.fValues && results.fValues.length > 0) {
    const numRows = results.fValues.length;

    const fValues = results.fValues.map(val => [val]);
    const gValues = results.gValues.map(val => [val]);
    const hValues = results.hValues.map(val => [val]);
    const iValues = results.iValues.map(val => [val]);
    const oValues = results.oValues.map(val => [val]); // O열(기간) 값 가져오기

    currentSheet.getRange(HEADER_ROW + 1, COL_F_BALANCE, numRows, 1).setValues(fValues);
    currentSheet.getRange(HEADER_ROW + 1, COL_G_REMAINING, numRows, 1).setValues(gValues);
    currentSheet.getRange(HEADER_ROW + 1, COL_H_FIFO_COST, numRows, 1).setValues(hValues);
    currentSheet.getRange(HEADER_ROW + 1, COL_I_PROFIT_LOSS, numRows, 1).setValues(iValues);
    currentSheet.getRange(HEADER_ROW + 1, COL_O_DURATION, numRows, 1).setValues(oValues); // O열(기간)에 쓰기
    
    SpreadsheetApp.flush(); // 변경사항 즉시 반영
  } else {
    //logger.log(`No results to write or empty results for sheet: ${sheetName}.`);
    // 이 경우에도 기존 계산된 값을 지울 수 있습니다.
    // const numColsToClear = COL_O_DURATION - COL_F_BALANCE + 1;
    // currentSheet.getRange(HEADER_ROW + 1, COL_F_BALANCE, data.length, numColsToClear).clearContent();
  }
}


/**
 * FIFO 계산 로직을 수행하는 함수입니다.
 * @param {Array[]} data - 시트에서 읽어온 데이터 (2차원 배열).
 * @param {string} sheetNameForLog - 로그 메시지에 사용할 시트 이름.
 * @return {object|null} 계산 결과 { fValues, gValues, hValues, iValues, oValues } 또는 오류 시 null.
 */
function calculateFIFO(data, sheetNameForLog) {
  let buyLots = [];
  let lastResetRowIndex = -1;

  let fValues = [];
  let gValues = [];
  let hValues = [];
  let iValues = [];
  let oValues = []; // 기간 값을 저장할 배열 - 추가

  let currentTotalBalance = 0;
  const logPrefix = `[${sheetNameForLog}] `;

  for (let i = 0; i < data.length; i++) {
    const currentRowIndexInData = i; // data 배열 내 현재 행 인덱스 (0부터 시작)
    const currentRowInSheet = currentRowIndexInData + HEADER_ROW + 1; // 실제 시트 행 번호

    // 데이터 가져오기 (B열: 날짜, C열: 외화금액, D열: 환율, E열: 원화금액)
    const dateValue = data[i][COL_B_DATE - 1]; // B열 (날짜) - 추가
    const amount = Number(data[i][COL_C_AMOUNT - 1] || 0);
    const rate = Number(data[i][COL_D_RATE - 1] || 0);
    const krwAmount = Number(data[i][COL_E_KRW - 1] || 0);

    let hValue = null; // FIFO 원가
    let iValue = null; // 환차손익
    let oValue = null; // 기간 - 추가 (현재 행의 기간 값)
    const currentLastResetIndex = lastResetRowIndex;

    // 현재 행의 날짜 객체 가져오기 (유효성 검사 포함) - 추가
    let currentDate = null;
    if (dateValue && (typeof dateValue.getTime === 'function' || !isNaN(new Date(dateValue).getTime()))) {
      currentDate = new Date(dateValue);
    } else {
      //logger.log(logPrefix + "Invalid or missing date at sheet row " + currentRowInSheet + ". Value: " + dateValue);
    }

    // C열 또는 D열(매도 시 환율)이 유효한 숫자가 아니면 계산 로직 건너뛰기
    if (isNaN(amount) || (amount !== 0 && isNaN(rate) && amount < 0 )) { // 매도 시(amount < 0)에는 rate가 없으면 FIFO 원가 계산 불가
      // hValue, iValue, oValue는 null로 유지
    } else if (amount > 0) { // 매수 건
      if (currentDate && !isNaN(currentDate.getTime())) { // 유효한 날짜가 있을 때만 buyLots에 추가
        buyLots.push({
          originalQty: amount, 
          rate: rate, 
          remainingQty: amount, 
          rowIndex: currentRowIndexInData,
          date: currentDate // 매수 날짜 저장 - 추가
        });
        //logger.log(logPrefix + "Buy Lot Added: SheetRow " + currentRowInSheet + ", Qty: " + amount + ", Rate: " + rate + ", Date: " + currentDate.toISOString().slice(0,10));
      } else {
         //logger.log(logPrefix + "Skipping buy lot due to invalid date at sheet row " + currentRowInSheet);
      }
      // 매수 건이므로 hValue, iValue, oValue는 null
    } else if (amount < 0) { // 매도 건
        //logger.log(logPrefix + "--- Processing Sell SheetRow: " + currentRowInSheet + ", Amount: " + amount + (currentDate ? ", Date: " + currentDate.toISOString().slice(0,10) : ", Date: N/A") + " ---");
        
        const amountToCover = Math.abs(amount);
        let totalFifoCost = 0;
        let remainingSellQty = amountToCover;
        let foundExactMatch = false;

        let consumedLotsDetailsForDuration = []; // 이번 매도 건에 소진된 로트 정보 (기간 계산용) - 추가

        // [1] 같은 금액 매칭 시도
        //logger.log(logPrefix + "Checking for exact match for amount: " + amountToCover.toPrecision(10));
        for (let lotIndex = 0; lotIndex < buyLots.length; lotIndex++) {
          let lot = buyLots[lotIndex];
          if (lot.rowIndex > currentLastResetIndex &&
              lot.remainingQty > RESET_THRESHOLD && // 남은 수량이 거의 0이 아닌 경우
              Math.abs(lot.originalQty - amountToCover) < RESET_THRESHOLD &&
              Math.abs(lot.remainingQty - lot.originalQty) < RESET_THRESHOLD) // 한 번도 사용 안 된 로트
          {
            totalFifoCost = amountToCover * lot.rate;
            lot.remainingQty = 0; // 해당 로트 전부 소진
            remainingSellQty = 0;
            foundExactMatch = true;

            // 기간 계산용 정보 저장 - 추가
            if (lot.date && currentDate && !isNaN(lot.date.getTime()) && !isNaN(currentDate.getTime())) {
                const daysHeld = (currentDate.getTime() - lot.date.getTime()) / (1000 * 60 * 60 * 24);
                consumedLotsDetailsForDuration.push({ qty: amountToCover, days: daysHeld });
            }
            //logger.log(logPrefix + " Exact match FOUND! Consumed lot from original sheetRow " + (lot.rowIndex + HEADER_ROW + 1) + " (Date: " + (lot.date ? lot.date.toISOString().slice(0,10) : 'N/A') + "). Cost: " + totalFifoCost.toPrecision(10) + ". Sell qty left: " + remainingSellQty.toPrecision(10));
            break;
          }
        }

        // [2] 표준 FIFO 처리
        if (!foundExactMatch && remainingSellQty > RESET_THRESHOLD) {
          //logger.log(logPrefix + "Exact match not found or sell qty > 0. Standard FIFO for: " + remainingSellQty.toPrecision(10));
          for (let lotIndex = 0; lotIndex < buyLots.length; lotIndex++) {
            let lot = buyLots[lotIndex];
            if (lot.rowIndex > currentLastResetIndex && lot.remainingQty > RESET_THRESHOLD) { // 남은 수량이 유의미한 경우
              const consumeAmount = Math.min(remainingSellQty, lot.remainingQty);
              
              totalFifoCost += consumeAmount * lot.rate;
              lot.remainingQty -= consumeAmount;
              remainingSellQty -= consumeAmount;

              // 기간 계산용 정보 저장 - 추가
              if (lot.date && currentDate && !isNaN(lot.date.getTime()) && !isNaN(currentDate.getTime())) {
                  const daysHeld = (currentDate.getTime() - lot.date.getTime()) / (1000 * 60 * 60 * 24);
                  consumedLotsDetailsForDuration.push({ qty: consumeAmount, days: daysHeld });
              }
              //logger.log(logPrefix + "  FIFO: Consumed " + consumeAmount.toPrecision(10) + " from lot at original sheetRow " + (lot.rowIndex + HEADER_ROW + 1) + " (Date: " + (lot.date ? lot.date.toISOString().slice(0,10) : 'N/A') + ")" + ". New lot remaining: " + lot.remainingQty.toPrecision(10) + ". Sell qty left: " + remainingSellQty.toPrecision(10));
              
              if (remainingSellQty < RESET_THRESHOLD) {
                  remainingSellQty = 0; // 매우 작은 값은 0으로 처리
                  //logger.log(logPrefix + "  Sell qty near zero, set to 0. Breaking FIFO loop.");
                  break;
              }
            }
          }
        }

        //logger.log(logPrefix + "After FIFO, final remainingSellQty: " + remainingSellQty.toPrecision(10));

        if (remainingSellQty > RESET_THRESHOLD) { // 매도할 재고 부족
          hValue = "#N/A"; iValue = "#N/A"; oValue = "#N/A"; // 기간도 N/A 처리 - 추가
          //logger.log(logPrefix + "Insufficient buy lots for sell at sheetRow " + currentRowInSheet + ". Needed: " + remainingSellQty.toPrecision(10) + ". H, I, O set to #N/A.");
        } else { // 매도 재고 충분 또는 정확히 매도 완료
          hValue = totalFifoCost;
          if (!isNaN(krwAmount) && typeof hValue === 'number') {
            iValue = krwAmount - hValue;
          } else if (hValue === "#N/A") { // hValue가 이미 #N/A면 iValue도 #N/A
            iValue = "#N/A";
          } else { // 기타 계산 오류
            iValue = "#ERROR_I";
            //logger.log(logPrefix + "Error calculating I value for sheetRow " + currentRowInSheet + ". E=" + krwAmount + ", H=" + hValue);
          }

          // 기간 계산 (가중 평균) - 추가
          if (currentDate && !isNaN(currentDate.getTime()) && consumedLotsDetailsForDuration.length > 0) {
              let totalWeightedDays = 0;
              let totalQuantityForDuration = 0;
              for (const detail of consumedLotsDetailsForDuration) {
                  if (typeof detail.qty === 'number' && typeof detail.days === 'number' && !isNaN(detail.qty) && !isNaN(detail.days)) {
                      totalWeightedDays += detail.qty * detail.days;
                      totalQuantityForDuration += detail.qty;
                  }
              }
              if (totalQuantityForDuration > RESET_THRESHOLD) {
                  oValue = totalWeightedDays / totalQuantityForDuration;
                  oValue = Math.round(oValue); // 일 단위로 반올림
                  //logger.log(logPrefix + "Calculated weighted avg duration: " + oValue + " days for sold qty: " + totalQuantityForDuration.toPrecision(10) + " at sheetRow " + currentRowInSheet);
              } else {
                  oValue = null; // 또는 에러 표시
                  //logger.log(logPrefix + "Cannot calculate duration: total qty for duration is zero or invalid at sheetRow " + currentRowInSheet);
              }
          } else if (Math.abs(amountToCover) > RESET_THRESHOLD && (!currentDate || isNaN(currentDate.getTime()))) { // 매도 수량은 있지만 현재 날짜가 유효하지 않은 경우
              oValue = "#NO_DATE"; // 기간 계산 불가
              //logger.log(logPrefix + "Cannot calculate duration due to invalid current date at sheetRow " + currentRowInSheet);
          } else if (Math.abs(amountToCover) > RESET_THRESHOLD && consumedLotsDetailsForDuration.length === 0) { // 매도 수량은 있으나 소진된 로트 정보가 없는 경우 (로직 오류 또는 초기 재고 문제)
              oValue = "#NO_LOTS"; 
              //logger.log(logPrefix + "Warning: Sold amount exists but no consumed lot details for duration calculation at sheetRow " + currentRowInSheet);
          }
          // 만약 amountToCover 자체가 0에 가까웠다면 (C열에 0 입력) oValue는 null로 유지됨
        }
    }
    // else (amount is 0 or invalid), hValue, iValue, oValue remain null

    currentTotalBalance += amount;
    fValues.push(currentTotalBalance);
    hValues.push(hValue);
    iValues.push(iValue);
    oValues.push(oValue); // 계산된 기간 값 또는 null을 배열에 추가 - 추가

    if (!isNaN(currentTotalBalance) && Math.abs(currentTotalBalance) < RESET_THRESHOLD) {
      lastResetRowIndex = currentRowIndexInData;
      //logger.log(logPrefix + "Reset index updated to " + lastResetRowIndex + " (data index) based on total balance at sheetRow " + currentRowInSheet + ". New Balance: " + currentTotalBalance.toPrecision(10));
    } else if (isNaN(currentTotalBalance)) {
        //logger.log(logPrefix + "Warning: currentTotalBalance is NaN at sheetRow " + currentRowInSheet);
    }
  } // End of data loop

  // --- 2단계: 최종 매수 재고 상태를 기반으로 G열 (거래후잔액) 계산 ---
  //logger.log(logPrefix + "--- Calculating G values (Pass 2) ---");
  for (let i = 0; i < data.length; i++) {
      const amount = Number(data[i][COL_C_AMOUNT - 1] || 0);
      let gValue = null;

      if (isNaN(amount)) {
          gValue = null;
      } else if (amount > 0) {
          const lotState = buyLots.find(lot => lot.rowIndex === i); // rowIndex는 data 배열 기준 인덱스
          if (lotState) {
              gValue = (Math.abs(lotState.remainingQty) < RESET_THRESHOLD) ? 0 : lotState.remainingQty;
          } else {
              gValue = 0; // buyLots에 추가되지 않은 매수건 (예: 날짜 오류) 또는 로직 오류
              //logger.log(logPrefix + "Warning: Could not find buyLot state for data index " + i + " (sheetRow " + (i + HEADER_ROW + 1) + ") for G value.");
          }
      } else { // 매도 건이거나 금액이 0인 경우 G열은 0
          gValue = 0;
      }
      gValues.push(gValue);
  }
  //logger.log(logPrefix + "--- FIFO Calculation Complete ---");

  return { fValues, gValues, hValues, iValues, oValues }; // oValues 포함하여 반환 - 추가
}