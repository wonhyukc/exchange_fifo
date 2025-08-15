function showonlyactive() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .whenNumberGreaterThan(0.2)
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(6, criteria);
};





/**
 * 활성 시트의 A3부터 Q열 마지막 행까지 데이터를
 * C열, D열, A열 순서로 오름차순 정렬합니다.종류 / 산날짜 / no
 */
function myFunction() {
  // 1. 현재 활성화된 스프레드시트 및 시트 가져오기
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 2. 데이터가 있는 마지막 행 번호 찾기
  const lastRow = sheet.getLastRow();

  // 만약 데이터가 A3 아래로 없다면 (헤더만 있거나 비어있다면) 함수 종료
  if (lastRow < 3) {
    Logger.log("A3 아래로 정렬할 데이터가 없습니다.");
    return;
  }

  // 3. A3부터 Q열 마지막 행까지 범위 정의
  // getRange(시작 행, 시작 열, 행 개수, 열 개수)
  const rangeToSort = sheet.getRange(3, 1, lastRow - 2, 17); // A열=1, Q열=17

  // 4. 정의된 범위를 지정된 열 순서로 정렬 (오름차순)
  // C열(3번째 열), D열(4번째 열), A열(1번째 열) 순서
  rangeToSort.sort([
    { column: 3, ascending: true }, // C열 오름차순 외화종류
    { column: 6, ascending: true }, // G열 오름차순 판날짜
    // { column: 4, ascending: true }, // D열 오름차순 산날짜
    { column: 1, ascending: true }  // A열 오름차순 번호
  ]);

  Logger.log("데이터 정렬 완료!");
};