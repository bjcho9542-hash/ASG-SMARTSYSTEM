// 바보

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ASG_도구')
    .addItem('데이터 통합(푸드)', 'mergeAllRegionsFood')
    .addItem('데이터 통합(논푸드)', 'mergeAllRegionsNonFood')
    .addToUi();
}

// 푸드 통합
function mergeAllRegionsFood() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const regionSheets = [
    '충청도', '경기도', '경상도', '전라도',
    '부산시', '대구시', '대전시', '광주시',
    '울산시', '인천시', '강원도', '제주도',
    '서울시', '세종시'
  ];

  let mergedSheet = ss.getSheetByName('전체통합');

  if (mergedSheet) {
    mergedSheet.clear();
  } else {
    mergedSheet = ss.insertSheet('전체통합');
  }

  // 헤더 (카테고리 추가)
  const headers = ['사업자번호', '상호명', '상세주소', '전화번호', '가입된 플랫폼', '타입', '카테고리'];
  mergedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  mergedSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');

  let currentRow = 2;

  regionSheets.forEach(regionName => {
    const sheet = ss.getSheetByName(regionName);

    if (!sheet) {
      Logger.log(`시트를 찾을 수 없습니다: ${regionName}`);
      return;
    }

    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      Logger.log(`데이터가 없습니다: ${regionName}`);
      return;
    }

    // 푸드 원본: A(빈칸), B(타입), C(구), D(사업자번호), E(상호명), F(주소)
    const dataRange = sheet.getRange(2, 2, lastRow - 1, 5);
    const data = dataRange.getValues();

    // 매핑: 사업자번호(D), 상호명(E), 상세주소(F), 전화번호(빈칸), 가입된플랫폼(빈칸), 타입(B), 카테고리(빈칸)
    const mappedData = data.map(row => [
      row[2],  // 사업자번호 (D열)
      row[3],  // 상호명 (E열)
      row[4],  // 상세주소 (F열)
      '',      // 전화번호
      '',      // 가입된 플랫폼
      row[0],  // 타입 (B열)
      ''       // 카테고리 (푸드는 없음)
    ]);

    if (mappedData.length > 0) {
      mergedSheet.getRange(currentRow, 1, mappedData.length, 7)
        .setValues(mappedData);
      currentRow += mappedData.length;
    }

    Logger.log(`${regionName}: ${mappedData.length}개 행 추가 완료`);
  });

  for (let i = 1; i <= 7; i++) {
    mergedSheet.autoResizeColumn(i);
  }

  Logger.log(`푸드 통합 완료: 총 ${currentRow - 2}개 행`);
  SpreadsheetApp.getUi().alert(`푸드 통합 완료!\n총 ${currentRow - 2}개의 데이터가 통합되었습니다.`);
}

// 논푸드 통합
function mergeAllRegionsNonFood() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const regionSheets = [
    '충청도', '경기도', '경상도', '전라도',
    '부산시', '대구시', '대전시', '광주시',
    '울산시', '인천시', '강원도', '제주도',
    '서울시', '세종시'
  ];

  let mergedSheet = ss.getSheetByName('전체통합');

  if (mergedSheet) {
    mergedSheet.clear();
  } else {
    mergedSheet = ss.insertSheet('전체통합');
  }

  // 헤더 (카테고리 추가)
  const headers = ['사업자번호', '상호명', '상세주소', '전화번호', '가입된 플랫폼', '타입', '카테고리'];
  mergedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  mergedSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');

  let currentRow = 2;

  regionSheets.forEach(regionName => {
    const sheet = ss.getSheetByName(regionName);

    if (!sheet) {
      Logger.log(`시트를 찾을 수 없습니다: ${regionName}`);
      return;
    }

    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      Logger.log(`데이터가 없습니다: ${regionName}`);
      return;
    }

    // 논푸드 원본: A(빈칸), B(타입), C(사업자번호), D(스토어명), E(광역시도명), F(시군구), G(주소), H(카테고리)
    const dataRange = sheet.getRange(2, 2, lastRow - 1, 7);
    const data = dataRange.getValues();

    // 매핑: 사업자번호(C), 상호명(D), 상세주소(G), 전화번호(빈칸), 가입된플랫폼(빈칸), 타입(B), 카테고리(H)
    const mappedData = data.map(row => [
      row[1],  // 사업자번호 (C열 = index 1)
      row[2],  // 상호명/스토어명 (D열 = index 2)
      row[5],  // 상세주소 (G열 = index 5)
      '',      // 전화번호
      '',      // 가입된 플랫폼
      row[0],  // 타입 (B열 = index 0)
      row[6]   // 카테고리 (H열 = index 6)
    ]);

    if (mappedData.length > 0) {
      mergedSheet.getRange(currentRow, 1, mappedData.length, 7)
        .setValues(mappedData);
      currentRow += mappedData.length;
    }

    Logger.log(`${regionName}: ${mappedData.length}개 행 추가 완료`);
  });

  for (let i = 1; i <= 7; i++) {
    mergedSheet.autoResizeColumn(i);
  }

  Logger.log(`논푸드 통합 완료: 총 ${currentRow - 2}개 행`);
  SpreadsheetApp.getUi().alert(`논푸드 통합 완료!\n총 ${currentRow - 2}개의 데이터가 통합되었습니다.`);
}
