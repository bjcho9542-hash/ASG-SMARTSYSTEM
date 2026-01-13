// 바보

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ASG_도구')
    .addItem('데이터 통합', 'mergeAllRegions')
    .addToUi();
}

function mergeAllRegions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 통합할 시트명 리스트
  const regionSheets = [
    '충청도', '경기도', '경상도', '전라도',
    '부산시', '대구시', '대전시', '광주시',
    '울산시', '인천시', '강원도', '제주도',
    '서울시', '세종시'
  ];

  // 통합 시트 준비
  let mergedSheet = ss.getSheetByName('전체통합');

  if (mergedSheet) {
    mergedSheet.clear();
  } else {
    mergedSheet = ss.insertSheet('전체통합');
  }

  // 헤더 작성 (새 구조)
  const headers = ['사업자번호', '상호명', '상세주소', '전화번호', '가입된 플랫폼', '타입'];
  mergedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 헤더 스타일링
  mergedSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');

  let currentRow = 2;

  // 각 지역 시트 순회
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

    // 원본: A(빈칸), B(타입), C(구), D(사업자번호), E(상호명), F(주소)
    // B~F열 가져오기 (2행부터)
    const dataRange = sheet.getRange(2, 2, lastRow - 1, 5);
    const data = dataRange.getValues();

    // 매핑: 사업자번호(D=index2), 상호명(E=index3), 상세주소(F=index4), 전화번호(빈칸), 가입된플랫폼(빈칸), 타입(B=index0)
    const mappedData = data.map(row => [
      row[2],  // 사업자번호 (D열 = index 2)
      row[3],  // 상호명 (E열 = index 3)
      row[4],  // 상세주소 (F열 = index 4)
      '',      // 전화번호 (없음)
      '',      // 가입된 플랫폼 (없음)
      row[0]   // 타입 (B열 = index 0)
    ]);

    if (mappedData.length > 0) {
      mergedSheet.getRange(currentRow, 1, mappedData.length, 6)
        .setValues(mappedData);
      currentRow += mappedData.length;
    }

    Logger.log(`${regionName}: ${mappedData.length}개 행 추가 완료`);
  });

  // 열 너비 자동 조정
  for (let i = 1; i <= 6; i++) {
    mergedSheet.autoResizeColumn(i);
  }

  Logger.log(`통합 완료: 총 ${currentRow - 2}개 행`);
  SpreadsheetApp.getUi().alert(`통합 완료!\n총 ${currentRow - 2}개의 데이터가 통합되었습니다.`);
}
