function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('도구')
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
    // 기존 시트가 있으면 내용 삭제
    mergedSheet.clear();
  } else {
    // 없으면 새로 생성
    mergedSheet = ss.insertSheet('전체통합');
  }

  // 헤더 작성
  const headers = ['분류', '사업자등록번호', '업체명', '상호명', '주소', '지역'];
  mergedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 헤더 스타일링
  mergedSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');

  let currentRow = 2; // 데이터 시작 행

  // 각 지역 시트 순회
  regionSheets.forEach(regionName => {
    const sheet = ss.getSheetByName(regionName);

    if (!sheet) {
      Logger.log(`시트를 찾을 수 없습니다: ${regionName}`);
      return;
    }

    const lastRow = sheet.getLastRow();

    // 헤더만 있거나 데이터가 없는 경우 스킵
    if (lastRow <= 1) {
      Logger.log(`데이터가 없습니다: ${regionName}`);
      return;
    }

    // 데이터 범위 가져오기 (2행부터 마지막 행까지, A~E열)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5);
    const data = dataRange.getValues();

    // 각 행에 지역명 추가
    const dataWithRegion = data.map(row => [...row, regionName]);

    // 통합 시트에 데이터 추가
    if (dataWithRegion.length > 0) {
      mergedSheet.getRange(currentRow, 1, dataWithRegion.length, 6)
        .setValues(dataWithRegion);
      currentRow += dataWithRegion.length;
    }

    Logger.log(`${regionName}: ${dataWithRegion.length}개 행 추가 완료`);
  });

  // 열 너비 자동 조정
  for (let i = 1; i <= 6; i++) {
    mergedSheet.autoResizeColumn(i);
  }

  Logger.log(`통합 완료: 총 ${currentRow - 2}개 행`);
  SpreadsheetApp.getUi().alert(`통합 완료!\n총 ${currentRow - 2}개의 데이터가 통합되었습니다.`);
}
