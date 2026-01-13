// 바보

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ASG_도구')
    .addItem('데이터 통합(푸드)', 'mergeAllRegionsFood')
    .addItem('데이터 통합(논푸드)', 'mergeAllRegionsNonFood')
    .addToUi();
}

// 진행률 팝업 표시
function showProgress(title) {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
          .progress-container { width: 100%; background-color: #e0e0e0; border-radius: 10px; margin: 20px 0; }
          .progress-bar { width: 0%; height: 30px; background-color: #4a86e8; border-radius: 10px; transition: width 0.3s; }
          .status { font-size: 14px; color: #666; margin-top: 10px; }
          .percent { font-size: 24px; font-weight: bold; color: #4a86e8; }
        </style>
      </head>
      <body>
        <h3>${title}</h3>
        <div class="percent" id="percent">0%</div>
        <div class="progress-container">
          <div class="progress-bar" id="progressBar"></div>
        </div>
        <div class="status" id="status">준비 중...</div>
        <script>
          function updateProgress(percent, status) {
            document.getElementById('percent').textContent = percent + '%';
            document.getElementById('progressBar').style.width = percent + '%';
            document.getElementById('status').textContent = status;
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(400)
  .setHeight(200);

  SpreadsheetApp.getUi().showModelessDialog(html, title);
  return html;
}

// 푸드 통합
function mergeAllRegionsFood() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const regionSheets = [
    '충청도', '경기도', '경상도', '전라도',
    '부산시', '대구시', '대전시', '광주시',
    '울산시', '인천시', '강원도', '제주도',
    '서울시', '세종시'
  ];

  // 진행률 시트 생성 (팝업과 통신용)
  let progressSheet = ss.getSheetByName('_진행률');
  if (!progressSheet) {
    progressSheet = ss.insertSheet('_진행률');
  }
  progressSheet.clear();
  progressSheet.getRange('A1').setValue(0);
  progressSheet.getRange('B1').setValue('시작...');

  // 팝업 표시
  showProgressDialog('푸드 데이터 통합 중...');

  let mergedSheet = ss.getSheetByName('전체통합');

  if (mergedSheet) {
    mergedSheet.clear();
  } else {
    mergedSheet = ss.insertSheet('전체통합');
  }

  const headers = ['사업자번호', '상호명', '상세주소', '전화번호', '가입된 플랫폼', '타입', '카테고리'];
  mergedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  mergedSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');

  let currentRow = 2;
  const totalRegions = regionSheets.length;

  regionSheets.forEach((regionName, index) => {
    // 진행률 업데이트
    const percent = Math.round(((index) / totalRegions) * 100);
    progressSheet.getRange('A1').setValue(percent);
    progressSheet.getRange('B1').setValue(`${regionName} 처리 중...`);
    SpreadsheetApp.flush();

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

    const dataRange = sheet.getRange(2, 2, lastRow - 1, 5);
    const data = dataRange.getValues();

    const mappedData = data.map(row => [
      row[2],
      row[3],
      row[4],
      '',
      '',
      row[0],
      ''
    ]);

    if (mappedData.length > 0) {
      mergedSheet.getRange(currentRow, 1, mappedData.length, 7)
        .setValues(mappedData);
      currentRow += mappedData.length;
    }

    Logger.log(`${regionName}: ${mappedData.length}개 행 추가 완료`);
  });

  // 완료
  progressSheet.getRange('A1').setValue(100);
  progressSheet.getRange('B1').setValue('완료!');
  SpreadsheetApp.flush();

  for (let i = 1; i <= 7; i++) {
    mergedSheet.autoResizeColumn(i);
  }

  // 진행률 시트 삭제
  ss.deleteSheet(progressSheet);

  Logger.log(`푸드 통합 완료: 총 ${currentRow - 2}개 행`);
  ui.alert(`푸드 통합 완료!\n총 ${currentRow - 2}개의 데이터가 통합되었습니다.`);
}

// 논푸드 통합
function mergeAllRegionsNonFood() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const regionSheets = [
    '충청도', '경기도', '경상도', '전라도',
    '부산시', '대구시', '대전시', '광주시',
    '울산시', '인천시', '강원도', '제주도',
    '서울시', '세종시'
  ];

  // 진행률 시트 생성
  let progressSheet = ss.getSheetByName('_진행률');
  if (!progressSheet) {
    progressSheet = ss.insertSheet('_진행률');
  }
  progressSheet.clear();
  progressSheet.getRange('A1').setValue(0);
  progressSheet.getRange('B1').setValue('시작...');

  // 팝업 표시
  showProgressDialog('논푸드 데이터 통합 중...');

  let mergedSheet = ss.getSheetByName('전체통합');

  if (mergedSheet) {
    mergedSheet.clear();
  } else {
    mergedSheet = ss.insertSheet('전체통합');
  }

  const headers = ['사업자번호', '상호명', '상세주소', '전화번호', '가입된 플랫폼', '타입', '카테고리'];
  mergedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  mergedSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');

  let currentRow = 2;
  const totalRegions = regionSheets.length;

  regionSheets.forEach((regionName, index) => {
    // 진행률 업데이트
    const percent = Math.round(((index) / totalRegions) * 100);
    progressSheet.getRange('A1').setValue(percent);
    progressSheet.getRange('B1').setValue(`${regionName} 처리 중...`);
    SpreadsheetApp.flush();

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

    const dataRange = sheet.getRange(2, 2, lastRow - 1, 7);
    const data = dataRange.getValues();

    const mappedData = data.map(row => [
      row[1],
      row[2],
      row[5],
      '',
      '',
      row[0],
      row[6]
    ]);

    if (mappedData.length > 0) {
      mergedSheet.getRange(currentRow, 1, mappedData.length, 7)
        .setValues(mappedData);
      currentRow += mappedData.length;
    }

    Logger.log(`${regionName}: ${mappedData.length}개 행 추가 완료`);
  });

  // 완료
  progressSheet.getRange('A1').setValue(100);
  progressSheet.getRange('B1').setValue('완료!');
  SpreadsheetApp.flush();

  for (let i = 1; i <= 7; i++) {
    mergedSheet.autoResizeColumn(i);
  }

  // 진행률 시트 삭제
  ss.deleteSheet(progressSheet);

  Logger.log(`논푸드 통합 완료: 총 ${currentRow - 2}개 행`);
  ui.alert(`논푸드 통합 완료!\n총 ${currentRow - 2}개의 데이터가 통합되었습니다.`);
}

// 진행률 다이얼로그
function showProgressDialog(title) {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
          .progress-container { width: 100%; background-color: #e0e0e0; border-radius: 10px; margin: 20px 0; height: 30px; }
          .progress-bar { width: 0%; height: 30px; background-color: #4a86e8; border-radius: 10px; transition: width 0.5s; }
          .status { font-size: 14px; color: #666; margin-top: 10px; }
          .percent { font-size: 28px; font-weight: bold; color: #4a86e8; }
        </style>
        <script>
          function poll() {
            google.script.run
              .withSuccessHandler(function(result) {
                if (result) {
                  document.getElementById('percent').textContent = result.percent + '%';
                  document.getElementById('progressBar').style.width = result.percent + '%';
                  document.getElementById('status').textContent = result.status;

                  if (result.percent < 100) {
                    setTimeout(poll, 500);
                  } else {
                    document.getElementById('status').textContent = '완료! 이 창을 닫아주세요.';
                    document.getElementById('progressBar').style.backgroundColor = '#34a853';
                  }
                } else {
                  setTimeout(poll, 500);
                }
              })
              .withFailureHandler(function() {
                setTimeout(poll, 1000);
              })
              .getProgress();
          }
          poll();
        </script>
      </head>
      <body>
        <h3>${title}</h3>
        <div class="percent" id="percent">0%</div>
        <div class="progress-container">
          <div class="progress-bar" id="progressBar"></div>
        </div>
        <div class="status" id="status">준비 중...</div>
      </body>
    </html>
  `)
  .setWidth(400)
  .setHeight(220);

  SpreadsheetApp.getUi().showModelessDialog(html, title);
}

// 진행률 조회 (팝업에서 호출)
function getProgress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheet = ss.getSheetByName('_진행률');

  if (!progressSheet) {
    return { percent: 0, status: '준비 중...' };
  }

  const percent = progressSheet.getRange('A1').getValue() || 0;
  const status = progressSheet.getRange('B1').getValue() || '처리 중...';

  return { percent: percent, status: status };
}
