PRD: 구글 스프레드시트 14개 시도 데이터 통합 자동화
1. 프로젝트 개요
14개 시도군별로 분리된 시트의 데이터를 하나의 통합 시트로 자동 병합하는 Apps Script 작성
2. 기술 스펙

플랫폼: Google Apps Script
대상: Google Spreadsheet
언어: JavaScript

3. 기능 요구사항
3.1 통합 대상 시트

총 14개 시트 (시도군별 분리)
시트명 목록:

충청도
경기도
경상도
전라도
부산시
대구시
대전시
광주시
울산시
인천시
강원도
제주도
서울시
세종시



3.2 데이터 구조
원본 시트 열 구조:

A열: 분류 (Type H-다, 공주시 등)
B열: 사업자등록번호
C열: 업체명
D열: 상호명
E열: 주소

통합 시트 열 구조:

A열: 분류
B열: 사업자등록번호
C열: 업체명
D열: 상호명
E열: 주소
F열: 지역 (새로 추가)

3.3 통합 로직

새 시트 생성: "전체통합" (이미 존재하면 기존 데이터 삭제 후 재작성)
헤더 행 작성
14개 시트를 순회하며:

각 시트의 2행부터 마지막 행까지 데이터 추출 (1행은 헤더 제외)
각 행의 마지막 열(F열)에 해당 시트명(지역명) 추가
통합 시트에 순차적으로 추가



3.4 예외 처리

빈 시트는 스킵
헤더만 있고 데이터가 없는 시트는 스킵
존재하지 않는 시트명은 로그 출력 후 스킵

4. 실행 방법

Apps Script 에디터에서 함수 실행
함수명: mergeAllRegions()

5. 코드 구현
javascriptfunction mergeAllRegions() {
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
6. 사용 방법

구글 스프레드시트에서 확장 프로그램 > Apps Script 메뉴 선택
코드 에디터에 위 코드 붙여넣기
저장 (Ctrl+S 또는 Cmd+S)
함수 선택: mergeAllRegions 선택
실행 버튼 클릭 (▶️)
첫 실행 시 권한 승인 필요
완료 후 "전체통합" 시트 확인

7. 주요 기능

✅ 14개 시트 자동 병합
✅ 지역명 자동 추가 (F열)
✅ 헤더 스타일링
✅ 열 너비 자동 조정
✅ 진행 상황 로그 출력
✅ 완료 알림 팝업

8. 검증 포인트

 14개 시트 모두 데이터 수집 확인
 지역명 정확히 매칭 확인
 데이터 누락 없는지 확인
 헤더 행 제외 여부 확인