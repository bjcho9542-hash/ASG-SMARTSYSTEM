# 프로젝트 규칙

## 코드 변경 시 자동 실행 규칙

코드가 변경(수정, 편집, 추가, 삭제)될 때마다 반드시 아래 순서로 실행:

1. **Git 커밋 & 푸시**
   ```bash
   git add . && git commit -m "변경 내용" && git push
   ```

2. **Clasp 푸시** (Google Apps Script 동기화)
   ```bash
   ~/.npm-global/bin/clasp push
   ```

3. **웹 배포** (doGet/doPost 함수가 있는 경우에만)
   ```bash
   ~/.npm-global/bin/clasp deploy
   ```

## 프로젝트 정보

- 플랫폼: Google Apps Script
- 목적: 14개 시도 데이터 통합 자동화
- 주요 함수: `mergeAllRegions()`
