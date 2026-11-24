# 정왕고 체험학습 작성기

학교 원본 양식(서식 1·2·활동사진 별지)을 배경 그대로 쓰고 빈칸에만 값을 얹어 PDF로 출력합니다.
담임 검토 기능은 구글 시트 + Apps Script를 서버로 씁니다(Vercel은 정적 파일만 제공).

## 파일
- `index.html` 학생 작성기 / `teacher.html` 담임 검토 화면
- `config.js` 검토 서버 주소(Apps Script 웹 앱 URL)
- `p1~p3.png` 원본 양식 배경, `schedule.png` 학사일정표, `og.png` 링크 미리보기
- `gas/Code.gs` 구글 시트에 붙여넣을 Apps Script 코드(배포되지 않음)

## 담임 검토 서버 설치(1회)
1. 새 구글 시트 → 확장 프로그램 → Apps Script → `gas/Code.gs` 내용 붙여넣기 → 저장
2. 함수 `setup` 실행 → 권한 허용 ('신청'·'설정' 탭과 드라이브 PDF 폴더 생성)
3. 배포 → 새 배포 → 웹 앱 (실행: 나, 액세스: 모든 사용자) → URL 복사
4. `config.js`의 `window.JW_API`에 URL 입력
5. '설정' 탭에 반별 담임 비밀번호 입력
