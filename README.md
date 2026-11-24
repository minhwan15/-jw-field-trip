# 정왕고 교외체험학습 작성기

학교 원본 양식(한글에서 PDF로 저장한 서식 1·2·활동사진 별지)을 배경 그대로 쓰고, 빈칸 위치에만 입력값·서명·사진을 얹어 PDF로 출력합니다.

- 파일: `index.html`(작성기), `p1.png`~`p3.png`(원본 양식 배경, 200dpi), `vercel.json`
- 서버·API 없음(정적 파일). 서명·사진은 브라우저 밖으로 전송되지 않습니다.
- Vercel: Framework Preset `Other`, Build Command·Output Directory 비움.
- 양식이 바뀌면 배경 PNG와 `index.html` 안의 좌표(spec)를 함께 다시 만들어야 합니다.
