# SAJU 웹앱 — 로컬 Dropbox TXT 버전

기존 Next.js 사주 웹앱과 화면 디자인을 유지하면서, 상담 대기열과 채팅/입금 음성 호출을 **PC Dropbox 동기화 폴더의 TXT 3개**에서 직접 읽도록 변경한 버전입니다.

자세한 사용법: [LOCAL_TXT_SETUP.md](./LOCAL_TXT_SETUP.md)

## 실행

```bash
npm ci
npm run dev
```

Vercel 배포는 기존 방식 그대로 유지합니다. 데스크톱 Chrome/Edge에서 `로컬 TXT` 설정을 열고 폴더를 선택해야 로컬 파일을 읽습니다. Vercel에 Dropbox 경로를 등록할 필요는 없습니다.

## 확인

```bash
npm run verify:local-txt
npm run verify:filters
npm run build
```

소스 ZIP에는 `.env.local`, 실사용자 백업 JSON 등 개인 데이터가 포함되어 있지 않습니다. 현재 운영 중인 Vercel 환경변수는 기존 값을 유지하십시오.
