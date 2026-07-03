# Trade Remedy Monitoring System

Google Apps Script와 Gemini API를 활용한 무역구제 동향 자동 모니터링 시스템입니다.

## 발송 이메일 샘플

<img width="1287" height="802" alt="image" src="https://github.com/user-attachments/assets/1cdb485e-9abd-4c87-a0ae-f998f468c2ec" />
<img width="1276" height="651" alt="image" src="https://github.com/user-attachments/assets/ddb9e9b9-4229-436a-866f-73f274c68565" />

> 매일 오전 6시 자동 발송되는 무역 동향 AI 요약 이메일입니다.

## 배경

미국 세관청(CBP), 무역위원회(KTC), 인도 DGTR 등 주요 기관의 무역구제 동향을
매일 수동으로 확인하던 작업을 자동화했습니다.
매일 오전 6시에 자동으로 실행되어 AI 요약 리포트를 이메일로 발송합니다.

## 주요 기능

- 주요 무역기관 및 인사 SNS 동향 자동 수집
- Gemini API를 활용한 AI 기반 요약 및 분류
- 매일 오전 6시 자동 실행 (Google Apps Script 트리거)
- 결과를 이메일로 자동 발송

## 기술 스택

- Google Apps Script
- Gemini API
- Naver News API
- Google Sheets

## 시스템 구조

```
trade-monitoring-system/
├── gas/
│   └── monitor.gs        # 수집·AI요약·이메일 발송 통합 스크립트
└── README.md
```

## 설정 방법

1. Google Apps Script 새 프로젝트 생성
2. Script Properties에 아래 API 키 등록
   - `GEMINI_API_KEY`
   - `NAVER_CLIENT_ID`
   - `NAVER_CLIENT_SECRET`
3. `monitor.gs` 코드 붙여넣기
4. 트리거 설정: 매일 오전 6시 실행
