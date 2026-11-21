# HTML in PPT — 슬라이드 안에서 실행되는 HTML (시험장)

이 폴더는 **분필팩의 `live-pptx` 스킬**을 만들고 실제 PowerPoint에서 검증하는 시험장이다.
완성본은 `nalssam/bunpil_pack`의 `src/plugins/bunpil-class/skills/live-pptx/`와 `site/live-html/`에 있다.

| 경로 | 내용 |
|---|---|
| `scripts/` | 변환기(`build_live_pptx.py`), 점검(`qa_live_pptx.py`), 등록(`register_addin.py`), 자동 시험(`test_live_pptx.py`) — 분필팩 스킬과 같은 파일 |
| `addin/` | PowerPoint 콘텐츠 추가 기능 페이지(`index.html`, 분필팩에서는 `app.html`)와 등록용 한 줄 스크립트 |
| `demo/deck.html` | 시험용 덱 5장 |
| `ci/` | 실제 PowerPoint 실행 시험: PowerPoint 설치, 로컬 https 서버(측정용 탐침 포함), COM 자동화, 대조 문서 생성 |
| `../.github/workflows/ppt-live-test.yml` | GitHub Actions Windows 러너에서 위 시험을 돌린다 |

## 실제 PowerPoint 시험이 확인하는 것

1. Windows에서 변환·점검·등록 스크립트가 그대로 동작한다.
2. PowerPoint가 변환된 pptx의 추가 기능을 개발자 등록(레지스트리)에서 찾아 활성화한다(Office 알림 로그).
3. 추가 기능 페이지가 PowerPoint의 WebView2 안에서 열리고, pptx에 저장된 HTML을 읽어 슬라이드별로 실행한다(탐침 로그).
4. 실제 마우스·키보드 입력이 슬라이드 안의 HTML에 전달되는지(편집 화면, 전체 화면 쇼, 창 모드 쇼).
5. 미등록·페이지 접속 불가·구형 Office·큰 HTML에서 무엇이 보이는지(화면 캡처).
6. `hosted` 작업: 분필팩 PR의 미리보기 사이트에서 배포용 스킬 zip과 추가 기능 페이지를 받아 끝까지 돌린다.

PowerPoint는 Office 배포 도구로 설치만 하고 정품 인증 없이(유예 상태) 잠깐 실행한다. 러너에는 그래픽 카드가 없다.

## 지금까지 잰 결과 (Windows 러너)

| PowerPoint | 편집 화면 | 전체 화면 쇼 | 창 모드 쇼 |
|---|---|---|---|
| Microsoft 365 최신 채널 16.0.20625 (한국어판) | 클릭·입력·키 넘기기 | 모두 동작 | 화면만 |
| Office LTSC 2024 16.0.17932 | 동작 | 화면만(클릭이 닿지 않음) | 화면만 |
| Office LTSC 2021 16.0.14334 | 동작 | 화면만 | 화면만 |
| Office 2019 볼륨판 16.0.10417 (IE11 엔진) | 미리보기 그림 + 안내 줄 | 같음 | 같음 |

Mac·웹용 PowerPoint·낱개 구매판 Office는 재지 못했다.

## 손으로 돌려 보기

```bash
pip install python-pptx
python scripts/build_live_pptx.py demo/deck.html -o out/deck.pptx --shots out/preview
python scripts/qa_live_pptx.py out/deck.pptx
cd scripts && python -m unittest test_live_pptx
```
