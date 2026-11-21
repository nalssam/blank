# HTML in PPT — 슬라이드 안에서 실행되는 HTML

HTML 한 파일을, 슬라이드마다 그 HTML이 **실제로 동작하는** PowerPoint 파일로 바꿉니다.
입력창에 글자를 치고, 슬라이더를 움직이고, 퀴즈 보기를 누를 수 있습니다.

## 구조

| 파일 | 역할 |
|---|---|
| `demo/deck.html` | 예시 덱 (슬라이드 5장, 한 파일) |
| `build.py` | HTML → pptx 변환. HTML 원문을 pptx 안에 저장하고 미리보기 그림을 찍어 넣음 |
| `addin/` | **Live HTML** 콘텐츠 추가 기능 (pptx 안의 HTML을 꺼내 그리는 작은 웹 페이지 + manifest) |
| `install-windows.ps1` | 추가 기능을 내 PC PowerPoint에 등록 |
| `demo/live-demo.pptx` | 변환 결과 예시 |

HTML은 pptx 파일 안(`ppt/webextensions/`)에 들어 있어서, 덱마다 따로 웹에 올릴 필요가 없습니다.
웹에 올라가 있어야 하는 것은 `addin/` 폴더 하나뿐이고, 한 번만 하면 됩니다.

## 처음 한 번 (Windows)

1. **추가 기능 페이지 호스팅** — `addin/` 폴더가 https 주소로 열려야 합니다.
   이 저장소에서 GitHub Pages를 켜면(Settings → Pages → `main` 브랜치) 주소는
   `https://nalssam.github.io/blank/html-in-ppt/addin` 이 됩니다. 다른 곳에 올렸다면
   `python build.py <html> --host-url <주소>` 로 `addin/manifest.xml` 을 다시 만드세요.
2. **PC에 등록** — PowerShell에서 `.\install-windows.ps1` 실행 후 PowerPoint 재시작.
   (해제: `.\install-windows.ps1 -Remove`)

## 덱 만들기

```bash
npm install                # 저장소 루트에서 (playwright)
pip install python-pptx lxml
python build.py demo/deck.html -o demo/live-demo.pptx
```

HTML 규칙은 두 가지입니다.

- 슬라이드 하나 = `<section class="slide">` 하나 (또는 `--slides N` 으로 장수 지정)
- 페이지가 열릴 때 `window.__LIVE_SLIDE__` (1부터)를 읽어 그 슬라이드만 보여 주기.
  `demo/deck.html` 맨 아래 스크립트 첫 부분을 그대로 가져다 쓰면 됩니다.

## 수동으로 넣기 (변환 파일이 추가 기능을 못 찾을 때)

PowerPoint에서 **삽입 → 추가 기능 → 내 추가 기능 → 개발자 추가 기능 → Live HTML** 을 슬라이드에
놓고, 나타나는 화면에서 HTML 파일과 슬라이드 번호를 고르면 됩니다. 이 경로는 `build.py` 없이도 동작합니다.

## 활용 아이디어

- 숫자를 만지는 발표: 슬라이더로 가정을 바꾸면 차트가 바로 다시 그려짐
- 수업·연수 퀴즈: 보기를 눌러 정답 확인, 다음 문제로
- 진행 도구: 활동 타이머, 발표자 뽑기
- 제품 데모: 실제 화면 프로토타입을 슬라이드에서 바로 클릭
- 코드·수식 시연: 입력값을 바꿔 결과를 그 자리에서 보여 주기

## 알아 둘 점

- **실제 PowerPoint에서의 동작은 아직 확인하지 못했습니다.** 변환 결과의 파일 구조 검증과,
  PowerPoint를 흉내 낸 브라우저 환경에서의 추가 기능 동작(입력·슬라이더·퀴즈·타이머)까지만 확인했습니다.
  특히 pptx가 등록된 추가 기능을 가리키는 방식(`store="developer" storeType="Registry"`)은 실기 확인이 필요합니다.
  안 될 때는 위 "수동으로 넣기"로 한 장을 넣어 저장한 파일을 보내 주시면 그 형식에 맞춥니다.
- 추가 기능이 없는 PC·뷰어에서는 미리보기 그림이 보입니다.
- 슬라이드 쇼 중 HTML을 클릭하면 키보드가 HTML 쪽으로 넘어갑니다. 다음 슬라이드로 갈 때는 리모컨이나 슬라이드 바깥 클릭을 쓰세요.
- HTML은 격리된 프레임에서 실행되어 `localStorage`·쿠키를 쓸 수 없습니다. 외부 글꼴·CDN은 인터넷이 될 때만 불러옵니다.
- Mac은 등록 방법이 다릅니다 (`manifest.xml` 을 `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` 에 복사).
