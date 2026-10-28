# 사무 자동화: PDF 를 엑셀 장부로 (office-automation-kr)

> [!IMPORTANT]
> **채팅 탭·Cowork 가 아니라 Code 탭(Local, 내 PC)에서만 씁니다.**
> 채팅 탭에 PDF 를 올리거나 Cowork 에 맡기면 회사 문서가 PC 밖으로 나갑니다. 이유는 [01장](01_구조_이해.md).

## 시작: PowerShell 에 두 줄

시작 메뉴에 `powershell` 을 쳐서 **Windows PowerShell** 을 엽니다. 상자 오른쪽 위 복사 단추로 복사해 붙여 넣고 Enter.

**① 기본 설치** (Claude Desktop · Node.js · uv·uvx) → 끝나면 Claude 에 로그인. 자세히는 [02장](02_기본_설치.md).

```powershell
[Net.ServicePointManager]::SecurityProtocol='Tls12'; irm https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/setup-base.ps1 | iex
```

**② 프로그램 설치** (사무 자동화 프로그램 · 작업 폴더 `C:\사무자동화` · Claude 기본 약속). 업데이트도 이 줄. 자세히는 [03장](03_프로그램_설치.md).

```powershell
[Net.ServicePointManager]::SecurityProtocol='Tls12'; irm https://raw.githubusercontent.com/dgkim3333-hash/office-automation-kr/main/setup-program.ps1 | iex
```

## 무엇을 해 주나

**Invoice · DM(Debit Memo) PDF** 를 `C:\사무자동화\받은문서` 에 넣고 `run.bat` 을 두 번 누르면

- `누적장부.xlsx` 에 한 줄씩 쌓이고
- PDF 는 `정리문서\고객명\2026-10\` 으로 정리되고
- 못 읽은 문서는 `확인필요\` 에 모입니다.

매일 처리에는 Claude 가 끼지 않습니다. 프로그램은 이 PC 안에서만 돌고 인터넷으로 아무것도 보내지 않습니다.

## 필요한 것

Windows 10 · 11, Claude **Pro 이상** 요금제, 프로그램 설치 권한.

## 읽는 순서

**파란 글씨를 순서대로** 누르세요. 각 장 맨 아래에 「다음 →」 이 있습니다.

| 장 | 무엇을 | 시간 |
|---|---|---|
| **[01](01_구조_이해.md)** | **[구조 이해](01_구조_이해.md)** (Code 와 Cowork, 무엇이 PC 밖으로 가는가) | 5분 |
| **[02](02_기본_설치.md)** | **[기본 설치](02_기본_설치.md)** (Claude Desktop · Node.js · uv, 자동) | 20분 |
| **[03](03_프로그램_설치.md)** | **[프로그램 설치](03_프로그램_설치.md)** (프로그램 · 작업 폴더 · 기본 약속, 자동) | 5분 |
| **[04](04_첫_실행_10분.md)** | **[첫 실행 10분](04_첫_실행_10분.md)** (가짜 샘플로 체험) | 10분 |
| **[05](05_내_양식_가르치기.md)** | **[내 양식 가르치기](05_내_양식_가르치기.md)** | 양식 하나당 20분 |
| **[06](06_매일_쓰기.md)** | **[매일 쓰기](06_매일_쓰기.md)** | 5분 |
| [07](07_프로젝트와_예정됨.md) | [프로젝트와 예정됨](07_프로젝트와_예정됨.md) (선택) | 10분 |
| [08](08_선택_Outlook_자동수집.md) | [Outlook 자동 수집](08_선택_Outlook_자동수집.md) (선택) | 15분 |
| [09](09_선택_엑셀PO변환.md) | [엑셀 PO 변환](09_선택_엑셀PO변환.md) (2차 예정) | — |
| [99](99_막혔을_때.md) | [막혔을 때](99_막혔을_때.md) | 필요할 때 |

> [!NOTE]
> 이 저장소에는 회사 정보가 없습니다. 샘플은 전부 가짜입니다. 여러분 규칙·장부·PDF 는 **여러분 PC 의 작업 폴더에만** 생깁니다.

## 폴더 안내

| 이름 | 무엇 |
|---|---|
| `setup-base.ps1` · `setup-program.ps1` | 02장 기본 설치 · 03장 프로그램 설치 (둘 다 `install.ps1` 을 부름) |
| `claude\전역_CLAUDE.md` | Claude 기본 약속. 설치 때 `C:\Users\내이름\.claude\CLAUDE.md` 로 복사(이미 있으면 건드리지 않음) |
| `run.bat` | 받은문서 처리 (매일) |
| `설치.bat` | 작업 폴더를 바꿀 때 |
| `수집.bat` | (선택) Outlook 첨부 자동 저장, 08장 |
| `samples\` | 가짜 샘플 (04장) |
| `rules\규칙.예시.yaml` | 가짜 양식용 규칙 |
| `app\` | 프로그램 본체 (고치지 않음) |
| `CLAUDE.md` | Code 탭의 Claude 가 지킬 약속 |
