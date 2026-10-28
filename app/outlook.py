"""(선택) Outlook 데스크톱 앱의 '받은 편지함/자동화' 폴더 메일에서 PDF·엑셀 첨부를 받은문서/ 로 복사한다.

PC 안의 Outlook 앱과만 대화하며, 외부로 아무것도 보내지 않는다.
실행: 수집.bat (= uv run --extra outlook python -m app outlook)
"""
from pathlib import Path

from app.organize import free_name

FOLDER_NAME = "자동화"
DONE_CATEGORY = "자동화완료"
ATTACHMENT_SUFFIXES = (".pdf", ".xlsx", ".xls")  # .xls 는 받은 뒤 확인필요에서 다시 저장 안내
OL_FOLDER_INBOX = 6

NO_OUTLOOK = (
    "Outlook 데스크톱 앱에 연결하지 못했습니다.\n"
    "- Outlook(새 Outlook·웹 메일이 아닌 기존 데스크톱 앱)이 설치돼 있는지\n"
    "- 수집.bat 로 실행했는지 확인하세요(필요한 부품을 처음 실행 때 자동 설치). (08장)"
)


def _categories(mail):
    return [c.strip() for c in (mail.Categories or "").split(",") if c.strip()]


def collect(mails, inbox_dir) -> list[Path]:
    """메일 목록에서 PDF·엑셀 첨부를 저장하고, 처리한 메일에 범주 '자동화완료' 를 붙인다."""
    inbox_dir = Path(inbox_dir)
    inbox_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for mail in [_item(mails, i) for i in range(_count(mails))]:  # 범주를 바꾸기 전에 목록을 고정
        categories = _categories(mail)
        if DONE_CATEGORY in categories:
            continue
        for i in range(_count(mail.Attachments)):
            attachment = _item(mail.Attachments, i)
            if not attachment.FileName.lower().endswith(ATTACHMENT_SUFFIXES):
                continue
            dst = free_name(inbox_dir / attachment.FileName)
            attachment.SaveAsFile(str(dst))
            saved.append(dst)
        mail.Categories = ", ".join(categories + [DONE_CATEGORY])
        mail.Save()
    return saved


def _count(collection):
    return collection.Count if hasattr(collection, "Count") else len(collection)


def _item(collection, i):
    return collection.Item(i + 1) if hasattr(collection, "Item") else collection[i]  # COM 은 1부터


def find_subfolder(folders, name):
    """이름이 같은 하위 폴더. 없으면 LookupError."""
    for i in range(_count(folders)):
        folder = _item(folders, i)
        if folder.Name == name:
            return folder
    raise LookupError(name)


def _connect():
    import win32com.client

    namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    return find_subfolder(inbox.Folders, FOLDER_NAME).Items


def collect_from_outlook(inbox_dir, connect=_connect) -> list[Path]:
    try:
        mails = connect()
    except Exception:  # pywin32 없음, Outlook 없음, '자동화' 폴더 없음
        print(NO_OUTLOOK)
        print(f"- 받은 편지함 아래에 '{FOLDER_NAME}' 폴더가 있는지도 확인하세요.")
        return []
    saved = collect(mails, inbox_dir)
    print(f"Outlook 에서 첨부(PDF·엑셀) {len(saved)}개를 받은문서에 저장했습니다.")
    return saved
