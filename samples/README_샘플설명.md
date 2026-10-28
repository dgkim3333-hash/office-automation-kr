# samples — 예제 파일 설명

작성 2026-10-28(주소 교체 2차). 모든 회사·고객·금액·번호·주소·도시명(Cheongun-si·Hanbit-si·Haneul-si·Prairie Bend)은 **가짜**이며 실제 사업장과 무관하다. 실제 양식이 없어서 반도체·데이터센터 쿨링 장비 거래에 흔한 항목으로 구성했다.
Code 세션은 이 폴더를 저장소 `samples/` 로 그대로 옮겨 쓴다 (계획 1.1·1.2 는 "생성"이 아니라 "이 파일을 가져와 검증"으로 바뀐다).

## 등장 회사 (가짜)

| 역할 | 이름 | 비고 |
|---|---|---|
| 공급사 (미국 본사) | Vertex Thermal Systems, Inc. (Prairie Bend, TX — 가상 도시) | 한국지사: Vertex Thermal Systems Korea (Branch) |
| 고객 1 | ACME Semiconductor Korea Co., Ltd. | 고객코드 C-KR-0417, 영문 PO 양식 |
| 고객 2 | Hanil DataCenter Co., Ltd. | 고객코드 C-KR-0522, 한국어 PO 양식 |

거래 2건: ACME `PO-ACME-2026-0417`(CDU 2대 외, 항공) · HANIL `HDC-PUR-26-0908`(칠러 1대 외, 해상).

## 폴더별 파일

### 01_고객오더_엑셀 — 고객이 보내온 발주서 (고객마다 양식이 다름)
| 파일 | 특징 |
|---|---|
| 고객오더_ACME_PO-ACME-2026-0417.xlsx | 영문. 머리 블록(A3:B11) + 품목표(13행부터). 열: No · Supplier Part No · Description · Qty · UoM · Unit Price · Amount · Need-by Date |
| 고객오더_HANIL_HDC-PUR-26-0908.xlsx | 한국어. 머리 블록 2열 배치(A4:D8) + 품목표(10행부터). 열: 순번 · 품번 · 품명 · 수량 · 단위 · 단가 · 금액 · 납기 |

→ 두 고객의 양식이 다르다는 점이 핵심. "고객별 엑셀 규칙"이 필요하다.

### 02_회사양식_변환 — 본사 시스템 업로드용 SO 입력 양식 (열 고정)
| 파일 | 특징 |
|---|---|
| 회사양식_SO입력_PO-ACME-2026-0417.xlsx | 시트 `SO_Upload`: 한 행 = 품목 1줄, 22개 열 고정 (Order Date · Customer Code · Customer PO No · Line No(10단위) · Part No · Qty · Unit Price · Ext Price · Requested Ship Date · Incoterms · Payment Terms … · 맨 뒤 본사용 HQ Unit Price · HQ Amount · Markup %). 시트 `열설명` 포함. 본사 단가는 가짜 마크업표(app/markup.py 예시: 0 · 2 · 3%, VTS-FLT-50 만 소수 2자리 반올림)로 계산 |
| 회사양식_SO입력_HDC-PUR-26-0908.xlsx | 같은 양식 |

→ 01 의 고객 양식을 이 양식으로 옮기는 것이 "발주 → 엑셀 정리 → 시스템 입력" 업무의 자동화 대상이다. Customer Code 는 고객마스터(가짜: 02 안 `열설명` 참조)에서 가져온다.

### 03_Invoice_DM_PDF — 본사가 발행한 문서 (PDF, 텍스트 추출 가능)
| 파일 | 종류 | 포인트 |
|---|---|---|
| INV-2026-10231.pdf | Invoice (ACME) | 품목 4줄, Freight 있음, 항공 AWB |
| INV-2026-10278.pdf | Invoice (HANIL) | 품목 3줄, 해상 — Vessel/Voyage 줄이 하나 더 있음, 추적번호에 컨테이너 번호 포함 |
| DM-2026-0092.pdf | Debit Memo (ACME) | 긴급 항공 운임 추가분 청구. **본문에 "Reference Invoice" 가 있어 판별 순서 함정** |
| DM-2026-0097.pdf | Debit Memo (HANIL) | 단가 인상분 청구 (20개 × $15.00) |
| Letter_warranty_notice.pdf | 판별 불가 | 안내 서신. → 확인필요/ 로 가야 정상 |
| Scan_signed_packing_list.pdf | 텍스트 0자 | 스캔본 흉내. → 확인필요/ "텍스트 없음" |

DM = **Debit Memo**(운임·가격 차액 청구서)로 확정 (2026-10-28).

### 04_장부_엑셀
| 파일 | 용도 |
|---|---|
| 누적장부_빈양식.xlsx | `run.bat` 가 처음 만들 장부의 모양. 시트 Invoice · PO · DM · 품목 · 처리로그 |
| 누적장부_기대결과.xlsx | 03 의 PDF 6개 + 01 의 엑셀 2개를 처리한 뒤 **나와야 하는** 상태. test_e2e 의 기대값 (1차에서는 PO 시트·PO 품목·엑셀 2개 로그 제외, 엑셀 PO 는 2차) |
| 추출결과_단건_INV-2026-10231.xlsx | `check` 명령이 화면에 찍는 것과 같은 내용을 엑셀로 (필드/값/비고 + 품목) |

### 규칙.예시.yaml (저장소 `rules\규칙.예시.yaml` 로 옮김)
03 의 PDF 에 맞춘 추출 규칙. 2026-10-28 검증: Invoice 2건·DM 2건 모든 필드·표 추출 성공, 서신·스캔본은 판별 불가(None).

설계 5절과 달라진 점 두 가지 (설계에 반영 필요):
1. 표 추출에 `행패턴`(정규식) 방식을 추가했다. 품명에 공백이 많아 "열 개수로 쪼개기"로는 안 된다.
2. 문서종류 검사 순서가 의미를 가진다 — DM 을 Invoice 보다 먼저. Invoice 키워드는 `Invoice No` 로 좁힌다.

## 설계에 추가해야 할 것 (브레인스토밍 범위 밖에서 드러남)

- **고객 PO 가 엑셀로 온다.** 지금 설계는 PDF 만 다룬다. "고객별 엑셀 규칙"(머리 블록 셀 위치 + 품목표 시작행·열 매핑)과 "회사 SO 양식 변환" 모듈이 필요하다. 1차에 넣을지, 2차로 미룰지 결정 필요.
- 02 의 Customer Code 처럼 **고객마스터**(고객명 → 코드 · 결제조건 · 인도조건 기본값)가 있어야 변환이 완성된다. 작은 엑셀 한 장이면 된다.

## 다시 만들기
`python make_samples.py samples` (reportlab · openpyxl 필요). 금액·날짜를 바꾸려면 스크립트 상단 `ORDERS` 만 고친다.
