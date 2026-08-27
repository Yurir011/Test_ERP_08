# 리팩토링 구조

## 파일 의존 관계

```
        common.py  (공용 헬퍼 + 베이스 클래스, 다른 모듈을 import 안 함)
           ↑   ↑   ↑
   quotation.py  delivery_note.py  doc.py   (서로 import 안 함, 독립적)
           ↑   ↑   ↑
        main.py  (HomeWindow + 앱 진입점, 위 4개를 조합)
```

핵심 원칙: **문서 타입 세 파일(quotation/delivery_note/doc.py)은 서로를 절대 참조하지 않습니다.** 공통으로 필요한 건 전부 `common.py`로 올려서, 한 문서 타입을 고치다가 다른 타입이 실수로 깨지는 경로를 원천 차단했습니다. (예: 이전엔 `DELIVERY_NOTE_BOTTOM_CLUSTER_WIDGET_NAMES`를 만들 때 `quotation.py`에서 상수를 가져와야 했는데, 그 상수를 `common.py`의 `COMPACT_BOTTOM_CLUSTER_WIDGET_NAMES`로 옮겨서 이 의존을 끊었습니다.)

## common.py — 공용 인프라

- **경로/상수**: `BASE_DIR`, `UI_DIR`, `CLICK_SOUND`, 컴팩트 폼 스타일시트, 품목행 레이아웃 상수
- **엑셀/PDF 내보내기**: `export_pdf`(Excel COM으로 PDF 변환), `export_widget_pdf`/`print_widget`(화면 캡처 → PDF/프린터, `_paint_widget_onto_printer`로 로직 공유), `export_document_files`(저장 3종 세트: xlsx+pdf+screen.pdf), `append_to_list_xlsx`(품목별 이력 로그 — 견적서/거래명세서가 공유)
- **창 관리**: `move_centered_on`(멀티모니터에서 화면 밖으로 안 나가게 clamp), `make_window_scrollable`(풀페이지 창에 스크롤 달기)
- **폼 로직**: `style_item_form`/`connect_amount_calculations`/`build_entry_chain`/`save_filled_document` — 견적서·거래명세서 풀페이지 창이 거의 그대로 재사용하는 저장 파이프라인
- **`CompactItemFormWindow`**: 컴팩트 입력폼(견적서/거래명세서 공통 골격)의 베이스 클래스. `max_item_rows`를 서브클래스가 오버라이드하는 구조라, 두 폼의 품목행 개수가 서로 달라져도(현재는 둘 다 6개, 원래는 8개였던 걸 각각 다르게 줄임) 서로 안 엮입니다.

## quotation.py / delivery_note.py — 각자 독립

둘 다 같은 패턴: `Quotation2Window`/`DeliveryNote2Window`(풀페이지, `common`의 헬퍼 조합해서 `__init__` 구성) + `QuotationWindow`/`DeliveryNoteWindow`(`CompactItemFormWindow` 상속, `ui_path`/`number_prefix`/`max_item_rows` 등만 클래스 속성으로 지정).

`delivery_note.py`에는 "문서참조"(저장된 xlsx/pdf 다시 불러오기) 파서들(`parse_saved_xlsx`, `parse_saved_pdf`, `_cluster_pdf_words*`)도 같이 있습니다 — 이 기능은 거래명세서에만 있어서.

## doc.py — 나머지 둘과 다른 이유

`DocWindow`/`Doc2Window`는 `CompactItemFormWindow`를 상속하지 않습니다. 공문서는 품목행 add/remove 개념이 없고(고정 필드 몇 개 + 자유 서식 본문), 대신 4단계 결재란(`DOC_APPROVAL_STAGES`, `_cycle_approval`)이라는 이 문서만의 로직이 있어서 원래부터 별도 구조였고, 분리 과정에서도 그대로 유지했습니다.

## main.py — 96줄로 축소

`HomeWindow`(홈 화면 버튼 연결), `_show_uncaught_exception`(전역 에러 다이얼로그), `main()` 진입점만 남았습니다. `run.bat`은 그대로 `python -u main.py`라 실행 방법은 안 바뀌었습니다.
