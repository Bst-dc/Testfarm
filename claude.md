# 프로젝트 컨텍스트 요약 (AI Handover Document)

이 문서는 사용자와의 작업이 마무리된 후 생성된 최신 상태 요약본입니다. Claude 등 다른 AI 에이전트가 이어서 작업할 때 이 문서를 기준으로 삼으세요.
(작성 기준: 2026-10-01, Supabase 전환 반영)

## 1. 시스템 개요
- **목적**: 한우 시험농장의 개체 입식, 사육 상태 변경, 비용 정산(월말 결산) 등을 관리하기 위한 ERP 시스템. 혼자 사용하며, 로컬 PC 실행과 외부(핸드폰/PC) 접속을 함께 지원하는 것을 목표로 함.
- **기술 스택**: Python, Streamlit, Pandas, DB는 Supabase(PostgreSQL) — `DATABASE_URL` 이 없으면 로컬 SQLite.
- **배포**: Streamlit Community Cloud(share.streamlit.io) + Supabase(서울 리전, 세션 풀러 5432). Secrets 에 `DATABASE_URL`(필요시 `ERP_PASSWORD`) 설정.
- **주요 파일**:
  - `erp_ui.py`: 전체 UI 렌더링, DB DDL(테이블 생성), CRUD 로직, 엑셀 파싱, 백업/복원, 정산 로직 등 메인 로직.
  - `erp_sunsan.db`, `erp_goa.db`: 농장별 독립된 SQLite 데이터베이스 파일.
  - `Dockerfile`, `fly.toml`, `requirements.txt`, `.dockerignore`: Fly.io 클라우드 배포용 구성.
  - `배포_가이드.md`: 로컬 실행 / Fly.io 배포 / 백업·복원 절차 안내.
- **저장소**: https://github.com/Bst-dc/Testfarm (Private 권장, 실데이터(.db, .xlsx)는 `.gitignore`로 제외됨)

## 2. DB 저장 위치 (중요 — 반드시 지킬 것)
- **Supabase 모드 (현재 운영)**: `DATABASE_URL` 이 secrets/환경변수에 있으면 `erp_ui.py` 의 `sqlite3` 이름이 `db_adapter` 로 바뀐다.
  - 농장 DB 파일명 → 스키마: `erp_sunsan.db` → `sunsan`, `erp_goa.db` → `goa`. 파일은 실제로 없으므로 **`os.path.exists(db_file)` 대신 `db_exists()`** 를 써야 한다(이걸 안 해서 클라우드에서 개체가 0으로 보였던 이력 있음).
  - `db_adapter.py`: SQL 변환(? → %s, strftime → substr, 대문자/작은따옴표 별칭 → 큰따옴표, sqlite_master/PRAGMA), 명령마다 SAVEPOINT(오류 난 명령만 되돌림 — SQLite처럼 '중복은 건너뛰고 계속'), 숫자 파라미터를 타입 미정 리터럴로 전송, NUMERIC → int/float, 연결 풀.
  - `db_schema.py`: `SQLITE_DDL`, `PG_DDL`(날짜는 TEXT, NUMERIC 자릿수 제한 없음, 매입 트리거는 PL/pgSQL), `DB_TABLES`(복사 순서).
  - 백업 = Supabase → SQLite 파일 내보내기, 복원 = SQLite 파일 → 스키마 통째 교체(`import_from_sqlite`, 한 트랜잭션).
  - 로컬 SQLite → Supabase 이관: `python migrate_to_supabase.py` (확인만) / `--run` / `--run --drop-extra`. 원본은 `(사용자 폴더)\시험농장DB` 의 farms.json 농장들(프로젝트 폴더의 옛 사본 아님).
- 아래는 SQLite 모드 설명:
- DB 파일은 **구글 드라이브 폴더에 두지 않는다.** 동기화 프로그램이 파일을 잠가 `database is locked` 오류와 파일 손상을 유발했던 이력이 있음.
- 저장 위치는 환경변수 `ERP_DB_DIR` 로 분리되어 있음:
  - 로컬: 기본값 `(사용자 폴더)\시험농장DB`
  - 클라우드(Fly.io): `/data` (영구 볼륨, `fly.toml`에 설정됨)
- 백업 폴더: `(ERP_DB_DIR)\backup` — 매일 첫 접속 시 자동 백업 + 사이드바에서 수동 백업/다운로드/복원 가능.

## 3. 최근 완료된 주요 작업 (최신순)
- **Supabase 전환 후 '개체가 하나도 없음' 문제 수정 (2026-10-01)**: 원인은 ① 이관 스크립트가 프로젝트 폴더의 옛 사본만 옮기고 고아농장은 누락, ② 앱의 `os.path.exists` 검사로 클라우드에서 DB가 '없음' 처리, ③ `to_sql` 이관으로 기본키·제약이 모두 빠짐. 어댑터/스키마/이관 스크립트를 새로 작성.
- **엑셀 일괄등록이 "이미 등록되어 건너뜀"으로만 끝나던 문제 수정**:
  - 원인은 중복이 아니라 **예전 DB의 CHECK 제약**이었음. 옛 스키마의 `cattle.feed_type`/`roughage_grade` 는 `('제한형','증량형')`/`('고급','저급')` 만 허용했는데, 일괄등록은 빈 값일 때 `'표준'` 을 기본값으로 넣어 모든 행이 `CHECK constraint failed` 로 거부됐다. 그런데 예전 코드가 모든 `sqlite3.IntegrityError` 를 '건너뜀'으로 집계해 중복인 것처럼 보였다.
  - `migrate_schema()` 에 `_migrate_cattle_default_check()` 추가: CHECK 제약은 `ALTER TABLE` 로 못 고치므로 `cattle` 표를 새로 만들어 데이터를 옮긴다. 이때 **`PRAGMA legacy_alter_table = ON` 필수** — 안 하면 RENAME 시 `disease_record`/`cattle_cost_log`/`cattle_item_usage_log` 의 `REFERENCES cattle(...)` 이 임시 표 이름으로 따라 바뀌어 참조가 끊긴다. 그 상태로 남은 DB를 되돌리는 `_repair_dangling_cattle_refs()` 도 함께 둠.
  - 일괄등록 오류는 `UNIQUE` 만 '건너뜀'으로 집계하고, 그 외 제약 위반은 사유별로 모아 결과 메시지에 함께 표시(`st.rerun()` 이 화면을 지워 per-row `st.error` 는 보이지 않았음).
- **결산 리포트(HTML/PDF) 생성 기능 구현** (이전까지 Next Steps 1순위였던 항목):
  - `generate_settlement_report()`: 정산월의 `cattle_cost_log`를 `cattle`/`testgroup_master`와 조인해, 개체별 변동비·고정비·합계와 시험군별 색상 배지가 들어간 인쇄용 HTML(Tailwind CDN, `report_2023-10.html` 목업 디자인 계승)을 생성. 사용자 입력값은 `html.escape`로 이스케이프.
  - `list_settled_months()`: 정산 완료된 연월 목록(최신순) 조회.
  - "🧾 결산 리포트" 탭 신설: 연월 선택 → 생성 → `st.iframe` 미리보기 + `.html` 다운로드.
- **정산 화면 안내 문구 정리 + SQL 파라미터 바인딩**:
  - 탭 라벨/버튼/안내문의 "1/n로 나누어" 표현을 실제 로직(사육일수 비례 일할계산)에 맞게 수정.
  - `settlement_month`/`target_month`/`target_code`/`item` 등을 f-string으로 SQL에 직접 끼워 넣던 8곳을 전부 `?` 파라미터 바인딩으로 전환.
- **월말 정산(원가 분배) 로직의 계산 오류 수정**:
  - 기존에는 `status = '사육'` 필터 때문에 정산월 도중에 출하/폐사한 개체가 정산에서 통째로 빠지고, 그 개체가 실제로 쓴 사료비·고정비 몫이 남은 개체에게 전가되는 문제가 있었음.
  - 상태와 무관하게 "해당 정산월에 사육 이력이 하루라도 겹치는 개체"를 모두 포함하도록 쿼리를 바꾸고, `calc_days()`에서 `closure_date`가 있으면 그 날짜까지만 사육일수로 계산하도록 수정.
- **리셋 시 가짜 데모 데이터 재삽입 문제 해결**:
  - `init_db()`에 하드코딩돼 있던 품목(ITEM1~3)·매입·월별 사용량·고정비 예시(2023-10 기준 가짜 금액) 데이터를 전부 제거. 리셋 후에는 완전히 빈 상태에서 시작.
- **DB 안정성**:
  - 연결 관리 헬퍼(`db_connect`/`close_stale_connections`) 도입. Streamlit의 `st.rerun()`이 예외로 동작해 연결 close가 스킵되던 문제를 해결.
  - `PRAGMA journal_mode=WAL`, `PRAGMA foreign_keys=ON`, `busy_timeout=30000` 적용.
- **백업/복원**:
  - 매일 자동 백업(`daily_backup`), 수동 백업 생성/다운로드, 백업 파일 업로드 복원(`restore_db`, 파일 무결성·스키마 검증 포함) 구현.
- **리셋 안전장치**:
  - 사이드바 리셋 버튼 → 관리자 비밀번호 확인 + 삭제될 데이터 건수 미리보기(`reset_preview`) → 실행 직전 자동 백업(`reset_dialog`).
- **외부 접속 보호**: 환경변수 `ERP_PASSWORD` 설정 시에만 비밀번호 화면 노출(`require_password`). 로컬 사용 시에는 그대로 통과.
- **클라우드 배포 준비**: Fly.io 기준 `Dockerfile`/`fly.toml` 작성 완료. **아직 `fly deploy` 실행 전 (미배포 상태).**
- **GitHub 연동**: 저장소 생성 및 최초 push 완료, 이후 커밋들도 반영됨.
- (이전 세션 작업) UI 레이아웃 최적화, 엑셀 일괄등록 신·구 양식 호환 및 중복 이표번호 자동 업데이트, 상태 변경/시험군 이동 UI, 금액 컬럼 콤마 포맷.

## 4. DB 스키마 구조 요약
- `testgroup_master`: 시험군 명칭, 시작/종료일.
- `cattle`: 개체 정보(이표번호, KPN, 시험군, 사료구분, 보험가입금액 등). 상태('사육', '출하', '폐사'), `closure_date`(출하/폐사일)는 정산 일할계산에 사용됨.
- `disease_record`: 질병 및 투약(약품1, 2, 3 등) 기록.
- `item_master`, `purchase`, `monthly_usage`, `monthly_fixedcost`, `cattle_cost_log`: 사료 및 비용 정산 로그 테이블. (리셋 시 더 이상 예시 데이터가 채워지지 않음 — 빈 상태로 시작)

## 5. 알려진 남은 이슈 / 향후 작업 (Next Steps)
- **클라우드 응답 속도**: 화면 한 번 그릴 때 쿼리 약 32회. Streamlit Cloud(미국) ↔ Supabase(서울) 왕복이 길어 클릭마다 수 초 걸릴 수 있음. 필요하면 조회 쿼리 묶기/캐시로 줄일 것.
- **farms.json 은 클라우드에서 휘발성**: '새 농장 추가'는 Streamlit Cloud 재시작 시 사라짐(기본 2개 농장만 유지). 영구 보관하려면 Supabase 표로 옮겨야 함.
- **Fly.io 실제 배포는 아직 미실행**. 가이드(`배포_가이드.md`)만 작성돼 있고, flyctl 설치·계정·결제 등록이 필요해 AI가 대신 실행할 수 없음. 사용자가 직접: `fly apps create` → `fly volumes create` → `fly secrets set ERP_PASSWORD` → `fly deploy`.
  - 클라우드로 옮길 경우, 기존 로컬 데이터는 백업 파일을 만들어 "백업 파일로 복원" 기능으로 옮겨야 함(자동 이관되지 않음).
- **결산 리포트 개선 여지**: 현재는 시험군당 색상 배지가 팔레트를 순환하는 방식이라 시험군이 7개 이상이면 색이 겹침(치명적이진 않음). PDF는 브라우저의 "인쇄하기"에 의존(자동 PDF 저장 기능은 없음).
- 그 외 알려진 버그는 없음. 새로 발견되는 문제는 이 항목에 추가할 것.
