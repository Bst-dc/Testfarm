# 허브(포털) — 시험농장 + 스마트웹시스템 통합 진입점

나중에 GCP에 두 시스템을 함께 올릴 때, 창 하나를 열면 앱을 선택하고
들어가도록 만들기 위한 틀입니다. 지금 당장 쓰는 파일이 아니라 **준비된
템플릿**이며, 실제 도메인과 스마트웹시스템 정보가 정해지면 아래 순서로
채워 넣으면 됩니다.

## 구조

```
mydomain.com          → index.html (앱 선택 화면)
farm.mydomain.com     → 시험농장 (이 저장소, Streamlit)
web.mydomain.com      → 스마트웹시스템
```

서브도메인으로 나누는 이유: Streamlit은 `mydomain.com/farm` 같은
경로 방식보다 자기 도메인을 통째로 갖는 방식(서브도메인)에서 웹소켓·
정적 리소스 문제 없이 가장 깔끔하게 동작합니다.

## 사용 순서 (실제로 배포할 때)

1. 도메인을 준비하고 DNS에 A 레코드 3개를 GCP VM의 고정 IP로 연결합니다.
   - `mydomain.com`, `farm.mydomain.com`, `web.mydomain.com`
2. `index.html` 에서 두 곳의 `https://farm.mydomain.com`,
   `https://web.mydomain.com` 을 실제 도메인으로 바꾸고, 스마트웹시스템
   카드의 설명 문구도 채웁니다.
3. `Caddyfile.example` → `Caddyfile` 로 복사하고 `mydomain.com` 을
   실제 도메인으로 바꿉니다.
4. `docker-compose.example.yml` → `docker-compose.yml` 로 복사하고
   `smart-web` 서비스의 `build.context`/포트를 실제 스마트웹시스템에
   맞게 수정합니다. (스마트웹시스템이 Docker로 안 돌아가는 앱이면
   먼저 Dockerfile부터 만들어야 합니다 — 필요하면 그때 도와드립니다.)
5. VM에서 `docker compose up -d` 한 번이면 세 개(허브 페이지, 시험농장,
   스마트웹시스템)가 함께 뜨고, Caddy가 HTTPS 인증서까지 자동으로
   받아옵니다.

## 주의할 점

- 허브 페이지 자체는 로그인 화면이 아닙니다. 각 앱은 지금처럼 자체
  `ERP_PASSWORD` 등으로 따로 보호합니다.
- SQLite를 쓰는 이 시험농장 앱은 반드시 볼륨(`farm_data`)에 DB를 두어야
  컨테이너를 새로 올려도 데이터가 유지됩니다 (이미 `ERP_DB_DIR=/data`
  로 반영돼 있음).
- 스마트웹시스템도 상태를 저장한다면 마찬가지로 볼륨을 반드시 붙여야
  합니다.
