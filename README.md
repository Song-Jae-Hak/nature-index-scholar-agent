# Nature Index → PNU Scholar 수집 에이전트

Nature가 관리하는 **Nature Index 저널 목록**(현재 178종)을 수집하고, 각 저널을
**부산대 PNU Scholar**(`scholar.pusan.ac.kr`)에서 찾아
① ISSN·eISSN, ② Scholar 등록 여부, ③ `Registered DB`에 **Nature Index** 태그가
있는지(Y/N), ④ `journal_id` 및 상세 URL 을 정리해 **엑셀 한 파일**로 만드는 자동화 스크립트입니다.

---

## 1. 한 줄 실행 (재실행도 동일)

```powershell
# 이 폴더(nature_index_agent)에서
powershell -ExecutionPolicy Bypass -File run.ps1
```
또는 탐색기에서 **`run.cmd` 더블클릭**.

- 실행하면 오늘 날짜 폴더(예: `..\natureindex\20260726\`)가 자동 생성되고 결과물이 그 안에 저장됩니다.
- 날짜/폴더를 지정하려면: `run.ps1 -Date 20260726`
- 외부 데이터를 새로 받아 캐시를 재생성하려면: `run.cmd -FreshCache`

실행 소요: 캐시가 있으면 1~2분, 완전히 처음이면 약 8~15분(외부 API·Scholar 조회량 때문).

---

## 2. 처리 단계

1. **FAQ 수집** — `agent-browser`(실제 Chrome)로
   `https://www.nature.com/nature-index/faq` 를 열고 **Q12(“What publications are in the Nature Index?”)**
   의 저널 목록(178종)을 파싱합니다.
   > 이 페이지는 `idp.nature.com` 인증 쿠키 리다이렉트가 걸려 단순 HTTP(WebFetch/requests)로는 열리지 않습니다.
   > 그래서 실제 브라우저인 agent-browser 를 사용합니다. (검색엔진은 캡차 때문에 사용하지 않음)
2. **ISSN/eISSN 수집** — 저널명으로 **Crossref**, **OpenAlex**(가능할 때) API를 조회해
   독립적인 ISSN 후보를 모읍니다.
3. **PNU Scholar 매칭** — Scholar 의 공개 REST API
   (`/wp-json/rm/v1/journals?ks=…`) 를 ISSN·저널명으로 검색하여 같은 저널을 찾고,
   `registered_db` 에 **Nature Index** 가 있으면 Y, 없으면 N 으로 표기하며 `journal_id`·상세 URL을 기록합니다.
4. **엑셀 저장** — 결과를 `.xlsx` 로 정리합니다.

### 같은 저널인지 확인(매칭) 규칙
- **ISSN 또는 eISSN 이 하나라도 일치하면 같은 저널로 확정**합니다(가장 강한 근거).
- 저널명 비교는 **대소문자 무시 + `&`↔`and` 통일 + 앞의 `The` 제거** 후 수행하며,
  Scholar 특유의 `약어-정식명`(예: `BJS-BRITISH JOURNAL OF SURGERY`),
  `정식명: 부제`, `&`↔`and` 표기 차이를 처리합니다.
- 확정된 매칭은 Scholar 저널명도 함께 엑셀에 기록합니다.
- `match_method` 열로 근거를 남깁니다: `issn_match`(ISSN 일치) / `title_exact`(제목 완전일치)
  / `title_match`(구문·약어 포함일치) / `title_fuzzy`(토큰 유사) / `not_found`(미검색).

---

## 3. 산출물 (오늘 날짜 폴더 안)

| 파일 | 내용 |
|---|---|
| `nature_index_scholar_<날짜>.xlsx` | **최종 결과**. 시트 4개: `result`(전체 178), `to_add (NI=N)`(Nature Index 태그가 아직 없는 저널 = 추가 대상), `not_found`(Scholar 미검색), `already_NI (NI=Y)`(이미 태그된 저널) |
| `01_nature_index_faq_<날짜>.xlsx` | 2단계 산출물: FAQ Q12 원본 저널 목록(178종) |
| `faq_raw.txt` | agent-browser 가 받아온 FAQ 페이지 텍스트(원본) |
| `_http_cache.json` | 외부/Scholar API 응답 캐시(재실행 가속·서버 부하 감소). 지워도 됩니다. |
| `_run_log.txt` | 실행 로그 |

### `result` 시트 주요 열
`no · Journal(Nature Index FAQ) · Scholar Title · Publisher · ISSN · eISSN ·
Nature Index DB?(Y/N) · journal_id · Scholar URL · Registered DB(all) · 등재년도 ·
articles(12M) · match_method · ext ISSN · ISSN verified`

---

## 4. 사전 환경 (이미 설치되어 있음)

- **Node.js 20** + **agent-browser** CLI + **Chrome for Testing**
  - 설치: `nvm use 20.19.6` → `npm install -g agent-browser` → `agent-browser install`
  - (agent-browser 는 Node 24 이상을 권장하지만, 프리빌드 바이너리라 Node 20 에서 설치·동작합니다.)
- **Python 3.10+** + `openpyxl` (`pip install openpyxl`) — 그 외는 표준 라이브러리만 사용.

`run.ps1` 은 `agent-browser` 실행 파일을 자동으로 찾습니다(Node 버전이 바뀌어도 v20.19.6 경로를 직접 참조).

> **이메일(선택)**: OpenAlex/Crossref polite-pool 용 연락 이메일은 환경변수
> `NATURE_INDEX_MAILTO` 로 지정합니다(미지정 시 중립 기본값 사용). 예:
> `$env:NATURE_INDEX_MAILTO = "you@example.com"`

이 저장소의 `sample_output/` 폴더에는 예시 결과(2026-07-26 수집본)가 들어 있습니다.

---

## 5. 정기 실행(선택)

Windows 작업 스케줄러 예시(매월 1일 09:00):
```
schtasks /Create /TN "NatureIndexAgent" /SC MONTHLY /D 1 /ST 09:00 ^
  /TR "cmd /c \"%~dp0run.cmd\""
```

---

## 6. 참고 / 한계

- **IROS 컨퍼런스**(`IEEE/RSJ International Conference on Intelligent Robots and Systems`)는
  학술지가 아닌 학회 프로시딩이라 Scholar 저널 DB에 없어 `not_found` 로 남습니다(정상).
- **OpenAlex** 는 2026년부터 일일 크레딧(예산) 제한이 생겨, 소진 시 해당 일에는 자동으로 건너뛰고
  Crossref + Scholar 만으로 매칭합니다(정확도에 큰 영향 없음 — 아래 검증 참고).
- **정확도 검증**: 기존 보유 파일(`현재nature저널_*.xlsx`, `신규로 추가된 NatureIndex 리스트_*.xlsx`)의
  검증 세트와 대조한 결과, 겹치는 저널 **145건의 journal_id 가 100% 일치**했습니다.
- Scholar 매칭 최종 ISSN/eISSN 은 **Scholar 등재 레코드 값**을 기준으로 하며,
  외부(OpenAlex/Crossref) ISSN 은 `ext ISSN` 열에 교차검증용으로 함께 표기합니다.
