# 보안 사고 보고서 자동 생성기

원시 보안 로그를 LLM에 넘겨 분석한 결과를, 사용자가 지정한 **MS Word 템플릿**에
채워 넣어 보고서를 자동 생성한다.

## 파일 구성

```
security_report/
├── models.py              # IncidentReport / Technique / IOC Pydantic 스키마
├── analyzer.py            # LangChain + with_structured_output 로 로그 → IncidentReport
├── threat_intel.py        # VirusTotal / AbuseIPDB 로 IOC 보강 (선택)
├── docx_filler.py         # 템플릿의 {{placeholder}} 안전 치환 (핵심)
├── main.py                # CLI 진입점
├── make_sample_template.py# 예시 템플릿 생성기
├── test_filler.py         # docx_filler 단위 테스트 (LLM 호출 없음)
├── template_ko.docx       # 예시 템플릿 (자유롭게 교체)
├── requirements.txt
└── .env.example
```

## 첨부 코드 대비 수정한 부분

| # | 첨부 코드의 문제 | 해결 위치 |
|---|---|---|
| 1 | `p.text = p.text.replace(...)` → Run 단위 서식(폰트·굵기·색)이 사라짐 | `docx_filler._replace_in_paragraph` 가 첫 Run의 `rPr`을 그대로 두고 텍스트만 교체 |
| 2 | placeholder가 여러 Run에 쪼개진 상태에서 `replace`가 동작하지 않음 (Word 자동 교정·맞춤법 검사 영향) | paragraph 전체 텍스트에서 정규식으로 찾고, 걸쳐 있는 Run들을 일관되게 재기록 |
| 3 | 표·머리글·바닥글의 placeholder 미처리 | `iter_all_paragraphs` 가 body/table/header/footer/even·first page까지 재귀 순회 |
| 4 | `"\n".join(...)` 결과를 `p.text` 에 넣으면 줄바꿈이 깨짐 | `_set_run_text_with_breaks` 가 `\n` 을 `<w:br/>` soft line break로 삽입 |
| 5 | image 1/4의 Pydantic 필드 불일치 (`techniques` vs `mitre_techniques`) | `models.IncidentReport` 로 통일, 필드와 placeholder를 1:1 매핑 |
| 6 | `report.date` 가 date 객체일 때 그대로 join → TypeError | `_fmt_date` 가 date/datetime/str 모두 처리 |
| 7 | LLM 출력 후 JSON 파싱 실패 위험 | `with_structured_output(IncidentReport)` 로 Pydantic 검증까지 한 번에 |
| 8 | 위협 인텔을 MCP 별도 프로세스로만 부를 수 있던 점 | `threat_intel.enrich_iocs_sync` 로 in-process 호출 가능, API 키 없으면 자동 스킵 |

## 사용

### 1. 설치

```bash
pip install -r requirements.txt
cp .env.example .env       # OPENAI_API_KEY 등을 채워 넣는다
```

### 2. 템플릿 준비

회사 양식이 있다면 **Word에서** 채우고 싶은 자리에 `{{key}}` 형태로
placeholder만 박아 두면 된다. 사용 가능한 키:

| placeholder | 형태 | 설명 |
|---|---|---|
| `{{title}}`            | str  | 사고 제목 |
| `{{incident_id}}`      | str  | 사고 ID |
| `{{date}}`             | str  | 발생일 (YYYY-MM-DD) |
| `{{severity}}`         | str  | Low/Medium/High/Critical |
| `{{summary}}`          | str  | 사고 개요 (3~5문장) |
| `{{assets}}`           | 여러 줄 | 영향받은 자산 |
| `{{techniques}}`       | 여러 줄 | MITRE 기법 상세 |
| `{{techniques_short}}` | str  | T1110, T1078.003 형식 (표 셀에 적합) |
| `{{iocs}}`             | 여러 줄 | IOC 목록 (위협 인텔 보강 포함) |
| `{{timeline}}`         | 여러 줄 | 시간순 사건 |
| `{{root_cause}}`       | str  | 근본 원인 |
| `{{recommendations}}`  | 여러 줄 | 번호 매겨진 권고사항 |

회사 양식이 없으면:
```bash
python make_sample_template.py template_ko.docx
```
을 실행해 예시 템플릿을 만들고 Word로 열어 자유롭게 수정.

### 3. 실행

```bash
python main.py \
    --template ./template_ko.docx \
    --logs ./samples/ssh_brute.log \
    --output ./out/incident_2025_001.docx
```

위협 인텔 보강을 건너뛰고 싶거나 키가 없으면:
```bash
python main.py ... --skip-enrichment
```

저비용 모델로:
```bash
python main.py ... --model gpt-4o-mini
```

### 4. 검증

```bash
python test_filler.py
```
LLM 호출 없이 docx_filler 의 5개 케이스(서식 보존, Run 분할,
다중 placeholder, 미정의 placeholder, e2e)를 모두 검증한다.

## PDF로 변환하려면

코드 안에서 자동 변환은 환경 의존성(LibreOffice/Word)이 있어 분리해 두었다.
별도 명령으로:

```bash
# LibreOffice 가 설치되어 있다면
soffice --headless --convert-to pdf out/incident_2025_001.docx

# 또는 Windows + MS Word 환경이면 docx2pdf
pip install docx2pdf
python -c "from docx2pdf import convert; convert('out/incident_2025_001.docx')"
```

## 코드를 라이브러리로 가져다 쓰기

```python
from analyzer import analyze_logs
from docx_filler import fill_template
from threat_intel import enrich_iocs_sync

raw = open("logs.txt", encoding="utf-8").read()
report = analyze_logs(raw)
report = report.model_copy(update={"iocs": enrich_iocs_sync(report.iocs)})
fill_template("template_ko.docx", report, "out.docx",
              extra_values={"company": "ACME"})  # 임의 키 추가도 가능
```
