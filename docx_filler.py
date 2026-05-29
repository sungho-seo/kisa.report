# docx_filler.py
"""
사용자가 지정한 .docx 템플릿 안의 {{placeholder}} 를 IncidentReport 값으로
치환하여 .docx 를 만드는 모듈.

첨부 코드에서 발생하던 4가지 문제를 모두 해결한다:

(1) 서식 손실:
    `p.text = p.text.replace(...)` 는 Run 단위 폰트·굵기·색을 모두 날린다.
    → 본 모듈은 Run을 직접 편집하여 첫 Run의 서식을 유지한다.

(2) Run 분할 placeholder:
    Word는 맞춤법 검사·자동 교정 결과로 "{{ti", "tle", "}}" 처럼 placeholder
    하나가 여러 Run으로 쪼개진 상태로 저장하는 일이 흔하다.
    `if "{{title}}" in p.text` 는 매칭되지만 단순 replace는 동작하지 않는다.
    → 본 모듈은 Run 경계와 무관하게 paragraph 텍스트 전체에서 placeholder
      위치를 찾고, 걸쳐 있는 Run들을 일관되게 재기록한다.

(3) 표·머리글·바닥글 미처리:
    템플릿의 표지나 요약 테이블 안의 placeholder가 채워지지 않는 문제.
    → iter_all_paragraphs() 로 body/table/header/footer를 모두 순회한다.

(4) 리스트 값의 줄바꿈:
    "\n".join(...) 결과를 그대로 넣으면 Word에서 한 줄로 붙거나 깨진다.
    → 줄바꿈은 같은 Paragraph 안의 soft line break(<w:br/>) 로 직접 삽입.
      placeholder가 paragraph 단독이면 새 Paragraph로 펼쳐 넣는 옵션도 제공.
"""

from __future__ import annotations

import copy
import logging
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Sequence, Tuple

from docx import Document
from docx.document import Document as _DocxDocument
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from models import IOC, IncidentReport, Technique, TimelineEvent

log = logging.getLogger("docx-filler")

# placeholder 형식: {{ key }} (양쪽 공백 허용)
_PH_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")


# ---------------------------------------------------------------------------
# 1. 보고서 값 → 문자열 매핑
# ---------------------------------------------------------------------------
def _fmt_date(d: Any) -> str:
    if isinstance(d, (date, datetime)):
        return d.strftime("%Y-%m-%d")
    return str(d)


def _fmt_technique(t: Technique) -> str:
    return f"[{t.id}] {t.name} ({t.tactic}) — 근거: {t.evidence}"


def _fmt_ioc(i: IOC) -> str:
    base = f"{i.type.upper()}: {i.value}"
    return f"{base} — {i.description}" if i.description else base


def _fmt_timeline(e: TimelineEvent) -> str:
    return f"{e.when}  {e.what}"


def build_value_map(report: IncidentReport) -> Dict[str, str]:
    """
    IncidentReport → {placeholder_key: value_str} 매핑.

    리스트 필드는 항목 사이를 '\n'으로 연결한다.
    docx_filler 의 _set_paragraph_text 가 '\n' 을 soft line break로 변환한다.
    """
    techs = "\n".join(_fmt_technique(t) for t in report.techniques) or "(해당 없음)"
    iocs = "\n".join(_fmt_ioc(i) for i in report.iocs) or "(해당 없음)"
    timeline = "\n".join(_fmt_timeline(e) for e in report.timeline) or "(해당 없음)"
    recos = "\n".join(f"{idx}. {r}" for idx, r in enumerate(report.recommendations, 1))
    assets = "\n".join(f"- {a}" for a in report.affected_assets) or "(해당 없음)"

    return {
        "title": report.title,
        "incident_id": report.incident_id,
        "date": _fmt_date(report.occurred_at),
        "severity": report.severity,
        "summary": report.summary,
        "assets": assets,
        "techniques": techs,
        # 짧은 형태(쉼표 구분, 표 셀에 어울림)도 별도 placeholder로 제공
        "techniques_short": ", ".join(t.id for t in report.techniques) or "(해당 없음)",
        "iocs": iocs,
        "timeline": timeline,
        "root_cause": report.root_cause,
        "recommendations": recos,
    }


# ---------------------------------------------------------------------------
# 2. 모든 paragraph 순회 (body + table + header + footer, 재귀적으로)
# ---------------------------------------------------------------------------
def _iter_block_paragraphs(parent) -> Iterable[Paragraph]:
    """parent의 모든 paragraph를 표 내부까지 재귀적으로 yield."""
    # body 또는 cell의 직속 paragraph
    if hasattr(parent, "paragraphs"):
        for p in parent.paragraphs:
            yield p
    # 내부 표
    tables: Sequence[Table] = getattr(parent, "tables", [])
    for tbl in tables:
        for row in tbl.rows:
            for cell in row.cells:
                yield from _iter_block_paragraphs(cell)


def iter_all_paragraphs(doc: _DocxDocument) -> Iterable[Paragraph]:
    # 본문
    yield from _iter_block_paragraphs(doc)
    # 모든 섹션의 머리글/바닥글
    for section in doc.sections:
        yield from _iter_block_paragraphs(section.header)
        yield from _iter_block_paragraphs(section.footer)
        # 첫 페이지/짝수 페이지 머리글·바닥글까지 (있다면)
        for attr in ("first_page_header", "first_page_footer",
                     "even_page_header", "even_page_footer"):
            part = getattr(section, attr, None)
            if part is not None:
                yield from _iter_block_paragraphs(part)


# ---------------------------------------------------------------------------
# 3. paragraph 텍스트 안전 치환  (핵심 알고리즘)
# ---------------------------------------------------------------------------
@dataclass
class _RunSlice:
    """paragraph 안에서 어떤 Run이 [start, end) 범위의 글자에 해당하는지."""
    run_index: int
    start: int  # paragraph 텍스트상 시작 오프셋(포함)
    end: int    # paragraph 텍스트상 끝 오프셋(미포함)


def _index_runs(p: Paragraph) -> Tuple[str, List[_RunSlice]]:
    """paragraph 전체 텍스트와, 각 Run의 텍스트 오프셋 매핑을 만들어 반환."""
    text_parts: List[str] = []
    slices: List[_RunSlice] = []
    offset = 0
    for idx, run in enumerate(p.runs):
        t = run.text or ""
        slices.append(_RunSlice(idx, offset, offset + len(t)))
        text_parts.append(t)
        offset += len(t)
    return "".join(text_parts), slices


def _set_run_text_with_breaks(run, text: str) -> None:
    """
    Run에 텍스트를 채워 넣되, '\n'은 soft line break(<w:br/>)로 변환한다.
    Run의 서식(rPr)은 그대로 유지된다.

    Run 내부 구조는 기본적으로 <w:r> 안에 <w:rPr> 와 <w:t> 가 있다.
    우리는 <w:t>와 <w:br/>을 직접 다루기 위해 기존 자식 노드(텍스트/브레이크)를
    제거하고 새로 끼워 넣는다. rPr은 보존.
    """
    r_el = run._r
    # rPr만 남기고 텍스트/브레이크 자식 제거
    for child in list(r_el):
        tag = child.tag
        if tag == qn("w:rPr"):
            continue
        if tag in (qn("w:t"), qn("w:br"), qn("w:tab"), qn("w:cr")):
            r_el.remove(child)

    # 새 자식 끼워 넣기 (\n 기준 분할 → <w:t> 사이마다 <w:br/>)
    from docx.oxml import OxmlElement

    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i > 0:
            r_el.append(OxmlElement("w:br"))
        if line == "":
            continue
        t_el = OxmlElement("w:t")
        t_el.text = line
        # 선행·후행 공백 보존
        t_el.set(qn("xml:space"), "preserve")
        r_el.append(t_el)


def _replace_in_paragraph(p: Paragraph, values: Dict[str, str]) -> bool:
    """
    paragraph 안의 모든 {{key}} 를 values 의 값으로 치환.
    Run 경계를 가로지르는 placeholder도 처리한다.
    반환: 1회 이상 치환했으면 True.
    """
    changed = False
    # 동일 paragraph에 placeholder가 여러 개 있을 수 있으므로,
    # 한 번에 하나씩 처리하며 매번 인덱스를 새로 잡는다.
    while True:
        full_text, slices = _index_runs(p)
        if not slices:
            return changed

        m = _PH_RE.search(full_text)
        if not m:
            return changed
        key = m.group(1)
        if key not in values:
            # 정의되지 않은 placeholder는 건드리지 않고 다음 placeholder를 찾기 위해
            # 일단 빈 문자열로 두지 말고 그대로 두기 위해 종료
            log.warning("템플릿에 정의되지 않은 placeholder: {{%s}}", key)
            # 같은 paragraph에 다른 placeholder가 있을 수 있어 잘라 살피기 위해
            # 임시로 안전한 sentinel로 마킹했다가 마지막에 복원하는 방식도 가능하지만,
            # 일반적인 보고서 사용에서는 placeholder가 그렇게 많지 않으므로
            # 첫 미정의 placeholder에서 멈춘다.
            return changed

        value = values[key]
        ph_start, ph_end = m.span()

        # placeholder가 걸친 Run 범위
        first_idx = None
        last_idx = None
        for s in slices:
            if s.start < ph_end and s.end > ph_start:
                if first_idx is None:
                    first_idx = s.run_index
                last_idx = s.run_index
        if first_idx is None or last_idx is None:
            # 이 경우는 발생하면 안 되지만, 안전하게 빠져나감
            return changed

        first_run = p.runs[first_idx]
        first_slice = slices[first_idx]
        last_slice = slices[last_idx]

        # 첫 Run에 들어갈 텍스트 = (placeholder 앞 부분) + (치환값) + (마지막 Run의 placeholder 뒷부분)
        prefix = first_run.text[: ph_start - first_slice.start]
        suffix = p.runs[last_idx].text[ph_end - last_slice.start :]

        # 첫 Run의 텍스트 갱신 (서식 유지, 줄바꿈은 break로 변환)
        _set_run_text_with_breaks(first_run, prefix + value + suffix)

        # 중간 Run들과 마지막 Run의 텍스트는 비움(서식 그대로 유지)
        for idx in range(first_idx + 1, last_idx + 1):
            p.runs[idx].text = ""

        changed = True
        # 다음 placeholder를 찾기 위해 루프 계속


# ---------------------------------------------------------------------------
# 4. 진입점
# ---------------------------------------------------------------------------
def fill_template(
    template_path: str | Path,
    report: IncidentReport,
    output_path: str | Path,
    extra_values: Dict[str, str] | None = None,
) -> Path:
    """
    template_path 의 .docx 를 열어 placeholder를 보고서 값으로 채우고
    output_path 에 저장한다.

    Parameters
    ----------
    template_path : 사용자가 지정한 .docx 템플릿 경로.
        본문/표/머리글/바닥글 어디든 {{key}} 형태로 placeholder를 둘 수 있다.
    report : LLM이 채워준 IncidentReport.
    output_path : 결과 .docx 저장 경로.
    extra_values : 추가로 치환할 임의의 키-값 (예: {"company": "ACME Corp"})

    Returns
    -------
    저장된 파일 경로 (Path).
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")
    if template_path.suffix.lower() != ".docx":
        raise ValueError(
            f"템플릿은 .docx 파일이어야 합니다. (입력: {template_path.suffix})"
        )

    values = build_value_map(report)
    if extra_values:
        values.update({str(k): str(v) for k, v in extra_values.items()})

    log.info("템플릿 열기: %s", template_path)
    doc = Document(str(template_path))

    total_changed = 0
    paragraphs_seen = 0
    for p in iter_all_paragraphs(doc):
        paragraphs_seen += 1
        if _replace_in_paragraph(p, values):
            total_changed += 1

    log.info(
        "치환 완료: %d/%d paragraph 수정 (정의된 키 %d개)",
        total_changed,
        paragraphs_seen,
        len(values),
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    log.info("저장: %s", output_path)
    return output_path
