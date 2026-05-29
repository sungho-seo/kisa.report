# analyzer.py
"""
LLM을 호출해 raw 보안 로그를 IncidentReport 객체로 변환한다.

핵심:
* ChatOpenAI(...).with_structured_output(IncidentReport) 를 사용 → LLM 응답이
  자동으로 Pydantic 객체로 파싱·검증된다. 사후 JSON 파싱 로직이 필요 없다.
* 프롬프트는 시스템/유저로 분리하고, 출력 언어를 명시적으로 한국어로 지정.
* 선택적 RAG: 유사 사례를 vector_store에서 가져와 컨텍스트로 추가할 수 있음.
"""

from __future__ import annotations

import logging
import os
from typing import Optional

from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI

from models import IncidentReport

log = logging.getLogger("analyzer")

_SYSTEM_PROMPT = """\
당신은 한국어로 보고하는 시니어 SOC(Security Operations Center) 분석가다.
입력으로 주어진 원시 보안 로그/이벤트를 읽고, 사고 보고서를 IncidentReport
스키마에 맞춰 채워라.

작성 원칙:
- 모든 문자열 필드는 한국어로 작성한다(고유명사·기술 용어는 영문 유지).
- MITRE ATT&CK 기법 ID는 실제 존재하는 ID만 사용한다(예: T1110, T1059.001).
  근거(evidence)에는 해당 판단의 출처가 되는 로그 라인이나 패턴을 인용한다.
- IOC(ip/domain/url/hash/email)는 로그에서 확인 가능한 것만 기록하고,
  description에는 사고와의 연관성을 한 줄로 설명한다.
- timeline 은 가능한 한 실제 로그 타임스탬프를 사용해 시간 순으로 정렬한다.
- recommendations 는 실행 가능한 단위(누가/무엇을/언제까지 가능 여부)로 적는다.
- 확신이 없는 항목은 추정으로 채우지 말고 "추정 불가" 또는 "로그상 확인 불가"
  로 명시한다.
"""


def build_analyzer(model: str = "gpt-4o", temperature: float = 0.0):
    """
    설정된 LLM을 IncidentReport 구조화 출력에 묶어 반환한다.

    OPENAI_API_KEY 환경변수가 필요하다.
    """
    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError(
            "OPENAI_API_KEY 환경변수가 설정되어 있지 않습니다. "
            ".env 파일에 OPENAI_API_KEY=sk-... 형식으로 추가하세요."
        )
    llm = ChatOpenAI(model=model, temperature=temperature)
    # with_structured_output: LLM의 출력을 곧바로 IncidentReport 객체로 받음
    return llm.with_structured_output(IncidentReport)


def analyze_logs(
    raw_logs: str,
    *,
    similar_cases: Optional[str] = None,
    model: str = "gpt-4o",
) -> IncidentReport:
    """
    원시 보안 로그를 받아 IncidentReport 로 변환.

    Parameters
    ----------
    raw_logs : 분석 대상이 되는 보안 로그 텍스트.
    similar_cases : (선택) vector store에서 가져온 유사 사례 텍스트.
        None이면 컨텍스트 없이 LLM이 단독 판단한다.
    model : 사용 모델명. 기본 gpt-4o.
    """
    if not raw_logs or not raw_logs.strip():
        raise ValueError("raw_logs 가 비어 있습니다.")

    analyzer = build_analyzer(model=model)

    user_content_parts = []
    if similar_cases:
        user_content_parts.append(f"### 유사 사례\n{similar_cases.strip()}")
    user_content_parts.append(
        "### 분석 대상 로그\n다음 로그를 IncidentReport 형식으로 정리하라.\n\n"
        f"{raw_logs.strip()}"
    )
    user_content = "\n\n".join(user_content_parts)

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=user_content),
    ]

    log.info("LLM 호출 (모델=%s, 입력 길이=%d자)", model, len(raw_logs))
    report = analyzer.invoke(messages)
    # with_structured_output 덕분에 report 는 이미 IncidentReport 인스턴스
    if not isinstance(report, IncidentReport):  # 방어적 체크
        raise RuntimeError(f"LLM이 IncidentReport가 아닌 타입을 반환: {type(report)}")
    log.info("LLM 응답 파싱 완료: %s (severity=%s)", report.title, report.severity)
    return report
