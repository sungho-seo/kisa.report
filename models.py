# models.py
"""
보안 사고 보고서용 Pydantic 모델.

LangChain의 with_structured_output(IncidentReport)에 그대로 넣어 사용하면
LLM의 출력이 항상 이 스키마에 맞춰 검증된 객체로 돌아온다.

템플릿(.docx)의 placeholder와 1:1로 매칭되도록 필드명을 정했다.
- 스칼라 필드 ({{title}}, {{date}}, ...) : str/Literal
- 리스트 필드 ({{techniques}}, {{timeline}}, ...) : List[...]
  → docx_filler 가 줄바꿈으로 펼쳐 넣는다.
"""

from __future__ import annotations

from datetime import date
from typing import List, Literal, Optional

from pydantic import BaseModel, Field, field_validator

Severity = Literal["Low", "Medium", "High", "Critical"]
Tactic = Literal[
    "Reconnaissance",
    "Initial Access",
    "Execution",
    "Persistence",
    "Privilege Escalation",
    "Defense Evasion",
    "Credential Access",
    "Discovery",
    "Lateral Movement",
    "Collection",
    "Command and Control",
    "Exfiltration",
    "Impact",
]


class Technique(BaseModel):
    """MITRE ATT&CK 기법 1건."""

    id: str = Field(
        ...,
        pattern=r"^T\d{4}(\.\d{3})?$",
        description="MITRE ATT&CK 기법 ID (예: T1059, T1059.001)",
    )
    name: str = Field(..., description="기법 이름 (예: Command and Scripting Interpreter)")
    tactic: Tactic = Field(..., description="해당 기법이 속한 ATT&CK 전술")
    evidence: str = Field(..., description="해당 기법이라고 판단한 근거 (로그 라인 등)")


class IOC(BaseModel):
    """Indicator of Compromise 1건."""

    type: Literal["ip", "domain", "url", "hash", "email"] = Field(
        ..., description="IOC 종류"
    )
    value: str = Field(..., description="실제 값 (예: 8.8.8.8, evil.example.com)")
    description: Optional[str] = Field(
        default=None, description="해당 IOC가 사고와 어떻게 연관되는지 설명"
    )


class TimelineEvent(BaseModel):
    """사고 타임라인 한 줄."""

    when: str = Field(..., description="YYYY-MM-DD HH:MM (대략적 시각 가능)")
    what: str = Field(..., description="해당 시각에 발생한 사건 요약")


class IncidentReport(BaseModel):
    """보안 사고 보고서 전체. LLM이 채워 넣는 최종 객체."""

    title: str = Field(..., description="보고서 제목 (예: 외부 SSH 무차별 대입 시도 사고)")
    incident_id: str = Field(
        ...,
        description="사고 식별자 (예: INC-2025-001). 명시되지 않으면 임의 부여",
    )
    occurred_at: date = Field(..., description="사고 발생일 (YYYY-MM-DD)")
    severity: Severity = Field(..., description="심각도")
    summary: str = Field(
        ...,
        min_length=20,
        description="3~5문장 정도의 사고 개요. 임원 보고용으로 평이하게.",
    )
    affected_assets: List[str] = Field(
        default_factory=list, description="영향받은 자산 목록 (호스트명/IP/서비스명)"
    )
    techniques: List[Technique] = Field(
        default_factory=list, description="식별된 MITRE ATT&CK 기법 목록"
    )
    iocs: List[IOC] = Field(default_factory=list, description="관찰된 IOC 목록")
    timeline: List[TimelineEvent] = Field(
        default_factory=list, description="시간순 사고 진행 요약"
    )
    root_cause: str = Field(..., description="근본 원인 분석 (현재까지 파악된 범위)")
    recommendations: List[str] = Field(
        ..., min_length=1, description="재발 방지 권고사항 (실행 가능한 단위로 작성)"
    )

    # --- 검증 ---
    @field_validator("recommendations")
    @classmethod
    def _strip_empty_recos(cls, v: List[str]) -> List[str]:
        cleaned = [s.strip() for s in v if s and s.strip()]
        if not cleaned:
            raise ValueError("recommendations must contain at least one non-empty item")
        return cleaned
