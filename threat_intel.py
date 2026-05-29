# threat_intel.py
"""
첨부하신 threat_intel_mcp.py 와 동일한 로직을, 보고서 생성 파이프라인에서
in-process로 부르기 위해 가볍게 재포장한 모듈.

* MCP 서버를 별도 프로세스로 띄울 필요 없이 enrich_iocs() 한 줄로 호출 가능
* 환경변수 VT_API_KEY / ABUSE_API_KEY 가 없으면 enrichment를 조용히 건너뜀
  (보고서는 정상 생성됨)
"""

from __future__ import annotations

import asyncio
import logging
import os
from typing import Dict, List

import httpx

from models import IOC

log = logging.getLogger("threat-intel")

_TIMEOUT = httpx.Timeout(10.0)


async def _virustotal_ip(client: httpx.AsyncClient, ip: str) -> Dict:
    key = os.getenv("VT_API_KEY")
    if not key:
        return {}
    try:
        r = await client.get(
            f"https://www.virustotal.com/api/v3/ip_addresses/{ip}",
            headers={"x-apikey": key},
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        stats = r.json()["data"]["attributes"]["last_analysis_stats"]
        return {
            "vt_malicious": stats.get("malicious", 0),
            "vt_suspicious": stats.get("suspicious", 0),
        }
    except (httpx.HTTPError, KeyError) as e:
        log.warning("VirusTotal lookup failed for %s: %s", ip, e)
        return {}


async def _abuseipdb(client: httpx.AsyncClient, ip: str) -> Dict:
    key = os.getenv("ABUSE_API_KEY")
    if not key:
        return {}
    try:
        r = await client.get(
            "https://api.abuseipdb.com/api/v2/check",
            params={"ipAddress": ip, "maxAgeInDays": 90},
            headers={"Key": key, "Accept": "application/json"},
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        data = r.json()["data"]
        return {
            "abuse_score": data.get("abuseConfidenceScore", 0),
            "abuse_reports": data.get("totalReports", 0),
        }
    except (httpx.HTTPError, KeyError) as e:
        log.warning("AbuseIPDB lookup failed for %s: %s", ip, e)
        return {}


async def _enrich_one(client: httpx.AsyncClient, ioc: IOC) -> IOC:
    if ioc.type != "ip":
        return ioc
    vt, ab = await asyncio.gather(
        _virustotal_ip(client, ioc.value),
        _abuseipdb(client, ioc.value),
    )
    if not vt and not ab:
        return ioc
    parts: List[str] = []
    if vt:
        parts.append(f"VT mal/susp={vt['vt_malicious']}/{vt['vt_suspicious']}")
    if ab:
        parts.append(f"AbuseIPDB score={ab['abuse_score']} reports={ab['abuse_reports']}")
    extra = " | ".join(parts)
    new_desc = f"{ioc.description} ({extra})" if ioc.description else extra
    return ioc.model_copy(update={"description": new_desc})


async def enrich_iocs(iocs: List[IOC]) -> List[IOC]:
    """IOC 목록에 평판 정보를 비동기로 부착해 반환."""
    if not iocs:
        return iocs
    if not (os.getenv("VT_API_KEY") or os.getenv("ABUSE_API_KEY")):
        log.info("위협 인텔 API 키가 없어 enrichment를 건너뜁니다.")
        return iocs
    async with httpx.AsyncClient() as client:
        return await asyncio.gather(*[_enrich_one(client, i) for i in iocs])


def enrich_iocs_sync(iocs: List[IOC]) -> List[IOC]:
    """동기 코드에서 부르기 좋은 래퍼."""
    try:
        return asyncio.run(enrich_iocs(iocs))
    except RuntimeError:
        # 이미 실행 중인 이벤트 루프가 있는 환경(주피터 등) 대응
        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(enrich_iocs(iocs))
        finally:
            loop.close()
