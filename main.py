# main.py
"""
명령줄 진입점.

예시:
    python main.py \
        --template ./template_ko.docx \
        --logs ./samples/ssh_brute.log \
        --output ./out/incident_2025_001.docx

옵션:
    --skip-enrichment    위협 인텔(VirusTotal/AbuseIPDB) 보강 건너뛰기
    --model gpt-4o-mini  모델 변경
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from dotenv import load_dotenv

# .env 는 main.py 와 같은 폴더에 두는 것을 기본 가정
load_dotenv(Path(__file__).parent / ".env")

from analyzer import analyze_logs  # noqa: E402
from docx_filler import fill_template  # noqa: E402
from threat_intel import enrich_iocs_sync  # noqa: E402


def _setup_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        stream=sys.stderr,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="보안 로그 → LLM 분석 → MS Word 보고서 자동 생성"
    )
    parser.add_argument("--template", required=True, help="입력 .docx 템플릿 경로")
    parser.add_argument("--logs", required=True, help="분석 대상 로그 파일 경로(텍스트)")
    parser.add_argument("--output", required=True, help="출력 .docx 경로")
    parser.add_argument("--model", default="gpt-4o", help="OpenAI 모델명 (기본 gpt-4o)")
    parser.add_argument(
        "--similar-cases",
        default=None,
        help="(선택) 유사 사례 텍스트 파일 경로 - vector store 결과 등을 미리 추출해 전달",
    )
    parser.add_argument(
        "--skip-enrichment",
        action="store_true",
        help="VirusTotal/AbuseIPDB로 IOC 보강하지 않기",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="DEBUG 로그 출력")
    args = parser.parse_args(argv)

    _setup_logging(args.verbose)
    log = logging.getLogger("main")

    log_path = Path(args.logs)
    if not log_path.exists():
        log.error("로그 파일을 찾을 수 없습니다: %s", log_path)
        return 2

    raw_logs = log_path.read_text(encoding="utf-8", errors="replace")

    similar_cases = None
    if args.similar_cases:
        sc = Path(args.similar_cases)
        if not sc.exists():
            log.error("유사 사례 파일을 찾을 수 없습니다: %s", sc)
            return 2
        similar_cases = sc.read_text(encoding="utf-8", errors="replace")

    # 1. LLM 분석
    report = analyze_logs(raw_logs, similar_cases=similar_cases, model=args.model)

    # 2. 위협 인텔 보강 (선택)
    if not args.skip_enrichment and report.iocs:
        log.info("IOC %d건에 대해 위협 인텔 조회", len(report.iocs))
        report = report.model_copy(update={"iocs": enrich_iocs_sync(report.iocs)})

    # 3. Word 보고서 생성
    out_path = fill_template(args.template, report, args.output)
    log.info("✅ 보고서 생성 완료: %s", out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
