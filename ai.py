"""Claude API helper: report section drafts and intake interview summaries.

Optional. It switches on when Streamlit secrets hold

    [anthropic]
    api_key = "sk-ant-..."
    model = "claude-sonnet-5-5"      # optional

Only the data needed for the one section being written is sent. Nothing is invented on purpose: the prompts
tell the model to use the given data only and to mark gaps with (확인 필요).
"""
from __future__ import annotations

import json
import os
import re

import streamlit as st

DEFAULT_MODEL = "claude-sonnet-5-5"

SYSTEM = {
    "ko": (
        "너는 Plug and Play 액셀러레이터 프로그램 매니저를 돕는 보고서 작성 보조다. "
        "한국 공공기관 보고서의 개조식 문체로 쓴다. 문장은 '~함', '~임', '~예정', '~완료'처럼 명사형으로 끝낸다. "
        "주어진 데이터에 있는 사실, 숫자, 회사명, 날짜만 쓴다. 데이터에 없는 내용은 만들지 말고, 꼭 필요하면 '(확인 필요)'라고 쓴다. "
        "과장하거나 홍보하는 표현은 쓰지 않는다. 출력은 Markdown 본문만 쓴다. '#' 제목은 쓰지 않는다 (섹션 제목은 이미 있다). "
        "글머리표는 '- '를 쓰고 하위 항목은 두 칸 들여쓴다. 표가 더 읽기 쉬우면 Markdown 표를 써도 된다. "
        "앞뒤 인사말이나 설명 없이 본문만 출력한다."
    ),
    "en": (
        "You help an accelerator program manager at Plug and Play write program reports. Write concise report "
        "bullets in plain business English. Use only the facts, numbers, company names and dates in the data you "
        "are given; never invent anything, and write '(to be confirmed)' where something needed is missing. "
        "No marketing language. Output Markdown body text only, no '#' headings (the section title already "
        "exists), '- ' bullets with two-space indented sub bullets, and a Markdown table where it reads better. "
        "No greeting or explanation before or after."
    ),
}


def _settings() -> tuple[str | None, str]:
    key, model = None, DEFAULT_MODEL
    try:
        sec = st.secrets["anthropic"]
        key = sec.get("api_key") or None
        model = sec.get("model") or DEFAULT_MODEL
    except Exception:  # noqa: BLE001  (no secrets file or no [anthropic] section)
        pass
    return key or os.environ.get("ANTHROPIC_API_KEY"), model


def available() -> bool:
    return bool(_settings()[0]) or os.environ.get("BIZMATCH_FAKE_AI") == "1"


def model_name() -> str:
    return _settings()[1]


def _call(system: str, prompt: str, max_tokens: int = 2000) -> str:
    if os.environ.get("BIZMATCH_FAKE_AI") == "1":          # used by the automated tests only
        return _fake(prompt)
    import anthropic

    key, model = _settings()
    client = anthropic.Anthropic(api_key=key)
    msg = client.messages.create(model=model, max_tokens=max_tokens, system=system,
                                 messages=[{"role": "user", "content": prompt}])
    return "".join(getattr(b, "text", "") for b in msg.content).strip()


def _fake(prompt: str) -> str:
    if "INTAKE" in prompt:
        return ("- **기업 현황**: 인터뷰 답변 기준으로 정리함\n- **프로그램 목표**: 미국 시장 파트너 발굴\n"
                "- **니즈**: 투자 유치, 현지 PoC\n\n```json\n{\"tags\": \"AI, B2B\", \"stage\": \"Seed\", "
                "\"one_liner\": \"AI tool for factories\", \"partner_types\": \"Corporate, VC\"}\n```")
    m = re.search(r"섹션: (.+)|Section: (.+)", prompt)
    title = (m.group(1) or m.group(2)).strip() if m else "section"
    return f"- {title} 내용을 데이터 기준으로 정리함\n  - 세부 내용 (확인 필요)"


def _clean(md: str) -> str:
    md = re.sub(r"^```(?:markdown|md)?\s*\n", "", md.strip())
    md = re.sub(r"\n```\s*$", "", md)
    return "\n".join(line for line in md.splitlines() if not line.lstrip().startswith("# ")).strip()


def draft_section(report_kind: str, title: str, instruction: str, data: dict, current: str = "",
                  lang: str = "ko") -> str:
    """Write one report section from the program data (Markdown body, no heading)."""
    if lang == "ko":
        prompt = (f"보고서 종류: {report_kind}\n섹션: {title}\n작성 지침: {instruction}\n\n"
                  f"데이터 (JSON):\n{json.dumps(data, ensure_ascii=False, default=str)}\n")
        if current.strip():
            prompt += f"\n지금 초안 (데이터와 맞는 부분은 살리고 다듬기):\n{current}\n"
    else:
        prompt = (f"Report: {report_kind}\nSection: {title}\nInstructions: {instruction}\n\n"
                  f"Data (JSON):\n{json.dumps(data, ensure_ascii=False, default=str)}\n")
        if current.strip():
            prompt += f"\nCurrent draft (keep what matches the data, improve the rest):\n{current}\n"
    return _clean(_call(SYSTEM.get(lang, SYSTEM["ko"]), prompt))


def summarize_intake(startup: str, profile: dict, qa: list[tuple[str, str, str]], notes: str,
                     lang: str = "ko") -> tuple[str, dict]:
    """Interview notes -> (summary Markdown, suggested directory fields)."""
    lines = "\n".join(f"[{sec}] Q. {q}\nA. {a or '(답변 없음)'}" for sec, q, a in qa)
    if lang == "ko":
        prompt = (
            f"INTAKE 인터뷰 기록을 정리해줘. 기업: {startup}\n회사 정보: {json.dumps(profile, ensure_ascii=False)}\n\n"
            f"질의응답:\n{lines}\n\n추가 메모:\n{notes or '(없음)'}\n\n"
            "다음 순서로 굵은 소제목 글머리표를 써줘: **기업 현황**, **프로그램 목표**, **니즈** (투자, 협력, 시장 진출, 기타), "
            "**추천 파트너 유형**, **후속 조치**. 각 항목은 1~3줄, 개조식. 답변에 없는 내용은 쓰지 마.\n"
            "마지막에 회사 정보에 넣을 값을 ```json 블록으로 줘. 키는 tags (쉼표로 구분한 분야와 B2B/B2C), stage (투자 단계), "
            "one_liner (영문 한 줄 소개), partner_types (만나면 좋은 파트너 유형). 모르면 빈 문자열."
        )
    else:
        prompt = (
            f"Summarize this INTAKE interview. Company: {startup}\nProfile: {json.dumps(profile, ensure_ascii=False)}\n\n"
            f"Q&A:\n{lines}\n\nExtra notes:\n{notes or '(none)'}\n\n"
            "Write bullets with bold labels in this order: **Company status**, **Program goals**, **Needs** "
            "(investment, partnerships, market entry, other), **Recommended partner types**, **Follow-ups**. 1 to 3 "
            "lines each. Nothing that is not in the answers.\nThen give values for the company directory in a "
            "```json block with keys tags (comma separated sectors plus B2B/B2C), stage, one_liner, partner_types. "
            "Empty string when unknown."
        )
    out = _call(SYSTEM.get(lang, SYSTEM["ko"]), prompt, max_tokens=1600)
    fields = {}
    m = re.search(r"```json\s*(\{.*?\})\s*```", out, re.S)
    if m:
        try:
            fields = {k: str(v).strip() for k, v in json.loads(m.group(1)).items() if str(v).strip()}
        except Exception:  # noqa: BLE001
            fields = {}
        out = out[:m.start()] + out[m.end():]
    return _clean(out), fields
