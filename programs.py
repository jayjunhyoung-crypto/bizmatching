"""Accelerator programs: stages and schedule, participants, intake interviews, KPIs and outcomes, reports.

app.py fills API with the helpers this module needs (database, current user, company directory, matching
results), so this file never imports app.py.
"""
from __future__ import annotations

import copy
import hashlib
import json
import time
import uuid
from datetime import date, timedelta

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import ai
import docgen
from i18n import lang

API: dict = {}   # store, me, norm, dir_rows, dir_name, dir_update, matching_stats, event_choices, pages


def ss():
    return st.session_state


def T(key: str, **kw) -> str:
    en, ko = STR.get(key, (key, key))
    s = ko if lang() == "ko" else en
    return s.format(**kw) if kw else s


def L2(pair) -> str:
    """(en, ko) pair in the current language."""
    return pair[1] if lang() == "ko" else pair[0]


def norm(x) -> str:
    return API["norm"](x) if API.get("norm") else str(x or "").strip().lower()


STR = {
    "programs": ("Programs", "프로그램"),
    "programs_cap": ("Every accelerator program in one place: stages and deadlines, startups, intake interviews, "
                     "results and reports.", "액셀러레이터 프로그램을 한곳에서 관리해요. 단계와 마감, 참가사, intake interview, 성과, 보고서까지."),
    "new_program": ("＋ New program", "＋ 새 프로그램"),
    "p_name": ("Program name", "프로그램 이름"), "p_client": ("Client", "발주처"),
    "p_template": ("Template", "템플릿"), "p_start": ("Start", "시작일"), "p_end": ("End", "종료일"),
    "p_pm": ("PM", "PM"), "p_team": ("Team", "팀원"), "p_budget": ("Budget", "예산"),
    "p_summary": ("Objective and scope (one point per line)", "사업 목적과 범위 (한 줄에 하나씩)"),
    "create": ("Create", "만들기"), "need_name": ("Give the program a name.", "프로그램 이름을 적어주세요."),
    "empty_hub": ("No programs yet. Create one from a template, or start with an example to look around.",
                  "아직 프로그램이 없어요. 템플릿으로 새로 만들거나, 예시 프로그램으로 둘러보세요."),
    "example": ("Create an example program", "예시 프로그램 만들기"),
    "upcoming": ("Coming up", "다가오는 마감"), "nothing_due": ("Nothing due in the next two weeks.", "2주 안에 마감되는 단계가 없어요."),
    "overdue": ("overdue {n}d", "{n}일 지남"), "dday": ("D-{n}", "D-{n}"), "today": ("today", "오늘"),
    "open": ("Open", "열기"), "back": ("← All programs", "← 프로그램 목록"),
    "n_startups": ("{n} startups", "{n}개사"), "stage_now": ("Now: {s}", "진행 중: {s}"),
    "next_due": ("Next: {s} {d}", "다음 마감: {s} {d}"), "progress": ("{d}/{n} stages done", "{n}단계 중 {d}단계 완료"),
    "all_done": ("All stages done", "모든 단계 완료"),
    "tab_overview": ("Overview", "개요"), "tab_stages": ("Stages & schedule", "단계·일정"),
    "tab_startups": ("Startups", "참가사"), "tab_intake": ("Intake interview", "Intake interview"),
    "tab_kpi": ("Results & KPIs", "성과·KPI"), "tab_report": ("Reports", "보고서"),
    "events_link": ("Business matching events in this program", "이 프로그램의 비즈매칭 일정"),
    "events_help": ("Their meetings count toward KPIs and reports.", "여기서 고른 일정의 미팅이 KPI와 보고서에 들어가요."),
    "no_events": ("No business matching schedule yet. Make one on the Business matching page.",
                  "아직 비즈매칭 일정이 없어요. 비즈매칭 페이지에서 만들면 여기서 연결할 수 있어요."),
    "manage": ("Manage", "관리"), "delete_prog": ("Delete this program", "이 프로그램 삭제"),
    "delete_confirm": ("Yes, delete it for everyone", "팀 전체에서 삭제할게요"),
    "deleted": ("Program deleted.", "프로그램을 삭제했어요."), "missing": ("That program no longer exists.", "그 프로그램은 삭제됐어요."),
    "col_status": ("Status", "상태"), "col_stage": ("Stage", "단계"), "col_start": ("Start", "시작"),
    "col_due": ("Due", "마감"), "col_owner": ("Owner", "담당"), "col_note": ("Notes", "메모"),
    "checklist": ("Checklist", "체크리스트"), "ck_item": ("Item", "항목"), "ck_done": ("Done", "완료"),
    "ck_for": ("Checklist for", "체크리스트 단계"), "ck_count": ("{d}/{n} checked", "{n}개 중 {d}개 완료"),
    "no_dates": ("Add start and due dates to see the timeline.", "시작일과 마감일을 넣으면 일정표가 보여요."),
    "stages_help": ("Add, remove or reorder stages freely. Every program can have its own steps.",
                    "단계는 자유롭게 추가하고 지울 수 있어요. 프로그램마다 단계가 달라도 돼요."),
    "pick_startups": ("Startups in this program", "이 프로그램 참가사"),
    "pick_help": ("From the company directory. Add or fix company details on the Companies page.",
                  "회사 정보에서 골라요. 회사 정보 수정은 회사 정보 페이지에서 해요."),
    "dir_empty": ("The company directory is empty. Add the startups there first.", "회사 정보가 비어 있어요. 먼저 회사 정보에 스타트업을 넣어주세요."),
    "go_companies": ("Open Companies", "회사 정보 열기"),
    "c_company": ("Company", "회사"), "c_kr": ("Korean name", "국문명"), "c_tags": ("Tags", "분야"),
    "c_intake": ("Intake", "Intake"), "c_meet": ("Meetings", "미팅"), "c_done": ("Done", "완료"),
    "c_out": ("Outcomes", "성과"),
    "ik_none": ("Add startups to the program first (Startups tab).", "먼저 참가사 탭에서 스타트업을 넣어주세요."),
    "ik_pick": ("Startup", "스타트업"), "ik_progress": ("{d} of {n} interviews done", "{n}곳 중 {d}곳 인터뷰 완료"),
    "ik_date": ("Date", "일시"), "ik_by": ("Interviewer", "면담자"), "ik_with": ("Attendees (startup side)", "참석자 (스타트업)"),
    "ik_status": ("Status", "상태"), "ik_todo": ("Planned", "예정"), "ik_done": ("Done", "완료"),
    "ik_notes": ("Other notes", "기타 메모"), "ik_questions": ("Questions", "질문 관리"),
    "ik_q_help": ("The same questions are used for every startup in this program.", "이 프로그램의 모든 스타트업에 같은 질문을 써요."),
    "q_section": ("Section", "구분"), "q_text": ("Question", "질문"),
    "ik_ai": ("✨ Summarize with AI", "✨ AI로 정리"), "ik_ai_off": ("AI summary is off (add an API key, see README).",
                                                              "AI 정리는 꺼져 있어요 (API 키를 넣으면 켜져요, README 참고)."),
    "ik_summary": ("Summary", "요약"), "ik_suggest": ("Suggested company details", "회사 정보 제안"),
    "ik_apply": ("Apply to the company directory", "회사 정보에 반영"), "ik_applied": ("Company details updated.", "회사 정보를 갱신했어요."),
    "ik_now": ("now: {v}", "지금: {v}"), "ik_doc": ("⬇ This startup (.hwpx)", "⬇ 이 기업 한글 파일"),
    "ik_doc_all": ("⬇ All startups (.hwpx)", "⬇ 전체 기업 한글 파일"),
    "ik_summary_help": ("Edit freely. It goes into the intake file and the reports.", "자유롭게 고쳐도 돼요. intake 파일과 보고서에 들어가요."),
    "ai_fail": ("The AI call failed: {e}", "AI 호출이 실패했어요: {e}"),
    "kpi_title": ("Indicators", "성과 지표"), "kpi_help": ("Auto indicators count from meetings, interviews and the outcome log. "
                                                        "Manual ones use the number you type.",
                                                        "자동 지표는 미팅, 인터뷰, 성과 기록에서 세고, 수동 지표는 적은 숫자를 써요."),
    "k_name": ("Indicator", "지표"), "k_target": ("Target", "목표"), "k_unit": ("Unit", "단위"),
    "k_source": ("Counts from", "집계 방식"), "k_manual": ("Manual value", "수동 값"),
    "k_actual": ("Actual", "실적"), "k_rate": ("Achieved", "달성률"),
    "out_title": ("Outcome log", "성과 기록"), "out_help": ("Follow-up meetings, MOUs, PoCs, investments… one row each.",
                                                        "후속 미팅, MOU, PoC, 투자 등을 한 줄씩 적어요."),
    "o_date": ("Date", "일자"), "o_startup": ("Startup", "스타트업"), "o_partner": ("Partner", "파트너"),
    "o_type": ("Type", "유형"), "o_amount": ("Amount", "금액"), "o_note": ("Notes", "메모"),
    "m_title": ("Business matching", "비즈매칭"), "m_total": ("Meetings", "전체 미팅"), "m_done": ("Done", "완료"),
    "m_noshow": ("No-shows", "노쇼"), "m_up": ("Upcoming", "예정"), "m_partners": ("Partners", "파트너"),
    "rp_kind": ("Document", "문서"), "rp_proposal": ("Proposal and quote", "제안서·견적"),
    "rp_interim": ("Interim report", "중간 보고서"), "rp_final": ("Final report", "최종 보고서"),
    "rp_doc_proposal": ("Proposal", "사업 제안서"), "rp_doc_interim": ("Interim report", "중간 보고서"),
    "rp_doc_final": ("Final report", "최종 보고서"),
    "rp_ai_wait": ("Writing drafts… this can take a minute.", "초안 쓰는 중이에요. 1분 정도 걸릴 수 있어요."),
    "rp_preview_btn": ("Preview pages", "쪽 미리보기"),
    "rp_pv_old": ("The text changed after this preview. Click Preview pages again to update it.",
                  "미리보기 뒤에 내용이 바뀌었어요. 쪽 미리보기를 다시 누르면 새로 그려요."),
    "qt_title": ("Quote", "견적"), "qt_help": ("Line items for the proposal. Amount = unit price × quantity.",
                                              "제안서에 들어갈 견적 항목이에요. 금액은 단가 × 수량으로 계산돼요."),
    "qt_currency": ("Currency", "통화"), "qt_vat": ("Add VAT (10%)", "부가세 10% 포함"),
    "qt_cat": ("Category", "구분"), "qt_item": ("Item", "내역"), "qt_price": ("Unit price", "단가"),
    "qt_qty": ("Qty", "수량"), "qt_unit": ("Unit", "단위"), "qt_note": ("Notes", "비고"),
    "qt_sub": ("Subtotal", "소계"), "qt_vatv": ("VAT", "부가세"), "qt_total": ("Total", "합계"),
    "rp_title": ("Title", "제목"), "rp_asof": ("As of", "기준일"), "rp_byline": ("Byline (under the title)", "제목 아래 줄"),
    "rp_fill_all": ("Refill every section from the data", "모든 섹션 데이터로 다시 채우기"),
    "rp_ai_all": ("✨ Draft every section with AI", "✨ 모든 섹션 AI 초안"),
    "rp_fill": ("Refill from data", "데이터로 채우기"), "rp_ai": ("✨ AI draft", "✨ AI 초안"),
    "rp_help": ("Each section starts filled from the program data. Edit the text freely (Markdown: '- ' bullets, "
                "two spaces for sub bullets, | tables |).", "섹션마다 프로그램 데이터로 먼저 채워져요. 자유롭게 고치면 돼요 "
                "(Markdown: '- ' 글머리표, 두 칸 들여쓰면 하위 항목, | 표 |)."),
    "rp_download": ("⬇ 한글 file (.hwpx)", "⬇ 한글 파일 받기"),
    "rp_hwp_note": ("Opens in 한글 2014 or later. If the client needs .hwp, open it in 한글 and Save As HWP.",
                    "한글 2014 이상에서 열려요. 발주처가 .hwp를 원하면 한글에서 열고 '다른 이름으로 저장'에서 HWP를 고르면 돼요."),
    "rp_preview": ("Page preview (approximate: 한글 draws the □ and - bullets, this preview does not)",
                   "쪽 미리보기 (대략적인 모습이에요. □, - 글머리표는 한글에서 열면 보여요)"),
    "rp_issues": ("The file check found problems: {e}", "파일 점검에서 문제가 나왔어요: {e}"),
    "rp_ai_done": ("AI drafts are in. Check numbers and names before sending.", "AI 초안을 넣었어요. 보내기 전에 숫자와 이름을 꼭 확인해 주세요."),
    "conflict": ("{who} changed the same thing at the same moment, so their version was kept there. Please check it.",
                 "{who} 님이 같은 부분을 같은 순간에 고쳐서 그 부분은 그쪽 내용으로 두었어요. 한 번 확인해 주세요."),
    "saved_by": ("Saved {when} · {who}", "저장됨 {when} · {who}"),
}

STATUS = {"todo": ("Planned", "예정"), "doing": ("In progress", "진행 중"), "done": ("Done", "완료"), "hold": ("On hold", "보류")}
STATUS_COLORS = {"todo": "#cbd2e0", "doing": "#88a1eb", "done": "#7cc49a", "hold": "#f2b38b"}
OUTCOMES = {"followup": ("Follow-up meeting", "후속 미팅"), "nda": ("NDA", "NDA"), "mou": ("MOU / LOI", "MOU·LOI"),
            "poc": ("PoC", "PoC"), "inv_review": ("Investment review", "투자 검토"),
            "inv_closed": ("Investment closed", "투자 유치"), "contract": ("Contract / sales", "계약·매출"),
            "other": ("Other", "기타")}
SOURCES = {"": ("Manual", "수동 입력"), "meetings_done": ("Matching: meetings done", "비즈매칭 완료 미팅 수"),
           "meetings_all": ("Matching: all meetings", "비즈매칭 전체 미팅 수"), "intake_done": ("Intake interviews done", "Intake 완료 기업 수"),
           **{f"outcome:{k}": (f"Outcome log: {v[0]}", f"성과 기록: {v[1]}") for k, v in OUTCOMES.items()},
           "amount:inv_closed": ("Outcome log: investment amount", "성과 기록: 투자 유치 금액 합계"),
           "amount:contract": ("Outcome log: contract amount", "성과 기록: 계약·매출 금액 합계")}


def label_of(table: dict, code: str) -> str:
    return L2(table.get(code, (code, code)))


def code_of(table: dict, label) -> str:
    for k, (en, ko) in table.items():
        if label in (en, ko, k):
            return k
    return "" if label is None else str(label)


# ---------------------------------------------------------------- templates
TEMPLATES = {
    "accel": (("Global acceleration", "글로벌 액셀러레이팅"), [
        (("Proposal and quote", "제안서·견적 제출"), 0.00, 0.04, ["제안서 작성", "견적서 작성", "제출"]),
        (("Contract", "계약 체결"), 0.04, 0.08, ["계약서 검토", "서명·날인"]),
        (("Startup selection", "참가기업 선발"), 0.06, 0.15, ["모집 공고", "서류 심사", "발표 평가", "선정 통보"]),
        (("Kick-off", "킥오프·OT"), 0.15, 0.18, ["일정 공지", "OT 자료 준비"]),
        (("Intake interviews", "Intake interview"), 0.16, 0.28, ["질문지 확정", "인터뷰 일정 잡기", "결과 정리"]),
        (("Workshops and mentoring", "워크숍·멘토링"), 0.25, 0.75, ["연사·멘토 섭외", "세션 운영", "참석 기록"]),
        (("Business matching", "비즈매칭"), 0.40, 0.80, ["파트너 섭외", "매칭 일정 생성", "행사 운영", "후속 성과 조사"]),
        (("Interim report", "중간 보고"), 0.48, 0.52, ["실적 집계", "보고서 작성", "제출"]),
        (("Demo day / IR", "데모데이·IR"), 0.80, 0.88, ["피칭 코칭", "심사위원 섭외", "행사 운영"]),
        (("Final report", "최종 보고"), 0.90, 0.98, ["성과 조사", "보고서 작성", "제출"]),
        (("Settlement", "정산"), 0.95, 1.00, ["증빙 정리", "정산 보고"]),
    ]),
    "expo": (("Exhibition / CES delegation", "전시회·CES 참가 지원"), [
        (("Proposal and contract", "제안·계약"), 0.00, 0.08, ["제안서·견적 제출", "계약 체결"]),
        (("Confirm startups", "참가기업 확정"), 0.05, 0.15, ["참가 확정", "참가비·서류 수합"]),
        (("Intake interviews", "Intake interview"), 0.12, 0.25, ["질문지 확정", "인터뷰", "결과 정리"]),
        (("Booth and logistics", "부스·물류 준비"), 0.15, 0.60, ["부스 디자인", "물품 운송", "숙소·항공"]),
        (("Pre-arranged matching", "사전 비즈매칭 섭외"), 0.25, 0.70, ["파트너 리스트업", "미팅 요청", "매칭 일정 생성"]),
        (("Interim report", "중간 보고"), 0.45, 0.50, ["실적 집계", "보고서 작성", "제출"]),
        (("On-site schedule", "현지 일정 확정"), 0.60, 0.75, ["현장 운영 계획", "비상 연락망"]),
        (("Show days and matching", "전시 운영·현장 비즈매칭"), 0.78, 0.86, ["부스 운영", "미팅 진행", "현장 사진·기록"]),
        (("Follow-ups", "후속 연결·성과 조사"), 0.86, 0.95, ["후속 미팅 연결", "성과 설문"]),
        (("Final report", "최종 보고"), 0.92, 0.99, ["성과 집계", "보고서 작성", "제출"]),
        (("Settlement", "정산"), 0.95, 1.00, ["증빙 정리", "정산 보고"]),
    ]),
    "matching": (("Business matching event", "비즈매칭 행사"), [
        (("Partner outreach", "파트너 섭외"), 0.00, 0.40, ["리스트업", "초청 메일", "참가 확정"]),
        (("Startup registration", "참가사 접수"), 0.10, 0.45, ["신청 접수", "참가 확정"]),
        (("Intake and needs", "Intake·니즈 파악"), 0.30, 0.55, ["인터뷰", "희망 파트너 정리"]),
        (("Build the schedule", "매칭 일정 생성"), 0.55, 0.70, ["일정 생성", "참가자 공지"]),
        (("Event day", "행사 운영"), 0.70, 0.75, ["현장 운영", "노쇼 대체"]),
        (("Follow-up survey", "후속 성과 조사"), 0.75, 0.92, ["설문", "후속 미팅 확인"]),
        (("Result report", "결과 보고"), 0.92, 1.00, ["보고서 작성", "제출"]),
    ]),
    "blank": (("Blank", "빈 프로그램"), []),
}

QUESTIONS = [  # (section, question) in (en, ko)
    (("Company", "기업 현황"), ("Main product or service, and key customers?", "주요 제품·서비스와 핵심 고객은?")),
    (("Company", "기업 현황"), ("Current stage: revenue, customers, key metrics?", "현재 단계 (매출, 고객 수, 주요 지표)는?")),
    (("Company", "기업 현황"), ("Team and overseas experience?", "팀 구성과 해외 경험은?")),
    (("Global", "글로벌 진출"), ("Target markets and progress so far?", "목표 시장과 지금까지 진출 현황은?")),
    (("Global", "글로벌 진출"), ("What do you most want from this program?", "이번 프로그램에서 가장 얻고 싶은 것은?")),
    (("Needs", "니즈"), ("Fundraising plans: round, amount, timing?", "투자 유치 계획 (라운드, 금액, 시기)은?")),
    (("Needs", "니즈"), ("Partners you want to meet (type or named companies)?", "만나고 싶은 파트너 (기업·투자자·기관 유형이나 구체 기업)는?")),
    (("Needs", "니즈"), ("Support needed: entity, certification, localization, hiring…?", "필요한 지원 (법인, 인증, 현지화, 채용 등)은?")),
    (("Other", "기타"), ("Main risks or difficulties?", "주요 리스크나 애로사항은?")),
    (("Other", "기타"), ("Anything else?", "기타 요청사항은?")),
]

KPIS = [  # (name en, name ko, target, unit en, unit ko, source)
    ("Business matching meetings done", "비즈매칭 미팅 (완료)", 100, "meetings", "건", "meetings_done"),
    ("Intake interviews", "Intake interview", None, "startups", "개사", "intake_done"),
    ("Follow-up meetings", "후속 미팅", 20, "", "건", "outcome:followup"),
    ("MOU / LOI", "MOU·LOI", 5, "", "건", "outcome:mou"),
    ("PoC", "PoC", 3, "", "건", "outcome:poc"),
    ("Investments closed", "투자 유치", 2, "", "건", "outcome:inv_closed"),
    ("Startup satisfaction", "참가기업 만족도", 4.5, "/ 5", "점", ""),
]


def new_id(prefix: str = "") -> str:
    return prefix + uuid.uuid4().hex[:10]


def make_program(name: str, tpl: str, start: date | None, end: date | None, client: str = "",
                 startups: list | None = None) -> dict:
    ko = lang() == "ko"
    start = start or date.today()
    end = end if end and end > start else start + timedelta(days=90)
    span = (end - start).days
    stages = []
    for (names, a, b, ck) in TEMPLATES.get(tpl, TEMPLATES["blank"])[1]:
        stages.append({"id": new_id("s"), "name": names[1] if ko else names[0],
                       "start": (start + timedelta(days=round(span * a))).isoformat(),
                       "due": (start + timedelta(days=max(round(span * b), round(span * a)))).isoformat(),
                       "owner": "", "status": "todo", "note": "",
                       "checklist": [{"text": x, "done": False} for x in ck]})
    kpis = [] if tpl == "blank" else [
        {"id": new_id("k"), "name": ko and k[1] or k[0], "target": k[2] if k[2] is not None else len(startups or []) or None,
         "unit": ko and k[4] or k[3], "source": k[5], "manual": None} for k in KPIS]
    return {"name": name, "client": client, "template": tpl, "start": start.isoformat(), "end": end.isoformat(),
            "pm": API["me"]() if API.get("me") else "", "team": "", "budget": "", "summary": "",
            "startups": list(startups or []), "stages": stages, "kpis": kpis, "outcomes": [], "events": [],
            "questions": [{"id": f"q{i + 1}", "section": L2(sec), "text": L2(q)} for i, (sec, q) in enumerate(QUESTIONS)]}


# ---------------------------------------------------------------- storage: team database or this browser
class Docs:
    def __init__(self):
        self.S = API["store"]() if API.get("store") else None

    @property
    def local(self) -> dict:
        return ss().setdefault("local_docs", {})

    def list(self, kind: str, parent: str | None = None, with_data: bool = False) -> list[dict]:
        if self.S:
            rows = self.S.doc_list(kind, parent, with_data)
            for r in rows:
                if with_data:
                    r["data"] = json.loads(r.get("data") or "{}")
            return rows
        out = []
        for k, v in self.local.items():
            kd, _, id_ = k.partition("|")
            if kd == kind and not v.get("archived") and (parent is None or v.get("parent") == parent):
                out.append({"id": id_, "name": v.get("name", ""), "parent": v.get("parent", ""),
                            "version": v["version"], "updated_at": v.get("updated_at"), "updated_by": v.get("updated_by"),
                            **({"data": copy.deepcopy(v["data"])} if with_data else {})})
        return sorted(out, key=lambda r: -(r.get("updated_at") or 0))

    def get(self, kind: str, id_: str) -> tuple[dict | None, int, dict]:
        if self.S:
            r = self.S.doc_get(kind, id_)
            if not r or r.get("archived"):
                return None, 0, {}
            return json.loads(r["data"]), r["version"], r
        v = self.local.get(f"{kind}|{id_}")
        if not v or v.get("archived"):
            return None, 0, {}
        return copy.deepcopy(v["data"]), v["version"], v

    def version(self, kind: str, id_: str) -> int:
        if self.S:
            return self.S.doc_version(kind, id_)
        v = self.local.get(f"{kind}|{id_}")
        return v["version"] if v else 0

    def put(self, kind: str, id_: str, data: dict, version: int, parent: str = "", name: str = "") -> int | None:
        who = API["me"]() if API.get("me") else ""
        if self.S:
            return self.S.doc_put(kind, id_, json.dumps(data, ensure_ascii=False, default=str), version, who,
                                  parent, name)
        k = f"{kind}|{id_}"
        cur = self.local.get(k)
        if (cur["version"] if cur else 0) != version:
            return None
        self.local[k] = {"data": copy.deepcopy(data), "version": version + 1, "parent": parent, "name": name,
                         "updated_at": time.time(), "updated_by": who, "archived": False}
        return version + 1

    def archive(self, kind: str, id_: str) -> None:
        if self.S:
            self.S.doc_archive(kind, id_, API["me"]() if API.get("me") else "")
        elif f"{kind}|{id_}" in self.local:
            self.local[f"{kind}|{id_}"]["archived"] = True


def h(data: dict) -> str:
    return hashlib.sha1(json.dumps(data, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()


class Working:
    """An open document. `data` is edited live. A widget starts from `data` whenever Streamlit holds no value
    for it (first draw, or a value Streamlit dropped), and `gen` is part of every widget key: when the data is
    replaced from outside the widgets (a reload, an AI draft, a refill) `rebase` bumps it so every widget starts
    over from the new data."""

    def __init__(self, kind: str, id_: str, parent: str = ""):
        self.kind, self.id, self.parent = kind, id_, parent
        self.data, self.version, self.meta = Docs().get(kind, id_)
        self.saved = h(self.data) if self.data is not None else None
        self.loaded = copy.deepcopy(self.data)   # the document as last read or written (for merging)
        self.gen = int(time.time() * 1000) % 10_000_000
        self.reg: dict = {}   # widget key -> how its value goes into data (see text_box, table_box, capture)
        self.epochs: dict = {}   # table -> restarts after the browser dropped its edits (see table_box)

    def rebase(self, keep_last: bool = False):
        """New widget keys, so every widget starts over from the data. `keep_last` keeps the bindings of the
        keys being retired for one more run: after a page switch Streamlit can still deliver a value typed
        just before the click, under the old key."""
        old = self.gen
        self.gen += 1
        self.reg = {k: v for k, v in self.reg.items() if v["gen"] == old} if keep_last else {}
        self.epochs = {}

    def reload(self):
        self.data, self.version, self.meta = Docs().get(self.kind, self.id)
        self.saved = h(self.data) if self.data is not None else None
        self.loaded = copy.deepcopy(self.data)
        self.rebase()

    def refresh_if_stale(self):
        """Someone else saved a newer version and we have nothing unsaved: take theirs."""
        docs = Docs()
        if docs.S and self.data is not None and h(self.data) == self.saved and docs.version(self.kind, self.id) > self.version:
            self.reload()

    def save(self, name: str = "") -> None:
        if self.data is None or h(self.data) == self.saved:
            return
        docs = Docs()
        v = docs.put(self.kind, self.id, self.data, self.version, self.parent, name)
        if v:
            self._written(v)
            return
        # A teammate saved in between. Put our changes on top of theirs; where both changed the same thing,
        # theirs stays and we say so.
        for _ in range(3):
            theirs, tv, meta = docs.get(self.kind, self.id)
            if theirs is None:
                self.reload()
                return
            merged, clash = merge3(self.loaded if self.loaded is not None else {}, self.data, theirs)
            if merged == theirs:      # nothing of ours left to add: just take theirs
                self.data, self.version, self.meta = theirs, tv, meta
                self.saved, self.loaded = h(theirs), copy.deepcopy(theirs)
                self.rebase()
                ss()["_pg_rerun"] = True
                if clash:
                    ss()["pg_flash"] = ("warning", T("conflict", who=(meta or {}).get("updated_by") or "?"))
                return
            v = docs.put(self.kind, self.id, merged, tv, self.parent, name)
            if v:
                self.data = merged
                self._written(v)
                self.rebase()
                ss()["_pg_rerun"] = True   # draw the widgets again from the merged document
                if clash:
                    ss()["pg_flash"] = ("warning", T("conflict", who=(meta or {}).get("updated_by") or "?"))
                return
        self.reload()
        ss()["pg_flash"] = ("warning", T("conflict", who="?"))

    def _written(self, v: int) -> None:
        self.version, self.saved, self.loaded = v, h(self.data), copy.deepcopy(self.data)
        self.meta = {"updated_by": API["me"]() if API.get("me") else "", "updated_at": time.time()}


_MISSING = object()


def merge3(base, mine, theirs):
    """Three way merge of JSON values: what we changed since `base` goes on top of `theirs`, key by key (nested
    dicts too). A key both sides changed differently keeps their value and is reported as a clash."""
    if not (isinstance(base, dict) and isinstance(mine, dict) and isinstance(theirs, dict)):
        if mine == base or theirs == mine:
            return theirs, []
        if theirs == base:
            return mine, []
        return theirs, ["*"]
    out, clash = dict(theirs), []
    for k in set(base) | set(mine):
        b, m, t = base.get(k, _MISSING), mine.get(k, _MISSING), theirs.get(k, _MISSING)
        if m == b:
            continue
        if t == b or t == m:
            if m is _MISSING:
                out.pop(k, None)
            else:
                out[k] = m
        elif isinstance(b, dict) and isinstance(m, dict) and isinstance(t, dict):
            out[k], sub = merge3(b, m, t)
            clash += [f"{k}.{x}" for x in sub]
        else:
            clash.append(k)
    return out, clash


def wc(slot: str, kind: str, id_: str, parent: str = "", default: dict | None = None) -> Working:
    """The working copy in session slot `slot` for document (kind, id); created from `default` when new
    (and only written once something in it changes)."""
    w = ss().get(slot)
    if not isinstance(w, Working) or (w.kind, w.id) != (kind, id_):
        w = Working(kind, id_, parent)
        if w.data is None and default is not None:
            w.data, w.version = copy.deepcopy(default), 0
            w.saved, w.loaded = h(w.data), copy.deepcopy(default)
        ss()[slot] = w
    else:
        w.refresh_if_stale()
    return w


SLOTS = ("pg_w", "ik_w", "rp_w")


def txt(v) -> str:
    """A box or cell value as clean text ('' for None, NaN and the like)."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def num(v):
    if v is None or v == "":
        return None
    try:
        if pd.isna(v):
            return None
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f == int(f) else f


def flag(v) -> bool:
    if v is None:
        return False
    try:
        if pd.isna(v):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        return v.strip().lower() in ("true", "1", "yes", "y", "o", "✓", "✅")
    return bool(v)


def _get(slot: str, data: dict, f: str) -> str:
    if slot == "ik_w":
        if f in ("by", "with"):
            return data.get("interviewer" if f == "by" else "attendees") or ""
        if f in ("notes", "summary"):
            return data.get(f) or ""
        return (data.get("answers") or {}).get(f) or ""
    if slot == "rp_w" and f not in ("title", "byline"):
        return (data.get("sections") or {}).get(f) or ""
    return data.get(f) or ""


def _put(slot: str, data: dict, f: str, v: str) -> None:
    if slot == "pg_w":
        if f != "name" or v.strip():
            data[f] = v
    elif slot == "ik_w":
        if f in ("by", "with"):
            data["interviewer" if f == "by" else "attendees"] = v
        elif f in ("notes", "summary"):
            data[f] = v
        else:
            data.setdefault("answers", {})[f] = v
    elif f in ("title", "byline"):
        data[f] = v
    else:
        data.setdefault("sections", {})[f] = v


def wkey(slot: str, f: str) -> str:
    return f"{slot[:2]}_{ss()[slot].gen}_{f}"


def text_box(slot: str, where, label: str, f: str, area: bool = False, **kw) -> str:
    """A text box bound to field `f` of the working copy in `slot`. Its default is the live value, so whenever
    the browser has no value of its own for the box (first draw, or it dropped it) it shows the document."""
    w = ss()[slot]
    key = wkey(slot, f)
    v = (where.text_area if area else where.text_input)(label, value=_get(slot, w.data, f), key=key, **kw)
    v = v if isinstance(v, str) else ""
    _put(slot, w.data, f, v)
    w.reg[key] = {"kind": "text", "f": f, "gen": w.gen}
    return v


def date_box(slot: str, where, label: str, f: str, data: dict) -> str:
    v = where.date_input(label, value=pdate(data.get(f)), key=wkey(slot, f))
    data[f] = v.isoformat() if isinstance(v, date) else ""
    return data[f]


def drop_if_invalid(key: str, ok) -> None:
    """Forget a picker's value that is no longer among its options, so it starts from its default."""
    if key in ss() and not ok(ss()[key]):
        del ss()[key]


def _delta(v) -> dict:
    v = v if isinstance(v, dict) else {}
    return {"edited_rows": {str(k): dict(x or {}) for k, x in (v.get("edited_rows") or {}).items()},
            "added_rows": [dict(x) for x in v.get("added_rows") or [] if isinstance(x, dict)],
            "deleted_rows": sorted(int(x) for x in v.get("deleted_rows") or [] if str(x).lstrip("-").isdigit())}


def lost_edits(last: dict | None, cur: dict) -> bool:
    """True when the browser's copy of a table no longer holds edits it sent before: it was drawn again from
    scratch (a page switch can race an edit). Applying it would undo those edits, so the table starts over
    from the document instead."""
    if not last:
        return False
    le, ce = last["edited_rows"], cur["edited_rows"]
    if any(not set(cols) <= set(ce.get(r, {})) for r, cols in le.items()):
        return True
    if set(last["deleted_rows"]) - set(cur["deleted_rows"]):
        return True
    had = bool(le or last["added_rows"] or last["deleted_rows"])
    return had and not (ce or cur["added_rows"] or cur["deleted_rows"])


def _has_edits(d: dict | None) -> bool:
    return bool(d and (d["edited_rows"] or d["added_rows"] or d["deleted_rows"]))


def _browser_ids() -> set | None:
    """IDs of the widgets whose values the browser sent for this run (None when it sent none, e.g. a page
    change or st.rerun, or when Streamlit's internals can't be read)."""
    try:
        from streamlit.runtime.state.session_state_proxy import get_session_state
        fresh = get_session_state()._state._new_widget_state.states
        return set(fresh) if fresh else None
    except Exception:  # noqa: BLE001
        return None


def _dropped(key: str, e: dict) -> bool:
    """The browser sent values this run but none for this table, although the table holds edits: it drew the
    table again from scratch and shows it without them."""
    sent = ss().get("_pg_sent")
    if not sent or not _has_edits(e.get("last")):
        return False
    try:
        from streamlit.runtime.state.session_state_proxy import get_session_state
        wid = get_session_state()._state._key_id_mapper.get_id_from_key(key, None)
    except Exception:  # noqa: BLE001
        return False
    return wid is not None and wid not in sent


def table_box(slot: str, where, f: str, rows_of, cols: list[str], to_doc, **kw) -> list[dict]:
    """A table editor. It starts from `rows_of(data)` and keeps those starting rows while it lives, because the
    edits Streamlit stores for it are relative to them; `to_doc(data, records)` writes its rows back. When the
    browser lost its copy of the edits, the table starts over (new key) from the document, which has them."""
    w = ss()[slot]
    key = f"{wkey(slot, f)}_{w.epochs.get(f, 0)}"
    e = w.reg.get(key)
    if e and key in ss() and (e.get("restart") or lost_edits(e.get("last"), _delta(ss()[key])) or _dropped(key, e)):
        w.epochs[f] = w.epochs.get(f, 0) + 1
        key, e = f"{wkey(slot, f)}_{w.epochs[f]}", None
    rows = e["rows"] if e and key in ss() else rows_of(w.data)
    ed = where.data_editor(pd.DataFrame(rows, columns=cols), key=key, **kw)
    recs = ed.to_dict("records")
    to_doc(w.data, recs)
    w.reg[key] = {"kind": "table", "to_doc": to_doc, "rows": rows, "cols": cols, "rows_of": rows_of, "gen": w.gen,
                  "last": _delta(ss().get(key)), "restart": False}
    return recs


def apply_delta(rows: list[dict], delta: dict, cols: list[str]) -> list[dict]:
    """What a table editor shows: its starting rows with the edits it keeps in session state
    (cell edits, then deleted rows, then added rows, the order Streamlit uses)."""
    out = [dict(r) for r in rows]
    for i, ch in (delta.get("edited_rows") or {}).items():
        try:
            i = int(i)
        except (TypeError, ValueError):
            continue
        if 0 <= i < len(out) and isinstance(ch, dict):
            out[i].update(ch)
    gone = set()
    for i in delta.get("deleted_rows") or []:
        try:
            gone.add(int(i))
        except (TypeError, ValueError):
            pass
    out = [r for i, r in enumerate(out) if i not in gone]
    out += [{c: a.get(c) for c in cols} for a in delta.get("added_rows") or [] if isinstance(a, dict)]
    return out


def apply_onto(snap: list[dict], now: list[dict], delta: dict, cols: list[str]) -> list[dict]:
    """Edits made against the starting rows `snap`, put on the rows the document has `now` (matched by id when
    the table has one, else by position)."""
    rows = [dict(r) for r in now]
    if "id" in cols:
        pos = {r.get("id"): j for j, r in enumerate(rows) if r.get("id")}

        def at(i: int):
            return pos.get(snap[i].get("id")) if 0 <= i < len(snap) else None
    else:
        def at(i: int):
            return i if 0 <= i < min(len(snap), len(rows)) else None
    for i, ch in (delta.get("edited_rows") or {}).items():
        try:
            j = at(int(i))
        except (TypeError, ValueError):
            continue
        if j is not None and isinstance(ch, dict):
            rows[j].update(ch)
    gone = set()
    for i in delta.get("deleted_rows") or []:
        try:
            gone.add(at(int(i)))
        except (TypeError, ValueError):
            pass
    rows = [r for j, r in enumerate(rows) if j not in gone]
    rows += [{c: a.get(c) for c in cols} for a in delta.get("added_rows") or [] if isinstance(a, dict)]
    return rows


def begin_run() -> None:
    """app.py calls this first in every run, before anything is drawn."""
    ss()["_pg_sent"] = _browser_ids()
    capture()


def capture(slots=SLOTS) -> None:
    """Copy the latest box and table values of the open documents into their working copies, then save.

    app.py calls this first in every run, before anything is drawn. When a click leaves the page right after
    typing, Streamlit may run the new page first and send the typed value a moment later in a run of its own,
    or cut the run the box started short; the start of a run is the one place such a value is always seen."""
    for slot in slots:
        w = ss().get(slot)
        if not isinstance(w, Working) or w.data is None or not w.reg:
            continue
        for key, e in list(w.reg.items()):
            if key not in ss():
                continue
            v = ss()[key]
            if e["kind"] == "text":
                if isinstance(v, str):
                    _put(slot, w.data, e["f"], v)
            elif isinstance(v, dict):
                cur = _delta(v)
                try:
                    if lost_edits(e.get("last"), cur):
                        # drawn again from scratch: its new edits go on top of the document, then it starts over
                        e["to_doc"](w.data, apply_onto(e["rows"], e["rows_of"](w.data), v, e["cols"]))
                        e["restart"] = True
                    else:
                        e["to_doc"](w.data, apply_delta(e["rows"], v, e["cols"]))
                    e["last"] = cur
                except Exception:  # noqa: BLE001
                    pass
        w.save(w.data.get("name") or w.data.get("startup") or w.data.get("title") or "")
    ss().pop("_pg_rerun", None)   # nothing is drawn yet, so a merged document needs no extra run


def harvest_all() -> None:
    capture()


def on_page_switch():
    """Called by app.py when the person moves between pages: keep what was typed last, then give every widget a
    new key so it is drawn again from the live document (the browser may have dropped its own copy)."""
    capture()
    for slot in SLOTS:
        w = ss().get(slot)
        if isinstance(w, Working) and w.data is not None:
            w.rebase(keep_last=True)


# ---------------------------------------------------------------- helpers
def pdate(x) -> date | None:
    if isinstance(x, date):
        return x
    try:
        return date.fromisoformat(str(x)[:10]) if x else None
    except ValueError:
        return None


def iso(v) -> str:
    """A date cell from a table editor as YYYY-MM-DD ('' for empty, NaT or junk)."""
    try:
        if v is None or pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    d = pdate(v.date() if hasattr(v, "date") and callable(v.date) else v)
    return d.isoformat() if d else ""


def fmt_num(v) -> str:
    if v is None or v == "":
        return "-"
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{int(f):,}" if f == int(f) else f"{f:,.1f}"


def stage_summary(d: dict) -> dict:
    st_ = d.get("stages") or []
    done = sum(1 for s in st_ if s.get("status") == "done")
    now = next((s for s in st_ if s.get("status") == "doing"), None)
    nxt = sorted([s for s in st_ if s.get("status") != "done" and pdate(s.get("due"))], key=lambda s: s["due"])
    return {"n": len(st_), "done": done, "now": now, "next": nxt[0] if nxt else None}


def dday(due: date) -> str:
    n = (due - date.today()).days
    return T("today") if n == 0 else T("dday", n=n) if n > 0 else T("overdue", n=-n)


def directory_startups() -> list[dict]:
    return API["dir_rows"]("startup") if API.get("dir_rows") else []


def startup_rows(d: dict) -> list[dict]:
    """The program's startups with their directory details (names that left the directory are kept)."""
    by = {norm(API["dir_name"]("startup", r)): r for r in directory_startups()}
    out = []
    for n in d.get("startups") or []:
        r = by.get(norm(n), {})
        out.append({"key": norm(n), "name": n, "kr": r.get("Company (KR)", ""), "tags": r.get("Tags", ""),
                    "ceo": r.get("CEO", ""), "web": r.get("Website", ""), "stage": r.get("Stage", ""),
                    "one_liner": r.get("One-liner", "")})
    return out


def _cache() -> dict:
    return ss().setdefault("_pg_cache", {})


def matching_for(d: dict) -> dict:
    keys = {norm(n) for n in d.get("startups") or []}
    ck = ("m", tuple(d.get("events") or []))
    if ck not in _cache():
        _cache()[ck] = API["matching_stats"](d.get("events") or []) if API.get("matching_stats") else None
    stats = _cache()[ck]
    if not stats:
        return {"events": [], "total": 0, "by": {}, "partners": set(), "startups": {}}
    if keys:   # only this program's startups
        sts = {k: v for k, v in stats["startups"].items() if k in keys}
        by: dict = {}
        for v in sts.values():
            for s_, n in v["by"].items():
                by[s_] = by.get(s_, 0) + n
        partners = set().union(*[set(v["partners"]) for v in sts.values()]) if sts else set()
        return {"events": stats["events"], "total": sum(v["total"] for v in sts.values()), "by": by,
                "partners": partners, "startups": sts}
    return stats


def intake_map(pid: str) -> dict:
    ck = ("ik", pid)
    if ck not in _cache():
        _cache()[ck] = {norm(r["data"].get("startup")): r["data"] for r in Docs().list("intake", pid, with_data=True)
                        if r.get("data")}
    out = dict(_cache()[ck])
    w = ss().get("ik_w")
    if isinstance(w, Working) and w.parent == pid and w.data:     # the open interview as it is right now
        out[norm(w.data.get("startup"))] = w.data
    return out


def kpi_values(d: dict, match: dict, intakes: dict) -> list[dict]:
    outs = d.get("outcomes") or []
    rows = []
    for k in d.get("kpis") or []:
        src = k.get("source") or ""
        if src == "meetings_done":
            actual = match["by"].get("Done", 0)
        elif src == "meetings_all":
            actual = match["total"]
        elif src == "intake_done":
            actual = sum(1 for v in intakes.values() if v.get("status") == "done")
        elif src.startswith("outcome:"):
            actual = sum(1 for o in outs if o.get("type") == src.split(":", 1)[1])
        elif src.startswith("amount:"):
            actual = sum(float(o.get("amount") or 0) for o in outs if o.get("type") == src.split(":", 1)[1])
        else:
            actual = k.get("manual")
        try:
            rate = float(actual) / float(k["target"]) if k.get("target") not in (None, "", 0) and actual not in (None, "") else None
        except (TypeError, ValueError, ZeroDivisionError):
            rate = None
        unit = k.get("unit") or ""
        rows.append({**k, "actual": actual, "rate": rate,
                     "target_s": f"{fmt_num(k.get('target'))}{unit if unit and not unit.startswith('/') else ''}"
                     if k.get("target") not in (None, "") else "-",
                     "actual_s": f"{fmt_num(actual)}{unit if unit and not unit.startswith('/') else ''}"
                     if actual not in (None, "") else "-",
                     "rate_s": f"{rate * 100:.1f}%" if rate is not None else "-"})
    return rows


# ---------------------------------------------------------------- table rows -> document
STAGE_COLS = ["id", "status", "name", "start", "due", "owner", "note"]


def _rid(r: dict, prefix: str, g: int, i: int) -> str:
    v = r.get("id")
    return v if isinstance(v, str) and v else f"{prefix}{g}n{i}"


def stage_rows(data: dict) -> list[dict]:
    return [{"id": s_.get("id"), "status": label_of(STATUS, s_.get("status", "todo")), "name": s_.get("name", ""),
             "start": pdate(s_.get("start")), "due": pdate(s_.get("due")), "owner": s_.get("owner", ""),
             "note": s_.get("note", "")} for s_ in data.get("stages") or []]


def stages_to_doc(g: int):
    def conv(data: dict, recs: list[dict]) -> None:
        cur = {s_.get("id"): s_ for s_ in data.get("stages") or []}
        out = []
        for i, r in enumerate(recs):
            if not txt(r.get("name")):
                continue
            sid = _rid(r, "s", g, i)
            out.append({"id": sid, "name": txt(r.get("name")), "status": code_of(STATUS, txt(r.get("status"))) or "todo",
                        "start": iso(r.get("start")), "due": iso(r.get("due")), "owner": txt(r.get("owner")),
                        "note": txt(r.get("note")), "checklist": (cur.get(sid) or {}).get("checklist", [])})
        data["stages"] = out
    return conv


def checklist_rows(pick: str):
    def rows_of(data: dict) -> list[dict]:
        ck = next((s_.get("checklist", []) for s_ in data.get("stages") or [] if s_.get("id") == pick), [])
        return [{"done": bool(x.get("done")), "text": x.get("text", "")} for x in ck]
    return rows_of


def checklist_to_doc(pick: str):
    def conv(data: dict, recs: list[dict]) -> None:
        items = [{"text": txt(r.get("text")), "done": flag(r.get("done"))} for r in recs if txt(r.get("text"))]
        for s_ in data.get("stages") or []:
            if s_.get("id") == pick:
                s_["checklist"] = items
    return conv


def questions_to_doc(g: int):
    def conv(data: dict, recs: list[dict]) -> None:
        data["questions"] = [{"id": _rid(r, "q", g, i), "section": txt(r.get("section")), "text": txt(r.get("text"))}
                             for i, r in enumerate(recs) if txt(r.get("text"))]
    return conv


def kpi_rows(src_labels: dict):
    def rows_of(data: dict) -> list[dict]:
        return [{"id": k.get("id"), "name": k.get("name"), "target": k.get("target"), "unit": k.get("unit"),
                 "source": src_labels.get(k.get("source") or "", k.get("source")), "manual": k.get("manual")}
                for k in data.get("kpis") or []]
    return rows_of


def kpis_to_doc(g: int, inv: dict):
    def conv(data: dict, recs: list[dict]) -> None:
        data["kpis"] = [{"id": _rid(r, "k", g, i), "name": txt(r.get("name")), "target": num(r.get("target")),
                         "unit": txt(r.get("unit")), "source": inv.get(txt(r.get("source")), ""),
                         "manual": num(r.get("manual"))}
                        for i, r in enumerate(recs) if txt(r.get("name"))]
    return conv


def outcome_rows(olab: dict):
    def rows_of(data: dict) -> list[dict]:
        return [{"id": o.get("id"), "date": pdate(o.get("date")), "startup": o.get("startup"),
                 "partner": o.get("partner"), "type": olab.get(o.get("type"), o.get("type")),
                 "amount": o.get("amount"), "note": o.get("note")} for o in data.get("outcomes") or []]
    return rows_of


def outcomes_to_doc(g: int, oinv: dict):
    def conv(data: dict, recs: list[dict]) -> None:
        data["outcomes"] = [{"id": _rid(r, "o", g, i), "date": iso(r.get("date")), "startup": txt(r.get("startup")),
                             "partner": txt(r.get("partner")),
                             "type": oinv.get(txt(r.get("type")), txt(r.get("type")) or "other"),
                             "amount": num(r.get("amount")), "note": txt(r.get("note"))}
                            for i, r in enumerate(recs) if txt(r.get("startup")) or txt(r.get("partner"))]
    return conv


def quote_to_doc(g: int):
    def conv(data: dict, recs: list[dict]) -> None:
        q = data.setdefault("quote", {"currency": "KRW", "vat": True, "items": []})
        q["items"] = [{"id": _rid(r, "i", g, i), "cat": txt(r.get("cat")), "item": txt(r.get("item")),
                       "price": num(r.get("price")), "qty": num(r.get("qty")), "unit": txt(r.get("unit")),
                       "note": txt(r.get("note"))}
                      for i, r in enumerate(recs) if txt(r.get("item")) or txt(r.get("cat"))]
    return conv


# ---------------------------------------------------------------- hub
def page():
    ss()["_pg_cache"] = {}
    box = st.container()
    for fk in ("flash", "pg_flash"):
        flash = ss().pop(fk, None)
        if flash:
            getattr(box, flash[0])(flash[1])
    qp = st.query_params.get("program")
    if qp and ss().get("pg_open") != qp and not ss().get("pg_closed"):
        ss().pg_open = qp
    pid = ss().get("pg_open")
    if pid:
        program_view(pid)
    else:
        hub()
    if ss().pop("_pg_rerun", False):
        st.rerun()


def cb_open(pid: str):
    ss().pg_open = pid
    ss().pg_closed = False
    st.query_params["program"] = pid


def cb_close():
    harvest_all()
    ss().pg_open = None
    ss().pg_closed = True
    st.query_params.pop("program", None)


def cb_create(example: bool = False):
    s = ss()
    name = (s.get("np_name") or "").strip()
    if example:
        today = date.today()
        names = [API["dir_name"]("startup", r) for r in directory_startups()][:33]
        d = make_program("예시 · CES 2027 참가 지원" if lang() == "ko" else "Example · CES 2027 delegation", "expo",
                         today - timedelta(days=40), today + timedelta(days=60), "창업진흥원" if lang() == "ko" else "Client",
                         names)
        for st_ in d["stages"]:   # stages before today look done, the current one in progress
            if pdate(st_["due"]) < today:
                st_["status"] = "done"
                for c in st_["checklist"]:
                    c["done"] = True
            elif pdate(st_["start"]) <= today:
                st_["status"] = "doing"
    else:
        if not name:
            s.pg_flash = ("error", T("need_name"))
            return
        d = make_program(name, s.get("np_tpl", "accel"), s.get("np_start"), s.get("np_end"), s.get("np_client", ""))
    pid = new_id("p")
    Docs().put("program", pid, d, 0, "", d["name"])
    for k in ("np_name", "np_client"):
        s[k] = ""
    cb_open(pid)


def new_program_form():
    st.text_input(T("p_name"), key="np_name", placeholder="CES 2027 Korea Startup Delegation")
    st.text_input(T("p_client"), key="np_client", placeholder="창업진흥원")
    st.selectbox(T("p_template"), list(TEMPLATES), key="np_tpl", format_func=lambda k: L2(TEMPLATES[k][0]))
    a, b = st.columns(2)
    a.date_input(T("p_start"), value=date.today(), key="np_start")
    b.date_input(T("p_end"), value=date.today() + timedelta(days=90), key="np_end")
    st.button(T("create"), type="primary", on_click=cb_create)


def hub():
    docs = Docs()
    progs = docs.list("program", with_data=True)
    h1, h2 = st.columns([4, 1.3], vertical_alignment="bottom")
    h1.subheader(T("programs"), anchor=False)
    h1.caption(T("programs_cap"))
    with h2.popover(T("new_program"), width="stretch"):
        new_program_form()
    if not progs:
        st.info(T("empty_hub"))
        st.button(T("example"), on_click=cb_create, args=(True,))
        return

    # deadlines across every program
    today = date.today()
    due = []
    for p in progs:
        for s_ in p["data"].get("stages") or []:
            dd = pdate(s_.get("due"))
            if s_.get("status") != "done" and dd and dd <= today + timedelta(days=14):
                due.append((dd, p, s_))
    with st.container(border=True):
        st.markdown(f"**{T('upcoming')}**")
        if not due:
            st.caption(T("nothing_due"))
        for dd, p, s_ in sorted(due, key=lambda x: x[0])[:10]:
            late = dd < today
            a, b, c = st.columns([1.1, 5, 1.2], vertical_alignment="center")
            a.markdown(f":{'red' if late else 'orange' if (dd - today).days <= 3 else 'gray'}-badge[{dday(dd)}]")
            b.markdown(f"**{s_.get('name')}** · {p['data'].get('name')}"
                       + (f" · {s_['owner']}" if s_.get("owner") else ""))
            c.button(T("open"), key=f"due_{p['id']}_{s_.get('id')}", on_click=cb_open, args=(p["id"],), width="stretch")

    cols = st.columns(3)
    for i, p in enumerate(progs):
        d = p["data"]
        sm = stage_summary(d)
        with cols[i % 3].container(border=True):
            st.markdown(f"#### {d.get('name')}")
            st.caption(" · ".join(x for x in (d.get("client"), f"{docgen.kdate(d.get('start'))} ~ {docgen.kdate(d.get('end'))}",
                                              T("n_startups", n=len(d.get('startups') or []))) if x))
            st.progress(sm["done"] / sm["n"] if sm["n"] else 0.0,
                        text=T("progress", d=sm["done"], n=sm["n"]) if sm["n"] else "")
            if sm["now"]:
                st.write(T("stage_now", s=sm["now"]["name"]))
            if sm["next"]:
                st.write(T("next_due", s=sm["next"]["name"], d=dday(pdate(sm["next"]["due"]))))
            elif sm["n"]:
                st.write(T("all_done"))
            st.button(T("open"), key=f"open_{p['id']}", on_click=cb_open, args=(p["id"],), type="primary",
                      width="stretch")


# ---------------------------------------------------------------- one program
TAB_KEYS = ("tab_overview", "tab_stages", "tab_startups", "tab_intake", "tab_kpi", "tab_report")


def program_view(pid: str):
    w = wc("pg_w", "program", pid)
    if w.data is None:
        st.warning(T("missing"))
        st.button(T("back"), on_click=cb_close)
        return
    d, g = w.data, w.gen
    top1, top2 = st.columns([5, 1.4], vertical_alignment="center")
    top1.button(T("back"), on_click=cb_close)
    if w.meta.get("updated_by"):
        top2.caption(T("saved_by", when=time.strftime("%m-%d %H:%M", time.localtime(w.meta.get("updated_at") or time.time())),
                       who=w.meta.get("updated_by")))
    st.markdown(f"## {d.get('name')}")
    sm = stage_summary(d)
    st.caption(" · ".join(x for x in (d.get("client"), f"{docgen.kdate(d.get('start'))} ~ {docgen.kdate(d.get('end'))}",
                                      T("n_startups", n=len(d.get('startups') or [])),
                                      T("progress", d=sm["done"], n=sm["n"]) if sm["n"] else "") if x))
    # The open tab survives reruns (it is tracked by key) and coming back from another page (remembered here).
    labels = [T(k) for k in TAB_KEYS]
    tkey, mem, first = f"pg_tab_{pid}", ss().setdefault("pg_tab_mem", {}), ss().setdefault("pg_tab_first", {})
    if tkey not in ss() or first.get(pid) not in labels:
        first[pid] = mem.get(pid) if mem.get(pid) in labels else labels[0]
    tabs = st.tabs(labels, key=tkey, on_change="rerun", default=first[pid])
    if ss().get(tkey) in labels:
        mem[pid] = ss()[tkey]
    with tabs[0]:
        tab_overview(pid, d)
    with tabs[1]:
        tab_stages(d, g)
    with tabs[2]:
        tab_startups(pid, d)
    with tabs[3]:
        tab_intake(pid, d, g)
    with tabs[4]:
        tab_kpi(pid, d, g)
    with tabs[5]:
        tab_report(pid, d, g)
    if ss().get("pg_w") is w and ss().get("pg_open"):
        w.save(d.get("name", ""))


def cb_delete(pid: str):
    if ss().get("pg_del_ok"):
        Docs().archive("program", pid)
        ss().pg_del_ok = False
        ss().pg_flash = ("success", T("deleted"))
        cb_close()


def tab_overview(pid: str, d: dict):
    c1, c2 = st.columns(2)
    text_box("pg_w", c1, T("p_name"), "name")
    text_box("pg_w", c2, T("p_client"), "client")
    c3, c4, c5, c6 = st.columns(4)
    date_box("pg_w", c3, T("p_start"), "start", d)
    date_box("pg_w", c4, T("p_end"), "end", d)
    text_box("pg_w", c5, T("p_pm"), "pm")
    text_box("pg_w", c6, T("p_team"), "team")
    text_box("pg_w", st, T("p_summary"), "summary", area=True, height=110)
    text_box("pg_w", st, T("p_budget"), "budget")
    choices = API["event_choices"]() if API.get("event_choices") else []
    if choices:
        names = dict(choices)
        k = wkey("pg_w", "events")
        drop_if_invalid(k, lambda v: all(e in names for e in v or []))
        d["events"] = st.multiselect(T("events_link"), list(names), default=[e for e in d.get("events") or [] if e in names],
                                     format_func=lambda e: names.get(e, e), key=k, help=T("events_help"))
    else:
        st.caption(T("no_events"))
    with st.popover(T("manage")):
        st.checkbox(T("delete_confirm"), key="pg_del_ok")
        st.button(T("delete_prog"), on_click=cb_delete, args=(pid,), disabled=not ss().get("pg_del_ok"))


def tab_stages(d: dict, g: int):
    st.caption(T("stages_help"))
    labels = [L2(v) for v in STATUS.values()]
    table_box("pg_w", st, "stages", stage_rows, STAGE_COLS, stages_to_doc(g),
              num_rows="dynamic", hide_index=True, width="stretch",
              column_config={"id": None,
                             "status": st.column_config.SelectboxColumn(T("col_status"), options=labels, width="small"),
                             "name": st.column_config.TextColumn(T("col_stage"), width="medium"),
                             "start": st.column_config.DateColumn(T("col_start"), format="YYYY-MM-DD"),
                             "due": st.column_config.DateColumn(T("col_due"), format="YYYY-MM-DD"),
                             "owner": st.column_config.TextColumn(T("col_owner"), width="small"),
                             "note": st.column_config.TextColumn(T("col_note"), width="large")})
    out = d.get("stages") or []

    # timeline
    tl = [s_ for s_ in out if pdate(s_["due"])]
    if tl:
        import altair as alt
        data = pd.DataFrame([{"stage": s_["name"], "from": pdate(s_["start"]) or pdate(s_["due"]) - timedelta(days=1),
                              "to": pdate(s_["due"]) + timedelta(days=1), "status": label_of(STATUS, s_["status"]),
                              "owner": s_["owner"]} for s_ in tl])
        bars = alt.Chart(data).mark_bar(cornerRadius=4).encode(
            x=alt.X("from:T", title=None, axis=alt.Axis(format="%m/%d", grid=True)), x2="to:T",
            y=alt.Y("stage:N", sort=None, title=None, axis=alt.Axis(labelOverlap=False, labelLimit=220, labelFontSize=12)),
            color=alt.Color("status:N", title=None, legend=alt.Legend(orient="top"),
                            scale=alt.Scale(domain=[L2(v) for v in STATUS.values()], range=list(STATUS_COLORS.values()))),
            tooltip=["stage", "status", "owner", alt.Tooltip("from:T", format="%Y-%m-%d"), alt.Tooltip("to:T", format="%Y-%m-%d")])
        rule = alt.Chart(pd.DataFrame({"t": [pd.Timestamp(date.today())]})).mark_rule(color="#d92d20", strokeDash=[4, 3]).encode(x="t:T")
        st.altair_chart((bars + rule).properties(height=max(160, 30 * len(tl) + 30)), width="stretch")
    else:
        st.caption(T("no_dates"))

    # checklist of one stage
    if out:
        st.markdown(f"**{T('checklist')}**")
        names = {s_["id"]: s_["name"] for s_ in out}
        k = wkey("pg_w", "ckpick")
        drop_if_invalid(k, lambda v: v in names)
        first_open = next((s_["id"] for s_ in out if s_["status"] == "doing"), out[0]["id"])
        pick = st.selectbox(T("ck_for"), list(names), index=list(names).index(first_open), format_func=lambda x: names[x],
                            key=k, label_visibility="collapsed")
        table_box("pg_w", st, f"ck_{pick}", checklist_rows(pick), ["done", "text"], checklist_to_doc(pick),
                  num_rows="dynamic", hide_index=True, width="stretch",
                  column_config={"done": st.column_config.CheckboxColumn(T("ck_done"), width="small"),
                                 "text": st.column_config.TextColumn(T("ck_item"), width="large")})
        items = next((s_.get("checklist", []) for s_ in d.get("stages") or [] if s_["id"] == pick), [])
        st.caption(T("ck_count", d=sum(1 for x in items if x["done"]), n=len(items)))


def tab_startups(pid: str, d: dict):
    rows = directory_startups()
    names = [API["dir_name"]("startup", r) for r in rows]
    if not names and not d.get("startups"):
        st.info(T("dir_empty"))
        if API.get("pages", {}).get("companies"):
            st.page_link(API["pages"]["companies"], label=T("go_companies"))
        return
    opts = names + [n for n in (d.get("startups") or []) if n not in names]
    k = wkey("pg_w", "startups")
    drop_if_invalid(k, lambda v: all(n in opts for n in v or []))
    d["startups"] = st.multiselect(T("pick_startups"), opts, default=[n for n in d.get("startups") or [] if n in opts],
                                   key=k, help=T("pick_help"))
    if not d["startups"]:
        return
    match = matching_for(d)
    intakes = intake_map(pid)
    outs = d.get("outcomes") or []
    table = []
    for s_ in startup_rows(d):
        m = match["startups"].get(s_["key"], {})
        table.append({T("c_company"): s_["name"], T("c_kr"): s_["kr"], T("c_tags"): s_["tags"],
                      T("c_intake"): "✅" if intakes.get(s_["key"], {}).get("status") == "done" else "·",
                      T("c_meet"): m.get("total", 0), T("c_done"): m.get("done", 0),
                      T("c_out"): sum(1 for o in outs if norm(o.get("startup")) == s_["key"])})
    st.dataframe(pd.DataFrame(table), hide_index=True, width="stretch")
    if API.get("pages", {}).get("companies"):
        st.page_link(API["pages"]["companies"], label=T("go_companies"), icon="📇")


# ---------------------------------------------------------------- intake interviews
def tab_intake(pid: str, d: dict, g: int):
    srows = startup_rows(d)
    if not srows:
        st.info(T("ik_none"))
        return
    with st.popover(T("ik_questions")):
        st.caption(T("ik_q_help"))
        table_box("pg_w", st, "questions", lambda data: data.get("questions") or [], ["id", "section", "text"],
                  questions_to_doc(g), num_rows="dynamic", hide_index=True, width="stretch",
                  column_config={"id": None, "section": st.column_config.TextColumn(T("q_section"), width="small"),
                                 "text": st.column_config.TextColumn(T("q_text"), width="large")})
    intakes = intake_map(pid)
    done = sum(1 for s_ in srows if intakes.get(s_["key"], {}).get("status") == "done")
    st.progress(done / len(srows), text=T("ik_progress", d=done, n=len(srows)))
    names = [s_["name"] for s_ in srows]
    pk, mem = f"ik_pick_{pid}", ss().setdefault("ik_pick_mem", {})
    drop_if_invalid(pk, lambda v: v in names)
    pick = st.selectbox(T("ik_pick"), names, index=names.index(mem[pid]) if mem.get(pid) in names else 0, key=pk,
                        format_func=lambda n: ("✅ " if intakes.get(norm(n), {}).get("status") == "done" else "⏳ ") + n)
    mem[pid] = pick
    srow = next(s_ for s_ in srows if s_["name"] == pick)
    default = {"startup": pick, "date": date.today().isoformat(), "interviewer": API["me"]() if API.get("me") else "",
               "attendees": "", "status": "todo", "answers": {q["id"]: "" for q in d.get("questions") or []},
               "notes": "", "summary": "", "suggest": {}}
    w = wc("ik_w", "intake", f"{pid}:{pick}", pid, default)
    r, rg = w.data, w.gen
    a1, a2, a3, a4 = st.columns([1, 1.2, 2, 1])
    date_box("ik_w", a1, T("ik_date"), "date", r)
    text_box("ik_w", a2, T("ik_by"), "by")
    text_box("ik_w", a3, T("ik_with"), "with")
    stl = {"todo": T("ik_todo"), "done": T("ik_done")}
    k = wkey("ik_w", "status")
    drop_if_invalid(k, lambda v: v in stl)
    r["status"] = a4.selectbox(T("ik_status"), list(stl), index=1 if r.get("status") == "done" else 0,
                               format_func=stl.get, key=k)
    section = None
    for q in d.get("questions") or []:
        if q.get("section") and q["section"] != section:
            section = q["section"]
            st.markdown(f"**{section}**")
        text_box("ik_w", st, q["text"], q["id"], area=True, height=80)
    text_box("ik_w", st, T("ik_notes"), "notes", area=True, height=80)

    c1, c2, c3 = st.columns([1.2, 1.2, 1.2])
    if ai.available():
        if c1.button(T("ik_ai"), key=f"ik_{rg}_ai", type="primary"):
            qa = [(q.get("section", ""), q["text"], (r.get("answers") or {}).get(q["id"], "")) for q in d.get("questions") or []]
            profile = {k_: srow.get(k_) for k_ in ("name", "kr", "tags", "stage", "one_liner")}
            try:
                with st.spinner(T("rp_ai_wait")):
                    r["summary"], r["suggest"] = ai.summarize_intake(pick, profile, qa, r.get("notes", ""), lang())
                w.rebase()
                w.save(pick)
                st.rerun()
            except Exception as e:  # noqa: BLE001
                st.error(T("ai_fail", e=e))
    else:
        c1.caption(T("ik_ai_off"))
    md_one = docgen.intake_markdown(srow, r, d.get("questions") or [], lang())
    c2.download_button(T("ik_doc"), deferred_hwpx(md_one), fname(f"Intake_{pick}"), key=f"ik_{rg}_dl",
                       mime=HWPX_MIME, width="stretch", on_click="ignore")
    allr = intake_map(pid)
    md_all = f"# {docgen.esc(d.get('name'))} Intake Interview\n\n" + "\n".join(
        docgen.intake_markdown(s_, allr.get(s_["key"], {}), d.get("questions") or [], lang(), heading=2) for s_ in srows)
    c3.download_button(T("ik_doc_all"), deferred_hwpx(md_all), fname(f"Intake_{d.get('name')}"), key=f"ik_{rg}_all",
                       mime=HWPX_MIME, width="stretch", on_click="ignore")

    if r.get("summary"):
        st.markdown(f"**{T('ik_summary')}**")
        text_box("ik_w", st, T("ik_summary"), "summary", area=True, height=200, label_visibility="collapsed",
                 help=T("ik_summary_help"))
    sug = r.get("suggest") or {}
    if sug:
        with st.expander(T("ik_suggest")):
            fields = {"tags": "Tags", "stage": "Stage", "one_liner": "One-liner"}
            for k_, col in fields.items():
                if sug.get(k_):
                    st.checkbox(f"{col} → {sug[k_]}  ({T('ik_now', v=srow.get(k_) or '-')})", value=not srow.get(k_),
                                key=f"ik_{rg}_ap_{k_}")
            if st.button(T("ik_apply"), key=f"ik_{rg}_apply"):
                upd = {col: sug[k_] for k_, col in fields.items() if sug.get(k_) and ss().get(f"ik_{rg}_ap_{k_}")}
                if upd and API.get("dir_update"):
                    API["dir_update"]("startup", pick, upd)
                    st.success(T("ik_applied"))
    w.save(pick)


HWPX_MIME = "application/hwp+zip"


def deferred_hwpx(md: str, byline: str | None = None):
    """The .hwpx file, made only when someone clicks download (it takes a moment to build)."""
    def make() -> bytes:
        return docgen.to_hwpx(md, byline)[0]
    return make


def fname(name: str, ext: str = "hwpx") -> str:
    bad = '\\/:*?"<>|\n\r\t'
    clean = "".join("_" if ch in bad else ch for ch in str(name or "document")).strip(" ._") or "document"
    return f"{clean[:80]}.{ext}"


# ---------------------------------------------------------------- results and KPIs
def tab_kpi(pid: str, d: dict, g: int):
    match = matching_for(d)
    intakes = intake_map(pid)
    st.markdown(f"**{T('m_title')}**")
    m1, m2, m3, m4, m5 = st.columns(5)
    by = match["by"]
    m1.metric(T("m_total"), match["total"])
    m2.metric(T("m_done"), by.get("Done", 0))
    m3.metric(T("m_noshow"), by.get("No-Show", 0))
    m4.metric(T("m_up"), by.get("Scheduled", 0) + by.get("Ongoing", 0))
    m5.metric(T("m_partners"), len(match["partners"]))
    if not match["events"]:
        st.caption(T("no_events") if not (API.get("event_choices") and API["event_choices"]()) else T("events_help"))

    st.markdown(f"**{T('kpi_title')}**")
    st.caption(T("kpi_help"))
    src_labels = {k: L2(v) for k, v in SOURCES.items()}
    inv = {v: k for k, v in src_labels.items()}
    table_box("pg_w", st, "kpis", kpi_rows(src_labels), ["id", "name", "target", "unit", "source", "manual"],
              kpis_to_doc(g, inv), num_rows="dynamic", hide_index=True, width="stretch",
              column_config={"id": None, "name": st.column_config.TextColumn(T("k_name"), width="medium"),
                             "target": st.column_config.NumberColumn(T("k_target")),
                             "unit": st.column_config.TextColumn(T("k_unit"), width="small"),
                             "source": st.column_config.SelectboxColumn(T("k_source"), options=list(src_labels.values()), width="medium"),
                             "manual": st.column_config.NumberColumn(T("k_manual"))})
    vals = kpi_values(d, match, intakes)
    if vals:
        st.dataframe(pd.DataFrame([{T("k_name"): v["name"], T("k_target"): v["target_s"], T("k_actual"): v["actual_s"],
                                    T("k_rate"): min(1.0, v["rate"]) if v["rate"] is not None else None} for v in vals]),
                     hide_index=True, width="stretch",
                     column_config={T("k_rate"): st.column_config.ProgressColumn(T("k_rate"), min_value=0.0, max_value=1.0, format="percent")})

    st.markdown(f"**{T('out_title')}**")
    st.caption(T("out_help"))
    olab = {k: L2(v) for k, v in OUTCOMES.items()}
    oinv = {v: k for k, v in olab.items()}
    table_box("pg_w", st, "outcomes", outcome_rows(olab), ["id", "date", "startup", "partner", "type", "amount", "note"],
              outcomes_to_doc(g, oinv), num_rows="dynamic", hide_index=True, width="stretch",
              column_config={"id": None, "date": st.column_config.DateColumn(T("o_date"), format="YYYY-MM-DD"),
                             "startup": st.column_config.SelectboxColumn(T("o_startup"), options=d.get("startups") or []),
                             "partner": st.column_config.TextColumn(T("o_partner")),
                             "type": st.column_config.SelectboxColumn(T("o_type"), options=list(olab.values())),
                             "amount": st.column_config.NumberColumn(T("o_amount")),
                             "note": st.column_config.TextColumn(T("o_note"), width="large")})


# ---------------------------------------------------------------- reports
def report_context(pid: str, d: dict, kind: str, asof: date) -> dict:
    match = matching_for(d)
    intakes = intake_map(pid)
    outs = [{**o, "startup_key": norm(o.get("startup")), "type_label": label_of(OUTCOMES, o.get("type", "other"))}
            for o in d.get("outcomes") or []]
    stages = [{**s_, "status_label": label_of(STATUS, s_.get("status", "todo"))} for s_ in d.get("stages") or []]
    return {"lang": lang(), "asof": asof.isoformat(), "kind": kind, "program": d, "startups": startup_rows(d),
            "stages": stages, "kpis": kpi_values(d, match, intakes), "matching": match, "outcomes": outs,
            "intake": intakes}


def ai_data(key: str, c: dict) -> dict:
    """Only what one section needs goes to the AI."""
    p = c["program"]
    base = {"program": {k: p.get(k) for k in ("name", "client", "start", "end", "pm", "team", "summary")},
            "as_of": c["asof"]}
    m = c["matching"]
    msum = {"events": m["events"], "total": m["total"], "by_status": m["by"], "partners": sorted(m["partners"])[:40]}
    kpis = [{k: v.get(k) for k in ("name", "target_s", "actual_s", "rate_s")} for v in c["kpis"]]
    if key == "overview":
        base["startups"] = [{"name": s["name"], "tags": s["tags"]} for s in c["startups"]]
    elif key in ("progress", "schedule"):
        base["stages"] = [{k: s.get(k) for k in ("name", "start", "due", "status_label", "owner", "note")} for s in c["stages"]]
    elif key == "plan":
        base["stages"] = [{"name": s.get("name"), "start": s.get("start"), "due": s.get("due"), "note": s.get("note"),
                           "checklist": [x.get("text") for x in s.get("checklist") or []]} for s in c["stages"]]
    elif key == "team":
        base["stage_owners"] = [{"stage": s.get("name"), "owner": s.get("owner")} for s in c["stages"] if s.get("owner")]
    elif key == "budget":
        q = p.get("quote") or {}
        sub, vat, tot = docgen.quote_totals(q)
        base["quote"] = {"currency": q.get("currency") or "KRW", "items": q.get("items") or [], "subtotal": sub,
                         "vat": vat, "total": tot, "budget_note": p.get("budget")}
    elif key in ("background",):
        base["startups_count"] = len(c["startups"])
    elif key == "outcomes":
        base["indicator_targets"] = [{k: v.get(k) for k in ("name", "target_s")} for v in c["kpis"]]
    elif key in ("kpi", "lessons"):
        base["indicators"] = kpis
        base["matching"] = msum
        if key == "lessons":
            base["outcomes"] = [{k: o.get(k) for k in ("startup", "partner", "type_label", "date", "amount", "note")} for o in c["outcomes"]]
    elif key == "matching":
        base["matching"] = msum
        base["by_startup"] = [{"name": v["name"], "total": v["total"], "done": v["done"], "no_show": v["noshow"],
                               "partners": v["partners"][:10]} for v in m["startups"].values()]
        base["outcomes"] = [{k: o.get(k) for k in ("startup", "partner", "type_label")} for o in c["outcomes"]]
    elif key == "startups":
        base["startups"] = [{"name": s["name"], "tags": s["tags"],
                             "intake_summary": (c["intake"].get(s["key"], {}).get("summary") or "")[:700],
                             "meetings": {k: m["startups"].get(s["key"], {}).get(k, 0) for k in ("total", "done", "noshow")},
                             "outcomes": [f"{o['type_label']} · {o.get('partner')}" for o in c["outcomes"] if o["startup_key"] == s["key"]]}
                            for s in c["startups"]]
    elif key == "next":
        base["remaining_stages"] = [{k: s.get(k) for k in ("name", "due", "status_label", "owner", "note")} for s in c["stages"] if s.get("status") != "done"]
        base["indicators"] = kpis
    return base


def default_kind(d: dict) -> str:
    """Which document a program most likely needs now."""
    stages = d.get("stages") or []
    if not stages or all(s_.get("status") == "todo" for s_ in stages):
        return "proposal"
    for s_ in stages:
        if s_.get("status") in ("doing", "done") and any(x in str(s_.get("name")) for x in ("최종", "Final", "final")):
            return "final"
    return "interim"


def quote_editor(d: dict, g: int):
    q = d.setdefault("quote", {"currency": "KRW", "vat": True, "items": []})
    with st.container(border=True):
        st.markdown(f"**{T('qt_title')}**")
        st.caption(T("qt_help"))
        c1, c2, _ = st.columns([1, 1.2, 3], vertical_alignment="bottom")
        curs = list(docgen.CURRENCY)
        k = wkey("pg_w", "qcur")
        drop_if_invalid(k, lambda v: v in curs)
        q["currency"] = c1.selectbox(T("qt_currency"), curs, index=curs.index(q.get("currency")) if q.get("currency") in curs else 0,
                                     key=k)
        q["vat"] = c2.toggle(T("qt_vat"), value=bool(q.get("vat", True)), key=wkey("pg_w", "qvat"))
        table_box("pg_w", st, "quote", lambda data: (data.get("quote") or {}).get("items") or [],
                  ["id", "cat", "item", "price", "qty", "unit", "note"], quote_to_doc(g),
                  num_rows="dynamic", hide_index=True, width="stretch",
                  column_config={"id": None, "cat": st.column_config.TextColumn(T("qt_cat"), width="small"),
                                 "item": st.column_config.TextColumn(T("qt_item"), width="medium"),
                                 "price": st.column_config.NumberColumn(T("qt_price"), format="localized"),
                                 "qty": st.column_config.NumberColumn(T("qt_qty")),
                                 "unit": st.column_config.TextColumn(T("qt_unit"), width="small"),
                                 "note": st.column_config.TextColumn(T("qt_note"), width="medium")})
        sub, vat, tot = docgen.quote_totals(q)
        unit = docgen.CURRENCY.get(q["currency"], ("", ""))[0 if lang() == "ko" else 1]
        m1, m2, m3 = st.columns(3)
        m1.metric(T("qt_sub"), f"{docgen.money(sub, q['currency'])} {unit}")
        m2.metric(T("qt_vatv"), f"{docgen.money(vat, q['currency'])} {unit}")
        m3.metric(T("qt_total"), f"{docgen.money(tot, q['currency'])} {unit}")


def refresh_auto(r: dict, titles: list, c: dict) -> bool:
    """Rebuild every section that still holds exactly what was last built from the data (nobody edited it)."""
    secs, auto = r.setdefault("sections", {}), r.setdefault("auto", {})
    changed = False
    for k, _ in titles:
        new = docgen.build_section(k, c)
        if k not in secs or (k in auto and secs[k] == auto[k] and new != auto[k]):
            secs[k] = auto[k] = new
            changed = True
    return changed


def tab_report(pid: str, d: dict, g: int):
    kinds = {"proposal": T("rp_proposal"), "interim": T("rp_interim"), "final": T("rp_final")}
    kk, mem = f"rp_kind_{pid}", ss().setdefault("rp_kind_mem", {})
    drop_if_invalid(kk, lambda v: v in kinds)
    kind = st.segmented_control(T("rp_kind"), list(kinds), format_func=kinds.get, key=kk, required=True,
                                default=mem[pid] if mem.get(pid) in kinds else default_kind(d),
                                label_visibility="collapsed") or "interim"
    mem[pid] = kind
    titles = docgen.section_titles(kind, lang())
    default = {"kind": kind, "title": f"{d.get('name')} {T('rp_doc_' + kind)}", "asof": date.today().isoformat(),
               "byline": f"{docgen.kdate(date.today())}  {d.get('pm') or 'Plug and Play'}", "sections": {}, "auto": {}}
    w = wc("rp_w", "report", f"{pid}:{kind}", pid, default)
    r = w.data
    if kind == "proposal":
        quote_editor(d, g)
    # Sections nobody has edited follow the program data. That is not a change worth saving: anyone who opens
    # the report gets them rebuilt the same way, so the document is only written once someone edits it.
    in_sync = h(r) == w.saved
    if refresh_auto(r, titles, report_context(pid, d, kind, pdate(r.get("asof")) or date.today())):
        w.rebase()
        if in_sync:
            w.saved = h(r)
    rg = w.gen
    a1, a2, a3 = st.columns([3, 1, 2])
    text_box("rp_w", a1, T("rp_title"), "title")
    asof = pdate(date_box("rp_w", a2, T("rp_asof"), "asof", r)) or date.today()
    text_box("rp_w", a3, T("rp_byline"), "byline")
    c = report_context(pid, d, kind, asof)

    def draft(k: str, name: str) -> str:
        return ai.draft_section(kinds[kind], name, docgen.AI_HINTS[k][0 if lang() == "ko" else 1], ai_data(k, c),
                                r["sections"].get(k, ""), lang())

    b1, b2, _ = st.columns([1.4, 1.4, 2])
    if b1.button(T("rp_fill_all"), key=f"rp_{rg}_fillall"):
        r["sections"] = {k: docgen.build_section(k, c) for k, _ in titles}
        r["auto"] = dict(r["sections"])
        w.rebase()
        st.rerun()
    if ai.available():
        if b2.button(T("rp_ai_all"), key=f"rp_{rg}_aiall", type="primary"):
            try:
                with st.spinner(T("rp_ai_wait")):
                    for k, name in titles:
                        r["sections"][k] = draft(k, name)
                        r.setdefault("auto", {}).pop(k, None)
                w.rebase()
                ss()["pg_flash"] = ("info", T("rp_ai_done"))
                st.rerun()
            except Exception as e:  # noqa: BLE001
                st.error(T("ai_fail", e=e))
    else:
        b2.caption(T("ik_ai_off"))
    st.caption(T("rp_help"))
    for k, name in titles:
        with st.expander(name, expanded=False):
            x1, x2, _ = st.columns([1, 1, 3])
            if x1.button(T("rp_fill"), key=f"rp_{rg}_fill_{k}"):
                r["sections"][k] = docgen.build_section(k, c)
                r.setdefault("auto", {})[k] = r["sections"][k]
                w.rebase()
                st.rerun()
            if ai.available() and x2.button(T("rp_ai"), key=f"rp_{rg}_ai_{k}"):
                try:
                    with st.spinner(T("rp_ai_wait")):
                        r["sections"][k] = draft(k, name)
                        r.setdefault("auto", {}).pop(k, None)
                    w.rebase()
                    st.rerun()
                except Exception as e:  # noqa: BLE001
                    st.error(T("ai_fail", e=e))
            text_box("rp_w", st, name, k, area=True, height=220, label_visibility="collapsed")
    md = docgen.report_markdown(r.get("title") or "", kind, r["sections"], lang())
    m1, m2, m3 = st.columns([1.2, 1.2, 2.4], vertical_alignment="center")
    m1.download_button(T("rp_download"), deferred_hwpx(md, r.get("byline")), fname(r.get("title") or kinds[kind]),
                       mime=HWPX_MIME, key=f"rp_{rg}_dl", type="primary", width="stretch", on_click="ignore")
    if m2.button(T("rp_preview_btn"), key=f"rp_{rg}_pv", width="stretch"):
        try:
            data, issues = docgen.to_hwpx(md, r.get("byline"))
            ss()["rp_pv"] = (w.id, h(r), docgen.preview_html(data), issues)
        except Exception as e:  # noqa: BLE001
            st.error(T("rp_issues", e=e))
    m3.caption(T("rp_hwp_note"))
    pv = ss().get("rp_pv")
    if pv and pv[0] == w.id:
        if pv[3]:
            st.warning(T("rp_issues", e="; ".join(pv[3])))
        if pv[1] != h(r):
            st.caption(T("rp_pv_old"))
        with st.expander(T("rp_preview"), expanded=True):
            components.html(pv[2], height=900, scrolling=True)
    w.save(r.get("title", ""))
