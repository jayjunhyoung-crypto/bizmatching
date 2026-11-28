"""Report and intake documents: Markdown built from program data, exported as 한글 (.hwpx).

The .hwpx file is made with python-hwpx-automation's government report look (title band, numbered section
boxes, □ / - / · bullets, ruled tables, page numbers), then checked with python-hwpx's open-safety check.
"""
from __future__ import annotations

import io
from collections import Counter
from datetime import date

L = {  # (ko, en)
    "name": ("사업명", "Program"), "client": ("발주처", "Client"), "period": ("사업기간", "Period"),
    "startups": ("참가기업", "Startups"), "lead": ("수행", "Run by"), "goal": ("사업 목적", "Objective"),
    "fields": ("분야", "Sectors"), "stage": ("단계", "Stage"), "when": ("기간", "Dates"), "status": ("상태", "Status"),
    "note": ("비고", "Notes"), "kpi": ("지표", "Indicator"), "target": ("목표", "Target"), "actual": ("실적", "Actual"),
    "rate": ("달성률", "Achieved"), "company": ("기업명", "Company"), "meet": ("미팅", "Meetings"),
    "done": ("완료", "Done"), "noshow": ("노쇼", "No-show"), "up": ("예정", "Upcoming"), "intake": ("Intake", "Intake"),
    "outcome": ("주요 성과", "Key outcomes"), "partner": ("파트너", "Partner"), "type": ("유형", "Type"),
    "date": ("일자", "Date"), "count_co": ("{n}개사", "{n} companies"), "count_case": ("{n}건", "{n}"),
    "of_stages": ("총 {n}단계 중 {d}단계 완료 (기준일 {a})", "{d} of {n} stages done (as of {a})"),
    "total_meet": ("총 미팅 {n}건 (완료 {d}, 노쇼 {x}, 진행·예정 {u})",
                   "{n} meetings in total ({d} done, {x} no-shows, {u} upcoming or ongoing)"),
    "partners_n": ("참여 파트너 {n}곳", "{n} partners took part"),
    "no_match": ("연결된 비즈매칭 일정이 없음", "No business matching event is linked"),
    "no_kpi": ("설정된 성과 지표 없음", "No indicators set"),
    "pending": ("{s} ({d}) 예정", "{s} ({d})"),
    "no_pending": ("남은 단계 없음", "No remaining stages"),
    "write_here": ("※ 성과 분석과 제언을 작성해 주세요 (AI 초안 사용 가능)",
                   "※ Write the analysis and recommendations here (AI draft available)"),
    "done_intake": ("완료", "Done"), "todo_intake": ("미완료", "Not yet"),
    "info": ("기업 기본 정보", "Company profile"), "overview_i": ("인터뷰 개요", "Interview"),
    "qa": ("질의응답", "Questions and answers"), "summary": ("요약", "Summary"), "memo": ("메모", "Notes"),
    "i_date": ("일시", "Date"), "i_by": ("면담자", "Interviewer"), "i_with": ("참석자", "Attendees"),
    "ceo": ("대표자", "CEO"), "kr": ("국문명", "Korean name"), "web": ("웹사이트", "Website"),
    "no_answer": ("(답변 없음)", "(no answer)"),
    "bg_write": ("※ 추진 배경과 필요성, 사업 목적을 작성해 주세요 (AI 초안 사용 가능)",
                 "※ Write the background and objectives here (AI draft available)"),
    "expect_write": ("※ 기대 효과를 작성해 주세요 (AI 초안 사용 가능)", "※ Write the expected impact here (AI draft available)"),
    "q_cat": ("구분", "Category"), "q_item": ("내역", "Item"), "q_price": ("단가", "Unit price"), "q_qty": ("수량", "Qty"),
    "q_amount": ("금액", "Amount"), "subtotal": ("소계", "Subtotal"), "vat": ("부가세 (10%)", "VAT (10%)"),
    "total": ("합계", "Total"), "budget_text": ("총 사업비", "Total budget"), "no_quote": ("견적 항목 없음", "No quote items yet"),
    "pm": ("PM", "PM"), "team": ("팀원", "Team"), "owner": ("담당", "Owner"), "by_stage": ("단계별 담당", "Owners by stage"),
}

CURRENCY = {"KRW": ("원", "KRW"), "USD": ("USD", "USD"), "SGD": ("SGD", "SGD")}


def money(v, cur: str = "KRW") -> str:
    try:
        f = float(v or 0)
    except (TypeError, ValueError):
        return str(v)
    return f"{f:,.0f}" if cur == "KRW" or f == int(f) else f"{f:,.2f}"


def quote_totals(q: dict) -> tuple[float, float, float]:
    sub = 0.0
    for it in (q or {}).get("items") or []:
        try:
            sub += float(it.get("price") or 0) * float(it.get("qty") or 0)
        except (TypeError, ValueError):
            pass
    vat = round(sub * 0.1, 2) if (q or {}).get("vat", True) else 0.0
    return sub, vat, sub + vat


def tr(key: str, lang: str, **kw) -> str:
    ko, en = L[key]
    s = ko if lang == "ko" else en
    return s.format(**kw) if kw else s


def esc(v) -> str:
    """Text that goes inside Markdown: no line breaks, no table pipes, no accidental markup."""
    s = "" if v is None else str(v)
    s = s.replace("\r", " ").replace("\n", " ").replace("|", "/").strip()
    for ch in ("*", "_", "`"):
        s = s.replace(ch, "\\" + ch)
    return s


def md_table(headers: list[str], rows: list[list], blank: str = "-") -> str:
    if not rows:
        return ""
    out = ["| " + " | ".join(esc(h) for h in headers) + " |", "|" + "---|" * len(headers)]
    out += ["| " + " | ".join(esc(c) if esc(c) else blank for c in r) + " |" for r in rows]
    return "\n".join(out)


def kdate(d) -> str:
    """2026-11-28 -> 2026. 11. 28. (Korean report style)"""
    if not d:
        return ""
    if isinstance(d, str):
        try:
            d = date.fromisoformat(d[:10])
        except ValueError:
            return d
    return f"{d.year}. {d.month}. {d.day}."


def short(d) -> str:
    if not d:
        return ""
    if isinstance(d, str):
        try:
            d = date.fromisoformat(d[:10])
        except ValueError:
            return d
    return f"{d.month}. {d.day}."


# ---------------------------------------------------------------- report sections (data drafts)
SECTIONS = {
    "proposal": [("background", "추진 배경 및 목적", "Background and objectives"), ("overview", "사업 개요", "Program overview"),
                 ("plan", "세부 추진 내용", "What we will do"), ("schedule", "추진 일정", "Schedule"),
                 ("team", "수행 체계", "Team"), ("budget", "소요 예산", "Budget"), ("outcomes", "기대 성과", "Expected results")],
    "interim": [("overview", "사업 개요", "Program overview"), ("progress", "추진 경과", "Progress"),
                ("kpi", "주요 성과 지표", "Key indicators"), ("matching", "비즈매칭 결과", "Business matching"),
                ("startups", "참가기업별 현황", "Startups"), ("next", "향후 계획", "Next steps")],
    "final": [("overview", "사업 개요", "Program overview"), ("progress", "추진 경과", "Progress"),
              ("kpi", "성과 지표 달성 현황", "Indicators achieved"), ("matching", "비즈매칭 결과", "Business matching"),
              ("startups", "참가기업별 성과", "Results by startup"), ("lessons", "성과 분석 및 제언", "Analysis and recommendations")],
}

AI_HINTS = {  # what the AI should write in each section
    "overview": ("사업명, 발주처, 기간, 참가기업 수와 분야 분포, 사업 목적을 3~6개 글머리표로 정리",
                 "Program, client, dates, number of startups and sectors, objective, in 3 to 6 bullets"),
    "progress": ("완료된 단계와 진행 중인 단계를 날짜와 함께 정리하고, 일정 지연이 있으면 사유를 (확인 필요)로 표시",
                 "Completed and ongoing stages with dates; flag delays with (to be confirmed)"),
    "kpi": ("지표별 목표 대비 실적과 달성률을 표로 쓰고, 눈에 띄는 점 2~3개를 글머리표로",
            "Targets against actuals as a table, then 2 or 3 bullets on what stands out"),
    "matching": ("미팅 총계, 완료와 노쇼, 많이 만난 기업과 파트너를 정리하고 후속 성과와 연결",
                 "Totals, done and no-shows, most active startups and partners, and follow-up outcomes"),
    "startups": ("기업별로 1~2줄씩 현황과 성과를 정리. intake 요약과 미팅, 성과 기록을 근거로",
                 "One or two lines per startup from the intake summary, meetings and outcomes"),
    "next": ("남은 단계와 날짜, 보완할 점을 정리", "Remaining stages with dates and what to improve"),
    "lessons": ("성과 분석, 잘된 점, 아쉬운 점, 다음 사업을 위한 제언을 정리",
                "Analysis, what worked, what did not, and recommendations for the next program"),
    "background": ("사업 목적 메모를 바탕으로 추진 배경, 필요성, 목적을 3~5개 글머리표로. 근거 없는 통계나 수치는 쓰지 말 것",
                   "Background, need and objectives in 3 to 5 bullets from the program notes; no invented statistics"),
    "plan": ("단계별 세부 추진 내용을 정리. 단계마다 1~2줄 설명과 체크리스트 항목을 하위 글머리표로",
             "What happens in each stage, one or two lines each, with the checklist items as sub bullets"),
    "schedule": ("단계별 일정을 표로 정리", "The stages and dates as a table"),
    "team": ("PM과 팀원, 역할 분담을 정리. 없는 사람을 만들지 말 것", "PM, team and roles; do not invent people"),
    "budget": ("견적 항목을 표로 정리하고 소계, 부가세, 합계를 표시. 금액은 절대 바꾸지 말 것",
               "The quote items as a table with subtotal, VAT and total; never change any amount"),
    "outcomes": ("성과 지표 목표를 표로 쓰고 기대 효과를 2~3개 글머리표로", "Indicator targets as a table, then 2 or 3 bullets on expected impact"),
}


def section_titles(kind: str, lang: str) -> list[tuple[str, str]]:
    return [(k, ko if lang == "ko" else en) for k, ko, en in SECTIONS[kind]]


def build_section(key: str, c: dict) -> str:
    """Markdown for one section, built only from the data (no AI)."""
    lg, p = c["lang"], c["program"]
    if key == "overview":
        tags = Counter(t.strip() for s in c["startups"] for t in (s.get("tags") or "").split(",") if t.strip())
        top = ", ".join(f"{k} {v}" for k, v in tags.most_common(6))
        lines = [f"- {tr('name', lg)} : {esc(p.get('name'))}"]
        if p.get("client"):
            lines.append(f"- {tr('client', lg)} : {esc(p['client'])}")
        if p.get("start") or p.get("end"):
            lines.append(f"- {tr('period', lg)} : {kdate(p.get('start'))} ~ {kdate(p.get('end'))}")
        if c["startups"]:
            lines.append(f"- {tr('startups', lg)} : {tr('count_co', lg, n=len(c['startups']))}")
        if top:
            lines.append(f"  - {tr('fields', lg)} : {esc(top)}")
        if p.get("pm") or p.get("team"):
            lines.append(f"- {tr('lead', lg)} : {esc(', '.join(x for x in (p.get('pm'), p.get('team')) if x))}")
        if p.get("summary") and c.get("kind") != "proposal":   # the proposal has its own background section
            lines.append(f"- {tr('goal', lg)}")
            lines += [f"  - {esc(x.strip(' -•'))}" for x in str(p['summary']).splitlines() if x.strip()]
        return "\n".join(lines)
    if key == "progress":
        st_ = c["stages"]
        d = sum(1 for s in st_ if s.get("status") == "done")
        head = f"- {tr('of_stages', lg, n=len(st_), d=d, a=kdate(c['asof']))}"
        rows = [[s.get("name"), f"{short(s.get('start'))} ~ {short(s.get('due'))}".strip(" ~"), s.get("status_label"),
                 s.get("note")] for s in st_]
        return head + "\n\n" + md_table([tr("stage", lg), tr("when", lg), tr("status", lg), tr("note", lg)], rows)
    if key == "kpi":
        if not c["kpis"]:
            return f"- {tr('no_kpi', lg)}"
        rows = [[k["name"], f"{k['target_s']}", f"{k['actual_s']}", k["rate_s"]] for k in c["kpis"]]
        return md_table([tr("kpi", lg), tr("target", lg), tr("actual", lg), tr("rate", lg)], rows)
    if key == "matching":
        m = c["matching"]
        if not m["events"]:
            return f"- {tr('no_match', lg)}"
        by = m["by"]
        lines = [f"- {tr('total_meet', lg, n=m['total'], d=by.get('Done', 0), x=by.get('No-Show', 0), u=by.get('Scheduled', 0) + by.get('Ongoing', 0))}",
                 f"- {tr('partners_n', lg, n=len(m['partners']))}"]
        rows = [[s["name"], s["total"], s["done"], s["noshow"]] for s in
                sorted(m["startups"].values(), key=lambda x: (-x["total"], x["name"]))]
        return "\n".join(lines) + "\n\n" + md_table([tr("company", lg), tr("meet", lg), tr("done", lg), tr("noshow", lg)], rows)
    if key == "startups":
        rows = []
        for s in c["startups"]:
            ms = c["matching"]["startups"].get(s["key"], {})
            outs = [o for o in c["outcomes"] if o.get("startup_key") == s["key"]]
            rows.append([s["name"], s.get("tags"), tr("done_intake" if c["intake"].get(s["key"], {}).get("status") == "done" else "todo_intake", lg),
                         f"{ms.get('done', 0)}/{ms.get('total', 0)}",
                         ", ".join(f"{o.get('type_label')} ({o.get('partner')})" for o in outs[:3])])
        return md_table([tr("company", lg), tr("fields", lg), tr("intake", lg), f"{tr('meet', lg)} ({tr('done', lg)})",
                         tr("outcome", lg)], rows)
    if key == "next":
        rest = [s for s in c["stages"] if s.get("status") != "done"]
        if not rest:
            return f"- {tr('no_pending', lg)}"
        return "\n".join(f"- {tr('pending', lg, s=esc(s.get('name')), d=kdate(s.get('due')) or '-')}" for s in rest)
    if key == "lessons":
        lines = [f"- {esc(k['name'])} : {k['actual_s']} / {k['target_s']} ({k['rate_s']})" for k in c["kpis"]]
        return "\n".join(lines + ["", tr("write_here", lg)])
    if key == "background":
        pts = [x.strip(" -•") for x in str(p.get("summary") or "").splitlines() if x.strip()]
        return "\n".join([f"- {esc(x)}" for x in pts] + (["", tr("bg_write", lg)] if len(pts) < 2 else []))
    if key == "plan":
        out = []
        for s in c["stages"]:
            when = f"{short(s.get('start'))} ~ {short(s.get('due'))}".strip(" ~")
            out.append(f"- {esc(s.get('name'))}" + (f" ({when})" if when else ""))
            out += [f"  - {esc(x.get('text'))}" for x in s.get("checklist") or [] if x.get("text")]
            if s.get("note"):
                out.append(f"  - {esc(s['note'])}")
        return "\n".join(out) or "-"
    if key == "schedule":
        rows = [[s.get("name"), f"{kdate(s.get('start'))} ~ {kdate(s.get('due'))}".strip(" ~"), s.get("owner")]
                for s in c["stages"]]
        return md_table([tr("stage", lg), tr("when", lg), tr("owner", lg)], rows) or "-"
    if key == "team":
        lines = []
        if p.get("pm"):
            lines.append(f"- {tr('pm', lg)} : {esc(p['pm'])}")
        if p.get("team"):
            lines.append(f"- {tr('team', lg)} : {esc(p['team'])}")
        owners = sorted({s.get("owner") for s in c["stages"] if s.get("owner")})
        if owners:
            lines.append(f"- {tr('by_stage', lg)}")
            for o in owners:
                lines.append(f"  - {esc(o)} : {esc(', '.join(s['name'] for s in c['stages'] if s.get('owner') == o))}")
        return "\n".join(lines) or "-"
    if key == "budget":
        q = p.get("quote") or {}
        cur = q.get("currency") or "KRW"
        unit = CURRENCY.get(cur, (cur, cur))[0 if lg == "ko" else 1]
        items = [it for it in q.get("items") or [] if it.get("item") or it.get("cat")]
        if not items:
            return f"- {tr('budget_text', lg)} : {esc(p['budget'])}" if p.get("budget") else f"- {tr('no_quote', lg)}"
        rows = []
        for it in items:
            try:
                amt = float(it.get("price") or 0) * float(it.get("qty") or 0)
            except (TypeError, ValueError):
                amt = 0
            rows.append([it.get("cat"), it.get("item"), money(it.get("price"), cur),
                         f"{money(it.get('qty'), 'USD')}{(' ' + it['unit']) if it.get('unit') else ''}", money(amt, cur), it.get("note")])
        sub, vat, tot = quote_totals(q)
        rows.append(["", tr("subtotal", lg), "", "", money(sub, cur), ""])
        if q.get("vat", True):
            rows.append(["", tr("vat", lg), "", "", money(vat, cur), ""])
        rows.append(["", tr("total", lg), "", "", money(tot, cur), ""])
        return md_table([tr("q_cat", lg), tr("q_item", lg), f"{tr('q_price', lg)} ({unit})", tr("q_qty", lg),
                         f"{tr('q_amount', lg)} ({unit})", tr("note", lg)], rows, blank=" ")
    if key == "outcomes":
        rows = [[k["name"], k["target_s"]] for k in c["kpis"] if k.get("target_s") not in (None, "", "-")]
        tbl = md_table([tr("kpi", lg), tr("target", lg)], rows)
        return (tbl + "\n\n" if tbl else "") + tr("expect_write", lg)
    return ""


def report_markdown(title: str, kind: str, sections: dict, lang: str) -> str:
    parts = [f"# {esc(title)}", ""]
    for key, name in section_titles(kind, lang):
        body = (sections.get(key) or "").strip()
        parts += [f"## {name}", "", body or "-", ""]
    return "\n".join(parts)


# ---------------------------------------------------------------- intake documents
def intake_markdown(startup: dict, rec: dict, questions: list[dict], lang: str, heading: int = 1) -> str:
    h = "#" * heading
    lg = lang
    out = [f"{h} {esc(startup.get('name'))}" + (f" ({esc(startup.get('kr'))})" if startup.get("kr") else ""), ""]
    out += [f"{h}# {tr('info', lg)}", "",
            md_table([tr("company", lg), tr("kr", lg), tr("ceo", lg), tr("fields", lg), tr("web", lg)],
                     [[startup.get("name"), startup.get("kr"), startup.get("ceo"), startup.get("tags"), startup.get("web")]]), ""]
    out += [f"{h}# {tr('overview_i', lg)}", "",
            f"- {tr('i_date', lg)} : {kdate(rec.get('date')) or '-'}",
            f"- {tr('i_by', lg)} : {esc(rec.get('interviewer')) or '-'}",
            f"- {tr('i_with', lg)} : {esc(rec.get('attendees')) or '-'}", ""]
    if rec.get("summary"):
        out += [f"{h}# {tr('summary', lg)}", "", rec["summary"].strip(), ""]
    out += [f"{h}# {tr('qa', lg)}", ""]
    ans = rec.get("answers") or {}
    for q in questions:
        a = (ans.get(q["id"]) or "").strip()
        out.append(f"- **{esc(q.get('text'))}**")
        if a:
            out += [f"  - {esc(x)}" for x in a.splitlines() if x.strip()]
        else:
            out.append(f"  - {tr('no_answer', lg)}")
    if (rec.get("notes") or "").strip():
        out += ["", f"{h}# {tr('memo', lg)}", ""] + [f"- {esc(x)}" for x in rec["notes"].splitlines() if x.strip()]
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------- export
def report_style():
    from dataclasses import replace

    from hwpx_automation.markdown import MarkdownStyle
    style = MarkdownStyle.report()
    try:   # navy and light blue title band instead of the four bright strips
        style = replace(style, banner_colors=("#1B1760", "#88A1EB", "#88A1EB", "#1B1760"))
    except Exception:  # noqa: BLE001
        pass
    return style


def to_hwpx(md: str, byline: str | None = None) -> tuple[bytes, list[str]]:
    """Markdown -> .hwpx bytes, plus any problems the open-safety check found (empty list = fine)."""
    import hwpx
    from hwpx_automation.markdown import from_markdown
    doc = from_markdown(md, style=report_style(), byline=byline or None)
    buf = io.BytesIO()
    doc.save(buf)
    data = buf.getvalue()
    issues: list[str] = []
    try:
        rep = hwpx.validate_editor_open_safety(data)
        if not rep.reopen_ok:
            issues.append(f"reopen failed: {rep.reopen_error}")
        issues += [str(e) for e in (rep.blocking_package_errors or ())]
        if rep.document_validation_error:
            issues.append(str(rep.document_validation_error))
    except Exception as e:  # noqa: BLE001
        issues.append(str(e))
    return data, issues


def preview_html(hwpx_bytes: bytes) -> str:
    """An approximate page preview (drawn by python-hwpx-automation, not by 한글)."""
    from hwpx_automation.preview import render_layout_preview
    return render_layout_preview(hwpx_bytes).html
