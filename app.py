"""
BizMatch Scheduler
==================
Generates optimal, conflict-free 1:1 business matching schedules
(startup x investor/corporate speed meetings) with Streamlit + Google OR-Tools CP-SAT.

Run locally:   streamlit run app.py
"""
from __future__ import annotations

import hashlib
import io
import json
import math
import uuid
import re
import time as _time
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import importlib
import i18n as _i18n
importlib.reload(_i18n)   # Streamlit Cloud reruns app.py after a push but can keep the old text module
from i18n import COL_LABELS, DIR_LABELS, LANGS, SUMMARY_KO, T, lang
from store import Store
import programs
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from ortools.sat.python import cp_model

# --------------------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------------------
APP_TITLE = "BizMatch"
BRAND = Path(__file__).parent / "static" / "brand"
LOCAL = components.declare_component("bizmatch_local", path=str(Path(__file__).parent / "local"))
SUMMARY = components.declare_component("bizmatch_summary", path=str(Path(__file__).parent / "summary"))
BOARD = components.declare_component("bizmatch_board", path=str(Path(__file__).parent / "board"))
STATUSES = ["Scheduled", "Ongoing", "Done", "No-Show"]
STATUS_CHOICES = ["Auto"] + STATUSES
LOCK_STATUSES = {"Ongoing", "Done", "No-Show"}
HOST_TYPES = ["Investor", "Corporate", "Agency", "Mentor", "Other"]
PACKING = {
    "front": "Front-packed: start at their availability, no gaps",
    "block": "Contiguous block: no gaps, may start later",
    "soft": "Soft: minimise gaps",
}
STARTUP_COLS = ["Name", "Presenter", "Domain", "Days", "Available From", "Available Until",
                "Blocked Times", "Priority", "Pitch Polishing", "Max Meetings"]
DEFAULT_NAMES = {"session": "Session", "time": "Time Slot", "table": "Table", "status": "Status",
                 "polish": "POLISHING SCHEDULE", "waiting": "WAITING LIST"}
ABSENT_WORDS = ("none", "no", "absent", "not here", "not attending", "cancelled", "canceled",
                "x", "dropped", "withdrawn", "불참", "미참석", "취소", "없음")
HOST_COLS = ["Name", "Contact", "Type", "Days", "Available From", "Available Until",
             "Blocked Times", "Only Domains", "Preferred Domains", "Excluded Domains", "Max Meetings"]
TIMEZONES = ["Asia/Singapore", "Asia/Seoul", "Asia/Tokyo", "Asia/Hong_Kong", "Asia/Shanghai",
             "Asia/Jakarta", "Asia/Bangkok", "Asia/Kolkata", "Asia/Dubai", "Europe/London",
             "Europe/Berlin", "America/New_York", "America/Los_Angeles", "UTC"]
TZ_HINT = {"America/Los_Angeles": "America/Los_Angeles (Las Vegas, CES)", "America/New_York": "America/New_York (Boston, NYC)",
           "Asia/Seoul": "Asia/Seoul (Seoul)", "Asia/Singapore": "Asia/Singapore (Singapore)"}
TIME_VALID = r"^$|^([01]?\d|2[0-4])[:.][0-5]\d$"
TIME_RE = re.compile(r"^(\d{1,2})[:.](\d{2})(?::\d{2})?$")
RANGE_RE = re.compile(
    r"^(?P<day>.*?)\s*(?P<a>\d{1,2}[:.]\d{2})\s*(?:-|–|to)\s*(?P<b>\d{1,2}[:.]\d{2})$", re.I)

# Objective weights. BASE dominates so the solver never trades meetings for fairness.
W_BASE = 10_000          # every meeting
W_PRIORITY_BONUS = 2_000  # extra per meeting for priority startups, per weight step
W_FAIR_STARTUP = 6_000    # concave fairness reward for startups: W/j for the j-th meeting
W_FAIR_HOST = 1_500       # concave fairness reward for hosts
W_GAP = 4_000             # soft packing: penalty per extra block
W_EARLY = 1               # tiny push toward earlier slots
W_POLISH = 60_000         # a pitch polishing session outranks any single meeting
W_PREFER = 3_000          # a host meeting a startup in its preferred domains (steers who, never costs a meeting)


# --------------------------------------------------------------------------------------
# Parsing helpers
# --------------------------------------------------------------------------------------
def is_blank(v) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip().lower() in ("", "nan", "none", "nat", "<na>", "-")


def to_min(v) -> int | None:
    """'14:30' / time(14,30) / '14.30' -> minutes since midnight."""
    if is_blank(v):
        return None
    if isinstance(v, (time, datetime)):
        return v.hour * 60 + v.minute
    s = str(v).strip()
    m = TIME_RE.match(s)
    if not m:
        raise ValueError(f"'{s}' is not a time, use HH:MM")
    h, mi = int(m.group(1)), int(m.group(2))
    if h > 24 or mi > 59 or (h == 24 and mi):
        raise ValueError(f"'{s}' is not a valid time")
    return h * 60 + mi


def fmt(m: float) -> str:
    """Minutes since midnight -> 'HH:MM' (or 'HH:MM:SS' when a gap has seconds)."""
    t = int(round(m * 60))
    h, rem = divmod(t, 3600)
    mi, sec = divmod(rem, 60)
    return f"{h:02d}:{mi:02d}" + (f":{sec:02d}" if sec else "")


def tstr(v) -> str:
    """Normalise any time-like cell to 'HH:MM' text (or '')."""
    try:
        m = to_min(v)
    except ValueError:
        return str(v)
    return "" if m is None else fmt(m)


def split_list(v) -> list[str]:
    if is_blank(v):
        return []
    return [t.strip() for t in re.split(r"[;,/|\n]+", str(v)) if t.strip()]


def norm(t: str) -> str:
    return re.sub(r"\s+", " ", str(t).strip().lower())


def to_bool(v) -> bool:
    if is_blank(v):
        return False
    if isinstance(v, bool):
        return v
    return norm(v) in ("true", "yes", "y", "1", "x", "o", "priority", "✓", "v")


def to_int(v) -> int | None:
    if is_blank(v):
        return None
    try:
        n = int(float(v))
    except (TypeError, ValueError):
        raise ValueError(f"'{v}' is not a number")
    return n if n > 0 else None


def parse_days(v, labels: list[str]) -> frozenset[int]:
    """Blank/All -> every day. 'Not here'/'None' -> no day. Accepts 'Day 2', '2', '1,2' or a day label."""
    if isinstance(v, str) and norm(v) in ABSENT_WORDS:
        return frozenset()
    if is_blank(v) or norm(v) in ("all", "any", "both", "every day", "전체", "모두", "전일", "양일"):
        return frozenset(range(len(labels)))
    if norm(v) in ABSENT_WORDS:
        return frozenset()
    out = set()
    lab = {norm(l): i for i, l in enumerate(labels)}
    for tok in re.split(r"[;,/|&+]+|\band\b", str(v)):
        t = norm(tok)
        if not t:
            continue
        if t in lab:
            out.add(lab[t])
            continue
        m = re.fullmatch(r"(?:day\s*|d)?(\d+)(?:\s*일차|\s*일)?", t)
        if m and 1 <= int(m.group(1)) <= len(labels):
            out.add(int(m.group(1)) - 1)
            continue
        raise ValueError(f"unknown day '{tok.strip()}'")
    return frozenset(out)


def parse_blocked(v, labels: list[str]) -> tuple:
    """'Day 1 14:00-15:00; 16:20-16:40' -> ((0, 840, 900), (None, 980, 1000))."""
    out = []
    for item in split_list(v):
        m = RANGE_RE.match(item)
        if not m:
            raise ValueError(f"blocked time '{item}' should look like 'Day 1 14:00-15:00'")
        a, b = to_min(m.group("a")), to_min(m.group("b"))
        if b <= a:
            raise ValueError(f"blocked time '{item}' ends before it starts")
        day = m.group("day").strip()
        if day:
            for d in parse_days(day, labels):
                out.append((d, a, b))
        else:
            out.append((None, a, b))
    return tuple(out)


# --------------------------------------------------------------------------------------
# Domain model
# --------------------------------------------------------------------------------------
@dataclass
class EventConfig:
    name: str = "Business Matching"
    days: list = field(default_factory=lambda: [{"label": "Day 1", "date": ""}])
    start: int = 600
    end: int = 1020
    interval: int = 20
    buffer: int = 0
    lunch_start: int | None = 720
    lunch_end: int | None = 780
    tables: int = 6
    tz: str = "Asia/Singapore"
    sessions: list = field(default_factory=list)      # [{"name": "Session A", "start": 600, "end": 720}]
    pol_days: list = field(default_factory=list)      # day indexes with pitch polishing
    pol_start: int | None = None
    pol_end: int | None = None
    pol_cap: int = 1                                  # startups polished at the same time
    names: dict = field(default_factory=dict)         # custom column names for the schedule

    @property
    def labels(self) -> list[str]:
        return [d["label"] for d in self.days]

    def has_lunch(self) -> bool:
        return (self.lunch_start is not None and self.lunch_end is not None
                and self.lunch_end > self.lunch_start
                and self.lunch_start < self.end and self.lunch_end > self.start)

    def day_date(self, d: int) -> date | None:
        v = self.days[d].get("date") if d < len(self.days) else None
        if not v:
            return None
        try:
            return date.fromisoformat(str(v)[:10])
        except ValueError:
            return None

    def day_title(self, d: int) -> str:
        dd = self.day_date(d)
        return f"{self.days[d]['label']}" + (f" · {dd.strftime('%a %d %b %Y')}" if dd else "")

    def nm(self, key: str) -> str:
        return (self.names or {}).get(key) or DEFAULT_NAMES[key]

    def day_short(self, d: int) -> str:
        dd = self.day_date(d)
        return dd.strftime("%d %b") if dd else self.days[d]["label"]

    def session_list(self) -> list[dict]:
        if self.sessions:
            return sorted(self.sessions, key=lambda x: x["start"])
        if self.has_lunch():
            return [{"name": "Session A", "start": self.start, "end": self.lunch_start},
                    {"name": "Session B", "start": self.lunch_end, "end": self.end}]
        return [{"name": "Session A", "start": self.start, "end": self.end}]

    def session_of(self, start: int) -> int:
        for i, x in enumerate(self.session_list()):
            if x["start"] <= start < x["end"]:
                return i
        return -1

    def polish_ok(self, sl: "Slot") -> bool:
        return (sl.day in self.pol_days and (self.pol_start is None or sl.start >= self.pol_start)
                and (self.pol_end is None or sl.end <= self.pol_end))


def parse_sessions(v) -> list[dict]:
    """'Session A 10:00-12:00; Session B 13:00-15:00' -> list of sessions."""
    out = []
    for i, item in enumerate(split_list(v)):
        m = RANGE_RE.match(item)
        if not m:
            raise ValueError(f"session '{item}' should look like 'Session A 10:00-12:00'")
        a, b = to_min(m.group("a")), to_min(m.group("b"))
        if b <= a:
            raise ValueError(f"session '{item}' ends before it starts")
        out.append({"name": m.group("day").strip() or f"Session {chr(65 + i)}", "start": a, "end": b})
    return out

@dataclass(frozen=True)
class Slot:
    k: int        # global index
    day: int
    idx: int      # index within the day
    start: int
    end: int


@dataclass(frozen=True)
class Party:
    name: str
    role: str                     # "startup" | "host"
    days: frozenset
    frm: int | None = None
    until: int | None = None
    blocked: tuple = ()
    max_m: int | None = None
    domains: frozenset = frozenset()   # startups
    excluded: frozenset = frozenset()  # hosts: never these tags
    only: frozenset = frozenset()      # hosts: only startups with one of these tags
    prefer: frozenset = frozenset()    # hosts: rather startups with these tags
    only_text: str = ""
    prefer_text: str = ""
    priority: bool = False
    kind: str = ""                     # host type
    contact: str = ""                  # host contact person
    domain_text: str = ""
    polish: bool = True                # startup takes part in pitch polishing
    presenter: str = ""                # startup person who pitches / attends

    @property
    def absent(self) -> bool:
        return not self.days

    def available(self, sl: Slot) -> bool:
        if sl.day not in self.days:
            return False
        if self.frm is not None and sl.start < self.frm:
            return False
        if self.until is not None and sl.end > self.until:
            return False
        for d, a, b in self.blocked:
            if (d is None or d == sl.day) and sl.start < b and sl.end > a:
                return False
        return True


def host_label(h) -> str:
    """'Raymond · Focus Tech' when a contact person is set."""
    if isinstance(h, str):
        return h
    return f"{h.contact} · {h.name}" if h.contact else h.name


def _tok(tag: str) -> frozenset:
    return frozenset(re.findall(r"[0-9a-z가-힣]+", tag.lower()))


def tags_hit(wanted, have) -> bool:
    """True when any wanted tag matches any of the startup's tags. 'AI' matches 'AI', 'AI/ML' and 'Gen AI',
    'B2C' matches 'B2C' and 'B2C Commerce' (every word of the wanted tag must appear in the startup tag)."""
    for w in wanted:
        wt = _tok(w)
        for h in have:
            if w == h or (wt and wt <= _tok(h)):
                return True
    return False


def compatible(s: Party, h: Party) -> bool:
    if h.excluded and tags_hit(h.excluded, s.domains):
        return False
    return not h.only or tags_hit(h.only, s.domains)


def mismatch(s: Party, h: Party) -> list | None:
    """Why a startup falls outside a partner's domain rules: ["ex", tags] or ["only", tags], else None."""
    if h.excluded and tags_hit(h.excluded, s.domains):
        raw = [t.strip() for t in s.domain_text.split(",") if t.strip()]
        return ["ex", ", ".join(t for t in raw if tags_hit(h.excluded, {norm(t)})) or ", ".join(sorted(h.excluded))]
    if h.only and not tags_hit(h.only, s.domains):
        return ["only", h.only_text]
    return None


def prefers(h: Party, s: Party) -> bool:
    return bool(h.prefer) and tags_hit(h.prefer, s.domains)


@dataclass
class Meeting:
    day: int
    k: int
    start: int
    end: int
    startup: str
    host: str
    table: int = 0
    lock: bool = False     # moved by hand: kept when the schedule is generated again

    @property
    def mid(self) -> str:
        """Stable id. A startup and a host meet at most once, so the pair identifies the meeting."""
        return f"{self.startup}|{self.host}"


@dataclass
class Polish:
    day: int
    k: int
    start: int
    end: int
    startup: str
    n: int = 1          # 1 = first polishing session for this startup, 2 = second, ...

    @property
    def pid(self) -> str:
        return f"{self.startup}|D{self.day}"


def number_polish(polish: list[Polish]) -> list[Polish]:
    seen = defaultdict(int)
    for p in sorted(polish, key=lambda p: (p.k, p.startup)):
        seen[p.startup] += 1
        p.n = seen[p.startup]
    return sorted(polish, key=lambda p: (p.k, p.startup))


def build_slots(cfg: EventConfig) -> list[Slot]:
    out: list[Slot] = []
    if cfg.interval < 5 or cfg.end <= cfg.start:
        return out
    # work in whole seconds so gaps like 1 min 10 s add up exactly
    L, B = int(round(cfg.interval * 60)), int(round(max(0, cfg.buffer) * 60))
    end = int(cfg.end * 60)
    ls, le = (int(cfg.lunch_start * 60), int(cfg.lunch_end * 60)) if cfg.has_lunch() else (None, None)
    for d in range(len(cfg.days)):
        t, i = int(cfg.start * 60), 0
        while t + L <= end and len(out) < 5000:
            if ls is not None and t < le and t + L > ls:
                t = le
                continue
            out.append(Slot(len(out), d, i, t / 60, (t + L) / 60))
            i += 1
            t += L + B
    return out


def parse_parties(df: pd.DataFrame, role: str, labels: list[str]) -> tuple[list[Party], list[str]]:
    parties, errors, seen = [], [], set()
    what = "Startups" if role == "startup" else "Hosts"
    for i, row in df.reset_index(drop=True).iterrows():
        name = "" if is_blank(row.get("Name")) else str(row.get("Name")).strip()
        if not name:
            if any(not is_blank(row.get(c)) for c in df.columns if c not in ("Priority",)):
                errors.append(f"{what} row {i + 1}: name is missing")
            continue
        where = f"{what} row {i + 1} ({name})"
        if norm(name) in seen:
            errors.append(f"{where}: duplicate name")
            continue
        seen.add(norm(name))
        try:
            frm, until = to_min(row.get("Available From")), to_min(row.get("Available Until"))
            if frm is not None and until is not None and until <= frm:
                raise ValueError("Available Until must be after Available From")
            kw = dict(
                name=name, role=role,
                days=parse_days(row.get("Days"), labels),
                frm=frm, until=until,
                blocked=parse_blocked(row.get("Blocked Times"), labels),
                max_m=to_int(row.get("Max Meetings")),
            )
            if role == "startup":
                doms = split_list(row.get("Domain"))
                pv = row.get("Pitch Polishing")
                kw.update(domains=frozenset(norm(x) for x in doms),
                          priority=to_bool(row.get("Priority")), domain_text=", ".join(doms),
                          polish=True if is_blank(pv) else to_bool(pv),
                          presenter="" if is_blank(row.get("Presenter")) else str(row.get("Presenter")).strip())
            else:
                kind = "" if is_blank(row.get("Type")) else str(row.get("Type")).strip()
                kw.update(excluded=frozenset(norm(x) for x in split_list(row.get("Excluded Domains"))),
                          only=frozenset(norm(x) for x in split_list(row.get("Only Domains"))),
                          prefer=frozenset(norm(x) for x in split_list(row.get("Preferred Domains"))),
                          only_text=", ".join(split_list(row.get("Only Domains"))),
                          prefer_text=", ".join(split_list(row.get("Preferred Domains"))),
                          kind=kind or "Investor",
                          contact="" if is_blank(row.get("Contact")) else str(row.get("Contact")).strip())
            parties.append(Party(**kw))
        except ValueError as e:
            errors.append(f"{where}: {e}")
    return parties, errors


# --------------------------------------------------------------------------------------
# Solver
# --------------------------------------------------------------------------------------
@dataclass
class SolverSettings:
    packing: str = "front"
    time_limit: float = 20.0
    priority_weight: int = 3
    host_fairness: bool = True
    prefer_early: bool = True
    workers: int = 8


@dataclass
class SolveResult:
    status: str
    meetings: list
    message: str
    wall: float
    n_vars: int
    objective: float | None = None
    bound: float | None = None
    kept: int = 0
    dropped_fixed: list = field(default_factory=list)
    polish: list = field(default_factory=list)


def greedy_schedule(cfg, slots, startups, hosts, settings, x, q, fixed_x=(), fixed_q=()):
    """Fast first schedule, slot by slot, that keeps every hard rule. It is handed to CP-SAT as a starting
    point (big events: 30+ partners) and is also the fallback if the solver cannot beat it in time.
    Returns (chosen x keys, chosen q keys)."""
    S, H = len(startups), len(hosts)
    busy_s, busy_h = set(), set()            # (si, k), (hi, k)
    used_pair, n_s, n_h = set(), [0] * S, [0] * H
    load, pol_load = defaultdict(int), defaultdict(int)
    chosen_x, chosen_q = [], []

    def take(si, hi, k):
        chosen_x.append((si, hi, k)); busy_s.add((si, k)); busy_h.add((hi, k)); used_pair.add((si, hi))
        n_s[si] += 1; n_h[hi] += 1; load[k] += 1

    for key in fixed_x:
        if key in x and key[:2] not in used_pair:
            take(*key)
    for si, k in fixed_q:
        if (si, k) in q:
            chosen_q.append((si, k)); busy_s.add((si, k)); pol_load[k] += 1
    # pitch polishing: one per startup per polishing day, spread over the allowed slots
    pol_days_done = {(si, slots[k].day) for si, k in chosen_q}
    by_sd = defaultdict(list)
    for (si, k) in q:
        by_sd[si, slots[k].day].append(k)
    for (si, d), ks in sorted(by_sd.items(), key=lambda t: (t[0][1], len(t[1]))):
        if (si, d) in pol_days_done:
            continue
        ok = [k for k in ks if (si, k) not in busy_s and pol_load[k] < max(1, cfg.pol_cap)]
        if ok:
            k = min(ok, key=lambda k: (pol_load[k], k))
            chosen_q.append((si, k)); busy_s.add((si, k)); pol_load[k] += 1

    pairs_at = defaultdict(list)                  # k -> [(si, hi)]
    for (si, hi, k) in x:
        pairs_at[k].append((si, hi))
    cap_s = [s.max_m if s.max_m is not None else 10**9 for s in startups]
    cap_h = [h.max_m if h.max_m is not None else 10**9 for h in hosts]
    av_h = [[h.available(sl) for sl in slots] for h in hosts]
    first_k, prev_av = {}, {}                     # per (hi, day): first available slot, previous available slot
    for hi in range(H):
        last = {}
        for sl in slots:
            if av_h[hi][sl.k]:
                first_k.setdefault((hi, sl.day), sl.k)
                prev_av[hi, sl.k] = last.get(sl.day)
                last[sl.day] = sl.k
    started = set()                               # (hi, day) for block packing
    for sl in slots:
        k = sl.k
        cand = defaultdict(list)
        for si, hi in pairs_at[k]:
            cand[hi].append(si)
        def may_meet(hi):
            if (hi, k) in busy_h or n_h[hi] >= cap_h[hi]:
                return False
            pk = prev_av.get((hi, k))
            if settings.packing == "front":
                return k == first_k.get((hi, sl.day)) or (pk is not None and (hi, pk) in busy_h)
            if settings.packing == "block":
                return (hi, sl.day) not in started or (pk is not None and (hi, pk) in busy_h)
            return True
        hs = sorted((hi for hi in cand if may_meet(hi)),
                    key=lambda hi: (0 if prev_av.get((hi, k)) is not None and (hi, prev_av[hi, k]) in busy_h else 1,
                                    n_h[hi], hi))
        for hi in hs:
            if load[k] >= cfg.tables:
                break
            ok = [si for si in cand[hi] if (si, k) not in busy_s and (si, hi) not in used_pair and n_s[si] < cap_s[si]]
            if not ok:
                continue
            si = min(ok, key=lambda si: (n_s[si] - (0.5 if startups[si].priority else 0),
                                         0 if prefers(hosts[hi], startups[si]) else 1, (si + k) % S))
            take(si, hi, k)
            started.add((hi, sl.day))
    return chosen_x, chosen_q


def solve_schedule(cfg: EventConfig, slots: list[Slot], startups: list[Party], hosts: list[Party],
                   settings: SolverSettings, fixed: list[tuple] = (),
                   fixed_polish: list[tuple] = ()) -> SolveResult:
    """fixed: [(startup, host, day, start_minute)] meetings that must stay as they are.
    fixed_polish: [(startup, day, start_minute)] polishing sessions that must stay."""
    t0 = _time.time()
    model = cp_model.CpModel()
    av_s = [[s.available(sl) for sl in slots] for s in startups]
    av_h = [[h.available(sl) for sl in slots] for h in hosts]

    x: dict[tuple[int, int, int], cp_model.IntVar] = {}
    by_pair, by_sk, by_hk, by_k = defaultdict(list), defaultdict(list), defaultdict(list), defaultdict(list)
    by_s, by_h = defaultdict(list), defaultdict(list)
    for si, s in enumerate(startups):
        for hi, h in enumerate(hosts):
            if not compatible(s, h):
                continue
            for sl in slots:
                if av_s[si][sl.k] and av_h[hi][sl.k]:
                    v = model.NewBoolVar(f"x_{si}_{hi}_{sl.k}")
                    x[si, hi, sl.k] = v
                    by_pair[si, hi].append(v)
                    by_sk[si, sl.k].append(v)
                    by_hk[hi, sl.k].append(v)
                    by_k[sl.k].append(v)
                    by_s[si].append(v)
                    by_h[hi].append(v)

    # Pitch polishing: q[si, k] = startup si is in polishing during slot k
    q: dict[tuple[int, int], cp_model.IntVar] = {}
    q_k, q_sd = defaultdict(list), defaultdict(list)
    for si, s in enumerate(startups):
        if not s.polish:
            continue
        for sl in slots:
            if av_s[si][sl.k] and cfg.polish_ok(sl):
                v = model.NewBoolVar(f"q_{si}_{sl.k}")
                q[si, sl.k] = v
                q_k[sl.k].append(v)
                q_sd[si, sl.day].append(v)
    for vs in q_sd.values():
        model.AddAtMostOne(vs)                     # one polishing session per startup per day
    for vs in q_k.values():
        if len(vs) > cfg.pol_cap:
            model.Add(sum(vs) <= max(1, cfg.pol_cap))

    # Fixed meetings (live re-planning)
    s_ix = {s.name: i for i, s in enumerate(startups)}
    h_ix = {h.name: i for i, h in enumerate(hosts)}
    k_ix = {(sl.day, sl.start): sl.k for sl in slots}
    kept, dropped = 0, []
    for sn, hn, d, start in fixed:
        key = (s_ix.get(sn), h_ix.get(hn), k_ix.get((d, start)))
        if key in x:
            model.Add(x[key] == 1)
            kept += 1
        else:
            dropped.append(f"{sn} with {hn} at {fmt(start)}")
    for sn, d, start in fixed_polish:
        key = (s_ix.get(sn), k_ix.get((d, start)))
        if key in q:
            model.Add(q[key] == 1)
            kept += 1
        else:
            dropped.append(f"{sn} pitch polishing at {fmt(start)}")

    # 1. Each pair meets at most once across the whole event
    for vs in by_pair.values():
        model.AddAtMostOne(vs)
    # 2. No double booking for startups (meetings and polishing) or hosts
    for key in set(by_sk) | set(q):
        vs = by_sk.get(key, []) + ([q[key]] if key in q else [])
        if len(vs) > 1:
            model.AddAtMostOne(vs)
    for vs in by_hk.values():
        model.AddAtMostOne(vs)
    # 3. Table capacity per slot
    for vs in by_k.values():
        if len(vs) > cfg.tables:
            model.Add(sum(vs) <= cfg.tables)
    # 4. Max meetings per participant
    for si, s in enumerate(startups):
        if s.max_m is not None and len(by_s[si]) > s.max_m:
            model.Add(sum(by_s[si]) <= s.max_m)
    for hi, h in enumerate(hosts):
        if h.max_m is not None and len(by_h[hi]) > h.max_m:
            model.Add(sum(by_h[hi]) <= h.max_m)

    terms: list[tuple] = []

    # 5. Host slot packing (per host, per day, over the host's available slots in order;
    #    the lunch break is not counted as a gap)
    for hi, h in enumerate(hosts):
        for d in range(len(cfg.days)):
            ks = [sl.k for sl in slots if sl.day == d and av_h[hi][sl.k]]
            if settings.packing == "front":
                for a, b in zip(ks, ks[1:]):
                    if by_hk[hi, b]:
                        model.Add(sum(by_hk[hi, b]) <= sum(by_hk[hi, a]))
            else:
                starts, prev = [], None
                for k in ks:
                    if by_hk[hi, k]:
                        stv = model.NewBoolVar(f"start_{hi}_{k}")
                        if prev is not None and by_hk[hi, prev]:
                            model.Add(stv >= sum(by_hk[hi, k]) - sum(by_hk[hi, prev]))
                        else:
                            model.Add(stv >= sum(by_hk[hi, k]))
                        starts.append(stv)
                    prev = k
                if settings.packing == "block" and starts:
                    model.Add(sum(starts) <= 1)
                elif settings.packing == "soft":
                    terms += [(v, -W_GAP) for v in starts]

    # Objective: total meetings first, then priority, then fairness, then earliness
    pw = max(1, int(settings.priority_weight))
    for (si, hi, _k), v in x.items():
        if prefers(hosts[hi], startups[si]):
            terms.append((v, W_PREFER))
    for si, s in enumerate(startups):
        vs = by_s[si]
        if not vs:
            continue
        w = pw if s.priority else 1
        base = W_BASE + (W_PRIORITY_BONUS * (pw - 1) if s.priority else 0)
        terms += [(v, base) for v in vs]
        ub = min(len({hi for (a, hi) in by_pair if a == si}),
                 len({k for (a, k) in by_sk if a == si}),
                 s.max_m or 10**9)
        z = [model.NewBoolVar(f"zs_{si}_{j}") for j in range(ub)]
        for a, b in zip(z, z[1:]):
            model.AddImplication(b, a)
        model.Add(sum(z) == sum(vs))
        terms += [(zj, int(W_FAIR_STARTUP * w / j)) for j, zj in enumerate(z, 1)]
    if settings.host_fairness:
        for hi, h in enumerate(hosts):
            vs = by_h[hi]
            if not vs:
                continue
            ub = min(len({k for (a, k) in by_hk if a == hi and by_hk[a, k]}), h.max_m or 10**9,
                     len({si for (si, a) in by_pair if a == hi}))
            z = [model.NewBoolVar(f"zh_{hi}_{j}") for j in range(ub)]
            for a, b in zip(z, z[1:]):
                model.AddImplication(b, a)
            model.Add(sum(z) == sum(vs))
            terms += [(zj, int(W_FAIR_HOST / j)) for j, zj in enumerate(z, 1)]
    terms += [(v, W_POLISH) for v in q.values()]
    if settings.prefer_early:
        idx = {sl.k: sl.idx for sl in slots}
        terms += [(v, -W_EARLY * idx[k]) for (si, hi, k), v in x.items() if idx[k]]

    if terms:
        model.Maximize(cp_model.LinearExpr.WeightedSum([t[0] for t in terms], [t[1] for t in terms]))

    # Warm start: a quick greedy schedule as a hint, so big events start from a full schedule
    fx = [(s_ix.get(sn), h_ix.get(hn), k_ix.get((d, st_))) for sn, hn, d, st_ in fixed]
    fq = [(s_ix.get(sn), k_ix.get((d, st_))) for sn, d, st_ in fixed_polish]
    gx, gq = greedy_schedule(cfg, slots, startups, hosts, settings, x, q, fx, fq)
    gxs, gqs = set(gx), set(gq)
    for key, v in x.items():
        model.AddHint(v, 1 if key in gxs else 0)
    for key, v in q.items():
        model.AddHint(v, 1 if key in gqs else 0)

    solver = cp_model.CpSolver()
    # bigger events get more time automatically (about 1 extra second per 1,500 choices, up to 2 minutes)
    solver.parameters.max_time_in_seconds = float(max(settings.time_limit, min(120.0, 10 + len(x) / 1500)))
    solver.parameters.num_workers = int(settings.workers)
    code = solver.Solve(model)
    name = solver.StatusName(code)
    wall = _time.time() - t0
    use_greedy = code not in (cp_model.OPTIMAL, cp_model.FEASIBLE) or \
        sum(1 for v in x.values() if solver.BooleanValue(v)) < len(gx)
    if use_greedy and gx and not dropped:
        sl_by_k = {sl.k: sl for sl in slots}
        meetings = [Meeting(sl_by_k[k].day, k, sl_by_k[k].start, sl_by_k[k].end, startups[si].name, hosts[hi].name)
                    for si, hi, k in gx]
        assign_tables(meetings, cfg.tables, hosts)
        polish = number_polish([Polish(sl_by_k[k].day, k, sl_by_k[k].start, sl_by_k[k].end, startups[si].name)
                                for si, k in gq])
        return SolveResult("FEASIBLE", meetings, "Quick schedule (the solver could not improve it within the time limit).",
                           wall, len(x) + len(q), None, None, kept, dropped, polish)

    if code not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        msg = {
            "INFEASIBLE": "No schedule satisfies every hard rule. Usually the meetings kept from the "
                          "live board clash with the packing rule. Try the Soft packing mode.",
            "UNKNOWN": "The solver ran out of time before finding a schedule. Raise the time limit.",
            "MODEL_INVALID": "The model is invalid. Please report this with your input file.",
        }.get(name, name)
        return SolveResult(name, [], msg, wall, len(x), kept=kept, dropped_fixed=dropped)

    sl_by_k = {sl.k: sl for sl in slots}
    meetings = [Meeting(sl_by_k[k].day, k, sl_by_k[k].start, sl_by_k[k].end,
                        startups[si].name, hosts[hi].name)
                for (si, hi, k), v in x.items() if solver.BooleanValue(v)]
    assign_tables(meetings, cfg.tables, hosts)
    polish = number_polish([Polish(sl_by_k[k].day, k, sl_by_k[k].start, sl_by_k[k].end, startups[si].name)
                            for (si, k), v in q.items() if solver.BooleanValue(v)])
    msg = ("Optimal schedule found." if code == cp_model.OPTIMAL else
           "Good schedule found within the time limit (not proven optimal).")
    return SolveResult(name, meetings, msg, wall, len(x) + len(q), solver.ObjectiveValue(),
                       solver.BestObjectiveBound(), kept, dropped, polish)


def assign_tables(meetings: list[Meeting], tables: int, hosts: list[Party]) -> None:
    """Give each meeting a table so hosts move as little as possible.
    1. A host continuing from the previous slot keeps its table.
    2. A host coming back keeps its last table if it is free.
    3. Anyone else gets the free table that has been empty the longest."""
    order = {h.name: i for i, h in enumerate(hosts)}
    by_k = defaultdict(list)
    for m in meetings:
        by_k[m.k].append(m)
    last: dict[tuple[int, str], tuple[int, int]] = {}   # (day, host) -> (table, k)
    freed: dict[tuple[int, int], int] = {}              # (day, table) -> k when it was last used
    prev_k: dict[int, int] = {}
    for k in sorted(by_k):
        ms = sorted(by_k[k], key=lambda m: order.get(m.host, 0))
        day = ms[0].day
        used = set()
        cont = [m for m in ms if last.get((day, m.host), (0, -9))[1] == prev_k.get(day)]
        back = [m for m in ms if m not in cont and (day, m.host) in last]
        for m in cont + back:
            t = last[(day, m.host)][0]
            if t not in used:
                m.table = t
                used.add(t)
        for m in ms:
            if m.table:
                continue
            free = [t for t in range(1, tables + 1) if t not in used]
            t = min(free, key=lambda t: (freed.get((day, t), -1), t))
            m.table = t
            used.add(t)
        for m in ms:
            last[(day, m.host)] = (m.table, k)
            freed[(day, m.table)] = k
        prev_k[day] = k
    meetings.sort(key=lambda m: (m.k, m.table))


def validate(cfg: EventConfig, slots: list[Slot], startups: list[Party], hosts: list[Party],
             meetings: list[Meeting], packing: str, polish: list[Polish] = ()) -> list[str]:
    """Independent check of every hard rule on the final schedule."""
    issues = []
    S = {s.name: s for s in startups}
    H = {h.name: h for h in hosts}
    sl = {s.k: s for s in slots}
    pairs, sk, hk, tk = set(), set(), set(), set()
    count_s, count_h = defaultdict(int), defaultdict(int)
    for m in meetings:
        s, h = S.get(m.startup), H.get(m.host)
        if not s or not h:
            issues.append(f"Unknown participant in {m.mid}")
            continue
        if (m.startup, m.host) in pairs:
            issues.append(f"Duplicate meeting: {m.startup} with {m.host}")
        pairs.add((m.startup, m.host))
        for key, bag, label in (((m.startup, m.k), sk, m.startup), ((m.host, m.k), hk, m.host)):
            if key in bag:
                issues.append(f"Double booking: {label} at {fmt(m.start)}")
            bag.add(key)
        if not 1 <= m.table <= cfg.tables or (m.table, m.k) in tk:
            issues.append(f"Table clash at {fmt(m.start)} table {m.table}")
        tk.add((m.table, m.k))
        if not (s.available(sl[m.k]) and h.available(sl[m.k])):
            issues.append(f"Outside availability: {m.startup} with {m.host} at {fmt(m.start)}")
        if not compatible(s, h):
            issues.append(f"Note: {m.startup} with {m.host} is outside {m.host}'s domain rules (added by hand)")
        count_s[m.startup] += 1
        count_h[m.host] += 1
    psd, pcap = set(), defaultdict(int)
    for p in polish:
        s = S.get(p.startup)
        if not s or not s.available(sl[p.k]) or not cfg.polish_ok(sl[p.k]):
            issues.append(f"Pitch polishing outside availability: {p.startup} at {fmt(p.start)}")
        if (p.startup, p.k) in sk:
            issues.append(f"Double booking: {p.startup} has a meeting during pitch polishing at {fmt(p.start)}")
        if (p.startup, p.day) in psd:
            issues.append(f"{p.startup} has two pitch polishing sessions on {cfg.days[p.day]['label']}")
        psd.add((p.startup, p.day))
        pcap[p.k] += 1
        if pcap[p.k] > max(1, cfg.pol_cap):
            issues.append(f"Too many pitch polishing sessions at {fmt(p.start)}")
    for p in startups + hosts:
        c = count_s[p.name] if p.role == "startup" else count_h[p.name]
        if p.max_m and c > p.max_m:
            issues.append(f"{p.name} exceeds max meetings ({c} > {p.max_m})")
    if packing in ("front", "block"):
        for h in hosts:
            for d in range(len(cfg.days)):
                pat = "".join("1" if (h.name, s.k) in hk else "0"
                              for s in slots if s.day == d and h.available(s))
                core = pat.rstrip("0") if packing == "front" else pat.strip("0")
                if "0" in core:
                    issues.append(f"Gap in {h.name}'s {cfg.days[d]['label']} schedule")
    return issues


# --------------------------------------------------------------------------------------
# Derived views
# --------------------------------------------------------------------------------------
def now_for(cfg: EventConfig) -> datetime:
    try:
        return datetime.now(ZoneInfo(cfg.tz))
    except Exception:
        return datetime.now(ZoneInfo("UTC"))


def auto_status(day: int, start: int, end: int, cfg: EventConfig, now: datetime) -> str:
    dd = cfg.day_date(day)
    if dd is None:
        return "Scheduled"
    today, nm = now.date(), now.hour * 60 + now.minute
    if dd < today:
        return "Done"
    if dd > today:
        return "Scheduled"
    return "Scheduled" if nm < start else ("Ongoing" if nm < end else "Done")


def eff_status(m: Meeting, cfg: EventConfig, statuses: dict, now: datetime) -> str:
    s = statuses.get(m.mid)
    return s if s in STATUSES else auto_status(m.day, m.start, m.end, cfg, now)


def pol_status(p: Polish, cfg: EventConfig, statuses: dict, now: datetime) -> str:
    s = statuses.get(p.pid)
    return s if s in STATUSES else auto_status(p.day, p.start, p.end, cfg, now)


def day_rows(day: int, cfg: EventConfig, slots: list[Slot], startups: list[Party],
             hosts: list[Party], meetings: list[Meeting], compact: bool = False,
             polish: list[Polish] = ()) -> list[dict]:
    """Rows for one day's timeline.
    Row kinds: slot rows, 'lunch', 'divider' (start of a new session without a lunch before it)
    and 'done' (compact=True folds the empty tail into one row)."""
    by_k, pol_k = defaultdict(list), defaultdict(list)
    for m in meetings:
        by_k[m.k].append(m)
    for p in polish:
        pol_k[p.k].append(p)
    ds = [s for s in slots if s.day == day]
    rows, lunch_added = [], not cfg.has_lunch()
    lunch = {"lunch": True, "start": cfg.lunch_start, "end": cfg.lunch_end,
             "time": f"{fmt(cfg.lunch_start)}–{fmt(cfg.lunch_end)}" if cfg.has_lunch() else ""}
    sessions = cfg.session_list()
    prev_sess = None
    for sl in ds:
        if not lunch_added and sl.start >= cfg.lunch_end:
            rows.append(lunch)
            lunch_added = True
        si = cfg.session_of(sl.start)
        if prev_sess is not None and si != prev_sess and si >= 0 and not rows[-1].get("lunch"):
            rows.append({"lunch": False, "divider": True, "session": si, "time": "",
                         "label": sessions[si]["name"]})
        prev_sess = si
        tables = [None] * cfg.tables
        for m in by_k[sl.k]:
            if 1 <= m.table <= cfg.tables:
                tables[m.table - 1] = m
        busy_s = {m.startup for m in by_k[sl.k]} | {p.startup for p in pol_k[sl.k]}
        busy_h = {m.host for m in by_k[sl.k]}
        rows.append({
            "lunch": False, "slot": sl, "start": sl.start, "end": sl.end, "session": si,
            "time": f"{fmt(sl.start)}–{fmt(sl.end)}", "tables": tables,
            "polish": sorted(pol_k[sl.k], key=lambda p: p.startup),
            "waiting": [s.name for s in startups if s.available(sl) and s.name not in busy_s],
            "resting": [h.name for h in hosts if h.available(sl) and h.name not in busy_h],
        })
    if not lunch_added and ds and ds[-1].end <= cfg.lunch_start:
        rows.append(lunch)
    if compact:
        last = max((i for i, r in enumerate(rows) if r.get("tables") and (any(r["tables"]) or r["polish"])),
                   default=-1)
        if last < len(rows) - 1:
            at = rows[last]["end"] if last >= 0 else None
            rows = rows[:last + 1] + [{"lunch": False, "done": True, "at": at, "time": ""}]
    return rows


def attendance(cfg: EventConfig, p: Party) -> str:
    if p.absent:
        return "Not attending"
    if len(p.days) == len(cfg.days):
        return "All days"
    return ", ".join(cfg.day_short(d) for d in sorted(p.days)) + " only"


def startup_summary(cfg, startups, hosts, meetings, statuses, now, sep=" • ", polish=()) -> pd.DataFrame:
    hl = {h.name: host_label(h) for h in hosts}
    pol = defaultdict(list)
    for p in polish:
        pol[p.startup].append(p)
    by_s = defaultdict(list)
    for m in meetings:
        by_s[m.startup].append(m)
    recs = []
    for s in startups:
        ms = sorted(by_s[s.name], key=lambda m: m.k)
        rec = {
            "Startup Name": s.name,
            "Presenter": s.presenter,
            "Attendance": attendance(cfg, s),
            "Domain": s.domain_text,
            "Priority": "★" if s.priority else "",
            "Total Meetings": len(ms),
            "No-Shows": sum(eff_status(m, cfg, statuses, now) == "No-Show" for m in ms),
            "Eligible Hosts": sum(compatible(s, h) for h in hosts),
            "Pitch Polishing": sep.join(f"{cfg.day_short(p.day)} {fmt(p.start)}" for p in pol[s.name]) or "-",
        }
        for d, dd in enumerate(cfg.days):
            rec[f"{dd['label']} Partners"] = sep.join(
                f"{fmt(m.start)}  {hl.get(m.host, m.host)}" for m in ms if m.day == d) or "-"
        recs.append(rec)
    return pd.DataFrame(recs)


def host_summary(cfg, hosts, meetings) -> pd.DataFrame:
    by_h = defaultdict(list)
    for m in meetings:
        by_h[m.host].append(m)
    recs = []
    for h in hosts:
        ms = sorted(by_h[h.name], key=lambda m: m.k)
        rec = {"Host": h.name, "Contact": h.contact, "Type": h.kind, "Total Meetings": len(ms)}
        for d, dd in enumerate(cfg.days):
            dm = [m for m in ms if m.day == d]
            rec[f"{dd['label']}"] = (f"{len(dm)} meetings, {fmt(dm[0].start)} to {fmt(dm[-1].end)}"
                                    if dm else "-")
        rec["Startups Met"] = ", ".join(m.startup for m in ms) or "-"
        recs.append(rec)
    return pd.DataFrame(recs)


def meetings_table(cfg, meetings, statuses, now, hosts=()) -> pd.DataFrame:
    hc = {h.name: h.contact for h in hosts}
    return pd.DataFrame([{
        "Day": cfg.days[m.day]["label"], "Date": str(cfg.day_date(m.day) or ""),
        "Start": fmt(m.start), "End": fmt(m.end), "Table": m.table,
        "Startup": m.startup, "Host": m.host, "Contact": hc.get(m.host, ""),
        "Status": eff_status(m, cfg, statuses, now),
    } for m in meetings], columns=["Day", "Date", "Start", "End", "Table", "Startup", "Host", "Contact", "Status"])


# --------------------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------------------
XF = "Arial"
C_NAVY, C_COB, C_INK, C_MUT = "152C5B", "2B65D9", "191F28", "4E5968"
C_LINE, C_ZONE, C_FAINT, C_BAND = "E1E6EF", "EEF2FF", "B8C0CF", "F7F9FC"
ST_STYLE = {"Scheduled": ("EEF0F4", "4A5168"), "Ongoing": ("2B65D9", "FFFFFF"),
            "Done": ("DDF2E3", "1E7B34"), "No-Show": ("FDE2E6", "B4132D")}
CELL_TINT = {"Ongoing": "EAF0FF", "Done": "F4F5F7", "No-Show": "FFF0F2"}
NAME_COLORS = ["2B65D9", "C2255C", "0B7A75", "B45309", "6D28D9", "0E7490", "A21CAF",
               "3F6212", "9A3412", "1D4ED8", "BE123C", "047857", "7C3AED", "155E75"]


def _f(size=10, bold=False, color=C_INK, italic=False) -> Font:
    return Font(name=XF, size=size, bold=bold, color=color, italic=italic)


def _if(size=10, bold=False, color=C_INK) -> InlineFont:
    return InlineFont(rFont=XF, sz=size, b=bold, color=Color(rgb="FF" + color))


def _fill(c: str) -> PatternFill:
    return PatternFill("solid", start_color=c, end_color=c)


def _lines(text: str, chars: float) -> int:
    return sum(max(1, -(-len(p) // max(1, int(chars)))) for p in str(text).split("\n"))


def startup_colors(startups) -> dict:
    return {s.name: NAME_COLORS[i % len(NAME_COLORS)] for i, s in enumerate(startups)}


def _table_sheet(wb, title, heading, sub, df, widths, center=(), wrap=(), bar=None):
    w = wb.create_sheet(title)
    w.sheet_view.showGridLines = False
    w.column_dimensions["A"].width = 2
    w.row_dimensions[1].height = 10
    w.cell(2, 2, heading).font = _f(16, True, C_NAVY)
    w.row_dimensions[2].height = 26
    w.cell(3, 2, sub).font = _f(10, color=C_MUT)
    hr, line = 5, Side(style="thin", color=C_LINE)
    for j, col in enumerate(df.columns):
        c = w.cell(hr, 2 + j, col)
        c.fill, c.font = _fill(C_HEAD), _f(10, True, C_HEAD_TXT)
        c.alignment = Alignment(horizontal="center" if col in center else "left", vertical="center", indent=0 if col in center else 1)
        w.column_dimensions[get_column_letter(2 + j)].width = widths.get(col, 16)
    w.row_dimensions[hr].height = 26
    for i, rec in enumerate(df.itertuples(index=False)):
        r, maxl = hr + 1 + i, 1
        for j, (col, v) in enumerate(zip(df.columns, rec)):
            c = w.cell(r, 2 + j, v)
            c.font = _f(10, bold=(j == 0), color=("B8860B" if v == "★" else C_INK))
            c.alignment = Alignment(vertical="center", horizontal="center" if col in center else "left",
                                    wrap_text=col in wrap, indent=0 if col in center else 1)
            c.border = Border(bottom=line)
            if i % 2:
                c.fill = _fill(C_BAND)
            if col in wrap and isinstance(v, str):
                maxl = max(maxl, _lines(v, widths.get(col, 16) * 1.1))
        w.row_dimensions[r].height = max(24, 13.5 * maxl + 9)
    last = hr + len(df)
    if bar and bar in df.columns and len(df):
        L = get_column_letter(2 + list(df.columns).index(bar))
        w.conditional_formatting.add(f"{L}{hr + 1}:{L}{last}", DataBarRule(
            start_type="num", start_value=0, end_type="max", color="8FA6F5", showValue=True))
    w.freeze_panes = w.cell(hr + 1, 3)
    if len(df):
        w.auto_filter.ref = f"B{hr}:{get_column_letter(1 + len(df.columns))}{last}"
    return w, hr


# Pastel by default, one hue apart. Polishing uses grey, so pink and yellow are free here
HOST_PALETTE = ["CFE5FA", "CDEFD9", "E0D7F5", "FDDCC4", "C9ECEC", "E4EFC9", "D6DCF7", "F0E2D0",
                "D3E8E1", "E9D8EF", "FBEFC0", "DDEBD3", "F9D5DC", "F6E7C9", "C7E6F2", "E3DDF0"]
# Pitch polishing is neutral slate grey (1st light, 2nd deeper) so every pastel stays free for partners and startups
POLISH_FILL = {1: "E4E8EF", 2: "C9D0DC"}
POLISH_FILL_MORE = "F1F3F6"
C_HEAD, C_HEAD_TXT, C_GRID, C_BREAK = "EEF1F6", "334155", "D9DDE5", "F6F7F9"


def text_on(hex_fill: str) -> str:
    r, g, b = (int(hex_fill[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return "FFFFFF" if 0.2126 * r + 0.7152 * g + 0.0722 * b < 0.5 else "1F2430"


def pastel_tones(n: int) -> list[str]:
    """n clearly different pastel colours: hues spread evenly round the colour wheel, and hues next to each
    other get a different lightness, so even 40 partners never share a colour."""
    import colorsys
    out = []
    step = max(1, round(n * 0.382))
    while math.gcd(step, n) != 1:
        step += 1
    for i in range(n):
        j = (i * step) % n
        lt, st_ = [(0.90, 0.75), (0.85, 0.66), (0.80, 0.56)][j % 3]
        r, g, b = colorsys.hls_to_rgb(j / n, lt, st_)
        out.append(f"{round(r * 255):02X}{round(g * 255):02X}{round(b * 255):02X}")
    return out


def host_colors(hosts) -> dict:
    """Each host (organisation) keeps one colour everywhere. Up to 16 hosts use the hand picked palette,
    bigger events get as many evenly spread pastel colours as there are hosts."""
    pal = HOST_PALETTE if len(hosts) <= len(HOST_PALETTE) else pastel_tones(len(hosts))
    return {h.name: pal[i % len(pal)] for i, h in enumerate(hosts)}


def polish_fill(n: int) -> str:
    return POLISH_FILL.get(n, POLISH_FILL_MORE)


def host_paren(h) -> str:
    """'(Raymond - Focus Tech)' like the team's sheet."""
    if isinstance(h, str):
        return f"({h})"
    return f"({h.contact} - {h.name})" if h.contact else f"({h.name})"


def day_banner(cfg: EventConfig, d: int) -> str:
    dd = cfg.day_date(d)
    if not dd:
        return f"{cfg.days[d]['label']} - Matching Schedule"
    return f"{dd:%a, %b} {dd.day}, {dd.year} - Matching Schedule"


def attendance_notes(cfg: EventConfig, startups) -> list[tuple[str, str]]:
    notes = []
    absent = [s.name for s in startups if s.absent]
    if absent:
        notes.append(("NOT HERE", ", ".join(absent)))
    if len(cfg.days) > 1:
        groups = defaultdict(list)
        for s in startups:
            if s.days and len(s.days) < len(cfg.days):
                groups[tuple(sorted(s.days))].append(s.name)
        for ds_, names in sorted(groups.items()):
            notes.append((f"ONLY {' & '.join(cfg.day_short(d).upper() for d in ds_)}", ", ".join(names)))
    return notes


PP_TITLE = "Pitching Practice Schedule"


def build_practice_sheet(wb, cfg, slots, startups, polish, statuses, now, notes, index=1):
    """Pitching practice tab, one row per polishing seat. Returns {(day, k): [startup cell refs]}
    so the main schedule can point at it with formulas."""
    days = [d for d in range(len(cfg.days)) if d in cfg.pol_days or any(p.day == d for p in polish)]
    if not days:
        return {}
    ws = wb.create_sheet(PP_TITLE, index)
    ws.sheet_view.showGridLines = False
    P = {x.name: x for x in startups}
    grid = Side(style="thin", color=C_GRID)
    box = Border(left=grid, right=grid, top=grid, bottom=grid)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for col, w in zip("ABCDEF", (2, 16, 26, 20, 14, 64)):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 10
    ws.cell(2, 2, f"{cfg.name} - Pitching Practice Schedule ({cfg.interval}-Min Rehearsals)").font = _f(16, True, C_NAVY)
    ws.row_dimensions[2].height = 28
    cap = max(1, cfg.pol_cap)
    by_k = defaultdict(list)
    for p in polish:
        by_k[p.k].append(p)
    refs, status_cells, r = {}, [], 4
    for d in days:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        c = ws.cell(r, 2, day_banner(cfg, d).replace("Matching Schedule",
                                                     f"Pitching Practice Schedule ({cfg.days[d]['label']})"))
        c.font, c.alignment = _f(12, True, C_NAVY), Alignment(vertical="center", indent=1)
        for cc in range(2, 7):
            ws.cell(r, cc).fill = _fill("E8EDF7")
        ws.row_dimensions[r].height = 24
        r += 1
        for cc, txt in enumerate((f"Time Slot ({cfg.interval}m)", "Startup", "Presenter", "Status",
                                  "Focus / Mentor Note"), 2):
            c = ws.cell(r, cc, txt)
            c.fill, c.font, c.alignment, c.border = _fill(C_HEAD), _f(10, True, C_HEAD_TXT), center, box
        ws.row_dimensions[r].height = 24
        r += 1
        for row in day_rows(d, cfg, slots, startups, [], []):
            if row["lunch"]:
                vals = (row["time"].replace("–", " - "), "LUNCH BREAK", "-", "-", "Lunch Break")
                for cc, v in enumerate(vals, 2):
                    c = ws.cell(r, cc, v)
                    c.fill, c.border = _fill(C_BREAK), box
                    c.font = _f(10, False, "6B7280", italic=True)
                    c.alignment = left if cc in (3, 6) else center
                r += 1
                continue
            if row.get("tables") is None:
                continue
            k, first = row["slot"].k, r
            seats = sorted(by_k[k], key=lambda p: p.startup)
            for i in range(cap):
                p = seats[i] if i < len(seats) else None
                ws.cell(r, 2, f"{fmt(row['start'])} - {fmt(row['end'])}" if i == 0 else None)
                ws.cell(r, 3, p.startup if p else None).font = _f(10.5, True, C_INK)
                ws.cell(r, 4, (P[p.startup].presenter if p and p.startup in P else None) or None).font = _f(10)
                sc = ws.cell(r, 5, pol_status(p, cfg, statuses, now) if p else None)
                sc.font = _f(10)
                ws.cell(r, 6, notes.get(p.pid) if p else None).font = _f(10)
                for cc in range(2, 7):
                    ws.cell(r, cc).border = box
                    ws.cell(r, cc).alignment = left if cc in (3, 6) else center
                ws.cell(r, 2).font = _f(10.5, True)
                status_cells.append(sc.coordinate)
                refs.setdefault((d, k), []).append(f"C{r}")
                ws.row_dimensions[r].height = max(22, 13 * _lines(notes.get(p.pid, "") if p else "", 64 * 1.1) + 9)
                r += 1
            if cap > 1:
                ws.merge_cells(start_row=first, start_column=2, end_row=r - 1, end_column=2)
        r += 1
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    for coord in status_cells:
        dv.add(coord)
    ws.add_data_validation(dv)
    rng = f"E4:E{r}"
    for s_, (bg, fg) in {**ST_STYLE, "Scheduled": ("EBF1FD", "2B65D9")}.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{s_}"'],
                                                      fill=_fill(bg), font=Font(color=fg, bold=True)))
    ws.freeze_panes = "C1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return refs


def build_excel(cfg, slots, startups, hosts, meetings, statuses, now, polish=(), pol_notes=None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Matching Schedule"
    ws.sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True
    pp_refs = build_practice_sheet(wb, cfg, slots, startups, polish, statuses, now, pol_notes or {})
    pol_empty = []
    T = cfg.tables
    H = {h.name: h for h in hosts}
    hcol = host_colors(hosts)
    L = get_column_letter
    SE, TM = 2, 3
    MC, SC = (lambda i: 4 + 2 * i), (lambda i: 5 + 2 * i)
    PC, WC = 4 + 2 * T, 5 + 2 * T
    grid = Side(style="thin", color=C_GRID)
    box = Border(left=grid, right=grid, top=grid, bottom=grid)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    W_MEET, W_STAT, W_POL, W_WAIT = 27, 12, 22, 48

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions[L(SE)].width = 15
    ws.column_dimensions[L(TM)].width = 15
    for i in range(T):
        ws.column_dimensions[L(MC(i))].width = W_MEET
        ws.column_dimensions[L(SC(i))].width = W_STAT
    ws.column_dimensions[L(PC)].width = W_POL
    ws.column_dimensions[L(WC)].width = W_WAIT

    ws.row_dimensions[1].height = 10
    ws.cell(2, SE, f"{cfg.name} ({cfg.interval}-Min Slots)").font = _f(18, True, C_NAVY)
    ws.row_dimensions[2].height = 30
    notes = attendance_notes(cfg, startups)
    r = 4
    if notes:
        # Hidden behind the row-group "+" button in the left margin (Excel and Google Sheets)
        c = ws.cell(3, SE, f"Attendance notes ({len(notes)}): click + on the left to show")
        c.font = _f(9, False, C_MUT, italic=True)
        for j, (head, names) in enumerate(notes):
            c = ws.cell(4 + j, SE)
            c.value = CellRichText(TextBlock(_if(10.5, True, "B4132D" if head == "NOT HERE" else C_NAVY), head + ": "),
                                   TextBlock(_if(10.5, False, C_INK), names))
            c.alignment = Alignment(vertical="center")
        ws.row_dimensions.group(4, 3 + len(notes), outline_level=1, hidden=True)
        ws.sheet_properties.outlinePr.summaryBelow = False
        r = 5 + len(notes)

    status_cells, meet_ranges, first = [], [], r
    for d in range(len(cfg.days)):
        ws.merge_cells(start_row=r, start_column=SE, end_row=r, end_column=WC)
        c = ws.cell(r, SE, day_banner(cfg, d))
        c.font, c.alignment = _f(12, True, C_NAVY), Alignment(vertical="center", indent=1)
        for cc in range(SE, WC + 1):
            ws.cell(r, cc).fill = _fill("E8EDF7")
        ws.row_dimensions[r].height = 24
        r += 1
        heads = ([(SE, cfg.nm("session")), (TM, f"{cfg.nm('time')} ({cfg.interval:g}m)")]
                 + [x for i in range(T) for x in ((MC(i), f"{cfg.nm('table')} {i + 1}"),
                                                  (SC(i), f"{cfg.nm('status')} {i + 1}"))]
                 + [(PC, cfg.nm("polish")), (WC, cfg.nm("waiting"))])
        for c_, txt in heads:
            c = ws.cell(r, c_, txt)
            c.fill, c.font, c.alignment, c.border = _fill(C_HEAD), _f(10, True, C_HEAD_TXT), center, box
        ws.row_dimensions[r].height = 26
        r += 1

        runs = []   # (session index, first row, last row)
        sessions = cfg.session_list()
        for row in day_rows(d, cfg, slots, startups, hosts, meetings, polish=polish):
            if row["lunch"] or row.get("divider"):
                label = "LUNCH BREAK" if row["lunch"] else str(row["label"]).upper()
                vals = {SE: "-", TM: row["time"].replace("–", " - ")}
                for i in range(T):
                    vals[MC(i)], vals[SC(i)] = label, "-"
                vals[PC], vals[WC] = label, label
                for cc in range(SE, WC + 1):
                    c = ws.cell(r, cc, vals.get(cc, ""))
                    c.fill, c.border, c.alignment = _fill(C_BREAK), box, center
                    c.font = _f(10, cc != SE and vals.get(cc) not in ("-", ""), "6B7280")
                ws.cell(r, TM).font = _f(10, True, C_INK)
                ws.row_dimensions[r].height = 24
                r += 1
                continue
            if runs and runs[-1][0] == row["session"] and runs[-1][2] == r - 1:
                runs[-1][2] = r
            else:
                runs.append([row["session"], r, r])
            c = ws.cell(r, TM, f"{fmt(row['start'])} - {fmt(row['end'])}")
            c.font, c.alignment, c.border = _f(10.5, True), center, box
            need = 2
            for i, m in enumerate(row["tables"]):
                mc, sc = ws.cell(r, MC(i)), ws.cell(r, SC(i))
                if m:
                    h = H.get(m.host, m.host)
                    fill = hcol.get(m.host, "FFFFFF")
                    mc.value = f"{m.startup}\n{host_paren(h)}"
                    mc.fill, mc.font = _fill(fill), _f(10, False, text_on(fill))
                    sc.value = eff_status(m, cfg, statuses, now)
                    sc.font = _f(10, False, "374151")
                    status_cells.append(sc.coordinate)
                    need = max(need, _lines(m.startup, W_MEET * 1.1) + _lines(host_paren(h), W_MEET * 1.15))
                else:
                    mc.value, sc.value = "-", "-"
                    mc.font = sc.font = _f(10, color="9CA3AF")
                mc.alignment, sc.alignment = center, center
                mc.border = sc.border = box
            pc = ws.cell(r, PC)
            ref = pp_refs.get((d, row["slot"].k))
            if ref:
                # linked to the Pitching Practice sheet: edit the startup there and this cell follows
                q = f"'{PP_TITLE}'!"
                pc.value = (f'=IF({q}{ref[0]}="","",{q}{ref[0]})' if len(ref) == 1
                            else f"=_xlfn.TEXTJOIN(CHAR(10),TRUE,{q}{ref[0]}:{ref[-1]})")
                if not row["polish"]:
                    pol_empty.append(pc.coordinate)
            elif row["polish"]:
                pc.value = "\n".join(p.startup for p in row["polish"])
            pc.font = _f(10.5, True, C_INK)
            if row["polish"]:
                pc.fill = _fill(polish_fill(max(p.n for p in row["polish"])))
                need = max(need, len(row["polish"]))
            pc.alignment, pc.border = center, box
            wc = ws.cell(r, WC, " • ".join(row["waiting"]) if row["waiting"] else "None (All in Sessions)")
            wc.font = _f(9.5, False, "374151", italic=not row["waiting"])
            wc.alignment, wc.border = Alignment(horizontal="center", vertical="center", wrap_text=True), box
            need = max(need, _lines(wc.value, W_WAIT * 1.25))
            ws.row_dimensions[r].height = max(34, 13 * need + 8)
            r += 1
        for si, a, b in runs:
            if b > a:
                ws.merge_cells(start_row=a, start_column=SE, end_row=b, end_column=SE)
            sname = sessions[si] if 0 <= si < len(sessions) else None
            c = ws.cell(a, SE, f"{sname['name']}\n({fmt(sname['start'])} - {fmt(sname['end'])})" if sname else "-")
            c.font, c.alignment = _f(10.5, True), center
            for rr in range(a, b + 1):
                ws.cell(rr, SE).border = box
        for i in range(T):
            meet_ranges.append((L(MC(i)), L(SC(i))))
        r += 1
    last = r

    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True,
                        showErrorMessage=True, errorTitle="Status", error="Pick a status from the list")
    for coord in status_cells:
        dv.add(coord)
    ws.add_data_validation(dv)
    srng = " ".join(f"{L(SC(i))}{first}:{L(SC(i))}{last}" for i in range(T))
    for s_ in ("Ongoing", "Done", "No-Show"):
        bg, fg = ST_STYLE[s_]
        ws.conditional_formatting.add(srng, CellIsRule(operator="equal", formula=[f'"{s_}"'],
                                                       fill=_fill(bg), font=Font(color=fg, bold=True)))
    for coord in pol_empty:
        ws.conditional_formatting.add(coord, FormulaRule(formula=[f"LEN({coord})>0"],
                                                         fill=_fill(POLISH_FILL[1])))
    for i in range(T):
        rng = f"{L(MC(i))}{first}:{L(MC(i))}{last}"
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'${L(SC(i))}{first}="No-Show"'], fill=_fill("EDEEF1"),
            font=Font(color="9CA3AF", strike=True)))
    ws.freeze_panes = ws.cell(1, MC(0))
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    sdf = startup_summary(cfg, startups, hosts, meetings, statuses, now, sep="\n", polish=polish)
    _table_sheet(wb, "Startup Summary", "Startup summary",
                 f"{len(meetings)} meetings across {sum(not s.absent for s in startups)} attending startups", sdf,
                 {"Startup Name": 24, "Presenter": 16, "Attendance": 16, "Domain": 18, "Priority": 9, "Total Meetings": 15,
                  "No-Shows": 11, "Eligible Hosts": 14, "Pitch Polishing": 18,
                  **{c_: 40 for c_ in sdf.columns if c_.endswith("Partners")}},
                 center=("Priority", "Total Meetings", "No-Shows", "Eligible Hosts"),
                 wrap=tuple(c_ for c_ in sdf.columns if c_.endswith("Partners")) + ("Pitch Polishing",),
                 bar="Total Meetings")
    hdf = host_summary(cfg, hosts, meetings)
    w, hr = _table_sheet(wb, "Host Summary", "Host summary", f"{len(hosts)} hosts", hdf,
                         {"Host": 26, "Contact": 18, "Type": 12, "Total Meetings": 15, "Startups Met": 60,
                          **{d_["label"]: 26 for d_ in cfg.days}},
                         center=("Type", "Total Meetings"), wrap=("Startups Met",), bar="Total Meetings")
    for i, h in enumerate(hosts):
        c = w.cell(hr + 1 + i, 2)
        c.fill, c.font = _fill(hcol[h.name]), _f(10, True, text_on(hcol[h.name]))
    mdf = meetings_table(cfg, meetings, statuses, now, hosts)
    w, hr = _table_sheet(wb, "All Meetings", "All meetings", "One row per meeting. Status can be changed here too",
                         mdf, {"Day": 9, "Date": 12, "Start": 9, "End": 9, "Table": 8, "Startup": 24,
                               "Host": 26, "Contact": 18, "Status": 13},
                         center=("Day", "Date", "Start", "End", "Table", "Status"))
    if len(mdf):
        Ls = get_column_letter(2 + list(mdf.columns).index("Status"))
        rng = f"{Ls}{hr + 1}:{Ls}{hr + len(mdf)}"
        dv2 = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
        dv2.add(rng)
        w.add_data_validation(dv2)
        for s_, (bg, fg) in ST_STYLE.items():
            w.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{s_}"'],
                                                         fill=_fill(bg), font=Font(color=fg, bold=True)))
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()



# --------------------------------------------------------------------------------------
# Sample data and file import
# --------------------------------------------------------------------------------------
def sample_frames() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    d0 = date.today() + timedelta(days=7)
    days = pd.DataFrame({"Label": ["Day 1", "Day 2"], "Date": [d0, d0 + timedelta(days=1)]})
    s = [
        ("Voltra Energy", "Climate, B2B", "", "", "", "", True),
        ("MediKey", "HealthTech, AI, B2B", "", "", "", "", False),
        ("Kanto Robotics", "Hardware, AI", "", "", "", "", False),
        ("Seedline Agritech", "AgriTech", "", "", "", "", False),
        ("Parcelo", "Logistics, B2B", "", "", "", "", False),
        ("FinchPay", "Fintech, B2C", "", "", "", "", True),
        ("Lumen Health", "HealthTech, B2C", "", "", "", "", False),
        ("Tidewave Marine", "Climate", "", "", "", "", False),
        ("Glowlab Cosmetics", "K-Beauty, B2C", "", "", "", "", False),
        ("Aster Skin", "K-Beauty, B2C", "Day 2", "", "", "", False),
        ("Railmind", "Mobility, AI", "", "", "16:00", "", False),
        ("Quanta Ledger", "Fintech, AI, B2B", "", "", "", "", False),
        ("Brightcart", "SaaS, B2C", "", "", "", "", False),
        ("Nimbus Grid", "Climate, Hardware", "", "", "", "", False),
        ("Coral AI", "SaaS, AI/ML, B2B", "", "", "", "", True),
        ("Haneul Foods", "FoodTech, B2C", "", "", "", "Day 1 13:00-14:00", False),
        ("Pathwise", "EdTech", "Not here", "", "", "", False),
        ("Ondo Mobility", "Mobility", "Day 1", "", "", "", False),
        ("Silo Sense", "AgriTech, Hardware", "", "11:00", "", "", False),
        ("Kite Insure", "Fintech, B2C", "", "", "", "", False),
    ]
    startups = pd.DataFrame(s, columns=["Name", "Domain", "Days", "Available From", "Available Until",
                                        "Blocked Times", "Priority"])
    startups["Presenter"] = ["Aiden", "Hana", "Kenji", "Ravi", "Mia", "Leon", "Sora", "Noah", "Yuna", "Clara",
                             "Arun", "Jiwoo", "Ethan", "Min & Ivy", "Zoe", "Daniel", "Lily", "Ken", "Mateo", "Ava"]
    startups["Pitch Polishing"] = True
    startups["Max Meetings"] = None
    h = [
        ("Sunrise Ventures", "Grace Lim", "Investor", "", "", "", "", "", "B2C", "K-Beauty", None),
        ("Meridian Capital", "Daniel Ong", "Investor", "Day 2", "15:00", "17:00", "", "AI, Fintech", "", "", None),
        ("Halcyon Logistics", "Priya Nair", "Corporate", "", "", "", "", "", "Logistics, B2B", "K-Beauty, EdTech", None),
        ("Orchid Bank", "Marcus Tan", "Corporate", "Day 1", "", "", "", "", "Fintech", "K-Beauty, Hardware", None),
        ("Harbor Innovation Agency", "Mei Ling", "Agency", "", "", "", "Day 2 10:00-11:00", "", "", "", None),
        ("Lumen Corporate Ventures", "Jason Koh", "Investor", "", "", "", "", "", "AI", "", 16),
    ]
    hosts = pd.DataFrame(h, columns=HOST_COLS)
    return days, clean_frame(startups, "startup"), clean_frame(hosts, "host")


def blank_frames():
    days = pd.DataFrame({"Label": ["Day 1"], "Date": [date.today() + timedelta(days=7)]})
    return (days, clean_frame(pd.DataFrame(columns=STARTUP_COLS), "startup"),
            clean_frame(pd.DataFrame(columns=HOST_COLS), "host"))


ALIASES = {
    "name": "Name", "startup": "Name", "startup name": "Name", "company": "Name", "host": "Name",
    "host name": "Name", "investor": "Name", "partner": "Name",
    "domain": "Domain", "domains": "Domain", "sector": "Domain", "industry": "Domain",
    "days": "Days", "day": "Days", "attending days": "Days", "available days": "Days",
    "available from": "Available From", "from": "Available From", "start": "Available From",
    "available until": "Available Until", "until": "Available Until", "to": "Available Until",
    "cutoff": "Available Until", "end": "Available Until",
    "blocked times": "Blocked Times", "blocked": "Blocked Times", "unavailable": "Blocked Times",
    "restricted times": "Blocked Times",
    "priority": "Priority", "max meetings": "Max Meetings", "max": "Max Meetings",
    "type": "Type", "host type": "Type",
    "presenter": "Presenter", "speaker": "Presenter", "founder": "Presenter", "pitcher": "Presenter",
    "contact": "Contact", "contact person": "Contact", "person": "Contact", "representative": "Contact",
    "attendee": "Contact",
    "excluded domains": "Excluded Domains", "excludes": "Excluded Domains", "exclude": "Excluded Domains",
    "only domains": "Only Domains", "only meets": "Only Domains", "only": "Only Domains", "must be": "Only Domains",
    "preferred domains": "Preferred Domains", "prefers": "Preferred Domains", "preferred": "Preferred Domains",
    "interests": "Preferred Domains", "interest": "Preferred Domains", "focus areas": "Preferred Domains",
    "희망 분야": "Preferred Domains", "선호 분야": "Preferred Domains", "관심 분야": "Preferred Domains",
    "전용 분야": "Only Domains", "만나는 분야": "Only Domains", "필수 분야": "Only Domains",
    "스타트업": "Name", "스타트업명": "Name", "회사": "Name", "회사명": "Name", "기업": "Name", "기업명": "Name",
    "이름": "Name", "소속": "Name", "투자사": "Name", "발표자": "Presenter", "분야": "Domain", "섹터": "Domain",
    "산업": "Domain", "참석일": "Days", "참석": "Days", "일자": "Days", "날짜": "Days", "담당자": "Contact",
    "유형": "Type", "구분": "Type", "제외 분야": "Excluded Domains", "제외": "Excluded Domains",
    "우선": "Priority", "우선순위": "Priority", "불가 시간": "Blocked Times", "가능 시작": "Available From",
    "가능 종료": "Available Until", "최대 미팅": "Max Meetings", "피치 폴리싱": "Pitch Polishing",
}


def host_type(v) -> str:
    if is_blank(v):
        return "Investor"
    t = norm(v)
    for k in HOST_TYPES:
        if t == k.lower():
            return k
    if t in ("vc", "cvc", "fund", "angel", "investors", "family office", "pe"):
        return "Investor"
    if t in ("corporate partner", "corporates", "company", "enterprise", "corp"):
        return "Corporate"
    if t in ("government", "gov", "agency partner", "public"):
        return "Agency"
    return "Other"


def clean_frame(df: pd.DataFrame, role: str) -> pd.DataFrame:
    cols = STARTUP_COLS if role == "startup" else HOST_COLS
    df = df.rename(columns={c: ALIASES.get(norm(c), c) for c in df.columns})
    df = df.loc[:, ~df.columns.duplicated()]
    out = pd.DataFrame(index=range(len(df)))
    for c in cols:
        src = df[c] if c in df.columns else pd.Series([None] * len(df))
        src = src.reset_index(drop=True)
        if c == "Priority":
            out[c] = src.map(to_bool).astype(bool)
        elif c == "Max Meetings":
            out[c] = pd.to_numeric(src, errors="coerce").astype("float64")
        elif c in ("Available From", "Available Until"):
            out[c] = src.map(lambda v: "" if is_blank(v) else tstr(v)).astype(object)
        elif c == "Type":
            out[c] = src.map(host_type).astype(object)
        elif c == "Days":
            out[c] = src.map(lambda v: "Not here" if isinstance(v, str) and norm(v) in ABSENT_WORDS
                             else ("" if is_blank(v) else str(v).strip())).astype(object)
        elif c == "Pitch Polishing":
            out[c] = src.map(lambda v: True if is_blank(v) else to_bool(v)).astype(bool)
        else:
            out[c] = src.map(lambda v: "" if is_blank(v) else str(v).strip()).astype(object)
    return out


def read_upload(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file, dtype=object, keep_default_na=False, na_values=["", "#N/A", "nan", "NaN"])
    raw = file.getvalue()
    for enc in ("utf-8-sig", "cp949", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), dtype=object, encoding=enc, sep=None, engine="python",
                               keep_default_na=False, na_values=["", "#N/A", "nan", "NaN"])
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    raise ValueError("Could not read the file. Save it as UTF-8 CSV or .xlsx")


# ---- workbooks with several tabs ----
SHEET_HINTS = {
    "startup": ("startup", "스타트업", "참가", "참여", "portfolio", "포트폴리오", "company list", "기업 목록", "기업목록"),
    "partner": ("partner", "파트너", "investor", "투자", "vc", "corporate", "corp", "buyer", "바이어", "기관", "host",
                "호스트", "lp"),
}


def _known_header(c: str) -> bool:
    n = norm(c)
    if not n:
        return False
    if n in ALIASES:
        return True
    return any(w in n for k in DIR_ALIASES for _, words in DIR_ALIASES[k] for w in words if len(w) > 2)


def detect_header(raw: pd.DataFrame) -> pd.DataFrame:
    """Sheets often start with a title or notes. Use the row that looks most like column names as the header."""
    raw = raw.map(lambda v: "" if is_blank(v) else v)
    best_i, best = 0, -1.0
    for i in range(min(15, len(raw))):
        cells = [str(v).strip() for v in raw.iloc[i] if str(v).strip()]
        if len(cells) < 2:
            continue
        texty = sum(1 for c in cells if not re.fullmatch(r"[\d.,:/\- ]+", c))
        score = texty + 3 * sum(_known_header(c) for c in cells)
        if score > best:
            best, best_i = score, i
    hdr, seen, cols = [str(v).strip() for v in raw.iloc[best_i]], {}, []
    for j, h in enumerate(hdr):
        h = h or f"Column {j + 1}"
        seen[h] = seen.get(h, 0) + 1
        cols.append(h if seen[h] == 1 else f"{h} ({seen[h]})")
    df = raw.iloc[best_i + 1:].copy()
    df.columns = cols
    df = df.loc[:, [not c.startswith("Column ") or any(str(v).strip() for v in df[c]) for c in df.columns]]
    df = df[[any(str(v).strip() for v in r) for r in df.itertuples(index=False)]].reset_index(drop=True)
    return df


def read_book(file) -> dict:
    """Every tab of a workbook (or the one table in a CSV) as {name: DataFrame}, empty tabs left out."""
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        sheets = pd.read_excel(file, sheet_name=None, header=None, dtype=object, keep_default_na=False,
                               na_values=["", "#N/A", "nan", "NaN"])
        out = {}
        for sn, raw in sheets.items():
            if raw.dropna(how="all").empty:
                continue
            df = detect_header(raw)
            if len(df.columns) >= 2 and len(df):      # skip cover / notes tabs
                out[str(sn)] = df
        if not out:
            raise ValueError("The workbook has no rows")
        return out
    return {"CSV": read_upload(file)}


def sheet_score(name: str, df: pd.DataFrame, target: str) -> float:
    kind = "partner" if target in ("host", "partner") else "startup"
    other = "startup" if kind == "partner" else "partner"
    n = norm(name)
    sc = 6.0 if any(w in n for w in SHEET_HINTS[kind]) else 0.0
    sc -= 4.0 if any(w in n for w in SHEET_HINTS[other]) else 0.0
    taken = set()
    for c in df.columns:
        f = guess_dir(c, kind, taken)
        if f not in ("(new)", "(ignore)"):
            taken.add(f)
    return sc + len(taken) + min(len(df), 500) / 1000


def book_text(book: dict, names: list) -> str:
    names = [n for n in names if n in book]
    if not names:
        return ""
    if len(names) == 1:
        return book[names[0]].to_csv(sep="\t", index=False)
    parts = []
    for n in names:
        d = book[n].copy()
        d.insert(0, "Sheet", n)
        parts.append(d)
    return pd.concat(parts, ignore_index=True).fillna("").to_csv(sep="\t", index=False)


def load_book(tag: str, book: dict, target: str, paste_key: str) -> str:
    """Remember the workbook, pick the tab that fits best and put it in the paste box. Returns the tab name."""
    ss()[f"book_{tag}"] = book
    best = max(book, key=lambda n: sheet_score(n, book[n], target))
    ss()[f"sheets_{tag}"] = [best]
    ss()[paste_key] = book_text(book, [best])
    return best


def cb_sheets(tag: str, paste_key: str):
    ss()[paste_key] = book_text(ss().get(f"book_{tag}", {}), ss().get(f"sheets_{tag}") or [])


def sheet_picker(tag: str, paste_key: str):
    """When a workbook has several tabs: choose one or more (same columns are combined)."""
    book = ss().get(f"book_{tag}")
    if not book or len(book) < 2:
        return
    st.multiselect(T("sheets_pick", n=len(book)), list(book), key=f"sheets_{tag}", on_change=cb_sheets,
                   args=(tag, paste_key), format_func=lambda n: T("sheet_rows", s=n, n=len(book[n])),
                   help=T("sheets_help"))


SPLIT_HOST = "Contact - Organisation (split)"
GUESS = [  # (words in the header, field)
    (("present", "speaker", "pitcher", "founder"), "Presenter"),
    (("contact", "attendee", "representative", "person"), "Contact"),
    (("exclud", "avoid", "not interested"), "Excluded Domains"),
    (("only", "must", "required"), "Only Domains"),
    (("prefer", "interest", "focus area", "thesis", "wish"), "Preferred Domains"),
    (("type", "kind", "category"), "Type"),            # hosts only (startups have no Type)
    (("domain", "sector", "industry", "category", "vertical"), "Domain"),
    (("block", "unavail", "busy", "restricted"), "Blocked Times"),
    (("polish", "practice", "rehears"), "Pitch Polishing"),
    (("priority", "vip", "focus startup"), "Priority"),
    (("max",), "Max Meetings"),
    (("from", "arriv", "start time"), "Available From"),
    (("until", "leave", "cutoff", "cut-off", "end time"), "Available Until"),
    (("day", "date", "attend"), "Days"),
    (("startup", "company", "name", "organisation", "organization", "investor", "partner", "host", "firm"), "Name"),
]


def parse_paste(raw: str):
    """Cells copied from Google Sheets / Excel (tab separated, first row = headers) -> DataFrame."""
    raw = (raw or "").strip("\n")
    if not raw.strip():
        return None
    sep = "\t" if "\t" in raw.splitlines()[0] else ","
    try:
        df = pd.read_csv(io.StringIO(raw), sep=sep, dtype=object, keep_default_na=False, engine="python")
    except Exception:  # noqa: BLE001
        return None
    df.columns = [str(c).strip() or f"Column {i + 1}" for i, c in enumerate(df.columns)]
    df = df.loc[:, [any(not is_blank(v) for v in df[c]) or not c.startswith("Unnamed") for c in df.columns]]
    df = df[[any(not is_blank(v) for v in r) for r in df.itertuples(index=False)]].reset_index(drop=True)
    return df if len(df.columns) else None


def guess_field(header: str, role: str, taken: set) -> str:
    cols = STARTUP_COLS if role == "startup" else HOST_COLS
    n = norm(header)
    own = next((c for c in cols if n in {norm(col_label(role, c)), norm(default_label(role, c, "en")),
                                         norm(default_label(role, c, "ko"))}), None)
    if own and own not in taken:
        return own
    cand = ALIASES.get(n) if ALIASES.get(n) in cols else next((c for c in cols if c.lower() == n), None)
    if not cand:
        for words, f in GUESS:
            if f in cols and any(w in n for w in words):
                cand = f
                break
    return cand if cand and cand not in taken else "(ignore)"


def default_label(role: str, col: str, lg: str | None = None) -> str:
    L = COL_LABELS[lg or lang()]
    return L.get("_host_name", "Organisation") if role == "host" and col == "Name" else L.get(col, col)


def col_label(role: str, col: str) -> str:
    """Heading shown in the participant tables. Renaming only changes the display, never the data."""
    return (st.session_state.get(f"lbl_{role}_{col}") or "").strip() or default_label(role, col)


def labelled(role: str, config: dict) -> dict:
    for col, cc in config.items():
        cc["label"] = col_label(role, col)
    return config


def rename_popover(role: str, where):
    cols = STARTUP_COLS if role == "startup" else HOST_COLS
    with where.popover(T("rename")):
        st.caption(T("rename_cap"))
        for c in cols:
            st.text_input(default_label(role, c), key=f"lbl_{role}_{c}", placeholder=default_label(role, c))


def map_key(role: str, i: int, header: str) -> str:
    return f"map_{role}_{i}_{hashlib.md5(header.encode()).hexdigest()[:6]}"


# --------------------------------------------------------------------------------------
# Streamlit state helpers
# --------------------------------------------------------------------------------------
def ss():
    return st.session_state


def set_frames(days, startups, hosts, cfg: dict | None = None):
    s = ss()
    s.days_base, s.s_base, s.h_base = days, startups, hosts
    s.days_df, s.s_df, s.h_df = days, startups, hosts
    s.ver = s.get("ver", 0) + 1
    c = cfg or {}
    s.ev_name = c.get("name", s.get("ev_name", "Smart Cities Selection Days"))
    s.t_start = time.fromisoformat(c.get("start", "10:00"))
    s.t_end = time.fromisoformat(c.get("end", "17:00"))
    s.interval = int(c.get("interval", 20))
    buf = float(c.get("buffer", 0) or 0)
    s.buffer_m, s.buffer_s = int(buf), int(round((buf - int(buf)) * 60))
    labels_ = c.get("col_labels", {}) or {}
    for role_, cols_ in (("startup", STARTUP_COLS), ("host", HOST_COLS)):
        for c_ in cols_:
            s[f"lbl_{role_}_{c_}"] = (labels_.get(role_, {}) or {}).get(c_) or ""
    names = c.get("names", {}) or {}
    for k_ in DEFAULT_NAMES:
        s[f"nm_{k_}"] = names.get(k_) or DEFAULT_NAMES[k_]
    s.lunch_on = bool(c.get("lunch_on", True))
    s.t_lunch_s = time.fromisoformat(c.get("lunch_start", "12:00"))
    s.t_lunch_e = time.fromisoformat(c.get("lunch_end", "13:00"))
    s.tables = int(c.get("tables", 6))
    s.tz = c.get("tz", "Asia/Singapore")
    s.sessions_txt = c.get("sessions", "Session A 10:00-12:00; Session B 13:00-15:00; Session C 15:00-17:00")
    s.pol_on = bool(c.get("pol_on", True))
    s.pol_days = list(c.get("pol_days", list(days["Label"]) if len(days) else []))
    s.t_pol_s = time.fromisoformat(c.get("pol_start", "10:00"))
    s.t_pol_e = time.fromisoformat(c.get("pol_end", "17:00"))
    s.pol_cap = int(c.get("pol_cap", 1))


def init_state():
    s = ss()
    if "init" in s:
        return
    s.init = True
    s.result = None
    s.statuses = {}
    s.pol_notes = {}
    s.auto_refresh = True
    set_frames(*sample_frames())


def current_cfg() -> EventConfig:
    s = ss()
    days = []
    for _, r in s.days_df.reset_index(drop=True).iterrows():
        lab = "" if is_blank(r.get("Label")) else str(r.get("Label")).strip()
        if not lab:
            lab = f"Day {len(days) + 1}"
        dv = r.get("Date")
        days.append({"label": lab, "date": "" if is_blank(dv) else str(pd.Timestamp(dv).date())})
    if not days:
        days = [{"label": "Day 1", "date": ""}]
    tm = lambda t: t.hour * 60 + t.minute
    try:
        sessions, s.sess_err = parse_sessions(s.get("sessions_txt", "")), None
    except ValueError as e:
        sessions, s.sess_err = [], str(e)
    labels = [d["label"] for d in days]
    pol_days = [labels.index(x) for x in s.get("pol_days", []) if x in labels] if s.get("pol_on") else []
    return EventConfig(
        name=s.ev_name or "Business Matching", days=days, start=tm(s.t_start), end=tm(s.t_end),
        interval=int(s.interval), buffer=int(s.buffer_m) + int(s.buffer_s) / 60,
        lunch_start=tm(s.t_lunch_s) if s.lunch_on else None,
        lunch_end=tm(s.t_lunch_e) if s.lunch_on else None,
        tables=int(s.tables), tz=s.tz, sessions=sessions, pol_days=pol_days,
        pol_start=tm(s.t_pol_s), pol_end=tm(s.t_pol_e), pol_cap=int(s.pol_cap),
        names={k_: (s.get(f"nm_{k_}") or "").strip() for k_ in DEFAULT_NAMES
               if (s.get(f"nm_{k_}") or "").strip() and s.get(f"nm_{k_}").strip() != DEFAULT_NAMES[k_]})


def frames_hash(cfg: EventConfig, sdf: pd.DataFrame, hdf: pd.DataFrame) -> str:
    blob = json.dumps([asdict(cfg), sdf.to_dict("records"), hdf.to_dict("records")], default=str)
    return hashlib.sha1(blob.encode()).hexdigest()


def project_json(live: bool = True) -> str:
    d = project_dict(live)
    if "dir_df" in ss():
        d["directory"] = {k: json.loads(dir_doc(k)) for k in DIR_FIELDS}
    if ss().get("local_docs"):
        d["program_docs"] = ss().local_docs
    return json.dumps(d, default=str, indent=1)


def project_dict(live: bool = True) -> dict:
    """Everything needed to restore the event. live=False leaves out statuses and notes, which the team
    workspace stores per item."""
    s = ss()
    cfg = current_cfg()
    d = {
        "app": APP_TITLE, "version": 1,
        "config": {"name": cfg.name, "start": fmt(cfg.start), "end": fmt(cfg.end),
                   "interval": cfg.interval, "buffer": cfg.buffer, "lunch_on": bool(s.lunch_on),
                   "lunch_start": s.t_lunch_s.strftime("%H:%M"), "lunch_end": s.t_lunch_e.strftime("%H:%M"),
                   "tables": cfg.tables, "tz": cfg.tz, "sessions": s.sessions_txt,
                   "pol_on": bool(s.pol_on), "pol_days": list(s.pol_days),
                   "pol_start": s.t_pol_s.strftime("%H:%M"), "pol_end": s.t_pol_e.strftime("%H:%M"),
                   "pol_cap": int(s.pol_cap), "names": cfg.names,
                   "col_labels": {role_: {c_: s.get(f"lbl_{role_}_{c_}").strip() for c_ in cols_
                                          if (s.get(f"lbl_{role_}_{c_}") or "").strip()}
                                  for role_, cols_ in (("startup", STARTUP_COLS), ("host", HOST_COLS))}},
        "days": cfg.days,
        "startups": s.s_df.to_dict("records"),
        "hosts": s.h_df.to_dict("records"),
        "result": s.result and {k: v for k, v in s.result.items()
                                if k in ("meetings", "polish", "parked", "packing", "status", "hash", "cfg")},
    }
    if live:
        d["statuses"] = s.statuses
        d["pol_notes"] = s.get("pol_notes", {})
    return d


# Callbacks (run before the script body, so they may reset widget state)
def cb_sample():
    set_frames(*sample_frames())
    ss().result, ss().statuses, ss().pol_notes = None, {}, {}


def cb_blank():
    set_frames(*blank_frames())
    ss().ev_name = "New Matching Event"
    ss().result, ss().statuses, ss().pol_notes = None, {}, {}


def cb_upload(role: str):
    """A file goes through the same column check as pasted cells."""
    f = ss().get(f"up_{role}")
    if not f:
        return
    try:
        book = read_book(f)
    except Exception as e:  # noqa: BLE001
        ss().flash = ("error", T("f_read_fail", f=f.name, e=e))
        return
    if len(book) == 1:
        ss().pop(f"book_{role}", None)
        ss()[f"paste_{role}"] = book_text(book, list(book))
        ss().flash = ("info", T("f_loaded_file", f=f.name))
    else:
        best = load_book(role, book, role, f"paste_{role}")
        ss().flash = ("info", T("f_loaded_tab", f=f.name, n=len(book), s=best))


def cb_paste(role: str):
    s = ss()
    df = parse_paste(s.get(f"paste_{role}", ""))
    if df is None:
        s.flash = ("error", T("f_nothing"))
        return
    out = pd.DataFrame(index=range(len(df)))
    for i, c in enumerate(df.columns):
        f = s.get(map_key(role, i, c), "(ignore)")
        col = df[c].reset_index(drop=True)
        if f == SPLIT_HOST:
            parts = col.map(lambda v: [x.strip() for x in str(v).split(" - ", 1)] if " - " in str(v)
                            else ["", str(v).strip()])
            out["Contact"], out["Name"] = parts.map(lambda p: p[0]), parts.map(lambda p: p[1])
        elif f != "(ignore)" and f not in out.columns:
            out[f] = col
    if "Name" not in out.columns:
        s.flash = ("error", T("f_need_name"))
        return
    new = clean_frame(out, role)
    new = new[[not is_blank(v) for v in new["Name"]]].reset_index(drop=True)
    key_base, key_df = ("s_base", "s_df") if role == "startup" else ("h_base", "h_df")
    if s.get(f"pmode_{role}", "replace") == "add":
        old = s[key_df]
        names = {norm(v) for v in new["Name"]}
        new = pd.concat([old[[norm(v) not in names for v in old["Name"]]], new], ignore_index=True)
    s[key_base] = s[key_df] = new
    s.ver += 1
    s[f"paste_{role}"] = ""
    s.flash = ("success", T("f_imported_s" if role == "startup" else "f_imported_h", n=len(new)))


def paste_box(role: str):
    s = ss()
    raw = s.get(f"paste_{role}", "")
    what = T("what_s" if role == "startup" else "what_h")
    with st.expander(T("paste_title", what=what), expanded=bool(raw)):
        sheet_picker(role, f"paste_{role}")
        st.text_area(T("paste_label"),
                     key=f"paste_{role}", height=140,
                     placeholder="Startup\tPresenter\tSector\nFlexcuit\tMijin\tHardware")
        df = parse_paste(s.get(f"paste_{role}", ""))
        if df is None:
            return
        cols = STARTUP_COLS if role == "startup" else HOST_COLS
        opts = ["(ignore)"] + cols + ([SPLIT_HOST] if role == "host" else [])
        st.caption(T("rows_found", n=len(df)))
        show = {"(ignore)": T("ignore"), SPLIT_HOST: T("split_host"), **{c: col_label(role, c) for c in cols}}
        taken, grid = set(), st.columns(min(4, len(df.columns)))
        for i, c in enumerate(df.columns):
            k = map_key(role, i, c)
            if k not in s:
                g = guess_field(c, role, taken)
                if role == "host" and g == "Name" and "Contact" not in taken and \
                        sum(" - " in str(v) for v in df[c]) > len(df) / 2:
                    g = SPLIT_HOST
                s[k] = g
            taken.add(s[k])
            sample = ", ".join(str(v) for v in df[c].head(2) if not is_blank(v))[:40]
            grid[i % len(grid)].selectbox(f"{c}", opts, key=k, format_func=lambda o, show=show: show.get(o, o),
                                          help=T("eg", x=sample) if sample else None)
        a, b = st.columns([2, 1])
        a.radio("mode", ["replace", "add"], horizontal=True, key=f"pmode_{role}", label_visibility="collapsed",
                format_func=lambda o: T("mode_" + o))
        b.button(T("import_n", n=len(df)), key=f"pbtn_{role}", type="primary", on_click=cb_paste, args=(role,))


def load_project(data: dict, live: bool = True) -> None:
    """Restore an event from project_dict() output (a project file or the team workspace)."""
    days = pd.DataFrame({"Label": [d["label"] for d in data["days"]],
                         "Date": [date.fromisoformat(d["date"]) if d.get("date") else None
                                  for d in data["days"]]})
    set_frames(days, clean_frame(pd.DataFrame(data.get("startups", [])), "startup"),
               clean_frame(pd.DataFrame(data.get("hosts", [])), "host"), data.get("config"))
    res = data.get("result")
    if res and res.get("meetings"):
        res = dict(res)
        cfg = current_cfg()
        res["cfg"] = asdict(cfg)
        res["s_df"], res["h_df"] = ss().s_df, ss().h_df
        res["hash"] = frames_hash(cfg, ss().s_df, ss().h_df)
        res["id"] = int(_time.time() * 1000)
        res.setdefault("message", "")
        res.setdefault("wall", 0.0)
        slots_ = build_slots(cfg)
        st_, _ = parse_parties(ss().s_df, "startup", cfg.labels)
        ho_, _ = parse_parties(ss().h_df, "host", cfg.labels)
        res["issues"] = validate(cfg, slots_, st_, ho_, [Meeting(**m) for m in res["meetings"]],
                                 res.get("packing", "soft"), [Polish(**p) for p in res.get("polish", [])])
        ss().result = res
    else:
        ss().result = None
    if live:
        ss().statuses = data.get("statuses", {}) or {}
        ss().pol_notes = data.get("pol_notes", {}) or {}


def cb_project():
    f = ss().get("up_project")
    if not f:
        return
    try:
        data = json.loads(f.getvalue().decode("utf-8"))
        load_project(data)
        if data.get("directory") and not store():
            init_dir()
            for k, doc in data["directory"].items():
                if k in DIR_FIELDS:
                    dir_load_doc(k, doc)
        if data.get("program_docs") and not store():
            ss().local_docs = data["program_docs"]
        ss().flash = ("success", T("f_loaded_project", f=f.name))
    except Exception as e:  # noqa: BLE001
        ss().flash = ("error", T("f_project_fail", e=e))


# --------------------------------------------------------------------------------------
# Team workspace: shared database + login (active when [database] url is in secrets)
# --------------------------------------------------------------------------------------
def _secret(section: str, key: str | None = None, default=None):
    try:
        sec = st.secrets[section]
        return sec if key is None else sec.get(key, default)
    except Exception:  # noqa: BLE001  (no secrets file, or section missing)
        return default


@st.cache_resource(show_spinner=False)
def get_store(url: str) -> Store:
    return Store(url)


def store() -> Store | None:
    url = _secret("database", "url")
    if not url:
        return None
    try:
        return get_store(url)
    except Exception as e:  # noqa: BLE001
        ss().db_error = str(e)
        return None


def auth_on() -> bool:
    a = _secret("auth")
    return bool(a) and bool(a.get("client_id") or any(isinstance(v, dict) for v in dict(a).values()))


def me() -> str:
    return ss().get("user_name") or "someone"


def _user_field(name: str) -> str:
    try:
        return str(getattr(st.user, name, None) or st.user.get(name) or "")
    except Exception:  # noqa: BLE001
        return ""


def email_allowed(email: str) -> bool:
    doms = [str(d).lower().lstrip("@") for d in (_secret("access", "allowed_domains") or [])]
    emails = [str(e).lower() for e in (_secret("access", "allowed_emails") or [])]
    if not doms and not emails:
        return True
    return email in emails or email.split("@")[-1] in doms


def gate_ready() -> bool:
    """Is this person already through the login / passcode? (no UI)"""
    if not store():
        return True
    if auth_on():
        if not getattr(st.user, "is_logged_in", False):
            return False
        email = _user_field("email").lower()
        if not email_allowed(email):
            return False
        ss().user_name = _user_field("name") or email
        return True
    pc = _secret("access", "passcode")
    return not ((pc and not ss().get("pass_ok")) or not ss().get("user_name"))


def gate() -> bool:
    """Login (Google) or team passcode before the shared workspace opens."""
    if not store():
        return True
    _, mid, _ = st.columns([1, 1.15, 1])   # a narrow, centred sign-in card
    if auth_on():
        if not getattr(st.user, "is_logged_in", False):
            with mid.container(border=True):
                st.markdown(f"#### {T('ws_signin')}")
                st.caption(T("login_needed"))
                st.button(T("login_google"), type="primary", on_click=st.login, width="stretch")
            return False
        email = _user_field("email").lower()
        if not email_allowed(email):
            with mid.container(border=True):
                st.error(T("not_allowed", email=email))
                st.button(T("logout"), on_click=st.logout)
            return False
        ss().user_name = _user_field("name") or email
        return True
    pc = _secret("access", "passcode")
    if (pc and not ss().get("pass_ok")) or not ss().get("user_name"):
        with mid.form("enter_ws"):
            st.markdown(f"#### {T('ws_signin')}")
            v = st.text_input(T("passcode"), type="password") if pc and not ss().get("pass_ok") else None
            n = st.text_input(T("your_name"), value=ss().get("user_name", ""))
            if st.form_submit_button(T("enter"), type="primary", width="stretch"):
                if (v is None or v == str(pc)) and n.strip():
                    ss().pass_ok, ss().user_name = True, n.strip()
                    st.rerun()
                st.error(T("pass_wrong"))
        return False
    return True


def canon(d: dict) -> str:
    return json.dumps(d, sort_keys=True, default=str)


def items_map() -> dict:
    m = {(k, "status"): v for k, v in ss().statuses.items() if v in STATUSES}
    m.update({(k, "note"): v for k, v in ss().get("pol_notes", {}).items() if v})
    return m


def open_event(eid: str) -> bool:
    """Load a shared event into this session (safe to call from callbacks and before widgets render)."""
    S = store()
    row = S.get(eid) if S else None
    if not row or row.get("archived"):
        ss().flash = ("error", T("ws_missing"))
        ss().event_id = None
        return False
    load_project(json.loads(row["data"]), live=False)
    items = S.items(eid)
    ss().statuses = {r["item_id"]: r["value"] for r in items if r["kind"] == "status" and r["value"] in STATUSES}
    ss().pol_notes = {r["item_id"]: r["value"] for r in items if r["kind"] == "note" and r["value"]}
    ss().items_synced = items_map()
    ss().items_t = max([r["updated_at"] for r in items], default=0.0)
    ss().event_id, ss().ev_version = eid, row["version"]
    ss().ev_by, ss().ev_at = row.get("updated_by") or "", row.get("updated_at") or 0
    ss().saved_hash = None          # the first run after loading sets the baseline
    ss().ev_pick = eid
    return True


def cb_pick():
    if ss().get("ev_pick") and ss().ev_pick != ss().get("event_id"):
        open_event(ss().ev_pick)


def cb_new():
    S, s = store(), ss()
    name = (s.get("new_ev_name") or "").strip() or "New matching event"
    src = s.get("new_ev_from", "blank")
    if src == "copy":
        d = project_dict(live=False)
    else:
        set_frames(*(sample_frames() if src == "sample" else blank_frames()))
        s.result, s.statuses, s.pol_notes = None, {}, {}
        d = project_dict(live=False)
    d["config"]["name"] = name
    eid, _ = S.create(name, canon(d), me())
    open_event(eid)
    s.new_ev_name = ""
    s.flash = ("success", T("ws_created", name=name))


def cb_archive():
    S, eid = store(), ss().get("event_id")
    if S and eid and ss().get("del_confirm"):
        S.archive(eid, me())
        ss().event_id, ss().del_confirm = None, False


def workspace_bar():
    S = store()
    if not S:
        if _secret("database", "url"):
            st.error(T("ws_db_error", e=ss().get("db_error", "")))
        else:
            st.caption(T("ws_off"))
        return
    evs = S.list_events()
    names = {e["id"]: e["name"] or "Untitled" for e in evs}
    if not evs or ss().get("event_id") not in names:
        return
    ss().ev_pick = ss().event_id
    c1, c2, c3, c4 = st.columns([3, 1, 1, 2.2], vertical_alignment="bottom")
    c1.selectbox(T("ws_event"), list(names), format_func=lambda i: names.get(i, "?"), key="ev_pick",
                 on_change=cb_pick)
    with c2.popover(T("ws_new"), width="stretch"):
        st.text_input(T("ws_new_name"), key="new_ev_name")
        st.radio("from", ["blank", "sample", "copy"], key="new_ev_from", label_visibility="collapsed",
                 format_func=lambda o: T("ws_from_" + o))
        st.button(T("ws_create"), type="primary", on_click=cb_new)
    with c3.popover(T("ws_more"), width="stretch"):
        st.checkbox(T("ws_delete_confirm"), key="del_confirm")
        st.button(T("ws_delete"), on_click=cb_archive, disabled=not ss().get("del_confirm"))
    when = datetime.fromtimestamp(ss().get("ev_at") or _time.time(), ZoneInfo(ss().get("tz") or "UTC"))
    c4.caption(T("ws_saved", when=when.strftime("%d %b %H:%M"), who=ss().get("ev_by") or "-") + "  \n"
               + T("ws_user", who=me()))
    if auth_on():
        c4.button(T("logout"), on_click=st.logout)


def autosave():
    """Runs at the end of every rerun: saves the event if it changed, then pushes status / note changes."""
    S, eid = store(), ss().get("event_id")
    if not S or not eid:
        return
    d = project_dict(live=False)
    txt = canon(d)
    h = hashlib.sha1(txt.encode()).hexdigest()
    if ss().get("saved_hash") is None:
        ss().saved_hash = h
    elif h != ss().saved_hash:
        v = S.save(eid, d["config"]["name"], txt, ss().ev_version, me())
        if v:
            ss().ev_version, ss().saved_hash, ss().ev_by, ss().ev_at = v, h, me(), _time.time()
        else:
            row = S.get(eid)
            if row and canon(json.loads(row["data"])) == txt:
                ss().ev_version, ss().saved_hash = row["version"], h
            else:
                ss().pending_reload = ("conflict", (row or {}).get("updated_by") or "")
                st.rerun()
    cur, old = items_map(), ss().get("items_synced", {})
    rows = [(k[0], k[1], v) for k, v in cur.items() if old.get(k) != v] + \
           [(k[0], k[1], "") for k in old if k not in cur]
    if rows:
        S.put_items(eid, rows, me())
        ss().items_synced = cur


@st.fragment(run_every=10)
def live_sync():
    """Every 10 s: pick up other people's saves and status changes."""
    S, eid = store(), ss().get("event_id")
    if not S or not eid:
        return
    h = S.head(eid)
    if not h or h.get("archived"):
        return
    if h["version"] > ss().get("ev_version", 0):
        ss().pending_reload = ("remote", h.get("updated_by") or "")
        st.rerun(scope="app")
    new = S.items(eid, ss().get("items_t", 0.0))
    if not new:
        return
    ss().items_t = max(r["updated_at"] for r in new)
    synced, changed = ss().setdefault("items_synced", {}), False
    for r in new:
        target = ss().statuses if r["kind"] == "status" else ss().pol_notes
        val = r["value"]
        if target.get(r["item_id"], "") != val:
            changed = True
            if val:
                target[r["item_id"]] = val
            else:
                target.pop(r["item_id"], None)
        if val:
            synced[(r["item_id"], r["kind"])] = val
        else:
            synced.pop((r["item_id"], r["kind"]), None)
    if changed:
        if ss().get("result"):
            ss().result["version"] = f"py-sync-{int(_time.time() * 1000)}"
        st.rerun(scope="app")


# --------------------------------------------------------------------------------------
# Drag-and-drop board (custom component in ./board)
# --------------------------------------------------------------------------------------
def cb_practice_edit(key: str, ids: list[str]):
    """Status / note edits in the pitching practice table. Bumping the version reloads the board."""
    s = ss()
    for i, ch in s[key].get("edited_rows", {}).items():
        if i >= len(ids):
            continue
        pid = ids[i]
        if "Status" in ch:
            if ch["Status"] in STATUSES:
                s.statuses[pid] = ch["Status"]
            else:
                s.statuses.pop(pid, None)
        if "Focus / Mentor Note" in ch:
            v = (ch["Focus / Mentor Note"] or "").strip()
            if v:
                s.pol_notes[pid] = v
            else:
                s.pol_notes.pop(pid, None)
    if s.result:
        s.result["version"] = f"py-{int(_time.time() * 1000)}"


def board_payload(day, cfg, slots, startups, hosts, r) -> dict:
    ds = [x for x in slots if x.day == day]
    polish = [Polish(**p) for p in r.get("polish", [])]
    rows = []
    for row in day_rows(day, cfg, slots, startups, hosts, []):
        if row["lunch"]:
            rows.append({"kind": "lunch", "label": T("lunch_row"), "time": row["time"].replace("–", " – ")})
        elif row.get("divider"):
            rows.append({"kind": "divider", "label": row["label"], "time": ""})
        elif row.get("tables") is not None:
            rows.append({"kind": "slot", "k": row["slot"].k, "start": row["start"], "end": row["end"],
                         "session": row["session"]})
    hcol = host_colors(hosts)
    prior = defaultdict(int)
    for p in polish:
        if p.day < day:
            prior[p.startup] += 1
    day_pol = [p for p in polish if p.day == day]
    other_n = defaultdict(int)
    for m in r["meetings"]:
        if m["day"] != day:
            other_n[m["startup"]] += 1
    return {
        "version": r.get("version") or f"py-{r['id']}",
        "day": day, "tables": cfg.tables, "tz": cfg.tz, "date": str(cfg.day_date(day) or ""),
        "rows": rows, "sessions": cfg.session_list(),
        "startups": [{"name": x.name, "priority": x.priority, "presenter": x.presenter,
                      "avail": [z.k for z in ds if x.available(z)],
                      "no": [h.name for h in hosts if not compatible(x, h)],
                      "noWhy": {h.name: mismatch(x, h) for h in hosts if not compatible(x, h)},
                      "pref": [h.name for h in hosts if prefers(h, x)],
                      "tags": x.domain_text, "otherN": other_n[x.name]}
                     for x in startups if not x.absent],
        # pairs already scheduled on other days (each startup meets a partner once)
        "other": {**{f"{m['startup']}|{m['host']}": "-" for m in r.get("parked", []) if m["day"] != day},
                  **{f"{m['startup']}|{m['host']}": f"{cfg.day_short(m['day'])} {fmt(m['start'])}"
                     for m in r["meetings"] if m["day"] != day}},
        "hosts": [{"name": h.name, "contact": h.contact, "color": hcol[h.name], "text": text_on(hcol[h.name]),
                   "avail": [z.k for z in ds if h.available(z)], "kind": h.kind, "only": h.only_text,
                   "prefer": h.prefer_text, "maxM": h.max_m} for h in hosts],
        "meetings": [{"id": f"{m['startup']}|{m['host']}", "startup": m["startup"], "host": m["host"],
                      "k": m["k"], "table": m["table"], "lock": bool(m.get("lock"))}
                     for m in r["meetings"] if m["day"] == day],
        "parked": [{"id": f"{m['startup']}|{m['host']}", "startup": m["startup"], "host": m["host"],
                    "k": None, "table": 0, "lock": False} for m in r.get("parked", []) if m["day"] == day],
        "polish": [{"id": p.pid, "startup": p.startup, "k": p.k, "note": ss().get("pol_notes", {}).get(p.pid, "")}
                   for p in day_pol],
        "polOn": day in cfg.pol_days or bool(day_pol),
        "polOk": [z.k for z in ds if cfg.polish_ok(z)], "polCap": max(1, cfg.pol_cap),
        "polPrior": dict(prior), "statuses": dict(ss().statuses),
        "notes": attendance_notes(cfg, startups),
        "names": {k_: (cfg.names or {}).get(k_) or T(f"nm_{k_}") for k_ in DEFAULT_NAMES},
        "lang": lang(),
    }


def apply_board(r, cfg, slots, startups, hosts, day, val) -> bool:
    """Take the edits made on the board (moves, pushes, statuses) into the stored result."""
    if not isinstance(val, dict) or val.get("day") != day or not val.get("version"):
        return False
    if val["version"] == r.get("version") or val["version"] in r.get("applied", []):
        return False
    r.setdefault("applied", []).append(val["version"])
    sl = {x.k: x for x in slots if x.day == day}
    kept = [m for m in r["meetings"] if m["day"] != day]
    new = [asdict(Meeting(day, x["k"], sl[x["k"]].start, sl[x["k"]].end, x["startup"], x["host"],
                          int(x.get("table") or 0), bool(x.get("lock"))))
           for x in val.get("meetings", []) if x.get("k") in sl]
    r["meetings"] = sorted(kept + new, key=lambda m: (m["k"], m["table"]))
    r["parked"] = [m for m in r.get("parked", []) if m["day"] != day] + [
        asdict(Meeting(day, -1, 0, 0, x["startup"], x["host"])) for x in val.get("parked", [])]
    pol = [Polish(**p) for p in r.get("polish", []) if p["day"] != day] + [
        Polish(day, x["k"], sl[x["k"]].start, sl[x["k"]].end, x["startup"])
        for x in val.get("polish", []) if x.get("k") in sl]
    r["polish"] = [asdict(p) for p in number_polish(pol)]
    ss().statuses = {k: v for k, v in (val.get("statuses") or {}).items() if v in STATUSES}
    r["version"] = val["version"]
    r["issues"] = validate(cfg, slots, startups, hosts, [Meeting(**m) for m in r["meetings"]],
                           r.get("packing", "soft"), [Polish(**p) for p in r["polish"]])
    return True


def summary_payload(cfg, startups, hosts, r) -> dict:
    hcol = host_colors(hosts)
    days = []
    for d in range(len(cfg.days)):
        dd = cfg.day_date(d)
        days.append({"title": cfg.days[d]["label"] + (f" · {dd:%a %d %b}" if dd else ""),
                     "short": cfg.day_short(d), "date": str(dd or "")})
    meetings = [{"id": f"{m['startup']}|{m['host']}", "day": m["day"], "start": m["start"], "end": m["end"],
                 "startup": m["startup"], "host": m["host"], "table": m["table"]} for m in r["meetings"]]
    polish = [{"id": p.pid, "day": p.day, "start": p.start, "end": p.end, "startup": p.startup, "n": p.n,
               "note": ss().get("pol_notes", {}).get(p.pid, "")} for p in [Polish(**x) for x in r.get("polish", [])]]
    tz_short = (cfg.tz or "").split("/")[-1].replace("_", " ")
    sts = {k: v for k, v in ss().statuses.items() if v in STATUSES}
    version = hashlib.sha1(json.dumps([sts, meetings, polish], sort_keys=True, default=str).encode()).hexdigest()[:16]
    return {
        "version": version, "lang": lang(), "tz": cfg.tz, "tzShort": tz_short, "days": days,
        "startups": [{"name": x.name, "priority": x.priority, "presenter": x.presenter, "days": sorted(x.days)}
                     for x in startups],
        "hosts": [{"name": h.name, "contact": h.contact, "kind": h.kind, "color": hcol[h.name], "days": sorted(h.days),
                   "only": h.only_text, "prefer": h.prefer_text} for h in hosts],
        "meetings": meetings, "polish": polish, "statuses": sts,
    }


def apply_summary(r, val) -> bool:
    """Status changes made on the summary dashboard."""
    if not isinstance(val, dict) or not val.get("version") or val["version"] in r.get("sum_applied", []):
        return False
    r.setdefault("sum_applied", []).append(val["version"])
    for k, v in (val.get("set") or {}).items():
        if v in STATUSES:
            ss().statuses[k] = v
        else:
            ss().statuses.pop(k, None)
    r["version"] = f"py-{int(_time.time() * 1000)}"     # the board picks up the new statuses
    return True


def summary_tab():
    s = ss()
    r = s.result
    if not r or not r.get("meetings"):
        st.info(T("gen_first"))
        return
    cfg = EventConfig(**r["cfg"])
    startups, _ = parse_parties(r["s_df"], "startup", cfg.labels)
    hosts, _ = parse_parties(r["h_df"], "host", cfg.labels)
    key = f"sum_{r['id']}"
    if apply_summary(r, s.get(key)):
        st.rerun()
    st.caption(T("sum_cap"))
    SUMMARY(data=summary_payload(cfg, startups, hosts, r), key=key, default=None)


I18N_KEYS = set(__import__("i18n").S)


def ko_headers(df: pd.DataFrame) -> pd.DataFrame:
    if lang() != "ko":
        return df
    def h(c):
        if c in SUMMARY_KO:
            return SUMMARY_KO[c]
        if c.endswith(" Partners"):
            return c[:-9] + " 파트너"
        return c
    return df.rename(columns=h)


def cb_lang():
    st.query_params["lang"] = ss().lang


def setup_tab():
    s = ss()
    st.subheader(T("event"))
    c1, c2, c3 = st.columns([1.2, 1, 1])
    with c1:
        st.text_input(T("event_name"), key="ev_name")
        s.days_df = st.data_editor(
            s.days_base, key=f"days_{s.ver}", num_rows="dynamic", hide_index=True,
            column_config={"Label": st.column_config.TextColumn(T("col_day"), required=True),
                           "Date": st.column_config.DateColumn(T("col_date"), format="ddd D MMM YYYY",
                                                               help=T("date_help"))})
    with c2:
        a, b = st.columns(2)
        a.time_input(T("day_starts"), key="t_start", step=timedelta(minutes=5))
        b.time_input(T("day_ends"), key="t_end", step=timedelta(minutes=5))
        a.number_input(T("meeting_len"), 5, 180, step=1, key="interval")
        g1, g2 = b.columns(2)
        g1.number_input(T("gap_min"), 0, 60, step=1, key="buffer_m", help=T("gap_help"))
        g2.number_input(T("gap_sec"), 0, 50, step=10, key="buffer_s")
        st.number_input(T("tables"), 1, 200, step=1, key="tables")
    with c3:
        st.checkbox(T("lunch"), key="lunch_on")
        a, b = st.columns(2)
        a.time_input(T("lunch_from"), key="t_lunch_s", step=timedelta(minutes=5), disabled=not s.lunch_on)
        b.time_input(T("lunch_until"), key="t_lunch_e", step=timedelta(minutes=5), disabled=not s.lunch_on)
        st.selectbox(T("tz"), TIMEZONES, key="tz", help=T("tz_help"),
                     format_func=lambda z: TZ_HINT.get(z, z))

    cfg = current_cfg()
    slots = build_slots(cfg)
    per_day = len(slots) // max(1, len(cfg.days))
    if slots:
        times = " ".join(fmt(x.start) for x in slots if x.day == 0)
        st.caption(T("slots_caption", per=per_day, days=len(cfg.days), n=len(slots), times=times))
    else:
        st.warning(T("no_slot"))

    st.subheader(T("sess_pol"))
    c1, c2 = st.columns([1, 1.4])
    with c1:
        st.text_input(T("sessions"), key="sessions_txt", help=T("sessions_help"))
        if s.get("sess_err"):
            st.error(s.sess_err)
        else:
            st.caption(" · ".join(f"{x['name']} {fmt(x['start'])}–{fmt(x['end'])}" for x in cfg.session_list()))
    with c2:
        st.checkbox(T("pol_on"), key="pol_on", help=T("pol_help"))
        s.pol_days = [x for x in s.get("pol_days", []) if x in cfg.labels]
        a, b, c, d = st.columns([1.4, 1, 1, 0.8])
        a.multiselect(T("pol_days"), cfg.labels, key="pol_days", disabled=not s.pol_on)
        b.time_input(T("from"), key="t_pol_s", step=timedelta(minutes=5), disabled=not s.pol_on)
        c.time_input(T("until"), key="t_pol_e", step=timedelta(minutes=5), disabled=not s.pol_on)
        d.number_input(T("coaches"), 1, 20, key="pol_cap", disabled=not s.pol_on, help=T("coaches_help"))

    with st.expander(T("colnames")):
        st.caption(T("colnames_cap"))
        cols_ = st.columns(len(DEFAULT_NAMES))
        for c_, k_ in zip(cols_, DEFAULT_NAMES):
            c_.text_input(T(f"nm_{k_}"), key=f"nm_{k_}")

    st.divider()
    st.subheader(T("startups"))
    u1, u0, u2, u3 = st.columns([3, 1.2, 1, 1])
    dir_pick_popover("startup", u0)
    rename_popover("startup", u3)
    u1.file_uploader(T("import_file"), type=["csv", "xlsx", "xls"], key="up_startup",
                     on_change=cb_upload, args=("startup",), label_visibility="collapsed")
    u2.download_button(T("template_s"), sample_frames()[1].to_csv(index=False).encode("utf-8-sig"),
                       "startups_template.csv", "text/csv")
    paste_box("startup")
    s.s_df = st.data_editor(
        s.s_base, key=f"s_{s.ver}", num_rows="dynamic", hide_index=True,
        column_config=labelled("startup", {
            "Name": st.column_config.TextColumn(required=True, width="medium"),
            "Domain": st.column_config.TextColumn(help=T("h_domain")),
            "Days": st.column_config.TextColumn(help=T("h_days_s")),
            "Available From": st.column_config.TextColumn(help=T("h_from"), validate=TIME_VALID),
            "Available Until": st.column_config.TextColumn(help=T("h_until"), validate=TIME_VALID),
            "Blocked Times": st.column_config.TextColumn(help=T("h_blocked"), width="medium"),
            "Priority": st.column_config.CheckboxColumn(help=T("h_priority")),
            "Presenter": st.column_config.TextColumn(help=T("h_presenter")),
            "Pitch Polishing": st.column_config.CheckboxColumn(default=True, help=T("h_polish")),
            "Max Meetings": st.column_config.NumberColumn(min_value=0, step=1, format="%d", help=T("h_max")),
        }))

    st.subheader(T("hosts"))
    u1, u0, u2, u3 = st.columns([3, 1.2, 1, 1])
    dir_pick_popover("host", u0)
    rename_popover("host", u3)
    u1.file_uploader(T("import_file"), type=["csv", "xlsx", "xls"], key="up_host",
                     on_change=cb_upload, args=("host",), label_visibility="collapsed")
    u2.download_button(T("template_h"), sample_frames()[2].to_csv(index=False).encode("utf-8-sig"),
                       "hosts_template.csv", "text/csv")
    paste_box("host")
    s.h_df = st.data_editor(
        s.h_base, key=f"h_{s.ver}", num_rows="dynamic", hide_index=True,
        column_config=labelled("host", {
            "Name": st.column_config.TextColumn(required=True, width="medium"),
            "Contact": st.column_config.TextColumn(help=T("h_contact")),
            "Type": st.column_config.SelectboxColumn(options=HOST_TYPES, default="Investor"),
            "Days": st.column_config.TextColumn(help=T("h_days_h")),
            "Available From": st.column_config.TextColumn(help=T("h_from"), validate=TIME_VALID),
            "Available Until": st.column_config.TextColumn(help=T("h_until"), validate=TIME_VALID),
            "Blocked Times": st.column_config.TextColumn(help=T("h_blocked"), width="medium"),
            "Only Domains": st.column_config.TextColumn(help=T("h_only"), width="medium"),
            "Preferred Domains": st.column_config.TextColumn(help=T("h_prefer"), width="medium"),
            "Excluded Domains": st.column_config.TextColumn(help=T("h_excluded"), width="medium"),
            "Max Meetings": st.column_config.NumberColumn(min_value=0, step=1, format="%d"),
        }))

    st.divider()
    a, b, c, d = st.columns(4)
    if not store():   # in the team workspace, new events come from the New event menu instead
        a.button(T("load_sample"), on_click=cb_sample)
        b.button(T("start_blank"), on_click=cb_blank)
    c.download_button(T("save_project"), project_json(), "bizmatch_project.json", "application/json")
    d.file_uploader(T("open_project"), type=["json"], key="up_project", on_change=cb_project,
                    label_visibility="collapsed")


def generate_tab():
    s = ss()
    cfg = current_cfg()
    slots = build_slots(cfg)
    startups, e1 = parse_parties(s.s_df, "startup", cfg.labels)
    hosts, e2 = parse_parties(s.h_df, "host", cfg.labels)
    errors = e1 + e2
    if errors:
        st.error(T("fix_errors") + "\n\n" + "\n".join(f"- {e}" for e in errors))

    attending = [x for x in startups if not x.absent]
    for h in hosts:
        if h.only and not h.absent and not any(compatible(a, h) for a in attending):
            st.warning(T("only_none", h=h.name, tags=h.only_text))
    pairs = sum(compatible(a, b) for a in attending for b in hosts if not b.absent)
    peak = max((sum(h.available(x) for h in hosts) for x in slots), default=0)
    c = st.columns(5)
    c[0].metric(T("m_attending"), len(attending),
                T("not_here", n=len(startups) - len(attending)) if len(startups) > len(attending) else None,
                delta_color="off")
    c[1].metric(T("m_hosts"), len(hosts))
    c[2].metric(T("m_slots"), len(slots))
    c[3].metric(T("m_pairs"), pairs)
    c[4].metric(T("m_poldays"), len(cfg.pol_days) if cfg.pol_days else T("off"))

    st.subheader(T("rules"))
    default = "front" if peak <= cfg.tables else "block"
    if s.get("packing_mode") not in PACKING:
        s.packing_mode = default
    a, b = st.columns([1.3, 1])
    with a:
        st.radio(T("packing"), list(PACKING), format_func=lambda k: T("pack_" + k), key="packing_mode")
        if s.packing_mode == "front" and peak > cfg.tables:
            st.warning(T("warn_front", peak=peak, tables=cfg.tables))
    with b:
        for k_, v_ in (("pw", 3), ("tl", 20), ("hf", True), ("pe", True)):
            ss().setdefault(k_, v_)
        st.slider(T("pw"), 1, 5, key="pw", help=T("pw_help"))
        st.slider(T("tl"), 5, 120, key="tl")
        st.checkbox(T("hf"), key="hf")
        st.checkbox(T("pe"), key="pe")

    keep_live = False
    if s.result and s.result.get("meetings"):
        ss().setdefault("keep_live", True)
        keep_live = st.checkbox(T("keep_live"), key="keep_live")

    st.caption(T("always"))
    go = st.button(T("generate"), type="primary",
                   disabled=bool(errors) or not startups or not hosts or not slots)
    if go:
        fixed, fixed_pol = [], []
        if s.result and s.result.get("meetings"):
            old_cfg = EventConfig(**s.result["cfg"])
            now = now_for(old_cfg)
            for md in s.result["meetings"]:
                m = Meeting(**md)
                # pinned meetings always stay; started / finished ones stay when the box is ticked
                if m.lock or (keep_live and eff_status(m, old_cfg, s.statuses, now) in LOCK_STATUSES):
                    fixed.append((m.startup, m.host, m.day, m.start))
            for pd_ in (s.result.get("polish", []) if keep_live else []):
                p_ = Polish(**pd_)
                dd = old_cfg.day_date(p_.day)
                if dd and (dd < now.date() or (dd == now.date() and p_.start <= now.hour * 60 + now.minute)):
                    fixed_pol.append((p_.startup, p_.day, p_.start))
        settings = SolverSettings(packing=s.packing_mode, time_limit=s.tl, priority_weight=s.pw,
                                  host_fairness=s.hf, prefer_early=s.pe)
        with st.spinner(T("spinner")):
            res = solve_schedule(cfg, slots, startups, hosts, settings, fixed, fixed_pol)
        if res.meetings:
            pinned = {(md["startup"], md["host"]) for md in (s.result or {}).get("meetings", []) if md.get("lock")}
            for m in res.meetings:
                m.lock = (m.startup, m.host) in pinned
            s.result = {
                "id": int(_time.time() * 1000), "meetings": [asdict(m) for m in res.meetings],
                "polish": [asdict(p_) for p_ in res.polish], "parked": [],
                "cfg": asdict(cfg), "s_df": s.s_df.copy(), "h_df": s.h_df.copy(),
                "packing": s.packing_mode, "status": res.status, "message": res.message,
                "wall": res.wall, "n_vars": res.n_vars, "kept": res.kept,
                "dropped": res.dropped_fixed, "hash": frames_hash(cfg, s.s_df, s.h_df),
                "issues": validate(cfg, slots, startups, hosts, res.meetings, s.packing_mode, res.polish),
            }
            pol_txt = T("gen_pol", n=len(res.polish)) if res.polish else ""
            st.success(T("gen_success", msg=T("solve_" + res.status), n=len(res.meetings), pol=pol_txt,
                         sec=f"{res.wall:.1f}"))
            if res.dropped_fixed:
                st.warning(T("dropped") + "; ".join(res.dropped_fixed))
        else:
            st.error(f"{res.status}: " + (T("solve_" + res.status) if "solve_" + res.status in I18N_KEYS
                                          else res.message))

    if s.result:
        r = s.result
        st.caption(T("last_run", st=r["status"], n=len(r["meetings"]), sec=f"{r.get('wall', 0):.1f}",
                     v=f"{r.get('n_vars', 0):,}"))


def results_tab():
    s = ss()
    r = s.result
    if not r or not r.get("meetings"):
        st.info(T("gen_first"))
        return
    cfg = EventConfig(**r["cfg"])
    slots = build_slots(cfg)
    startups, _ = parse_parties(r["s_df"], "startup", cfg.labels)
    hosts, _ = parse_parties(r["h_df"], "host", cfg.labels)
    now = now_for(cfg)

    dc, hc = st.columns([6, 1], vertical_alignment="bottom")
    day = dc.radio(T("day"), list(range(len(cfg.days))), format_func=cfg.day_title, horizontal=True, key="day_sel")
    with hc.popover(T("how_board"), width="stretch"):
        st.markdown(T("board_cap"))
    key = f"board_{r['id']}_{day}"
    apply_board(r, cfg, slots, startups, hosts, day, s.get(key))

    meetings = [Meeting(**m) for m in r["meetings"]]
    polish = [Polish(**p) for p in r.get("polish", [])]
    parked = r.get("parked", [])

    if r.get("hash") and r["hash"] != frames_hash(current_cfg(), s.s_df, s.h_df):
        st.warning(T("setup_changed"))
    issues = r.get("issues", [])
    hard = [i for i in issues if not i.startswith(("Gap in", "Note:"))]
    soft = [i for i in issues if i.startswith(("Gap in", "Note:"))]

    counts = defaultdict(int)
    for m in meetings:
        counts[m.startup] += 1
    att = [x for x in startups if not x.absent]
    BOARD(data=board_payload(day, cfg, slots, startups, hosts, r), key=key, default=None)

    c = st.columns(5)
    c[0].metric(T("m_meetings"), len(meetings), T("unplaced", n=len(parked)) if parked else None,
                delta_color="inverse")
    c[1].metric(T("m_with"), f"{sum(1 for x in att if counts[x.name])}/{len(att)}")
    c[2].metric(T("m_avg"), f"{len(meetings) / max(1, len(att)):.1f}")
    c[3].metric(T("m_pol"), len(polish) if polish else T("off"))
    c[4].metric(T("m_clash"), len(hard))

    if hard:
        with st.expander(T("clashes_exp", n=len(hard)), expanded=False):
            st.markdown("\n".join(f"- {i}" for i in hard[:40]))
    elif not parked:
        st.success(T("verified") + (T("verified_pol") if polish else ""))
    if soft:
        st.caption(T("gaps_note", n=len(soft)))

    day_pol = sorted([p for p in polish if p.day == day], key=lambda p: (p.k, p.startup))
    if day_pol or day in cfg.pol_days:
        st.subheader(T("practice"))
        st.caption(T("practice_cap"))
        P_ = {x.name: x for x in startups}
        pdf = pd.DataFrame([{
            "Time": f"{fmt(p.start)}–{fmt(p.end)}", "Startup": p.startup,
            "Presenter": P_[p.startup].presenter if p.startup in P_ else "",
            "Round": T("round_1") if p.n == 1 else T("round_2") if p.n == 2 else T("round_n", n=p.n),
            "Status": s.statuses.get(p.pid, "Auto"),
            "Focus / Mentor Note": s.pol_notes.get(p.pid, ""),
        } for p in day_pol], columns=["Time", "Startup", "Presenter", "Round", "Status", "Focus / Mentor Note"])
        pkey = f"pp_{r['id']}_{day}_{r.get('version', '')}"
        st.data_editor(pdf, key=pkey, hide_index=True, disabled=["Time", "Startup", "Presenter", "Round"],
                       on_change=cb_practice_edit, args=(pkey, [p.pid for p in day_pol]),
                       column_config={
                           "Time": st.column_config.TextColumn(T("pc_time")),
                           "Startup": st.column_config.TextColumn(T("pc_startup")),
                           "Round": st.column_config.TextColumn(T("pc_round")),
                           "Status": st.column_config.SelectboxColumn(T("pc_status"), options=STATUS_CHOICES,
                                                                      required=True, help=T("auto_help")),
                           "Presenter": st.column_config.TextColumn(T("pc_presenter"), help=T("presenter_help")),
                           "Focus / Mentor Note": st.column_config.TextColumn(T("pc_note"), width="large")})

    stamp = now.strftime("%Y%m%d_%H%M")
    a, b = st.columns(2)
    a.download_button(T("dl_excel"),
                      build_excel(cfg, slots, startups, hosts, meetings, s.statuses, now, polish,
                                  s.get("pol_notes", {})),
                      f"matching_schedule_{stamp}.xlsx",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    b.download_button(T("dl_csv"),
                      meetings_table(cfg, meetings, s.statuses, now, hosts).to_csv(index=False).encode("utf-8-sig"),
                      f"meetings_{stamp}.csv", "text/csv")
    if lang() == "ko":
        st.caption(T("excel_lang"))


# --------------------------------------------------------------------------------------
# Company directory: the team's master list of startups and partners. Events pull from it.
# --------------------------------------------------------------------------------------
DIR_FIELDS = {
    "startup": ["Company (EN)", "Company (KR)", "CEO", "CEO (EN)", "Email", "Phone", "Website", "Tags",
                "One-liner", "Founded", "Stage", "Employees", "Presenter", "Booth", "Deck", "Notes"],
    "partner": ["Organisation", "Contact", "Title", "Email", "Phone", "Type", "Country", "Prefers",
                "Only meets", "Excludes", "Website", "Notes"],
}
DIR_NAME = {"startup": "Company (EN)", "partner": "Organisation"}
ST_KO = {"Scheduled": "예정", "Ongoing": "진행 중", "Done": "완료", "No-Show": "노쇼"}


def st_label(x: str) -> str:
    return ST_KO.get(x, x) if lang() == "ko" else x
PAGES: dict = {}
DIR_HELP = {"Company (EN)": "dir_h_name_en", "Tags": "dir_h_tags", "Presenter": "dir_h_presenter",
            "Prefers": "h_prefer", "Only meets": "h_only", "Excludes": "h_excluded", "Type": "dir_h_type"}
DIR_ALIASES = {  # (field, words that identify a header), checked in order; Korean and English
    "startup": [
        ("CEO (EN)", ("대표자 영문", "대표 영문", "ceo (en)", "ceo en", "ceo english", "representative (en)")),
        ("Company (EN)", ("영문", "english", "company (en)", "company name", "startup name", "startup", "company", "name")),
        ("Company (KR)", ("국문", "회사명 (국문)", "기업명", "회사명", "한글")),
        ("CEO", ("대표", "ceo", "founder", "representative")),
        ("Email", ("mail", "메일")),
        ("Phone", ("phone", "전화", "연락처", "mobile", "휴대")),
        ("Website", ("website", "홈페이지", "웹사이트", "url", "web")),
        ("Tags", ("tag", "domain", "sector", "industry", "분야", "업종", "카테고리", "category")),
        ("One-liner", ("one-liner", "one liner", "한줄", "한 줄", "소개", "description", "summary", "pitch")),
        ("Founded", ("found", "설립")),
        ("Stage", ("stage", "단계", "round", "라운드")),
        ("Employees", ("employee", "직원", "인원", "headcount")),
        ("Presenter", ("present", "발표", "피칭", "attendee", "참석")),
        ("Booth", ("booth", "부스")),
        ("Deck", ("deck", "ir", "소개서", "자료")),
        ("Notes", ("note", "메모", "비고", "remark")),
    ],
    "partner": [
        ("Only meets", ("only", "전용", "만 만남", "이 분야만")),
        ("Excludes", ("exclud", "제외", "avoid")),
        ("Prefers", ("prefer", "interest", "관심", "선호", "focus", "thesis")),
        ("Organisation", ("organi", "기관", "회사", "company", "investor", "firm", "partner", "corporate")),
        ("Title", ("title", "직함", "직책", "position", "role")),
        ("Contact", ("contact", "담당", "이름", "name", "person")),
        ("Email", ("mail", "메일")),
        ("Phone", ("phone", "전화", "연락처", "mobile", "휴대")),
        ("Type", ("type", "유형", "구분", "kind")),
        ("Country", ("country", "국가", "region", "지역")),
        ("Website", ("website", "홈페이지", "웹사이트", "url", "web")),
        ("Notes", ("note", "메모", "비고", "remark")),
    ],
}


def dir_cols(kind: str) -> list[str]:
    extra = (ss().get("dir_extra") or {}).get(kind, [])
    return DIR_FIELDS[kind] + [c for c in extra if c not in DIR_FIELDS[kind]]


def dir_label(col: str) -> str:
    return DIR_LABELS.get(lang(), {}).get(col) or DIR_LABELS["en"].get(col) or col


def dir_tidy(df: pd.DataFrame, kind: str, ver: int | None = None) -> pd.DataFrame:
    """Every column as text, blank rows dropped, a stable id on every row."""
    cols = dir_cols(kind)
    df = df.copy()
    for c in ["_id"] + cols:
        if c not in df.columns:
            df[c] = ""
    df = df[["_id"] + cols].astype(object).where(pd.notna(df[["_id"] + cols]), "")
    df = df.map(lambda v: "" if is_blank(v) else str(v).strip())
    keep = pd.Series([any(r[c] for c in cols) for _, r in df.iterrows()], index=df.index, dtype=bool)
    df = df[keep].reset_index(drop=True)
    tag = f"{ss().setdefault('sid', uuid.uuid4().hex[:6])}{ver if ver is not None else ss().get('dir_ver', 0)}"
    df["_id"] = [v or f"{tag}r{i}" for i, v in enumerate(df["_id"])]
    return df


def dir_set(kind: str, df: pd.DataFrame) -> None:
    ss().dir_base[kind] = ss().dir_df[kind] = dir_tidy(df, kind)
    ss().dir_ver = ss().get("dir_ver", 0) + 1


def dir_doc(kind: str) -> str:
    df = ss().dir_df[kind]
    return json.dumps({"fields": ss().dir_extra.get(kind, []), "rows": df.to_dict("records")},
                      sort_keys=True, ensure_ascii=False)


def init_dir() -> None:
    s = ss()
    if "dir_base" not in s:
        s.dir_extra = {"startup": [], "partner": []}
        s.dir_base, s.dir_df = {}, {}
        s.dir_version, s.dir_saved = {"startup": 0, "partner": 0}, {}
        for k in DIR_FIELDS:
            s.dir_base[k] = s.dir_df[k] = dir_tidy(pd.DataFrame(), k, 0)
            s.dir_saved[k] = hashlib.sha1(dir_doc(k).encode()).hexdigest()
        s.dir_ver = 1


def dir_load_doc(kind: str, doc: dict) -> None:
    ss().dir_extra[kind] = list(doc.get("fields", []))
    dir_set(kind, pd.DataFrame(doc.get("rows", [])))


def dir_pull(force: bool = False) -> None:
    """Team workspace: load the shared directory when someone else saved a newer one."""
    S = store()
    if not S:
        return
    init_dir()
    for kind in DIR_FIELDS:
        v = S.dir_version(kind)
        unsaved = hashlib.sha1(dir_doc(kind).encode()).hexdigest() != ss().dir_saved.get(kind)
        if v > ss().dir_version.get(kind, 0) and (force or not unsaved):
            row = S.get_dir(kind)
            dir_load_doc(kind, json.loads(row["data"]))
            ss().dir_version[kind] = row["version"]
            ss().dir_saved[kind] = hashlib.sha1(dir_doc(kind).encode()).hexdigest()


def dir_autosave() -> None:
    S = store()
    if not S:
        return
    for kind in DIR_FIELDS:
        doc = dir_doc(kind)
        h = hashlib.sha1(doc.encode()).hexdigest()
        if h == ss().dir_saved.get(kind):
            continue
        v = S.save_dir(kind, doc, ss().dir_version.get(kind, 0), me())
        if v:
            ss().dir_version[kind], ss().dir_saved[kind] = v, h
        else:   # someone saved at the same moment: take theirs and say so
            row = S.get_dir(kind)
            dir_load_doc(kind, json.loads(row["data"]))
            ss().dir_version[kind] = row["version"]
            ss().dir_saved[kind] = hashlib.sha1(dir_doc(kind).encode()).hexdigest()
            ss().flash = ("warning", T("dir_conflict", who=row.get("updated_by") or "?"))
            st.rerun()


def dir_commit() -> None:
    """Make the current edits the new starting point (before the filter or the view changes)."""
    for kind in DIR_FIELDS:
        ss().dir_base[kind] = ss().dir_df[kind]
    ss().dir_ver += 1


def guess_dir(header: str, kind: str, taken: set) -> str:
    n = norm(header)
    for f in dir_cols(kind):
        if n in {norm(f), norm(dir_label(f)), norm(DIR_LABELS["ko"].get(f, f)), norm(DIR_LABELS["en"].get(f, f))} \
                and f not in taken:
            return f
    for f, words in DIR_ALIASES[kind]:
        if f not in taken and any(w in n for w in words):
            return f
    if n in {"", "no", "no.", "#", "번호", "순번", "순서", "index", "sheet"} or n.startswith(("unnamed", "column ")):
        return "(ignore)"
    return "(new)"     # keep everything: a column we do not know becomes its own field


def cb_dir_upload(kind: str):
    f = ss().get(f"dir_up_{kind}")
    if not f:
        return
    try:
        book = read_book(f)
    except Exception as e:  # noqa: BLE001
        ss().flash = ("error", T("f_read_fail", f=f.name, e=e))
        return
    if len(book) == 1:
        ss().pop(f"book_dir_{kind}", None)
        ss()[f"dir_paste_{kind}"] = book_text(book, list(book))
        return
    best = load_book(f"dir_{kind}", book, kind, f"dir_paste_{kind}")
    msg = T("f_loaded_tab", f=f.name, n=len(book), s=best)
    # a second tab that clearly holds the other list is prepared on the other side too
    other = "partner" if kind == "startup" else "startup"
    cands = [n for n in book if n != best and any(w in norm(n) for w in SHEET_HINTS[other])]
    if cands and not ss().get(f"dir_paste_{other}"):
        o = load_book(f"dir_{other}", book, other, f"dir_paste_{other}")
        msg += " " + T("f_other_tab", s=o, side=T("dir_" + other + "s"))
    ss().flash = ("info", msg)


def cb_dir_import(kind: str):
    s = ss()
    df = parse_paste(s.get(f"dir_paste_{kind}", ""))
    if df is None:
        s.flash = ("error", T("f_nothing"))
        return
    out = pd.DataFrame(index=range(len(df)))
    for i, c in enumerate(df.columns):
        f = s.get(f"dmap_{kind}_{i}_{hashlib.md5(c.encode()).hexdigest()[:6]}", "(ignore)")
        if f == "(new)":
            f = c.strip()
            if f not in dir_cols(kind):
                s.dir_extra[kind].append(f)
        if f != "(ignore)" and f not in out.columns:
            out[f] = df[c].reset_index(drop=True)
    name = DIR_NAME[kind]
    if name not in out.columns and not (kind == "startup" and "Company (KR)" in out.columns):
        s.flash = ("error", T("dir_need_name", col=dir_label(name)))
        return
    new = dir_tidy(out, kind, s.dir_ver + 1)
    old = dir_tidy(s.dir_df[kind], kind)
    if s.get(f"dir_mode_{kind}", "merge") == "merge":
        key = (lambda r: norm(r.get(name) or r.get("Company (KR)", ""))) if kind == "startup" else (lambda r: norm(r.get(name)))
        idx = {key(r): i for i, r in old.iterrows() if key(r)}
        added = updated = 0
        rows = old.to_dict("records")
        for r in new.to_dict("records"):
            k = key(r)
            if k in idx:
                cur = rows[idx[k]]
                for c in out.columns:
                    if r.get(c):
                        cur[c] = r[c]
                updated += 1
            else:
                rows.append(r)
                added += 1
        dir_set(kind, pd.DataFrame(rows))
        s.flash = ("success", T("dir_merged", a=added, u=updated))
    else:
        dir_set(kind, new)
        s.flash = ("success", T("dir_replaced", n=len(new)))
    s[f"dir_paste_{kind}"] = ""


def dir_import_box(kind: str):
    s = ss()
    st.file_uploader(T("import_file"), type=["csv", "xlsx", "xls"], key=f"dir_up_{kind}",
                     on_change=cb_dir_upload, args=(kind,))
    sheet_picker(f"dir_{kind}", f"dir_paste_{kind}")
    st.text_area(T("dir_paste"), key=f"dir_paste_{kind}", height=110,
                 placeholder="회사명\t영문명\t대표자\t이메일\n플렉스컷\tFlexcuit\t김민지\tmj@flexcuit.com")
    df = parse_paste(s.get(f"dir_paste_{kind}", ""))
    if df is None:
        return
    st.caption(T("rows_found", n=len(df)))
    opts = ["(ignore)", "(new)"] + dir_cols(kind)
    show = {"(ignore)": T("ignore"), "(new)": T("dir_new_field"), **{c: dir_label(c) for c in dir_cols(kind)}}
    taken, grid = set(), st.columns(min(3, len(df.columns)))
    for i, c in enumerate(df.columns):
        k = f"dmap_{kind}_{i}_{hashlib.md5(c.encode()).hexdigest()[:6]}"
        if k not in s:
            s[k] = guess_dir(c, kind, taken)
        if s[k] not in ("(ignore)", "(new)"):
            taken.add(s[k])
        sample = ", ".join(str(v) for v in df[c].head(2) if not is_blank(v))[:40]
        grid[i % len(grid)].selectbox(c, opts, key=k, format_func=lambda o, show=show: show.get(o, o),
                                      help=T("eg", x=sample) if sample else None)
    st.radio("mode", ["merge", "replace"], horizontal=True, key=f"dir_mode_{kind}", label_visibility="collapsed",
             format_func=lambda o: T("dir_mode_" + o))
    st.button(T("import_n", n=len(df)), key=f"dir_go_{kind}", type="primary", on_click=cb_dir_import, args=(kind,))


def cb_dir_field(kind: str, action: str):
    s = ss()
    dir_commit()
    if action == "add":
        name = (s.get(f"dir_newf_{kind}") or "").strip()
        if name and name not in dir_cols(kind):
            s.dir_extra[kind].append(name)
            dir_set(kind, s.dir_df[kind])
        s[f"dir_newf_{kind}"] = ""
    else:
        name = s.get(f"dir_delf_{kind}")
        if name in s.dir_extra[kind]:
            s.dir_extra[kind].remove(name)
            dir_set(kind, s.dir_df[kind].drop(columns=[name], errors="ignore"))


def dir_excel() -> bytes:
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    for kind, title in (("startup", "Startups"), ("partner", "Partners")):
        ws = wb.create_sheet(title)
        cols = dir_cols(kind)
        for j, c in enumerate(cols, 1):
            cell = ws.cell(1, j, dir_label(c))
            cell.fill, cell.font = _fill(C_HEAD), _f(10, True, C_HEAD_TXT)
            ws.column_dimensions[get_column_letter(j)].width = 34 if c in ("One-liner", "Notes") else 18
        for i, r in enumerate(ss().dir_df[kind].to_dict("records"), 2):
            for j, c in enumerate(cols, 1):
                ws.cell(i, j, r.get(c, ""))
        ws.freeze_panes = "B2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def dir_display_name(kind: str, r: dict) -> str:
    if kind == "startup":
        return r.get("Company (EN)") or r.get("Company (KR)") or "?"
    return r.get("Organisation") or "?"


def dir_history(kind: str, name: str) -> list[dict]:
    """Meetings this company has in the open event (and, in the team workspace, in every saved event)."""
    out, key = [], norm(name)
    who = "startup" if kind == "startup" else "host"
    other = "host" if who == "startup" else "startup"
    r = ss().get("result")
    if r and r.get("meetings"):
        cfg = EventConfig(**r["cfg"])
        for m in r["meetings"]:
            if norm(m[who]) == key:
                out.append({"event": cfg.name, "with": m[other], "when": f"{cfg.day_short(m['day'])} {fmt(m['start'])}",
                            "status": st_label(eff_status(Meeting(**m), cfg, ss().statuses, now_for(cfg)))})
    S = store()
    if S:
        cur = ss().get("event_id")
        for e in S.list_events():
            if e["id"] == cur:
                continue
            row = S.get(e["id"])
            try:
                d = json.loads(row["data"])
            except Exception:  # noqa: BLE001
                continue
            for m in (d.get("result") or {}).get("meetings", []) or []:
                if norm(m.get(who)) == key:
                    out.append({"event": e["name"], "with": m.get(other), "when": fmt(m.get("start", 0)), "status": ""})
    return out


def dir_profile(kind: str):
    df = ss().dir_df[kind]
    if df.empty:
        return
    names = [dir_display_name(kind, r) for r in df.to_dict("records")]
    pick = st.selectbox(T("dir_profile"), ["—"] + names, key=f"dir_pick_{kind}")
    if pick == "—":
        return
    r = df.iloc[names.index(pick)].to_dict()
    with st.container(border=True):
        a, b = st.columns([3, 2])
        if kind == "startup":
            a.markdown(f"### {r.get('Company (EN)') or r.get('Company (KR)')}")
            if r.get("Company (KR)") and r.get("Company (EN)"):
                a.caption(r["Company (KR)"])
            ceo = " / ".join(x for x in (r.get("CEO"), r.get("CEO (EN)")) if x)
            if ceo:
                a.write(f"**{dir_label('CEO')}** {ceo}")
            if r.get("One-liner"):
                a.write(r["One-liner"])
            tags = [t.strip() for t in (r.get("Tags") or "").split(",") if t.strip()]
        else:
            a.markdown(f"### {r.get('Organisation')}")
            a.caption(" · ".join(x for x in (r.get("Contact"), r.get("Title"), r.get("Type"), r.get("Country")) if x))
            tags = [f"{dir_label('Prefers')} {r['Prefers']}"] if r.get("Prefers") else []
            if r.get("Only meets"):
                tags.append(f"{dir_label('Only meets')} {r['Only meets']}")
            if r.get("Excludes"):
                tags.append(f"{dir_label('Excludes')} {r['Excludes']}")
        if tags:
            a.markdown(" ".join(f":blue-badge[{t}]" for t in tags))
        shown = {"Company (EN)", "Company (KR)", "CEO", "CEO (EN)", "One-liner", "Tags", "Organisation", "Contact",
                 "Title", "Type", "Country", "Prefers", "Only meets", "Excludes", "_id"}
        for c in dir_cols(kind):
            v = r.get(c)
            if not v or c in shown:
                continue
            if c in ("Website", "Deck") or str(v).startswith("http"):
                link = v if str(v).startswith("http") else f"https://{v}"
                b.markdown(f"**{dir_label(c)}** [{v}]({link})")
            else:
                b.markdown(f"**{dir_label(c)}** {v}")
        hist = dir_history(kind, pick)
        st.markdown(f"**{T('dir_history')}** · {len(hist)}")
        if hist:
            st.dataframe(pd.DataFrame(hist).rename(columns={"event": T("dir_h_event"), "with": T("dir_h_with"),
                                                            "when": T("dir_h_when"), "status": T("dir_h_status")}),
                         hide_index=True, width="stretch")
        else:
            st.caption(T("dir_no_history"))


def dir_to_event_row(kind: str, r: dict) -> dict:
    if kind == "startup":
        return {"Name": dir_display_name(kind, r), "Presenter": r.get("Presenter") or r.get("CEO (EN)") or r.get("CEO") or "",
                "Domain": r.get("Tags", "")}
    return {"Name": r.get("Organisation", ""), "Contact": r.get("Contact", ""),
            "Type": r.get("Type") if r.get("Type") in HOST_TYPES else "Investor",
            "Preferred Domains": r.get("Prefers", ""), "Only Domains": r.get("Only meets", ""),
            "Excluded Domains": r.get("Excludes", "")}


def cb_dir_to_event(role: str, mode: str):
    s = ss()
    kind = "startup" if role == "startup" else "partner"
    key_base, key_df = ("s_base", "s_df") if role == "startup" else ("h_base", "h_df")
    cur = s[key_df].copy()
    rows = s.dir_df[kind].to_dict("records")
    by_name = {}
    for r in rows:
        for n in (dir_display_name(kind, r), r.get("Company (KR)", "")):
            if n:
                by_name[norm(n)] = r
    if mode == "add":
        picked = s.get(f"dimp_{role}") or []
        new = [dir_to_event_row(kind, by_name[norm(n)]) for n in picked if norm(n) in by_name]
        have = {norm(v) for v in cur["Name"]}
        new = [r for r in new if norm(r["Name"]) not in have]
        out = pd.concat([cur, clean_frame(pd.DataFrame(new), role)], ignore_index=True) if new else cur
        s.flash = ("success", T("dir_added_to_event", n=len(new)))
    else:
        n_up = 0
        for i, name in enumerate(cur["Name"]):
            r = by_name.get(norm(name))
            if not r:
                continue
            for c, v in dir_to_event_row(kind, r).items():
                if c != "Name" and v and c in cur.columns:
                    cur.at[i, c] = v
            n_up += 1
        out = cur
        s.flash = ("success", T("dir_refreshed", n=n_up))
    s[key_base] = s[key_df] = clean_frame(out, role)
    s.ver += 1
    s[f"dimp_{role}"] = []


def dir_pick_popover(role: str, where):
    """In an event: bring companies in from the directory, or refresh their details from it."""
    kind = "startup" if role == "startup" else "partner"
    with where.popover(T("dir_from"), width="stretch"):
        dir_pull()
        df = ss().dir_df[kind]
        if df.empty:
            st.caption(T("dir_empty_hint"))
            if PAGES.get("companies"):
                st.page_link(PAGES["companies"], label=T("dir_go"))
            return
        cur = ss().s_df if role == "startup" else ss().h_df
        have = {norm(v) for v in cur["Name"]}
        names = [dir_display_name(kind, r) for r in df.to_dict("records")]
        missing = [n for n in names if norm(n) not in have]
        st.caption(T("dir_from_cap", n=len(missing), t=len(names)))
        st.multiselect(T("dir_pick"), names, default=missing, key=f"dimp_{role}", label_visibility="collapsed")
        a, b = st.columns(2)
        a.button(T("dir_add_btn"), key=f"dimp_go_{role}", type="primary", on_click=cb_dir_to_event, args=(role, "add"))
        b.button(T("dir_refresh_btn"), key=f"dimp_rf_{role}", on_click=cb_dir_to_event, args=(role, "refresh"),
                 help=T("dir_refresh_help"))


def matching_stats(event_ids: list) -> dict:
    """Meetings of the chosen business matching events, counted by live status, overall and per startup."""
    out = {"events": [], "total": 0, "by": {}, "partners": set(), "startups": {}}
    S = store()

    def cfg_of(res: dict, data: dict | None) -> EventConfig | None:
        """The event's settings, enough to tell from the clock which meetings are over."""
        try:
            if res.get("cfg"):
                return EventConfig(**res["cfg"])
            if data:
                return EventConfig(days=data.get("days") or [{"label": "Day 1", "date": ""}],
                                   tz=(data.get("config") or {}).get("tz") or "Asia/Singapore")
        except Exception:  # noqa: BLE001
            pass
        return None

    def add(name: str, cfg: EventConfig | None, meetings: list, statuses: dict):
        if not meetings:
            return
        out["events"].append(name)
        for m in meetings:
            if not isinstance(m, dict) or not m.get("startup") or not m.get("host"):
                continue
            mid = f"{m['startup']}|{m['host']}"
            stt = None
            if cfg:
                try:
                    stt = eff_status(Meeting(**m), cfg, statuses, now_for(cfg))
                except Exception:  # noqa: BLE001  (a meeting saved by an older version)
                    stt = None
            if stt is None:
                stt = statuses.get(mid) if statuses.get(mid) in STATUSES else "Scheduled"
            out["total"] += 1
            out["by"][stt] = out["by"].get(stt, 0) + 1
            out["partners"].add(m["host"])
            k = norm(m["startup"])
            e = out["startups"].setdefault(k, {"name": m["startup"], "total": 0, "done": 0, "noshow": 0, "by": {},
                                              "partners": []})
            e["total"] += 1
            e["by"][stt] = e["by"].get(stt, 0) + 1
            e["done"] += stt == "Done"
            e["noshow"] += stt == "No-Show"
            if stt == "Done":
                e["partners"].append(m["host"])

    for eid in event_ids or []:
        if eid == "current" or (S and eid == ss().get("event_id")):
            r = ss().get("result") or {}
            try:
                cfg = current_cfg() if "days_df" in ss() else cfg_of(r, None)
            except Exception:  # noqa: BLE001
                cfg = cfg_of(r, None)
            add(ss().get("ev_name") or "Business matching", cfg, r.get("meetings") or [], ss().get("statuses", {}))
        elif S:
            row = S.get(eid)
            if not row or row.get("archived"):
                continue
            try:
                d = json.loads(row["data"])
            except Exception:  # noqa: BLE001
                continue
            res = d.get("result") or {}
            items = S.items(eid)
            sts = {i["item_id"]: i["value"] for i in items if i["kind"] == "status" and i["value"] in STATUSES}
            add(row.get("name") or "?", cfg_of(res, d), res.get("meetings") or [], sts)
    return out


def event_choices() -> list:
    S = store()
    if S:
        return [(e["id"], e["name"] or "Untitled") for e in S.list_events()]
    r = ss().get("result")
    return [("current", ss().get("ev_name") or "Business matching")] if r and r.get("meetings") else []


def dir_rows(kind: str) -> list[dict]:
    init_dir()
    dir_pull()
    return ss().dir_df[kind].to_dict("records")


def dir_update(kind: str, name: str, fields: dict) -> None:
    """Fill in company details from another page (e.g. an intake summary), then save the directory."""
    init_dir()
    df = ss().dir_df[kind].copy()
    hit = [i for i, r in df.iterrows() if norm(dir_display_name(kind, r.to_dict())) == norm(name)]
    if not hit:
        row = {DIR_NAME[kind]: name, **fields}
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    else:
        for c, v in fields.items():
            if c in df.columns:
                df.at[hit[0], c] = v
    dir_set(kind, df)
    dir_autosave()


def companies_page():
    s = ss()
    init_dir()
    dir_pull()
    box = st.container()
    show_flash(box, s.pop("flash", None))
    st.subheader(T("dir_title"), anchor=False)
    st.caption(T("dir_cap"))
    counts = {k: len(s.dir_df[k]) for k in DIR_FIELDS}
    s.setdefault("dir_kind", "startup")
    kind = st.segmented_control("kind", list(DIR_FIELDS), key="dir_kind", on_change=dir_commit,
                                format_func=lambda k: f"{T('dir_' + k + 's')} {counts[k]}",
                                label_visibility="collapsed") or "startup"
    t1, t2, t3, t4 = st.columns([3, 1, 1, 1], vertical_alignment="bottom")
    q = t1.text_input(T("dir_search"), key=f"dir_q_{kind}", on_change=dir_commit, placeholder=T("dir_search_ph"),
                      label_visibility="collapsed")
    with t2.popover(T("dir_import"), width="stretch"):
        dir_import_box(kind)
    t3.download_button(T("dir_export"), dir_excel(), "companies.xlsx", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with t4.popover(T("dir_fields"), width="stretch"):
        st.text_input(T("dir_new_field_name"), key=f"dir_newf_{kind}", placeholder=T("dir_new_field_ph"))
        st.button(T("dir_add_field"), key=f"dir_addf_{kind}", on_click=cb_dir_field, args=(kind, "add"))
        if s.dir_extra[kind]:
            st.selectbox(T("dir_remove_field"), s.dir_extra[kind], key=f"dir_delf_{kind}")
            st.button(T("dir_remove_btn"), key=f"dir_rmf_{kind}", on_click=cb_dir_field, args=(kind, "del"))

    base = s.dir_base[kind]
    ql = (q or "").strip().lower()
    view = base[pd.Series([any(ql in str(v).lower() for v in r) for r in base.itertuples(index=False)],
                          index=base.index, dtype=bool)] if ql else base
    cfg = {"_id": None}
    for c in dir_cols(kind):
        if kind == "partner" and c == "Type":
            cfg[c] = st.column_config.SelectboxColumn(dir_label(c), options=[""] + list(HOST_TYPES))
        elif c in ("One-liner", "Notes"):
            cfg[c] = st.column_config.TextColumn(dir_label(c), width="large")
        else:
            cfg[c] = st.column_config.TextColumn(dir_label(c), help=T(DIR_HELP[c]) if c in DIR_HELP else None,
                                                 width="medium" if c == DIR_NAME[kind] else None)
    edited = st.data_editor(view.reset_index(drop=True), key=f"dir_{kind}_{s.dir_ver}", num_rows="dynamic",
                            hide_index=True, column_config=cfg, width="stretch",
                            height=min(640, 38 + 35 * (len(view) + 2)))
    if ql:
        merged = pd.concat([base[~base["_id"].isin(view["_id"])], edited], ignore_index=True)
    else:
        merged = edited
    s.dir_df[kind] = dir_tidy(merged, kind)
    if ql:
        st.caption(T("dir_filtered", n=len(view), t=len(base)))
    dir_profile(kind)
    dir_autosave()


KEEP_KEYS = {"tz", "tl", "tables", "t_start", "t_end", "t_pol_s", "t_pol_e", "t_lunch_s", "t_lunch_e", "sessions_txt",
             "pw", "pol_on", "pol_days", "pol_cap", "pe", "packing_mode", "lunch_on", "keep_live", "interval", "hf",
             "ev_name", "dir_kind", "day_sel", "buffer_s", "buffer_m", "lang"}
KEEP_PREFIXES = ("lbl_", "nm_", "dir_q_", "pmode_", "dir_mode_", "paste_", "dir_paste_", "map_", "dmap_", "dir_pick_")


def keep_widget_state() -> None:
    """Streamlit forgets the value of a widget whose page is not on screen. Re-saving the values at the start of
    every run keeps the event settings when you hop between Companies and Business matching."""
    s = ss()
    for k in list(s.keys()):
        if isinstance(k, str) and (k in KEEP_KEYS or k.startswith(KEEP_PREFIXES)):
            try:
                s[k] = s[k]
            except Exception:  # noqa: BLE001
                pass


def page_switched() -> None:
    """Tables are edited on top of a starting copy. When you leave a page, make the edits the new start,
    so nothing typed in a table is lost when you come back."""
    s = ss()
    if "s_df" in s:
        s.days_base, s.s_base, s.h_base = s.days_df, s.s_df, s.h_df
        s.ver = s.get("ver", 0) + 1
    if "dir_df" in s:
        dir_commit()
    programs.on_page_switch()


def local_backup_restore() -> None:
    """Personal mode: bring back the last saved copy from this browser, once per visit."""
    if store() or ss().get("_restored"):
        return
    val = ss().get("bm_local")
    if not isinstance(val, dict):
        return
    ss()._restored = True
    raw = val.get("loaded") or ""
    if not raw:
        return
    try:
        data = json.loads(raw)
        load_project(data)
        init_dir()
        for k, doc in (data.get("directory") or {}).items():
            if k in DIR_FIELDS:
                dir_load_doc(k, doc)
        if data.get("program_docs"):
            ss().local_docs = data["program_docs"]
        ss().flash = ("success", T("restored_local"))
    except Exception:  # noqa: BLE001
        pass


def local_backup_save() -> None:
    """Personal mode: keep a copy of the event and the directory in this browser after every change."""
    if store():
        return
    save = None
    if ss().get("_restored"):
        try:
            doc = project_json()
            h = hashlib.sha1(doc.encode()).hexdigest()
            if h != ss().get("_local_saved"):
                save, ss()._local_saved = doc, h
        except Exception:  # noqa: BLE001
            save = None
    LOCAL(key="bm_local", data={"key": "bizmatch-v1", "save": save}, default=None)


def main():
    programs.begin_run()   # first, before anything is drawn: keeps what was typed right before a click that leaves the page
    st.set_page_config(page_title=APP_TITLE, page_icon=str(BRAND / "mark.png"), layout="wide")
    keep_widget_state()
    if "lang" not in ss():   # the link's ?lang=, else the browser's language
        q = st.query_params.get("lang")
        if q not in LANGS:
            q = "ko" if str(getattr(st.context, "locale", "") or "").lower().startswith("ko") else "en"
        ss().lang = q
    init_state()
    st.logo(str(BRAND / "logo.png"), size="large")
    st.markdown("<style>[data-testid='stMainBlockContainer'],.block-container{max-width:100%;"
                "padding-left:2rem;padding-right:2rem;padding-top:4.75rem}</style>", unsafe_allow_html=True)
    ready = gate_ready()
    programs.API.update(store=store, me=me, norm=norm, dir_rows=dir_rows, dir_name=dir_display_name,
                        dir_update=dir_update, matching_stats=matching_stats, event_choices=event_choices, pages=PAGES)
    PAGES["programs"] = st.Page(programs.page, title=T("nav_programs"), icon=":material/view_timeline:",
                                url_path="programs", default=True)
    PAGES["companies"] = st.Page(companies_page, title=T("nav_companies"), icon=":material/apartment:",
                                 url_path="companies")
    PAGES["matching"] = st.Page(matching_page, title=T("nav_matching"), icon=":material/handshake:", url_path="matching")
    nav = st.navigation(list(PAGES.values()), position="top" if ready else "hidden")   # keeps /matching links after login
    if not ready:
        gate()
        footer()
        return
    init_dir()
    local_backup_restore()
    here = getattr(nav, "url_path", "") or "home"
    if ss().get("_page") not in (None, here):
        page_switched()
    ss()._page = here
    nav.run()
    footer()
    local_backup_save()


def show_flash(box, flash) -> None:
    """Confirmations slide in as a small toast and leave by themselves; anything to read or act on stays."""
    if not flash:
        return
    kind, msg = flash
    if kind == "success":
        st.toast(msg, icon=":material/check_circle:")
    else:
        getattr(box, kind)(msg)


def footer():
    """A quiet last line: what the app is, and the language (it starts in the browser's language)."""
    st.write("")
    st.divider()
    c1, c2 = st.columns([6, 1.5], vertical_alignment="center")
    c1.caption(T("app_caption"))
    c2.segmented_control(T("language"), list(LANGS), format_func=LANGS.get, key="lang", on_change=cb_lang,
                         required=True, label_visibility="collapsed", width="stretch")


def matching_page():
    S = store()
    if S:
        pend = ss().pop("pending_reload", None)
        if pend and ss().get("event_id"):
            open_event(ss().event_id)
            msg = T("ws_conflict" if pend[0] == "conflict" else "ws_remote", who=pend[1] or "?")
            if pend[0] == "conflict":
                ss().flash = ("warning", msg)
            elif pend[1] != me():
                st.toast(msg)
        if not ss().get("event_id"):
            evs = S.list_events()
            if evs:
                open_event(evs[0]["id"])
            else:
                set_frames(*sample_frames())
                ss().result, ss().statuses, ss().pol_notes = None, {}, {}
                d = project_dict(live=False)
                eid, _ = S.create(d["config"]["name"], canon(d), me())
                open_event(eid)
    workspace_bar()
    box = st.container()   # a fixed slot, so a message never moves the tabs (moving them resets the open tab)
    show_flash(box, ss().pop("flash", None))
    t1, t2, t3, t4 = st.tabs([T("tab1"), T("tab2"), T("tab3"), T("tab4")])
    with t1:
        setup_tab()
    with t2:
        generate_tab()
    with t3:
        results_tab()
    with t4:
        summary_tab()
    if S:
        autosave()
        live_sync()


if __name__ == "__main__":
    main()
