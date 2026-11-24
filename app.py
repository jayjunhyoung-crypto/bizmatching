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
APP_TITLE = "BizMatch Scheduler"
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
                "x", "dropped", "withdrawn")
HOST_COLS = ["Name", "Contact", "Type", "Days", "Available From", "Available Until",
             "Blocked Times", "Excluded Domains", "Max Meetings"]
TIMEZONES = ["Asia/Singapore", "Asia/Seoul", "Asia/Tokyo", "Asia/Hong_Kong", "Asia/Shanghai",
             "Asia/Jakarta", "Asia/Bangkok", "Asia/Kolkata", "Asia/Dubai", "Europe/London",
             "Europe/Berlin", "America/New_York", "America/Los_Angeles", "UTC"]
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
    if is_blank(v) or norm(v) in ("all", "any", "both", "every day"):
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
        m = re.fullmatch(r"(?:day\s*|d)?(\d+)", t)
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
    excluded: frozenset = frozenset()  # hosts
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


def compatible(s: Party, h: Party) -> bool:
    return not (s.domains & h.excluded)


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

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(settings.time_limit)
    solver.parameters.num_workers = int(settings.workers)
    code = solver.Solve(model)
    name = solver.StatusName(code)
    wall = _time.time() - t0

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
            issues.append(f"Excluded domain: {m.startup} with {m.host}")
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
C_NAVY, C_COB, C_INK, C_MUT = "18255C", "1240E0", "111827", "5A6276"
C_LINE, C_ZONE, C_FAINT, C_BAND = "E1E6EF", "EEF2FF", "B8C0CF", "F7F9FC"
ST_STYLE = {"Scheduled": ("EEF0F4", "4A5168"), "Ongoing": ("1240E0", "FFFFFF"),
            "Done": ("DDF2E3", "1E7B34"), "No-Show": ("FDE2E6", "B4132D")}
CELL_TINT = {"Ongoing": "EAF0FF", "Done": "F4F5F7", "No-Show": "FFF0F2"}
NAME_COLORS = ["1240E0", "C2255C", "0B7A75", "B45309", "6D28D9", "0E7490", "A21CAF",
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
        c.fill, c.font = _fill(C_NAVY), _f(10, True, "FFFFFF")
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


# Host colours avoid pink and yellow, which are reserved for pitch polishing
HOST_PALETTE = ["D9EAD3", "CFE2F3", "D9D2E9", "76A5AF", "134F5C", "F6B26B", "8E7CC3", "E69138",
                "B6D7A8", "351C75", "93C47D", "A4C2F4", "B4A7D6", "45818E", "6FA8DC", "6AA84F",
                "3D85C6", "B45F06", "A2C4C9", "38761D"]
POLISH_FILL = {1: "F4CCCC", 2: "FFE599"}
POLISH_FILL_MORE = "D0E0E3"
C_HEAD, C_GRID, C_BREAK = "2B3A67", "D9DDE5", "F1F2F5"


def text_on(hex_fill: str) -> str:
    r, g, b = (int(hex_fill[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return "FFFFFF" if 0.2126 * r + 0.7152 * g + 0.0722 * b < 0.5 else "1F2430"


def host_colors(hosts) -> dict:
    """Each host (organisation) keeps one colour everywhere."""
    return {h.name: HOST_PALETTE[i % len(HOST_PALETTE)] for i, h in enumerate(hosts)}


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
            c.fill, c.font, c.alignment, c.border = _fill(C_HEAD), _f(10, True, "FFFFFF"), center, box
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
    for s_, (bg, fg) in {**ST_STYLE, "Scheduled": ("E3EAFB", "1240E0")}.items():
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
            c.fill, c.font, c.alignment, c.border = _fill(C_HEAD), _f(10, True, "FFFFFF"), center, box
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
        ("Voltra Energy", "Climate", "", "", "", "", True),
        ("MediKey", "HealthTech", "", "", "", "", False),
        ("Kanto Robotics", "Hardware", "", "", "", "", False),
        ("Seedline Agritech", "AgriTech", "", "", "", "", False),
        ("Parcelo", "Logistics", "", "", "", "", False),
        ("FinchPay", "Fintech", "", "", "", "", True),
        ("Lumen Health", "HealthTech", "", "", "", "", False),
        ("Tidewave Marine", "Climate", "", "", "", "", False),
        ("Glowlab Cosmetics", "K-Beauty", "", "", "", "", False),
        ("Aster Skin", "K-Beauty", "Day 2", "", "", "", False),
        ("Railmind", "Mobility", "", "", "16:00", "", False),
        ("Quanta Ledger", "Fintech", "", "", "", "", False),
        ("Brightcart", "SaaS", "", "", "", "", False),
        ("Nimbus Grid", "Climate, Hardware", "", "", "", "", False),
        ("Coral AI", "SaaS", "", "", "", "", True),
        ("Haneul Foods", "FoodTech", "", "", "", "Day 1 13:00-14:00", False),
        ("Pathwise", "EdTech", "Not here", "", "", "", False),
        ("Ondo Mobility", "Mobility", "Day 1", "", "", "", False),
        ("Silo Sense", "AgriTech, Hardware", "", "11:00", "", "", False),
        ("Kite Insure", "Fintech", "", "", "", "", False),
    ]
    startups = pd.DataFrame(s, columns=["Name", "Domain", "Days", "Available From", "Available Until",
                                        "Blocked Times", "Priority"])
    startups["Presenter"] = ["Aiden", "Hana", "Kenji", "Ravi", "Mia", "Leon", "Sora", "Noah", "Yuna", "Clara",
                             "Arun", "Jiwoo", "Ethan", "Min & Ivy", "Zoe", "Daniel", "Lily", "Ken", "Mateo", "Ava"]
    startups["Pitch Polishing"] = True
    startups["Max Meetings"] = None
    h = [
        ("Sunrise Ventures", "Grace Lim", "Investor", "", "", "", "", "K-Beauty", None),
        ("Meridian Capital", "Daniel Ong", "Investor", "Day 2", "15:00", "17:00", "", "", None),
        ("Halcyon Logistics", "Priya Nair", "Corporate", "", "", "", "", "K-Beauty, EdTech", None),
        ("Orchid Bank", "Marcus Tan", "Corporate", "Day 1", "", "", "", "K-Beauty, Hardware", None),
        ("Harbor Innovation Agency", "Mei Ling", "Agency", "", "", "", "Day 2 10:00-11:00", "", None),
        ("Lumen Corporate Ventures", "Jason Koh", "Investor", "", "", "", "", "", 16),
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


SPLIT_HOST = "Contact - Organisation (split)"
GUESS = [  # (words in the header, field)
    (("present", "speaker", "pitcher", "founder"), "Presenter"),
    (("contact", "attendee", "representative", "person"), "Contact"),
    (("exclud", "avoid", "not interested"), "Excluded Domains"),
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
    cand = ALIASES.get(n) if ALIASES.get(n) in cols else next((c for c in cols if c.lower() == n), None)
    if not cand:
        for words, f in GUESS:
            if f in cols and any(w in n for w in words):
                cand = f
                break
    return cand if cand and cand not in taken else "(ignore)"


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


def project_json() -> str:
    s = ss()
    cfg = current_cfg()
    return json.dumps({
        "app": APP_TITLE, "version": 1,
        "config": {"name": cfg.name, "start": fmt(cfg.start), "end": fmt(cfg.end),
                   "interval": cfg.interval, "buffer": cfg.buffer, "lunch_on": bool(s.lunch_on),
                   "lunch_start": s.t_lunch_s.strftime("%H:%M"), "lunch_end": s.t_lunch_e.strftime("%H:%M"),
                   "tables": cfg.tables, "tz": cfg.tz, "sessions": s.sessions_txt,
                   "pol_on": bool(s.pol_on), "pol_days": list(s.pol_days),
                   "pol_start": s.t_pol_s.strftime("%H:%M"), "pol_end": s.t_pol_e.strftime("%H:%M"),
                   "pol_cap": int(s.pol_cap), "names": cfg.names},
        "days": cfg.days,
        "startups": s.s_df.to_dict("records"),
        "hosts": s.h_df.to_dict("records"),
        "result": s.result and {k: v for k, v in s.result.items() if k in ("meetings", "polish", "parked", "packing", "status", "hash")},
        "statuses": s.statuses,
        "pol_notes": s.get("pol_notes", {}),
    }, default=str, indent=1)


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
        df = read_upload(f)
    except Exception as e:  # noqa: BLE001
        ss().flash = ("error", f"Could not read {f.name}: {e}")
        return
    ss()[f"paste_{role}"] = df.to_csv(sep="\t", index=False)
    ss().flash = ("info", f"Loaded {f.name}. Check which column is which below, then press Import.")


def cb_paste(role: str):
    s = ss()
    df = parse_paste(s.get(f"paste_{role}", ""))
    if df is None:
        s.flash = ("error", "Nothing to import. Paste cells including the header row.")
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
        s.flash = ("error", "Pick which column holds the startup or organisation name.")
        return
    new = clean_frame(out, role)
    new = new[[not is_blank(v) for v in new["Name"]]].reset_index(drop=True)
    key_base, key_df = ("s_base", "s_df") if role == "startup" else ("h_base", "h_df")
    if s.get(f"pmode_{role}", "Replace the list") == "Add to the list":
        old = s[key_df]
        names = {norm(v) for v in new["Name"]}
        new = pd.concat([old[[norm(v) not in names for v in old["Name"]]], new], ignore_index=True)
    s[key_base] = s[key_df] = new
    s.ver += 1
    s[f"paste_{role}"] = ""
    s.flash = ("success", f"Imported {len(new)} {'startups' if role == 'startup' else 'hosts'}.")


def paste_box(role: str):
    s = ss()
    raw = s.get(f"paste_{role}", "")
    what = "startups" if role == "startup" else "hosts"
    with st.expander(f"Paste {what} from Google Sheets or Excel (your own column names are fine)", expanded=bool(raw)):
        st.text_area("Select the cells in your sheet including the header row, copy, and paste here",
                     key=f"paste_{role}", height=140,
                     placeholder="Startup\tPresenter\tSector\nFlexcuit\tMijin\tHardware")
        df = parse_paste(s.get(f"paste_{role}", ""))
        if df is None:
            return
        cols = STARTUP_COLS if role == "startup" else HOST_COLS
        opts = ["(ignore)"] + cols + ([SPLIT_HOST] if role == "host" else [])
        st.caption(f"{len(df)} rows found. Tell the app what each of your columns means:")
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
            grid[i % len(grid)].selectbox(f"{c}", opts, key=k, help=f"e.g. {sample}" if sample else None)
        a, b = st.columns([2, 1])
        a.radio("Import mode", ["Replace the list", "Add to the list"], horizontal=True, key=f"pmode_{role}",
                label_visibility="collapsed")
        b.button(f"Import {len(df)} rows", key=f"pbtn_{role}", type="primary", on_click=cb_paste, args=(role,))


def cb_project():
    f = ss().get("up_project")
    if not f:
        return
    try:
        data = json.loads(f.getvalue().decode("utf-8"))
        days = pd.DataFrame({"Label": [d["label"] for d in data["days"]],
                             "Date": [date.fromisoformat(d["date"]) if d.get("date") else None
                                      for d in data["days"]]})
        set_frames(days, clean_frame(pd.DataFrame(data.get("startups", [])), "startup"),
                   clean_frame(pd.DataFrame(data.get("hosts", [])), "host"), data.get("config"))
        res = data.get("result")
        if res and res.get("meetings"):
            cfg = current_cfg()
            res["cfg"] = asdict(cfg)
            res["s_df"], res["h_df"] = ss().s_df, ss().h_df
            res["hash"] = frames_hash(cfg, ss().s_df, ss().h_df)
            res["id"] = int(_time.time())
            res.setdefault("message", "Loaded from project file")
            res.setdefault("wall", 0.0)
            res.setdefault("issues", [])
            ss().result = res
        else:
            ss().result = None
        ss().statuses = data.get("statuses", {}) or {}
        ss().pol_notes = data.get("pol_notes", {}) or {}
        ss().flash = ("success", f"Loaded project {f.name}")
    except Exception as e:  # noqa: BLE001
        ss().flash = ("error", f"Could not load project: {e}")


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
            rows.append({"kind": "lunch", "label": "Lunch break", "time": row["time"].replace("–", " – ")})
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
    return {
        "version": r.get("version") or f"py-{r['id']}",
        "day": day, "tables": cfg.tables, "tz": cfg.tz, "date": str(cfg.day_date(day) or ""),
        "rows": rows, "sessions": cfg.session_list(),
        "startups": [{"name": x.name, "priority": x.priority, "presenter": x.presenter,
                      "avail": [z.k for z in ds if x.available(z)]}
                     for x in startups if not x.absent],
        "hosts": [{"name": h.name, "contact": h.contact, "color": hcol[h.name], "text": text_on(hcol[h.name]),
                   "avail": [z.k for z in ds if h.available(z)]} for h in hosts],
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
        "notes": attendance_notes(cfg, startups), "maxHeight": 880,
        "names": {k_: cfg.nm(k_) for k_ in DEFAULT_NAMES},
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


def setup_tab():
    s = ss()
    st.subheader("Event")
    c1, c2, c3 = st.columns([1.2, 1, 1])
    with c1:
        st.text_input("Event name", key="ev_name")
        s.days_df = st.data_editor(
            s.days_base, key=f"days_{s.ver}", num_rows="dynamic", hide_index=True,
            column_config={"Label": st.column_config.TextColumn("Day", required=True),
                           "Date": st.column_config.DateColumn("Date", format="ddd D MMM YYYY",
                                                               help="Used for live status")})
    with c2:
        a, b = st.columns(2)
        a.time_input("Day starts", key="t_start", step=timedelta(minutes=5))
        b.time_input("Day ends", key="t_end", step=timedelta(minutes=5))
        a.number_input("Meeting length (min)", 5, 180, step=1, key="interval")
        g1, g2 = b.columns(2)
        g1.number_input("Gap (min)", 0, 60, step=1, key="buffer_m", help="Break between two meetings")
        g2.number_input("+ sec", 0, 50, step=10, key="buffer_s")
        st.number_input("Tables (concurrent meetings)", 1, 200, step=1, key="tables")
    with c3:
        st.checkbox("Lunch break", key="lunch_on")
        a, b = st.columns(2)
        a.time_input("Lunch from", key="t_lunch_s", step=timedelta(minutes=5), disabled=not s.lunch_on)
        b.time_input("Lunch until", key="t_lunch_e", step=timedelta(minutes=5), disabled=not s.lunch_on)
        st.selectbox("Time zone", TIMEZONES, key="tz", help="Drives the live Ongoing and Done status")

    cfg = current_cfg()
    slots = build_slots(cfg)
    per_day = len(slots) // max(1, len(cfg.days))
    if slots:
        times = " ".join(fmt(x.start) for x in slots if x.day == 0)
        st.caption(f"{per_day} slots per day × {len(cfg.days)} day(s) = {len(slots)} slots. Day 1 starts at {times}")
    else:
        st.warning("No meeting slot fits. Check the day start, end and meeting length.")

    st.subheader("Sessions and pitch polishing")
    c1, c2 = st.columns([1, 1.4])
    with c1:
        st.text_input("Sessions", key="sessions_txt",
                      help="Groups the timeline like your sheet. Leave blank for before and after lunch")
        if s.get("sess_err"):
            st.error(s.sess_err)
        else:
            st.caption(" · ".join(f"{x['name']} {fmt(x['start'])}–{fmt(x['end'])}" for x in cfg.session_list()))
    with c2:
        st.checkbox("Run pitch polishing alongside the meetings", key="pol_on",
                    help="Each startup gets one polishing slot per selected day and has no meeting in that slot")
        s.pol_days = [x for x in s.get("pol_days", []) if x in cfg.labels]
        a, b, c, d = st.columns([1.4, 1, 1, 0.8])
        a.multiselect("Days", cfg.labels, key="pol_days", disabled=not s.pol_on)
        b.time_input("From", key="t_pol_s", step=timedelta(minutes=5), disabled=not s.pol_on)
        c.time_input("Until", key="t_pol_e", step=timedelta(minutes=5), disabled=not s.pol_on)
        d.number_input("Coaches", 1, 20, key="pol_cap", disabled=not s.pol_on,
                       help="How many startups can be polished at the same time")

    with st.expander("Column names in the schedule"):
        st.caption("Rename the headings used on the board and in the Excel file, e.g. Table → Booth")
        cols_ = st.columns(len(DEFAULT_NAMES))
        for c_, (k_, v_) in zip(cols_, DEFAULT_NAMES.items()):
            c_.text_input(v_.title() if v_.isupper() else v_, key=f"nm_{k_}")

    st.divider()
    st.subheader("Startups")
    u1, u2 = st.columns([3, 1])
    u1.file_uploader("Import startups (CSV or Excel)", type=["csv", "xlsx", "xls"], key="up_startup",
                     on_change=cb_upload, args=("startup",), label_visibility="collapsed")
    u2.download_button("Startup template", sample_frames()[1].to_csv(index=False).encode("utf-8-sig"),
                       "startups_template.csv", "text/csv")
    paste_box("startup")
    s.s_df = st.data_editor(
        s.s_base, key=f"s_{s.ver}", num_rows="dynamic", hide_index=True,
        column_config={
            "Name": st.column_config.TextColumn(required=True, width="medium"),
            "Domain": st.column_config.TextColumn(help="One or more, comma separated. Hosts can exclude domains"),
            "Days": st.column_config.TextColumn(
                help="Blank = all days. Day 2 or 1,2 for some days. Not here = not attending"),
            "Available From": st.column_config.TextColumn(help="HH:MM. Blank = day start", validate=TIME_VALID),
            "Available Until": st.column_config.TextColumn(
                help="HH:MM cutoff. No meeting may end after this", validate=TIME_VALID),
            "Blocked Times": st.column_config.TextColumn(
                help="Restricted windows, e.g. Day 1 14:00-15:00; 16:20-16:40", width="medium"),
            "Priority": st.column_config.CheckboxColumn(help="Gets more meetings when slots are scarce"),
            "Presenter": st.column_config.TextColumn(help="Who pitches, e.g. Jay & Serena"),
            "Pitch Polishing": st.column_config.CheckboxColumn(default=True,
                                                               help="Gets a pitch polishing slot each polishing day"),
            "Max Meetings": st.column_config.NumberColumn(min_value=0, step=1, format="%d",
                                                          help="Blank = no limit"),
        })

    st.subheader("Hosts (investors, corporates, agencies)")
    u1, u2 = st.columns([3, 1])
    u1.file_uploader("Import hosts (CSV or Excel)", type=["csv", "xlsx", "xls"], key="up_host",
                     on_change=cb_upload, args=("host",), label_visibility="collapsed")
    u2.download_button("Host template", sample_frames()[2].to_csv(index=False).encode("utf-8-sig"),
                       "hosts_template.csv", "text/csv")
    paste_box("host")
    s.h_df = st.data_editor(
        s.h_base, key=f"h_{s.ver}", num_rows="dynamic", hide_index=True,
        column_config={
            "Name": st.column_config.TextColumn("Organisation", required=True, width="medium"),
            "Contact": st.column_config.TextColumn(help="Person attending, e.g. Raymond"),
            "Type": st.column_config.SelectboxColumn(options=HOST_TYPES, default="Investor"),
            "Days": st.column_config.TextColumn(help="Blank = all days. Day 1 for one day. Not here = not attending"),
            "Available From": st.column_config.TextColumn(help="HH:MM", validate=TIME_VALID),
            "Available Until": st.column_config.TextColumn(help="HH:MM", validate=TIME_VALID),
            "Blocked Times": st.column_config.TextColumn(help="e.g. Day 2 10:00-11:00", width="medium"),
            "Excluded Domains": st.column_config.TextColumn(
                help="Never matched with startups in these domains, e.g. K-Beauty, Hardware", width="medium"),
            "Max Meetings": st.column_config.NumberColumn(min_value=0, step=1, format="%d"),
        })

    st.divider()
    a, b, c, d = st.columns(4)
    a.button("Load sample event", on_click=cb_sample)
    b.button("Start blank", on_click=cb_blank)
    c.download_button("Save project (.json)", project_json(), "bizmatch_project.json", "application/json")
    d.file_uploader("Open project", type=["json"], key="up_project", on_change=cb_project,
                    label_visibility="collapsed")


def generate_tab():
    s = ss()
    cfg = current_cfg()
    slots = build_slots(cfg)
    startups, e1 = parse_parties(s.s_df, "startup", cfg.labels)
    hosts, e2 = parse_parties(s.h_df, "host", cfg.labels)
    errors = e1 + e2
    if errors:
        st.error("Fix these before generating:\n\n" + "\n".join(f"- {e}" for e in errors))

    attending = [x for x in startups if not x.absent]
    pairs = sum(compatible(a, b) for a in attending for b in hosts if not b.absent)
    peak = max((sum(h.available(x) for h in hosts) for x in slots), default=0)
    c = st.columns(5)
    c[0].metric("Startups attending", len(attending),
                f"{len(startups) - len(attending)} not here" if len(startups) > len(attending) else None,
                delta_color="off")
    c[1].metric("Hosts", len(hosts))
    c[2].metric("Slots", len(slots))
    c[3].metric("Eligible pairs", pairs)
    c[4].metric("Polishing days", len(cfg.pol_days) if cfg.pol_days else "Off")

    st.subheader("Rules")
    default = "front" if peak <= cfg.tables else "block"
    if s.get("packing_mode") not in PACKING:
        s.packing_mode = default
    a, b = st.columns([1.3, 1])
    with a:
        st.radio("Host slot packing", list(PACKING), format_func=PACKING.get, key="packing_mode")
        if s.packing_mode == "front" and peak > cfg.tables:
            st.warning(f"Up to {peak} hosts are available at the same time but there are only {cfg.tables} "
                       "tables, so they cannot all start at their first slot. Front-packed will leave some "
                       "hosts empty. Use Contiguous block instead.")
    with b:
        st.slider("Priority weight", 1, 5, 3, key="pw",
                  help="How strongly priority startups are favoured when slots are scarce")
        st.slider("Solver time limit (seconds)", 5, 120, 20, key="tl")
        st.checkbox("Spread meetings fairly across hosts too", True, key="hf")
        st.checkbox("Prefer earlier slots", True, key="pe")

    keep_live = False
    if s.result and s.result.get("meetings"):
        has_pins = any(m.get("lock") for m in s.result["meetings"])
        keep_live = st.checkbox(
            "Keep pinned meetings (moved by hand) and anything already Ongoing, Done or No-Show, "
            "and reschedule the rest", value=has_pins, key="keep_live")

    st.caption("Always enforced: each startup and host meet at most once, nobody is double booked, "
               "tables are never over capacity, availability, cutoffs, blocked times and domain "
               "exclusions are respected.")
    go = st.button("Generate schedule", type="primary",
                   disabled=bool(errors) or not startups or not hosts or not slots)
    if go:
        fixed, fixed_pol = [], []
        if keep_live:
            old_cfg = EventConfig(**s.result["cfg"])
            now = now_for(old_cfg)
            for md in s.result["meetings"]:
                m = Meeting(**md)
                if m.lock or eff_status(m, old_cfg, s.statuses, now) in LOCK_STATUSES:
                    fixed.append((m.startup, m.host, m.day, m.start))
            for pd_ in s.result.get("polish", []):
                p_ = Polish(**pd_)
                dd = old_cfg.day_date(p_.day)
                if dd and (dd < now.date() or (dd == now.date() and p_.start <= now.hour * 60 + now.minute)):
                    fixed_pol.append((p_.startup, p_.day, p_.start))
        settings = SolverSettings(packing=s.packing_mode, time_limit=s.tl, priority_weight=s.pw,
                                  host_fairness=s.hf, prefer_early=s.pe)
        with st.spinner("Optimising with OR-Tools CP-SAT…"):
            res = solve_schedule(cfg, slots, startups, hosts, settings, fixed, fixed_pol)
        if res.meetings:
            pinned = {(f[0], f[1]) for f in fixed if any(md.get("lock") and md["startup"] == f[0] and
                                                         md["host"] == f[1] for md in s.result["meetings"])} \
                if keep_live else set()
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
            pol_txt = f" and {len(res.polish)} pitch polishing sessions" if res.polish else ""
            st.success(f"{res.message} {len(res.meetings)} meetings{pol_txt} in {res.wall:.1f}s. "
                       "Open the Results tab.")
            if res.dropped_fixed:
                st.warning("These kept meetings no longer fit the setup and were dropped: "
                           + "; ".join(res.dropped_fixed))
        else:
            st.error(f"{res.status}: {res.message}")

    if s.result:
        r = s.result
        st.caption(f"Last run: {r['status']} · {len(r['meetings'])} meetings · "
                   f"{r.get('wall', 0):.1f}s · {r.get('n_vars', 0):,} decision variables")


def results_tab():
    s = ss()
    r = s.result
    if not r or not r.get("meetings"):
        st.info("Generate a schedule first.")
        return
    cfg = EventConfig(**r["cfg"])
    slots = build_slots(cfg)
    startups, _ = parse_parties(r["s_df"], "startup", cfg.labels)
    hosts, _ = parse_parties(r["h_df"], "host", cfg.labels)
    now = now_for(cfg)

    day = st.radio("Day", list(range(len(cfg.days))), format_func=cfg.day_title, horizontal=True, key="day_sel")
    key = f"board_{r['id']}_{day}"
    apply_board(r, cfg, slots, startups, hosts, day, s.get(key))

    meetings = [Meeting(**m) for m in r["meetings"]]
    polish = [Polish(**p) for p in r.get("polish", [])]
    parked = r.get("parked", [])

    if r.get("hash") and r["hash"] != frames_hash(current_cfg(), s.s_df, s.h_df):
        st.warning("Setup changed since this schedule was generated. Generate again to apply the changes.")
    issues = r.get("issues", [])
    hard = [i for i in issues if not i.startswith("Gap in")]
    soft = [i for i in issues if i.startswith("Gap in")]

    counts = defaultdict(int)
    for m in meetings:
        counts[m.startup] += 1
    att = [x for x in startups if not x.absent]
    c = st.columns(5)
    c[0].metric("Meetings", len(meetings), f"{len(parked)} unplaced" if parked else None, delta_color="inverse")
    c[1].metric("Startups with meetings", f"{sum(1 for x in att if counts[x.name])}/{len(att)}")
    c[2].metric("Avg per startup", f"{len(meetings) / max(1, len(att)):.1f}")
    c[3].metric("Pitch polishing", len(polish) if polish else "Off")
    c[4].metric("Clashes", len(hard))

    st.caption("Drag a meeting onto another table or time. Green fits, amber pushes the meetings in the way to "
               "the nearest free time, red is a clash. Click a meeting to set its status, pin it or unschedule it. "
               "Ctrl+Z undoes.")
    BOARD(data=board_payload(day, cfg, slots, startups, hosts, r), key=key, default=None)

    if hard:
        with st.expander(f"{len(hard)} clash{'es' if len(hard) > 1 else ''} to resolve", expanded=False):
            st.markdown("\n".join(f"- {i}" for i in hard[:40]))
    elif not parked:
        st.success("Verified: no duplicate pairs, no double booking, table capacity, availability, cutoffs, "
                   "blocked times, domain exclusions" + (", pitch polishing" if polish else ""))
    if soft:
        st.caption(f"Note: {len(soft)} host schedule(s) now have a gap after manual moves.")

    day_pol = sorted([p for p in polish if p.day == day], key=lambda p: (p.k, p.startup))
    if day_pol or day in cfg.pol_days:
        st.subheader("Pitching practice")
        st.caption("The same sessions as the Pitch polishing column on the board. Drag them on the board to change "
                   "the time. Status and Focus / Mentor Note edited here show on the board and in the Excel file.")
        P_ = {x.name: x for x in startups}
        pdf = pd.DataFrame([{
            "Time": f"{fmt(p.start)}–{fmt(p.end)}", "Startup": p.startup,
            "Presenter": P_[p.startup].presenter if p.startup in P_ else "",
            "Round": {1: "1st", 2: "2nd"}.get(p.n, f"{p.n}th"),
            "Status": s.statuses.get(p.pid, "Auto"),
            "Focus / Mentor Note": s.pol_notes.get(p.pid, ""),
        } for p in day_pol], columns=["Time", "Startup", "Presenter", "Round", "Status", "Focus / Mentor Note"])
        pkey = f"pp_{r['id']}_{day}_{r.get('version', '')}"
        st.data_editor(pdf, key=pkey, hide_index=True, disabled=["Time", "Startup", "Presenter", "Round"],
                       on_change=cb_practice_edit, args=(pkey, [p.pid for p in day_pol]),
                       column_config={
                           "Status": st.column_config.SelectboxColumn(options=STATUS_CHOICES, required=True,
                                                                      help="Auto follows the clock"),
                           "Presenter": st.column_config.TextColumn(help="Edit presenters in tab 1"),
                           "Focus / Mentor Note": st.column_config.TextColumn(width="large")})

    st.subheader("Startup summary")
    st.dataframe(startup_summary(cfg, startups, hosts, meetings, s.statuses, now, polish=polish), hide_index=True)
    with st.expander("Host summary"):
        st.dataframe(host_summary(cfg, hosts, meetings), hide_index=True)

    stamp = now.strftime("%Y%m%d_%H%M")
    a, b = st.columns(2)
    a.download_button("Download Excel (.xlsx)",
                      build_excel(cfg, slots, startups, hosts, meetings, s.statuses, now, polish,
                                  s.get("pol_notes", {})),
                      f"matching_schedule_{stamp}.xlsx",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    b.download_button("Download meetings (.csv)",
                      meetings_table(cfg, meetings, s.statuses, now, hosts).to_csv(index=False).encode("utf-8-sig"),
                      f"meetings_{stamp}.csv", "text/csv")


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🗓️", layout="wide")
    init_state()
    st.title(APP_TITLE)
    st.caption("1:1 business matching schedules, optimised with Google OR-Tools CP-SAT")
    flash = ss().pop("flash", None)
    if flash:
        getattr(st, flash[0])(flash[1])
    t1, t2, t3 = st.tabs(["1. Event & Participants", "2. Generate", "3. Results & Live Board"])
    with t1:
        setup_tab()
    with t2:
        generate_tab()
    with t3:
        results_tab()


if __name__ == "__main__":
    main()
