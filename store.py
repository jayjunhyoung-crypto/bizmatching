"""Team workspace storage. Works with any SQLAlchemy URL: Postgres (Supabase, Neon) in production,
SQLite for local testing. Tables are created automatically on first use.

events     one row per event: the whole setup + schedule as JSON, with a version number for safe saves
event_items one row per meeting/polishing status or note, so several people can update the live board
           at the same time without overwriting each other
"""
from __future__ import annotations

import time
import uuid

import sqlalchemy as sa

META = sa.MetaData()
EVENTS = sa.Table(
    "bm_events", META,
    sa.Column("id", sa.String(36), primary_key=True),
    sa.Column("name", sa.Text, nullable=False, default=""),
    sa.Column("data", sa.Text, nullable=False, default="{}"),
    sa.Column("version", sa.Integer, nullable=False, default=1),
    sa.Column("created_at", sa.Float), sa.Column("created_by", sa.Text),
    sa.Column("updated_at", sa.Float), sa.Column("updated_by", sa.Text),
    sa.Column("archived", sa.Boolean, nullable=False, default=False),
)
ITEMS = sa.Table(
    "bm_event_items", META,
    sa.Column("event_id", sa.String(36), primary_key=True),
    sa.Column("item_id", sa.Text, primary_key=True),
    sa.Column("kind", sa.String(16), primary_key=True),       # "status" or "note"
    sa.Column("value", sa.Text, nullable=False, default=""),  # "" means cleared
    sa.Column("updated_at", sa.Float, index=True), sa.Column("updated_by", sa.Text),
)

DIRECTORY = sa.Table(       # the team's company directory: one row per kind ("startup", "partner")
    "bm_directory", META,
    sa.Column("kind", sa.String(16), primary_key=True),
    sa.Column("data", sa.Text, nullable=False, default="{}"),
    sa.Column("version", sa.Integer, nullable=False, default=0),
    sa.Column("updated_at", sa.Float), sa.Column("updated_by", sa.Text),
)

DOCS = sa.Table(            # programs, intake records, report drafts: one JSON document per row
    "bm_docs", META,
    sa.Column("kind", sa.String(24), primary_key=True),
    sa.Column("id", sa.String(200), primary_key=True),
    sa.Column("parent", sa.String(64), index=True),
    sa.Column("name", sa.Text, nullable=False, default=""),
    sa.Column("data", sa.Text, nullable=False, default="{}"),
    sa.Column("version", sa.Integer, nullable=False, default=1),
    sa.Column("updated_at", sa.Float), sa.Column("updated_by", sa.Text),
    sa.Column("archived", sa.Boolean, nullable=False, default=False),
)


class Store:
    def __init__(self, url: str):
        url = url.strip()
        for pre in ("postgres://", "postgresql://"):   # use the psycopg 3 driver for plain Postgres URLs
            if url.startswith(pre):
                url = "postgresql+psycopg://" + url[len(pre):]
        kw = {"pool_pre_ping": True}
        if url.startswith("sqlite"):
            kw["connect_args"] = {"check_same_thread": False, "timeout": 15}
        else:
            kw.update(pool_size=3, max_overflow=2, pool_recycle=300)
        self.eng = sa.create_engine(url, **kw)
        META.create_all(self.eng)
        if self.eng.dialect.name == "postgresql":
            # Supabase also serves the public schema over its REST API. Row level security with no policies
            # shuts that door, while this app (the table owner) keeps full access.
            with self.eng.begin() as c:
                for t in META.sorted_tables:
                    c.execute(sa.text(f'alter table "{t.name}" enable row level security'))

    # ---------------- events ----------------
    def list_events(self) -> list[dict]:
        q = (sa.select(EVENTS.c.id, EVENTS.c.name, EVENTS.c.updated_at, EVENTS.c.updated_by, EVENTS.c.version)
             .where(EVENTS.c.archived.is_(False)).order_by(EVENTS.c.updated_at.desc()))
        with self.eng.connect() as c:
            return [dict(r._mapping) for r in c.execute(q)]

    def create(self, name: str, data: str, user: str) -> tuple[str, int]:
        eid, now = str(uuid.uuid4()), time.time()
        with self.eng.begin() as c:
            c.execute(EVENTS.insert().values(id=eid, name=name, data=data, version=1, created_at=now,
                                             created_by=user, updated_at=now, updated_by=user, archived=False))
        return eid, 1

    def get(self, eid: str) -> dict | None:
        with self.eng.connect() as c:
            r = c.execute(sa.select(EVENTS).where(EVENTS.c.id == eid)).first()
            return dict(r._mapping) if r else None

    def head(self, eid: str) -> dict | None:
        q = sa.select(EVENTS.c.version, EVENTS.c.updated_at, EVENTS.c.updated_by, EVENTS.c.archived) \
            .where(EVENTS.c.id == eid)
        with self.eng.connect() as c:
            r = c.execute(q).first()
            return dict(r._mapping) if r else None

    def save(self, eid: str, name: str, data: str, version: int, user: str) -> int | None:
        """Save only if nobody else saved since `version`. Returns the new version, or None on conflict."""
        with self.eng.begin() as c:
            res = c.execute(EVENTS.update()
                            .where(EVENTS.c.id == eid, EVENTS.c.version == version)
                            .values(name=name, data=data, version=version + 1,
                                    updated_at=time.time(), updated_by=user))
            return version + 1 if res.rowcount == 1 else None

    def archive(self, eid: str, user: str) -> None:
        with self.eng.begin() as c:
            c.execute(EVENTS.update().where(EVENTS.c.id == eid)
                      .values(archived=True, updated_at=time.time(), updated_by=user))

    # ---------------- per-item statuses and notes ----------------
    def items(self, eid: str, since: float = 0.0) -> list[dict]:
        q = sa.select(ITEMS).where(ITEMS.c.event_id == eid, ITEMS.c.updated_at > since)
        with self.eng.connect() as c:
            return [dict(r._mapping) for r in c.execute(q)]

    def put_items(self, eid: str, rows: list[tuple[str, str, str]], user: str) -> float:
        """rows: (item_id, kind, value). Last write wins per item."""
        now = time.time()
        with self.eng.begin() as c:
            for item_id, kind, value in rows:
                cond = (ITEMS.c.event_id == eid, ITEMS.c.item_id == item_id, ITEMS.c.kind == kind)
                res = c.execute(ITEMS.update().where(*cond).values(value=value, updated_at=now, updated_by=user))
                if res.rowcount == 0:
                    c.execute(ITEMS.insert().values(event_id=eid, item_id=item_id, kind=kind, value=value,
                                                    updated_at=now, updated_by=user))
        return now

    # ---------------- company directory ----------------
    def get_dir(self, kind: str) -> dict | None:
        with self.eng.connect() as c:
            r = c.execute(sa.select(DIRECTORY).where(DIRECTORY.c.kind == kind)).first()
            return dict(r._mapping) if r else None

    def dir_version(self, kind: str) -> int:
        with self.eng.connect() as c:
            v = c.execute(sa.select(DIRECTORY.c.version).where(DIRECTORY.c.kind == kind)).scalar()
            return int(v or 0)

    def save_dir(self, kind: str, data: str, version: int, user: str) -> int | None:
        """Save only if nobody else saved since `version` (0 = never saved). Returns the new version or None."""
        now = time.time()
        with self.eng.begin() as c:
            if version == 0:
                exists = c.execute(sa.select(DIRECTORY.c.kind).where(DIRECTORY.c.kind == kind)).first()
                if exists:
                    return None
                c.execute(DIRECTORY.insert().values(kind=kind, data=data, version=1, updated_at=now, updated_by=user))
                return 1
            res = c.execute(DIRECTORY.update().where(DIRECTORY.c.kind == kind, DIRECTORY.c.version == version)
                            .values(data=data, version=version + 1, updated_at=now, updated_by=user))
            return version + 1 if res.rowcount == 1 else None

    # ---------------- generic documents (programs, intake records, report drafts) ----------------
    def doc_list(self, kind: str, parent: str | None = None, with_data: bool = False) -> list[dict]:
        cols = [DOCS.c.id, DOCS.c.name, DOCS.c.parent, DOCS.c.version, DOCS.c.updated_at, DOCS.c.updated_by]
        q = sa.select(*(cols + ([DOCS.c.data] if with_data else []))).where(DOCS.c.kind == kind,
                                                                             DOCS.c.archived.is_(False))
        if parent is not None:
            q = q.where(DOCS.c.parent == parent)
        with self.eng.connect() as c:
            return [dict(r._mapping) for r in c.execute(q.order_by(DOCS.c.updated_at.desc()))]

    def doc_get(self, kind: str, id_: str) -> dict | None:
        with self.eng.connect() as c:
            r = c.execute(sa.select(DOCS).where(DOCS.c.kind == kind, DOCS.c.id == id_)).first()
            return dict(r._mapping) if r else None

    def doc_version(self, kind: str, id_: str) -> int:
        with self.eng.connect() as c:
            v = c.execute(sa.select(DOCS.c.version).where(DOCS.c.kind == kind, DOCS.c.id == id_)).scalar()
            return int(v or 0)

    def doc_put(self, kind: str, id_: str, data: str, version: int, user: str, parent: str = "",
                name: str = "") -> int | None:
        """Create (version 0) or update only if nobody saved since `version`. Returns the new version or None."""
        now = time.time()
        if version == 0:
            try:
                with self.eng.begin() as c:
                    if c.execute(sa.select(DOCS.c.id).where(DOCS.c.kind == kind, DOCS.c.id == id_)).first():
                        return None
                    c.execute(DOCS.insert().values(kind=kind, id=id_, parent=parent, name=name, data=data, version=1,
                                                   updated_at=now, updated_by=user, archived=False))
                return 1
            except sa.exc.IntegrityError:   # someone created it at the same moment
                return None
        with self.eng.begin() as c:
            res = c.execute(DOCS.update().where(DOCS.c.kind == kind, DOCS.c.id == id_, DOCS.c.version == version)
                            .values(data=data, name=name, version=version + 1, updated_at=now, updated_by=user))
            return version + 1 if res.rowcount == 1 else None

    def doc_archive(self, kind: str, id_: str, user: str) -> None:
        with self.eng.begin() as c:
            c.execute(DOCS.update().where(DOCS.c.kind == kind, DOCS.c.id == id_)
                      .values(archived=True, updated_at=time.time(), updated_by=user))
