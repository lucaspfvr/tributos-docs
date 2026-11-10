"""Local templates, revision history, PDF archive and SQLite backups."""

import json
import re
import sqlite3
import unicodedata
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from .models import ASSETS, Model

SCHEMA = """
CREATE TABLE IF NOT EXISTS models (
    id TEXT PRIMARY KEY, title TEXT NOT NULL, kind TEXT NOT NULL, document BLOB NOT NULL
);
CREATE TABLE IF NOT EXISTS records (
    id INTEGER PRIMARY KEY AUTOINCREMENT
);
CREATE TABLE IF NOT EXISTS revisions (
    record_id INTEGER NOT NULL REFERENCES records(id), version INTEGER NOT NULL,
    created TEXT NOT NULL, owner TEXT NOT NULL, reference TEXT NOT NULL, model_id TEXT NOT NULL,
    title TEXT NOT NULL, kind TEXT NOT NULL, document BLOB NOT NULL,
    form TEXT NOT NULL, fractions TEXT NOT NULL,
    file TEXT NOT NULL UNIQUE, pdf BLOB NOT NULL, PRIMARY KEY(record_id, version)
);
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY);
"""
APP_ID = 0x54524443
FOLDERS = {"desdobro": "DESDOBRO", "retificacao": "RETIFICAÇÃO", "avaliacao": "AVALIAÇÃO"}


def folder_name(label):
    ascii_label = unicodedata.normalize("NFKD", label).encode("ascii", "ignore").decode().lower()
    clean = re.sub(r"[^a-z0-9]+", "-", ascii_label).strip("-")[:48] or "documentos"
    if re.fullmatch(r"con|prn|aux|nul|com[0-9]|lpt[0-9]", clean):
        return "modelo-" + clean
    return clean


def searchable(text):
    return "".join(c for c in unicodedata.normalize("NFKD", text.casefold())
                   if not unicodedata.combining(c))


def matches_history(row, query):
    code = f"TD-{row['record_id']:06d}" if row["record_id"] else ""
    fields = (row["path"].stem, row["category"], row["title"], row["owner"], row["reference"], code)
    return searchable(query.strip()) in searchable(" ".join(fields))


def atomic_write(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    staging = path.with_name(f".{uuid4().hex}.tmp")
    try:
        with staging.open("xb") as handle:
            handle.write(data)
        staging.replace(path)
    finally:
        staging.unlink(missing_ok=True)


class Storage:
    def __init__(self, directory, archive=None):
        self.root = Path(directory)
        self.archive = Path(archive) if archive is not None else self.root / "Repositório"
        self.root.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.database = self.root / "tributos.sqlite3"
        self.connection = sqlite3.connect(self.database)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.connection.executescript(SCHEMA)
        self.connection.execute(f"PRAGMA application_id = {APP_ID}")
        with self.connection:
            initialized = self.connection.execute("PRAGMA user_version").fetchone()[0]
            if not initialized:
                for kind, title in ASSETS.items():
                    document = Path(__file__).with_name("assets").joinpath(kind + ".docx").read_bytes()
                    self.connection.execute("INSERT INTO models VALUES (?, ?, ?, ?)", (kind, title, kind, document))
                self.connection.execute("PRAGMA user_version = 1")

    def models(self):
        rows = self.connection.execute("SELECT id, title, kind, document FROM models ORDER BY title COLLATE NOCASE")
        return [Model(*row) for row in rows]

    def model_count(self):
        return self.connection.execute("SELECT count(*) FROM models").fetchone()[0]

    def save_model(self, model):
        if not model.title.strip():
            raise ValueError("Informe o nome do modelo.")
        if not model.keys():
            raise ValueError("Adicione ao menos uma chave ao modelo, por exemplo <CimImovel>.")
        with self.connection:
            self.connection.execute("""INSERT INTO models VALUES (?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET title=excluded.title, document=excluded.document""",
                (model.id, model.title.strip(), model.kind, model.document))

    def delete_model(self, model_id):
        with self.connection:
            self.connection.execute("DELETE FROM models WHERE id = ?", (model_id,))

    def save_record(self, model, form, fractions, pdf, reference, record_id=None):
        self.prepare_archive()
        now = datetime.now()
        owner = form.get("NomeProprietario", "").strip()
        path = None
        try:
            with self.connection:
                self.connection.execute("BEGIN IMMEDIATE")
                if record_id is None:
                    record_id = self.connection.execute("INSERT INTO records DEFAULT VALUES").lastrowid
                version = self.connection.execute(
                    "SELECT coalesce(max(version), 0)+1 FROM revisions WHERE record_id=?", (record_id,)
                ).fetchone()[0]
                file = f"{now.year}/{folder_name(model.title)}/TD-{record_id:06d}-v{version:02d}-{uuid4().hex[:12]}.pdf"
                self.connection.execute("INSERT INTO revisions VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (record_id, version, now.isoformat(timespec="seconds"), owner, reference, model.id,
                     model.title, model.kind, model.document, json.dumps(form, ensure_ascii=False),
                     json.dumps(fractions, ensure_ascii=False), file, pdf))
                path = self.ensure_pdf(file, pdf, model.kind)
        except Exception:
            if path is not None:
                path.unlink(missing_ok=True)
            raise
        return record_id, version, path

    def history(self, query=""):
        self.prepare_archive()
        metadata = {}
        for row in self.connection.execute("""
            SELECT record_id, version, title, owner, created, reference, file FROM revisions
        """):
            metadata[Path(row["file"]).name.casefold()] = dict(row)
        rows = []
        for folder in self.categories():
            for path in (self.archive / folder).rglob("*"):
                if path.suffix.casefold() != ".pdf":
                    continue
                try:
                    if not path.is_file():
                        continue
                    modified = path.stat().st_mtime
                except FileNotFoundError:
                    continue
                row = dict(metadata.get(path.name.casefold(), {}))
                if not row:
                    row = {"record_id": None, "version": None, "title": folder,
                           "owner": "", "reference": "",
                           "created": datetime.fromtimestamp(modified).isoformat(timespec="seconds")}
                row.update(path=path, category=folder)
                if matches_history(row, query):
                    rows.append(row)
        return sorted(rows, key=lambda row: (row["created"], row["path"].name), reverse=True)

    def record(self, record_id, version):
        return self.connection.execute("""
            SELECT v.* FROM revisions v
            WHERE v.record_id=? AND v.version=?""", (record_id, version)).fetchone()

    @staticmethod
    def categories():
        return (*FOLDERS.values(), "OUTROS")

    def destination(self, kind):
        return self.archive / FOLDERS.get(kind, "OUTROS")

    def prepare_archive(self):
        for name in self.categories():
            (self.archive / name).mkdir(parents=True, exist_ok=True)
        with self.connection:
            self.connection.execute("BEGIN IMMEDIATE")
            if self.connection.execute("SELECT 1 FROM settings WHERE key='archive-v2'").fetchone():
                return
            self.copy_archived_pdfs()
            self.connection.execute("INSERT INTO settings VALUES ('archive-v2')")

    def copy_archived_pdfs(self):
        for row in self.connection.execute("SELECT file, pdf, kind FROM revisions"):
            self.ensure_pdf(row["file"], row["pdf"], row["kind"])

    def ensure_pdf(self, relative, pdf, kind=None):
        # Paths in restored backups are treated as data.
        if not re.fullmatch(r"[0-9]{4}/[a-z0-9-]{1,48}/TD-[0-9]+-v[0-9]+-[0-9a-f]{12}\.pdf", relative):
            raise ValueError("O caminho do PDF arquivado é inválido.")
        if kind is None:
            row = self.connection.execute("SELECT kind FROM revisions WHERE file=?", (relative,)).fetchone()
            kind = row["kind"] if row else None
        path = self.destination(kind) / Path(relative).name
        if not path.exists():
            atomic_write(path, pdf)
        return path

    def backup(self, destination):
        destination = Path(destination)
        if destination.resolve() == self.database.resolve():
            raise ValueError("Selecione uma pasta de backup diferente do banco em uso.")
        temp = destination.with_name(f".{uuid4().hex}.tmp")
        try:
            target = sqlite3.connect(temp)
            try:
                self.connection.backup(target)
            finally:
                target.close()
            temp.replace(destination)
        finally:
            temp.unlink(missing_ok=True)

    def restore(self, source):
        source = Path(source).resolve()
        if source == self.database.resolve():
            raise ValueError("Selecione um arquivo de backup diferente do banco em uso.")
        backup = sqlite3.connect(source.as_uri() + "?mode=ro", uri=True)
        try:
            if (backup.execute("PRAGMA application_id").fetchone()[0] != APP_ID or
                backup.execute("PRAGMA user_version").fetchone()[0] != 1 or
                backup.execute("PRAGMA quick_check").fetchone()[0] != "ok"):
                raise ValueError("O arquivo não é um backup compatível do Tributos Docs.")
            backup.backup(self.connection)
        finally:
            backup.close()
        self.connection.executescript(SCHEMA)
        self.copy_archived_pdfs()
