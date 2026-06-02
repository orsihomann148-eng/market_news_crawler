from __future__ import annotations

import json
import hashlib
import sqlite3
import time
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DB_PATH = BASE_DIR / "news_crawler.db"
SCHEMA_VERSION = 1
STAR_JSON_MIGRATION_NAME = "article_star_store_json_migrated"


def normalize_text(value: Any) -> str:
    return str(value or "").strip()


def json_dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def json_loads(value: Any, default: Any = None) -> Any:
    try:
        return json.loads(str(value or ""))
    except Exception:
        return default


def connect(db_path: Path | str | None = None) -> sqlite3.Connection:
    path = Path(db_path or DEFAULT_DB_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA foreign_keys=ON")
    return connection


def init_db(db_path: Path | str | None = None) -> None:
    with connect(db_path) as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS schema_migrations (
                name TEXT PRIMARY KEY,
                applied_at INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS runs (
                run_id TEXT PRIMARY KEY,
                country_code TEXT NOT NULL,
                country_label TEXT,
                output_dir TEXT NOT NULL,
                generated_at TEXT,
                side_label TEXT,
                metadata_json TEXT NOT NULL,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS articles (
                run_id TEXT NOT NULL,
                stage TEXT NOT NULL,
                row_index INTEGER NOT NULL,
                country_code TEXT NOT NULL,
                platform_label TEXT,
                matched_brands TEXT,
                title TEXT,
                title_translated TEXT,
                summary TEXT,
                summary_translated TEXT,
                source_name TEXT,
                source_url TEXT,
                article_url TEXT,
                published_at TEXT,
                category TEXT,
                survey_dimensions TEXT,
                survey_question_ids TEXT,
                survey_indicator_examples TEXT,
                briefing_sentiment TEXT,
                briefing_sentiment_reason TEXT,
                raw_json TEXT NOT NULL,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL,
                PRIMARY KEY (run_id, stage, row_index),
                FOREIGN KEY (run_id) REFERENCES runs(run_id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_articles_country_stage
                ON articles(country_code, stage, published_at);
            CREATE INDEX IF NOT EXISTS idx_articles_run_stage
                ON articles(run_id, stage);

            CREATE TABLE IF NOT EXISTS article_stars (
                article_id TEXT PRIMARY KEY,
                starred INTEGER NOT NULL DEFAULT 1,
                updated_at INTEGER NOT NULL,
                country_code TEXT,
                title TEXT,
                article_url TEXT,
                payload_json TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS manual_articles (
                manual_article_id TEXT PRIMARY KEY,
                run_id TEXT NOT NULL,
                stage TEXT NOT NULL DEFAULT 'after',
                country_code TEXT NOT NULL,
                article_url TEXT,
                title TEXT,
                platform_label TEXT,
                raw_json TEXT NOT NULL,
                enabled INTEGER NOT NULL DEFAULT 1,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL,
                FOREIGN KEY (run_id) REFERENCES runs(run_id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_manual_articles_run_stage
                ON manual_articles(run_id, stage, enabled);
            CREATE INDEX IF NOT EXISTS idx_manual_articles_url
                ON manual_articles(run_id, article_url);
            """
        )
        connection.execute(
            "INSERT OR IGNORE INTO schema_migrations(name, applied_at) VALUES (?, ?)",
            (f"schema_v{SCHEMA_VERSION}", int(time.time())),
        )


def list_table_names(db_path: Path | str | None = None) -> list[str]:
    init_db(db_path)
    with connect(db_path) as connection:
        rows = connection.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall()
    return [str(row["name"]) for row in rows]


def article_column_value(row: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = row.get(key)
        if value not in (None, ""):
            if isinstance(value, (list, tuple, dict)):
                return json_dumps(value)
            return normalize_text(value)
    return ""


def save_run(
    output_dir: Path | str,
    metadata: dict[str, Any],
    before_articles: list[dict[str, Any]],
    after_articles: list[dict[str, Any]],
    *,
    db_path: Path | str | None = None,
) -> str:
    init_db(db_path)
    output_path = Path(output_dir)
    run_id = output_path.name
    now = int(time.time())
    country_code = normalize_text(metadata.get("country_code") or metadata.get("country"))
    country_label = normalize_text(metadata.get("country_label") or metadata.get("country"))
    with connect(db_path) as connection:
        connection.execute(
            """
            INSERT INTO runs(run_id, country_code, country_label, output_dir, generated_at, side_label, metadata_json, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(run_id) DO UPDATE SET
                country_code=excluded.country_code,
                country_label=excluded.country_label,
                output_dir=excluded.output_dir,
                generated_at=excluded.generated_at,
                side_label=excluded.side_label,
                metadata_json=excluded.metadata_json,
                updated_at=excluded.updated_at
            """,
            (
                run_id,
                country_code,
                country_label,
                str(output_path),
                normalize_text(metadata.get("generated_at")),
                normalize_text(metadata.get("side_label")),
                json_dumps(metadata),
                now,
                now,
            ),
        )
        for stage, rows in (("before", before_articles), ("after", after_articles)):
            connection.execute("DELETE FROM articles WHERE run_id=? AND stage=?", (run_id, stage))
            for index, row in enumerate(rows):
                raw = dict(row)
                connection.execute(
                    """
                    INSERT INTO articles(
                        run_id, stage, row_index, country_code, platform_label, matched_brands,
                        title, title_translated, summary, summary_translated, source_name,
                        source_url, article_url, published_at, category, survey_dimensions,
                        survey_question_ids, survey_indicator_examples, briefing_sentiment,
                        briefing_sentiment_reason, raw_json, created_at, updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        run_id,
                        stage,
                        index,
                        country_code,
                        article_column_value(raw, "platform_label", "platform", "brand"),
                        article_column_value(raw, "matched_brands"),
                        article_column_value(raw, "title"),
                        article_column_value(raw, "title_translated"),
                        article_column_value(raw, "summary"),
                        article_column_value(raw, "summary_translated"),
                        article_column_value(raw, "source_name", "source_site"),
                        article_column_value(raw, "source_url"),
                        article_column_value(raw, "verification_final_url", "article_url"),
                        article_column_value(raw, "published_at"),
                        article_column_value(raw, "category"),
                        article_column_value(raw, "survey_dimensions"),
                        article_column_value(raw, "survey_question_ids"),
                        article_column_value(raw, "survey_indicator_examples"),
                        article_column_value(raw, "briefing_sentiment"),
                        article_column_value(raw, "briefing_sentiment_reason"),
                        json_dumps(raw),
                        now,
                        now,
                    ),
                )
    return run_id


def list_runs(country_code: str | None = None, *, db_path: Path | str | None = None, limit: int | None = None) -> list[dict[str, Any]]:
    init_db(db_path)
    params: list[Any] = []
    where = ""
    if normalize_text(country_code):
        where = "WHERE r.country_code=?"
        params.append(normalize_text(country_code))
    sql = f"""
        SELECT
            r.run_id, r.country_code, r.country_label, r.output_dir, r.generated_at,
            r.side_label, r.updated_at, r.metadata_json,
            COALESCE(article_counts.before_count, 0) AS before_count,
            COALESCE(article_counts.after_count, 0) + COALESCE(manual_counts.after_count, 0) AS after_count
        FROM runs r
        LEFT JOIN (
            SELECT
                run_id,
                SUM(CASE WHEN stage='before' THEN 1 ELSE 0 END) AS before_count,
                SUM(CASE WHEN stage='after' THEN 1 ELSE 0 END) AS after_count
            FROM articles
            GROUP BY run_id
        ) article_counts ON article_counts.run_id = r.run_id
        LEFT JOIN (
            SELECT run_id, COUNT(*) AS after_count
            FROM manual_articles
            WHERE stage='after' AND enabled=1
            GROUP BY run_id
        ) manual_counts ON manual_counts.run_id = r.run_id
        {where}
        ORDER BY r.updated_at DESC
    """
    if limit is not None:
        sql += " LIMIT ?"
        params.append(int(limit))
    with connect(db_path) as connection:
        rows = connection.execute(sql, params).fetchall()
    return [dict(row) for row in rows]


def get_run(run_id: str, *, db_path: Path | str | None = None) -> dict[str, Any] | None:
    init_db(db_path)
    with connect(db_path) as connection:
        row = connection.execute("SELECT * FROM runs WHERE run_id=?", (run_id,)).fetchone()
    return dict(row) if row else None


def load_articles(run_id: str, stage: str = "after", *, db_path: Path | str | None = None) -> list[dict[str, Any]]:
    init_db(db_path)
    normalized_stage = "before" if stage == "before" else "after"
    with connect(db_path) as connection:
        rows = connection.execute(
            """
            SELECT raw_json
            FROM articles
            WHERE run_id=? AND stage=?
            ORDER BY row_index ASC
            """,
            (run_id, normalized_stage),
        ).fetchall()
    articles: list[dict[str, Any]] = []
    for row in rows:
        payload = json_loads(row["raw_json"], {})
        if isinstance(payload, dict):
            articles.append(payload)
    return articles


def manual_article_identity(run_id: str, article: dict[str, Any]) -> str:
    url = normalize_text(article.get("verification_final_url") or article.get("article_url")).lower()
    if url:
        raw = f"{normalize_text(run_id)}|url|{url}"
    else:
        raw = "|".join(
            [
                normalize_text(run_id),
                normalize_text(article.get("platform_label") or article.get("platform")),
                normalize_text(article.get("published_at")),
                normalize_text(article.get("title")),
                normalize_text(article.get("source_name") or article.get("source_site")),
            ]
        ).lower()
    return "manual_" + hashlib.sha1(raw.encode("utf-8")).hexdigest()


def save_manual_article(
    run_id: str,
    country_code: str,
    article: dict[str, Any],
    *,
    manual_article_id: str | None = None,
    db_path: Path | str | None = None,
) -> tuple[str, bool]:
    init_db(db_path)
    normalized_run_id = normalize_text(run_id)
    if not normalized_run_id:
        raise ValueError("missing_run_id")
    if not get_run(normalized_run_id, db_path=db_path):
        raise ValueError("run_not_found")

    raw = dict(article)
    raw["manual_added"] = True
    raw["source_discovery"] = raw.get("source_discovery") or "manual_added"
    raw["stage"] = "after"
    raw["country_code"] = normalize_text(country_code)
    normalized_id = normalize_text(manual_article_id) or manual_article_identity(normalized_run_id, raw)
    raw["manual_article_id"] = normalized_id
    now = int(time.time())

    with connect(db_path) as connection:
        existing = connection.execute(
            "SELECT 1 FROM manual_articles WHERE manual_article_id=?",
            (normalized_id,),
        ).fetchone()
        connection.execute(
            """
            INSERT INTO manual_articles(
                manual_article_id, run_id, stage, country_code, article_url, title,
                platform_label, raw_json, enabled, created_at, updated_at
            )
            VALUES (?, ?, 'after', ?, ?, ?, ?, ?, 1, ?, ?)
            ON CONFLICT(manual_article_id) DO UPDATE SET
                country_code=excluded.country_code,
                article_url=excluded.article_url,
                title=excluded.title,
                platform_label=excluded.platform_label,
                raw_json=excluded.raw_json,
                enabled=1,
                updated_at=excluded.updated_at
            """,
            (
                normalized_id,
                normalized_run_id,
                normalize_text(country_code),
                article_column_value(raw, "verification_final_url", "article_url"),
                article_column_value(raw, "title"),
                article_column_value(raw, "platform_label", "platform", "brand"),
                json_dumps(raw),
                now,
                now,
            ),
        )
    return normalized_id, bool(existing)


def load_manual_articles(
    run_id: str,
    stage: str = "after",
    *,
    include_disabled: bool = False,
    db_path: Path | str | None = None,
) -> list[dict[str, Any]]:
    init_db(db_path)
    normalized_stage = "before" if stage == "before" else "after"
    sql = """
        SELECT raw_json, manual_article_id, enabled
        FROM manual_articles
        WHERE run_id=? AND stage=?
    """
    params: list[Any] = [normalize_text(run_id), normalized_stage]
    if not include_disabled:
        sql += " AND enabled=1"
    sql += " ORDER BY updated_at ASC"
    with connect(db_path) as connection:
        rows = connection.execute(sql, params).fetchall()

    articles: list[dict[str, Any]] = []
    for row in rows:
        payload = json_loads(row["raw_json"], {})
        if not isinstance(payload, dict):
            payload = {}
        payload["manual_added"] = True
        payload["manual_article_id"] = normalize_text(payload.get("manual_article_id") or row["manual_article_id"])
        payload["manual_enabled"] = bool(row["enabled"])
        articles.append(payload)
    return articles


def set_manual_article_enabled(
    manual_article_id: str,
    enabled: bool,
    *,
    db_path: Path | str | None = None,
) -> bool:
    init_db(db_path)
    with connect(db_path) as connection:
        cursor = connection.execute(
            "UPDATE manual_articles SET enabled=?, updated_at=? WHERE manual_article_id=?",
            (1 if enabled else 0, int(time.time()), normalize_text(manual_article_id)),
        )
        return cursor.rowcount > 0


def migration_applied(name: str, *, db_path: Path | str | None = None) -> bool:
    init_db(db_path)
    with connect(db_path) as connection:
        row = connection.execute("SELECT 1 FROM schema_migrations WHERE name=?", (name,)).fetchone()
    return bool(row)


def mark_migration_applied(name: str, *, db_path: Path | str | None = None) -> None:
    init_db(db_path)
    with connect(db_path) as connection:
        connection.execute(
            "INSERT OR IGNORE INTO schema_migrations(name, applied_at) VALUES (?, ?)",
            (name, int(time.time())),
        )


def migrate_star_json_if_needed(json_path: Path | str, *, db_path: Path | str | None = None) -> int:
    init_db(db_path)
    if migration_applied(STAR_JSON_MIGRATION_NAME, db_path=db_path):
        return 0
    path = Path(json_path)
    migrated_count = 0
    if path.exists():
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            payload = {}
        if isinstance(payload, dict):
            for article_id, value in payload.items():
                if not isinstance(value, dict):
                    continue
                set_article_star(str(article_id), value, True, db_path=db_path)
                migrated_count += 1
    mark_migration_applied(STAR_JSON_MIGRATION_NAME, db_path=db_path)
    return migrated_count


def load_star_store(*, db_path: Path | str | None = None) -> dict[str, dict[str, Any]]:
    init_db(db_path)
    with connect(db_path) as connection:
        rows = connection.execute(
            "SELECT article_id, payload_json, updated_at FROM article_stars WHERE starred=1"
        ).fetchall()
    store: dict[str, dict[str, Any]] = {}
    for row in rows:
        payload = json_loads(row["payload_json"], {})
        if not isinstance(payload, dict):
            payload = {}
        payload["starred"] = True
        payload["updated_at"] = int(payload.get("updated_at") or row["updated_at"] or 0)
        store[str(row["article_id"])] = payload
    return store


def load_starred_article_ids(*, db_path: Path | str | None = None) -> set[str]:
    return set(load_star_store(db_path=db_path))


def set_article_star(
    article_id: str,
    payload: dict[str, Any] | None,
    starred: bool,
    *,
    db_path: Path | str | None = None,
) -> None:
    init_db(db_path)
    normalized_id = normalize_text(article_id)
    if not normalized_id:
        return
    now = int(time.time())
    payload = dict(payload or {})
    payload["starred"] = bool(starred)
    payload["updated_at"] = int(payload.get("updated_at") or now)
    with connect(db_path) as connection:
        if starred:
            connection.execute(
                """
                INSERT INTO article_stars(article_id, starred, updated_at, country_code, title, article_url, payload_json)
                VALUES (?, 1, ?, ?, ?, ?, ?)
                ON CONFLICT(article_id) DO UPDATE SET
                    starred=1,
                    updated_at=excluded.updated_at,
                    country_code=excluded.country_code,
                    title=excluded.title,
                    article_url=excluded.article_url,
                    payload_json=excluded.payload_json
                """,
                (
                    normalized_id,
                    int(payload["updated_at"]),
                    normalize_text(payload.get("country_code")),
                    normalize_text(payload.get("title")),
                    normalize_text(payload.get("article_url")),
                    json_dumps(payload),
                ),
            )
        else:
            connection.execute("DELETE FROM article_stars WHERE article_id=?", (normalized_id,))
