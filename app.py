from __future__ import annotations

import csv
import io
import json
import os
import sqlite3
from datetime import date, datetime, timedelta, timezone
from functools import wraps
from typing import Any

from flask import (
    Flask,
    Response,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

try:
    from openpyxl import Workbook, load_workbook
except ImportError:  # pragma: no cover - optional until requirements are installed
    Workbook = None
    load_workbook = None


DATABASE = os.environ.get("LIFECYCLE_DB", os.path.join(os.path.dirname(__file__), "lifecycle_tracker.db"))
DATE_FIELDS = {"start_date", "due_date", "completed_date"}
REQUIRED_FIELDS = {"tracker_type", "title", "stage", "status", "owner"}
TRACKER_TYPES = {"I20", "D20"}
ROLES = {"Admin": "admin", "Editor": "editor", "Viewer": "viewer"}

FIELD_ALIASES = {
    "tracker type": "tracker_type",
    "trackertype": "tracker_type",
    "tracker": "tracker_type",
    "type": "tracker_type",
    "activity": "activity",
    "opportunity name": "title",
    "opportunity": "title",
    "title": "title",
    "name": "title",
    "account name": "customer",
    "customer": "customer",
    "customer name": "customer",
    "account": "customer",
    "owner": "owner",
    "assigned to": "owner",
    "cortave owner": "cortave_owner",
    "innovator owner": "innovator_owner",
    "who @ cortave": "who_at_cortave",
    "who cortave": "who_at_cortave",
    "stage": "stage",
    "status": "status",
    "description": "description",
    "notes": "notes",
    "start day": "start_day",
    "start date": "start_date",
    "startdate": "start_date",
    "actual start date": "actual_start_date",
    "actual start date edit kick off date for automated deadlines": "actual_start_date",
    "due date": "due_date",
    "duedate": "due_date",
    "target date": "due_date",
    "target due day": "target_due_day",
    "completed date": "completed_date",
    "completion date": "completed_date",
    "completeddate": "completed_date",
    "priority": "priority",
    "next action": "next_action",
    "nextaction": "next_action",
    "link": "link",
    "url": "link",
}

RECORD_FIELDS = [
    "tracker_type",
    "stage",
    "status",
    "title",
    "activity",
    "owner",
    "who_at_cortave",
    "cortave_owner",
    "innovator_owner",
    "customer",
    "description",
    "notes",
    "start_day",
    "start_date",
    "actual_start_date",
    "target_due_day",
    "due_date",
    "completed_date",
    "priority",
    "next_action",
    "link",
    "additional_fields",
]

PARTNER_FIELDS = [
    "partner_name",
    "owner",
    "territory",
    "master_stage",
    "date_of_stage_change",
    "age_of_stage",
    "days_overdue",
    "workbook_link",
    "sf_account",
    "next_steps_notes",
    "partner_type",
    "csm_involved",
]

PARTNER_HEADER_ALIASES = {
    "partner": "partner_name",
    "owner": "owner",
    "territory": "territory",
    "active": "master_stage",
    "current partner stage": "master_stage",
    "date of stage change": "date_of_stage_change",
    "age of stage": "age_of_stage",
    "days overdue": "days_overdue",
    "workbook link": "workbook_link",
    "sf account": "sf_account",
    "next steps notes": "next_steps_notes",
    "next steps note": "next_steps_notes",
    "next steps": "next_steps_notes",
    "notes": "next_steps_notes",
    "partner type": "partner_type",
    "csm involved": "csm_involved",
}


def create_app(test_config: dict[str, Any] | None = None) -> Flask:
    app = Flask(__name__)
    app.config.update(
        SECRET_KEY=os.environ.get("SECRET_KEY", "dev-lifecycle-tracker-key"),
        DATABASE=DATABASE,
        MAX_CONTENT_LENGTH=16 * 1024 * 1024,
    )
    if test_config:
        app.config.update(test_config)

    @app.before_request
    def ensure_database_and_default_user() -> None:
        init_db(app.config["DATABASE"])
        if "user" not in session:
            session["user"] = {"name": "Lifecycle Admin", "role": "admin"}

    @app.context_processor
    def inject_user() -> dict[str, Any]:
        return {"current_user": session.get("user", {"name": "Lifecycle Admin", "role": "admin"})}

    @app.route("/")
    def dashboard() -> str:
        return render_template("dashboard.html", page_title="Overview Dashboard")

    @app.route("/trackers/<tracker_type>")
    def tracker_page(tracker_type: str) -> str:
        tracker_type = tracker_type.upper()
        if tracker_type not in TRACKER_TYPES:
            return redirect(url_for("dashboard"))
        return render_template("tracker.html", page_title=f"{tracker_type} Tracker", tracker_type=tracker_type)

    @app.route("/settings")
    @role_required("admin")
    def settings_page() -> str:
        return render_template("settings.html", page_title="Data Validation / Settings")

    @app.route("/import-export")
    @role_required("admin")
    def import_export_page() -> str:
        return render_template("import_export.html", page_title="Import / Export")

    @app.route("/partners")
    def partners_page() -> str:
        return render_template("partners.html", page_title="Master Partner Dashboard")

    @app.route("/admin")
    @role_required("admin")
    def admin_page() -> str:
        return render_template("admin.html", page_title="Admin Settings", roles=ROLES)

    @app.route("/login", methods=["GET", "POST"])
    def login() -> Response | str:
        if request.method == "POST":
            role = request.form.get("role", "admin")
            if role not in ROLES.values():
                flash("Invalid role selected.", "error")
                return redirect(url_for("login"))
            name = request.form.get("name", "").strip() or role.title()
            session["user"] = {"name": name, "role": role}
            return redirect(url_for("dashboard"))
        return render_template("login.html", page_title="Sign in", roles=ROLES)

    @app.route("/logout", methods=["POST"])
    def logout() -> Response:
        session.clear()
        return redirect(url_for("login"))

    @app.get("/api/options")
    def get_options() -> Response:
        db = get_db(app.config["DATABASE"])
        rows = db.execute(
            """
            SELECT id, category, value, label, color, sort_order
            FROM settings_options
            WHERE is_active = 1
            ORDER BY category, sort_order, label
            """
        ).fetchall()
        grouped: dict[str, list[dict[str, Any]]] = {}
        for row in rows:
            grouped.setdefault(row["category"], []).append(dict(row))
        return jsonify(grouped)

    @app.post("/api/options")
    @role_required("admin")
    def create_option() -> Response:
        payload = request.get_json(force=True)
        category = clean_text(payload.get("category"))
        value = clean_text(payload.get("value"))
        label = clean_text(payload.get("label")) or value
        color = clean_text(payload.get("color"))
        sort_order = int(payload.get("sortOrder") or payload.get("sort_order") or 100)
        if category not in {"status", "stage", "owner", "priority"} or not value:
            return jsonify({"error": "Category and value are required."}), 400
        db = get_db(app.config["DATABASE"])
        try:
            cursor = db.execute(
                """
                INSERT INTO settings_options (category, value, label, color, sort_order, is_active)
                VALUES (?, ?, ?, ?, ?, 1)
                """,
                (category, value, label, color, sort_order),
            )
            db.commit()
        except sqlite3.IntegrityError:
            return jsonify({"error": "That option already exists."}), 409
        return jsonify({"id": cursor.lastrowid}), 201

    @app.put("/api/options/<int:option_id>")
    @role_required("admin")
    def update_option(option_id: int) -> Response:
        payload = request.get_json(force=True)
        value = clean_text(payload.get("value"))
        label = clean_text(payload.get("label")) or value
        color = clean_text(payload.get("color"))
        sort_order = int(payload.get("sortOrder") or payload.get("sort_order") or 100)
        if not value:
            return jsonify({"error": "Value is required."}), 400
        db = get_db(app.config["DATABASE"])
        db.execute(
            """
            UPDATE settings_options
            SET value = ?, label = ?, color = ?, sort_order = ?
            WHERE id = ?
            """,
            (value, label, color, sort_order, option_id),
        )
        db.commit()
        return jsonify({"ok": True})

    @app.delete("/api/options/<int:option_id>")
    @role_required("admin")
    def delete_option(option_id: int) -> Response:
        db = get_db(app.config["DATABASE"])
        db.execute("UPDATE settings_options SET is_active = 0 WHERE id = ?", (option_id,))
        db.commit()
        return jsonify({"ok": True})

    @app.get("/api/dashboard")
    def dashboard_data() -> Response:
        db = get_db(app.config["DATABASE"])
        today = date.today().isoformat()
        week_end = (date.today() + timedelta(days=7)).isoformat()
        records = [record_to_dict(row) for row in db.execute("SELECT * FROM tracker_records ORDER BY updated_at DESC").fetchall()]
        summary = {
            "totalI20": count_where(records, trackerType="I20"),
            "totalD20": count_where(records, trackerType="D20"),
            "byStatus": group_count(records, "status"),
            "byStage": group_count(records, "stage"),
            "overdue": [
                record for record in records if record.get("dueDate") and record["dueDate"] < today and record["status"] != "Completed"
            ],
            "dueThisWeek": [
                record
                for record in records
                if record.get("dueDate") and today <= record["dueDate"] <= week_end and record["status"] != "Completed"
            ],
            "recentlyUpdated": records[:8],
            "completed": count_where(records, status="Completed"),
            "onHold": count_where(records, status="On Hold"),
            "inProgress": count_where(records, status="In Progress"),
            "notStarted": count_where(records, status="Not Started"),
        }
        return jsonify(summary)

    @app.get("/api/records")
    def list_records() -> Response:
        filters = parse_record_filters(request.args)
        records = query_records(app.config["DATABASE"], filters)
        return jsonify({"records": records})

    @app.get("/api/records/<int:record_id>")
    def get_record(record_id: int) -> Response:
        db = get_db(app.config["DATABASE"])
        row = db.execute("SELECT * FROM tracker_records WHERE id = ?", (record_id,)).fetchone()
        if not row:
            return jsonify({"error": "Record not found."}), 404
        history = [
            dict(item)
            for item in db.execute(
                "SELECT * FROM audit_history WHERE tracker_record_id = ? ORDER BY changed_at DESC",
                (record_id,),
            ).fetchall()
        ]
        return jsonify({"record": record_to_dict(row), "history": history})

    @app.post("/api/records")
    @role_required("admin", "editor")
    def create_record() -> Response:
        payload = request.get_json(force=True)
        record, errors = validate_record_payload(payload, app.config["DATABASE"])
        if errors:
            return jsonify({"errors": errors}), 400
        now = utc_now()
        user = session["user"]["name"]
        db = get_db(app.config["DATABASE"])
        values = [record.get(field) for field in RECORD_FIELDS]
        cursor = db.execute(
            f"""
            INSERT INTO tracker_records ({", ".join(RECORD_FIELDS)}, created_at, updated_at, updated_by)
            VALUES ({", ".join(["?"] * len(RECORD_FIELDS))}, ?, ?, ?)
            """,
            (*values, now, now, user),
        )
        record_id = cursor.lastrowid
        db.execute(
            """
            INSERT INTO audit_history
                (tracker_record_id, changed_at, changed_by, previous_status, new_status, previous_stage, new_stage, change_summary)
            VALUES (?, ?, ?, NULL, ?, NULL, ?, ?)
            """,
            (record_id, now, user, record["status"], record["stage"], json.dumps({"action": "created"})),
        )
        db.commit()
        return jsonify({"id": record_id}), 201

    @app.put("/api/records/<int:record_id>")
    @role_required("admin", "editor")
    def update_record(record_id: int) -> Response:
        payload = request.get_json(force=True)
        record, errors = validate_record_payload(payload, app.config["DATABASE"])
        if errors:
            return jsonify({"errors": errors}), 400
        db = get_db(app.config["DATABASE"])
        existing = db.execute("SELECT * FROM tracker_records WHERE id = ?", (record_id,)).fetchone()
        if not existing:
            return jsonify({"error": "Record not found."}), 404
        now = utc_now()
        user = session["user"]["name"]
        db.execute(
            f"""
            UPDATE tracker_records
            SET {", ".join([field + " = ?" for field in RECORD_FIELDS])},
                updated_at = ?,
                updated_by = ?
            WHERE id = ?
            """,
            (*[record.get(field) for field in RECORD_FIELDS], now, user, record_id),
        )
        changes = summarize_changes(existing, record)
        if changes:
            db.execute(
                """
                INSERT INTO audit_history
                    (tracker_record_id, changed_at, changed_by, previous_status, new_status, previous_stage, new_stage, change_summary)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record_id,
                    now,
                    user,
                    existing["status"],
                    record["status"],
                    existing["stage"],
                    record["stage"],
                    json.dumps(changes),
                ),
            )
        db.commit()
        return jsonify({"ok": True})

    @app.delete("/api/records/<int:record_id>")
    @role_required("admin")
    def delete_record(record_id: int) -> Response:
        db = get_db(app.config["DATABASE"])
        db.execute("DELETE FROM tracker_records WHERE id = ?", (record_id,))
        db.commit()
        return jsonify({"ok": True})

    @app.post("/api/import")
    @role_required("admin")
    def import_records() -> Response:
        upload = request.files.get("file")
        tracker_type = clean_text(request.form.get("trackerType") or request.form.get("tracker_type")).upper()
        if tracker_type not in TRACKER_TYPES:
            return jsonify({"error": "Choose I20 or D20 as the import target."}), 400
        if not upload or not upload.filename:
            return jsonify({"error": "Upload a CSV or Excel file."}), 400

        rows = parse_uploaded_file(upload, tracker_type)
        imported = 0
        errors = []
        for index, row in enumerate(rows, start=2):
            payload = map_import_row(row, tracker_type)
            record, row_errors = validate_record_payload(payload, app.config["DATABASE"])
            if row_errors:
                errors.append({"row": index, "errors": row_errors})
                continue
            create_record_from_import(record, app.config["DATABASE"], session["user"]["name"])
            imported += 1
        return jsonify({"imported": imported, "errors": errors})

    @app.get("/api/partners")
    def list_partners() -> Response:
        filters = parse_partner_filters(request.args)
        partners = query_partners(app.config["DATABASE"], filters)
        summary = partner_summary(app.config["DATABASE"], partners)
        return jsonify({"partners": partners, "summary": summary})

    @app.get("/api/partners/options")
    def partner_options() -> Response:
        db = get_db(app.config["DATABASE"])
        options = {}
        for key, column in {
            "owners": "owner",
            "territories": "territory",
            "stages": "master_stage",
            "partnerTypes": "partner_type",
            "csmInvolved": "csm_involved",
        }.items():
            rows = db.execute(
                f"SELECT DISTINCT {column} AS value FROM partners WHERE COALESCE({column}, '') != '' ORDER BY {column}"
            ).fetchall()
            options[key] = [row["value"] for row in rows]
        db.close()
        return jsonify(options)

    @app.get("/api/partners/<int:partner_id>")
    def get_partner(partner_id: int) -> Response:
        db = get_db(app.config["DATABASE"])
        row = db.execute("SELECT * FROM partners WHERE id = ?", (partner_id,)).fetchone()
        db.close()
        if not row:
            return jsonify({"error": "Partner not found."}), 404
        return jsonify({"partner": partner_to_dict(row)})

    @app.post("/api/partners/import")
    @role_required("admin")
    def import_partners() -> Response:
        upload = request.files.get("file")
        if not upload or not upload.filename:
            return jsonify({"error": "Upload the Partner Dashboard Excel file."}), 400
        if not upload.filename.lower().endswith(".xlsx"):
            return jsonify({"error": "Partner import expects an .xlsx workbook."}), 400
        try:
            rows = parse_partner_dashboard_file(upload)
            upload.stream.seek(0)
            org_impact = parse_org_impact_file(upload)
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        created, updated, errors = upsert_partner_rows(app.config["DATABASE"], rows, session["user"]["name"], upload.filename)
        org_rows = save_org_impact(app.config["DATABASE"], org_impact)
        return jsonify({"created": created, "updated": updated, "errors": errors, "orgImpactRows": org_rows})

    @app.get("/api/partner-import-history")
    @role_required("admin")
    def partner_import_history() -> Response:
        db = get_db(app.config["DATABASE"])
        rows = db.execute(
            """
            SELECT * FROM partner_import_history
            ORDER BY imported_at DESC
            LIMIT 25
            """
        ).fetchall()
        db.close()
        return jsonify({"history": [partner_import_history_to_dict(row) for row in rows]})

    @app.get("/api/org-impact")
    def get_org_impact() -> Response:
        return jsonify(load_org_impact(app.config["DATABASE"]))

    @app.get("/api/lifecycle-workbook")
    def get_lifecycle_workbook() -> Response:
        workbook = load_default_lifecycle_workbook(app.config["DATABASE"])
        return jsonify({"workbook": workbook})

    @app.post("/api/lifecycle-workbook/link")
    @role_required("admin", "editor")
    def link_lifecycle_workbook() -> Response:
        payload = request.get_json(force=True)
        partner_id = int(payload.get("partnerId") or payload.get("partner_id") or 0)
        if partner_id <= 0:
            return jsonify({"error": "Choose a partner to link."}), 400
        db = get_db(app.config["DATABASE"])
        partner = db.execute("SELECT * FROM partners WHERE id = ?", (partner_id,)).fetchone()
        if not partner:
            db.close()
            return jsonify({"error": "Partner not found."}), 404
        now = utc_now()
        db.execute(
            """
            UPDATE lifecycle_workbooks
            SET partner_id = ?, partner_name = ?, master_workbook_link = ?, updated_at = ?
            WHERE id = 1
            """,
            (partner_id, partner["partner_name"], partner["workbook_link"], now),
        )
        db.commit()
        db.close()
        return jsonify({"workbook": load_default_lifecycle_workbook(app.config["DATABASE"])})

    @app.post("/api/lifecycle-workbook/unlink")
    @role_required("admin", "editor")
    def unlink_lifecycle_workbook() -> Response:
        db = get_db(app.config["DATABASE"])
        db.execute(
            """
            UPDATE lifecycle_workbooks
            SET partner_id = NULL, partner_name = '', master_workbook_link = '', updated_at = ?
            WHERE id = 1
            """,
            (utc_now(),),
        )
        db.commit()
        db.close()
        return jsonify({"workbook": load_default_lifecycle_workbook(app.config["DATABASE"])})

    @app.get("/api/export")
    @role_required("admin", "editor", "viewer")
    def export_records() -> Response:
        filters = parse_record_filters(request.args)
        records = query_records(app.config["DATABASE"], filters)
        export_format = request.args.get("format", "csv").lower()
        filename_prefix = (filters.get("tracker_type") or "lifecycle").lower()
        flat_records = flatten_export_records(records)
        if export_format == "xlsx":
            if Workbook is None:
                return jsonify({"error": "Excel export requires openpyxl."}), 500
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Tracker Records"
            headers = export_headers(records)
            sheet.append(headers)
            for record in flat_records:
                sheet.append([record.get(header) for header in headers])
            output = io.BytesIO()
            workbook.save(output)
            output.seek(0)
            return send_file(
                output,
                as_attachment=True,
                download_name=f"{filename_prefix}-tracker-export.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        output = io.StringIO()
        headers = export_headers(records)
        writer = csv.DictWriter(output, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(flat_records)
        return Response(
            output.getvalue(),
            mimetype="text/csv",
            headers={"Content-Disposition": f"attachment; filename={filename_prefix}-tracker-export.csv"},
        )

    return app


def role_required(*roles: str):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            user = session.get("user", {"role": "admin"})
            if user.get("role") not in roles:
                if request.path.startswith("/api/"):
                    return jsonify({"error": "You do not have permission for this action."}), 403
                flash("You do not have permission for that page.", "error")
                return redirect(url_for("dashboard"))
            return view(*args, **kwargs)

        return wrapped

    return decorator


def get_db(database_path: str) -> sqlite3.Connection:
    connection = sqlite3.connect(database_path)
    connection.row_factory = sqlite3.Row
    return connection


def init_db(database_path: str) -> None:
    db = get_db(database_path)
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS tracker_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tracker_type TEXT NOT NULL,
            stage TEXT NOT NULL,
            status TEXT NOT NULL,
            title TEXT NOT NULL,
            activity TEXT,
            owner TEXT NOT NULL,
            who_at_cortave TEXT,
            cortave_owner TEXT,
            innovator_owner TEXT,
            customer TEXT,
            description TEXT,
            notes TEXT,
            start_day TEXT,
            start_date TEXT,
            actual_start_date TEXT,
            target_due_day TEXT,
            due_date TEXT,
            completed_date TEXT,
            priority TEXT,
            next_action TEXT,
            link TEXT,
            additional_fields TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            updated_by TEXT
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS settings_options (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            value TEXT NOT NULL,
            label TEXT NOT NULL,
            color TEXT,
            sort_order INTEGER NOT NULL DEFAULT 100,
            is_active INTEGER NOT NULL DEFAULT 1,
            UNIQUE(category, value)
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tracker_record_id INTEGER NOT NULL,
            changed_at TEXT NOT NULL,
            changed_by TEXT,
            previous_status TEXT,
            new_status TEXT,
            previous_stage TEXT,
            new_stage TEXT,
            change_summary TEXT,
            FOREIGN KEY(tracker_record_id) REFERENCES tracker_records(id) ON DELETE CASCADE
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS partners (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            partner_name TEXT NOT NULL UNIQUE,
            owner TEXT,
            territory TEXT,
            master_stage TEXT,
            date_of_stage_change TEXT,
            age_of_stage INTEGER,
            days_overdue INTEGER,
            workbook_link TEXT,
            sf_account TEXT,
            next_steps_notes TEXT,
            partner_type TEXT,
            csm_involved TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS lifecycle_workbooks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            workbook_name TEXT NOT NULL,
            partner_id INTEGER,
            partner_name TEXT,
            master_workbook_link TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(partner_id) REFERENCES partners(id) ON DELETE SET NULL
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS org_impact_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            section TEXT NOT NULL,
            row_label TEXT NOT NULL,
            values_json TEXT NOT NULL,
            sort_order INTEGER NOT NULL DEFAULT 0,
            imported_at TEXT NOT NULL
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS partner_import_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            imported_at TEXT NOT NULL,
            imported_by TEXT,
            created_count INTEGER NOT NULL DEFAULT 0,
            updated_count INTEGER NOT NULL DEFAULT 0,
            error_count INTEGER NOT NULL DEFAULT 0,
            notes TEXT
        )
        """
    )
    ensure_partner_columns(db)
    ensure_default_lifecycle_workbook(db)
    ensure_tracker_record_columns(db)
    seed_options(db)
    db.commit()
    db.close()


def ensure_tracker_record_columns(db: sqlite3.Connection) -> None:
    existing = {row[1] for row in db.execute("PRAGMA table_info(tracker_records)").fetchall()}
    columns = {
        "activity": "TEXT",
        "who_at_cortave": "TEXT",
        "cortave_owner": "TEXT",
        "innovator_owner": "TEXT",
        "start_day": "TEXT",
        "actual_start_date": "TEXT",
        "target_due_day": "TEXT",
        "link": "TEXT",
    }
    for name, column_type in columns.items():
        if name not in existing:
            db.execute(f"ALTER TABLE tracker_records ADD COLUMN {name} {column_type}")


def seed_options(db: sqlite3.Connection) -> None:
    defaults = [
        ("status", "Not Started", "Not Started", "#f4c7c3", 10),
        ("status", "In Progress", "In Progress", "#fce8b2", 20),
        ("status", "On Hold", "On Hold", "#c9daf8", 30),
        ("status", "Completed", "Completed", "#b7e1cd", 40),
        ("stage", "I20", "I20", "", 10),
        ("stage", "I50", "I50", "", 20),
        ("stage", "D20", "D20", "", 30),
        ("stage", "Qualified Out", "Qualified Out", "", 40),
        ("stage", "I0", "I0", "", 50),
        ("stage", "D0", "D0", "", 60),
        ("stage", "D50", "D50", "", 70),
        ("owner", "Unassigned", "Unassigned", "", 10),
        ("owner", "Wade / Sales", "Wade / Sales", "", 20),
        ("owner", "Melissa / Marketing", "Melissa / Marketing", "", 30),
        ("owner", "Mark / Innovator", "Mark / Innovator", "", 40),
        ("owner", "Bobby / Experience", "Bobby / Experience", "", 50),
        ("owner", "Nick / Sales", "Nick / Sales", "", 60),
        ("owner", "Customer", "Customer", "", 70),
        ("owner", "Legal", "Legal", "", 80),
        ("owner", "Finance", "Finance", "", 90),
        ("owner", "Innovator", "Innovator", "", 100),
        ("owner", "Innovator / cortave", "Innovator / cortave", "", 110),
        ("priority", "Low", "Low", "", 10),
        ("priority", "Medium", "Medium", "", 20),
        ("priority", "High", "High", "", 30),
        ("priority", "Critical", "Critical", "", 40),
    ]
    db.executemany(
        """
        INSERT OR IGNORE INTO settings_options (category, value, label, color, sort_order)
        VALUES (?, ?, ?, ?, ?)
        """,
        defaults,
    )


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def validate_record_payload(payload: dict[str, Any], database_path: str) -> tuple[dict[str, Any], list[str]]:
    raw_dates = {
        "start_date": payload.get("startDate") or payload.get("start_date"),
        "actual_start_date": payload.get("actualStartDate") or payload.get("actual_start_date"),
        "due_date": payload.get("dueDate") or payload.get("due_date"),
        "completed_date": payload.get("completedDate") or payload.get("completed_date"),
    }
    record = {
        "tracker_type": clean_text(payload.get("trackerType") or payload.get("tracker_type")).upper(),
        "stage": clean_text(payload.get("stage")),
        "status": clean_text(payload.get("status")),
        "title": clean_text(payload.get("title") or payload.get("opportunityName") or payload.get("activity")),
        "activity": clean_text(payload.get("activity") or payload.get("title") or payload.get("opportunityName")),
        "owner": clean_text(payload.get("owner") or payload.get("cortaveOwner") or payload.get("cortave_owner") or payload.get("whoAtCortave") or payload.get("who_at_cortave")),
        "who_at_cortave": clean_text(payload.get("whoAtCortave") or payload.get("who_at_cortave")),
        "cortave_owner": clean_text(payload.get("cortaveOwner") or payload.get("cortave_owner")),
        "innovator_owner": clean_text(payload.get("innovatorOwner") or payload.get("innovator_owner")),
        "customer": clean_text(payload.get("customer") or payload.get("accountName")),
        "description": clean_text(payload.get("description")),
        "notes": clean_text(payload.get("notes")),
        "start_day": clean_text(payload.get("startDay") or payload.get("start_day")),
        "start_date": clean_date(raw_dates["start_date"]),
        "actual_start_date": clean_date(raw_dates["actual_start_date"]),
        "target_due_day": clean_text(payload.get("targetDueDay") or payload.get("target_due_day")),
        "due_date": clean_date(raw_dates["due_date"]),
        "completed_date": clean_date(raw_dates["completed_date"]),
        "priority": clean_text(payload.get("priority")),
        "next_action": clean_text(payload.get("nextAction") or payload.get("next_action")),
        "link": clean_text(payload.get("link")),
        "additional_fields": normalize_additional_fields(payload.get("additionalFields") or payload.get("additional_fields")),
    }
    errors = []
    for field in sorted(REQUIRED_FIELDS):
        if not record[field]:
            errors.append(f"{field.replace('_', ' ').title()} is required.")
    for field, raw_value in raw_dates.items():
        if clean_text(raw_value) and not record[field]:
            errors.append(f"{field.replace('_', ' ').title()} must be a valid date.")
    if record["tracker_type"] and record["tracker_type"] not in TRACKER_TYPES:
        errors.append("Tracker type must be I20 or D20.")
    option_values = load_option_values(database_path)
    if record["status"] and record["status"] not in option_values["status"]:
        errors.append("Status must be one of the configured status options.")
    if record["stage"] and record["stage"] not in option_values["stage"]:
        errors.append("Stage must be one of the configured stage options.")
    if record["priority"] and record["priority"] not in option_values["priority"]:
        errors.append("Priority must be one of the configured priority options.")
    if record["status"] == "Completed" and not record["completed_date"]:
        record["completed_date"] = date.today().isoformat()
    return record, errors


def clean_date(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return ""


def normalize_additional_fields(value: Any) -> str:
    if isinstance(value, dict):
        return json.dumps(value, sort_keys=True)
    text = clean_text(value)
    if not text:
        return "{}"
    try:
        parsed = json.loads(text)
        if isinstance(parsed, dict):
            return json.dumps(parsed, sort_keys=True)
    except json.JSONDecodeError:
        pass
    return json.dumps({"notes": text})


def load_option_values(database_path: str) -> dict[str, set[str]]:
    db = get_db(database_path)
    rows = db.execute("SELECT category, value FROM settings_options WHERE is_active = 1").fetchall()
    values = {"status": set(), "stage": set(), "owner": set(), "priority": set()}
    for row in rows:
        values.setdefault(row["category"], set()).add(row["value"])
    db.close()
    return values


def parse_record_filters(args: Any) -> dict[str, Any]:
    return {
        "tracker_type": clean_text(args.get("trackerType") or args.get("tracker_type")).upper(),
        "status": clean_text(args.get("status")),
        "stage": clean_text(args.get("stage")),
        "owner": clean_text(args.get("owner")),
        "search": clean_text(args.get("search")),
        "due_before": clean_date(args.get("dueBefore") or args.get("due_before")),
        "due_after": clean_date(args.get("dueAfter") or args.get("due_after")),
        "overdue": clean_text(args.get("overdue")).lower() in {"true", "1", "yes"},
        "sort": clean_text(args.get("sort")) or "updated_at",
        "direction": "ASC" if clean_text(args.get("direction")).lower() == "asc" else "DESC",
    }


def query_records(database_path: str, filters: dict[str, Any]) -> list[dict[str, Any]]:
    clauses = []
    params: list[Any] = []
    if filters.get("tracker_type") in TRACKER_TYPES:
        clauses.append("tracker_type = ?")
        params.append(filters["tracker_type"])
    for key in ("status", "stage", "owner"):
        if filters.get(key):
            clauses.append(f"{key} = ?")
            params.append(filters[key])
    if filters.get("search"):
        clauses.append("(title LIKE ? OR customer LIKE ? OR owner LIKE ? OR description LIKE ? OR notes LIKE ?)")
        search = f"%{filters['search']}%"
        params.extend([search, search, search, search, search])
    if filters.get("due_before"):
        clauses.append("due_date <= ?")
        params.append(filters["due_before"])
    if filters.get("due_after"):
        clauses.append("due_date >= ?")
        params.append(filters["due_after"])
    if filters.get("overdue"):
        clauses.append("due_date < ? AND status != 'Completed'")
        params.append(date.today().isoformat())
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    sort_map = {
        "title": "title",
        "stage": "stage",
        "status": "status",
        "owner": "owner",
        "customer": "customer",
        "dueDate": "due_date",
        "due_date": "due_date",
        "updatedAt": "updated_at",
        "updated_at": "updated_at",
        "priority": "priority",
    }
    sort_column = sort_map.get(filters.get("sort"), "updated_at")
    db = get_db(database_path)
    rows = db.execute(
        f"SELECT * FROM tracker_records {where} ORDER BY {sort_column} {filters['direction']}",
        params,
    ).fetchall()
    db.close()
    return [record_to_dict(row) for row in rows]


def record_to_dict(row: sqlite3.Row) -> dict[str, Any]:
    try:
        additional = json.loads(row["additional_fields"] or "{}")
    except json.JSONDecodeError:
        additional = {}
    return {
        "id": row["id"],
        "trackerType": row["tracker_type"],
        "stage": row["stage"],
        "status": row["status"],
        "title": row["title"],
        "activity": row["activity"] or row["title"],
        "owner": row["owner"],
        "whoAtCortave": row["who_at_cortave"] or "",
        "cortaveOwner": row["cortave_owner"] or "",
        "innovatorOwner": row["innovator_owner"] or "",
        "customer": row["customer"] or "",
        "description": row["description"] or "",
        "notes": row["notes"] or "",
        "startDay": row["start_day"] or "",
        "startDate": row["start_date"] or "",
        "actualStartDate": row["actual_start_date"] or "",
        "targetDueDay": row["target_due_day"] or "",
        "dueDate": row["due_date"] or "",
        "completedDate": row["completed_date"] or "",
        "priority": row["priority"] or "",
        "nextAction": row["next_action"] or "",
        "link": row["link"] or "",
        "additionalFields": additional,
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
        "updatedBy": row["updated_by"] or "",
        "isOverdue": bool(row["due_date"] and row["due_date"] < date.today().isoformat() and row["status"] != "Completed"),
    }


def count_where(records: list[dict[str, Any]], **criteria: Any) -> int:
    return sum(1 for record in records if all(record.get(key) == value for key, value in criteria.items()))


def group_count(records: list[dict[str, Any]], field: str) -> dict[str, int]:
    counts: dict[str, int] = {}
    for record in records:
        counts[record.get(field) or "Unspecified"] = counts.get(record.get(field) or "Unspecified", 0) + 1
    return counts


def summarize_changes(existing: sqlite3.Row, updated: dict[str, Any]) -> dict[str, Any]:
    changes = {}
    for field in RECORD_FIELDS:
        if clean_text(existing[field]) != clean_text(updated.get(field)):
            changes[field] = {"previous": existing[field], "new": updated.get(field)}
    return changes


def parse_uploaded_file(upload: Any, tracker_type: str = "") -> list[dict[str, Any]]:
    filename = upload.filename.lower()
    if filename.endswith(".xlsx"):
        if load_workbook is None:
            raise ValueError("Excel import requires openpyxl.")
        workbook = load_workbook(upload.stream, data_only=True)
        preferred_sheet = f"{tracker_type.upper()} Tracker" if tracker_type else ""
        sheet = workbook[preferred_sheet] if preferred_sheet in workbook.sheetnames else workbook.active
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return []
        header_index = find_header_row(rows)
        headers = [clean_text(value) for value in rows[header_index]]
        return [dict(zip(headers, row)) for row in rows[header_index + 1 :] if any(cell not in (None, "") for cell in row)]
    stream = io.StringIO(upload.stream.read().decode("utf-8-sig"))
    return list(csv.DictReader(stream))


def find_header_row(rows: list[tuple[Any, ...]]) -> int:
    for index, row in enumerate(rows[:20]):
        normalized = {normalize_header(value) for value in row if clean_text(value)}
        if {"stage", "status"}.issubset(normalized) and ("activity" in normalized or "opportunity name" in normalized or "title" in normalized):
            return index
    return 0


def map_import_row(row: dict[str, Any], tracker_type: str) -> dict[str, Any]:
    payload: dict[str, Any] = {"trackerType": tracker_type}
    extra: dict[str, Any] = {}
    for raw_key, value in row.items():
        key = clean_text(raw_key)
        mapped = FIELD_ALIASES.get(normalize_header(key))
        if mapped:
            payload[to_camel(mapped)] = value
        elif key:
            extra[key] = value
    payload["trackerType"] = tracker_type
    payload["additionalFields"] = extra
    return payload


def normalize_header(value: Any) -> str:
    text = clean_text(value).lower().replace("_", " ").replace("@", " @ ")
    normalized = "".join(char if char.isalnum() or char in {" ", "@"} else " " for char in text)
    return " ".join(normalized.split())


def to_camel(field: str) -> str:
    parts = field.split("_")
    return parts[0] + "".join(part.title() for part in parts[1:])


def create_record_from_import(record: dict[str, Any], database_path: str, user: str) -> None:
    now = utc_now()
    db = get_db(database_path)
    values = [record.get(field) for field in RECORD_FIELDS]
    cursor = db.execute(
        f"""
        INSERT INTO tracker_records ({", ".join(RECORD_FIELDS)}, created_at, updated_at, updated_by)
        VALUES ({", ".join(["?"] * len(RECORD_FIELDS))}, ?, ?, ?)
        """,
        (*values, now, now, user),
    )
    db.execute(
        """
        INSERT INTO audit_history
            (tracker_record_id, changed_at, changed_by, previous_status, new_status, previous_stage, new_stage, change_summary)
        VALUES (?, ?, ?, NULL, ?, NULL, ?, ?)
        """,
        (cursor.lastrowid, now, user, record["status"], record["stage"], json.dumps({"action": "imported"})),
    )
    db.commit()
    db.close()


def export_headers(records: list[dict[str, Any]]) -> list[str]:
    headers = [
        "id",
        "trackerType",
        "stage",
        "status",
        "title",
        "activity",
        "owner",
        "whoAtCortave",
        "cortaveOwner",
        "innovatorOwner",
        "customer",
        "description",
        "notes",
        "startDay",
        "startDate",
        "actualStartDate",
        "targetDueDay",
        "dueDate",
        "completedDate",
        "priority",
        "nextAction",
        "link",
        "createdAt",
        "updatedAt",
        "updatedBy",
    ]
    extra_keys = sorted({key for record in records for key in record.get("additionalFields", {}).keys()})
    return headers + extra_keys


def flatten_export_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    flattened = []
    for record in records:
        row = {key: value for key, value in record.items() if key != "additionalFields"}
        row.update(record.get("additionalFields", {}))
        flattened.append(row)
    return flattened


def ensure_partner_columns(db: sqlite3.Connection) -> None:
    existing = {row[1] for row in db.execute("PRAGMA table_info(partners)").fetchall()}
    columns = {
        "partner_name": "TEXT",
        "owner": "TEXT",
        "territory": "TEXT",
        "master_stage": "TEXT",
        "date_of_stage_change": "TEXT",
        "age_of_stage": "INTEGER",
        "days_overdue": "INTEGER",
        "workbook_link": "TEXT",
        "sf_account": "TEXT",
        "next_steps_notes": "TEXT",
        "partner_type": "TEXT",
        "csm_involved": "TEXT",
        "created_at": "TEXT",
        "updated_at": "TEXT",
    }
    for name, column_type in columns.items():
        if name not in existing:
            db.execute(f"ALTER TABLE partners ADD COLUMN {name} {column_type}")


def ensure_default_lifecycle_workbook(db: sqlite3.Connection) -> None:
    now = utc_now()
    db.execute(
        """
        INSERT OR IGNORE INTO lifecycle_workbooks
            (id, workbook_name, partner_id, partner_name, master_workbook_link, created_at, updated_at)
        VALUES (1, 'Lifecycle Tracker', NULL, '', '', ?, ?)
        """,
        (now, now),
    )


def parse_partner_filters(args: Any) -> dict[str, Any]:
    return {
        "search": clean_text(args.get("search")),
        "owner": clean_text(args.get("owner")),
        "territory": clean_text(args.get("territory")),
        "master_stage": clean_text(args.get("masterStage") or args.get("master_stage")),
        "partner_type": clean_text(args.get("partnerType") or args.get("partner_type")),
        "csm_involved": clean_text(args.get("csmInvolved") or args.get("csm_involved")),
        "overdue": clean_text(args.get("overdue")).lower() in {"true", "1", "yes"},
        "sort": clean_text(args.get("sort")) or "partnerName",
        "direction": "ASC" if clean_text(args.get("direction")).lower() == "asc" else "DESC",
    }


def query_partners(database_path: str, filters: dict[str, Any]) -> list[dict[str, Any]]:
    clauses = []
    params: list[Any] = []
    if filters.get("search"):
        clauses.append("(p.partner_name LIKE ? OR p.owner LIKE ? OR p.sf_account LIKE ? OR p.next_steps_notes LIKE ?)")
        search = f"%{filters['search']}%"
        params.extend([search, search, search, search])
    for key, column in (
        ("owner", "owner"),
        ("territory", "territory"),
        ("master_stage", "master_stage"),
        ("partner_type", "partner_type"),
        ("csm_involved", "csm_involved"),
    ):
        if filters.get(key):
            clauses.append(f"p.{column} = ?")
            params.append(filters[key])
    if filters.get("overdue"):
        clauses.append("COALESCE(p.days_overdue, 0) > 0")
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    sort_map = {
        "partnerName": "p.partner_name",
        "owner": "p.owner",
        "territory": "p.territory",
        "masterStage": "p.master_stage",
        "partnerType": "p.partner_type",
        "daysOverdue": "COALESCE(p.days_overdue, 0)",
        "dateOfStageChange": "p.date_of_stage_change",
        "updatedAt": "p.updated_at",
    }
    sort_column = sort_map.get(filters.get("sort"), "partner_name")
    db = get_db(database_path)
    rows = db.execute(
        f"""
        SELECT p.*, COUNT(lw.id) AS linked_lifecycle_count
        FROM partners p
        LEFT JOIN lifecycle_workbooks lw ON lw.partner_id = p.id
        {where}
        GROUP BY p.id
        ORDER BY {sort_column} {filters['direction']}, p.partner_name ASC
        """,
        params,
    ).fetchall()
    db.close()
    return [partner_to_dict(row) for row in rows]


def partner_summary(database_path: str, partners: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "totalPartners": len(partners),
        "byStage": group_count(partners, "masterStage"),
        "byOwner": group_count(partners, "owner"),
        "byTerritory": group_count(partners, "territory"),
        "overduePartners": sum(1 for partner in partners if partner.get("daysOverdue", 0) > 0),
        "withLifecycleWorkbook": sum(1 for partner in partners if partner.get("linkedLifecycleCount", 0) > 0),
        "withoutLifecycleWorkbook": sum(1 for partner in partners if partner.get("linkedLifecycleCount", 0) == 0),
        "qualifiedOutPartners": sum(1 for partner in partners if clean_text(partner.get("masterStage")).lower() == "qualified out"),
        "activePartners": sum(1 for partner in partners if clean_text(partner.get("masterStage")).lower() != "qualified out"),
    }


def partner_to_dict(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "partnerName": row["partner_name"] or "",
        "owner": row["owner"] or "",
        "territory": row["territory"] or "",
        "masterStage": row["master_stage"] or "",
        "dateOfStageChange": row["date_of_stage_change"] or "",
        "ageOfStage": row["age_of_stage"] if row["age_of_stage"] is not None else "",
        "daysOverdue": row["days_overdue"] if row["days_overdue"] is not None else 0,
        "workbookLink": row["workbook_link"] or "",
        "sfAccount": row["sf_account"] or "",
        "nextStepsNotes": row["next_steps_notes"] or "",
        "partnerType": row["partner_type"] or "",
        "csmInvolved": row["csm_involved"] or "",
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
        "linkedLifecycleCount": row["linked_lifecycle_count"] if "linked_lifecycle_count" in row.keys() else 0,
        "isWorkbookLinkUrl": is_url(row["workbook_link"] or ""),
    }


def load_default_lifecycle_workbook(database_path: str) -> dict[str, Any]:
    db = get_db(database_path)
    row = db.execute(
        """
        SELECT lw.*, p.owner, p.territory, p.master_stage, p.date_of_stage_change,
               p.age_of_stage, p.days_overdue, p.workbook_link, p.sf_account,
               p.next_steps_notes, p.partner_type, p.csm_involved,
               p.created_at AS partner_created_at, p.updated_at AS partner_updated_at
        FROM lifecycle_workbooks lw
        LEFT JOIN partners p ON p.id = lw.partner_id
        WHERE lw.id = 1
        """
    ).fetchone()
    db.close()
    if not row:
        db = get_db(database_path)
        ensure_default_lifecycle_workbook(db)
        db.commit()
        db.close()
        return load_default_lifecycle_workbook(database_path)
    partner = None
    if row["partner_id"]:
        partner = {
            "id": row["partner_id"],
            "partnerName": row["partner_name"] or "",
            "owner": row["owner"] or "",
            "territory": row["territory"] or "",
            "masterStage": row["master_stage"] or "",
            "dateOfStageChange": row["date_of_stage_change"] or "",
            "ageOfStage": row["age_of_stage"] if row["age_of_stage"] is not None else "",
            "daysOverdue": row["days_overdue"] if row["days_overdue"] is not None else 0,
            "workbookLink": row["workbook_link"] or row["master_workbook_link"] or "",
            "sfAccount": row["sf_account"] or "",
            "nextStepsNotes": row["next_steps_notes"] or "",
            "partnerType": row["partner_type"] or "",
            "csmInvolved": row["csm_involved"] or "",
            "updatedAt": row["partner_updated_at"] or "",
            "isWorkbookLinkUrl": is_url(row["workbook_link"] or row["master_workbook_link"] or ""),
        }
    return {
        "id": row["id"],
        "workbookName": row["workbook_name"],
        "partnerId": row["partner_id"],
        "partnerName": row["partner_name"] or "",
        "masterWorkbookLink": row["master_workbook_link"] or "",
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
        "partner": partner,
    }


def parse_partner_dashboard_file(upload: Any) -> list[dict[str, Any]]:
    if load_workbook is None:
        raise ValueError("Partner import requires openpyxl.")
    workbook = load_workbook(upload.stream, data_only=True)
    if "Partners" not in workbook.sheetnames:
        raise ValueError("The workbook must contain a Partners sheet.")
    sheet = workbook["Partners"]
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    header_index = find_partner_header_row(rows)
    headers = [PARTNER_HEADER_ALIASES.get(normalize_header(value), "") for value in rows[header_index]]
    imported_rows = []
    for row in rows[header_index + 1 :]:
        mapped: dict[str, Any] = {}
        for index, value in enumerate(row[: len(headers)]):
            field = headers[index]
            if field:
                mapped[field] = clean_partner_value(value)
        if any(clean_text(value) for value in mapped.values()):
            imported_rows.append(mapped)
    return imported_rows


def find_partner_header_row(rows: list[tuple[Any, ...]]) -> int:
    for index, row in enumerate(rows[:20]):
        normalized = {normalize_header(value) for value in row if clean_text(value)}
        if {"partner", "owner", "territory", "workbook link"}.issubset(normalized):
            return index
    return 0


def upsert_partner_rows(database_path: str, rows: list[dict[str, Any]], user: str, filename: str) -> tuple[int, int, list[dict[str, Any]]]:
    db = get_db(database_path)
    created = 0
    updated = 0
    errors: list[dict[str, Any]] = []
    now = utc_now()
    for row_number, row in enumerate(rows, start=2):
        partner_name = clean_text(row.get("partner_name"))
        if not partner_name:
            errors.append({"row": row_number, "errors": ["Partner name is required."]})
            continue
        values = {
            "partner_name": partner_name,
            "owner": clean_text(row.get("owner")),
            "territory": clean_text(row.get("territory")),
            "master_stage": clean_text(row.get("master_stage")),
            "date_of_stage_change": clean_date(row.get("date_of_stage_change")),
            "age_of_stage": clean_int(row.get("age_of_stage")),
            "days_overdue": clean_int(row.get("days_overdue")),
            "workbook_link": clean_text(row.get("workbook_link")),
            "sf_account": clean_text(row.get("sf_account")),
            "next_steps_notes": clean_text(row.get("next_steps_notes")),
            "partner_type": clean_text(row.get("partner_type")),
            "csm_involved": clean_text(row.get("csm_involved")),
        }
        existing = db.execute("SELECT id FROM partners WHERE lower(partner_name) = lower(?)", (partner_name,)).fetchone()
        if existing:
            db.execute(
                f"""
                UPDATE partners
                SET {', '.join([field + ' = ?' for field in PARTNER_FIELDS])}, updated_at = ?
                WHERE id = ?
                """,
                (*[values[field] for field in PARTNER_FIELDS], now, existing["id"]),
            )
            updated += 1
        else:
            db.execute(
                f"""
                INSERT INTO partners ({', '.join(PARTNER_FIELDS)}, created_at, updated_at)
                VALUES ({', '.join(['?'] * len(PARTNER_FIELDS))}, ?, ?)
                """,
                (*[values[field] for field in PARTNER_FIELDS], now, now),
            )
            created += 1
    db.execute(
        """
        INSERT INTO partner_import_history
            (filename, imported_at, imported_by, created_count, updated_count, error_count, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (filename, now, user, created, updated, len(errors), json.dumps({"errors": errors[:25]})),
    )
    db.commit()
    db.close()
    return created, updated, errors


def clean_partner_value(value: Any) -> str:
    text = clean_text(value)
    if text.startswith("#") and text.endswith("?"):
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return text


def clean_int(value: Any) -> int | None:
    text = clean_partner_value(value)
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def is_url(value: str) -> bool:
    text = clean_text(value).lower()
    return text.startswith("http://") or text.startswith("https://")


def partner_import_history_to_dict(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "filename": row["filename"] or "",
        "importedAt": row["imported_at"],
        "importedBy": row["imported_by"] or "",
        "createdCount": row["created_count"],
        "updatedCount": row["updated_count"],
        "errorCount": row["error_count"],
        "notes": row["notes"] or "",
    }


def parse_org_impact_file(upload: Any) -> dict[str, list[dict[str, Any]]]:
    if load_workbook is None:
        raise ValueError("Org Impact import requires openpyxl.")
    workbook = load_workbook(upload.stream, data_only=True)
    sheet_name = next((name for name in workbook.sheetnames if normalize_header(name) == "org impact"), "")
    if not sheet_name:
        return {"parameters": [], "assumptions": [], "averages": [], "summary": []}
    sheet = workbook[sheet_name]
    parameters = []
    for row in range(4, 12):
        label = clean_text(sheet.cell(row, 4).value)
        value = clean_number_or_text(sheet.cell(row, 3).value)
        if label:
            parameters.append({"label": label, "value": value})
    assumptions = parse_org_impact_matrix(sheet, 16, 26, "baseAssumptions", [
        "recruitmentDays", "onboardingDays", "activeDays", "recruitmentWeeks", "onboardingWeeks", "activeWeeks"
    ])
    averages = parse_org_impact_matrix(sheet, 31, 41, "estimatedAverages", [
        "recruitmentHours", "onboardingHours", "activeHours", "recruitmentCost", "onboardingCost", "activeCost"
    ])
    summary = []
    for row in range(45, 50):
        stage = clean_text(sheet.cell(row, 3).value)
        if stage:
            summary.append({
                "label": stage,
                "countInStage": clean_number_or_text(sheet.cell(row, 4).value),
                "totalHours": clean_number_or_text(sheet.cell(row, 5).value),
                "totalDays": clean_number_or_text(sheet.cell(row, 6).value),
                "totalYears": clean_number_or_text(sheet.cell(row, 7).value),
            })
    return {"parameters": parameters, "assumptions": assumptions, "averages": averages, "summary": summary}


def parse_org_impact_matrix(sheet: Any, start_row: int, end_row: int, section: str, keys: list[str]) -> list[dict[str, Any]]:
    rows = []
    for row in range(start_row, end_row + 1):
        label = clean_text(sheet.cell(row, 3).value)
        if not label:
            continue
        values = {"label": label, "section": section}
        for offset, key in enumerate(keys, start=4):
            values[key] = clean_number_or_text(sheet.cell(row, offset).value)
        rows.append(values)
    return rows


def save_org_impact(database_path: str, org_impact: dict[str, list[dict[str, Any]]]) -> int:
    now = utc_now()
    db = get_db(database_path)
    db.execute("DELETE FROM org_impact_rows")
    inserted = 0
    sort_order = 0
    for section, rows in org_impact.items():
        for row in rows:
            label = clean_text(row.get("label"))
            if not label:
                continue
            sort_order += 1
            db.execute(
                """
                INSERT INTO org_impact_rows (section, row_label, values_json, sort_order, imported_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (section, label, json.dumps(row, sort_keys=True), sort_order, now),
            )
            inserted += 1
    db.commit()
    db.close()
    return inserted


def load_org_impact(database_path: str) -> dict[str, Any]:
    db = get_db(database_path)
    rows = db.execute("SELECT * FROM org_impact_rows ORDER BY sort_order").fetchall()
    db.close()
    grouped = {"parameters": [], "assumptions": [], "averages": [], "summary": []}
    imported_at = ""
    for row in rows:
        imported_at = row["imported_at"]
        try:
            values = json.loads(row["values_json"])
        except json.JSONDecodeError:
            values = {"label": row["row_label"]}
        section = row["section"]
        if section == "baseAssumptions":
            grouped["assumptions"].append(values)
        elif section == "estimatedAverages":
            grouped["averages"].append(values)
        else:
            grouped.setdefault(section, []).append(values)
    grouped["importedAt"] = imported_at
    grouped["totals"] = next((row for row in grouped["summary"] if clean_text(row.get("label")).lower() == "total"), {})
    return grouped


def clean_number_or_text(value: Any) -> Any:
    text = clean_partner_value(value)
    if text == "":
        return ""
    try:
        number = float(text)
        if number.is_integer():
            return int(number)
        return round(number, 2)
    except (ValueError, TypeError):
        return text


if __name__ == "__main__":
    app = create_app()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=os.environ.get("FLASK_DEBUG") == "1")
