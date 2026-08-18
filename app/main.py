"""Flask application: public check-in kiosk + staff admin dashboard."""

from __future__ import annotations

import io
import os
from datetime import date, datetime

from flask import (
    Flask, abort, flash, redirect, render_template, request,
    send_file, send_from_directory, url_for, g,
)
from werkzeug.utils import secure_filename

from .database import DEFAULT_DB_PATH, get_connection, init_db
from .mapping import (
    HOUSEHOLD_TYPE_CHOICES, POPULATION_CHOICES, RACE_CHOICES,
    VISITOR_TYPE_CHOICES, suggest_household_type, suggest_population_category,
)
from .models import normalize_name, parse_dob
from .services import checkin_service, client_service, excel_service, intake_service

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UPLOAD_DIR = os.path.join(BASE_DIR, "data", "uploads")


def create_app(db_path: str = DEFAULT_DB_PATH) -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("SECRET_KEY", "employment-center-local-dev")
    app.config["DB_PATH"] = db_path
    app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB workbook cap
    init_db(db_path)
    os.makedirs(UPLOAD_DIR, exist_ok=True)

    # ---- request-scoped DB connection -----------------------------------
    def db():
        if "conn" not in g:
            g.conn = get_connection(app.config["DB_PATH"])
        return g.conn

    @app.teardown_appcontext
    def _close_db(exc):
        conn = g.pop("conn", None)
        if conn is not None:
            if exc is None:
                conn.commit()
            conn.close()

    # ---- template helpers ----------------------------------------------
    @app.context_processor
    def _inject():
        return {
            "RACE_CHOICES": RACE_CHOICES,
            "HOUSEHOLD_TYPE_CHOICES": HOUSEHOLD_TYPE_CHOICES,
            "POPULATION_CHOICES": POPULATION_CHOICES,
            "VISITOR_TYPE_CHOICES": VISITOR_TYPE_CHOICES,
            "now": datetime.now(),
        }

    app.jinja_env.filters["fmt_dob"] = lambda v: (
        parse_dob(v).strftime("%m/%d/%Y") if parse_dob(v) else ""
    )
    import calendar as _cal
    app.jinja_env.filters["month_name"] = lambda m: (
        _cal.month_name[int(m)] if m else ""
    )

    # =====================================================================
    # Public kiosk
    # =====================================================================

    @app.route("/")
    def home():
        return render_template("checkin.html")

    @app.route("/identify", methods=["POST"])
    def identify():
        first = (request.form.get("first_name") or "").strip()
        last = (request.form.get("last_name") or "").strip()
        dob_raw = (request.form.get("date_of_birth") or "").strip()
        dob = parse_dob(dob_raw)
        if not first or not last or not dob:
            flash("Please enter first name, last name, and date of birth.")
            return redirect(url_for("home"))

        dob_iso = dob.isoformat()
        matches = client_service.find_matches(db(), first, last, dob_iso)

        if not matches:
            # New client -> full intake
            return redirect(url_for("intake_form", first=first, last=last, dob=dob_iso))
        if len(matches) > 1:
            return render_template("select_client.html", matches=matches)
        return _route_existing_client(matches[0]["id"])

    def _route_existing_client(client_id: int):
        today = date.today()
        client = client_service.get_client(db(), client_id)
        if client is None:
            abort(404)
        done = intake_service.has_completed_intake(
            db(), client_id, today.year, today.month)
        if done:
            return render_template("welcome_back.html", client=client)
        return redirect(url_for("intake_form", client_id=client_id))

    @app.route("/continue/<int:client_id>")
    def continue_client(client_id):
        return _route_existing_client(client_id)

    # ---- intake form ----------------------------------------------------
    @app.route("/intake")
    def intake_form():
        client_id = request.args.get("client_id", type=int)
        prefill = {"first_name": "", "last_name": "", "date_of_birth": "",
                   "uid": ""}
        intake = {}
        dob = None

        if client_id:
            client = client_service.get_client(db(), client_id)
            if client is None:
                abort(404)
            prefill.update({
                "first_name": client["first_name"],
                "last_name": client["last_name"],
                "date_of_birth": client["date_of_birth"],
                "uid": client["uid"] or "",
            })
            dob = parse_dob(client["date_of_birth"])
            latest = intake_service.get_latest_intake(db(), client_id)
            if latest is not None:
                intake = dict(latest)  # prefill from most recent month
        else:
            prefill.update({
                "first_name": request.args.get("first", ""),
                "last_name": request.args.get("last", ""),
                "date_of_birth": request.args.get("dob", ""),
            })
            dob = parse_dob(prefill["date_of_birth"])

        defaults = {
            "population_category": intake.get("population_category")
            or suggest_population_category(dob),
            "household_type": intake.get("household_type")
            or suggest_household_type(dob),
        }
        return render_template("intake.html", client_id=client_id,
                               prefill=prefill, intake=intake, defaults=defaults)

    @app.route("/intake/submit", methods=["POST"])
    def intake_submit():
        first = (request.form.get("first_name") or "").strip()
        last = (request.form.get("last_name") or "").strip()
        dob = parse_dob(request.form.get("date_of_birth"))
        uid = (request.form.get("uid") or "").strip()
        client_id = request.form.get("client_id", type=int)

        if not first or not last or not dob:
            flash("First name, last name, and date of birth are required.")
            return redirect(url_for("intake_form", client_id=client_id))

        dob_iso = dob.isoformat()
        conn = db()

        # Create or update the permanent client record.
        if not client_id:
            existing = client_service.find_matches(conn, first, last, dob_iso)
            if existing:
                client_id = existing[0]["id"]
                client_service.update_client(conn, client_id, first, last, dob_iso, uid)
            else:
                client_id = client_service.create_client(conn, first, last, dob_iso, uid)
        else:
            client_service.update_client(conn, client_id, first, last, dob_iso, uid)

        data = _parse_intake_form(request.form)
        today = date.today()
        intake_id, _created = intake_service.create_or_get_intake(
            conn, client_id, today.year, today.month, data)

        checkin_service.create_visit(
            conn, client_id, monthly_intake_id=intake_id,
            visitor_type=request.form.get("visitor_type"),
            services=request.form.get("services"))
        conn.commit()

        _safe_sync(conn, today.year, today.month)
        return redirect(url_for("done", name=first))

    # ---- simple check-in (returning client, intake already done) --------
    @app.route("/checkin/<int:client_id>", methods=["POST"])
    def checkin(client_id):
        conn = db()
        client = client_service.get_client(conn, client_id)
        if client is None:
            abort(404)
        today = date.today()
        intake = intake_service.get_intake_for_month(
            conn, client_id, today.year, today.month)
        if intake is None:
            # Safety: no intake this month -> send them through intake instead.
            return redirect(url_for("intake_form", client_id=client_id))
        checkin_service.create_visit(
            conn, client_id, monthly_intake_id=intake["id"],
            visitor_type=request.form.get("visitor_type"))
        conn.commit()
        _safe_sync(conn, today.year, today.month)
        return redirect(url_for("done", name=client["first_name"]))

    @app.route("/done")
    def done():
        return render_template("done.html", name=request.args.get("name", ""))

    # =====================================================================
    # Admin
    # =====================================================================

    @app.route("/admin")
    def admin_dashboard():
        stats = _dashboard_stats(db())
        return render_template("admin/dashboard.html", stats=stats)

    @app.route("/admin/clients")
    def admin_clients():
        q = (request.args.get("q") or "").strip()
        results = client_service.search_clients(db(), q) if q else []
        return render_template("admin/clients.html", q=q, results=results)

    @app.route("/admin/clients/<int:client_id>")
    def admin_client_detail(client_id):
        conn = db()
        client = client_service.get_client(conn, client_id)
        if client is None:
            abort(404)
        intakes = intake_service.list_intakes(conn, client_id)
        visits = checkin_service.list_visits(conn, client_id)
        return render_template("admin/client_detail.html", client=client,
                               intakes=intakes, visits=visits)

    @app.route("/admin/clients/<int:client_id>/edit", methods=["POST"])
    def admin_client_edit(client_id):
        conn = db()
        client_service.update_client(
            conn, client_id,
            request.form.get("first_name", ""), request.form.get("last_name", ""),
            parse_dob(request.form.get("date_of_birth")).isoformat()
            if parse_dob(request.form.get("date_of_birth")) else client_service
            .get_client(conn, client_id)["date_of_birth"],
            request.form.get("uid", ""))
        conn.commit()
        flash("Client record updated.")
        return redirect(url_for("admin_client_detail", client_id=client_id))

    @app.route("/admin/intake/<int:intake_id>/edit", methods=["POST"])
    def admin_intake_edit(intake_id):
        conn = db()
        row = conn.execute(
            "SELECT * FROM monthly_intakes WHERE id = ?", (intake_id,)).fetchone()
        if row is None:
            abort(404)
        data = _parse_intake_form(request.form)
        intake_service.update_intake(conn, intake_id, data)
        conn.commit()
        _safe_sync(conn, row["reporting_year"], row["reporting_month"])
        flash("Monthly intake corrected.")
        return redirect(url_for("admin_client_detail", client_id=row["client_id"]))

    # ---- Excel export / sync / upload -----------------------------------
    @app.route("/admin/sync", methods=["POST"])
    def admin_sync():
        today = date.today()
        result = excel_service.sync_default_workbook(db(), today.year, today.month)
        flash(f"Synced {result['intake_rows']} intakes and "
              f"{result['visit_rows']} visits into "
              f"'{result['intake_sheet']}' / '{result['signin_sheet']}'.")
        return redirect(url_for("admin_dashboard"))

    @app.route("/admin/export")
    def admin_export():
        today = date.today()
        result = excel_service.sync_default_workbook(db(), today.year, today.month)
        return send_file(result["path"], as_attachment=True,
                         download_name=os.path.basename(result["path"]))

    @app.route("/admin/upload", methods=["POST"])
    def admin_upload():
        file = request.files.get("workbook")
        if not file or not file.filename:
            flash("Please choose a workbook (.xlsx) to upload.")
            return redirect(url_for("admin_dashboard"))
        if not file.filename.lower().endswith(".xlsx"):
            flash("Only .xlsx workbooks are supported.")
            return redirect(url_for("admin_dashboard"))

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        src = os.path.join(UPLOAD_DIR, f"upload_{stamp}_{secure_filename(file.filename)}")
        file.save(src)

        today = date.today()
        # optional month override from the form (YYYY-MM)
        ym = (request.form.get("month") or "").strip()
        year, month = today.year, today.month
        if ym:
            try:
                year, month = int(ym[:4]), int(ym[5:7])
            except (ValueError, IndexError):
                pass

        dest = os.path.join(UPLOAD_DIR, f"updated_{stamp}_{secure_filename(file.filename)}")
        excel_service.apply_to_uploaded_workbook(db(), src, dest, year, month)
        return send_file(dest, as_attachment=True, download_name=file.filename)

    @app.route("/backups/<path:name>")
    def download_backup(name):
        return send_from_directory(excel_service.BACKUP_DIR, name, as_attachment=True)

    # =====================================================================
    # Internal helpers
    # =====================================================================

    def _safe_sync(conn, year, month):
        """Sync to the working workbook but never let an Excel error break
        the kiosk flow (the DB remains the source of truth)."""
        try:
            excel_service.sync_default_workbook(conn, year, month)
        except Exception as exc:  # pragma: no cover - defensive
            app.logger.warning("Excel sync failed: %s", exc)

    return app


# ---------------------------------------------------------------------------
# Form parsing / stats (module-level, no app state)
# ---------------------------------------------------------------------------

def _parse_intake_form(form) -> dict:
    def flag(name):
        return 1 if form.get(name) in ("1", "on", "true", "yes") else 0

    house_size = form.get("house_size", type=int)
    return {
        "veteran": flag("veteran"),
        "hispanic": flag("hispanic"),
        "race": form.get("race") or None,
        "house_size": house_size if house_size and house_size > 0 else None,
        "household_type": form.get("household_type") or None,
        "population_category": form.get("population_category") or None,
        "female_head": flag("female_head"),
        "disabled": flag("disabled"),
    }


def _dashboard_stats(conn) -> dict:
    today = date.today()
    today_iso = today.isoformat()
    y, m = today.year, today.month

    def scalar(sql, params=()):
        return conn.execute(sql, params).fetchone()[0]

    return {
        "today": today.strftime("%A, %B %d, %Y"),
        "visits_today": scalar(
            "SELECT COUNT(*) FROM visits WHERE visit_date = ?", (today_iso,)),
        "unique_clients_today": scalar(
            "SELECT COUNT(DISTINCT client_id) FROM visits WHERE visit_date = ?",
            (today_iso,)),
        "unique_clients_month": scalar(
            "SELECT COUNT(DISTINCT client_id) FROM visits "
            "WHERE reporting_year = ? AND reporting_month = ?", (y, m)),
        "visits_month": scalar(
            "SELECT COUNT(*) FROM visits "
            "WHERE reporting_year = ? AND reporting_month = ?", (y, m)),
        "new_clients_month": scalar(
            "SELECT COUNT(*) FROM clients "
            "WHERE substr(created_at,1,7) = ?", (f"{y:04d}-{m:02d}",)),
        "completed_intakes_month": scalar(
            "SELECT COUNT(*) FROM monthly_intakes "
            "WHERE reporting_year = ? AND reporting_month = ?", (y, m)),
        "intake_required_month": scalar(
            "SELECT COUNT(DISTINCT v.client_id) FROM visits v "
            "WHERE v.reporting_year = ? AND v.reporting_month = ? "
            "AND NOT EXISTS (SELECT 1 FROM monthly_intakes mi "
            "  WHERE mi.client_id = v.client_id AND mi.reporting_year = ? "
            "  AND mi.reporting_month = ?)", (y, m, y, m)),
        "month_label": today.strftime("%B %Y"),
    }
