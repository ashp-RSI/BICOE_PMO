import datetime
import logging
import re
from datetime import date

import numpy as np
import pandas as pd
from flask import Blueprint, current_app, jsonify, render_template, request
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

from app import cache
from app.services.sharepoint_service import SharePointService

logger = logging.getLogger(__name__)
api_bp = Blueprint("api", __name__)

_sp_service = None

BENCH_STATUSES = {"Non-Billable"}


def _get_service() -> SharePointService:
    global _sp_service
    if _sp_service is None:
        _sp_service = SharePointService(current_app.config)
    return _sp_service


def _normalize_emp_code(raw):
    """Normalize an emp code value (which may be float like 19027.0) to a
    clean string like '19027' for reliable comparison."""
    if raw is None:
        return ""
    if isinstance(raw, float) and np.isnan(raw):
        return ""
    s = str(raw).strip()
    if not s or s.lower() == "nan":
        return ""
    try:
        return str(int(float(s)))
    except (ValueError, OverflowError):
        return s


def _clean_value(v):
    """Convert numpy/pandas types to JSON-safe Python types."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if hasattr(v, "isoformat"):
        if pd.isna(v):
            return None
        return v.isoformat()
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        if np.isnan(v):
            return None
        return float(v)
    return v


def _parse_experience_range(value):
    """Parse experience range strings into (min, max) floats.
    Handles: '5-10', '5 to 10', '5+', '5', '5-10 yrs', None/empty."""
    if not value or (isinstance(value, float) and np.isnan(value)):
        return (0, 99)

    text = str(value).lower().replace("years", "").replace("yrs", "").strip()

    match = re.match(r"(\d+\.?\d*)\s*[-–]\s*(\d+\.?\d*)", text)
    if match:
        return (float(match.group(1)), float(match.group(2)))

    match = re.match(r"(\d+\.?\d*)\s+to\s+(\d+\.?\d*)", text)
    if match:
        return (float(match.group(1)), float(match.group(2)))

    if text.endswith("+"):
        try:
            return (float(text[:-1].strip()), 99)
        except ValueError:
            return (0, 99)

    try:
        v = float(text)
        return (v, v)
    except ValueError:
        return (0, 99)


def _tokenize_skills(skill_str):
    """Split a comma/slash-separated skill string into normalized lowercase tokens.
    Removes internal spaces so 'Power BI' and 'PowerBI' both become 'powerbi'."""
    if not skill_str or (isinstance(skill_str, float) and np.isnan(skill_str)):
        return set()
    return {re.sub(r"\s+", "", t).lower()
            for t in re.split(r"[,/;|]+", str(skill_str)) if t.strip()}


def _parse_employee_experience(raw_value):
    """Convert the Experience column (Y:M:D format) to total years as a float.
    Excel may store Y:M:D as datetime.time (H:M:S), datetime.datetime,
    pd.Timestamp, timedelta, string, or number depending on cell formatting."""
    if raw_value is None:
        return None
    if isinstance(raw_value, float) and np.isnan(raw_value):
        return None
    if isinstance(raw_value, (int, float)):
        return round(float(raw_value), 2)

    if isinstance(raw_value, datetime.time):
        return round(raw_value.hour + raw_value.minute / 12
                     + raw_value.second / 365, 2)

    if isinstance(raw_value, datetime.datetime):
        return round(raw_value.hour + raw_value.minute / 12
                     + raw_value.second / 365, 2)

    if isinstance(raw_value, pd.Timestamp):
        return round(raw_value.hour + raw_value.minute / 12
                     + raw_value.second / 365, 2)

    if isinstance(raw_value, datetime.timedelta):
        total_secs = raw_value.total_seconds()
        h = int(total_secs // 3600)
        m = int((total_secs % 3600) // 60)
        s = int(total_secs % 60)
        return round(h + m / 12 + s / 365, 2)

    text = str(raw_value).strip()
    if not text:
        return None

    # Handle "1900-01-01 04:03:23" style strings pandas sometimes produces
    if " " in text and "-" in text.split(" ")[0]:
        time_part = text.split(" ")[-1]
        tparts = time_part.split(":")
        if len(tparts) >= 2:
            try:
                h = int(tparts[0])
                m = int(tparts[1])
                s = int(tparts[2]) if len(tparts) > 2 else 0
                return round(h + m / 12 + s / 365, 2)
            except ValueError:
                pass

    parts = text.split(":")
    if len(parts) == 3:
        try:
            return round(int(parts[0]) + int(parts[1]) / 12
                         + int(parts[2]) / 365, 2)
        except ValueError:
            pass
    if len(parts) == 2:
        try:
            return round(int(parts[0]) + int(parts[1]) / 12, 2)
        except ValueError:
            pass

    try:
        return round(float(text), 2)
    except ValueError:
        logger.warning("Could not parse experience value: %r (type=%s)",
                       raw_value, type(raw_value).__name__)
        return None


CACHE_TIMEOUT = 900


def _get_cached_df():
    key = "headcount_df"
    df = cache.get(key)
    if df is None:
        sp = _get_service()
        df = sp.get_dataframe()
        cache.set(key, df, timeout=CACHE_TIMEOUT)
    return df


def _get_cached_demand_df():
    key = "demand_df"
    df = cache.get(key)
    if df is None:
        sp = _get_service()
        sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")
        df = sp.get_demand_dataframe(sheet)
        cache.set(key, df, timeout=CACHE_TIMEOUT)
    return df


@api_bp.route("/headcount", methods=["GET"])
def get_headcount():
    """Return head count data with optional filters and pagination."""
    try:
        df = _get_cached_df()

        sub_practice = request.args.get("sub_practice")
        billable = request.args.get("billable")
        project = request.args.get("project")
        kpi = request.args.get("kpi")
        search_term = request.args.get("search", "").strip().lower()
        page = int(request.args.get("page", 1))
        per_page = int(request.args.get("per_page", 25))

        if sub_practice and sub_practice != "All":
            df = df[df["Sub Practice"] == sub_practice]
        if billable and billable != "All":
            df = df[df["Billable/Non Billable"] == billable]
        if project and project != "All" and "Projects" in df.columns:
            df = df[df["Projects"].fillna("").astype(str).str.strip() == project.strip()]

        if "Status" in df.columns:
            df = df[df["Status"].fillna("").astype(str).str.strip() != "Resigned"]
        if "Billable/Non Billable" in df.columns:
            df = df[df["Billable/Non Billable"].fillna("").astype(str).str.strip() != "Resigned"]

        if kpi and kpi != "all":
            if kpi == "other":
                known_statuses = {
                    "Billable", "Non-Billable", "Blocked", "Proposed",
                    "Internal project", "Solution Offerings",
                }
                df = df[~df["Billable/Non Billable"].fillna("").astype(str).str.strip().isin(known_statuses)]
            elif kpi.startswith("sp-"):
                sp_value = kpi[3:]
                df = df[df["Sub Practice"] == sp_value]
            else:
                df = df[df["Billable/Non Billable"].fillna("").astype(str).str.strip() == kpi]
        if search_term:
            mask = df.apply(
                lambda row: row.astype(str).str.lower().str.contains(search_term).any(),
                axis=1,
            )
            df = df[mask]

        total = len(df)
        start = (page - 1) * per_page
        end = start + per_page
        page_df = df.iloc[start:end]

        notif_map = _get_notif_map_for_codes(
            page_df["Emp Code"].tolist() if "Emp Code" in page_df.columns else []
        )

        records = []
        for idx, row in page_df.iterrows():
            record = {"row_index": int(idx)}
            for col in df.columns:
                record[col] = _clean_value(row[col])
            emp_code = _normalize_emp_code(row.get("Emp Code"))
            if emp_code and emp_code in notif_map:
                record["_notification"] = notif_map[emp_code]
            records.append(record)

        return jsonify({
            "data": records,
            "total": total,
            "page": page,
            "per_page": per_page,
            "total_pages": max(1, -(-total // per_page)),
        })
    except Exception as e:
        logger.exception("Error fetching headcount data")
        return jsonify({"error": str(e)}), 500


def _get_notif_map_for_codes(emp_codes):
    """Return {emp_code: {status, sent_at, reminder_count, id}} for the
    most recent notification of each emp_code in the list. Used to attach
    badge state to headcount rows in a single SQLite read."""
    try:
        from app import notification_store
        if notification_store is None or not emp_codes:
            return {}
        all_active = notification_store.list_all_active_codes()
        normalized_codes = {
            _normalize_emp_code(c) for c in emp_codes if c is not None
        }
        return {
            code: data for code, data in all_active.items()
            if code in normalized_codes
        }
    except Exception:
        logger.exception("Failed to attach notification state to headcount")
        return {}


@api_bp.route("/filters", methods=["GET"])
def get_filters():
    """Return distinct values for filter dropdowns."""
    try:
        df = _get_cached_df()
        sub_practices = sorted(df["Sub Practice"].dropna().unique().tolist())
        billable_vals = sorted(
            df["Billable/Non Billable"].dropna().unique().tolist()
        )
        projects = sorted(
            df["Projects"].dropna().astype(str).str.strip().loc[lambda s: s != ""].unique().tolist()
        ) if "Projects" in df.columns else []
        return jsonify({
            "sub_practices": sub_practices,
            "billable_options": billable_vals,
            "projects": projects,
        })
    except Exception as e:
        logger.exception("Error fetching filters")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/summary", methods=["GET"])
def get_summary():
    """Return aggregated stats for dashboard cards and charts."""
    try:
        df = _get_cached_df()

        sub_practice = request.args.get("sub_practice")
        billable = request.args.get("billable")
        project = request.args.get("project")
        kpi = request.args.get("kpi")

        if sub_practice and sub_practice != "All":
            df = df[df["Sub Practice"] == sub_practice]
        if billable and billable != "All":
            df = df[df["Billable/Non Billable"] == billable]
        if project and project != "All" and "Projects" in df.columns:
            df = df[df["Projects"].fillna("").astype(str).str.strip() == project.strip()]

        if "Status" in df.columns:
            df = df[df["Status"].fillna("").astype(str).str.strip() != "Resigned"]
        if "Billable/Non Billable" in df.columns:
            df = df[df["Billable/Non Billable"].fillna("").astype(str).str.strip() != "Resigned"]

        if kpi and kpi != "all":
            if kpi == "other":
                known_statuses = {
                    "Billable", "Non-Billable", "Blocked", "Proposed",
                    "Internal project", "Solution Offerings",
                }
                df = df[~df["Billable/Non Billable"].fillna("").astype(str).str.strip().isin(known_statuses)]
            elif kpi.startswith("sp-"):
                sp_value = kpi[3:]
                df = df[df["Sub Practice"] == sp_value]
            else:
                df = df[df["Billable/Non Billable"].fillna("").astype(str).str.strip() == kpi]

        total = len(df)

        bl_col = df["Billable/Non Billable"].fillna("").astype(str).str.strip()

        billable_count = int((bl_col == "Billable").sum())
        non_billable_count = int((bl_col == "Non-Billable").sum())
        blocked_count = int((bl_col == "Blocked").sum())
        proposed_count = int((bl_col == "Proposed").sum())
        internal_project_count = int((bl_col == "Internal project").sum())
        solution_offering_count = int((bl_col == "Solution Offerings").sum())
        other_count = total - billable_count - non_billable_count - blocked_count - proposed_count - internal_project_count - solution_offering_count

        sp_col = df["Sub Practice"].fillna("").astype(str).str.strip()
        data_count = int((sp_col == "DE").sum())
        ai_count = int((sp_col == "AI").sum())
        bi_count = int((sp_col == "BI").sum())
        core_count = int((sp_col == "Core").sum())
        not_confirmed_count = int((sp_col == "Not Confirmed").sum())

        return jsonify({
            "total": total,
            "data_count": data_count,
            "ai_count": ai_count,
            "bi_count": bi_count,
            "core_count": core_count,
            "not_confirmed_count": not_confirmed_count,
            "billable": billable_count,
            "non_billable": non_billable_count,
            "blocked": blocked_count,
            "proposed": proposed_count,
            "internal_project_count": internal_project_count,
            "solution_offering_count": solution_offering_count,
            "other": other_count,
        })
    except Exception as e:
        logger.exception("Error fetching summary")
        return jsonify({"error": str(e)}), 500


def _sync_demand_status_for_employee(hc_row_index, new_billable_status,
                                     old_billable_status=None):
    """When a mapped employee's Billable/Non Billable changes, reverse-sync
    the corresponding demand requisition's Demand Status.

    Three transitions are handled:
      1. Employee no longer Billable while demand is Fulfilled
         → revert demand to "In Progress".
      2. Employee becomes Billable while demand is In Progress
         → mark demand "Fulfilled".
      3. (NEW) Employee leaves "Proposed" without becoming Billable
         (e.g. set back to Non-Billable / Blocked / etc.)
         → fully un-map the demand: clear Mapped Emp Code/Name/Date,
           reset Fulfillment Type, set Demand Status to "Open".
         This prevents ghost mappings where a demand row keeps pointing
         at an employee whose status no longer reflects the proposal.
    """
    try:
        hc_df = _get_cached_df()
        if hc_row_index not in hc_df.index:
            return
        emp_code = _normalize_emp_code(hc_df.loc[hc_row_index].get("Emp Code"))
        if not emp_code:
            return

        demand_df = _get_cached_demand_df()
        if demand_df.empty or "Mapped Emp Code" not in demand_df.columns:
            return

        matched = demand_df[
            demand_df["Mapped Emp Code"].apply(_normalize_emp_code) == emp_code
        ]
        if matched.empty:
            return

        sp = _get_service()
        sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")

        new_norm = (new_billable_status or "").strip()
        old_norm = (old_billable_status or "").strip()

        for demand_idx, demand_row in matched.iterrows():
            ds = str(demand_row.get("Demand Status", "")).strip()

            if new_norm != "Billable" and ds == "Fulfilled":
                sp.update_demand_row(sheet, int(demand_idx),
                                     {"Demand Status": "In Progress"})
                logger.info("Demand row %d reverted to 'In Progress' "
                            "(employee %s no longer Billable)",
                            demand_idx, emp_code)
            elif new_norm == "Billable" and ds == "In Progress":
                sp.update_demand_row(sheet, int(demand_idx),
                                     {"Demand Status": "Fulfilled"})
                logger.info("Demand row %d set to 'Fulfilled' "
                            "(employee %s now Billable)",
                            demand_idx, emp_code)
            elif (
                old_norm == "Proposed"
                and new_norm not in ("Proposed", "Billable")
                and ds in ("In Progress", "Open", "")
            ):
                sp.update_demand_row(sheet, int(demand_idx), {
                    "Demand Status": "Open",
                    "Mapped Emp Code": "",
                    "Mapped Emp Name": "",
                    "Mapping Date": "",
                    "Fulfillment Type": "",
                })
                logger.info("Demand row %d un-mapped (employee %s left "
                            "'Proposed' for '%s')",
                            demand_idx, emp_code, new_norm)
                # Cancel any active manager notification for this employee
                _cancel_active_notification_for_emp_code(
                    emp_code,
                    note=f"Mapping reverted: employee status changed to '{new_norm}'",
                )

        cache.delete("demand_df")
    except Exception:
        logger.exception("Error syncing demand status for HC row %d", hc_row_index)


def _cancel_active_notification_for_emp_code(emp_code, note=""):
    """If there's an active manager notification for this emp_code, cancel
    it so reminders stop firing for a mapping that no longer exists."""
    try:
        from app import notification_store
        from app.services.notification_store import STATUS_CANCELLED
        if notification_store is None:
            return
        active = notification_store.get_active_for_employee(emp_code)
        if active:
            notification_store.update_status(
                active["id"], STATUS_CANCELLED,
                decided_by="system", decision_note=note,
            )
            logger.info("Cancelled active notification %d for emp_code %s "
                        "(reason: %s)", active["id"], emp_code, note)
    except Exception:
        logger.exception("Failed to cancel active notification for %s",
                         emp_code)


@api_bp.route("/headcount/<int:row_index>", methods=["PUT"])
def update_headcount(row_index):
    """Update a single row and sync back to SharePoint."""
    try:
        updates = request.get_json()
        if not updates:
            return jsonify({"error": "No update data provided"}), 400

        protected_cols = {"S.No", "Emp Code", "row_index"}
        clean_updates = {
            k: v for k, v in updates.items() if k not in protected_cols
        }

        if not clean_updates:
            return jsonify({"error": "No valid fields to update"}), 400

        if clean_updates.get("Billable/Non Billable") == "Billable":
            clean_updates.setdefault("Customer interview happened(Yes/No)", "Yes")
            clean_updates.setdefault("Customer Selected(Yes/No)", "Yes")

        df = _get_cached_df()
        if row_index in df.index:
            row = df.loc[row_index]
            interview = clean_updates.get(
                "Customer interview happened(Yes/No)",
                str(row.get("Customer interview happened(Yes/No)", "")).strip(),
            )
            selected = clean_updates.get(
                "Customer Selected(Yes/No)",
                str(row.get("Customer Selected(Yes/No)", "")).strip(),
            )
            if interview == "Yes" and selected == "Yes":
                clean_updates.setdefault("Billable/Non Billable", "Billable")

        old_billable = ""
        if row_index in df.index:
            old_billable = str(
                df.loc[row_index].get("Billable/Non Billable", "") or ""
            ).strip()

        sp = _get_service()
        sp.update_row(row_index, clean_updates)

        cache.delete("headcount_df")

        if "Billable/Non Billable" in clean_updates:
            _sync_demand_status_for_employee(
                row_index,
                clean_updates["Billable/Non Billable"],
                old_billable_status=old_billable,
            )

        return jsonify({"success": True, "message": "Row updated successfully"})
    except Exception as e:
        logger.exception("Error updating row %d", row_index)
        return jsonify({"error": str(e)}), 500


HEADCOUNT_ADD_FIELDS = [
    "Division", "Emp Code", "Emp Name", "Status", "LWD", "Skills",
    "Fresher/Lateral", "Offshore/Onsite", "Experience", "Designation",
    "Grade", "DOJ", "Gender", "First Line Manager", "Skip Level Manager",
    "Company Email", "Sub Practice", "Remarks", "Empower SL",
    "Billable/Non Billable", "Billable Till Date", "Projects", "Remarks2",
    "Customer Name", "Customer interview happened(Yes/No)",
    "Customer Selected(Yes/No)", "Comments",
]


@api_bp.route("/headcount/bulk-add", methods=["POST"])
def bulk_add_headcount():
    """Add one or more new employees to the Head Count Report sheet in a
    single SharePoint download-upload cycle."""
    try:
        payload = request.get_json() or {}
        employees = payload.get("employees", [])
        if not isinstance(employees, list) or not employees:
            return jsonify({"error": "No employees provided"}), 400

        required = ["Emp Code", "Emp Name", "Sub Practice"]
        for i, emp in enumerate(employees, start=1):
            if not isinstance(emp, dict):
                return jsonify({"error": f"Row {i}: invalid payload"}), 400
            missing = [f for f in required
                       if not str(emp.get(f, "") or "").strip()]
            if missing:
                return jsonify({
                    "error": f"Row {i} missing required field(s): "
                             f"{', '.join(missing)}"
                }), 400

        hc_df = _get_cached_df()
        existing_codes = set()
        if "Emp Code" in hc_df.columns:
            existing_codes = {
                _normalize_emp_code(v) for v in hc_df["Emp Code"].tolist()
                if _normalize_emp_code(v)
            }

        seen_in_batch = set()
        for i, emp in enumerate(employees, start=1):
            code = _normalize_emp_code(emp.get("Emp Code"))
            if not code:
                return jsonify({"error": f"Row {i}: invalid Emp Code"}), 400
            if code in existing_codes:
                return jsonify({
                    "error": f"Row {i}: Emp Code '{code}' already exists "
                             f"in Head Count Report"
                }), 400
            if code in seen_in_batch:
                return jsonify({
                    "error": f"Row {i}: duplicate Emp Code '{code}' "
                             f"in this batch"
                }), 400
            seen_in_batch.add(code)

        next_sno = 1
        if "S.No" in hc_df.columns and not hc_df["S.No"].dropna().empty:
            try:
                next_sno = int(hc_df["S.No"].dropna().max()) + 1
            except (ValueError, TypeError):
                next_sno = len(hc_df) + 1

        rows_to_append = []
        for emp in employees:
            row = {"S.No": next_sno}
            for field in HEADCOUNT_ADD_FIELDS:
                val = emp.get(field, "")
                row[field] = str(val).strip() if val is not None else ""

            if not row.get("Status"):
                row["Status"] = "Confirmed"
            if not row.get("Billable/Non Billable"):
                row["Billable/Non Billable"] = "Non-Billable"
            if not row.get("Fresher/Lateral"):
                row["Fresher/Lateral"] = "Lateral"
            if not row.get("Offshore/Onsite"):
                row["Offshore/Onsite"] = "Offshore"

            if row.get("Billable/Non Billable") == "Billable":
                if not row.get("Customer interview happened(Yes/No)"):
                    row["Customer interview happened(Yes/No)"] = "Yes"
                if not row.get("Customer Selected(Yes/No)"):
                    row["Customer Selected(Yes/No)"] = "Yes"

            rows_to_append.append(row)
            next_sno += 1

        sp = _get_service()
        sheet_name = current_app.config["SHEET_NAME"]
        sp.append_rows(sheet_name, rows_to_append)

        cache.delete("headcount_df")

        return jsonify({
            "success": True,
            "added": len(rows_to_append),
            "message": f"{len(rows_to_append)} employee(s) added to "
                       f"Head Count Report.",
        })
    except Exception as e:
        logger.exception("Error in bulk add headcount")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/headcount/<int:row_index>", methods=["DELETE"])
def delete_headcount(row_index):
    """Delete a single employee row from the headcount sheet."""
    try:
        df = _get_cached_df()
        if row_index not in df.index:
            return jsonify({"error": "Employee not found"}), 404

        emp_name = _clean_value(df.loc[row_index].get("Emp Name")) or ""
        emp_code = _clean_value(df.loc[row_index].get("Emp Code")) or ""

        sp = _get_service()
        sp.delete_headcount_row(row_index)

        cache.delete("headcount_df")

        return jsonify({
            "success": True,
            "message": f"Employee {emp_name} ({emp_code}) deleted successfully",
        })
    except Exception as e:
        logger.exception("Error deleting headcount row %d", row_index)
        return jsonify({"error": str(e)}), 500


DEMAND_HEADERS = [
    "Requisition ID", "Yrs of Exp", "Skillset", "Demand Status", "Notes",
    "Customer Name", "Fulfillment Type", "Mapped Emp Code", "Mapped Emp Name",
    "Mapping Date",
]


@api_bp.route("/demand-requisition", methods=["POST"])
def create_demand_requisition():
    """Save a new demand requisition row to the Demand Requisition sheet."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data provided"}), 400

        required_fields = ["Requisition ID", "Skillset", "Demand Status", "Customer Name"]
        missing = [f for f in required_fields if not data.get(f)]
        if missing:
            return jsonify({"error": f"Missing required fields: {', '.join(missing)}"}), 400

        sp = _get_service()
        sheet_name = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")

        row_data = {h: data.get(h, "") for h in DEMAND_HEADERS}
        new_row = sp.append_row(sheet_name, row_data, DEMAND_HEADERS)

        cache.delete("demand_df")

        return jsonify({
            "success": True,
            "message": "Demand requisition saved to SharePoint",
            "row": new_row,
        })
    except Exception as e:
        logger.exception("Error saving demand requisition")
        return jsonify({"error": str(e)}), 500


# ── Demand Requisition List ──────────────────────────────────────────

@api_bp.route("/demand-requisitions", methods=["GET"])
def list_demand_requisitions():
    """Return all demand requisitions with optional filters and pagination."""
    try:
        df = _get_cached_demand_df()
        if df.empty:
            return jsonify({"data": [], "total": 0, "page": 1,
                            "per_page": 25, "total_pages": 1})

        status_filter = request.args.get("demand_status")
        search_term = request.args.get("search", "").strip().lower()
        df_kpi = request.args.get("df_kpi", "all")
        page = int(request.args.get("page", 1))
        per_page = int(request.args.get("per_page", 25))

        if df_kpi == "open":
            if "Demand Status" in df.columns:
                df = df[df["Demand Status"].isin(["Open", "In Progress"])]
        elif df_kpi == "internally_fulfilled":
            ds = df["Demand Status"].fillna("").astype(str) if "Demand Status" in df.columns else pd.Series(dtype=str)
            ft = df["Fulfillment Type"].fillna("").astype(str) if "Fulfillment Type" in df.columns else pd.Series(dtype=str)
            df = df[(ds == "Fulfilled") & (ft == "Internal")]
        elif df_kpi == "externally_fulfilled":
            ds = df["Demand Status"].fillna("").astype(str) if "Demand Status" in df.columns else pd.Series(dtype=str)
            ft = df["Fulfillment Type"].fillna("").astype(str) if "Fulfillment Type" in df.columns else pd.Series(dtype=str)
            df = df[(ds == "Fulfilled") & (ft == "External")]
        elif df_kpi == "external_raised":
            if "Demand Status" in df.columns:
                df = df[df["Demand Status"].fillna("").astype(str) == "External"]

        if status_filter and status_filter != "All":
            if "Demand Status" in df.columns:
                df = df[df["Demand Status"] == status_filter]

        if search_term:
            mask = df.apply(
                lambda row: row.astype(str).str.lower().str.contains(
                    search_term, na=False).any(), axis=1,
            )
            df = df[mask]

        total = len(df)
        start = (page - 1) * per_page
        page_df = df.iloc[start:start + per_page]

        records = []
        for idx, row in page_df.iterrows():
            record = {"row_index": int(idx)}
            for col in df.columns:
                record[col] = _clean_value(row[col])
            records.append(record)

        return jsonify({
            "data": records,
            "total": total,
            "page": page,
            "per_page": per_page,
            "total_pages": max(1, -(-total // per_page)),
        })
    except Exception as e:
        logger.exception("Error listing demand requisitions")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/demand-requisition/<int:row_index>", methods=["PUT"])
def update_demand_requisition(row_index):
    """Update a demand requisition row in SharePoint."""
    try:
        updates = request.get_json()
        if not updates:
            return jsonify({"error": "No update data provided"}), 400

        sp = _get_service()
        sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")
        sp.update_demand_row(sheet, row_index, updates)

        cache.delete("demand_df")
        return jsonify({"success": True, "message": "Demand requisition updated"})
    except Exception as e:
        logger.exception("Error updating demand requisition row %d", row_index)
        return jsonify({"error": str(e)}), 500


@api_bp.route("/demand-requisition/<int:row_index>", methods=["DELETE"])
def delete_demand_requisition(row_index):
    """Delete a demand requisition and revert mapped employee if applicable."""
    try:
        demand_df = _get_cached_demand_df()
        if demand_df.empty or row_index not in demand_df.index:
            return jsonify({"error": "Demand requisition not found"}), 404

        demand_row = demand_df.loc[row_index]

        revert_emp_row = None
        mapped_code = _normalize_emp_code(demand_row.get("Mapped Emp Code"))
        if mapped_code:
            hc_df = _get_cached_df()
            match = hc_df[hc_df["Emp Code"].apply(_normalize_emp_code) == mapped_code]
            if not match.empty:
                revert_emp_row = int(match.index[0])

        sp = _get_service()
        sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")
        sp.delete_demand_row(sheet, row_index, revert_emp_row)

        cache.delete("demand_df")
        cache.delete("headcount_df")

        msg = "Demand requisition deleted"
        if revert_emp_row is not None:
            msg += " and mapped employee reverted to Non-Billable"

        return jsonify({"success": True, "message": msg})
    except Exception as e:
        logger.exception("Error deleting demand requisition row %d", row_index)
        return jsonify({"error": str(e)}), 500


# ── Demand Fulfillment ───────────────────────────────────────────────

@api_bp.route("/demand-fulfillment/summary", methods=["GET"])
def demand_fulfillment_summary():
    """Return KPIs: total demands, open, internally/externally fulfilled, external raised."""
    try:
        df = _get_cached_demand_df()
        if df.empty:
            return jsonify({"total_demands": 0, "open_demands": 0,
                            "internally_fulfilled": 0,
                            "externally_fulfilled": 0,
                            "external_raised": 0})

        total = len(df)

        open_demands = 0
        internally_fulfilled = 0
        externally_fulfilled = 0
        external_raised = 0

        ds = df["Demand Status"].fillna("").astype(str) if "Demand Status" in df.columns else pd.Series(dtype=str)
        ft = df["Fulfillment Type"].fillna("").astype(str) if "Fulfillment Type" in df.columns else pd.Series(dtype=str)

        open_demands = int(ds.isin(["Open", "In Progress"]).sum())
        internally_fulfilled = int(((ds == "Fulfilled") & (ft == "Internal")).sum())
        externally_fulfilled = int(((ds == "Fulfilled") & (ft == "External")).sum())
        external_raised = int((ds == "External").sum())

        return jsonify({
            "total_demands": total,
            "open_demands": open_demands,
            "internally_fulfilled": internally_fulfilled,
            "externally_fulfilled": externally_fulfilled,
            "external_raised": external_raised,
        })
    except Exception as e:
        logger.exception("Error fetching demand fulfillment summary")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/demand-fulfillment/suggestions/<int:demand_row_index>",
              methods=["GET"])
def find_matching_employees(demand_row_index):
    """Find bench employees whose skills and experience match a demand."""
    try:
        demand_df = _get_cached_demand_df()
        if demand_df.empty or demand_row_index not in demand_df.index:
            return jsonify({"error": "Demand requisition not found"}), 404

        demand_row = demand_df.loc[demand_row_index]
        demand_skills = _tokenize_skills(demand_row.get("Skillset"))
        exp_min, exp_max = _parse_experience_range(demand_row.get("Yrs of Exp"))
        demand_has_exp_range = exp_min > 0 or exp_max < 99
        demand_has_skills = len(demand_skills) > 0

        logger.info("Demand match criteria — skills=%s, exp_range=(%s,%s), "
                     "has_exp=%s, has_skills=%s",
                     demand_skills, exp_min, exp_max,
                     demand_has_exp_range, demand_has_skills)

        hc_df = _get_cached_df()

        if "Billable/Non Billable" in hc_df.columns:
            bench = hc_df[hc_df["Billable/Non Billable"].isin(BENCH_STATUSES)].copy()
        else:
            bench = hc_df.copy()

        if "Status" in bench.columns:
            bench = bench[bench["Status"] != "Resigned"]

        suggestions = []
        for idx, emp in bench.iterrows():
            emp_exp = _parse_employee_experience(emp.get("Experience"))

            if demand_has_exp_range:
                if emp_exp is None:
                    continue
                if not (exp_min <= emp_exp <= exp_max):
                    continue

            emp_skills = _tokenize_skills(emp.get("Skills"))
            overlap = demand_skills & emp_skills
            skill_score = len(overlap)

            if demand_has_skills and skill_score == 0:
                continue

            exp_match = (emp_exp is not None and exp_min <= emp_exp <= exp_max
                         ) if demand_has_exp_range else True

            record = {"row_index": int(idx), "skill_match_count": skill_score,
                      "matched_skills": sorted(overlap),
                      "exp_match": exp_match,
                      "parsed_experience": emp_exp}
            for col in hc_df.columns:
                record[col] = _clean_value(emp[col])
            suggestions.append(record)

        suggestions.sort(key=lambda x: (-x["skill_match_count"],
                                        -int(x["exp_match"])))

        return jsonify({"suggestions": suggestions[:50],
                        "total_matches": len(suggestions)})
    except Exception as e:
        logger.exception("Error finding suggestions for demand %d",
                         demand_row_index)
        return jsonify({"error": str(e)}), 500


@api_bp.route("/demand-fulfillment/map", methods=["POST"])
def map_employee_to_demand():
    """Map a bench employee to a demand requisition. Updates both sheets."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data provided"}), 400

        demand_row_index = data.get("demand_row_index")
        emp_row_index = data.get("emp_row_index")
        if demand_row_index is None or emp_row_index is None:
            return jsonify({"error": "demand_row_index and emp_row_index required"}), 400

        hc_df = _get_cached_df()
        if emp_row_index not in hc_df.index:
            return jsonify({"error": "Employee not found"}), 404

        emp = hc_df.loc[emp_row_index]
        emp_code = _clean_value(emp.get("Emp Code")) or ""
        emp_name = _clean_value(emp.get("Emp Name")) or ""

        demand_updates = {
            "Fulfillment Type": "Internal",
            "Mapped Emp Code": emp_code,
            "Mapped Emp Name": emp_name,
            "Mapping Date": date.today().isoformat(),
            "Demand Status": "In Progress",
        }
        headcount_updates = {
            "Billable/Non Billable": "Proposed",
        }

        sp = _get_service()
        sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")
        sp.map_employee_to_demand(sheet, int(demand_row_index), demand_updates,
                                  int(emp_row_index), headcount_updates)

        cache.delete("demand_df")
        cache.delete("headcount_df")

        return jsonify({
            "success": True,
            "message": f"{emp_name} ({emp_code}) mapped to demand. "
                       f"Status changed to Proposed.",
        })
    except Exception as e:
        logger.exception("Error mapping employee to demand")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/demand-fulfillment/confirm", methods=["POST"])
def confirm_demand_fulfillment():
    """Confirm a mapped demand: set Demand Status to Fulfilled and employee to Billable."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data provided"}), 400

        demand_row_index = data.get("demand_row_index")
        if demand_row_index is None:
            return jsonify({"error": "demand_row_index is required"}), 400

        demand_df = _get_cached_demand_df()
        if demand_df.empty or demand_row_index not in demand_df.index:
            return jsonify({"error": "Demand requisition not found"}), 404

        demand_row = demand_df.loc[demand_row_index]
        raw_code = demand_row.get("Mapped Emp Code")
        mapped_code = _normalize_emp_code(raw_code)
        logger.info("Confirm — raw Mapped Emp Code: %r (type=%s), normalized: %r",
                     raw_code, type(raw_code).__name__, mapped_code)
        if not mapped_code:
            return jsonify({"error": "No employee mapped to this demand"}), 400

        hc_df = _get_cached_df()
        hc_codes = hc_df["Emp Code"].apply(_normalize_emp_code)
        logger.info("Confirm — looking for '%s' in HC Emp Codes (sample: %s)",
                     mapped_code, hc_codes.head(5).tolist())
        match = hc_df[hc_codes == mapped_code]
        if match.empty:
            return jsonify({"error": f"Mapped employee {mapped_code} not found in head count"}), 404

        emp_row_index = int(match.index[0])
        emp_name = _clean_value(match.iloc[0].get("Emp Name")) or ""

        demand_updates = {"Demand Status": "Fulfilled"}
        headcount_updates = {
            "Billable/Non Billable": "Billable",
            "Customer interview happened(Yes/No)": "Yes",
            "Customer Selected(Yes/No)": "Yes",
        }

        sp = _get_service()
        sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")
        sp.map_employee_to_demand(sheet, int(demand_row_index), demand_updates,
                                  emp_row_index, headcount_updates)

        cache.delete("demand_df")
        cache.delete("headcount_df")

        return jsonify({
            "success": True,
            "message": f"Demand fulfilled. {emp_name} ({mapped_code}) is now Billable.",
        })
    except Exception as e:
        logger.exception("Error confirming demand fulfillment")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/demand-fulfillment/fulfill-external", methods=["POST"])
def fulfill_external_demand():
    """Add an external employee to Head Count and mark the demand as Fulfilled."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data provided"}), 400

        demand_row_index = data.get("demand_row_index")
        if demand_row_index is None:
            return jsonify({"error": "demand_row_index is required"}), 400

        emp_name = (data.get("emp_name") or "").strip()
        emp_code = (data.get("emp_code") or "").strip()
        skills = (data.get("skills") or "").strip()
        sub_practice = (data.get("sub_practice") or "").strip()
        customer_name = (data.get("customer_name") or "").strip()

        if not emp_name or not emp_code or not sub_practice:
            return jsonify({"error": "Emp Code, Emp Name, and Sub Practice are required"}), 400

        demand_df = _get_cached_demand_df()
        if demand_df.empty or demand_row_index not in demand_df.index:
            return jsonify({"error": "Demand requisition not found"}), 404

        hc_df = _get_cached_df()
        next_sno = int(hc_df["S.No"].dropna().max()) + 1 if "S.No" in hc_df.columns and not hc_df["S.No"].dropna().empty else 1

        headcount_row = {
            "S.No": next_sno,
            "Emp Code": emp_code,
            "Emp Name": emp_name,
            "Status": "Confirmed",
            "Skills": skills,
            "Experience": data.get("experience", ""),
            "Designation": data.get("designation", ""),
            "Sub Practice": sub_practice,
            "Billable/Non Billable": "Billable",
            "Projects": data.get("projects", ""),
            "Customer Name": customer_name,
            "Fresher/Lateral": "Lateral",
            "Offshore/Onsite": data.get("offshore_onsite", "Offshore"),
            "Customer interview happened(Yes/No)": "Yes",
            "Customer Selected(Yes/No)": "Yes",
        }

        demand_updates = {
            "Demand Status": "Fulfilled",
            "Mapped Emp Code": emp_code,
            "Mapped Emp Name": emp_name,
            "Mapping Date": date.today().isoformat(),
        }

        sp = _get_service()
        sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")
        sp.fulfill_external_demand(sheet, int(demand_row_index),
                                   demand_updates, headcount_row)

        cache.delete("demand_df")
        cache.delete("headcount_df")

        return jsonify({
            "success": True,
            "message": f"External employee {emp_name} ({emp_code}) added to headcount "
                       f"and demand marked as Fulfilled.",
        })
    except Exception as e:
        logger.exception("Error fulfilling external demand")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/refresh", methods=["POST"])
def refresh_data():
    """Force refresh data from SharePoint."""
    try:
        cache.delete("headcount_df")
        cache.delete("demand_df")
        df = _get_cached_df()
        if "Status" in df.columns:
            df = df[df["Status"].fillna("").astype(str).str.strip() != "Resigned"]
        if "Billable/Non Billable" in df.columns:
            df = df[df["Billable/Non Billable"].fillna("").astype(str).str.strip() != "Resigned"]
        return jsonify({"success": True, "total_rows": len(df)})
    except Exception as e:
        logger.exception("Error refreshing data")
        return jsonify({"error": str(e)}), 500


# ── Manager Notifications ────────────────────────────────────────────

def _get_notification_store():
    from app import notification_store
    if notification_store is None:
        raise RuntimeError(
            "Manager notifications are disabled or failed to initialize"
        )
    return notification_store


_ACTION_TOKEN_SALT = "notif-action"


def _generate_action_token(notification_id, action):
    """Create a signed, time-limited token encoding {nid, act}."""
    s = URLSafeTimedSerializer(current_app.config["SECRET_KEY"])
    return s.dumps({"nid": notification_id, "act": action},
                   salt=_ACTION_TOKEN_SALT)


def _verify_action_token(token):
    """Verify and decode an action token.

    Returns (payload_dict, error_message). On success error_message is None.
    """
    s = URLSafeTimedSerializer(current_app.config["SECRET_KEY"])
    max_age = current_app.config.get("NOTIF_TOKEN_MAX_AGE_DAYS", 30) * 86400
    try:
        data = s.loads(token, salt=_ACTION_TOKEN_SALT, max_age=max_age)
        return data, None
    except SignatureExpired:
        return None, "This link has expired. Please ask the team to resend."
    except BadSignature:
        return None, "Invalid or corrupted link."


def _build_action_buttons_html(notification_id):
    """Return an HTML snippet with Yes / No buttons for an email."""
    base_url = current_app.config["APP_BASE_URL"]
    approve_token = _generate_action_token(notification_id, "approve")
    reject_token = _generate_action_token(notification_id, "reject")
    approve_url = f"{base_url}/api/notifications/action?token={approve_token}"
    reject_url = f"{base_url}/api/notifications/action?token={reject_token}"
    return f"""
        <table cellpadding="0" cellspacing="0" border="0" style="margin-top:20px">
          <tr>
            <td style="padding-right:12px">
              <a href="{approve_url}"
                 style="background-color:#28a745;color:#ffffff;padding:12px 28px;
                        text-decoration:none;border-radius:5px;font-weight:bold;
                        display:inline-block;font-size:14px">
                &#10004; Yes
              </a>
            </td>
            <td>
              <a href="{reject_url}"
                 style="background-color:#dc3545;color:#ffffff;padding:12px 28px;
                        text-decoration:none;border-radius:5px;font-weight:bold;
                        display:inline-block;font-size:14px">
                &#10008; No
              </a>
            </td>
          </tr>
        </table>
        <p style="color:#888;font-size:11px;margin-top:8px">
          Click a button above to respond directly, or reply to this email.
        </p>
    """


def _build_default_email(emp_row, demand_row):
    """Build a default subject + HTML body the user can edit in the modal."""
    emp_name = str(emp_row.get("Emp Name") or "").strip()
    emp_code = str(emp_row.get("Emp Code") or "").strip()
    skills = str(emp_row.get("Skills") or "").strip()
    grade = str(emp_row.get("Grade") or "").strip()
    sub_practice = str(emp_row.get("Sub Practice") or "").strip()
    experience = str(emp_row.get("Experience") or "").strip()
    projects = str(emp_row.get("Projects") or "").strip()
    customer_on_emp = str(emp_row.get("Customer Name") or "").strip()

    project_name = projects or customer_on_emp or "an upcoming allocation"

    subject = f"Allocation Approval Needed — {emp_name} ({emp_code})"
    html = f"""
        <p>Hi,</p>
        <p><b>{emp_name}</b> (Emp Code: <b>{emp_code}</b>, Sub-Practice: {sub_practice or 'N/A'})
           has been <b>proposed</b> for {project_name}.</p>
        <p>Please confirm whether the allocation should proceed?</p>
        <table cellpadding="6" style="border-collapse:collapse;font-size:13px;border:1px solid #ddd">
          <tr><td><b>Employee Skills:</b></td><td>{skills or 'N/A'}</td></tr>
          <tr><td><b>Employee Experience:</b></td><td>{experience or 'N/A'}</td></tr>
          <tr><td><b>Grade:</b></td><td>{grade or 'N/A'}</td></tr>
        </table>
        <p>Please <u>Note</u> : Allocation priority will be considered based on the earliest confirmed start
           date in case of multiple proposals.</p>
        <p>Reply with <b>Yes</b> to approve or <b>No</b> to reject.</p>
    """
    return subject, html


@api_bp.route("/notify-manager/preview/<int:emp_row_index>", methods=["GET"])
def notify_manager_preview(emp_row_index):
    """Resolve manager email and build a default subject/body for the
    Notify Manager modal. Does NOT send anything."""
    try:
        from app.services import manager_resolver

        hc_df = _get_cached_df()
        if emp_row_index not in hc_df.index:
            return jsonify({"error": "Employee not found"}), 404

        emp_row = hc_df.loc[emp_row_index]
        bl_status = str(emp_row.get("Billable/Non Billable") or "").strip()
        if bl_status != "Proposed":
            return jsonify({
                "error": (
                    "Employee is not in 'Proposed' status. "
                    f"Current status: '{bl_status or 'unknown'}'."
                )
            }), 400

        emp_code = _normalize_emp_code(emp_row.get("Emp Code"))

        demand_row = None
        demand_row_index = None
        demand_df = _get_cached_demand_df()
        if not demand_df.empty and "Mapped Emp Code" in demand_df.columns:
            matches = demand_df[
                demand_df["Mapped Emp Code"].apply(_normalize_emp_code) == emp_code
            ]
            if not matches.empty:
                demand_row_index = int(matches.index[0])
                demand_row = matches.iloc[0]

        threshold = current_app.config.get("NOTIF_FUZZY_THRESHOLD", 90)
        resolution = manager_resolver.resolve_manager_email(
            emp_row, hc_df, fuzzy_threshold=threshold
        )

        subject, html_body = _build_default_email(emp_row, demand_row)

        company_email = str(emp_row.get("Company Email") or "").strip()

        return jsonify({
            "emp_row_index": emp_row_index,
            "emp_code": str(emp_row.get("Emp Code") or ""),
            "emp_name": str(emp_row.get("Emp Name") or ""),
            "has_demand_link": demand_row is not None,
            "demand_row_index": demand_row_index,
            "demand_req_id": (
                str(demand_row.get("Requisition ID") or "")
                if demand_row is not None else ""
            ),
            "customer_name": (
                str(demand_row.get("Customer Name") or "")
                if demand_row is not None else ""
            ),
            "manager_resolution": resolution,
            "default_to": resolution.get("manager_email", ""),
            "default_cc": [company_email] if company_email else [],
            "subject": subject,
            "body_html": html_body,
        })
    except Exception as e:
        logger.exception("Notify preview failed")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/notify-manager", methods=["POST"])
def notify_manager_send():
    """Send the initial Notify-Manager email and create a tracking record.

    Performs pre-flight checks:
      - Employee is currently in 'Proposed' status.
      - The demand still has Mapped Emp Code matching this employee.
      - No active notification is already pending for this employee
        (use /resend instead).
    """
    try:
        from app.services.email_service import EmailService, EmailServiceError
        from app.services.notification_store import STATUS_AWAITING

        store = _get_notification_store()
        data = request.get_json() or {}

        emp_row_index = data.get("emp_row_index")
        manager_email = (data.get("manager_email") or "").strip()
        cc_emails = data.get("cc_emails") or []
        subject = (data.get("subject") or "").strip()
        body_html = data.get("body_html") or ""
        resolution_method = (data.get("resolution_method") or "manual").strip()
        manager_name_override = (data.get("manager_name") or "").strip()

        if emp_row_index is None:
            return jsonify({"error": "emp_row_index is required"}), 400
        if not manager_email:
            return jsonify({"error": "Manager email is required"}), 400
        if not subject or not body_html:
            return jsonify({"error": "Subject and body cannot be empty"}), 400

        hc_df = _get_cached_df()
        if emp_row_index not in hc_df.index:
            return jsonify({"error": "Employee not found"}), 404

        emp_row = hc_df.loc[emp_row_index]
        bl_status = str(emp_row.get("Billable/Non Billable") or "").strip()
        if bl_status != "Proposed":
            return jsonify({
                "error": (
                    "Employee is no longer in 'Proposed' status. "
                    f"Current status: '{bl_status or 'unknown'}'. "
                    "Please re-map the employee before notifying."
                )
            }), 400

        emp_code = _normalize_emp_code(emp_row.get("Emp Code"))
        emp_name = _clean_value(emp_row.get("Emp Name")) or ""
        manager_name = (
            manager_name_override
            or str(emp_row.get("First Line Manager") or "").strip()
        )

        # Demand link is optional — Proposed status alone is enough to notify.
        # When a demand row exists and points back at this employee, we attach
        # its context so the email is richer and Approve/Reject can update
        # both sheets atomically. When no demand row exists, we still send
        # the notification (an employee can be manually marked "Proposed"
        # without going through demand fulfillment).
        demand_df = _get_cached_demand_df()
        demand_row_index = None
        demand_req_id = ""
        customer_name = ""
        if not demand_df.empty and "Mapped Emp Code" in demand_df.columns:
            matches = demand_df[
                demand_df["Mapped Emp Code"].apply(_normalize_emp_code) == emp_code
            ]
            if not matches.empty:
                demand_row_index = int(matches.index[0])
                demand_req_id = str(matches.iloc[0].get("Requisition ID") or "")
                customer_name = str(matches.iloc[0].get("Customer Name") or "")

        # Reject if there's already an active notification
        existing = store.get_active_for_employee(emp_code)
        if existing:
            return jsonify({
                "error": (
                    f"An active notification already exists for this "
                    f"employee (sent {existing['sent_at']}). Use Resend "
                    f"or Stop Reminders instead."
                )
            }), 409

        # Create notification first so we have an ID for the action buttons
        notif_id = store.create(
            emp_code=emp_code, emp_name=emp_name,
            emp_row_index=int(emp_row_index),
            demand_row_index=demand_row_index,
            demand_req_id=demand_req_id, customer_name=customer_name,
            manager_name=manager_name, manager_email=manager_email,
            resolution_method=resolution_method,
            cc_emails=cc_emails, subject=subject, body_html=body_html,
        )

        buttons_html = _build_action_buttons_html(notif_id)
        email_body = body_html + buttons_html

        email_svc = EmailService(_get_service())
        try:
            send_result = email_svc.send_mail(
                to_email=manager_email, subject=subject,
                html_body=email_body, cc_emails=cc_emails, save_to_sent=True,
            )
        except EmailServiceError as e:
            store.delete(notif_id)
            return jsonify({"error": str(e)}), 502

        store.record_send_ids(
            notif_id,
            conversation_id=send_result.get("conversation_id"),
            message_id=send_result.get("message_id"),
        )

        return jsonify({
            "success": True,
            "notification_id": notif_id,
            "status": STATUS_AWAITING,
            "message": f"Email sent to {manager_email}",
        })
    except Exception as e:
        logger.exception("Notify-manager send failed")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/notifications", methods=["GET"])
def list_notifications():
    """List all notifications, filterable by status / emp_code."""
    try:
        store = _get_notification_store()
        status = request.args.get("status") or None
        emp_code = request.args.get("emp_code") or None
        limit = int(request.args.get("limit", 500))
        rows = store.list(status=status, emp_code=emp_code, limit=limit)
        return jsonify({"data": rows, "total": len(rows)})
    except Exception as e:
        logger.exception("List notifications failed")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/notifications/by-employee/<int:emp_row_index>", methods=["GET"])
def get_notification_for_employee(emp_row_index):
    """Return the latest notification (if any) for the given headcount row."""
    try:
        store = _get_notification_store()
        hc_df = _get_cached_df()
        if emp_row_index not in hc_df.index:
            return jsonify({"error": "Employee not found"}), 404
        emp_code = _normalize_emp_code(hc_df.loc[emp_row_index].get("Emp Code"))
        if not emp_code:
            return jsonify({"notification": None})
        latest = store.get_latest_for_employee(emp_code)
        return jsonify({"notification": latest})
    except Exception as e:
        logger.exception("Get notification for employee failed")
        return jsonify({"error": str(e)}), 500


def _perform_approve(notification_id, decided_by="user", note=""):
    """Core approve logic shared by the dashboard API and email-link handler.

    Returns (result_dict, http_status_code).
    """
    from app.services.notification_store import STATUS_APPROVED

    store = _get_notification_store()
    n = store.get(notification_id)
    if not n:
        return {"error": "Notification not found"}, 404

    if n["status"] != "awaiting_reply":
        return {
            "error": f"Already resolved ({n['status']})",
            "already_decided": True,
        }, 409

    emp_code = _normalize_emp_code(n["emp_code"])
    hc_df = _get_cached_df()
    match = hc_df[
        hc_df["Emp Code"].apply(_normalize_emp_code) == emp_code
    ]
    if match.empty:
        return {"error": "Employee no longer in headcount"}, 404
    emp_row_index = int(match.index[0])
    emp_name = _clean_value(match.iloc[0].get("Emp Name")) or ""

    headcount_updates = {
        "Billable/Non Billable": "Billable",
        "Customer interview happened(Yes/No)": "Yes",
        "Customer Selected(Yes/No)": "Yes",
    }

    sp = _get_service()
    sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")

    demand_idx = n.get("demand_row_index")
    demand_still_valid = False
    if demand_idx is not None:
        demand_df = _get_cached_demand_df()
        if demand_idx in demand_df.index:
            demand_row = demand_df.loc[demand_idx]
            mapped_code = _normalize_emp_code(
                demand_row.get("Mapped Emp Code")
            )
            if mapped_code == emp_code:
                demand_still_valid = True

    if demand_still_valid:
        demand_updates = {"Demand Status": "Fulfilled"}
        sp.map_employee_to_demand(
            sheet, int(demand_idx), demand_updates,
            emp_row_index, headcount_updates,
        )
        msg = f"{emp_name} ({emp_code}) approved — marked Billable, demand Fulfilled."
    else:
        sp.update_row(emp_row_index, headcount_updates)
        msg = (
            f"{emp_name} ({emp_code}) approved — marked Billable. "
            f"(No demand row was linked.)"
        )

    cache.delete("demand_df")
    cache.delete("headcount_df")

    store.update_status(
        notification_id, STATUS_APPROVED,
        decided_by=decided_by, decision_note=note,
    )

    return {"success": True, "message": msg}, 200


def _perform_reject(notification_id, decided_by="user", note=""):
    """Core reject logic shared by the dashboard API and email-link handler.

    Returns (result_dict, http_status_code).
    """
    from app.services.notification_store import STATUS_REJECTED

    store = _get_notification_store()
    n = store.get(notification_id)
    if not n:
        return {"error": "Notification not found"}, 404

    if n["status"] != "awaiting_reply":
        return {
            "error": f"Already resolved ({n['status']})",
            "already_decided": True,
        }, 409

    emp_code = n["emp_code"]
    hc_df = _get_cached_df()
    match = hc_df[hc_df["Emp Code"].apply(_normalize_emp_code) == emp_code]
    if match.empty:
        store.update_status(
            notification_id, STATUS_REJECTED,
            decided_by=decided_by,
            decision_note=note + " (employee no longer in headcount)",
        )
        return {
            "success": True,
            "message": "Marked rejected (employee not in headcount)",
        }, 200

    emp_row_index = int(match.index[0])

    sp = _get_service()
    sheet = current_app.config.get("DEMAND_SHEET_NAME", "Demand Requisition")

    demand_idx = n.get("demand_row_index")
    if demand_idx is not None:
        demand_df = _get_cached_demand_df()
        if demand_idx in demand_df.index:
            demand_updates = {
                "Demand Status": "Open",
                "Mapped Emp Code": "",
                "Mapped Emp Name": "",
                "Mapping Date": "",
                "Fulfillment Type": "",
            }
            hc_updates = {"Billable/Non Billable": "Non-Billable"}
            sp.map_employee_to_demand(
                sheet, int(demand_idx), demand_updates,
                emp_row_index, hc_updates,
            )
        else:
            sp.update_row(
                emp_row_index, {"Billable/Non Billable": "Non-Billable"}
            )
    else:
        sp.update_row(
            emp_row_index, {"Billable/Non Billable": "Non-Billable"}
        )

    cache.delete("demand_df")
    cache.delete("headcount_df")

    store.update_status(
        notification_id, STATUS_REJECTED,
        decided_by=decided_by, decision_note=note,
    )

    return {
        "success": True,
        "message": (
            "Marked rejected. Demand reopened and employee reverted to "
            "Non-Billable."
        ),
    }, 200


# ── Email Action Link (one-click approve/reject from email) ──────────

@api_bp.route("/notifications/action", methods=["GET"])
def notification_action_from_email():
    """Handle a manager clicking an Approve / Reject button in the email.

    The token encodes the notification ID and action, signed with SECRET_KEY.
    On success the manager sees a simple confirmation page.
    """
    token = request.args.get("token", "")
    if not token:
        return render_template("notification_action.html",
                               success=False,
                               message="Missing or invalid link."), 400

    data, err = _verify_action_token(token)
    if err:
        return render_template("notification_action.html",
                               success=False, message=err), 400

    notification_id = data.get("nid")
    action = data.get("act")
    if action not in ("approve", "reject"):
        return render_template("notification_action.html",
                               success=False,
                               message="Invalid action in link."), 400

    try:
        if action == "approve":
            result, status = _perform_approve(
                notification_id, decided_by="manager_email_link")
        else:
            result, status = _perform_reject(
                notification_id, decided_by="manager_email_link")
    except Exception as e:
        logger.exception("Email action failed for notification %d",
                         notification_id)
        return render_template("notification_action.html",
                               success=False,
                               message=f"Something went wrong: {e}"), 500

    if result.get("already_decided"):
        return render_template("notification_action.html",
                               success=False,
                               message="This request has already been processed. "
                                       "No further action is needed.")

    if status >= 400:
        return render_template("notification_action.html",
                               success=False,
                               message=result.get("error",
                                                   "An error occurred.")), status

    action_label = "Approved" if action == "approve" else "Rejected"
    return render_template("notification_action.html",
                           success=True,
                           message=f"{action_label} successfully. "
                                   f"{result.get('message', '')}")


# ── Dashboard API wrappers (existing POST endpoints) ─────────────────

@api_bp.route("/notifications/<int:notification_id>/approve", methods=["POST"])
def notification_approve(notification_id):
    """Manager approved via dashboard → mark employee Billable."""
    try:
        decided_by = (request.get_json() or {}).get("decided_by") or "user"
        note = (request.get_json() or {}).get("note") or ""
        result, status = _perform_approve(notification_id, decided_by, note)
        return jsonify(result), status
    except Exception as e:
        logger.exception("Approve notification %d failed", notification_id)
        return jsonify({"error": str(e)}), 500


@api_bp.route("/notifications/<int:notification_id>/reject", methods=["POST"])
def notification_reject(notification_id):
    """Manager rejected via dashboard → un-map demand and revert employee."""
    try:
        decided_by = (request.get_json() or {}).get("decided_by") or "user"
        note = (request.get_json() or {}).get("note") or ""
        result, status = _perform_reject(notification_id, decided_by, note)
        return jsonify(result), status
    except Exception as e:
        logger.exception("Reject notification %d failed", notification_id)
        return jsonify({"error": str(e)}), 500


@api_bp.route("/notifications/<int:notification_id>/cancel", methods=["POST"])
def notification_cancel(notification_id):
    """Stop reminders without changing any sheet data."""
    try:
        from app.services.notification_store import STATUS_CANCELLED

        store = _get_notification_store()
        n = store.get(notification_id)
        if not n:
            return jsonify({"error": "Notification not found"}), 404

        decided_by = (request.get_json() or {}).get("decided_by") or "user"
        note = (request.get_json() or {}).get("note") or "Reminders stopped"

        store.update_status(
            notification_id, STATUS_CANCELLED,
            decided_by=decided_by, decision_note=note,
        )
        return jsonify({"success": True, "message": "Reminders stopped"})
    except Exception as e:
        logger.exception("Cancel notification %d failed", notification_id)
        return jsonify({"error": str(e)}), 500


@api_bp.route("/notifications/<int:notification_id>/resend", methods=["POST"])
def notification_resend(notification_id):
    """Resend the original email with fresh action buttons and reset reminders."""
    try:
        from app.services.email_service import EmailService, EmailServiceError

        store = _get_notification_store()
        n = store.get(notification_id)
        if not n:
            return jsonify({"error": "Notification not found"}), 404

        buttons_html = _build_action_buttons_html(notification_id)
        email_body = n["body_html"] + buttons_html

        email_svc = EmailService(_get_service())
        try:
            send_result = email_svc.send_mail(
                to_email=n["manager_email"],
                subject=n["subject"],
                html_body=email_body,
                cc_emails=n.get("cc_emails") or [],
                save_to_sent=True,
            )
        except EmailServiceError as e:
            return jsonify({"error": str(e)}), 502

        store.reset_for_resend(
            notification_id,
            conversation_id=send_result.get("conversation_id"),
            message_id=send_result.get("message_id"),
        )
        return jsonify({"success": True, "message": "Email resent"})
    except Exception as e:
        logger.exception("Resend notification %d failed", notification_id)
        return jsonify({"error": str(e)}), 500


@api_bp.route("/manager-email-audit", methods=["GET"])
def manager_email_audit():
    """List employees whose manager email could not be auto-resolved."""
    try:
        from app.services import manager_resolver

        hc_df = _get_cached_df()
        threshold = current_app.config.get("NOTIF_FUZZY_THRESHOLD", 90)
        rows = manager_resolver.audit_unresolved(hc_df, threshold)
        return jsonify({"data": rows, "total": len(rows)})
    except Exception as e:
        logger.exception("Manager email audit failed")
        return jsonify({"error": str(e)}), 500


# ── Blocked Employee Allocation Check ────────────────────────────────

def _build_blocked_email(emp_row):
    """Build default subject + HTML body for blocked employee allocation check."""
    emp_name = str(emp_row.get("Emp Name") or "").strip()
    emp_code = str(emp_row.get("Emp Code") or "").strip()
    skills = str(emp_row.get("Skills") or "").strip()
    grade = str(emp_row.get("Grade") or "").strip()
    sub_practice = str(emp_row.get("Sub Practice") or "").strip()
    experience = str(emp_row.get("Experience") or "").strip()
    projects = str(emp_row.get("Projects") or "").strip()
    customer = str(emp_row.get("Customer Name") or "").strip()

    project_name = projects or customer or "an upcoming allocation"

    subject = f"Allocation Status Check — {emp_name} ({emp_code})"
    html = f"""
        <p>Hi,</p>
        <p><b>{emp_name}</b> (Emp Code: <b>{emp_code}</b>, Sub-Practice: {sub_practice or 'N/A'})
           has been <b>Blocked</b> for {project_name}.</p>
        <p>Please confirm the allocation start date.</p>
        <table cellpadding="6" style="border-collapse:collapse;font-size:13px;border:1px solid #ddd">
          <tr><td><b>Employee Skills:</b></td><td>{skills or 'N/A'}</td></tr>
          <tr><td><b>Employee Experience:</b></td><td>{experience or 'N/A'}</td></tr>
          <tr><td><b>Grade:</b></td><td>{grade or 'N/A'}</td></tr>
        </table>
    """
    return subject, html


def _build_blocked_action_buttons_html(notification_id):
    """Return HTML with Yes (opens form) / No (one-click) buttons for blocked flow."""
    base_url = current_app.config["APP_BASE_URL"]
    confirm_token = _generate_action_token(notification_id, "confirm_blocked")
    reject_token = _generate_action_token(notification_id, "reject_blocked")
    confirm_url = f"{base_url}/api/notifications/blocked-action?token={confirm_token}"
    reject_url = f"{base_url}/api/notifications/blocked-action?token={reject_token}"
    return f"""
        <table cellpadding="0" cellspacing="0" border="0" style="margin-top:20px">
          <tr>
            <td style="padding-right:12px">
              <a href="{confirm_url}"
                 style="background-color:#28a745;color:#ffffff;padding:12px 28px;
                        text-decoration:none;border-radius:5px;font-weight:bold;
                        display:inline-block;font-size:14px">
                &#10004; Yes, Allocation Completed
              </a>
            </td>
            <td>
              <a href="{reject_url}"
                 style="background-color:#dc3545;color:#ffffff;padding:12px 28px;
                        text-decoration:none;border-radius:5px;font-weight:bold;
                        display:inline-block;font-size:14px">
                &#10008; No
              </a>
            </td>
          </tr>
        </table>
        <p style="color:#888;font-size:11px;margin-top:8px">
          Click a button above to respond directly, or reply to this email.
        </p>
    """


@api_bp.route("/notify-blocked/preview/<int:emp_row_index>", methods=["GET"])
def notify_blocked_preview(emp_row_index):
    """Resolve manager email and build default email for a Blocked employee."""
    try:
        from app.services import manager_resolver

        hc_df = _get_cached_df()
        if emp_row_index not in hc_df.index:
            return jsonify({"error": "Employee not found"}), 404

        emp_row = hc_df.loc[emp_row_index]
        bl_status = str(emp_row.get("Billable/Non Billable") or "").strip()
        if bl_status != "Blocked":
            return jsonify({
                "error": (
                    "Employee is not in 'Blocked' status. "
                    f"Current status: '{bl_status or 'unknown'}'."
                )
            }), 400

        threshold = current_app.config.get("NOTIF_FUZZY_THRESHOLD", 90)
        resolution = manager_resolver.resolve_manager_email(
            emp_row, hc_df, fuzzy_threshold=threshold
        )

        subject, html_body = _build_blocked_email(emp_row)
        company_email = str(emp_row.get("Company Email") or "").strip()

        return jsonify({
            "emp_row_index": emp_row_index,
            "emp_code": str(emp_row.get("Emp Code") or ""),
            "emp_name": str(emp_row.get("Emp Name") or ""),
            "manager_resolution": resolution,
            "default_to": resolution.get("manager_email", ""),
            "default_cc": [company_email] if company_email else [],
            "subject": subject,
            "body_html": html_body,
        })
    except Exception as e:
        logger.exception("Notify blocked preview failed")
        return jsonify({"error": str(e)}), 500


@api_bp.route("/notify-blocked", methods=["POST"])
def notify_blocked_send():
    """Send allocation-check email for a Blocked employee and create tracking record."""
    try:
        from app.services.email_service import EmailService, EmailServiceError
        from app.services.notification_store import (
            STATUS_AWAITING, NOTIF_TYPE_BLOCKED,
        )

        store = _get_notification_store()
        data = request.get_json() or {}

        emp_row_index = data.get("emp_row_index")
        manager_email = (data.get("manager_email") or "").strip()
        cc_emails = data.get("cc_emails") or []
        subject = (data.get("subject") or "").strip()
        body_html = data.get("body_html") or ""
        resolution_method = (data.get("resolution_method") or "manual").strip()
        manager_name_override = (data.get("manager_name") or "").strip()

        if emp_row_index is None:
            return jsonify({"error": "emp_row_index is required"}), 400
        if not manager_email:
            return jsonify({"error": "Manager email is required"}), 400
        if not subject or not body_html:
            return jsonify({"error": "Subject and body cannot be empty"}), 400

        hc_df = _get_cached_df()
        if emp_row_index not in hc_df.index:
            return jsonify({"error": "Employee not found"}), 404

        emp_row = hc_df.loc[emp_row_index]
        bl_status = str(emp_row.get("Billable/Non Billable") or "").strip()
        if bl_status != "Blocked":
            return jsonify({
                "error": (
                    "Employee is no longer in 'Blocked' status. "
                    f"Current status: '{bl_status or 'unknown'}'."
                )
            }), 400

        emp_code = _normalize_emp_code(emp_row.get("Emp Code"))
        emp_name = _clean_value(emp_row.get("Emp Name")) or ""
        manager_name = (
            manager_name_override
            or str(emp_row.get("First Line Manager") or "").strip()
        )

        existing = store.get_active_for_employee(emp_code)
        if existing:
            return jsonify({
                "error": (
                    f"An active notification already exists for this "
                    f"employee (sent {existing['sent_at']}). Use Resend "
                    f"or Stop Reminders instead."
                )
            }), 409

        notif_id = store.create(
            emp_code=emp_code, emp_name=emp_name,
            emp_row_index=int(emp_row_index),
            demand_row_index=None, demand_req_id="",
            customer_name="", manager_name=manager_name,
            manager_email=manager_email,
            resolution_method=resolution_method,
            cc_emails=cc_emails, subject=subject, body_html=body_html,
            notification_type=NOTIF_TYPE_BLOCKED,
        )

        buttons_html = _build_blocked_action_buttons_html(notif_id)
        email_body = body_html + buttons_html

        email_svc = EmailService(_get_service())
        try:
            send_result = email_svc.send_mail(
                to_email=manager_email, subject=subject,
                html_body=email_body, cc_emails=cc_emails, save_to_sent=True,
            )
        except EmailServiceError as e:
            store.delete(notif_id)
            return jsonify({"error": str(e)}), 502

        store.record_send_ids(
            notif_id,
            conversation_id=send_result.get("conversation_id"),
            message_id=send_result.get("message_id"),
        )

        return jsonify({
            "success": True,
            "notification_id": notif_id,
            "status": STATUS_AWAITING,
            "message": f"Allocation check email sent to {manager_email}",
        })
    except Exception as e:
        logger.exception("Notify-blocked send failed")
        return jsonify({"error": str(e)}), 500


def _perform_blocked_approve(notification_id, allocation_date,
                             decided_by="user", note=""):
    """Blocked employee confirmed as allocated → update to Billable."""
    from app.services.notification_store import STATUS_APPROVED

    store = _get_notification_store()
    n = store.get(notification_id)
    if not n:
        return {"error": "Notification not found"}, 404

    if n["status"] != "awaiting_reply":
        return {
            "error": f"Already resolved ({n['status']})",
            "already_decided": True,
        }, 409

    emp_code = _normalize_emp_code(n["emp_code"])
    hc_df = _get_cached_df()
    match = hc_df[hc_df["Emp Code"].apply(_normalize_emp_code) == emp_code]
    if match.empty:
        return {"error": "Employee no longer in headcount"}, 404

    emp_row_index = int(match.index[0])
    emp_name = _clean_value(match.iloc[0].get("Emp Name")) or ""

    sp = _get_service()
    sp.update_row(emp_row_index, {
        "Billable/Non Billable": "Billable",
        "Billable Till Date": allocation_date,
        "Customer interview happened(Yes/No)": "Yes",
        "Customer Selected(Yes/No)": "Yes",
    })

    cache.delete("headcount_df")

    store.record_allocation_response(notification_id, allocation_date)
    store.update_status(
        notification_id, STATUS_APPROVED,
        decided_by=decided_by,
        decision_note=note or f"Allocation confirmed. Date: {allocation_date}",
    )

    return {
        "success": True,
        "message": (
            f"{emp_name} ({emp_code}) allocation confirmed — "
            f"status updated to Billable (date: {allocation_date})."
        ),
    }, 200


def _perform_blocked_reject(notification_id, decided_by="user", note=""):
    """Manager says allocation NOT completed → keep as Blocked."""
    from app.services.notification_store import STATUS_REJECTED

    store = _get_notification_store()
    n = store.get(notification_id)
    if not n:
        return {"error": "Notification not found"}, 404

    if n["status"] != "awaiting_reply":
        return {
            "error": f"Already resolved ({n['status']})",
            "already_decided": True,
        }, 409

    store.update_status(
        notification_id, STATUS_REJECTED,
        decided_by=decided_by,
        decision_note=note or "Manager confirmed allocation not completed",
    )

    return {
        "success": True,
        "message": "Noted — employee will remain as Blocked.",
    }, 200


@api_bp.route("/notifications/blocked-action", methods=["GET"])
def blocked_action_from_email():
    """Handle manager clicking Yes/No in blocked employee allocation email."""
    token = request.args.get("token", "")
    if not token:
        return render_template("blocked_action.html",
                               step="error",
                               message="Missing or invalid link."), 400

    data, err = _verify_action_token(token)
    if err:
        return render_template("blocked_action.html",
                               step="error", message=err), 400

    notification_id = data.get("nid")
    action = data.get("act")

    if action == "reject_blocked":
        try:
            result, status = _perform_blocked_reject(
                notification_id, decided_by="manager_email_link")
        except Exception as e:
            logger.exception("Blocked reject failed for notification %d",
                             notification_id)
            return render_template("blocked_action.html", step="error",
                                   message=f"Something went wrong: {e}"), 500

        if result.get("already_decided"):
            return render_template("blocked_action.html", step="done",
                                   success=False,
                                   message="This request has already been processed.")
        return render_template("blocked_action.html", step="done",
                               success=status < 400,
                               message=result.get("message") or result.get("error"))

    if action == "confirm_blocked":
        store = _get_notification_store()
        n = store.get(notification_id)
        if n and n["status"] != "awaiting_reply":
            return render_template("blocked_action.html", step="done",
                                   success=False,
                                   message="This request has already been processed.")
        return render_template("blocked_action.html", step="form",
                               token=token, notification_id=notification_id)

    return render_template("blocked_action.html", step="error",
                           message="Invalid action in link."), 400


@api_bp.route("/notifications/blocked-action", methods=["POST"])
def blocked_action_submit():
    """Process the allocation-confirmed form submission from the manager."""
    token = request.form.get("token", "")
    allocation_date = request.form.get("allocation_date", "").strip()

    data, err = _verify_action_token(token)
    if err:
        return render_template("blocked_action.html",
                               step="error", message=err), 400

    notification_id = data.get("nid")

    if not allocation_date:
        return render_template("blocked_action.html", step="form",
                               token=token, notification_id=notification_id,
                               error="Please provide the allocation completion date.")

    try:
        result, status = _perform_blocked_approve(
            notification_id, allocation_date=allocation_date,
            decided_by="manager_email_link")
    except Exception as e:
        logger.exception("Blocked approve failed for notification %d",
                         notification_id)
        return render_template("blocked_action.html", step="error",
                               message=f"Something went wrong: {e}"), 500

    if result.get("already_decided"):
        return render_template("blocked_action.html", step="done",
                               success=False,
                               message="This request has already been processed.")

    return render_template("blocked_action.html", step="done",
                           success=status < 400,
                           message=result.get("message") or result.get("error"))


@api_bp.route("/notifications/<int:notification_id>/blocked-approve",
              methods=["POST"])
def notification_blocked_approve(notification_id):
    """Dashboard-triggered approve for a blocked employee notification."""
    try:
        data = request.get_json() or {}
        allocation_date = (data.get("allocation_date") or "").strip()
        if not allocation_date:
            return jsonify({"error": "allocation_date is required"}), 400
        decided_by = data.get("decided_by") or "user"
        note = data.get("note") or ""
        result, status = _perform_blocked_approve(
            notification_id, allocation_date, decided_by, note)
        return jsonify(result), status
    except Exception as e:
        logger.exception("Blocked approve notification %d failed",
                         notification_id)
        return jsonify({"error": str(e)}), 500


@api_bp.route("/notifications/<int:notification_id>/blocked-reject",
              methods=["POST"])
def notification_blocked_reject(notification_id):
    """Dashboard-triggered reject for a blocked employee notification."""
    try:
        decided_by = (request.get_json() or {}).get("decided_by") or "user"
        note = (request.get_json() or {}).get("note") or ""
        result, status = _perform_blocked_reject(
            notification_id, decided_by, note)
        return jsonify(result), status
    except Exception as e:
        logger.exception("Blocked reject notification %d failed",
                         notification_id)
        return jsonify({"error": str(e)}), 500
