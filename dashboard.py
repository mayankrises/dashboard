"""
BPCL Uran - Document Control Index (DCI) Dashboard
Now backed by Supabase (PostgreSQL + Storage) for persistence.

Run with:  streamlit run dashboard.py

Required Streamlit secrets (see secrets.toml.example):
    SUPABASE_URL
    SUPABASE_SERVICE_ROLE_KEY
"""

import hashlib
import hmac
import os
import io
import zipfile
import secrets
from datetime import datetime, date, timedelta, timezone

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(page_title="DCI Project Hub", layout="wide")


def _early_secret(name):
    try:
        value = st.secrets.get(name)
    except Exception:
        value = None
    return str(value or os.getenv(name, "")).strip()


try:
    from streamlit_cookies_manager import EncryptedCookieManager
except ImportError:
    st.error(
        "Persistent sign-in requires `streamlit-cookies-manager`. "
        "Add `streamlit-cookies-manager==0.2.0` to requirements.txt and redeploy."
    )
    st.stop()

_COOKIE_PASSWORD = _early_secret("COOKIES_PASSWORD")
if not _COOKIE_PASSWORD:
    st.error("COOKIES_PASSWORD is not configured in Streamlit secrets or the environment.")
    st.stop()

AUTH_COOKIES = EncryptedCookieManager(
    prefix="dci-project-hub/",
    password=_COOKIE_PASSWORD,
)
if not AUTH_COOKIES.ready():
    st.stop()

AUTH_COOKIE_NAME = "app_session"
ACCOUNT_SESSION_DAYS = 30
MASTER_SESSION_HOURS = 8


APPROVED_CODES = {"1", "1 WITH COMMENTS"}
RETAINED_CODES = {"R"}
CODE2_CODES = {"2"}
CODE3_CODES = {"3"}
VOID_CODES = {"V"}

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

AGING_BUCKET_ORDER = ["Not due yet", "0-7 days", "8-14 days", "15-21 days", ">21 days"]

CHART_PALETTE = ["#B4A7D6", "#F1948A", "#8EE4D0", "#f5c26b", "#8ab4f8", "#c9a0dc"]

STORAGE_BUCKET = "dci-workbooks"
MAX_UPLOAD_BYTES = 25 * 1024 * 1024
MAX_CONSECUTIVE_EMPTY_ROWS = 25
TEMPLATE_ENGINE_VERSION = 2

# Original Excel-derived fields (unchanged shape expected by all the existing
# analytical tabs) plus the new dashboard-managed / editable fields, so a
# document loaded from Supabase looks exactly like a document parsed from Excel.
RECORD_COLUMNS = [
    "sr_no", "doc_no", "title", "discipline", "poc_submission",
    "latest_rev", "latest_uploaded_date", "eil_review_status", "latest_code",
    "latest_from_eil_date", "category", "l4_marked", "schedule_submission_date",
    "status", "poc", "eil_mandays", "eil_manhours", "stspl_mandays",
    "stspl_manhours", "revision_cycle", "expected_manhours",
    # -- new, dashboard-managed fields --
    "remarks", "assigned_engineer", "action_owner", "expected_completion_date",
    "forecast_month", "approval_date",
    # -- linkage fields (not shown in charts, used for Supabase round-trips) --
    "document_key", "document_id", "original_row_number",
]

HISTORY_COLUMNS = ["document_id", "doc_no", "rev_no", "uploaded_date", "code", "from_eil_date"]

NUMERIC_COLUMNS = [
    "eil_mandays", "eil_manhours", "stspl_mandays",
    "stspl_manhours", "revision_cycle", "expected_manhours",
]

# Maps documents-table (Supabase) column names <-> dashboard RECORD_COLUMNS names.
DOC_TO_RECORD_FIELD = {
    "document_number": "doc_no",
    "document_title": "title",
    "discipline": "discipline",
    "vendor": "poc",
    "revision": "latest_rev",
    "dci_status": "status",
    "returned_code": "latest_code",
    "planned_submission_date": "schedule_submission_date",
    "actual_submission_date": "latest_uploaded_date",
    "review_date": "latest_from_eil_date",
    "approval_date": "approval_date",
    "forecast_month": "forecast_month",
    "remarks": "remarks",
    "assigned_engineer": "assigned_engineer",
    "action_owner": "action_owner",
    "expected_completion_date": "expected_completion_date",
}
DEFAULT_REVISION_OPTIONS = [str(i) for i in range(0, 21)]
DEFAULT_CODE_TO_DCI_STATUS = {
    "1": "APPROVED",
    "1 WITH COMMENTS": "APPROVED WITH COMMENTS",
    "2": "COMMENTED",
    "3": "RESUBMISSION",
    "R": "RETAINED FOR RECORDS",
    "V": "VOID",
}
FORM_CONFIG_KEY = "revision_update_form"


def get_admin_password():
    """Read configuration-admin password from Streamlit secrets or environment."""
    try:
        secret_value = st.secrets.get("ADMIN_PASSWORD")
    except Exception:
        secret_value = None
    return str(secret_value or os.getenv("ADMIN_PASSWORD", "")).strip()


def admin_password_configured():
    return bool(get_admin_password())


def _clean_unique_options(values):
    cleaned, seen = [], set()
    for value in values or []:
        item = str(value).strip()
        if item and item not in seen:
            cleaned.append(item)
            seen.add(item)
    return cleaned


def _values_from_documents(df, column):
    if df is None or df.empty or column not in df.columns:
        return []
    return _clean_unique_options(df[column].dropna().astype(str).tolist())


def default_form_config(df=None):
    return {
        "revision_options": DEFAULT_REVISION_OPTIONS.copy(),
        "code_status_map": DEFAULT_CODE_TO_DCI_STATUS.copy(),
        "discipline_options": _values_from_documents(df, "discipline"),
        "vendor_options": _values_from_documents(df, "poc"),
    }


def normalize_form_config(value, df=None):
    defaults = default_form_config(df)
    if not isinstance(value, dict):
        return defaults

    revisions = _clean_unique_options(value.get("revision_options"))
    disciplines = _clean_unique_options(value.get("discipline_options"))
    vendors = _clean_unique_options(value.get("vendor_options"))

    code_status_map = {}
    raw_map = value.get("code_status_map") or {}
    if isinstance(raw_map, dict):
        for code, status in raw_map.items():
            clean_code = normalize_code(code)
            clean_status = " ".join(str(status or "").strip().upper().split())
            if clean_code and clean_status:
                code_status_map[clean_code] = clean_status

    return {
        "revision_options": revisions or defaults["revision_options"],
        "code_status_map": code_status_map or defaults["code_status_map"],
        "discipline_options": disciplines or defaults["discipline_options"],
        "vendor_options": vendors or defaults["vendor_options"],
    }


@st.cache_data(ttl=60, show_spinner=False)
def load_form_config(_supabase, _df_signature=None):
    try:
        result = (
            _supabase.table("dashboard_settings")
            .select("setting_value")
            .eq("setting_key", FORM_CONFIG_KEY)
            .limit(1)
            .execute()
        )
        if result.data:
            return result.data[0].get("setting_value") or {}
    except Exception:
        pass
    return {}


def save_form_config(supabase, config):
    if not globals().get("CAN_ADMIN", False):
        raise PermissionError("Admin access is required to change form configuration.")
    normalized = normalize_form_config(config)
    supabase.table("dashboard_settings").upsert(
        {
            "setting_key": FORM_CONFIG_KEY,
            "setting_value": normalized,
            "updated_by": "admin-user",
            "updated_at": datetime.now().isoformat(),
        },
        on_conflict="setting_key",
    ).execute()
    st.cache_data.clear()
    return normalized

ROOT_CAUSE_NOTES = [
    "Change in Tank Anchor Plate Design",
    "Dome Roof Rafter Section & Numbers and Compression plate thickness",
    "Suspended deck Design and Instrument Nozzle and Stillwell Size",
    "Intank Pump Selection and its document finalization",
    "Product Bottom Inlet Pipe Size reduction and Vapour Disengagement Design",
    "VMS Selection and its document finalization",
    "Tank Shell Insulation Arrangement change",
    "Tank Cool Down Sensor Arrangement change",
    "Change in Nos. of PSV & VSV and Change in their Platform arrangement",
    "Change in the Pump Roof Top maintenance platform arrangement",
    "Change in ROV arrangement at Liquid Platform and Multilevel Pipe Rack",
    "BOG line Orientation and Supporting on Shell and Trestle Design",
]


# ----------------------------------------------------------------------------
# SUPABASE CLIENT
# ----------------------------------------------------------------------------
@st.cache_resource
def get_supabase():
    from supabase import create_client

    url = st.secrets.get("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        return None
    return create_client(url, key)


@st.cache_resource
def get_auth_supabase():
    """Public Supabase client used only for end-user authentication."""
    from supabase import create_client

    try:
        url = st.secrets.get("SUPABASE_URL")
        anon_key = st.secrets.get("SUPABASE_ANON_KEY")
    except Exception:
        url = os.getenv("SUPABASE_URL")
        anon_key = os.getenv("SUPABASE_ANON_KEY")
    if not url or not anon_key:
        return None
    return create_client(url, anon_key)


def _get_secret(name):
    try:
        value = st.secrets.get(name)
    except Exception:
        value = None
    return str(value or os.getenv(name, "")).strip()


def _master_key():
    return _get_secret("MASTER_ACCESS_KEY")


def _session_token_hash(token):
    return hashlib.sha256(str(token).encode("utf-8")).hexdigest()


def _cookie_session_token():
    try:
        return str(AUTH_COOKIES.get(AUTH_COOKIE_NAME) or "").strip()
    except Exception:
        return ""


def _write_session_cookie(token):
    AUTH_COOKIES[AUTH_COOKIE_NAME] = str(token)
    AUTH_COOKIES.save()


def _delete_session_cookie():
    try:
        if AUTH_COOKIE_NAME in AUTH_COOKIES:
            del AUTH_COOKIES[AUTH_COOKIE_NAME]
            AUTH_COOKIES.save()
    except Exception:
        pass


def _revoke_persistent_session(service_client):
    token = _cookie_session_token()
    if token:
        try:
            service_client.table("app_sessions").update({
                "revoked_at": datetime.now(timezone.utc).isoformat(),
            }).eq("token_hash", _session_token_hash(token)).execute()
        except Exception:
            pass
    _delete_session_cookie()


def _create_persistent_session(service_client, context):
    """Create a server-tracked opaque browser session.

    The browser receives only a random opaque token. Supabase refresh/access tokens
    are never placed in the browser cookie. A stolen database dump does not reveal
    usable browser tokens because only SHA-256 hashes are stored.
    """
    raw_token = secrets.token_urlsafe(48)
    now = datetime.now(timezone.utc)
    is_master = context.get("mode") == "master"
    expires_at = now + (timedelta(hours=MASTER_SESSION_HOURS) if is_master else timedelta(days=ACCOUNT_SESSION_DAYS))
    service_client.table("app_sessions").insert({
        "token_hash": _session_token_hash(raw_token),
        "user_id": context.get("user_id"),
        "access_mode": "master" if is_master else "account",
        "expires_at": expires_at.isoformat(),
        "created_at": now.isoformat(),
        "last_seen_at": now.isoformat(),
    }).execute()
    _write_session_cookie(raw_token)


def _clear_access_session(service_client=None, revoke=True):
    if revoke and service_client is not None:
        _revoke_persistent_session(service_client)
    for key in (
        "access_context", "auth_access_token", "auth_refresh_token",
        "active_project_id", "show_project_home", "pending_doc_changes",
        "master_data_configuration_unlocked",
    ):
        st.session_state.pop(key, None)


def _profile_for_user(service_client, user_id):
    result = (
        service_client.table("user_profiles")
        .select("*")
        .eq("id", str(user_id))
        .limit(1)
        .execute()
    )
    return result.data[0] if result.data else None


def _ensure_profile(service_client, user, full_name=""):
    if user is None:
        return None
    existing = _profile_for_user(service_client, user.id)
    if existing:
        return existing
    metadata = getattr(user, "user_metadata", None) or {}
    row = {
        "id": str(user.id),
        "email": getattr(user, "email", None),
        "full_name": str(full_name or metadata.get("full_name") or "").strip() or None,
        "role": "viewer",
        "approval_status": "pending",
        "is_active": True,
    }
    service_client.table("user_profiles").upsert(row, on_conflict="id").execute()
    return _profile_for_user(service_client, user.id)


def _set_user_access(user, session, profile):
    st.session_state["auth_access_token"] = getattr(session, "access_token", None)
    st.session_state["auth_refresh_token"] = getattr(session, "refresh_token", None)
    st.session_state["access_context"] = {
        "mode": "account",
        "user_id": str(user.id),
        "email": getattr(user, "email", "") or "",
        "name": (profile or {}).get("full_name") or getattr(user, "email", "") or "User",
        "role": (profile or {}).get("role") or "viewer",
        "approval_status": (profile or {}).get("approval_status") or "pending",
        "is_active": bool((profile or {}).get("is_active", True)),
    }


def _refresh_access_profile(service_client):
    context = st.session_state.get("access_context") or {}
    if context.get("mode") != "account" or not context.get("user_id"):
        return context
    profile = _profile_for_user(service_client, context["user_id"])
    if profile:
        context.update({
            "name": profile.get("full_name") or context.get("email") or "User",
            "email": profile.get("email") or context.get("email") or "",
            "role": profile.get("role") or "viewer",
            "approval_status": profile.get("approval_status") or "pending",
            "is_active": bool(profile.get("is_active", True)),
        })
        st.session_state["access_context"] = context
    return context


def _restore_persistent_access(service_client):
    if st.session_state.get("access_context"):
        return _refresh_access_profile(service_client)
    raw_token = _cookie_session_token()
    if not raw_token:
        return {}
    token_hash = _session_token_hash(raw_token)
    result = (
        service_client.table("app_sessions")
        .select("*")
        .eq("token_hash", token_hash)
        .is_("revoked_at", "null")
        .limit(1)
        .execute()
    )
    row = result.data[0] if result.data else None
    now = datetime.now(timezone.utc)
    if not row:
        _delete_session_cookie()
        return {}
    expires_at = pd.to_datetime(row.get("expires_at"), utc=True, errors="coerce")
    if pd.isna(expires_at) or expires_at.to_pydatetime() <= now:
        service_client.table("app_sessions").update({"revoked_at": now.isoformat()}).eq("id", row["id"]).execute()
        _delete_session_cookie()
        return {}

    if row.get("access_mode") == "master":
        context = {
            "mode": "master", "user_id": None, "email": "master-access",
            "name": "Master Access", "role": "master",
            "approval_status": "approved", "is_active": True,
        }
    else:
        profile = _profile_for_user(service_client, row.get("user_id"))
        if not profile:
            _delete_session_cookie()
            return {}
        context = {
            "mode": "account",
            "user_id": str(profile["id"]),
            "email": profile.get("email") or "",
            "name": profile.get("full_name") or profile.get("email") or "User",
            "role": profile.get("role") or "viewer",
            "approval_status": profile.get("approval_status") or "pending",
            "is_active": bool(profile.get("is_active", True)),
        }
    st.session_state["access_context"] = context
    service_client.table("app_sessions").update({"last_seen_at": now.isoformat()}).eq("id", row["id"]).execute()
    return context


def render_auth_gate(service_client):
    """Sign-in, sign-up, approval gate, master override, and persistent sessions."""
    auth_client = get_auth_supabase()
    context = _restore_persistent_access(service_client)

    if context:
        if context.get("mode") == "master":
            return context
        if not context.get("is_active", True):
            st.error("This account has been deactivated. Contact an administrator.")
        elif context.get("approval_status") == "approved":
            return context
        elif context.get("approval_status") == "rejected":
            st.error("This account request was rejected. Contact an administrator.")
        else:
            st.title("Account awaiting approval")
            st.info(
                f"You are signed in as **{context.get('email')}**, but an administrator "
                "must approve your account and assign project access before you can open project data."
            )
        if st.button("Sign out"):
            try:
                if auth_client:
                    auth_client.auth.sign_out()
            except Exception:
                pass
            _clear_access_session(service_client)
            st.rerun()
        st.stop()

    st.title("DCI Project Hub")
    st.caption("Secure sign-in for project dashboards and document control.")
    signin_tab, signup_tab, master_tab = st.tabs(["Sign in", "Sign up", "Master access"])

    with signin_tab:
        if auth_client is None:
            st.error("SUPABASE_ANON_KEY is not configured, so account sign-in is unavailable.")
        email = st.text_input("Email", key="signin_email")
        password = st.text_input("Password", type="password", key="signin_password")
        if st.button("Sign in", type="primary", disabled=auth_client is None, key="signin_button"):
            try:
                response = auth_client.auth.sign_in_with_password({
                    "email": email.strip(), "password": password,
                })
                user = response.user
                profile = _ensure_profile(service_client, user)
                _set_user_access(user, response.session, profile)
                _create_persistent_session(service_client, st.session_state["access_context"])
                st.rerun()
            except Exception as exc:
                st.error(f"Sign-in failed: {exc}")

        reset_email = st.text_input("Password-reset email", key="reset_email")
        if st.button("Send password reset", disabled=auth_client is None, key="reset_button"):
            try:
                auth_client.auth.reset_password_for_email(reset_email.strip())
                st.success("If that account exists, Supabase has sent a password-reset email.")
            except Exception as exc:
                st.error(f"Could not send reset email: {exc}")

    with signup_tab:
        if auth_client is None:
            st.error("SUPABASE_ANON_KEY is not configured, so account registration is unavailable.")
        full_name = st.text_input("Full name", key="signup_name")
        signup_email = st.text_input("Work email", key="signup_email")
        signup_password = st.text_input("Create password", type="password", key="signup_password")
        signup_password_2 = st.text_input("Confirm password", type="password", key="signup_password_2")
        signup_errors = []
        if signup_password and len(signup_password) < 8:
            signup_errors.append("Password must contain at least 8 characters.")
        if signup_password_2 and signup_password != signup_password_2:
            signup_errors.append("The passwords do not match.")
        for message in signup_errors:
            st.error(message)
        if st.button(
            "Create account", type="primary", key="signup_button",
            disabled=auth_client is None or bool(signup_errors),
        ):
            if not full_name.strip() or not signup_email.strip() or not signup_password:
                st.error("Full name, email, and password are required.")
            else:
                try:
                    response = auth_client.auth.sign_up({
                        "email": signup_email.strip(),
                        "password": signup_password,
                        "options": {"data": {"full_name": full_name.strip()}},
                    })
                    if response.user:
                        profile = _ensure_profile(service_client, response.user, full_name)
                        if response.session:
                            _set_user_access(response.user, response.session, profile)
                            _create_persistent_session(service_client, st.session_state["access_context"])
                            st.rerun()
                        else:
                            st.success(
                                "Account created. Confirm your email if Supabase sent a verification message, "
                                "then sign in. An administrator must approve the account and assign projects."
                            )
                except Exception as exc:
                    st.error(f"Could not create account: {exc}")

    with master_tab:
        st.warning("Master access grants full control. It expires after eight hours.")
        entered = st.text_input("Master key", type="password", key="master_access_key")
        if not _master_key():
            st.error("MASTER_ACCESS_KEY is not configured in secrets or the environment.")
        if st.button("Enter with master access", type="primary", key="master_access_button"):
            configured = _master_key()
            if configured and hmac.compare_digest(entered, configured):
                st.session_state["access_context"] = {
                    "mode": "master", "user_id": None, "email": "master-access",
                    "name": "Master Access", "role": "master",
                    "approval_status": "approved", "is_active": True,
                }
                _create_persistent_session(service_client, st.session_state["access_context"])
                st.rerun()
            else:
                st.error("Incorrect master key.")

    st.stop()


def _can_edit(context):
    return context.get("role") in {"editor", "admin", "master"}


def _can_admin(context):
    return context.get("role") in {"admin", "master"}


def _can_export(context):
    return context.get("role") in {"viewer", "editor", "admin", "master"}


def _actor_label(context):
    if context.get("mode") == "master":
        return "master-access"
    return context.get("email") or context.get("name") or "authenticated-user"


def _verify_admin_safety_check(context, entered_secret):
    """Require fresh credentials before changing access controls."""
    entered_secret = str(entered_secret or "")
    if context.get("mode") == "master":
        configured = _master_key()
        return bool(configured and hmac.compare_digest(entered_secret, configured))
    email = context.get("email") or ""
    if not email or not entered_secret:
        return False
    try:
        client = get_auth_supabase()
        response = client.auth.sign_in_with_password({"email": email, "password": entered_secret})
        return bool(response and response.user and str(response.user.id) == str(context.get("user_id")))
    except Exception:
        return False


def _direct_project_ids(service_client, user_id):
    result = service_client.table("user_project_access").select("project_id").eq("user_id", str(user_id)).execute()
    return {row["project_id"] for row in (result.data or [])}


def _group_project_ids(service_client, user_id):
    membership = service_client.table("project_group_members").select("group_id").eq("user_id", str(user_id)).execute()
    group_ids = [row["group_id"] for row in (membership.data or [])]
    if not group_ids:
        return set()
    groups = service_client.table("project_groups").select("project_id").in_("id", group_ids).eq("is_active", True).execute()
    return {row["project_id"] for row in (groups.data or [])}


def accessible_project_ids(service_client, context):
    if _can_admin(context):
        return None
    user_id = context.get("user_id")
    if not user_id:
        return set()
    return _direct_project_ids(service_client, user_id) | _group_project_ids(service_client, user_id)


def _user_group_ids(service_client, user_id):
    result = service_client.table("project_group_members").select("group_id").eq("user_id", str(user_id)).execute()
    return {row["group_id"] for row in (result.data or [])}


def render_user_management(service_client):
    """Advanced approval, project assignment, role and project-group management."""
    st.subheader("User access management")
    st.caption(
        "Search for one account by name or email, then review and update that user's access. "
        "Assign users directly to projects or place them in project groups."
    )

    projects = list_projects(service_client)
    project_by_id = {p["id"]: p for p in projects}
    project_label_to_id = {
        f"{p.get('project_name') or 'Unnamed project'} · #{p['id']}": p["id"] for p in projects
    }

    st.markdown("#### Project groups")
    groups_result = (
        service_client.table("project_groups")
        .select("*")
        .eq("is_active", True)
        .order("group_name")
        .execute()
    )
    groups = groups_result.data or []

    with st.container(border=True):
        st.markdown("##### ＋ Create project group")
        group_name = st.text_input("Group name", key="new_access_group_name")
        group_project_label = st.selectbox(
            "Project",
            list(project_label_to_id.keys()) if project_label_to_id else ["No projects available"],
            key="new_access_group_project",
        )
        group_description = st.text_area(
            "Description",
            placeholder="Example: Mechanical document-control team",
            key="new_access_group_description",
        )
        if st.button("Create group", key="create_access_group", disabled=not project_label_to_id):
            if not group_name.strip():
                st.error("Group name is required.")
            else:
                service_client.table("project_groups").insert({
                    "project_id": project_label_to_id[group_project_label],
                    "group_name": group_name.strip(),
                    "description": group_description.strip() or None,
                    "created_by": CURRENT_ACTOR,
                    "is_active": True,
                }).execute()
                st.success("Project group created.")
                st.rerun()

    if groups:
        group_rows = []
        for group in groups:
            members = (
                service_client.table("project_group_members")
                .select("user_id", count="exact")
                .eq("group_id", group["id"])
                .execute()
            )
            group_rows.append({
                "Group": group.get("group_name"),
                "Project": (project_by_id.get(group.get("project_id")) or {}).get(
                    "project_name", group.get("project_id")
                ),
                "Members": members.count or 0,
                "Description": group.get("description") or "",
                "Email scope": "Planned",
            })
        st.dataframe(pd.DataFrame(group_rows), use_container_width=True, hide_index=True)

    result = service_client.table("user_profiles").select("*").order("created_at", desc=True).execute()
    profiles = result.data or []
    if not profiles:
        st.caption("No registered account profiles yet.")
        return

    group_label_to_id = {
        f"{g.get('group_name')} · {(project_by_id.get(g.get('project_id')) or {}).get('project_name', 'Unknown project')}": g["id"]
        for g in groups
    }

    st.markdown("#### Find a user")
    st.caption("Click the dropdown and type any part of the person's name or email address.")

    # Keep user IDs as stable option values. Streamlit's selectbox supports type-to-search,
    # while format_func displays a useful name/email label.
    profile_by_id = {str(profile["id"]): profile for profile in profiles}
    ordered_user_ids = sorted(
        profile_by_id,
        key=lambda uid: (
            str(profile_by_id[uid].get("approval_status") or "pending") != "pending",
            str(profile_by_id[uid].get("full_name") or "").casefold(),
            str(profile_by_id[uid].get("email") or "").casefold(),
        ),
    )

    def _user_option_label(user_id):
        profile = profile_by_id[user_id]
        name = str(profile.get("full_name") or "Unnamed user").strip()
        email = str(profile.get("email") or user_id).strip()
        role = str(profile.get("role") or "viewer").title()
        status = str(profile.get("approval_status") or "pending").title()
        return f"{name} — {email} · {role} · {status}"

    selected_user_id = st.selectbox(
        "Search name or email",
        ordered_user_ids,
        format_func=_user_option_label,
        key="user_management_selected_user",
        placeholder="Choose or type to search for a user",
    )
    profile = profile_by_id.get(str(selected_user_id))
    if not profile:
        st.info("Select an account to manage its access.")
        return

    user_id = str(profile["id"])
    with st.container(border=True):
        top1, top2, top3 = st.columns([3, 1, 1])
        top1.markdown(f"### {profile.get('full_name') or 'Unnamed user'}")
        top1.markdown(f"[{profile.get('email') or user_id}](mailto:{profile.get('email') or ''})")
        top2.metric("Current role", str(profile.get("role") or "viewer").title())
        top3.metric("Status", str(profile.get("approval_status") or "pending").title())
        st.caption(f"Account created: {str(profile.get('created_at') or '')[:10]}")

        c1, c2, c3 = st.columns(3)
        role_options = ["viewer", "editor", "admin"]
        current_role = profile.get("role") if profile.get("role") in role_options else "viewer"
        new_role = c1.selectbox(
            "Role",
            role_options,
            index=role_options.index(current_role),
            key=f"profile_role_{user_id}",
        )
        status_options = ["pending", "approved", "rejected"]
        current_status = (
            profile.get("approval_status")
            if profile.get("approval_status") in status_options
            else "pending"
        )
        new_status = c2.selectbox(
            "Status",
            status_options,
            index=status_options.index(current_status),
            key=f"profile_status_{user_id}",
        )
        active = c3.checkbox(
            "Account active",
            value=bool(profile.get("is_active", True)),
            key=f"profile_active_{user_id}",
        )

        current_direct = _direct_project_ids(service_client, user_id)
        current_project_labels = [
            label for label, pid in project_label_to_id.items() if pid in current_direct
        ]
        selected_project_labels = st.multiselect(
            "Direct project access",
            list(project_label_to_id.keys()),
            default=current_project_labels,
            key=f"profile_projects_{user_id}",
            help="A user can access several projects. Group access is added on top of these selections.",
        )

        current_groups = _user_group_ids(service_client, user_id)
        current_group_labels = [
            label for label, gid in group_label_to_id.items() if gid in current_groups
        ]
        selected_group_labels = st.multiselect(
            "Project groups",
            list(group_label_to_id.keys()),
            default=current_group_labels,
            key=f"profile_groups_{user_id}",
            help="Membership grants access to the project linked to each group.",
        )

        # Show the effective project list before the admin saves anything.
        direct_project_ids = {project_label_to_id[label] for label in selected_project_labels}
        selected_group_ids = {group_label_to_id[label] for label in selected_group_labels}
        grouped_project_ids = {
            group.get("project_id") for group in groups if group.get("id") in selected_group_ids
        }
        effective_project_ids = direct_project_ids | grouped_project_ids
        effective_names = [
            (project_by_id.get(pid) or {}).get("project_name", f"Project #{pid}")
            for pid in sorted(effective_project_ids)
        ]
        st.info(
            "Effective project access: "
            + (", ".join(effective_names) if effective_names else "No projects assigned")
        )

        credentials_label = (
            "Confirm master key"
            if ACCESS_CONTEXT.get("mode") == "master"
            else "Confirm your admin password"
        )
        safety_secret = st.text_input(
            credentials_label,
            type="password",
            key=f"profile_safety_{user_id}",
            help=(
                "Fresh credentials are required before role, approval, active-state, "
                "project or group access changes are applied."
            ),
        )

        save_col, reset_col = st.columns([1, 5])
        if save_col.button("Save user access", key=f"save_profile_{user_id}", type="primary"):
            if not _verify_admin_safety_check(ACCESS_CONTEXT, safety_secret):
                st.error("Safety check failed. Enter your current admin password or master key.")
                return

            service_client.table("user_profiles").update({
                "role": new_role,
                "approval_status": new_status,
                "is_active": active,
                "approved_at": (
                    datetime.now(timezone.utc).isoformat() if new_status == "approved" else None
                ),
                "approved_by": CURRENT_ACTOR,
                "updated_at": datetime.now(timezone.utc).isoformat(),
            }).eq("id", user_id).execute()

            desired_projects = {
                project_label_to_id[label] for label in selected_project_labels
            }
            service_client.table("user_project_access").delete().eq("user_id", user_id).execute()
            if desired_projects:
                service_client.table("user_project_access").insert([
                    {"user_id": user_id, "project_id": pid, "granted_by": CURRENT_ACTOR}
                    for pid in sorted(desired_projects)
                ]).execute()

            desired_groups = {group_label_to_id[label] for label in selected_group_labels}
            service_client.table("project_group_members").delete().eq("user_id", user_id).execute()
            if desired_groups:
                service_client.table("project_group_members").insert([
                    {"user_id": user_id, "group_id": gid, "added_by": CURRENT_ACTOR}
                    for gid in sorted(desired_groups)
                ]).execute()

            service_client.table("app_sessions").update({
                "revoked_at": datetime.now(timezone.utc).isoformat(),
            }).eq("user_id", user_id).is_("revoked_at", "null").execute()
            st.success("User access updated. Existing remembered sessions were revoked.")
            st.rerun()

        if reset_col.button("Reload saved access", key=f"reload_profile_{user_id}"):
            # Clear this user's widget state so values are reconstructed from Supabase.
            for prefix in (
                "profile_role_", "profile_status_", "profile_active_",
                "profile_projects_", "profile_groups_", "profile_safety_",
            ):
                st.session_state.pop(f"{prefix}{user_id}", None)
            st.rerun()


# ----------------------------------------------------------------------------
# EXCEL TEMPLATE DISCOVERY + PARSING
# ----------------------------------------------------------------------------
def _normalise_header(value):
    """Normalise a worksheet header so harmless punctuation/case changes do not matter."""
    import re

    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ").strip().upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return " ".join(text.split())


HEADER_ALIASES = {
    "sr_no": {"SR NO", "SERIAL NO", "S NO"},
    "doc_no": {"DOCUMENT NUMBER", "DOCUMENT NO", "DOC NO", "STS DOCUMENT NUMBER"},
    "title": {"DOCUMENT TITLE", "TITLE", "DOCUMENT DESCRIPTION"},
    "discipline": {"DISCIPLINE"},
    "vendor": {"VENDOR NAME", "VENDOR", "POC", "VENDOR CONSULTANT POC"},
    "status": {"DCI STATUS", "STATUS", "FRESH RESUBMISSION"},
    "latest_rev": {"REV", "LATEST REV", "LATEST REVISION", "CURRENT REVISION"},
    "latest_uploaded_date": {"UPLOADED DATE", "ACTUAL SUBMISSION DATE", "REVISION DATE"},
    "eil_review_status": {"EIL REVIEW STATUS", "REVIEW STATUS"},
    "latest_code": {"CODE", "RETURN CODE", "RETURNED CODE"},
    "latest_from_eil_date": {"RECEIVED DATE", "FROM EIL DATE", "RETURN DATE", "REVIEW DATE"},
    "category": {"CATEGORY"},
    "l4_marked": {"L4 MARKED", "L 4 MARKED"},
    "schedule_submission_date": {"SCHEDULE SUBMISSION DATE", "PLANNED SUBMISSION DATE"},
    "remarks": {"REMARKS", "COMMENTS"},
    "eil_mandays": {"EIL MANDAYS", "EIL MAN DAYS"},
    "eil_manhours": {"EIL MANHOURS", "EIL MAN HOURS"},
    "stspl_mandays": {"STSPL MANDAYS", "STSPL MAN DAYS"},
    "stspl_manhours": {"STSPL MANHOURS", "STSPL MAN HOURS"},
    "revision_cycle": {"REVISION CYCLE"},
    "expected_manhours": {"EXPECTED MANHOURS", "EXPECTED MAN HOURS"},
}


def _header_matches(value, aliases):
    return _normalise_header(value) in aliases


def _looks_like_revision_header(value):
    """Recognise REV-0 / Rev 1 / LATEST REV. style history-block headers."""
    import re

    h = _normalise_header(value)
    if h in {"LATEST REV", "REVISION", "REV"}:
        return True
    return bool(re.fullmatch(r"REV(?:ISION)?\s*\d+", h))


def _looks_like_upload_header(value):
    return _normalise_header(value) in {
        "UPLOADED DATE", "UPLOAD DATE", "REVISION DATE", "ACTUAL SUBMISSION DATE"
    }


def _looks_like_code_header(value):
    return _normalise_header(value) in {"CODE", "RETURN CODE", "RETURNED CODE"}


def _looks_like_return_header(value):
    return _normalise_header(value) in {
        "FROM EIL DATE", "RECEIVED DATE", "RETURN DATE", "REVIEW DATE"
    }


def _row_header_values(ws, row):
    """Return normalized values for one row, expanding merged-cell anchors."""
    merged_lookup = {}
    for merged in ws.merged_cells.ranges:
        anchor = ws.cell(merged.min_row, merged.min_col).value
        for rr in range(merged.min_row, merged.max_row + 1):
            if rr != row:
                continue
            for cc in range(merged.min_col, merged.max_col + 1):
                merged_lookup[cc] = anchor

    values = []
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=row, column=col).value
        if value in (None, "") and col in merged_lookup:
            value = merged_lookup[col]
        values.append(_normalise_header(value))
    return values


def _header_row_score(ws, row):
    """Score how likely a row is to be the main DCI table header."""
    values = set(_row_header_values(ws, row))
    has_doc = bool(values & HEADER_ALIASES["doc_no"])
    supporting = sum(
        bool(values & HEADER_ALIASES[key])
        for key in ("title", "discipline", "vendor", "latest_rev", "latest_code")
    )
    revision_signals = sum(1 for value in values if _looks_like_revision_header(value))
    score = (100 if has_doc else 0) + supporting * 12 + min(revision_signals, 10)
    return score, has_doc, supporting


def _find_header_row(ws, scan_rows=60):
    """Discover the table-header row from semantic header names.

    The engine scans the top part of the sheet and chooses the strongest row
    containing Document Number plus at least two supporting DCI headers. It does
    not assume row 5, row 8, or any other fixed position.
    """
    candidates = []
    limit = min(scan_rows, ws.max_row)
    for row in range(1, limit + 1):
        score, has_doc, supporting = _header_row_score(ws, row)
        if has_doc and supporting >= 2:
            candidates.append((score, row))

    if not candidates:
        raise ValueError(
            "Could not locate the DCI table header. The workbook must contain a "
            "Document Number header and at least two of Document Title, Discipline, "
            "Vendor, Revision, or Code."
        )

    candidates.sort(reverse=True)
    return candidates[0][1]


def _extract_revision_label(*values):
    """Extract an explicit revision identifier from REV-0 / Revision A labels.

    Generic summary headers such as ``REV.`` or ``LATEST REV.`` intentionally
    return None so they are not mistaken for revision-history blocks.
    """
    import re

    for value in values:
        text = _normalise_header(value)
        match = re.fullmatch(r"REV(?:ISION)?\s*[-_:]?\s*([A-Z0-9][A-Z0-9._-]*)", text)
        if match:
            label = match.group(1).strip("._-")
            if label:
                return label
    return None


def _discover_revision_blocks(ws, header_row):
    """Discover revision-history blocks from a two/three-row header band.

    Supported examples include:
      REV-0 | UPLOADED DATE | CODE | FROM EIL DATE
    and a merged parent label above a detail row such as:
      [Rev-0 merged across four columns]
      REV.  | REVISION DATE | RETURN CODE | RETURN DATE

    Columns can move and extra unrelated columns can appear between blocks.
    """
    detail_rows = range(max(1, header_row - 2), min(ws.max_row, header_row + 1) + 1)
    blocks = []
    used_revision_columns = set()

    for detail_row in detail_rows:
        for col in range(1, ws.max_column + 1):
            first = ws.cell(detail_row, col).value
            parent1 = ws.cell(detail_row - 1, col).value if detail_row > 1 else None
            parent2 = ws.cell(detail_row - 2, col).value if detail_row > 2 else None

            revision_label = _extract_revision_label(first, parent1, parent2)
            if revision_label is None:
                continue

            # Locate the three companion headers within a bounded window. This
            # tolerates inserted helper columns without pairing with another block.
            window_end = min(ws.max_column, col + 7)
            upload_col = code_col = return_col = None
            for candidate in range(col + 1, window_end + 1):
                value = ws.cell(detail_row, candidate).value
                if upload_col is None and _looks_like_upload_header(value):
                    upload_col = candidate
                elif code_col is None and _looks_like_code_header(value):
                    code_col = candidate
                elif return_col is None and _looks_like_return_header(value):
                    return_col = candidate

            if not all((upload_col, code_col, return_col)):
                continue
            if not (col < upload_col < code_col < return_col):
                continue
            if col in used_revision_columns:
                continue

            blocks.append({
                "revision": col,
                "uploaded_date": upload_col,
                "code": code_col,
                "from_eil_date": return_col,
                "label": revision_label,
                "header_row": detail_row,
            })
            used_revision_columns.add(col)

    blocks.sort(key=lambda block: block["revision"])
    if not blocks:
        raise ValueError(
            "No revision-history blocks were found. Expected headers equivalent to "
            "Revision | Uploaded/Revision Date | Code | From EIL/Return Date."
        )
    return blocks


def _find_first_matching_column(ws, header_row, aliases, *, before_col=None, after_col=None):
    start = max(1, (after_col or 0) + 1)
    end = min(ws.max_column, (before_col - 1) if before_col else ws.max_column)
    for row in range(max(1, header_row - 2), min(ws.max_row, header_row + 1) + 1):
        for col in range(start, end + 1):
            if _header_matches(ws.cell(row, col).value, aliases):
                return col
    return None


def _discover_data_start_row(ws, header_row, doc_no_col, scan_limit=30):
    """Find the first actual document row below the discovered header."""
    max_row = min(ws.max_row, header_row + scan_limit)
    for row in range(header_row + 1, max_row + 1):
        value = ws.cell(row=row, column=doc_no_col).value
        if value is not None and str(value).strip():
            return row
    # Empty templates are valid: data will begin immediately below the header.
    return header_row + 1


def discover_dci_layout(ws):
    """Build a complete, runtime-discovered template profile for a DCI sheet."""
    header_row = _find_header_row(ws)
    revision_blocks = _discover_revision_blocks(ws, header_row)
    first_revision_col = min(block["revision"] for block in revision_blocks)
    last_revision_col = max(block["from_eil_date"] for block in revision_blocks)

    columns = {}
    for key in (
        "sr_no", "doc_no", "title", "discipline", "vendor", "status",
        "latest_rev", "latest_uploaded_date", "eil_review_status", "latest_code",
        "latest_from_eil_date", "category", "l4_marked", "schedule_submission_date",
    ):
        columns[key] = _find_first_matching_column(
            ws, header_row, HEADER_ALIASES[key], before_col=first_revision_col
        )

    for key in (
        "remarks", "eil_mandays", "eil_manhours", "stspl_mandays",
        "stspl_manhours", "revision_cycle", "expected_manhours",
    ):
        columns[key] = _find_first_matching_column(
            ws, header_row, HEADER_ALIASES[key], after_col=last_revision_col
        )
        if columns[key] is None:
            columns[key] = _find_first_matching_column(ws, header_row, HEADER_ALIASES[key])

    missing = [name for name in ("doc_no", "title", "discipline") if columns.get(name) is None]
    if missing:
        raise ValueError(
            "The DCI header was found, but required columns are missing: " + ", ".join(missing)
        )

    data_start_row = _discover_data_start_row(ws, header_row, columns["doc_no"])
    return {
        "engine_version": TEMPLATE_ENGINE_VERSION,
        "sheet_name": ws.title,
        "header_row": header_row,
        "data_start_row": data_start_row,
        "columns": columns,
        "revision_blocks": revision_blocks,
    }


def discover_dci_sheet_and_layout(workbook):
    """Discover the DCI sheet and its layout from any supplied workbook.

    Exact sheet name 'DCI' is preferred, but a normal Excel file may use another
    sheet name. In that case every visible worksheet is scored semantically.
    """
    candidates = []
    preferred = ["DCI"] if "DCI" in workbook.sheetnames else []
    sheet_names = preferred + [name for name in workbook.sheetnames if name not in preferred]

    errors = []
    for name in sheet_names:
        ws = workbook[name]
        if ws.sheet_state != "visible":
            continue
        try:
            layout = discover_dci_layout(ws)
            score, _, supporting = _header_row_score(ws, layout["header_row"])
            score += len(layout["revision_blocks"]) * 5 + supporting
            if name == "DCI":
                score += 1000
            candidates.append((score, ws, layout))
        except ValueError as exc:
            errors.append(f"{name}: {exc}")

    if not candidates:
        detail = "; ".join(errors[:4])
        raise ValueError(
            "No compatible DCI worksheet could be discovered in this workbook. "
            + (f"Checked sheets: {detail}" if detail else "")
        )

    candidates.sort(key=lambda item: item[0], reverse=True)
    _, ws, layout = candidates[0]
    return ws, layout


def layout_summary(layout):
    """Return human-readable discovery diagnostics for import preview/debugging."""
    from openpyxl.utils import get_column_letter

    detected = {}
    for field, column in layout["columns"].items():
        if column:
            detected[field] = f"{get_column_letter(column)} ({column})"
    blocks = []
    for block in layout["revision_blocks"]:
        blocks.append({
            "label": block.get("label"),
            "revision": get_column_letter(block["revision"]),
            "revision_date": get_column_letter(block["uploaded_date"]),
            "code": get_column_letter(block["code"]),
            "return_date": get_column_letter(block["from_eil_date"]),
        })
    return {
        "sheet": layout["sheet_name"],
        "header_row": layout["header_row"],
        "data_start_row": layout["data_start_row"],
        "fields": detected,
        "revision_blocks": blocks,
    }

def _find_label_value(ws, label, scan_rows=12):
    wanted = _normalise_header(label)
    for row in range(1, min(scan_rows, ws.max_row) + 1):
        for col in range(1, min(ws.max_column, 20) + 1):
            if _normalise_header(ws.cell(row, col).value) != wanted:
                continue
            # Values can be in the next cell or farther right because of merges.
            for next_col in range(col + 1, min(ws.max_column, col + 12) + 1):
                value = ws.cell(row, next_col).value
                if value not in (None, ""):
                    return value
    return None


def _cell_value(ws, row, columns, field):
    col = columns.get(field)
    return ws.cell(row=row, column=col).value if col else None



def _normalise_import_code(value):
    """Convert template labels such as Code-2 / CODE 1 to canonical DCI codes."""
    import re

    if value is None or value == "":
        return None
    text = str(value).strip().upper()
    text = re.sub(r"^CODE\s*[-:]?\s*", "", text)
    text = " ".join(text.split())
    return text or None

def parse_workbook_bytes(file_bytes):
    """Parse any DCI workbook matching the supplied template's header names.

    The header row, data start row and revision blocks are discovered at runtime;
    moving the table from row 8 to row 5 no longer breaks imports.
    """
    import openpyxl

    _validate_xlsx_container(file_bytes)
    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes),
        data_only=True,
        read_only=False,
        keep_links=True,
    )
    ws, layout = discover_dci_sheet_and_layout(wb)
    columns = layout["columns"]

    project_name = _find_label_value(ws, "PROJECT")
    contractor = _find_label_value(ws, "CONTRACTOR")
    records, history = [], []

    consecutive_empty_rows = 0
    for r in range(layout["data_start_row"], ws.max_row + 1):
        doc_no = _cell_value(ws, r, columns, "doc_no")
        if doc_no is None or str(doc_no).strip() == "":
            consecutive_empty_rows += 1
            if records and consecutive_empty_rows >= MAX_CONSECUTIVE_EMPTY_ROWS:
                break
            continue
        consecutive_empty_rows = 0
        if _normalise_header(doc_no) in HEADER_ALIASES["doc_no"]:
            # Ignore repeated headers sometimes inserted between print sections.
            continue

        vendor = _cell_value(ws, r, columns, "vendor")
        rec = {
            "sr_no": _cell_value(ws, r, columns, "sr_no"),
            "doc_no": str(doc_no).strip(),
            "title": _cell_value(ws, r, columns, "title"),
            "discipline": _cell_value(ws, r, columns, "discipline"),
            "poc_submission": vendor,
            "latest_rev": _cell_value(ws, r, columns, "latest_rev"),
            "latest_uploaded_date": _date_jsonable(_cell_value(ws, r, columns, "latest_uploaded_date")),
            "eil_review_status": _cell_value(ws, r, columns, "eil_review_status"),
            "latest_code": _normalise_import_code(_cell_value(ws, r, columns, "latest_code")),
            "latest_from_eil_date": _date_jsonable(_cell_value(ws, r, columns, "latest_from_eil_date")),
            "category": _cell_value(ws, r, columns, "category"),
            "l4_marked": _cell_value(ws, r, columns, "l4_marked"),
            "schedule_submission_date": _date_jsonable(_cell_value(ws, r, columns, "schedule_submission_date")),
            "status": _cell_value(ws, r, columns, "status"),
            "poc": vendor,
            "eil_mandays": _cell_value(ws, r, columns, "eil_mandays"),
            "eil_manhours": _cell_value(ws, r, columns, "eil_manhours"),
            "stspl_mandays": _cell_value(ws, r, columns, "stspl_mandays"),
            "stspl_manhours": _cell_value(ws, r, columns, "stspl_manhours"),
            "revision_cycle": _cell_value(ws, r, columns, "revision_cycle"),
            "expected_manhours": _cell_value(ws, r, columns, "expected_manhours"),
            "remarks": _cell_value(ws, r, columns, "remarks") or "",
            "assigned_engineer": "",
            "action_owner": "",
            "expected_completion_date": None,
            "forecast_month": "",
            "approval_date": None,
            "original_row_number": r,
        }
        rec["document_key"] = f"{rec['doc_no']}::ROW-{r}"

        row_history = []
        for block in layout["revision_blocks"]:
            rev_no = ws.cell(row=r, column=block["revision"]).value
            uploaded = ws.cell(row=r, column=block["uploaded_date"]).value
            code = ws.cell(row=r, column=block["code"]).value
            from_eil = ws.cell(row=r, column=block["from_eil_date"]).value
            if all(value in (None, "") for value in (rev_no, uploaded, code, from_eil)):
                continue
            row_history.append({
                "doc_no": rec["doc_no"],
                "rev_no": rev_no,
                "uploaded_date": _date_jsonable(uploaded),
                "code": _normalise_import_code(code),
                "from_eil_date": _date_jsonable(from_eil),
            })

        # Some templates calculate the current summary fields from revision blocks.
        # If cached formula results are absent, fall back to the last populated block.
        if row_history:
            latest_history = row_history[-1]
            if rec["latest_rev"] in (None, ""):
                rec["latest_rev"] = latest_history["rev_no"]
            if rec["latest_uploaded_date"] in (None, ""):
                rec["latest_uploaded_date"] = latest_history["uploaded_date"]
            if rec["latest_code"] in (None, ""):
                rec["latest_code"] = latest_history["code"]
            if rec["latest_from_eil_date"] in (None, ""):
                rec["latest_from_eil_date"] = latest_history["from_eil_date"]

        rec["_history_rows"] = row_history
        records.append(rec)
        history.extend(row_history)

    return records, history, project_name, contractor


def _jsonable(value):
    """Make an openpyxl cell value safe to store in JSONB."""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return value


def _date_jsonable(value):
    """Convert genuine Excel/Python dates to ISO without accepting return codes."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, pd.Timestamp):
        return None if pd.isna(value) else value.date().isoformat()

    # Numeric Excel serials may appear when the cell lacks a date format.
    # Accept only plausible modern serials; short numeric return codes remain invalid.
    if isinstance(value, (int, float, np.integer, np.floating)):
        if pd.isna(value) or not (20000 <= float(value) <= 80000):
            return None
        try:
            return pd.to_datetime(float(value), unit="D", origin="1899-12-30").date().isoformat()
        except Exception:
            return None

    text = str(value).strip()
    if not text or _normalise_header(text) in {"R", "V", "CODE 1", "CODE 2", "CODE 3"}:
        return None
    try:
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
        return None if pd.isna(parsed) else parsed.date().isoformat()
    except Exception:
        return None



def _validate_xlsx_container(file_bytes):
    """Reject oversized, renamed, encrypted, or malformed uploads before parsing."""
    if not isinstance(file_bytes, (bytes, bytearray)) or not file_bytes:
        raise ValueError("The uploaded workbook is empty.")
    if len(file_bytes) > MAX_UPLOAD_BYTES:
        raise ValueError(
            f"The workbook is larger than the allowed {MAX_UPLOAD_BYTES // (1024 * 1024)} MB."
        )
    if not zipfile.is_zipfile(io.BytesIO(file_bytes)):
        raise ValueError("This is not a valid .xlsx workbook.")
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as archive:
            names = set(archive.namelist())
            required = {"[Content_Types].xml", "xl/workbook.xml"}
            if not required.issubset(names):
                raise ValueError("The uploaded file is not a valid Excel .xlsx package.")
            encrypted_markers = {"EncryptedPackage", "EncryptionInfo"}
            if names & encrypted_markers:
                raise ValueError("Password-protected Excel workbooks are not supported.")
    except zipfile.BadZipFile as exc:
        raise ValueError("The uploaded workbook is damaged or incomplete.") from exc


def workbook_hash(file_bytes):
    """Stable content fingerprint used for exact-duplicate detection."""
    return hashlib.sha256(file_bytes).hexdigest()


def validate_records(records):
    """Validate parsed document rows without guessing or silently discarding problems."""
    warnings, errors = [], []
    if not records:
        errors.append(
            "The workbook structure was recognised, but no document rows containing "
            "a Document Number were found."
        )
        return warnings, errors

    seen = {}
    duplicate_numbers = set()
    for record in records:
        doc_no = str(record.get("doc_no") or "").strip()
        if not doc_no:
            errors.append(
                f"Excel row {record.get('original_row_number', '?')} has no Document Number."
            )
            continue
        if doc_no in seen:
            duplicate_numbers.add(doc_no)
        else:
            seen[doc_no] = record.get("original_row_number")

    if duplicate_numbers:
        sample = ", ".join(sorted(duplicate_numbers)[:10])
        warnings.append(
            f"{len(duplicate_numbers)} duplicate Document Number value(s) were found: {sample}"
            + (" …" if len(duplicate_numbers) > 10 else "")
            + ". They will remain separate records using their original Excel row numbers."
        )

    missing_titles = [
        record.get("original_row_number")
        for record in records
        if not str(record.get("title") or "").strip()
    ]
    if missing_titles:
        warnings.append(f"{len(missing_titles)} document row(s) have no Document Title.")

    invalid_history_dates = 0
    for record in records:
        for event in record.get("_history_rows", []):
            uploaded = pd.to_datetime(event.get("uploaded_date"), errors="coerce")
            returned = pd.to_datetime(event.get("from_eil_date"), errors="coerce")
            if pd.notna(uploaded) and pd.notna(returned) and returned < uploaded:
                invalid_history_dates += 1
    if invalid_history_dates:
        warnings.append(
            f"{invalid_history_dates} revision block(s) have a Return Date earlier than "
            "their Revision Date. They will be retained but excluded from turnaround metrics."
        )

    return warnings, errors


def inspect_workbook_bytes(file_bytes):
    """Validate and discover a workbook once, returning reusable diagnostics."""
    import openpyxl

    _validate_xlsx_container(file_bytes)
    workbook = openpyxl.load_workbook(
        io.BytesIO(file_bytes),
        data_only=False,
        read_only=False,
        keep_links=True,
    )
    worksheet, layout = discover_dci_sheet_and_layout(workbook)
    return worksheet, layout, layout_summary(layout)


# ----------------------------------------------------------------------------
# IMPORT WORKFLOW (Supabase writes)
# ----------------------------------------------------------------------------
def find_import_by_hash(supabase, file_hash, project_id=None):
    query = (
        supabase.table("imports")
        .select("*")
        .eq("workbook_hash", file_hash)
        .eq("is_active", True)
    )
    if project_id is not None:
        query = query.eq("project_id", project_id)
    res = query.order("uploaded_at", desc=True).limit(1).execute()
    return res.data[0] if res.data else None


def list_imports(supabase, project_id=None):
    query = (
        supabase.table("imports")
        .select("*")
        .eq("is_active", True)
    )
    if project_id is not None:
        query = query.eq("project_id", project_id)
    res = query.order("uploaded_at", desc=True).execute()
    return res.data or []


def list_projects(supabase, access_context=None):
    query = (
        supabase.table("projects")
        .select("*")
        .eq("is_active", True)
    )
    if access_context is not None:
        allowed = accessible_project_ids(supabase, access_context)
        if allowed is not None:
            if not allowed:
                return []
            query = query.in_("id", sorted(allowed))
    res = query.order("created_at", desc=True).execute()
    return res.data or []


def get_project(supabase, project_id):
    res = supabase.table("projects").select("*").eq("id", project_id).single().execute()
    return res.data


def list_templates(supabase):
    res = (
        supabase.table("dci_templates")
        .select("*")
        .eq("is_active", True)
        .order("created_at", desc=True)
        .execute()
    )
    return res.data or []


def create_template(supabase, *, name, file_bytes, filename, description=""):
    """Validate, store, and register a reusable DCI template."""
    if not globals().get("CAN_ADMIN", False):
        raise PermissionError("Admin access is required to create templates.")
    _ws, layout, diagnostics = inspect_workbook_bytes(file_bytes)
    template_row = {
        "template_name": str(name).strip() or filename,
        "description": str(description or "").strip(),
        "original_filename": filename,
        "storage_bucket": STORAGE_BUCKET,
        "storage_path": "pending",
        "workbook_hash": workbook_hash(file_bytes),
        "engine_version": int(layout.get("engine_version", TEMPLATE_ENGINE_VERSION)),
        "layout_definition": layout,
        "is_active": True,
    }
    created = supabase.table("dci_templates").insert(template_row).execute()
    if not created.data:
        raise RuntimeError("Supabase did not return the new template row.")
    template_id = created.data[0]["id"]
    path = f"templates/{template_id}/template.xlsx"
    try:
        supabase.storage.from_(STORAGE_BUCKET).upload(
            path=path,
            file=file_bytes,
            file_options={
                "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "upsert": "true",
            },
        )
        supabase.table("dci_templates").update({"storage_path": path}).eq("id", template_id).execute()
        return template_id, diagnostics
    except Exception:
        supabase.table("dci_templates").update({"is_active": False}).eq("id", template_id).execute()
        raise


def create_project(supabase, *, project_name, contractor, client, review_period_days, template_id):
    if not globals().get("CAN_ADMIN", False):
        raise PermissionError("Admin access is required to create projects.")
    row = {
        "project_name": str(project_name).strip(),
        "contractor": str(contractor or "").strip() or None,
        "client": str(client or "").strip() or None,
        "review_period_days": int(review_period_days),
        "template_id": template_id,
        "is_active": True,
    }
    created = supabase.table("projects").insert(row).execute()
    if not created.data:
        raise RuntimeError("Supabase did not return the new project row.")
    st.cache_data.clear()
    return created.data[0]


def get_template_bytes(supabase, template_id):
    row = supabase.table("dci_templates").select("*").eq("id", template_id).single().execute().data
    if not row:
        raise ValueError("The selected template no longer exists.")
    return supabase.storage.from_(row.get("storage_bucket") or STORAGE_BUCKET).download(row["storage_path"]), row


def upload_workbook_to_storage(supabase, import_id, file_bytes, filename):
    path = f"imports/{import_id}/original.xlsx"
    supabase.storage.from_(STORAGE_BUCKET).upload(
        path=path,
        file=file_bytes,
        file_options={
            "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "upsert": "true",
        },
    )
    return path


def run_import(supabase, file_bytes, filename, project_name, contractor,
                review_period_days, version_mode, parent_import_id=None,
                project_id=None, template_id=None):
    if not globals().get("CAN_ADMIN", False):
        raise PermissionError("Admin access is required to import workbooks.")
    """Import a workbook without deactivating the current version prematurely.

    Supabase/PostgREST calls are not a single database transaction from Python, so
    this function uses a safe publish sequence: create pending import, upload and
    populate it, mark it complete, and only then archive the replaced import.
    """
    _worksheet, discovered_layout, _diagnostics = inspect_workbook_bytes(file_bytes)
    records, _history, wb_project_name, wb_contractor = parse_workbook_bytes(file_bytes)
    warnings, errors = validate_records(records)
    if errors:
        raise ValueError("; ".join(errors))

    file_hash = workbook_hash(file_bytes)
    version_number = 1
    if version_mode == "new_version" and parent_import_id:
        parent = (
            supabase.table("imports")
            .select("version_number")
            .eq("id", parent_import_id)
            .limit(1)
            .execute()
        )
        if parent.data:
            version_number = int(parent.data[0].get("version_number") or 1) + 1

    import_row = {
        "project_id": project_id,
        "template_id": template_id,
        "original_filename": filename,
        "display_name": filename,
        "project_name": project_name or wb_project_name,
        "contractor": contractor or wb_contractor,
        "version_number": version_number,
        "parent_import_id": parent_import_id if version_mode == "new_version" else None,
        "source_sheet_name": discovered_layout["sheet_name"],
        "row_count": len(records),
        "workbook_hash": file_hash,
        "storage_bucket": STORAGE_BUCKET,
        "original_storage_path": "pending",
        "review_period_days": int(review_period_days),
        "validation_status": "warnings" if warnings else "clean",
        "import_status": "pending",
        "warning_count": len(warnings),
        "error_count": 0,
        "is_active": True,
    }
    created = supabase.table("imports").insert(import_row).execute()
    if not created.data:
        raise RuntimeError("Supabase did not return the newly created import row.")
    import_id = created.data[0]["id"]

    storage_path = None
    try:
        storage_path = upload_workbook_to_storage(supabase, import_id, file_bytes, filename)
        supabase.table("imports").update({"original_storage_path": storage_path}).eq("id", import_id).execute()
        _bulk_import_documents(supabase, import_id, records)
        supabase.table("imports").update({"import_status": "complete"}).eq("id", import_id).execute()

        if version_mode == "replace" and parent_import_id:
            supabase.table("imports").update({
                "is_active": False,
                "import_status": "archived",
            }).eq("id", parent_import_id).execute()

        st.cache_data.clear()
        return import_id, warnings
    except Exception:
        # Keep the previous active import untouched. Mark this import failed so it
        # cannot be selected as a valid workbook. Cascading cleanup can be done by
        # an admin without destroying evidence needed to diagnose the failure.
        supabase.table("imports").update({
            "is_active": False,
            "import_status": "failed",
        }).eq("id", import_id).execute()
        raise


def _bulk_import_documents(supabase, import_id, records, batch_size=200):
    for batch_start in range(0, len(records), batch_size):
        batch = records[batch_start:batch_start + batch_size]

        doc_rows = []
        for rec in batch:
            doc_rows.append({
                "import_id": import_id,
                "document_key": rec["document_key"],
                "original_row_number": rec["original_row_number"],
                "document_number": rec["doc_no"],
                "document_title": rec["title"],
                "discipline": rec["discipline"],
                "vendor": rec["poc"],
                "revision": None if rec["latest_rev"] is None else str(rec["latest_rev"]),
                "dci_status": rec["status"],
                "returned_code": None if rec["latest_code"] is None else str(rec["latest_code"]),
                "planned_submission_date": _date_jsonable(rec["schedule_submission_date"]),
                "actual_submission_date": _date_jsonable(rec["latest_uploaded_date"]),
                "review_date": _date_jsonable(rec["latest_from_eil_date"]),
                "approval_date": None,
                "forecast_month": "",
                "remarks": "",
                "assigned_engineer": "",
                "action_owner": "",
                "expected_completion_date": None,
                "last_changed_by": "import",
                "is_active": True,
            })

        inserted = supabase.table("documents").insert(doc_rows).execute()
        inserted_rows = inserted.data or []
        if len(inserted_rows) != len(doc_rows):
            raise RuntimeError(
                f"Inserted {len(inserted_rows)} of {len(doc_rows)} documents in a batch."
            )
        inserted_by_key = {row["document_key"]: row for row in inserted_rows}

        raw_rows, history_rows = [], []
        for rec in batch:
            doc = inserted_by_key.get(rec["document_key"])
            if doc is None:
                raise RuntimeError(f"Could not resolve inserted document: {rec['document_key']}")
            record_json = {k: _jsonable(v) for k, v in rec.items() if not k.startswith("_")}
            raw_rows.append({
                "document_id": doc["id"],
                "row_data": {"record": record_json, "history_rows": rec["_history_rows"]},
                "original_column_names": list(record_json.keys()),
                "original_row_number": rec["original_row_number"],
            })
            history_rows.append({
                "document_id": doc["id"],
                "import_id": import_id,
                "event_type": "imported",
                "field_name": None,
                "old_value": None,
                "new_value": None,
                "revision": doc.get("revision"),
                "returned_code": doc.get("returned_code"),
                "status": doc.get("dci_status"),
                "remarks": None,
                "changed_by": "import",
                "source": "import",
            })

        if raw_rows:
            supabase.table("document_raw_data").insert(raw_rows).execute()
        if history_rows:
            supabase.table("document_history").insert(history_rows).execute()


# ----------------------------------------------------------------------------
# REPOSITORY — load documents back out of Supabase into the shapes the
# existing analytical code (derive_fields / derive_history) already expects.
# ----------------------------------------------------------------------------
def _fetch_all(query_builder, page_size=1000):
    """Page through a Supabase query using .range() since PostgREST caps at 1000 rows."""
    rows, start = [], 0
    while True:
        res = query_builder.range(start, start + page_size - 1).execute()
        chunk = res.data or []
        rows.extend(chunk)
        if len(chunk) < page_size:
            break
        start += page_size
    return rows


@st.cache_data(ttl=60, show_spinner=False)
def load_documents(_supabase, import_id):
    """Returns (raw_df, raw_hist) shaped exactly like the original Excel parse,
    reconstructed from Supabase (raw JSON snapshot + current editable fields).
    The leading underscore on _supabase tells st.cache_data not to hash the client.
    """
    doc_rows = _fetch_all(_supabase.table("documents").select("*").eq("import_id", import_id).eq("is_active", True))
    if not doc_rows:
        return pd.DataFrame(columns=RECORD_COLUMNS), pd.DataFrame(columns=HISTORY_COLUMNS)

    doc_ids = [d["id"] for d in doc_rows]
    raw_rows = []
    for i in range(0, len(doc_ids), 200):
        chunk_ids = doc_ids[i:i + 200]
        res = _supabase.table("document_raw_data").select("*").in_("document_id", chunk_ids).execute()
        raw_rows.extend(res.data or [])
    raw_by_doc_id = {r["document_id"]: r for r in raw_rows}

    records, history = [], []
    for doc in doc_rows:
        raw = raw_by_doc_id.get(doc["id"], {})
        row_data = raw.get("row_data") or {}
        record = dict(row_data.get("record") or {})

        # Overlay the current (possibly edited) values from the documents table,
        # since those are the source of truth after any dashboard edit.
        for doc_field, record_field in DOC_TO_RECORD_FIELD.items():
            if doc_field in doc:
                record[record_field] = doc[doc_field]

        record["document_key"] = doc["document_key"]
        record["document_id"] = doc["id"]
        record["original_row_number"] = doc["original_row_number"]
        records.append(record)

        for h in (row_data.get("history_rows") or []):
            history_item = dict(h)
            history_item["document_id"] = doc["id"]
            history.append(history_item)

    df = pd.DataFrame(records)
    for col in RECORD_COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df[RECORD_COLUMNS]

    hist = pd.DataFrame(history)
    for col in HISTORY_COLUMNS:
        if col not in hist.columns:
            hist[col] = None
    hist = hist[HISTORY_COLUMNS] if not hist.empty else pd.DataFrame(columns=HISTORY_COLUMNS)

    return df, hist


def get_import_meta(supabase, import_id):
    res = supabase.table("imports").select("*").eq("id", import_id).single().execute()
    return res.data


def get_document_history(supabase, document_id):
    res = (
        supabase.table("document_history")
        .select("*")
        .eq("document_id", document_id)
        .order("changed_at", desc=False)
        .execute()
    )
    return res.data or []


def get_history_for_documents(supabase, document_ids, batch_size=200):
    """Fetch history for many documents in a handful of batched requests
    (one .in_() query per batch_size ids) instead of one request per document.
    Use this instead of looping get_document_history() over a whole workbook —
    with hundreds of documents that loop is slow and prone to timing out.
    """
    document_ids = [d for d in document_ids if d is not None and not (isinstance(d, float) and pd.isna(d))]
    rows = []
    for i in range(0, len(document_ids), batch_size):
        chunk = document_ids[i:i + batch_size]
        res = (
            supabase.table("document_history")
            .select("*")
            .in_("document_id", chunk)
            .order("changed_at", desc=False)
            .execute()
        )
        rows.extend(res.data or [])
    return rows


DATE_RECORD_FIELDS = {
    "schedule_submission_date", "latest_uploaded_date", "latest_from_eil_date",
    "approval_date", "expected_completion_date",
}


def _normalise_record_value(field, value):
    """Normalize form/database scalars so previews and saves compare identically."""
    if field in DATE_RECORD_FIELDS:
        return _date_to_iso(value)
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def build_change_diff(form_values, current_record):
    diff = []
    for record_field, new_value in form_values.items():
        old_value = current_record.get(record_field)
        new_norm = _normalise_record_value(record_field, new_value)
        old_norm = _normalise_record_value(record_field, old_value)
        if new_norm != old_norm:
            diff.append({
                "field": record_field,
                "old": old_norm,
                "new": new_norm,
            })
    return diff


def save_document_changes(supabase, document_id, form_values, current_record, changed_by, event_type):
    """Persist changed fields atomically through the database RPC."""
    if not globals().get("CAN_EDIT", False):
        raise PermissionError("Editor access is required to update documents.")
    record_diff = build_change_diff(form_values, current_record)
    if not record_diff:
        return False, []

    changes = {}
    old_values = {}
    for item in record_diff:
        doc_field = next(
            (db_field for db_field, record_field in DOC_TO_RECORD_FIELD.items()
             if record_field == item["field"]),
            None,
        )
        if doc_field is None:
            continue
        changes[doc_field] = item["new"] or None
        old_values[doc_field] = item["old"] or None

    if not changes:
        return False, []

    supabase.rpc("fn_save_document_changes", {
        "p_document_id": document_id,
        "p_changes": changes,
        "p_old_values": old_values,
        "p_event_type": event_type,
        "p_changed_by": changed_by or "dashboard-user",
    }).execute()

    st.cache_data.clear()
    display_diff = [
        {
            "field": item["field"],
            "old": item["old"] or "(blank)",
            "new": item["new"] or "(blank)",
        }
        for item in record_diff
    ]
    return True, display_diff


def _date_to_iso(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, pd.Timestamp):
        if pd.isna(val):
            return None
        return val.date().isoformat()
    if isinstance(val, (date, datetime)):
        return val.isoformat()[:10]
    if isinstance(val, str):
        text = val.strip()
        if not text:
            return None
        parsed = pd.to_datetime(text, errors="coerce")
        return None if pd.isna(parsed) else parsed.date().isoformat()
    return None


# ----------------------------------------------------------------------------
# EXPORT — regenerate an .xlsx from the original workbook + current DB values
# ----------------------------------------------------------------------------
# Export targets are discovered from the workbook headers at runtime.
# No Excel letters or fixed numeric column positions are used.
# Fields with no home in the original layout get appended as new columns.
EXTRA_EXPORT_COLUMNS = ["remarks", "assigned_engineer", "action_owner",
                         "expected_completion_date", "forecast_month", "approval_date"]
EXTRA_EXPORT_HEADERS = {
    "remarks": "Remarks (Dashboard)",
    "assigned_engineer": "Assigned Engineer (Dashboard)",
    "action_owner": "Action Owner (Dashboard)",
    "expected_completion_date": "Expected Completion Date (Dashboard)",
    "forecast_month": "Forecast Month (Dashboard)",
    "approval_date": "Approval Date (Dashboard)",
}


def _is_formula_cell(cell):
    """Return True when an openpyxl cell contains an Excel formula."""
    return isinstance(cell.value, str) and cell.value.startswith("=")


def _excel_scalar(value):
    """Convert database ISO dates back to genuine Excel date values where possible."""
    if value in (None, ""):
        return None
    if isinstance(value, pd.Timestamp):
        return None if pd.isna(value) else value.to_pydatetime()
    if isinstance(value, (date, datetime)):
        return value
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed) and len(text) >= 8:
            return parsed.to_pydatetime()
        return text
    return value


def _write_preserving_formula(ws, row, column, value):
    """Write only into input cells, never over an existing Excel formula."""
    if not column:
        return False
    cell = ws.cell(row=row, column=column)
    if _is_formula_cell(cell):
        return False
    old_number_format = cell.number_format
    converted = _excel_scalar(value)
    cell.value = converted
    if isinstance(converted, (date, datetime)) and old_number_format == "General":
        cell.number_format = "dd-mmm-yyyy"
    return True


def _normalise_revision_for_match(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip().upper()


def _find_revision_block(ws, row_number, revision, revision_blocks):
    """Find the existing block for revision, otherwise the first empty block."""
    wanted = _normalise_revision_for_match(revision)
    first_empty = None

    for block in revision_blocks:
        existing_rev = ws.cell(row=row_number, column=block["revision"]).value
        existing_upload = ws.cell(row=row_number, column=block["uploaded_date"]).value
        existing_code = ws.cell(row=row_number, column=block["code"]).value
        existing_return = ws.cell(row=row_number, column=block["from_eil_date"]).value

        if wanted and _normalise_revision_for_match(existing_rev) == wanted:
            return block

        if first_empty is None and all(v in (None, "") for v in (
            existing_rev, existing_upload, existing_code, existing_return
        )):
            first_empty = block

    return first_empty


def generate_updated_workbook(supabase, import_id):
    if not globals().get("CAN_EXPORT", False):
        raise PermissionError("An approved Viewer, Editor, Admin, or Master account is required to export workbooks.")
    import openpyxl

    meta = get_import_meta(supabase, import_id)
    original_bytes = supabase.storage.from_(STORAGE_BUCKET).download(meta["original_storage_path"])

    # data_only=False is essential: it loads formula expressions, not cached results.
    wb = openpyxl.load_workbook(io.BytesIO(original_bytes), data_only=False)
    ws, layout = discover_dci_sheet_and_layout(wb)
    columns = layout["columns"]

    doc_rows = _fetch_all(
        supabase.table("documents").select("*").eq("import_id", import_id).eq("is_active", True)
    )

    header_row = layout["header_row"]
    existing_extra_columns = {}
    for col in range(1, ws.max_column + 1):
        header = _normalise_header(ws.cell(header_row, col).value)
        for field, label in EXTRA_EXPORT_HEADERS.items():
            if header == _normalise_header(label):
                existing_extra_columns[field] = col

    next_extra_col = ws.max_column + 1
    extra_columns = {}
    for field in EXTRA_EXPORT_COLUMNS:
        if field in existing_extra_columns:
            extra_columns[field] = existing_extra_columns[field]
        else:
            extra_columns[field] = next_extra_col
            ws.cell(row=header_row, column=next_extra_col, value=EXTRA_EXPORT_HEADERS[field])
            next_extra_col += 1

    history_block_warnings = []
    for doc in doc_rows:
        r = doc["original_row_number"]

        # Update discovered summary cells only when they are not formulas.
        # Formula cells are retained and will recalculate from history blocks.
        export_fields = {
            "title": doc.get("document_title"),
            "latest_rev": doc.get("revision"),
            "latest_uploaded_date": doc.get("actual_submission_date"),
            "status": doc.get("dci_status"),
            "latest_code": doc.get("returned_code"),
            "latest_from_eil_date": doc.get("review_date"),
            "schedule_submission_date": doc.get("planned_submission_date"),
        }
        for layout_field, value in export_fields.items():
            col = columns.get(layout_field)
            if col:
                _write_preserving_formula(ws, r, col, value)

        # Write the current revision into its actual four-column revision block:
        # revision | revision/upload date | return code | return date.
        revision = doc.get("revision")
        if revision not in (None, ""):
            block = _find_revision_block(ws, r, revision, layout["revision_blocks"])
            if block is None:
                history_block_warnings.append(
                    f"{doc.get('document_number')}: no empty revision block remained for revision {revision}"
                )
            else:
                _write_preserving_formula(ws, r, block["revision"], revision)
                _write_preserving_formula(ws, r, block["uploaded_date"], doc.get("actual_submission_date"))
                _write_preserving_formula(ws, r, block["code"], doc.get("returned_code"))
                _write_preserving_formula(ws, r, block["from_eil_date"], doc.get("review_date"))

        for i, field in enumerate(EXTRA_EXPORT_COLUMNS):
            ws.cell(row=r, column=extra_columns[field], value=doc.get(field))

    # History sheet
    history_sheet_name = "Dashboard Change History"
    if history_sheet_name in wb.sheetnames:
        del wb[history_sheet_name]
    hs = wb.create_sheet(history_sheet_name)
    headers = ["Document Number", "Document Title", "Event Type", "Field Changed", "Old Value",
               "New Value", "Revision", "Returned Code", "Status", "Remarks", "Changed By", "Changed At"]
    hs.append(headers)

    doc_lookup = {d["id"]: d for d in doc_rows}
    all_history = []
    for i in range(0, len(doc_rows), 200):
        chunk_ids = [d["id"] for d in doc_rows[i:i + 200]]
        res = supabase.table("document_history").select("*").in_("document_id", chunk_ids) \
            .order("changed_at").execute()
        all_history.extend(res.data or [])

    for h in all_history:
        doc = doc_lookup.get(h["document_id"], {})
        hs.append([
            doc.get("document_number"), doc.get("document_title"), h.get("event_type"),
            h.get("field_name"), h.get("old_value"), h.get("new_value"), h.get("revision"),
            h.get("returned_code"), h.get("status"), h.get("remarks"), h.get("changed_by"),
            h.get("changed_at"),
        ])

    if history_block_warnings:
        warning_sheet = "Dashboard Export Warnings"
        if warning_sheet in wb.sheetnames:
            del wb[warning_sheet]
        warning_ws = wb.create_sheet(warning_sheet)
        warning_ws.append(["Warning"])
        for warning in history_block_warnings:
            warning_ws.append([warning])

    # Ask Excel to recalculate retained formulas when the exported workbook opens.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"{(meta.get('project_name') or 'DCI').replace(' ', '_')}_DCI_Updated_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
    return buf.getvalue(), filename


def record_export(supabase, import_id, export_type, filename, row_count, exported_by=None):
    supabase.table("export_records").insert({
        "import_id": import_id,
        "export_type": export_type,
        "exported_by": exported_by or globals().get("CURRENT_ACTOR", "authenticated-user"),
        "filename": filename,
        "row_count": row_count,
        "success": True,
    }).execute()


# ----------------------------------------------------------------------------
# DERIVED-FIELD CALCULATIONS (unchanged from the original dashboard)
# ----------------------------------------------------------------------------
def clean_dates(df, cols):
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NaT
        else:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def normalize_code(code):
    if pd.isna(code):
        return ""
    if isinstance(code, (int, float)) and float(code).is_integer():
        return str(int(code))
    return " ".join(str(code).strip().upper().split())


def classify_code(code):
    value = normalize_code(code)
    if value in APPROVED_CODES:
        return "Approved (Code 1)"
    if value in RETAINED_CODES:
        return "Retained for Record (Code R)"
    if value in CODE2_CODES:
        return "Code 2 - Comments/Resubmit"
    if value in CODE3_CODES:
        return "Code 3 - Rejected/Resubmit"
    if value in VOID_CODES:
        return "Void"
    return "No code yet"


def normalize_status(value):
    if pd.isna(value):
        return ""
    status = " ".join(str(value).strip().upper().split())
    aliases = {
        "WITH EIL FOR REVIEW": "WITH EIL",
        "PENDING WITH EIL": "WITH EIL",
        "FRESH SUBMISSION PENDING": "FRESH SUBMISSION",
        "CODE 1": "APPROVED",
    }
    return aliases.get(status, status)


def status_fallback(row):
    status = normalize_status(row.get("status"))
    if status:
        return status
    if pd.isna(row.get("latest_uploaded_date")):
        return "FRESH SUBMISSION"
    if row["code_bucket"] in ("Approved (Code 1)", "Retained for Record (Code R)"):
        return "APPROVED"
    if pd.notna(row.get("latest_from_eil_date")):
        return "RETURNED BY EIL"
    return "WITH EIL"


def aging_bucket(age):
    if pd.isna(age):
        return None
    if age < 0:
        return "Not due yet"
    if age <= 7:
        return "0-7 days"
    if age <= 14:
        return "8-14 days"
    if age <= 21:
        return "15-21 days"
    return ">21 days"


def derive_fields(df, review_period_days, today):
    df = df.copy()
    df = clean_dates(df, ["latest_uploaded_date", "latest_from_eil_date", "schedule_submission_date"])

    for col in NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["code_bucket"] = df["latest_code"].apply(classify_code)
    df["is_void"] = df["code_bucket"] == "Void"

    df["status_clean"] = df.apply(status_fallback, axis=1)

    df["is_submitted"] = df["latest_uploaded_date"].notna() | (df["status_clean"] != "FRESH SUBMISSION")
    df["is_reviewed"] = df["latest_from_eil_date"].notna()
    df["is_approved_final"] = df["code_bucket"].isin(
        ["Approved (Code 1)", "Retained for Record (Code R)"]
    )

    df["submission_delay_days"] = np.where(
        df["latest_uploaded_date"].isna() & df["schedule_submission_date"].notna(),
        (today - df["schedule_submission_date"]).dt.days,
        np.nan,
    )
    df["review_age_days"] = np.where(
        df["latest_uploaded_date"].notna()
        & df["latest_from_eil_date"].isna()
        & ~df["is_approved_final"]
        & ~df["is_void"],
        (today - df["latest_uploaded_date"]).dt.days,
        np.nan,
    )
    df["is_submission_overdue"] = df["submission_delay_days"] > 0
    df["is_review_overdue"] = df["review_age_days"] > review_period_days

    def compute_age(row):
        if row["is_void"] or row["is_approved_final"]:
            return np.nan
        if row["status_clean"] == "FRESH SUBMISSION":
            base = row["schedule_submission_date"]
        else:
            base = row["latest_uploaded_date"]
        if pd.isna(base):
            return np.nan
        return (today - base).days

    df["age_days"] = df.apply(compute_age, axis=1)
    df["aging_bucket"] = df["age_days"].apply(aging_bucket)
    df["is_overdue"] = df["is_submission_overdue"].fillna(False) | df["is_review_overdue"].fillna(False)

    df["turnaround_days"] = (df["latest_from_eil_date"] - df["latest_uploaded_date"]).dt.days

    for c in ["latest_uploaded_date", "latest_from_eil_date", "schedule_submission_date"]:
        df[c + "_month"] = df[c].dt.to_period("M").astype(str)

    df["latest_rev_display"] = df["latest_rev"].apply(
        lambda v: "" if pd.isna(v) else str(v).strip()
    )

    return df


def derive_history(hist):
    hist = hist.copy()
    hist = clean_dates(hist, ["uploaded_date", "from_eil_date"])
    hist["code_bucket"] = hist["code"].apply(classify_code)
    hist["turnaround_days"] = (hist["from_eil_date"] - hist["uploaded_date"]).dt.days

    invalid_turnaround = hist["turnaround_days"] < 0
    invalid_count = int(invalid_turnaround.sum())
    if invalid_count:
        hist.loc[invalid_turnaround, "turnaround_days"] = np.nan
    st.session_state["_invalid_turnaround_count"] = invalid_count

    hist["uploaded_month"] = hist["uploaded_date"].dt.to_period("M").astype(str)
    hist["returned_month"] = hist["from_eil_date"].dt.to_period("M").astype(str)
    return hist


def _timeline_date(value):
    """Return a timezone-naive Timestamp for safe timeline sorting."""
    ts = pd.to_datetime(value, errors="coerce", utc=True)
    if pd.isna(ts):
        return pd.NaT
    return ts.tz_convert(None)


def _revision_sort_key(value):
    """Sort common revision values naturally: 0, 1, 2, A, B, etc."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return (2, "")
    text = str(value).strip()
    try:
        return (0, float(text))
    except ValueError:
        return (1, text.casefold())


def _next_sequential_revision(current_revision):
    """Return the only valid next numeric revision, or None for non-numeric schemes."""
    value = "" if current_revision is None else str(current_revision).strip()
    if value.isdigit():
        return str(int(value) + 1)
    return None


def _revision_transition_is_valid(current_revision, selected_revision):
    """Allow correcting the current revision or advancing by exactly one numeric step."""
    current = "" if current_revision is None else str(current_revision).strip()
    selected = "" if selected_revision is None else str(selected_revision).strip()
    if selected == current:
        return True
    expected = _next_sequential_revision(current)
    return expected is not None and selected == expected


def build_document_timeline(revision_history, audit_events):
    """Merge Excel revision history with later Supabase audit events.

    Excel dates represent the real document lifecycle. The import audit event only
    represents when that historical record was added to this dashboard.
    """
    timeline = []

    if revision_history is not None and not revision_history.empty:
        revisions = revision_history.copy()
        revisions = revisions.sort_values(
            by="rev_no", key=lambda col: col.map(_revision_sort_key)
        )

        for _, rev in revisions.iterrows():
            rev_text = "" if pd.isna(rev.get("rev_no")) else str(rev.get("rev_no")).strip()
            code = normalize_code(rev.get("code"))
            uploaded = _timeline_date(rev.get("uploaded_date"))
            returned = _timeline_date(rev.get("from_eil_date"))

            if pd.notna(uploaded):
                timeline.append({
                    "event_date": uploaded,
                    "event": "Revision submitted",
                    "revision": rev_text,
                    "code": "",
                    "status": "Submitted to EIL",
                    "details": f"Revision {rev_text or 'N/A'} uploaded/submitted for review.",
                    "source": "Excel revision history",
                })

            if pd.notna(returned):
                bucket = classify_code(code)
                if bucket in ("Approved (Code 1)", "Retained for Record (Code R)"):
                    event_name = "Revision approved"
                elif bucket == "Void":
                    event_name = "Revision voided"
                else:
                    event_name = "Revision returned by EIL"

                details = f"EIL returned Revision {rev_text or 'N/A'}"
                if code:
                    details += f" with Code {code}"
                if pd.notna(uploaded):
                    turnaround = (returned.normalize() - uploaded.normalize()).days
                    if turnaround >= 0:
                        details += f" after {turnaround} day{'s' if turnaround != 1 else ''}"
                details += "."

                timeline.append({
                    "event_date": returned,
                    "event": event_name,
                    "revision": rev_text,
                    "code": _normalise_import_code(code),
                    "status": bucket,
                    "details": details,
                    "source": "Excel revision history",
                })

            # Preserve partially populated revision blocks instead of hiding them.
            if pd.isna(uploaded) and pd.isna(returned) and (rev_text or code):
                timeline.append({
                    "event_date": pd.NaT,
                    "event": "Revision information recorded",
                    "revision": rev_text,
                    "code": _normalise_import_code(code),
                    "status": classify_code(code),
                    "details": "Revision data exists in Excel, but no usable event date was available.",
                    "source": "Excel revision history",
                })

    for ev in audit_events or []:
        event_type = ev.get("event_type") or "updated"
        changed_at = _timeline_date(ev.get("changed_at"))

        if event_type == "imported":
            timeline.append({
                "event_date": changed_at,
                "event": "Added to dashboard",
                "revision": "" if ev.get("revision") is None else str(ev.get("revision")),
                "code": normalize_code(ev.get("returned_code")),
                "status": ev.get("status") or "",
                "details": "The existing Excel document and its earlier revision history were imported into Supabase.",
                "source": "Dashboard audit log",
            })
            continue

        field_name = (ev.get("field_name") or "document").replace("_", " ").title()
        old_value = ev.get("old_value")
        new_value = ev.get("new_value")
        if old_value is not None or new_value is not None:
            details = f"{field_name}: {old_value or '(blank)'} → {new_value or '(blank)'}"
        else:
            details = field_name

        timeline.append({
            "event_date": changed_at,
            "event": event_type.replace("_", " ").title(),
            "revision": "" if ev.get("revision") is None else str(ev.get("revision")),
            "code": normalize_code(ev.get("returned_code")),
            "status": ev.get("status") or "",
            "details": details,
            "source": "Dashboard audit log",
        })

    # Dated events first in chronological order; incomplete undated Excel rows last.
    timeline.sort(key=lambda e: (pd.isna(e["event_date"]), e["event_date"] if pd.notna(e["event_date"]) else pd.Timestamp.max))
    return timeline


# ----------------------------------------------------------------------------
# SHARED CHART HELPER
# ----------------------------------------------------------------------------
def render_monthly_grouped_chart(data, date_col, title, count_label="Count"):
    d = data.dropna(subset=[date_col]).copy()
    if d.empty:
        st.info(f"No data available for '{title}' with the current filters.")
        return

    d["year"] = d[date_col].dt.year
    d["month_num"] = d[date_col].dt.month

    years = sorted(d["year"].unique(), reverse=True)
    grid_index = pd.MultiIndex.from_product([years, range(1, 13)], names=["year", "month_num"])
    grid = (
        d.groupby(["year", "month_num"]).size()
        .reindex(grid_index, fill_value=0)
        .reset_index(name=count_label)
    )
    grid["month_name"] = grid["month_num"].apply(lambda m: MONTH_ORDER[m - 1])
    grid["year"] = grid["year"].astype(str)

    year_labels = [str(y) for y in years]
    color_map = {y: CHART_PALETTE[i % len(CHART_PALETTE)] for i, y in enumerate(year_labels)}

    with st.container(border=True):
        fig = px.bar(
            grid,
            x="month_name",
            y=count_label,
            color="year",
            barmode="group",
            category_orders={"month_name": MONTH_ORDER, "year": year_labels},
            color_discrete_map=color_map,
        )
        fig.update_traces(marker_line_color="rgba(255,255,255,0.6)", marker_line_width=1)
        fig.update_layout(
            title=dict(
                text=title,
                x=0.5,
                xanchor="center",
                font=dict(family="Arial Black, Arial, sans-serif", size=22, color="#595959"),
            ),
            yaxis_title="No. of Deliverables",
            xaxis_title="",
            yaxis=dict(gridcolor="#e6e6e6", zeroline=False),
            xaxis=dict(showgrid=False),
            plot_bgcolor="white",
            paper_bgcolor="white",
            bargap=0.28,
            bargroupgap=0.04,
            legend=dict(
                orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5,
                title_text="", font=dict(size=13),
            ),
            margin=dict(t=70, b=10, l=10, r=10),
        )
        st.plotly_chart(fig, use_container_width=True)

        header_cells = "".join(
            f"<th style='padding:5px 10px;border:1px solid #d9d9d9;text-align:center;'>{m}</th>"
            for m in MONTH_ORDER
        )
        row_html = []
        for y in year_labels:
            color = color_map[y]
            series = grid.loc[grid["year"] == y].set_index("month_name").reindex(MONTH_ORDER)[count_label]
            value_cells = "".join(
                f"<td style='padding:5px 10px;border:1px solid #d9d9d9;text-align:center;'>{int(v)}</td>"
                for v in series
            )
            row_html.append(
                "<tr>"
                f"<td style='padding:5px 10px;border:1px solid #d9d9d9;white-space:nowrap;'>"
                f"<span style='display:inline-block;width:10px;height:10px;background:{color};"
                f"border:1px solid #999;margin-right:6px;'></span>{y}</td>"
                f"{value_cells}</tr>"
            )
        table_html = f"""
        <table style='border-collapse:collapse;width:100%;font-size:14px;color:#333;margin-top:4px;'>
            <tr><th style='padding:5px 10px;border:1px solid #d9d9d9;'></th>{header_cells}</tr>
            {''.join(row_html)}
        </table>
        """
        st.markdown(table_html, unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# SUPABASE CONNECTION GATE
# ----------------------------------------------------------------------------
supabase = get_supabase()
if supabase is None:
    st.title("BPCL Uran - Document Control Index Dashboard")
    st.error("Supabase is not configured yet.")
    st.markdown(
        "Add these to your Streamlit secrets (`.streamlit/secrets.toml` locally, "
        "or **Settings → Secrets** on Streamlit Community Cloud) and reload:\n\n"
        "```toml\n"
        'SUPABASE_URL = "https://your-project.supabase.co"\n'
        'SUPABASE_SERVICE_ROLE_KEY = "your-service-role-key"\n'
        "```\n\n"
        "Also run `supabase_schema.sql` once in your Supabase project's SQL editor "
        "(Table Editor → SQL Editor) to create the required tables, and create a "
        f"**private** Storage bucket named `{STORAGE_BUCKET}`."
    )
    st.stop()

# Authenticate before any project data or server-side write controls are rendered.
ACCESS_CONTEXT = render_auth_gate(supabase)
CAN_EDIT = _can_edit(ACCESS_CONTEXT)
CAN_ADMIN = _can_admin(ACCESS_CONTEXT)
CAN_EXPORT = _can_export(ACCESS_CONTEXT)
CURRENT_ACTOR = _actor_label(ACCESS_CONTEXT)

# ----------------------------------------------------------------------------
# PROJECT HOME / PROJECT SELECTION
# ----------------------------------------------------------------------------
def _clear_pending_upload():
    for key in (
        "pending_upload_bytes", "pending_upload_name", "create_project_template_bytes",
        "create_project_template_name",
    ):
        st.session_state.pop(key, None)


def render_project_home(supabase):
    st.title("DCI Project Home")
    st.caption("Open an existing project or create a new one from a reusable Excel template.")
    st.caption(f"Signed in as **{ACCESS_CONTEXT.get('name')}** · Role: **{ACCESS_CONTEXT.get('role').title()}**")

    if CAN_ADMIN:
        with st.expander("👥 User Access Management", expanded=False):
            render_user_management(supabase)

    projects = list_projects(supabase, ACCESS_CONTEXT)
    if projects:
        st.subheader("Projects")
        cards_per_row = 3
        for start in range(0, len(projects), cards_per_row):
            cols = st.columns(cards_per_row)
            for col, project in zip(cols, projects[start:start + cards_per_row]):
                with col.container(border=True):
                    st.markdown(f"### {project.get('project_name') or 'Unnamed Project'}")
                    details = []
                    if project.get("client"):
                        details.append(f"Client: {project['client']}")
                    if project.get("contractor"):
                        details.append(f"Contractor: {project['contractor']}")
                    if details:
                        st.caption("  \n".join(details))
                    p_imports = list_imports(supabase, project["id"])
                    latest = p_imports[0] if p_imports else None
                    if latest:
                        st.metric("Documents", int(latest.get("row_count") or 0))
                        st.caption(f"Latest workbook: {latest.get('display_name') or latest.get('original_filename')}")
                    else:
                        st.metric("Documents", 0)
                        st.caption("Template ready · no workbook imported yet")
                    if st.button("Open Project", key=f"open_project_{project['id']}", type="primary", use_container_width=True):
                        st.session_state["active_project_id"] = project["id"]
                        st.session_state["show_project_home"] = False
                        _clear_pending_upload()
                        st.rerun()
    else:
        if CAN_ADMIN:
            st.info("No projects exist yet. Create the first project below.")
        else:
            st.warning("Your account is approved, but no project has been assigned to you yet. Contact an administrator.")

    st.markdown("---")
    if not CAN_ADMIN:
        st.info("Only an Admin or Master user can create a new project.")
        return
    st.subheader("＋ Create New Project")
    with st.container(border=True):
        project_name = st.text_input("Project name *", key="new_project_name")
        c1, c2, c3 = st.columns(3)
        contractor = c1.text_input("Contractor", key="new_project_contractor")
        client = c2.text_input("Client", key="new_project_client")
        review_days = c3.number_input(
            "Allowed EIL review period (days)", min_value=1, max_value=90, value=21,
            key="new_project_review_days",
        )

        templates = list_templates(supabase)
        template_mode_options = ["Upload a new template"]
        if templates:
            template_mode_options.insert(0, "Use an existing template")
        template_mode = st.radio("Template source", template_mode_options, horizontal=True)

        selected_template_id = None
        uploaded_template = None
        template_name = ""
        template_description = ""
        if template_mode == "Use an existing template":
            template_labels = {
                f"{t['template_name']} · {t.get('original_filename') or 'Excel template'}": t["id"]
                for t in templates
            }
            selected_label = st.selectbox("Choose template", list(template_labels.keys()))
            selected_template_id = template_labels[selected_label]
            selected = next(t for t in templates if t["id"] == selected_template_id)
            st.caption(selected.get("description") or "This template will define the workbook structure for the project.")
        else:
            template_name = st.text_input("Template name *", placeholder="Example: Standard BPCL DCI Template")
            template_description = st.text_area("Template description", height=80)
            uploaded_template = st.file_uploader(
                "Upload an empty template or a normal populated DCI workbook (.xlsx)",
                type=["xlsx"], key="create_project_template_uploader",
            )
            if uploaded_template is not None:
                try:
                    _ws, _layout, diagnostic = inspect_workbook_bytes(uploaded_template.getvalue())
                    st.success(
                        f"Template detected: sheet {diagnostic['sheet']}, header row "
                        f"{diagnostic['header_row']}, {len(diagnostic['revision_blocks'])} revision blocks."
                    )
                except Exception as exc:
                    st.error(f"Template discovery failed: {exc}")

        if st.button("Create Project", type="primary", key="create_project_button"):
            if not str(project_name).strip():
                st.error("Project name is required.")
            elif template_mode == "Upload a new template" and uploaded_template is None:
                st.error("Upload a template workbook before creating the project.")
            elif template_mode == "Upload a new template" and not str(template_name).strip():
                st.error("Template name is required.")
            else:
                try:
                    with st.spinner("Creating project..."):
                        if template_mode == "Upload a new template":
                            selected_template_id, _diagnostic = create_template(
                                supabase,
                                name=template_name,
                                file_bytes=uploaded_template.getvalue(),
                                filename=uploaded_template.name,
                                description=template_description,
                            )
                        created_project = create_project(
                            supabase,
                            project_name=project_name,
                            contractor=contractor,
                            client=client,
                            review_period_days=review_days,
                            template_id=selected_template_id,
                        )
                    st.session_state["active_project_id"] = created_project["id"]
                    st.session_state["show_project_home"] = False
                    _clear_pending_upload()
                    st.success("Project created.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Could not create project: {exc}")


try:
    _available_projects = list_projects(supabase, ACCESS_CONTEXT)
except Exception as exc:
    st.error("The multi-project database migration has not been applied yet.")
    st.code(str(exc))
    st.info("Run `project_home_migration.sql` in the Supabase SQL Editor, then reload the app.")
    st.stop()

if "show_project_home" not in st.session_state:
    st.session_state["show_project_home"] = st.session_state.get("active_project_id") is None

if st.session_state.get("show_project_home") or st.session_state.get("active_project_id") is None:
    render_project_home(supabase)
    st.stop()

active_project_id = st.session_state.get("active_project_id")
try:
    active_project = get_project(supabase, active_project_id)
except Exception:
    active_project = None
if not active_project:
    st.session_state.pop("active_project_id", None)
    st.session_state["show_project_home"] = True
    st.rerun()
_allowed_project_ids = accessible_project_ids(supabase, ACCESS_CONTEXT)
if _allowed_project_ids is not None and active_project_id not in _allowed_project_ids:
    st.session_state.pop("active_project_id", None)
    st.session_state["show_project_home"] = True
    st.error("You no longer have access to that project.")
    st.rerun()

# ----------------------------------------------------------------------------
# SIDEBAR
# ----------------------------------------------------------------------------
st.sidebar.title("DCI Dashboard")
st.sidebar.caption(f"{ACCESS_CONTEXT.get('name')} · {ACCESS_CONTEXT.get('role').title()}")
if st.sidebar.button("Sign out", use_container_width=True):
    try:
        auth_client = get_auth_supabase()
        if auth_client and ACCESS_CONTEXT.get("mode") == "account":
            auth_client.auth.sign_out()
    except Exception:
        pass
    _clear_access_session(supabase)
    st.rerun()
if st.sidebar.button("← Project Home", use_container_width=True):
    st.session_state["show_project_home"] = True
    _clear_pending_upload()
    st.rerun()
st.sidebar.caption(f"Project: **{active_project.get('project_name', 'Unnamed Project')}**")

with st.sidebar.expander("Workbook", expanded=True):
    imports = list_imports(supabase, active_project_id)
    import_options = {f"{i['display_name']} (v{i['version_number']}, {i['row_count']} docs)": i["id"] for i in imports}

    if imports:
        selected_label = st.selectbox("Active workbook", list(import_options.keys()))
        active_import_id = import_options[selected_label]
        active_import = next(i for i in imports if i["id"] == active_import_id)
        st.caption(
            f"Uploaded: {active_import['uploaded_at'][:10]}  \n"
            f"Status: {active_import['import_status']}  \n"
            f"Rows: {active_import['row_count']}"
        )
    else:
        active_import_id = None
        active_import = None
        st.info("No workbook imported yet. Upload one in the **Import / Export** tab.")

    if CAN_ADMIN:
        new_upload = st.file_uploader("Upload a new DCI workbook (.xlsx)", type=["xlsx"], key="sidebar_uploader")
        if new_upload is not None:
            st.session_state["pending_upload_bytes"] = new_upload.getvalue()
            st.session_state["pending_upload_name"] = new_upload.name
            st.info("Go to the **Import / Export** tab to validate and confirm this upload.")
    else:
        st.caption("Workbook import is available to Admin and Master users only.")

st.sidebar.markdown("---")
with st.sidebar.expander("Project Settings", expanded=True):
    default_review_days = active_import["review_period_days"] if active_import else int(active_project.get("review_period_days") or 21)
    review_period_days = st.number_input(
        "Allowed EIL review period (days)", min_value=1, max_value=90, value=default_review_days
    )

today = pd.Timestamp(datetime.today().date())

if active_import_id is not None:
    raw_df, raw_hist = load_documents(supabase, active_import_id)
else:
    raw_df, raw_hist = pd.DataFrame(columns=RECORD_COLUMNS), pd.DataFrame(columns=HISTORY_COLUMNS)

has_data = not raw_df.empty

if has_data:
    df = derive_fields(raw_df, review_period_days, today)
    hist = derive_history(raw_hist)
    ever_reviewed_docs = set(hist.loc[hist["from_eil_date"].notna(), "doc_no"])
    df["ever_reviewed"] = df["doc_no"].isin(ever_reviewed_docs)
else:
    df = pd.DataFrame(columns=RECORD_COLUMNS)
    hist = pd.DataFrame(columns=HISTORY_COLUMNS)

# Load persisted form/master-data configuration. Existing workbook values are
# retained as fallbacks so old documents never disappear from selectors.
_df_signature = (len(df), tuple(sorted(df["discipline"].dropna().astype(str).unique())) if has_data else (),
                 tuple(sorted(df["poc"].dropna().astype(str).unique())) if has_data else ())
_stored_form_config = load_form_config(supabase, _df_signature)
form_config = normalize_form_config(_stored_form_config, df if has_data else None)
REVISION_OPTIONS = form_config["revision_options"]
CODE_TO_DCI_STATUS = form_config["code_status_map"]
RETURNED_CODE_OPTIONS = list(CODE_TO_DCI_STATUS.keys())
CONFIGURED_DISCIPLINES = form_config["discipline_options"]
CONFIGURED_VENDORS = form_config["vendor_options"]

st.sidebar.markdown("---")
st.sidebar.subheader("Filters")

if has_data:
    # Master-data options come from Admin Configuration. Values already used by
    # imported documents are appended so legacy records remain filterable.
    disciplines = _clean_unique_options(CONFIGURED_DISCIPLINES + _values_from_documents(df, "discipline"))
    vendors = _clean_unique_options(CONFIGURED_VENDORS + _values_from_documents(df, "poc"))
    statuses = sorted([s for s in df["status_clean"].dropna().unique()])
    code_buckets = sorted([c for c in df["code_bucket"].dropna().unique()])
    revisions = _clean_unique_options(REVISION_OPTIONS + [r for r in df["latest_rev_display"].unique() if r != ""])
else:
    disciplines = CONFIGURED_DISCIPLINES
    vendors = CONFIGURED_VENDORS
    statuses = code_buckets = []
    revisions = REVISION_OPTIONS
aging_buckets = AGING_BUCKET_ORDER

f_discipline = st.sidebar.multiselect("Discipline", disciplines)
f_vendor = st.sidebar.multiselect("Vendor / Consultant (POC)", vendors)
f_status = st.sidebar.multiselect("Status", statuses)
f_code = st.sidebar.multiselect("Approval code", code_buckets)
f_revision = st.sidebar.multiselect("Revision", revisions)
f_aging = st.sidebar.multiselect("Aging bucket (open docs only)", aging_buckets)
search = st.sidebar.text_input("Search document number / title")

f_date_range, include_undated = None, True
if has_data:
    date_cols_available = df["latest_uploaded_date"].dropna()
    if not date_cols_available.empty:
        min_d, max_d = date_cols_available.min().date(), date_cols_available.max().date()
        f_date_range = st.sidebar.date_input(
            "Uploaded date range", value=(min_d, max_d), min_value=min_d, max_value=max_d
        )
        include_undated = st.sidebar.checkbox("Include documents without an uploaded date", value=True)

filtered = df.copy()
if has_data:
    if f_discipline:
        filtered = filtered[filtered["discipline"].isin(f_discipline)]
    if f_vendor:
        filtered = filtered[filtered["poc"].isin(f_vendor)]
    if f_status:
        filtered = filtered[filtered["status_clean"].isin(f_status)]
    if f_code:
        filtered = filtered[filtered["code_bucket"].isin(f_code)]
    if f_revision:
        filtered = filtered[filtered["latest_rev_display"].isin(f_revision)]
    if f_aging:
        filtered = filtered[filtered["aging_bucket"].isin(f_aging)]
    if search:
        s = search.lower()
        filtered = filtered[
            filtered["doc_no"].str.lower().str.contains(s, na=False)
            | filtered["title"].fillna("").str.lower().str.contains(s, na=False)
        ]
    if f_date_range and isinstance(f_date_range, tuple) and len(f_date_range) == 2:
        start_d, end_d = pd.Timestamp(f_date_range[0]), pd.Timestamp(f_date_range[1])
        date_match = filtered["latest_uploaded_date"].between(start_d, end_d)
        if include_undated:
            date_match = date_match | filtered["latest_uploaded_date"].isna()
        filtered = filtered[date_match]

filtered_doc_numbers = set(filtered["doc_no"]) if has_data else set()
filtered_hist = hist[hist["doc_no"].isin(filtered_doc_numbers)].copy() if has_data else hist

st.sidebar.markdown("---")
with st.sidebar.expander("Quick Statistics", expanded=False):
    if has_data:
        st.metric("Filtered Documents", len(filtered))
        st.metric("Approved", int((filtered["code_bucket"] == "Approved (Code 1)").sum()))
        st.metric("Returned", int((filtered["status_clean"] == "RETURNED BY EIL").sum()))
        st.metric("Delayed", int(filtered["is_overdue"].sum()))
        st.metric("Under Review", int((filtered["status_clean"] == "WITH EIL").sum()))
    else:
        st.caption("No data loaded yet.")

# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
project_name = active_import["project_name"] if active_import else active_project.get("project_name")
contractor = active_import["contractor"] if active_import else active_project.get("contractor")

if has_data:
    as_of_candidates = pd.concat([df["latest_uploaded_date"], df["latest_from_eil_date"]]).dropna()
    as_of_date = as_of_candidates.max().strftime("%d %b %Y") if not as_of_candidates.empty else "N/A"
else:
    as_of_date = "N/A"

st.title(project_name or "Document Control Index Dashboard")
st.caption(f"Contractor: {contractor or 'N/A'}  |  Data as of: {as_of_date}  |  Total documents: {len(df)}")

tabs = st.tabs(
    [
        "Executive Summary",
        "Submission Progress",
        "Review Progress",
        "Approval Status",
        "Discipline Analysis",
        "Vendor Analysis",
        "Delays & Aging",
        "Manhours",
        "Document Register",
        "Update Document",
        "Document Timeline",
        "Import / Export",
    ]
)


# Mouse-friendly horizontal controls for the tab bar. Streamlit does not expose
# tab scrolling controls, so a tiny component augments the existing tablist.
import streamlit.components.v1 as components
components.html(
    """
    <script>
    (function addTabScrollButtons() {
      const doc = window.parent.document;
      const install = () => {
        const tabList = doc.querySelector('[data-testid="stTabs"] [role="tablist"]');
        if (!tabList || doc.getElementById('dci-tab-scroll-right')) return false;

        const host = tabList.parentElement;
        host.style.display = 'flex';
        host.style.alignItems = 'center';
        host.style.gap = '6px';
        tabList.style.flex = '1 1 auto';
        tabList.style.minWidth = '0';
        tabList.style.overflowX = 'auto';
        tabList.style.scrollBehavior = 'smooth';
        tabList.style.scrollbarWidth = 'thin';

        const makeButton = (id, label, direction, title) => {
          const button = doc.createElement('button');
          button.id = id;
          button.type = 'button';
          button.textContent = label;
          button.title = title;
          button.style.cssText = [
            'flex:0 0 auto', 'height:34px', 'min-width:42px', 'padding:0 10px',
            'border:1px solid rgba(49,51,63,.25)', 'border-radius:8px',
            'background:white', 'color:#31333f', 'font-weight:700',
            'cursor:pointer', 'position:sticky', direction < 0 ? 'left:0' : 'right:0',
            'z-index:20'
          ].join(';');
          button.onclick = () => tabList.scrollBy({left: direction * Math.max(320, tabList.clientWidth * 0.7), behavior:'smooth'});
          return button;
        };

        host.insertBefore(makeButton('dci-tab-scroll-left', '<<', -1, 'Scroll tabs left'), tabList);
        host.appendChild(makeButton('dci-tab-scroll-right', '>>', 1, 'Scroll tabs right'));
        return true;
      };

      if (!install()) {
        const observer = new MutationObserver(() => { if (install()) observer.disconnect(); });
        observer.observe(doc.body, {childList:true, subtree:true});
        setTimeout(() => observer.disconnect(), 10000);
      }
    })();
    </script>
    """,
    height=0,
)

def build_document_selector_options(dataframe):
    """Return stable labels keyed by document_id, even when doc numbers repeat."""
    options = []
    duplicate_counts = dataframe["doc_no"].value_counts(dropna=False)
    for _, row in dataframe.sort_values(["doc_no", "original_row_number"], na_position="last").iterrows():
        doc_no = str(row.get("doc_no") or "(no document number)")
        label = doc_no
        if duplicate_counts.get(row.get("doc_no"), 0) > 1:
            label = f"{doc_no} — Excel row {row.get('original_row_number')}"
        options.append((label, row.get("document_id")))
    return options


try:
    # ------------------------------------------------------------------
    # TAB 1: EXECUTIVE SUMMARY
    # ------------------------------------------------------------------
    with tabs[0]:
        if not has_data:
            st.info("No workbook imported yet. Go to **Import / Export** to upload one.")
        else:
            total = len(filtered)
            submitted = int(filtered["is_submitted"].sum())
            reviewed = int(filtered["is_reviewed"].sum())
            approved1 = int((filtered["code_bucket"] == "Approved (Code 1)").sum())
            retained = int((filtered["code_bucket"] == "Retained for Record (Code R)").sum())
            under_review = int((filtered["status_clean"] == "WITH EIL").sum())
            pending_first_sub = int((filtered["status_clean"] == "FRESH SUBMISSION").sum())
            overdue = int(filtered["is_overdue"].sum())

            r1 = st.columns(5)
            r1[0].metric("Total Documents", total)
            r1[1].metric("Submitted", submitted, f"{submitted/total*100:.0f}%" if total else None)
            r1[2].metric("Reviewed (latest rev)", reviewed, f"{reviewed/total*100:.0f}%" if total else None)
            r1[3].metric("Approved (Code 1)", approved1, f"{approved1/total*100:.0f}%" if total else None)
            r1[4].metric("Retained for Record", retained)

            r2 = st.columns(5)
            r2[0].metric("Under Review (with EIL)", under_review)
            r2[1].metric("Pending First Submission", pending_first_sub)
            r2[2].metric("Code 2 (comments)", int((filtered["code_bucket"] == "Code 2 - Comments/Resubmit").sum()))
            r2[3].metric("Void", int(filtered["is_void"].sum()))
            r2[4].metric("Overdue (submission or review)", overdue, delta_color="inverse")

            st.markdown("---")
            c1, c2 = st.columns([2, 1])
            with c1:
                breakdown = filtered["code_bucket"].value_counts().reset_index()
                breakdown.columns = ["Code", "Count"]
                fig = px.pie(breakdown, names="Code", values="Count", title="Approval Status Breakdown", hole=0.4)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                st.subheader("Root Cause - Manhour Overshoot")
                st.caption("Carried over from the previous Excel dashboard (view-only).")
                for i, note in enumerate(ROOT_CAUSE_NOTES, 1):
                    st.markdown(f"**{i}.** {note}")

    # ------------------------------------------------------------------
    # TAB 2: SUBMISSION PROGRESS
    # ------------------------------------------------------------------
    with tabs[1]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Monthly Submission Chart")
            render_monthly_grouped_chart(filtered_hist, "uploaded_date", "Monthly Submission Chart", "Submissions")

            st.markdown("---")
            st.subheader("Planned vs Actual First Submission")

            first_submission = (
                filtered_hist.dropna(subset=["uploaded_date"])
                .groupby("doc_no", as_index=False)["uploaded_date"].min()
                .rename(columns={"uploaded_date": "first_uploaded_date"})
            )

            planned = filtered.dropna(subset=["schedule_submission_date"]).copy()
            planned["planned_month"] = planned["schedule_submission_date"].dt.to_period("M").astype(str)
            planned_counts = planned.groupby("planned_month").size().reset_index(name="Planned")

            actual = filtered.merge(first_submission, on="doc_no", how="left").dropna(subset=["first_uploaded_date"])
            actual["actual_month"] = actual["first_uploaded_date"].dt.to_period("M").astype(str)
            actual_counts = actual.groupby("actual_month").size().reset_index(name="Actual")

            merged = pd.merge(
                planned_counts, actual_counts, left_on="planned_month", right_on="actual_month", how="outer"
            )
            merged["month"] = merged["planned_month"].combine_first(merged["actual_month"])
            merged = merged[["month", "Planned", "Actual"]].fillna(0).sort_values("month")
            fig2 = px.line(
                merged, x="month", y=["Planned", "Actual"], markers=True,
                title="Planned vs Actual First Submission by Month",
            )
            st.plotly_chart(fig2, use_container_width=True)

            cum = merged.copy()
            cum["Planned Cum"] = cum["Planned"].cumsum()
            cum["Actual Cum"] = cum["Actual"].cumsum()
            fig3 = px.line(cum, x="month", y=["Planned Cum", "Actual Cum"], markers=True, title="Cumulative Submission Progress")
            st.plotly_chart(fig3, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 3: REVIEW PROGRESS
    # ------------------------------------------------------------------
    with tabs[2]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Monthly Review Chart")
            render_monthly_grouped_chart(filtered_hist, "from_eil_date", "Monthly Review Chart", "Reviews")

            st.markdown("---")
            st.subheader("Submission-to-Review Turnaround")
            invalid_count = st.session_state.get("_invalid_turnaround_count", 0)
            if invalid_count:
                st.warning(
                    f"{invalid_count} revision record(s) had an EIL return date earlier than their "
                    "upload date and were excluded from turnaround calculations."
                )
            turn = filtered_hist.dropna(subset=["turnaround_days"])
            if not turn.empty:
                avg_turn = turn.groupby("returned_month")["turnaround_days"].mean().reset_index()
                fig2 = px.line(avg_turn, x="returned_month", y="turnaround_days", markers=True,
                                title="Average Turnaround Time (days) by Month")
                st.plotly_chart(fig2, use_container_width=True)
                st.metric("Average turnaround (all revisions, days)", f"{turn['turnaround_days'].mean():.1f}")
            else:
                st.info("No turnaround data available yet.")

    # ------------------------------------------------------------------
    # TAB 4: APPROVAL STATUS
    # ------------------------------------------------------------------
    with tabs[3]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Approval Code Breakdown (current / latest revision)")
            code_counts = filtered["code_bucket"].value_counts().reset_index()
            code_counts.columns = ["Code", "Count"]
            fig = px.bar(code_counts, x="Code", y="Count", title="Documents by Approval Code", text="Count")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Approval Code by Revision (full history)")
            rc = filtered_hist.dropna(subset=["rev_no"]).groupby(["rev_no", "code_bucket"]).size().reset_index(name="Count")
            fig2 = px.bar(rc, x="rev_no", y="Count", color="code_bucket", title="Code outcome at each revision number", barmode="stack")
            st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Revision Distribution")
            rev_dist = (
                filtered.loc[filtered["latest_rev_display"] != "", "latest_rev_display"]
                .value_counts()
                .reset_index()
            )
            rev_dist.columns = ["Latest Revision", "Count"]
            rev_dist = rev_dist.sort_values("Latest Revision", key=lambda s: s.str.casefold())
            fig3 = px.bar(rev_dist, x="Latest Revision", y="Count", title="Documents by Current Revision Number")
            st.plotly_chart(fig3, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 5: DISCIPLINE ANALYSIS
    # ------------------------------------------------------------------
    with tabs[4]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Documents by Discipline")
            disc_counts = filtered["discipline"].value_counts().reset_index()
            disc_counts.columns = ["Discipline", "Count"]
            fig = px.bar(disc_counts.sort_values("Count"), x="Count", y="Discipline", orientation="h",
                         title="Total Documents by Discipline", height=600)
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Pending Documents by Discipline")
            pending = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
            pend_disc = pending["discipline"].value_counts().reset_index()
            pend_disc.columns = ["Discipline", "Pending"]
            fig2 = px.bar(pend_disc.sort_values("Pending"), x="Pending", y="Discipline", orientation="h",
                          title="Pending Documents by Discipline", height=600)
            st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Overdue Documents by Discipline")
            overdue_df = filtered[filtered["is_overdue"] == True]
            od_disc = overdue_df["discipline"].value_counts().reset_index()
            od_disc.columns = ["Discipline", "Overdue"]
            if not od_disc.empty:
                fig3 = px.bar(od_disc.sort_values("Overdue"), x="Overdue", y="Discipline", orientation="h",
                              title="Overdue Documents by Discipline", height=500)
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("No overdue documents for the current filter.")

    # ------------------------------------------------------------------
    # TAB 6: VENDOR ANALYSIS
    # ------------------------------------------------------------------
    with tabs[5]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Documents by Vendor / Consultant")
            vend_counts = filtered["poc"].value_counts().reset_index()
            vend_counts.columns = ["Vendor / POC", "Count"]
            fig = px.bar(vend_counts.sort_values("Count").tail(20), x="Count", y="Vendor / POC", orientation="h",
                         title="Documents by Vendor (Top 20)", height=600)
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Top Vendors with Pending Documents")
            pending = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
            pend_vend = pending["poc"].value_counts().reset_index().head(15)
            pend_vend.columns = ["Vendor / POC", "Pending"]
            fig2 = px.bar(pend_vend.sort_values("Pending"), x="Pending", y="Vendor / POC", orientation="h",
                          title="Top 15 Vendors by Pending Documents", height=500)
            st.plotly_chart(fig2, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 7: DELAYS & AGING
    # ------------------------------------------------------------------
    with tabs[6]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            open_docs = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
            st.subheader("Aging Buckets (open documents)")
            bucket_counts = open_docs["aging_bucket"].value_counts().reindex(AGING_BUCKET_ORDER).fillna(0).reset_index()
            bucket_counts.columns = ["Aging Bucket", "Count"]
            fig = px.bar(bucket_counts, x="Aging Bucket", y="Count", text="Count", title="Open Documents by Aging Bucket")
            st.plotly_chart(fig, use_container_width=True)

            m = st.columns(3)
            m[0].metric("Total open (pending) documents", len(open_docs))
            m[1].metric("Late First Submissions", int(filtered["is_submission_overdue"].fillna(False).sum()))
            m[2].metric(f"EIL Reviews Overdue (> {review_period_days} days)", int(filtered["is_review_overdue"].fillna(False).sum()))
            st.caption(
                "\"Late First Submissions\" compares today against the scheduled submission date for "
                "documents not yet uploaded. \"EIL Reviews Overdue\" compares today against the upload "
                "date for documents still awaiting EIL return."
            )

            st.subheader("Overdue Documents")
            overdue_table = open_docs[open_docs["is_overdue"] == True][
                ["doc_no", "title", "discipline", "poc", "status_clean", "age_days", "schedule_submission_date", "latest_uploaded_date"]
            ].sort_values("age_days", ascending=False).copy()
            for date_col in ["schedule_submission_date", "latest_uploaded_date"]:
                overdue_table[date_col] = overdue_table[date_col].dt.strftime("%d-%b-%Y").fillna("")
            st.dataframe(overdue_table, use_container_width=True, hide_index=True)

    # ------------------------------------------------------------------
    # TAB 8: MANHOURS
    # ------------------------------------------------------------------
    with tabs[7]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            total_expected = filtered["expected_manhours"].fillna(0).sum()
            total_actual = (filtered["eil_manhours"].fillna(0) + filtered["stspl_manhours"].fillna(0)).sum()
            variance = total_actual - total_expected

            c = st.columns(3)
            c[0].metric("Planned Manhours (Expected)", f"{total_expected:,.0f}")
            c[1].metric("Actual Manhours (EIL + STSPL)", f"{total_actual:,.0f}")
            c[2].metric("Variance", f"{variance:,.0f}", delta_color="inverse")

            st.subheader("Manhours by Discipline")
            mh_disc = filtered.groupby("discipline")[["expected_manhours", "eil_manhours", "stspl_manhours"]].sum().reset_index()
            mh_disc["actual_manhours"] = mh_disc["eil_manhours"].fillna(0) + mh_disc["stspl_manhours"].fillna(0)
            fig = px.bar(mh_disc, x="discipline", y=["expected_manhours", "actual_manhours"], barmode="group",
                         title="Planned vs Actual Manhours by Discipline")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Manhours by Vendor")
            mh_vend = filtered.groupby("poc")[["expected_manhours", "eil_manhours", "stspl_manhours"]].sum().reset_index()
            mh_vend["actual_manhours"] = mh_vend["eil_manhours"].fillna(0) + mh_vend["stspl_manhours"].fillna(0)
            mh_vend = mh_vend.sort_values("actual_manhours", ascending=False).head(15)
            fig2 = px.bar(mh_vend, x="poc", y=["expected_manhours", "actual_manhours"], barmode="group",
                          title="Planned vs Actual Manhours - Top 15 Vendors")
            st.plotly_chart(fig2, use_container_width=True)

            st.markdown("---")
            st.subheader("Root Cause Analysis for Manhour Overshoot")
            st.caption("Carried over from the original Excel dashboard (view-only reference list).")
            for i, note in enumerate(ROOT_CAUSE_NOTES, 1):
                st.markdown(f"**{i}.** {note}")

    # ------------------------------------------------------------------
    # TAB 9: DOCUMENT REGISTER
    # ------------------------------------------------------------------
    with tabs[8]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader(f"Document Register ({len(filtered)} documents matching filters)")

            register_cols = [
                "doc_no", "title", "discipline", "poc", "latest_rev_display", "status_clean",
                "code_bucket", "schedule_submission_date", "latest_uploaded_date", "latest_from_eil_date",
                "age_days",
            ]
            register_df = filtered[register_cols].rename(columns={
                "doc_no": "Document Number", "title": "Document Title", "discipline": "Discipline",
                "poc": "Vendor", "latest_rev_display": "Current Revision", "status_clean": "Current Status",
                "code_bucket": "Approval Code", "schedule_submission_date": "Submission Date",
                "latest_uploaded_date": "Last Uploaded", "latest_from_eil_date": "Review Date",
                "age_days": "Aging Days",
            }).copy()
            for c in ["Submission Date", "Last Uploaded", "Review Date"]:
                register_df[c] = pd.to_datetime(register_df[c]).dt.strftime("%d-%b-%Y").fillna("")
            register_df = register_df.sort_values("Document Number")

            st.dataframe(register_df, use_container_width=True, hide_index=True, height=500)

            st.markdown("---")
            st.caption("Select a document below to jump to Update Document or Document Timeline.")
            pick_options = ["-- select --"] + sorted(filtered["doc_no"].dropna().unique().tolist())
            picked = st.selectbox("Open document", pick_options, key="register_pick")
            if picked != "-- select --":
                st.session_state["selected_doc_no"] = picked
                st.success(f"'{picked}' selected. Open the **Update Document** or **Document Timeline** tab.")

    # ------------------------------------------------------------------
    # TAB 10: UPDATE DOCUMENT
    # ------------------------------------------------------------------
    with tabs[9]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Update Document")
            if not CAN_EDIT:
                st.warning("Your Viewer role is read-only. An Editor, Admin, or Master user can save document updates.")
            selector_options = build_document_selector_options(df)
            option_labels = [label for label, _ in selector_options]
            id_by_label = dict(selector_options)
            default_idx = 0
            preselected = st.session_state.get("selected_doc_no")
            if preselected:
                matches = [i for i, label in enumerate(option_labels)
                           if label == preselected or label.startswith(f"{preselected} —")]
                if matches:
                    default_idx = matches[0]

            selected_label = st.selectbox("Document", option_labels, index=default_idx, key="update_doc_select")
            document_id = id_by_label.get(selected_label)

            if document_id is not None:
                doc_row = df[df["document_id"] == document_id].iloc[0]

                st.markdown("#### Current Information")
                info1, info2 = st.columns(2)
                info1.text_input("Document Number", value=str(doc_row["doc_no"] or ""), disabled=True,
                                 key=f"document_number_readonly_{document_id}")
                info2.text_input("Document Title", value=str(doc_row["title"] or ""), disabled=True,
                                 key=f"document_title_readonly_{document_id}")

                current_discipline = str(doc_row["discipline"] or "").strip()
                current_vendor = str(doc_row["poc"] or "").strip()

                info3, info4 = st.columns(2)
                info3.text_input(
                    "Discipline", value=current_discipline, disabled=True,
                    key=f"discipline_readonly_{document_id}",
                )
                info4.text_input(
                    "Vendor / Consultant (POC)", value=current_vendor, disabled=True,
                    key=f"vendor_readonly_{document_id}",
                )

                st.markdown("#### Revision Update")
                st.caption(
                    "Choose the configured values and dates. DCI Status is calculated automatically from Return Code."
                )

                current_revision = "" if pd.isna(doc_row["latest_rev"]) else str(doc_row["latest_rev"]).strip()
                next_revision = _next_sequential_revision(current_revision)
                if next_revision is not None:
                    # For numeric revision schemes, users may correct the current
                    # revision or advance by exactly one step—never skip revisions.
                    revision_choices = _clean_unique_options([current_revision, next_revision])
                else:
                    # Non-numeric schemes remain controlled by the admin master list.
                    revision_choices = _clean_unique_options(REVISION_OPTIONS + [current_revision])
                current_code = normalize_code(doc_row["latest_code"])
                code_choices = _clean_unique_options(RETURNED_CODE_OPTIONS + [current_code])

                def _safe_date(v):
                    ts = pd.to_datetime(v, errors="coerce")
                    return None if pd.isna(ts) else ts.date()

                left, right = st.columns(2)
                with left:
                    new_revision = st.selectbox(
                        "Revision", revision_choices,
                        index=revision_choices.index(current_revision) if current_revision in revision_choices else 0,
                        key=f"revision_{document_id}", disabled=not CAN_EDIT,
                    )
                    new_code = st.selectbox(
                        "Return Code", code_choices,
                        index=code_choices.index(current_code) if current_code in code_choices else 0,
                        key=f"return_code_{document_id}", disabled=not CAN_EDIT,
                    )
                    derived_status = CODE_TO_DCI_STATUS.get(new_code, "UNMAPPED")
                    st.text_input("DCI Status (automatic)", value=derived_status, disabled=True,
                                  key=f"derived_status_{document_id}_{new_code}")
                with right:
                    revision_date = st.date_input(
                        "Revision Date", value=_safe_date(doc_row["latest_uploaded_date"]),
                        key=f"revision_date_{document_id}", disabled=not CAN_EDIT,
                    )
                    return_date = st.date_input(
                        "Return Date", value=_safe_date(doc_row["latest_from_eil_date"]),
                        key=f"return_date_{document_id}", disabled=not CAN_EDIT,
                    )

                validation_errors = []
                if not new_revision:
                    validation_errors.append("A revision must be selected.")
                elif not _revision_transition_is_valid(current_revision, new_revision):
                    expected_revision = _next_sequential_revision(current_revision)
                    if expected_revision is not None:
                        validation_errors.append(
                            f"Revision {current_revision} can only remain {current_revision} for a correction "
                            f"or advance to revision {expected_revision}."
                        )
                if not new_code:
                    validation_errors.append("A return code must be selected.")
                if derived_status == "UNMAPPED":
                    validation_errors.append("This return code has no DCI Status mapping. Ask an administrator to configure it.")
                if return_date and revision_date and return_date < revision_date:
                    validation_errors.append("Return Date cannot be earlier than Revision Date.")
                for message in validation_errors:
                    st.error(message)

                st.markdown("#### Comments")
                existing_remarks = "" if pd.isna(doc_row.get("remarks")) else str(doc_row.get("remarks") or "")
                if existing_remarks:
                    with st.expander("View existing comments", expanded=False):
                        st.text(existing_remarks)
                else:
                    st.caption("No additional comments have been added.")
                new_comment = st.text_area(
                    "Add Comment", placeholder="Enter an optional comment for this revision update…",
                    key=f"new_comment_{document_id}", disabled=not CAN_EDIT,
                )

                if st.button("Save Revision Update", type="primary", disabled=bool(validation_errors) or not CAN_EDIT,
                             key=f"save_revision_{document_id}"):
                    remarks_value = existing_remarks
                    clean_comment = new_comment.strip() if new_comment else ""
                    if clean_comment:
                        stamp = datetime.now().strftime("%d %b %Y %H:%M")
                        remarks_value = (existing_remarks + "\n" + f"[{stamp}] {clean_comment}").strip()

                    form_values = {
                        "status": derived_status,
                        "latest_rev": new_revision,
                        "latest_code": new_code,
                        "latest_uploaded_date": revision_date,
                        "latest_from_eil_date": return_date,
                    }
                    if clean_comment:
                        form_values["remarks"] = remarks_value

                    changed, diff = save_document_changes(
                        supabase, document_id, form_values, doc_row.to_dict(),
                        changed_by=CURRENT_ACTOR, event_type="revision_updated",
                    )
                    if changed:
                        st.success(f"Saved {len(diff)} change(s) to {doc_row['doc_no']}.")
                        st.rerun()
                    else:
                        st.info("No changes were detected.")

                st.markdown("---")
                if CAN_ADMIN:
                    with st.expander("⚙ Admin Configuration", expanded=False):
                        st.caption("Admin/master users can manage revision, code, discipline, and vendor master data.")
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown("##### Revision options")
                            edited_revisions = st.data_editor(
                                pd.DataFrame({"Revision": REVISION_OPTIONS}), num_rows="dynamic",
                                hide_index=True, use_container_width=True, key="revision_options_editor",
                            )
                            st.markdown("##### Discipline options")
                            edited_disciplines = st.data_editor(
                                pd.DataFrame({"Discipline": CONFIGURED_DISCIPLINES}), num_rows="dynamic",
                                hide_index=True, use_container_width=True, key="discipline_options_editor",
                            )
                        with c2:
                            st.markdown("##### Return-code → DCI-status mapping")
                            edited_mapping = st.data_editor(
                                pd.DataFrame([{"Return Code": c, "DCI Status": v}
                                              for c, v in CODE_TO_DCI_STATUS.items()]),
                                num_rows="dynamic", hide_index=True, use_container_width=True,
                                key="code_status_mapping_editor",
                            )
                            st.markdown("##### Vendor / Consultant options")
                            edited_vendors = st.data_editor(
                                pd.DataFrame({"Vendor / Consultant": CONFIGURED_VENDORS}), num_rows="dynamic",
                                hide_index=True, use_container_width=True, key="vendor_options_editor",
                            )

                        revision_values = _clean_unique_options(edited_revisions.get("Revision", []).tolist())
                        discipline_values = _clean_unique_options(edited_disciplines.get("Discipline", []).tolist())
                        vendor_values = _clean_unique_options(edited_vendors.get("Vendor / Consultant", []).tolist())
                        raw_codes, new_mapping, config_errors = [], {}, []
                        for _, row in edited_mapping.iterrows():
                            code = normalize_code(row.get("Return Code"))
                            status = " ".join(str(row.get("DCI Status") or "").strip().upper().split())
                            if not code and not status:
                                continue
                            if not code or not status:
                                config_errors.append("Every mapping row must contain both Return Code and DCI Status.")
                                continue
                            raw_codes.append(code)
                            new_mapping[code] = status

                        if not revision_values: config_errors.append("At least one revision option is required.")
                        if not discipline_values: config_errors.append("At least one discipline option is required.")
                        if not vendor_values: config_errors.append("At least one vendor option is required.")
                        if not new_mapping: config_errors.append("At least one return-code mapping is required.")
                        duplicates = sorted({x for x in raw_codes if raw_codes.count(x) > 1})
                        if duplicates:
                            config_errors.append("Duplicate return codes are not allowed: " + ", ".join(duplicates))
                        for message in dict.fromkeys(config_errors):
                            st.error(message)

                        save_col, reset_col = st.columns(2)
                        if save_col.button("Save Configuration", type="primary",
                                           disabled=bool(config_errors), key="save_master_data_configuration"):
                            try:
                                save_form_config(supabase, {
                                    "revision_options": revision_values,
                                    "code_status_map": new_mapping,
                                    "discipline_options": discipline_values,
                                    "vendor_options": vendor_values,
                                })
                                st.success("Configuration saved. Update-form and sidebar options are now refreshed.")
                                st.rerun()
                            except Exception as exc:
                                st.error("Could not save configuration. Ensure dashboard_settings exists. " + str(exc))

                        if reset_col.button("Reset to Workbook Defaults", key="reset_master_data_configuration"):
                            try:
                                save_form_config(supabase, default_form_config(df))
                                st.success("Configuration reset from the current workbook.")
                                st.rerun()
                            except Exception as exc:
                                st.error("Could not reset configuration: " + str(exc))
    # TAB 11: DOCUMENT TIMELINE
    # ------------------------------------------------------------------
    with tabs[10]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Document Timeline")
            selector_options = build_document_selector_options(df)
            option_labels = [label for label, _ in selector_options]
            id_by_label = dict(selector_options)
            default_idx = 0
            preselected = st.session_state.get("selected_doc_no")
            if preselected:
                matching = [i for i, label in enumerate(option_labels) if label == preselected or label.startswith(f"{preselected} —")]
                if matching:
                    default_idx = matching[0]
            timeline_label = st.selectbox(
                "Document", option_labels, index=default_idx, key="timeline_doc_select"
            )
            document_id = id_by_label.get(timeline_label)

            if document_id is not None:
                doc_row = df[df["document_id"] == document_id].iloc[0]
                timeline_doc_no = doc_row["doc_no"]
                st.caption(f"{doc_row['title'] or ''}  |  Discipline: {doc_row['discipline'] or 'N/A'}  |  Vendor: {doc_row['poc'] or 'N/A'}")

                audit_events = get_document_history(supabase, document_id)
                revision_history = hist[hist["document_id"] == document_id].copy()
                timeline_events = build_document_timeline(revision_history, audit_events)

                if not timeline_events:
                    st.info("No revision dates or dashboard history were found for this document.")
                else:
                    st.caption(
                        "Historical submission and EIL-return events are reconstructed from the "
                        "revision blocks in the imported Excel file. Dashboard edits are then added "
                        "from the Supabase audit log."
                    )

                    for i, ev in enumerate(timeline_events):
                        is_last = i == len(timeline_events) - 1
                        marker = "◉" if is_last else "●"
                        bar = " " if is_last else "│"
                        ts = ev["event_date"]
                        ts_display = "Date unavailable" if pd.isna(ts) else ts.strftime("%d %b %Y")

                        revision_label = f" — Revision {ev['revision']}" if ev.get("revision") else ""
                        st.markdown(f"**{marker} {ev['event']}{revision_label}**")

                        meta = [ts_display]
                        if ev.get("code"):
                            meta.append(f"Code: {ev['code']}")
                        if ev.get("status"):
                            meta.append(f"Status: {ev['status']}")
                        st.caption("  |  ".join(meta))
                        st.write(ev["details"])
                        if not is_last:
                            st.markdown("│")

                    st.markdown("---")
                    st.markdown("#### Revision-wise History Table")
                    timeline_df = pd.DataFrame(timeline_events).rename(columns={
                        "event_date": "Date",
                        "event": "Event",
                        "revision": "Revision",
                        "code": "Returned Code",
                        "status": "Status",
                        "details": "Details",
                        "source": "Source",
                    })
                    timeline_df["Date"] = timeline_df["Date"].apply(
                        lambda value: "" if pd.isna(value) else value.strftime("%d-%b-%Y")
                    )
                    st.dataframe(
                        timeline_df[["Date", "Revision", "Event", "Returned Code", "Status", "Details", "Source"]],
                        use_container_width=True,
                        hide_index=True,
                    )

    # ------------------------------------------------------------------
    # TAB 12: IMPORT / EXPORT
    # ------------------------------------------------------------------
    with tabs[11]:
        if not CAN_ADMIN:
            st.info("Import is restricted to Admin and Master users. Export is available below.")
        else:
            st.subheader("Import")

            pending_bytes = st.session_state.get("pending_upload_bytes")
            pending_name = st.session_state.get("pending_upload_name")

            tab_upload = st.file_uploader("Upload a DCI workbook (.xlsx)", type=["xlsx"], key="import_tab_uploader")
            if tab_upload is not None:
                pending_bytes = tab_upload.getvalue()
                pending_name = tab_upload.name
                st.session_state["pending_upload_bytes"] = pending_bytes
                st.session_state["pending_upload_name"] = pending_name

            if pending_bytes is None:
                st.caption("No pending workbook. Upload a populated workbook above when project documents are ready.")
                if active_project.get("template_id"):
                    try:
                        template_bytes, template_meta = get_template_bytes(supabase, active_project["template_id"])
                        st.download_button(
                            "Download Project Template",
                            data=template_bytes,
                            file_name=template_meta.get("original_filename") or "DCI_Template.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception as template_exc:
                        st.warning(f"The project template could not be loaded: {template_exc}")
            else:
                try:
                    _preview_ws, discovered_layout, discovery_info = inspect_workbook_bytes(pending_bytes)
                    records, history, wb_project, wb_contractor = parse_workbook_bytes(pending_bytes)
                    warnings, errors = validate_records(records)
                    file_hash = workbook_hash(pending_bytes)
                    dup = find_import_by_hash(supabase, file_hash, active_project_id)

                    st.markdown("#### Preview")
                    p1, p2, p3 = st.columns(3)
                    p1.metric("Filename", pending_name)
                    p2.metric("Rows Found", len(records))
                    p3.metric("Warnings", len(warnings))

                    with st.expander("Detected workbook structure", expanded=False):
                        d1, d2, d3 = st.columns(3)
                        d1.metric("Worksheet", discovery_info["sheet"])
                        d2.metric("Header row", discovery_info["header_row"])
                        d3.metric("Revision blocks", len(discovery_info["revision_blocks"]))
                        st.caption(
                            f"First data row: {discovery_info['data_start_row']} · "
                            f"Template engine: v{discovered_layout.get('engine_version', TEMPLATE_ENGINE_VERSION)}"
                        )
                        field_rows = [
                            {"Logical field": field.replace("_", " ").title(), "Detected column": column}
                            for field, column in discovery_info["fields"].items()
                        ]
                        if field_rows:
                            st.dataframe(pd.DataFrame(field_rows), use_container_width=True, hide_index=True)
                        if discovery_info["revision_blocks"]:
                            st.dataframe(
                                pd.DataFrame(discovery_info["revision_blocks"]),
                                use_container_width=True,
                                hide_index=True,
                            )

                    if errors:
                        for e in errors:
                            st.error(e)
                    if warnings:
                        for w in warnings:
                            st.warning(w)

                    preview_df = pd.DataFrame(records[:10])
                    cols_to_show = [c for c in ["doc_no", "title", "discipline", "status", "latest_rev", "latest_code"] if c in preview_df.columns]
                    preview_display = preview_df[cols_to_show] if cols_to_show else preview_df
                    # Raw Excel cells often mix plain numbers (1, 2, 3) with text ("R", "V",
                    # "1 WITH COMMENTS") in the same column (e.g. latest_code, latest_rev).
                    # Arrow can't serialize a mixed int/str object column, so stringify
                    # everything for display purposes only -- the underlying import still
                    # uses the original typed values.
                    preview_display = preview_display.astype(object).where(preview_display.notna(), "").astype(str)
                    st.dataframe(preview_display, use_container_width=True, hide_index=True)

                    if dup:
                        st.info(
                            f"This exact workbook was already imported as **{dup['display_name']}** "
                            f"(v{dup['version_number']}, {dup['uploaded_at'][:10]})."
                        )

                    if not errors:
                        st.markdown("#### Confirm Import")
                        mode_options = ["Import as new workbook"]
                        if imports:
                            mode_options = ["Create New Version", "Replace Current Import", "Import as Separate Workbook"]
                        version_choice = st.radio("Import mode", mode_options, horizontal=True)
                        ic1, ic2 = st.columns([1, 4])
                        if ic1.button("Confirm Import", type="primary"):
                            mode_map = {
                                "Create New Version": "new_version",
                                "Replace Current Import": "replace",
                                "Import as Separate Workbook": "separate",
                                "Import as new workbook": "separate",
                            }
                            mode = mode_map[version_choice]
                            parent_id = active_import_id if mode in ("new_version", "replace") else None
                            try:
                                with st.spinner("Importing..."):
                                    new_import_id, import_warnings = run_import(
                                        supabase, pending_bytes, pending_name, wb_project, wb_contractor,
                                        review_period_days, mode, parent_id,
                                        project_id=active_project_id,
                                        template_id=active_project.get("template_id"),
                                    )
                                st.session_state.pop("pending_upload_bytes", None)
                                st.session_state.pop("pending_upload_name", None)
                                st.cache_data.clear()
                                st.success(f"Import complete: {len(records)} documents imported.")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"Import failed: {ex}")
                        if ic2.button("Cancel Upload"):
                            st.session_state.pop("pending_upload_bytes", None)
                            st.session_state.pop("pending_upload_name", None)
                            st.rerun()
                except Exception as ex:
                    st.error(f"Could not read this workbook: {ex}")

        st.markdown("---")
        st.subheader("Export")

        if not active_import_id:
            st.caption("No active workbook to export.")
        else:
            ec1, ec2 = st.columns(2)
            with ec1:
                if st.button("Generate Updated Workbook (.xlsx)"):
                    with st.spinner("Regenerating workbook from current data..."):
                        try:
                            wb_bytes, wb_filename = generate_updated_workbook(supabase, active_import_id)
                            record_export(supabase, active_import_id, "workbook", wb_filename, len(df))
                            st.download_button(
                                "Download Updated Workbook", data=wb_bytes, file_name=wb_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                        except Exception as ex:
                            st.error(f"Export failed: {ex}")
            with ec2:
                if has_data:
                    all_history_rows = get_history_for_documents(supabase, df["document_id"].dropna().unique().tolist())
                    if all_history_rows:
                        hist_csv = pd.DataFrame(all_history_rows).to_csv(index=False).encode("utf-8")
                        st.download_button(
                            "Download History Report (CSV)", data=hist_csv,
                            file_name=f"dci_history_report_{datetime.now().strftime('%Y-%m-%d')}.csv",
                            mime="text/csv",
                        )
                    else:
                        st.caption("No history events recorded yet.")

            if has_data:
                filtered_csv = filtered.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Download Filtered Document Register (CSV)", data=filtered_csv,
                    file_name=f"dci_filtered_register_{datetime.now().strftime('%Y-%m-%d')}.csv",
                    mime="text/csv",
                )

            st.markdown("#### Previous Exports")
            exp_res = supabase.table("export_records").select("*").eq("import_id", active_import_id) \
                .order("exported_at", desc=True).limit(20).execute()
            if exp_res.data:
                st.dataframe(pd.DataFrame(exp_res.data), use_container_width=True, hide_index=True)
            else:
                st.caption("No exports recorded yet for this workbook.")
except Exception as e:
    st.error(
        "Something went wrong while rendering the dashboard tabs. "
        "The details below should help pin down the cause."
    )
    st.exception(e)
