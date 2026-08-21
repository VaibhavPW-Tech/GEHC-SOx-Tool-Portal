import streamlit as st
import re
import time
import os
import glob
import zipfile
import tempfile
import datetime
import shutil
import subprocess
import io

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ============================================================
# PAGE CONFIG (must be the very first Streamlit call)
# ============================================================
st.set_page_config(
    page_title="GEHC SOX Tools",
    layout="wide",
    page_icon="🚀",
    initial_sidebar_state="expanded",
)

# ============================================================
# GLOBAL CUSTOM STYLING
# ============================================================
st.markdown(
    """
    <style>
    /* ---------- General page polish ---------- */
    .stApp {
        background: linear-gradient(180deg, #0b1020 0%, #10162b 100%);
    }
    #MainMenu, footer {visibility: hidden;}

    /* ---------- Hero banner ---------- */
    .portal-hero {
        background: linear-gradient(135deg, #4338ca 0%, #6366f1 45%, #06b6d4 100%);
        padding: 2.2rem 2.4rem;
        border-radius: 1rem;
        margin-bottom: 1.6rem;
        box-shadow: 0 10px 30px rgba(67, 56, 202, 0.35);
    }
    .portal-hero h1 {
        color: #ffffff;
        font-size: 2.1rem;
        margin: 0;
        font-weight: 800;
    }
    .portal-hero p {
        color: #e0e7ff;
        margin-top: 0.4rem;
        font-size: 1.02rem;
    }

    /* ---------- Page banner (per-tool) ---------- */
    .page-banner {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08);
        padding: 1.1rem 1.5rem;
        border-radius: 0.8rem;
        margin-bottom: 1.4rem;
    }
    .page-banner h2 {
        margin: 0;
        color: #f8fafc;
        font-size: 1.5rem;
    }
    .page-banner p {
        margin: 0.2rem 0 0 0;
        color: #94a3b8;
        font-size: 0.95rem;
    }

    /* ---------- Tool cards on home page ---------- */
    .tool-card {
        background: rgba(255,255,255,0.045);
        border: 1px solid rgba(255,255,255,0.09);
        border-radius: 1rem;
        padding: 1.4rem 1.3rem;
        height: 100%;
        transition: transform 0.15s ease, border-color 0.15s ease;
    }
    .tool-card:hover {
        border-color: #6366f1;
        transform: translateY(-3px);
    }
    .tool-card .icon {
        font-size: 2.2rem;
    }
    .tool-card h3 {
        color: #f1f5f9;
        margin: 0.5rem 0 0.3rem 0;
        font-size: 1.15rem;
    }
    .tool-card p {
        color: #94a3b8;
        font-size: 0.88rem;
        min-height: 3.4rem;
    }

    /* ---------- Sidebar polish ---------- */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f1428 0%, #141a33 100%);
        border-right: 1px solid rgba(255,255,255,0.06);
    }
    .avatar-circle {
        width: 46px; height: 46px;
        border-radius: 50%;
        background: linear-gradient(135deg, #6366f1, #06b6d4);
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 1.05rem;
        margin-right: 0.7rem;
    }
    .sidebar-user-box {
        display: flex;
        align-items: center;
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 0.7rem;
        padding: 0.6rem 0.7rem;
        margin-bottom: 0.9rem;
    }
    .sidebar-user-box .info .name {
        color: #f1f5f9;
        font-weight: 700;
        font-size: 0.95rem;
        line-height: 1.1;
    }
    .sidebar-user-box .info .sso {
        color: #94a3b8;
        font-size: 0.75rem;
    }

    /* ---------- Buttons ---------- */
    .stButton>button {
        border-radius: 0.55rem;
        font-weight: 600;
        transition: all 0.15s ease;
    }
    .stButton>button:hover {
        transform: translateY(-1px);
    }

    /* ---------- Login card ---------- */
    .login-card {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.09);
        border-radius: 1.1rem;
        padding: 2.2rem 2.2rem 1.6rem 2.2rem;
        box-shadow: 0 12px 35px rgba(0,0,0,0.35);
    }
    .login-logo {
        font-size: 2.6rem;
        text-align: center;
        margin-bottom: 0.2rem;
    }
    .login-title {
        text-align: center;
        color: #f8fafc;
        font-weight: 800;
        font-size: 1.5rem;
        margin-bottom: 0.15rem;
    }
    .login-subtitle {
        text-align: center;
        color: #94a3b8;
        font-size: 0.92rem;
        margin-bottom: 1.4rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# CONFIG: Hardcoded valid SSO -> Name mapping
# ============================================================
VALID_SSO_DICT = {

  "550005636": "Harshita Chawla",
  "550005637": "Sidak Kaur",
  "550025844": "Jahanvi Vaid",
  "550025846": "Ashna Tandon",
  "550020293": "Vaibhav Goyal",
  "550020292": "Jeevika Tuli",
  "550015783": "Russh Ahluwalia",
  "550025845": "Anagh Kaura",
  "550005634": "Amanpreet Singh",
  "550005836": "Parikshit Ghosh",
  "550005635": "Anmol Sharma",
  "223118537": "Akshay Aggarwal",
  "250022036": "Shriya Gupta",
  "550010987": "Arjita Sharma"

    # Add more valid SSOs here as "sso.id": "Full Name"
}

# ============================================================
# TOOL METADATA (icons, descriptions) — used on Home + Sidebar
# ============================================================
TOOL_INFO = {
    "Saviynt Tool": {
        "icon": "✅",
        "tagline": "SOX Access Comparison Tool",
        "description": "Upload Saviynt & dump files to run automated SOX access checks and comparisons.",
    },
    "CM Tool 1": {
        "icon": "🔄",
        "tagline": "Sequential Ticket Automation (Pulse)",
        "description": "Automate login and bulk PDF download of Pulse change tickets.",
    },
    "CM Tool 2": {
        "icon": "📝",
        "tagline": "Change Management Report Generator",
        "description": "Upload change ticket PDFs and generate a formatted Excel control-testing report.",
    },
}

# ============================================================
# SESSION STATE INIT (login/navigation only)
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_sso" not in st.session_state:
    st.session_state.user_sso = None
if "user_name" not in st.session_state:
    st.session_state.user_name = None
if "selected_tool" not in st.session_state:
    st.session_state.selected_tool = "Home"
if "login_attempts" not in st.session_state:
    st.session_state.login_attempts = 0
if "login_time" not in st.session_state:
    st.session_state.login_time = None


def _get_initials(name: str) -> str:
    parts = [p for p in name.strip().split(" ") if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][0].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def _greeting() -> str:
    hour = datetime.datetime.now().hour
    if hour < 12:
        return "Good morning"
    elif hour < 17:
        return "Good afternoon"
    else:
        return "Good evening"


# ============================================================
# LOGIN PAGE
# ============================================================
def login():
    left, mid, right = st.columns([1, 1.3, 1])
    with mid:
        #st.markdown('<div class="login-card">', unsafe_allow_html=True)
        st.markdown('<div class="login-logo">🚀</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-title">GEHC SOx Tools</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="login-subtitle">Sign in with your SSO ID to access Tools </div>',
            unsafe_allow_html=True,
        )

        sso = st.text_input("SSO ID", key="login_sso_input", placeholder="e.g. john.doe")

        col_a, col_b = st.columns([1, 1])
        with col_a:
            login_clicked = st.button("🔓 Login", use_container_width=True, type="primary")
        with col_b:
            clear_clicked = st.button("Clear", use_container_width=True)

        if clear_clicked:
            st.session_state.login_sso_input = ""
            st.rerun()

        if login_clicked:
            entered_sso = sso.strip()
            if entered_sso and entered_sso in VALID_SSO_DICT:
                st.session_state.authenticated = True
                st.session_state.user_sso = entered_sso
                st.session_state.user_name = VALID_SSO_DICT[entered_sso]
                st.session_state.login_attempts = 0
                st.session_state.login_time = datetime.datetime.now()
                st.session_state.selected_tool = "Home"
                st.success(f"✅ Welcome, {st.session_state.user_name}! Redirecting...")
                time.sleep(0.6)
                st.rerun()
            else:
                st.session_state.login_attempts += 1
                st.error(
                    "❌ SSO not recognized. Please re-enter correct details. "
                    f"(Attempt {st.session_state.login_attempts})"
                )

        with st.expander("ℹ️ Need help logging in?"):
            st.write(
                "- Enter your organizational SSO ID exactly as issued (case-sensitive).\n"
                "- If your SSO isn't recognized, contact your portal administrator "
                "to be added to the access list.\n"
                "- This portal grants access to: **Saviynt Tool**, **CM Tool 1**, and **CM Tool 2**."
            )

        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown(
            "<p style='text-align:center; color:#475569; font-size:0.78rem; margin-top:0.8rem;'>"
            "🔒 Secure internal access only</p>",
            unsafe_allow_html=True,
        )


# ============================================================
# LOGOUT
# ============================================================
def logout():
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.user_sso = None
        st.session_state.user_name = None
        st.session_state.selected_tool = "Home"
        st.rerun()


# ============================================================
# PAGE BANNER helper (per-tool header)
# ============================================================
def render_page_banner(title: str, subtitle: str, icon: str):
    st.markdown(
        f"""
        <div class="page-banner">
            <h2>{icon} {title}</h2>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# ============================================================
#   TOOL 1: SAVIYNT TOOL  (SOX Access Comparison Tool)
# ============================================================
# ============================================================
def run_saviynt_tool():

    # =========================================
    #  Page config & basic styling
    # =========================================

    st.markdown(
        """
        <style>
        .main {
            background-color: #0b1120;
            color: #e5e7eb;
        }
        .report-box {
            background-color: #020617;
            padding: 1.2rem 1.5rem;
            border-radius: 0.5rem;
            border: 1px solid #1f2937;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.6);
        }
        .stButton>button {
            border-radius: 0.4rem;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.title("SOX Access Comparison Tool")
    st.caption("Upload the Analytics Summary and Dump files, run checks, and see per‑check summaries.")

    # =========================================
    #  Helper functions
    # =========================================
    def make_unique(cols):
        """Make column names unique (for Streamlit/pyarrow display)."""
        seen = {}
        new_cols = []
        for c in cols:
            if c not in seen:
                seen[c] = 0
                new_cols.append(c)
            else:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
        return new_cols


    def extract_id_from_username(username: str) -> str:
        """
        Extract ID from a string like 'Name (12345)' => '12345'.
        Returns upper-stripped string or '' if not found.
        """
        if pd.isna(username):
            return ""
        text = str(username)
        match = re.search(r"\(([^)]+)\)", text)
        if match:
            return match.group(1).strip().upper()
        return ""


    def parse_multi_values(value, treat_as_name_id=False):
        """
        Split a cell into a set of normalized approver/owner values.
        Handles comma-separated values like 'a, b, c'
        or 'Name (ID), Name2 (ID2)' — even if names themselves contain commas
        like 'Smith, John (12345), Doe, Jane (67890)'.

        If treat_as_name_id=True:
            - Splits on commas ONLY when they are NOT inside parentheses
              (so 'Smith, John (12345)' stays as one chunk before ID extraction).
            - Extracts the ID from inside brackets for each chunk.

        If treat_as_name_id=False:
            - Splits on every comma (simple case, plain IDs/names).
        """
        if pd.isna(value):
            return set()
        text = str(value)

        if treat_as_name_id:
            # Split on comma only if NOT followed eventually by ')' before another '('
            parts = re.split(r',\s*(?![^()]*\))', text)
        else:
            parts = text.split(",")

        parts = [p.strip() for p in parts if p.strip()]

        result = set()
        for p in parts:
            if treat_as_name_id:
                id_ = extract_id_from_username(p)
                result.add(id_ if id_ else p.strip().upper())
            else:
                result.add(p.strip().upper())
        return result


    def build_role_approver_dict(df, role_col, approver_cols, treat_as_name_id=False):
        """
        Build {role: set(approvers)} from a dataframe (used as a reference/lookup table).
        approver_cols: list of column names (values may be comma-separated).
        Duplicate roles across rows are merged (union of approvers).
        """
        result = {}
        for _, row in df.iterrows():
            role_val = row[role_col]
            if pd.isna(role_val) or str(role_val).strip() == "":
                continue
            role = str(role_val).strip().upper()

            approvers = set()
            for col in approver_cols:
                approvers |= parse_multi_values(row[col], treat_as_name_id)

            if role in result:
                result[role] |= approvers
            else:
                result[role] = approvers
        return result


    def format_approval_message(role_display, common_approvers):
        """
        Build the human-readable message:
        - 0 matches  -> 'has no matching approvers'
        - 1 match    -> 'is approved by X'
        - 2 matches  -> 'is approved by both X and Y'
        - 3+ matches -> 'is approved by X, Y, and Z'
        """
        approvers = sorted(common_approvers)
        if len(approvers) == 0:
            return f"{role_display} has no matching approvers"
        elif len(approvers) == 1:
            return f"{role_display} is approved by {approvers[0]}"
        elif len(approvers) == 2:
            return f"{role_display} is approved by both {approvers[0]} and {approvers[1]}"
        else:
            return f"{role_display} is approved by {', '.join(approvers[:-1])}, and {approvers[-1]}"


    def style_dataframe_safe(df, status_col=None, valid_cols=None,
                              status_style_fn=None, valid_style_fn=None):
        """
        Apply cell-level styling to a dataframe using pandas Styler,
        compatible with both old pandas (.applymap) and new pandas (.map).
        Requires df to have unique column names.
        """
        if df.columns.duplicated().any():
            raise ValueError("Cannot style a dataframe with duplicate column names.")

        styled = df.style

        use_map = hasattr(styled, "map")

        if status_col and status_col in df.columns and status_style_fn:
            if use_map:
                styled = styled.map(status_style_fn, subset=[status_col])
            else:
                styled = styled.applymap(status_style_fn, subset=[status_col])

        if valid_cols and valid_style_fn:
            valid_cols_present = [c for c in valid_cols if c in df.columns]
            if valid_cols_present:
                if use_map:
                    styled = styled.map(valid_style_fn, subset=valid_cols_present)
                else:
                    styled = styled.applymap(valid_style_fn, subset=valid_cols_present)

        return styled


    # =========================================
    #  1. File upload (all files used across checks)
    # =========================================
    st.header("Upload Files")

    col_up1, col_up2, col_up3 = st.columns(3)

    with col_up1:
        analytics_file = st.file_uploader(
            "Upload Saviynt Summary file (e.g. Analytics_Summary_*.xlsx)",
            type=["xlsx"],
            key="analytics"
        )

    with col_up2:
        sbl_file = st.file_uploader(
            "Upload Dump Automated file (e.g. sbl_am_automated.xlsx)",
            type=["xlsx"],
            key="sbl"
        )

    with col_up3:
        entitlement_file = st.file_uploader(
            "Upload Entitlement Owner List (optional, for Check 4)",
            type=["xlsx", "csv"],
            key="entitlement"
        )

    df_analytics, df_sbl, df_entitlement = None, None, None

    if analytics_file is not None:
        df_analytics = pd.read_excel(analytics_file)

    if sbl_file is not None:
        df_sbl = pd.read_excel(sbl_file)

    if entitlement_file is not None:
        if entitlement_file.name.endswith(".csv"):
            df_entitlement = pd.read_csv(entitlement_file)
        else:
            df_entitlement = pd.read_excel(entitlement_file)

    # Extra safety: stop execution until both mandatory files are uploaded
    if df_analytics is None or df_sbl is None:
        st.info("Please upload both files (Saviynt Summary and Dump Automated) to continue.")
        st.stop()

    # Row counts summary (now guaranteed not None)
    analytics_rows = len(df_analytics)
    sbl_rows = len(df_sbl)
    entitlement_msg = (
        f" Entitlement Owner file has {len(df_entitlement)} rows."
        if df_entitlement is not None
        else " Entitlement Owner file not uploaded (Check 4 will be skipped)."
    )
    st.success(
        f"Files uploaded successfully. "
        f"Saviynt file has {analytics_rows} rows. "
        f"Dump Automated file has {sbl_rows} rows."
        f"{entitlement_msg}"
    )

    all_results = []  # collect partial results from all checks

    # =========================================
    #  2. Check 1 – Date comparisons (Mode A & Mode B)
    # =========================================
    st.header("Check 1 – Access is provisioned post approvals only")

    check1_mode = st.radio(
        "Choose Check 1 mode:",
        options=[
            "A: Approval vs Task Completion (Saviynt only)",
            "B: Application Created Date vs Saviynt Request Approved Date (Unique by SSO)"
        ],
        index=0,
        key="check1_mode"
    )

    # ----------------- MODE A: Existing logic -----------------
    if check1_mode.startswith("A"):
        with st.expander("Configure Check 1 (Mode A)", expanded=True):
            c1_1, c1_2, c1_3 = st.columns(3)

            with c1_1:
                approval_col = st.selectbox(
                    "Approval Date column (Analytics)",
                    options=df_analytics.columns,
                    key="approval_date_col"
                )

            with c1_2:
                completion_col = st.selectbox(
                    "Task Completion Date column (Analytics)",
                    options=df_analytics.columns,
                    key="completion_date_col"
                )

            with c1_3:
                key_col_1 = st.selectbox(
                    "Key/ID column (Analytics, for report)",
                    options=df_analytics.columns,
                    key="key_col_1"
                )

            run_check1 = st.checkbox("Run Check 1 (Mode A, ALL rows)", value=True, key="run_check1A")

        if run_check1:
            df_check1 = df_analytics.copy()

            df_check1["_approval_datetime"] = pd.to_datetime(df_check1[approval_col], errors="coerce")
            df_check1["_completion_datetime"] = pd.to_datetime(df_check1[completion_col], errors="coerce")

            def check_approval_before_completion(row):
                a = row["_approval_datetime"]
                c = row["_completion_datetime"]
                if pd.isna(a) or pd.isna(c):
                    return "Failed - Missing date"
                return "Passed" if a <= c else "Failed - Approval after completion"

            df_check1["Check1_Result"] = df_check1.apply(check_approval_before_completion, axis=1)
            df_check1["Check1_Reason"] = df_check1["Check1_Result"]

            all_results.append(df_check1[[key_col_1, "Check1_Result", "Check1_Reason"]])

            total_rows_c1 = len(df_check1)
            failed_rows_c1 = df_check1[df_check1["Check1_Result"].str.startswith("Failed")]
            passed_rows_c1 = total_rows_c1 - len(failed_rows_c1)

            st.subheader("Check 1 (Mode A) Summary")
            st.write(f"Total rows evaluated: {total_rows_c1}")
            st.write(f"Passed: {passed_rows_c1}")
            st.write(f"Failed: {len(failed_rows_c1)}")

            if len(failed_rows_c1) > 0:
                st.write("Failed rows for Check 1 (Mode A):")
                failed_view = failed_rows_c1[[key_col_1, approval_col, completion_col, "Check1_Result"]].copy()
                failed_view.columns = make_unique(failed_view.columns)
                st.dataframe(failed_view)
            else:
                st.success("There were no issues or failed cases for this check (Mode A).")

            st.write("Full Check 1 (Mode A) results (ALL rows):")
            preview_cols = [key_col_1, approval_col, completion_col, "Check1_Result"]
            preview = df_check1[preview_cols].copy()
            preview.columns = make_unique(preview.columns)
            st.dataframe(preview)

    # ----------------- MODE B: SBL Created Date vs Analytics Request Approved Date -----------------
    else:
        with st.expander("Configure Check 1 (Mode B)", expanded=True):
            c1b_1, c1b_2, c1b_3, c1b_4 = st.columns(4)

            with c1b_1:
                sbl_sso_col = st.selectbox(
                    "User SSO / ID column in Dump (Application file)",
                    options=df_sbl.columns,
                    key="sbl_sso_col_c1"
                )

            with c1b_2:
                sbl_created_date_col = st.selectbox(
                    "Created Date column in Dump (Application file)",
                    options=df_sbl.columns,
                    key="sbl_created_date_col_c1"
                )

            with c1b_3:
                analytics_sso_col = st.selectbox(
                    "User SSO / ID column in Saviynt file",
                    options=df_analytics.columns,
                    key="analytics_sso_col_c1"
                )

            with c1b_4:
                analytics_approved_date_col = st.selectbox(
                    "Request Approved Date column in Saviynt file",
                    options=df_analytics.columns,
                    key="analytics_approved_date_col_c1"
                )

            key_col_1b = st.selectbox(
                "Key/ID column (Saviynt, for report)",
                options=df_analytics.columns,
                key="key_col_1b"
            )

            username_is_pure_id_c1 = st.checkbox(
                "Saviynt SSO column already contains pure ID (no brackets)",
                value=True,
                key="username_is_pure_id_c1",
                help="Tick if the selected Saviynt user column is just the ID (e.g., 223055402), not 'Name (223055402)'."
            )

            run_check1B = st.checkbox(
                "Run Check 1 (Mode B, ALL rows)",
                value=True,
                key="run_check1B"
            )

        if run_check1B:
            # --- Normalize SSO/ID in SBL (Dump) ---
            sbl_tmp = df_sbl[[sbl_sso_col, sbl_created_date_col]].copy()

            # Ensure SSO column is a Series
            sso_series = sbl_tmp[sbl_sso_col]
            if isinstance(sso_series, pd.DataFrame):
                sso_series = sso_series.iloc[:, 0]

            sbl_tmp["_sso_norm"] = sso_series.astype(str).str.strip().str.upper()

            # Ensure created-date column is a Series
            created_series = sbl_tmp[sbl_created_date_col]
            if isinstance(created_series, pd.DataFrame):
                created_series = created_series.iloc[:, 0]

            # Convert created date to date (no time)
            sbl_tmp["_created_dt"] = pd.to_datetime(created_series, errors="coerce").dt.date

            # Map: SSO -> earliest created date (if multiple rows in SBL)
            sbl_created_map = (
                sbl_tmp.groupby("_sso_norm")["_created_dt"]
                .min()
                .to_dict()
            )

            # --- Normalize SSO/ID in Analytics ---
            df_check1B = df_analytics[[analytics_sso_col, analytics_approved_date_col, key_col_1b]].copy()

            analytics_user_series = df_check1B[analytics_sso_col]
            if isinstance(analytics_user_series, pd.DataFrame):
                analytics_user_series = analytics_user_series.iloc[:, 0]

            if username_is_pure_id_c1:
                df_check1B["_sso_norm"] = analytics_user_series.astype(str).str.strip().str.upper()
            else:
                df_check1B["_sso_norm"] = analytics_user_series.apply(extract_id_from_username)

            # Ensure approved-date column is a Series
            approved_series = df_check1B[analytics_approved_date_col]
            if isinstance(approved_series, pd.DataFrame):
                approved_series = approved_series.iloc[:, 0]

            # Convert approved date to date (no time)
            df_check1B["_approved_dt"] = pd.to_datetime(approved_series, errors="coerce").dt.date

            # --- Row-wise comparison: Approved Date vs Created Date ---
            def compare_approved_vs_created(row):
                sso = row["_sso_norm"]
                approved_dt = row["_approved_dt"]

                # If we can't extract any SSO at all from Analytics -> fail
                if not sso:
                    return "Failed - No SSO/ID extracted", None, ""

                created_dt = sbl_created_map.get(sso, None)

                # If SSO not found in SBL -> fail
                if created_dt is None:
                    return "Failed - SSO/ID not found in Application File", None, sso

                # If created date exists in mapping but is NaN -> fail
                if pd.isna(created_dt):
                    return "Failed - Created date missing in Application File", None, sso

                # From this point, we HAVE a created date in SBL
                if pd.isna(approved_dt):
                    return "Failed - Approved date missing in Saviynt File", created_dt, sso

                # Rule: Approved date must be <= Created date
                if approved_dt <= created_dt:
                    return "Passed", created_dt, sso
                else:
                    return "Failed - Approved date is after created date", created_dt, sso

            results_b = df_check1B.apply(
                lambda r: compare_approved_vs_created(r),
                axis=1,
                result_type="expand"
            )

            df_check1B["Check1_Result"] = results_b[0]
            df_check1B["Created Date"] = results_b[1]
            df_check1B["SSO ID"] = results_b[2]
            df_check1B["Approved Date"] = df_check1B["_approved_dt"]
            df_check1B["Check1_Reason"] = df_check1B["Check1_Result"]

            # --- Collect for final download / combined results ---
            all_results.append(
                df_check1B[
                    [
                        key_col_1b,
                        "SSO ID",
                        "Created Date",
                        "Approved Date",
                        "Check1_Result",
                        "Check1_Reason",
                    ]
                ]
            )

            total_rows_c1b = len(df_check1B)
            failed_rows_c1b = df_check1B[df_check1B["Check1_Result"].str.startswith("Failed")]
            passed_rows_c1b = total_rows_c1b - len(failed_rows_c1b)

            st.subheader("Check 1 (Mode B) Summary")
            st.write(f"Total rows evaluated (Analytics): {total_rows_c1b}")
            st.write(f"Passed: {passed_rows_c1b}")
            st.write(f"Failed: {len(failed_rows_c1b)}")

            if len(failed_rows_c1b) > 0:
                st.write("Failed rows for Check 1 (Mode B):")
                failed_view_b = failed_rows_c1b[
                    [
                        key_col_1b,
                        analytics_sso_col,
                        "SSO ID",
                        "Created Date",
                        "Approved Date",
                        "Check1_Result",
                    ]
                ].copy()
                failed_view_b.columns = make_unique(failed_view_b.columns)
                st.dataframe(failed_view_b)
            else:
                st.success("There were no issues or failed cases for this check (Mode B).")

            st.write("Full Check 1 (Mode B) results (ALL rows):")
            preview_cols_b = [
                key_col_1b,
                analytics_sso_col,
                "SSO ID",
                "Created Date",
                "Approved Date",
                "Check1_Result",
            ]
            preview_b = df_check1B[preview_cols_b].copy()
            preview_b.columns = make_unique(preview_b.columns)
            st.dataframe(preview_b)

    # =========================================
    #  3. Check 2 – SOD: Inter-check between Requested / Approvals / Granted
    # =========================================
    st.header("Check 2 – SOD: Requested vs Approvals vs Granted")

    with st.expander("Configure Check 2", expanded=True):
        st.markdown("**Approval levels (up to 5, Leave unused ones blank ).**")
        c2_cols = st.columns(5)
        approval_cols = []

        for i, col in enumerate(c2_cols, start=1):
            with col:
                selected = st.selectbox(
                    f"Approval Level {i} column (optional)",
                    options=[""] + list(df_analytics.columns),
                    key=f"appr_level_{i}"
                )
                approval_cols.append(selected if selected != "" else None)

        requested_by_col = st.selectbox(
            "Requested By column (Analytics)",
            options=df_analytics.columns,
            key="requested_by_col"
        )

        # NEW: format of Requested For
        requested_is_pure_id = st.checkbox(
            "Requested By already contains pure ID (no Name(ID) format)",
            value=True,
            key="requested_is_pure_id_c2",
            help="Tick if Requested By is just the ID (e.g. 223055402), not 'Name (223055402)'."
        )

        st.markdown("**Granted By source (choose one option):**")
        c2_g1, c2_g2 = st.columns(2)
        with c2_g1:
            granted_by_col = st.selectbox(
                "Granted By column (optional)",
                options=[""] + list(df_analytics.columns),
                key="granted_by_col"
            )
        with c2_g2:
            granted_by_id_input = st.text_input(
                "OR enter a single Granted By ID (applies to all rows)",
                value="",
                key="granted_by_id_input"
            )

        # Format of approval & granted columns
        values_are_name_id = st.checkbox(
            "Approval/Granted columns use 'Name (ID)' format (e.g. John Smith (12345))",
            value=True,
            key="values_are_name_id_c2",
            help="If ticked, the tool will extract the ID inside brackets and compare only IDs."
        )

        key_col_2 = st.selectbox(
            "Key/ID column (Analytics, for report)",
            options=df_analytics.columns,
            key="key_col_2"
        )

        run_check2 = st.checkbox("Run Check 2 (for ALL rows)", value=True)

    if run_check2:
        df_check2 = df_analytics.copy()

        # ---------- Normalization helper (to ID) ----------
        def normalize_id(val, treat_as_name_id: bool):
            """
            Convert a cell value to a comparable ID.
            - If NaN/blank -> "" (ignored in comparisons).
            - If treat_as_name_id=True -> extract 'ID' from 'Name (ID)'.
            - Else -> just strip & uppercase.
            """
            if pd.isna(val):
                return ""
            text = str(val).strip()
            if text == "":
                return ""
            if treat_as_name_id:
                return extract_id_from_username(text)
            else:
                return text.upper()

        # ---------- Requested By (ID) ----------
        rb_series = df_check2[requested_by_col]
        if isinstance(rb_series, pd.DataFrame):
            rb_series = rb_series.iloc[:, 0]
        df_check2["_requested_id"] = rb_series.apply(
            lambda v: normalize_id(v, treat_as_name_id=not requested_is_pure_id)
        )

        # ---------- Granted By source ----------
        granted_by_col_selected = granted_by_col if granted_by_col != "" else None

        granted_by_id_manual = granted_by_id_input.strip()
        granted_by_id_manual_norm = (
            normalize_id(granted_by_id_manual, treat_as_name_id=False)
            if granted_by_id_manual
            else ""
        )

        if granted_by_col_selected:
            gb_series = df_check2[granted_by_col_selected]
            if isinstance(gb_series, pd.DataFrame):
                gb_series = gb_series.iloc[:, 0]
            df_check2["_granted_id_col"] = gb_series.apply(
                lambda v: normalize_id(v, treat_as_name_id=values_are_name_id)
            )
        else:
            df_check2["_granted_id_col"] = ""

        df_check2["_granted_id"] = df_check2.apply(
            lambda r: granted_by_id_manual_norm if granted_by_id_manual_norm else r["_granted_id_col"],
            axis=1,
        )

        # ---------- Approvals (ID) ----------
        norm_approval_cols = []
        for idx, col_name in enumerate(approval_cols, start=1):
            if col_name is None:
                continue
            series = df_check2[col_name]
            if isinstance(series, pd.DataFrame):
                series = series.iloc[:, 0]
            norm_name = f"_appr_{idx}_id"
            df_check2[norm_name] = series.apply(
                lambda v: normalize_id(v, treat_as_name_id=values_are_name_id)
            )
            norm_approval_cols.append((col_name, norm_name))

        # ---------- Row-wise SoD check (by ID) ----------
        def sod_intercheck(row):
            req_id = row["_requested_id"]
            grn_id = row["_granted_id"]

            conflicts = []

            if req_id and grn_id and req_id == grn_id:
                conflicts.append("Requested ID = Granted ID")

            for orig_col, norm_col in norm_approval_cols:
                appr_id = row[norm_col]
                if not appr_id:
                    continue

                if req_id and appr_id == req_id:
                    conflicts.append(f"Approval ({orig_col}) ID = Requested ID")
                if grn_id and appr_id == grn_id:
                    conflicts.append(f"Approval ({orig_col}) ID = Granted ID")

            if not conflicts:
                return "Passed", ""

            reason = "Failed - SoD conflict: " + "; ".join(conflicts)
            return reason, "; ".join(conflicts)

        results = df_check2.apply(sod_intercheck, axis=1, result_type="expand")
        df_check2["Check2_Result"] = results[0]
        df_check2["Check2_Conflicts"] = results[1]
        df_check2["Check2_Reason"] = df_check2["Check2_Result"]

        cols_for_output = [key_col_2, requested_by_col]
        if granted_by_col_selected:
            cols_for_output.append(granted_by_col_selected)
        for orig_col, _norm_col in norm_approval_cols:
            cols_for_output.append(orig_col)

        cols_for_output.append("_requested_id")
        cols_for_output.append("_granted_id")
        for _, norm_col in norm_approval_cols:
            cols_for_output.append(norm_col)

        cols_for_output += ["Check2_Result", "Check2_Reason", "Check2_Conflicts"]

        all_results.append(df_check2[cols_for_output])

        total_rows_c2 = len(df_check2)
        failed_rows_c2 = df_check2[df_check2["Check2_Result"].str.startswith("Failed")]
        passed_rows_c2 = total_rows_c2 - len(failed_rows_c2)

        st.subheader("Check 2 Summary")
        st.write(f"Total rows evaluated: {total_rows_c2}")
        st.write(f"Passed: {passed_rows_c2}")
        st.write(f"Failed: {len(failed_rows_c2)}")

        if len(failed_rows_c2) > 0:
            st.write("Failed rows for Check 2:")
            failed_view2 = failed_rows_c2[cols_for_output].copy()
            failed_view2.columns = make_unique(failed_view2.columns)
            st.dataframe(failed_view2)
        else:
            st.success("There were no issues or failed cases for this check.")

        st.write("Full Check 2 results (ALL rows):")
        preview2 = df_check2[cols_for_output].copy()
        preview2.columns = make_unique(preview2.columns)
        st.dataframe(preview2)

    # =========================================
    #  4. Check 3 – Role-by-role comparison (multiple roles per user, ALL rows)
    # =========================================
    st.header("Check 3 – Role Requested = Role Provisioned")

    with st.expander("Configure Check 3 (User & Role Comparison)", expanded=True):
        c3_1, c3_2 = st.columns(2)
        c3_3, c3_4 = st.columns(2)

        with c3_1:
            sbl_userid_col = st.selectbox(
                "User ID column in Dump File",
                options=df_sbl.columns.tolist(),
                key="sbl_userid_col"
            )

        with c3_2:
            sbl_role_col = st.selectbox(
                "Role column in Dump File",
                options=df_sbl.columns.tolist(),
                key="sbl_role_col"
            )

        with c3_3:
            analytics_username_col = st.selectbox(
                "User identifier column in Saviynt  (either Name(ID) or pure ID)",
                options=df_analytics.columns.tolist(),
                key="analytics_username_col"
            )

        with c3_4:
            analytics_role_col = st.selectbox(
                "Role column in Saviynt",
                options=df_analytics.columns.tolist(),
                key="analytics_role_col"
            )

        key_col_3 = st.selectbox(
            "Key/ID column for reporting (from Saviynt Summary)",
            options=df_analytics.columns.tolist(),
            key="key_col_3"
        )

        username_is_pure_id = st.checkbox(
            "Saviynt user column already contains pure ID (no brackets)",
            value=True,
            help="Tick this if the selected Analytics user column is just the ID (e.g., 223055402), "
                 "not 'Name (223055402)'."
        )

        run_check3 = st.checkbox("Run Check 3 (uses ALL rows in both files)", value=True)

    if run_check3:
        sbl_id_role = df_sbl[[sbl_userid_col, sbl_role_col]].copy()

        col_user = sbl_id_role[sbl_userid_col]
        if isinstance(col_user, pd.DataFrame):
            col_user = col_user.iloc[:, 0]

        col_role = sbl_id_role[sbl_role_col]
        if isinstance(col_role, pd.DataFrame):
            col_role = col_role.iloc[:, 0]

        sbl_id_role["_user_id"] = col_user.astype(str).str.strip().str.upper()
        sbl_id_role["_role_norm"] = col_role.astype(str).str.strip().str.upper()

        sbl_map = (
            sbl_id_role.groupby("_user_id")["_role_norm"]
            .apply(lambda x: set(r for r in x if r))
            .to_dict()
        )

        df_check3 = df_analytics[[analytics_username_col, analytics_role_col, key_col_3]].copy()

        user_col_series = df_check3[analytics_username_col]
        if isinstance(user_col_series, pd.DataFrame):
            user_col_series = user_col_series.iloc[:, 0]

        role_col_series = df_check3[analytics_role_col]
        if isinstance(role_col_series, pd.DataFrame):
            role_col_series = role_col_series.iloc[:, 0]

        if username_is_pure_id:
            df_check3["_user_id"] = user_col_series.astype(str).str.strip().str.upper()
        else:
            df_check3["_user_id"] = user_col_series.apply(extract_id_from_username)

        df_check3["_role_norm"] = role_col_series.astype(str).str.strip().str.upper()

        def compare_row(row):
            uid = row["_user_id"]
            role = row["_role_norm"]

            if not uid:
                return "Failed - No ID extracted from username"

            roles_in_sbl = sbl_map.get(uid, None)
            if roles_in_sbl is None or len(roles_in_sbl) == 0:
                return "Failed - User ID not found in Application File"

            if role in roles_in_sbl:
                return "Passed"
            else:
                return "Failed - Role not found for this ID in Application File"

        df_check3["Check3_Result"] = df_check3.apply(compare_row, axis=1)
        df_check3["Check3_Reason"] = df_check3["Check3_Result"]

        all_results.append(
            df_check3[[key_col_3, "_user_id", analytics_role_col, "Check3_Result", "Check3_Reason"]]
        )

        total_rows_c3 = len(df_check3)
        failed_rows_c3 = df_check3[df_check3["Check3_Result"].str.startswith("Failed")]
        passed_rows_c3 = df_check3[df_check3["Check3_Result"] == "Passed"]

        st.subheader("Check 3 Summary")
        st.write(f"Total rows evaluated (Analytics): {total_rows_c3}")
        st.write(f"Passed: {len(passed_rows_c3)}")
        st.write(f"Failed: {len(failed_rows_c3)}")

        st.write("Failed rows for Check 3 (ALL):")
        if not failed_rows_c3.empty:
            failed_view3 = failed_rows_c3[
                [key_col_3, analytics_username_col, "_user_id", analytics_role_col, "Check3_Result"]
            ].copy()
            failed_view3.columns = make_unique(failed_view3.columns)
            st.dataframe(failed_view3)
        else:
            st.success("There were no failed cases for Check 3.")

        st.write("Passed rows for Check 3 (ALL):")
        if not passed_rows_c3.empty:
            passed_view3 = passed_rows_c3[
                [key_col_3, analytics_username_col, "_user_id", analytics_role_col, "Check3_Result"]
            ].copy()
            passed_view3.columns = make_unique(passed_view3.columns)
            st.dataframe(passed_view3)
        else:
            st.info("There were no passed cases for Check 3.")

        st.write("All rows for Check 3 (ALL):")
        all_view3 = df_check3[
            [key_col_3, analytics_username_col, "_user_id", analytics_role_col, "Check3_Result"]
        ].copy()
        all_view3.columns = make_unique(all_view3.columns)
        st.dataframe(all_view3)

    # =========================================
    #  5. Check 4 – Role & Approver Comparison (Per Approval Level vs Entitlement Owner List)
    # =========================================
    st.header("Check 4 – Role Approval Comparison (Saviynt vs Entitlement Owner List)")

    if df_entitlement is None:
        st.info("Upload an Entitlement Owner List file above to enable Check 4.")
    else:
        with st.expander("Configure Check 4", expanded=True):
            st.markdown("**Saviynt (Analytics) side — each approval level is checked separately**")
            c4_0, c4_1 = st.columns(2)

            with c4_0:
                key_col_4 = st.selectbox(
                    "Key/ID column (Saviynt, for report)",
                    options=df_analytics.columns.tolist(),
                    key="key_col_4"
                )

            with c4_1:
                analytics_role_col_c4 = st.selectbox(
                    "Role column (Saviynt)",
                    options=df_analytics.columns.tolist(),
                    key="analytics_role_col_c4"
                )

            analytics_approver_cols_c4 = st.multiselect(
                "Approval Level column(s) in Saviynt — select ALL levels "
                "(e.g., Approval Level 1, Approval Level 2, Approval Level 3...). "
                "Each one will be validated independently.",
                options=df_analytics.columns.tolist(),
                key="analytics_approver_cols_c4"
            )

            analytics_name_id_c4 = st.checkbox(
                "Saviynt approver values use 'Name (ID)' format",
                value=True,
                key="analytics_name_id_c4",
                help="Tick if values look like 'John Smith (12345)' instead of plain names/IDs."
            )

            st.markdown("---")
            st.markdown("**Entitlement Owner List side — used as reference/lookup table (by Role)**")
            c4_3, c4_4 = st.columns(2)

            with c4_3:
                entitlement_role_col = st.selectbox(
                    "Role column (Entitlement Owner List)",
                    options=df_entitlement.columns.tolist(),
                    key="entitlement_role_col"
                )

            with c4_4:
                entitlement_approver_cols = st.multiselect(
                    "Approver / Owner column(s) (Entitlement Owner List)",
                    options=df_entitlement.columns.tolist(),
                    key="entitlement_approver_cols"
                )

            entitlement_name_id_c4 = st.checkbox(
                "Entitlement Owner List values use 'Name (ID)' format",
                value=False,
                key="entitlement_name_id_c4"
            )

            st.markdown("---")
            st.markdown("**Overall row verdict rule**")
            overall_rule_c4 = st.radio(
                "How should the OVERALL row Status be decided from the individual approval levels?",
                options=[
                    "ALL non-blank levels must be valid owners (strict)",
                    "AT LEAST ONE level must be a valid owner (lenient)"
                ],
                index=0,
                key="overall_rule_c4"
            )

            st.markdown("---")
            st.markdown("**Display options**")
            c4_disp1, c4_disp2 = st.columns(2)
            with c4_disp1:
                hide_blank_roles = st.checkbox(
                    "Hide rows with no Role value (still counted in summary)",
                    value=True,
                    key="hide_blank_roles_c4"
                )
            with c4_disp2:
                status_filter_c4 = st.selectbox(
                    "Show rows",
                    options=["All", "Passed only", "Failed only", "Error only", "NA only"],
                    index=0,
                    key="status_filter_c4"
                )

            run_check4 = st.checkbox("Run Check 4 (evaluates ALL rows, ALL approval levels)", value=True, key="run_check4")

        if run_check4:
            if not analytics_approver_cols_c4 or not entitlement_approver_cols:
                st.warning("Please select at least one Approval Level column and one Entitlement Owner column.")
            else:
                # ---------- Build Entitlement Owner reference dict: {role: set(approvers)} ----------
                entitlement_role_dict = build_role_approver_dict(
                    df_entitlement,
                    entitlement_role_col,
                    entitlement_approver_cols,
                    treat_as_name_id=entitlement_name_id_c4
                )

                strict_mode = overall_rule_c4.startswith("ALL")

                # Fixed, guaranteed-unique names for reference/summary columns
                ENTITLEMENT_COL_NAME = "Entitlement Owner Approver(s) [Reference]"

                # ---------- Row-level evaluation: check EACH approval level independently ----------
                def evaluate_check4_row(row):
                    role_val = row[analytics_role_col_c4]
                    role_norm = "" if pd.isna(role_val) or str(role_val).strip() == "" else str(role_val).strip().upper()

                    result = {"Role": role_val if role_norm else "(No Role)"}

                    # Build unique column names per level using a "Saviynt - " prefix,
                    # so they can never collide with the fixed ENTITLEMENT_COL_NAME
                    # or with each other.
                    def approver_col_name(col):
                        return f"Saviynt - {col} (Approver ID)"

                    def valid_col_name(col):
                        return f"Saviynt - {col} (Valid?)"

                    # -------- Case: No role at all -> Status = NA --------
                    if not role_norm:
                        for col in analytics_approver_cols_c4:
                            result[approver_col_name(col)] = "-"
                            result[valid_col_name(col)] = "-"
                        result[ENTITLEMENT_COL_NAME] = "-"
                        result["Status"] = "NA"
                        result["Details"] = "No role value found in this row"
                        return pd.Series(result)

                    entitlement_approvers = entitlement_role_dict.get(role_norm, set())
                    role_exists_in_entitlement = role_norm in entitlement_role_dict

                    # If the role exists in the Entitlement Owner List but has NO
                    # approvers listed against it, there is nothing to validate against —
                    # this is now treated as an "Error" state (data issue), not Pass/Fail.
                    no_owners_defined = role_exists_in_entitlement and len(entitlement_approvers) == 0

                    level_statuses = []       # True/False per non-blank level
                    failed_levels = []        # names of levels that failed
                    passed_levels = []        # names of levels that passed

                    for col in analytics_approver_cols_c4:
                        level_approvers = parse_multi_values(row[col], treat_as_name_id=analytics_name_id_c4)

                        a_col = approver_col_name(col)
                        v_col = valid_col_name(col)

                        if not level_approvers:
                            # blank level -> not counted towards pass/fail
                            result[a_col] = "-"
                            result[v_col] = "-"
                            continue

                        result[a_col] = ", ".join(sorted(level_approvers))

                        if not role_exists_in_entitlement:
                            result[v_col] = "❌ Role not found"
                            level_statuses.append(False)
                            failed_levels.append(col)
                            continue

                        if no_owners_defined:
                            # Nothing to validate against -> mark as N/A at level, contributes to Error status
                            result[v_col] = "⚠️ No owners defined"
                            continue

                        is_valid = level_approvers.issubset(entitlement_approvers) and len(level_approvers) > 0
                        # "valid" here means every approver listed in this level is a recognized entitlement owner
                        if is_valid:
                            result[v_col] = "✅ Valid"
                            level_statuses.append(True)
                            passed_levels.append(col)
                        else:
                            invalid_ones = level_approvers - entitlement_approvers
                            result[v_col] = f"❌ Invalid ({', '.join(sorted(invalid_ones))})"
                            level_statuses.append(False)
                            failed_levels.append(col)

                    result[ENTITLEMENT_COL_NAME] = (
                        ", ".join(sorted(entitlement_approvers)) if entitlement_approvers else "-"
                    )

                    # ---------- Decide overall status ----------
                    if not role_exists_in_entitlement:
                        result["Status"] = "Failed"
                        result["Details"] = "Role not found in Entitlement Owner List"
                    elif no_owners_defined:
                        # Data issue: role exists but has no approvers listed against it -> Error
                        result["Status"] = "Error"
                        result["Details"] = "Role found, but no approvers listed in Entitlement Owner List (auto-passed)"
                    elif not level_statuses:
                        result["Status"] = "Failed"
                        result["Details"] = "No approver value found in any approval level"
                    else:
                        if strict_mode:
                            overall_pass = all(level_statuses)
                        else:
                            overall_pass = any(level_statuses)

                        if overall_pass:
                            result["Status"] = "Passed"
                            result["Details"] = f"Valid level(s): {', '.join(passed_levels)}"
                        else:
                            result["Status"] = "Failed"
                            if strict_mode:
                                result["Details"] = f"Invalid approver(s) at level(s): {', '.join(failed_levels)}"
                            else:
                                result["Details"] = "None of the approval levels matched a valid entitlement owner"

                    return pd.Series(result)

                check4_results = df_analytics.apply(evaluate_check4_row, axis=1)
                df_check4 = pd.concat(
                    [df_analytics[[key_col_4]].reset_index(drop=True), check4_results.reset_index(drop=True)],
                    axis=1
                )
                df_check4.rename(columns={key_col_4: "Request ID"}, inplace=True)

                # ---------- Safety net: guarantee unique column names no matter what ----------
                if df_check4.columns.duplicated().any():
                    df_check4.columns = make_unique(df_check4.columns.tolist())

                # keep for combined download (full, unfiltered) -- includes Valid? columns for audit trail
                all_results.append(df_check4.copy())

                # ---------- Summary ----------
                total_rows_c4 = len(df_check4)
                passed_rows_c4 = df_check4[df_check4["Status"] == "Passed"]
                failed_rows_c4 = df_check4[df_check4["Status"] == "Failed"]
                error_rows_c4 = df_check4[df_check4["Status"] == "Error"]
                na_rows_c4 = df_check4[df_check4["Status"] == "NA"]

                st.subheader("Check 4 Summary")
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Total Rows", total_rows_c4)
                m2.metric("✅ Passed", len(passed_rows_c4))
                m3.metric("❌ Failed", len(failed_rows_c4))
                m4.metric("⚠️ Error", len(error_rows_c4))
                m5.metric("⬜ NA", len(na_rows_c4))

                # ---------- Apply display filters ----------
                display_df = df_check4.copy()

                if hide_blank_roles:
                    display_df = display_df[display_df["Role"] != "(No Role)"]

                if status_filter_c4 == "Passed only":
                    display_df = display_df[display_df["Status"] == "Passed"]
                elif status_filter_c4 == "Failed only":
                    display_df = display_df[display_df["Status"] == "Failed"]
                elif status_filter_c4 == "Error only":
                    display_df = display_df[display_df["Status"] == "Error"]
                elif status_filter_c4 == "NA only":
                    display_df = display_df[display_df["Status"] == "NA"]

                # ---------- Column order: Request ID, Role, [Approver ID cols only], Entitlement Owners, Status, Details ----------
                # NOTE: "(Valid?)" columns are intentionally EXCLUDED from the on-screen table
                # (they remain in the full df_check4 / CSV download for audit purposes).
                level_cols_ordered = []
                for col in analytics_approver_cols_c4:
                    level_cols_ordered.append(f"Saviynt - {col} (Approver ID)")

                final_col_order = (
                    ["Request ID", "Role"]
                    + level_cols_ordered
                    + [ENTITLEMENT_COL_NAME, "Status", "Details"]
                )
                final_col_order = [c for c in final_col_order if c in display_df.columns]
                display_df = display_df[final_col_order]

                # ---------- Color-coded Status (version-safe styling) ----------
                def highlight_status(val):
                    if val == "Passed":
                        return "background-color: #14532d; color: #dcfce7; font-weight: 600;"
                    elif val == "Failed":
                        return "background-color: #7f1d1d; color: #fee2e2; font-weight: 600;"
                    elif val == "Error":
                        return "background-color: #78350f; color: #fef3c7; font-weight: 600;"
                    elif val == "NA":
                        return "background-color: #374151; color: #d1d5db; font-weight: 600;"
                    return ""

                def highlight_valid_cols(val):
                    if isinstance(val, str):
                        if val.startswith("✅"):
                            return "color: #86efac;"
                        elif val.startswith("❌"):
                            return "color: #fca5a5;"
                        elif val.startswith("⚠️"):
                            return "color: #fde68a;"
                    return ""

                st.write(f"Showing {len(display_df)} of {total_rows_c4} rows")

                # No "(Valid?)" columns remain in display_df, so this will just be an empty list
                valid_cols_to_style = [c for c in display_df.columns if c.endswith("(Valid?)")]

                try:
                    if display_df.columns.duplicated().any():
                        raise ValueError("Duplicate columns detected in display table.")

                    styled_df = style_dataframe_safe(
                        display_df,
                        status_col="Status",
                        valid_cols=valid_cols_to_style,
                        status_style_fn=highlight_status,
                        valid_style_fn=highlight_valid_cols
                    )
                    st.dataframe(styled_df, use_container_width=True, hide_index=True)
                except Exception as style_err:
                    # Fallback: if styling fails for any reason, show plain dataframe instead of crashing
                    st.warning(f"Could not apply cell styling (showing plain table instead). Details: {style_err}")
                    st.dataframe(display_df, use_container_width=True, hide_index=True)

                # ---------- Download this check's results separately (FULL data incl. Valid? columns) ----------
                csv_c4 = df_check4.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "📥 Download Check 4 Results (CSV) — includes full audit detail",
                    data=csv_c4,
                    file_name="check4_role_approval_results.csv",
                    mime="text/csv",
                    key="download_check4"
                )

    # =========================================
    #  6. Combined Download (all checks)
    # =========================================
    st.header("Download Combined Results")

    if all_results:
        try:
            combined_df = pd.concat(all_results, ignore_index=True, sort=False)

            st.write(f"Combined report has {len(combined_df)} rows across all checks run.")

            csv_data = combined_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Combined Results (CSV)",
                data=csv_data,
                file_name="sox_access_comparison_results.csv",
                mime="text/csv"
            )
        except Exception as e:
            st.warning(f"Could not combine all results into a single file: {e}")
    else:
        st.info("Run at least one check to enable combined download.")

# ============================================================
# ============================================================
#   TOOL 2: CM TOOL 1  (Sequential Ticket Automation - Pulse)
# ============================================================
# ============================================================
def run_cm_tool_1():



    TICKET_TYPE_LABELS = {
        "Pulse Tickets": {"Yes": "Normal Ticket", "No": "GXP Standard Ticket"},
        "Other Pulse Tickets": {"Yes": "Standard Ticket", "No": "GXP Normal Ticket"},
    }

    LOG_FILE_PATH = os.path.join(os.getcwd(), "automation_log.txt")

    defaults = {
        "step": 0,
        "sso_id": "",
        "password": "",
        "platform": "Pulse Tickets",
        "ticket_type_key": "Yes",
        "tickets": [],
        "logs": [],
        "download_dir": None,
        "headless": True,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    _log_placeholder = None
    _image_placeholder = None


    def log(msg: str):
        timestamped = f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}"
        st.session_state.logs.append(msg)
        print(msg)
        try:
            with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
                f.write(timestamped + "\n")
        except Exception as e:
            print(f"Could not write to log file: {e}")
        if _log_placeholder is not None:
            _log_placeholder.text("\n".join(st.session_state.logs[-400:]))


    def show_screenshot(driver, caption: str = ""):
        if _image_placeholder is None:
            return
        try:
            png_bytes = driver.get_screenshot_as_png()
            _image_placeholder.image(png_bytes, caption=caption, use_container_width=True)
        except Exception as e:
            log(f"  (could not capture screenshot: {e})")


    def poll_until(driver, ticket, description, check_fn, timeout=60, interval=2):
        start = time.time()
        while time.time() - start < timeout:
            show_screenshot(driver, f"[{ticket}] {description}...")
            try:
                result = check_fn()
                if result:
                    return result
            except StaleElementReferenceException:
                pass
            except Exception:
                pass
            time.sleep(interval)
        return None


    # ==========================================================
    # Shadow-DOM-aware element finders
    # ==========================================================
    DEEP_QS_JS = """
    function deepQuerySelectorAll(selector, root) {
        root = root || document;
        let results = [];
        try { results = Array.from(root.querySelectorAll(selector)); } catch (e) {}
        const all = root.querySelectorAll('*');
        for (const el of all) {
            if (el.shadowRoot) {
                results = results.concat(deepQuerySelectorAll(selector, el.shadowRoot));
            }
        }
        return results;
    }
    return deepQuerySelectorAll(arguments[0]);
    """

    DEEP_TEXT_JS = """
    function getDirectText(el) {
        let text = '';
        for (const node of el.childNodes) {
            if (node.nodeType === Node.TEXT_NODE) { text += node.textContent; }
        }
        return text.trim();
    }
    function isVisible(el) {
        return !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
    }
    function deepFindByText(text, root) {
        root = root || document;
        const all = root.querySelectorAll('*');
        for (const el of all) {
            if (el.shadowRoot) {
                const res = deepFindByText(text, el.shadowRoot);
                if (res) return res;
            }
            if (getDirectText(el) === text && isVisible(el)) {
                return el;
            }
        }
        return null;
    }
    return deepFindByText(arguments[0]);
    """

    DEEP_ANY_VISIBLE_JS = """
    function isVisible(el) {
        return !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
    }
    function deepAnyVisible(selector, root) {
        root = root || document;
        let found = [];
        try { found = Array.from(root.querySelectorAll(selector)); } catch (e) {}
        for (const el of found) {
            if (isVisible(el)) return true;
        }
        const all = root.querySelectorAll('*');
        for (const el of all) {
            if (el.shadowRoot) {
                if (deepAnyVisible(selector, el.shadowRoot)) return true;
            }
        }
        return false;
    }
    return deepAnyVisible(arguments[0]);
    """


    def deep_find_all(driver, css_selector):
        try:
            return driver.execute_script(DEEP_QS_JS, css_selector) or []
        except Exception:
            return []


    def deep_find(driver, css_selector):
        els = deep_find_all(driver, css_selector)
        for el in els:
            try:
                if el.is_displayed():
                    return el
            except Exception:
                continue
        return els[0] if els else None


    def deep_find_by_text(driver, text):
        try:
            return driver.execute_script(DEEP_TEXT_JS, text)
        except Exception:
            return None


    def deep_any_visible(driver, css_selector):
        try:
            return bool(driver.execute_script(DEEP_ANY_VISIBLE_JS, css_selector))
        except Exception:
            return False


    def safe_click(driver, el):
        try:
            el.click()
            return
        except Exception:
            pass
        try:
            driver.execute_script("arguments[0].click();", el)
            return
        except Exception:
            pass
        try:
            ActionChains(driver).move_to_element(el).click().perform()
        except Exception as e:
            raise e


    # ==========================================================
    # Iframe-aware search wrapper (iframes found via shadow-DOM-aware search)
    # ==========================================================
    def find_across_frames(driver, finder_fn, max_depth=3):
        el = finder_fn(driver)
        if el:
            return el

        driver.switch_to.default_content()
        el = finder_fn(driver)
        if el:
            return el

        def search_frames(depth):
            if depth > max_depth:
                return None
            frames = deep_find_all(driver, "iframe")
            for frame in frames:
                try:
                    driver.switch_to.frame(frame)
                except Exception:
                    continue
                found = finder_fn(driver)
                if found:
                    return found
                nested = search_frames(depth + 1)
                if nested:
                    return nested
                driver.switch_to.parent_frame()
            return None

        result = search_frames(1)
        if result:
            return result

        driver.switch_to.default_content()
        return None


    # ==========================================================
    # Loading-spinner / step-completion helpers
    # ==========================================================
    SPINNER_SELECTORS = [
        "[class*='spinner' i]",
        "[class*='loading' i]",
        "[aria-busy='true']",
        ".now-loading-indicator",
    ]


    def is_loading(driver):
        for sel in SPINNER_SELECTORS:
            if deep_any_visible(driver, sel):
                return True
        return False


    def wait_for_spinner_to_clear(driver, ticket, context_label, timeout=30, interval=1):
        start = time.time()
        while time.time() - start < timeout:
            if not is_loading(driver):
                return True
            show_screenshot(driver, f"[{ticket}] {context_label}: waiting for loading to finish...")
            time.sleep(interval)
        return not is_loading(driver)


    def wait_until_gone(driver, ticket, description, css_selector, timeout=30, interval=1):
        start = time.time()
        while time.time() - start < timeout:
            if not deep_any_visible(driver, css_selector):
                return True
            show_screenshot(driver, f"[{ticket}] {description}: waiting to close...")
            time.sleep(interval)
        return not deep_any_visible(driver, css_selector)


    # ==========================================================
    # Diagnostics
    # ==========================================================
    def debug_dump_page_state(driver, ticket):
        try:
            driver.switch_to.default_content()
            log(f"  [DEBUG] === Diagnostics for ticket {ticket} ===")
            log(f"  [DEBUG] Current URL: {driver.current_url}")
            log(f"  [DEBUG] Page title: {driver.title}")
            log(f"  [DEBUG] Number of window handles: {len(driver.window_handles)}")

            iframes = deep_find_all(driver, "iframe")
            log(f"  [DEBUG] Found {len(iframes)} iframe(s) on the page (incl. shadow DOM):")
            for i, f in enumerate(iframes):
                try:
                    fid = f.get_attribute("id")
                    fsrc = f.get_attribute("src")
                    fname = f.get_attribute("name")
                    log(f"    iframe[{i}] id={fid!r} name={fname!r} src={str(fsrc)[:150]!r}")
                except Exception as e:
                    log(f"    iframe[{i}] (could not read attributes: {e})")

            candidate_selectors = [
                "button.additional-actions-context-menu-button",
                "button[aria-label='additional actions']",
                "button[aria-label*='action' i]",
                "button[title*='action' i]",
                "[aria-haspopup='true']",
                "button.icon-menu",
            ]

            def scan_frame(label):
                for sel in candidate_selectors:
                    els = deep_find_all(driver, sel)
                    if els:
                        details = []
                        for e in els:
                            try:
                                details.append(
                                    f"(visible={e.is_displayed()}, text={e.text!r}, "
                                    f"aria-label={e.get_attribute('aria-label')!r})"
                                )
                            except Exception as ex:
                                details.append(f"(error reading: {ex})")
                        log(f"    [{label}] {sel!r} -> {len(els)} match(es): {details}")
                    else:
                        log(f"    [{label}] {sel!r} -> 0 matches")

            scan_frame("top-level")
            for i, f in enumerate(iframes):
                try:
                    driver.switch_to.frame(f)
                    scan_frame(f"iframe[{i}]")
                    driver.switch_to.default_content()
                except Exception as e:
                    log(f"    [iframe[{i}]] could not switch into frame: {e}")
                    driver.switch_to.default_content()

            log(f"  [DEBUG] === End diagnostics for {ticket} ===")

        except Exception as e:
            log(f"  [DEBUG] debug_dump_page_state itself failed: {e}")
        finally:
            driver.switch_to.default_content()


    # ==========================================================
    # Binary detection (HARDENED: known paths -> PATH -> filesystem-wide search)
    # ==========================================================
    def _filesystem_search(patterns, roots=("/usr", "/opt", "/snap", "/root", "/home")):
        """Last-resort recursive search for a binary matching any of `patterns`."""
        for root in roots:
            if not os.path.isdir(root):
                continue
            for pattern in patterns:
                try:
                    matches = glob.glob(os.path.join(root, "**", pattern), recursive=True)
                except Exception:
                    matches = []
                for m in matches:
                    if os.path.isfile(m) and os.access(m, os.X_OK) or os.path.isfile(m):
                        return m
        return None


    def find_browser_binary():
        candidates = [
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/lib/chromium/chromium",
            "/usr/lib/chromium-browser/chromium-browser",
            "/snap/bin/chromium",
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
        for name in ["chromium", "chromium-browser", "google-chrome", "google-chrome-stable"]:
            found = shutil.which(name)
            if found:
                return found
        # Last resort: search the filesystem (covers unusual apt install locations)
        found = _filesystem_search(["chromium", "chromium-browser", "google-chrome*"])
        return found


    def find_chromedriver_binary():
        candidates = [
            "/usr/bin/chromedriver",
            "/usr/lib/chromium/chromedriver",
            "/usr/lib/chromium-browser/chromedriver",
            "/snap/bin/chromedriver",
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
        found = shutil.which("chromedriver")
        if found:
            return found
        # Last resort: search the filesystem
        found = _filesystem_search(["chromedriver"])
        if found:
            try:
                os.chmod(found, 0o755)  # ensure it's executable
            except Exception:
                pass
        return found


    def run_environment_diagnostics():
        """Logs what's actually installed, to make future failures easy to diagnose."""
        log("=== Environment diagnostics ===")
        try:
            which_chromium = shutil.which("chromium") or shutil.which("chromium-browser")
            which_chromedriver = shutil.which("chromedriver")
            log(f"  which chromium: {which_chromium}")
            log(f"  which chromedriver: {which_chromedriver}")
        except Exception as e:
            log(f"  which check failed: {e}")

        try:
            result = subprocess.run(
                ["dpkg", "-l"], capture_output=True, text=True, timeout=10
            )
            lines = [l for l in result.stdout.splitlines() if "chrom" in l.lower()]
            if lines:
                log("  dpkg packages matching 'chrom':")
                for l in lines:
                    log(f"    {l}")
            else:
                log("  dpkg: no chromium/chromedriver packages found "
                    "(packages.txt may be missing or app wasn't rebooted after adding it)")
        except Exception as e:
            log(f"  dpkg check failed or unavailable: {e}")

        log("=== End environment diagnostics ===")


    # ==========================================================
    # Selenium driver setup
    # ==========================================================
    def build_driver(download_dir: str, headless: bool = True):
        run_environment_diagnostics()

        options = Options()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")

        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled": True,
        }
        options.add_experimental_option("prefs", prefs)

        chromium_path = find_browser_binary()
        chromedriver_path = find_chromedriver_binary()

        log(f"Detected chromium binary: {chromium_path}")
        log(f"Detected chromedriver binary: {chromedriver_path}")

        if chromium_path:
            options.binary_location = chromium_path

        driver = None

        if chromedriver_path:
            try:
                service = Service(chromedriver_path)
                driver = webdriver.Chrome(service=service, options=options)
                log(f"Using detected chromedriver at {chromedriver_path}")
            except Exception as e:
                log(f"Detected chromedriver failed to launch: {e}")
        else:
            log("No chromedriver binary found via known paths, system PATH, or filesystem search. "
                "This almost certainly means packages.txt is missing 'chromium-driver' "
                "or the app has not been rebooted since adding it.")

        if driver is None:
            try:
                driver = webdriver.Chrome(options=options)
                log("Using Selenium Manager auto-resolved driver")
            except Exception as e:
                log(f"Selenium Manager failed: {e}")

        if driver is None:
            try:
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=options)
                log("Using webdriver-manager driver")
            except Exception as e:
                log(f"webdriver-manager also failed: {e}")
                raise RuntimeError(
                    "Could not start Chrome/Chromium via any method. "
                    "Check that packages.txt contains 'chromium' and 'chromium-driver', "
                    "and that you have rebooted the app (Manage app -> Reboot) after adding it."
                )

        try:
            driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": download_dir
            })
            log(f"Download behavior explicitly set via CDP to: {download_dir}")
        except Exception as e:
            log(f"Warning: could not set CDP download behavior: {e}")

        return driver


    # ==========================================================
    # Login — fully automated username + password, then wait for MFA push
    # ==========================================================
    def login(driver, sso_id: str, password: str):
        login_url = "https://pulse.service-now.com/now/nav/ui/home"
        expected_domain = "pulse.service-now.com"

        log("Opening login page...")
        driver.get(login_url)
        show_screenshot(driver, "Login page loaded")

        try:
            username_field = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))
            )
            username_field.send_keys(sso_id)
            show_screenshot(driver, "Username entered")
            log("Entered SSO ID, clicking Next...")

            next_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Next')]"))
            )
            next_btn.click()
        except TimeoutException:
            show_screenshot(driver, "FAILED at username step")
            raise TimeoutException("Could not find username field or Next button on login page")

        log("Waiting 2 seconds before entering password...")
        time.sleep(2)
        show_screenshot(driver, "Waiting before password entry")

        try:
            password_field = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))
            )
            password_field.send_keys(password)
            show_screenshot(driver, "Password entered")
            log("Entered password, submitting...")

            login_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Log In & Remember Me')]"))
            )
            login_btn.click()
            log("Clicked 'Log In & Remember Me'")
        except TimeoutException:
            show_screenshot(driver, "FAILED at password step")
            raise TimeoutException("Could not find password field or Login button")

        log("Password submitted. A push notification should now be sent to your "
            "authenticator app. Please approve it on your phone.")
        log("Waiting for you to approve the push notification... (up to 5 minutes)")

        start = time.time()
        timeout = 300
        while time.time() - start < timeout:
            if expected_domain in driver.current_url and "login" not in driver.current_url.lower():
                break
            show_screenshot(driver, "Waiting for MFA push approval on your phone...")
            time.sleep(3)
        else:
            raise TimeoutException(
                "Login was not completed within 5 minutes. "
                "Did you approve the push notification on your phone?"
            )

        log("Login successful! Resuming automation...")
        show_screenshot(driver, "Logged in successfully")


    # ==========================================================
    # Step 1: search box (shadow-DOM aware)
    # ==========================================================
    def open_search_and_type(driver, ticket):
        def find_input():
            return deep_find(driver, "#sncwsgs-typeahead-input")

        search_box = poll_until(driver, ticket, "Step 1a: finding search input", find_input, timeout=15, interval=1)

        if search_box is None:
            def find_trigger():
                for sel in ["input[placeholder='Search']", "[aria-label='Search']",
                            "button[aria-label*='search' i]", ".search-box", ".global-search"]:
                    el = deep_find(driver, sel)
                    if el:
                        return el
                return None

            trigger = poll_until(driver, ticket, "Step 1a: finding search trigger", find_trigger, timeout=15, interval=1)
            if trigger is None:
                raise TimeoutException("Could not find search box/icon (even searching shadow DOM)")
            safe_click(driver, trigger)
            show_screenshot(driver, f"[{ticket}] Step 1a: clicked search trigger")

            search_box = poll_until(driver, ticket, "Step 1b: waiting for expanded search input", find_input, timeout=15, interval=1)
            if search_box is None:
                raise TimeoutException("Expanded search input never appeared after clicking search")

        safe_click(driver, search_box)
        search_box.clear()
        search_box.send_keys(ticket)
        search_box.send_keys(Keys.ENTER)
        show_screenshot(driver, f"[{ticket}] Step 1: searched, waiting for results")
        return search_box


    # ==========================================================
    # Full ticket download flow
    # ==========================================================
    def download_ticket_pdf(driver, ticket: str):
        driver.switch_to.default_content()

        log("  Step 1: clicking search box and typing ticket")
        open_search_and_type(driver, ticket)

        log("  Step 2: waiting for loading spinner to clear, then results")
        wait_for_spinner_to_clear(driver, ticket, "Step 2 (search)", timeout=30)

        def check_results(d):
            if is_loading(d):
                return None
            el = deep_find(d, "ul[aria-labelledby='section-EXACT_MATCH_SECTION'] li[data-testclass='sn-global-search-record']")
            if el:
                return el
            return deep_find(d, "li[data-testclass='sn-global-search-record']")

        result = poll_until(
            driver, ticket, "Step 2: waiting for search results",
            lambda: check_results(driver), timeout=60, interval=2
        )
        if result is None:
            show_screenshot(driver, f"[{ticket}] Step 2: FAILED — no results after 60s")
            raise TimeoutException(f"No search results appeared for ticket {ticket} within 60s")

        log("  Step 2: result found, clicking it")
        safe_click(driver, result)
        show_screenshot(driver, f"[{ticket}] Step 2: opened ticket")

        time.sleep(3)

        log("  Step 2b: confirming ticket detail page loaded")
        wait_for_spinner_to_clear(driver, ticket, "Step 2b (opening ticket)", timeout=30)

        menu_selectors = [
            "button.additional-actions-context-menu-button",
            "button[aria-label='additional actions']",
            "button[aria-label='Additional actions']",
            "button[aria-label*='additional actions' i]",
            "button[title*='additional actions' i]",
        ]

        def check_menu_button(d):
            if is_loading(d):
                return None
            for sel in menu_selectors:
                el = deep_find(d, sel)
                if el:
                    return el
            return None

        menu_btn = poll_until(
            driver, ticket, "Step 2b: confirming ticket page loaded",
            lambda: find_across_frames(driver, check_menu_button),
            timeout=40, interval=2
        )
        if menu_btn is None:
            log(f"  Step 2b: FAILED for {ticket} — running diagnostics...")
            debug_dump_page_state(driver, ticket)
            show_screenshot(driver, f"[{ticket}] Step 2b: FAILED — see automation_log.txt")
            raise TimeoutException(
                f"Ticket detail page never finished loading for {ticket}. "
                f"See {LOG_FILE_PATH} for diagnostics."
            )

        log("  Step 3: opening 'Additional actions' menu")
        safe_click(driver, menu_btn)
        show_screenshot(driver, f"[{ticket}] Step 3: menu opened")

        time.sleep(1)

        log("  Step 4: hovering 'Export'")

        def check_export_item(d):
            return deep_find_by_text(d, "Export")

        export_item = poll_until(
            driver, ticket, "Step 4: waiting for Export menu item",
            lambda: find_across_frames(driver, check_export_item),
            timeout=20, interval=1
        )
        if export_item is None:
            debug_dump_page_state(driver, ticket)
            raise TimeoutException(f"'Export' menu item never appeared for ticket {ticket}")
        ActionChains(driver).move_to_element(export_item).perform()
        show_screenshot(driver, f"[{ticket}] Step 4: hovering Export")

        log("  Step 5: clicking 'PDF' in flyout")

        def check_pdf_item(d):
            return deep_find_by_text(d, "PDF")

        pdf_item = poll_until(
            driver, ticket, "Step 5: waiting for PDF flyout item",
            lambda: find_across_frames(driver, check_pdf_item),
            timeout=20, interval=1
        )
        if pdf_item is None:
            debug_dump_page_state(driver, ticket)
            raise TimeoutException(f"'PDF' flyout item never appeared for ticket {ticket}")
        safe_click(driver, pdf_item)
        show_screenshot(driver, f"[{ticket}] Step 5: clicked PDF")

        log("  Step 5b: confirming 'Export to PDF' dialog fully opened")

        def check_dialog_open(d):
            return deep_find(d, "#ok_button")

        ok_btn = poll_until(
            driver, ticket, "Step 5b: waiting for Export dialog to appear",
            lambda: find_across_frames(driver, check_dialog_open),
            timeout=20, interval=1
        )
        if ok_btn is None:
            raise TimeoutException(f"Export dialog 'ok_button' never appeared for ticket {ticket}")

        log("  Step 6: clicking Export button in dialog")
        safe_click(driver, ok_btn)
        show_screenshot(driver, f"[{ticket}] Step 6: export triggered")

        log("  Step 6b: confirming orientation dialog closed")
        wait_until_gone(driver, ticket, "Step 6b (orientation dialog)", "#ok_button", timeout=15)

        log("  Step 7: waiting for Download button to appear AND become enabled "
            "(PDF generation can take time)")

        _logged_once = {"done": False}

        def check_download_button(d):
            if is_loading(d):
                return None
            el = deep_find(d, "#download_button")
            if el is None:
                return None
            try:
                disabled_attr = el.get_attribute("disabled")
                aria_disabled = el.get_attribute("aria-disabled")
                classes = (el.get_attribute("class") or "").lower()
                if not _logged_once["done"]:
                    log(f"    [debug] download_button found: disabled={disabled_attr!r} "
                        f"aria-disabled={aria_disabled!r} class={classes!r} "
                        f"is_enabled={el.is_enabled()}")
                    _logged_once["done"] = True
                if disabled_attr or aria_disabled == "true" or "disabled" in classes:
                    return None
                if not el.is_enabled():
                    return None
            except Exception:
                return None
            return el

        download_btn = poll_until(
            driver, ticket, "Step 7: waiting for PDF generation to finish (button enabled)",
            lambda: find_across_frames(driver, check_download_button),
            timeout=120, interval=3
        )
        if download_btn is None:
            raise TimeoutException(f"Download button never became enabled for ticket {ticket}")
        show_screenshot(driver, f"[{ticket}] Step 7: export ready (button enabled)")

        time.sleep(1)

        safe_click(driver, download_btn)
        log("  Step 8: clicked Download")
        show_screenshot(driver, f"[{ticket}] Step 8: downloading...")

        log("  Step 8b: waiting for 'Export Complete' dialog to close on its own "
            "(this can take a while depending on file size)")
        dialog_closed = wait_until_gone(
            driver, ticket, "Step 8b (export complete dialog)", "#download_button", timeout=90
        )
        if dialog_closed:
            log(f"  Step 8b: dialog closed normally for {ticket}")
        else:
            log(f"  Step 8b: dialog did NOT close within 90s for {ticket} — "
                f"will check if the file downloaded anyway, and try to force-close the dialog.")
        show_screenshot(driver, f"[{ticket}] Step 8b: done waiting on dialog")

        driver.switch_to.default_content()


    # ==========================================================
    # Download detection — checks multiple candidate folders as a safety net
    # ==========================================================
    def get_candidate_download_dirs(download_dir: str):
        candidates = [download_dir]
        home = os.path.expanduser("~")
        default_downloads = os.path.join(home, "Downloads")
        if default_downloads not in candidates:
            candidates.append(default_downloads)
        return candidates


    def wait_for_new_download(download_dir: str, before_files: set, timeout: int = 120):
        candidate_dirs = get_candidate_download_dirs(download_dir)
        start = time.time()
        while time.time() - start < timeout:
            current_files = set()
            for d in candidate_dirs:
                current_files |= set(glob.glob(os.path.join(d, "*")))
            new_files = current_files - before_files
            finished = [f for f in new_files if not f.endswith(".crdownload") and not f.endswith(".tmp")]
            still_downloading = any(
                f.endswith(".crdownload") or f.endswith(".tmp") for f in current_files
            )
            if finished and not still_downloading:
                return finished[0]
            time.sleep(1)
        return None


    def rename_downloaded_file(filepath: str, ticket_number: str, ticket_type_label: str, download_dir: str):
        if not filepath or not os.path.exists(filepath):
            return None
        safe_label = ticket_type_label.replace(" ", "_")
        new_name = os.path.join(download_dir, f"{ticket_number}_{safe_label}.pdf")
        try:
            shutil.move(filepath, new_name)
            return new_name
        except Exception as e:
            log(f"Error moving/renaming file: {e}")
            return None


    def close_export_dialog(driver, ticket):
        def check_still_open(d):
            return deep_find(d, "#download_button")

        still_open = find_across_frames(driver, check_still_open)
        if not still_open:
            driver.switch_to.default_content()
            return

        def find_cancel_or_close(d):
            el = deep_find_by_text(d, "Cancel")
            if el:
                return el
            for sel in ["button[aria-label='Close']", "button[title='Close']", ".close", "[aria-label*='close' i]"]:
                el = deep_find(d, sel)
                if el:
                    return el
            return None

        btn = find_across_frames(driver, find_cancel_or_close)
        if btn:
            try:
                safe_click(driver, btn)
                log(f"  Force-closed 'Export Complete' dialog for {ticket}")
                show_screenshot(driver, f"[{ticket}] Dialog force-closed")
            except Exception as e:
                log(f"  Could not click Cancel/Close on dialog for {ticket}: {e}")
        else:
            log(f"  Dialog still open for {ticket} but no Cancel/Close button found")
        driver.switch_to.default_content()


    def run_automation(sso_id, password, tickets, ticket_type_label, download_dir, headless):
        driver = None
        try:
            try:
                with open(LOG_FILE_PATH, "w", encoding="utf-8") as f:
                    f.write(f"=== Automation run started {datetime.datetime.now()} ===\n")
            except Exception:
                pass

            driver = build_driver(download_dir, headless=headless)
            login(driver, sso_id, password)

            for ticket in tickets:
                try:
                    log(f"Processing ticket: {ticket}")
                    before_files = set()
                    for d in get_candidate_download_dirs(download_dir):
                        before_files |= set(glob.glob(os.path.join(d, "*")))

                    download_ticket_pdf(driver, ticket)

                    downloaded_file = wait_for_new_download(download_dir, before_files, timeout=120)

                    close_export_dialog(driver, ticket)

                    if downloaded_file:
                        renamed = rename_downloaded_file(downloaded_file, ticket, ticket_type_label, download_dir)
                        if renamed:
                            log(f"Ticket {ticket} downloaded and renamed to: {os.path.basename(renamed)}")
                        else:
                            log(f"Ticket {ticket} downloaded but renaming failed. File was at: {downloaded_file}")
                    else:
                        log(f"Ticket {ticket}: download did not complete within 120s. "
                            f"Checked folders: {get_candidate_download_dirs(download_dir)}")

                except TimeoutException as e:
                    msg = str(e).strip() or "(no additional details)"
                    log(f"Ticket {ticket}: TIMEOUT — {msg}")
                    show_screenshot(driver, f"[{ticket}] TIMEOUT")
                except Exception as ex:
                    log(f"Ticket {ticket}: ERROR — {type(ex).__name__}: {ex}")
                    show_screenshot(driver, f"[{ticket}] ERROR")

            log("All tickets processed.")
            log(f"Full diagnostic log saved to: {LOG_FILE_PATH}")

        except Exception as ex:
            log(f"Error: {str(ex)}")
        finally:
            if driver is not None:
                driver.quit()
            log("Automation completed. Browser closed.")


    def make_zip(download_dir: str) -> str:
        zip_path = os.path.join(download_dir, "tickets.zip")
        with zipfile.ZipFile(zip_path, "w") as zf:
            for f in glob.glob(os.path.join(download_dir, "*.pdf")):
                zf.write(f, os.path.basename(f))
        return zip_path


    # ==========================================================
    # UI
    # ==========================================================
    st.title("Sequential Ticket Automation")

    st.caption(f"Detailed logs are also saved to: `{LOG_FILE_PATH}`")

    step = st.session_state.step

    if step == 0:
        st.header("Login")
        st.session_state.sso_id = st.text_input("SSO ID", value=st.session_state.sso_id)
        st.session_state.password = st.text_input("Password", type="password", value=st.session_state.password)

        st.info(
            "Your SSO ID and password will be entered automatically. After that, "
            "you'll just need to approve the push notification on your authenticator "
            "app (e.g. Okta Verify / Duo / Microsoft Authenticator) on your phone — "
            "no browser interaction needed."
        )

        st.session_state.headless = st.checkbox(
            "Run headless (leave checked for cloud deployment)",
            value=st.session_state.headless,
        )

        if st.button("Next"):
            if st.session_state.sso_id and st.session_state.password:
                st.session_state.step = 1
                st.rerun()
            else:
                st.warning("Please enter your SSO ID and password.")

    elif step == 1:
        st.subheader("Step 1: Select platform")
        st.session_state.platform = st.radio(
            "Platform",
            list(TICKET_TYPE_LABELS.keys()),
            index=list(TICKET_TYPE_LABELS.keys()).index(st.session_state.platform),
        )

        st.subheader("Step 2: Select ticket type")
        labels = TICKET_TYPE_LABELS[st.session_state.platform]
        label_values = list(labels.values())
        default_label = labels[st.session_state.ticket_type_key] if st.session_state.ticket_type_key in labels else label_values[0]
        choice_label = st.radio("Ticket type", label_values, index=label_values.index(default_label))
        st.session_state.ticket_type_key = [k for k, v in labels.items() if v == choice_label][0]

        st.subheader("Step 3: Paste TicketNumber column (one per line)")
        raw = st.text_area("Tickets", height=200)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Back"):
                st.session_state.step = 0
                st.rerun()
        with col2:
            if st.button("Load Tickets"):
                tickets = [l.strip() for l in raw.splitlines() if l.strip()]
                if tickets:
                    st.session_state.tickets = tickets
                    st.session_state.step = 2
                    st.rerun()
                else:
                    st.warning("No tickets found in pasted text.")

    elif step == 2:
        st.subheader("Ready to run")

        labels = TICKET_TYPE_LABELS[st.session_state.platform]
        ticket_type_label = labels[st.session_state.ticket_type_key]

        st.write(f"**Platform:** {st.session_state.platform}")
        st.write(f"**Ticket type:** {ticket_type_label}")

        n = len(st.session_state.tickets)
        est = 60 + n * 150
        m, s = divmod(est, 60)
        st.info(f"Estimated time (excluding manual MFA approval): up to {m} min {s} sec")

        st.write("**Ticket List**")
        st.write(st.session_state.tickets)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Back"):
                st.session_state.step = 1
                st.rerun()
        with col2:
            start_clicked = st.button("Start Process")

        col_left, col_right = st.columns([1, 1])
        with col_left:
            st.write("**Live Browser View**")
            _image_placeholder = st.empty()
        with col_right:
            st.write("**Process Log**")
            _log_placeholder = st.empty()
            _log_placeholder.text("\n".join(st.session_state.logs[-400:]))

        if start_clicked:
            tmp_dir = tempfile.mkdtemp()
            st.session_state.download_dir = tmp_dir
            st.session_state.logs = []
            run_automation(
                st.session_state.sso_id,
                st.session_state.password,
                st.session_state.tickets,
                ticket_type_label,
                tmp_dir,
                st.session_state.headless,
            )
            st.session_state.password = ""

        if os.path.exists(LOG_FILE_PATH):
            with open(LOG_FILE_PATH, "r", encoding="utf-8") as f:
                log_content = f.read()
            st.download_button("Download full diagnostic log (automation_log.txt)",
                                log_content, file_name="automation_log.txt")

        if st.session_state.download_dir:
            pdf_files = sorted(glob.glob(os.path.join(st.session_state.download_dir, "*.pdf")))
            if pdf_files:
                st.write("### Downloaded Tickets")
                st.write(f"{len(pdf_files)} PDF(s) ready:")

                for pdf_path in pdf_files:
                    fname = os.path.basename(pdf_path)
                    fsize_kb = os.path.getsize(pdf_path) / 1024
                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            f"⬇ {fname} ({fsize_kb:.0f} KB)",
                            f,
                            file_name=fname,
                            mime="application/pdf",
                            key=f"dl_{fname}",
                        )

                zip_path = make_zip(st.session_state.download_dir)
                with open(zip_path, "rb") as f:
                    st.download_button(
                        "⬇ Download ALL as ZIP",
                        f,
                        file_name="tickets.zip",
                        mime="application/zip",
                        key="dl_zip_all",
                    )
            else:
                st.info("No PDFs downloaded yet in this session.")


# ============================================================
# ============================================================
#   TOOL 3: CM TOOL 2  (Change Management Tool - Excel report)
# ============================================================
# ============================================================

# ---------------------------------------------------------------------------
# Logging helper – writes into st.session_state so it survives reruns
# ---------------------------------------------------------------------------
def log_to_frontend(message: str):
    if "conversion_log" not in st.session_state:
        st.session_state.conversion_log = []
    st.session_state.conversion_log.append(message)


# ---------------------------------------------------------------------------
# Extraction helpers (identical logic to the original Flet version)
# ---------------------------------------------------------------------------
def extract_field_value(text, field_label, length_limit=None):
    try:
        start_index = text.index(field_label) + len(field_label)
        end_index = text.index("\n", start_index)
        value = text[start_index:end_index].strip()
        if length_limit:
            value = value[:length_limit]
        log_to_frontend(f"Extracted value for '{field_label}': {value}")
        return value
    except ValueError:
        log_to_frontend(f"Field '{field_label}' not found in text.")
        return None


def extract_requestor_value(text, field_label):
    try:
        start_index = text.index(field_label) + len(field_label)
        end_index = text.index("\n", start_index)
        value = text[start_index:end_index].strip()
        end_paren_index = value.find(")")
        if end_paren_index != -1:
            value = value[: end_paren_index + 1]
        if "Type:" in value:
            value = value.split("Type:")[0].strip()

        log_to_frontend(f"Extracted requestor for '{field_label}': {value}")
        return value
    except ValueError:
        log_to_frontend(f"Field '{field_label}' not found in text.")
        return None


def extract_value_after_line(text, field_label):
    try:
        start_index = text.index(field_label) + len(field_label)
        end_index = text.index("\n", start_index)
        next_line_start = end_index + 1
        next_line_end = text.find("\n", next_line_start)
        if next_line_end == -1:
            next_line_end = len(text)
        value = text[next_line_start:next_line_end].strip()
        log_to_frontend(f"Extracted value after '{field_label}': {value}")
        return value
    except ValueError:
        log_to_frontend(f"Field '{field_label}' not found in text.")
        return None


def extract_data_from_pdf(file_bytes: bytes, filename: str):
    text = ""
    approvers = []
    approval_dates = []
    approver_details = []
    change_developer = ""
    change_implementor = ""
    change_implemented_on = ""
    cab_approval_provided = "False"
    cab_approver = None
    cab_approval_date = None

    # ---- pass 1: developer / implementor from the "Planning" / "Implement" rows
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for pdf_page in pdf.pages:
                for table in pdf_page.extract_tables():
                    for row in table:
                        if len(row) >= 10:
                            type_value = (
                                (" ".join((row[2] or "").split())).lower() if row[2] else None
                            )
                            assigned_to_value = (
                                " ".join((row[9] or "").split()) if row[9] else None
                            )
                            assigned_end_date = row[8].strip() if row[8] else None

                            if type_value and type_value.startswith("planning"):
                                change_developer = assigned_to_value
                            elif type_value and type_value.startswith("implement"):
                                change_implementor = assigned_to_value
                                change_implemented_on = assigned_end_date
    except Exception as e:
        log_to_frontend(f"pdfplumber failed to extract tables from {filename}: {e}")

    # ---- pass 2: full text + approver table
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for pdf_page in pdf.pages:
                page_text = pdf_page.extract_text()
                if page_text:
                    text += page_text + "\n"

                for table in pdf_page.extract_tables():
                    if table and "Approver" in table[0]:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        approved_rows = df[df.iloc[:, 0].str.lower() == "approved"]

                        approvers.extend(approved_rows.iloc[:, 1].tolist())
                        approval_dates.extend(approved_rows.iloc[:, 4].tolist())
                        approver_details.extend(approved_rows.iloc[:, 2].tolist())

                        for idx, detail in enumerate(approved_rows.iloc[:, 2]):
                            normalized_detail = str(detail).replace("\n", "").strip()
                            if "@IT CAB Approvers" in normalized_detail:
                                cab_approval_provided = "Yes"
                                cab_approver = approved_rows.iloc[idx, 1]
                                cab_approval_date = approved_rows.iloc[idx, 4]
                                log_to_frontend(
                                    "CAB Approval keyword found in approver details. "
                                    f"CAB Approver: {cab_approver}, Date: {cab_approval_date}"
                                )
                                break
                        if cab_approver:
                            break
    except Exception as e:
        log_to_frontend(f"pdfplumber failed to extract text from {filename}: {e}")

    ticket_no = extract_field_value(text, "Number:", length_limit=10)
    requestor = extract_requestor_value(text, "Requested by:")
    planned_start_date = extract_field_value(text, "Planned start date:", length_limit=19)
    extract_field_value(text, "Planned end date:")            # extracted for parity/log only
    extract_field_value(text, "Approval obtained on:")        # extracted for parity/log only
    change_type = extract_field_value(text, "Type:")
    change_description = extract_value_after_line(text, "Short description:")
    extract_field_value(text, "Actual start date:")           # extracted for parity/log only
    extract_field_value(text, "Actual end date:")              # extracted for parity/log only

    change_approved_by = ", ".join(approvers)
    change_approved_on = ", ".join(approval_dates)

    attribute_checks = {
        "Approvals Obtained": "Approval: Approved" in text,
        "Segregation of Duty": change_implementor.strip().lower() != change_developer.strip().lower(),
        "Exception Noted": "",
        "CAB Approval Provided (Yes/No)?": cab_approval_provided,
    }

    return (
        ticket_no,
        requestor,
        planned_start_date,
        change_type,
        change_description,
        change_approved_by,
        change_approved_on,
        change_developer,
        change_implementor,
        change_implemented_on,
        attribute_checks,
        cab_approver,
        cab_approval_date,
        cab_approval_provided,
    )


# ---------------------------------------------------------------------------
# Excel report generator (same formatting/formulas as the original)
# ---------------------------------------------------------------------------
def generate_excel_report(data_list) -> bytes:
    df = pd.DataFrame(
        data_list,
        columns=[
            "Sl No", "Change Request Ticket", "Change Description", "Change Requestor",
            "Planned Start Date", "Change Type", "Change Approved By", "Change Approved On",
            "Approvals are obtained for all changes?",
            "Change Implemented By", "Change Implemented On", "Change Developed By",
            "Whether SOD is maintained between the developer and implementor? (Yes/No)",
            "CAB Approval Provided (Yes/No)?",
            "CAB Approver", "CAB Approval Obtained on", "Exception Noted",
        ],
    )

    df["Whether SOD is maintained between the developer and implementor? (Yes/No)"] = df.apply(
        lambda row: "True"
        if str(row["Change Implemented By"]).strip().lower() != str(row["Change Developed By"]).strip().lower()
        else "False",
        axis=1,
    )

    df["Exception Noted"] = df.apply(
        lambda row: "Yes" if row.astype(str).str.contains("False", case=False).any() else "No",
        axis=1,
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Testing Table"
    ws.sheet_view.showGridLines = False

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=21)
    control_cell = ws.cell(row=1, column=1, value="Control Description: Testing & approval is required...")
    control_cell.font = Font(name="Source Sans Pro", size=12, bold=True)

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    attributes = [
        "Attribute 1", "Approvals are obtained for all changes",
        "Attribute 2", "Segregation of duty is ensured between developers and implementors",
        "Attribute 3", "CAB Approval Provided (Yes/No)?",
    ]
    for i in range(0, len(attributes), 2):
        attr_number = ws.cell(row=3 + i // 2, column=1, value=attributes[i])
        attr_number.font = Font(name="Source Sans Pro", bold=True, color="FFFFFF")
        attr_number.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
        attr_number.border = thin_border
        attr_description = ws.cell(row=3 + i // 2, column=2, value=attributes[i + 1])
        attr_description.border = thin_border
        attr_description.font = Font(name="Source Sans Pro")

    start_row = 8

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
        for c_idx, value in enumerate(row, 1):
            if r_idx > start_row and c_idx == 14:
                value = ""
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == start_row:
                cell.font = Font(bold=True, color="FFFFFF", name="Source Sans Pro")
                cell.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
            else:
                cell.font = Font(name="Source Sans Pro")
                cell.border = Border(
                    left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"),
                )
                if c_idx == 14 and r_idx >= start_row + 1:
                    cell.value = f'=IF(O{r_idx}<>"", "Yes", "No")'

    for col in ws.iter_cols(min_row=3, max_row=start_row):
        col_values = [len(str(cell.value)) for cell in col if cell.value is not None]
        if col_values:
            ws.column_dimensions[col[0].column_letter].width = max(col_values) + 2

    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)

    ws.protection.sheet = True
    ws.protection.enable()

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------
def run_cm_tool_2():

    if "conversion_log" not in st.session_state:
        st.session_state.conversion_log = []
    if "excel_bytes" not in st.session_state:
        st.session_state.excel_bytes = None

    st.title("🛠️ Change Management Tool")
    st.caption(
        "Upload Pulse change tickets (PDF), extract key control-testing details, "
        "and generate a formatted Excel report."
    )
    st.divider()

    uploaded_files = st.file_uploader(
        "📤 Upload Tickets (PDF)", type=["pdf"], accept_multiple_files=True
    )

    st.subheader(f"Ticket Count: {len(uploaded_files) if uploaded_files else 0}")

    if uploaded_files:
        st.write("**Pulse Tickets Uploaded**")
        file_df = pd.DataFrame(
            {"#": range(1, len(uploaded_files) + 1), "File Name": [f.name for f in uploaded_files]}
        )
        st.dataframe(file_df, use_container_width=True, hide_index=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        generate_clicked = st.button("⚙️ Generate Report", use_container_width=True)
    with col2:
        download_slot = st.container()
    with col3:
        if st.button("🗑️ Clear Log", use_container_width=True):
            st.session_state.conversion_log = []

    # ------------------------- Generate report -------------------------
    if generate_clicked:
        st.session_state.conversion_log = []
        if not uploaded_files:
            log_to_frontend("No tickets uploaded. Please upload tickets first.")
        else:
            data_list = []
            progress = st.progress(0.0)
            with st.spinner("Processing tickets..."):
                for i, file in enumerate(uploaded_files, start=1):
                    try:
                        file_bytes = file.getvalue()
                        (
                            ticket_no, requestor, planned_start_date, change_type,
                            change_description, change_approved_by, change_approved_on,
                            change_developer, change_implementor, change_implemented_on,
                            attribute_checks, cab_approver, cab_approval_date, cab_approval_provided,
                        ) = extract_data_from_pdf(file_bytes, file.name)

                        data_list.append(
                            [
                                i, ticket_no, change_description, requestor, planned_start_date,
                                change_type, change_approved_by, change_approved_on,
                                attribute_checks["Approvals Obtained"],
                                change_implementor, change_implemented_on, change_developer,
                                attribute_checks["Segregation of Duty"], cab_approval_provided,
                                cab_approver, cab_approval_date, "",
                            ]
                        )
                    except Exception as e:
                        log_to_frontend(f"Failed to process {file.name}: {e}")

                    progress.progress(i / len(uploaded_files))

            if data_list:
                st.session_state.excel_bytes = generate_excel_report(data_list)
                log_to_frontend("Excel report generated successfully: change_management_report.xlsx")
                st.success("✅ Report Generated Successfully!")
            else:
                log_to_frontend("No valid data to generate report.")
                st.session_state.excel_bytes = None

    # ------------------------- Download button -------------------------
    with download_slot:
        if st.session_state.excel_bytes:
            st.download_button(
                label="⬇️ Download Report",
                data=st.session_state.excel_bytes,
                file_name="change_management_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("⬇️ Download Report", disabled=True, use_container_width=True)

    # ------------------------- Conversion log -------------------------
    st.divider()
    st.subheader("📋 Conversion Log")
    log_text = "\n".join(st.session_state.conversion_log) if st.session_state.conversion_log else "No log messages yet."
    st.text_area("Conversion log", value=log_text, height=280, disabled=True, label_visibility="collapsed")

# ============================================================
# MAIN APP LOGIC
# ============================================================
def render_home():
    name = st.session_state.user_name or "there"
    st.markdown(
        f"""
        <div class="portal-hero">
            <h1>{_greeting()}, {name}! 👋</h1>
            <p>Welcome to the GEHC SOX TOOL pick a tool below to get started, or use the sidebar anytime.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("🧰 Tools Available", len(TOOL_INFO))
    with m2:
        login_time_str = (
            st.session_state.login_time.strftime("%I:%M %p")
            if st.session_state.login_time else "—"
        )
        st.metric("🕒 Session Started", login_time_str)
    with m3:
        st.metric("👤 Logged in as", st.session_state.user_name or "—")

    st.write("")
    st.subheader("Choose a tool to launch")

    cols = st.columns(3)
    tool_names = list(TOOL_INFO.keys())
    for col, tname in zip(cols, tool_names):
        info = TOOL_INFO[tname]
        with col:
            st.markdown(
                f"""
                <div class="tool-card">
                    <div class="icon">{info['icon']}</div>
                    <h3>{tname}</h3>
                    <p>{info['description']}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(f"Open {tname} ➜", key=f"open_{tname}", use_container_width=True):
                st.session_state.selected_tool = tname
                st.rerun()

    st.write("")
    st.info(
        "💡 **Tip:** You can switch between tools anytime using the navigation menu in the sidebar "
    )


def render_sidebar_nav():
    name = st.session_state.user_name or "User"
    sso = st.session_state.user_sso or ""
    initials = _get_initials(name)

    st.sidebar.markdown(
        f"""
        <div class="sidebar-user-box">
            <div class="avatar-circle">{initials}</div>
            <div class="info">
                <div class="name">{name}</div>
                <div class="sso">SSO: {sso}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    logout()

    st.sidebar.markdown("---")
    st.sidebar.caption("NAVIGATION")

    nav_options = ["Home"] + list(TOOL_INFO.keys())
    icons_map = {"Home": "🏠", **{k: v["icon"] for k, v in TOOL_INFO.items()}}

    current = st.session_state.selected_tool
    if current not in nav_options:
        current = "Home"

    choice = st.sidebar.radio(
        "Go to",
        nav_options,
        index=nav_options.index(current),
        format_func=lambda x: f"{icons_map.get(x, '')}  {x}",
        label_visibility="collapsed",
    )
    st.session_state.selected_tool = choice

    st.sidebar.markdown("---")
    with st.sidebar.expander("ℹ️ About this portal"):
        st.write(
            "This portal gives SSO-verified users quick access to internal "
            "compliance & change-management tools:\n\n"
            "- ✅ **Saviynt Tool** — SOX access checks\n"
            "- 🎫 **CM Tool 1** — Ticket automation\n"
            "- 🛠️ **CM Tool 2** — Excel report generator"
        )
    st.sidebar.caption("v1.0 · Internal use only")

    return choice


def main():
    if not st.session_state.authenticated:
        login()
        return

    choice = render_sidebar_nav()

    if choice == "Home":
        render_home()
    elif choice == "Saviynt Tool":
        render_page_banner(
            "Saviynt Tool",
            TOOL_INFO["Saviynt Tool"]["tagline"],
            TOOL_INFO["Saviynt Tool"]["icon"],
        )
        run_saviynt_tool()
    elif choice == "CM Tool 1":
        render_page_banner(
            "CM Tool 1",
            TOOL_INFO["CM Tool 1"]["tagline"],
            TOOL_INFO["CM Tool 1"]["icon"],
        )
        run_cm_tool_1()
    elif choice == "CM Tool 2":
        render_page_banner(
            "CM Tool 2",
            TOOL_INFO["CM Tool 2"]["tagline"],
            TOOL_INFO["CM Tool 2"]["icon"],
        )
        run_cm_tool_2()


if __name__ == "__main__":
    main()
