import calendar
import html
import io
import os
import pickle
import re
import threading
import time
from datetime import date, datetime
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION & MOBILE NAVIGATION STATE
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="MI Overdue Works (Mobile)",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

if "mob_active_page" not in st.session_state:
    st.session_state["mob_active_page"] = "dashboard"


def _switch_mob_page(target_page: str):
    st.session_state["mob_active_page"] = target_page


def clean_html(raw_html: str) -> str:
    """Strips leading line indentation so Markdown never renders HTML as a code block."""
    return "\n".join(
        line.strip() for line in raw_html.splitlines() if line.strip()
    )


# -----------------------------------------------------------------------------
# INSTANT CSS KILL-SWITCH FOR INACTIVE PAGE CONTAINER
# (Hides only the inactive page so dropdown changes never flash or jump scroll)
# -----------------------------------------------------------------------------
_active_page = st.session_state["mob_active_page"]
_inactive_key = (
    "mob_page_analytics" if _active_page == "dashboard" else "mob_page_dashboard"
)

st.markdown(
    clean_html(
        f"""
        <style>
            .st-key-{_inactive_key},
            div.element-container:has(.st-key-{_inactive_key}) {{
                display: none !important;
            }}
        </style>
        """
    ),
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# GLOBAL MOBILE-OPTIMIZED DASHBOARD STYLES
# -----------------------------------------------------------------------------
st.markdown(
    clean_html(
        """
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Montserrat:wght@700;800&display=swap');

            .block-container {
                max-width: 100% !important;
                padding-left: 0.65rem !important;
                padding-right: 0.75rem !important;
                padding-top: 2.1rem !important;
                padding-bottom: 1.8rem !important;
            }
            [data-testid="stPlotlyChart"] {
                width: 100% !important;
            }

            /* Sleek Sidebar (220px on Mobile) */
            section[data-testid="stSidebar"][aria-expanded="true"] {
                width: 220px !important;
                min-width: 220px !important;
                max-width: 220px !important;
                background: linear-gradient(180deg, rgba(15, 23, 42, 0.98) 0%, rgba(15, 23, 42, 0.94) 100%) !important;
                border-right: 1px solid rgba(148, 163, 184, 0.16) !important;
            }

            /* Always-Visible Sidebar Expand/Collapse Button */
            [data-testid="stSidebarCollapseButton"],
            [data-testid="collapsedControl"],
            [data-testid="stExpandSidebarButton"] {
                opacity: 1 !important;
                visibility: visible !important;
                display: flex !important;
                z-index: 999995 !important;
            }
            [data-testid="stSidebarCollapseButton"] button,
            [data-testid="collapsedControl"] button,
            [data-testid="stExpandSidebarButton"] {
                opacity: 1 !important;
                visibility: visible !important;
                background: #1E293B !important;
                border: 1px solid rgba(148, 163, 184, 0.45) !important;
                border-radius: 10px !important;
                color: #E2E8F0 !important;
                width: 34px !important;
                height: 34px !important;
                min-height: 34px !important;
            }

            /* Filter Control Cards */
            div[data-testid="stVerticalBlockBorderWrapper"] {
                border-radius: 11px !important;
                padding: 0.35rem 0.5rem !important;
                background: linear-gradient(145deg, rgba(30, 41, 59, 0.52), rgba(15, 23, 42, 0.75)) !important;
                border: 1px solid rgba(148, 163, 184, 0.18) !important;
                box-shadow: 0 3px 10px rgba(0, 0, 0, 0.22) !important;
            }

            /* Mobile Header Banner */
            .mob-header-wrap {
                display: flex;
                align-items: center;
                justify-content: center;
                padding: 0.15rem 0;
                margin-bottom: 0.35rem;
                width: 100%;
            }
            .mob-header-title {
                font-family: 'Montserrat', sans-serif !important;
                font-size: clamp(0.95rem, 4.4vw, 1.45rem) !important;
                font-weight: 800 !important;
                background: linear-gradient(90deg, #FFFFFF 0%, #CBD5E1 100%) !important;
                -webkit-background-clip: text !important;
                -webkit-text-fill-color: transparent !important;
                margin: 0 !important;
                text-align: center !important;
                line-height: 1.25 !important;
            }

            /* Matte Navigation Button */
            div[data-testid="stButton"] > button {
                background: linear-gradient(180deg, rgba(30, 41, 59, 0.92) 0%, rgba(15, 23, 42, 0.98) 100%) !important;
                border: 1px solid rgba(148, 163, 184, 0.38) !important;
                border-top: 1px solid rgba(226, 232, 240, 0.45) !important;
                border-radius: 10px !important;
                color: #F8FAFC !important;
                font-family: 'Montserrat', sans-serif !important;
                font-weight: 700 !important;
                font-size: 0.84rem !important;
                min-height: 38px !important;
                box-shadow: none !important;
            }

            /* Status Bar */
            .mi-status-bar {
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                gap: 0.35rem 0.75rem;
                padding: 0.38rem 0.65rem;
                margin-bottom: 0.6rem;
                border-radius: 9px;
                background: rgba(30, 41, 59, 0.48);
                border: 1px solid rgba(148, 163, 184, 0.14);
                font-size: 0.74rem;
                color: #CBD5E1;
            }
            .mi-status-pill {
                display: inline-flex;
                align-items: center;
                padding: 0.1rem 0.45rem;
                border-radius: 999px;
                background: rgba(59, 130, 246, 0.15);
                border: 1px solid rgba(59, 130, 246, 0.35);
                color: #93C5FD;
                font-weight: 600;
                font-size: 0.72rem;
            }

            /* Mobile Responsive 2x3 KPI Grid */
            .mob-kpi-grid {
                display: grid;
                grid-template-columns: repeat(2, 1fr);
                gap: 8px;
                margin-top: 4px;
                margin-bottom: 10px;
            }
            @media (min-width: 768px) {
                .mob-kpi-grid {
                    grid-template-columns: repeat(6, 1fr);
                    gap: 10px;
                }
            }
            .mob-kpi-card {
                background: linear-gradient(145deg, rgba(30, 41, 59, 0.75), rgba(15, 23, 42, 0.92));
                border: 1px solid rgba(148, 163, 184, 0.18);
                border-radius: 10px;
                padding: 9px 11px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.26);
            }
            .mob-kpi-label {
                font-size: 0.68rem;
                font-weight: 700;
                color: #94A3B8;
                text-transform: uppercase;
                letter-spacing: 0.3px;
                margin-bottom: 2px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .mob-kpi-val {
                font-size: 1.18rem;
                font-weight: 800;
                color: #F8FAFC;
                line-height: 1.15;
            }
            .mob-kpi-sub {
                font-size: 0.66rem;
                font-weight: 600;
                color: #CBD5E1;
                margin-top: 2px;
            }

            /* Mobile Responsive 3x2 Overdue Ageing Tree Grid */
            .mob-age-grid {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 7px;
                margin-top: 4px;
                margin-bottom: 10px;
            }
            @media (min-width: 768px) {
                .mob-age-grid {
                    grid-template-columns: repeat(6, 1fr);
                }
            }
            .mob-age-card {
                background: linear-gradient(145deg, rgba(239, 68, 68, 0.14), rgba(15, 23, 42, 0.88));
                border: 1px solid rgba(239, 68, 68, 0.32);
                border-top: 2.5px solid #EF4444;
                border-radius: 9px;
                padding: 7px 8px;
                text-align: center;
            }
            .mob-age-title {
                font-size: 0.63rem;
                font-weight: 700;
                color: #FCA5A5;
                text-transform: uppercase;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .mob-age-count {
                font-size: 1.05rem;
                font-weight: 800;
                color: #FFFFFF;
                margin-top: 1px;
            }
            .mob-age-pct {
                font-size: 0.62rem;
                font-weight: 700;
                color: #FCA5A5;
            }

            /* Section Header Banners */
            .exec-section-banner {
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                justify-content: space-between;
                gap: 0.35rem 0.75rem;
                padding: 0.5rem 0.75rem;
                margin-top: 0.2rem;
                margin-bottom: 0.55rem;
                border-radius: 9px;
                background: rgba(15, 23, 42, 0.65);
                border: 1px solid rgba(148, 163, 184, 0.16);
                border-left: 4px solid #3B82F6;
            }
            .exec-section-title {
                font-size: 1.02rem;
                font-weight: 700;
                color: #F8FAFC;
            }
            .exec-scope-badge {
                font-size: 0.70rem;
                font-weight: 600;
                padding: 0.18rem 0.55rem;
                border-radius: 999px;
                background: rgba(59, 130, 246, 0.12);
                border: 1px solid rgba(96, 165, 250, 0.3);
                color: #93C5FD;
            }

            /* Touch-Scrollable CE Combined Overdue Works Table */
            .ce-table-wrapper {
                max-height: 500px;
                overflow-y: auto;
                overflow-x: auto;
                -webkit-overflow-scrolling: touch;
                border: 1px solid rgba(148, 163, 184, 0.22);
                border-radius: 10px;
                background: rgba(15, 23, 42, 0.78);
                margin-top: 0.35rem;
                margin-bottom: 0.7rem;
            }
            table.ce-overdue-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 0.78rem;
                color: #E2E8F0;
            }
            table.ce-overdue-table thead th {
                position: sticky;
                top: 0;
                z-index: 10;
                background: #273346;
                color: #F1F5F9;
                font-weight: 700;
                text-align: left;
                padding: 9px 8px;
                border-bottom: 2px solid #3B82F6;
                white-space: nowrap;
                font-size: 0.76rem;
            }
            table.ce-overdue-table tbody tr {
                border-bottom: 1px solid rgba(148, 163, 184, 0.12);
            }
            table.ce-overdue-table tbody tr:nth-child(even) {
                background: rgba(30, 41, 59, 0.32);
            }
            table.ce-overdue-table td {
                padding: 7px 8px;
                vertical-align: middle;
                line-height: 1.3;
            }
            .ce-work-name-cell {
                min-width: 220px;
                max-width: 300px;
                white-space: normal;
            }
            .ce-work-name-clamp {
                display: -webkit-box;
                -webkit-line-clamp: 3;
                -webkit-box-orient: vertical;
                overflow: hidden;
                text-overflow: ellipsis;
                line-height: 1.3;
                font-weight: 600;
                color: #F8FAFC;
            }
            .ce-compact-cell {
                white-space: nowrap;
                font-size: 0.76rem;
            }
            .ce-date-overdue {
                color: #FCA5A5;
                font-weight: 700;
                white-space: nowrap;
            }
            .ce-sl-cell {
                text-align: center;
                font-weight: 600;
                color: #94A3B8;
                width: 36px;
            }

            /* Truth Audit & Leaderboard Panels */
            .truth-audit-box {
                background: linear-gradient(135deg, rgba(30, 41, 59, 0.80), rgba(15, 23, 42, 0.95));
                border: 1px solid rgba(148, 163, 184, 0.25);
                border-left: 4px solid #EF4444;
                border-radius: 10px;
                padding: 12px 13px;
                margin-top: 6px;
                margin-bottom: 12px;
            }
            .exec-lb-panel {
                background: linear-gradient(160deg, rgba(30, 41, 59, 0.65), rgba(15, 23, 42, 0.94));
                border: 1px solid rgba(148, 163, 184, 0.20);
                border-radius: 12px;
                padding: 12px 13px;
                margin-bottom: 12px;
            }
            .exec-lb-panel-danger {
                border-top: 3.5px solid #EF4444;
            }
            .exec-lb-panel-success {
                border-top: 3.5px solid #10B981;
            }
            .exec-lb-header {
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                justify-content: space-between;
                gap: 4px;
                margin-bottom: 10px;
                padding-bottom: 6px;
                border-bottom: 1px solid rgba(148, 163, 184, 0.16);
            }
            .exec-lb-title {
                font-size: 0.90rem;
                font-weight: 800;
                color: #F8FAFC;
            }
            .exec-lb-sub {
                font-size: 0.68rem;
                font-weight: 600;
                color: #94A3B8;
                background: rgba(15, 23, 42, 0.7);
                padding: 2px 7px;
                border-radius: 999px;
            }
            .exec-lb-row {
                background: rgba(15, 23, 42, 0.65);
                border: 1px solid rgba(148, 163, 184, 0.14);
                border-radius: 9px;
                padding: 8px 10px;
                margin-bottom: 7px;
            }
            .exec-lb-row-top {
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                justify-content: space-between;
                gap: 6px;
                margin-bottom: 5px;
            }
            .exec-lb-officer-wrap {
                display: flex;
                align-items: center;
                gap: 6px;
                min-width: 0;
            }
            .exec-lb-rank {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                min-width: 24px;
                height: 21px;
                border-radius: 5px;
                font-size: 0.70rem;
                font-weight: 800;
                background: rgba(59, 130, 246, 0.18);
                color: #93C5FD;
                border: 1px solid rgba(59, 130, 246, 0.35);
            }
            .exec-lb-name {
                font-size: 0.82rem;
                font-weight: 700;
                color: #F8FAFC;
            }
            .exec-lb-div-tag {
                font-size: 0.66rem;
                font-weight: 600;
                color: #94A3B8;
                background: rgba(148, 163, 184, 0.12);
                padding: 1px 6px;
                border-radius: 5px;
            }
            .exec-lb-badges {
                display: flex;
                align-items: center;
                gap: 5px;
            }
            .exec-pill-danger {
                background: rgba(239, 68, 68, 0.18);
                color: #FCA5A5;
                border: 1px solid rgba(239, 68, 68, 0.42);
                padding: 1px 7px;
                border-radius: 999px;
                font-weight: 800;
                font-size: 0.70rem;
            }
            .exec-pill-success {
                background: rgba(16, 185, 129, 0.18);
                color: #6EE7B7;
                border: 1px solid rgba(16, 185, 129, 0.42);
                padding: 1px 7px;
                border-radius: 999px;
                font-weight: 800;
                font-size: 0.70rem;
            }
            .exec-pill-ets {
                background: rgba(245, 158, 11, 0.16);
                color: #FCD34D;
                border: 1px solid rgba(245, 158, 11, 0.35);
                padding: 1px 6px;
                border-radius: 999px;
                font-weight: 700;
                font-size: 0.68rem;
            }
            .exec-pill-total {
                color: #38BDF8;
                font-weight: 700;
                font-size: 0.70rem;
            }
            .exec-lb-bar-bg {
                width: 100%;
                height: 5px;
                background: rgba(148, 163, 184, 0.16);
                border-radius: 999px;
                overflow: hidden;
                display: flex;
            }
            .exec-lb-bar-od {
                background: linear-gradient(90deg, #EF4444, #F87171);
                height: 100%;
            }
            .exec-lb-bar-ets {
                background: linear-gradient(90deg, #10B981, #34D399);
                height: 100%;
            }
        </style>
        """
    ),
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# SAFE TOUCH-ONLY KEYBOARD SUPPRESSION FOR MOBILE SELECTBOXES
# (Uses passive event delegation after mount so Autocomplete.js never crashes)
# -----------------------------------------------------------------------------
def inject_safe_mobile_touch_helpers(show_hud: bool = True):
    components.html(
        f"""
        <script>
        (function() {{
            const parentDoc = window.parent.document;
            const parentWin = window.parent;

            // 1. Prevent virtual keyboard pop-up on touch devices without mutating React initialization
            const isTouchDevice = ('ontouchstart' in parentWin) || (navigator.maxTouchPoints > 0);
            if (isTouchDevice && !parentWin.__miSafeTouchSelectBound) {{
                parentWin.__miSafeTouchSelectBound = true;
                parentDoc.addEventListener("focusin", function(e) {{
                    const t = e.target;
                    if (t && t.tagName === "INPUT" && t.closest('[data-baseweb="select"]')) {{
                        t.setAttribute("inputmode", "none");
                        t.setAttribute("readonly", "readonly");
                    }}
                }}, true);
            }}

            // 2. Right-Edge Scroll HUD Ruler
            if (parentWin.__miScrollHudRAF) {{
                parentWin.cancelAnimationFrame(parentWin.__miScrollHudRAF);
            }}
            const oldHud = parentDoc.getElementById("mi-scroll-hud-root");
            if (oldHud) oldHud.remove();
            const oldStyle = parentDoc.getElementById("mi-scroll-hud-style");
            if (oldStyle) oldStyle.remove();

            if (!{str(show_hud).lower()}) return;

            const style = parentDoc.createElement("style");
            style.id = "mi-scroll-hud-style";
            style.innerHTML = `
                #mi-scroll-hud-root {{
                    position: fixed !important;
                    top: 0 !important;
                    right: 0 !important;
                    width: 12px !important;
                    height: 100vh !important;
                    z-index: 999990 !important;
                    pointer-events: none !important;
                    user-select: none !important;
                    background: transparent !important;
                    font-family: monospace !important;
                }}
                #mi-scroll-hud-indicator {{
                    position: absolute;
                    right: 0;
                    width: 12px;
                    top: 50%;
                    z-index: 6;
                }}
                #mi-scroll-hud-blue-line {{
                    width: 12px;
                    height: 2px;
                    background: #3B82F6;
                }}
                #mi-scroll-hud-pct {{
                    margin-top: 2px;
                    padding-right: 1px;
                    text-align: right;
                    font-size: 7px;
                    font-weight: 700;
                    color: #93C5FD;
                }}
            `;
            parentDoc.head.appendChild(style);

            const hudRoot = parentDoc.createElement("div");
            hudRoot.id = "mi-scroll-hud-root";
            hudRoot.innerHTML = `
                <div id="mi-scroll-hud-indicator">
                    <div id="mi-scroll-hud-blue-line"></div>
                    <div id="mi-scroll-hud-pct">0%</div>
                </div>
            `;
            parentDoc.body.appendChild(hudRoot);

            function getScrollPct() {{
                const el = parentDoc.querySelector('[data-testid="stMain"]') || parentDoc.scrollingElement || parentDoc.documentElement;
                if (!el) return 0;
                const maxS = Math.max(1, el.scrollHeight - el.clientHeight);
                return Math.max(0, Math.min(100, (el.scrollTop / maxS) * 100));
            }}

            function loop() {{
                const root = parentDoc.getElementById("mi-scroll-hud-root");
                const ind = parentDoc.getElementById("mi-scroll-hud-indicator");
                const pct = parentDoc.getElementById("mi-scroll-hud-pct");
                if (root && ind && pct) {{
                    const p = getScrollPct();
                    const y = root.clientHeight * (0.20 + 0.60 * (p / 100.0));
                    ind.style.top = y.toFixed(1) + "px";
                    pct.textContent = Math.round(p) + "%";
                }}
                parentWin.__miScrollHudRAF = parentWin.requestAnimationFrame(loop);
            }}
            parentWin.__miScrollHudRAF = parentWin.requestAnimationFrame(loop);
        }})();
        </script>
        """,
        height=0,
        width=0,
    )


# -----------------------------------------------------------------------------
# CONSTANTS & TARGET SHEETS
# -----------------------------------------------------------------------------
TARGET_SHEETS = [
    "4702_MOT",
    "4702-New Tanks",
    "4702-LIS",
    "4702_A&P",
    "PMKSY-SMI-RRR",
    "4702_Paschima Vahini",
    "4711__Other Flood Control",
    "4702-Kharland",
    "NABARD - New Tank",
    "SDMF",
    "KMERC",
    "NABARD - A&P",
    "ATAL BHUJAL",
    "NABARD_MOT",
    "NABARD - LIS",
    "other dept funds",
    "SCP (LIS)",
    "SCP (Borewell)",
    "SCP (Checkdam)",
    "TSP (Checkdam)",
    "TSP (Borewell)",
    "TSP (LIS)",
    "NABARD - Other Flood Control",
]

COL_WORK_NAME_KN = "Work Name"
COL_DIV_KN = "ವಿಭಾಗ"
COL_SUBDIV_KN = "ಉಪವಿಭಾಗ"
COL_SO_KN = "ಶಾಖಾಧಿಕಾರಿಯ ಹೆಸರು"
COL_AEE_KN = "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್ ಹೆಸರು"
COL_START_DATE_KN = "ಕಾಮಗಾರಿ ಪ್ರಾರಂಭಿಸಿದ ದಿನಾಂಕ"
COL_TARGET_DATE_KN = "ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ"
COL_ACTUAL_DATE_KN = "ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ"
COL_HOA_KN = "ಲೆಕ್ಕಶೀರ್ಷಿಕೆ"

LABEL_DIV_EN = "Division"
LABEL_SUBDIV_EN = "Sub-division"
LABEL_AE_EN = "Assistant Engineer"
LABEL_AEE_EN = "Assistant Executive Engineer"
LABEL_EE_EN = "Executive Engineer"
LABEL_TA_EN = "Technical Assistant"
LABEL_AS_EN = "Account Superintendent"

ALL_OPTION = "ALL"
UNSPECIFIED_LABEL = "Not Specified (Empty)"
UNASSIGNED_LABEL = "Unassigned (Empty)"

BAR_SHRINK_THRESHOLD = 250
BAR_SHRINK_MAX_EXTRA = 18.0


def get_default_sheet_url() -> str:
    """Safely reads SHEET_URL from Render environment variables without requiring secrets.toml."""
    fallback = "https://docs.google.com/spreadsheets/d/13W04z-fdU3RhuVJ5vQFRpK4BBSQRJWfscPqpr7Oj2Y8/edit?gid=270221191#gid=270221191"
    env_url = os.environ.get("SHEET_URL")
    if env_url and env_url.strip():
        return env_url.strip()
    try:
        sec_url = st.secrets.get("SHEET_URL")
        if sec_url and str(sec_url).strip():
            return str(sec_url).strip()
    except Exception:
        pass
    return fallback


DEFAULT_SHEET_URL = get_default_sheet_url()

SCHEMA_VERSION = "v14_render_mobile_clean_2page"
DISK_BACKUP_FILE = "/tmp/mi_overdue_mobile_snapshot_v14.pkl"


# -----------------------------------------------------------------------------
# SHARED SERVER MEMORY & LOCK
# -----------------------------------------------------------------------------
@st.cache_resource
def get_shared_state_mob_v14() -> dict:
    state = {
        "version": SCHEMA_VERSION,
        "lock": threading.Lock(),
        "data": None,
        "last_synced": None,
        "source_label": None,
    }
    if os.path.exists(DISK_BACKUP_FILE):
        try:
            with open(DISK_BACKUP_FILE, "rb") as f:
                saved = pickle.load(f)
                if saved.get("version") == SCHEMA_VERSION:
                    state["data"] = saved.get("data")
                    state["last_synced"] = saved.get("last_synced")
                    state["source_label"] = saved.get("source_label")
        except Exception:
            pass
    return state


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def calc_overdue_duration(target_date_obj, current_date_obj) -> str:
    if not isinstance(target_date_obj, date) or target_date_obj >= current_date_obj:
        return "-"

    y = current_date_obj.year - target_date_obj.year
    m = current_date_obj.month - target_date_obj.month
    d = current_date_obj.day - target_date_obj.day + 1

    if d <= 0:
        m -= 1
        prev_month = current_date_obj.month - 1 if current_date_obj.month > 1 else 12
        prev_year = (
            current_date_obj.year
            if current_date_obj.month > 1
            else current_date_obj.year - 1
        )
        d += calendar.monthrange(prev_year, prev_month)[1]

    if m < 0:
        y -= 1
        m += 12

    days_in_current_month = calendar.monthrange(
        current_date_obj.year, current_date_obj.month
    )[1]
    if d > days_in_current_month:
        d -= days_in_current_month
        m += 1
        if m >= 12:
            m -= 12
            y += 1

    parts = []
    if y > 0:
        parts.append(f"{y} yr{'s' if y > 1 else ''}")
    if m > 0:
        parts.append(f"{m} mo{'s' if m > 1 else ''}")
    if d > 0:
        parts.append(f"{d} d")

    return "(" + ", ".join(parts) + ")" if parts else "(0 d)"


def compress_bar_length(val: float, threshold: int = BAR_SHRINK_THRESHOLD) -> float:
    v = float(val)
    t = float(threshold)
    if v <= t:
        return v
    excess = v - t
    return t + BAR_SHRINK_MAX_EXTRA * (excess / (excess + 120.0))


def to_safe_python_date(val):
    if val is None:
        return None
    if isinstance(val, pd.Timestamp):
        return date(val.year, val.month, val.day)
    if isinstance(val, datetime):
        return date(val.year, val.month, val.day)
    if isinstance(val, date):
        return val
    return val


def sanitize_dataframe_cells(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_rows = [
        [to_safe_python_date(cell) for cell in row]
        for row in df.to_numpy().tolist()
    ]
    return pd.DataFrame(cleaned_rows, dtype=object)


def read_sheet_safely(
    excel_file: pd.ExcelFile, workbook_bytes: bytes, sheet_name: str
) -> pd.DataFrame:
    try:
        raw_df = pd.read_excel(
            excel_file, sheet_name=sheet_name, header=None, dtype=object
        )
        return sanitize_dataframe_cells(raw_df)
    except Exception:
        wb = openpyxl.load_workbook(
            io.BytesIO(workbook_bytes), data_only=True, read_only=True
        )
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            cleaned_row = [to_safe_python_date(cell) for cell in row]
            if any(c is not None and str(c).strip() != "" for c in cleaned_row):
                rows.append(cleaned_row)
            elif len(rows) < 35:
                rows.append(cleaned_row)
        wb.close()
        return pd.DataFrame(rows, dtype=object)


def normalize_text(val) -> str:
    val = to_safe_python_date(val)
    if val is None:
        return ""
    if isinstance(val, date):
        return f"{val.day:02d}-{val.month:02d}-{val.year:04d}"
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(val)).strip()


def format_cell_for_display(val) -> str:
    val = to_safe_python_date(val)
    if val is None:
        return ""
    if isinstance(val, date):
        return f"{val.day:02d}-{val.month:02d}-{val.year:04d}"
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return normalize_text(val)


def extract_spreadsheet_key(sheet_url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise ValueError("Invalid Google Sheet URL.")
    return match.group(1)


def download_xlsx_bytes(sheet_url: str) -> bytes:
    key = extract_spreadsheet_key(sheet_url)
    urls = [
        f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx",
        f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx&id={key}",
    ]
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Linux; Android 14) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/128.0.0.0 Mobile Safari/537.36"
        ),
    }

    last_err = None
    for attempt in range(3):
        url = urls[attempt % len(urls)]
        if attempt > 0:
            url = f"{url}&t={int(time.time())}"
        try:
            resp = requests.get(url, headers=headers, timeout=120)
            resp.raise_for_status()
            if len(resp.content) > 1000:
                return resp.content
        except Exception as e:
            last_err = e
            time.sleep(2 * (attempt + 1))

    raise RuntimeError(f"Could not download workbook after 3 attempts: {last_err}")


def find_matching_column_indices(headers: list[str]) -> dict:
    col_idx_map = {}

    for idx, col in enumerate(headers):
        norm_col = normalize_text(col)
        upper_col = norm_col.upper()
        compact_col = re.sub(r"[\s\-_]+", "", norm_col)

        if norm_col in {"ವಿಭಾಗ", "ವಿಭಾಗದ ಹೆಸರು"} and "division" not in col_idx_map:
            col_idx_map["division"] = idx
        elif (
            compact_col in {"ಉಪವಿಭಾಗ", "ಉಪವಿಭಾಗದಹೆಸರು"}
            and "subdivision" not in col_idx_map
        ):
            col_idx_map["subdivision"] = idx
        elif upper_col == "EXECUTIVE ENGINEER" and "ee" not in col_idx_map:
            col_idx_map["ee"] = idx
        elif upper_col == "TECHNICAL ASSISTANT" and "ta" not in col_idx_map:
            col_idx_map["ta"] = idx
        elif (
            upper_col in {"ACCOUNT SUPERINTENDENT", "ACCOUNTS SUPERINTENDENT"}
            and "as_off" not in col_idx_map
        ):
            col_idx_map["as_off"] = idx
        elif (
            norm_col
            in {
                "ಕಾಮಗಾರಿ ಹೆಸರು",
                "ಕಾಮಗಾರಿಯ ಹೆಸರು",
                "ಕಾಮಗಾರಿಗಳ ಹೆಸರು",
                "WORK NAME",
                "NAME OF WORK",
            }
            and "work_name" not in col_idx_map
        ):
            col_idx_map["work_name"] = idx
        elif (
            compact_col in {"ಲೆಕ್ಕಶೀರ್ಷಿಕೆ", "ಲೆಕ್ಕಶೀರ್ಷಿಕೆಹೆಸರು"}
            and "hoa" not in col_idx_map
        ):
            col_idx_map["hoa"] = idx

    for idx, col in enumerate(headers):
        norm_col = normalize_text(col)
        upper_col = norm_col.upper()
        compact_upper = re.sub(r"\s+", "", upper_col)

        if (
            ("ಉಪವಿಭಾಗ" in norm_col or "ಉಪ ವಿಭಾಗ" in norm_col or "ಉಪ-ವಿಭಾಗ" in norm_col)
            and "subdivision" not in col_idx_map
        ):
            col_idx_map["subdivision"] = idx
        elif (
            "ವಿಭಾಗ" in norm_col
            and "ಉಪ" not in norm_col
            and "SUB" not in upper_col
            and "division" not in col_idx_map
        ):
            col_idx_map["division"] = idx
        elif (COL_SO_KN in norm_col or "ಶಾಖಾಧಿಕಾರಿ" in norm_col) and "so" not in col_idx_map:
            col_idx_map["so"] = idx
        elif (
            COL_AEE_KN in norm_col or "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್" in norm_col
        ) and "aee" not in col_idx_map:
            col_idx_map["aee"] = idx
        elif (
            ("EXECUTIVE ENGINEER" in upper_col and "ASSISTANT" not in upper_col)
            or ("ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್" in norm_col and "ಸಹಾಯಕ" not in norm_col)
        ) and "ee" not in col_idx_map:
            col_idx_map["ee"] = idx
        elif (
            "TECHNICAL ASSISTANT" in upper_col or "ತಾಂತ್ರಿಕ ಸಹಾಯಕ" in norm_col
        ) and "ta" not in col_idx_map:
            col_idx_map["ta"] = idx
        elif (
            "ACCOUNT SUPERINTENDENT" in upper_col
            or "ACCOUNTS SUPERINTENDENT" in upper_col
            or "ಲೆಕ್ಕ ಅಧೀಕ್ಷಕ" in norm_col
        ) and "as_off" not in col_idx_map:
            col_idx_map["as_off"] = idx
        elif (
            COL_TARGET_DATE_KN in norm_col or "ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ" in norm_col
        ) and "target_date" not in col_idx_map:
            col_idx_map["target_date"] = idx
        elif (
            COL_ACTUAL_DATE_KN in norm_col
            and "ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ" not in norm_col
            and "actual_date" not in col_idx_map
        ):
            col_idx_map["actual_date"] = idx
        elif (
            "ಪ್ರಾರಂಭಿಸಿದ ದಿನಾಂಕ" in norm_col
            or "ಪ್ರಾರಂಭವಾದ ದಿನಾಂಕ" in norm_col
            or "ಪ್ರಾರಂಭ ದಿನಾಂಕ" in norm_col
        ) and "start_date" not in col_idx_map:
            col_idx_map["start_date"] = idx
        elif (
            (
                "ಕಾಮಗಾರಿ ಹೆಸರು" in norm_col
                or "ಕಾಮಗಾರಿಯ ಹೆಸರು" in norm_col
                or "ಕಾಮಗಾರಿಗಳ ಹೆಸರು" in norm_col
            )
            and "ದಿನಾಂಕ" not in norm_col
            and "work_name" not in col_idx_map
        ):
            col_idx_map["work_name"] = idx
        elif (
            ("ಲೆಕ್ಕಶೀರ್ಷಿಕೆ" in norm_col or "ಲೆಕ್ಕ ಶೀರ್ಷಿಕೆ" in norm_col)
            and "hoa" not in col_idx_map
        ):
            col_idx_map["hoa"] = idx
        elif "CP/CF" in compact_upper and "status" not in col_idx_map:
            col_idx_map["status"] = idx

    if "work_name" not in col_idx_map:
        for idx, col in enumerate(headers):
            norm_col = normalize_text(col)
            upper_col = norm_col.upper()
            if (
                (
                    "ಕಾಮಗಾರಿ" in norm_col
                    or "WORK NAME" in upper_col
                    or "NAME OF WORK" in upper_col
                )
                and "ದಿನಾಂಕ" not in norm_col
                and "ಮೊತ್ತ" not in norm_col
                and "ಸಂಖ್ಯೆ" not in norm_col
                and "ಹಂತ" not in norm_col
                and idx not in col_idx_map.values()
            ):
                col_idx_map["work_name"] = idx
                break

    return col_idx_map


def parse_indian_date_list(series: pd.Series) -> list:
    parsed_dates = []
    for raw_val in series.tolist():
        val = to_safe_python_date(raw_val)
        if val is None:
            parsed_dates.append(None)
            continue

        if isinstance(val, date):
            parsed_dates.append(val)
            continue

        s = normalize_text(val)
        if not s or s.upper() in {"-", "--", "---", "NIL", "NA", "N/A", "NONE", ""}:
            parsed_dates.append(None)
            continue

        s = re.sub(r"[./]", "-", s).strip(" -")
        match = re.search(r"(\d{1,4})-(\d{1,2})-(\d{1,4})", s)
        if match:
            p1, p2, p3 = int(match.group(1)), int(match.group(2)), int(match.group(3))
            try:
                if p1 > 31:
                    parsed_dates.append(date(max(1, min(9999, p1)), p2, p3))
                else:
                    yr = p3 + 2000 if p3 < 100 else p3
                    parsed_dates.append(date(max(1, min(9999, yr)), p2, p1))
                continue
            except Exception:
                pass

        try:
            dt = pd.to_datetime(s, dayfirst=True, utc=True, errors="coerce")
            if pd.isna(dt):
                parsed_dates.append(None)
            else:
                parsed_dates.append(date(dt.year, dt.month, dt.day))
        except Exception:
            parsed_dates.append(None)

    return parsed_dates


def is_empty_date(series: pd.Series) -> pd.Series:
    results = []
    for raw_val in series.tolist():
        val = to_safe_python_date(raw_val)
        if val is None:
            results.append(True)
        elif isinstance(val, date):
            results.append(False)
        else:
            s = normalize_text(val)
            results.append(
                s.upper() in {"", "-", "--", "---", "NIL", "NA", "N/A", "NONE", "NAT"}
            )
    return pd.Series(results, index=series.index, dtype=bool)


def is_valid_division_text(val: str) -> bool:
    if not val:
        return False
    upper = val.upper().strip()
    if upper in {"-", "--", "---", "NIL", "NA", "N/A", "NONE", "NAN", "NAT", "NULL", "0"}:
        return False
    if re.fullmatch(r"[\d\W_]+", val):
        return False
    if val in {COL_DIV_KN, "ವಿಭಾಗದ ಹೆಸರು", "ಒಟ್ಟು", "ಒಟ್ಟು ಮೊತ್ತ", "ಮಹಾ ಒಟ್ಟು"}:
        return False
    if upper in {"DIVISION", "DIVISION NAME", "TOTAL", "GRAND TOTAL", "SUB TOTAL"}:
        return False
    if val.startswith("ಒಟ್ಟು ") or val.endswith(" ಒಟ್ಟು") or upper.startswith("TOTAL "):
        return False
    return True


def map_optional_field(name: str) -> str:
    if not name:
        return UNSPECIFIED_LABEL
    upper = name.upper().strip()
    empty_tokens = {
        "NAN", "NONE", "-", "--", "---", "NIL", "NA", "N/A", "0", "ಇಲ್ಲ", "ಖಾಲಿ"
    }
    if upper in empty_tokens or re.fullmatch(r"[\d\W_]+", name):
        return UNSPECIFIED_LABEL
    if (
        len(name) > 75
        or "ಒಟ್ಟು" in name
        or "TOTAL" in upper
        or name in {COL_SUBDIV_KN, "ಉಪ ವಿಭಾಗ", "ಉಪವಿಭಾಗದ ಹೆಸರು"}
        or upper
        in {
            "EXECUTIVE ENGINEER",
            "TECHNICAL ASSISTANT",
            "ACCOUNT SUPERINTENDENT",
            "ACCOUNTS SUPERINTENDENT",
        }
    ):
        return UNSPECIFIED_LABEL
    return name


def map_officer_or_unassigned(name: str) -> str:
    if not name:
        return UNASSIGNED_LABEL
    upper = name.upper()
    empty_tokens = {
        "NAN", "NONE", "-", "--", "---", "NIL", "NA", "N/A", "0", "ಇಲ್ಲ", "ಖಾಲಿ"
    }
    if upper in empty_tokens or name.isdigit():
        return UNASSIGNED_LABEL
    if (
        len(name) > 65
        or "ಒಟ್ಟು" in name
        or "TOTAL" in upper
        or COL_SO_KN in name
        or COL_AEE_KN in name
    ):
        return UNASSIGNED_LABEL
    return name


def shorten_label(name: str, max_len: int = 18) -> str:
    if name == UNASSIGNED_LABEL:
        return "⚠️ Unassigned"
    return name if len(name) <= max_len else name[: max_len - 1] + "…"


def display_clean_cell(val: str) -> str:
    if not val or val in {UNSPECIFIED_LABEL, UNASSIGNED_LABEL}:
        return "-"
    return val


def render_ce_overdue_html_table(df_display: pd.DataFrame) -> str:
    headers_html = "".join(
        f"<th>{html.escape(str(col))}</th>" for col in df_display.columns
    )

    rows_html = []
    for idx_row, (_, row) in enumerate(df_display.iterrows(), start=1):
        cells_html = [f'<td class="ce-sl-cell">{idx_row}</td>']
        for col in df_display.columns[1:]:
            raw_text = str(row[col]) if pd.notna(row[col]) else "-"
            escaped_text = html.escape(raw_text if raw_text else "-")

            if col == "Work Name":
                cells_html.append(
                    f'<td class="ce-work-name-cell" title="{escaped_text}">'
                    f'<div class="ce-work-name-clamp">{escaped_text}</div>'
                    f"</td>"
                )
            elif col in ["Target Date", "Overdue By"]:
                cells_html.append(
                    f'<td class="ce-date-overdue">{escaped_text}</td>'
                )
            else:
                cells_html.append(
                    f'<td class="ce-compact-cell">{escaped_text}</td>'
                )
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

    return clean_html(
        f"""
        <div class="ce-table-wrapper">
            <table class="ce-overdue-table">
                <thead><tr>{headers_html}</tr></thead>
                <tbody>{''.join(rows_html)}</tbody>
            </table>
        </div>
        """
    )


def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Overdue_Works")
    return output.getvalue()


def _fast_age_bucket_from_date(t_dt, today_dt: date) -> str:
    if not isinstance(t_dt, date) or t_dt >= today_dt:
        return "Unknown"
    days_od = (today_dt - t_dt).days
    if days_od <= 180:
        return "1. < 6 Months"
    if days_od <= 365:
        return "2. 6 – 12 Months"
    if days_od <= 730:
        return "3. 1 – 2 Years"
    if days_od <= 1825:
        return "4. 2 – 5 Years"
    return "5. > 5 Years (Chronic)"


def _fast_tree_bucket_from_date(t_dt, today_dt: date) -> str:
    if not isinstance(t_dt, date) or t_dt >= today_dt:
        return "Unknown"
    days = (today_dt - t_dt).days
    if days <= 180:
        return "Under 6 Months"
    if days <= 365:
        return "6 Months to 1 Year"
    if days <= 730:
        return "1 - 2 Years"
    if days <= 1095:
        return "2 - 3 Years"
    if days <= 1460:
        return "3 - 4 Years"
    return "4+ Years"


def ensure_precomputed_age_columns(
    df_master: pd.DataFrame, today_dt: date
) -> pd.DataFrame:
    if df_master.empty:
        return df_master
    if "age_bucket" in df_master.columns and "tree_bucket" in df_master.columns:
        return df_master

    parsed_dates = parse_indian_date_list(df_master["target_date_raw"])
    df_master["age_bucket"] = [
        _fast_age_bucket_from_date(d, today_dt) for d in parsed_dates
    ]
    df_master["tree_bucket"] = [
        _fast_tree_bucket_from_date(d, today_dt) for d in parsed_dates
    ]
    return df_master


def parse_xlsx_workbook(workbook_bytes: bytes):
    progress = st.progress(0, text="Opening workbook...")

    try:
        excel_file = pd.ExcelFile(io.BytesIO(workbook_bytes), engine="calamine")
    except Exception:
        excel_file = pd.ExcelFile(io.BytesIO(workbook_bytes), engine="openpyxl")

    available_sheets = {
        normalize_text(name).lower(): name for name in excel_file.sheet_names
    }

    today_ist_date = datetime.now(ZoneInfo("Asia/Kolkata")).date()

    compiled_rows = []
    full_overdue_rows = []
    sheet_diagnostics = []

    total_targets = len(TARGET_SHEETS)
    for idx_sheet, target in enumerate(TARGET_SHEETS, start=1):
        progress.progress(
            idx_sheet / total_targets,
            text=f"Loading {idx_sheet}/{total_targets}: {target}...",
        )

        norm_target = normalize_text(target).lower()
        actual_sheet_name = available_sheets.get(norm_target)

        if not actual_sheet_name:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Missing in Workbook",
                    "Rows Loaded": 0,
                    "Overdue ('P')": 0,
                    "ETS": 0,
                }
            )
            continue

        try:
            raw_df = read_sheet_safely(excel_file, workbook_bytes, actual_sheet_name)
            if raw_df.empty:
                continue

            header_idx = None
            for idx in range(min(30, len(raw_df))):
                row_text = " | ".join(
                    normalize_text(v) for v in raw_df.iloc[idx].tolist()
                )
                if (
                    COL_SO_KN in row_text
                    or COL_AEE_KN in row_text
                    or "ಶಾಖಾಧಿಕಾರಿ" in row_text
                ):
                    header_idx = idx
                    break

            if header_idx is None:
                continue

            raw_headers = [
                normalize_text(col) for col in raw_df.iloc[header_idx].tolist()
            ]
            col_idx_map = find_matching_column_indices(raw_headers)

            required_keys = {
                "division",
                "so",
                "aee",
                "target_date",
                "actual_date",
                "status",
            }
            optional_keys = {
                "subdivision",
                "ee",
                "ta",
                "as_off",
                "work_name",
                "start_date",
                "hoa",
            }

            if (
                (required_keys - set(col_idx_map.keys()))
                or (optional_keys - set(col_idx_map.keys()))
            ) and (header_idx + 1 < len(raw_df)):
                next_row = [
                    normalize_text(col) for col in raw_df.iloc[header_idx + 1].tolist()
                ]
                combined_headers = [
                    f"{h1} {h2}".strip() for h1, h2 in zip(raw_headers, next_row)
                ]
                alt_map = find_matching_column_indices(combined_headers)
                for k, v in alt_map.items():
                    if k not in col_idx_map:
                        col_idx_map[k] = v

            if header_idx > 0:
                prev_row = [
                    normalize_text(col) for col in raw_df.iloc[header_idx - 1].tolist()
                ]
                combined_prev = [
                    f"{h1} {h2}".strip() for h1, h2 in zip(prev_row, raw_headers)
                ]
                alt_prev_map = find_matching_column_indices(combined_prev)
                for k, v in alt_prev_map.items():
                    if k not in col_idx_map:
                        col_idx_map[k] = v

            if required_keys - set(col_idx_map.keys()):
                continue

            df = raw_df.iloc[header_idx + 1 :].reset_index(drop=True).copy()

            def _extract_col_series(key: str, default_val: str = "") -> pd.Series:
                if key in col_idx_map:
                    return pd.Series(
                        [
                            format_cell_for_display(v)
                            for v in df.iloc[:, col_idx_map[key]].tolist()
                        ],
                        index=df.index,
                        dtype=str,
                    )
                return pd.Series([default_val] * len(df), index=df.index, dtype=str)

            s_div_raw = _extract_col_series("division")
            s_subdiv_raw = _extract_col_series("subdivision")
            s_so_raw = _extract_col_series("so")
            s_aee_raw = _extract_col_series("aee")
            s_ee_raw = _extract_col_series("ee")
            s_ta_raw = _extract_col_series("ta")
            s_as_raw = _extract_col_series("as_off")
            s_work_name = _extract_col_series("work_name")
            s_start_date = _extract_col_series("start_date")
            s_hoa_raw = _extract_col_series("hoa", default_val=target)
            s_hoa = s_hoa_raw.map(lambda x: x if x != "" else target)

            s_status = _extract_col_series("status")
            s_status_upper = s_status.str.upper().str.strip()
            s_target_raw = df.iloc[:, col_idx_map["target_date"]]
            s_actual_raw = df.iloc[:, col_idx_map["actual_date"]]

            target_date_list = parse_indian_date_list(s_target_raw)
            actual_empty = is_empty_date(s_actual_raw)

            duration_list = [
                calc_overdue_duration(td, today_ist_date) for td in target_date_list
            ]
            s_overdue_duration = pd.Series(duration_list, index=df.index, dtype=str)

            age_bucket_list = [
                _fast_age_bucket_from_date(td, today_ist_date)
                for td in target_date_list
            ]
            tree_bucket_list = [
                _fast_tree_bucket_from_date(td, today_ist_date)
                for td in target_date_list
            ]

            is_status_p = s_status_upper == "P"
            is_status_ets = s_status_upper == "ETS"

            target_is_before_today = pd.Series(
                [
                    (d is not None) and (d < today_ist_date)
                    for d in target_date_list
                ],
                index=df.index,
                dtype=bool,
            )

            is_overdue = is_status_p & target_is_before_today & actual_empty
            valid_row_mask = s_div_raw.map(is_valid_division_text)

            s_subdiv_mapped = s_subdiv_raw.map(map_optional_field)
            s_ee_mapped = s_ee_raw.map(map_optional_field)
            s_ta_mapped = s_ta_raw.map(map_optional_field)
            s_as_mapped = s_as_raw.map(map_optional_field)
            s_so_mapped = s_so_raw.map(map_officer_or_unassigned)
            s_aee_mapped = s_aee_raw.map(map_officer_or_unassigned)
            s_target_Formatted = pd.Series(
                [format_cell_for_display(v) for v in s_target_raw.tolist()],
                index=df.index,
                dtype=str,
            )

            sub = pd.DataFrame(
                {
                    "sheet_name": target,
                    "ಕಾಮಗಾರಿ ಹೆಸರು": s_work_name,
                    "Work Name": s_work_name,
                    "division_name": s_div_raw,
                    "subdivision_name": s_subdiv_mapped,
                    "ee_name": s_ee_mapped,
                    "ta_name": s_ta_mapped,
                    "as_name": s_as_mapped,
                    "so_name": s_so_mapped,
                    "aee_name": s_aee_mapped,
                    "status_raw": s_status,
                    "target_date_raw": s_target_Formatted,
                    "Target Date": s_target_Formatted,
                    "Overdue By": s_overdue_duration,
                    "age_bucket": age_bucket_list,
                    "tree_bucket": tree_bucket_list,
                    "H.O.A.": s_hoa,
                    "actual_empty": actual_empty,
                    "is_status_p": is_status_p,
                    "is_status_ets": is_status_ets,
                    "is_overdue": is_overdue,
                }
            )
            sub = sub[valid_row_mask].reset_index(drop=True)
            compiled_rows.append(sub)

            overdue_mask = valid_row_mask & is_overdue
            if overdue_mask.any():
                overdue_curated = pd.DataFrame(
                    {
                        COL_WORK_NAME_KN: s_work_name.loc[overdue_mask].values,
                        COL_DIV_KN: s_div_raw.loc[overdue_mask].values,
                        COL_SUBDIV_KN: [
                            display_clean_cell(v)
                            for v in s_subdiv_mapped.loc[overdue_mask].values
                        ],
                        COL_SO_KN: [
                            display_clean_cell(v)
                            for v in s_so_mapped.loc[overdue_mask].values
                        ],
                        COL_AEE_KN: [
                            display_clean_cell(v)
                            for v in s_aee_mapped.loc[overdue_mask].values
                        ],
                        LABEL_TA_EN: [
                            display_clean_cell(v)
                            for v in s_ta_mapped.loc[overdue_mask].values
                        ],
                        LABEL_AS_EN: [
                            display_clean_cell(v)
                            for v in s_as_mapped.loc[overdue_mask].values
                        ],
                        LABEL_EE_EN: [
                            display_clean_cell(v)
                            for v in s_ee_mapped.loc[overdue_mask].values
                        ],
                        COL_START_DATE_KN: s_start_date.loc[overdue_mask].values,
                        COL_TARGET_DATE_KN: s_target_Formatted.loc[overdue_mask].values,
                        "Overdue By": s_overdue_duration.loc[overdue_mask].values,
                        COL_HOA_KN: s_hoa.loc[overdue_mask].values,
                        "_division_name": s_div_raw.loc[overdue_mask].values,
                        "_subdivision_name": s_subdiv_mapped.loc[overdue_mask].values,
                        "_ee_name": s_ee_mapped.loc[overdue_mask].values,
                        "_ta_name": s_ta_mapped.loc[overdue_mask].values,
                        "_as_name": s_as_mapped.loc[overdue_mask].values,
                        "_so_name": s_so_mapped.loc[overdue_mask].values,
                        "_aee_name": s_aee_mapped.loc[overdue_mask].values,
                    }
                )
                full_overdue_rows.append(overdue_curated.reset_index(drop=True))

            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Loaded",
                    "Rows Loaded": len(sub),
                    "Overdue ('P')": int(overdue_mask.sum()),
                    "ETS": int((valid_row_mask & is_status_ets).sum()),
                }
            )
        except Exception as sheet_err:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": f"Error: {sheet_err}",
                    "Rows Loaded": 0,
                    "Overdue ('P')": 0,
                    "ETS": 0,
                }
            )

    progress.empty()

    master_df = (
        pd.concat(compiled_rows, ignore_index=True)
        if compiled_rows
        else pd.DataFrame()
    )
    overdue_full_df = (
        pd.concat(full_overdue_rows, ignore_index=True).fillna("")
        if full_overdue_rows
        else pd.DataFrame()
    )
    diag_df = pd.DataFrame(sheet_diagnostics)
    return master_df, overdue_full_df, diag_df, today_ist_date.strftime("%d-%m-%Y")


def save_shared_state(state: dict, data_tuple, source_label: str):
    now_str = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %I:%M %p")
    state["version"] = SCHEMA_VERSION
    state["data"] = data_tuple
    state["last_synced"] = now_str
    state["source_label"] = source_label
    try:
        with open(DISK_BACKUP_FILE, "wb") as f:
            pickle.dump(
                {
                    "version": SCHEMA_VERSION,
                    "data": data_tuple,
                    "last_synced": now_str,
                    "source_label": source_label,
                },
                f,
            )
    except Exception:
        pass


def build_dropdown_options(df: pd.DataFrame, col_name: str) -> list[str]:
    raw_vals = [
        v
        for v in df[col_name].dropna().unique().tolist()
        if v != "" and v != UNSPECIFIED_LABEL
    ]
    options = [ALL_OPTION] + sorted(raw_vals)
    if (df[col_name] == UNSPECIFIED_LABEL).any():
        options.append(UNSPECIFIED_LABEL)
    return options


# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Dashboard Controls")

sheet_url_input = st.sidebar.text_input(
    "Google Spreadsheet Link",
    value=DEFAULT_SHEET_URL,
)

uploaded_xlsx = st.sidebar.file_uploader(
    "Or Upload Workbook (.xlsx)",
    type=["xlsx"],
)

sync_clicked = st.sidebar.button(
    "🔄 Sync Latest Data",
    use_container_width=True,
    type="primary",
)

denominator_mode = st.sidebar.radio(
    "Total Works Count Basis:",
    options=[
        "All Works Assigned (All Statuses)",
        "Only Ongoing Works (Status = 'Progress')",
    ],
    index=0,
)

chart_style = st.sidebar.radio(
    "Bar Width Mode:",
    options=[
        "Overlay (Overdue + ETS + Rest of Works)",
        "Scale by Overdue Works (Full-Width Bars)",
        "Scale by Overdue Percentage (%)",
    ],
    index=0,
)

shrink_outliers = st.sidebar.checkbox(
    f"Shrink long bars above {BAR_SHRINK_THRESHOLD} works",
    value=True,
)

hide_zero_overdue = st.sidebar.checkbox(
    "Show only officers with Overdue > 0",
    value=True,
)

show_unassigned_bar = st.sidebar.checkbox(
    "Include Unassigned Officer bar",
    value=True,
)

show_scroll_hud = st.sidebar.checkbox(
    "Show Right-Edge Scroll Ruler",
    value=True,
)

# -----------------------------------------------------------------------------
# MOBILE HEADER & PAGE SWITCHER
# -----------------------------------------------------------------------------
if st.session_state["mob_active_page"] == "dashboard":
    st.markdown(
        clean_html(
            """
            <div class="mob-header-wrap">
                <h1 class="mob-header-title">Minor Irrigation — Overdue Dashboard</h1>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )
    st.button(
        "📊 Switch to CE Analytics & ETS Truth Audit ➔",
        use_container_width=True,
        key="mob_nav_to_analytics",
        on_click=_switch_mob_page,
        args=("analytics",),
    )
else:
    st.markdown(
        clean_html(
            """
            <div class="mob-header-wrap">
                <h1 class="mob-header-title">MI — CE Analytics &amp; ETS Truth Audit</h1>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )
    st.button(
        "⬅️ Back to Main Overdue Works Dashboard",
        use_container_width=True,
        key="mob_nav_to_dashboard",
        on_click=_switch_mob_page,
        args=("dashboard",),
    )

# -----------------------------------------------------------------------------
# THREAD-SAFE DATA LOADING
# -----------------------------------------------------------------------------
shared_state = get_shared_state_mob_v14()

if uploaded_xlsx is not None:
    file_id = f"{uploaded_xlsx.name}_{uploaded_xlsx.size}_{SCHEMA_VERSION}"
    if st.session_state.get("_last_uploaded_id") != file_id:
        with shared_state["lock"]:
            parsed = parse_xlsx_workbook(uploaded_xlsx.getvalue())
            save_shared_state(
                shared_state, parsed, f"Uploaded ({uploaded_xlsx.name})"
            )
            st.session_state["_last_uploaded_id"] = file_id

elif (
    shared_state["data"] is None
    or shared_state.get("version") != SCHEMA_VERSION
    or sync_clicked
):
    with shared_state["lock"]:
        if (
            shared_state["data"] is None
            or shared_state.get("version") != SCHEMA_VERSION
            or sync_clicked
        ):
            with st.spinner("Syncing live snapshot from Google Sheets (~15s)..."):
                try:
                    xlsx_bytes = download_xlsx_bytes(sheet_url_input)
                    parsed = parse_xlsx_workbook(xlsx_bytes)
                    save_shared_state(shared_state, parsed, "Google Sheets Live")
                except Exception as e:
                    if shared_state["data"] is not None:
                        st.warning(
                            f"Sheet busy ({e}). Using snapshot from {shared_state['last_synced']}."
                        )
                    else:
                        st.error(f"Could not load Google Sheet: {e}")
                        st.stop()

master_df, overdue_full_df, diag_df, today_str = shared_state["data"]

today_ist_obj = datetime.now(ZoneInfo("Asia/Kolkata")).date()
if not master_df.empty and (
    "age_bucket" not in master_df.columns
    or "tree_bucket" not in master_df.columns
):
    master_df = ensure_precomputed_age_columns(master_df, today_ist_obj)
    shared_state["data"] = (master_df, overdue_full_df, diag_df, today_str)

st.markdown(
    clean_html(
        f"""
        <div class="mi-status-bar">
            <span class="mi-status-pill">⚡ {shared_state['source_label']}</span>
            <span>🕒 <b>Synced:</b> {shared_state['last_synced']} IST</span>
        </div>
        """
    ),
    unsafe_allow_html=True,
)

if master_df.empty:
    st.warning("No matching data rows found across the 23 target worksheets.")
    st.stop()


# =============================================================================
# ANALYTICS HELPER FUNCTIONS
# =============================================================================
def build_officer_analytics_df(
    df_scope: pd.DataFrame, col_key: str, label_name: str
) -> pd.DataFrame:
    sub = df_scope[
        (df_scope[col_key] != "")
        & (df_scope[col_key] != UNASSIGNED_LABEL)
        & (df_scope[col_key] != UNSPECIFIED_LABEL)
    ].copy()
    if sub.empty:
        return pd.DataFrame()

    grp = (
        sub.groupby(col_key, as_index=False)
        .agg(
            total_works=("is_overdue", "count"),
            ongoing_p_works=("is_status_p", "sum"),
            ets_works=("is_status_ets", "sum"),
            overdue_works=("is_overdue", "sum"),
            primary_div=(
                "division_name",
                lambda s: s.mode().iloc[0] if not s.mode().empty else "-",
            ),
        )
        .rename(columns={col_key: label_name})
    )

    grp["on_track_p"] = (grp["ongoing_p_works"] - grp["overdue_works"]).clip(lower=0)
    grp["rest_works"] = (
        grp["total_works"] - grp["overdue_works"] - grp["ets_works"]
    ).clip(lower=0)
    grp["overdue_pct"] = (
        (grp["overdue_works"] / grp["total_works"]) * 100
    ).round(1).fillna(0.0)
    grp["ets_pct"] = (
        (grp["ets_works"] / grp["total_works"]) * 100
    ).round(1).fillna(0.0)
    grp["delayed_exposure"] = grp["overdue_works"] + grp["ets_works"]
    grp["unexcused_ratio_pct"] = (
        (grp["overdue_works"] / grp["delayed_exposure"].replace(0, pd.NA)) * 100
    ).fillna(0.0).round(1)
    return grp


def render_executive_leaderboard_panel(
    df_slice: pd.DataFrame,
    officer_col: str,
    title_text: str,
    subtitle_text: str,
    is_danger: bool = True,
    primary_metric: str = "count",
) -> str:
    panel_cls = "exec-lb-panel-danger" if is_danger else "exec-lb-panel-success"
    pill_cls = "exec-pill-danger" if is_danger else "exec-pill-success"

    if df_slice.empty:
        return clean_html(
            f"""
            <div class="exec-lb-panel {panel_cls}">
                <div class="exec-lb-header">
                    <span class="exec-lb-title">{html.escape(title_text)}</span>
                    <span class="exec-lb-sub">{html.escape(subtitle_text)}</span>
                </div>
                <div style="color:#94A3B8;font-size:0.80rem;padding:8px 0;">
                    No officers match the current filter criteria.
                </div>
            </div>
            """
        )

    max_od = max(int(df_slice["overdue_works"].max()), 1)
    rows_html = []

    for rank_idx, (_, r) in enumerate(df_slice.iterrows(), start=1):
        off_name = html.escape(str(r[officer_col]))
        div_name = html.escape(str(r.get("primary_div", "-")))
        od_cnt = int(r["overdue_works"])
        od_pct = float(r["overdue_pct"])
        ets_cnt = int(r["ets_works"])
        tot_cnt = max(int(r["total_works"]), 1)

        if primary_metric == "pct":
            main_badge = f'<span class="{pill_cls}">{od_pct:.1f}% ({od_cnt} OD)</span>'
            od_bar_w = min(100.0, max(2.0 if od_cnt > 0 else 0.0, od_pct))
        else:
            main_badge = f'<span class="{pill_cls}">{od_cnt} OD ({od_pct:.1f}%)</span>'
            od_bar_w = (
                min(100.0, max(4.0, (od_cnt / max_od) * 100.0))
                if is_danger and od_cnt > 0
                else min(100.0, (od_cnt / tot_cnt) * 100.0)
            )

        ets_bar_w = min(100.0 - od_bar_w, (ets_cnt / tot_cnt) * 100.0)
        ets_badge = (
            f'<span class="exec-pill-ets">{ets_cnt} ETS</span>'
            if ets_cnt > 0
            else '<span style="color:#64748B;font-size:0.68rem;">0 ETS</span>'
        )

        rows_html.append(
            f"""
            <div class="exec-lb-row">
                <div class="exec-lb-row-top">
                    <div class="exec-lb-officer-wrap">
                        <span class="exec-lb-rank">#{rank_idx}</span>
                        <span class="exec-lb-name">{off_name}</span>
                        <span class="exec-lb-div-tag">🏛️ {div_name}</span>
                    </div>
                    <div class="exec-lb-badges">
                        {main_badge}
                        {ets_badge}
                        <span class="exec-pill-total">{tot_cnt} Tot</span>
                    </div>
                </div>
                <div class="exec-lb-bar-bg">
                    <div class="exec-lb-bar-od" style="width:{od_bar_w:.1f}%;"></div>
                    <div class="exec-lb-bar-ets" style="width:{ets_bar_w:.1f}%;"></div>
                </div>
            </div>
            """
        )

    return clean_html(
        f"""
        <div class="exec-lb-panel {panel_cls}">
            <div class="exec-lb-header">
                <span class="exec-lb-title">{html.escape(title_text)}</span>
                <span class="exec-lb-sub">{html.escape(subtitle_text)}</span>
            </div>
            {''.join(rows_html)}
        </div>
        """
    )


# =============================================================================
# PAGE 2 VIEW: CE EXECUTIVE ANALYTICS & ETS TRUTH AUDIT SUITE (MOBILE)
# =============================================================================
if st.session_state["mob_active_page"] == "analytics":
    with st.container(key="mob_page_analytics"):
        af_c1, af_c2, af_c3, af_c4, af_c5 = st.columns(5, gap="small")

        div_opts_an = [ALL_OPTION] + sorted(
            [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
        )
        with af_c1:
            with st.container(border=True):
                an_div = st.selectbox(
                    f"🏛️ {LABEL_DIV_EN}:",
                    options=div_opts_an,
                    index=0,
                    key="an_f_div",
                )

        an_s1 = (
            master_df[master_df["division_name"] == an_div].copy()
            if an_div != ALL_OPTION
            else master_df.copy()
        )

        subdiv_opts_an = build_dropdown_options(an_s1, "subdivision_name")
        with af_c2:
            with st.container(border=True):
                an_subdiv = st.selectbox(
                    f"🏢 {LABEL_SUBDIV_EN}:",
                    options=subdiv_opts_an,
                    index=0,
                    key="an_f_subdiv",
                )

        an_s2 = (
            an_s1[an_s1["subdivision_name"] == an_subdiv].copy()
            if an_subdiv != ALL_OPTION
            else an_s1.copy()
        )

        ee_opts_an = build_dropdown_options(an_s2, "ee_name")
        with af_c3:
            with st.container(border=True):
                an_ee = st.selectbox(
                    f"👔 {LABEL_EE_EN}:",
                    options=ee_opts_an,
                    index=0,
                    key="an_f_ee",
                )

        an_s3 = (
            an_s2[an_s2["ee_name"] == an_ee].copy()
            if an_ee != ALL_OPTION
            else an_s2.copy()
        )

        ta_opts_an = build_dropdown_options(an_s3, "ta_name")
        with af_c4:
            with st.container(border=True):
                an_ta = st.selectbox(
                    f"🛠️ {LABEL_TA_EN}:",
                    options=ta_opts_an,
                    index=0,
                    key="an_f_ta",
                )

        an_s4 = (
            an_s3[an_s3["ta_name"] == an_ta].copy()
            if an_ta != ALL_OPTION
            else an_s3.copy()
        )

        as_opts_an = build_dropdown_options(an_s4, "as_name")
        with af_c5:
            with st.container(border=True):
                an_as = st.selectbox(
                    f"📑 {LABEL_AS_EN}:",
                    options=as_opts_an,
                    index=0,
                    key="an_f_as",
                )

        an_scope_df = (
            an_s4[an_s4["as_name"] == an_as].copy()
            if an_as != ALL_OPTION
            else an_s4.copy()
        )

        if denominator_mode.startswith("Only Ongoing Works"):
            an_scope_df = an_scope_df[an_scope_df["is_status_p"]].copy()

        if an_scope_df.empty:
            st.warning("No works found matching the selected Analytics filter scope.")
            inject_safe_mobile_touch_helpers(show_hud=show_scroll_hud)
            st.stop()

        ae_stats_df = build_officer_analytics_df(an_scope_df, "so_name", LABEL_AE_EN)
        aee_stats_df = build_officer_analytics_df(an_scope_df, "aee_name", LABEL_AEE_EN)

        tot_works_an = len(an_scope_df)
        tot_od_an = int(an_scope_df["is_overdue"].sum())
        tot_ets_an = int(an_scope_df["is_status_ets"].sum())
        od_pct_an = (
            round((tot_od_an / tot_works_an) * 100, 1) if tot_works_an > 0 else 0.0
        )
        ets_pct_an = (
            round((tot_ets_an / tot_works_an) * 100, 1) if tot_works_an > 0 else 0.0
        )
        tot_delayed_an = tot_od_an + tot_ets_an
        unexcused_share_an = (
            round((tot_od_an / tot_delayed_an) * 100, 1) if tot_delayed_an > 0 else 0.0
        )

        od_rows_in_scope = an_scope_df[an_scope_df["is_overdue"]]
        chronic_2yr_plus = (
            int(
                od_rows_in_scope["age_bucket"]
                .isin(["4. 2 – 5 Years", "5. > 5 Years (Chronic)"])
                .sum()
            )
            if not od_rows_in_scope.empty
            else 0
        )

        # --- Mobile 2x3 KPI Ribbon ---
        st.markdown(
            clean_html(
                f"""
                <div class="mob-kpi-grid">
                    <div class="mob-kpi-card" style="border-top: 3px solid #06B6D4;">
                        <div class="mob-kpi-label">📋 Total Works</div>
                        <div class="mob-kpi-val">{tot_works_an:,}</div>
                        <div class="mob-kpi-sub">{len(ae_stats_df)} AEs &bull; {len(aee_stats_df)} AEEs</div>
                    </div>
                    <div class="mob-kpi-card" style="border-top: 3px solid #EF4444;">
                        <div class="mob-kpi-label">🔴 Overdue ('P')</div>
                        <div class="mob-kpi-val" style="color:#FCA5A5;">{tot_od_an:,} ({od_pct_an}%)</div>
                        <div class="mob-kpi-sub">Past Target Date</div>
                    </div>
                    <div class="mob-kpi-card" style="border-top: 3px solid #10B981;">
                        <div class="mob-kpi-label">🟢 Claimed 'ETS'</div>
                        <div class="mob-kpi-val" style="color:#6EE7B7;">{tot_ets_an:,} ({ets_pct_an}%)</div>
                        <div class="mob-kpi-sub">Extension Status</div>
                    </div>
                    <div class="mob-kpi-card" style="border-top: 3px solid #F59E0B;">
                        <div class="mob-kpi-label">⚠️ Delayed Total</div>
                        <div class="mob-kpi-val" style="color:#FCD34D;">{tot_delayed_an:,}</div>
                        <div class="mob-kpi-sub">Overdue + ETS</div>
                    </div>
                    <div class="mob-kpi-card" style="border-top: 3px solid #EC4899;">
                        <div class="mob-kpi-label">🕵️ Unexcused %</div>
                        <div class="mob-kpi-val" style="color:#F472B6;">{unexcused_share_an}%</div>
                        <div class="mob-kpi-sub">Not in ETS</div>
                    </div>
                    <div class="mob-kpi-card" style="border-top: 3px solid #8B5CF6;">
                        <div class="mob-kpi-label">⏳ Chronic (&gt;2Y)</div>
                        <div class="mob-kpi-val" style="color:#C4B5FD;">{chronic_2yr_plus:,}</div>
                        <div class="mob-kpi-sub">2+ Yrs Delayed</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        # =====================================================================
        # ANALYTICS SECTION 1: AE / AEE WISE ANALYTICS & AGEING BREAKDOWN
        # =====================================================================
        with st.container(border=True):
            audit_role = st.radio(
                "👤 Select Officer Cadre:",
                options=[LABEL_AE_EN, LABEL_AEE_EN],
                horizontal=True,
                key="an_truth_role_radio",
            )
            active_audit_df = (
                ae_stats_df if audit_role == LABEL_AE_EN else aee_stats_df
            )

            if not active_audit_df.empty:
                audit_sorted_opts = active_audit_df.sort_values(
                    by=["overdue_works", "ets_works"], ascending=[False, False]
                )[audit_role].tolist()

                picked_off = st.selectbox(
                    f"Select {audit_role} to verify ETS claim:",
                    options=audit_sorted_opts,
                    index=0,
                    key="an_fact_check_select",
                )

                off_row = active_audit_df[
                    active_audit_df[audit_role] == picked_off
                ].iloc[0]
                o_tot = int(off_row["total_works"])
                o_od = int(off_row["overdue_works"])
                o_ets = int(off_row["ets_works"])
                o_ontrack = int(off_row["on_track_p"])
                o_rest = int(off_row["rest_works"])
                o_od_pct = float(off_row["overdue_pct"])
                o_unexcused_pct = float(off_row["unexcused_ratio_pct"])

                if o_od == 0 and o_ets == 0:
                    verdict_color = "#10B981"
                    verdict_title = "✅ CLEAN RECORD (0 Overdue, 0 ETS)"
                    verdict_msg = f"<b>{html.escape(picked_off)}</b> has 0 overdue and 0 ETS across <b>{o_tot}</b> works."
                elif o_od > 0 and o_ets == 0:
                    verdict_color = "#EF4444"
                    verdict_title = "🚨 100% UNEXCUSED OVERDUE (0 ETS!)"
                    verdict_msg = f"<b>{html.escape(picked_off)}</b> has <b>{o_od} Overdue ('P')</b> work(s) and <b>0 works in ETS</b>."
                elif o_od > o_ets:
                    verdict_color = "#F97316"
                    verdict_title = f"⚠️ MORE OVERDUE THAN ETS ({o_unexcused_pct:.1f}% Unexcused)"
                    verdict_msg = f"Out of <b>{o_od + o_ets}</b> delayed works, <b>{o_od}</b> are Overdue ('P') vs. only <b>{o_ets}</b> in ETS."
                else:
                    verdict_color = "#F59E0B"
                    verdict_title = f"📌 ETS HEAVY ({o_ets} ETS vs. {o_od} Overdue)"
                    verdict_msg = f"<b>{html.escape(picked_off)}</b> has <b>{o_ets}</b> in ETS and <b>{o_od}</b> Overdue ('P') ({o_od_pct:.1f}%)."

                target_col = "so_name" if audit_role == LABEL_AE_EN else "aee_name"
                off_od_df = od_rows_in_scope[
                    od_rows_in_scope[target_col] == picked_off
                ]

                od_u6 = od_6_12 = od_1_2 = od_2_3 = od_3_4 = od_4_plus = 0
                if not off_od_df.empty:
                    tree_counts = off_od_df["tree_bucket"].value_counts()
                    od_u6 = int(tree_counts.get("Under 6 Months", 0))
                    od_6_12 = int(tree_counts.get("6 Months to 1 Year", 0))
                    od_1_2 = int(tree_counts.get("1 - 2 Years", 0))
                    od_2_3 = int(tree_counts.get("2 - 3 Years", 0))
                    od_3_4 = int(tree_counts.get("3 - 4 Years", 0))
                    od_4_plus = int(tree_counts.get("4+ Years", 0))

                st.markdown(
                    clean_html(
                        f"""
                        <div class="truth-audit-box" style="border-left-color: {verdict_color};">
                            <div style="font-size:0.88rem;font-weight:800;color:{verdict_color};margin-bottom:4px;">
                                {verdict_title}
                            </div>
                            <div style="font-size:0.80rem;color:#E2E8F0;line-height:1.4;margin-bottom:10px;">
                                {verdict_msg}
                            </div>
                            <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:8px;text-align:center;margin-bottom:10px;">
                                <div style="background:rgba(239,68,68,0.18);border:1px solid rgba(239,68,68,0.55);border-radius:8px;padding:8px;">
                                    <div style="font-size:0.68rem;color:#FCA5A5;font-weight:700;">OVERDUE ('P')</div>
                                    <div style="font-size:1.3rem;font-weight:800;color:#FCA5A5;">{o_od}</div>
                                </div>
                                <div style="background:rgba(16,185,129,0.16);border:1px solid rgba(16,185,129,0.4);border-radius:8px;padding:8px;">
                                    <div style="font-size:0.68rem;color:#6EE7B7;font-weight:700;">IN 'ETS'</div>
                                    <div style="font-size:1.3rem;font-weight:800;color:#6EE7B7;">{o_ets}</div>
                                </div>
                                <div style="background:rgba(59,130,246,0.16);border:1px solid rgba(59,130,246,0.4);border-radius:8px;padding:8px;">
                                    <div style="font-size:0.68rem;color:#93C5FD;font-weight:700;">ON-TRACK ('P')</div>
                                    <div style="font-size:1.3rem;font-weight:800;color:#93C5FD;">{o_ontrack}</div>
                                </div>
                                <div style="background:rgba(148,163,184,0.16);border:1px solid rgba(148,163,184,0.3);border-radius:8px;padding:8px;">
                                    <div style="font-size:0.68rem;color:#CBD5E1;font-weight:700;">TOTAL WORKS</div>
                                    <div style="font-size:1.3rem;font-weight:800;color:#F8FAFC;">{o_tot}</div>
                                </div>
                            </div>
                            <div class="mob-age-grid">
                                <div class="mob-age-card"><div class="mob-age-title">&lt; 6 Mos</div><div class="mob-age-count">{od_u6}</div></div>
                                <div class="mob-age-card"><div class="mob-age-title">6–12 Mos</div><div class="mob-age-count">{od_6_12}</div></div>
                                <div class="mob-age-card"><div class="mob-age-title">1–2 Yrs</div><div class="mob-age-count">{od_1_2}</div></div>
                                <div class="mob-age-card"><div class="mob-age-title">2–3 Yrs</div><div class="mob-age-count">{od_2_3}</div></div>
                                <div class="mob-age-card"><div class="mob-age-title">3–4 Yrs</div><div class="mob-age-count">{od_3_4}</div></div>
                                <div class="mob-age-card"><div class="mob-age-title">4+ Yrs</div><div class="mob-age-count">{od_4_plus}</div></div>
                            </div>
                        </div>
                        """
                    ),
                    unsafe_allow_html=True,
                )

                off_od_col = "_so_name" if audit_role == LABEL_AE_EN else "_aee_name"
                proof_df = pd.DataFrame()
                if not overdue_full_df.empty:
                    proof_df = overdue_full_df[
                        overdue_full_df[off_od_col] == picked_off
                    ].copy()
                    if an_div != ALL_OPTION:
                        proof_df = proof_df[proof_df["_division_name"] == an_div]
                    if an_subdiv != ALL_OPTION:
                        proof_df = proof_df[proof_df["_subdivision_name"] == an_subdiv]
                    if an_ee != ALL_OPTION:
                        proof_df = proof_df[proof_df["_ee_name"] == an_ee]
                    if an_ta != ALL_OPTION:
                        proof_df = proof_df[proof_df["_ta_name"] == an_ta]
                    if an_as != ALL_OPTION:
                        proof_df = proof_df[proof_df["_as_name"] == an_as]

                if not proof_df.empty:
                    st.markdown(
                        f"**📂 {len(proof_df)} Actual Overdue Work(s) — {html.escape(picked_off)}**"
                    )
                    st.dataframe(
                        proof_df[
                            [
                                COL_WORK_NAME_KN,
                                COL_TARGET_DATE_KN,
                                "Overdue By",
                                COL_HOA_KN,
                            ]
                        ].rename(
                            columns={
                                COL_WORK_NAME_KN: "Work Name",
                                COL_TARGET_DATE_KN: "Target Date",
                                COL_HOA_KN: "H.O.A.",
                            }
                        ),
                        use_container_width=True,
                        hide_index=True,
                        height=250,
                    )

        # =====================================================================
        # ANALYTICS SECTION 2: AE & AEE LEADERBOARDS
        # =====================================================================
        st.markdown(
            clean_html(
                """
                <div class="exec-section-banner">
                    <div class="exec-section-title">🏆 Executive AE &amp; AEE Leaderboards</div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        lb_c1, lb_c2 = st.columns(2, gap="small")
        with lb_c1:
            min_works_threshold = st.slider(
                "⚖️ Min. Works for Ranking:",
                min_value=1,
                max_value=50,
                value=5,
            )
        with lb_c2:
            top_n_count = st.radio(
                "Leaderboard Depth:",
                options=[5, 10],
                format_func=lambda x: f"Top {x}",
                horizontal=True,
                index=0,
            )

        ae_qual = ae_stats_df[ae_stats_df["total_works"] >= min_works_threshold].copy()
        if ae_qual.empty:
            ae_qual = ae_stats_df.copy()

        ae_high_cnt = ae_stats_df.sort_values(
            by=["overdue_works", "overdue_pct", "total_works"],
            ascending=[False, False, False],
        ).head(top_n_count)
        ae_high_pct = ae_qual.sort_values(
            by=["overdue_pct", "overdue_works", "total_works"],
            ascending=[False, False, False],
        ).head(top_n_count)
        ae_least_cnt = ae_qual.sort_values(
            by=["overdue_works", "overdue_pct", "total_works"],
            ascending=[True, True, False],
        ).head(top_n_count)
        ae_least_pct = ae_qual.sort_values(
            by=["overdue_pct", "overdue_works", "total_works"],
            ascending=[True, True, False],
        ).head(top_n_count)

        aee_qual = aee_stats_df[
            aee_stats_df["total_works"] >= min_works_threshold
        ].copy()
        if aee_qual.empty:
            aee_qual = aee_stats_df.copy()

        aee_high_cnt = aee_stats_df.sort_values(
            by=["overdue_works", "overdue_pct", "total_works"],
            ascending=[False, False, False],
        ).head(top_n_count)
        aee_high_pct = aee_qual.sort_values(
            by=["overdue_pct", "overdue_works", "total_works"],
            ascending=[False, False, False],
        ).head(top_n_count)
        aee_least_cnt = aee_qual.sort_values(
            by=["overdue_works", "overdue_pct", "total_works"],
            ascending=[True, True, False],
        ).head(top_n_count)
        aee_least_pct = aee_qual.sort_values(
            by=["overdue_pct", "overdue_works", "total_works"],
            ascending=[True, True, False],
        ).head(top_n_count)

        tab_ae_lb, tab_aee_lb = st.tabs(
            [
                f"👷 AE Top {top_n_count}",
                f"👔 AEE Top {top_n_count}",
            ]
        )

        with tab_ae_lb:
            st.markdown(
                render_executive_leaderboard_panel(
                    ae_high_cnt,
                    LABEL_AE_EN,
                    f"🔴 Top {top_n_count} AEs — Highest Overdue (Count)",
                    "By Overdue Count",
                    is_danger=True,
                    primary_metric="count",
                ),
                unsafe_allow_html=True,
            )
            st.markdown(
                render_executive_leaderboard_panel(
                    ae_high_pct,
                    LABEL_AE_EN,
                    f"🔴 Top {top_n_count} AEs — Highest Overdue (%)",
                    f"Min. {min_works_threshold} Works",
                    is_danger=True,
                    primary_metric="pct",
                ),
                unsafe_allow_html=True,
            )
            st.markdown(
                render_executive_leaderboard_panel(
                    ae_least_cnt,
                    LABEL_AE_EN,
                    f"🟢 Top {top_n_count} AEs — Least Overdue (Count)",
                    f"Min. {min_works_threshold} Works",
                    is_danger=False,
                    primary_metric="count",
                ),
                unsafe_allow_html=True,
            )
            st.markdown(
                render_executive_leaderboard_panel(
                    ae_least_pct,
                    LABEL_AE_EN,
                    f"🟢 Top {top_n_count} AEs — Least Overdue (%)",
                    f"Min. {min_works_threshold} Works",
                    is_danger=False,
                    primary_metric="pct",
                ),
                unsafe_allow_html=True,
            )

        with tab_aee_lb:
            st.markdown(
                render_executive_leaderboard_panel(
                    aee_high_cnt,
                    LABEL_AEE_EN,
                    f"🔴 Top {top_n_count} AEEs — Highest Overdue (Count)",
                    "By Overdue Count",
                    is_danger=True,
                    primary_metric="count",
                ),
                unsafe_allow_html=True,
            )
            st.markdown(
                render_executive_leaderboard_panel(
                    aee_high_pct,
                    LABEL_AEE_EN,
                    f"🔴 Top {top_n_count} AEEs — Highest Overdue (%)",
                    f"Min. {min_works_threshold} Works",
                    is_danger=True,
                    primary_metric="pct",
                ),
                unsafe_allow_html=True,
            )
            st.markdown(
                render_executive_leaderboard_panel(
                    aee_least_cnt,
                    LABEL_AEE_EN,
                    f"🟢 Top {top_n_count} AEEs — Least Overdue (Count)",
                    f"Min. {min_works_threshold} Works",
                    is_danger=False,
                    primary_metric="count",
                ),
                unsafe_allow_html=True,
            )
            st.markdown(
                render_executive_leaderboard_panel(
                    aee_least_pct,
                    LABEL_AEE_EN,
                    f"🟢 Top {top_n_count} AEEs — Least Overdue (%)",
                    f"Min. {min_works_threshold} Works",
                    is_danger=False,
                    primary_metric="pct",
                ),
                unsafe_allow_html=True,
            )

    inject_safe_mobile_touch_helpers(show_hud=show_scroll_hud)
    st.stop()


# =============================================================================
# PAGE 1 VIEW: MAIN OVERDUE WORKS DASHBOARD (MOBILE)
# =============================================================================
with st.container(key="mob_page_dashboard"):
    f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns(5, gap="small")

    unique_divisions = sorted(
        [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
    )
    division_options = [ALL_OPTION] + unique_divisions

    with f_col1:
        with st.container(border=True):
            selected_division = st.selectbox(
                f"🏛️ {LABEL_DIV_EN}:",
                options=division_options,
                index=0,
            )

    df_step1 = (
        master_df[master_df["division_name"] == selected_division].copy()
        if selected_division != ALL_OPTION
        else master_df.copy()
    )

    subdivision_options = build_dropdown_options(df_step1, "subdivision_name")
    with f_col2:
        with st.container(border=True):
            selected_subdivision = st.selectbox(
                f"🏢 {LABEL_SUBDIV_EN}:",
                options=subdivision_options,
                index=0,
            )

    df_step2 = (
        df_step1[df_step1["subdivision_name"] == selected_subdivision].copy()
        if selected_subdivision != ALL_OPTION
        else df_step1.copy()
    )

    ee_options = build_dropdown_options(df_step2, "ee_name")
    with f_col3:
        with st.container(border=True):
            selected_ee = st.selectbox(
                f"👔 {LABEL_EE_EN}:",
                options=ee_options,
                index=0,
            )

    df_step3 = (
        df_step2[df_step2["ee_name"] == selected_ee].copy()
        if selected_ee != ALL_OPTION
        else df_step2.copy()
    )

    ta_options = build_dropdown_options(df_step3, "ta_name")
    with f_col4:
        with st.container(border=True):
            selected_ta = st.selectbox(
                f"🛠️ {LABEL_TA_EN}:",
                options=ta_options,
                index=0,
            )

    df_step4 = (
        df_step3[df_step3["ta_name"] == selected_ta].copy()
        if selected_ta != ALL_OPTION
        else df_step3.copy()
    )

    as_options = build_dropdown_options(df_step4, "as_name")
    with f_col5:
        with st.container(border=True):
            selected_as = st.selectbox(
                f"📑 {LABEL_AS_EN}:",
                options=as_options,
                index=0,
            )

    filtered_master_df = (
        df_step4[df_step4["as_name"] == selected_as].copy()
        if selected_as != ALL_OPTION
        else df_step4.copy()
    )

    r_col1, r_col2 = st.columns(2, gap="small")
    with r_col1:
        view_choice = st.radio(
            "Officer View",
            options=[LABEL_AE_EN, LABEL_AEE_EN],
            horizontal=True,
            label_visibility="collapsed",
        )
    with r_col2:
        sort_choice = st.radio(
            "Sort Order",
            options=[
                "Highest Overdue Count at Top",
                "Highest Overdue Percentage (%) at Top",
            ],
            horizontal=True,
            label_visibility="collapsed",
        )

    group_col = "so_name" if view_choice == LABEL_AE_EN else "aee_name"
    full_filter_col = "_so_name" if view_choice == LABEL_AE_EN else "_aee_name"

    valid_df = filtered_master_df[filtered_master_df[group_col] != ""].copy()

    if not show_unassigned_bar:
        valid_df = valid_df[valid_df[group_col] != UNASSIGNED_LABEL].copy()

    if denominator_mode.startswith("Only Ongoing Works"):
        valid_df = valid_df[valid_df["is_status_p"]].copy()

    if valid_df.empty:
        st.warning("No works found matching the selected filters.")
        inject_safe_mobile_touch_helpers(show_hud=show_scroll_hud)
        st.stop()

    summary_df = (
        valid_df.groupby(group_col, as_index=False)
        .agg(
            total_works=("is_overdue", "count"),
            ongoing_p_works=("is_status_p", "sum"),
            ets_works=("is_status_ets", "sum"),
            overdue_works=("is_overdue", "sum"),
        )
        .rename(columns={group_col: view_choice})
    )

    summary_df["rest_works"] = (
        summary_df["total_works"]
        - summary_df["overdue_works"]
        - summary_df["ets_works"]
    ).clip(lower=0)

    summary_df["overdue_pct"] = (
        (summary_df["overdue_works"] / summary_df["total_works"]) * 100
    ).round(1).fillna(0.0)

    summary_df["_is_unassigned"] = summary_df[view_choice] == UNASSIGNED_LABEL

    if sort_choice == "Highest Overdue Percentage (%) at Top":
        summary_df = summary_df.sort_values(
            by=["_is_unassigned", "overdue_pct", "overdue_works", "total_works"],
            ascending=[True, False, False, False],
        ).reset_index(drop=True)
    else:
        summary_df = summary_df.sort_values(
            by=["_is_unassigned", "overdue_works", "overdue_pct", "total_works"],
            ascending=[True, False, False, False],
        ).reset_index(drop=True)

    # -------------------------------------------------------------------------
    # MOBILE KPI SUMMARY CARDS (2x3 GRID ON PHONES)
    # -------------------------------------------------------------------------
    total_assigned_all = int(summary_df["total_works"].sum())
    total_overdue_all = int(summary_df["overdue_works"].sum())
    total_ets_all = int(summary_df["ets_works"].sum())

    if hide_zero_overdue:
        named_officers_count = int(
            ((~summary_df["_is_unassigned"]) & (summary_df["overdue_works"] > 0)).sum()
        )
    else:
        named_officers_count = int((~summary_df["_is_unassigned"]).sum())

    overall_overdue_pct = (
        round((total_overdue_all / total_assigned_all) * 100, 1)
        if total_assigned_all > 0
        else 0.0
    )

    active_scope_parts = [
        f"Div: {selected_division}",
        f"Sub: {selected_subdivision}",
    ]
    filter_scope_label = " | ".join(active_scope_parts)

    page1_od_df = valid_df[valid_df["is_overdue"]]
    chronic_2yr_plus_p1 = (
        int(
            page1_od_df["age_bucket"]
            .isin(["4. 2 – 5 Years", "5. > 5 Years (Chronic)"])
            .sum()
        )
        if not page1_od_df.empty
        else 0
    )

    st.markdown(
        clean_html(
            f"""
            <div class="mob-kpi-grid">
                <div class="mob-kpi-card" style="border-top: 3px solid #3B82F6;">
                    <div class="mob-kpi-label">📅 Date (Today)</div>
                    <div class="mob-kpi-val">{today_str}</div>
                </div>
                <div class="mob-kpi-card" style="border-top: 3px solid #8B5CF6;">
                    <div class="mob-kpi-label">👥 {html.escape(view_choice)}s</div>
                    <div class="mob-kpi-val">{named_officers_count:,}</div>
                </div>
                <div class="mob-kpi-card" style="border-top: 3px solid #06B6D4;">
                    <div class="mob-kpi-label">📋 Total Works</div>
                    <div class="mob-kpi-val">{total_assigned_all:,}</div>
                </div>
                <div class="mob-kpi-card" style="border-top: 3px solid #EF4444;">
                    <div class="mob-kpi-label">🔴 Overdue ('P')</div>
                    <div class="mob-kpi-val" style="color:#FCA5A5;">{total_overdue_all:,} ({overall_overdue_pct}%)</div>
                </div>
                <div class="mob-kpi-card" style="border-top: 3px solid #F59E0B;">
                    <div class="mob-kpi-label">⏳ Chronic (&gt;2Y)</div>
                    <div class="mob-kpi-val" style="color:#FCD34D;">{chronic_2yr_plus_p1:,}</div>
                </div>
                <div class="mob-kpi-card" style="border-top: 3px solid #4ADE80;">
                    <div class="mob-kpi-label">🟢 'ETS' Works</div>
                    <div class="mob-kpi-val" style="color:#6EE7B7;">{total_ets_all:,}</div>
                </div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    # -------------------------------------------------------------------------
    # MOBILE 3x2 OVERDUE AGEING BUCKET CARDS
    # -------------------------------------------------------------------------
    od_u6_p1 = od_6_12_p1 = od_1_2_p1 = od_2_3_p1 = od_3_4_p1 = od_4_plus_p1 = 0
    if not page1_od_df.empty:
        p1_tree_counts = page1_od_df["tree_bucket"].value_counts()
        od_u6_p1 = int(p1_tree_counts.get("Under 6 Months", 0))
        od_6_12_p1 = int(p1_tree_counts.get("6 Months to 1 Year", 0))
        od_1_2_p1 = int(p1_tree_counts.get("1 - 2 Years", 0))
        od_2_3_p1 = int(p1_tree_counts.get("2 - 3 Years", 0))
        od_3_4_p1 = int(p1_tree_counts.get("3 - 4 Years", 0))
        od_4_plus_p1 = int(p1_tree_counts.get("4+ Years", 0))

    denom_od_p1 = max(total_overdue_all, 1)
    age_items_p1 = [
        ("&lt; 6 Months", od_u6_p1),
        ("6–12 Months", od_6_12_p1),
        ("1–2 Years", od_1_2_p1),
        ("2–3 Years", od_2_3_p1),
        ("3–4 Years", od_3_4_p1),
        ("4+ Years", od_4_plus_p1),
    ]
    age_cards_html = []
    for lbl, val in age_items_p1:
        pct = round((val / denom_od_p1) * 100, 1) if total_overdue_all > 0 else 0.0
        age_cards_html.append(
            f'<div class="mob-age-card">'
            f'<div class="mob-age-title">{lbl}</div>'
            f'<div class="mob-age-count">{val:,}</div>'
            f'<div class="mob-age-pct">{pct:.1f}%</div>'
            f"</div>"
        )

    st.markdown(
        clean_html(f'<div class="mob-age-grid">{"".join(age_cards_html)}</div>'),
        unsafe_allow_html=True,
    )

    # =========================================================================
    # SECTION 1: MOBILE-FITTED HORIZONTAL BAR CHART
    # =========================================================================
    st.markdown(
        clean_html(
            f"""
            <div class="exec-section-banner">
                <div class="exec-section-title">📈 Overdue &amp; ETS by {html.escape(view_choice)}</div>
                <div class="exec-scope-badge">{html.escape(filter_scope_label)}</div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    if hide_zero_overdue:
        chart_source_df = summary_df[
            (summary_df["overdue_works"] > 0) | summary_df["_is_unassigned"]
        ].copy()
    else:
        chart_source_df = summary_df.copy()

    if chart_source_df.empty:
        st.info("No officers match the current filter.")
    else:
        plot_df = chart_source_df.iloc[::-1].copy()
        plot_df["y_short"] = plot_df[view_choice].map(lambda s: shorten_label(s, 17))
        plot_df["is_shrunk"] = shrink_outliers & (
            plot_df["total_works"] > BAR_SHRINK_THRESHOLD
        )

        def _build_bar_label(r) -> str:
            if sort_choice == "Highest Overdue Percentage (%) at Top":
                return (
                    f" <b>{r['overdue_pct']:.0f}%</b> "
                    f"({int(r['overdue_works'])}/{int(r['total_works'])})"
                )
            return (
                f" <b>{int(r['overdue_works'])} OD</b> | "
                f"<span style='color:#4ADE80'>{int(r['ets_works'])} ETS</span> / "
                f"{int(r['total_works'])} ({r['overdue_pct']:.0f}%)"
            )

        plot_df["bar_label"] = plot_df.apply(_build_bar_label, axis=1)

        fig = go.Figure()

        if chart_style == "Scale by Overdue Works (Full-Width Bars)":
            fig.add_trace(
                go.Bar(
                    y=plot_df[view_choice],
                    x=plot_df["overdue_works"],
                    name="Overdue ('P')",
                    orientation="h",
                    marker=dict(color="#EF4444"),
                    text=plot_df["bar_label"],
                    textposition="outside",
                    textfont=dict(size=11),
                    cliponaxis=False,
                )
            )
            max_x = max(int(plot_df["overdue_works"].max()), 1)
            x_range_max = max_x * 1.38
            x_axis_title = "Overdue Works"

        elif chart_style == "Scale by Overdue Percentage (%)":
            fig.add_trace(
                go.Bar(
                    y=plot_df[view_choice],
                    x=plot_df["overdue_pct"],
                    name="Overdue %",
                    orientation="h",
                    marker=dict(color="#EF4444"),
                    text=plot_df["bar_label"],
                    textposition="outside",
                    textfont=dict(size=11),
                    cliponaxis=False,
                )
            )
            max_x = max(float(plot_df["overdue_pct"].max()), 1.0)
            x_range_max = max_x * 1.38
            x_axis_title = "Overdue %"

        else:
            plot_df["overdue_plus_ets"] = (
                plot_df["overdue_works"] + plot_df["ets_works"]
            )
            if shrink_outliers:
                plot_df["vis_total"] = plot_df["total_works"].map(compress_bar_length)
                plot_df["vis_overdue_plus_ets"] = plot_df["overdue_plus_ets"].map(
                    compress_bar_length
                )
                plot_df["vis_overdue"] = plot_df["overdue_works"].map(
                    compress_bar_length
                )
            else:
                plot_df["vis_total"] = plot_df["total_works"].astype(float)
                plot_df["vis_overdue_plus_ets"] = plot_df["overdue_plus_ets"].astype(
                    float
                )
                plot_df["vis_overdue"] = plot_df["overdue_works"].astype(float)

            fig.add_trace(
                go.Bar(
                    y=plot_df[view_choice],
                    x=plot_df["vis_total"],
                    name="Rest",
                    orientation="h",
                    marker=dict(color="rgba(148, 163, 184, 0.30)"),
                    text=plot_df["bar_label"],
                    textposition="outside",
                    textfont=dict(size=11),
                    cliponaxis=False,
                )
            )
            fig.add_trace(
                go.Bar(
                    y=plot_df[view_choice],
                    x=plot_df["vis_overdue_plus_ets"],
                    name="ETS",
                    orientation="h",
                    marker=dict(color="#4ADE80"),
                )
            )
            fig.add_trace(
                go.Bar(
                    y=plot_df[view_choice],
                    x=plot_df["vis_overdue"],
                    name="Overdue ('P')",
                    orientation="h",
                    marker=dict(color="#EF4444"),
                )
            )
            max_x = max(float(plot_df["vis_total"].max()), 1.0)
            x_range_max = max_x * 1.38
            x_axis_title = "Works Count"

        chart_height = max(420, len(plot_df) * 32 + 80)

        fig.update_layout(
            barmode="overlay",
            height=chart_height,
            bargap=0.22,
            dragmode=False,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(15, 23, 42, 0.35)",
            xaxis=dict(
                title=x_axis_title,
                range=[0, x_range_max],
                gridcolor="rgba(148, 163, 184, 0.14)",
            ),
            yaxis=dict(
                title="",
                automargin=False,
                tickmode="array",
                tickvals=plot_df[view_choice],
                ticktext=plot_df["y_short"],
                tickfont=dict(size=11),
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.01,
                xanchor="left",
                x=0,
                font=dict(size=11),
                traceorder="reversed",
            ),
            margin=dict(l=118, r=10, t=32, b=35),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar": False},
        )

    # =========================================================================
    # SECTION 2: COMBINED OVERDUE WORKS TABLE
    # =========================================================================
    st.markdown(
        clean_html(
            f"""
            <div class="exec-section-banner">
                <div class="exec-section-title">📂 Overdue Works by {html.escape(view_choice)}</div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    officer_options = (
        chart_source_df[view_choice].tolist()
        if not chart_source_df.empty
        else summary_df[view_choice].tolist()
    )

    officer_stats_map = {
        row[view_choice]: (
            int(row["overdue_works"]),
            int(row["ets_works"]),
            int(row["total_works"]),
            float(row["overdue_pct"]),
        )
        for _, row in summary_df.iterrows()
    }

    selected_officer = st.selectbox(
        f"Select {view_choice}:",
        options=officer_options,
        index=0 if officer_options else None,
        format_func=lambda name: (
            f"{name} ({officer_stats_map[name][0]} OD | "
            f"{officer_stats_map[name][1]} ETS / {officer_stats_map[name][2]} Tot)"
        ),
    )

    if selected_officer and not overdue_full_df.empty:
        officer_overdue_df = overdue_full_df[
            overdue_full_df[full_filter_col] == selected_officer
        ].copy()

        if selected_division != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_division_name"] == selected_division
            ].copy()
        if selected_subdivision != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_subdivision_name"] == selected_subdivision
            ].copy()
        if selected_ee != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_ee_name"] == selected_ee
            ].copy()
        if selected_ta != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_ta_name"] == selected_ta
            ].copy()
        if selected_as != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_as_name"] == selected_as
            ].copy()

        opposite_officer_col = (
            COL_AEE_KN if view_choice == LABEL_AE_EN else COL_SO_KN
        )
        opposite_officer_en = (
            LABEL_AEE_EN if view_choice == LABEL_AE_EN else LABEL_AE_EN
        )

        curated_cols = [
            COL_WORK_NAME_KN,
            COL_TARGET_DATE_KN,
            "Overdue By",
            COL_HOA_KN,
            COL_DIV_KN,
            COL_SUBDIV_KN,
            opposite_officer_col,
            COL_START_DATE_KN,
        ]

        officer_overdue_df = officer_overdue_df[curated_cols].reset_index(drop=True)
        officer_overdue_df.rename(
            columns={
                COL_WORK_NAME_KN: "Work Name",
                COL_TARGET_DATE_KN: "Target Date",
                COL_HOA_KN: "H.O.A.",
                COL_DIV_KN: LABEL_DIV_EN,
                COL_SUBDIV_KN: LABEL_SUBDIV_EN,
                opposite_officer_col: opposite_officer_en,
                COL_START_DATE_KN: "Start Date",
            },
            inplace=True,
        )
        officer_overdue_df.insert(
            0, "Sl.", range(1, len(officer_overdue_df) + 1)
        )

        if officer_overdue_df.empty:
            st.success(f"🎉 {selected_officer} has 0 overdue works in this scope!")
        else:
            safe_name = (
                re.sub(r"[^\w\-]+", "_", selected_officer).strip("_")
                or "Unassigned"
            )
            dl_c1, dl_c2 = st.columns(2, gap="small")
            with dl_c1:
                st.download_button(
                    label="📥 Excel (.xlsx)",
                    data=convert_df_to_excel_bytes(officer_overdue_df),
                    file_name=f"Overdue_{safe_name}_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl_c2:
                st.download_button(
                    label="📥 CSV (.csv)",
                    data=officer_overdue_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"Overdue_{safe_name}_{today_str}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            st.markdown(
                render_ce_overdue_html_table(officer_overdue_df),
                unsafe_allow_html=True,
            )

    # =========================================================================
    # SECTION 3: OFFICER-WISE SUMMARY TABLE
    # =========================================================================
    st.markdown(
        clean_html(
            """
            <div class="exec-section-banner">
                <div class="exec-section-title">📋 Officer-Wise Summary Table</div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    display_table = summary_df[
        [
            view_choice,
            "total_works",
            "overdue_works",
            "ets_works",
            "overdue_pct",
        ]
    ].rename(
        columns={
            "total_works": "Total",
            "overdue_works": "Overdue ('P')",
            "ets_works": "ETS",
            "overdue_pct": "Overdue %",
        }
    )
    display_table.index = display_table.index + 1

    st.dataframe(
        display_table.style.format({"Overdue %": "{:.1f}%"}),
        use_container_width=True,
    )

    with st.expander("🔍 Inspect Worksheet Loading Status (23 Sheets)"):
        st.dataframe(diag_df, use_container_width=True)

    inject_safe_mobile_touch_helpers(show_hud=show_scroll_hud)
