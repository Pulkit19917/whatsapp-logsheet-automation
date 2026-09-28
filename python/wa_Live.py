"""
Live Photo Monitoring Dashboard
===============================

Shows, for one date, which site/group folders have received photos and which
are still pending.

Everything that changes between setups lives in config/config.json:
    paths.output_dir          folder that holds one sub-folder per date
    dashboard.*               branding, colours, formats, refresh, labels

Every key is optional - if it is missing, the default below is used, so the
dashboard also runs with no "dashboard" section at all.

Run:
    streamlit run python/wa_Live.py

Optional environment variables:
    PROJECT_ROOT              force the project root folder
    DASHBOARD_CONFIG_FILE     use a different config file
"""

from __future__ import annotations

import html
import json
import os
import re
import time
from datetime import date, datetime
from pathlib import Path

import streamlit as st


# ============================================================
# 1. PROJECT ROOT & CONFIG LOADING
# ============================================================

CURRENT_FILE = Path(__file__).resolve()


def find_project_root() -> Path:
    """
    Project root = the nearest parent folder that contains
    config/config.json (so the dashboard can live in the project root or in
    a sub-folder such as python/).
    """

    forced = os.environ.get("PROJECT_ROOT")

    if forced:
        return Path(forced)

    for folder in CURRENT_FILE.parents:
        if (folder / "config" / "config.json").is_file():
            return folder

    return CURRENT_FILE.parent


PROJECT_ROOT = find_project_root()

CONFIG_FILE = Path(
    os.environ.get("DASHBOARD_CONFIG_FILE")
    or PROJECT_ROOT / "config" / "config.json"
)


def load_config() -> tuple[dict, str | None]:
    """Return (config, error_message). Never raises."""

    if not CONFIG_FILE.is_file():
        return {}, f"Config file not found, using defaults: {CONFIG_FILE}"

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        return (data if isinstance(data, dict) else {}), None

    except json.JSONDecodeError as e:
        return {}, f"config.json is not valid JSON, using defaults: {e}"

    except OSError as e:
        return {}, f"Could not read config.json, using defaults: {e}"


CONFIG, CONFIG_ERROR = load_config()

PATH_CONFIG = CONFIG.get("paths") or {}
DASH_CONFIG = CONFIG.get("dashboard") or {}
BRAND_CONFIG = DASH_CONFIG.get("branding") or {}
THEME_CONFIG = DASH_CONFIG.get("theme") or {}
FORMAT_CONFIG = DASH_CONFIG.get("formats") or {}
REFRESH_CONFIG = DASH_CONFIG.get("refresh") or {}
LABEL_CONFIG = DASH_CONFIG.get("labels") or {}


# ============================================================
# 2. CONFIG HELPERS
# ============================================================

def cfg_str(source: dict, key: str, default: str) -> str:
    value = source.get(key)

    if value is None or str(value).strip() == "":
        return default

    return str(value).strip()


def cfg_int(source: dict, key: str, default: int) -> int:
    try:
        return int(source.get(key))
    except (TypeError, ValueError):
        return default


def cfg_bool(source: dict, key: str, default: bool) -> bool:
    value = source.get(key)

    if isinstance(value, bool):
        return value

    if isinstance(value, str):
        text = value.strip().lower()

        if text in {"1", "true", "yes", "y", "on"}:
            return True

        if text in {"0", "false", "no", "n", "off"}:
            return False

    return default


def cfg_list(source: dict, key: str, default: list[str]) -> list[str]:
    value = source.get(key)

    if isinstance(value, list) and value:
        return [str(v) for v in value]

    return list(default)


def resolve_path(value, default_relative: str) -> Path:
    """Absolute path as given; relative path from the project root."""

    if value is None or str(value).strip() == "":
        value = default_relative

    path = Path(str(value))

    return path if path.is_absolute() else PROJECT_ROOT / path


def esc(value) -> str:
    """HTML-escape text that is placed inside custom HTML."""

    return html.escape(str(value), quote=True)


SAFE_CSS_VALUE = re.compile(r"^[#\w\s(),.%-]+$")


# ============================================================
# 3. SETTINGS (config with defaults)
# ============================================================

OUTPUT_DIR = resolve_path(PATH_CONFIG.get("output_dir"), "data/output")

IMAGE_EXTENSIONS = {
    (e if e.startswith(".") else "." + e).lower()
    for e in cfg_list(
        DASH_CONFIG,
        "image_extensions",
        [".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff"],
    )
}

SCAN_SUBFOLDERS = cfg_bool(DASH_CONFIG, "scan_subfolders", True)
SHOW_FOLDER_INFO = cfg_bool(DASH_CONFIG, "show_folder_info", True)
SHOW_SYSTEM_INFO = cfg_bool(DASH_CONFIG, "show_system_info", True)
SHOW_FOLDER_COLUMN = cfg_bool(DASH_CONFIG, "show_folder_column", True)
MAX_CONTENT_WIDTH = cfg_int(DASH_CONFIG, "max_content_width_px", 1600)

# ---- formats ------------------------------------------------
DATE_FOLDER_FORMAT = cfg_str(FORMAT_CONFIG, "date_folder", "%Y-%m-%d")
DATE_INPUT_FORMAT = cfg_str(FORMAT_CONFIG, "date_input", "%Y-%m-%d")
DATE_DISPLAY_FORMAT = cfg_str(FORMAT_CONFIG, "date_display", "%d-%m-%Y")
TIME_DISPLAY_FORMAT = cfg_str(FORMAT_CONFIG, "time_display", "%I:%M:%S %p")
DATETIME_DISPLAY_FORMAT = cfg_str(
    FORMAT_CONFIG, "datetime_display", "%d-%m-%Y %I:%M:%S %p"
)

DATE_INPUT_HINT = (
    DATE_INPUT_FORMAT.replace("%Y", "YYYY")
    .replace("%y", "YY")
    .replace("%m", "MM")
    .replace("%d", "DD")
)

# ---- branding ----------------------------------------------
PAGE_TITLE = cfg_str(BRAND_CONFIG, "page_title", "Live Photo Monitoring")
PAGE_ICON = cfg_str(BRAND_CONFIG, "page_icon", "📊")
COMPANY_NAME = cfg_str(BRAND_CONFIG, "company_name", "Site Operations")
TAGLINE = cfg_str(BRAND_CONFIG, "tagline", "Operations • Fleet • Site Monitoring")
LOGO_TEXT = cfg_str(BRAND_CONFIG, "logo_text", "OPS")
SIDEBAR_TITLE = cfg_str(BRAND_CONFIG, "sidebar_title", LOGO_TEXT)
SIDEBAR_SUBTITLE = cfg_str(BRAND_CONFIG, "sidebar_subtitle", "SITE MONITORING")
DASHBOARD_TITLE = cfg_str(
    BRAND_CONFIG, "dashboard_title", "WhatsApp Site Photo Monitoring Dashboard"
)
DASHBOARD_SUBTITLE = cfg_str(
    BRAND_CONFIG,
    "dashboard_subtitle",
    "Monitor daily site photo submissions and pending groups "
    "from one central dashboard.",
)
LIVE_BADGE_TEXT = cfg_str(BRAND_CONFIG, "live_badge", "LIVE MONITORING")
SYSTEM_STATUS_TEXT = cfg_str(BRAND_CONFIG, "system_status", "Online")

# ---- labels -------------------------------------------------
LABEL_TOTAL = cfg_str(LABEL_CONFIG, "total_sites", "Total Groups / Sites")
LABEL_RECEIVED = cfg_str(LABEL_CONFIG, "received_metric", "Groups Received")
LABEL_PENDING = cfg_str(LABEL_CONFIG, "pending_metric", "Pending Groups")
LABEL_COMPLETION = cfg_str(LABEL_CONFIG, "completion_metric", "Completion")
STATUS_RECEIVED = cfg_str(LABEL_CONFIG, "status_received", "Received")
STATUS_PENDING = cfg_str(LABEL_CONFIG, "status_pending", "Pending")

# ---- refresh ------------------------------------------------
REFRESH_ENABLED = cfg_bool(REFRESH_CONFIG, "enabled", True)
REFRESH_DEFAULT = cfg_int(REFRESH_CONFIG, "default_seconds", 30)
REFRESH_MIN = cfg_int(REFRESH_CONFIG, "min_seconds", 10)
REFRESH_MAX = cfg_int(REFRESH_CONFIG, "max_seconds", 300)
REFRESH_STEP = max(1, cfg_int(REFRESH_CONFIG, "step_seconds", 10))

if REFRESH_MAX <= REFRESH_MIN:
    REFRESH_MAX = REFRESH_MIN + REFRESH_STEP

REFRESH_DEFAULT = min(max(REFRESH_DEFAULT, REFRESH_MIN), REFRESH_MAX)

# ---- theme --------------------------------------------------
DEFAULT_THEME = {
    "primary": "#0e3d30",
    "primary_dark": "#09211b",
    "primary_light": "#155745",
    "sidebar_top": "#091f19",
    "sidebar_bottom": "#11362e",
    "accent": "#00e676",
    "accent_hover": "#00c853",
    "background": "#f4f7f6",
    "card_border": "#e1e8e5",
    "input_background": "#051410",
    "muted_text": "#5c7069",
    "header_muted": "#c2e5d9",
}

THEME = {}

for _key, _default in DEFAULT_THEME.items():
    _value = str(THEME_CONFIG.get(_key, _default)).strip()
    THEME[_key] = _value if SAFE_CSS_VALUE.match(_value) else _default


# ============================================================
# 4. PAGE CONFIGURATION  (must be the first Streamlit call)
# ============================================================

st.set_page_config(
    page_title=PAGE_TITLE,
    page_icon=PAGE_ICON,
    layout="wide",
    initial_sidebar_state="expanded",
)

if CONFIG_ERROR:
    st.warning(CONFIG_ERROR)


# ============================================================
# 5. STYLING  (theme colours come from config)
# ============================================================

CSS_TEMPLATE = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

.stApp { background-color: __background__ !important; }

.block-container {
    padding-top: 3.5rem !important;
    padding-bottom: 2.5rem !important;
    max-width: __max_width__px !important;
}

header[data-testid="stHeader"] {
    background-color: __background__ !important;
    z-index: 99 !important;
}

/* ---------- sidebar ---------- */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, __sidebar_top__ 0%, __sidebar_bottom__ 100%) !important;
    border-right: 1px solid __primary_dark__;
}

[data-testid="stSidebar"] *,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #ffffff !important; }

.sidebar-brand {
    text-align: center;
    padding: 10px 0 15px 0;
    border-bottom: 1px solid rgba(255, 255, 255, 0.12);
    margin-bottom: 15px;
}
.sidebar-brand-title {
    font-size: 26px; font-weight: 900; letter-spacing: 2px;
    color: #ffffff !important;
}
.sidebar-brand-sub {
    font-size: 10px; letter-spacing: 2px; font-weight: 700;
    color: __accent__ !important; text-transform: uppercase; margin-top: 2px;
}

[data-testid="stSidebar"] div[data-baseweb="input"],
[data-testid="stSidebar"] div[data-baseweb="base-input"] {
    background-color: __input_background__ !important;
    border: 1px solid __accent__ !important;
    border-radius: 8px !important;
}

[data-testid="stSidebar"] input[type="text"],
[data-testid="stSidebar"] input {
    background-color: __input_background__ !important;
    color: __accent__ !important;
    -webkit-text-fill-color: __accent__ !important;
    font-weight: 800 !important;
    font-size: 1rem !important;
    letter-spacing: 0.5px !important;
}

[data-testid="stSidebar"] [data-testid="stCodeBlock"],
[data-testid="stSidebar"] [data-testid="stCode"],
[data-testid="stSidebar"] pre,
[data-testid="stSidebar"] code {
    background-color: __input_background__ !important;
    color: __accent__ !important;
    border: 1px solid rgba(255, 255, 255, 0.15) !important;
    border-radius: 8px !important;
    font-family: monospace !important;
    font-size: 0.8rem !important;
    white-space: pre-wrap !important;
    word-break: break-all !important;
}

[data-testid="stSidebar"] button[aria-label="Copy to clipboard"] { display: none !important; }
[data-testid="stSidebar"] [data-baseweb="checkbox"] span { color: #ffffff !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255, 255, 255, 0.12) !important; }

/* ---------- header banner ---------- */
.app-header {
    background: linear-gradient(135deg, __primary_dark__ 0%, __primary__ 50%, __primary_light__ 100%);
    border-radius: 14px;
    padding: 24px 30px;
    margin-bottom: 25px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.15);
    color: #ffffff;
    border: 1px solid rgba(255, 255, 255, 0.08);
}
.app-header-top {
    display: flex; justify-content: space-between; align-items: center;
    gap: 16px; flex-wrap: wrap;
}
.app-brand-section { display: flex; align-items: center; gap: 16px; }
.app-logo {
    min-width: 54px; height: 54px; padding: 0 8px;
    display: flex; align-items: center; justify-content: center;
    border-radius: 12px; background: #ffffff; color: __primary__;
    font-size: 18px; font-weight: 900; letter-spacing: -0.5px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
}
.app-company { font-size: 1.35rem; font-weight: 800; line-height: 1.2; color: #ffffff; margin-bottom: 2px; }
.app-company-subtitle { font-size: 0.85rem; color: __header_muted__; font-weight: 500; }
.app-live-badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: rgba(255, 255, 255, 0.10);
    border: 1px solid __accent__;
    padding: 6px 14px; border-radius: 30px;
    color: __accent__; font-size: 0.78rem; font-weight: 700; letter-spacing: 0.5px;
}
.app-live-dot {
    width: 8px; height: 8px; border-radius: 50%;
    background: __accent__; box-shadow: 0 0 10px __accent__;
}
.app-header-divider { height: 1px; background: rgba(255, 255, 255, 0.12); margin: 16px 0; }
.app-dashboard-title { font-size: 1.15rem; font-weight: 700; color: #ffffff; margin-bottom: 4px; }
.app-dashboard-subtitle { font-size: 0.85rem; color: __header_muted__; }
.app-info-row { display: flex; gap: 12px; flex-wrap: wrap; margin-top: 16px; }
.app-info-box {
    background: rgba(0, 0, 0, 0.22);
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-radius: 8px; padding: 8px 14px; min-width: 140px;
}
.app-info-label { color: __header_muted__; font-size: 0.68rem; font-weight: 700; letter-spacing: 0.5px; margin-bottom: 2px; }
.app-info-value { color: #ffffff; font-size: 0.9rem; font-weight: 700; }

/* ---------- main content ---------- */
.section-heading {
    color: __primary__; font-size: 1.1rem; font-weight: 800;
    margin-top: 18px; margin-bottom: 12px; letter-spacing: -0.3px;
}

[data-testid="stMetric"] {
    background: #ffffff !important;
    border: 1px solid __card_border__ !important;
    border-radius: 12px !important;
    padding: 16px 20px !important;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04) !important;
}
[data-testid="stMetricLabel"] {
    color: __muted_text__ !important; font-size: 0.8rem !important;
    font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.5px;
}
[data-testid="stMetricValue"] {
    color: __primary__ !important; font-size: 1.8rem !important; font-weight: 800 !important;
}

[data-testid="stProgressBar"] > div > div > div {
    background-color: __primary__ !important; border-radius: 10px;
}

[data-testid="stDataFrame"] {
    background: #ffffff; border-radius: 12px;
    border: 1px solid __card_border__; padding: 6px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03);
}

.stButton > button {
    background: __accent__ !important; color: __input_background__ !important;
    border-radius: 8px !important; border: none !important;
    font-weight: 800 !important; transition: all 0.2s ease;
}
.stButton > button:hover {
    background: __accent_hover__ !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
}
</style>
"""


def build_css() -> str:
    css = CSS_TEMPLATE.replace("__max_width__", str(MAX_CONTENT_WIDTH))

    for key, value in THEME.items():
        css = css.replace(f"__{key}__", value)

    return css


st.markdown(build_css(), unsafe_allow_html=True)


# ============================================================
# 6. HELPER FUNCTIONS
# ============================================================

def is_image_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS


def get_image_files(folder: Path) -> list[Path]:
    if not folder.is_dir():
        return []

    try:
        candidates = folder.rglob("*") if SCAN_SUBFOLDERS else folder.iterdir()
        return [p for p in candidates if is_image_file(p)]
    except OSError:
        return []


def get_date_folder(selected_date: date) -> Path:
    return OUTPUT_DIR / selected_date.strftime(DATE_FOLDER_FORMAT)


def get_site_folders(selected_date: date) -> list[Path]:
    date_folder = get_date_folder(selected_date)

    if not date_folder.is_dir():
        return []

    try:
        return sorted(
            (p for p in date_folder.iterdir() if p.is_dir()),
            key=lambda p: p.name.lower(),
        )
    except OSError:
        return []


def get_last_received_time(images: list[Path]) -> str:
    latest_ts = None

    for img in images:
        try:
            ts = img.stat().st_mtime
        except OSError:
            continue

        if latest_ts is None or ts > latest_ts:
            latest_ts = ts

    if latest_ts is None:
        return "-"

    return datetime.fromtimestamp(latest_ts).strftime(DATETIME_DISPLAY_FORMAT)


def scan_dashboard(selected_date: date) -> list[dict]:
    rows = []

    for site_folder in get_site_folders(selected_date):

        images = get_image_files(site_folder)
        received = len(images) > 0

        rows.append(
            {
                "Site / Group": site_folder.name,
                "Expected": "Yes",
                "Photo Received": "Yes" if received else "No",
                "Status": STATUS_RECEIVED if received else STATUS_PENDING,
                "Photo Count": len(images),
                "Last Received": get_last_received_time(images),
                "Folder": str(site_folder),
            }
        )

    return rows


def scan_all_dates() -> list[str]:
    if not OUTPUT_DIR.is_dir():
        return []

    names = []

    for folder in OUTPUT_DIR.iterdir():

        if not folder.is_dir():
            continue

        try:
            datetime.strptime(folder.name, DATE_FOLDER_FORMAT)
        except ValueError:
            continue

        names.append(folder.name)

    return sorted(names, reverse=True)


# ============================================================
# 7. SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        '<div class="sidebar-brand">'
        f'<div class="sidebar-brand-title">{esc(SIDEBAR_TITLE)}</div>'
        f'<div class="sidebar-brand-sub">{esc(SIDEBAR_SUBTITLE)}</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    st.header("📊 Dashboard Settings")

    date_str = st.text_input(
        f"Monitoring Date ({DATE_INPUT_HINT})",
        value=date.today().strftime(DATE_INPUT_FORMAT),
    )

    try:
        selected_date = datetime.strptime(
            date_str.strip(), DATE_INPUT_FORMAT
        ).date()
    except ValueError:
        selected_date = date.today()
        st.error(f"Format: {DATE_INPUT_HINT}")

    if SHOW_FOLDER_INFO:
        st.divider()
        st.subheader("Folder Information")
        st.write("Output folder:")
        st.code(str(OUTPUT_DIR), language="text")
        st.write("Selected date folder:")
        st.code(str(get_date_folder(selected_date)), language="text")

    st.divider()
    st.subheader("Auto Refresh")

    auto_refresh = st.checkbox("Enable auto refresh", value=REFRESH_ENABLED)

    refresh_seconds = st.slider(
        "Refresh interval (seconds)",
        REFRESH_MIN,
        REFRESH_MAX,
        REFRESH_DEFAULT,
        REFRESH_STEP,
    )

    st.divider()

    if st.button("🔄 Refresh Now", use_container_width=True):
        st.rerun()


# ============================================================
# 8. HEADER BANNER
# ============================================================

current_time = datetime.now().strftime(TIME_DISPLAY_FORMAT)
formatted_date = selected_date.strftime(DATE_DISPLAY_FORMAT)

st.markdown(
    f"""
    <div class="app-header">
        <div class="app-header-top">
            <div class="app-brand-section">
                <div class="app-logo">{esc(LOGO_TEXT)}</div>
                <div>
                    <div class="app-company">{esc(COMPANY_NAME)}</div>
                    <div class="app-company-subtitle">{esc(TAGLINE)}</div>
                </div>
            </div>
            <div class="app-live-badge">
                <span class="app-live-dot"></span>
                {esc(LIVE_BADGE_TEXT)}
            </div>
        </div>
        <div class="app-header-divider"></div>
        <div class="app-dashboard-title">{esc(DASHBOARD_TITLE)}</div>
        <div class="app-dashboard-subtitle">{esc(DASHBOARD_SUBTITLE)}</div>
        <div class="app-info-row">
            <div class="app-info-box">
                <div class="app-info-label">MONITORING DATE</div>
                <div class="app-info-value">📅 {esc(formatted_date)}</div>
            </div>
            <div class="app-info-box">
                <div class="app-info-label">LAST UPDATED</div>
                <div class="app-info-value">🕒 {esc(current_time)}</div>
            </div>
            <div class="app-info-box">
                <div class="app-info-label">SYSTEM STATUS</div>
                <div class="app-info-value">🟢 {esc(SYSTEM_STATUS_TEXT)}</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 9. MAIN CONTENT & SCANS
# ============================================================

if not OUTPUT_DIR.exists():
    st.error("Output folder not found!")
    st.code(str(OUTPUT_DIR), language="text")
    st.info("Set paths.output_dir in config/config.json to the folder "
            "where the photo bot saves its daily output.")
    st.stop()

DATE_FOLDER = get_date_folder(selected_date)

if not DATE_FOLDER.exists():
    st.warning(f"No folder found for {formatted_date}.")
    st.code(str(DATE_FOLDER), language="text")
    st.info("Choose another date or wait until photos are received.")
    st.stop()

rows = scan_dashboard(selected_date)

total_sites = len(rows)
received_sites = sum(1 for r in rows if r["Status"] == STATUS_RECEIVED)
pending_sites = sum(1 for r in rows if r["Status"] == STATUS_PENDING)
total_photos = sum(r["Photo Count"] for r in rows)
completion_percentage = (
    received_sites / total_sites * 100 if total_sites > 0 else 0
)


# ============================================================
# 10. SUMMARY CARDS
# ============================================================

st.markdown(
    '<div class="section-heading">📈 Photo Summary</div>',
    unsafe_allow_html=True,
)

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(LABEL_TOTAL, total_sites)
with col2:
    st.metric(LABEL_RECEIVED, received_sites)
with col3:
    st.metric(LABEL_PENDING, pending_sites)
with col4:
    st.metric(LABEL_COMPLETION, f"{completion_percentage:.1f}%")

st.progress(completion_percentage / 100 if total_sites > 0 else 0)
st.caption(f"Total image files found: {total_photos}")


# ============================================================
# 11. SITE / GROUP STATUS TABLE
# ============================================================

st.markdown(
    '<div class="section-heading">📋 Site / Group Status</div>',
    unsafe_allow_html=True,
)

if not rows:
    st.warning("No site/group folders found for this date.")

else:
    filter_col1, filter_col2 = st.columns([2, 1])

    with filter_col1:
        search_text = st.text_input(
            "🔎 Search Site / Group", placeholder="Enter site name..."
        )

    with filter_col2:
        status_filter = st.selectbox(
            "Filter Status", ["All", STATUS_RECEIVED, STATUS_PENDING]
        )

    filtered_rows = rows

    if search_text.strip():
        needle = search_text.strip().lower()
        filtered_rows = [
            r for r in filtered_rows if needle in r["Site / Group"].lower()
        ]

    if status_filter != "All":
        filtered_rows = [
            r for r in filtered_rows if r["Status"] == status_filter
        ]

    st.write(f"Showing **{len(filtered_rows)}** of **{len(rows)}** groups/sites")

    display_rows = [
        {
            "Site / Group": r["Site / Group"],
            "Expected": r["Expected"],
            "Photo Received": r["Photo Received"],
            "Status": r["Status"],
            "Photo Count": r["Photo Count"],
            "Last Received": r["Last Received"],
        }
        for r in filtered_rows
    ]

    st.dataframe(
        display_rows,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Site / Group": st.column_config.TextColumn(
                "Site / Group", help="Name of the monitored site"
            ),
            "Status": st.column_config.TextColumn("Status"),
            "Photo Count": st.column_config.NumberColumn(
                "Photo Count", format="%d 📷"
            ),
        },
    )


# ============================================================
# 12. PENDING GROUPS
# ============================================================

st.divider()

st.markdown(
    '<div class="section-heading">⚠️ Pending Groups</div>',
    unsafe_allow_html=True,
)

pending_rows = [r for r in rows if r["Status"] == STATUS_PENDING]

if pending_rows:
    st.warning(f"{len(pending_rows)} group/site(s) have not received photos.")

    pending_display = []

    for r in pending_rows:
        item = {
            "Site / Group": r["Site / Group"],
            "Status": r["Status"],
            "Photo Count": r["Photo Count"],
        }

        if SHOW_FOLDER_COLUMN:
            item["Folder"] = r["Folder"]

        pending_display.append(item)

    st.dataframe(pending_display, use_container_width=True, hide_index=True)

else:
    st.success("All discovered groups/sites have received photos.")


# ============================================================
# 13. FOOTER / EXPANDERS
# ============================================================

with st.expander("📅 Available Date Folders"):
    available_dates = scan_all_dates()

    if available_dates:
        st.write(available_dates)
    else:
        st.write("No date folders found.")

if SHOW_SYSTEM_INFO:
    with st.expander("⚙️ System Information"):
        st.write("Dashboard file:")
        st.code(str(CURRENT_FILE), language="text")
        st.write("Project root:")
        st.code(str(PROJECT_ROOT), language="text")
        st.write("Config file:")
        st.code(str(CONFIG_FILE), language="text")
        st.write("Output folder:")
        st.code(str(OUTPUT_DIR), language="text")


# ============================================================
# 14. AUTO REFRESH
# ============================================================

if auto_refresh:
    time.sleep(refresh_seconds)
    st.rerun()