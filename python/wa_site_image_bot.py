"""
WhatsApp Site-Group Image Downloader Bot
=========================================

Configuration-friendly version.

Project structure expected:

    site-operations-automation/
    │
    ├── python/
    │   └── wa_site_image_bot.py
    │
    ├── config/
    │   ├── config.json
    │   ├── config.example.json
    │   ├── site_names.json
    │   └── site_names.example.json
    │
    └── data/
        └── generated files

The program reads paths and settings from config/config.json.

Relative paths in config.json are resolved relative to the project root.
Absolute paths are used exactly as supplied.

------------------------------------------------------------------------
"""

import difflib
import io
import json
import logging
import os
import re
import shutil
import textwrap
from datetime import datetime, date
from pathlib import Path


# ============================================================================
# PROJECT / CONFIGURATION
# ============================================================================

# This file is expected to be inside:
#     project_root/python/wa_site_image_bot.py
#
# Therefore:
#     PROJECT_ROOT = project_root/
PROJECT_ROOT = Path(__file__).resolve().parent.parent

CONFIG_FILE = PROJECT_ROOT / "config" / "config.json"


# ============================================================================
# LOGGING
# ============================================================================

try:
    from colorama import init as _colorama_init, Fore, Style

    _colorama_init()
    _COLOR = True

except ImportError:
    _COLOR = False

    class _NoColor:
        def __getattr__(self, _):
            return ""

    Fore = Style = _NoColor()


class _ColorFormatter(logging.Formatter):
    """Compact, colored log lines."""

    _LEVEL_STYLE = {
        logging.DEBUG: (Fore.LIGHTBLACK_EX, "DEBUG"),
        logging.INFO: (Fore.CYAN, "INFO "),
        logging.WARNING: (Fore.YELLOW, "WARN "),
        logging.ERROR: (Fore.RED, "ERROR"),
        logging.CRITICAL: (Fore.RED + Style.BRIGHT, "CRIT "),
    }

    def format(self, record):
        color, label = self._LEVEL_STYLE.get(
            record.levelno,
            ("", record.levelname)
        )

        ts = self.formatTime(record, "%H:%M:%S")
        msg = record.getMessage()

        if record.exc_info:
            msg += " " + self.formatException(record.exc_info)

        reset = Style.RESET_ALL if _COLOR else ""

        return (
            f"{Fore.LIGHTBLACK_EX}{ts}{reset} "
            f"{color}{label}{reset} {msg}"
        )


_handler = logging.StreamHandler()
_handler.setFormatter(_ColorFormatter())

logging.basicConfig(
    level=logging.INFO,
    handlers=[_handler]
)

logger = logging.getLogger("wa-site-image-bot")


class _SuppressBenignWhatsmeowWarnings(logging.Filter):
    """
    Hides known harmless WhatsApp history-sync warnings.
    """

    _SUPPRESSED_SUBSTRINGS = (
        "Failed to delete history sync media from server",
    )

    def filter(self, record):
        msg = record.getMessage()

        return not any(
            s in msg
            for s in self._SUPPRESSED_SUBSTRINGS
        )


_handler.addFilter(_SuppressBenignWhatsmeowWarnings())


# ============================================================================
# CONFIGURATION HELPERS
# ============================================================================

def _load_json_config(path: Path) -> dict:
    """
    Load JSON configuration.

    If config.json does not exist, return an empty dictionary.
    """

    if not path.is_file():
        logger.warning(
            "Configuration file not found: %s",
            path
        )

        logger.warning(
            "Using built-in defaults."
        )

        return {}

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, dict):
            raise ValueError(
                "Configuration root must be a JSON object."
            )

        return data

    except Exception:
        logger.exception(
            "Failed to load configuration file: %s",
            path
        )

        return {}


CONFIG = _load_json_config(CONFIG_FILE)


def _config_section(name: str) -> dict:
    value = CONFIG.get(name, {})

    return value if isinstance(value, dict) else {}


PATH_CONFIG = _config_section("paths")
WHATSAPP_CONFIG = _config_section("whatsapp")
BACKFILL_CONFIG = _config_section("backfill")
TEXT_CONFIG = _config_section("text_capture")
OCR_CONFIG = _config_section("ocr")
FONT_CONFIG = _config_section("fonts")


def _resolve_config_path(value, default: str) -> Path:
    """
    Convert a config path into an absolute Path.

    Rules:

    1. Empty / missing value -> default.
    2. Absolute path -> use as-is.
    3. Relative path -> relative to PROJECT_ROOT.
    """

    if value is None or str(value).strip() == "":
        value = default

    path = Path(str(value))

    if path.is_absolute():
        return path

    return PROJECT_ROOT / path


# ============================================================================
# CONFIGURED PATHS
# ============================================================================

OUTPUT_DIR = _resolve_config_path(
    PATH_CONFIG.get("output_dir"),
    "data/output"
)

SESSION_DB = _resolve_config_path(
    PATH_CONFIG.get("session_db"),
    "data/session.sqlite3"
)

SITE_NAMES_FILE = _resolve_config_path(
    PATH_CONFIG.get("site_names_file"),
    "config/site_names.json"
)

SITE_NAMES_EXAMPLE_FILE = _resolve_config_path(
    PATH_CONFIG.get("site_names_example_file"),
    "config/site_names.example.json"
)

GROUP_OVERRIDES_FILE = _resolve_config_path(
    PATH_CONFIG.get("group_overrides_file"),
    "data/group_folder_overrides.json"
)

SITE_GROUP_FOLDERS_FILE = _resolve_config_path(
    PATH_CONFIG.get("site_group_folders_file"),
    "data/site_group_folders.json"
)

NUMBERING_FILE = _resolve_config_path(
    PATH_CONFIG.get("numbering_file"),
    "data/site_numbering.json"
)


# ============================================================================
# WHATSAPP GROUP FILTER
# ============================================================================

ONLY_THESE_GROUPS = WHATSAPP_CONFIG.get(
    "only_these_groups"
)

if ONLY_THESE_GROUPS is not None:

    if not isinstance(ONLY_THESE_GROUPS, list):
        logger.warning(
            "'only_these_groups' must be a JSON list or null. "
            "Ignoring invalid value."
        )

        ONLY_THESE_GROUPS = None


# ============================================================================
# OCR / TESSERACT
# ============================================================================

def _find_tesseract():
    """
    Find Tesseract automatically.

    Priority:

    1. config.json path
    2. PATH environment
    3. Windows Program Files locations
    """

    configured = OCR_CONFIG.get("tesseract_path")

    if configured:
        configured_path = Path(str(configured))

        if not configured_path.is_absolute():
            configured_path = PROJECT_ROOT / configured_path

        if configured_path.is_file():
            return configured_path

        logger.warning(
            "Configured Tesseract path does not exist: %s",
            configured_path
        )

    # Search PATH.
    discovered = shutil.which("tesseract")

    if discovered:
        return Path(discovered)

    # Windows common installation locations.
    if os.name == "nt":

        program_files = os.environ.get("ProgramFiles")

        if program_files:
            candidate = (
                Path(program_files)
                / "Tesseract-OCR"
                / "tesseract.exe"
            )

            if candidate.is_file():
                return candidate

        program_files_x86 = os.environ.get(
            "ProgramFiles(x86)"
        )

        if program_files_x86:
            candidate = (
                Path(program_files_x86)
                / "Tesseract-OCR"
                / "tesseract.exe"
            )

            if candidate.is_file():
                return candidate

    return None


# ============================================================================
# PIL / OCR
# ============================================================================

from PIL import Image, ImageOps, ImageDraw, ImageFont

try:
    import pytesseract
except ImportError:
    pytesseract = None


TESSERACT_PATH = _find_tesseract()

if pytesseract is not None and TESSERACT_PATH:
    pytesseract.pytesseract.tesseract_cmd = str(
        TESSERACT_PATH
    )

    logger.info(
        "Tesseract found: %s",
        TESSERACT_PATH
    )

elif pytesseract is not None:
    logger.warning(
        "Tesseract was not found. "
        "OCR orientation detection will be unavailable."
    )


# ============================================================================
# NEONIZE
# ============================================================================

from neonize.client import NewClient
from neonize.events import ConnectedEv, MessageEv, HistorySyncEv


# ============================================================================
# SITE NAMES
# ============================================================================

_FALLBACK_EXAMPLE_SITE_NAMES = [
    "Site-A",
    "Site-B",
    "Warehouse-1",
    "Depot-2"
]


def _load_site_names() -> list:
    """
    Load real site names from configured site_names.json.

    If the real file is unavailable, use site_names.example.json.
    """

    if SITE_NAMES_FILE.is_file():

        try:

            with open(
                SITE_NAMES_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                names = json.load(f)

            if isinstance(names, list):
                return names

            logger.warning(
                "%s does not contain a JSON list.",
                SITE_NAMES_FILE
            )

        except Exception:

            logger.exception(
                "Failed to load %s",
                SITE_NAMES_FILE
            )

    if SITE_NAMES_EXAMPLE_FILE.is_file():

        try:

            with open(
                SITE_NAMES_EXAMPLE_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                names = json.load(f)

            logger.info(
                "%s not found - using example site names "
                "from %s.",
                SITE_NAMES_FILE,
                SITE_NAMES_EXAMPLE_FILE
            )

            if isinstance(names, list):
                return names

        except Exception:

            logger.exception(
                "Failed to load %s",
                SITE_NAMES_EXAMPLE_FILE
            )

    logger.warning(
        "No site-name file found. "
        "Using built-in example names."
    )

    return list(_FALLBACK_EXAMPLE_SITE_NAMES)


def _save_site_names(names: list):

    try:

        SITE_NAMES_FILE.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        with open(
            SITE_NAMES_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                names,
                f,
                indent=2,
                ensure_ascii=False
            )

    except Exception:

        logger.exception(
            "Failed to save %s",
            SITE_NAMES_FILE
        )


SITE_NAMES = _load_site_names()

_SITE_NAMES_SORTED = sorted(
    set(SITE_NAMES),
    key=len,
    reverse=True
)


def match_site_name(group_name: str):
    """
    Return the SITE_NAMES entry contained in group_name.
    """

    gn_lower = (
        group_name or ""
    ).lower()

    for site in _SITE_NAMES_SORTED:

        if site.lower() in gn_lower:
            return site

    return None


# ============================================================================
# GROUP OVERRIDES
# ============================================================================

def _load_group_overrides() -> dict:

    if GROUP_OVERRIDES_FILE.is_file():

        try:

            with open(
                GROUP_OVERRIDES_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                return json.load(f)

        except Exception:

            logger.exception(
                "Failed to load %s - starting fresh",
                GROUP_OVERRIDES_FILE
            )

    return {}


def _save_group_overrides(overrides: dict):

    try:

        GROUP_OVERRIDES_FILE.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        with open(
            GROUP_OVERRIDES_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                overrides,
                f,
                indent=2,
                ensure_ascii=False
            )

    except Exception:

        logger.exception(
            "Failed to save %s",
            GROUP_OVERRIDES_FILE
        )


_group_overrides = _load_group_overrides()


def _add_site_name_permanently(new_name: str):

    global SITE_NAMES
    global _SITE_NAMES_SORTED

    if new_name in SITE_NAMES:
        return

    SITE_NAMES.append(new_name)

    _SITE_NAMES_SORTED = sorted(
        set(SITE_NAMES),
        key=len,
        reverse=True
    )

    _save_site_names(SITE_NAMES)

    logger.info(
        "Added '%s' to %s",
        new_name,
        SITE_NAMES_FILE
    )


def resolve_folder_name(group_name: str):

    matched = match_site_name(group_name)

    if matched:
        return matched

    if group_name in _group_overrides:
        return _group_overrides[group_name]

    print("\n" + "=" * 60)

    print(
        "  New WhatsApp group - no match in your site list:"
    )

    print(
        f'    "{group_name}"'
    )

    print(
        "  Type the site/folder name to use for this group's images,"
    )

    print(
        "  or just press Enter to SKIP this chat "
        "(won't ask again)."
    )

    print("=" * 60)

    _root_logger = logging.getLogger()

    previous_level = _root_logger.level

    _root_logger.setLevel(logging.ERROR)

    try:

        answer = input(
            "  Folder name (blank = skip): "
        ).strip()

    except EOFError:

        answer = ""

    finally:

        _root_logger.setLevel(previous_level)

    chosen = answer if answer else None

    if chosen:

        existing = next(
            (
                s
                for s in SITE_NAMES
                if s.lower() == chosen.lower()
            ),
            None
        )

        if existing:
            chosen = existing

    _group_overrides[group_name] = chosen

    _save_group_overrides(
        _group_overrides
    )

    if chosen:

        _add_site_name_permanently(
            chosen
        )

        logger.info(
            "Group '%s' -> saving under '%s' "
            "(remembered)",
            group_name,
            chosen
        )

    else:

        logger.info(
            "Group '%s' -> SKIPPING from now on "
            "(remembered)",
            group_name
        )

    return chosen


# ============================================================================
# ONE FOLDER PER DISTINCT GROUP
# ============================================================================

def _load_site_group_folders() -> dict:

    if SITE_GROUP_FOLDERS_FILE.is_file():

        try:

            with open(
                SITE_GROUP_FOLDERS_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                return json.load(f)

        except Exception:

            logger.exception(
                "Failed to load %s - starting fresh",
                SITE_GROUP_FOLDERS_FILE
            )

    return {}


def _save_site_group_folders(data: dict):

    try:

        SITE_GROUP_FOLDERS_FILE.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        with open(
            SITE_GROUP_FOLDERS_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                data,
                f,
                indent=2,
                ensure_ascii=False
            )

    except Exception:

        logger.exception(
            "Failed to save %s",
            SITE_GROUP_FOLDERS_FILE
        )


_site_group_folders = _load_site_group_folders()


def assign_top_folder(
    site_key: str,
    group_name: str
) -> str:

    groups_for_site = _site_group_folders.setdefault(
        site_key,
        {}
    )

    if group_name in groups_for_site:
        return groups_for_site[group_name]

    slot = len(groups_for_site) + 1

    top_folder = (
        site_key
        if slot == 1
        else f"{site_key}-{slot}"
    )

    groups_for_site[group_name] = top_folder

    _save_site_group_folders(
        _site_group_folders
    )

    logger.info(
        "Group '%s' -> new site folder '%s' "
        "(site '%s')",
        group_name,
        top_folder,
        site_key
    )

    return top_folder


# ============================================================================
# TEXT MESSAGE CONFIGURATION
# ============================================================================

TEXT_CAPTURE_KEYWORDS = TEXT_CONFIG.get(
    "keywords",
    ["diesel"]
)

if not isinstance(
    TEXT_CAPTURE_KEYWORDS,
    list
):
    TEXT_CAPTURE_KEYWORDS = ["diesel"]


try:

    TEXT_CAPTURE_FUZZY_THRESHOLD = float(
        TEXT_CONFIG.get(
            "fuzzy_threshold",
            0.75
        )
    )

except (TypeError, ValueError):

    TEXT_CAPTURE_FUZZY_THRESHOLD = 0.75


# ============================================================================
# OCR FUNCTIONS
# ============================================================================

def _ocr_confidence(
    img: Image.Image
) -> float:

    if pytesseract is None:
        return -1.0

    data = pytesseract.image_to_data(
        img,
        output_type=pytesseract.Output.DICT
    )

    confs = [
        int(c)
        for c in data.get("conf", [])
        if str(c).lstrip("-").isdigit()
        and int(c) >= 0
    ]

    return (
        sum(confs) / len(confs)
        if confs
        else -1.0
    )


def detect_rotation_angle(
    img: Image.Image
):

    if pytesseract is None:
        return None

    try:

        osd = pytesseract.image_to_osd(img)

        for line in osd.splitlines():

            if line.startswith("Rotate:"):

                return int(
                    line.split(":")[1].strip()
                )

    except Exception:

        logger.debug(
            "OCR orientation detection inconclusive"
        )

    return None


def pick_rotation_direction(
    option_left: Image.Image,
    option_right: Image.Image
) -> Image.Image:

    try:

        if pytesseract is None:
            raise RuntimeError(
                "pytesseract not installed"
            )

        conf_left = _ocr_confidence(
            option_left
        )

        conf_right = _ocr_confidence(
            option_right
        )

        if conf_right > conf_left:
            return option_right

    except Exception:

        logger.debug(
            "OCR direction tie-breaker skipped"
        )

    return option_left


def correct_orientation(
    data: bytes
) -> bytes:

    try:

        img = Image.open(
            io.BytesIO(data)
        )

        img = ImageOps.exif_transpose(
            img
        )

        angle = detect_rotation_angle(
            img
        )

        if angle:

            img = img.rotate(
                -angle,
                expand=True
            )

        elif (
            angle is None
            and img.width > img.height
        ):

            option_left = img.rotate(
                90,
                expand=True
            )

            option_right = img.rotate(
                -90,
                expand=True
            )

            img = pick_rotation_direction(
                option_left,
                option_right
            )

        buf = io.BytesIO()

        save_format = (
            img.format or "JPEG"
        ).upper()

        if save_format in (
            "JPEG",
            "JPG"
        ):

            img = img.convert("RGB")

            img.save(
                buf,
                format="JPEG",
                quality=95
            )

        else:

            img.save(
                buf,
                format=save_format
            )

        return buf.getvalue()

    except Exception:

        logger.exception(
            "Orientation correction failed - "
            "saving original image unchanged"
        )

        return data


# ============================================================================
# TEXT MESSAGE CAPTURE
# ============================================================================

def _is_fuzzy_keyword_match(
    word: str,
    keyword: str
) -> bool:

    word = word.lower().strip(
        ".,:;!?()[]{}\"'"
    )

    if not word:
        return False

    if word == keyword:
        return True

    return (
        difflib.SequenceMatcher(
            None,
            word,
            keyword
        ).ratio()
        >= TEXT_CAPTURE_FUZZY_THRESHOLD
    )


def text_contains_keyword(
    text: str
) -> bool:

    if not text:
        return False

    for word in re.findall(
        r"[A-Za-z]+",
        text
    ):

        for keyword in TEXT_CAPTURE_KEYWORDS:

            if _is_fuzzy_keyword_match(
                word,
                keyword
            ):
                return True

    return False


# ============================================================================
# TEXT CARD FONTS
# ============================================================================

_text_card_font = None
_text_card_font_bold = None


def _find_font(
    configured_path,
    candidates,
    size: int
):

    if configured_path:

        configured = Path(
            str(configured_path)
        )

        if not configured.is_absolute():
            configured = PROJECT_ROOT / configured

        if configured.is_file():

            try:
                return ImageFont.truetype(
                    str(configured),
                    size
                )

            except Exception:
                logger.warning(
                    "Could not load configured font: %s",
                    configured
                )

    for candidate in candidates:

        candidate_path = Path(candidate)

        if candidate_path.is_file():

            try:

                return ImageFont.truetype(
                    str(candidate_path),
                    size
                )

            except Exception:
                continue

    return ImageFont.load_default()


def _load_text_card_fonts():

    global _text_card_font
    global _text_card_font_bold

    if _text_card_font is not None:
        return

    regular_config = FONT_CONFIG.get(
        "regular_path"
    )

    bold_config = FONT_CONFIG.get(
        "bold_path"
    )

    regular_candidates = []

    bold_candidates = []

    # Windows
    if os.name == "nt":

        windir = os.environ.get(
            "WINDIR"
        )

        if windir:

            fonts_dir = Path(
                windir
            ) / "Fonts"

            regular_candidates.extend([
                fonts_dir / "segoeui.ttf",
                fonts_dir / "arial.ttf",
            ])

            bold_candidates.extend([
                fonts_dir / "segoeuib.ttf",
                fonts_dir / "arialbd.ttf",
            ])

    # Linux
    regular_candidates.extend([
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ])

    bold_candidates.extend([
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ])

    # macOS
    regular_candidates.extend([
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ])

    bold_candidates.extend([
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    ])

    _text_card_font = _find_font(
        regular_config,
        regular_candidates,
        30
    )

    _text_card_font_bold = _find_font(
        bold_config,
        bold_candidates,
        32
    )


def render_text_message_image(
    sender_name: str,
    text: str,
    msg_dt: datetime
) -> bytes:

    _load_text_card_fonts()

    width = 900
    margin = 30

    wrapped_lines = []

    for paragraph in (
        text.splitlines() or [""]
    ):

        wrapped_lines.extend(
            textwrap.wrap(
                paragraph,
                width=48
            ) or [""]
        )

    line_height = 40
    header_height = 70

    height = (
        header_height
        + margin * 2
        + line_height
        * max(len(wrapped_lines), 1)
    )

    img = Image.new(
        "RGB",
        (width, height),
        "white"
    )

    draw = ImageDraw.Draw(img)

    header = (
        f"{sender_name or 'Unknown sender'}"
        f"  -  "
        f"{msg_dt.strftime('%d-%b-%Y %I:%M %p')}"
    )

    draw.text(
        (margin, margin),
        header,
        fill=(30, 90, 40),
        font=_text_card_font_bold
    )

    draw.line(
        [
            (margin, header_height),
            (width - margin, header_height)
        ],
        fill=(220, 220, 220),
        width=2
    )

    y = header_height + 15

    for line in wrapped_lines:

        draw.text(
            (margin, y),
            line,
            fill=(20, 20, 20),
            font=_text_card_font
        )

        y += line_height

    buf = io.BytesIO()

    img.save(
        buf,
        format="JPEG",
        quality=95
    )

    return buf.getvalue()


# ============================================================================
# FILE NUMBERING
# ============================================================================

def _load_numbering() -> dict:

    if NUMBERING_FILE.is_file():

        try:

            with open(
                NUMBERING_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                return json.load(f)

        except Exception:

            logger.exception(
                "Failed to load %s - starting fresh",
                NUMBERING_FILE
            )

    return {}


def _save_numbering(data: dict):

    try:

        NUMBERING_FILE.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        with open(
            NUMBERING_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                data,
                f,
                indent=2,
                ensure_ascii=False
            )

    except Exception:

        logger.exception(
            "Failed to save %s",
            NUMBERING_FILE
        )


_numbering = _load_numbering()


def next_time_based_name(
    folder_key: str,
    msg_id: str,
    time_label: str
) -> str:

    folder_map = _numbering.setdefault(
        folder_key,
        {}
    )

    if msg_id in folder_map:
        return folder_map[msg_id]

    existing_for_label = [
        v
        for v in folder_map.values()
        if isinstance(v, str)
        and (
            v == time_label
            or v.startswith(
                time_label + "-"
            )
        )
    ]

    label = (
        time_label
        if not existing_for_label
        else f"{time_label}-{len(existing_for_label) + 1}"
    )

    folder_map[msg_id] = label

    _save_numbering(
        _numbering
    )

    return label


# ============================================================================
# IMAGE SAVING
# ============================================================================

def sanitize(name: str) -> str:

    name = (
        name or ""
    ).strip()

    name = re.sub(
        r"[\r\n\t]+",
        " ",
        name
    )

    name = re.sub(
        r'[\\/*?:"<>|]',
        "_",
        name
    )

    name = re.sub(
        r"\s+",
        " ",
        name
    ).strip()

    name = name[:100]

    return name or "Unknown Group"


def guess_extension(
    mimetype: str
) -> str:

    mapping = {
        "image/jpeg": ".jpg",
        "image/png": ".png",
        "image/webp": ".webp",
        "image/gif": ".gif",
    }

    return mapping.get(
        (mimetype or "")
        .split(";")[0]
        .strip(),
        ".jpg"
    )


def save_image(
    top_folder: str,
    filename_prefix: str,
    msg_dt: datetime,
    mimetype: str,
    data: bytes,
    msg_id: str = None
) -> Path:

    top = sanitize(
        top_folder
    )

    date_str = (
        msg_dt.date().isoformat()
    )

    day_folder = (
        OUTPUT_DIR
        / date_str
        / top
    )

    folder_key = (
        f"{date_str}/{top}"
    )

    day_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    key = (
        msg_id
        or f"t{msg_dt.strftime('%H%M%S%f')}"
    )

    time_label = msg_dt.strftime(
        "%H-%M"
    )

    label = next_time_based_name(
        folder_key,
        key,
        time_label
    )

    filename = (
        f"{sanitize(filename_prefix)}"
        f"-{label}"
        f"{guess_extension(mimetype)}"
    )

    path = day_folder / filename

    with open(
        path,
        "wb"
    ) as f:

        f.write(data)

    return path


# ============================================================================
# WHATSAPP MEDIA DOWNLOAD
# ============================================================================

def download_media_bytes(
    client: NewClient,
    proto_message
) -> bytes:

    """
    Pull raw image bytes from a WhatsApp Message proto.

    Supports neonize versions exposing either download_any
    or download_media.
    """

    if hasattr(
        client,
        "download_any"
    ):

        return client.download_any(
            proto_message
        )

    if hasattr(
        client,
        "download_media"
    ):

        return client.download_media(
            proto_message
        )

    raise RuntimeError(
        "Could not find a download method on the client. "
        "Check the installed neonize version."
    )


# ============================================================================
# WHATSAPP CLIENT
# ============================================================================

# Make sure the parent folder exists before opening the database.
SESSION_DB.parent.mkdir(
    parents=True,
    exist_ok=True
)

client = NewClient(
    str(SESSION_DB)
)


# ============================================================================
# BACKFILL
# ============================================================================

BACKFILL_START_DATE = None


@client.event(ConnectedEv)
def on_connected(
    client: NewClient,
    event: ConnectedEv
):

    logger.info(
        "Connected to WhatsApp. "
        "Watching groups for images..."
    )

    if BACKFILL_START_DATE is not None:

        logger.info(
            "Backfill is ON, from %s through now "
            "(then live).",
            BACKFILL_START_DATE.isoformat()
        )

    else:

        logger.info(
            "Backfill is OFF - only new images "
            "from now on will be saved."
        )


# ============================================================================
# HISTORY SYNC / BACKFILL
# ============================================================================

@client.event(HistorySyncEv)
def on_history_sync(
    client: NewClient,
    event: HistorySyncEv
):

    if BACKFILL_START_DATE is None:
        return

    try:

        conversations = getattr(
            getattr(
                event,
                "Data",
                None
            ),
            "conversations",
            None
        )

        if not conversations:
            return

        saved_count = 0

        for conv in conversations:

            group_name = (
                getattr(
                    conv,
                    "name",
                    None
                )
                or
                f"UnknownGroup_"
                f"{getattr(conv, 'ID', 'unknown')}"
            )

            if (
                ONLY_THESE_GROUPS
                and
                group_name.strip().upper()
                not in {
                    g.strip().upper()
                    for g in ONLY_THESE_GROUPS
                }
            ):
                continue

            backlog_items = []

            for hist_msg in getattr(
                conv,
                "messages",
                []
            ):

                wmi = getattr(
                    hist_msg,
                    "message",
                    None
                )

                if wmi is None:
                    continue

                proto_msg = getattr(
                    wmi,
                    "message",
                    None
                )

                if proto_msg is None:
                    continue

                is_image = proto_msg.HasField(
                    "imageMessage"
                )

                text_body = None

                if not is_image:

                    if proto_msg.HasField(
                        "conversation"
                    ):

                        text_body = (
                            proto_msg.conversation
                        )

                    elif proto_msg.HasField(
                        "extendedTextMessage"
                    ):

                        text_body = (
                            proto_msg
                            .extendedTextMessage
                            .text
                        )

                is_keyword_text = (
                    bool(text_body)
                    and text_contains_keyword(
                        text_body
                    )
                )

                if (
                    not is_image
                    and not is_keyword_text
                ):
                    continue

                ts = getattr(
                    wmi,
                    "messageTimestamp",
                    None
                )

                if not ts:
                    continue

                msg_dt = datetime.fromtimestamp(
                    ts
                )

                if (
                    msg_dt.date()
                    < BACKFILL_START_DATE
                ):
                    continue

                backlog_items.append(
                    (
                        wmi,
                        proto_msg,
                        msg_dt,
                        ts,
                        is_image,
                        text_body
                    )
                )

            if not backlog_items:
                continue

            site_key = resolve_folder_name(
                group_name
            )

            if site_key is None:
                continue

            top_folder = assign_top_folder(
                site_key,
                group_name
            )

            for (
                wmi,
                proto_msg,
                msg_dt,
                ts,
                is_image,
                text_body
            ) in backlog_items:

                try:

                    msg_id = (
                        getattr(
                            getattr(
                                wmi,
                                "key",
                                None
                            ),
                            "ID",
                            ""
                        )
                        or
                        str(ts)
                    )

                    if is_image:

                        image_msg = (
                            proto_msg.imageMessage
                        )

                        data = (
                            download_media_bytes(
                                client,
                                proto_msg
                            )
                        )

                        data = correct_orientation(
                            data
                        )

                        mimetype = getattr(
                            image_msg,
                            "mimetype",
                            ""
                        )

                        note = ""

                    else:

                        sender_name = getattr(
                            wmi,
                            "pushName",
                            None
                        )

                        data = (
                            render_text_message_image(
                                sender_name,
                                text_body,
                                msg_dt
                            )
                        )

                        mimetype = "image/jpeg"

                        note = "  [text report]"

                    saved_path = save_image(
                        top_folder,
                        site_key,
                        msg_dt,
                        mimetype,
                        data,
                        msg_id=msg_id
                    )

                    saved_count += 1

                    tick = (
                        f"{Fore.GREEN}OK"
                        f"{Style.RESET_ALL}"
                        if _COLOR
                        else "OK"
                    )

                    logger.info(
                        "%s %-22s %s%s",
                        tick,
                        top_folder,
                        saved_path.name,
                        note
                    )

                except Exception:

                    logger.exception(
                        "Failed to backfill one item "
                        "from group '%s'",
                        group_name
                    )

        if saved_count:

            logger.info(
                "%s Backfill complete - %d image(s) "
                "from %s onward.",
                (
                    Fore.GREEN
                    + "=="
                    + Style.RESET_ALL
                )
                if _COLOR
                else "==",
                saved_count,
                BACKFILL_START_DATE.isoformat()
            )

    except Exception:

        logger.exception(
            "Failed to process HistorySyncEv "
            "for backfill"
        )


# ============================================================================
# LIVE MESSAGE HANDLER
# ============================================================================

@client.event(MessageEv)
def on_message(
    client: NewClient,
    message: MessageEv
):

    try:

        chat_jid = (
            message
            .Info
            .MessageSource
            .Chat
        )

        if (
            getattr(
                chat_jid,
                "Server",
                ""
            )
            != "g.us"
        ):
            return

        msg = message.Message

        is_image = msg.HasField(
            "imageMessage"
        )

        text_body = None

        if not is_image:

            if msg.HasField(
                "conversation"
            ):

                text_body = (
                    msg.conversation
                )

            elif msg.HasField(
                "extendedTextMessage"
            ):

                text_body = (
                    msg.extendedTextMessage.text
                )

        is_keyword_text = (
            bool(text_body)
            and text_contains_keyword(
                text_body
            )
        )

        if (
            not is_image
            and not is_keyword_text
        ):
            return

        image_msg = (
            msg.imageMessage
            if is_image
            else None
        )

        group_name = None

        try:

            group_info = (
                client.get_group_info(
                    chat_jid
                )
            )

            group_name = getattr(
                getattr(
                    group_info,
                    "GroupName",
                    None
                ),
                "Name",
                None
            )

            if not group_name:

                logger.warning(
                    "get_group_info returned no "
                    "GroupName.Name - raw object: %r",
                    group_info
                )

        except Exception as e:

            logger.warning(
                "get_group_info failed for %s: %s",
                getattr(
                    chat_jid,
                    "User",
                    chat_jid
                ),
                e
            )

        if not group_name:

            group_name = (
                f"UnknownGroup_"
                f"{getattr(chat_jid, 'User', 'unknown')}"
            )

        site_key = resolve_folder_name(
            group_name
        )

        if site_key is None:
            return

        if (
            ONLY_THESE_GROUPS
            and
            group_name.strip().upper()
            not in {
                g.strip().upper()
                for g in ONLY_THESE_GROUPS
            }
        ):
            return

        top_folder = assign_top_folder(
            site_key,
            group_name
        )

        ts = getattr(
            message.Info,
            "Timestamp",
            None
        )

        msg_dt = (
            ts
            if isinstance(ts, datetime)
            else datetime.now()
        )

        msg_id = getattr(
            message.Info,
            "ID",
            None
        )

        if is_image:

            data = download_media_bytes(
                client,
                msg
            )

            data = correct_orientation(
                data
            )

            mimetype = getattr(
                image_msg,
                "mimetype",
                ""
            )

            note = ""

        else:

            sender_name = getattr(
                message.Info,
                "PushName",
                None
            )

            data = render_text_message_image(
                sender_name,
                text_body,
                msg_dt
            )

            mimetype = "image/jpeg"

            note = "  [text report]"

        saved_path = save_image(
            top_folder,
            site_key,
            msg_dt,
            mimetype,
            data,
            msg_id=msg_id
        )

        tick = (
            f"{Fore.GREEN}OK"
            f"{Style.RESET_ALL}"
            if _COLOR
            else "OK"
        )

        logger.info(
            "%s %-22s %s  (new)%s",
            tick,
            top_folder,
            saved_path.name,
            note
        )

    except Exception:

        logger.exception(
            "Failed to process an incoming message"
        )


# ============================================================================
# BACKFILL PROMPT
# ============================================================================

def _prompt_backfill_start_date():

    global BACKFILL_START_DATE

    print("\n" + "=" * 58)

    print(
        "  Backfill from history?"
    )

    print(
        "  Press Enter for TODAY's images only (default)."
    )

    print(
        "  Or type a start date (YYYY-MM-DD) to pull every"
    )

    print(
        "  image from that date through now."
    )

    print(
        "  Type 'skip' to disable backfill."
    )

    print("=" * 58)

    try:

        answer = input(
            "  Backfill start date (blank = today): "
        ).strip()

    except EOFError:

        answer = ""

    if not answer:

        BACKFILL_START_DATE = date.today()

        logger.info(
            "No date entered - backfilling today (%s), "
            "then live.",
            BACKFILL_START_DATE.isoformat()
        )

        return

    if answer.lower() in (
        "skip",
        "none",
        "no"
    ):

        BACKFILL_START_DATE = None

        logger.info(
            "Backfill disabled - capturing new "
            "images only."
        )

        return

    try:

        BACKFILL_START_DATE = (
            datetime.strptime(
                answer,
                "%Y-%m-%d"
            ).date()
        )

        logger.info(
            "Backfill enabled from %s through now, "
            "then live.",
            BACKFILL_START_DATE.isoformat()
        )

    except ValueError:

        logger.warning(
            "Couldn't parse '%s' as YYYY-MM-DD - "
            "backfilling today instead.",
            answer
        )

        BACKFILL_START_DATE = date.today()


# ============================================================================
# STARTUP
# ============================================================================

def _print_banner():

    if os.name == "nt":

        os.system(
            "title WhatsApp Site Image Bot"
        )

    line = "=" * 58

    c = (
        Fore.CYAN
        if _COLOR
        else ""
    )

    r = (
        Style.RESET_ALL
        if _COLOR
        else ""
    )

    print(
        f"\n{c}{line}"
    )

    print(
        "   WhatsApp Site Image Bot"
    )

    print(
        f"   Project: {PROJECT_ROOT}"
    )

    print(
        f"   Saving into: {OUTPUT_DIR.resolve()}"
    )

    print(
        f"{line}{r}\n"
    )


if __name__ == "__main__":

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    _print_banner()

    _prompt_backfill_start_date()

    client.connect()