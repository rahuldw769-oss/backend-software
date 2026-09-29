from fastapi import FastAPI, HTTPException, Dependsfrom fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

from datetime import datetime, date
import os
import hmac
import hashlib
import base64
import json
import requests
import calendar
import random


# =========================================================
# APP
# =========================================================

app = FastAPI(
    title="Rahul Software API"
)


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


security = HTTPBearer(
    auto_error=False
)


# =========================================================
# LOGIN / ENVIRONMENT
# =========================================================

USERNAME = os.getenv(
    "RAHUL_USERNAME",
    "rahul"
)

PASSWORD = os.getenv(
    "RAHUL_PASSWORD",
    "rahul123"
)

SECRET = os.getenv(
    "RAHUL_SECRET",
    "change-this-secret"
)


SUPABASE_URL = os.getenv(
    "SUPABASE_URL"
)

SUPABASE_KEY = os.getenv(
    "SUPABASE_KEY"
)


# =========================================================
# LOCAL AI - OLLAMA
# =========================================================

OLLAMA_URL = os.getenv(
    "OLLAMA_URL",
    "http://127.0.0.1:11434"
)

OLLAMA_MODEL = os.getenv(
    "OLLAMA_MODEL",
    "qwen3:4b"
)


TABLE = "excel_rows"


# =========================================================
# INDEPENDENT FILE SNAPSHOTS
# =========================================================
# These sources are intentionally kept separate from the existing
# 2026-27 Supabase data.  FY/date logic below is not changed.
# A snapshot is made from displayed Excel values and then served
# independently, so the UI does not need to open the Excel file.

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

STOCK_SNAPSHOT_FILE = os.path.join(
    DATA_DIR, "stock_snapshot.json"
)

TRANSPORT_SNAPSHOT_FILE = os.path.join(
    DATA_DIR, "transport_snapshot.json"
)

OLD_2025_26_SNAPSHOT_FILE = os.path.join(
    DATA_DIR, "2025-26_snapshot.json"
)


def _read_snapshot(path):

    if not os.path.exists(path):
        return None

    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _excel_source_candidates(*names):

    candidates = []

    for name in names:
        candidates.extend([
            os.path.join(BASE_DIR, name),
            os.path.join(DATA_DIR, name),
        ])

    return candidates


def _find_existing_file(*names):

    for path in _excel_source_candidates(*names):
        if os.path.isfile(path):
            return path

    return None


def _excel_value(value):

    if isinstance(value, (datetime, date)):
        return value.isoformat()

    if value is None:
        return None

    return value


def _normalise_sheet_name(name):
    return " ".join(str(name or "").strip().upper().split())


def _find_sheet_name(wb, wanted_name):
    """Find an Excel sheet by name without changing its original casing."""
    wanted = _normalise_sheet_name(wanted_name)
    for name in wb.sheetnames:
        if _normalise_sheet_name(name) == wanted:
            return name
    return None


def _style_snapshot(cell):
    """Return only presentation information needed to draw Excel-like cells."""
    def rgb(value):
        if not value:
            return None
        return getattr(value, "rgb", None) or getattr(value, "indexed", None) or getattr(value, "theme", None)

    font = cell.font
    fill = cell.fill
    alignment = cell.alignment
    border = cell.border

    def side_data(side):
        return {
            "style": side.style,
            "color": rgb(side.color),
        }

    return {
        "font": {
            "name": font.name,
            "size": font.sz,
            "bold": bool(font.bold),
            "italic": bool(font.italic),
            "underline": font.underline,
            "strike": bool(font.strike),
            "color": rgb(font.color),
        },
        "fill": {
            "type": fill.fill_type,
            "fg": rgb(fill.fgColor),
            "bg": rgb(fill.bgColor),
        },
        "alignment": {
            "horizontal": alignment.horizontal,
            "vertical": alignment.vertical,
            "wrap_text": bool(alignment.wrap_text),
            "text_rotation": alignment.text_rotation,
            "shrink_to_fit": bool(alignment.shrink_to_fit),
            "indent": alignment.indent,
        },
        "border": {
            "left": side_data(border.left),
            "right": side_data(border.right),
            "top": side_data(border.top),
            "bottom": side_data(border.bottom),
        },
        "number_format": cell.number_format,
    }


def _make_sheet_snapshot(path, sheet_name=None):
    """Create a static, Excel-like snapshot of one sheet.

    Important: values are read with data_only=True. Therefore formula cells
    contribute their cached/displayed Excel values, not formulas that the
    backend tries to recalculate. The JSON snapshot becomes the runtime
    source; Excel is not opened during normal API reads.
    """
    try:
        from openpyxl import load_workbook
    except Exception as exc:
        raise RuntimeError(
            "openpyxl is required to create an Excel snapshot: " + str(exc)
        )

    wb = load_workbook(
        path,
        data_only=True,
        read_only=False,
        keep_vba=path.lower().endswith(".xlsm")
    )

    actual_sheet = _find_sheet_name(wb, sheet_name) if sheet_name else None
    ws = wb[actual_sheet] if actual_sheet else wb.active

    cells = []
    styles = {}
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0

    for row in ws.iter_rows():
        row_values = []
        for cell in row:
            row_values.append(_excel_value(cell.value))
            # Keep presentation data for non-empty/styled cells only so the
            # snapshot stays reasonably small even for large sheets.
            if cell.value is not None or cell.has_style:
                styles[cell.coordinate] = _style_snapshot(cell)
        cells.append(row_values)

    merges = [str(rng) for rng in ws.merged_cells.ranges]

    widths = {}
    hidden_columns = []
    for key, dim in ws.column_dimensions.items():
        if dim.width is not None:
            widths[key] = dim.width
        if dim.hidden:
            hidden_columns.append(key)

    heights = {}
    hidden_rows = []
    for key, dim in ws.row_dimensions.items():
        if dim.height is not None:
            heights[str(key)] = dim.height
        if dim.hidden:
            hidden_rows.append(str(key))

    freeze = ws.freeze_panes
    freeze = str(freeze) if freeze else None

    sheet_view = ws.sheet_view
    tab_color = None
    if ws.sheet_properties.tabColor:
        tab_color = (
            getattr(ws.sheet_properties.tabColor, "rgb", None)
            or getattr(ws.sheet_properties.tabColor, "indexed", None)
            or getattr(ws.sheet_properties.tabColor, "theme", None)
        )

    wb.close()

    return {
        "sheet": ws.title,
        "max_row": max_row,
        "max_column": max_col,
        "values": cells,
        "styles": styles,
        "merged_cells": merges,
        "column_widths": widths,
        "row_heights": heights,
        "hidden_columns": hidden_columns,
        "hidden_rows": hidden_rows,
        "freeze_panes": freeze,
        "show_grid_lines": bool(sheet_view.showGridLines),
        "show_row_col_headers": bool(sheet_view.showRowColHeaders),
        "tab_color": tab_color,
    }


def _ensure_snapshot_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def _source_2026_27_workbook():
    """Single source for STOCK, TRANSPORT and 2025-26 sheets."""
    return _find_existing_file("2026-27.xlsm", "2026-27.xlsx")


def _write_sheet_snapshot(snapshot_file, source, sheet_name, mode):
    if _read_snapshot(snapshot_file) is not None:
        return

    if not source:
        return

    _ensure_snapshot_dir()
    try:
        snapshot = _make_sheet_snapshot(source, sheet_name)
    except Exception:
        return

    # Never silently fall back to another sheet: the requested sheet is the
    # source of truth for this independent section.
    if _normalise_sheet_name(snapshot.get("sheet")) != _normalise_sheet_name(sheet_name):
        return

    snapshot["source"] = os.path.basename(source)
    snapshot["source_sheet"] = snapshot.get("sheet")
    snapshot["mode"] = mode
    snapshot["runtime_source"] = "static_snapshot"
    snapshot["values_are_cached_display_values"] = True

    with open(snapshot_file, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)


def create_stock_snapshot_if_needed():
    # STOCK is a sheet INSIDE 2026-27.xlsm. No STOCK.xlsx dependency.
    source = _source_2026_27_workbook()
    _write_sheet_snapshot(
        STOCK_SNAPSHOT_FILE,
        source,
        "STOCK",
        "graphical_static_stock_snapshot"
    )


def create_transport_snapshot_if_needed():
    # TRANSPORT is a separate independent sheet INSIDE 2026-27.xlsm.
    source = _source_2026_27_workbook()
    _write_sheet_snapshot(
        TRANSPORT_SNAPSHOT_FILE,
        source,
        "TRANSPORT",
        "independent_transport_snapshot"
    )


def create_2025_26_snapshot_if_needed():
    # 2025-26 is also a sheet INSIDE 2026-27.xlsm, not another workbook.
    source = _source_2026_27_workbook()
    _write_sheet_snapshot(
        OLD_2025_26_SNAPSHOT_FILE,
        source,
        "2025-26",
        "independent_2025_26_snapshot"
    )


def _snapshot_to_rows(snapshot, sheet_name):
    if not snapshot:
        return []

    values = snapshot.get("values") or []
    if not values:
        return []

    headers = []
    for value in values[0]:
        headers.append(text(value))

    rows = []
    for raw in values[1:]:
        if not any(v not in (None, "") for v in raw):
            continue
        padded = list(raw) + [None] * max(0, len(headers) - len(raw))
        data = {
            headers[i] if headers[i] else f"Column {i + 1}": padded[i]
            for i in range(len(headers))
        }
        rows.append({
            "sheet": sheet_name,
            "headers": list(data.keys()),
            "row": padded[:len(headers)],
            "data": data,
            "cancelled": False
        })

    return rows


def rows_for_financial_year(financial_year, current_rows=None):
    """Keep 2026-27 untouched; use the independent 2025-26 sheet snapshot."""
    fy = text(financial_year)

    if fy == "2025-26":
        create_2025_26_snapshot_if_needed()
        snapshot = _read_snapshot(OLD_2025_26_SNAPSHOT_FILE)
        if snapshot:
            return _snapshot_to_rows(snapshot, "2025-26")

    if current_rows is not None:
        return current_rows

    return get_all_rows()


@app.get("/stock/snapshot")
def stock_snapshot(
    authenticated: bool = Depends(require_auth)
):
    create_stock_snapshot_if_needed()
    snapshot = _read_snapshot(STOCK_SNAPSHOT_FILE)

    if snapshot is None:
        raise HTTPException(
            status_code=404,
            detail="STOCK sheet snapshot not available. Put 2026-27.xlsm beside server.py so the STOCK sheet can be captured."
        )

    return snapshot


@app.get("/transport")
def transport_data(
    limit: int = 1000,
    authenticated: bool = Depends(require_auth)
):
    create_transport_snapshot_if_needed()
    snapshot = _read_snapshot(TRANSPORT_SNAPSHOT_FILE)

    if snapshot is None:
        raise HTTPException(
            status_code=404,
            detail="TRANSPORT sheet snapshot not available in 2026-27.xlsm."
        )

    rows = _snapshot_to_rows(snapshot, "TRANSPORT")
    limit = max(1, min(int(limit), 5000))

    return {
        "source": "2026-27.xlsm",
        "independent": True,
        "sheet": snapshot.get("sheet"),
        "headers": snapshot.get("values", [[]])[0] if snapshot.get("values") else [],
        "snapshot": snapshot,
        "count": len(rows[:limit]),
        "data": [item["data"] for item in rows[:limit]]
    }


@app.get("/old-data/2025-26")
def old_data_2025_26(
    limit: int = 1000,
    authenticated: bool = Depends(require_auth)
):
    create_2025_26_snapshot_if_needed()
    snapshot = _read_snapshot(OLD_2025_26_SNAPSHOT_FILE)

    if snapshot is None:
        raise HTTPException(
            status_code=404,
            detail="2025-26 sheet snapshot not available in 2026-27.xlsm."
        )

    rows = _snapshot_to_rows(snapshot, "2025-26")
    limit = max(1, min(int(limit), 5000))

    return {
        "financial_year": "2025-26",
        "independent": True,
        "source": snapshot.get("source"),
        "sheet": snapshot.get("sheet"),
        "snapshot": snapshot,
        "headers": snapshot.get("values", [[]])[0] if snapshot.get("values") else [],
        "count": len(rows[:limit]),
        "data": [item["data"] for item in rows[:limit]]
    }


# =========================================================
# MONTHLY MATERIALS
# =========================================================

MONTHLY_MATERIALS = [

    "SLIPSHEET",

    "FRESCO PAD",

    "M FOLD",

    "CORE PIPE SCRAP",

    "PAPER SCRAP",

    "TOILET ROLL",

    "KRAFT PAPER",

    "PET GRIPSHEET",

    "TISSUE PAPER/   NAPKIN",

    "KITCHEN ROLL",

    "PLASTIC SHEET",

    "JRT",

    "Z FOLD",

]


MATERIAL_MAPPING = {

    "SLIPSHEET":
        "SLIPSHEET",

    "SLIP SHEET":
        "SLIPSHEET",

    "FRESCOPAD":
        "FRESCO PAD",

    "FRESCO PAD":
        "FRESCO PAD",

    "MFOLD":
        "M FOLD",

    "M FOLD":
        "M FOLD",

    "COREPIPESCRAP":
        "CORE PIPE SCRAP",

    "CORE PIPE SCRAP":
        "CORE PIPE SCRAP",

    "PAPERSCRAP":
        "PAPER SCRAP",

    "PAPER SCRAP":
        "PAPER SCRAP",

    "TOILETROLL":
        "TOILET ROLL",

    "TOILET ROLL":
        "TOILET ROLL",

    "KRAFTPAPER":
        "KRAFT PAPER",

    "KRAFT PAPER":
        "KRAFT PAPER",

    "PETGRIPSHEET":
        "PET GRIPSHEET",

    "PET GRIP SHEET":
        "PET GRIPSHEET",

    "PETGRIP SHEET":
        "PET GRIPSHEET",

    "NAPKIN":
        "TISSUE PAPER/   NAPKIN",

    "TISSUEPAPER":
        "TISSUE PAPER/   NAPKIN",

    "TISSUE PAPER":
        "TISSUE PAPER/   NAPKIN",

    "TISSUEPAPERNAPKIN":
        "TISSUE PAPER/   NAPKIN",

    "TISSUE PAPER NAPKIN":
        "TISSUE PAPER/   NAPKIN",

    "TISSUE PAPER/NAPKIN":
        "TISSUE PAPER/   NAPKIN",

    "KITCHENROLL":
        "KITCHEN ROLL",

    "KITCHEN ROLL":
        "KITCHEN ROLL",

    "PLASTICSHEET":
        "PLASTIC SHEET",

    "PLASTIC SHEET":
        "PLASTIC SHEET",

    "JRT":
        "JRT",

    "ZFOLD":
        "Z FOLD",

    "Z FOLD":
        "Z FOLD",

}


# =========================================================
# BASIC HELPERS
# =========================================================

def text(value):

    if value is None:
        return ""

    return str(value).strip()


def clean_value(value):

    if value is None:
        return None

    if isinstance(
        value,
        (datetime, date)
    ):
        return value.isoformat()

    return value


def number(value):

    try:

        if (
            value is None
            or value == ""
        ):
            return 0

        return float(
            str(value)
            .replace(",", "")
            .strip()
        )

    except Exception:

        return 0


def clean_number(value):

    value = float(value)

    if value.is_integer():
        return int(value)

    return round(
        value,
        2
    )


def date_only(value):

    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    value = str(value).strip()

    if not value:
        return None

    # Extra spaces remove
    value = " ".join(value.split())

    # ISO datetime:
    # 2026-09-28T10:30:00
    if "T" in value:
        value = value.split("T", 1)[0].strip()

    # Normal datetime:
    # 2026-09-28 10:30:00
    if " " in value:
        value = value.split(" ", 1)[0].strip()

    # Remove timezone if present
    if "+" in value:
        value = value.split("+", 1)[0].strip()

    # Remove trailing Z
    if value.endswith("Z"):
        value = value[:-1].strip()

    formats = (
        "%Y-%m-%d",   # 2026-09-28
        "%Y/%m/%d",   # 2026/09/28
        "%d-%m-%Y",   # 28-09-2026
        "%d/%m/%Y",   # 28/09/2026
        "%d.%m.%Y",   # 28.09.2026
        "%Y.%m.%d",   # 2026.09.28
    )

    for fmt in formats:
        try:
            return datetime.strptime(
                value,
                fmt
            ).date()

        except ValueError:
            pass

    return None


def first_existing(
    record,
    names
):

    for name in names:

        if name in record:
            return record[name]

    return None


def normalize_material(
    value
):

    value = text(
        value
    ).upper()


    value = (
        value
        .replace("\u00a0", " ")
        .replace("\t", " ")
        .replace("-", " ")
        .replace("_", " ")
        .replace("/", " ")
        .replace("\\", " ")
        .replace("(", " ")
        .replace(")", " ")
    )


    parts = value.split()


    parts = [
        part
        for part in parts
        if part not in {

            "PCS",
            "PC",
            "KG",
            "KGS",
            "NOS",
            "NO",
            "KILOGRAM",
            "KILOGRAMS",

        }
    ]


    return " ".join(
        parts
    )


def normalize_size(
    value
):

    return text(
        value
    )


def get_row_position(
    row,
    position
):

    index = position - 1


    if index < 0:
        return None


    if index >= len(row):
        return None


    return row[index]


# =========================================================
# TOKEN
# =========================================================

def make_token(
    username
):

    payload = {

        "username":
            username,

        "exp":
            int(
                datetime.now().timestamp()
            ) + 86400

    }


    raw = json.dumps(
            payload,
            separators=(
                ",",
                ":"
            )
        ).encode()


    encoded = base64.urlsafe_b64encode(
            raw
        ).decode().rstrip("=")


    signature = hmac.new(
            SECRET.encode(),
            encoded.encode(),
            hashlib.sha256
        ).hexdigest()


    return (
        encoded + "." +
        signature
    )


def verify_token(
    token
):

    try:

        encoded, signature = token.split(
                ".",
                1
            )


        expected = hmac.new(
            SECRET.encode(),
            encoded.encode(),
            hashlib.sha256
        ).hexdigest()


        if not hmac.compare_digest(
            signature,
            expected
        ):

            return False


        padding = "=" * (
                -len(encoded) % 4
            )


        raw = base64.urlsafe_b64decode(
                encoded + padding
            )


        payload = json.loads(
                raw.decode()
            )


        if (
            payload["exp"] <
            int(
                datetime.now().timestamp()
            )
        ):

            return False


        return True


    except Exception:

        return False


def require_auth(
    credentials:
        HTTPAuthorizationCredentials = Depends(security)
):

    if credentials is None:

        raise HTTPException(
            status_code=401,
            detail="Authentication required"
        )


    if (
        credentials.scheme.lower()
        != "bearer"
    ):

        raise HTTPException(
            status_code=401,
            detail="Invalid authentication"
        )


    if not verify_token(
        credentials.credentials
    ):

        raise HTTPException(
            status_code=401,
            detail="Invalid or expired token"
        )


    return True


# =========================================================
# SUPABASE
# =========================================================

def supabase_headers():

    if (
        not SUPABASE_URL
        or
        not SUPABASE_KEY
    ):

        raise HTTPException(
            status_code=500,
            detail="Supabase environment variables missing"
        )


    return {

        "apikey":
            SUPABASE_KEY,

        "Authorization":
            "Bearer " + SUPABASE_KEY,

        "Content-Type":
            "application/json",

        "Prefer":
            "return=representation"

    }


def get_all_rows():

    url = f"{SUPABASE_URL}/rest/v1/{TABLE}"


    headers = supabase_headers()


    all_rows = []


    offset = 0

    page_size = 1000


    while True:

        response = requests.get(

                url,

                headers=headers,

                params={

                    "select":
                        "*",

                    "offset":
                        offset,

                    "limit":
                        page_size

                },

                timeout=60

            )


        if not response.ok:

            raise HTTPException(

                status_code=500,

                detail="Supabase read error: " + response.text

            )


        rows = response.json()


        if not rows:
            break


        all_rows.extend(
            rows
        )


        if len(rows) < page_size:
            break


        offset += page_size


    return all_rows


# =========================================================
# DATABASE ROW
# =========================================================

def convert_database_row(
    item
):

    sheet_name = item.get(
            "sheet",
            ""
        )


    headers = item.get(
            "headers",
            []
        )


    row = item.get(
            "row",
            []
        )


    if not isinstance(
        headers,
        list
    ):

        headers = []


    if not isinstance(
        row,
        list
    ):

        row = []


    data = {}


    for i in range(
        len(headers)
    ):

        data[
            str(
                headers[i]
            )
        ] = clean_value(
                row[i]
                if i < len(row)
                else None
            )


    return {

        "sheet":
            sheet_name,

        "headers":
            headers,

        "row":
            row,

        "data":
            data

    }


# =========================================================
# HOME
# =========================================================

@app.get("/")
def home():

    return {

        "status":
            "online",

        "message":
            "Rahul Software API is running",

        "database":
            "Supabase",

        "ai":
            "Ollama / " + OLLAMA_MODEL

    }


# =========================================================
# LOGIN
# =========================================================

@app.post("/login")
def login(
    data: dict
):

    username = text(
            data.get(
                "username"
            )
        )


    password = text(
            data.get(
                "password"
            )
        )


    if (

        hmac.compare_digest(
            username,
            USERNAME
        )

        and

        hmac.compare_digest(
            password,
            PASSWORD
        )

    ):

        token = make_token(
                username
            )


        return {

            "status":
                "success",

            "access_token":
                token,

            "token":
                token,

            "token_type":
                "bearer"

        }


    raise HTTPException(

        status_code=401,

        detail="Invalid username or password"

    )


# =========================================================
# SHEETS
# =========================================================

@app.get("/sheets")
def sheets(
    authenticated:
        bool = Depends(require_auth)
):

    try:

        rows = get_all_rows()


        names = set()


        for item in rows:

            name = text(
                    item.get(
                        "sheet"
                    )
                )


            if name:
                names.add(
                    name
                )


        names = sorted(
                names
            )


        return {

            "count":
                len(names),

            "sheets":
                names

        }


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# SINGLE SHEET
# =========================================================

@app.get(
    "/sheet/{sheet_name}"
)
def get_sheet(

    sheet_name: str,

    limit: int = 100,

    authenticated:
        bool = Depends(require_auth)

):

    try:

        rows = get_all_rows()


        results = []

        headers = []


        for item in rows:

            if (
                text(
                    item.get(
                        "sheet"
                    )
                )
                != sheet_name
            ):
                continue


            converted = convert_database_row(
                    item
                )


            if not headers:

                headers = converted[
                        "headers"
                    ]


            results.append(
                converted[
                    "data"
                ]
            )


        if (
            not results
            and
            not headers
        ):

            raise HTTPException(

                status_code=404,

                detail=f"Sheet '{sheet_name}' not found"

            )


        limit = max(
                1,
                min(
                    int(limit),
                    1000
                )
            )


        results = results[
                :limit
            ]


        return {

            "sheet":
                sheet_name,

            "headers":
                headers,

            "count":
                len(results),

            "data":
                results

        }


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# FINANCIAL YEAR HELPERS (SEARCH + MONTHLY REPORT ONLY)
# =========================================================

FINANCIAL_YEARS = [
    "2025-26",
    "2026-27",
    "2027-28",
    "2028-29"
]


def financial_year_sheets(financial_year, available_sheets=None):

    fy = text(financial_year)

    if fy == "2026-27":
        return [
            "ORDERS",
            "Dispatched Orders",
            "Holded Orders"
        ]

    if available_sheets is not None and fy in available_sheets:
        return [fy]

    return [fy]


@app.get("/financial-years")
def financial_years(
    authenticated:
        bool = Depends(require_auth)
):

    return {
        "financial_years":
            FINANCIAL_YEARS
    }


# =========================================================
# SEARCH
# =========================================================

@app.get("/search")
def search(

    q: str,

    sheet:
        str = "ALL SHEETS",

    financial_year:
        str = "2026-27",

    authenticated:
        bool = Depends(require_auth)

):

    try:

        search_text = text(
                q
            ).lower()


        if not search_text:

            return {

                "query":
                    q,

                "sheet":
                    sheet,

                "financial_year":
                    financial_year,

                "count":
                    0,

                "results":
                    []

            }


        rows = rows_for_financial_year(
            financial_year,
            get_all_rows()
        )

        available_sheets = {
            text(item.get("sheet"))
            for item in rows
            if text(item.get("sheet"))
        }

        target_sheets = financial_year_sheets(
            financial_year,
            available_sheets
        )

        results = []


        for item in rows:

            sheet_name = text(
                    item.get(
                        "sheet"
                    )
                )


            if sheet_name not in target_sheets:
                continue


            if (
                sheet != "ALL SHEETS"
                and
                sheet_name != sheet
            ):

                continue


            converted = convert_database_row(
                    item
                )


            record = converted[
                    "data"
                ]


            found = False


            for value in record.values():

                if (
                    search_text
                    in
                    text(
                        value
                    ).lower()
                ):

                    found = True

                    break


            if found:

                results.append({

                    "sheet":
                        sheet_name,

                    "headers":
                        converted[
                            "headers"
                        ],

                    "row":
                        converted[
                            "row"
                        ],

                    "data":
                        record

                })


        return {

            "query":
                q,

            "sheet":
                sheet,

            "financial_year":
                financial_year,

            "count":
                len(results),

            "results":
                results

        }


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# DASHBOARD
# =========================================================

@app.get("/dashboard")
def dashboard(
    authenticated:
        bool = Depends(require_auth)
):

    try:

        rows = get_all_rows()


        today = datetime.now().date()


        orders = []

        dispatch = []

        hold = []


        for item in rows:

            sheet = text(
                    item.get(
                        "sheet"
                    )
                )


            converted = convert_database_row(
                    item
                )


            if sheet == "ORDERS":

                orders.append(
                    converted
                )


            elif sheet == "Dispatched Orders":

                dispatch.append({

                    **converted,

                    "cancelled":
                        item.get(
                            "cancelled",
                            False
                        )
                        is True

                })


            elif sheet == "Holded Orders":

                hold.append(
                    converted
                )


        order_quantity = 0

        order_numbers = set()

        today_orders = 0

        today_order_quantity = 0

        size_data = {}

        party_data = {}


        for item in orders:

            row = item["row"]

            record = item["data"]


            order_no = first_existing(
                    record,
                    [
                        "Order No.",
                        "Order No"
                    ]
                )


            party = text(
                    first_existing(
                        record,
                        [
                            "PARTY NAME",
                            "Party Name",
                            "PARTY"
                        ]
                    )
                )


            size = normalize_size(
                    first_existing(
                        record,
                        [
                            "Size",
                            "SIZE"
                        ]
                    )
                )


            quantity = number(
                    first_existing(
                        record,
                        [
                            "Quantity Pcs.",
                            "Quantity",
                            "QUANTITY"
                        ]
                    )
                )


            row_date = date_only(
                    first_existing(
                        record,
                        [
                            "Date",
                            "DATE"
                        ]
                    )
                )


            destination = text(
                    get_row_position(
                        row,
                        9
                    )
                )


            if order_no not in (
                None,
                ""
            ):

                order_numbers.add(
                    text(
                        order_no
                    )
                )


            order_quantity += quantity


            if row_date == today:

                today_orders += 1

                today_order_quantity += quantity


            if size:

                if size not in size_data:

                    size_data[size] = {

                        "size":
                            size,

                        "orders":
                            0,

                        "quantity":
                            0

                    }


                size_data[
                    size
                ]["orders"] += 1


                size_data[
                    size
                ]["quantity"] += quantity


            if party:

                if party not in party_data:

                    party_data[
                        party
                    ] = {

                        "party":
                            party,

                        "orders":
                            0,

                        "quantity":
                            0,

                        "destinations":
                            {}

                    }


                party_data[
                    party
                ]["orders"] += 1


                party_data[
                    party
                ]["quantity"] += quantity


                destination_key = destination or "-"


                if (
                    destination_key
                    not in
                    party_data[
                        party
                    ]["destinations"]
                ):

                    party_data[
                        party
                    ]["destinations"][
                        destination_key
                    ] = {

                        "destination":
                            destination_key,

                        "orders":
                            0,

                        "quantity":
                            0,

                        "sizes":
                            {}

                    }


                destination_item = party_data[
                        party
                    ]["destinations"][
                        destination_key
                    ]


                destination_item[
                    "orders"
                ] += 1


                destination_item[
                    "quantity"
                ] += quantity


                size_key = size or "-"


                if (
                    size_key
                    not in
                    destination_item[
                        "sizes"
                    ]
                ):

                    destination_item[
                        "sizes"
                    ][size_key] = {

                        "size":
                            size_key,

                        "orders":
                            0,

                        "quantity":
                            0

                    }


                destination_item[
                    "sizes"
                ][
                    size_key
                ]["orders"] += 1


                destination_item[
                    "sizes"
                ][
                    size_key
                ]["quantity"] += quantity


        dispatch_quantity = 0

        today_dispatch = 0

        today_dispatch_quantity = 0

        cancelled_dispatch_rows = 0


        for item in dispatch:

            if item.get(
                "cancelled",
                False
            ):

                cancelled_dispatch_rows += 1

                continue


            row = item["row"]

            record = item["data"]


            quantity = number(
                    get_row_position(
                        row,
                        6
                    )
                )


            row_date = date_only(
                    get_row_position(
                        row,
                        11
                    )
                )


            dispatch_quantity += quantity


            if row_date == today:

                today_dispatch += 1

                today_dispatch_quantity += quantity


        hold_quantity = 0


        for item in hold:

            record = item["data"]


            hold_quantity += \
                number(
                    first_existing(
                        record,
                        [
                            "Quantity Pcs.",
                            "Quantity",
                            "QUANTITY"
                        ]
                    )
                )


        size_wise = list(
            size_data.values()
        )


        for item in size_wise:

            item["quantity"] = clean_number(
                    item["quantity"]
                )


        size_wise.sort(
            key=lambda x:
                x["quantity"],
            reverse=True
        )


        party_wise = []


        for item in party_data.values():

            destinations = []


            for destination_item in \
                item[
                    "destinations"
                ].values():

                sizes = []


                for size_item in \
                    destination_item[
                        "sizes"
                    ].values():

                    sizes.append({

                        "size":
                            size_item[
                                "size"
                            ],

                        "orders":
                            size_item[
                                "orders"
                            ],

                        "quantity":
                            clean_number(
                                size_item[
                                    "quantity"
                                ]
                            )

                    })


                sizes.sort(
                    key=lambda x:
                        x["quantity"],
                    reverse=True
                )


                destinations.append({

                    "destination":
                        destination_item[
                            "destination"
                        ],

                    "orders":
                        destination_item[
                            "orders"
                        ],

                    "quantity":
                        clean_number(
                            destination_item[
                                "quantity"
                            ]
                        ),

                    "sizes":
                        sizes

                })


            destinations.sort(
                key=lambda x:
                    x["quantity"],
                reverse=True
            )


            party_wise.append({

                "party":
                    item[
                        "party"
                    ],

                "orders":
                    item[
                        "orders"
                    ],

                "quantity":
                    clean_number(
                        item[
                            "quantity"
                        ]
                    ),

                "destinations":
                    destinations

            })


        party_wise.sort(
            key=lambda x:
                x["quantity"],
            reverse=True
        )


        return {

            "date":
                today.isoformat(),


            "today": {

                "orders":
                    today_orders,

                "order_quantity":
                    clean_number(
                        today_order_quantity
                    ),

                "dispatch":
                    today_dispatch,

                "dispatch_quantity":
                    clean_number(
                        today_dispatch_quantity
                    )

            },


            "orders": {

                "total_rows":
                    len(orders),

                "unique_orders":
                    len(
                        order_numbers
                    ),

                "total_quantity":
                    clean_number(
                        order_quantity
                    )

            },


            "dispatch": {

                "total_rows":
                    len(dispatch)
                    - cancelled_dispatch_rows,

                "total_quantity":
                    clean_number(
                        dispatch_quantity
                    ),

                "today_rows":
                    today_dispatch,

                "today_quantity":
                    clean_number(
                        today_dispatch_quantity
                    ),

                "cancelled_rows":
                    cancelled_dispatch_rows

            },


            "hold": {

                "orders":
                    len(hold),

                "quantity":
                    clean_number(
                        hold_quantity
                    )

            },


            "size_wise":
                size_wise,


            "party_wise":
                party_wise,


            "last_updated":
                datetime.now()
                .astimezone()
                .isoformat()

        }


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# DASHBOARD - DISPATCH QUANTITY BY FINANCIAL YEAR
# =========================================================

@app.get("/dashboard/dispatch-quantity")
def dashboard_dispatch_quantity(

    financial_year:
        str = "2026-27",

    authenticated:
        bool = Depends(require_auth)

):

    try:

        rows = rows_for_financial_year(
            financial_year,
            get_all_rows()
        )

        fy = text(financial_year)

        if fy == "2026-27":
            target_sheets = [
                "Dispatched Orders"
            ]
        else:
            target_sheets = [
                fy
            ]

        quantity_total = 0

        for item in rows:

            if text(item.get("sheet")) not in target_sheets:
                continue

            if item.get("cancelled", False) is True:
                continue

            converted = convert_database_row(item)

            quantity_total += number(
                get_row_position(
                    converted["row"],
                    6
                )
            )

        return {

            "financial_year":
                fy,

            "dispatch_quantity":
                clean_number(quantity_total)

        }

    except HTTPException:
        raise

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# MONTHS
# =========================================================

@app.get(
    "/monthly-report/months"
)
def monthly_report_months(

    financial_year:
        str = "2026-27",

    authenticated:
        bool = Depends(require_auth)
):

    try:

        rows = rows_for_financial_year(
            financial_year,
            get_all_rows()
        )

        available_sheets = {
            text(item.get("sheet"))
            for item in rows
            if text(item.get("sheet"))
        }

        target_sheets = financial_year_sheets(
            financial_year,
            available_sheets
        )

        # Monthly dispatch data comes from the same positional
        # columns in the old 2025-26 sheet and current dispatch sheet.
        months = set()


        for item in rows:

            if text(item.get("sheet")) not in target_sheets:
                continue


            if item.get(
                "cancelled",
                False
            ) is True:

                continue


            converted = convert_database_row(
                    item
                )


            row_date = date_only(
                    get_row_position(
                        converted[
                            "row"
                        ],
                        11
                    )
                )


            if row_date:

                months.add(
                    (
                        row_date.year,
                        row_date.month
                    )
                )


        month_list = []


        for year, month in sorted(
            months,
            reverse=True
        ):

            month_list.append({

                "year":
                    year,

                "month":
                    month,

                "month_name":
                    calendar.month_name[
                        month
                    ],

                "label":
                    f"{calendar.month_name[month]} {year}"

            })


        return {

            "count":
                len(month_list),

            "months":
                month_list

        }


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# DAILY PCS DISPATCH
# =========================================================

@app.get(
    "/monthly-report/daily-pcs-dispatch"
)
def daily_pcs_dispatch(

    year: int,

    month: int,

    financial_year:
        str = "2026-27",

    authenticated:
        bool = Depends(require_auth)

):

    try:

        if month < 1 or month > 12:

            raise HTTPException(
                status_code=400,
                detail="Invalid month"
            )


        if year < 2000 or year > 2100:

            raise HTTPException(
                status_code=400,
                detail="Invalid year"
            )


        rows = rows_for_financial_year(
            financial_year,
            get_all_rows()
        )

        available_sheets = {
            text(item.get("sheet"))
            for item in rows
            if text(item.get("sheet"))
        }

        target_sheets = financial_year_sheets(
            financial_year,
            available_sheets
        )


        days_in_month = calendar.monthrange(
                year,
                month
            )[1]


        daily_data = {}


        for day_number in range(
            1,
            days_in_month + 1
        ):

            daily_data[
                day_number
            ] = {

                "date":
                    day_number

            }


            for material in \
                MONTHLY_MATERIALS:

                daily_data[
                    day_number
                ][material] = 0


        matched_rows = 0

        skipped_cancelled_rows = 0


        for item in rows:

            if text(item.get("sheet")) not in target_sheets:

                continue

                continue


            if item.get(
                "cancelled",
                False
            ) is True:

                skipped_cancelled_rows += 1

                continue


            converted = convert_database_row(
                    item
                )


            row = converted[
                    "row"
                ]


            row_date = date_only(
                    get_row_position(
                        row,
                        11
                    )
                )


            if row_date is None:
                continue


            if (
                row_date.year != year
                or
                row_date.month != month
            ):

                continue


            material_type = text(
                    get_row_position(
                        row,
                        13
                    )
                )


            normalized = normalize_material(
                    material_type
                )


            report_column = MATERIAL_MAPPING.get(
                    normalized
                )


            if not report_column:
                continue


            quantity = number(
                    get_row_position(
                        row,
                        6
                    )
                )


            daily_data[
                row_date.day
            ][
                report_column
            ] += quantity


            matched_rows += 1


        total = {

            "date":
                "TOTAL"

        }


        for material in \
            MONTHLY_MATERIALS:

            value = 0


            for day_number in \
                daily_data:

                value += number(
                    daily_data[
                        day_number
                    ][material]
                )


            total[
                material
            ] = clean_number(
                    value
                )


        days = []


        for day_number in range(
            1,
            days_in_month + 1
        ):

            item = {

                "date":
                    day_number

            }


            for material in \
                MONTHLY_MATERIALS:

                item[
                    material
                ] = clean_number(
                        daily_data[
                            day_number
                        ][material]
                    )


            days.append(
                item
            )


        month_title = (
            f"{calendar.month_name[month].upper()}- "
            f"{year} MONTHLY PCS DISPATCH QUANTITY"
        )


        return {

            "year":
                year,

            "month":
                month,

            "month_name":
                calendar.month_name[
                    month
                ],

            "title":
                month_title,

            "columns":
                MONTHLY_MATERIALS,

            "days":
                days,

            "total":
                total,

            "matched_dispatch_rows":
                matched_rows,

            "skipped_cancelled_rows":
                skipped_cancelled_rows,

            "last_updated":
                datetime.now()
                .astimezone()
                .isoformat()

        }


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# CURRENT MONTH
# =========================================================

@app.get(
    "/monthly-report/daily-pcs-dispatch/current"
)
def current_month_daily_pcs_dispatch(
    authenticated:
        bool = Depends(require_auth)
):

    today = datetime.now().date()


    return daily_pcs_dispatch(

        year=today.year,

        month=today.month,

        authenticated=True

    )


# =========================================================
# MATERIAL MAPPING
# =========================================================

@app.get(
    "/monthly-report/material-mapping"
)
def material_mapping(
    authenticated:
        bool = Depends(require_auth)
):

    try:

        rows = get_all_rows()


        actual_types = {}


        for item in rows:

            if (
                text(
                    item.get(
                        "sheet"
                    )
                )
                != "Dispatched Orders"
            ):

                continue


            converted = convert_database_row(
                    item
                )


            record = converted[
                    "data"
                ]


            material_type = text(
                    first_existing(
                        record,
                        [
                            "TYPE OF MATERIAL",
                            "Type Of Material",
                            "TYPE OF MATERIAL "
                        ]
                    )
                )


            if not material_type:
                continue


            normalized = normalize_material(
                    material_type
                )


            mapped_to = MATERIAL_MAPPING.get(
                    normalized
                )


            if material_type not in actual_types:

                actual_types[
                    material_type
                ] = {

                    "excel_value":
                        material_type,

                    "report_column":
                        mapped_to,

                    "mapped":
                        bool(
                            mapped_to
                        )

                }


        return {

            "count":
                len(actual_types),

            "materials":
                list(
                    actual_types.values()
                )

        }


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# TEST REPORT - PARTY MASTER
# =========================================================

def _test_report_master():

    rows = get_all_rows()


    parties = {}


    for item in rows:

        if (
            text(
                item.get(
                    "sheet"
                )
            )
            != "PARTIES"
        ):

            continue


        row = item.get(
                "row",
                []
            )


        if not isinstance(
            row,
            list
        ):

            continue


        # PARTIES SHEET:
        # A = PARTY NAME
        # B = ADDRESS


        party = text(
                get_row_position(
                    row,
                    1
                )
            )


        address = text(
                get_row_position(
                    row,
                    2
                )
            )


        if not party:
            continue


        header = party.upper()


        if header in {

            "PARTY",
            "PARTY NAME",
            "PARTIES",
            "PARTIES PLANTS",
            "PARTIES PLANTS NAME",
            "CUSTOMER NAME"

        }:

            continue


        parties[
            party
        ] = address


    result = [

        {

            "party":
                party,

            "address":
                address

        }

        for party, address
        in sorted(
            parties.items(),
            key=lambda x:
                x[0].upper()
        )

    ]


    return {

        "sheet":
            "PARTIES",

        "count":
            len(result),

        "parties":
            result

    }


@app.get(
    "/test-report/master-data"
)
def test_report_master_data(
    authenticated:
        bool = Depends(require_auth)
):

    try:

        return _test_report_master()


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# TEST REPORT - SIZE PARSER
# =========================================================

def parse_size_pattern(
    value
):

    s = text(
            value
        ).upper()


    s = s.replace(
            " ",
            ""
        )


    s = s.replace(
            "×",
            "X"
        )


    if s.count("X") != 1:

        raise ValueError(
            "Size format: 1200+75X1000+75"
        )


    left, right = s.split(
            "X",
            1
        )


    def nums(part):

        values = part.split(
                "+"
            )


        if any(
            v == ""
            for v in values
        ):

            raise ValueError(
                "Invalid size format"
            )


        result = [

            float(v)

            for v in values

        ]


        if any(
            v <= 0
            for v in result
        ):

            raise ValueError(
                "Size values must be greater than zero"
            )


        return result


    a = nums(
            left
        )


    b = nums(
            right
        )


    if (
        len(a) == 1
        and
        len(b) == 1
    ):

        return {

            "type":
                "none",

            "length":
                a[0],

            "width":
                b[0]

        }


    if (
        len(a) == 2
        and
        len(b) == 1
    ):

        return {

            "type":
                "one",

            "length":
                a[0],

            "tab_length":
                a[1],

            "width":
                b[0]

        }


    if (
        len(a) == 3
        and
        len(b) == 1
    ):

        return {

            "type":
                "two_length",

            "tab_length_1":
                a[0],

            "length":
                a[1],

            "tab_length_2":
                a[2],

            "width":
                b[0]

        }


    if (
        len(a) == 1
        and
        len(b) == 3
    ):

        return {

            "type":
                "two_width",

            "length":
                a[0],

            "tab_width_1":
                b[0],

            "width":
                b[1],

            "tab_width_2":
                b[2]

        }


    if (
        len(a) == 2
        and
        len(b) == 2
    ):

        return {

            "type":
                "two_mixed",

            "length":
                a[0],

            "tab_length":
                a[1],

            "width":
                b[0],

            "tab_width":
                b[1]

        }


    if (
        len(a) == 3
        and
        len(b) == 3
    ):

        return {

            "type":
                "four",

            "tab_length_1":
                a[0],

            "length":
                a[1],

            "tab_length_2":
                a[2],

            "tab_width_1":
                b[0],

            "width":
                b[1],

            "tab_width_2":
                b[2]

        }


    raise ValueError(
        "Unsupported size format"
    )


@app.get(
    "/test-report/parse-size"
)
def test_report_parse_size(

    size: str,

    authenticated:
        bool = Depends(require_auth)

):

    try:

        return parse_size_pattern(
            size
        )


    except ValueError as e:

        raise HTTPException(
            status_code=400,
            detail=str(e)
        )


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# TEST REPORT - BILL NUMBER
# =========================================================

def next_bill_number(
    prefix="MP/26-27/"
):

    rows = get_all_rows()


    highest = 408


    for item in rows:

        if (
            text(
                item.get(
                    "sheet"
                )
            )
            != "Dispatched Orders"
        ):

            continue


        converted = convert_database_row(
                item
            )


        record = converted[
                "data"
            ]


        invoice = text(
                first_existing(
                    record,
                    [
                        "Inovoice",
                        "Invoice",
                        "INVOICE"
                    ]
                )
            )


        if not invoice:
            continue


        value = invoice.strip()


        if not value.upper().startswith(
            prefix.upper()
        ):

            continue


        tail = value[
                len(prefix):
            ].strip()


        if tail.isdigit():

            highest = max(
                    highest,
                    int(tail)
                )


    return (
        prefix + str(
            highest + 1
        )
    )


@app.post(
    "/test-report/reserve-bill"
)
def reserve_test_bill(

    data: dict,

    authenticated:
        bool = Depends(require_auth)

):

    try:

        mode = text(
                data.get(
                    "mode"
                )
                or
                "AUTO NEXT"
            ).upper()


        manual_bill = text(
                data.get(
                    "bill_no"
                )
            )


        if mode == "MANUAL":

            if not manual_bill:

                raise HTTPException(
                    status_code=400,
                    detail="Manual Bill No. required"
                )


            return {

                "bill_no":
                    manual_bill

            }


        if mode != "AUTO NEXT":

            raise HTTPException(
                status_code=400,
                detail="Invalid bill mode"
            )


        return {

            "bill_no":
                next_bill_number(
                    "MP/26-27/"
                )

        }


    except HTTPException:

        raise


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# TEST REPORT - GENERATION
# =========================================================

def rand_int(
    minimum,
    maximum
):

    return random.randint(
        minimum,
        maximum
    )


def fmt_mm(
    value
):

    value = float(
            value
        )


    if value.is_integer():

        return str(
            int(value)
        )


    return str(
        round(
            value,
            2
        )
    )


def actual_total(
    value
):

    return fmt_mm(
        float(value)
        + rand_int(
            -5,
            5
        )
    )


def actual_tab(
    value
):

    return fmt_mm(
        float(value)
        + rand_int(
            -3,
            3
        )
    )


def actual_thickness(
    value
):

    value = float(
            value
        )


    actual = value + random.uniform(
            -0.05,
            0.05
        )


    return fmt_mm(
        max(
            0.01,
            actual
        )
    )


def actual_moisture():

    return (
        f"{rand_int(10,12)} %"
    )


def tc_standard(
    thickness
):

    t = float(
            thickness
        )


    if abs(
        t - 0.9
    ) < 0.001:

        return 800


    if (
        abs(
            t - 1.2
        ) < 0.001
        or
        abs(
            t - 1.5
        ) < 0.001
    ):

        return 1000


    return None


def actual_tc(
    thickness
):

    standard = tc_standard(
            thickness
        )


    if standard is None:

        return "—"


    return (
        f"{standard + rand_int(-150,150)} kg"
    )


def test_report_rows(
    parsed,
    thickness
):

    names = [
        "Total Length"
    ]


    if parsed["type"] == "none":

        names += [
            "Total Width"
        ]


    elif parsed["type"] == "one":

        names += [
            "Total Width",
            "1 Tab at Length"
        ]


    elif parsed["type"] == "two_length":

        names += [
            "Total Width",
            "1 Tab at Length",
            "2 Tab at Length"
        ]


    elif parsed["type"] == "two_width":

        names += [
            "Total Width",
            "1 Tab at Width",
            "2 Tab at Width"
        ]


    elif parsed["type"] == "two_mixed":

        names += [
            "Total Width",
            "1 Tab at Length",
            "1 Tab at Width"
        ]


    elif parsed["type"] == "four":

        names += [
            "Total Width",
            "1 Tab at Length",
            "2 Tab at Length",
            "1 Tab at Width",
            "2 Tab at Width"
        ]


    names += [

        "Thickness",

        "Moisture",

        "TC"

    ]


    def standard_for(
        name
    ):

        if name == "Total Length":

            return (
                fmt_mm(
                    parsed["length"]
                ) + "mm",

                "±10mm"
            )


        if name == "Total Width":

            return (
                fmt_mm(
                    parsed["width"]
                ) + "mm",

                "±10mm"
            )


        if name == "1 Tab at Length":

            value = parsed.get(
                    "tab_length",
                    parsed.get(
                        "tab_length_1"
                    )
                )


            return (
                fmt_mm(
                    value
                ) + "mm",

                "±5mm"
            )


        if name == "2 Tab at Length":

            return (
                fmt_mm(
                    parsed[
                        "tab_length_2"
                    ]
                ) + "mm",

                "±5mm"
            )


        if name == "1 Tab at Width":

            value = parsed.get(
                    "tab_width",
                    parsed.get(
                        "tab_width_1"
                    )
                )


            return (
                fmt_mm(
                    value
                ) + "mm",

                "±5mm"
            )


        if name == "2 Tab at Width":

            return (
                fmt_mm(
                    parsed[
                        "tab_width_2"
                    ]
                ) + "mm",

                "±5mm"
            )


        if name == "Thickness":

            return (
                fmt_mm(
                    thickness
                ) + "mm",

                "±0.3mm"
            )


        if name == "Moisture":

            return (
                "10 % to 12 %",
                "±2 %"
            )


        if name == "TC":

            standard = tc_standard(
                    thickness
                )


            return (

                (
                    f"{standard} kg"
                    if standard is not None
                    else "—"
                ),

                "±150 kg"

            )


        return (
            "",
            ""
        )


    actual_map = {

        "Total Length":
            [
                actual_total(
                    parsed[
                        "length"
                    ]
                )
                for _ in range(4)
            ],

        "Total Width":
            [
                actual_total(
                    parsed[
                        "width"
                    ]
                )
                for _ in range(4)
            ],

        "Thickness":
            [
                actual_thickness(
                    thickness
                )
                for _ in range(4)
            ],

        "Moisture":
            [
                actual_moisture()
                for _ in range(4)
            ],

        "TC":
            [
                actual_tc(
                    thickness
                )
                for _ in range(4)
            ]

    }


    if parsed["type"] == "one":

        actual_map[
            "1 Tab at Length"
        ] = [

            actual_tab(
                parsed[
                    "tab_length"
                ]
            )

            for _ in range(4)

        ]


    elif parsed["type"] == "two_length":

        actual_map[
            "1 Tab at Length"
        ] = [

            actual_tab(
                parsed[
                    "tab_length_1"
                ]
            )

            for _ in range(4)

        ]


        actual_map[
            "2 Tab at Length"
        ] = [

            actual_tab(
                parsed[
                    "tab_length_2"
                ]
            )

            for _ in range(4)

        ]


    elif parsed["type"] == "two_width":

        actual_map[
            "1 Tab at Width"
        ] = [

            actual_tab(
                parsed[
                    "tab_width_1"
                ]
            )

            for _ in range(4)

        ]


        actual_map[
            "2 Tab at Width"
        ] = [

            actual_tab(
                parsed[
                    "tab_width_2"
                ]
            )

            for _ in range(4)

        ]


    elif parsed["type"] == "two_mixed":

        actual_map[
            "1 Tab at Length"
        ] = [

            actual_tab(
                parsed[
                    "tab_length"
                ]
            )

            for _ in range(4)

        ]


        actual_map[
            "1 Tab at Width"
        ] = [

            actual_tab(
                parsed[
                    "tab_width"
                ]
            )

            for _ in range(4)

        ]


    elif parsed["type"] == "four":

        actual_map[
            "1 Tab at Length"
        ] = [

            actual_tab(
                parsed[
                    "tab_length_1"
                ]
            )

            for _ in range(4)

        ]


        actual_map[
            "2 Tab at Length"
        ] = [

            actual_tab(
                parsed[
                    "tab_length_2"
                ]
            )

            for _ in range(4)

        ]


        actual_map[
            "1 Tab at Width"
        ] = [

            actual_tab(
                parsed[
                    "tab_width_1"
                ]
            )

            for _ in range(4)

        ]


        actual_map[
            "2 Tab at Width"
        ] = [

            actual_tab(
                parsed[
                    "tab_width_2"
                ]
            )

            for _ in range(4)

        ]


    result = []


    for name in names:

        standard, tolerance = standard_for(
                name
            )


        result.append({

            "name":
                name,

            "standard":
                standard,

            "tolerance":
                tolerance,

            "actual":
                actual_map.get(
                    name,
                    [
                        "",
                        "",
                        "",
                        ""
                    ]
                )

        })


    return result


@app.post(
    "/test-report/generate"
)
def generate_test_report(

    data: dict,

    authenticated:
        bool = Depends(require_auth)

):

    try:

        party = text(
                data.get(
                    "party"
                )
            )


        address = text(
                data.get(
                    "address"
                )
            )


        size = text(
                data.get(
                    "size"
                )
            )


        thickness = number(
                data.get(
                    "thickness"
                )
            )


        gsm = text(
                data.get(
                    "gsm"
                )
            )


        po_mode = text(
                data.get(
                    "po_mode"
                )
                or
                "AS PER SHEET"
            ).upper()


        po_no = text(
                data.get(
                    "po_no"
                )
            )


        po_date = text(
                data.get(
                    "po_date"
                )
            )


        bill_mode = text(
                data.get(
                    "bill_mode"
                )
                or
                "AUTO NEXT"
            ).upper()


        manual_bill = text(
                data.get(
                    "bill_no"
                )
            )


        bill_date = text(
                data.get(
                    "bill_date"
                )
            )


        drawing = text(
                data.get(
                    "drawing"
                )
            )


        specification = text(
                data.get(
                    "specification"
                )
            )


        if not party:

            raise HTTPException(
                status_code=400,
                detail="Party select karo"
            )


        if not size:

            raise HTTPException(
                status_code=400,
                detail="Size bharo"
            )


        if thickness <= 0:

            raise HTTPException(
                status_code=400,
                detail="Thickness sahi bharo"
            )


        if po_mode not in {

            "AS PER SHEET",
            "VERBAL",
            "MANUAL"

        }:

            raise HTTPException(
                status_code=400,
                detail="Invalid PO mode"
            )


        if (
            po_mode == "MANUAL"
            and
            not po_no
        ):

            raise HTTPException(
                status_code=400,
                detail="Manual PO No. bharo"
            )


        if bill_mode not in {

            "AUTO NEXT",
            "MANUAL"

        }:

            raise HTTPException(
                status_code=400,
                detail="Invalid bill mode"
            )


        parsed = parse_size_pattern(
                size
            )


        if bill_mode == "AUTO NEXT":

            bill_no = next_bill_number(
                    "MP/26-27/"
                )

        else:

            if not manual_bill:

                raise HTTPException(
                    status_code=400,
                    detail="Manual Bill No. bharo"
                )

            bill_no = manual_bill


        if po_mode == "AS PER SHEET":

            display_po = "AS PER SHEET"

        elif po_mode == "VERBAL":

            display_po = "VERBAL"

        else:

            display_po = po_no


        return {

            "party":
                party,

            "address":
                address,

            "size":
                size,

            "thickness":
                thickness,

            "gsm":
                gsm,

            "po_no":
                display_po,

            "po_date":
                po_date,

            "bill_no":
                bill_no,

            "bill_date":
                bill_date,

            "drawing":
                drawing,

            "specification":
                specification,

            "rows":
                test_report_rows(
                    parsed,
                    thickness
                )

        }


    except HTTPException:

        raise


    except ValueError as e:

        raise HTTPException(
            status_code=400,
            detail=str(e)
        )


    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# =========================================================
# AI CHAT - LOCAL OLLAMA
# =========================================================

@app.post("/ai-chat")
def ai_chat(
    data: dict,
    authenticated:
        bool = Depends(require_auth)
):

    try:

        message = text(
            data.get(
                "message"
            )
        )


        if not message:

            raise HTTPException(
                status_code=400,
                detail="Message is required"
            )


        if len(message) > 10000:

            raise HTTPException(
                status_code=400,
                detail="Message is too long"
            )


        system_prompt = """
You are Rahul AI Assistant inside Rahul Software.

Rahul Software is a business operations system.

The business works with:

- Orders
- Dispatch
- Kraft paper
- Paper products
- Slipsheets
- Gripsheets
- PET gripsheets
- Packaging
- GSM
- TC
- Thickness
- Caliper
- ECT
- Weight calculations
- Roll calculations
- Monthly reports
- Test reports
- Business operations

Answer the user's questions clearly and practically.

Use simple Indian business English/Hinglish when appropriate.

If the user asks for a calculation:

- show the formula
- show the calculation
- show the final answer

If the user asks a technical paper/packaging question:

- explain it simply
- give a practical example when useful

IMPORTANT:

Do not invent Rahul Software's business data.

You do NOT automatically have access to Supabase data
just because you are inside Rahul Software.

If the user asks about specific company data, orders,
parties, dispatch quantities, stock, invoices, or reports
and that data has not been provided to you in the request,
clearly say that the required data is not currently
available to the AI.

Do not claim that you changed, deleted, created, or updated
any business data.

Keep answers concise unless the user asks for detailed
explanation.

You are running locally through Ollama.

Do not mention Gemini, Google Gemini API, API keys,
or cloud AI unless the user specifically asks about
the AI architecture.
"""


        ollama_payload = {

            "model":
                OLLAMA_MODEL,

            "stream":
                False,

            "messages": [

                {
                    "role":
                        "system",

                    "content":
                        system_prompt

                },

                {
                    "role":
                        "user",

                    "content":
                        message

                }

            ],

            "options": {

                "temperature":
                    0.3

            }

        }


        response = requests.post(

            OLLAMA_URL + "/api/chat",

            json=ollama_payload,

            timeout=180

        )


        if not response.ok:

            raise HTTPException(

                status_code=500,

                detail=(
                    "Ollama error: "
                    + response.text
                )

            )


        result = response.json()


        answer = ""


        if isinstance(
            result,
            dict
        ):

            message_data = result.get(
                "message"
            )


            if isinstance(
                message_data,
                dict
            ):

                answer = text(
                    message_data.get(
                        "content"
                    )
                )


            if not answer:

                answer = text(
                    result.get(
                        "response"
                    )
                )


        if not answer:

            answer = (
                "AI ne koi response "
                "generate nahi kiya."
            )


        return {

            "status":
                "success",

            "answer":
                answer

        }


    except HTTPException:

        raise


    except requests.exceptions.ConnectionError:

        raise HTTPException(

            status_code=503,

            detail=(
                "Local AI server (Ollama) "
                "connect nahi ho raha. "
                "Check karo ki Ollama running hai "
                "aur OLLAMA_URL sahi hai."
            )

        )


    except requests.exceptions.Timeout:

        raise HTTPException(

            status_code=504,

            detail=(
                "Local AI response mein "
                "bahut time lag raha hai. "
                "Ollama model ya computer load check karo."
            )

        )


    except Exception as e:

        raise HTTPException(

            status_code=500,

            detail="AI error: " + str(e)

        )
