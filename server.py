from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

from datetime import datetime, date, timedelta
import os
import hmac
import hashlib
import base64
import json
import requests
import calendar


# =========================================================
# APP
# =========================================================

app = FastAPI(title="Rahul Software API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

security = HTTPBearer(auto_error=False)


# =========================================================
# ENVIRONMENT VARIABLES
# =========================================================

USERNAME = os.getenv("RAHUL_USERNAME", "rahul")
PASSWORD = os.getenv("RAHUL_PASSWORD", "rahul123")
SECRET = os.getenv("RAHUL_SECRET", "change-this-secret")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

TABLE = "excel_rows"


# =========================================================
# MONTHLY REPORT MATERIAL COLUMNS
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


# =========================================================
# MATERIAL MAPPING
# =========================================================

MATERIAL_MAPPING = {
    "SLIPSHEET": "SLIPSHEET",
    "SLIP SHEET": "SLIPSHEET",

    "FRESCOPAD": "FRESCO PAD",
    "FRESCO PAD": "FRESCO PAD",

    "MFOLD": "M FOLD",
    "M FOLD": "M FOLD",

    "COREPIPESCRAP": "CORE PIPE SCRAP",
    "CORE PIPE SCRAP": "CORE PIPE SCRAP",

    "PAPERSCRAP": "PAPER SCRAP",
    "PAPER SCRAP": "PAPER SCRAP",

    "TOILETROLL": "TOILET ROLL",
    "TOILET ROLL": "TOILET ROLL",

    "KRAFTPAPER": "KRAFT PAPER",
    "KRAFT PAPER": "KRAFT PAPER",

    "PETGRIPSHEET": "PET GRIPSHEET",
    "PET GRIP SHEET": "PET GRIPSHEET",
    "PETGRIP SHEET": "PET GRIPSHEET",

    "NAPKIN": "TISSUE PAPER/   NAPKIN",
    "TISSUEPAPER": "TISSUE PAPER/   NAPKIN",
    "TISSUE PAPER": "TISSUE PAPER/   NAPKIN",
    "TISSUEPAPERNAPKIN": "TISSUE PAPER/   NAPKIN",
    "TISSUE PAPER NAPKIN": "TISSUE PAPER/   NAPKIN",

    "KITCHENROLL": "KITCHEN ROLL",
    "KITCHEN ROLL": "KITCHEN ROLL",

    "PLASTICSHEET": "PLASTIC SHEET",
    "PLASTIC SHEET": "PLASTIC SHEET",

    "JRT": "JRT",

    "ZFOLD": "Z FOLD",
    "Z FOLD": "Z FOLD",
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

    if isinstance(value, (datetime, date)):
        return value.isoformat()

    return value


def number(value):
    try:
        if value is None or value == "":
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

    return round(value, 2)


def date_only(value):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    value = str(value).strip()

    # ISO datetime:
    # 2026-08-02T00:00:00
    # 2026-08-02T00:00:00+00:00
    try:
        return datetime.fromisoformat(
            value.replace("Z", "+00:00")
        ).date()
    except Exception:
        pass

    formats = (
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d.%m.%Y",
    )

    for fmt in formats:
        try:
            return datetime.strptime(
                value,
                fmt
            ).date()
        except Exception:
            pass

    return None


def first_existing(record, names):
    for name in names:
        if name in record:
            return record[name]

    return None


# =========================================================
# MATERIAL NORMALIZATION
# =========================================================

def normalize_material(value):

    value = text(value).upper()

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

    units_to_remove = {
        "PCS",
        "PC",
        "KG",
        "KGS",
        "NOS",
        "NO",
        "KILOGRAM",
        "KILOGRAMS",
    }

    parts = [
        part
        for part in parts
        if part not in units_to_remove
    ]

    return " ".join(parts)


def normalize_size(value):

    value = text(value)

    value = " ".join(
        value.split()
    )

    return value


def get_row_position(row, position):

    index = position - 1

    if index < 0:
        return None

    if index >= len(row):
        return None

    return row[index]


# =========================================================
# TOKEN AUTHENTICATION
# =========================================================

def make_token(username):

    payload = {
        "username": username,
        "exp": int(
            datetime.now().timestamp()
        ) + 86400
    }

    raw = json.dumps(
        payload,
        separators=(",", ":")
    ).encode()

    encoded = base64.urlsafe_b64encode(
        raw
    ).decode().rstrip("=")

    signature = hmac.new(
        SECRET.encode(),
        encoded.encode(),
        hashlib.sha256
    ).hexdigest()

    return encoded + "." + signature


def verify_token(token):

    try:

        encoded, signature = token.split(".", 1)

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

        padding = "=" * (-len(encoded) % 4)

        raw = base64.urlsafe_b64decode(
            encoded + padding
        )

        payload = json.loads(
            raw.decode()
        )

        if payload["exp"] < int(
            datetime.now().timestamp()
        ):
            return False

        return True

    except Exception:
        return False


def require_auth(
    credentials: HTTPAuthorizationCredentials = Depends(security)
):

    if credentials is None:

        raise HTTPException(
            status_code=401,
            detail="Authentication required"
        )

    if credentials.scheme.lower() != "bearer":

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

    if not SUPABASE_URL or not SUPABASE_KEY:

        raise HTTPException(
            status_code=500,
            detail="Supabase environment variables missing"
        )

    return {
        "apikey": SUPABASE_KEY,
        "Authorization":
            f"Bearer {SUPABASE_KEY}",
        "Content-Type":
            "application/json"
    }


def get_all_rows():

    url = (
        f"{SUPABASE_URL}"
        f"/rest/v1/{TABLE}"
    )

    headers = supabase_headers()

    all_rows = []

    offset = 0
    page_size = 1000

    while True:

        response = requests.get(

            url,

            headers=headers,

            params={
                "select": "*",
                "offset": offset,
                "limit": page_size
            },

            timeout=60

        )

        if not response.ok:

            raise HTTPException(

                status_code=500,

                detail=(
                    "Supabase read error: "
                    + response.text
                )

            )

        rows = response.json()

        if not rows:
            break

        all_rows.extend(rows)

        if len(rows) < page_size:
            break

        offset += page_size

    return all_rows


# =========================================================
# DATABASE ROW CONVERSION
# =========================================================

def convert_database_row(item):

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

    if not isinstance(headers, list):
        headers = []

    if not isinstance(row, list):
        row = []

    return {

        "sheet": sheet_name,

        "headers": headers,

        "row": row,

        "data": {
            str(headers[i]):
                row[i] if i < len(row) else None
            for i in range(len(headers))
        }

    }


# =========================================================
# HOME
# =========================================================

@app.get("/")
def home():

    return {

        "status": "online",

        "message":
            "Rahul Software API is running",

        "database":
            "Supabase"

    }


# =========================================================
# LOGIN
# =========================================================

@app.post("/login")
def login(data: dict):

    username = text(
        data.get("username")
    )

    password = text(
        data.get("password")
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

            "status": "success",

            "access_token": token,

            "token": token,

            "token_type": "bearer"

        }

    raise HTTPException(

        status_code=401,

        detail=
            "Invalid username or password"

    )


# =========================================================
# SHEETS
# =========================================================

@app.get("/sheets")
def sheets(
    authenticated: bool = Depends(
        require_auth
    )
):

    try:

        rows = get_all_rows()

        names = set()

        for item in rows:

            sheet_name = text(
                item.get("sheet")
            )

            if sheet_name:
                names.add(sheet_name)

        names = sorted(names)

        return {

            "count": len(names),

            "sheets": names

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

@app.get("/sheet/{sheet_name}")
def get_sheet(

    sheet_name: str,

    limit: int = 100,

    authenticated: bool = Depends(
        require_auth
    )

):

    try:

        rows = get_all_rows()

        results = []

        headers = []

        for item in rows:

            if text(
                item.get("sheet")
            ) != sheet_name:

                continue

            converted = convert_database_row(
                item
            )

            if not headers:

                headers = converted[
                    "headers"
                ]

            results.append(
                converted["data"]
            )

        if not results and not headers:

            raise HTTPException(

                status_code=404,

                detail=
                    f"Sheet '{sheet_name}' not found"

            )

        limit = max(
            1,
            min(
                int(limit),
                1000
            )
        )

        results = results[:limit]

        return {

            "sheet": sheet_name,

            "headers": headers,

            "count": len(results),

            "data": results

        }

    except HTTPException:
        raise

    except Exception as e:

        raise HTTPException(

            status_code=500,

            detail=str(e)

        )


# =========================================================
# GLOBAL SEARCH
# =========================================================

@app.get("/search")
def search(

    q: str,

    sheet: str = "ALL SHEETS",

    authenticated: bool = Depends(
        require_auth
    )

):

    try:

        search_text = text(q).lower()

        if not search_text:

            return {

                "query": q,

                "sheet": sheet,

                "count": 0,

                "results": []

            }

        rows = get_all_rows()

        results = []

        for item in rows:

            sheet_name = text(
                item.get("sheet")
            )

            if (
                sheet != "ALL SHEETS"
                and sheet_name != sheet
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
                    in text(value).lower()
                ):

                    found = True
                    break

            if found:

                results.append({

                    "sheet":
                        sheet_name,

                    "headers":
                        converted["headers"],

                    "row":
                        converted["row"],

                    "data":
                        record

                })

        return {

            "query": q,

            "sheet": sheet,

            "count": len(results),

            "results": results

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

    authenticated: bool = Depends(
        require_auth
    )

):

    try:

        rows = get_all_rows()

        today = datetime.now().date()

        orders = []

        for item in rows:

            if text(
                item.get("sheet")
            ) == "ORDERS":

                converted = convert_database_row(
                    item
                )

                orders.append({

                    "data":
                        converted["data"],

                    "row":
                        converted["row"]

                })


        order_quantity = 0

        order_numbers = set()

        today_orders = 0

        today_order_quantity = 0

        size_data = {}

        party_data = {}


        for order_item in orders:

            row = order_item["row"]

            record = order_item["data"]


            order_no = get_row_position(
                row,
                1
            )


            party = text(
                get_row_position(
                    row,
                    3
                )
            )


            raw_size = get_row_position(
                row,
                4
            )

            size = normalize_size(
                raw_size
            )


            quantity = number(
                get_row_position(
                    row,
                    6
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
                    text(order_no)
                )


            order_quantity += quantity


            if row_date == today:

                today_orders += 1

                today_order_quantity += (
                    quantity
                )


            if size:

                if size not in size_data:

                    size_data[size] = {

                        "size": size,

                        "orders": 0,

                        "quantity": 0,

                        "order_numbers": set()

                    }


                size_data[size][
                    "orders"
                ] += 1


                size_data[size][
                    "quantity"
                ] += quantity


                if order_no not in (
                    None,
                    ""
                ):

                    size_data[size][
                        "order_numbers"
                    ].add(
                        text(order_no)
                    )


            # =================================================
            # PARTY -> DESTINATION -> SIZE -> QUANTITY
            # =================================================

            if party:

                if party not in party_data:

                    party_data[party] = {

                        "party":
                            party,

                        "orders":
                            0,

                        "quantity":
                            0,

                        "destinations":
                            {}

                    }


                party_data[party][
                    "orders"
                ] += 1


                party_data[party][
                    "quantity"
                ] += quantity


                destination_name = (

                    destination

                    if destination

                    else
                    "NO DESTINATION"

                )


                if destination_name not in (
                    party_data[
                        party
                    ]["destinations"]
                ):

                    party_data[
                        party
                    ]["destinations"][destination_name] = {

                        "destination":
                            destination_name,

                        "orders":
                            0,

                        "quantity":
                            0,

                        "sizes":
                            {}

                    }


                destination_data = (
                    party_data[
                        party
                    ]["destinations"][destination_name]
                )


                destination_data[
                    "orders"
                ] += 1


                destination_data[
                    "quantity"
                ] += quantity


                size_name = (
                    size
                    if size
                    else "UNKNOWN"
                )


                if size_name not in (
                    destination_data[
                        "sizes"
                    ]
                ):

                    destination_data[
                        "sizes"
                    ][size_name] = {

                        "size":
                            size_name,

                        "orders":
                            0,

                        "quantity":
                            0

                    }


                size_data_item = (
                    destination_data[
                        "sizes"
                    ][size_name]
                )


                size_data_item[
                    "orders"
                ] += 1


                size_data_item[
                    "quantity"
                ] += quantity


        # =========================================================
        # DISPATCH
        # =========================================================

        dispatch = []

        for item in rows:

            if text(
                item.get("sheet")
            ) == "Dispatched Orders":

                dispatch.append(
                    convert_database_row(
                        item
                    )
                )


        dispatch_quantity = 0

        today_dispatch = 0

        today_dispatch_quantity = 0


        for dispatch_item in dispatch:

            row = dispatch_item["row"]

            record = dispatch_item["data"]


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


            dispatch_quantity += quantity


            row_date = date_only(

                first_existing(

                    record,

                    [
                        "Date",
                        "BILL DATE",
                        "IN/OUT DATE"
                    ]

                )

            )


            if row_date == today:

                today_dispatch += 1

                today_dispatch_quantity += (
                    quantity
                )


        # =========================================================
        # HOLD ORDERS
        # =========================================================

        hold = []

        for item in rows:

            if text(
                item.get("sheet")
            ) == "Holded Orders":

                hold.append(
                    convert_database_row(
                        item
                    )["data"]
                )


        hold_quantity = 0


        for row in hold:

            hold_quantity += number(

                first_existing(

                    row,

                    [
                        "Quantity Pcs.",
                        "Quantity",
                        "QUANTITY"
                    ]

                )

            )


        # =========================================================
        # SIZE WISE FINAL
        # =========================================================

        size_wise = []


        for item in size_data.values():

            unique_size_orders = len(
                item["order_numbers"]
            )

            if unique_size_orders == 0:

                unique_size_orders = (
                    item["orders"]
                )


            size_wise.append({

                "size":
                    item["size"],

                "orders":
                    unique_size_orders,

                "quantity":
                    clean_number(
                        item["quantity"]
                    )

            })


        size_wise.sort(

            key=lambda x:
                x["quantity"],

            reverse=True

        )


        # =========================================================
        # PARTY WISE FINAL
        # =========================================================

        party_wise = []


        for item in party_data.values():

            destinations = []


            for destination_item in (
                item[
                    "destinations"
                ].values()
            ):

                sizes = []


                for size_item in (
                    destination_item[
                        "sizes"
                    ].values()
                ):

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
                    item["party"],

                "orders":
                    item["orders"],

                "quantity":
                    clean_number(
                        item["quantity"]
                    ),

                "destinations":
                    destinations

            })


        party_wise.sort(

            key=lambda x:
                x["quantity"],

            reverse=True

        )


        # =========================================================
        # FINAL DASHBOARD RESPONSE
        # =========================================================

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
                    len(order_numbers),

                "total_quantity":
                    clean_number(
                        order_quantity
                    )

            },


            "dispatch": {

                "total_rows":
                    len(dispatch),

                "total_quantity":
                    clean_number(
                        dispatch_quantity
                    ),

                "today_rows":
                    today_dispatch,

                "today_quantity":
                    clean_number(
                        today_dispatch_quantity
                    )

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
# MONTHLY REPORT - AVAILABLE MONTHS
# =========================================================

@app.get("/monthly-report/months")
def monthly_report_months(

    authenticated: bool = Depends(
        require_auth
    )

):

    try:

        rows = get_all_rows()

        months = set()


        for item in rows:

            if text(
                item.get("sheet")
            ) != "Dispatched Orders":

                continue


            converted = convert_database_row(
                item
            )

            row = converted[
                "row"
            ]


            # COLUMN K = DATE

            row_date = date_only(
                get_row_position(
                    row,
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
# MONTHLY REPORT - DAILY PCS DISPATCH
#
# K = DATE
# M = MATERIAL
# F = QUANTITY
#
# SAME DATE + SAME MATERIAL = PLUS
# =========================================================

@app.get("/monthly-report/daily-pcs-dispatch")
def daily_pcs_dispatch(

    year: int,

    month: int,

    authenticated: bool = Depends(
        require_auth
    )

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


        rows = get_all_rows()


        days_in_month = calendar.monthrange(
            year,
            month
        )[1]


        # =========================================================
        # DAILY DATA
        # =========================================================

        daily_data = {}

        for day_number in range(
            1,
            days_in_month + 1
        ):

            daily_data[day_number] = {
                "date": day_number
            }

            for material in MONTHLY_MATERIALS:

                daily_data[day_number][
                    material
                ] = 0


        matched_rows = 0

        skipped_unknown_material = 0

        skipped_no_date = 0


        # =========================================================
        # READ DISPATCHED ORDERS
        # =========================================================

        for item in rows:

            if text(
                item.get("sheet")
            ) != "Dispatched Orders":

                continue


            converted = convert_database_row(
                item
            )

            row = converted[
                "row"
            ]


            # =====================================================
            # K = DATE
            # =====================================================

            row_date = date_only(

                get_row_position(
                    row,
                    11
                )

            )


            if row_date is None:

                skipped_no_date += 1

                continue


            if row_date.year != year:
                continue

            if row_date.month != month:
                continue


            # =====================================================
            # M = MATERIAL
            # =====================================================

            material_type = text(

                get_row_position(
                    row,
                    13
                )

            )


            if not material_type:

                skipped_unknown_material += 1

                continue


            normalized = normalize_material(
                material_type
            )


            # =====================================================
            # MAP MATERIAL
            # =====================================================

            report_column = (
                MATERIAL_MAPPING.get(
                    normalized
                )
            )


            if not report_column:

                skipped_unknown_material += 1

                continue


            # =====================================================
            # F = QUANTITY
            # =====================================================

            quantity = number(

                get_row_position(
                    row,
                    6
                )

            )


            # =====================================================
            # SAME DATE + SAME MATERIAL = PLUS
            # =====================================================

            daily_data[
                row_date.day
            ][
                report_column
            ] += quantity


            matched_rows += 1


        # =========================================================
        # TOTAL
        # =========================================================

        total = {
            "date": "TOTAL"
        }


        for material in MONTHLY_MATERIALS:

            value = 0

            for day_number in daily_data:

                value += number(

                    daily_data[
                        day_number
                    ][material]

                )

            total[material] = clean_number(
                value
            )


        # =========================================================
        # DAYS
        # =========================================================

        days = []


        for day_number in range(
            1,
            days_in_month + 1
        ):

            item = {
                "date": day_number
            }


            for material in MONTHLY_MATERIALS:

                item[material] = clean_number(

                    daily_data[
                        day_number
                    ][material]

                )


            days.append(item)


        # =========================================================
        # TITLE
        # =========================================================

        month_title = (

            f"{calendar.month_name[month].upper()}"
            f"- {year} MONTHLY PCS DISPATCH QUANTITY"

        )


        return {

            "year":
                year,

            "month":
                month,

            "month_name":
                calendar.month_name[month],

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

            "skipped_unknown_material":
                skipped_unknown_material,

            "skipped_no_date":
                skipped_no_date,

            "source_columns": {

                "date":
                    "K",

                "material":
                    "M",

                "quantity":
                    "F"

            },

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
# CURRENT MONTH SHORTCUT
# =========================================================

@app.get("/monthly-report/daily-pcs-dispatch/current")
def current_month_daily_pcs_dispatch(

    authenticated: bool = Depends(
        require_auth
    )

):

    today = datetime.now().date()

    return daily_pcs_dispatch(

        year=today.year,

        month=today.month,

        authenticated=True

    )


# =========================================================
# MATERIAL MAPPING CHECK
# =========================================================

@app.get("/monthly-report/material-mapping")
def material_mapping(

    authenticated: bool = Depends(
        require_auth
    )

):

    try:

        rows = get_all_rows()

        actual_types = {}


        for item in rows:

            if text(
                item.get("sheet")
            ) != "Dispatched Orders":

                continue


            converted = convert_database_row(
                item
            )

            row = converted[
                "row"
            ]


            # M = TYPE OF MATERIAL

            material_type = text(

                get_row_position(
                    row,
                    13
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

                    "normalized":
                        normalized,

                    "report_column":
                        mapped_to,

                    "mapped":
                        bool(mapped_to)

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
