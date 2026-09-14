from fastapi import FastAPI, HTTPException, Depends
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


TABLE = "excel_rows"


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

    if isinstance(
        value,
        datetime
    ):
        return value.date()

    if isinstance(
        value,
        date
    ):
        return value

    value = str(
        value
    ).strip()


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


    parts =
        value.split()


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

    index =
        position - 1


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


    raw =
        json.dumps(
            payload,
            separators=(
                ",",
                ":"
            )
        ).encode()


    encoded =
        base64.urlsafe_b64encode(
            raw
        ).decode().rstrip("=")


    signature =
        hmac.new(
            SECRET.encode(),
            encoded.encode(),
            hashlib.sha256
        ).hexdigest()


    return (
        encoded +
        "." +
        signature
    )


def verify_token(
    token
):

    try:

        encoded, signature =
            token.split(
                ".",
                1
            )


        expected =
            hmac.new(
                SECRET.encode(),
                encoded.encode(),
                hashlib.sha256
            ).hexdigest()


        if not hmac.compare_digest(
            signature,
            expected
        ):

            return False


        padding =
            "=" * (
                -len(encoded) % 4
            )


        raw =
            base64.urlsafe_b64decode(
                encoded +
                padding
            )


        payload =
            json.loads(
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
        HTTPAuthorizationCredentials =
        Depends(security)
):

    if credentials is None:

        raise HTTPException(
            status_code=401,
            detail=
                "Authentication required"
        )


    if (
        credentials.scheme.lower()
        !=
        "bearer"
    ):

        raise HTTPException(
            status_code=401,
            detail=
                "Invalid authentication"
        )


    if not verify_token(
        credentials.credentials
    ):

        raise HTTPException(
            status_code=401,
            detail=
                "Invalid or expired token"
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
            detail=
                "Supabase environment variables missing"
        )


    return {

        "apikey":
            SUPABASE_KEY,

        "Authorization":
            "Bearer " +
            SUPABASE_KEY,

        "Content-Type":
            "application/json",

        "Prefer":
            "return=representation"

    }


def get_all_rows():

    url =
        f"{SUPABASE_URL}/rest/v1/{TABLE}"


    headers =
        supabase_headers()


    all_rows = []


    offset = 0

    page_size = 1000


    while True:

        response =
            requests.get(

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

                detail=
                    "Supabase read error: " +
                    response.text

            )


        rows =
            response.json()


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

    sheet_name =
        item.get(
            "sheet",
            ""
        )


    headers =
        item.get(
            "headers",
            []
        )


    row =
        item.get(
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
        ] =
            clean_value(
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
            "Supabase"

    }


# =========================================================
# LOGIN
# =========================================================

@app.post("/login")
def login(
    data: dict
):

    username =
        text(
            data.get(
                "username"
            )
        )


    password =
        text(
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

        token =
            make_token(
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

        detail=
            "Invalid username or password"

    )


# =========================================================
# SHEETS
# =========================================================

@app.get("/sheets")
def sheets(
    authenticated:
        bool =
        Depends(require_auth)
):

    try:

        rows =
            get_all_rows()


        names =
            set()


        for item in rows:

            name =
                text(
                    item.get(
                        "sheet"
                    )
                )


            if name:
                names.add(
                    name
                )


        names =
            sorted(
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
        bool =
        Depends(require_auth)

):

    try:

        rows =
            get_all_rows()


        results = []

        headers = []


        for item in rows:

            if (
                text(
                    item.get(
                        "sheet"
                    )
                )
                !=
                sheet_name
            ):
                continue


            converted =
                convert_database_row(
                    item
                )


            if not headers:

                headers =
                    converted[
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

                detail=
                    f"Sheet '{sheet_name}' not found"

            )


        limit =
            max(
                1,
                min(
                    int(limit),
                    1000
                )
            )


        results =
            results[
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
# SEARCH
# =========================================================

@app.get("/search")
def search(

    q: str,

    sheet:
        str =
        "ALL SHEETS",

    authenticated:
        bool =
        Depends(require_auth)

):

    try:

        search_text =
            text(
                q
            ).lower()


        if not search_text:

            return {

                "query":
                    q,

                "sheet":
                    sheet,

                "count":
                    0,

                "results":
                    []

            }


        rows =
            get_all_rows()


        results = []


        for item in rows:

            sheet_name =
                text(
                    item.get(
                        "sheet"
                    )
                )


            if (
                sheet !=
                "ALL SHEETS"
                and
                sheet_name !=
                sheet
            ):

                continue


            converted =
                convert_database_row(
                    item
                )


            record =
                converted[
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
        bool =
        Depends(require_auth)
):

    try:

        rows =
            get_all_rows()


        today =
            datetime.now().date()


        orders = []

        dispatch = []

        hold = []


        for item in rows:

            sheet =
                text(
                    item.get(
                        "sheet"
                    )
                )


            converted =
                convert_database_row(
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

        order_numbers =
            set()

        today_orders = 0

        today_order_quantity = 0

        size_data = {}

        party_data = {}


        for item in orders:

            row =
                item["row"]

            record =
                item["data"]


            order_no =
                first_existing(
                    record,
                    [
                        "Order No.",
                        "Order No"
                    ]
                )


            party =
                text(
                    first_existing(
                        record,
                        [
                            "PARTY NAME",
                            "Party Name",
                            "PARTY"
                        ]
                    )
                )


            size =
                normalize_size(
                    first_existing(
                        record,
                        [
                            "Size",
                            "SIZE"
                        ]
                    )
                )


            quantity =
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


            row_date =
                date_only(
                    first_existing(
                        record,
                        [
                            "Date",
                            "DATE"
                        ]
                    )
                )


            destination =
                text(
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

                today_order_quantity += \
                    quantity


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
                ]["quantity"] += \
                    quantity


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
                ]["quantity"] += \
                    quantity


                destination_key =
                    destination or "-"


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


                destination_item =
                    party_data[
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


                size_key =
                    size or "-"


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
                ]["quantity"] += \
                    quantity


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


            row =
                item["row"]

            record =
                item["data"]


            quantity =
                number(
                    get_row_position(
                        row,
                        6
                    )
                )


            row_date =
                date_only(
                    get_row_position(
                        row,
                        11
                    )
                )


            dispatch_quantity += \
                quantity


            if row_date == today:

                today_dispatch += 1

                today_dispatch_quantity += \
                    quantity


        hold_quantity = 0


        for item in hold:

            record =
                item["data"]


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

            item["quantity"] =
                clean_number(
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
                    -
                    cancelled_dispatch_rows,

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
# MONTHS
# =========================================================

@app.get(
    "/monthly-report/months"
)
def monthly_report_months(
    authenticated:
        bool =
        Depends(require_auth)
):

    try:

        rows =
            get_all_rows()


        months =
            set()


        for item in rows:

            if (
                text(
                    item.get(
                        "sheet"
                    )
                )
                !=
                "Dispatched Orders"
            ):

                continue


            if item.get(
                "cancelled",
                False
            ) is True:

                continue


            converted =
                convert_database_row(
                    item
                )


            record =
                converted[
                    "data"
                ]


            row_date =
                date_only(
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

    authenticated:
        bool =
        Depends(require_auth)

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


        rows =
            get_all_rows()


        days_in_month =
            calendar.monthrange(
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

            if (
                text(
                    item.get(
                        "sheet"
                    )
                )
                !=
                "Dispatched Orders"
            ):

                continue


            if item.get(
                "cancelled",
                False
            ) is True:

                skipped_cancelled_rows += 1

                continue


            converted =
                convert_database_row(
                    item
                )


            row =
                converted[
                    "row"
                ]


            row_date =
                date_only(
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


            material_type =
                text(
                    get_row_position(
                        row,
                        13
                    )
                )


            normalized =
                normalize_material(
                    material_type
                )


            report_column =
                MATERIAL_MAPPING.get(
                    normalized
                )


            if not report_column:
                continue


            quantity =
                number(
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
            ] =
                clean_number(
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
                ] =
                    clean_number(
                        daily_data[
                            day_number
                        ][material]
                    )


            days.append(
                item
            )


        month_title =
            f"{calendar.month_name[month].upper()}- {year} MONTHLY PCS DISPATCH QUANTITY"


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
        bool =
        Depends(require_auth)
):

    today =
        datetime.now().date()


    return daily_pcs_dispatch(

        year=
            today.year,

        month=
            today.month,

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
        bool =
        Depends(require_auth)
):

    try:

        rows =
            get_all_rows()


        actual_types = {}


        for item in rows:

            if (
                text(
                    item.get(
                        "sheet"
                    )
                )
                !=
                "Dispatched Orders"
            ):

                continue


            converted =
                convert_database_row(
                    item
                )


            record =
                converted[
                    "data"
                ]


            material_type =
                text(
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


            normalized =
                normalize_material(
                    material_type
                )


            mapped_to =
                MATERIAL_MAPPING.get(
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
            !=
            "PARTIES"
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
        bool =
        Depends(require_auth)
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

    s =
        text(
            value
        ).upper()


    s =
        s.replace(
            " ",
            ""
        )


    s =
        s.replace(
            "×",
            "X"
        )


    if s.count("X") != 1:

        raise ValueError(
            "Size format: 1200+75X1000+75"
        )


    left, right =
        s.split(
            "X",
            1
        )


    def nums(part):

        values =
            part.split(
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


    a =
        nums(
            left
        )


    b =
        nums(
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
        bool =
        Depends(require_auth)

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

    rows =
        get_all_rows()


    highest =
        408


    for item in rows:

        if (
            text(
                item.get(
                    "sheet"
                )
            )
            !=
            "Dispatched Orders"
        ):

            continue


        converted =
            convert_database_row(
                item
            )


        record =
            converted[
                "data"
            ]


        invoice =
            text(
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


        value =
            invoice.strip()


        if not value.upper().startswith(
            prefix.upper()
        ):

            continue


        tail =
            value[
                len(prefix):
            ].strip()


        if tail.isdigit():

            highest =
                max(
                    highest,
                    int(tail)
                )


    return (
        prefix +
        str(
            highest + 1
        )
    )


@app.post(
    "/test-report/reserve-bill"
)
def reserve_test_bill(

    data: dict,

    authenticated:
        bool =
        Depends(require_auth)

):

    try:

        mode =
            text(
                data.get(
                    "mode"
                )
                or
                "AUTO NEXT"
            ).upper()


        manual_bill =
            text(
                data.get(
                    "bill_no"
                )
            )


        if mode == "MANUAL":

            if not manual_bill:

                raise HTTPException(
                    status_code=400,
                    detail=
                        "Manual Bill No. required"
                )


            return {

                "bill_no":
                    manual_bill

            }


        if mode != "AUTO NEXT":

            raise HTTPException(
                status_code=400,
                detail=
                    "Invalid bill mode"
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

    value =
        float(
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
        +
        rand_int(
            -5,
            5
        )
    )


def actual_tab(
    value
):

    return fmt_mm(
        float(value)
        +
        rand_int(
            -3,
            3
        )
    )


def actual_thickness(
    value
):

    value =
        float(value)


    actual =
        value +
        random.uniform(
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

    t =
        float(
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

    standard =
        tc_standard(
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
                ) +
                "mm",

                "±10mm"
            )


        if name == "Total Width":

            return (
                fmt_mm(
                    parsed["width"]
                ) +
                "mm",

                "±10mm"
            )


        if name == "1 Tab at Length":

            value =
                parsed.get(
                    "tab_length",
                    parsed.get(
                        "tab_length_1"
                    )
                )


            return (
                fmt_mm(
                    value
                ) +
                "mm",

                "±5mm"
            )


        if name == "2 Tab at Length":

            return (
                fmt_mm(
                    parsed[
                        "tab_length_2"
                    ]
                ) +
                "mm",

                "±5mm"
            )


        if name == "1 Tab at Width":

            value =
                parsed.get(
                    "tab_width",
                    parsed.get(
                        "tab_width_1"
                    )
                )


            return (
                fmt_mm(
                    value
                ) +
                "mm",

                "±5mm"
            )


        if name == "2 Tab at Width":

            return (
                fmt_mm(
                    parsed[
                        "tab_width_2"
                    ]
                ) +
                "mm",

                "±5mm"
            )


        if name == "Thickness":

            return (
                fmt_mm(
                    thickness
                ) +
                "mm",

                "±0.3mm"
            )


        if name == "Moisture":

            return (
                "10 % to 12 %",
                "±2 %"
            )


        if name == "TC":

            standard =
                tc_standard(
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

        standard, tolerance =
            standard_for(
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
        bool =
        Depends(require_auth)

):

    try:

        party =
            text(
                data.get(
                    "party"
                )
            )


        address =
            text(
                data.get(
                    "address"
                )
            )


        size =
            text(
                data.get(
                    "size"
                )
            )


        thickness =
            number(
                data.get(
                    "thickness"
                )
            )


        gsm =
            text(
                data.get(
                    "gsm"
                )
            )


        po_mode =
            text(
                data.get(
                    "po_mode"
                )
                or
                "AS PER SHEET"
            ).upper()


        po_no =
            text(
                data.get(
                    "po_no"
                )
            )


        po_date =
            text(
                data.get(
                    "po_date"
                )
            )


        bill_mode =
            text(
                data.get(
                    "bill_mode"
                )
                or
                "AUTO NEXT"
            ).upper()


        manual_bill =
            text(
                data.get(
                    "bill_no"
                )
            )


        bill_date =
            text(
                data.get(
                    "bill_date"
                )
            )


        drawing =
            text(
                data.get(
                    "drawing"
                )
            )


        specification =
            text(
                data.get(
                    "specification"
                )
            )


        if not party:

            raise HTTPException(
                status_code=400,
                detail=
                    "Party select karo"
            )


        if not size:

            raise HTTPException(
                status_code=400,
                detail=
                    "Size bharo"
            )


        if thickness <= 0:

            raise HTTPException(
                status_code=400,
                detail=
                    "Thickness sahi bharo"
            )


        if po_mode not in {

            "AS PER SHEET",
            "VERBAL",
            "MANUAL"

        }:

            raise HTTPException(
                status_code=400,
                detail=
                    "Invalid PO mode"
            )


        if (
            po_mode == "MANUAL"
            and
            not po_no
        ):

            raise HTTPException(
                status_code=400,
                detail=
                    "Manual PO No. bharo"
            )


        if bill_mode not in {

            "AUTO NEXT",
            "MANUAL"

        }:

            raise HTTPException(
                status_code=400,
                detail=
                    "Invalid bill mode"
            )


        parsed =
            parse_size_pattern(
                size
            )


        if bill_mode == "AUTO NEXT":

            bill_no =
                next_bill_number(
                    "MP/26-27/"
                )

        else:

            if not manual_bill:

                raise HTTPException(
                    status_code=400,
                    detail=
                        "Manual Bill No. bharo"
                )


            bill_no =
                manual_bill


        if po_mode == "AS PER SHEET":

            display_po =
                "AS PER SHEET"

        elif po_mode == "VERBAL":

            display_po =
                "VERBAL"

        else:

            display_po =
                po_no


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
