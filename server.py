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

    formats = (
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
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


        # =================================================
        # ORDERS
        # =================================================

        orders = []

        for item in rows:

            if text(
                item.get("sheet")
            ) == "ORDERS":

                orders.append(
                    convert_database_row(
                        item
                    )["data"]
                )


        order_quantity = 0

        order_numbers = set()

        today_orders = 0

        today_order_quantity = 0

        size_data = {}

        party_data = {}


        for row in orders:

            order_no = first_existing(

                row,

                [
                    "Order No.",
                    "Order No"
                ]

            )

            quantity = number(

                first_existing(

                    row,

                    [
                        "Quantity Pcs.",
                        "Quantity",
                        "QUANTITY"
                    ]

                )

            )

            row_date = date_only(

                first_existing(

                    row,

                    [
                        "Date",
                        "DATE"
                    ]

                )

            )

            size = text(

                first_existing(

                    row,

                    [
                        "Size",
                        "SIZE"
                    ]

                )

            )

            party = text(

                first_existing(

                    row,

                    [
                        "PARTY NAME",
                        "Party Name",
                        "PARTY"
                    ]

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

                        "quantity": 0

                    }

                size_data[size][
                    "orders"
                ] += 1

                size_data[size][
                    "quantity"
                ] += quantity


            if party:

                if party not in party_data:

                    party_data[party] = {

                        "party": party,

                        "orders": 0,

                        "quantity": 0

                    }

                party_data[party][
                    "orders"
                ] += 1

                party_data[party][
                    "quantity"
                ] += quantity


        # =================================================
        # DISPATCH
        # =================================================

        dispatch = []

        for item in rows:

            if text(
                item.get("sheet")
            ) == "Dispatched Orders":

                dispatch.append(
                    convert_database_row(
                        item
                    )["data"]
                )


        dispatch_quantity = 0

        today_dispatch = 0

        today_dispatch_quantity = 0


        for row in dispatch:

            quantity = number(

                first_existing(

                    row,

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

                    row,

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


        # =================================================
        # HOLD ORDERS
        # =================================================

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


        # =================================================
        # SIZE WISE
        # =================================================

        size_wise = sorted(

            [

                {

                    "size":
                        item["size"],

                    "orders":
                        item["orders"],

                    "quantity":
                        clean_number(
                            item["quantity"]
                        )

                }

                for item in
                size_data.values()

            ],

            key=lambda x:
                x["quantity"],

            reverse=True

        )


        # =================================================
        # PARTY WISE
        # =================================================

        party_wise = sorted(

            [

                {

                    "party":
                        item["party"],

                    "orders":
                        item["orders"],

                    "quantity":
                        clean_number(
                            item["quantity"]
                        )

                }

                for item in
                party_data.values()

            ],

            key=lambda x:
                x["quantity"],

            reverse=True

        )


        # =================================================
        # FINAL RESPONSE
        # =================================================

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
