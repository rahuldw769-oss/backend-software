from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
import os
import requests
import jwt

app = FastAPI(title="Rahul Software API")

SUPABASE_URL = os.getenv("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
APP_USERNAME = os.getenv("APP_USERNAME", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "")
JWT_SECRET = os.getenv("JWT_SECRET", "")
TABLE = "excel_rows"

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

security = HTTPBearer()


class LoginData(BaseModel):
    username: str
    password: str


@app.post("/login")
def login(data: LoginData):

    if not APP_USERNAME or not APP_PASSWORD or not JWT_SECRET:
        raise HTTPException(
            status_code=500,
            detail="Login environment variables are not configured"
        )

    if data.username != APP_USERNAME or data.password != APP_PASSWORD:
        raise HTTPException(
            status_code=401,
            detail="Invalid username or password"
        )

    expire = datetime.now(timezone.utc) + timedelta(hours=12)

    token = jwt.encode(
        {
            "sub": data.username,
            "exp": expire
        },
        JWT_SECRET,
        algorithm="HS256"
    )

    return {
        "success": True,
        "token": token
    }


def verify_token(
    credentials: HTTPAuthorizationCredentials = Depends(security)
):

    token = credentials.credentials

    try:

        payload = jwt.decode(
            token,
            JWT_SECRET,
            algorithms=["HS256"]
        )

        username = payload.get("sub")

        if not username:
            raise HTTPException(
                status_code=401,
                detail="Invalid token"
            )

        return username

    except jwt.ExpiredSignatureError:

        raise HTTPException(
            status_code=401,
            detail="Token expired"
        )

    except jwt.InvalidTokenError:

        raise HTTPException(
            status_code=401,
            detail="Invalid token"
        )


def supabase_get(params=None, offset=0):

    if not SUPABASE_URL or not SUPABASE_KEY:
        raise HTTPException(
            status_code=500,
            detail="Cloud database is not configured"
        )

    url = f"{SUPABASE_URL}/rest/v1/{TABLE}"

    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}"
    }

    if params is None:
        params = {}

    params = params.copy()
    params["offset"] = offset
    params["limit"] = 1000

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=30
    )

    if not response.ok:

        raise HTTPException(
            status_code=response.status_code,
            detail=response.text
        )

    return response.json()


def supabase_get_all(params=None):

    all_rows = []
    offset = 0

    while True:

        rows = supabase_get(params, offset)

        if not rows:
            break

        all_rows.extend(rows)

        if len(rows) < 1000:
            break

        offset += 1000

    return all_rows


def normalize_header(value):

    if value is None:
        return ""

    return str(value).strip().lower()


def get_column_index(headers, possible_names):

    normalized = [
        normalize_header(x)
        for x in headers
    ]

    for name in possible_names:

        name = normalize_header(name)

        if name in normalized:
            return normalized.index(name)

    return None


def parse_date(value):

    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    # ISO format coming from sync_excel.py
    try:
        return datetime.fromisoformat(
            text.replace("Z", "+00:00")
        ).date()
    except Exception:
        pass

    # Other common formats
    formats = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%d.%m.%Y"
    ]

    for fmt in formats:

        try:
            return datetime.strptime(
                text[:10],
                fmt
            ).date()
        except Exception:
            continue

    return None


def number_value(value):

    if value is None:
        return 0

    try:
        if isinstance(value, bool):
            return 0

        return float(value)

    except Exception:
        try:
            text = str(value).replace(",", "").strip()
            return float(text)
        except Exception:
            return 0


def clean_size(value):

    if value is None:
        return "Not specified"

    text = str(value).strip()

    if not text:
        return "Not specified"

    return text


def clean_party(value):

    if value is None:
        return "Not specified"

    text = str(value).strip()

    if not text:
        return "Not specified"

    return text


def build_sheet_info(sheet_rows):

    if not sheet_rows:
        return None

    headers = sheet_rows[0].get("headers", [])

    return {
        "headers": headers,
        "rows": [
            item.get("row", [])
            for item in sheet_rows
        ]
    }


@app.get("/")
def home():

    return {
        "status": "online",
        "app": "Rahul Software API"
    }


@app.get("/sheets")
def sheets(username: str = Depends(verify_token)):

    names = set()

    rows = supabase_get_all({
        "select": "sheet"
    })

    for item in rows:

        sheet = item.get("sheet")

        if sheet:
            names.add(str(sheet))

    return {
        "count": len(names),
        "sheets": sorted(names)
    }


@app.get("/sheet/{sheet_name}")
def get_sheet(
    sheet_name: str,
    limit: int = 100,
    username: str = Depends(verify_token)
):

    limit = min(limit, 1000)

    rows = supabase_get(
        {
            "sheet": f"eq.{sheet_name}",
            "select": "headers,row"
        },
        0
    )

    rows = rows[:limit]

    return {
        "sheet": sheet_name,
        "count": len(rows),
        "data": rows
    }


@app.get("/search")
def search(
    q: str,
    sheet: str = "ALL SHEETS",
    limit: int = 100,
    username: str = Depends(verify_token)
):

    q = q.strip().lower()

    if not q:

        return {
            "count": 0,
            "results": []
        }

    results = []
    offset = 0

    while True:

        params = {
            "select": "sheet,headers,row"
        }

        if sheet != "ALL SHEETS":

            params["sheet"] = f"eq.{sheet}"

        rows = supabase_get(
            params,
            offset
        )

        if not rows:
            break

        for item in rows:

            row = item.get("row", [])

            found = False

            for cell in row:

                if cell is not None:

                    if q in str(cell).lower():

                        found = True
                        break

            if found:

                results.append(item)

                if len(results) >= limit:

                    return {
                        "count": len(results),
                        "results": results
                    }

        if len(rows) < 1000:
            break

        offset += 1000

    return {
        "count": len(results),
        "results": results
    }


# ==========================================================
# DASHBOARD
# ==========================================================

@app.get("/dashboard")
def dashboard(username: str = Depends(verify_token)):

    # India date
    india_now = datetime.now(
        ZoneInfo("Asia/Kolkata")
    )

    today = india_now.date()

    # ------------------------------------------------------
    # ORDERS
    # ------------------------------------------------------

    order_rows = supabase_get_all({
        "sheet": "eq.ORDERS",
        "select": "headers,row"
    })

    total_orders = 0
    total_quantity = 0

    today_orders = 0
    today_quantity = 0

    size_data = {}
    party_data = {}

    unique_order_numbers = set()
    today_order_numbers = set()

    if order_rows:

        headers = order_rows[0].get(
            "headers",
            []
        )

        date_index = get_column_index(
            headers,
            ["Date"]
        )

        order_no_index = get_column_index(
            headers,
            ["Order No."]
        )

        size_index = get_column_index(
            headers,
            ["Size"]
        )

        quantity_index = get_column_index(
            headers,
            ["Quantity Pcs."]
        )

        party_index = get_column_index(
            headers,
            ["PARTY NAME", "Party Name"]
        )

        for item in order_rows:

            row = item.get("row", [])

            if not row:
                continue

            order_no = ""

            if (
                order_no_index is not None
                and order_no_index < len(row)
            ):
                order_no = str(
                    row[order_no_index]
                ).strip()

            if order_no:
                unique_order_numbers.add(
                    order_no
                )

            order_date = None

            if (
                date_index is not None
                and date_index < len(row)
            ):
                order_date = parse_date(
                    row[date_index]
                )

            quantity = 0

            if (
                quantity_index is not None
                and quantity_index < len(row)
            ):
                quantity = number_value(
                    row[quantity_index]
                )

            size = "Not specified"

            if (
                size_index is not None
                and size_index < len(row)
            ):
                size = clean_size(
                    row[size_index]
                )

            party = "Not specified"

            if (
                party_index is not None
                and party_index < len(row)
            ):
                party = clean_party(
                    row[party_index]
                )

            total_orders += 1
            total_quantity += quantity

            # Size wise
            if size not in size_data:

                size_data[size] = {
                    "size": size,
                    "orders": 0,
                    "quantity": 0
                }

            size_data[size]["orders"] += 1
            size_data[size]["quantity"] += quantity

            # Party wise
            if party not in party_data:

                party_data[party] = {
                    "party": party,
                    "orders": 0,
                    "quantity": 0
                }

            party_data[party]["orders"] += 1
            party_data[party]["quantity"] += quantity

            # Today
            if order_date == today:

                today_orders += 1
                today_quantity += quantity

                if order_no:
                    today_order_numbers.add(
                        order_no
                    )

    # ------------------------------------------------------
    # DISPATCH
    # ------------------------------------------------------

    dispatch_rows = supabase_get_all({
        "sheet": "eq.Dispatched Orders",
        "select": "headers,row"
    })

    dispatched_total = 0
    dispatched_total_quantity = 0

    dispatched_today = 0
    dispatched_today_quantity = 0

    if dispatch_rows:

        headers = dispatch_rows[0].get(
            "headers",
            []
        )

        # First Date column is the order date.
        # Second Date column is invoice/dispatch date.
        date_indexes = []

        for index, header in enumerate(headers):

            if normalize_header(header) == "date":
                date_indexes.append(index)

        dispatch_date_index = None

        if len(date_indexes) >= 2:
            dispatch_date_index = date_indexes[1]

        elif len(date_indexes) == 1:
            dispatch_date_index = date_indexes[0]

        quantity_index = get_column_index(
            headers,
            ["Quantity Pcs."]
        )

        for item in dispatch_rows:

            row = item.get("row", [])

            if not row:
                continue

            quantity = 0

            if (
                quantity_index is not None
                and quantity_index < len(row)
            ):
                quantity = number_value(
                    row[quantity_index]
                )

            dispatched_total += 1
            dispatched_total_quantity += quantity

            dispatch_date = None

            if (
                dispatch_date_index is not None
                and dispatch_date_index < len(row)
            ):
                dispatch_date = parse_date(
                    row[dispatch_date_index]
                )

            if dispatch_date == today:

                dispatched_today += 1
                dispatched_today_quantity += quantity

    # ------------------------------------------------------
    # HOLD ORDERS
    # ------------------------------------------------------

    hold_rows = supabase_get_all({
        "sheet": "eq.Holded Orders",
        "select": "headers,row"
    })

    hold_orders = 0
    hold_quantity = 0

    if hold_rows:

        headers = hold_rows[0].get(
            "headers",
            []
        )

        quantity_index = get_column_index(
            headers,
            ["Quantity Pcs."]
        )

        for item in hold_rows:

            row = item.get("row", [])

            if not row:
                continue

            quantity = 0

            if (
                quantity_index is not None
                and quantity_index < len(row)
            ):
                quantity = number_value(
                    row[quantity_index]
                )

            hold_orders += 1
            hold_quantity += quantity

    # ------------------------------------------------------
    # SORT DATA
    # ------------------------------------------------------

    size_wise = sorted(
        size_data.values(),
        key=lambda x: x["quantity"],
        reverse=True
    )

    party_wise = sorted(
        party_data.values(),
        key=lambda x: x["quantity"],
        reverse=True
    )

    # Keep dashboard lightweight
    party_wise = party_wise[:20]

    # ------------------------------------------------------
    # RESPONSE
    # ------------------------------------------------------

    return {

        "date": today.isoformat(),

        "today": {
            "orders": today_orders,
            "order_quantity": today_quantity,
            "dispatch": dispatched_today,
            "dispatch_quantity": dispatched_today_quantity
        },

        "orders": {
            "total_rows": total_orders,
            "unique_orders": len(unique_order_numbers),
            "total_quantity": total_quantity
        },

        "dispatch": {
            "total_rows": dispatched_total,
            "total_quantity": dispatched_total_quantity,
            "today_rows": dispatched_today,
            "today_quantity": dispatched_today_quantity
        },

        "hold": {
            "orders": hold_orders,
            "quantity": hold_quantity
        },

        "size_wise": size_wise,

        "party_wise": party_wise,

        "last_updated": india_now.isoformat()
    }
