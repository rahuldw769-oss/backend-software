from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel
from datetime import datetime, timedelta, timezone
import os
import requests
import jwt


app = FastAPI(title="Rahul Software API")


# =========================
# ENVIRONMENT VARIABLES
# =========================

SUPABASE_URL = os.getenv("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

APP_USERNAME = os.getenv("APP_USERNAME", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "")
JWT_SECRET = os.getenv("JWT_SECRET", "")

TABLE = "excel_rows"


# =========================
# CORS
# =========================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =========================
# LOGIN
# =========================

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


# =========================
# AUTH CHECK
# =========================

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


# =========================
# SUPABASE
# =========================

def supabase_get(params=None, offset=0):

    if not SUPABASE_URL or not SUPABASE_KEY:

        raise HTTPException(
            status_code=500,
            detail="Cloud database is not configured"
        )

    url = f"{SUPABASE_URL}/rest/v1/{TABLE}"

    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
    }

    if params is None:
        params = {}

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


# =========================
# HOME
# =========================

@app.get("/")
def home():

    return {
        "status": "online",
        "app": "Rahul Software API"
    }


# =========================
# SHEETS
# =========================

@app.get("/sheets")
def sheets(username: str = Depends(verify_token)):

    names = set()

    offset = 0

    while True:

        rows = supabase_get(
            {
                "select": "sheet"
            },
            offset
        )

        if not rows:
            break

        for item in rows:

            sheet = item.get("sheet")

            if sheet:
                names.add(str(sheet))

        if len(rows) < 1000:
            break

        offset += 1000

    return {
        "count": len(names),
        "sheets": sorted(names)
    }


# =========================
# GET SHEET
# =========================

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


# =========================
# SEARCH
# =========================

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
