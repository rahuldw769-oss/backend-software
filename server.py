from fastapi import FastAPI, HTTPException
import os
import requests

app = FastAPI(title="Rahul Software API")

SUPABASE_URL = os.getenv("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

TABLE = "excel_rows"


def supabase_get(params=None):
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

    response = requests.get(
        url,
        headers=headers,
        params=params or {},
        timeout=30
    )

    if not response.ok:
        raise HTTPException(
            status_code=response.status_code,
            detail=response.text
        )

    return response.json()


@app.get("/")
def home():
    return {
        "status": "online",
        "app": "Rahul Software API"
    }


@app.get("/sheets")
def sheets():
    rows = supabase_get({
        "select": "sheet",
        "limit": 10000
    })

    names = sorted(
        set(
            str(x["sheet"])
            for x in rows
            if x.get("sheet")
        )
    )

    return {
        "count": len(names),
        "sheets": names
    }


@app.get("/sheet/{sheet_name}")
def get_sheet(sheet_name: str, limit: int = 100):
    rows = supabase_get({
        "sheet": f"eq.{sheet_name}",
        "select": "headers,row",
        "limit": min(limit, 1000)
    })

    return {
        "sheet": sheet_name,
        "count": len(rows),
        "data": rows
    }


@app.get("/search")
def search(q: str, sheet: str = "ALL SHEETS", limit: int = 100):
    q = q.strip().lower()

    if not q:
        return {
            "count": 0,
            "results": []
        }

    params = {
        "select": "sheet,headers,row",
        "limit": 10000
    }

    if sheet != "ALL SHEETS":
        params["sheet"] = f"eq.{sheet}"

    rows = supabase_get(params)

    results = []

    for item in rows:
        row = item.get("row", [])

        found = False

        for cell in row:
            if cell is not None and q in str(cell).lower():
                found = True
                break

        if found:
            results.append(item)

            if len(results) >= limit:
                break

    return {
        "count": len(results),
        "results": results
    }