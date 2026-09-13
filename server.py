from fastapi import FastAPI,HTTPException,Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer,HTTPAuthorizationCredentials
from datetime import datetime,date
import os,hmac,hashlib,base64,json,requests,calendar

app=FastAPI(title="Rahul Software API")
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_credentials=True,allow_methods=["*"],allow_headers=["*"])
security=HTTPBearer(auto_error=False)

USERNAME=os.getenv("RAHUL_USERNAME","rahul")
PASSWORD=os.getenv("RAHUL_PASSWORD","rahul123")
SECRET=os.getenv("RAHUL_SECRET","change-this-secret")
SUPABASE_URL=os.getenv("SUPABASE_URL")
SUPABASE_KEY=os.getenv("SUPABASE_KEY")
TABLE="excel_rows"

MONTHLY_MATERIALS=[
"SLIPSHEET","FRESCO PAD","M FOLD","CORE PIPE SCRAP","PAPER SCRAP",
"TOILET ROLL","KRAFT PAPER","PET GRIPSHEET","TISSUE PAPER/   NAPKIN",
"KITCHEN ROLL","PLASTIC SHEET","JRT","Z FOLD"
]

MATERIAL_MAPPING={
"SLIPSHEET":"SLIPSHEET","SLIP SHEET":"SLIPSHEET",
"FRESCOPAD":"FRESCO PAD","FRESCO PAD":"FRESCO PAD",
"MFOLD":"M FOLD","M FOLD":"M FOLD",
"COREPIPESCRAP":"CORE PIPE SCRAP","CORE PIPE SCRAP":"CORE PIPE SCRAP",
"PAPERSCRAP":"PAPER SCRAP","PAPER SCRAP":"PAPER SCRAP",
"TOILETROLL":"TOILET ROLL","TOILET ROLL":"TOILET ROLL",
"KRAFTPAPER":"KRAFT PAPER","KRAFT PAPER":"KRAFT PAPER",
"PETGRIPSHEET":"PET GRIPSHEET","PET GRIP SHEET":"PET GRIPSHEET","PETGRIP SHEET":"PET GRIPSHEET",
"NAPKIN":"TISSUE PAPER/   NAPKIN","TISSUEPAPER":"TISSUE PAPER/   NAPKIN",
"TISSUE PAPER":"TISSUE PAPER/   NAPKIN","TISSUEPAPERNAPKIN":"TISSUE PAPER/   NAPKIN",
"TISSUE PAPER NAPKIN":"TISSUE PAPER/   NAPKIN",
"KITCHENROLL":"KITCHEN ROLL","KITCHEN ROLL":"KITCHEN ROLL",
"PLASTICSHEET":"PLASTIC SHEET","PLASTIC SHEET":"PLASTIC SHEET",
"JRT":"JRT","ZFOLD":"Z FOLD","Z FOLD":"Z FOLD"
}

def text(v):
    return "" if v is None else str(v).strip()

def number(v):
    try:
        return 0 if v in (None,"") else float(str(v).replace(",","").strip())
    except:
        return 0

def clean_number(v):
    v=float(v)
    return int(v) if v.is_integer() else round(v,2)

def date_only(v):
    if v is None:return None
    if isinstance(v,datetime):return v.date()
    if isinstance(v,date):return v
    s=str(v).strip()
    try:return datetime.fromisoformat(s.replace("Z","+00:00")).date()
    except:pass
    for f in ("%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%m/%d/%Y","%d.%m.%Y"):
        try:return datetime.strptime(s,f).date()
        except:pass
    return None

def first_existing(r,names):
    for n in names:
        if n in r:return r[n]
    return None

def pos(row,n):
    return row[n-1] if 0<n<=len(row) else None

def normalize_material(v):
    s=text(v).upper()
    s=s.replace("\u00a0"," ").replace("\t"," ").replace("-"," ").replace("_"," ")
    s=s.replace("/"," ").replace("\\"," ").replace("("," ").replace(")"," ")
    remove={"PCS","PC","KG","KGS","NOS","NO","KILOGRAM","KILOGRAMS"}
    return " ".join(x for x in s.split() if x not in remove)

def normalize_size(v):
    return " ".join(text(v).split())

def make_token(user):
    payload={"username":user,"exp":int(datetime.now().timestamp())+86400}
    raw=json.dumps(payload,separators=(",",":")).encode()
    enc=base64.urlsafe_b64encode(raw).decode().rstrip("=")
    sig=hmac.new(SECRET.encode(),enc.encode(),hashlib.sha256).hexdigest()
    return enc+"."+sig

def verify_token(t):
    try:
        enc,sig=t.split(".",1)
        expected=hmac.new(SECRET.encode(),enc.encode(),hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig,expected):return False
        raw=base64.urlsafe_b64decode(enc+"="*(-len(enc)%4))
        return json.loads(raw.decode())["exp"]>=int(datetime.now().timestamp())
    except:
        return False

def require_auth(c:HTTPAuthorizationCredentials=Depends(security)):
    if c is None or c.scheme.lower()!="bearer" or not verify_token(c.credentials):
        raise HTTPException(401,"Authentication required")

def supabase_headers():
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise HTTPException(500,"Supabase environment variables missing")
    return {
        "apikey":SUPABASE_KEY,
        "Authorization":f"Bearer {SUPABASE_KEY}",
        "Content-Type":"application/json"
    }

def get_all_rows():
    url=f"{SUPABASE_URL}/rest/v1/{TABLE}"
    result=[]
    offset=0
    while True:
        r=requests.get(
            url,
            headers=supabase_headers(),
            params={"select":"*","offset":offset,"limit":1000},
            timeout=60
        )
        if not r.ok:
            raise HTTPException(500,"Supabase read error: "+r.text)
        rows=r.json()
        if not rows:break
        result.extend(rows)
        if len(rows)<1000:break
        offset+=1000
    return result

def convert_database_row(item):
    headers=item.get("headers",[])
    row=item.get("row",[])
    if not isinstance(headers,list):headers=[]
    if not isinstance(row,list):row=[]
    return {
        "sheet":item.get("sheet",""),
        "headers":headers,
        "row":row,
        "data":{str(headers[i]):row[i] if i<len(row) else None for i in range(len(headers))}
    }

@app.get("/")
def home():
    return {"status":"online","message":"Rahul Software API is running","database":"Supabase"}

@app.post("/login")
def login(data:dict):
    username=text(data.get("username"))
    password=text(data.get("password"))
    if hmac.compare_digest(username,USERNAME) and hmac.compare_digest(password,PASSWORD):
        t=make_token(username)
        return {"status":"success","access_token":t,"token":t,"token_type":"bearer"}
    raise HTTPException(401,"Invalid username or password")

@app.get("/sheets")
def sheets(authenticated:bool=Depends(require_auth)):
    names={text(i.get("sheet")) for i in get_all_rows() if text(i.get("sheet"))}
    return {"count":len(names),"sheets":sorted(names)}

@app.get("/sheet/{sheet_name}")
def get_sheet(sheet_name:str,limit:int=100,authenticated:bool=Depends(require_auth)):
    result=[];headers=[]
    for item in get_all_rows():
        if text(item.get("sheet"))==sheet_name:
            c=convert_database_row(item)
            if not headers:headers=c["headers"]
            result.append(c["data"])
    if not result and not headers:
        raise HTTPException(404,f"Sheet '{sheet_name}' not found")
    limit=max(1,min(int(limit),1000))
    return {"sheet":sheet_name,"headers":headers,"count":min(len(result),limit),"data":result[:limit]}

@app.get("/search")
def search(q:str,sheet:str="ALL SHEETS",authenticated:bool=Depends(require_auth)):
    q=text(q).lower()
    if not q:return {"query":q,"sheet":sheet,"count":0,"results":[]}
    result=[]
    for item in get_all_rows():
        sn=text(item.get("sheet"))
        if sheet!="ALL SHEETS" and sn!=sheet:continue
        c=convert_database_row(item)
        if any(q in text(v).lower() for v in c["data"].values()):
            result.append({
                "sheet":sn,
                "headers":c["headers"],
                "row":c["row"],
                "data":c["data"]
            })
    return {"query":q,"sheet":sheet,"count":len(result),"results":result}

@app.get("/dashboard")
def dashboard(authenticated:bool=Depends(require_auth)):
    rows=get_all_rows()
    today=datetime.now().date()
    orders=[]
    dispatch=[]
    hold=[]

    for item in rows:
        sn=text(item.get("sheet"))
        c=convert_database_row(item)
        if sn=="ORDERS":
            orders.append(c)
        elif sn=="Dispatched Orders":
            dispatch.append({"c":c,"cancelled":item.get("cancelled",False) is True})
        elif sn=="Holded Orders":
            hold.append(c["data"])

    order_qty=0
    today_orders=0
    today_order_qty=0
    order_numbers=set()
    size_data={}
    party_data={}

    for c in orders:
        row=c["row"]
        record=c["data"]
        order_no=pos(row,1)
        party=text(pos(row,3))
        size=normalize_size(pos(row,4))
        qty=number(pos(row,6))
        row_date=date_only(first_existing(record,["Date","DATE"]))
        destination=text(pos(row,9))

        order_qty+=qty

        if order_no not in (None,""):
            order_numbers.add(text(order_no))

        if row_date==today:
            today_orders+=1
            today_order_qty+=qty

        if size:
            x=size_data.setdefault(size,{"orders":0,"quantity":0,"nos":set()})
            x["orders"]+=1
            x["quantity"]+=qty
            if order_no not in (None,""):
                x["nos"].add(text(order_no))

        if party:
            x=party_data.setdefault(
                party,{"orders":0,"quantity":0,"destinations":{}}
            )
            x["orders"]+=1
            x["quantity"]+=qty

            destination_name=destination or "NO DESTINATION"
            y=x["destinations"].setdefault(
                destination_name,{"orders":0,"quantity":0,"sizes":{}}
            )
            y["orders"]+=1
            y["quantity"]+=qty

            size_name=size or "UNKNOWN"
            z=y["sizes"].setdefault(
                size_name,{"orders":0,"quantity":0}
            )
            z["orders"]+=1
            z["quantity"]+=qty

    dispatch_qty=0
    today_dispatch=0
    today_dispatch_qty=0
    cancelled=0

    for item in dispatch:
        if item["cancelled"]:
            cancelled+=1
            continue

        c=item["c"]
        qty=number(first_existing(c["data"],["Quantity Pcs.","Quantity","QUANTITY"]))
        row_date=date_only(first_existing(c["data"],["Date","BILL DATE","IN/OUT DATE"]))

        dispatch_qty+=qty

        if row_date==today:
            today_dispatch+=1
            today_dispatch_qty+=qty

    hold_qty=sum(
        number(first_existing(x,["Quantity Pcs.","Quantity","QUANTITY"]))
        for x in hold
    )

    size_wise=[]
    for k,v in size_data.items():
        size_wise.append({
            "size":k,
            "orders":len(v["nos"]) or v["orders"],
            "quantity":clean_number(v["quantity"])
        })
    size_wise.sort(key=lambda x:x["quantity"],reverse=True)

    party_wise=[]
    for p,x in party_data.items():
        destinations=[]

        for dn,y in x["destinations"].items():
            sizes=[]
            for sn,z in y["sizes"].items():
                sizes.append({
                    "size":sn,
                    "orders":z["orders"],
                    "quantity":clean_number(z["quantity"])
                })

            sizes.sort(key=lambda x:x["quantity"],reverse=True)

            destinations.append({
                "destination":dn,
                "orders":y["orders"],
                "quantity":clean_number(y["quantity"]),
                "sizes":sizes
            })

        destinations.sort(key=lambda x:x["quantity"],reverse=True)

        party_wise.append({
            "party":p,
            "orders":x["orders"],
            "quantity":clean_number(x["quantity"]),
            "destinations":destinations
        })

    party_wise.sort(key=lambda x:x["quantity"],reverse=True)

    return {
        "date":today.isoformat(),
        "today":{
            "orders":today_orders,
            "order_quantity":clean_number(today_order_qty),
            "dispatch":today_dispatch,
            "dispatch_quantity":clean_number(today_dispatch_qty)
        },
        "orders":{
            "total_rows":len(orders),
            "unique_orders":len(order_numbers),
            "total_quantity":clean_number(order_qty)
        },
        "dispatch":{
            "total_rows":len(dispatch)-cancelled,
            "total_quantity":clean_number(dispatch_qty),
            "today_rows":today_dispatch,
            "today_quantity":clean_number(today_dispatch_qty),
            "cancelled_rows":cancelled
        },
        "hold":{
            "orders":len(hold),
            "quantity":clean_number(hold_qty)
        },
        "size_wise":size_wise,
        "party_wise":party_wise,
        "last_updated":datetime.now().astimezone().isoformat()
    }

@app.get("/monthly-report/months")
def monthly_report_months(authenticated:bool=Depends(require_auth)):
    months=set()

    for item in get_all_rows():
        if text(item.get("sheet"))!="Dispatched Orders":
            continue
        if item.get("cancelled",False) is True:
            continue

        row=convert_database_row(item)["row"]
        d=date_only(pos(row,11))

        if d:
            months.add((d.year,d.month))

    return {
        "count":len(months),
        "months":[
            {
                "year":y,
                "month":m,
                "month_name":calendar.month_name[m],
                "label":f"{calendar.month_name[m]} {y}"
            }
            for y,m in sorted(months,reverse=True)
        ]
    }

@app.get("/monthly-report/daily-pcs-dispatch")
def daily_pcs_dispatch(
    year:int,
    month:int,
    authenticated:bool=Depends(require_auth)
):
    if month<1 or month>12:
        raise HTTPException(400,"Invalid month")
    if year<2000 or year>2100:
        raise HTTPException(400,"Invalid year")

    days={
        d:{
            "date":d,
            **{m:0 for m in MONTHLY_MATERIALS}
        }
        for d in range(1,calendar.monthrange(year,month)[1]+1)
    }

    matched=0
    cancelled=0
    unknown=0
    no_date=0

    for item in get_all_rows():
        if text(item.get("sheet"))!="Dispatched Orders":
            continue

        if item.get("cancelled",False) is True:
            cancelled+=1
            continue

        row=convert_database_row(item)["row"]
        row_date=date_only(pos(row,11))

        if row_date is None:
            no_date+=1
            continue

        if row_date.year!=year or row_date.month!=month:
            continue

        material=normalize_material(pos(row,13))
        report_column=MATERIAL_MAPPING.get(material)

        if not report_column:
            unknown+=1
            continue

        days[row_date.day][report_column]+=number(pos(row,6))
        matched+=1

    total={
        "date":"TOTAL",
        **{
            m:clean_number(sum(days[d][m] for d in days))
            for m in MONTHLY_MATERIALS
        }
    }

    return {
        "year":year,
        "month":month,
        "month_name":calendar.month_name[month],
        "title":f"{calendar.month_name[month].upper()}- {year} MONTHLY PCS DISPATCH QUANTITY",
        "columns":MONTHLY_MATERIALS,
        "days":[
            {
                k:clean_number(v) if k!="date" else v
                for k,v in x.items()
            }
            for x in days.values()
        ],
        "total":total,
        "matched_dispatch_rows":matched,
        "skipped_cancelled_rows":cancelled,
        "skipped_unknown_material":unknown,
        "skipped_no_date":no_date,
        "source_columns":{
            "date":"K",
            "material":"M",
            "quantity":"F"
        },
        "last_updated":datetime.now().astimezone().isoformat()
    }

@app.get("/monthly-report/daily-pcs-dispatch/current")
def current_month_daily_pcs_dispatch(authenticated:bool=Depends(require_auth)):
    d=datetime.now().date()
    return daily_pcs_dispatch(d.year,d.month,True)

@app.get("/monthly-report/material-mapping")
def material_mapping(authenticated:bool=Depends(require_auth)):
    result={}

    for item in get_all_rows():
        if text(item.get("sheet"))!="Dispatched Orders":
            continue

        value=text(pos(convert_database_row(item)["row"],13))

        if value:
            normalized=normalize_material(value)
            mapped=MATERIAL_MAPPING.get(normalized)
            result[value]={
                "excel_value":value,
                "normalized":normalized,
                "report_column":mapped,
                "mapped":bool(mapped)
            }

    return {
        "count":len(result),
        "materials":list(result.values())
    }

TEST_REPORT_MASTER_SHEETS=[
    "TEST REPORT DATA",
    "TEST REPORT PARTY MASTER",
    "PARTY MASTER",
    "PARTY MASTER DATA"
]

def find_test_report_master(rows):
    for item in rows:
        headers=item.get("headers",[])
        if not isinstance(headers,list):
            continue

        normalized=[
            normalize_material(x).replace(" ","")
            for x in headers
        ]

        party_index=None
        address_index=None

        for i,name in enumerate(normalized):
            if name in {"PARTYNAME","CUSTOMERNAME","PARTY"} and party_index is None:
                party_index=i
            if name in {"ADDRESS","PARTYADDRESS","CUSTOMERADDRESS"} and address_index is None:
                address_index=i

        if party_index is not None and address_index is not None:
            return party_index,address_index,text(item.get("sheet"))

    for sheet_name in TEST_REPORT_MASTER_SHEETS:
        if any(text(i.get("sheet"))==sheet_name for i in rows):
            return 0,1,sheet_name

    return None,None,None

@app.get("/test-report/master-data")
def test_report_master_data(authenticated:bool=Depends(require_auth)):
    rows=get_all_rows()
    party_index,address_index,sheet_name=find_test_report_master(rows)

    if party_index is None:
        return {
            "sheet":None,
            "count":0,
            "parties":[],
            "message":"A=PARTY NAME aur B=ADDRESS wali master sheet nahi mili."
        }

    parties={}

    for item in rows:
        if text(item.get("sheet"))!=sheet_name:
            continue

        row=item.get("row",[])
        if not isinstance(row,list):
            continue

        party=text(pos(row,party_index+1))
        address=text(pos(row,address_index+1))

        if not party:
            continue

        check=normalize_material(party).replace(" ","")

        if check in {"PARTYNAME","CUSTOMERNAME","PARTY"}:
            continue

        parties[party]=address

    result=[
        {"party":party,"address":address}
        for party,address in sorted(
            parties.items(),
            key=lambda x:x[0].upper()
        )
    ]

    return {
        "sheet":sheet_name,
        "count":len(result),
        "parties":result
    }

@app.post("/test-report/reserve-bill")
def test_report_reserve_bill(authenticated:bool=Depends(require_auth)):
    try:
        prefix="MP/26-27/"
        highest=408

        for item in get_all_rows():
            if text(item.get("sheet"))!="Dispatched Orders":
                continue

            row=convert_database_row(item)["row"]

            for value in row:
                value=text(value)

                if value.startswith(prefix):
                    try:
                        highest=max(
                            highest,
                            int(value[len(prefix):])
                        )
                    except:
                        pass

        return {
            "bill_no":prefix+str(highest+1)
        }

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500,str(e))
