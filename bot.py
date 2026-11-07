import json
import os
import re
import smtplib
import sqlite3
import ssl
from datetime import datetime
from email.message import EmailMessage
import numpy as np
import pandas as pd
import reportlab

REPORT_DIR = "reports"                 
SUPPLIER_FILE = "Suppliers.xlsx"      
ALLOWED_RECIPIENTS = []                
SMTP_HOST, SMTP_PORT = "smtp.gmail.com", 465

FILE = "Inventory Data.xlsx"        

orders = pd.read_excel(FILE, "Past Orders")
stock = pd.read_excel(FILE, "Stock")
orders["SKU ID"] = orders["SKU ID"].astype(str).str.upper().str.strip()
stock["SKU ID"] = stock["SKU ID"].astype(str).str.upper().str.strip()


end = orders["Order Date"].max()
recent = orders[orders["Order Date"] > end - pd.Timedelta(days=90)]
daily = (recent.groupby("SKU ID")["Order Quantity"].sum() / 90).rename("daily")

t = stock.merge(daily, on="SKU ID", how="left")
t["daily"] = t["daily"].fillna(0.0)

t = t.rename(columns={
    "Current Stock Quantity": "stock",
    "Average Lead Time (days)": "lead_avg",
    "Maximum Lead Time (days)": "lead_max",
    "Unit Price": "price",
})


t["value"] = t["stock"] * t["price"]
t["reorder_point"] = t["daily"] * t["lead_max"]
t["cover"] = np.where(t["daily"] > 0, t["stock"] / t["daily"].replace(0, np.nan), np.inf)

def status(r):
    if r.daily == 0:
        return "DEAD" if r.stock > 0 else "EMPTY-NO DEMAND"
    if r.stock <= 0:
        return "STOCKOUT"
    if r.cover < r.lead_avg:
        return "CRITICAL"
    if r.cover < r.lead_max:
        return "WARNING"
    if r.cover > 365:
        return "OVERSTOCK"
    return "OK"

t["status"] = t.apply(status, axis=1)
t["order_qty"] = (t["daily"] * (t["lead_max"] + 30) - t["stock"]).clip(lower=0)


v = t.sort_values("value", ascending=False)
cum = v["value"].cumsum() / v["value"].sum()
abc = dict(zip(v["SKU ID"], np.where(cum <= 0.8, "A", np.where(cum <= 0.95, "B", "C"))))
t["abc"] = t["SKU ID"].map(abc)

HAS_SUPPLIERS = False
if os.path.exists(SUPPLIER_FILE):
    sup = pd.read_excel(SUPPLIER_FILE)
    sup["SKU ID"] = sup["SKU ID"].astype(str).str.upper().str.strip()
    sup = sup.dropna(subset=["Supplier"]).drop_duplicates("SKU ID")[["SKU ID", "Supplier"]]
    if len(sup):
        t = t.merge(sup.rename(columns={"Supplier": "supplier"}), on="SKU ID", how="left")
        HAS_SUPPLIERS = True


def save_count(sku, system_qty, counted_qty):
    var = counted_qty - system_qty
    con = sqlite3.connect("counts.db")
    con.execute("""CREATE TABLE IF NOT EXISTS counts(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT, system_qty REAL, counted_qty REAL, variance REAL,
        counted_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    con.execute("INSERT INTO counts(sku, system_qty, counted_qty, variance) VALUES (?,?,?,?)",
                (sku, system_qty, counted_qty, var))
    con.commit()
    con.close()
    return var



def restock_list(limit=10):
    bad = t[t["status"].isin(["STOCKOUT", "CRITICAL", "WARNING"])].copy()
    order = {"STOCKOUT": 0, "CRITICAL": 1, "WARNING": 2}
    bad["rank"] = bad["status"].map(order)
    bad["gap"] = bad["lead_avg"] - bad["cover"]         
    bad = bad.sort_values(["rank", "gap"], ascending=[True, False])
    return bad.head(limit)



NOT_UPDATED = "Not updated"          

SKU_RE = re.compile(r"\b(\d{4})\s?([A-Za-z]{2})\b")

def find_sku(msg):
    m = SKU_RE.search(msg)
    return (m.group(1) + m.group(2)).upper() if m else None

def has(low, words):
    return any(w in low for w in words)

def get_n(low, default=5):
    m = re.search(r"\b(\d+)\b", low)
    return int(m.group(1)) if m else default

def table_lines(df, make_line):
    return "\n".join(make_line(r) for _, r in df.iterrows())


# ---------- STEP 8: questions about ONE SKU ----------
# (keywords, label, function that reads the answer from the table)
FIELDS = [
    (["price", "cost", "rate"],              "Unit price",      lambda r, o: f"{r['price']:,.2f}"),
    (["lead time", "lead", "delivery"],      "Lead time",       lambda r, o: f"{r['lead_avg']} days average, {r['lead_max']} days maximum"),
    (["demand", "sales", "sold", "selling"], "Demand",          lambda r, o: f"{r['daily']:.1f} Kg/day (last 90 days)"),
    (["cover", "how long"],                  "Days of cover",   lambda r, o: "no recent demand" if r["daily"] == 0 else f"{r['cover']:.0f} days"),
    (["status"],                             "Status",          lambda r, o: r["status"]),
    (["class", "abc"],                       "ABC class",       lambda r, o: r["abc"]),
    (["value", "worth"],                     "Stock value",     lambda r, o: f"{r['value']:,.0f}"),
    (["last order", "latest order"],         "Last order date", lambda r, o: str(o["Order Date"].max().date()) if len(o) else NOT_UPDATED),
    (["how many orders", "number of orders", "total orders"], "Number of orders", lambda r, o: f"{len(o):,}"),
    (["total ordered"],                      "Total ordered",   lambda r, o: f"{o['Order Quantity'].sum():,.0f} Kg"),
    (["stock", "quantity", "qty", "available", "have", "left", "balance"], "Stock", lambda r, o: f"{r['stock']:,.0f} Kg"),
]

# things people may ask about that are NOT in the Excel file
NOT_IN_DATABASE = ["supplier", "vendor", "location", "rack", "shelf", "bin", "expiry", "expire",
                   "batch", "brand", "category", "barcode", "warehouse", "weight", "color",
                   "description", "name of", "manufacturer", "country"]

def answer_sku(sku, low):
    row = t[t["SKU ID"] == sku]
    if row.empty:
        return NOT_UPDATED                     
    if has(low, NOT_IN_DATABASE):
        return NOT_UPDATED                     
    r = row.iloc[0]
    o = orders[orders["SKU ID"] == sku]         
    out = []
    for words, label, fn in FIELDS:
        if has(low, words):
            out.append(f"{label} of {sku}: {fn(r, o)}")
    if out:
        return "\n".join(out)
    return (f"{sku}: stock {r['stock']:,.0f} Kg | price {r['price']:,.2f} | demand {r['daily']:.1f} Kg/day | "
            f"lead time {r['lead_avg']}/{r['lead_max']} days | status {r['status']} | class {r['abc']}")



def handle(msg):
    low = msg.lower().strip()
    sku = find_sku(msg)

    if low in {"hi", "hello", "help", "hey", ""}:
        return "Ask me about stock, price, lead time, demand, orders, reorder, overstock... Anything else: 'Not updated'."

    if sku and has(low, ["counted", "found"]):
        rest = SKU_RE.sub(" ", msg)
        numbers = re.findall(r"\d+\.?\d*", rest)
        row = t[t["SKU ID"] == sku]
        if row.empty:
            return NOT_UPDATED
        if not numbers:
            return "What quantity did you count?"
        system_qty = float(row.iloc[0]["stock"])
        counted = float(numbers[-1])
        var = save_count(sku, system_qty, counted)
        return f"{sku}: system {system_qty:,.0f}, counted {counted:,.0f}, variance {var:+,.0f} Kg"


    if sku:
        return answer_sku(sku, low)

 
    n = get_n(low)
    if has(low, ["how many sku", "number of sku", "total sku", "how many items", "how many products"]):
        return f"{len(t)} SKUs in the database."
    if has(low, ["total stock value", "total value", "inventory value", "stock worth"]):
        return f"Total stock value: {t['value'].sum():,.0f}"
    if has(low, ["fast", "top sell", "best sell", "highest demand", "most sold", "most ordered"]):
        d = t.sort_values("daily", ascending=False).head(n)
        return f"Top {len(d)} by demand:\n" + table_lines(d, lambda r: f"{r['SKU ID']}  {r['daily']:,.1f} Kg/day")
    if has(low, ["highest value", "most valuable", "top value", "biggest value"]):
        d = t.sort_values("value", ascending=False).head(n)
        return f"Top {len(d)} by stock value:\n" + table_lines(d, lambda r: f"{r['SKU ID']}  {r['value']:,.0f}")
    if has(low, ["zero stock", "out of stock", "no stock", "stockout", "stock out"]):
        d = t[t["stock"] <= 0]
        return f"{len(d)} SKUs with zero stock:\n" + ", ".join(d["SKU ID"].head(30))
    if has(low, ["overstock", "too much", "excess"]):
        d = t[t["status"] == "OVERSTOCK"].sort_values("value", ascending=False)
        return (f"{len(d)} overstock SKUs. Top {min(n, len(d))}:\n"
                + table_lines(d.head(n), lambda r: f"{r['SKU ID']}  cover {r['cover']:.0f} days  value {r['value']:,.0f}"))
    if has(low, ["dead", "no demand", "not moving"]):
        d = t[t["status"] == "DEAD"]
        return f"{len(d)} dead-stock SKUs (stock but no orders in 90 days):\n" + ", ".join(d["SKU ID"])
    if has(low, ["low", "reorder", "restock", "running out", "short", "critical"]):
        d = restock_list(n)
        total = t["status"].isin(["STOCKOUT", "CRITICAL", "WARNING"]).sum()
        return (f"{total} SKUs need attention. Top {len(d)}:\n"
                + table_lines(d, lambda r: f"{r['SKU ID']}  {r['status']:<9} stock {r['stock']:>8,.0f}  cover {r['cover']:>4.0f} d  order {r['order_qty']:>8,.0f}"))
    if has(low, ["summary", "overview"]):
        return t["status"].value_counts().to_string()

    return NOT_UPDATED               




MODEL = "claude-sonnet-5-5"         


COLS = ["SKU ID", "stock", "price", "daily", "cover", "lead_avg", "lead_max",
        "reorder_point", "value", "status", "abc", "order_qty"]
if HAS_SUPPLIERS:
    COLS.append("supplier")

def rows(df):
    """DataFrame -> list of plain dicts (inf/NaN become null)."""
    d = df[COLS].replace([np.inf, -np.inf], np.nan).round(2)
    return json.loads(d.to_json(orient="records"))

def tool_get_sku(sku):
    sku = str(sku).upper().replace(" ", "")
    row = t[t["SKU ID"] == sku]
    if row.empty:
        return {"error": "SKU not found"}
    o = orders[orders["SKU ID"] == sku]
    info = rows(row)[0]
    info["number_of_orders"] = int(len(o))
    info["total_ordered_kg"] = float(o["Order Quantity"].sum())
    info["first_order_date"] = str(o["Order Date"].min().date()) if len(o) else None
    info["last_order_date"] = str(o["Order Date"].max().date()) if len(o) else None
    return info

def tool_query_skus(status=None, abc=None, sort_by="daily", descending=True, limit=10):
    d = t
    if status:
        d = d[d["status"] == status.upper()]
    if abc:
        d = d[d["abc"] == abc.upper()]
    if sort_by not in COLS:
        return {"error": "sort_by must be one of " + ", ".join(COLS)}
    matching = len(d)                         
    d = d.sort_values(sort_by, ascending=not descending).head(min(int(limit), 50))
    return {"count_matching": int(matching), "rows": rows(d)}

def tool_summary():
    return {
        "number_of_skus": int(len(t)),
        "total_stock_value": round(float(t["value"].sum()), 2),
        "status_counts": {str(k): int(v) for k, v in t["status"].value_counts().items()},
        "abc_counts": {str(k): int(v) for k, v in t["abc"].value_counts().items()},
        "orders_data_from": str(orders["Order Date"].min().date()),
        "orders_data_to": str(orders["Order Date"].max().date()),
    }

def tool_monthly_demand(sku, months=6):
    sku = str(sku).upper().replace(" ", "")
    o = orders[orders["SKU ID"] == sku]
    if o.empty:
        return {"error": "no orders for this SKU"}
    m = o.groupby(o["Order Date"].dt.to_period("M"))["Order Quantity"].sum().tail(int(months))
    return {str(k): float(v) for k, v in m.items()}

def tool_record_count(sku, counted_qty):
    sku = str(sku).upper().replace(" ", "")
    row = t[t["SKU ID"] == sku]
    if row.empty:
        return {"error": "SKU not found"}
    system_qty = float(row.iloc[0]["stock"])
    var = save_count(sku, system_qty, float(counted_qty))
    return {"sku": sku, "system_qty": system_qty, "counted_qty": float(counted_qty), "variance": float(var)}




def num(x, digits=1):
    """Number -> rounded float, or None for NaN/inf (JSON-safe)."""
    x = float(x)
    return None if (np.isnan(x) or np.isinf(x)) else round(x, digits)



def monthly_series(sku=None):
    """Kg ordered per month (all SKUs if sku is None). Drops a last month that is only partly finished."""
    o = orders if sku is None else orders[orders["SKU ID"] == sku]
    if o.empty:
        return None
    m = o.groupby(o["Order Date"].dt.to_period("M"))["Order Quantity"].sum()
    last_day = orders["Order Date"].max()
    last_month = last_day.to_period("M")
    if last_day.normalize() < last_month.end_time.normalize():     
        m = m[m.index < last_month]
    if m.empty:
        return None
    return m.reindex(pd.period_range(m.index.min(), m.index.max(), freq="M"), fill_value=0)

def tool_forecast_demand(sku=None, months_ahead=1):
    sku = str(sku).upper().replace(" ", "") if sku else None
    if sku and t[t["SKU ID"] == sku].empty:
        return {"error": "SKU not found"}
    m = monthly_series(sku)
    if m is None or len(m) < 3:
        return {"error": "not enough monthly history to forecast"}
    ma3 = float(m.tail(3).mean())                     
    level = float(m.iloc[0])
    for x in m.iloc[1:]:                              
        level = 0.5 * float(x) + 0.5 * level
    prev3 = m.iloc[-6:-3]
    trend = None
    if len(prev3) == 3 and prev3.mean() > 0:
        trend = round((ma3 / float(prev3.mean()) - 1) * 100, 1)
    out = {
        "scope": sku or "ALL SKUs",
        "last_full_month": str(m.index[-1]),
        "history_kg_per_month": {str(k): float(v) for k, v in m.tail(6).items()},
        "moving_average_3m_kg": round(ma3, 1),
        "forecast_kg_per_month": round(level, 1),
        "forecast_months": int(months_ahead),
        "forecast_total_kg": round(level * int(months_ahead), 1),
        "change_last3_vs_previous3_pct": trend,
        "method": "exponential smoothing (alpha 0.5); flat estimate, no seasonality",
    }
    if sku:
        row = t[t["SKU ID"] == sku].iloc[0]
        out["stock_kg"] = float(row["stock"])
        out["months_of_stock_at_forecast"] = num(row["stock"] / level) if level > 0 else None
        out["average_lead_time_months"] = num(row["lead_avg"] / 30)
    return out



def tool_what_if(lead_time_multiplier=1.0, demand_multiplier=1.0, stock_change_pct=0.0, sku=None, limit=10):
    s = t.copy()
    s["daily"] = s["daily"] * float(demand_multiplier)
    s["lead_avg"] = s["lead_avg"] * float(lead_time_multiplier)
    s["lead_max"] = s["lead_max"] * float(lead_time_multiplier)
    s["stock"] = s["stock"] * (1 + float(stock_change_pct) / 100)
    s["cover"] = s["stock"] / s["daily"]
    s["status"] = s.apply(status, axis=1)                
    s["order_qty"] = (s["daily"] * (s["lead_max"] + 30) - s["stock"]).clip(lower=0)

    before, after = t.set_index("SKU ID"), s.set_index("SKU ID")
    scenario = {"lead_time_multiplier": lead_time_multiplier, "demand_multiplier": demand_multiplier,
                "stock_change_pct": stock_change_pct}
    if sku:
        sku = str(sku).upper().replace(" ", "")
        if sku not in after.index:
            return {"error": "SKU not found"}
        b, a = before.loc[sku], after.loc[sku]
        return {"sku": sku, "scenario": scenario,
                "before": {"status": b["status"], "days_of_cover": num(b["cover"]), "lead_time_avg_days": num(b["lead_avg"]), "suggested_order_kg": num(b["order_qty"], 0)},
                "after": {"status": a["status"], "days_of_cover": num(a["cover"]), "lead_time_avg_days": num(a["lead_avg"]), "suggested_order_kg": num(a["order_qty"], 0)}}

    risk = ["STOCKOUT", "CRITICAL"]
    newly = after[after["status"].isin(risk) & ~before["status"].isin(risk)].sort_values("daily", ascending=False)
    return {
        "scenario": scenario,
        "status_counts_before": {str(k): int(v) for k, v in before["status"].value_counts().items()},
        "status_counts_after": {str(k): int(v) for k, v in after["status"].value_counts().items()},
        "newly_at_risk_count": int(len(newly)),
        "newly_at_risk_top": [{"SKU ID": k, "status_before": before.loc[k, "status"], "status_after": r["status"],
                               "days_of_cover": num(r["cover"]), "suggested_order_kg": num(r["order_qty"], 0)}
                              for k, r in newly.head(min(int(limit), 50)).iterrows()],
        "total_suggested_order_kg_before": num(before["order_qty"].sum(), 0),
        "total_suggested_order_kg_after": num(after["order_qty"].sum(), 0),
    }



def tool_export_report(report="reorder", file_format="excel"):
    cols = ["SKU ID", "status", "abc", "stock", "price", "daily", "cover", "lead_avg", "lead_max", "order_qty"]
    if HAS_SUPPLIERS:
        cols.insert(1, "supplier")
    if report == "reorder":
        d = t[t["status"].isin(["STOCKOUT", "CRITICAL", "WARNING"])].copy()
        d["_rank"] = d["status"].map({"STOCKOUT": 0, "CRITICAL": 1, "WARNING": 2})
        d = d.sort_values(["_rank", "cover"])
        title = "Reorder list"
    elif report == "overstock":
        d = t[t["status"] == "OVERSTOCK"].sort_values("value", ascending=False)
        title = "Overstock list"
    elif report == "full":
        d = t.sort_values("SKU ID")
        title = "Full inventory"
    else:
        return {"error": "report must be reorder, overstock or full"}
    d = d[cols].replace([np.inf, -np.inf], np.nan).round(1)

    os.makedirs(REPORT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if str(file_format).lower() == "pdf":
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        except ImportError:
            return {"error": "PDF needs the reportlab package: pip install reportlab"}
        path = os.path.join(REPORT_DIR, f"{report}_{stamp}.pdf")
        doc = SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
        styles = getSampleStyleSheet()
        data = [list(d.columns)] + d.astype(object).where(d.notna(), "-").values.tolist()
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a5f")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f5f9")]),
        ]))
        doc.build([Paragraph(title + " - " + datetime.now().strftime("%Y-%m-%d"), styles["Title"]),
                   Paragraph("Order data up to " + str(orders["Order Date"].max().date()) + ". Quantities in Kg.", styles["Normal"]),
                   Spacer(1, 8), table])
    else:
        path = os.path.join(REPORT_DIR, f"{report}_{stamp}.xlsx")
        d.to_excel(path, index=False)
    return {"file": os.path.abspath(path), "rows": int(len(d)), "report": report}


def ask_human(prompt):
    """Asks the person at the keyboard. In a web app you would replace this with a button."""
    return input(prompt)

def tool_send_email(to, subject, body, attach_file=None):
    if not ALLOWED_RECIPIENTS:
        return {"error": "Email is switched off. Add the allowed addresses to ALLOWED_RECIPIENTS in bot.py."}
    if str(to).strip().lower() not in [a.lower() for a in ALLOWED_RECIPIENTS]:
        return {"error": "That address is not in the ALLOWED_RECIPIENTS list."}
    path = None
    if attach_file:
        path = os.path.abspath(attach_file)
        if not path.startswith(os.path.abspath(REPORT_DIR) + os.sep) or not os.path.exists(path):
            return {"error": "Only files inside the reports folder can be attached."}
    sender, password = os.environ.get("EMAIL_ADDRESS"), os.environ.get("EMAIL_APP_PASSWORD")
    if not sender or not password:
        return {"error": "Set EMAIL_ADDRESS and EMAIL_APP_PASSWORD first."}

    print("\n--- EMAIL DRAFT ---\nTo: " + to + "\nSubject: " + subject + "\nAttachment: "
          + (os.path.basename(path) if path else "none") + "\n\n" + body + "\n-------------------")
    if ask_human("Send this email? (yes/no): ").strip().lower() not in ("yes", "y"):
        return {"sent": False, "note": "The user declined, nothing was sent."}

    msg = EmailMessage()
    msg["From"], msg["To"], msg["Subject"] = sender, to, subject
    msg.set_content(body)
    if path:
        subtype = "pdf" if path.lower().endswith(".pdf") else "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype=subtype, filename=os.path.basename(path))
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context()) as server:
        server.login(sender, password)
        server.send_message(msg)
    return {"sent": True, "to": to}



def tool_supplier_report(sort_by="avg_lead_time_days", limit=10):
    if not HAS_SUPPLIERS:
        return {"error": "No supplier data. Fill in Suppliers.xlsx (columns: SKU ID, Supplier)."}
    g = (t.dropna(subset=["supplier"]).groupby("supplier")
          .agg(skus=("SKU ID", "count"), avg_lead_time_days=("lead_avg", "mean"),
               avg_max_lead_time_days=("lead_max", "mean"), worst_lead_time_days=("lead_max", "max"),
               stock_value=("value", "sum"),
               critical_skus=("status", lambda s: int(s.isin(["STOCKOUT", "CRITICAL"]).sum())))
          .reset_index())
    if sort_by not in g.columns or sort_by == "supplier":
        return {"error": "sort_by must be one of " + ", ".join(c for c in g.columns if c != "supplier")}
    g = g.sort_values(sort_by, ascending=False).head(min(int(limit), 50)).round(1)
    return {"suppliers": json.loads(g.to_json(orient="records")),
            "note": "Speed is judged from the lead times in the Stock sheet, not from real delivery records."}


TOOL_FUNCS = {
    "get_sku": tool_get_sku,
    "query_skus": tool_query_skus,
    "database_summary": tool_summary,
    "monthly_demand": tool_monthly_demand,
    "record_count": tool_record_count,
    "forecast_demand": tool_forecast_demand,
    "what_if": tool_what_if,
    "export_report": tool_export_report,
    "send_email": tool_send_email,
    "supplier_report": tool_supplier_report,
}


TOOLS = [
    {"name": "get_sku",
     "description": "All facts about one SKU: stock, price, daily demand, days of cover, lead times, status, ABC class, suggested order quantity, order count and dates.",
     "input_schema": {"type": "object", "properties": {"sku": {"type": "string", "description": "e.g. 1193BA"}}, "required": ["sku"]}},
    {"name": "query_skus",
     "description": "List SKUs, filtered and sorted. status is one of STOCKOUT, CRITICAL, WARNING, OK, OVERSTOCK, DEAD. abc is A, B or C. Use for top/bottom lists, reorder lists, overstock lists, fast movers, most valuable SKUs.",
     "input_schema": {"type": "object", "properties": {
         "status": {"type": "string"}, "abc": {"type": "string"},
         "sort_by": {"type": "string", "description": "one of: " + ", ".join(COLS)},
         "descending": {"type": "boolean"}, "limit": {"type": "integer"}}}},
    {"name": "database_summary",
     "description": "Totals: number of SKUs, total stock value, counts by status and ABC class, date range of the order data.",
     "input_schema": {"type": "object", "properties": {}}},
    {"name": "monthly_demand",
     "description": "Total quantity ordered per month for one SKU (use for trends and recent months).",
     "input_schema": {"type": "object", "properties": {"sku": {"type": "string"}, "months": {"type": "integer"}}, "required": ["sku"]}},
    {"name": "record_count",
     "description": "Save a physical stock count for a SKU and return the variance against the system quantity. Only use when the user clearly reports a counted quantity.",
     "input_schema": {"type": "object", "properties": {"sku": {"type": "string"}, "counted_qty": {"type": "number"}}, "required": ["sku", "counted_qty"]}},
    {"name": "forecast_demand",
     "description": "Forecast demand (Kg per month) for one SKU, or for ALL SKUs if sku is omitted. Returns history, 3-month average, smoothed forecast, trend, and months of stock left.",
     "input_schema": {"type": "object", "properties": {"sku": {"type": "string"}, "months_ahead": {"type": "integer"}}}},
    {"name": "what_if",
     "description": "Scenario test. lead_time_multiplier 2 = delivery takes twice as long. demand_multiplier 1.2 = demand up 20%. stock_change_pct -30 = 30% less stock. Give sku for one SKU, or omit it to see the effect on all SKUs.",
     "input_schema": {"type": "object", "properties": {
         "lead_time_multiplier": {"type": "number"}, "demand_multiplier": {"type": "number"},
         "stock_change_pct": {"type": "number"}, "sku": {"type": "string"}, "limit": {"type": "integer"}}}},
    {"name": "export_report",
     "description": "Create a report file in the reports folder. report is reorder, overstock or full. file_format is excel or pdf. Returns the file path.",
     "input_schema": {"type": "object", "properties": {
         "report": {"type": "string"}, "file_format": {"type": "string"}}, "required": ["report"]}},
    {"name": "send_email",
     "description": "Email a message, optionally with a report file made by export_report (pass its file path as attach_file). The person at the keyboard must confirm before it is sent.",
     "input_schema": {"type": "object", "properties": {
         "to": {"type": "string"}, "subject": {"type": "string"}, "body": {"type": "string"},
         "attach_file": {"type": "string"}}, "required": ["to", "subject", "body"]}},
    {"name": "supplier_report",
     "description": "Compare suppliers: number of SKUs, average and worst lead time, stock value, critical SKUs. sort_by is one of avg_lead_time_days, avg_max_lead_time_days, worst_lead_time_days, skus, stock_value, critical_skus. Slowest supplier = highest avg_lead_time_days.",
     "input_schema": {"type": "object", "properties": {"sort_by": {"type": "string"}, "limit": {"type": "integer"}}}},
]

SYSTEM = f"""You are a warehouse inventory assistant.
You may answer ONLY with facts returned by your tools. The database has these fields only:
stock quantity (Kg), unit price, average and maximum lead time (days), order history (date, SKU, quantity),
and values calculated from them (daily demand, days of cover, status, ABC class, suggested order).
It has NO location, expiry, batch, brand or category information. Supplier names exist only if the supplier_report tool returns them.
If the question cannot be answered from the tool results (unrelated topics, missing information,
an unknown SKU), reply with exactly: {NOT_UPDATED}
Forecasts and what-if results are estimates: say so in one short phrase.
When you create a report, tell the user the file path. For an email, call send_email once; the program asks the user to confirm.
Never say an email was sent unless the tool result says sent is true.
Never guess or use outside knowledge. Keep answers short and clear. Quantities are in Kg."""



def ask_llm(question, history):
    import anthropic
    client = anthropic.Anthropic()              
    messages = history + [{"role": "user", "content": question}]
    for _ in range(6):                         
        resp = client.messages.create(model=MODEL, max_tokens=1500, system=SYSTEM,
                                      tools=TOOLS, messages=messages)
        if resp.stop_reason != "tool_use":
            return "".join(b.text for b in resp.content if b.type == "text").strip()
        messages.append({"role": "assistant", "content": resp.content})
        results = []
        for b in resp.content:
            if b.type == "tool_use":
                try:
                    out = TOOL_FUNCS[b.name](**b.input)
                except Exception as e:
                    out = {"error": str(e)}
                results.append({"type": "tool_result", "tool_use_id": b.id,
                                "content": json.dumps(out, default=str)})
        messages.append({"role": "user", "content": results})
    return NOT_UPDATED



chat_history = []                              

def ask(question):
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return handle(question)               
    try:
        answer = ask_llm(question, chat_history[-6:])
    except Exception as e:
        return "AI error: " + str(e)
    chat_history.append({"role": "user", "content": question})
    chat_history.append({"role": "assistant", "content": answer})
    return answer



if __name__ == "__main__":
    mode = "AI mode" if os.environ.get("ANTHROPIC_API_KEY") else "keyword mode (no ANTHROPIC_API_KEY set)"
    print("Inventory bot, " + mode + ". Type 'quit' to stop.")
    while True:
        msg = input("You: ")
        if msg.lower() == "quit":
            break
        print("Bot:", ask(msg))





def num(x, digits=1):
    """Number -> rounded float, or None for NaN/inf (JSON-safe)."""
    x = float(x)
    return None if (np.isnan(x) or np.isinf(x)) else round(x, digits)


# ---------- 1. FORECAST ----------
def monthly_series(sku=None):
    """Kg ordered per month (all SKUs if sku is None). Drops a last month that is only partly finished."""
    o = orders if sku is None else orders[orders["SKU ID"] == sku]
    if o.empty:
        return None
    m = o.groupby(o["Order Date"].dt.to_period("M"))["Order Quantity"].sum()
    last_day = orders["Order Date"].max()
    last_month = last_day.to_period("M")
    if last_day.normalize() < last_month.end_time.normalize():   
        m = m[m.index < last_month]
    if m.empty:
        return None
    return m.reindex(pd.period_range(m.index.min(), m.index.max(), freq="M"), fill_value=0)

def tool_forecast_demand(sku=None, months_ahead=1):
    sku = str(sku).upper().replace(" ", "") if sku else None
    if sku and t[t["SKU ID"] == sku].empty:
        return {"error": "SKU not found"}
    m = monthly_series(sku)
    if m is None or len(m) < 3:
        return {"error": "not enough monthly history to forecast"}
    ma3 = float(m.tail(3).mean())                      
    level = float(m.iloc[0])
    for x in m.iloc[1:]:                               
        level = 0.5 * float(x) + 0.5 * level
    prev3 = m.iloc[-6:-3]
    trend = None
    if len(prev3) == 3 and prev3.mean() > 0:
        trend = round((ma3 / float(prev3.mean()) - 1) * 100, 1)
    out = {
        "scope": sku or "ALL SKUs",
        "last_full_month": str(m.index[-1]),
        "history_kg_per_month": {str(k): float(v) for k, v in m.tail(6).items()},
        "moving_average_3m_kg": round(ma3, 1),
        "forecast_kg_per_month": round(level, 1),
        "forecast_months": int(months_ahead),
        "forecast_total_kg": round(level * int(months_ahead), 1),
        "change_last3_vs_previous3_pct": trend,
        "method": "exponential smoothing (alpha 0.5); flat estimate, no seasonality",
    }
    if sku:
        row = t[t["SKU ID"] == sku].iloc[0]
        out["stock_kg"] = float(row["stock"])
        out["months_of_stock_at_forecast"] = num(row["stock"] / level) if level > 0 else None
        out["average_lead_time_months"] = num(row["lead_avg"] / 30)
    return out



def tool_what_if(lead_time_multiplier=1.0, demand_multiplier=1.0, stock_change_pct=0.0, sku=None, limit=10):
    s = t.copy()
    s["daily"] = s["daily"] * float(demand_multiplier)
    s["lead_avg"] = s["lead_avg"] * float(lead_time_multiplier)
    s["lead_max"] = s["lead_max"] * float(lead_time_multiplier)
    s["stock"] = s["stock"] * (1 + float(stock_change_pct) / 100)
    s["cover"] = s["stock"] / s["daily"]
    s["status"] = s.apply(status, axis=1)                 
    s["order_qty"] = (s["daily"] * (s["lead_max"] + 30) - s["stock"]).clip(lower=0)

    before, after = t.set_index("SKU ID"), s.set_index("SKU ID")
    scenario = {"lead_time_multiplier": lead_time_multiplier, "demand_multiplier": demand_multiplier,
                "stock_change_pct": stock_change_pct}
    if sku:
        sku = str(sku).upper().replace(" ", "")
        if sku not in after.index:
            return {"error": "SKU not found"}
        b, a = before.loc[sku], after.loc[sku]
        return {"sku": sku, "scenario": scenario,
                "before": {"status": b["status"], "days_of_cover": num(b["cover"]), "lead_time_avg_days": num(b["lead_avg"]), "suggested_order_kg": num(b["order_qty"], 0)},
                "after": {"status": a["status"], "days_of_cover": num(a["cover"]), "lead_time_avg_days": num(a["lead_avg"]), "suggested_order_kg": num(a["order_qty"], 0)}}

    risk = ["STOCKOUT", "CRITICAL"]
    newly = after[after["status"].isin(risk) & ~before["status"].isin(risk)].sort_values("daily", ascending=False)
    return {
        "scenario": scenario,
        "status_counts_before": {str(k): int(v) for k, v in before["status"].value_counts().items()},
        "status_counts_after": {str(k): int(v) for k, v in after["status"].value_counts().items()},
        "newly_at_risk_count": int(len(newly)),
        "newly_at_risk_top": [{"SKU ID": k, "status_before": before.loc[k, "status"], "status_after": r["status"],
                               "days_of_cover": num(r["cover"]), "suggested_order_kg": num(r["order_qty"], 0)}
                              for k, r in newly.head(min(int(limit), 50)).iterrows()],
        "total_suggested_order_kg_before": num(before["order_qty"].sum(), 0),
        "total_suggested_order_kg_after": num(after["order_qty"].sum(), 0),
    }


def tool_export_report(report="reorder", file_format="excel"):
    cols = ["SKU ID", "status", "abc", "stock", "price", "daily", "cover", "lead_avg", "lead_max", "order_qty"]
    if HAS_SUPPLIERS:
        cols.insert(1, "supplier")
    if report == "reorder":
        d = t[t["status"].isin(["STOCKOUT", "CRITICAL", "WARNING"])].copy()
        d["_rank"] = d["status"].map({"STOCKOUT": 0, "CRITICAL": 1, "WARNING": 2})
        d = d.sort_values(["_rank", "cover"])
        title = "Reorder list"
    elif report == "overstock":
        d = t[t["status"] == "OVERSTOCK"].sort_values("value", ascending=False)
        title = "Overstock list"
    elif report == "full":
        d = t.sort_values("SKU ID")
        title = "Full inventory"
    else:
        return {"error": "report must be reorder, overstock or full"}
    d = d[cols].replace([np.inf, -np.inf], np.nan).round(1)

    os.makedirs(REPORT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if str(file_format).lower() == "pdf":
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        except ImportError:
            return {"error": "PDF needs the reportlab package: pip install reportlab"}
        path = os.path.join(REPORT_DIR, f"{report}_{stamp}.pdf")
        doc = SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
        styles = getSampleStyleSheet()
        data = [list(d.columns)] + d.astype(object).where(d.notna(), "-").values.tolist()
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a5f")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f5f9")]),
        ]))
        doc.build([Paragraph(title + " - " + datetime.now().strftime("%Y-%m-%d"), styles["Title"]),
                   Paragraph("Order data up to " + str(orders["Order Date"].max().date()) + ". Quantities in Kg.", styles["Normal"]),
                   Spacer(1, 8), table])
    else:
        path = os.path.join(REPORT_DIR, f"{report}_{stamp}.xlsx")
        d.to_excel(path, index=False)
    return {"file": os.path.abspath(path), "rows": int(len(d)), "report": report}



def ask_human(prompt):
    """Asks the person at the keyboard. In a web app you would replace this with a button."""
    return input(prompt)

def tool_send_email(to, subject, body, attach_file=None):
    if not ALLOWED_RECIPIENTS:
        return {"error": "Email is switched off. Add the allowed addresses to ALLOWED_RECIPIENTS in bot.py."}
    if str(to).strip().lower() not in [a.lower() for a in ALLOWED_RECIPIENTS]:
        return {"error": "That address is not in the ALLOWED_RECIPIENTS list."}
    path = None
    if attach_file:
        path = os.path.abspath(attach_file)
        if not path.startswith(os.path.abspath(REPORT_DIR) + os.sep) or not os.path.exists(path):
            return {"error": "Only files inside the reports folder can be attached."}
    sender, password = os.environ.get("EMAIL_ADDRESS"), os.environ.get("EMAIL_APP_PASSWORD")
    if not sender or not password:
        return {"error": "Set EMAIL_ADDRESS and EMAIL_APP_PASSWORD first."}

    print("\n--- EMAIL DRAFT ---\nTo: " + to + "\nSubject: " + subject + "\nAttachment: "
          + (os.path.basename(path) if path else "none") + "\n\n" + body + "\n-------------------")
    if ask_human("Send this email? (yes/no): ").strip().lower() not in ("yes", "y"):
        return {"sent": False, "note": "The user declined, nothing was sent."}

    msg = EmailMessage()
    msg["From"], msg["To"], msg["Subject"] = sender, to, subject
    msg.set_content(body)
    if path:
        subtype = "pdf" if path.lower().endswith(".pdf") else "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype=subtype, filename=os.path.basename(path))
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context()) as server:
        server.login(sender, password)
        server.send_message(msg)
    return {"sent": True, "to": to}



def tool_supplier_report(sort_by="avg_lead_time_days", limit=10):
    if not HAS_SUPPLIERS:
        return {"error": "No supplier data. Fill in Suppliers.xlsx (columns: SKU ID, Supplier)."}
    g = (t.dropna(subset=["supplier"]).groupby("supplier")
          .agg(skus=("SKU ID", "count"), avg_lead_time_days=("lead_avg", "mean"),
               avg_max_lead_time_days=("lead_max", "mean"), worst_lead_time_days=("lead_max", "max"),
               stock_value=("value", "sum"),
               critical_skus=("status", lambda s: int(s.isin(["STOCKOUT", "CRITICAL"]).sum())))
          .reset_index())
    if sort_by not in g.columns or sort_by == "supplier":
        return {"error": "sort_by must be one of " + ", ".join(c for c in g.columns if c != "supplier")}
    g = g.sort_values(sort_by, ascending=False).head(min(int(limit), 50)).round(1)
    return {"suppliers": json.loads(g.to_json(orient="records")),
            "note": "Speed is judged from the lead times in the Stock sheet, not from real delivery records."}

    
