# ============================================================
# main.py — Linen Tracker API
# Handles all HTTP endpoints for towel tracking
# Runs a background job to auto-flag missing towels
# ============================================================

from fastapi import FastAPI, HTTPException
from sqlalchemy.orm import Session
from database import SessionLocal, Towel, Event, init_db
from pydantic import BaseModel
from typing import Optional
import datetime
import asyncio
from contextlib import asynccontextmanager
from fastapi.middleware.cors import CORSMiddleware  # Allows browser dashboards to talk to this API from any webpage
import os  # lets us read environment variables like API_KEY
from fastapi.security.api_key import APIKeyHeader  # handles reading the key from request headers
from fastapi import Security  # used to attach security checks to endpoints
from fastapi import FastAPI, HTTPException, Depends
from fastapi.security import HTTPBasic, HTTPBasicCredentials # for staff login and security purposes- keep track of logs
import secrets
#for excel sheet 
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
#for emails
import smtplib  # Python's built-in email sending library — no pip install needed
from email.mime.text import MIMEText  # formats the email body as HTML
from email.mime.multipart import MIMEMultipart  # allows email to have subject + body

# ============================================================
# HTTP BASIC AUTH — Staff login
# Username and password stored as Railway environment variables
# Set STAFF_USER and STAFF_PASS in Railway Variables tab
# ============================================================
security = HTTPBasic()

def verify_staff(credentials: HTTPBasicCredentials = Depends(security)):
    # secrets.compare_digest is used instead of == for security
    # It prevents "timing attacks" — a way hackers can guess passwords
    correct_user = secrets.compare_digest(
        credentials.username, os.getenv("STAFF_USER", "staff")
    )
    correct_pass = secrets.compare_digest(
        credentials.password, os.getenv("STAFF_PASS", "linen2026")
    )
    if not (correct_user and correct_pass):
        raise HTTPException(
            status_code=401,
            detail="Incorrect username or password",
            headers={"WWW-Authenticate": "Basic"}, # This triggers the browser popup
        )
    return credentials.username  # Returns username so we can log who did what


# ============================================================
# AUTHENTICATION SETUP
# Reads the API_KEY from Railway's environment variables
# If not set, falls back to "dev-key" for local testing only
# ============================================================

# Read the API key from the server's environment variables
API_KEY = os.getenv("API_KEY", "dev-key")

# This tells FastAPI to look for a header called "X-API-Key" in every request
api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)

# This function runs on every protected endpoint
# If the key matches, the request goes through
# If not, it gets blocked with a 403 error
def verify_key(key: str = Security(api_key_header)):
    if key != API_KEY:
        raise HTTPException(
            status_code=403,
            detail="Invalid or missing API key"  # Don't tell them what the right key is
        )
        

# ============================================================
# STARTUP & SHUTDOWN
# lifespan runs once when the server starts, and once when
# it stops. Everything before yield = startup, after = shutdown.
# ============================================================

@asynccontextmanager
async def lifespan(app):
    init_db()                                   # Create DB tables if they don't exist yet
    asyncio.create_task(auto_mark_missing())    # Launch background job (runs forever in background)
    asyncio.create_task(daily_report())         #for sending daily reports to the managers
    print("Server started — background job running")
    yield                                       # Server is now live and handling requests
    print("Server shutting down")


app = FastAPI(lifespan=lifespan)
# ============================================================
# CORS MIDDLEWARE
# Without this, browsers block the dashboard from calling the API
# allow_origins=["*"] means any webpage can connect — fine for now
# In production you'd replace "*" with your exact dashboard URL
# ============================================================
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],     # Allow requests from any webpage
    allow_methods=["*"],     # Allow GET, POST, PATCH and all other methods
    allow_headers=["*"],     # Allow all headers
)


# ============================================================
# BACKGROUND JOB — AUTO MISSING DETECTION
# Wakes up every hour and checks if any towel has been
# in_use for more than 24 hours. If so, marks it "missing"
# and logs a MISSING event.
# ============================================================

async def auto_mark_missing():
    while True:
        await asyncio.sleep(3600)   # Wait 1 hour (3600 seconds), then run the check
                                    # "await" means: pause here but keep serving requests

        db = SessionLocal()
        try:
            # Calculate what time it was 24 hours ago
            cutoff = datetime.datetime.utcnow() - datetime.timedelta(hours=24)

            # Find all towels that are still out AND were dispatched before the cutoff
            missing_towels = db.query(Towel).filter(
                Towel.status == "in_use",
                Towel.dispatched_at < cutoff
            ).all()

            # Loop through each overdue towel and update it
            for towel in missing_towels:
                towel.status = "missing"
                event = Event(
                    tag_id=towel.tag_id,
                    event_type="MISSING",
                    location=towel.last_location,
                    created_at=datetime.datetime.utcnow()
                )
                db.add(event)

            db.commit()
            print(f"Missing check done — {len(missing_towels)} towel(s) flagged")

        finally:
            db.close()  # Always close the DB connection, even if an error occurred
#--------------------------------------------------------------------------------------------------------------------------------------------------------

# ============================================================
# EMAIL ALERTS
# Sends a daily summary email at 8am every morning
# If missing count is 5 or more, email is flagged as urgent
# Reads Gmail credentials from Railway environment variables
# ============================================================

def send_email(subject: str, html_body: str):
    gmail_user    = os.getenv("GMAIL_USER")
    gmail_pass    = os.getenv("GMAIL_PASS")
    manager_email = os.getenv("MANAGER_EMAIL")
    if not all([gmail_user, gmail_pass, manager_email]):
        print("Email not sent — GMAIL_USER, GMAIL_PASS or MANAGER_EMAIL not set in Railway")
        return
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = gmail_user
        msg["To"]      = manager_email
        msg.attach(MIMEText(html_body, "html"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gmail_user, gmail_pass)
            server.sendmail(gmail_user, manager_email, msg.as_string())
        print(f"Email sent to {manager_email}: {subject}")
    except Exception as e:
        print(f"Email failed: {e}")
async def daily_report():
    # ============================================================
    # DAILY REPORT JOB
    # Runs every 24 hours, timed to send at 8am
    # Waits until 8am on first run, then sends every 24 hours
    # ============================================================
    while True:
        # Calculate how many seconds until next 8am
        now     = datetime.datetime.utcnow()  # UTC time — Railway runs on UTC
        next_8am = now.replace(hour=8, minute=0, second=0, microsecond=0)

        # If 8am already passed today, schedule for tomorrow
        if now >= next_8am:
            next_8am += datetime.timedelta(days=1)

        # Sleep until 8am
        seconds_until_8am = (next_8am - now).total_seconds()
        print(f"Daily report scheduled in {seconds_until_8am/3600:.1f} hours")
        await asyncio.sleep(seconds_until_8am)

        # Fetch current inventory from database
        db = SessionLocal()
        try:
            total      = db.query(Towel).count()
            registered = db.query(Towel).filter(Towel.status == "registered").count()
            in_use     = db.query(Towel).filter(Towel.status == "in_use").count()
            in_laundry = db.query(Towel).filter(Towel.status == "in_laundry").count()
            missing    = db.query(Towel).filter(Towel.status == "missing").count()

            # Get list of missing towels for detail section
            missing_towels = db.query(Towel).filter(Towel.status == "missing").all()

        finally:
            db.close()

        # Decide if this is urgent (5 or more missing)
        is_urgent = missing >= 5
        subject   = f"🚨 URGENT — {missing} Towels Missing | Linen Report" if is_urgent else f"📊 Daily Linen Report — {datetime.datetime.utcnow().strftime('%d %b %Y')}"

        # Build missing towels detail rows for the email
        if missing_towels:
            missing_rows = "".join([
                f"<tr><td style='padding:8px;border-bottom:1px solid #fee2e2;font-family:monospace'>{t.tag_id}</td>"
                f"<td style='padding:8px;border-bottom:1px solid #fee2e2'>{t.towel_type or '—'}</td>"
                f"<td style='padding:8px;border-bottom:1px solid #fee2e2'>{t.last_location or '—'}</td></tr>"
                for t in missing_towels
            ])
            missing_section = f"""
                <h3 style='color:#991b1b;margin-top:24px'>Missing Towels</h3>
                <table style='width:100%;border-collapse:collapse;background:#fef2f2;border-radius:8px'>
                    <tr style='background:#991b1b;color:white'>
                        <th style='padding:8px;text-align:left'>Tag ID</th>
                        <th style='padding:8px;text-align:left'>Type</th>
                        <th style='padding:8px;text-align:left'>Last Location</th>
                    </tr>
                    {missing_rows}
                </table>
            """
        else:
            missing_section = "<p style='color:#166534;background:#f0fdf4;padding:12px;border-radius:8px'>✓ No missing towels today</p>"

        # Build the full HTML email
        urgent_banner = f"<div style='background:#dc2626;color:white;padding:12px;border-radius:8px;margin-bottom:20px;font-weight:bold'>⚠ ALERT: {missing} towels have been missing for over 24 hours. Immediate investigation required.</div>" if is_urgent else ""

        html_body = f"""
        <div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#374151'>
            <div style='background:#1e3a5f;padding:20px;border-radius:8px 8px 0 0'>
                <h1 style='color:white;margin:0;font-size:20px'>Smart Linen Tracking System</h1>
                <p style='color:#93c5fd;margin:4px 0 0'>Daily Report — {datetime.datetime.utcnow().strftime('%d %B %Y')}</p>
            </div>
            <div style='background:white;padding:24px;border:1px solid #e5e7eb;border-radius:0 0 8px 8px'>
                {urgent_banner}
                <h2 style='color:#1e3a5f;margin-top:0'>Inventory Summary</h2>
                <table style='width:100%;border-collapse:collapse'>
                    <tr style='background:#f8fafc'>
                        <td style='padding:10px;border:1px solid #e5e7eb'>Total Towels</td>
                        <td style='padding:10px;border:1px solid #e5e7eb;font-weight:bold;text-align:center'>{total}</td>
                    </tr>
                    <tr>
                        <td style='padding:10px;border:1px solid #e5e7eb'>Registered</td>
                        <td style='padding:10px;border:1px solid #e5e7eb;text-align:center;color:#166534'>{registered}</td>
                    </tr>
                    <tr style='background:#f8fafc'>
                        <td style='padding:10px;border:1px solid #e5e7eb'>In Use</td>
                        <td style='padding:10px;border:1px solid #e5e7eb;text-align:center;color:#1d4ed8'>{in_use}</td>
                    </tr>
                    <tr>
                        <td style='padding:10px;border:1px solid #e5e7eb'>In Laundry</td>
                        <td style='padding:10px;border:1px solid #e5e7eb;text-align:center;color:#92400e'>{in_laundry}</td>
                    </tr>
                    <tr style='background:#fef2f2'>
                        <td style='padding:10px;border:1px solid #e5e7eb;color:#991b1b;font-weight:bold'>Missing</td>
                        <td style='padding:10px;border:1px solid #e5e7eb;text-align:center;color:#991b1b;font-weight:bold'>{missing}</td>
                    </tr>
                </table>
                {missing_section}
                <p style='color:#9ca3af;font-size:12px;margin-top:24px;border-top:1px solid #e5e7eb;padding-top:12px'>
                    This is an automated report from your Smart Linen Tracking System.<br>
                    View live dashboard: <a href='https://web-production-51ae0.up.railway.app'>web-production-51ae0.up.railway.app</a>
                </p>
            </div>
        </div>
        """

        send_email(subject, html_body)

#----------------------------------------------------------------------------------------------------------------------------------------------------------

# ============================================================
# DATA MODEL — what a "register towel" request must include
# Pydantic checks this automatically before your function runs
# ============================================================

class TowelCreate(BaseModel):
    tag_id: str       # The RFID tag number, e.g. "TOWEL-001"
    towel_type: str   # e.g. "bath", "hand", "pool"


# ============================================================
# ENDPOINTS
# ============================================================

# --- Register a new towel ---
# POST /towels
# Body: { "tag_id": "TOWEL-001", "towel_type": "bath" }
@app.post("/towels")
def register_towel(towel: TowelCreate , _=Depends(verify_key)):
    db = SessionLocal()

    # Block duplicate registrations — each RFID tag is unique
    existing = db.query(Towel).filter(Towel.tag_id == towel.tag_id).first()
    if existing:
        raise HTTPException(status_code=400, detail="Tag ID already registered")

    new_towel = Towel(
        tag_id=towel.tag_id,
        towel_type=towel.towel_type,
        status="registered",    # Fresh towels start in "registered" state
        wash_count=0,
        created_at=datetime.datetime.utcnow()
    )
    db.add(new_towel)
    db.commit()

    # Log the registration as an event (for audit trail / history)
    event = Event(
        tag_id=towel.tag_id,
        event_type="REGISTERED",
        created_at=datetime.datetime.utcnow()
    )
    db.add(event)
    db.commit()
    db.close()
    return {"message": "Towel registered successfully", "tag_id": towel.tag_id}


# --- Get all towels ---
# GET /towels
# Returns every towel in the system
@app.get("/towels")
def get_all_towels():
    db = SessionLocal()
    towels = db.query(Towel).all()
    db.close()
    return towels


# --- Get one towel by tag ID ---
# GET /towels/TOWEL-001
@app.get("/towels/{tag_id}")
def get_towel(tag_id: str):
    db = SessionLocal()
    towel = db.query(Towel).filter(Towel.tag_id == tag_id).first()
    db.close()
    if not towel:
        raise HTTPException(status_code=404, detail="Towel not found")
    return towel


# --- Dispatch a towel to a room/location ---
# PATCH /towels/TOWEL-001/dispatch?location=Room202
# Blocks if towel is already out (prevents double dispatch)
@app.patch("/towels/{tag_id}/dispatch")
def dispatch_towel(tag_id: str, location: Optional[str] = None, _=Depends(verify_key)):
    db = SessionLocal()
    towel = db.query(Towel).filter(Towel.tag_id == tag_id).first()
    if not towel:
        raise HTTPException(status_code=404, detail="Towel not found")

    # Guard: block dispatch if towel is already out in a room
    if towel.status == "in_use":
        raise HTTPException(status_code=400, detail="Towel already dispatched — return it first")

    towel.status = "in_use"
    towel.dispatched_at = datetime.datetime.utcnow()   # Record when it left — used by missing detection
    towel.wash_count += 1                              # Increment lifetime wash counter
    if location:
        towel.last_location = location

    event = Event(
        tag_id=tag_id,
        event_type="DISPATCHED",
        location=location,
        created_at=datetime.datetime.utcnow()
    )
    db.add(event)
    db.commit()

    result = {
        "message": "Towel dispatched",
        "tag_id": tag_id,
        "wash_count": towel.wash_count,
        "location": towel.last_location,
        "status": towel.status
    }
    db.close()
    return result


# --- Return a towel from a room ---
# PATCH /towels/TOWEL-001/return
# Moves towel to "in_laundry" and clears its dispatch timestamp
@app.patch("/towels/{tag_id}/return")
def return_towel(tag_id: str, _=Depends(verify_key)):
    db = SessionLocal()
    towel = db.query(Towel).filter(Towel.tag_id == tag_id).first()
    if not towel:
        raise HTTPException(status_code=404, detail="Towel not found")

    towel.status = "in_laundry"
    towel.dispatched_at = None   # Clear the timestamp so missing detection ignores this towel

    event = Event(
        tag_id=tag_id,
        event_type="RETURNED",
        created_at=datetime.datetime.utcnow()
    )
    db.add(event)
    db.commit()

    result = {
        "message": "Towel returned to laundry",
        "tag_id": tag_id,
        "status": towel.status
    }
    db.close()
    return result


# --- Get currently missing towels ---
# GET /missing
# Returns towels that are in_use but overdue by more than 24 hours
# Note: this is a live query — it doesn't rely on the background job having run
@app.get("/missing")
def get_missing_towels():
    db = SessionLocal()
    threshold = datetime.datetime.utcnow() - datetime.timedelta(hours=24)

    # Same logic as the background job, but on-demand instead of scheduled
    missing = db.query(Towel).filter(
        Towel.status == "in_use",
        Towel.dispatched_at < threshold
    ).all()
    db.close()
    return {"missing_count": len(missing), "towels": missing}

  # --- Get inventory summary ---
# GET /inventory
# Returns a count breakdown of all towels by status
# Useful for dashboard views and quick stock checks
@app.get("/inventory")
def get_inventory():
    db = SessionLocal()
    try:
        # Get the total number of towels in the system
        total = db.query(Towel).count()

        # Count how many towels are in each status
        # .count() is like asking "how many rows match this filter?"
        registered = db.query(Towel).filter(Towel.status == "registered").count()
        in_use     = db.query(Towel).filter(Towel.status == "in_use").count()
        in_laundry = db.query(Towel).filter(Towel.status == "in_laundry").count()
        missing    = db.query(Towel).filter(Towel.status == "missing").count()

        return {
            "total":      total,
            "registered": registered,
            "in_use":     in_use,
            "in_laundry": in_laundry,
            "missing":    missing
        }
    finally:
        db.close()  # Always close, even if something goes wrong  

#-----------------------------------------------------------------------------------------------------------------------------

# --- Serve the staff terminal page ---
# GET /staff
# Browser will prompt for username and password automatically
# Returns the staff HTML page if credentials are correct
@app.get("/staff")
def staff_page(username: str = Depends(verify_staff)):
    from fastapi.responses import HTMLResponse
    # Read the staff.html file and serve it
    with open("staff.html", "r") as f:
        return HTMLResponse(content=f.read())

#-------------------------------------------------------------------------------------------------------------------------------
# --- Export full inventory report as Excel ---
@app.get("/export")
# GET /export
def export_excel(
    status:   Optional[str] = None,   # Filter by status e.g. "in_use"
    type:     Optional[str] = None,   # Filter by towel type e.g. "bath"
    location: Optional[str] = None,   # Filter by last location e.g. "Floor 3"
    search:   Optional[str] = None,   # Search by tag ID e.g. "TOWEL-00"
    _=Depends(verify_key)
):
    db = SessionLocal()
    try:
        # Start with all towels then apply filters one by one
        query = db.query(Towel)
        if status:   query = query.filter(Towel.status == status)
        if type:     query = query.filter(Towel.towel_type == type)
        if location: query = query.filter(Towel.last_location == location)
        if search:   query = query.filter(Towel.tag_id.ilike(f"%{search}%"))
        towels = query.all()

        # Events and missing always show full data regardless of filters
        events  = db.query(Event).order_by(Event.created_at.desc()).all()
        missing = db.query(Towel).filter(
            Towel.status == "in_use",
            Towel.dispatched_at < datetime.datetime.utcnow() - datetime.timedelta(hours=24)
        ).all()

        # Full inventory counts always reflect entire database, not filtered subset
        total_all   = db.query(Towel).count()
        registered  = db.query(Towel).filter(Towel.status == "registered").count()
        in_use      = db.query(Towel).filter(Towel.status == "in_use").count()
        in_laundry  = db.query(Towel).filter(Towel.status == "in_laundry").count()
        n_missing   = db.query(Towel).filter(Towel.status == "missing").count()

        # ---- Styles ---- (same as before)
        HOTEL_BLUE   = "1E3A5F"
        WHITE        = "FFFFFF"
        HEADER_GRAY  = "F8FAFC"
        LIGHT_BLUE   = "EFF6FF"
        GREEN_BG     = "F0FDF4"
        RED_BG       = "FEF2F2"
        ORANGE_BG    = "FFFBEB"
        BORDER_COLOR = "E2E8F0"
        STATUS_COLORS = {
            "in_use":     "DBEAFE",
            "in_laundry": "FEF9C3",
            "registered": "DCFCE7",
            "missing":    "FEE2E2",
        }
        EVENT_COLORS = {
            "DISPATCHED": "DBEAFE",
            "RETURNED":   "DCFCE7",
            "MISSING":    "FEE2E2",
            "REGISTERED": "F3F4F6",
        }

        thin   = Side(style='thin', color=BORDER_COLOR)
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        def hcell(ws, r, c, v, bg=HOTEL_BLUE, fg=WHITE):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font  = Font(name='Arial', bold=True, color=fg, size=11)
            cell.fill  = PatternFill('solid', start_color=bg)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
            return cell

        def dcell(ws, r, c, v, bg=WHITE, bold=False, align='left'):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font  = Font(name='Arial', bold=bold, size=10)
            cell.fill  = PatternFill('solid', start_color=bg)
            cell.alignment = Alignment(horizontal=align, vertical='center')
            cell.border = border
            return cell

        def title_block(ws, title, subtitle):
            ws.row_dimensions[1].height = 36
            ws.merge_cells('A1:H1')
            c = ws['A1']
            c.value = title
            c.font  = Font(name='Arial', bold=True, size=16, color=WHITE)
            c.fill  = PatternFill('solid', start_color=HOTEL_BLUE)
            c.alignment = Alignment(horizontal='left', vertical='center')
            ws.merge_cells('A2:H2')
            c2 = ws['A2']
            c2.value = subtitle
            c2.font  = Font(name='Arial', size=10, color="6B7280")
            c2.fill  = PatternFill('solid', start_color=HEADER_GRAY)
            c2.alignment = Alignment(horizontal='left', vertical='center')

        generated_at = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")

        # Show which filters were active in the subtitle
        filter_desc = []
        if status:   filter_desc.append(f"Status: {status}")
        if type:     filter_desc.append(f"Type: {type}")
        if location: filter_desc.append(f"Location: {location}")
        if search:   filter_desc.append(f"Search: {search}")
        filter_note = "  |  Filters: " + ", ".join(filter_desc) if filter_desc else "  |  No filters — all data"

        wb = Workbook()
        wb.remove(wb.active)

        # ---- Sheet 1: Summary (always full counts) ----
        ws1 = wb.create_sheet("Summary")
        title_block(ws1, "Smart Linen Tracking — Inventory Summary", f"Generated: {generated_at}{filter_note}")
        for col, h in enumerate(["Metric","Count","% of Total","Notes"], 1):
            hcell(ws1, 4, col, h)
        rows = [
            ("Total Towels", total_all,   None,       "All towels in system",           LIGHT_BLUE),
            ("Registered",   registered,  "=B6/B5",   "In inventory, not dispatched",   GREEN_BG),
            ("In Use",       in_use,      "=B7/B5",   "Currently out in rooms",         "DBEAFE"),
            ("In Laundry",   in_laundry,  "=B8/B5",   "Returned, being washed",         ORANGE_BG),
            ("Missing",      n_missing,   "=B9/B5",   "Over 24hrs — investigate",       RED_BG),
        ]
        for i, (metric, count, pct, note, bg) in enumerate(rows):
            r = i + 5
            dcell(ws1, r, 1, metric, bg=bg, bold=(i==0))
            dcell(ws1, r, 2, count,  bg=bg, align='center')
            if pct:
                cell = ws1.cell(row=r, column=3, value=pct)
                cell.number_format = '0.0%'
                cell.font  = Font(name='Arial', size=10)
                cell.fill  = PatternFill('solid', start_color=bg)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            else:
                dcell(ws1, r, 3, "—", bg=bg, align='center')
            dcell(ws1, r, 4, note, bg=bg)
        for col, w in enumerate([22,12,14,36], 1):
            ws1.column_dimensions[get_column_letter(col)].width = w

        # ---- Sheet 2: Filtered Towels ----
        ws2 = wb.create_sheet("All Towels")
        title_block(ws2, f"Towels ({len(towels)} records)", f"Generated: {generated_at}{filter_note}")
        for col, h in enumerate(["Tag ID","Type","Status","Last Location","Wash Count","Registered"], 1):
            hcell(ws2, 4, col, h)
        for i, t in enumerate(towels):
            r  = i + 5
            bg = STATUS_COLORS.get(t.status, WHITE)
            dcell(ws2, r, 1, t.tag_id, bg=bg)
            dcell(ws2, r, 2, (t.towel_type or "—").title(), bg=bg)
            dcell(ws2, r, 3, t.status.replace("_"," ").upper(), bg=bg, align='center')
            dcell(ws2, r, 4, t.last_location or "—", bg=bg)
            dcell(ws2, r, 5, t.wash_count or 0, bg=bg, align='center')
            dcell(ws2, r, 6, str(t.created_at or "—"), bg=bg)
        ws2.auto_filter.ref = f"A4:F{4+len(towels)}"
        ws2.freeze_panes   = "A5"
        for col, w in enumerate([16,14,16,18,14,20], 1):
            ws2.column_dimensions[get_column_letter(col)].width = w

        # ---- Sheet 3: Missing ----
        ws3 = wb.create_sheet("Missing Towels")
        title_block(ws3, "Missing Towels Report", f"Generated: {generated_at}")
        for col, h in enumerate(["Tag ID","Type","Last Location","Dispatched At","Action"], 1):
            hcell(ws3, 4, col, h, bg="991B1B")
        if missing:
            for i, t in enumerate(missing):
                r = i + 5
                dcell(ws3, r, 1, t.tag_id, bg=RED_BG)
                dcell(ws3, r, 2, (t.towel_type or "—").title(), bg=RED_BG)
                dcell(ws3, r, 3, t.last_location or "—", bg=RED_BG)
                dcell(ws3, r, 4, str(t.dispatched_at or "—"), bg=RED_BG)
                dcell(ws3, r, 5, "Investigate immediately", bg=RED_BG)
        else:
            ws3.merge_cells("A5:E5")
            c = ws3['A5']
            c.value = "✓  No missing towels at time of export"
            c.font  = Font(name='Arial', bold=True, size=11, color="166534")
            c.fill  = PatternFill('solid', start_color=GREEN_BG)
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws3.row_dimensions[5].height = 32
        ws3.auto_filter.ref = "A4:E4"
        ws3.freeze_panes   = "A5"
        for col, w in enumerate([16,14,18,22,24], 1):
            ws3.column_dimensions[get_column_letter(col)].width = w

        # ---- Sheet 4: Event History ----
        ws4 = wb.create_sheet("Event History")
        title_block(ws4, "Event History Log", f"Generated: {generated_at}  |  {len(events)} events")
        for col, h in enumerate(["#","Tag ID","Event","Location","Date & Time"], 1):
            hcell(ws4, 4, col, h)
        for i, e in enumerate(events):
            r  = i + 5
            bg = EVENT_COLORS.get(e.event_type, WHITE)
            dcell(ws4, r, 1, i+1, bg=bg, align='center')
            dcell(ws4, r, 2, e.tag_id, bg=bg)
            dcell(ws4, r, 3, e.event_type, bg=bg, align='center')
            dcell(ws4, r, 4, e.location or "—", bg=bg)
            dcell(ws4, r, 5, str(e.created_at or "—"), bg=bg)
        ws4.auto_filter.ref = f"A4:E{4+len(events)}"
        ws4.freeze_panes   = "A5"
        for col, w in enumerate([6,16,16,18,24], 1):
            ws4.column_dimensions[get_column_letter(col)].width = w

        # Stream directly to browser — no temp file saved on server
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        filename = f"linen_report_{datetime.datetime.utcnow().strftime('%Y%m%d_%H%M')}.xlsx"
        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    finally:
        db.close()

#----------------------------------------------------------------------------------------------------------------------------------------------

# TEMPORARY — test email, delete after confirming it works
@app.get("/test-email")
def test_email(_=Depends(verify_key)):
    db = SessionLocal()
    try:
        total      = db.query(Towel).count()
        registered = db.query(Towel).filter(Towel.status == "registered").count()
        in_use     = db.query(Towel).filter(Towel.status == "in_use").count()
        in_laundry = db.query(Towel).filter(Towel.status == "in_laundry").count()
        missing    = db.query(Towel).filter(Towel.status == "missing").count()
    finally:
        db.close()

    send_email(
        subject="🧪 Test — Linen Tracker Email Working",
        html_body=f"<p>Test email from your linen tracker!</p><p>Current counts: {total} total, {in_use} in use, {missing} missing.</p>"
    )
    return {"message": "Test email sent — check your inbox"}

#---------------------------------------------------------------------------------------------------------------------------------------------------
