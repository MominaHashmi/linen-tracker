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

# TEMPORARY DEBUG ROUTE — delete this after fixing
@app.get("/debug-env")
def debug_env():
    return {
        "STAFF_USER": os.getenv("STAFF_USER", "NOT SET"),
        "STAFF_PASS": os.getenv("STAFF_PASS", "NOT SET")
    }


