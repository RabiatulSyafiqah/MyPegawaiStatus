import asyncio
import os
import re
import json
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional

from dotenv import load_dotenv
load_dotenv()

from telegram import (
    Update,
    ReplyKeyboardRemove,
    ReplyKeyboardMarkup,
    KeyboardButton,
)
from telegram.ext import (
    Application,
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)
import gspread
from google.oauth2.service_account import Credentials

# Imports for Google Calendar integration and timezones
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from zoneinfo import ZoneInfo

from flask import Flask, request, Response
import threading

# -----------------------------
# Configuration
# -----------------------------
DATE_FMT = "%d/%m/%Y"
TIME_FMT = "%H:%M"  # 24-hour

# Available Officer
OFFICERS = [
    ("Pegawai Daerah", "DO"),
    ("Penolong Pegawai Daerah (Pentadbiran)", "ADO_PENTADBIRAN"),
    ("Penolong Pegawai Daerah (Pembangunan)", "ADO_PEMBANGUNAN"),
]

_admins = {}

# Load admin creds
a1_name = os.getenv("ADMIN1_NAME")
a1_pass = os.getenv("ADMIN1_PASS")
a2_name = os.getenv("ADMIN2_NAME")
a2_pass = os.getenv("ADMIN2_PASS")
a3_name = os.getenv("ADMIN3_NAME")
a3_pass = os.getenv("ADMIN3_PASS")

if a1_name and a1_pass:
    _admins[a1_name] = a1_pass
if a2_name and a2_pass:
    _admins[a2_name] = a2_pass
if a3_name and a3_pass:
    _admins[a3_name] = a3_pass

ADMIN_CREDENTIALS = _admins

BOT_TOKEN = os.getenv("BOT_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE", "service_account.json")

# Calendar mapping
try:
    from sheet import CALENDAR_IDS as _CALENDAR_IDS_OVERRIDE  # type: ignore
except Exception:
    _CALENDAR_IDS_OVERRIDE = None

CALENDAR_IDS = _CALENDAR_IDS_OVERRIDE or {
    "DO": os.getenv("CAL_DO"),
    "ADO_PENTADBIRAN": os.getenv("CAL_ADO_PENTADBIRAN"),
    "ADO_PEMBANGUNAN": os.getenv("CAL_ADO_PEMBANGUNAN"),
}

# Timezone for event creation 
BOT_TIMEZONE = os.getenv("TIMEZONE", "Asia/Kuching")

# -----------------------------
# Google Sheets helpers
# -----------------------------
def _get_ws():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SPREADSHEET_ID)
    # ensure tab
    try:
        ws = sh.worksheet("Sheet1")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Sheet1", rows=1000, cols=9)
        ws.append_row(
            [
                "date",
                "officer",
                "lokasi",
                "urusan rasmi",
                "status keahlian",
                "start_time",
                "end_time",
                "updated_by",
                "updated_at",
            ]
        )
    return ws

def save_status(
    date_str: str,
    officer_code: str,
    lokasi: str,
    urusan_rasmi: str,
    status_keahlian: str,
    start_time: str,
    end_time: str,
    updated_by: str,
):
    ws = _get_ws()
    ws.append_row(
        [
            date_str,
            officer_code,
            lokasi,
            urusan_rasmi,
            status_keahlian,
            start_time,  # This goes to "masa mula" column
            end_time,    # This goes to "masa tamat" column
            updated_by,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ]
    )

def query_status(date_str: str, officer_code: str) -> List[Dict[str, str]]:
    ws = _get_ws()
    rows = ws.get_all_records()
    results = []
    for r in rows:
        if (r.get("date") == date_str) and (r.get("officer") == officer_code):
            results.append(r)
    return results

def delete_status(date_str: str, officer_code: str, urusan_rasmi: str) -> bool:
    """
    Delete records from spreadsheet that match date, officer, and urusan rasmi.
    Returns True if any records were deleted, False otherwise.
    """
    ws = _get_ws()
    rows = ws.get_all_values()
    
    if not rows:
        return False
    
    headers = rows[0]
    indices_to_delete = []
    
    # Find matching rows
    for i, row in enumerate(rows[1:], start=2):  # start=2 because of 1-based indexing and header row
        if len(row) >= len(headers):
            row_dict = dict(zip(headers, row))
            if (row_dict.get("date") == date_str and 
                row_dict.get("officer") == officer_code and 
                row_dict.get("urusan rasmi") == urusan_rasmi):
                indices_to_delete.append(i)
    
    # Delete from bottom to top to maintain correct indices
    indices_to_delete.sort(reverse=True)
    for index in indices_to_delete:
        ws.delete_rows(index)
    
    return len(indices_to_delete) > 0
    
# -----------------------------
# Google Calendar helpers
# -----------------------------
def _get_calendar_service():
    """
    Build and return an authorized Google Calendar service using the
    same service account file. Requires Calendar API enabled for the project.
    """
    scopes = [
        "https://www.googleapis.com/auth/calendar",
        "https://www.googleapis.com/auth/calendar.events"
    ]
    
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    service = build("calendar", "v3", credentials=creds)
    return service

def _get_calendar_id_for_officer(officer_code: str) -> Optional[str]:
    return CALENDAR_IDS.get(officer_code)

def _code_to_label(code: str) -> str:
    for lab, c in OFFICERS:
        if c == code:
            return lab
    return code

def create_calendar_event_for_meeting(date_str: str, start_time: str, end_time: str, officer_code: str, lokasi: str, urusan_rasmi: str, status_keahlian: str) -> bool:
    import sys
    
    # Force flush output to ensure Render shows it
    def log_msg(msg):
        print(f"CALENDAR_DEBUG: {msg}")
        sys.stdout.flush()
        
    log_msg(f"Starting event creation for {officer_code} on {date_str}")
    
    cal_id = _get_calendar_id_for_officer(officer_code)
    if not cal_id:
        log_msg(f"ERROR: No calendar ID for {officer_code}")
        return False
        
    log_msg(f"Using calendar ID: {cal_id}")
    
    try:
        # Test service creation first
        log_msg("Creating calendar service...")
        service = _get_calendar_service()
        log_msg("Service created successfully")
        
        # Parse timezone
        log_msg("Setting up timezone...")
        try:
            tz = ZoneInfo(BOT_TIMEZONE)
            log_msg(f"Timezone set to: {BOT_TIMEZONE}")
        except Exception as e:
            log_msg(f"ZoneInfo error, falling back to naive datetimes: {e}")
            tz = None

        # Parse datetime
        log_msg("Parsing datetime...")
        dt_start = datetime.strptime(f"{date_str} {start_time}", f"{DATE_FMT} {TIME_FMT}")
        dt_end = datetime.strptime(f"{date_str} {end_time}", f"{DATE_FMT} {TIME_FMT}")
        log_msg(f"Parsed start: {dt_start}, end: {dt_end}")

        if tz:
            dt_start = dt_start.replace(tzinfo=tz)
            dt_end = dt_end.replace(tzinfo=tz)
            log_msg("Applied timezone to datetime objects")

        # Create event object with all data
        log_msg("Creating event object...")
        event = {
            "summary": f"KENINGAU — {_code_to_label(officer_code)}",
            "location": lokasi or "",
            "description": f"Urusan Rasmi: {urusan_rasmi}\nStatus Keahlian: {status_keahlian}\nLokasi: {lokasi}",
            "start": {"dateTime": dt_start.isoformat(), "timeZone": BOT_TIMEZONE},
            "end": {"dateTime": dt_end.isoformat(), "timeZone": BOT_TIMEZONE},
            "reminders": {"useDefault": True},
        }
        log_msg(f"Event object created: {json.dumps(event, indent=2)}")
        
        log_msg("Making API call...")
        created = service.events().insert(calendarId=cal_id, body=event).execute()
        log_msg(f"SUCCESS: Event created with ID: {created.get('id')}")
        return True
        
    except HttpError as e:
        log_msg(f"Google Calendar API HTTP error: {e}")
        log_msg(f"HTTP error details: status={e.resp.status}, reason={e.resp.reason}")
        return False
    except Exception as e:
        log_msg(f"EXCEPTION: {type(e).__name__}: {str(e)}")
        import traceback
        log_msg(f"TRACEBACK: {traceback.format_exc()}")
        return False

def create_calendar_event_for_luar_daerah(date_str: str, officer_code: str, urusan_rasmi: str, status_keahlian: str) -> bool:
    """
    Create an all-day event for LUAR DAERAH that blocks the whole day.
    """
    import sys
    
    def log_msg(msg):
        print(f"CALENDAR_DEBUG: {msg}")
        sys.stdout.flush()
        
    log_msg(f"Starting all-day event creation for {officer_code} on {date_str}")
    
    cal_id = _get_calendar_id_for_officer(officer_code)
    if not cal_id:
        log_msg(f"ERROR: No calendar ID for {officer_code}")
        return False
        
    log_msg(f"Using calendar ID: {cal_id}")
    
    try:
        service = _get_calendar_service()
        log_msg("Service created successfully")

        # Parse date for all-day event
        log_msg("Parsing date for all-day event...")
        event_date = datetime.strptime(date_str, DATE_FMT).date()
        log_msg(f"Parsed date: {event_date}")

        # Create all-day event with all data
        log_msg("Creating all-day event object...")
        event = {
            "summary": f"LUAR DAERAH — {_code_to_label(officer_code)}",
            "location": "LUAR DAERAH",
            "description": f"Urusan Rasmi: {urusan_rasmi}\nStatus Keahlian: {status_keahlian}\nLokasi: LUAR DAERAH",
            "start": {"date": event_date.isoformat()},
            "end": {"date": (event_date + timedelta(days=1)).isoformat()},  # All-day events use date only
            "reminders": {"useDefault": True},
        }
        log_msg(f"All-day event object created: {json.dumps(event, indent=2)}")
        
        log_msg("Making API call for all-day event...")
        created = service.events().insert(calendarId=cal_id, body=event).execute()
        log_msg(f"SUCCESS: All-day event created with ID: {created.get('id')}")
        return True
        
    except HttpError as e:
        log_msg(f"Google Calendar API HTTP error: {e}")
        log_msg(f"HTTP error details: status={e.resp.status}, reason={e.resp.reason}")
        return False
    except Exception as e:
        log_msg(f"EXCEPTION: {type(e).__name__}: {str(e)}")
        import traceback
        log_msg(f"TRACEBACK: {traceback.format_exc()}")
        return False

def delete_calendar_events(date_str: str, officer_code: str, urusan_rasmi: str) -> bool:
    """
    Delete calendar events that match the date, officer, and urusan rasmi.
    Returns True if events were found and deleted, False otherwise.
    """
    import sys
    
    def log_msg(msg):
        print(f"CALENDAR_DELETE_DEBUG: {msg}")
        sys.stdout.flush()
        
    log_msg(f"Starting event deletion for {officer_code} on {date_str}, urusan: {urusan_rasmi}")
    
    cal_id = _get_calendar_id_for_officer(officer_code)
    if not cal_id:
        log_msg(f"ERROR: No calendar ID for {officer_code}")
        return False
        
    log_msg(f"Using calendar ID: {cal_id}")
    
    try:
        service = _get_calendar_service()
        log_msg("Service created successfully")

        # Parse the date for the entire day range
        event_date = datetime.strptime(date_str, DATE_FMT).date()
        
        # Start of day in the bot's timezone
        time_min = datetime(event_date.year, event_date.month, event_date.day, 0, 0, 0)
        # End of day in the bot's timezone  
        time_max = datetime(event_date.year, event_date.month, event_date.day, 23, 59, 59)
        
        # Apply timezone if available
        if BOT_TIMEZONE:
            try:
                tz = ZoneInfo(BOT_TIMEZONE)
                time_min = time_min.replace(tzinfo=tz)
                time_max = time_max.replace(tzinfo=tz)
                log_msg(f"Applied timezone {BOT_TIMEZONE} to time range")
            except Exception as e:
                log_msg(f"Timezone error: {e}")

        log_msg(f"Searching events from {time_min} to {time_max}")
        
        # Get events for the entire day
        events_result = service.events().list(
            calendarId=cal_id,
            timeMin=time_min.isoformat(),
            timeMax=time_max.isoformat(),
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        
        events = events_result.get('items', [])
        log_msg(f"Found {len(events)} total events for the day")

        deleted_count = 0
        
        # Search for matching events
        for event in events:
            event_description = event.get('description', '')
            event_summary = event.get('summary', '')
            event_id = event.get('id')
            
            log_msg(f"Checking event: {event_summary}")
            log_msg(f"Event description: {event_description}")
            log_msg(f"Looking for urusan: '{urusan_rasmi}' in description")
            
            # Check if this event matches our criteria
            # Look for the urusan rasmi in the description
            if urusan_rasmi in event_description:
                log_msg(f"✓ MATCH FOUND: Event ID {event_id}")
                try:
                    log_msg(f"Deleting event: {event_summary} (ID: {event_id})")
                    service.events().delete(calendarId=cal_id, eventId=event_id).execute()
                    deleted_count += 1
                    log_msg(f"✓ Successfully deleted event {event_id}")
                except HttpError as e:
                    log_msg(f"✗ Failed to delete event {event_id}: {e}")
                    # Check if it's a 404 error (event not found)
                    if e.resp.status == 404:
                        log_msg(f"Event {event_id} not found in calendar - might have been already deleted")
                    else:
                        log_msg(f"HTTP error details: status={e.resp.status}, reason={e.resp.reason}")
                except Exception as e:
                    log_msg(f"✗ Error deleting event {event_id}: {e}")
            else:
                log_msg(f"✗ NO MATCH: '{urusan_rasmi}' not found in description")
        
        log_msg(f"Deleted {deleted_count} events")
        return deleted_count > 0
        
    except HttpError as e:
        log_msg(f"Google Calendar API HTTP error: {e}")
        log_msg(f"HTTP error details: status={e.resp.status}, reason={e.resp.reason}")
        return False
    except Exception as e:
        log_msg(f"EXCEPTION: {type(e).__name__}: {str(e)}")
        import traceback
        log_msg(f"TRACEBACK: {traceback.format_exc()}")
        return False

# -----------------------------
# Utilities
# -----------------------------
def parse_date_ddmmyyyy(s: str) -> Optional[str]:
    try:
        d = datetime.strptime(s.strip(), DATE_FMT)
        return d.strftime(DATE_FMT)
    except Exception:
        return None

def parse_time_hhmm(s: str) -> Optional[str]:
    if not re.fullmatch(r"\d{2}:\d{2}", s.strip()):
        return None
    try:
        t = datetime.strptime(s.strip(), TIME_FMT)
        return t.strftime(TIME_FMT)
    except Exception:
        return None

def _validate_not_past_and_not_weekend(date_str: str) -> Optional[str]:
    try:
        d = datetime.strptime(date_str, DATE_FMT).date()
    except Exception:
        return None

    today = datetime.now().date()
    # past check
    if d < today:
        return None

    # weekend check: weekday() -> 0=Mon .. 6=Sun ; weekend = 5,6
    if d.weekday() >= 5:
        return None

    return d.strftime(DATE_FMT)

# Map label -> code helpers (for reply-keyboard flow)
def officer_label_to_code(label: str) -> Optional[str]:
    for lab, code in OFFICERS:
        if lab == label:
            return code
    return None

def officer_keyboard() -> ReplyKeyboardMarkup:
    keyboard = [[KeyboardButton(label)] for (label, _) in OFFICERS]
    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

def officer_keyboard_simple() -> ReplyKeyboardMarkup:
    keyboard = [[KeyboardButton(label)] for (label, _) in OFFICERS]
    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

def post_check_keyboard() -> ReplyKeyboardMarkup:
    keyboard = [
        [KeyboardButton("Semak Pegawai Lain")],
        [KeyboardButton("Ubah Tarikh Semakan"), KeyboardButton("Semakan Tamat")],
    ]
    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

def yes_no_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [[KeyboardButton("YA"), KeyboardButton("TIDAK")]],
        one_time_keyboard=True,
        resize_keyboard=True,
    )

def attendance_keyboard() -> ReplyKeyboardMarkup:
    # Admin attendance input expects "KENINGAU" or "LUAR DAERAH"
    return ReplyKeyboardMarkup(
        [[KeyboardButton("KENINGAU"), KeyboardButton("LUAR DAERAH")]],
        one_time_keyboard=True,
        resize_keyboard=True,
    )

def membership_status_keyboard() -> ReplyKeyboardMarkup:
    # choices: Pengerusi, Ahli Biasa, Persami, Jemputan
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("Pengerusi"), KeyboardButton("Ahli Biasa")],
            [KeyboardButton("Perasmi"), KeyboardButton("Jemputan")],
        ],
        one_time_keyboard=True,
        resize_keyboard=True,
    )

def role_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [[KeyboardButton("Kakitangan Admin"), KeyboardButton("Kakitangan Biasa")]],
        one_time_keyboard=True,
        resize_keyboard=True,
    )

def admin_main_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("Kemaskini Jadual"), KeyboardButton("Padam Jadual")],
            [KeyboardButton("Selesai")],
        ],
        one_time_keyboard=True,
        resize_keyboard=True,
    )

# Flask integration
_EVENT_LOOP = None  # type: Optional[asyncio.AbstractEventLoop]

def create_flask_app(app_telegram: Application) -> Flask:
    flask_app = Flask(__name__)

    @flask_app.route("/health", methods=["GET"])  # simple uptime/health check
    def healthcheck():
        return ("ok", 200)

    @flask_app.route("/", methods=["GET", "HEAD"])  # simple root page
    def index():
        return "Service is live", 200

    webhook_path = os.getenv("WEBHOOK_PATH", BOT_TOKEN)
    if not webhook_path:
        raise RuntimeError("WEBHOOK_PATH or BOT_TOKEN must be set to define webhook route")
    route = f"/{webhook_path}"

    @flask_app.route(route, methods=["POST"])
    def telegram_webhook():
        data = request.get_json(force=True, silent=True)
        if not data:
            return Response(status=400)
        update = Update.de_json(data, app_telegram.bot)
        # schedule processing on the background event loop
        assert _EVENT_LOOP is not None
        asyncio.run_coroutine_threadsafe(app_telegram.process_update(update), _EVENT_LOOP)
        return Response(status=200)

    # Log available routes for debugging in Render logs
    print("Flask routes:", flask_app.url_map)

    return flask_app
    
# -----------------------------
# Conversation states
# -----------------------------
(
    CHOOSE_ROLE,
    ADMIN_USERNAME,
    ADMIN_PASSWORD,
    ADMIN_MAIN_MENU,
    ADMIN_DATE,
    ADMIN_OFFICER,
    ADMIN_LOCATION,
    ADMIN_OFFICIAL_BUSINESS,
    ADMIN_MEMBERSHIP_STATUS,
    ADMIN_OFFICIAL_BUSINESS_START,
    ADMIN_OFFICIAL_BUSINESS_END,
    STAFF_DATE,
    STAFF_OFFICER,
    ADMIN_CONTINUE_DECISION,
    ADMIN_DELETE_DATE,
    ADMIN_DELETE_OFFICER,
    ADMIN_DELETE_SELECT_EVENT,
    ADMIN_DELETE_CONFIRM,
) = range(18)

# -----------------------------
# Handlers
# -----------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "Selamat datang ke *MyPegawaiStatus!* 🏛️\n\n"
        "Sila pilih kategori kakitangan anda:\n"
        "1. Kakitangan Admin\n"
        "2. Kakitangan Biasa"\
    )
    
    if update.message:
        await update.message.reply_text(text, reply_markup=role_keyboard())
    else:
        await update.effective_chat.send_message(text, reply_markup=role_keyboard())
    return CHOOSE_ROLE

async def choose_role(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role_label = update.message.text.strip()
    context.user_data.clear()

    if role_label == "Kakitangan Admin":
        context.user_data["role"] = "admin"
        await update.message.reply_text("Sila masukkan nama pengguna (username) anda:", reply_markup=ReplyKeyboardRemove())
        return ADMIN_USERNAME
    elif role_label == "Kakitangan Biasa":
        context.user_data["role"] = "staff"
        await update.message.reply_text("Sila masukkan tarikh pilihan (DD/MM/YYYY):", reply_markup=ReplyKeyboardRemove())
        return STAFF_DATE
    else:
        await update.message.reply_text("Pilihan tidak dikenali. Sila pilih daripada papan kekunci.")
        return CHOOSE_ROLE

# --- Admin flow ---
async def admin_username(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["username"] = update.message.text.strip()
    await update.message.reply_text("Sila masukkan kata laluan:")
    return ADMIN_PASSWORD

async def admin_password(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = context.user_data.get("username")
    password = update.message.text.strip()
    if ADMIN_CREDENTIALS.get(username) != password:
        await update.message.reply_text("Kata laluan salah. Sila /start semula.")
        return ConversationHandler.END

    context.user_data["is_admin"] = True
    await update.message.reply_text(
        "Log masuk berjaya! Sila pilih tindakan:",
        reply_markup=admin_main_keyboard()
    )
    return ADMIN_MAIN_MENU

async def admin_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    choice = update.message.text.strip()
    
    if choice == "Kemaskini Jadual":
        await update.message.reply_text("Sila masukkan tarikh pilihan (DD/MM/YYYY):", reply_markup=ReplyKeyboardRemove())
        return ADMIN_DATE
    elif choice == "Padam Jadual":
        await update.message.reply_text("Sila masukkan tarikh untuk dipadam (DD/MM/YYYY):", reply_markup=ReplyKeyboardRemove())
        return ADMIN_DELETE_DATE
    elif choice == "Selesai":
        await update.message.reply_text("Sesi Admin Ditamatkan. Terima Kasih.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    else:
        await update.message.reply_text("Pilihan tidak dikenali. Sila pilih daripada papan kekunci.", reply_markup=admin_main_keyboard())
        return ADMIN_MAIN_MENU

async def admin_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text
    d = parse_date_ddmmyyyy(raw)
    if not d:
        await update.message.reply_text("Tarikh tidak sah. Sila gunakan format DD/MM/YYYY.")
        return ADMIN_DATE

    # validate not past and not weekend
    valid = _validate_not_past_and_not_weekend(d)
    if not valid:
        try:
            parsed = datetime.strptime(d, DATE_FMT).date()
            today = datetime.now().date()
            if parsed < today:
                await update.message.reply_text("Tarikh yang dimasukkan tidak sah! Sila masukkan tarikh pada hari ini/akan datang (DD/MM/YYYY):")
                return ADMIN_DATE
            if parsed.weekday() >= 5:
                await update.message.reply_text("Sila pilih tarikh bekerja (Isnin–Jumaat).")
                return ADMIN_DATE
        except Exception:
            await update.message.reply_text("Tarikh tidak sah. Sila cuba lagi dengan format DD/MM/YYYY.")
            return ADMIN_DATE

    context.user_data["date"] = valid
    await update.message.reply_text("Sila pilih pegawai untuk dikemaskini:", reply_markup=officer_keyboard())
    return ADMIN_OFFICER

async def admin_officer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    label = update.message.text.strip()
    code = officer_label_to_code(label)
    if not code:
        await update.message.reply_text("Pilihan pegawai tidak sah. Sila cuba sekali lagi.")
        return ADMIN_OFFICER
    context.user_data["officer"] = code

    await update.message.reply_text("Sila pilih lokasi:", reply_markup=attendance_keyboard())
    return ADMIN_LOCATION

async def admin_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    choice = update.message.text.strip().upper()
    if choice not in ("KENINGAU", "LUAR DAERAH"):
        await update.message.reply_text("Sila pilih KENINGAU atau LUAR DAERAH dari papan kekunci.")
        return ADMIN_LOCATION
    context.user_data["lokasi"] = choice

    # Both KENINGAU and LUAR DAERAH follow the same flow
    await update.message.reply_text("Namakan urusan rasmi pegawai:", reply_markup=ReplyKeyboardRemove())
    return ADMIN_OFFICIAL_BUSINESS

async def admin_official_business(update: Update, context: ContextTypes.DEFAULT_TYPE):
    urusan_rasmi = update.message.text.strip()
    context.user_data["urusan_rasmi"] = urusan_rasmi

    await update.message.reply_text(
        "Nyatakan status keahlian pegawai:", 
        reply_markup=membership_status_keyboard()
    )
    return ADMIN_MEMBERSHIP_STATUS

async def admin_membership_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    status_keahlian = update.message.text.strip()
    context.user_data["status_keahlian"] = status_keahlian

    # Check if it's LUAR DAERAH - if yes, skip time questions and save directly
    if context.user_data["lokasi"] == "LUAR DAERAH":
        # Save LUAR DAERAH to Sheets (no start/end time needed)
        save_status(
            date_str=context.user_data["date"],
            officer_code=context.user_data["officer"],
            lokasi=context.user_data["lokasi"],
            urusan_rasmi=context.user_data["urusan_rasmi"],
            status_keahlian=context.user_data["status_keahlian"],
            start_time="",  # Empty for LUAR DAERAH
            end_time="",    # Empty for LUAR DAERAH
            updated_by=context.user_data.get("username", "admin"),
        )

        # Add all-day event to Google Calendar for LUAR DAERAH
        cal_ok = create_calendar_event_for_luar_daerah(
            date_str=context.user_data["date"],
            officer_code=context.user_data["officer"],
            urusan_rasmi=context.user_data["urusan_rasmi"],
            status_keahlian=context.user_data["status_keahlian"]
        )

        if cal_ok:
            await update.message.reply_text(
                "Status LUAR DAERAH berjaya dikemaskini ke Google Calendar (acara sepanjang hari).\n\nTerima kasih.",
                reply_markup=ReplyKeyboardRemove()
            )
        else:
            await update.message.reply_text(
                "Status LUAR DAERAH berjaya dikemaskini. (Gagal menambah acara sepanjang hari ke Calendar — semak konfigurasi.)\n\nTerima kasih.",
                reply_markup=ReplyKeyboardRemove()
            )

        await _prompt_admin_continue(update)
        return ADMIN_CONTINUE_DECISION
    else:
        # For KENINGAU, ask for time details
        await update.message.reply_text(
            "Nyatakan masa mula urusan (HH:MM):", 
            reply_markup=ReplyKeyboardRemove()
        )
        return ADMIN_OFFICIAL_BUSINESS_START

async def admin_official_business_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    start_time = parse_time_hhmm(update.message.text)
    if not start_time:
        await update.message.reply_text("Format masa tidak sah. Gunakan HH:MM (cth 09:00).")
        return ADMIN_OFFICIAL_BUSINESS_START
    
    context.user_data["start_time"] = start_time
    await update.message.reply_text("Nyatakan masa tamat urusan (HH:MM):")
    return ADMIN_OFFICIAL_BUSINESS_END

async def admin_official_business_end(update: Update, context: ContextTypes.DEFAULT_TYPE):
    end_time = parse_time_hhmm(update.message.text)
    if not end_time:
        await update.message.reply_text("Format masa tidak sah. Gunakan HH:MM (cth 17:00).")
        return ADMIN_OFFICIAL_BUSINESS_END

    context.user_data["end_time"] = end_time

    # Save to Sheets (for KENINGAU)
    save_status(
        date_str=context.user_data["date"],
        officer_code=context.user_data["officer"],
        lokasi=context.user_data["lokasi"],
        urusan_rasmi=context.user_data["urusan_rasmi"],
        status_keahlian=context.user_data["status_keahlian"],
        start_time=context.user_data["start_time"],
        end_time=context.user_data["end_time"],
        updated_by=context.user_data.get("username", "admin"),
    )

    # Add to Google Calendar (for KENINGAU)
    cal_ok = create_calendar_event_for_meeting(
        date_str=context.user_data["date"],
        start_time=context.user_data["start_time"],
        end_time=context.user_data["end_time"],
        officer_code=context.user_data["officer"],
        lokasi=context.user_data["lokasi"],
        urusan_rasmi=context.user_data["urusan_rasmi"],
        status_keahlian=context.user_data["status_keahlian"]
    )

    if cal_ok:
        await update.message.reply_text(
            "Status berjaya dikemaskini ke Google Calendar.\n\nTerima kasih.",
            reply_markup=ReplyKeyboardRemove()
        )
    else:
        await update.message.reply_text(
            "Status berjaya dikemaskini. (Gagal menambah acara ke Calendar — semak konfigurasi.)\n\nTerima kasih.",
            reply_markup=ReplyKeyboardRemove()
        )

    await _prompt_admin_continue(update)
    return ADMIN_CONTINUE_DECISION

async def _prompt_admin_continue(update: Update):
    await update.message.reply_text(
        "Adakah anda ingin meneruskan kemaskini untuk tarikh atau pegawai lain?",
        reply_markup=yes_no_keyboard(),
    )

async def admin_continue_decision(update: Update, context: ContextTypes.DEFAULT_TYPE):
    choice = update.message.text.strip().upper()
    if choice == "YA":
        await update.message.reply_text("Sila masukkan tarikh pilihan (DD/MM/YYYY):", reply_markup=ReplyKeyboardRemove())
        return ADMIN_DATE
    elif choice == "TIDAK":
        await update.message.reply_text(
            "Sila pilih tindakan seterusnya:",
            reply_markup=admin_main_keyboard()
        )
        return ADMIN_MAIN_MENU
    else:
        await update.message.reply_text("Sila pilih YA atau TIDAK.", reply_markup=yes_no_keyboard())
        return ADMIN_CONTINUE_DECISION

# --- Delete flow ---
async def admin_delete_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text
    d = parse_date_ddmmyyyy(raw)
    if not d:
        await update.message.reply_text("Tarikh tidak sah. Sila gunakan format DD/MM/YYYY.")
        return ADMIN_DELETE_DATE

    context.user_data["delete_date"] = d
    await update.message.reply_text("Sila pilih pegawai untuk dipadam:", reply_markup=officer_keyboard())
    return ADMIN_DELETE_OFFICER

async def admin_delete_officer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    label = update.message.text.strip()
    code = officer_label_to_code(label)
    if not code:
        await update.message.reply_text("Pilihan pegawai tidak sah. Sila cuba sekali lagi.")
        return ADMIN_DELETE_OFFICER
    
    context.user_data["delete_officer"] = code
    
    # Get events for this date and officer
    date_str = context.user_data["delete_date"]
    records = query_status(date_str, code)
    
    if not records:
        await update.message.reply_text(
            f"Tiada rekod untuk {_code_to_label(code)} pada {date_str}.",
            reply_markup=admin_main_keyboard()
        )
        return ADMIN_MAIN_MENU
    
    # Store records for selection
    context.user_data["delete_records"] = records
    
    # Create selection keyboard
    keyboard = []
    for i, record in enumerate(records):
        urusan_rasmi = record.get("urusan rasmi", "Unknown")
        lokasi = record.get("lokasi", "")
        start_time = record.get("masa mula", "")
        end_time = record.get("masa tamat", "")
        
        if lokasi == "LUAR DAERAH":
            display_text = f"LUAR DAERAH: {urusan_rasmi}"
        elif start_time and end_time:
            display_text = f"{start_time}-{end_time}: {urusan_rasmi}"
        else:
            display_text = urusan_rasmi
            
        keyboard.append([KeyboardButton(display_text)])
    
    keyboard.append([KeyboardButton("Kembali ke Menu Utama")])
    
    await update.message.reply_text(
        f"Pilih urusan rasmi untuk dipadam bagi {_code_to_label(code)} pada {date_str}:",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
    )
    return ADMIN_DELETE_SELECT_EVENT

async def admin_delete_select_event(update: Update, context: ContextTypes.DEFAULT_TYPE):
    choice = update.message.text.strip()
    
    if choice == "Kembali ke Menu Utama":
        await update.message.reply_text(
            "Sila pilih tindakan:",
            reply_markup=admin_main_keyboard()
        )
        return ADMIN_MAIN_MENU
    
    records = context.user_data.get("delete_records", [])
    selected_record = None
    
    # Find the matching record
    for record in records:
        urusan_rasmi = record.get("urusan rasmi", "Unknown")
        lokasi = record.get("lokasi", "")
        start_time = record.get("masa mula", "")
        end_time = record.get("masa tamat", "")
        
        if lokasi == "LUAR DAERAH":
            display_text = f"LUAR DAERAH: {urusan_rasmi}"
        elif start_time and end_time:
            display_text = f"{start_time}-{end_time}: {urusan_rasmi}"
        else:
            display_text = urusan_rasmi
            
        if display_text == choice:
            selected_record = record
            break
    
    if not selected_record:
        await update.message.reply_text("Pilihan tidak sah. Sila cuba sekali lagi.")
        return ADMIN_DELETE_SELECT_EVENT
    
    context.user_data["delete_selected"] = selected_record
    
    # Show confirmation
    urusan_rasmi = selected_record.get("urusan rasmi", "")
    lokasi = selected_record.get("lokasi", "")
    status_keahlian = selected_record.get("status keahlian", "")
    
    confirmation_text = (
        f"Adakah anda pasti ingin memadam rekod ini?\n\n"
        f"Pegawai: {_code_to_label(context.user_data['delete_officer'])}\n"
        f"Tarikh: {context.user_data['delete_date']}\n"
        f"Lokasi: {lokasi}\n"
        f"Urusan Rasmi: {urusan_rasmi}\n"
        f"Status Keahlian: {status_keahlian}"
    )
    
    await update.message.reply_text(
        confirmation_text,
        reply_markup=yes_no_keyboard()
    )
    return ADMIN_DELETE_CONFIRM

async def admin_delete_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    choice = update.message.text.strip().upper()
    
    if choice == "YA":
        selected_record = context.user_data["delete_selected"]
        date_str = context.user_data["delete_date"]
        officer_code = context.user_data["delete_officer"]
        urusan_rasmi = selected_record.get("urusan rasmi", "")
        
        # Delete from database
        db_success = delete_status(date_str, officer_code, urusan_rasmi)
        
        # Delete from calendar
        cal_success = delete_calendar_events(date_str, officer_code, urusan_rasmi)
        
        if db_success:
            message = "Rekod berjaya dipadam dari pangkalan data."
            if cal_success:
                message += " Acara juga berjaya dipadam dari kalendar."
            else:
                message += " (Gagal memadam dari kalendar — semak konfigurasi.)"
        else:
            message = "Gagal memadam rekod. Sila cuba lagi."
        
        await update.message.reply_text(
            message,
            reply_markup=admin_main_keyboard()
        )
        return ADMIN_MAIN_MENU
        
    elif choice == "TIDAK":
        await update.message.reply_text(
            "Pemadaman dibatalkan.",
            reply_markup=admin_main_keyboard()
        )
        return ADMIN_MAIN_MENU
    else:
        await update.message.reply_text("Sila pilih YA atau TIDAK.", reply_markup=yes_no_keyboard())
        return ADMIN_DELETE_CONFIRM

# --- Staff flow ---
async def staff_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Expecting DD/MM/YYYY
    raw = update.message.text
    d = parse_date_ddmmyyyy(raw)
    if not d:
        await update.message.reply_text("Tarikh tidak sah. Sila gunakan format DD/MM/YYYY.")
        return STAFF_DATE

    # validate not past and not weekend
    valid = _validate_not_past_and_not_weekend(d)
    if not valid:
        try:
            parsed = datetime.strptime(d, DATE_FMT).date()
            today = datetime.now().date()
            if parsed < today:
                await update.message.reply_text("Tarikh telah berlalu. Sila pilih tarikh hari ini atau tarikh pada masa hadapan.")
                return STAFF_DATE
            if parsed.weekday() >= 5:
                await update.message.reply_text("Tarikh jatuh pada hujung minggu. Sila pilih hari bekerja (Isnin–Jumaat).")
                return STAFF_DATE
        except Exception:
            await update.message.reply_text("Tarikh tidak sah. Sila cuba lagi dengan format DD/MM/YYYY.")
            return STAFF_DATE

    context.user_data["date"] = valid
    # clear the checked-once flag so the flow behaves like first check
    context.user_data.pop("checked_once", None)
    await update.message.reply_text(
        f"Tarikh ditetapkan kepada {valid}. Sila pilih pegawai untuk disemak:",
        reply_markup=officer_keyboard_simple(),
    )
    return STAFF_OFFICER

async def staff_officer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    label = update.message.text.strip()

    # Post-check options
    if label == "Semak Pegawai Lain":
        await update.message.reply_text("Sila pilih pegawai yang ingin disemak:", reply_markup=officer_keyboard_simple())
        return STAFF_OFFICER

    if label == "Ubah Tarikh Semakan":
        await update.message.reply_text("Sila masukkan tarikh baru (DD/MM/YYYY):", reply_markup=ReplyKeyboardRemove())
        return STAFF_DATE

    if label == "Semakan Tamat":
        await update.message.reply_text("Semakan ditamatkan. Terima kasih.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END

    # Otherwise it's expected to be an officer label
    code = officer_label_to_code(label)
    if not code:
        if context.user_data.get("checked_once"):
            await update.message.reply_text(
                "Pilihan pegawai tidak sah. Sila pilih dari papan kekunci.",
                reply_markup=post_check_keyboard(),
            )
        else:
            await update.message.reply_text(
                "Pilihan pegawai tidak sah. Sila pilih dari papan kekunci.",
                reply_markup=officer_keyboard_simple(),
            )
        return STAFF_OFFICER

    # valid officer selected
    context.user_data["officer"] = code
    date_str = context.user_data.get("date")
    if not date_str:
        await update.message.reply_text("Tarikh tidak sah! Sila masukkan tarikh (DD/MM/YYYY):", reply_markup=ReplyKeyboardRemove())
        return STAFF_DATE

    records = query_status(date_str, code)

    if not records:
        await update.message.reply_text("Tiada rekod untuk tarikh tersebut.", reply_markup=post_check_keyboard())
    else:
        lines = []
        for r in records:
            # Use the correct column names from your Google Sheets
            lokasi = r.get("lokasi", "")
            urusan_rasmi = r.get("urusan rasmi", "")
            status_keahlian = r.get("status keahlian", "")
            masa_mula = r.get("masa mula", "")  # This is the start time
            masa_tamat = r.get("masa tamat", "")  # This is the end time
            
            lines.append(f"Lokasi: {lokasi}")
            lines.append(f"Urusan Rasmi: {urusan_rasmi}")
            lines.append(f"Status Keahlian: {status_keahlian}")
            
            # FIXED: Use the correct time column names
            if lokasi == "LUAR DAERAH":
                lines.append("Masa: Sepanjang Hari")
            elif masa_mula and masa_tamat:
                lines.append(f"Masa: {masa_mula} - {masa_tamat}")
            elif masa_mula:
                lines.append(f"Masa: {masa_mula} - (Tamat tidak dinyatakan)")
            elif masa_tamat:
                lines.append(f"Masa: (Mula tidak dinyatakan) - {masa_tamat}")
            else:
                lines.append("Masa: Tidak dinyatakan")
            
            lines.append("---")

        response_text = "\n".join(lines)
        await update.message.reply_text(response_text, reply_markup=post_check_keyboard())

    context.user_data["checked_once"] = True
    return STAFF_OFFICER
    
# --- Misc ---
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Universal cancel command:
    - Clears user_data
    - Removes any custom reply keyboard
    - Sends cancellation confirmation
    - Ends any active conversation
    Works for message updates and callback_query updates.
    """
    # Clear stored state for safety
    try:
        context.user_data.clear()
    except Exception:
        # defensive: if user_data is not available for some reason, ignore
        pass

    # Remove keyboards and notify user (handle message or callback_query)
    try:
        if update.message:
            await update.message.reply_text("Sesi Dibatalkan.", reply_markup=ReplyKeyboardRemove())
        elif update.callback_query:
            # answer callback query then send a message
            try:
                await update.callback_query.answer()
            except Exception:
                pass
            await update.effective_chat.send_message("Sesi Dibatalkan.", reply_markup=ReplyKeyboardRemove())
        else:
            # fallback
            await update.effective_chat.send_message("Sesi Dibatalkan.", reply_markup=ReplyKeyboardRemove())
    except Exception:
        # last resort: ignore errors sending the confirmation
        pass

    return ConversationHandler.END

def main():
    if not BOT_TOKEN or not SPREADSHEET_ID:
        raise RuntimeError("Please set BOT_TOKEN and SPREADSHEET_ID environment variables.")

    # For local development, use polling instead of webhook
    if os.getenv("RENDER"):
        # Webhook mode for Render
        application: Application = ApplicationBuilder().token(BOT_TOKEN).build()
        
        # Register global /cancel so it can be used anywhere
        application.add_handler(CommandHandler("cancel", cancel))

        conv = ConversationHandler(
            entry_points=[CommandHandler("start", start)],
            states={
                CHOOSE_ROLE: [MessageHandler(filters.TEXT & ~filters.COMMAND, choose_role)],

                # Admin states
                ADMIN_USERNAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_username)],
                ADMIN_PASSWORD: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_password)],
                ADMIN_MAIN_MENU: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_main_menu)],
                ADMIN_DATE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_date)],
                ADMIN_OFFICER:  [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_officer)],
                ADMIN_LOCATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_location)],
                ADMIN_OFFICIAL_BUSINESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_official_business)],
                ADMIN_MEMBERSHIP_STATUS: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_membership_status)],
                ADMIN_OFFICIAL_BUSINESS_START: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_official_business_start)],
                ADMIN_OFFICIAL_BUSINESS_END: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_official_business_end)],
                ADMIN_CONTINUE_DECISION: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_continue_decision)],
                
                # Delete states
                ADMIN_DELETE_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_date)],
                ADMIN_DELETE_OFFICER: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_officer)],
                ADMIN_DELETE_SELECT_EVENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_select_event)],
                ADMIN_DELETE_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_confirm)],

                # Staff states
                STAFF_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, staff_date)],
                STAFF_OFFICER: [MessageHandler(filters.TEXT & ~filters.COMMAND, staff_officer)],
            },
            fallbacks=[CommandHandler("cancel", cancel)],
            allow_reentry=True,
        )

        application.add_handler(conv)

        # Configure webhook for Render Web Service
        render_external_url = os.getenv("RENDER_EXTERNAL_URL")
        if not render_external_url:
            raise RuntimeError("RENDER_EXTERNAL_URL environment variable is required for webhook mode on Render.")

        port = int(os.getenv("PORT", "10000"))
        webhook_path = os.getenv("WEBHOOK_PATH", BOT_TOKEN)
        webhook_url = f"{render_external_url.rstrip('/')}/{webhook_path}"

        # Start Application on a dedicated asyncio loop in the background
        global _EVENT_LOOP
        _EVENT_LOOP = asyncio.new_event_loop()
        threading.Thread(target=_EVENT_LOOP.run_forever, daemon=True).start()

        # initialize/start application and set webhook
        asyncio.run_coroutine_threadsafe(application.initialize(), _EVENT_LOOP).result()
        asyncio.run_coroutine_threadsafe(application.start(), _EVENT_LOOP).result()
        asyncio.run_coroutine_threadsafe(application.bot.set_webhook(webhook_url), _EVENT_LOOP).result()

        flask_app = create_flask_app(application)

        print(f"Bot is running (Flask) on 0.0.0.0:{port} with path /{webhook_path}")
        flask_app.run(host="0.0.0.0", port=port)
    else:
        # Polling mode for local development
        application: Application = ApplicationBuilder().token(BOT_TOKEN).build()
        
        # Register global /cancel so it can be used anywhere
        application.add_handler(CommandHandler("cancel", cancel))

        conv = ConversationHandler(
            entry_points=[CommandHandler("start", start)],
            states={
                CHOOSE_ROLE: [MessageHandler(filters.TEXT & ~filters.COMMAND, choose_role)],

                # Admin states
                ADMIN_USERNAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_username)],
                ADMIN_PASSWORD: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_password)],
                ADMIN_MAIN_MENU: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_main_menu)],
                ADMIN_DATE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_date)],
                ADMIN_OFFICER:  [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_officer)],
                ADMIN_LOCATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_location)],
                ADMIN_OFFICIAL_BUSINESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_official_business)],
                ADMIN_MEMBERSHIP_STATUS: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_membership_status)],
                ADMIN_OFFICIAL_BUSINESS_START: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_official_business_start)],
                ADMIN_OFFICIAL_BUSINESS_END: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_official_business_end)],
                ADMIN_CONTINUE_DECISION: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_continue_decision)],
                
                # Delete states
                ADMIN_DELETE_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_date)],
                ADMIN_DELETE_OFFICER: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_officer)],
                ADMIN_DELETE_SELECT_EVENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_select_event)],
                ADMIN_DELETE_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_confirm)],

                # Staff states
                STAFF_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, staff_date)],
                STAFF_OFFICER: [MessageHandler(filters.TEXT & ~filters.COMMAND, staff_officer)],
            },
            fallbacks=[CommandHandler("cancel", cancel)],
            allow_reentry=True,
        )

        application.add_handler(conv)

        print("Bot is running in polling mode...")
        application.run_polling()

if __name__ == "__main__":
    main()
