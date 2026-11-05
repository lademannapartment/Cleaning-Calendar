import calendar
import io
import json
import os
import re
import smtplib
from collections import defaultdict
from datetime import date, datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

from ido_api import (
    fetch_all_reservations, fetch_availability_locks, fetch_client_names,
    stays_by_item, fetch_all_units, locks_by_item,
)

load_dotenv()

# ==================== USTAWIENIA ====================
EXCEL_FILENAME = "harmonogram_sprzatan.xlsx"
CALENDAR_ID = "b2d9ce8ce415f14476e74488efb830d690fec571755b76440bd0f6dbcb12c7f5@group.calendar.google.com"

SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/spreadsheets",
]
POLAND_TZ = ZoneInfo("Europe/Warsaw")

SMTP_SERVER, SMTP_PORT = "smtp.gmail.com", 587
SENDER_EMAIL = "ilona.overflow@gmail.com"
SENDER_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
RECIPIENT_EMAIL = "lademann.apartment@gmail.com"

SHEET_ID = "1S3NmWakTIjCmwHp36jvmT_PxTfQFNGqZVZfwex1yeIA"
SHEET_BASE = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid="
OBJECTS_SHEET = "OBAPI"
GID_NOTES = 1045522512
GID_EMPLOYEE_CALENDARS = 0  # ← gid listy "ID Calendar" (wpisz właściwy)

# Nowe jednostki z IdoBooking od razu jako TAK (True) albo NIE (False)
AUTO_ACTIVATE_NEW_UNITS = True

# Chemia: dopisywana wyłącznie do istniejących uborek (nigdy osobne zadania)
CHEM_MARK = "Comiesięczne zliczanie chemii"
CHEM_DAY_LIMIT = 4                      # maksymalnie tyle adresów z chemią w jednym dniu
CHEM_EXCLUDE_SUBSTR = ("SZEMUD",)       # adresy, których chemia nigdy nie dotyczy
CHEM_HEADER = f"🧴 {CHEM_MARK.upper()} - policz stany i uzupełnij braki:"
CHEM_DONE = "✅"
CHEM_CHECKED = ("TRUE", "TAK", "1", "YES")
CHEM_UNCHECKED = ("FALSE", "NIE", "0", "NO")

# Jednostki, które są jednym mieszkaniem - chemia liczona raz na całość.
# Klucz: początek nazwy (UPPER), wartość: wspólny adres do chemii.
CHEM_MERGE = {
    "TRAKT 99": "TRAKT ŚW. WOJCIECHA 99/3A",   # pokoje 1-3 oraz studia K2 (Rout) i K3 (Queem)
}

# Dopisek do opisu chemii dla wybranych adresów (dodawany pod standardowym tekstem)
CHEM_EXTRA_NOTES = {
    "POW 3A": "Duży magazyn + 16 i apartament, jeśli jest dostępny.",
    "POW 3B": "Duży magazyn + 65, 64 i apartamenty, które akurat są dostępne.",
}

DEFAULT_CHEM_TEXT = """Miejsca trudnodostępne: W pomieszczeniach o niskim standardzie (brak światła, piwnice, obecność insektów/gryzoni) nie ma obowiązku ręcznego liczenia na miejscu. Proszę o wykonanie wyraźnych zdjęć zapasów – stan zostanie zliczony online na ich podstawie.

Priorytet spożywczy: Artykuły takie jak kawa, cukier, herbata, sól i pieprz należy niezwłocznie dostarczyć „do góry” (do części użytkowej), aby były dostępne dla gości/pracowników.

Lista produktów do zliczenia:
Proszę o uzupełnienie stanów faktycznych (sztuki/opakowania) zgodnie z poniższą listą:

Papier toaletowy (szt.)
Płyn do szyb (szt.)
Płyn do podłóg (szt.)
Płyn do naczyń (szt.)
Mydło w płynie (szt.)
Płyn do WC (szt.)
Środek do kamienia (szt.)
Gąbki do naczyń (szt.)
Worki do śmieci 30l (opak.)
Worki do śmieci 60l (opak.)
Płyn do prania (szt.)
Kapsułki do zmywarki (szt.)
Kwasek cytrynowy (szt.)
Sól (szt.)
Pieprz (szt.)
Herbata (opak.)
Kawa (opak.)
Cukier (opak.)
Worki do odkurzacza (szt.)
Wkładki do mopa (szt.)"""

ELEKTROZAMEK = (
    "PASKA 8/8", "GROTA ROWECKIEGO 4/5", "KAMIENNA GROBLA 9/8", "SEREDYŃSKIEGO 17/1 M1",
    "SEREDYŃSKIEGO 17/1 M2", "GŁĘBOKA 14/5", "TRAKT 99", "PARTYZANTÓW 56/4", "CHROBREGO 5A/4",
    "ZELWEROWICZA 27", "WARSZAWSKA 49/10 M1", "GDYNIA MŚCIWOJA 4-6C/3", "GDYNIA MORSKA 151/7",
    "WIENIAWSKIEGO 15/6", "SEMPOŁOWSKA 12/5", "WARYŃSKIEGO 30/31", "KATARZYNKI 4/2",
)

# Blokady: zbyt długie okresy i tarify (zamknięte plany cenowe) nie są blokadami dostępności
MAX_LOCK_DAYS = 120
LOCK_EXCLUDE_NOTES = ("plan cenow", "cenowy")

# Okno skanowania: od dziś na tyle dni do przodu
SCAN_DAYS_AHEAD = 45

STATS = {"added": [], "cancelled": []}

# {ADRES_CHEM_UPPER: {dni, w których obiekt ma blokadę}} - w te dni nie dokładamy chemii
BLOCKED_DAYS = defaultdict(set)


# ==================== OKNO SKANOWANIA ====================

def month_window():
    today = datetime.now(POLAND_TZ).date()
    return today, today + timedelta(days=SCAN_DAYS_AHEAD)


def scan_window():
    return month_window()


def calendar_month_window(today):
    """Od dziś do ostatniego dnia bieżącego miesiąca."""
    last = calendar.monthrange(today.year, today.month)[1]
    return today, today.replace(day=last)


# ==================== NARZĘDZIA ====================

def clean(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    v = str(v).strip()
    return "" if v.lower() == "nan" else v


def norm_id(v):
    s = clean(v)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def read_sheet(gid, **kw):
    r = requests.get(f"{SHEET_BASE}{gid}", timeout=20)
    r.raise_for_status()
    return pd.read_csv(io.StringIO(r.text), **kw)


def read_sheet_by_name(sheet_name, **kw):
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq"
    r = requests.get(url, params={"tqx": "out:csv", "sheet": sheet_name}, timeout=20)
    r.raise_for_status()
    return pd.read_csv(io.StringIO(r.text), **kw)


def fmt_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return s


# ==================== RĘCZNE DOPISKI ====================
AUTO_BEGIN = "──── AUTO (nie edytuj tej sekcji) ────"
AUTO_END = "──── KONIEC AUTO ────"
LEGACY_AUTO_PREFIXES = ("💡 Uwagi:", "➕ Dodatkowe zadania:", "• ", "  📝", "  🔄", "🧴 ", "UID:")
CHECK_TAG = "🔎 SPRAWDŹ:"


def wrap_auto(text):
    return f"{AUTO_BEGIN}\n{text}\n{AUTO_END}"


def split_description(desc):
    lines = (desc or "").split("\n")
    if AUTO_BEGIN in lines and AUTO_END in lines:
        b, e = lines.index(AUTO_BEGIN), lines.index(AUTO_END)
        return lines[:b], lines[b + 1:e], lines[e + 1:]
    starts = [i for i, l in enumerate(lines) if l.startswith(LEGACY_AUTO_PREFIXES)]
    if not starts:
        return [], lines, []
    b = starts[0]
    u = max((i for i, l in enumerate(lines) if l.startswith("UID:")), default=len(lines) - 1)
    return lines[:b], lines[b:u + 1], lines[u + 1:]


def clean_lines(lines):
    return [l for l in lines if l.strip()]


EXT_TAG_RE = re.compile(r"(?:Pokój\s+\S+\s+)?przedluzenie", re.I)


def manual_title(summary):
    m = re.search(r"CZYTAJ OPIS(.*)$", summary or "", re.I)
    if not m:
        return ""
    text = m.group(1).replace("+ CHEMIA", "")
    text = EXT_TAG_RE.sub("", text)
    text = re.sub(r"\s*-\s*-\s*", " - ", text)
    return text.strip(" -+,")


def is_unit_marker(text):
    return bool(re.fullmatch(r"\s*(K[1-4]|[ABCD])\s*", text, re.I))


def summary_rooms(summary):
    head = (summary or "").split(" - CZYTAJ OPIS", 1)[0]
    m = re.search(r"\(([^()]*)\)\s*$", head)
    if not m or is_unit_marker(m.group(1)):
        return []
    inner = re.sub(r"\s*-\s*przedluzenie\b", "", m.group(1), flags=re.I)
    return [r.strip() for r in inner.split(",") if r.strip()]


def auto_rooms(description):
    _, auto, _ = split_description(description or "")
    rooms = set()
    for line in auto:
        # Ищет номер комнаты независимо от того, как начинается строка (с точки, замка или текста)
        m = re.search(r"Pokój\s+([A-Za-z0-9]+)", line, re.I)
        if m:
            rooms.add(m.group(1).strip())
    return rooms


def manual_rooms(existing):
    auto = auto_rooms(existing.get("description", ""))
    return [r for r in summary_rooms(existing.get("summary", "")) if r not in auto]


def add_rooms_to_summary(summary, rooms):
    if not rooms or " - CZYTAJ OPIS" not in summary:
        return summary
    head, tail = summary.split(" - CZYTAJ OPIS", 1)
    extra = ",".join(rooms)
    m = re.search(r"\(([^()]*)\)\s*$", head)
    if m and not is_unit_marker(m.group(1)):
        head = f"{head[:m.start()]}({m.group(1).strip()},{extra})"
    else:
        head = f"{head} ({extra})"
    return f"{head} - CZYTAJ OPIS{tail}"


def guest_key(name):
    return " ".join(sorted(re.sub(r"\s+", " ", (name or "").strip().lower()).split()))


def prepare_body(body, existing):
    if not existing:
        return body
    pre, _, post = split_description(existing.get("description", ""))
    body = {**body, "description": "\n".join(clean_lines(pre) + [body["description"]] + clean_lines(post))}

    new_rooms = set(summary_rooms(body["summary"]))
    kept = [r for r in manual_rooms(existing) if r not in new_rooms]
    body["summary"] = add_rooms_to_summary(body["summary"], kept)

    mt = manual_title(existing.get("summary", ""))
    if mt and "CZYTAJ OPIS" in body["summary"]:
        body["summary"] = f"{body['summary']} {mt}"

    if chem_is_done(existing.get("summary", "")) and not chem_is_done(body["summary"]):
        body["summary"] = f"{CHEM_DONE} {body['summary']}"
    return body


def review_manual(ev, body_summary, where):
    issues = []
    if place_key(ev["summary"]) != place_key(body_summary):
        issues.append(f"adres różni się: „{ev['summary']}” → „{body_summary}”")
    mt = manual_title(ev["summary"])
    if mt:
        issues.append(f"dopisek w tytule (zachowano): „{mt}”")
    pre, _, post = split_description(ev.get("description", ""))
    extra = clean_lines(pre) + clean_lines(post)
    if extra:
        issues.append("ręczne wpisy w opisie (zachowano): " + " | ".join(l.strip() for l in extra))
    return [f"[{ev['date']}] {where}: {i}" for i in issues]


def room_checks(ev, where):
    address = re.split(r"\s*\(|\s*-\s*CZYTAJ OPIS", ev.get("summary", ""), maxsplit=1, flags=re.I)[0].strip()
    return [
        f"{CHECK_TAG} [{ev['date']}] {address} – pokój {room}: dopisany ręcznie, BRAK rezerwacji w API "
        f"({where}). Sprawdź adres i pokój – czy sprzątanie jest potrzebne, czy to zbędny wpis."
        for room in manual_rooms(ev)
    ]


# ==================== DOPASOWANIE WYDARZEŃ (UID + ADRES) ====================

def place_key(text):
    text = re.sub(r"^✅\s*", "", text or "").strip()
    text = re.sub(r"^test\s+", "", text, flags=re.I)
    head = re.split(r"\s*-\s*CZYTAJ OPIS", text, maxsplit=1, flags=re.I)[0].strip()
    m = re.search(r"\(([^()]*)\)\s*$", head)
    if m and not is_unit_marker(m.group(1)):
        head = head[:m.start()].strip()
    return re.sub(r"\s+", " ", head).strip().upper()


def same_place(summary, base):
    return place_key(summary) == place_key(base)


def index_by_uid(events):
    idx = {}
    for ev in events:
        for u in extract_uids_from_description(ev["description"]):
            idx.setdefault(u, []).append(ev)
    return idx


def find_existing(uids, by_uid, base, where, warnings):
    conflict = False
    for u in uids:
        for ev in by_uid.get(u, []):
            if same_place(ev["summary"], base):
                return ev, False
            conflict = True
            warnings.append(
                f"{CHECK_TAG} [{ev['date']}] {where}: UID {u} jest w wydarzeniu „{ev['summary']}”, "
                f"ale rezerwacja należy do „{base}”. Wydarzenie NIE zostało zmienione ani zduplikowane – "
                f"sprawdź i popraw ręcznie."
            )
    return None, conflict


# ==================== GOOGLE: AUTORYZACJA ====================

def get_credentials():
    if os.environ.get("GOOGLE_SA_JSON"):
        info = json.loads(os.environ["GOOGLE_SA_JSON"])
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    base = Path(__file__).resolve().parent
    return service_account.Credentials.from_service_account_file(
        str(base / "cleaningtest-509509-12cde6c3a6f7.json"), scopes=SCOPES)


# ==================== OBAPI: ODCZYT I ZAPIS JEDNOSTEK ====================

def load_apartments():
    notes = []
    try:
        creds = get_credentials()
        values = build("sheets", "v4", credentials=creds).spreadsheets().values().get(
            spreadsheetId=SHEET_ID, range=f"{OBJECTS_SHEET}!A:G"
        ).execute().get("values", [])
    except Exception as e:
        return [], True, [f"Błąd odczytu zakładki {OBJECTS_SHEET}: {e}"]

    if not values:
        return [], True, [f"Zakładka {OBJECTS_SHEET} jest pusta"]

    header = values[0]
    rows = [dict(zip(header, r + [""] * (len(header) - len(r)))) for r in values[1:]]

    apts, inactive, incomplete = [], 0, False
    for row in rows:
        if str(row.get("aktywny", "")).strip().upper() != "TAK":
            inactive += 1
            continue
        item_id = norm_id(row.get("objectItemId", ""))
        own_name = clean(row.get("Nazwa do kalendarza", ""))
        name = own_name or clean(row.get("Nazwa z IdoBooking", ""))

        if not item_id and not own_name and not clean(row.get("Nazwa z IdoBooking", "")):
            continue  # całkowicie pusty wiersz z przypadkowym TAK
        if not item_id or not name:
            incomplete = True
            notes.append(
                f"Aktywny wiersz OBAPI bez nazwy lub objectItemId (ID: {item_id or '-'}) - usuwanie wstrzymane")
            continue
        if not own_name:
            notes.append(f"Brak „Nazwa do kalendarza” dla ID {item_id} - użyto nazwy z IdoBooking: {name}")
        apts.append({"name": name, "item_id": item_id})

    print(f"🏠 Aktywne jednostki: {len(apts)} (nieaktywne: {inactive})")
    return apts, incomplete, notes


def read_objects_rows(creds):
    sheets = build("sheets", "v4", credentials=creds).spreadsheets().values()
    res = sheets.get(spreadsheetId=SHEET_ID, range=f"{OBJECTS_SHEET}!A:G").execute()
    return res.get("values", [])


def add_new_units_to_sheet(creds, api_units):
    values = read_objects_rows(creds)
    if not values:
        return ["Arkusz OBAPI jest pusty lub nie udało się go odczytać - dopisywanie wstrzymane"], []

    header, body_rows = values[0], values[1:]
    pairs, item_to_object = set(), {}
    for row in body_rows:
        row = row + [""] * (7 - len(row))
        item_id, object_id = norm_id(row[0]), norm_id(row[1])
        if not item_id:
            continue
        pairs.add((item_id, object_id))
        item_to_object.setdefault(item_id, object_id)

    notes, to_add, seen = [], [], set()
    conflicts = 0
    for u in api_units:
        item_id, object_id = u["objectItemId"], u["objectId"]
        key = (item_id, object_id)
        if key in pairs or key in seen:
            continue
        if item_id in item_to_object:
            conflicts += 1
            notes.append(
                f"Konflikt w OBAPI: objectItemId {item_id} jest zapisany z objectId "
                f"{item_to_object[item_id]}, a w IdoBooking ma objectId {object_id} ({u['name']}) - nie dopisano"
            )
            continue
        seen.add(key)
        to_add.append(u)

    print(f"🔎 IdoBooking: {len(api_units)} jednostek, w OBAPI: {len(pairs)}, "
          f"nowe: {len(to_add)}, konflikty: {conflicts}")

    if not to_add:
        print("🆕 Brak nowych jednostek w IdoBooking")
        return notes, []

    status = "TAK" if AUTO_ACTIVATE_NEW_UNITS else "NIE"
    rows = [[u["objectItemId"], u["objectId"], u["name"], u["code"],
             status, u["name"], ""] for u in to_add]

    start_row = len(values) + 1
    last_row = start_row + len(rows) - 1
    try:
        sheets = build("sheets", "v4", credentials=creds).spreadsheets().values()
        sheets.update(
            spreadsheetId=SHEET_ID,
            range=f"{OBJECTS_SHEET}!A{start_row}:G{last_row}",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()
    except Exception as e:
        return notes + [f"Nie udało się dopisać nowych jednostek do OBAPI: {e}"], []

    print(f"🆕 Dopisano do OBAPI nowych jednostek: {len(to_add)} (wiersze {start_row}-{last_row}, aktywny = {status})")
    notes += [f"NOWA jednostka z IdoBooking dopisana do OBAPI jako {status}: {u['name']} "
              f"(objectItemId {u['objectItemId']}, objectId {u['objectId']})" for u in to_add]
    active_new = ([{"name": u["name"], "item_id": u["objectItemId"]} for u in to_add]
                  if AUTO_ACTIVATE_NEW_UNITS else [])
    return notes, active_new


# ==================== ARKUSZ UWAG, PRACOWNICY, ADRESY ====================

def load_notes():
    notes, extras, chem, chem_done = {}, {}, {}, []
    try:
        df = read_sheet(GID_NOTES)
        chem_i = next((i for i, c in enumerate(df.columns) if "chemi" in str(c).lower()), None)
        others = [i for i in range(1, df.shape[1]) if i != chem_i]
        notes_i = others[0] if others else None
        extra_i = others[1] if len(others) > 1 else None

        for _, r in df.iterrows():
            addr = clean(r.iloc[0]).upper()
            if not addr:
                continue
            for target, idx in ((notes, notes_i), (extras, extra_i)):
                text = clean(r.iloc[idx]) if idx is not None else ""
                if text:
                    target[addr] = text
            if chem_i is not None:
                raw = clean(r.iloc[chem_i]).upper()
                if raw in CHEM_CHECKED:
                    chem_done.append(addr)
                elif raw and raw not in CHEM_UNCHECKED:
                    chem[addr] = clean(r.iloc[chem_i])
        print(f"📋 Uwagi: {len(notes)}, dodatkowe zadania: {len(extras)}, "
              f"teksty chemii: {len(chem)}, chemia odhaczona: {len(chem_done)}")
    except Exception as e:
        print(f"⚠️ Błąd odczytu arkusza uwag: {e}")
    return notes, extras, chem, chem_done


EMPLOYEE_SHEET_NAME = "ID Cleaning"  # Точное название вкладки с картинки


def load_employee_calendars():
    calendars = {}
    try:
        # Читаем данные из вкладки "ID Cleaning"
        df = read_sheet_by_name(EMPLOYEE_SHEET_NAME, dtype=str).fillna("")
        for _, r in df.iterrows():
            name = clean(r.iloc[0])  # Колонка A (Cleaning / Имя)
            cal_id = clean(r.iloc[1])  # Колонка B (ID Calendar)
            if name and cal_id and "@" in cal_id:
                # Очищаем суффикс " Calendar" из имени, если он есть, чтобы осталось чистое имя
                clean_name = re.sub(r"\s+Calendar\s*$", "", name, flags=re.I).strip()
                calendars[clean_name or name] = cal_id
        print(f"👥 Kalendarze sprzątaczek (z zakładki '{EMPLOYEE_SHEET_NAME}'): {len(calendars)}")
    except Exception as e:
        print(f"⚠️ Błąd odczytu zakładki '{EMPLOYEE_SHEET_NAME}': {e}")

    return calendars


def addr_key(text):
    s = re.sub(r"\([^()]*\)", "", text or "").upper()
    s = re.sub(r"\bGDYNIA\b", "", s)
    s = re.sub(r"\bPOW\b", "POWSTAŃCÓW WARSZAWSKICH", s)
    s = re.sub(r"\bROWECKIEGO\b", "GROTA ROWECKIEGO", s)
    s = re.sub(r"\bM\.?\s*(?=\d)", "M", s)              # „M. 9” -> „M9”
    s = re.sub(r"(\d+)[A-Z]\b(?=/)", r"\1", s)          # „26A/7” -> „26/7”
    s = re.sub(r"\s+-\s+.*$", "", s)                    # wszystko po „ - ” (nazwy apartamentów)
    s = re.sub(r"\b(APARTMENT|APARTAMENT|ROOM|ELITE|OASIS|WINE|CAMEL)\b.*$", "", s)
    return re.sub(r"[^A-Z0-9]", "", s)


def addr_match(sheet_addr, base):
    a, b = addr_key(sheet_addr), addr_key(base)
    if not a or not b:
        return False
    if len(a) > len(b):
        a, b = b, a
    return b.startswith(a) and (len(a) == len(b) or not b[len(a)].isdigit())


def is_done_in_sheet(base, sheet_done):
    return any(addr_match(s, base) for s in sheet_done)


def parse_apartment_name(name):
    full = name.split("+")[0].strip()
    if re.search(r"\((?:A|B|C|D|K[1-4])\)", full, re.I) or "APARTMENT" in full.upper():
        return full, ""
    if "(" in full and ")" in full:
        base, room = full.split("(")[0].strip(), full.split("(")[1].split(")")[0].strip()
        return (full, "") if re.fullmatch(r"K[1-4]|[ABCD]", room, re.I) else (base, room)
    return full, ""


def cleaning_hours(name):
    u = name.upper()
    if any(a in u for a in ELEKTROZAMEK):
        return 9, 15
    if "NAWYPASIE" in u and "SZEMUD" in u:
        return 13, 17
    if "NAWYPASIE" in u:
        return 13, 16
    if any(k in u for k in ("APARTMENT", "STUDIO", "ŚWIĘTEGO DUCHA")):
        return 12, 17
    return 10, 15


def is_real_lock(lk):
    """Odrzuca tarify (Zamknięty plan cenowy) i zbyt długie okresy."""
    note = (lk.get("note") or "").lower()
    if any(x in note for x in LOCK_EXCLUDE_NOTES):
        return False
    start = datetime.strptime(lk["start"], "%Y-%m-%d").date()
    end = datetime.strptime(lk["end"], "%Y-%m-%d").date()
    return 0 <= (end - start).days <= MAX_LOCK_DAYS


# ==================== GRUPY SPRZĄTAŃ Z API ====================

def add_stays_to_groups(stays, groups, full_name, base, room, today, window_end):
    """
    Rezerwacje i blokady -> uborki (grupa = adres + dzień).
    Rezerwacja: dzień wyjazdu. Blokada: dzień zakończenia blokady.
    Przedłużenie: ten sam gość (imię i nazwisko) na tej samej jednostce wjeżdża w dniu wyjazdu.
    """
    by_arrival = {(guest_key(s["name"]), s["start"]): s
                  for s in stays if not s["blocked"] and s["name"] and s["start"]}

    for x in stays:
        try:
            if x["blocked"] and x["start"] and x["end"]:
                # Dni aktywnej blokady (bez dnia zakończenia, bo wtedy jest sprzątanie)
                addr_key = chem_addr(base).upper()
                first = max(x["start"], today)
                last = min(x["end"] - timedelta(days=1), window_end)
                d = first
                while d <= last:
                    BLOCKED_DAYS[addr_key].add(d)
                    d += timedelta(days=1)

            day = x["end"]
            if not day or not (today <= day <= window_end):
                continue

            nxt = None
            if not x["blocked"] and x["name"]:
                nxt = by_arrival.get((guest_key(x["name"]), day))
                nxt = nxt if nxt is not x else None

            start_h, end_h = cleaning_hours(full_name)
            start_dt = datetime(day.year, day.month, day.day, start_h, 0, tzinfo=POLAND_TZ)

            group_key = f"{day:%Y-%m-%d}_{base.upper()}"
            g = groups.setdefault(group_key, {
                "base": base, "start_dt": start_dt, "end_dt": start_dt.replace(hour=end_h),
                "rooms": [], "uids": [], "details": [], "blocked": False, "ext_rooms": set(),
            })

            g["blocked"] = g["blocked"] or x["blocked"]

            if room and room not in g["rooms"]:
                g["rooms"].append(room)
            if nxt:
                g["ext_rooms"].add(room)
            if x["uid"] in g["uids"]:
                continue
            g["uids"].append(x["uid"])

            if x.get("blocked"):
                note = x.get("note") or "Brak powódu"
                room_label = f"Pokój {room}" if room else "Obiekt"
                detail = f"• {room_label} - 🔒 Blokada, od {x['start']} do {x['end']}: {note}"
            else:
                notes = [x["name"]] if x["name"] else []
                detail = f"• {'Pokój ' + room if room else 'Obiekt'} | Rezerwacja: {x['uid_short']}"
                if notes:
                    detail += "\n  📝 Notatki IdoBooking:\n" + "\n".join(f"     {n}" for n in notes)
                if nxt:
                    detail += (f"\n  🔄 Przedłużenie gościa {x['name']} - sprzątanie NIEPOTRZEBNE. "
                               f"Sprzątanie po nowej rezerwacji {nxt['uid_short']} "
                               f"(wyjazd {nxt['end']:%d.%m.%Y}).")

            if detail not in g["details"]:
                g["details"].append(detail)
        except Exception as e:
            print(f"⚠️ Błąd przetwarzania {full_name}: {e}")


def build_groups(apartments):
    """
    Grupy sprzątań: rezerwacje + blokady w jednej uborce.
    Zwraca (groups, data_incomplete, notes). Gdy API nie odpowiada: (None, None, notes).
    """
    notes = []
    reservations = fetch_all_reservations()
    if reservations is None:
        return None, None, ["Nie udało się pobrać rezerwacji z API"]

    client_names = fetch_client_names()
    if client_names is None:
        client_names = {}
        notes.append("⚠️ Nie udało się pobrać nazwisk gości - przedłużenia po nazwisku nie będą wykrywane")

    locks = fetch_availability_locks()
    if locks is None:
        # Bez blokad nie wiemy, co jest zajęte - nie ruszamy kalendarza
        return None, None, ["Nie udało się pobrać blokad z API - kalendarz bez zmian"]

    total_locks = len(locks)
    valid_locks = [lk for lk in locks if is_real_lock(lk)]
    notes.append(f"🔒 Blokady wpisane: {len(valid_locks)} z {total_locks} "
                 f"(pominięto tarify i zbyt długie okresy)")
    locks_map = locks_by_item(valid_locks)

    by_item = stays_by_item(reservations, client_names)
    groups = {}
    BLOCKED_DAYS.clear()
    data_incomplete = False
    today, window_end = scan_window()

    for apt in apartments:
        full_name = apt["name"]
        base, room = parse_apartment_name(full_name)
        if not apt["item_id"]:
            data_incomplete = True
            notes.append(f"⚠️ Brak objectItemId dla „{full_name}” - usuwanie wstrzymane")
            continue

        stays = by_item.get(apt["item_id"], []) + locks_map.get(apt["item_id"], [])
        add_stays_to_groups(stays, groups, full_name, base, room, today, window_end)

    return groups, data_incomplete, notes


def build_summary(g):
    chem = " + CHEMIA" if g.get("chem") else ""
    ext = g.get("ext_rooms") or set()

    plain = sorted(r for r in g["rooms"] if r and r not in ext)
    head = g["base"] + (f" ({','.join(plain)})" if plain else "")

    tags = [f"Pokój {r} przedluzenie" for r in sorted(r for r in ext if r)]
    if "" in ext:
        tags.append("przedluzenie")

    tail = " - CZYTAJ OPIS"
    if tags:
        tail += " - " + ", ".join(tags)
    return f"{head}{tail}{chem}"


def find_in(base, mapping):
    b = base.upper()
    if b in mapping:
        return mapping[b]
    return next((v for k, v in mapping.items() if k and (k in b or b in k)), "")


def build_event_body(g, notes, extras, chem_texts):
    note, extra = find_in(g["base"], notes), find_in(g["base"], extras)
    lines = ([f"💡 Uwagi: {note}"] if note else []) + ([f"➕ Dodatkowe zadania: {extra}"] if extra else [])
    lines = (lines + [""] if lines else []) + g["details"]
    if g.get("chem"):
        chem_txt = find_in(g["base"], chem_texts) or DEFAULT_CHEM_TEXT
        extra = next((v for k, v in CHEM_EXTRA_NOTES.items() if addr_match(k, g["base"])), "")
        if extra:
            chem_txt = f"{chem_txt}\n\n{extra}"
        lines += ["", CHEM_HEADER, chem_txt]

    lines.append(f"\nUID:{','.join(g['uids'])}")
    summary, description = build_summary(g), wrap_auto("\n".join(lines))

    body = {
        "summary": summary,
        "description": description,
        "start": {"dateTime": g["start_dt"].isoformat(), "timeZone": "Europe/Warsaw"},
        "end": {"dateTime": g["end_dt"].isoformat(), "timeZone": "Europe/Warsaw"},
    }

    if g.get("ext_rooms") or g.get("blocked"):
        body["colorId"] = "4"  # Flamingo: przedłużenia i blokady

    return body


# ==================== GOOGLE CALENDAR ====================

def get_calendar_events(service, target_cal_id, start=None, end=None):
    if start is None or end is None:
        start, end = scan_window()
    t0 = datetime(start.year, start.month, start.day, tzinfo=POLAND_TZ).isoformat()
    t1 = (datetime(end.year, end.month, end.day, tzinfo=POLAND_TZ) + timedelta(days=1)).isoformat()
    events, token = [], None
    try:
        while True:
            r = service.events().list(calendarId=target_cal_id, timeMin=t0, timeMax=t1,
                                      singleEvents=True, maxResults=2500, pageToken=token).execute()
            for e in r.get("items", []):
                st = e.get("start", {})
                events.append({
                    "id": e["id"],
                    "date": (st.get("dateTime") or st.get("date", "")).split("T")[0],
                    "summary": (e.get("summary") or "").strip(),
                    "description": e.get("description") or "",
                })
            token = r.get("nextPageToken")
            if not token:
                return events
    except Exception as e:
        print(f"❌ Błąd pobierania wydarzeń z kalendarza {target_cal_id}: {e}")
        return None


def extract_uids_from_description(description):
    uids = []
    if "UID:" in description:
        for line in description.split("\n"):
            if line.startswith("UID:"):
                for u in line.replace("UID:", "").strip().split(","):
                    if u.strip():
                        uids.append(u.strip())
    return uids


def find_manual_events(events, groups, extra_uids=()):
    found = []
    all_group_uids = {u for g in groups.values() for u in g["uids"]} | set(extra_uids)
    for ev in events:
        ev_uids = extract_uids_from_description(ev["description"])
        if not ev_uids:
            found.append({"date": ev["date"], "summary": f"{ev['summary']} (wpisane ręcznie — brak ID rezerwacji)"})
        elif not any(u in all_group_uids for u in ev_uids):
            found.append({"date": ev["date"], "summary": f"{ev['summary']} (rezerwacja nie istnieje już w API)"})
    return found


# ==================== SYNCHRONIZACJA KALENDARZY SPRZĄTACZEK ====================

def sync_employee_calendars(service, groups, notes, extras, chem_texts, employee_cals, today):
    mismatches = []
    today_s = str(today)

    for emp_name, emp_cal_id in employee_cals.items():
        where = f"pracownik {emp_name}"
        emp_events = get_calendar_events(service, emp_cal_id, today, today)
        if emp_events is None:
            continue

        emp_active = [e for e in emp_events if e["date"] == today_s]
        if not emp_active:
            continue

        for key, g in groups.items():
            g_date = key.split("_")[0]
            if g_date != today_s:
                continue

            g_base = g["base"].upper()
            matching_emp_evs = [
                ev for ev in emp_active
                if g_base in place_key(ev["summary"]) or place_key(ev["summary"]) in g_base
            ]

            if not matching_emp_evs:
                continue

            existing_ev = matching_emp_evs[0]
            body = build_event_body(g, notes, extras, chem_texts)
            body = prepare_body(body, existing_ev)

            mismatches += review_manual(existing_ev, body["summary"], where)
            mismatches += room_checks(existing_ev, where)
            try:
                service.events().patch(calendarId=emp_cal_id, eventId=existing_ev["id"], body=body).execute()

                # Фиксируем обновление в статистике для отчета
                STATS["added"].append({
                    "date": today_s,
                    "summary": body["summary"],
                    "where": where
                })

                print(f"  🔄 Zaktualizowano w kalendarzu ({where}): [{today_s}] {body['summary']}")
            except Exception as e:
                print(f"  ❌ Błąd aktualizacji ({where}): {e}")

    return mismatches


def remove_cancelled_from_employee_calendars(service, employee_cals, employee_events,
                                             groups, data_incomplete):
    """
    Usuwa z kalendarzy sprzątaczek wydarzenia, których rezerwacje/blokady
    nie istnieją już w API (anulowane). Te same zabezpieczenia co w głównym kalendarzu:
      - wydarzenia bez UID (wpisane ręcznie) nie są ruszane,
      - przy niepełnych danych z API/OBAPI nic nie jest usuwane,
      - kalendarz, którego nie udało się pobrać, jest pomijany.
    Jeśli UID nadal istnieje, ale sprzątanie przesunęło się na inny dzień - tylko ostrzeżenie.
    """
    mismatches = []
    uid_dates = {u: key.split("_")[0] for key, g in groups.items() for u in g["uids"]}

    for emp_name in employee_cals:
        events = employee_events.get(emp_name)
        if events is None:
            continue
        where = f"pracownik {emp_name}"
        emp_cal_id = employee_cals[emp_name]

        for ev in events:
            ev_uids = extract_uids_from_description(ev["description"])
            if not ev_uids:
                continue  # wpis ręczny - nie ruszamy

            if any(u in uid_dates for u in ev_uids):
                new_dates = {uid_dates[u] for u in ev_uids if u in uid_dates}
                if ev["date"] not in new_dates:
                    mismatches.append(
                        f"{CHECK_TAG} [{ev['date']}] {where}: „{ev['summary']}” – rezerwacja ma teraz "
                        f"sprzątanie w dniu {', '.join(sorted(new_dates))}. "
                        f"Wydarzenie NIE zostało przesunięte – sprawdź i popraw ręcznie."
                    )
                continue

            if data_incomplete:
                print(f"⚠️ Pominięto usuwanie ({where}, niepełne dane z API/OBAPI): "
                      f"[{ev['date']}] {ev['summary']}")
                continue

            try:
                service.events().delete(calendarId=emp_cal_id, eventId=ev["id"]).execute()
                STATS["cancelled"].append({"date": ev["date"], "summary": ev["summary"], "where": where})
                print(f"🗑️ Usunięto z kalendarza ({where}): [{ev['date']}] {ev['summary']}")
            except Exception as e:
                print(f"⚠️ Błąd usuwania z kalendarza ({where}) {ev['id']}: {e}")

    return mismatches


# ==================== CHEMIA (DOPISYWANA DO UBOREK) ====================

def chem_addr(address):
    """Adres używany do chemii: jednostki z CHEM_MERGE są sklejane w jeden (np. Trakt 99 = jedno mieszkanie)."""
    u = address.upper()
    for frag, target in CHEM_MERGE.items():
        if u.startswith(frag):
            return target
    # Ujednolicenie zapisu: „ZIMNA 7A/5- SEASHELL” i „ZIMNA 7A/5 - SEASHELL”, końcowe „ -” itd.
    norm = re.sub(r"\s*-\s*", " - ", address)
    return re.sub(r"\s+", " ", norm).strip(" -")


def is_chem_event(summary):
    return CHEM_MARK.lower() in (summary or "").lower()


def chem_is_done(summary):
    return (summary or "").strip().startswith(CHEM_DONE)


def chem_excluded(address):
    """Adresy wyłączone z chemii (np. Szemud) - nigdy nie dostają chemii."""
    return any(x in address.upper() for x in CHEM_EXCLUDE_SUBSTR)


def cleanings_in_month(groups, today):
    """
    {ADRES_CHEM_UPPER: [daty uborek od dziś do końca miesiąca]}.
    Uborki wszystkich jednostek sklejonych w jeden adres (np. Trakt 99 pokoje 1-3 + K2 + K3)
    trafiają do wspólnej listy, więc chemia wychodzi raz na mieszkanie.
    """
    _, month_end = calendar_month_window(today)
    by_addr = defaultdict(set)
    for key in groups:
        day_s, b = key.split("_", 1)
        d = date.fromisoformat(day_s)
        addr = chem_addr(b).upper()
        if today <= d <= month_end and d not in BLOCKED_DAYS.get(addr, ()):
            by_addr[addr].add(d)
    return {b: sorted(ds) for b, ds in by_addr.items()}


def chem_done_set(addresses, sheet_done):
    """Adresy z TRUE w arkuszu „Cleaning Uwagi” - chemia policzona."""
    return {a.upper() for a in addresses if is_done_in_sheet(a, sheet_done)}


def chem_status(addresses, groups, today, sheet_done):
    """
    Stan chemii dla każdego adresu (adresy już po chem_addr):
      'sheet'        - policzone (TRUE w arkuszu)
      'todo'         - nie policzone, ma uborkę w tym miesiącu
      'no_cleaning'  - nie policzone, brak uborki w tym miesiącu
      'excluded'     - wyłączone (Szemud)
    """
    done = chem_done_set(addresses, sheet_done)
    cl = cleanings_in_month(groups, today)
    status = {}
    for a in addresses:
        b = a.upper()
        if chem_excluded(a):
            status[b] = "excluded"
        elif b in done:
            status[b] = "sheet"
        elif cl.get(b):
            status[b] = "todo"
        else:
            status[b] = "no_cleaning"
    return status


def print_chem_status(status):
    titles = {
        "sheet": "✅ policzone (TRUE w arkuszu)",
        "todo": "⏳ do zrobienia (jest uborka w tym miesiącu)",
        "no_cleaning": "⚪ brak uborki w tym miesiącu",
        "excluded": "🚫 wyłączone z chemii",
    }
    for kind in ("sheet", "todo", "no_cleaning", "excluded"):
        items = sorted(b for b, s in status.items() if s == kind)
        print(f"🧴 {titles[kind]}: {len(items)}")
        for b in items:
            print(f"   - {b}")


def check_sheet_matching(sheet_done, addresses):
    """Wiersze z TRUE w arkuszu, które nie pasują do żadnego adresu z OBAPI (literówki, inne nazwy)."""
    for s in sheet_done:
        if not any(addr_match(s, a) for a in addresses):
            print(f"⚠️ Arkusz: adres „{s}” nie pasuje do żadnego adresu z OBAPI - sprawdź pisownię")


def chem_month_plan(addresses, today, groups, sheet_done=()):
    """
    Plan chemii: {ADRES_CHEM_UPPER: data uborki}.
    - Chemia idzie wyłącznie na istniejącą uborkę danego adresu w tym miesiącu.
    - Adresy wg kolejności alfabetycznej (stała każdy miesiąc).
    - Najwcześniejsza uborka z wolnym miejscem (< CHEM_DAY_LIMIT adresów w danym dniu).
    - Gdy uborka minęła bez odhaczenia, kolejny adres trafia na następną dostępną uborkę.
    """
    status = chem_status(addresses, groups, today, sheet_done)
    cl = cleanings_in_month(groups, today)
    load = defaultdict(int)
    plan = {}

    for b in sorted(b for b, s in status.items() if s == "todo"):
        for d in cl[b]:
            if load[d] < CHEM_DAY_LIMIT:
                plan[b] = d
                load[d] += 1
                break

    unplanned = [b for b, s in status.items() if s == "todo" and b not in plan]
    if unplanned:
        print(f"⚠️ Chemia: brak wolnego miejsca w uborkach tego miesiąca dla {len(unplanned)} adresów")
    return plan


def apply_chem_to_groups(plan, groups):
    """Dopisuje „+ CHEMIA” tylko do wybranej uborki każdej jednostki sklejonej w adres chemii."""
    for b, day in plan.items():
        for key in list(groups):
            day_s, gb = key.split("_", 1)
            if day_s == f"{day:%Y-%m-%d}" and chem_addr(gb).upper() == b:
                groups[key]["chem"] = True
                break  # jedna uborka na adres = jedna chemia


# ==================== EXCEL + EMAIL ====================

def generate_excel(report):
    rows = [{"Data": d, "Godzina": i["time"], "Obiekt / Pokój": i["summary"]}
            for d in sorted(report) for i in report[d]]
    pd.DataFrame(rows).to_excel(EXCEL_FILENAME, index=False, sheet_name="Harmonogram")
    print(f"📊 Wygenerowano plik Excel: {EXCEL_FILENAME}")


WEEKDAYS_PL = ["poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota", "niedziela"]


def pl_date_label(d):
    """Np. 'Środa, 28.10.2026'."""
    return f"{WEEKDAYS_PL[d.weekday()].capitalize()}, {d:%d.%m.%Y}"


def send_email_report(report, manual, mismatches=(), run_date=None):
    now = datetime.now(POLAND_TZ)
    run_date = datetime.now(POLAND_TZ).date()
    added, cancelled = STATS["added"], STATS["cancelled"]
    checks = [m[len(CHECK_TAG):].strip() for m in mismatches if m.startswith(CHECK_TAG)]
    notes = [m for m in mismatches if not m.startswith(CHECK_TAG)]
    changed = bool(added or cancelled or manual or mismatches)

    msg = MIMEMultipart()
    msg["From"], msg["To"] = SENDER_EMAIL, RECIPIENT_EMAIL
    msg["Subject"] = (f"🧹 [SYNCHRONIZACJA] Harmonogram ({now:%d.%m.%Y})" if changed
                      else f"🧹 Harmonogram sprzątań ({now:%d.%m.%Y})")

    def by_where(items, label):
        main_n = sum(1 for i in items if i["where"] == "główny")
        return f"{len(items)} (główny: {main_n}, {label}: {len(items) - main_n})"

    body = "Hej,<br><br>"
    window_end = run_date + timedelta(days=SCAN_DAYS_AHEAD)
    body += (f"<b>Podsumowanie synchronizacji (okno: od {run_date:%d.%m.%Y} "
             f"do {window_end:%d.%m.%Y}):</b><br>\n")
    body += f"   • Dodano nowych sprzątań: <b>{by_where(added, 'panie sprzątające')}</b><br>\n"
    body += f"   • Anulowano sprzątań: <b>{by_where(cancelled, 'panie sprzątające')}</b><br><br>\n"

    if added:
        body += "🆕 <b>Nowe:</b><br>\n"
        body += "".join(f"   • <b>{fmt_date(i['date'])}</b> — {escape(i['summary'])} ({escape(i['where'])})<br>\n"
                        for i in sorted(added, key=lambda x: x["date"])) + "<br>\n"
    if cancelled:
        body += "❌ <b>Anulowane:</b><br>\n"
        body += "".join(f"   • <b>{fmt_date(i['date'])}</b> — {escape(i['summary'])} ({escape(i['where'])})<br>\n"
                        for i in sorted(cancelled, key=lambda x: x["date"])) + "<br>\n"
    if checks:
        body += "🔎 <b>SPRAWDŹ ADRES / POKÓJ:</b><br>\n"
        body += "".join(f"   • {escape(c)}<br>\n" for c in checks) + "<br>\n"
    if notes:
        body += "✏️ <b>Uwagi i rozbieżności:</b><br>\n"
        body += "".join(f"   • {escape(m)}<br>\n" for m in notes) + "<br>\n"

    body += "Raport na najbliższe dwa dni:<br><br>"
    for offset in (0, 1):
        d = run_date + timedelta(days=offset)
        items = report.get(d.isoformat(), [])
        body += f"📅 <b>{pl_date_label(d)}:</b><br>\n"
        if items:
            body += "".join(f"   • {escape(i['summary'])}<br>\n" for i in items) + "<br>\n"
        else:
            body += "   • brak zadań<br><br>\n"

    body += ("Pełny harmonogram znajdziesz w załączonym pliku Excel."
             "<br><br>Pozdrowienia,<br>Twój System Automatyzacji")
    msg.attach(MIMEText(body, "html", "utf-8"))

    try:
        generate_excel(report)
        with open(EXCEL_FILENAME, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {EXCEL_FILENAME}")
        msg.attach(part)
    except Exception as e:
        print(f"⚠️ Nie udało się dołączyć pliku Excel: {e}")

    if not SENDER_PASSWORD:
        return
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())
        print("📧 Raport e-mail został wysłany!")
    except Exception as e:
        print(f"❌ Błąd wysyłania e-maila: {e}")


# ==================== GŁÓWNA SYNCHRONIZACJA ====================

def sync_cleanings():
    creds = get_credentials()
    service = build("calendar", "v3", credentials=creds)
    STATS["added"].clear()
    STATS["cancelled"].clear()
    run_date = datetime.now(POLAND_TZ).date()

    # 1. Jednostki z IdoBooking -> OBAPI
    new_unit_notes, new_active = [], []
    api_units = fetch_all_units()
    if api_units is None:
        new_unit_notes = ["Nie udało się pobrać listy jednostek z IdoBooking - OBAPI nie zaktualizowano"]
    else:
        try:
            new_unit_notes, new_active = add_new_units_to_sheet(creds, api_units)
        except Exception as e:
            new_unit_notes = [f"Błąd dopisywania nowych jednostek: {e}"]

    # 2. Aktywne jednostki z OBAPI
    apartments, apt_incomplete, apt_notes = load_apartments()
    known = {a["item_id"] for a in apartments}
    apartments += [a for a in new_active if a["item_id"] not in known]

    if not apartments:
        print("❌ Brak aktywnych jednostek w OBAPI!")
        return

    today, window_end = scan_window()
    print(f"📅 Okno skanowania: od {today:%d.%m.%Y} do {window_end:%d.%m.%Y}")

    # Kalendarze sprzątaczek dla całego okna skanowania
    employee_cals = load_employee_calendars()
    employee_assigned_uids = set()
    employee_events = {}  # emp_name -> lista wydarzeń (tylko gdy pobranie się udało)
    for emp_name, emp_cal_id in employee_cals.items():
        emp_events = get_calendar_events(service, emp_cal_id, today, window_end)
        if emp_events is None:
            print(f"⚠️ Nie pobrano kalendarza pracownika {emp_name} - anulowania w nim nie będą sprawdzane")
            continue
        employee_events[emp_name] = emp_events
        for ev in emp_events:
            employee_assigned_uids.update(extract_uids_from_description(ev["description"]))
    print(f"👤 UID przypisane do sprzątaczek (w całym oknie): {len(employee_assigned_uids)}")

    m_start, m_end = calendar_month_window(today)
    all_events = get_calendar_events(service, CALENDAR_ID, min(today, m_start), max(window_end, m_end))
    if all_events is None:
        return

    notes, extras, chem_texts, sheet_done = load_notes()

    # 3. Rezerwacje + blokady (razem, w uborkach)
    groups, data_incomplete, build_notes = build_groups(apartments)
    if groups is None:
        print("❌ Brak danych z API - synchronizacja przerwana, kalendarz bez zmian")
        return
    data_incomplete = data_incomplete or apt_incomplete

    # 3b. Stare osobne wydarzenia blokad ("🔒 ...") usuwamy - blokady są teraz w uborkach
    old_locks = [ev for ev in all_events if ev["summary"].startswith("🔒")]
    for ev in old_locks:
        try:
            service.events().delete(calendarId=CALENDAR_ID, eventId=ev["id"]).execute()
            print(f"🔒➡️ Usunięto stare osobne wydarzenie blokady: [{ev['date']}] {ev['summary']}")
        except Exception as e:
            print(f"⚠️ Błąd usuwania starej blokady {ev['id']}: {e}")
    old_lock_ids = {ev["id"] for ev in old_locks}

    events = [e for e in all_events
              if e["id"] not in old_lock_ids
              and not is_chem_event(e["summary"])
              and str(today) <= e["date"] <= str(window_end)]

    mismatches = list(new_unit_notes) + list(apt_notes) + list(build_notes)
    if old_locks:
        mismatches.append(f"Usunięto {len(old_locks)} starych osobnych wydarzeń blokad - blokady są teraz w uborkach")

    # Adresy do chemii: jednostki sklejone (Trakt 99 + K2 + K3) dają jeden wpis
    addresses = sorted({chem_addr(parse_apartment_name(a["name"])[0]) for a in apartments})

    # 3c. Chemia: sprawdzenie arkusza, status i plan (tylko na istniejące uborki)
    check_sheet_matching(sheet_done, addresses)
    print_chem_status(chem_status(addresses, groups, today, sheet_done))

    chem_plan_now = chem_month_plan(addresses, today, groups, sheet_done)
    apply_chem_to_groups(chem_plan_now, groups)

    manual = find_manual_events(events, groups)
    main_by_uid = index_by_uid(events)
    processed_main_ids = set()

    # 4. Dodawanie / aktualizacja w głównym kalendarzu
    for key, g in groups.items():
        if any(u in employee_assigned_uids for u in g["uids"]):
            print(f"  ⏭️ Pominięto w głównym (przypisane do pań sprzątających): [{key.split('_')[0]}] {g['base']}")
            continue

        body = build_event_body(g, notes, extras, chem_texts)
        existing_ev, conflict = find_existing(g["uids"], main_by_uid, g["base"], "główny", mismatches)
        if conflict and existing_ev is None:
            print(f"  ⛔ Konflikt UID w głównym – pominięto: [{key.split('_')[0]}] {g['base']}")
            continue

        if existing_ev:
            body = prepare_body(body, existing_ev)
            mismatches += review_manual(existing_ev, body["summary"], "główny")
            mismatches += room_checks(existing_ev, "główny")
            try:
                service.events().patch(calendarId=CALENDAR_ID, eventId=existing_ev["id"], body=body).execute()
                processed_main_ids.add(existing_ev["id"])
                print(f"  🔄 ZAKTUALIZOWANO w głównym: [{key.split('_')[0]}] {body['summary']}")
            except Exception as e:
                print(f"  ❌ Błąd aktualizacji {body['summary']}: {e}")
        else:
            try:
                service.events().insert(calendarId=CALENDAR_ID, body=body).execute()
                STATS["added"].append({"date": key.split("_")[0], "summary": body["summary"], "where": "główny"})
                print(f"  ✅ DODANO do głównego: [{key.split('_')[0]}] {body['summary']}")
            except Exception as e:
                print(f"  ❌ Błąd dodawania {body['summary']}: {e}")

    # 5. Usuwanie wydarzeń, których nie ma już w danych z API (z ochroną)
    all_group_uids = {u for g in groups.values() for u in g["uids"]}
    for ev in events:
        ev_uids = extract_uids_from_description(ev["description"])
        if not ev_uids:
            print(f"🛡️ Zabezpieczono wpis ręczny (brak UID): [{ev['date']}] {ev['summary']}")
            continue
        if ev["id"] in processed_main_ids:
            continue
        if not any(u in all_group_uids for u in ev_uids):
            if data_incomplete:
                print(f"⚠️ Pominięto usuwanie (niepełne dane z API/OBAPI): [{ev['date']}] {ev['summary']}")
                continue
            try:
                service.events().delete(calendarId=CALENDAR_ID, eventId=ev["id"]).execute()
                STATS["cancelled"].append({"date": ev["date"], "summary": ev["summary"], "where": "główny"})
                print(f"🗑️ Usunięto z głównego: [{ev['date']}] {ev['summary']}")
            except Exception as e:
                print(f"⚠️ Błąd usuwania {ev['id']}: {e}")

    # 5b. Anulowane sprzątania w kalendarzach sprzątaczek
    mismatches += remove_cancelled_from_employee_calendars(
        service, employee_cals, employee_events, groups, data_incomplete)

    # 6. Synchronizacja kalendarzy sprzątaczek
    mismatches += sync_employee_calendars(service, groups, notes, extras, chem_texts, employee_cals, today)

    report = {}
    for g in groups.values():
        report.setdefault(f"{g['start_dt']:%Y-%m-%d}", []).append({
            "time": f"{g['start_dt']:%H:%M} - {g['end_dt']:%H:%M}",
            "summary": build_summary(g),
        })

    print(f"\n✨ Gotowe. Dodano: {len(STATS['added'])}, anulowano: {len(STATS['cancelled'])}")
    send_email_report(report, manual, mismatches, run_date)


if __name__ == "__main__":
    sync_cleanings()
