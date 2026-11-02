import hashlib
import json
import os
import re
import time
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv

load_dotenv()

POLAND_TZ = ZoneInfo("Europe/Warsaw")
IDO_DOMAIN = os.environ.get("IDO_DOMAIN", "")
IDO_LOGIN = os.environ.get("IDO_LOGIN", "")
IDO_KEY = os.environ.get("IDO_KEY", "")

API_VERSION = "36"
PER_PAGE = 100
PAGE_PAUSE_SEC = 0.2

# Równoległe pobieranie stron. Jeśli pojawią się błędy HTTP (429/5xx), zmniejsz do 2.
WORKERS = 4

# Cache nazwisk klientów (plik obok skryptu). Odświeżany po TTL lub gdy brakuje clientId.
CLIENTS_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "client_names_cache.json")
CLIENTS_CACHE_TTL = 3600  # sekundy

# Prefiks UID zapisywany w opisach wydarzeń. Sprawdź w istniejącej uborce linię "UID:".
# Jeśli UID to "reservation5063666758", zostaw "reservation". Jeśli to sama liczba, ustaw "".
UID_PREFIX = "reservation"
LOCK_UID_PREFIX = "lock"

# Jedna sesja HTTP: ponownie używa połączeń, każde zapytanie jest szybsze.
_session = requests.Session()


def _local(tag):
    return tag.split("}")[-1]


def _leaf(el):
    """Pola będące tekstem/liczbami (bez zagnieżdżonych struktur)."""
    return {_local(c.tag): c.text for c in el if len(c) == 0}


def _fault(root):
    for el in root.iter():
        if _local(el.tag) == "faultString" and el.text:
            return el.text
    return None


def _system_key():
    date_str = datetime.now(POLAND_TZ).strftime("%Y%m%d")
    inner = hashlib.sha1(IDO_KEY.encode("utf-8")).hexdigest()
    return hashlib.sha1((date_str + inner).encode("utf-8")).hexdigest()


def _soap(endpoint, method, inner_xml, lang="eng"):
    if not (IDO_DOMAIN and IDO_LOGIN and IDO_KEY):
        print("❌ IdoBooking: brak IDO_DOMAIN / IDO_LOGIN / IDO_KEY w .env")
        return None
    auth = (
        "<authenticate>"
        f"<systemKey>{_system_key()}</systemKey>"
        f"<systemLogin>{IDO_LOGIN}</systemLogin>"
        f"<lang>{lang}</lang>"
        "</authenticate>"
    )
    envelope = (
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:ns="ApiObjects"><soapenv:Body>'
        f"<ns:{method}><request>{auth}{inner_xml}</request></ns:{method}>"
        "</soapenv:Body></soapenv:Envelope>"
    )
    url = f"https://{IDO_DOMAIN}.idobooking.com/api/{endpoint}"
    headers = {"Content-Type": "text/xml; charset=utf-8", "SOAPAction": '""'}
    try:
        r = _session.post(url, data=envelope.encode("utf-8"), headers=headers, timeout=60)
    except requests.RequestException as e:
        print(f"❌ IdoBooking: błąd połączenia ({endpoint}): {e}")
        return None
    if r.status_code != 200:
        print(f"❌ IdoBooking: HTTP {r.status_code} ({endpoint})")
        return None
    return r.text


def _std_inner(page, params_xml):
    """Zapytanie w formacie reservations/get i availabilitylocks/get: result + paramsSearch."""
    return (f"<result><page>{page}</page><countOnPage>{PER_PAGE}</countOnPage></result>"
            f"<paramsSearch>{params_xml}</paramsSearch>")


def _locks_inner(page, params_xml):
    """Zapytanie dla availabilitylocks/get z uwzględnieniem zakresu dat."""
    return (f"<result><page>{page}</page><countOnPage>500</countOnPage></result>"
            f"<paramsSearch>{params_xml}</paramsSearch>")


def _fetch_page(endpoint, method, params_xml, build_inner, page):
    """Jedna strona: (root, pageAll) albo None przy błędzie."""
    xml_text = _soap(endpoint, method, build_inner(page, params_xml))
    if xml_text is None:
        return None
    root = ET.fromstring(xml_text)
    err = _fault(root)
    if err:
        print(f"❌ IdoBooking ({endpoint}, strona {page}): {err}")
        return None
    info = next((_leaf(el) for el in root.iter() if _local(el.tag) == "result"), {})
    return root, int(info.get("pageAll") or 1)


def _paged_get(endpoint, method, params_xml, parse, build_inner=_std_inner):
    """
    Pobiera wszystkie strony metody: pierwsza strona dla pageAll, reszta równolegle.
    parse(root) zwraca listę słowników z kluczem 'id'.
    build_inner(page, params_xml) buduje treść zapytania (różne metody mają różną strukturę).
    Zwraca listę albo None przy błędzie.
    """
    first = _fetch_page(endpoint, method, params_xml, build_inner, 1)
    if first is None:
        return None
    roots, page_all = [first[0]], first[1]

    if page_all > 1:
        pages = range(2, page_all + 1)
        with ThreadPoolExecutor(max_workers=WORKERS) as ex:
            results = list(ex.map(
                lambda p: _fetch_page(endpoint, method, params_xml, build_inner, p), pages))
        if any(r is None for r in results):
            return None
        roots += [r[0] for r in results]

    out, seen = [], set()
    for root in roots:  # kolejność stron zachowana
        for rec in parse(root):
            if rec["id"] not in seen:
                seen.add(rec["id"])
                out.append(rec)
    return out


def _day(text):
    """'2026-04-18 17:00' -> date(2026, 4, 18); None gdy brak lub błąd."""
    if not text:
        return None
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


# ==================== GOŚCIE (IMIĘ I NAZWISKO) ====================

def _clients_inner(page, params_xml):
    return (f"<settings><page>{page}</page><number>{PER_PAGE}</number></settings>"
            f"<params>{params_xml}</params>")


def _parse_clients(root):
    parents = {c: p for p in root.iter() for c in p}
    out = []
    for item in root.iter():
        if _local(item.tag) != "item":
            continue
        parent = parents.get(item)
        if parent is None or _local(parent.tag) != "clients":
            continue

        fields = _leaf(item)
        cid = (fields.get("id") or "").strip()
        if not cid:
            continue

        guest_name = ""
        guests = next((c for c in item if _local(c.tag) == "guests"), None)
        if guests is not None:
            for g in guests:
                if _local(g.tag) != "item":
                    continue
                gf = _leaf(g)
                guest_name = " ".join(x.strip() for x in (gf.get("firstName"), gf.get("lastName"))
                                      if x and x.strip())
                if guest_name:
                    break

        if not guest_name:
            guest_name = " ".join(x.strip() for x in (fields.get("firstName"), fields.get("lastName"))
                                  if x and x.strip())
        out.append({"id": cid, "name": guest_name})
    return out


def _load_clients_cache(needed_ids):
    try:
        if time.time() - os.path.getmtime(CLIENTS_CACHE) >= CLIENTS_CACHE_TTL:
            return None
        with open(CLIENTS_CACHE, encoding="utf-8") as f:
            cache = json.load(f)
    except (OSError, ValueError):
        return None
    if needed_ids and any(i not in cache for i in needed_ids):
        return None
    return cache


def fetch_client_names(needed_ids=None):
    cache = _load_clients_cache(needed_ids)
    if cache is not None:
        names = {k: v for k, v in cache.items() if v}
        print(f"👤 IdoBooking: nazwiska z cache: {len(names)} (z {len(cache)})")
        return names

    recs = _paged_get(f"clients/get/{API_VERSION}/soap", "get", "", _parse_clients,
                      build_inner=_clients_inner)
    if recs is None:
        return None

    cache = {r["id"]: r["name"] for r in recs}
    try:
        with open(CLIENTS_CACHE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except OSError as e:
        print(f"⚠️ Nie udało się zapisać cache klientów: {e}")

    names = {k: v for k, v in cache.items() if v}
    print(f"👤 IdoBooking: klientów z nazwiskiem: {len(names)} (z {len(recs)})")
    return names


# ==================== REZERWACJE ====================

def _parse_reservations(root):
    res_list = next((el for el in root.iter() if _local(el.tag) == "reservations"), None)
    if res_list is None:
        return []
    out = []
    for item in res_list:
        if _local(item.tag) != "item":
            continue
        rec = {"id": None, "details": {}, "items": []}
        for el in item:
            name = _local(el.tag)
            if name == "id":
                rec["id"] = el.text
            elif name == "reservationDetails":
                rec["details"] = _leaf(el)
            elif name == "items":
                rec["items"] = [_leaf(it) for it in el if _local(it.tag) == "item"]
        out.append(rec)
    return out


def fetch_all_reservations():
    res = _paged_get(f"reservations/get/{API_VERSION}/soap", "get", "", _parse_reservations)
    if res is not None:
        print(f"🛎️ IdoBooking: pobrano rezerwacji: {len(res)}")
    return res


def stays_by_item(reservations, client_names=None):
    client_names = client_names or {}
    by_item = {}
    for r in reservations:
        d = r["details"]
        if d.get("status") == "canceled":
            continue
        start, end = _day(d.get("dateFrom")), _day(d.get("dateTo"))
        if not start or not end:
            continue
        ext = (d.get("externalReservationId") or "").strip()
        short = ext if ext and ext != "None" else f"id{r['id']}"
        uid = re.sub(r"\s+", "", UID_PREFIX + short).lower()
        cid = (d.get("clientId") or "").strip()
        name = client_names.get(cid, "")
        for it in r["items"]:
            item_id = (it.get("objectItemId") or "").strip()
            if not item_id:
                continue
            by_item.setdefault(item_id, []).append({
                "uid": uid, "uid_short": short, "blocked": False,
                "name": name, "manual": [], "client": cid,
                "start": start, "end": end,
            })
    return by_item


# ==================== BLOKADY ====================

def _parse_locks(root):
    locks = []
    for item in root.findall(".//lock") or root.findall(".//item"):
        item_id = (item.findtext("objectItemId") or item.findtext("itemId") or "").strip()
        start = (item.findtext("dateFrom") or item.findtext("start") or "").strip()
        end = (item.findtext("dateTo") or item.findtext("end") or "").strip()
        if not (item_id and start and end):
            continue
        start = start.split(" ")[0]
        end = end.split(" ")[0]
        lock_id = (item.findtext("id") or item.findtext("lockId")
                   or f"{item_id}-{start}-{end}").strip()
        note = (item.findtext("note") or item.findtext("notatka")
                or item.findtext("description") or item.findtext("comment")
                or "Blokada dostępności").strip()
        locks.append({"id": lock_id, "item_id": item_id, "start": start, "end": end, "note": note})
    return locks


def fetch_availability_locks():
    """
    Pobiera blokady dostępności z IdoBooking: od (dziś - 15 dni) do (dziś + 45 dni).
    Zakres dzielony na kawałki po 14 dni, bo API zwraca najwyżej 100 rekordów na zapytanie.
    Jeśli któryś kawałek zwróci 100 rekordów, uznajemy dane za niepełne i zwracamy None
    (kalendarz bez zmian w tym przebiegu).
    """
    today = datetime.now(POLAND_TZ).date()
    range_start = today - timedelta(days=15)
    range_end = today + timedelta(days=45)

    all_locks, seen = [], set()
    chunk_start = range_start
    while chunk_start <= range_end:
        chunk_end = min(chunk_start + timedelta(days=13), range_end)
        date_from = chunk_start.strftime("%Y-%m-%d 00:00:00")
        date_to = chunk_end.strftime("%Y-%m-%d 23:59:59")
        params_xml = f"<dateFrom>{date_from}</dateFrom><dateTo>{date_to}</dateTo>"

        part = _paged_get(f"availabilitylocks/get/{API_VERSION}/soap", "get", params_xml,
                          _parse_locks, build_inner=_locks_inner)
        if part is None:
            return None  # None = błąd API. NIE zamieniamy na []
        if len(part) >= 100:
            print(f"⚠️ Blokady {date_from[:10]}-{date_to[:10]}: {len(part)} rekordów (limit), dane niepełne")
            return None
        for lk in part:
            if lk["id"] not in seen:
                seen.add(lk["id"])
                all_locks.append(lk)
        chunk_start = chunk_end + timedelta(days=1)

    print(f"🔒 IdoBooking: pobrano blokad: {len(all_locks)}")
    return all_locks


def locks_by_item(locks):
    by_item = {}
    for lk in locks or []:
        item_id = lk["item_id"]
        by_item.setdefault(item_id, []).append({
            "uid": f"lock{lk['id']}".lower(),
            "blocked": True,
            "name": f"🔒 BLOKADA: {lk.get('note', '')}",
            "note": lk.get("note", "Blokada dostępności"),
            "start": datetime.strptime(lk["start"], "%Y-%m-%d").date(),
            "end": datetime.strptime(lk["end"], "%Y-%m-%d").date(),
        })
    return by_item


# ==================== OBIEKTY I JEDNOSTKI ====================

def _get_name(object_id):
    inner = f"<paramsSearch><objectId>{object_id}</objectId></paramsSearch>"
    xml_text = _soap(f"objects/getDescriptions/{API_VERSION}/soap", "getDescriptions", inner, lang="pol")
    if not xml_text:
        return ""
    for el in ET.fromstring(xml_text).iter():
        if _local(el.tag) == "name" and el.text:
            return el.text.strip()
    return ""


def _get_items(object_id):
    inner = f"<paramsSearch><objectId>{object_id}</objectId></paramsSearch>"
    xml_text = _soap(f"objects/getItems/{API_VERSION}/soap", "getItems", inner, lang="pol")
    if not xml_text:
        return []
    root = ET.fromstring(xml_text)
    parents = {child: parent for parent in root.iter() for child in parent}
    out = []
    for el in root.iter():
        if _local(el.tag) == "objectItemId" and el.text:
            parent = parents.get(el)
            code = ""
            if parent is not None:
                code = next((c.text.strip() for c in parent
                             if _local(c.tag) == "objectItemCodeInternal" and c.text), "")
            out.append((el.text.strip(), code))
    return out


def _unit_rows(oid):
    name = _get_name(oid)
    return [{"objectItemId": item_id, "objectId": oid, "name": name, "code": code}
            for item_id, code in _get_items(oid)]


def fetch_all_units():
    today = datetime.now(POLAND_TZ).date()
    inner = (f"<dateFrom>{today.isoformat()}</dateFrom>"
             f"<dateTo>{(today + timedelta(days=2)).isoformat()}</dateTo>")
    xml_text = _soap(f"objects/getAvailabilityByDateRange/{API_VERSION}/soap",
                     "getAvailabilityByDateRange", inner)
    if xml_text is None:
        return None
    root = ET.fromstring(xml_text)
    if _fault(root):
        return None

    ids = sorted({el.text.strip() for el in root.iter()
                  if _local(el.tag) == "objectId" and el.text}, key=int)
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        chunks = list(ex.map(_unit_rows, ids))
    return [u for chunk in chunks for u in chunk]