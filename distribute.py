"""
Ręczne rozdzielanie sprzątań między sprzątaczki – dla jednego dnia.

Uruchamiane tylko przez workflow „Rozdzielanie sprzątań” (workflow_dispatch),
nigdy z cyklicznej synchronizacji. Główny kalendarz i sync_cleanings() nie są zmieniane.

Tryb dry-run (domyślny): tylko wypisuje plan, niczego nie zapisuje.
Tryb apply: tworzy wydarzenia w kalendarzach sprzątaczek (insert), nic nie usuwa.

Zasady przypisania (arkusz „Grafik Cleaning”):
- A – adres, D – główna sprzątaczka, E – zastępstwo (dni robocze),
  kolumna „Zastępstwo weekendowe” – dodatkowa zamiana w sobotę i niedzielę.
- Kolumna DATA – dzień; w tym wierszu sprzątaczki mają status z siatki.
- Pierwsza osoba z listy (D, E, a w weekend także F), która „pracuje”, dostaje sprzątanie.
- Sprzątanie, które już ma UID w którymś kalendarzu sprzątaczki, pomijamy.
- Ręczny wpis z tym samym adresem w kalendarzu sprzątaczki na ten dzień – pomijamy i zgłaszamy.
- Nikt z listy nie pracuje – nic nie przypisujemy, zgłaszamy.
"""
import argparse
import difflib
import re
import sys
from datetime import date, datetime, timedelta

from googleapiclient.discovery import build

import apicleaning as ac

GRAFIK_SHEET = "Grafik Cleaning"
UNAVAILABLE_PREFIXES = ("wolne", "pralnia", "moze pracowac", "w przypadku problemów", "w przypadku problemow")


def cell(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()


def norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


NO_CLEAN_WORDS = ("nie sprzątamy", "nie sprzatamy", "nie potrzeba sprzątań", "nie potrzeba sprzatan")


def no_clean_note(g, main_events):
    """Tekst ręcznej uwagi z głównego kalendarza, np. „nie sprzątamy przez remont” – albo None."""
    for ev in main_events or []:
        if not any(u in g["uids"] for u in ac.extract_uids_from_description(ev["description"])):
            continue
        mt = ac.manual_title(ev["summary"]) or ""
        if any(w in mt.lower() for w in NO_CLEAN_WORDS):
            return mt
    return None


def is_extension_only(g):
    """Samo przedłużenie (brak numeru pokoju w nawiasie przy przedłużeniu) – sprzątanie niepotrzebne."""
    ext = g.get("ext_rooms") or set()
    if not ext:
        return False
    plain = [r for r in g.get("rooms", set()) if r and r not in ext]
    return not plain


def is_working(status):
    s = norm(status)
    if not s or s.startswith(UNAVAILABLE_PREFIXES):
        return False
    return s.startswith("pracuje") or s.startswith("od 15:30") or s.startswith("po 15:30")


def read_grafik():
    """Zwraca (addr_rows, day_grid).
    addr_rows: {adres: (główna, zastępstwo, weekend_or_None)}
    day_grid:  {'dd.mm.rrrr': {nazwa_sprzątaczki_z_nagłówka: status}}
    """
    df = ac.read_sheet_by_name(GRAFIK_SHEET, header=None, dtype=str, keep_default_na=False)
    rows = df.values.tolist()

    hdr_i = next((i for i, r in enumerate(rows)
                  if "Główna Pani" in [cell(c) for c in r]), None)
    if hdr_i is None:
        sys.exit("❌ Nie znaleziono wiersza nagłówka (Główna Pani) w „Grafik Cleaning”. "
                 f"Pierwsze wiersze: {rows[:4]}")
    header = [cell(c) for c in rows[hdr_i]]

    def col_exact(name):
        return next(i for i, h in enumerate(header) if h == name)

    i_main = col_exact("Główna Pani")
    i_sub = col_exact("Zastępstwo")
    i_wknd = next((i for i, h in enumerate(header) if h.lower().startswith("zastępstwo weekend")), None)
    i_date = col_exact("DATA")
    i_count = col_exact("Ilość pań w pracy")
    people = {i: h for i, h in enumerate(header) if i > i_count and h}

    addr_rows, day_grid = {}, {}
    for r in rows[hdr_i + 1:]:
        addr = cell(r[0]) if r else ""
        main = cell(r[i_main]) if len(r) > i_main else ""
        if addr and main:
            sub = cell(r[i_sub]) if len(r) > i_sub else ""
            wknd = cell(r[i_wknd]) if (i_wknd is not None and len(r) > i_wknd) else ""
            entry = (main, sub, wknd)
            if addr in addr_rows and addr_rows[addr] != entry:
                print(f"⚠️ Różne dane w arkuszu dla „{addr}”: {addr_rows[addr]} vs {entry} – bierzemy pierwsze")
            addr_rows.setdefault(addr, entry)

        d = cell(r[i_date]) if len(r) > i_date else ""
        if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", d) and d not in day_grid:
            day_grid[d] = {people[i]: cell(r[i]) if i < len(r) else "" for i in people}
    return addr_rows, day_grid


def calendar_for(name, calendars):
    """Nazwa z siatki (np. „Amelia B.”) -> (nazwa kalendarza, id) albo (None, None)."""
    n = norm(name)
    for cal_name, cal_id in calendars.items():
        c = norm(cal_name)
        if c == n or c.startswith(n + " "):
            return cal_name, cal_id
    return None, None


def find_sheet_addr(base, addr_rows):
    hits = [a for a in addr_rows if ac.addr_match(a, base)]
    if hits:
        return hits[0], None
    keys = {ac.addr_key(a): a for a in addr_rows}
    close = difflib.get_close_matches(ac.addr_key(base), list(keys), n=1, cutoff=0.9)
    if close:
        return keys[close[0]], f"dopasowano przybliżenie „{keys[close[0]]}”"
    return None, "adres nie znaleziony w arkuszu"


def manual_entry_for(base, events):
    g_base = base.upper()
    for ev in events:
        if ac.extract_uids_from_description(ev["description"]):
            continue
        place = ac.place_key(ev["summary"])
        if g_base in place or place in g_base:
            return ev
    return None


def run(date_s, mode, only_cleaners):
    target = date.fromisoformat(date_s)
    dry = mode != "apply"
    print(f"{'DRY-RUN' if dry else 'APPLY'} – rozdzielanie na {target:%d.%m.%Y}")
    if only_cleaners:
        print(f"Tylko: {', '.join(only_cleaners)}")

    creds = ac.get_credentials()
    service = build("calendar", "v3", credentials=creds)
    today, _ = ac.scan_window()
    if target < today:
        sys.exit(f"❌ Data {target} jest w przeszłości (dziś {today}).")

    # 1. Kalendarze sprzątaczek – tylko odczyt, musi się udać dla każdego
    calendars = ac.load_employee_calendars()
    if not calendars:
        sys.exit("❌ Brak kalendarzy sprzątaczek w zakładce „ID Cleaning”.")
    employee_events, assigned_uid = {}, {}
    for name, cid in calendars.items():
        evs = ac.get_calendar_events(service, cid, target, target)
        if evs is None:
            sys.exit(f"❌ Nie udało się odczytać kalendarza „{name}” – przerwano, żeby nie zrobić duplikatów.")
        employee_events[name] = evs
        for ev in evs:
            for u in ac.extract_uids_from_description(ev["description"]):
                assigned_uid[u] = name

    # 2. Sprzątania z IdoBooking (ta sama logika co w sync_cleanings)
    apartments, apt_incomplete, apt_notes = ac.load_apartments()
    if not apartments:
        sys.exit("❌ Brak aktywnych jednostek w OBAPI.")
    groups, data_incomplete, build_notes = ac.build_groups(apartments)
    if groups is None:
        sys.exit("❌ Brak danych z API – przerwano.")
    for n in list(apt_notes) + list(build_notes):
        print(f"⚠️ {n}")
    if data_incomplete or apt_incomplete:
        if not dry:
            sys.exit("❌ Niepełne dane z API/OBAPI – zapis przerwany. Uruchom dry-run.")
        print("⚠️ Niepełne dane z API/OBAPI – plan może być niepełny")

    notes, extras, chem_texts, sheet_done = ac.load_notes()
    addresses = sorted({ac.chem_addr(ac.parse_apartment_name(a["name"])[0]) for a in apartments})
    ac.apply_chem_to_groups(ac.chem_month_plan(addresses, today, groups, sheet_done), groups)

    day_groups = [g for k, g in groups.items() if k.split("_")[0] == target.isoformat()]

    # Duplikaty: wydarzenia w głównym kalendarzu, których UID są już przypisane do sprzątaczki
    main_events = ac.get_calendar_events(service, ac.CALENDAR_ID, target, target)
    if main_events is None:
        sys.exit("❌ Nie udało się odczytać głównego kalendarza – przerwano.")
    main_dups = []
    for ev in main_events:
        if any(u in assigned_uid for u in ac.extract_uids_from_description(ev["description"])):
            main_dups.append(ev)
    print(f"Uborek na {target:%d.%m.%Y} w IdoBooking: {len(day_groups)}")

    # 3. Arkusz: adresy i siatka dostępności
    addr_rows, day_grid = read_grafik()
    dd = target.strftime("%d.%m.%Y")
    status = day_grid.get(dd)
    if status is None:
        sys.exit(f"❌ Brak wiersza {dd} w siatce „{GRAFIK_SHEET}”.")
    weekend = target.weekday() >= 5
    print(f"Dzień: {'weekend' if weekend else 'dzień roboczy'}")

    # 4. Przypisanie
    plan, skipped = [], []
    for g in sorted(day_groups, key=lambda x: x["base"]):
        label = f"{g['base']}"
        if any(u in assigned_uid for u in g["uids"]):
            skipped.append(f"⏭️ {label} – już przypisane do: {assigned_uid[next(u for u in g['uids'] if u in assigned_uid)]}")
            continue
        if is_extension_only(g):
            skipped.append(f"➖ {label} – przedłużenie, sprzątanie niepotrzebne")
            continue
        nc = no_clean_note(g, main_events)
        if nc:
            skipped.append(f"🚫 {label} – sprzątanie niepotrzebne: „{nc}”")
            continue
        manual = None
        for evs in employee_events.values():
            manual = manual_entry_for(g["base"], evs)
            if manual:
                break
        if manual:
            skipped.append(f"🔎 {label} – ręczny wpis w kalendarzu: „{manual['summary']}” ({manual['date']}), nie dublujemy")
            continue

        sheet_addr, note = find_sheet_addr(g["base"], addr_rows)
        if not sheet_addr:
            skipped.append(f"❗ {label} – {note}; przypisz ręcznie")
            continue
        main, sub, wknd = addr_rows[sheet_addr]
        cands = [main, sub] + ([wknd] if weekend and wknd else [])

        chosen = None
        for name in cands:
            st = status.get(name)
            if st is None:
                continue
            if is_working(st):
                cal_name, cal_id = calendar_for(name, calendars)
                if not cal_id:
                    skipped.append(f"❗ {label} – „{name}” pracuje, ale nie ma kalendarza w „ID Cleaning”")
                    chosen = "stop"
                    break
                if only_cleaners and cal_name not in only_cleaners and name not in only_cleaners:
                    continue  # filtr --cleaners: szukamy dalej w liście
                chosen = (name, cal_name, cal_id, st)
                break
        if chosen == "stop":
            continue
        if chosen is None:
            details = ", ".join(f"{n}: {status.get(n) or '—'}" for n in cands)
            skipped.append(f"❗ {label} – никто из [{', '.join(cands)}] не работает ({details}); не назначено")
            continue

        name, cal_name, cal_id, st = chosen
        role = "główna" if name == main else ("zastępstwo" if name == sub else "zastępstwo weekendowe")
        plan.append({"g": g, "name": name, "cal_name": cal_name, "cal_id": cal_id,
                     "role": role, "status": st, "note": note})

    # 5. Wynik
    print("\n=== PLAN ===")
    for p in plan:
        extra = f" ({p['note']})" if p["note"] else ""
        print(f"✅ {p['g']['base']} → {p['name']} [{p['role']}, {p['status']}]{extra}")
    if not plan:
        print("(nic do przypisania)")
    print("\n=== POMINIĘTE / DO SPRAWDZENIA ===")
    for s in skipped:
        print(s)

    print("\n=== DUPLIKATY W GŁÓWNYM KALENDARZU ===")
    if not main_dups:
        print("(brak)")
    for ev in main_dups:
        print(f"🗑️ {ev['date']} {ev['summary']} (przypisane do: {assigned_uid[next(u for u in ac.extract_uids_from_description(ev['description']) if u in assigned_uid)]})")

    if dry:
        print("\nDRY-RUN – nic nie zapisano.")
        return

    # 6. Zapis (tylko insert, w kalendarzach sprzątaczek)
    ok, fail = 0, 0
    inserted_uids = set()
    for p in plan:
        body = ac.build_event_body(p["g"], notes, extras, chem_texts)
        main_ev = next((ev for ev in main_events
                        if any(u in p["g"]["uids"] for u in ac.extract_uids_from_description(ev["description"]))), None)
        if main_ev:
            body = ac.prepare_body(body, main_ev)  # przenosi ręczne dopiski z głównego kalendarza
        try:
            service.events().insert(calendarId=p["cal_id"], body=body).execute()
            ok += 1
            inserted_uids.update(p["g"]["uids"])
            print(f"➕ DODANO → {p['name']}: [{target}] {body['summary']}")
        except Exception as e:
            fail += 1
            print(f"❌ Błąd dodawania → {p['name']}: {body['summary']}: {e}")

    # Ręczne dopiski z głównego kalendarza dopisujemy do wydarzenia sprzątaczki (jeśli jeszcze go tam nie ma)
    for ev in main_dups:
        mt = ac.manual_title(ev["summary"])
        if not mt:
            continue
        uid = next((u for u in ac.extract_uids_from_description(ev["description"]) if u in assigned_uid), None)
        if not uid:
            continue
        name = assigned_uid[uid]
        emp = next((e for e in employee_events[name]
                    if uid in ac.extract_uids_from_description(e["description"])), None)
        if emp and mt not in emp["summary"]:
            try:
                service.events().patch(calendarId=calendars[name], eventId=emp["id"],
                                       body={"summary": f"{emp['summary']} {mt}"}).execute()
                print(f"📝 Dopisano do wydarzenia {name}: „{mt}”")
            except Exception as e:
                print(f"❌ Błąd dopisywania notatki do {name}: {e}")

    # Usuwamy duplikaty z głównego kalendarza (tylko te, których UID są przypisane do sprzątaczki)
    removed = 0
    # Dodane przed chwilą (ich UID nie były w kalendarzach przy starcie) też trafiają na listę
    dups_now = list(main_dups)
    for ev in main_events:
        if ev in dups_now:
            continue
        if any(u in inserted_uids for u in ac.extract_uids_from_description(ev["description"])):
            dups_now.append(ev)
    for ev in dups_now:
        try:
            service.events().delete(calendarId=ac.CALENDAR_ID, eventId=ev["id"]).execute()
            removed += 1
            print(f"🗑️ Usunięto duplikat z głównego: {ev['date']} {ev['summary']}")
        except Exception as e:
            print(f"❌ Błąd usuwania duplikatu {ev['summary']}: {e}")
    print(f"Usunięto duplikatów z głównego: {removed}")
    print(f"\nGotowe. Dodano: {ok}, błędy: {fail}")


# ==================== PROPOZYCJA NA 14 DNI (tylko wydarzenie, bez przypisywania) ====================

SUGGEST_TITLE = "Propozycja przydziału personelu sprzątającego"
SUGGEST_OLD_TITLE = "Propozycja przypisania sprzątaczek"
SUGGEST_DAYS = 14
SUGGEST_MARK = "[AUTO: propozycja przypisania – nie edytować]"
ROLE_PL = {"main": "główna", "sub": "zastępstwo", "wknd": "zastępstwo weekendowe"}
SUGGEST_COLOR = "9"  # Blueberry (niebieski) – zawsze


PL_WEEKDAYS = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek", "Sobota", "Niedziela"]
PL_MONTHS = ["stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca", "lipca", "sierpnia",
             "września", "października", "listopada", "grudnia"]


def suggest_lines(target, groups, day_events, addr_rows, day_grid, calendars, main_events):
    """Zwraca (tekst, done). Tekst po polsku: osobno „do decyzji” i „proponowany przydział” (po osobach).
    done=True, gdy nie ma żadnej uborki czekającej na decyzję."""
    dd = target.strftime("%d.%m.%Y")
    day_groups = [g for k, g in groups.items() if k.split("_")[0] == target.isoformat()]
    if not day_groups:
        return "Brak sprzątań tego dnia.", True

    assigned_uid = {}
    for evs in day_events.values():
        for ev in evs:
            for u in ac.extract_uids_from_description(ev["description"]):
                assigned_uid[u] = True

    status = day_grid.get(dd)
    weekend = target.weekday() >= 5
    todo, extensions, no_clean = [], [], []
    for g in sorted(day_groups, key=lambda x: x["base"]):
        if any(u in assigned_uid for u in g["uids"]):
            continue
        if any(manual_entry_for(g["base"], evs) for evs in day_events.values()):
            continue
        if is_extension_only(g):
            extensions.append(g["base"])
            continue
        nc = no_clean_note(g, main_events)
        if nc:
            no_clean.append(f"• {g['base']} – {nc}")
            continue
        todo.append(g)

    if not todo and not extensions and not no_clean:
        return "Wszystkie sprzątania są już przypisane.", True

    decide, proposals = [], {}
    for g in todo:
        sheet_addr, _ = find_sheet_addr(g["base"], addr_rows)
        if not sheet_addr:
            decide.append(f"• {g['base']}")
            continue
        main, sub, wknd = addr_rows[sheet_addr]
        cands = [("main", main), ("sub", sub)] + ([("wknd", wknd)] if weekend and wknd else [])
        chosen = None
        if status:
            for role, name in cands:
                st = status.get(name)
                if st and is_working(st) and calendar_for(name, calendars)[1]:
                    chosen = (role, name)
                    break
        if chosen:
            proposals.setdefault(chosen[1], []).append((g["base"], ROLE_PL[chosen[0]]))
        else:
            decide.append(f"• {g['base']}")

    parts = []
    if no_clean:
        parts += ["", "🚫 SPRZĄTANIE NIEPOTRZEBNE"] + no_clean
    if decide:
        parts += ["", "⚠️ DO DECYZJI (nikt z listy nie pracuje)"] + decide
    if proposals:
        parts += ["", "✅ PROPONOWANY PRZYDZIAŁ"]
        for name in sorted(proposals, key=lambda n: (-len(proposals[n]), n)):
            items = proposals[name]
            parts.append(f"👤 {name} ({len(items)})")
            parts += [f"• {base} – {role}" for base, role in items]
    if status is None and todo:
        parts += ["", "ℹ️ Brak danych o dyspozycyjności w grafiku dla tej daty."]
    if extensions:
        parts += ["", "Przedłużenia (bez sprzątania):"] + [f"• {e}" for e in extensions]
    return "\n".join(parts).strip(), not decide


def upsert_suggestion(service, target, text, done):
    body = {
        "summary": SUGGEST_TITLE,
        "colorId": SUGGEST_COLOR,
        "description": f"{text}\n\n{SUGGEST_MARK}",
        "start": {"dateTime": datetime(target.year, target.month, target.day, 7, 0, tzinfo=ac.POLAND_TZ).isoformat(),
                  "timeZone": "Europe/Warsaw"},
        "end": {"dateTime": datetime(target.year, target.month, target.day, 8, 0, tzinfo=ac.POLAND_TZ).isoformat(),
                "timeZone": "Europe/Warsaw"},
    }
    existing = [e for e in ac.get_calendar_events(service, ac.CALENDAR_ID, target, target) or []
                if e["summary"] in (SUGGEST_TITLE, SUGGEST_OLD_TITLE)]
    if existing:
        service.events().patch(calendarId=ac.CALENDAR_ID, eventId=existing[0]["id"], body=body).execute()
        for extra in existing[1:]:
            service.events().delete(calendarId=ac.CALENDAR_ID, eventId=extra["id"]).execute()
        return "zaktualizowano"
    service.events().insert(calendarId=ac.CALENDAR_ID, body=body).execute()
    return "dodano"


def run_suggest():
    creds = ac.get_credentials()
    service = build("calendar", "v3", credentials=creds)
    today, _ = ac.scan_window()
    end = today + timedelta(days=SUGGEST_DAYS - 1)

    calendars = ac.load_employee_calendars()
    employee_events = {}
    for name, cid in calendars.items():
        evs = ac.get_calendar_events(service, cid, today, end)
        if evs is None:
            sys.exit(f"❌ Nie udało się odczytać kalendarza „{name}” – propozycji nie aktualizuję.")
        employee_events[name] = evs

    apartments, apt_incomplete, _ = ac.load_apartments()
    groups, data_incomplete, _ = ac.build_groups(apartments) if apartments else (None, True, [])
    if groups is None or data_incomplete or apt_incomplete:
        sys.exit("❌ Niepełne dane z API/OBAPI – propozycji nie aktualizuję.")

    addr_rows, day_grid = read_grafik()
    for i in range(SUGGEST_DAYS):
        target = today + timedelta(days=i)
        day_events = {n: [e for e in evs if e["date"] == target.isoformat()] for n, evs in employee_events.items()}
        main_events = ac.get_calendar_events(service, ac.CALENDAR_ID, target, target) or []
        text, done = suggest_lines(target, groups, day_events, addr_rows, day_grid, calendars, main_events)
        result = upsert_suggestion(service, target, text, done)
        print(f"{target:%d.%m.%Y}: {result}\n{text}\n")
    print("Gotowe.")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--date", default="", help="YYYY-MM-DD (dla dry-run/apply)")
    p.add_argument("--mode", choices=["dry-run", "apply", "suggest"], default="dry-run")
    p.add_argument("--cleaners", default="", help="opcjonalnie: tylko te panie, po przecinku")
    a = p.parse_args()
    if a.mode == "suggest":
        run_suggest()
        return
    if not a.date:
        sys.exit("❌ Podaj --date dla trybu dry-run/apply.")
    only = [x.strip() for x in a.cleaners.split(",") if x.strip()]
    run(a.date, a.mode, only)


if __name__ == "__main__":
    main()
