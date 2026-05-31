#!/usr/bin/env python3
"""
Hotel Price Scraper — Parallel Groups Edition
----------------------------------------------
Run once per hotel group (parallel GitHub Actions matrix jobs).
Each job scrapes one group and writes to its own worksheet tab.

Usage: python scraper.py <group_id>
  e.g. python scraper.py la_clef

GitHub Secrets required:
  CONFIG_JSON             - full config with groups and date_ranges
  GOOGLE_SERVICE_ACCOUNT  - service account JSON key
  GOOGLE_SHEET_ID         - Google Sheet ID
"""

import asyncio
import json
import os
import re
import sys
from dataclasses import dataclass, fields
from datetime import date, datetime
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials


# ── Args & config ─────────────────────────────────────────

if len(sys.argv) < 2:
    print("Usage: python scraper.py <group_id>")
    sys.exit(1)

GROUP_ID = sys.argv[1]


def load_config() -> dict:
    raw = os.environ.get("CONFIG_JSON", "").strip()
    if not raw:
        print("❌  CONFIG_JSON not set.")
        sys.exit(1)
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"❌  CONFIG_JSON invalid JSON: {e}")
        sys.exit(1)


CONFIG      = load_config()
DATE_RANGES = [tuple(d) for d in CONFIG["date_ranges"]]
ADULTS      = CONFIG.get("adults", 2)
ROOMS       = CONFIG.get("rooms", 1)

# Find the group matching GROUP_ID
GROUP = next((g for g in CONFIG["groups"] if g["id"] == GROUP_ID), None)
if not GROUP:
    print(f"❌  Group '{GROUP_ID}' not found in config. "
          f"Available: {[g['id'] for g in CONFIG['groups']]}")
    sys.exit(1)

HOTELS    = GROUP["hotels"]
SHEET_TAB = GROUP["sheet_tab"]

print(f"\n🏨  Group: {SHEET_TAB}")
print(f"    {len(HOTELS)} hotels × {len(DATE_RANGES)} dates = "
      f"{len(HOTELS) * len(DATE_RANGES)} combinations")


# ── Data model ────────────────────────────────────────────

@dataclass
class RoomResult:
    scraped_date:       str
    group:              str
    is_primary:         str   # "Primary" or "Compset"
    hotel:              str
    check_in:           str
    check_out:          str
    day_type:           str   # Weekday / Shoulder / Weekend
    nights:             int
    room_type:          str
    price_per_night:    str
    total_price:        str
    currency:           str
    free_cancellation:  str
    breakfast_included: str
    url:                str


HEADERS = [f.name.replace("_", " ").title() for f in fields(RoomResult)]


def classify_day_type(check_in: str) -> str:
    """Classify a check-in date as Weekday, Shoulder (Fri), or Weekend (Sat)."""
    d = datetime.strptime(check_in, "%Y-%m-%d")
    if d.weekday() == 4:   # Friday
        return "Shoulder (Fri-Sat)"
    elif d.weekday() == 5:  # Saturday
        return "Weekend (Sat-Sun)"
    else:
        return "Weekday (Tue-Wed)"


# ── Google Sheets ─────────────────────────────────────────

def get_sheet():
    sa_json  = os.environ.get("GOOGLE_SERVICE_ACCOUNT", "")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID", "")
    if not sa_json:
        print("❌  GOOGLE_SERVICE_ACCOUNT not set.")
        sys.exit(1)
    if not sheet_id:
        print("❌  GOOGLE_SHEET_ID not set.")
        sys.exit(1)

    creds = Credentials.from_service_account_info(
        json.loads(sa_json),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    client    = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    # Get or create the tab for this group
    try:
        ws = spreadsheet.worksheet(SHEET_TAB)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=SHEET_TAB, rows=50000, cols=20)

    # Add header if sheet is empty
    existing = ws.row_values(1) if ws.row_count > 0 else []
    if not existing:
        ws.append_row(HEADERS, value_input_option="RAW")
        ws.format("1:1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.122, "green": 0.22, "blue": 0.392},
        })
        print(f"  📋  Created worksheet tab: {SHEET_TAB}")

    return ws


def append_to_sheet(ws, results: list[RoomResult]) -> None:
    col_names = [f.name for f in fields(RoomResult)]
    rows = []
    for r in results:
        row = []
        for name in col_names:
            val = getattr(r, name)
            if name in ("price_per_night", "total_price", "nights"):
                try:
                    val = int(val)
                except (ValueError, TypeError):
                    pass
            row.append(val)
        rows.append(row)
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  📊  Appended {len(rows)} rows → '{SHEET_TAB}'")


# ── Playwright ────────────────────────────────────────────

async def dismiss_overlays(page) -> None:
    for sel in [
        'button[id*="onetrust-accept"]',
        'button[aria-label*="Dismiss"]',
        'button[aria-label*="Close"]',
        '[data-testid="user-account-ui-cta-close"]',
        'button.modal-mask-closeBtn',
    ]:
        try:
            btn = page.locator(sel).first
            if await btn.is_visible(timeout=1500):
                await btn.click()
                await page.wait_for_timeout(600)
        except Exception:
            pass


async def get_hotel_url(page, hotel: str, check_in: str, check_out: str, currency: str) -> Optional[str]:
    q = hotel.replace(" ", "+")
    url = (
        f"https://www.booking.com/search.html?ss={q}"
        f"&checkin={check_in}&checkout={check_out}"
        f"&group_adults={ADULTS}&no_rooms={ROOMS}"
        f"&selected_currency={currency}&lang=en-gb&order=popularity"
    )
    await page.goto(url, wait_until="domcontentloaded", timeout=35000)
    await page.wait_for_timeout(3500)
    await dismiss_overlays(page)

    card = page.locator('[data-testid="property-card"]').first
    try:
        await card.wait_for(timeout=8000)
    except Exception:
        return None

    href = await card.locator('[data-testid="title-link"]').first.get_attribute("href")
    if not href:
        return None
    if not href.startswith("http"):
        href = "https://www.booking.com" + href
    return href.split("?")[0]


async def scrape_room_table(page) -> list[dict]:
    rooms = []
    try:
        await page.wait_for_selector("#hprt-table, .hprt-table", timeout=12000)
    except Exception:
        return rooms

    rows = await page.query_selector_all(".hprt-table tr, #hprt-table tr")
    current_name = "N/A"

    for row in rows:
        name_el = await row.query_selector(".hprt-roomtype-icon-link, .room-name")
        if name_el:
            t = (await name_el.inner_text()).strip()
            if t:
                current_name = t

        price_el = await row.query_selector(
            ".bui-price-display__value, .prco-valign-middle-helper, "
            "[data-testid='price-and-discounted-price']"
        )
        if not price_el:
            continue

        nums = re.findall(r"[\d,]+(?:\.\d+)?",
                          (await price_el.inner_text()).replace(",", ""))
        if not nums:
            continue
        try:
            total_val = float(nums[-1])
        except ValueError:
            continue

        row_text   = (await row.inner_text()).lower()
        cond_el    = await row.query_selector(
            ".hprt-conditions, [data-testid='cancellation-and-charges']"
        )
        check_cond = (await cond_el.inner_text()).lower() if cond_el else row_text

        free_cancel = "Unknown"
        if "free cancellation" in check_cond:
            free_cancel = "Yes"
        elif "non-refundable" in check_cond or "no refund" in check_cond:
            free_cancel = "No"

        breakfast = "Unknown"
        if "breakfast included" in row_text or "includes breakfast" in row_text:
            breakfast = "Yes"
        elif any(x in row_text for x in ["room only", "no breakfast", "without breakfast"]):
            breakfast = "No"
        elif "breakfast" in row_text:
            breakfast = "Yes"

        rooms.append({
            "room_type": current_name,
            "total_val": total_val,
            "free_cancellation": free_cancel,
            "breakfast": breakfast,
        })

    rooms.sort(key=lambda r: r["total_val"])
    return rooms


async def fetch_hotel(
    page, hotel_name: str, currency: str, is_primary: bool,
    check_in: str, check_out: str
) -> Optional[RoomResult]:
    ci     = datetime.strptime(check_in, "%Y-%m-%d")
    co     = datetime.strptime(check_out, "%Y-%m-%d")
    nights = (co - ci).days
    label  = "PRIMARY" if is_primary else "compset"

    print(f"  🔍  [{label}] {hotel_name}  |  {check_in}  |  {currency}")
    try:
        hotel_url = await get_hotel_url(page, hotel_name, check_in, check_out, currency)
        if not hotel_url:
            print(f"    ⚠️  No result found")
            return None

        direct_url = (
            f"{hotel_url}?checkin={check_in}&checkout={check_out}"
            f"&group_adults={ADULTS}&no_rooms={ROOMS}"
            f"&selected_currency={currency}&lang=en-gb"
        )
        await page.goto(direct_url, wait_until="domcontentloaded", timeout=35000)
        await page.wait_for_timeout(3500)
        await dismiss_overlays(page)

        rooms = await scrape_room_table(page)
        if not rooms:
            print(f"    ⚠️  Room table empty")
            return None

        cheapest  = rooms[0]
        total     = cheapest["total_val"]
        per_night = round(total / nights, 0)

        print(f"    ✅  {cheapest['room_type'][:35]}  |  {currency} {per_night:.0f}/night")
        return RoomResult(
            scraped_date       = date.today().isoformat(),
            group              = SHEET_TAB,
            is_primary         = "Primary" if is_primary else "Compset",
            hotel              = hotel_name,
            check_in           = check_in,
            check_out          = check_out,
            day_type           = classify_day_type(check_in),
            nights             = nights,
            room_type          = cheapest["room_type"],
            price_per_night    = f"{per_night:.0f}",
            total_price        = f"{total:.0f}",
            currency           = currency,
            free_cancellation  = cheapest["free_cancellation"],
            breakfast_included = cheapest["breakfast"],
            url                = direct_url,
        )
    except Exception as e:
        print(f"    ❌  Error: {e}")
        return None


# ── Main ──────────────────────────────────────────────────

async def main():
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("playwright not installed.")
        sys.exit(1)

    results: list[RoomResult] = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="en-GB",
            viewport={"width": 1400, "height": 900},
            extra_http_headers={"Accept-Language": "en-GB,en;q=0.9"},
        )
        page = await context.new_page()

        for hotel in HOTELS:
            for check_in, check_out in DATE_RANGES:
                result = await fetch_hotel(
                    page,
                    hotel["name"],
                    hotel["currency"],
                    hotel.get("is_primary", False),
                    check_in,
                    check_out,
                )
                if result:
                    results.append(result)
                await asyncio.sleep(3)  # slightly longer delay for larger runs

        await browser.close()

    if not results:
        print("\n⚠️  No results scraped.")
        sys.exit(0)

    print(f"\n💾  Writing {len(results)} rows…")
    ws = get_sheet()
    append_to_sheet(ws, results)
    print(f"✅  Group '{GROUP_ID}' done.")


if __name__ == "__main__":
    asyncio.run(main())
