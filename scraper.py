#!/usr/bin/env python3
"""
Hotel Price Scraper — Google Sheets Edition
--------------------------------------------
Reads config from CONFIG_JSON secret.
Hotels can each have their own currency.
Appends results to a Google Sheet daily.

GitHub Secrets required:
  CONFIG_JSON             - your hotels/dates config (see config.template.json)
  GOOGLE_SERVICE_ACCOUNT  - full contents of your service account JSON key file
  GOOGLE_SHEET_ID         - the ID from your Google Sheet URL
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


# ── Load config ───────────────────────────────────────────

def load_config() -> dict:
    raw = os.environ.get("CONFIG_JSON", "").strip()
    if not raw:
        print("❌  CONFIG_JSON environment variable is not set.")
        sys.exit(1)
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"❌  CONFIG_JSON is not valid JSON: {e}")
        sys.exit(1)


CONFIG      = load_config()
DATE_RANGES = [tuple(d) for d in CONFIG["date_ranges"]]
ADULTS      = CONFIG.get("adults", 2)
ROOMS       = CONFIG.get("rooms", 1)

# Hotels is now a list of dicts: {"name": "...", "currency": "..."}
# Support both old format (list of strings) and new format (list of dicts)
raw_hotels = CONFIG["hotels"]
if raw_hotels and isinstance(raw_hotels[0], str):
    # Legacy format — all use same currency
    default_currency = CONFIG.get("currency", "SGD")
    HOTELS = [{"name": h, "currency": default_currency} for h in raw_hotels]
else:
    HOTELS = raw_hotels


# ── Data model ────────────────────────────────────────────

@dataclass
class RoomResult:
    scraped_date:       str
    hotel:              str
    check_in:           str
    check_out:          str
    nights:             int
    room_type:          str
    price_per_night:    str
    total_price:        str
    currency:           str
    free_cancellation:  str
    breakfast_included: str
    url:                str


HEADERS = [f.name.replace("_", " ").title() for f in fields(RoomResult)]


# ── Google Sheets helpers ─────────────────────────────────

def get_sheet():
    sa_json  = os.environ.get("GOOGLE_SERVICE_ACCOUNT", "")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID", "")

    if not sa_json:
        print("❌  GOOGLE_SERVICE_ACCOUNT secret is not set.")
        sys.exit(1)
    if not sheet_id:
        print("❌  GOOGLE_SHEET_ID secret is not set.")
        sys.exit(1)

    sa_dict = json.loads(sa_json)
    creds   = Credentials.from_service_account_info(
        sa_dict,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    client = gspread.authorize(creds)
    sheet  = client.open_by_key(sheet_id)

    try:
        ws = sheet.worksheet("Hotel Prices")
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title="Hotel Prices", rows=10000, cols=20)

    if ws.row_count == 0 or not ws.row_values(1):
        ws.append_row(HEADERS, value_input_option="RAW")
        ws.format("1:1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.122, "green": 0.22, "blue": 0.392},
        })

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
    print(f"  📊  Appended {len(rows)} rows to Google Sheet")


# ── Playwright helpers ────────────────────────────────────

async def dismiss_overlays(page) -> None:
    for selector in [
        'button[id*="onetrust-accept"]',
        'button[aria-label*="Dismiss"]',
        'button[aria-label*="Close"]',
        '[data-testid="user-account-ui-cta-close"]',
        'button.modal-mask-closeBtn',
    ]:
        try:
            btn = page.locator(selector).first
            if await btn.is_visible(timeout=1500):
                await btn.click()
                await page.wait_for_timeout(600)
        except Exception:
            pass


async def get_hotel_url(page, hotel: str, check_in: str, check_out: str, currency: str) -> Optional[str]:
    q = hotel.replace(" ", "+")
    search_url = (
        f"https://www.booking.com/search.html?ss={q}"
        f"&checkin={check_in}&checkout={check_out}"
        f"&group_adults={ADULTS}&no_rooms={ROOMS}"
        f"&selected_currency={currency}&lang=en-gb&order=popularity"
    )
    await page.goto(search_url, wait_until="domcontentloaded", timeout=35000)
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
    current_room_name = "N/A"

    for row in rows:
        name_el = await row.query_selector(".hprt-roomtype-icon-link, .room-name")
        if name_el:
            t = (await name_el.inner_text()).strip()
            if t:
                current_room_name = t

        price_el = await row.query_selector(
            ".bui-price-display__value, "
            ".prco-valign-middle-helper, "
            "[data-testid='price-and-discounted-price']"
        )
        if not price_el:
            continue

        price_raw = (await price_el.inner_text()).strip()
        nums = re.findall(r"[\d,]+(?:\.\d+)?", price_raw.replace(",", ""))
        if not nums:
            continue
        try:
            total_val = float(nums[-1])
        except ValueError:
            continue

        row_text   = (await row.inner_text()).lower()
        cond_el    = await row.query_selector(".hprt-conditions, [data-testid='cancellation-and-charges']")
        check_cond = (await cond_el.inner_text()).lower() if cond_el else row_text

        free_cancel = "Unknown"
        if "free cancellation" in check_cond:
            free_cancel = "Yes"
        elif "non-refundable" in check_cond or "no refund" in check_cond:
            free_cancel = "No"

        breakfast = "Unknown"
        if "breakfast included" in row_text or "includes breakfast" in row_text:
            breakfast = "Yes"
        elif "room only" in row_text or "no breakfast" in row_text or "without breakfast" in row_text:
            breakfast = "No"
        elif "breakfast" in row_text:
            breakfast = "Yes"

        rooms.append({
            "room_type":         current_room_name,
            "total_val":         total_val,
            "free_cancellation": free_cancel,
            "breakfast":         breakfast,
        })

    rooms.sort(key=lambda r: r["total_val"])
    return rooms


async def fetch_hotel(page, hotel_name: str, currency: str, check_in: str, check_out: str) -> Optional[RoomResult]:
    ci     = datetime.strptime(check_in, "%Y-%m-%d")
    co     = datetime.strptime(check_out, "%Y-%m-%d")
    nights = (co - ci).days

    print(f"  🔍  {hotel_name}  |  {check_in} → {check_out}  |  {currency}")
    try:
        hotel_url = await get_hotel_url(page, hotel_name, check_in, check_out, currency)
        if not hotel_url:
            print(f"    ⚠️  No search result found")
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

        print(f"    ✅  {cheapest['room_type'][:40]}  |  {currency} {per_night:.0f}/night")
        return RoomResult(
            scraped_date       = date.today().isoformat(),
            hotel              = hotel_name,
            check_in           = check_in,
            check_out          = check_out,
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

    total_combos = len(HOTELS) * len(DATE_RANGES)
    print(f"\n🏨  Hotel Price Scraper — {date.today()}")
    print(f"    {len(HOTELS)} hotels × {len(DATE_RANGES)} date ranges = {total_combos} combinations\n")

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
            hotel_name = hotel["name"]
            currency   = hotel["currency"]
            for check_in, check_out in DATE_RANGES:
                result = await fetch_hotel(page, hotel_name, currency, check_in, check_out)
                if result:
                    results.append(result)
                await asyncio.sleep(2)

        await browser.close()

    if not results:
        print("\n⚠️  No results scraped.")
        sys.exit(0)

    print(f"\n💾  Writing {len(results)} rows to Google Sheets…")
    ws = get_sheet()
    append_to_sheet(ws, results)
    print(f"\n✅  Done.")


if __name__ == "__main__":
    asyncio.run(main())
