#!/usr/bin/env python3
"""
Hotel Price Scraper — Google Sheets Config Edition
----------------------------------------------------
Reads config from "Config - Hotels" and "Config - Dates" tabs.
Writes results to:
  1. The group's own tab (e.g. "La Clef Group")
  2. "All Data" tab — single source for Looker Studio
  3. "Scrape Failures" tab — logs every hotel/date that returned no price

Usage: python scraper.py <group_id>

GitHub Secrets required:
  GOOGLE_SERVICE_ACCOUNT  - service account JSON key
  GOOGLE_SHEET_ID         - master Google Sheet ID
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

ALL_DATA_TAB  = "All Data"
FAILURE_TAB   = "Scrape Failures"

# ── Args ──────────────────────────────────────────────────

if len(sys.argv) < 2:
    print("Usage: python scraper.py <group_id>")
    sys.exit(1)

GROUP_ID = sys.argv[1]


# ── Google Sheets client ──────────────────────────────────

def get_spreadsheet():
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
    return gspread.authorize(creds).open_by_key(sheet_id)


# ── Load config from sheet ────────────────────────────────

def load_config(spreadsheet):
    try:
        hotels_ws = spreadsheet.worksheet("Config - Hotels")
    except gspread.exceptions.WorksheetNotFound:
        print("❌  'Config - Hotels' tab not found.")
        sys.exit(1)

    hotels_data = hotels_ws.get_all_records()
    if not hotels_data:
        print("❌  'Config - Hotels' tab is empty.")
        sys.exit(1)

    group_rows = [
        r for r in hotels_data
        if str(r.get("group_id", "")).strip().lower() == GROUP_ID.lower()
        and str(r.get("hotel_name", "")).strip()
    ]
    if not group_rows:
        available = list({r["group_id"] for r in hotels_data})
        print(f"❌  Group '{GROUP_ID}' not found. Available: {available}")
        sys.exit(1)

    sheet_tab = group_rows[0].get("sheet_tab", GROUP_ID).strip()
    hotels = [
        {
            "name":       str(r["hotel_name"]).strip(),
            "currency":   str(r.get("currency", "SGD")).strip().upper(),
            "is_primary": str(r.get("is_primary", "")).strip().lower()
                          in ("yes", "true", "1", "primary"),
        }
        for r in group_rows
    ]

    try:
        dates_ws = spreadsheet.worksheet("Config - Dates")
    except gspread.exceptions.WorksheetNotFound:
        print("❌  'Config - Dates' tab not found.")
        sys.exit(1)

    dates_data = dates_ws.get_all_records()
    date_ranges = [
        (str(r["check_in"]).strip(), str(r["check_out"]).strip())
        for r in dates_data
        if str(r.get("check_in", "")).strip() and str(r.get("check_out", "")).strip()
    ]

    print(f"  📋  {len(hotels)} hotels · {len(date_ranges)} dates loaded from sheet")
    return hotels, date_ranges, sheet_tab


# ── Data models ───────────────────────────────────────────

@dataclass
class RoomResult:
    scraped_date:       str
    group:              str
    is_primary:         str
    hotel:              str
    check_in:           str
    check_out:          str
    day_type:           str
    nights:             int
    room_type:          str
    price_per_night:    str
    total_price:        str
    currency:           str
    free_cancellation:  str
    breakfast_included: str
    url:                str


@dataclass
class FailureLog:
    logged_date:  str   # date this failure was recorded
    group:        str
    is_primary:   str
    hotel:        str
    check_in:     str
    check_out:    str
    day_type:     str
    reason:       str   # "No search result" | "Room table empty" | "Exception: ..."
    search_url:   str


RESULT_HEADERS  = [f.name.replace("_", " ").title() for f in fields(RoomResult)]
FAILURE_HEADERS = [f.name.replace("_", " ").title() for f in fields(FailureLog)]


def classify_day_type(check_in: str) -> str:
    d = datetime.strptime(check_in, "%Y-%m-%d")
    if d.weekday() == 4: return "Shoulder (Fri-Sat)"
    if d.weekday() == 5: return "Weekend (Sat-Sun)"
    return "Weekday (Tue-Wed)"


# ── Sheet write helpers ───────────────────────────────────

def get_or_create_tab(spreadsheet, tab_name: str, headers: list[str]):
    try:
        ws = spreadsheet.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=tab_name, rows=100000, cols=25)

    if not ws.row_values(1):
        ws.append_row(headers, value_input_option="RAW")
        ws.format("1:1", {
            "textFormat": {
                "bold": True,
                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
            },
            "backgroundColor": {"red": 0.122, "green": 0.22, "blue": 0.392},
        })
        print(f"  📋  Created tab: '{tab_name}'")
    return ws


def results_to_rows(results: list[RoomResult]) -> list[list]:
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
    return rows


def failures_to_rows(failures: list[FailureLog]) -> list[list]:
    col_names = [f.name for f in fields(FailureLog)]
    return [[getattr(f, n) for n in col_names] for f in failures]


def append_results(spreadsheet, results: list[RoomResult],
                   failures: list[FailureLog], group_tab: str) -> None:
    # 1. Group tab
    if results:
        ws_group = get_or_create_tab(spreadsheet, group_tab, RESULT_HEADERS)
        ws_group.append_rows(results_to_rows(results), value_input_option="USER_ENTERED")
        print(f"  📊  {len(results)} rows → '{group_tab}'")

    # 2. All Data tab
    if results:
        ws_all = get_or_create_tab(spreadsheet, ALL_DATA_TAB, RESULT_HEADERS)
        ws_all.append_rows(results_to_rows(results), value_input_option="USER_ENTERED")
        print(f"  📊  {len(results)} rows → '{ALL_DATA_TAB}'")

    # 3. Scrape Failures tab
    if failures:
        ws_fail = get_or_create_tab(spreadsheet, FAILURE_TAB, FAILURE_HEADERS)
        # Style failure header red instead of navy
        if ws_fail.row_values(1) == FAILURE_HEADERS:
            ws_fail.format("1:1", {
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                },
                "backgroundColor": {"red": 0.75, "green": 0.15, "blue": 0.15},
            })
        ws_fail.append_rows(failures_to_rows(failures), value_input_option="USER_ENTERED")
        print(f"  ⚠️   {len(failures)} failures → '{FAILURE_TAB}'")
    else:
        print(f"  ✅  No failures for this group")


# ── Playwright helpers ────────────────────────────────────

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


async def get_hotel_url(page, hotel, check_in, check_out, currency, adults, rooms):
    q = hotel.replace(" ", "+")
    url = (
        f"https://www.booking.com/search.html?ss={q}"
        f"&checkin={check_in}&checkout={check_out}"
        f"&group_adults={adults}&no_rooms={rooms}"
        f"&selected_currency={currency}&lang=en-gb&order=popularity"
    )
    await page.goto(url, wait_until="domcontentloaded", timeout=35000)
    await page.wait_for_timeout(3500)
    await dismiss_overlays(page)

    card = page.locator('[data-testid="property-card"]').first
    try:
        await card.wait_for(timeout=8000)
    except Exception:
        return None, url   # return search url for failure log

    href = await card.locator('[data-testid="title-link"]').first.get_attribute("href")
    if not href:
        return None, url
    if not href.startswith("http"):
        href = "https://www.booking.com" + href
    return href.split("?")[0], url


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
            "room_type":         current_name,
            "total_val":         total_val,
            "free_cancellation": free_cancel,
            "breakfast":         breakfast,
        })

    rooms.sort(key=lambda r: r["total_val"])
    return rooms


async def fetch_hotel(page, hotel_name, currency, is_primary,
                      check_in, check_out, sheet_tab, adults, rooms):
    ci     = datetime.strptime(check_in, "%Y-%m-%d")
    co     = datetime.strptime(check_out, "%Y-%m-%d")
    nights = (co - ci).days
    label  = "PRIMARY" if is_primary else "compset"
    day_type = classify_day_type(check_in)
    primary_str = "Primary" if is_primary else "Compset"

    def make_failure(reason: str, search_url: str = "") -> FailureLog:
        return FailureLog(
            logged_date = date.today().isoformat(),
            group       = sheet_tab,
            is_primary  = primary_str,
            hotel       = hotel_name,
            check_in    = check_in,
            check_out   = check_out,
            day_type    = day_type,
            reason      = reason,
            search_url  = search_url,
        )

    print(f"  🔍  [{label}] {hotel_name}  |  {check_in}  |  {currency}")
    try:
        hotel_url, search_url = await get_hotel_url(
            page, hotel_name, check_in, check_out, currency, adults, rooms
        )
        if not hotel_url:
            print(f"    ⚠️  No search result found")
            return None, make_failure("No search result", search_url)

        direct_url = (
            f"{hotel_url}?checkin={check_in}&checkout={check_out}"
            f"&group_adults={adults}&no_rooms={rooms}"
            f"&selected_currency={currency}&lang=en-gb"
        )
        await page.goto(direct_url, wait_until="domcontentloaded", timeout=35000)
        await page.wait_for_timeout(3500)
        await dismiss_overlays(page)

        room_list = await scrape_room_table(page)
        if not room_list:
            print(f"    ⚠️  Room table empty")
            return None, make_failure("Room table empty — hotel page loaded but no rooms shown", direct_url)

        cheapest  = room_list[0]
        total     = cheapest["total_val"]
        per_night = round(total / nights, 0)

        print(f"    ✅  {cheapest['room_type'][:35]}  |  {currency} {per_night:.0f}/night")
        return RoomResult(
            scraped_date       = date.today().isoformat(),
            group              = sheet_tab,
            is_primary         = primary_str,
            hotel              = hotel_name,
            check_in           = check_in,
            check_out          = check_out,
            day_type           = day_type,
            nights             = nights,
            room_type          = cheapest["room_type"],
            price_per_night    = f"{per_night:.0f}",
            total_price        = f"{total:.0f}",
            currency           = currency,
            free_cancellation  = cheapest["free_cancellation"],
            breakfast_included = cheapest["breakfast"],
            url                = direct_url,
        ), None

    except Exception as e:
        print(f"    ❌  Exception: {e}")
        return None, make_failure(f"Exception: {str(e)[:120]}", "")


# ── Main ──────────────────────────────────────────────────

async def main():
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("playwright not installed.")
        sys.exit(1)

    print(f"\n🏨  Hotel Price Scraper — {date.today()}  |  Group: {GROUP_ID}")

    spreadsheet = get_spreadsheet()
    hotels, date_ranges, sheet_tab = load_config(spreadsheet)

    adults = 2
    rooms  = 1
    total  = len(hotels) * len(date_ranges)
    print(f"    {len(hotels)} hotels × {len(date_ranges)} dates = {total} combinations\n")

    results:  list[RoomResult]  = []
    failures: list[FailureLog]  = []

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

        for hotel in hotels:
            for check_in, check_out in date_ranges:
                result, failure = await fetch_hotel(
                    page,
                    hotel["name"],
                    hotel["currency"],
                    hotel["is_primary"],
                    check_in,
                    check_out,
                    sheet_tab,
                    adults,
                    rooms,
                )
                if result:
                    results.append(result)
                if failure:
                    failures.append(failure)
                await asyncio.sleep(3)

        await browser.close()

    total_attempted = len(results) + len(failures)
    print(f"\n💾  {len(results)} succeeded / {len(failures)} failed out of {total_attempted}")
    append_results(spreadsheet, results, failures, sheet_tab)
    print(f"✅  Group '{GROUP_ID}' done.")


if __name__ == "__main__":
    asyncio.run(main())
