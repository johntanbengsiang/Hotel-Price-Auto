#!/usr/bin/env python3
"""
Hotel Price Scraper — Retry Failed Scrapes
-------------------------------------------
Reads the "Scrape Failures" tab, retries every entry,
removes resolved rows, updates persistent failures with
an incremented retry count.

Runs as a separate GitHub Actions workflow a few hours
after the main scraper.

GitHub Secrets required (same as main scraper):
  GOOGLE_SERVICE_ACCOUNT
  GOOGLE_SHEET_ID
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

ALL_DATA_TAB = "All Data"
FAILURE_TAB  = "Scrape Failures"

ADULTS = 2
ROOMS  = 1


# ── Google Sheets client ──────────────────────────────────

def get_spreadsheet():
    sa_json  = os.environ.get("GOOGLE_SERVICE_ACCOUNT", "")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID", "")
    if not sa_json or not sheet_id:
        print("❌  Missing GOOGLE_SERVICE_ACCOUNT or GOOGLE_SHEET_ID.")
        sys.exit(1)
    creds = Credentials.from_service_account_info(
        json.loads(sa_json),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return gspread.authorize(creds).open_by_key(sheet_id)


# ── Read failures ─────────────────────────────────────────

@dataclass
class FailureRow:
    row_index:    int    # 1-based sheet row number (for deletion)
    logged_date:  str
    group:        str
    is_primary:   str
    hotel:        str
    check_in:     str
    check_out:    str
    day_type:     str
    reason:       str
    search_url:   str
    retry_count:  int    # how many times this has been retried


def load_failures(spreadsheet) -> tuple[gspread.Worksheet, list[FailureRow]]:
    try:
        ws = spreadsheet.worksheet(FAILURE_TAB)
    except gspread.exceptions.WorksheetNotFound:
        print("✅  No Scrape Failures tab found — nothing to retry.")
        sys.exit(0)

    all_rows = ws.get_all_values()
    if len(all_rows) <= 1:
        print("✅  Scrape Failures tab is empty — nothing to retry.")
        sys.exit(0)

    headers = [h.lower().replace(" ", "_") for h in all_rows[0]]
    failures = []

    for i, row in enumerate(all_rows[1:], start=2):  # row 2 onwards (1-based)
        if not any(row):
            continue
        d = dict(zip(headers, row))
        failures.append(FailureRow(
            row_index   = i,
            logged_date = d.get("logged_date", ""),
            group       = d.get("group", ""),
            is_primary  = d.get("is_primary", ""),
            hotel       = d.get("hotel", ""),
            check_in    = d.get("check_in", "").strip()[:10],
            check_out   = d.get("check_out", "").strip()[:10],
            day_type    = d.get("day_type", ""),
            reason      = d.get("reason", ""),
            search_url  = d.get("search_url", ""),
            retry_count = int(d.get("retry_count", 0) or 0),
        ))

    print(f"  📋  {len(failures)} failures to retry")
    return ws, failures


# ── Data model ────────────────────────────────────────────

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


RESULT_HEADERS = [f.name.replace("_", " ").title() for f in fields(RoomResult)]


def classify_day_type(check_in: str) -> str:
    d = datetime.strptime(check_in, "%Y-%m-%d")
    if d.weekday() == 4: return "Shoulder (Fri-Sat)"
    if d.weekday() == 5: return "Weekend (Sat-Sun)"
    return "Weekday (Tue-Wed)"


def get_currency_for_hotel(spreadsheet, hotel_name: str) -> str:
    """Look up currency from Config - Hotels tab."""
    try:
        ws = spreadsheet.worksheet("Config - Hotels")
        for row in ws.get_all_records():
            if str(row.get("hotel_name", "")).strip() == hotel_name:
                return str(row.get("currency", "SGD")).strip().upper()
    except Exception:
        pass
    return "SGD"


# ── Sheet write helpers ───────────────────────────────────

def get_or_create_tab(spreadsheet, tab_name: str, headers: list[str]):
    try:
        ws = spreadsheet.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=tab_name, rows=5000, cols=20)

    if not ws.row_values(1):
        ws.append_row(headers, value_input_option="RAW")
        ws.format("1:1", {
            "textFormat": {
                "bold": True,
                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
            },
            "backgroundColor": {"red": 0.122, "green": 0.22, "blue": 0.392},
        })
    return ws


def write_result(spreadsheet, result: RoomResult) -> None:
    col_names = [f.name for f in fields(RoomResult)]
    row = []
    for name in col_names:
        val = getattr(result, name)
        if name in ("price_per_night", "total_price", "nights"):
            try:
                val = int(val)
            except (ValueError, TypeError):
                pass
        row.append(val)

    # Write to All Data tab
    ws_all = get_or_create_tab(spreadsheet, ALL_DATA_TAB, RESULT_HEADERS)
    ws_all.append_row(row, value_input_option="USER_ENTERED")

    # Write to group tab
    ws_group = get_or_create_tab(spreadsheet, result.group, RESULT_HEADERS)
    ws_group.append_row(row, value_input_option="USER_ENTERED")

    print(f"    ✅  Written to '{result.group}' and '{ALL_DATA_TAB}'")


def update_failures_tab(ws_failures: gspread.Worksheet,
                        resolved_indices: list[int],
                        still_failing: list[tuple[int, str]]) -> None:
    """
    resolved_indices : 1-based row numbers to delete
    still_failing    : list of (row_index, new_reason) to update
    """
    # Update reasons and increment retry count for persistent failures
    for row_idx, new_reason in still_failing:
        # Column 8 = reason (1-based), column 10 = retry_count
        ws_failures.update_cell(row_idx, 8, new_reason)
        # Get current retry count and increment
        try:
            current = int(ws_failures.cell(row_idx, 10).value or 0)
        except (ValueError, TypeError):
            current = 0
        ws_failures.update_cell(row_idx, 10, current + 1)

    # Delete resolved rows in reverse order so indices stay valid
    for row_idx in sorted(resolved_indices, reverse=True):
        ws_failures.delete_rows(row_idx)
        print(f"    🗑️   Removed resolved failure (row {row_idx})")


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


async def get_hotel_url(page, hotel, check_in, check_out, currency):
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
        return None, search_url

    href = await card.locator('[data-testid="title-link"]').first.get_attribute("href")
    if not href:
        return None, search_url
    if not href.startswith("http"):
        href = "https://www.booking.com" + href
    return href.split("?")[0], search_url


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


async def retry_one(page, failure: FailureRow, currency: str
                    ) -> tuple[Optional[RoomResult], Optional[str]]:
    """Returns (RoomResult, None) on success or (None, reason_string) on failure."""
    ci     = datetime.strptime(failure.check_in, "%Y-%m-%d")
    co     = datetime.strptime(failure.check_out, "%Y-%m-%d")
    nights = (co - ci).days

    print(f"  🔄  {failure.hotel}  |  {failure.check_in}  |  {currency}")
    try:
        hotel_url, search_url = await get_hotel_url(
            page, failure.hotel, failure.check_in, failure.check_out, currency
        )
        if not hotel_url:
            reason = "No search result"
            print(f"    ⚠️  {reason}")
            return None, reason

        direct_url = (
            f"{hotel_url}?checkin={failure.check_in}&checkout={failure.check_out}"
            f"&group_adults={ADULTS}&no_rooms={ROOMS}"
            f"&selected_currency={currency}&lang=en-gb"
        )
        await page.goto(direct_url, wait_until="domcontentloaded", timeout=35000)
        await page.wait_for_timeout(3500)
        await dismiss_overlays(page)

        room_list = await scrape_room_table(page)
        if not room_list:
            reason = "Room table empty — hotel page loaded but no rooms shown"
            print(f"    ⚠️  {reason}")
            return None, reason

        cheapest  = room_list[0]
        total     = cheapest["total_val"]
        per_night = round(total / nights, 0)

        print(f"    ✅  {cheapest['room_type'][:35]}  |  {currency} {per_night:.0f}/night")
        return RoomResult(
            scraped_date       = date.today().isoformat(),
            group              = failure.group,
            is_primary         = failure.is_primary,
            hotel              = failure.hotel,
            check_in           = failure.check_in,
            check_out          = failure.check_out,
            day_type           = classify_day_type(failure.check_in),
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
        reason = f"Exception: {str(e)[:120]}"
        print(f"    ❌  {reason}")
        return None, reason


# ── Main ──────────────────────────────────────────────────

async def main():
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("playwright not installed.")
        sys.exit(1)

    print(f"\n🔄  Retry Scraper — {date.today()}")

    spreadsheet = get_spreadsheet()
    ws_failures, failures = load_failures(spreadsheet)

    # Pre-load all currencies from Config - Hotels once
    currency_map: dict[str, str] = {}
    try:
        hotels_ws = spreadsheet.worksheet("Config - Hotels")
        for row in hotels_ws.get_all_records():
            name = str(row.get("hotel_name", "")).strip()
            curr = str(row.get("currency", "SGD")).strip().upper()
            if name:
                currency_map[name] = curr
    except Exception as e:
        print(f"  ⚠️  Could not load currency map: {e}")

    resolved_indices: list[int]         = []
    still_failing:    list[tuple[int, str]] = []

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

        for failure in failures:
            currency = currency_map.get(failure.hotel, "SGD")
            result, reason = await retry_one(page, failure, currency)

            if result:
                write_result(spreadsheet, result)
                resolved_indices.append(failure.row_index)
            else:
                still_failing.append((failure.row_index, reason))

            await asyncio.sleep(3)

        await browser.close()

    resolved = len(resolved_indices)
    still    = len(still_failing)
    print(f"\n📊  Retry complete: {resolved} resolved · {still} still failing")

    if resolved_indices or still_failing:
        print("  🗂️   Updating Scrape Failures tab…")
        update_failures_tab(ws_failures, resolved_indices, still_failing)

    print("✅  Done.")


if __name__ == "__main__":
    asyncio.run(main())
