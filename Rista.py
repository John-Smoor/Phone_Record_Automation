
import time
import jwt
import requests
import pandas as pd
from datetime import datetime, timedelta
import os
import sys

# ── API credentials ───────────────────────────────────────────────────────────

API_KEY    = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# ── Paths ─────────────────────────────────────────────────────────────────────
DATA_FILE   = r"C:/Users/jeryy/OneDrive - BLISS CHOCOLATES INDIA PRIVATE LIMITED/Documents/Phone No/Data/sales_rolling.csv"
BRANCH_FILE = "data/branch_codes.csv"
SEED_CSV    = r"C:/Users/jeryy/OneDrive - BLISS CHOCOLATES INDIA PRIVATE LIMITED/Documents/Phone No/Data/sales_rollings.csv"



from datetime import datetime, timedelta

ROLLING_DAYS = 10

YESTERDAY = (datetime.now() - timedelta(days=1)).date()
START_DATE = YESTERDAY - timedelta(days=ROLLING_DAYS - 1)
# ═════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def generate_jwt_token():
    payload = {"iss": API_KEY, "iat": int(time.time())}
    return jwt.encode(payload, SECRET_KEY, algorithm="HS256")


def daterange(start, end):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


# ═════════════════════════════════════════════════════════════════════════════
#  API FETCH
# ═════════════════════════════════════════════════════════════════════════════

def fetch_sales_for_branch(branch, day_str):
    url     = "https://api.ristaapps.com/v1/sales/page"
    headers = {
        "x-api-key":    API_KEY,
        "x-api-token":  generate_jwt_token(),
        "content-type": "application/json"
    }
    params   = {"branch": branch, "day": day_str}
    all_data = []
    page     = 1

    while True:
        print(f"  [{branch}] [{day_str}] page {page}...")
        r = requests.get(url, headers=headers, params=params)
        if r.status_code != 200:
            print(f"  Failed {r.status_code}: {r.text}")
            break
        rj = r.json()
        all_data.extend(rj.get("data", []))
        last_key = rj.get("lastKey")
        if not last_key:
            break
        params["lastKey"] = last_key
        page += 1
        time.sleep(1)

    return all_data


def extract_phone_number(record):
    checks = [
        record.get("customer", {}).get("phoneNumber"),
        record.get("customer", {}).get("phone"),
        record.get("customer", {}).get("mobile"),
        record.get("phoneNumber"),
        record.get("phone"),
        record.get("delivery", {}).get("phoneNumber"),
        record.get("delivery_address", {}).get("phoneNumber"),
        record.get("contactNumber"),
    ]
    for phone in checks:
        if phone:
            return str(phone).strip()
    return None


def extract_item_rows(record):
    phone   = extract_phone_number(record)
    invoice = {
        "branchName":                   record.get("branchName"),
        "branchCode":                   record.get("branchCode"),
        "invoiceNumber":                record.get("invoiceNumber"),
        "invoiceDate":                  record.get("invoiceDate"),
        "createdDate":                  record.get("createdDate"),
        "chargeTaxTotal":               record.get("chargeTaxTotal"),
        "channel":                      record.get("channel"),
        "itemCount":                    record.get("itemCount"),
        "itemTotalAmount":              record.get("itemTotalAmount"),
        "totalDiscountAmount":          record.get("totalDiscountAmount"),
        "grossAmount":                  record.get("grossAmount"),
        "netDiscountAmount":            record.get("netDiscountAmount"),
        "netAmount":                    record.get("netAmount"),
        "status":                       record.get("status"),
        "delivery_mode":                record.get("delivery_mode"),
        "delivery_address_addressLine": record.get("delivery_address", {}).get("addressLine"),
        "delivery_address_zip":         record.get("delivery_address", {}).get("zip"),
        "delivery_address_longitude":   record.get("delivery_address", {}).get("longitude"),
        "delivery_address_latitude":    record.get("delivery_address", {}).get("latitude"),
        "customer_name":                record.get("customer", {}).get("name"),
        "customer_email":               record.get("customer", {}).get("email"),
        "customer_phoneNumber":         phone,
        "customer_id":                  record.get("customer", {}).get("id"),
        "orderNumber":                  record.get("orderNumber"),
        "branch":                       record.get("branch"),
        "date":                         record.get("date"),
    }
    rows = []
    for item in record.get("items", []):
        row = invoice.copy()
        row.update({
            "item_shortName":         item.get("shortName"),
            "item_longName":          item.get("longName"),
            "item_skuCode":           item.get("skuCode"),
            "item_categoryName":      item.get("categoryName"),
            "item_subCategoryName":   item.get("subCategoryName"),
            "item_quantity":          item.get("quantity"),
            "item_unitPrice":         item.get("unitPrice"),
            "item_measuringUnit":     item.get("measuringUnit"),
            "item_discountAmount":    item.get("discountAmount"),
            "item_grossAmount":       item.get("grossAmount"),
            "item_netDiscountAmount": item.get("netDiscountAmount"),
            "item_netAmount":         item.get("netAmount"),
            "item_taxAmount":         item.get("taxAmount"),
            "item_itemNumber":        item.get("itemNumber"),
            "item_itemAmount":        item.get("itemAmount"),
            "item_itemTotalAmount":   item.get("itemTotalAmount"),
            "item_optionAmount":      item.get("optionAmount"),
            "item_baseGrossAmount":   item.get("baseGrossAmount"),
            "item_baseNetAmount":     item.get("baseNetAmount"),
            "item_baseTaxAmount":     item.get("baseTaxAmount"),
            "item_createdTime":       item.get("createdTime"),
            "item_kotNumber":         item.get("kotNumber"),
            "item_kotStatus":         item.get("kotStatus"),
        })
        taxes = item.get("taxes", [])
        if taxes:
            t = taxes[0]
            row.update({
                "item_tax_name":           t.get("name"),
                "item_tax_percentage":     t.get("percentage"),
                "item_tax_amountIncluded": t.get("amountIncluded"),
                "item_tax_amountExcluded": t.get("amountExcluded"),
            })
        rows.append(row)
    return rows


def scrape_dates(branches, dates):
    new_rows = []
    for branch in branches:
        print(f"\n=== Branch: {branch} ===")
        for day in dates:
            records = fetch_sales_for_branch(branch, day.strftime("%Y-%m-%d"))
            for record in records:
                record["branch"] = branch
                record["date"]   = day.strftime("%Y-%m-%d")
                new_rows.extend(extract_item_rows(record))
    return new_rows


# ═════════════════════════════════════════════════════════════════════════════
#  ROLLING WINDOW MANAGER
#  Always keeps exactly 21 days: adds newest day, removes oldest day
# ═════════════════════════════════════════════════════════════════════════════

def update_rolling_csv(new_rows):
    """
    1. Load existing sales_rolling.csv  (or seed from historical file)
    2. Append yesterday's fresh rows
    3. Deduplicate on invoiceNumber + item_itemNumber
    4. Drop any rows older than 21 days (the oldest day falls off)
    5. Save back to sales_rolling.csv
    """

    # ── Step 1: Load existing data ────────────────────────────────────────────
    if os.path.exists(DATA_FILE):
        existing = pd.read_csv(DATA_FILE, dtype=str, low_memory=False)
        print(f"\n  Loaded existing rolling CSV  →  {len(existing):,} rows")
        print(f"  Date range before update     →  {existing['date'].min()}  to  {existing['date'].max()}")
    elif os.path.exists(SEED_CSV):
        existing = pd.read_csv(SEED_CSV, dtype=str, low_memory=False)
        print(f"\n  No rolling CSV found — seeding from historical file  →  {len(existing):,} rows")
    else:
        existing = pd.DataFrame()
        print("\n  No existing file — starting fresh")

    # ── Step 2: Append yesterday's new data ───────────────────────────────────
    if new_rows:
        yesterday_df = pd.DataFrame(new_rows)
        print(f"  Yesterday's new rows fetched →  {len(yesterday_df):,} rows  ({YESTERDAY})")
        combined = pd.concat([existing, yesterday_df], ignore_index=True)
    else:
        print(f"  No new rows fetched for {YESTERDAY} — keeping existing data")
        combined = existing.copy()

    # ── Step 3: Deduplicate (safe to re-run without doubling data) ────────────
    before_dedup = len(combined)
    if "item_itemNumber" in combined.columns:
        combined = combined.drop_duplicates(
            subset=["invoiceNumber", "item_itemNumber"], keep="last"
        )
    else:
        combined = combined.drop_duplicates(subset=["invoiceNumber"], keep="last")
    dupes_removed = before_dedup - len(combined)
    if dupes_removed:
        print(f"  Duplicates removed           →  {dupes_removed:,} rows")

    # ── Step 4: Check window — only drop oldest day if we exceed 21 days ────────
    combined["date"] = pd.to_datetime(combined["date"], errors="coerce").dt.date

    distinct_days = sorted(combined["date"].dropna().unique())
    num_days      = len(distinct_days)

    print(f"  Distinct days in data        →  {num_days}  ({distinct_days[0]}  to  {distinct_days[-1]})")

    if num_days <= ROLLING_DAYS:
        print(f"  Window check                 →  {num_days} days ≤ {ROLLING_DAYS} — nothing to drop ✓")
    else:
        # More than 21 days — drop the oldest day only
        oldest_day   = distinct_days[0]
        before_trim  = len(combined)
        combined     = combined[combined["date"] > oldest_day].reset_index(drop=True)
        rows_dropped = before_trim - len(combined)
        print(f"  Window check                 →  {num_days} days > {ROLLING_DAYS} — dropping oldest day")
        print(f"  Oldest day removed           →  {oldest_day}  ({rows_dropped:,} rows dropped)")

    # ── Step 5: Save updated rolling CSV ─────────────────────────────────────
    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    combined.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

    print(f"\n✅  Rolling CSV updated")
    print(f"   Total rows saved  →  {len(combined):,}")
    print(f"   Date range now    →  {combined['date'].min()}  to  {combined['date'].max()}")
    print(f"   File              →  {DATA_FILE}")


# ═════════════════════════════════════════════════════════════════════════════
#  MODE 1 — BACKFILL  (run once to build the initial 21-day CSV from scratch)
# ═════════════════════════════════════════════════════════════════════════════

def run_backfill():
    print(f"\n{'='*60}")
    print(f"  BACKFILL  |  {START_DATE}  →  {YESTERDAY}  ({ROLLING_DAYS} days)")
    print(f"{'='*60}")

    branches = pd.read_csv(BRANCH_FILE)["branch"].dropna().unique()
    dates    = list(daterange(START_DATE, YESTERDAY))
    rows     = scrape_dates(branches, dates)

    if not rows:
        print("No data fetched.")
        return

    df = pd.DataFrame(rows)
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    df = df.drop_duplicates(subset=["invoiceNumber", "item_itemNumber"], keep="last")
    df = df[df["date"] >= START_DATE].reset_index(drop=True)

    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

    print(f"\n✅  Backfill complete")
    print(f"   Total rows  →  {len(df):,}")
    print(f"   Date range  →  {df['date'].min()}  to  {df['date'].max()}")
    print(f"   File        →  {DATA_FILE}")


# ═════════════════════════════════════════════════════════════════════════════
#  MODE 2 — DAILY  (runs every night)
#
#  What happens each run:
#    1. Fetch yesterday's data from Rista API (all branches)
#    2. Append it to sales_rolling.csv
#    3. Remove the oldest day's rows so window stays at 21 days
#    4. Save back to sales_rolling.csv
# ═════════════════════════════════════════════════════════════════════════════

def run_daily():
    print(f"\n{'='*60}")
    print(f"  DAILY UPDATE  |  Fetching data for  {YESTERDAY}")
    print(f"{'='*60}")

    branches = pd.read_csv(BRANCH_FILE)["branch"].dropna().unique()
    print(f"  Branches to scrape  →  {len(branches)}")

    # Fetch yesterday from API
    new_rows = scrape_dates(branches, [YESTERDAY])

    # Add to rolling CSV, drop oldest day
    update_rolling_csv(new_rows)


# ═════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
#
#  python sales_daily.py            →  daily mode  (add yesterday, drop oldest)
#  python sales_daily.py backfill   →  backfill mode  (build 21 days from scratch)
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "daily"
    if mode == "backfill":
        run_backfill()
    else:
        run_daily()
        
