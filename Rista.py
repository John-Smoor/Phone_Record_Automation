# # import time
# # import jwt  # PyJWT
# # import requests
# # import json
# # from pprint import pprint

# # # API credentials
# # API_KEY = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
# # SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# # def generate_jwt_token(api_key, secret_key):
# #     token_creation_time = int(time.time())
# #     payload = {
# #         "iss": api_key,
# #         "iat": token_creation_time
# #     }
# #     token = jwt.encode(payload, secret_key, algorithm="HS256")
# #     return token

# # def fetch_all_sales(branch='KOR', day='2025-09-09'):
# #     url = "https://api.ristaapps.com/v1/sales/page"
# #     jwt_token = generate_jwt_token(API_KEY, SECRET_KEY)
    
# #     headers = {
# #         'x-api-key': API_KEY,
# #         'x-api-token': jwt_token,
# #         'content-type': 'application/json'
# #     }
    
# #     all_data = []
# #     params = {
# #         'branch': branch,
# #         'day': day
# #     }

# #     iteration = 1
# #     while True:
# #         print(f"Fetching page {iteration}...")
# #         response = requests.get(url, headers=headers, params=params)
        
# #         if response.status_code != 200:
# #             print(f"Failed with status code {response.status_code}: {response.text}")
# #             break
        
# #         resp_json = response.json()
# #         page_data = resp_json.get('data', [])
# #         all_data.extend(page_data)

# #         last_key = resp_json.get('lastKey')
# #         if not last_key:
# #             print("No lastKey found. Finished fetching all pages.")
# #             break

# #         # Set lastKey for next request
# #         params['lastKey'] = last_key
# #         iteration += 1
# #         time.sleep(1)  # Avoid hitting rate limits

# #     # Save all fetched data into one JSON file
# #     with open('all_sales_data.json', 'w') as f:
# #         json.dump(all_data, f, indent=4)

# #     print(f"Fetched total records: {len(all_data)}")

# # if __name__ == "__main__":
# #     fetch_all_sales()

# # import time
# # import jwt
# # import requests
# # import json
# # import pandas as pd
# # from pandas import json_normalize
# # from datetime import datetime, timedelta

# # # API credentials
# # API_KEY = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
# # SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# # def generate_jwt_token(api_key, secret_key):
# #     payload = {
# #         "iss": api_key,
# #         "iat": int(time.time())
# #     }
# #     return jwt.encode(payload, secret_key, algorithm="HS256")

# # def fetch_sales_for_branch(branch, day):
# #     url = "https://api.ristaapps.com/v1/sales/page"
# #     jwt_token = generate_jwt_token(API_KEY, SECRET_KEY)

# #     headers = {
# #         'x-api-key': API_KEY,
# #         'x-api-token': jwt_token,
# #         'content-type': 'application/json'
# #     }

# #     params = {"branch": branch, "day": day}
# #     all_data = []
# #     iteration = 1

# #     while True:
# #         print(f"[{branch}] [{day}] Fetching page {iteration}...")
# #         response = requests.get(url, headers=headers, params=params)

# #         if response.status_code != 200:
# #             print(f"[{branch}] [{day}] Failed: {response.status_code} → {response.text}")
# #             break

# #         resp_json = response.json()
# #         page_data = resp_json.get("data", [])
# #         all_data.extend(page_data)

# #         last_key = resp_json.get("lastKey")
# #         if not last_key:
# #             print(f"[{branch}] [{day}] Completed fetching all pages.")
# #             break

# #         params["lastKey"] = last_key
# #         iteration += 1
# #         time.sleep(1)

# #     return all_data

# # def daterange(start_date, end_date):
# #     """Yield all dates from start_date to end_date inclusive."""
# #     current = start_date
# #     while current <= end_date:
# #         yield current
# #         current += timedelta(days=1)

# # def save_all_branches(start_date="2025-11-01", end_date="2025-11-02"):
# #     # Load branch codes
# #     df_branches = pd.read_csv("branch_codes.csv")
# #     branches = df_branches["branch"].dropna().unique()

# #     # Convert input strings to date objects
# #     start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
# #     end_date = datetime.strptime(end_date, "%Y-%m-%d").date()

# #     all_records = []

# #     for branch in branches:
# #         print(f"\n========== Processing branch: {branch} ==========")

# #         for day in daterange(start_date, end_date):
# #             day_str = day.strftime("%Y-%m-%d")

# #             print(f"\n--- Fetching for date: {day_str} ---")

# #             branch_data = fetch_sales_for_branch(branch, day_str)

# #             # Add metadata into each record
# #             for record in branch_data:
# #                 record["branch"] = branch
# #                 record["date"] = day_str

# #             all_records.extend(branch_data)

# #     print(f"\nTotal records fetched: {len(all_records)}")

# #     if not all_records:
# #         print("No data fetched. CSV not created.")
# #         return

# #     # Flatten JSON into columns
# #     df = json_normalize(all_records, sep="_")

# #     # Convert lists/dicts → JSON strings
# #     for col in df.columns:
# #         df[col] = df[col].apply(
# #             lambda x: json.dumps(x) if isinstance(x, (list, dict)) else x
# #         )

# #     output_file = f"all_branches_sales_{start_date}_{end_date}.csv"
# #     df.to_csv(output_file, index=False, encoding="utf-8-sig")

# #     print(f"\nCSV saved → {output_file}")

# # if __name__ == "__main__":
# #     save_all_branches()

# # import time
# # import jwt
# # import requests
# # import json
# # import pandas as pd
# # from pandas import json_normalize
# # from datetime import datetime, timedelta

# # # API credentials
# # API_KEY = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
# # SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# # def generate_jwt_token(api_key, secret_key):
# #     payload = {"iss": api_key, "iat": int(time.time())}
# #     return jwt.encode(payload, secret_key, algorithm="HS256")

# # def fetch_sales_for_branch(branch, day):
# #     url = "https://api.ristaapps.com/v1/sales/page"
# #     jwt_token = generate_jwt_token(API_KEY, SECRET_KEY)

# #     headers = {
# #         "x-api-key": API_KEY,
# #         "x-api-token": jwt_token,
# #         "content-type": "application/json"
# #     }

# #     params = {"branch": branch, "day": day}
# #     all_data = []
# #     iteration = 1

# #     while True:
# #         print(f"[{branch}] [{day}] Fetching page {iteration}...")
# #         response = requests.get(url, headers=headers, params=params)

# #         if response.status_code != 200:
# #             print(f"[{branch}] [{day}] Failed: {response.status_code} → {response.text}")
# #             break

# #         resp_json = response.json()
# #         page_data = resp_json.get("data", [])
# #         all_data.extend(page_data)

# #         last_key = resp_json.get("lastKey")
# #         if not last_key:
# #             print(f"[{branch}] [{day}] Completed fetching all pages.")
# #             break

# #         params["lastKey"] = last_key
# #         iteration += 1
# #         time.sleep(1)

# #     return all_data

# # def daterange(start_date, end_date):
# #     """Yield all dates from start_date to end_date inclusive."""
# #     current = start_date
# #     while current <= end_date:
# #         yield current
# #         current += timedelta(days=1)

# # def extract_item_rows(record):
# #     """Turn each item in record['items'] into its own row with invoice-level fields attached."""

# #     # --- INVOICE LEVEL FIELDS ---
# #     invoice = {
# #         "branchName": record.get("branchName"),
# #         "branchCode": record.get("branchCode"),
# #         "invoiceNumber": record.get("invoiceNumber"),
# #         "invoiceDate": record.get("invoiceDate"),
# #         "createdDate": record.get("createdDate"),
# #         "Date": record.get("date"),
# #         "chargeTaxTotal": record.get("chargeTaxTotal"),
# #         "channel": record.get("channel"),
# #         "itemCount": record.get("itemCount"),
# #         "itemTotalAmount": record.get("itemTotalAmount"),
# #         "totalDiscountAmount": record.get("totalDiscountAmount"),
# #         "grossAmount": record.get("grossAmount"),
# #         "netDiscountAmount": record.get("netDiscountAmount"),
# #         "netAmount": record.get("netAmount"),
# #         "status": record.get("status"),
# #         "delivery_mode": record.get("delivery_mode"),

# #         # delivery.address.*
# #         "delivery_address_addressLine": record.get("delivery_address", {}).get("addressLine"),
# #         "delivery_address_zip": record.get("delivery_address", {}).get("zip"),
# #         "delivery_address_longitude": record.get("delivery_address", {}).get("longitude"),
# #         "delivery_address_latitude": record.get("delivery_address", {}).get("latitude"),

# #         # customer.*
# #         "customer_name": record.get("customer", {}).get("name"),
# #         "customer_email": record.get("customer", {}).get("email"),
# #         "customer_phoneNumber": record.get("customer", {}).get("phoneNumber"),
# #         "customer_id": record.get("customer", {}).get("id"),

# #         "orderNumber": record.get("orderNumber"),

# #         # Metadata added by script
# #         "branch": record.get("branch"),
# #         "date": record.get("date"),
# #     }

# #     items = record.get("items", [])
# #     rows = []

# #     for item in items:
# #         row = invoice.copy()

# #         # ---- ITEM LEVEL FIELDS ----
# #         row.update({
# #             "item_shortName": item.get("shortName"),
# #             "item_longName": item.get("longName"),
# #             "item_skuCode": item.get("skuCode"),
# #             "item_categoryName": item.get("categoryName"),
# #             "item_subCategoryName": item.get("subCategoryName"),
# #             "item_quantity": item.get("quantity"),
# #             "item_unitPrice": item.get("unitPrice"),
# #             "item_measuringUnit": item.get("measuringUnit"),
# #             "item_discountAmount": item.get("discountAmount"),
# #             "item_grossAmount": item.get("grossAmount"),
# #             "item_netDiscountAmount": item.get("netDiscountAmount"),
# #             "item_netAmount": item.get("netAmount"),
# #             "item_taxAmount": item.get("taxAmount"),

# #             # Extra useful item fields you showed in the sample:
# #             "item_itemNumber": item.get("itemNumber"),
# #             "item_itemAmount": item.get("itemAmount"),
# #             "item_itemTotalAmount": item.get("itemTotalAmount"),
# #             "item_optionAmount": item.get("optionAmount"),
# #             "item_baseGrossAmount": item.get("baseGrossAmount"),
# #             "item_baseNetAmount": item.get("baseNetAmount"),
# #             "item_baseTaxAmount": item.get("baseTaxAmount"),
# #             "item_createdTime": item.get("createdTime"),
# #             "item_kotNumber": item.get("kotNumber"),
# #             "item_kotStatus": item.get("kotStatus"),
# #         })

# #         # if taxes list exists take first tax (common in Rista)
# #         taxes = item.get("taxes", [])
# #         if taxes:
# #             tax = taxes[0]
# #             row.update({
# #                 "item_tax_name": tax.get("name"),
# #                 "item_tax_percentage": tax.get("percentage"),
# #                 "item_tax_amountIncluded": tax.get("amountIncluded"),
# #                 "item_tax_amountExcluded": tax.get("amountExcluded"),
# #             })

# #         rows.append(row)

# #     return rows

# # def save_all_branches(start_date="2025-12-16", end_date="2025-12-17"):
# #     # Load branch codes
# #     df_branches = pd.read_csv("branch_codes.csv")
# #     branches = df_branches["branch"].dropna().unique()

# #     # Convert input strings to date objects
# #     start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
# #     end_date = datetime.strptime(end_date, "%Y-%m-%d").date()

# #     final_rows = []

# #     for branch in branches:
# #         print(f"\n========== Processing branch: {branch} ==========")

# #         for day in daterange(start_date, end_date):
# #             day_str = day.strftime("%Y-%m-%d")

# #             print(f"\n--- Fetching for date: {day_str} ---")

# #             branch_data = fetch_sales_for_branch(branch, day_str)

# #             # Add metadata + convert items → rows
# #             for record in branch_data:
# #                 record["branch"] = branch
# #                 record["date"] = day_str

# #                 rows = extract_item_rows(record)
# #                 final_rows.extend(rows)

# #     print(f"\nTotal item rows generated: {len(final_rows)}")

# #     if not final_rows:
# #         print("No data fetched. CSV not created.")
# #         return

# #     df = pd.DataFrame(final_rows)

# #     output_file = f"sales_item_level_{start_date}_{end_date}.csv"
# #     df.to_csv(output_file, index=False, encoding="utf-8-sig")

# #     print(f"\nCSV saved → {output_file}")

# # if __name__ == "__main__":
# #     save_all_branches()

# # import time
# # import jwt
# # import requests
# # import json
# # import pandas as pd
# # from pandas import json_normalize
# # from datetime import datetime, timedelta

# # # API credentials
# # API_KEY = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
# # SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# # def generate_jwt_token(api_key, secret_key):
# #     payload = {"iss": api_key, "iat": int(time.time())}
# #     return jwt.encode(payload, secret_key, algorithm="HS256")

# # def fetch_sales_for_branch(branch, day):
# #     url = "https://api.ristaapps.com/v1/sales/page"
# #     jwt_token = generate_jwt_token(API_KEY, SECRET_KEY)

# #     headers = {
# #         "x-api-key": API_KEY,
# #         "x-api-token": jwt_token,
# #         "content-type": "application/json"
# #     }

# #     params = {"branch": branch, "day": day}
# #     all_data = []
# #     iteration = 1

# #     while True:
# #         print(f"[{branch}] [{day}] Fetching page {iteration}...")
# #         response = requests.get(url, headers=headers, params=params)

# #         if response.status_code != 200:
# #             print(f"[{branch}] [{day}] Failed: {response.status_code} → {response.text}")
# #             break

# #         resp_json = response.json()
# #         page_data = resp_json.get("data", [])
# #         all_data.extend(page_data)

# #         last_key = resp_json.get("lastKey")
# #         if not last_key:
# #             print(f"[{branch}] [{day}] Completed fetching all pages.")
# #             break

# #         params["lastKey"] = last_key
# #         iteration += 1
# #         time.sleep(1)

# #     return all_data

# # def daterange(start_date, end_date):
# #     """Yield all dates from start_date to end_date inclusive."""
# #     current = start_date
# #     while current <= end_date:
# #         yield current
# #         current += timedelta(days=1)

# # def extract_phone_number(record):
# #     """
# #     Try multiple possible locations in the API response for a phone number.
# #     Returns the first non-empty value found, or None.
# #     """
# #     # 1. customer.phoneNumber (primary)
# #     phone = record.get("customer", {}).get("phoneNumber")
# #     if phone:
# #         return str(phone).strip()

# #     # 2. customer.phone
# #     phone = record.get("customer", {}).get("phone")
# #     if phone:
# #         return str(phone).strip()

# #     # 3. customer.mobile
# #     phone = record.get("customer", {}).get("mobile")
# #     if phone:
# #         return str(phone).strip()

# #     # 4. Top-level phoneNumber
# #     phone = record.get("phoneNumber")
# #     if phone:
# #         return str(phone).strip()

# #     # 5. Top-level phone
# #     phone = record.get("phone")
# #     if phone:
# #         return str(phone).strip()

# #     # 6. delivery.phoneNumber
# #     phone = record.get("delivery", {}).get("phoneNumber")
# #     if phone:
# #         return str(phone).strip()

# #     # 7. delivery_address.phoneNumber
# #     phone = record.get("delivery_address", {}).get("phoneNumber")
# #     if phone:
# #         return str(phone).strip()

# #     # 8. contactNumber
# #     phone = record.get("contactNumber")
# #     if phone:
# #         return str(phone).strip()

# #     return None

# # def extract_item_rows(record):
# #     """Turn each item in record['items'] into its own row with invoice-level fields attached."""

# #     # Extract phone using multi-location resolver
# #     phone_number = extract_phone_number(record)

# #     # --- INVOICE LEVEL FIELDS ---
# #     invoice = {
# #         "branchName": record.get("branchName"),
# #         "branchCode": record.get("branchCode"),
# #         "invoiceNumber": record.get("invoiceNumber"),
# #         "invoiceDate": record.get("invoiceDate"),
# #         "createdDate": record.get("createdDate"),
# #         "Date": record.get("date"),
# #         "chargeTaxTotal": record.get("chargeTaxTotal"),
# #         "channel": record.get("channel"),
# #         "itemCount": record.get("itemCount"),
# #         "itemTotalAmount": record.get("itemTotalAmount"),
# #         "totalDiscountAmount": record.get("totalDiscountAmount"),
# #         "grossAmount": record.get("grossAmount"),
# #         "netDiscountAmount": record.get("netDiscountAmount"),
# #         "netAmount": record.get("netAmount"),
# #         "status": record.get("status"),
# #         "delivery_mode": record.get("delivery_mode"),

# #         # delivery.address.*
# #         "delivery_address_addressLine": record.get("delivery_address", {}).get("addressLine"),
# #         "delivery_address_zip": record.get("delivery_address", {}).get("zip"),
# #         "delivery_address_longitude": record.get("delivery_address", {}).get("longitude"),
# #         "delivery_address_latitude": record.get("delivery_address", {}).get("latitude"),

# #         # customer.*
# #         "customer_name": record.get("customer", {}).get("name"),
# #         "customer_email": record.get("customer", {}).get("email"),
# #         "customer_phoneNumber": phone_number,   # ← resolved from all possible locations
# #         "customer_id": record.get("customer", {}).get("id"),

# #         "orderNumber": record.get("orderNumber"),

# #         # Metadata added by script
# #         "branch": record.get("branch"),
# #         "date": record.get("date"),
# #     }

# #     items = record.get("items", [])
# #     rows = []

# #     for item in items:
# #         row = invoice.copy()

# #         # ---- ITEM LEVEL FIELDS ----
# #         row.update({
# #             "item_shortName": item.get("shortName"),
# #             "item_longName": item.get("longName"),
# #             "item_skuCode": item.get("skuCode"),
# #             "item_categoryName": item.get("categoryName"),
# #             "item_subCategoryName": item.get("subCategoryName"),
# #             "item_quantity": item.get("quantity"),
# #             "item_unitPrice": item.get("unitPrice"),
# #             "item_measuringUnit": item.get("measuringUnit"),
# #             "item_discountAmount": item.get("discountAmount"),
# #             "item_grossAmount": item.get("grossAmount"),
# #             "item_netDiscountAmount": item.get("netDiscountAmount"),
# #             "item_netAmount": item.get("netAmount"),
# #             "item_taxAmount": item.get("taxAmount"),

# #             "item_itemNumber": item.get("itemNumber"),
# #             "item_itemAmount": item.get("itemAmount"),
# #             "item_itemTotalAmount": item.get("itemTotalAmount"),
# #             "item_optionAmount": item.get("optionAmount"),
# #             "item_baseGrossAmount": item.get("baseGrossAmount"),
# #             "item_baseNetAmount": item.get("baseNetAmount"),
# #             "item_baseTaxAmount": item.get("baseTaxAmount"),
# #             "item_createdTime": item.get("createdTime"),
# #             "item_kotNumber": item.get("kotNumber"),
# #             "item_kotStatus": item.get("kotStatus"),
# #         })

# #         # if taxes list exists take first tax (common in Rista)
# #         taxes = item.get("taxes", [])
# #         if taxes:
# #             tax = taxes[0]
# #             row.update({
# #                 "item_tax_name": tax.get("name"),
# #                 "item_tax_percentage": tax.get("percentage"),
# #                 "item_tax_amountIncluded": tax.get("amountIncluded"),
# #                 "item_tax_amountExcluded": tax.get("amountExcluded"),
# #             })

# #         rows.append(row)

# #     return rows

# # def debug_phone_fields(branch, day):
# #     """
# #     Helper: Print raw customer + top-level fields from the first record
# #     so you can see exactly where the phone number lives in the API response.
# #     """
# #     print(f"\n=== DEBUG: Checking phone field locations for {branch} on {day} ===")
# #     data = fetch_sales_for_branch(branch, day)
# #     if not data:
# #         print("No data returned.")
# #         return

# #     record = data[0]
# #     print("Top-level keys:", list(record.keys()))
# #     print("customer object:", json.dumps(record.get("customer", {}), indent=2))
# #     print("delivery object:", json.dumps(record.get("delivery", {}), indent=2))
# #     print("delivery_address:", json.dumps(record.get("delivery_address", {}), indent=2))
# #     print("phoneNumber (top-level):", record.get("phoneNumber"))
# #     print("phone (top-level):", record.get("phone"))
# #     print("contactNumber:", record.get("contactNumber"))

# # def save_all_branches(start_date="2025-12-16", end_date="2025-12-17"):
# #     # Load branch codes
# #     df_branches = pd.read_csv("branch_codes.csv")
# #     branches = df_branches["branch"].dropna().unique()

# #     # Convert input strings to date objects
# #     start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
# #     end_date = datetime.strptime(end_date, "%Y-%m-%d").date()

# #     final_rows = []
# #     missing_phone_count = 0

# #     for branch in branches:
# #         print(f"\n========== Processing branch: {branch} ==========")

# #         for day in daterange(start_date, end_date):
# #             day_str = day.strftime("%Y-%m-%d")
# #             print(f"\n--- Fetching for date: {day_str} ---")

# #             branch_data = fetch_sales_for_branch(branch, day_str)

# #             for record in branch_data:
# #                 record["branch"] = branch
# #                 record["date"] = day_str

# #                 # Warn if phone is missing
# #                 if not extract_phone_number(record):
# #                     missing_phone_count += 1

# #                 rows = extract_item_rows(record)
# #                 final_rows.extend(rows)

# #     print(f"\nTotal item rows generated: {len(final_rows)}")
# #     print(f"Records with no phone number found: {missing_phone_count}")

# #     if not final_rows:
# #         print("No data fetched. CSV not created.")
# #         return

# #     df = pd.DataFrame(final_rows)

# #     # Quick check: how many rows have a phone number
# #     filled = df["customer_phoneNumber"].notna().sum()
# #     print(f"Rows with customer_phoneNumber filled: {filled} / {len(df)}")

# #     output_file = f"sales_item_level_{start_date}_{end_date}.csv"
# #     df.to_csv(output_file, index=False, encoding="utf-8-sig")
# #     print(f"\nCSV saved → {output_file}")

# # if __name__ == "__main__":
# #     save_all_branches()


# # working 

# # import time
# # import jwt
# # import requests
# # import json
# # import pandas as pd
# # from datetime import datetime, timedelta

# # # API credentials
# # API_KEY = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
# # SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# # # ── Auto-compute yesterday ──────────────────────────────────────────────────
# # YESTERDAY = (datetime.now() - timedelta(days=1)).date()
# # YESTERDAY_STR = YESTERDAY.strftime("%Y-%m-%d")


# # def generate_jwt_token(api_key, secret_key):
# #     payload = {"iss": api_key, "iat": int(time.time())}
# #     return jwt.encode(payload, secret_key, algorithm="HS256")


# # def fetch_sales_for_branch(branch, day):
# #     url = "https://api.ristaapps.com/v1/sales/page"
# #     jwt_token = generate_jwt_token(API_KEY, SECRET_KEY)

# #     headers = {
# #         "x-api-key": API_KEY,
# #         "x-api-token": jwt_token,
# #         "content-type": "application/json"
# #     }

# #     params = {"branch": branch, "day": day}
# #     all_data = []
# #     iteration = 1

# #     while True:
# #         print(f"[{branch}] [{day}] Fetching page {iteration}...")
# #         response = requests.get(url, headers=headers, params=params)

# #         if response.status_code != 200:
# #             print(f"[{branch}] [{day}] Failed: {response.status_code} → {response.text}")
# #             break

# #         resp_json = response.json()
# #         page_data = resp_json.get("data", [])
# #         all_data.extend(page_data)

# #         last_key = resp_json.get("lastKey")
# #         if not last_key:
# #             print(f"[{branch}] [{day}] Completed fetching all pages.")
# #             break

# #         params["lastKey"] = last_key
# #         iteration += 1
# #         time.sleep(1)

# #     return all_data


# # def extract_phone_number(record):
# #     """Check multiple possible locations for a phone number."""
# #     checks = [
# #         record.get("customer", {}).get("phoneNumber"),
# #         record.get("customer", {}).get("phone"),
# #         record.get("customer", {}).get("mobile"),
# #         record.get("phoneNumber"),
# #         record.get("phone"),
# #         record.get("delivery", {}).get("phoneNumber"),
# #         record.get("delivery_address", {}).get("phoneNumber"),
# #         record.get("contactNumber"),
# #     ]
# #     for phone in checks:
# #         if phone:
# #             return str(phone).strip()
# #     return None


# # def extract_item_rows(record):
# #     """Turn each item in record['items'] into its own row with invoice-level fields attached."""

# #     phone_number = extract_phone_number(record)

# #     invoice = {
# #         "branchName":               record.get("branchName"),
# #         "branchCode":               record.get("branchCode"),
# #         "invoiceNumber":            record.get("invoiceNumber"),
# #         "invoiceDate":              record.get("invoiceDate"),
# #         "createdDate":              record.get("createdDate"),
# #         "Date":                     record.get("date"),
# #         "chargeTaxTotal":           record.get("chargeTaxTotal"),
# #         "channel":                  record.get("channel"),
# #         "itemCount":                record.get("itemCount"),
# #         "itemTotalAmount":          record.get("itemTotalAmount"),
# #         "totalDiscountAmount":      record.get("totalDiscountAmount"),
# #         "grossAmount":              record.get("grossAmount"),
# #         "netDiscountAmount":        record.get("netDiscountAmount"),
# #         "netAmount":                record.get("netAmount"),
# #         "status":                   record.get("status"),
# #         "delivery_mode":            record.get("delivery_mode"),

# #         "delivery_address_addressLine": record.get("delivery_address", {}).get("addressLine"),
# #         "delivery_address_zip":         record.get("delivery_address", {}).get("zip"),
# #         "delivery_address_longitude":   record.get("delivery_address", {}).get("longitude"),
# #         "delivery_address_latitude":    record.get("delivery_address", {}).get("latitude"),

# #         "customer_name":        record.get("customer", {}).get("name"),
# #         "customer_email":       record.get("customer", {}).get("email"),
# #         "customer_phoneNumber": phone_number,
# #         "customer_id":          record.get("customer", {}).get("id"),

# #         "orderNumber": record.get("orderNumber"),
# #         "branch":      record.get("branch"),
# #         "date":        record.get("date"),
# #     }

# #     rows = []
# #     for item in record.get("items", []):
# #         row = invoice.copy()
# #         row.update({
# #             "item_shortName":        item.get("shortName"),
# #             "item_longName":         item.get("longName"),
# #             "item_skuCode":          item.get("skuCode"),
# #             "item_categoryName":     item.get("categoryName"),
# #             "item_subCategoryName":  item.get("subCategoryName"),
# #             "item_quantity":         item.get("quantity"),
# #             "item_unitPrice":        item.get("unitPrice"),
# #             "item_measuringUnit":    item.get("measuringUnit"),
# #             "item_discountAmount":   item.get("discountAmount"),
# #             "item_grossAmount":      item.get("grossAmount"),
# #             "item_netDiscountAmount":item.get("netDiscountAmount"),
# #             "item_netAmount":        item.get("netAmount"),
# #             "item_taxAmount":        item.get("taxAmount"),
# #             "item_itemNumber":       item.get("itemNumber"),
# #             "item_itemAmount":       item.get("itemAmount"),
# #             "item_itemTotalAmount":  item.get("itemTotalAmount"),
# #             "item_optionAmount":     item.get("optionAmount"),
# #             "item_baseGrossAmount":  item.get("baseGrossAmount"),
# #             "item_baseNetAmount":    item.get("baseNetAmount"),
# #             "item_baseTaxAmount":    item.get("baseTaxAmount"),
# #             "item_createdTime":      item.get("createdTime"),
# #             "item_kotNumber":        item.get("kotNumber"),
# #             "item_kotStatus":        item.get("kotStatus"),
# #         })

# #         taxes = item.get("taxes", [])
# #         if taxes:
# #             tax = taxes[0]
# #             row.update({
# #                 "item_tax_name":          tax.get("name"),
# #                 "item_tax_percentage":    tax.get("percentage"),
# #                 "item_tax_amountIncluded":tax.get("amountIncluded"),
# #                 "item_tax_amountExcluded":tax.get("amountExcluded"),
# #             })

# #         rows.append(row)

# #     return rows


# # def run():
# #     print(f"\n{'='*55}")
# #     print(f"  Daily sales scrape — fetching data for: {YESTERDAY_STR}")
# #     print(f"{'='*55}\n")

# #     df_branches = pd.read_csv("data/branch_codes.csv")
# #     branches = df_branches["branch"].dropna().unique()

# #     final_rows = []
# #     missing_phone_count = 0

# #     for branch in branches:
# #         print(f"\n========== Processing branch: {branch} ==========")
# #         branch_data = fetch_sales_for_branch(branch, YESTERDAY_STR)

# #         for record in branch_data:
# #             record["branch"] = branch
# #             record["date"] = YESTERDAY_STR

# #             if not extract_phone_number(record):
# #                 missing_phone_count += 1

# #             final_rows.extend(extract_item_rows(record))

# #     print(f"\nTotal item rows generated : {len(final_rows)}")
# #     print(f"Records with no phone     : {missing_phone_count}")

# #     if not final_rows:
# #         print("No data fetched. CSV not created.")
# #         return

# #     df = pd.DataFrame(final_rows)

# #     filled = df["customer_phoneNumber"].notna().sum()
# #     print(f"Rows with phone filled    : {filled} / {len(df)}")

# #     output_file = f"sales_{YESTERDAY_STR}.csv"
# #     df.to_csv(output_file, index=False, encoding="utf-8-sig")
# #     print(f"\n✅ CSV saved → {output_file}")


# # if __name__ == "__main__":
# #     run()

# # import time
# # import jwt
# # import requests
# # import json
# # import pandas as pd
# # from datetime import datetime, timedelta

# # # API credentials
# # API_KEY = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
# # SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# # # ── Auto-compute date range: last 3 weeks back from yesterday ───────────────
# # YESTERDAY  = (datetime.now() - timedelta(days=1)).date()
# # START_DATE = YESTERDAY - timedelta(weeks=3)


# # def generate_jwt_token(api_key, secret_key):
# #     payload = {"iss": api_key, "iat": int(time.time())}
# #     return jwt.encode(payload, secret_key, algorithm="HS256")


# # def daterange(start, end):
# #     """Yield every date from start to end inclusive."""
# #     current = start
# #     while current <= end:
# #         yield current
# #         current += timedelta(days=1)


# # def fetch_sales_for_branch(branch, day):
# #     url = "https://api.ristaapps.com/v1/sales/page"
# #     jwt_token = generate_jwt_token(API_KEY, SECRET_KEY)

# #     headers = {
# #         "x-api-key": API_KEY,
# #         "x-api-token": jwt_token,
# #         "content-type": "application/json"
# #     }

# #     params = {"branch": branch, "day": day}
# #     all_data = []
# #     iteration = 1

# #     while True:
# #         print(f"[{branch}] [{day}] Fetching page {iteration}...")
# #         response = requests.get(url, headers=headers, params=params)

# #         if response.status_code != 200:
# #             print(f"[{branch}] [{day}] Failed: {response.status_code} → {response.text}")
# #             break

# #         resp_json = response.json()
# #         page_data = resp_json.get("data", [])
# #         all_data.extend(page_data)

# #         last_key = resp_json.get("lastKey")
# #         if not last_key:
# #             print(f"[{branch}] [{day}] Completed fetching all pages.")
# #             break

# #         params["lastKey"] = last_key
# #         iteration += 1
# #         time.sleep(1)

# #     return all_data


# # def extract_phone_number(record):
# #     """Check multiple possible locations for a phone number."""
# #     checks = [
# #         record.get("customer", {}).get("phoneNumber"),
# #         record.get("customer", {}).get("phone"),
# #         record.get("customer", {}).get("mobile"),
# #         record.get("phoneNumber"),
# #         record.get("phone"),
# #         record.get("delivery", {}).get("phoneNumber"),
# #         record.get("delivery_address", {}).get("phoneNumber"),
# #         record.get("contactNumber"),
# #     ]
# #     for phone in checks:
# #         if phone:
# #             return str(phone).strip()
# #     return None


# # def extract_item_rows(record):
# #     """Turn each item in record['items'] into its own row with invoice-level fields attached."""

# #     phone_number = extract_phone_number(record)

# #     invoice = {
# #         "branchName":               record.get("branchName"),
# #         "branchCode":               record.get("branchCode"),
# #         "invoiceNumber":            record.get("invoiceNumber"),
# #         "invoiceDate":              record.get("invoiceDate"),
# #         "createdDate":              record.get("createdDate"),
# #         "Date":                     record.get("date"),
# #         "chargeTaxTotal":           record.get("chargeTaxTotal"),
# #         "channel":                  record.get("channel"),
# #         "itemCount":                record.get("itemCount"),
# #         "itemTotalAmount":          record.get("itemTotalAmount"),
# #         "totalDiscountAmount":      record.get("totalDiscountAmount"),
# #         "grossAmount":              record.get("grossAmount"),
# #         "netDiscountAmount":        record.get("netDiscountAmount"),
# #         "netAmount":                record.get("netAmount"),
# #         "status":                   record.get("status"),
# #         "delivery_mode":            record.get("delivery_mode"),

# #         "delivery_address_addressLine": record.get("delivery_address", {}).get("addressLine"),
# #         "delivery_address_zip":         record.get("delivery_address", {}).get("zip"),
# #         "delivery_address_longitude":   record.get("delivery_address", {}).get("longitude"),
# #         "delivery_address_latitude":    record.get("delivery_address", {}).get("latitude"),

# #         "customer_name":        record.get("customer", {}).get("name"),
# #         "customer_email":       record.get("customer", {}).get("email"),
# #         "customer_phoneNumber": phone_number,
# #         "customer_id":          record.get("customer", {}).get("id"),

# #         "orderNumber": record.get("orderNumber"),
# #         "branch":      record.get("branch"),
# #         "date":        record.get("date"),
# #     }

# #     rows = []
# #     for item in record.get("items", []):
# #         row = invoice.copy()
# #         row.update({
# #             "item_shortName":         item.get("shortName"),
# #             "item_longName":          item.get("longName"),
# #             "item_skuCode":           item.get("skuCode"),
# #             "item_categoryName":      item.get("categoryName"),
# #             "item_subCategoryName":   item.get("subCategoryName"),
# #             "item_quantity":          item.get("quantity"),
# #             "item_unitPrice":         item.get("unitPrice"),
# #             "item_measuringUnit":     item.get("measuringUnit"),
# #             "item_discountAmount":    item.get("discountAmount"),
# #             "item_grossAmount":       item.get("grossAmount"),
# #             "item_netDiscountAmount": item.get("netDiscountAmount"),
# #             "item_netAmount":         item.get("netAmount"),
# #             "item_taxAmount":         item.get("taxAmount"),
# #             "item_itemNumber":        item.get("itemNumber"),
# #             "item_itemAmount":        item.get("itemAmount"),
# #             "item_itemTotalAmount":   item.get("itemTotalAmount"),
# #             "item_optionAmount":      item.get("optionAmount"),
# #             "item_baseGrossAmount":   item.get("baseGrossAmount"),
# #             "item_baseNetAmount":     item.get("baseNetAmount"),
# #             "item_baseTaxAmount":     item.get("baseTaxAmount"),
# #             "item_createdTime":       item.get("createdTime"),
# #             "item_kotNumber":         item.get("kotNumber"),
# #             "item_kotStatus":         item.get("kotStatus"),
# #         })

# #         taxes = item.get("taxes", [])
# #         if taxes:
# #             tax = taxes[0]
# #             row.update({
# #                 "item_tax_name":           tax.get("name"),
# #                 "item_tax_percentage":     tax.get("percentage"),
# #                 "item_tax_amountIncluded": tax.get("amountIncluded"),
# #                 "item_tax_amountExcluded": tax.get("amountExcluded"),
# #             })

# #         rows.append(row)

# #     return rows


# # def run():
# #     total_days = (YESTERDAY - START_DATE).days + 1

# #     print(f"\n{'='*60}")
# #     print(f"  Scraping last 3 weeks of sales data")
# #     print(f"  From : {START_DATE}  →  To : {YESTERDAY}  ({total_days} days)")
# #     print(f"{'='*60}\n")

# #     df_branches = pd.read_csv("data/branch_codes.csv")
# #     branches = df_branches["branch"].dropna().unique()

# #     final_rows = []
# #     missing_phone_count = 0

# #     for branch in branches:
# #         print(f"\n========== Branch: {branch} ==========")

# #         for day in daterange(START_DATE, YESTERDAY):
# #             day_str = day.strftime("%Y-%m-%d")
# #             branch_data = fetch_sales_for_branch(branch, day_str)

# #             for record in branch_data:
# #                 record["branch"] = branch
# #                 record["date"]   = day_str

# #                 if not extract_phone_number(record):
# #                     missing_phone_count += 1

# #                 final_rows.extend(extract_item_rows(record))

# #     print(f"\nTotal item rows generated : {len(final_rows)}")
# #     print(f"Records with no phone     : {missing_phone_count}")

# #     if not final_rows:
# #         print("No data fetched. CSV not created.")
# #         return

# #     df = pd.DataFrame(final_rows)

# #     filled = df["customer_phoneNumber"].notna().sum()
# #     print(f"Rows with phone filled    : {filled} / {len(df)}")

# #     # Single CSV covering the full 3-week window
# #     output_file = f"sales_{START_DATE}_to_{YESTERDAY}.csv"
# #     df.to_csv(output_file, index=False, encoding="utf-8-sig")
# #     print(f"\n✅ CSV saved → {output_file}")


# # if __name__ == "__main__":
# #     run()


# import time
# import jwt
# import requests
# import pandas as pd
# from datetime import datetime, timedelta
# import os
# import sys

# # API credentials
# API_KEY    = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
# SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# # ── Single rolling CSV — always kept at 21 days ─────────────────────────────
# DATA_FILE    = "data/sales_rolling.csv"
# BRANCH_FILE  = "data/branch_codes.csv"
# ROLLING_DAYS = 21

# YESTERDAY  = (datetime.now() - timedelta(days=1)).date()
# START_DATE = YESTERDAY - timedelta(days=ROLLING_DAYS - 1)   # used for backfill only


# # ── Helpers ──────────────────────────────────────────────────────────────────
# def generate_jwt_token():
#     payload = {"iss": API_KEY, "iat": int(time.time())}
#     return jwt.encode(payload, SECRET_KEY, algorithm="HS256")


# def daterange(start, end):
#     current = start
#     while current <= end:
#         yield current
#         current += timedelta(days=1)


# # ── API fetch ─────────────────────────────────────────────────────────────────
# def fetch_sales_for_branch(branch, day_str):
#     url     = "https://api.ristaapps.com/v1/sales/page"
#     headers = {
#         "x-api-key":    API_KEY,
#         "x-api-token":  generate_jwt_token(),
#         "content-type": "application/json"
#     }
#     params   = {"branch": branch, "day": day_str}
#     all_data = []
#     page     = 1

#     while True:
#         print(f"  [{branch}] [{day_str}] page {page}...")
#         r = requests.get(url, headers=headers, params=params)
#         if r.status_code != 200:
#             print(f"  Failed {r.status_code}: {r.text}")
#             break
#         rj = r.json()
#         all_data.extend(rj.get("data", []))
#         last_key = rj.get("lastKey")
#         if not last_key:
#             break
#         params["lastKey"] = last_key
#         page += 1
#         time.sleep(1)

#     return all_data


# def extract_phone_number(record):
#     checks = [
#         record.get("customer", {}).get("phoneNumber"),
#         record.get("customer", {}).get("phone"),
#         record.get("customer", {}).get("mobile"),
#         record.get("phoneNumber"),
#         record.get("phone"),
#         record.get("delivery", {}).get("phoneNumber"),
#         record.get("delivery_address", {}).get("phoneNumber"),
#         record.get("contactNumber"),
#     ]
#     for phone in checks:
#         if phone:
#             return str(phone).strip()
#     return None


# def extract_item_rows(record):
#     phone   = extract_phone_number(record)
#     invoice = {
#         "branchName":               record.get("branchName"),
#         "branchCode":               record.get("branchCode"),
#         "invoiceNumber":            record.get("invoiceNumber"),
#         "invoiceDate":              record.get("invoiceDate"),
#         "createdDate":              record.get("createdDate"),
#         "chargeTaxTotal":           record.get("chargeTaxTotal"),
#         "channel":                  record.get("channel"),
#         "itemCount":                record.get("itemCount"),
#         "itemTotalAmount":          record.get("itemTotalAmount"),
#         "totalDiscountAmount":      record.get("totalDiscountAmount"),
#         "grossAmount":              record.get("grossAmount"),
#         "netDiscountAmount":        record.get("netDiscountAmount"),
#         "netAmount":                record.get("netAmount"),
#         "status":                   record.get("status"),
#         "delivery_mode":            record.get("delivery_mode"),
#         "delivery_address_addressLine": record.get("delivery_address", {}).get("addressLine"),
#         "delivery_address_zip":         record.get("delivery_address", {}).get("zip"),
#         "delivery_address_longitude":   record.get("delivery_address", {}).get("longitude"),
#         "delivery_address_latitude":    record.get("delivery_address", {}).get("latitude"),
#         "customer_name":        record.get("customer", {}).get("name"),
#         "customer_email":       record.get("customer", {}).get("email"),
#         "customer_phoneNumber": phone,
#         "customer_id":          record.get("customer", {}).get("id"),
#         "orderNumber":          record.get("orderNumber"),
#         "branch":               record.get("branch"),
#         "date":                 record.get("date"),
#     }
#     rows = []
#     for item in record.get("items", []):
#         row = invoice.copy()
#         row.update({
#             "item_shortName":         item.get("shortName"),
#             "item_longName":          item.get("longName"),
#             "item_skuCode":           item.get("skuCode"),
#             "item_categoryName":      item.get("categoryName"),
#             "item_subCategoryName":   item.get("subCategoryName"),
#             "item_quantity":          item.get("quantity"),
#             "item_unitPrice":         item.get("unitPrice"),
#             "item_measuringUnit":     item.get("measuringUnit"),
#             "item_discountAmount":    item.get("discountAmount"),
#             "item_grossAmount":       item.get("grossAmount"),
#             "item_netDiscountAmount": item.get("netDiscountAmount"),
#             "item_netAmount":         item.get("netAmount"),
#             "item_taxAmount":         item.get("taxAmount"),
#             "item_itemNumber":        item.get("itemNumber"),
#             "item_itemAmount":        item.get("itemAmount"),
#             "item_itemTotalAmount":   item.get("itemTotalAmount"),
#             "item_optionAmount":      item.get("optionAmount"),
#             "item_baseGrossAmount":   item.get("baseGrossAmount"),
#             "item_baseNetAmount":     item.get("baseNetAmount"),
#             "item_baseTaxAmount":     item.get("baseTaxAmount"),
#             "item_createdTime":       item.get("createdTime"),
#             "item_kotNumber":         item.get("kotNumber"),
#             "item_kotStatus":         item.get("kotStatus"),
#         })
#         taxes = item.get("taxes", [])
#         if taxes:
#             t = taxes[0]
#             row.update({
#                 "item_tax_name":           t.get("name"),
#                 "item_tax_percentage":     t.get("percentage"),
#                 "item_tax_amountIncluded": t.get("amountIncluded"),
#                 "item_tax_amountExcluded": t.get("amountExcluded"),
#             })
#         rows.append(row)
#     return rows


# # ── Scrape a list of date objects across all branches ────────────────────────
# def scrape_dates(branches, dates):
#     new_rows = []
#     for branch in branches:
#         print(f"\n=== Branch: {branch} ===")
#         for day in dates:
#             records = fetch_sales_for_branch(branch, day.strftime("%Y-%m-%d"))
#             for record in records:
#                 record["branch"] = branch
#                 record["date"]   = day.strftime("%Y-%m-%d")
#                 new_rows.extend(extract_item_rows(record))
#     return new_rows


# # ── Drop rows older than 21 days ─────────────────────────────────────────────
# def trim_to_rolling_window(df):
#     df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
#     cutoff     = YESTERDAY - timedelta(days=ROLLING_DAYS - 1)
#     before     = len(df)
#     df         = df[df["date"] >= cutoff].reset_index(drop=True)
#     dropped    = before - len(df)
#     if dropped:
#         print(f"  Trimmed {dropped} rows older than {cutoff}")
#     return df


# def save(df):
#     os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
#     df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
#     print(f"\n✅  {len(df)} rows saved  |  {df['date'].min()} → {df['date'].max()}  →  {DATA_FILE}")


# # ════════════════════════════════════════════════════════════
# #   MODE 1 — BACKFILL  (run once to seed the rolling CSV)
# # ════════════════════════════════════════════════════════════
# def run_backfill():
#     print(f"\n{'='*60}")
#     print(f"  BACKFILL  |  {START_DATE} → {YESTERDAY}  ({ROLLING_DAYS} days)")
#     print(f"{'='*60}")

#     branches = pd.read_csv(BRANCH_FILE)["branch"].dropna().unique()
#     dates    = list(daterange(START_DATE, YESTERDAY))
#     rows     = scrape_dates(branches, dates)

#     if not rows:
#         print("No data fetched.")
#         return

#     df = trim_to_rolling_window(pd.DataFrame(rows))
#     save(df)


# # ════════════════════════════════════════════════════════════
# #   MODE 2 — DAILY  (runs every night via GitHub Actions)
# #            appends yesterday → trims anything > 21 days
# # ════════════════════════════════════════════════════════════
# def run_daily():
#     print(f"\n{'='*60}")
#     print(f"  DAILY  |  Scraping {YESTERDAY}")
#     print(f"{'='*60}")

#     branches = pd.read_csv(BRANCH_FILE)["branch"].dropna().unique()
#     rows     = scrape_dates(branches, [YESTERDAY])

#     # Load existing rolling CSV
#     if os.path.exists(DATA_FILE):
#         existing = pd.read_csv(DATA_FILE, dtype=str, low_memory=False)
#         print(f"  Loaded existing: {len(existing)} rows")
#     else:
#         existing = pd.DataFrame()
#         print("  No existing file — starting fresh")

#     combined = pd.concat(
#         [existing, pd.DataFrame(rows)] if rows else [existing],
#         ignore_index=True
#     )

#     # Deduplicate in case of re-runs (same invoice + item)
#     if "item_itemNumber" in combined.columns:
#         combined = combined.drop_duplicates(
#             subset=["invoiceNumber", "item_itemNumber"], keep="last"
#         )
#     else:
#         combined = combined.drop_duplicates(subset=["invoiceNumber"], keep="last")

#     combined = trim_to_rolling_window(combined.reset_index(drop=True))
#     save(combined)


# # ════════════════════════════════════════════════════════════
# #   ENTRY POINT
# #   python sales_daily.py            → daily (GitHub Actions default)
# #   python sales_daily.py backfill   → one-time seed
# # ════════════════════════════════════════════════════════════
# if __name__ == "__main__":
#     mode = sys.argv[1] if len(sys.argv) > 1 else "daily"
#     if mode == "backfill":
#         run_backfill()
#     else:
#         run_daily()


# import time
# import jwt
# import requests
# import pandas as pd
# from datetime import datetime, timedelta
# import os
# import sys

# # API credentials
# API_KEY    = 'dfaf6eea-3622-40d5-b61b-eacdd36fd921'
# SECRET_KEY = 'JsUVbyLjrQp1abOAO94pWk_PHphRwDXnVX1pWWzWfxI'

# # ── Single rolling CSV — always kept at 21 days (3 weeks) ──────────────────
# DATA_FILE    = r"C:/Users/jeryy/OneDrive - BLISS CHOCOLATES INDIA PRIVATE LIMITED/Desktop/Documents/Phone No/sales_rolling.csv"
# BRANCH_FILE  = "data/branch_codes.csv"
# ROLLING_DAYS = 21   # 3 weeks

# YESTERDAY  = (datetime.now() - timedelta(days=1)).date()
# START_DATE = YESTERDAY - timedelta(days=ROLLING_DAYS - 1)   # used for backfill only


# # ── Helpers ──────────────────────────────────────────────────────────────────
# def generate_jwt_token():
#     payload = {"iss": API_KEY, "iat": int(time.time())}
#     return jwt.encode(payload, SECRET_KEY, algorithm="HS256")


# def daterange(start, end):
#     current = start
#     while current <= end:
#         yield current
#         current += timedelta(days=1)


# # ── API fetch ─────────────────────────────────────────────────────────────────
# def fetch_sales_for_branch(branch, day_str):
#     url     = "https://api.ristaapps.com/v1/sales/page"
#     headers = {
#         "x-api-key":    API_KEY,
#         "x-api-token":  generate_jwt_token(),
#         "content-type": "application/json"
#     }
#     params   = {"branch": branch, "day": day_str}
#     all_data = []
#     page     = 1

#     while True:
#         print(f"  [{branch}] [{day_str}] page {page}...")
#         r = requests.get(url, headers=headers, params=params)
#         if r.status_code != 200:
#             print(f"  Failed {r.status_code}: {r.text}")
#             break
#         rj = r.json()
#         all_data.extend(rj.get("data", []))
#         last_key = rj.get("lastKey")
#         if not last_key:
#             break
#         params["lastKey"] = last_key
#         page += 1
#         time.sleep(1)

#     return all_data


# def extract_phone_number(record):
#     checks = [
#         record.get("customer", {}).get("phoneNumber"),
#         record.get("customer", {}).get("phone"),
#         record.get("customer", {}).get("mobile"),
#         record.get("phoneNumber"),
#         record.get("phone"),
#         record.get("delivery", {}).get("phoneNumber"),
#         record.get("delivery_address", {}).get("phoneNumber"),
#         record.get("contactNumber"),
#     ]
#     for phone in checks:
#         if phone:
#             return str(phone).strip()
#     return None


# def extract_item_rows(record):
#     phone   = extract_phone_number(record)
#     invoice = {
#         "branchName":               record.get("branchName"),
#         "branchCode":               record.get("branchCode"),
#         "invoiceNumber":            record.get("invoiceNumber"),
#         "invoiceDate":              record.get("invoiceDate"),
#         "createdDate":              record.get("createdDate"),
#         "chargeTaxTotal":           record.get("chargeTaxTotal"),
#         "channel":                  record.get("channel"),
#         "itemCount":                record.get("itemCount"),
#         "itemTotalAmount":          record.get("itemTotalAmount"),
#         "totalDiscountAmount":      record.get("totalDiscountAmount"),
#         "grossAmount":              record.get("grossAmount"),
#         "netDiscountAmount":        record.get("netDiscountAmount"),
#         "netAmount":                record.get("netAmount"),
#         "status":                   record.get("status"),
#         "delivery_mode":            record.get("delivery_mode"),
#         "delivery_address_addressLine": record.get("delivery_address", {}).get("addressLine"),
#         "delivery_address_zip":         record.get("delivery_address", {}).get("zip"),
#         "delivery_address_longitude":   record.get("delivery_address", {}).get("longitude"),
#         "delivery_address_latitude":    record.get("delivery_address", {}).get("latitude"),
#         "customer_name":        record.get("customer", {}).get("name"),
#         "customer_email":       record.get("customer", {}).get("email"),
#         "customer_phoneNumber": phone,
#         "customer_id":          record.get("customer", {}).get("id"),
#         "orderNumber":          record.get("orderNumber"),
#         "branch":               record.get("branch"),
#         "date":                 record.get("date"),
#     }
#     rows = []
#     for item in record.get("items", []):
#         row = invoice.copy()
#         row.update({
#             "item_shortName":         item.get("shortName"),
#             "item_longName":          item.get("longName"),
#             "item_skuCode":           item.get("skuCode"),
#             "item_categoryName":      item.get("categoryName"),
#             "item_subCategoryName":   item.get("subCategoryName"),
#             "item_quantity":          item.get("quantity"),
#             "item_unitPrice":         item.get("unitPrice"),
#             "item_measuringUnit":     item.get("measuringUnit"),
#             "item_discountAmount":    item.get("discountAmount"),
#             "item_grossAmount":       item.get("grossAmount"),
#             "item_netDiscountAmount": item.get("netDiscountAmount"),
#             "item_netAmount":         item.get("netAmount"),
#             "item_taxAmount":         item.get("taxAmount"),
#             "item_itemNumber":        item.get("itemNumber"),
#             "item_itemAmount":        item.get("itemAmount"),
#             "item_itemTotalAmount":   item.get("itemTotalAmount"),
#             "item_optionAmount":      item.get("optionAmount"),
#             "item_baseGrossAmount":   item.get("baseGrossAmount"),
#             "item_baseNetAmount":     item.get("baseNetAmount"),
#             "item_baseTaxAmount":     item.get("baseTaxAmount"),
#             "item_createdTime":       item.get("createdTime"),
#             "item_kotNumber":         item.get("kotNumber"),
#             "item_kotStatus":         item.get("kotStatus"),
#         })
#         taxes = item.get("taxes", [])
#         if taxes:
#             t = taxes[0]
#             row.update({
#                 "item_tax_name":           t.get("name"),
#                 "item_tax_percentage":     t.get("percentage"),
#                 "item_tax_amountIncluded": t.get("amountIncluded"),
#                 "item_tax_amountExcluded": t.get("amountExcluded"),
#             })
#         rows.append(row)
#     return rows


# # ── Scrape a list of date objects across all branches ────────────────────────
# def scrape_dates(branches, dates):
#     new_rows = []
#     for branch in branches:
#         print(f"\n=== Branch: {branch} ===")
#         for day in dates:
#             records = fetch_sales_for_branch(branch, day.strftime("%Y-%m-%d"))
#             for record in records:
#                 record["branch"] = branch
#                 record["date"]   = day.strftime("%Y-%m-%d")
#                 new_rows.extend(extract_item_rows(record))
#     return new_rows


# # ── Drop rows older than 21 days ─────────────────────────────────────────────
# def trim_to_rolling_window(df):
#     df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
#     cutoff     = YESTERDAY - timedelta(days=ROLLING_DAYS - 1)
#     before     = len(df)
#     df         = df[df["date"] >= cutoff].reset_index(drop=True)
#     dropped    = before - len(df)
#     if dropped:
#         print(f"  Trimmed {dropped} rows older than {cutoff}")
#     return df


# def save(df):
#     os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
#     df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
#     print(f"\n✅  {len(df)} rows saved  |  {df['date'].min()} → {df['date'].max()}  →  {DATA_FILE}")


# # ════════════════════════════════════════════════════════════
# #   MODE 1 — BACKFILL  (run once to seed the rolling CSV)
# # ════════════════════════════════════════════════════════════
# def run_backfill():
#     print(f"\n{'='*60}")
#     print(f"  BACKFILL  |  {START_DATE} → {YESTERDAY}  ({ROLLING_DAYS} days)")
#     print(f"{'='*60}")

#     branches = pd.read_csv(BRANCH_FILE)["branch"].dropna().unique()
#     dates    = list(daterange(START_DATE, YESTERDAY))
#     rows     = scrape_dates(branches, dates)

#     if not rows:
#         print("No data fetched.")
#         return

#     df = trim_to_rolling_window(pd.DataFrame(rows))
#     save(df)


# # ════════════════════════════════════════════════════════════
# #   MODE 2 — DAILY  (runs every night via GitHub Actions)
# #            appends yesterday → trims anything > 21 days
# # ════════════════════════════════════════════════════════════
# def run_daily():
#     print(f"\n{'='*60}")
#     print(f"  DAILY  |  Scraping {YESTERDAY}")
#     print(f"{'='*60}")

#     branches = pd.read_csv(BRANCH_FILE)["branch"].dropna().unique()
#     rows     = scrape_dates(branches, [YESTERDAY])

#     # Seed path: the large historical CSV from OneDrive (used first time only)
#     SEED_CSV = r"C:/Users/jeryy/OneDrive - BLISS CHOCOLATES INDIA PRIVATE LIMITED/Desktop/Documents/Phone No/sales_2026-02-16_to_2026-03-09.csv"

#     # Load existing rolling CSV, or fall back to seed file, or start fresh
#     if os.path.exists(DATA_FILE):
#         existing = pd.read_csv(DATA_FILE, dtype=str, low_memory=False)
#         print(f"  Loaded rolling CSV: {len(existing)} rows")
#     elif os.path.exists(SEED_CSV):
#         existing = pd.read_csv(SEED_CSV, dtype=str, low_memory=False)
#         print(f"  Rolling CSV not found — seeding from: {SEED_CSV}  ({len(existing)} rows)")
#     else:
#         existing = pd.DataFrame()
#         print("  No existing file — starting fresh")

#     combined = pd.concat(
#         [existing, pd.DataFrame(rows)] if rows else [existing],
#         ignore_index=True
#     )

#     # Deduplicate in case of re-runs (same invoice + item)
#     if "item_itemNumber" in combined.columns:
#         combined = combined.drop_duplicates(
#             subset=["invoiceNumber", "item_itemNumber"], keep="last"
#         )
#     else:
#         combined = combined.drop_duplicates(subset=["invoiceNumber"], keep="last")

#     combined = trim_to_rolling_window(combined.reset_index(drop=True))
#     save(combined)


# # ════════════════════════════════════════════════════════════
# #   ENTRY POINT
# #   python sales_daily.py            → daily (GitHub Actions default)
# #   python sales_daily.py backfill   → one-time seed
# # ════════════════════════════════════════════════════════════
# if __name__ == "__main__":
#     mode = sys.argv[1] if len(sys.argv) > 1 else "daily"
#     if mode == "backfill":
#         run_backfill()
#     else:
#         run_daily()




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
DATA_FILE   = r"C:/Users/jeryy/OneDrive - BLISS CHOCOLATES INDIA PRIVATE LIMITED/Desktop/Documents/Phone No/Data/sales_rolling.csv"
BRANCH_FILE = "data/branch_codes.csv"
SEED_CSV    = r"C:/Users/jeryy/OneDrive - BLISS CHOCOLATES INDIA PRIVATE LIMITED/Desktop/Documents/Phone No/Data/sales_rollings.csv"

# ── Rolling window ────────────────────────────────────────────────────────────
ROLLING_DAYS = 21   # always keep exactly 3 weeks

YESTERDAY  = (datetime.now() - timedelta(days=1)).date()
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