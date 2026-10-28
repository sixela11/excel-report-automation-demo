"""Generate fully synthetic monthly sales workbooks with deliberate data-quality issues.

All names are invented and prefixed "Sample"; no real customer or employer data is used.
Usage:
    python generate_sample_data.py [--output sample_data/input]
"""
from __future__ import annotations

import argparse
import random
from datetime import datetime
from pathlib import Path

import pandas as pd

SEED = 42
COLUMNS = ["invoice_id", "invoice_date", "customer_id", "customer_name",
           "postcode", "product_code", "quantity", "unit_price"]
CUSTOMERS = [  # (customer_id, name, postcode) - leading zeros are intentional
    ("000101", "Sample Harbour Supplies", "0800"),
    ("000102", "Sample Redgum Logistics", "2000"),
    ("000103", "Sample Coastal Trades", "3000"),
    ("000104", "Sample Outback Wholesale", "0870"),
    ("000105", "Sample Riverside Clinic", "4000"),
    ("000106", "Sample Summit Property", "5000"),
    ("000107", "Sample Bayview Retail", "6000"),
    ("000108", "Sample Highland Builders", "7000"),
    ("000109", "Sample Northgate Services", "0810"),
    ("000110", "Sample Meadow Distribution", "2600"),
    ("000111", "Sample Ironbark Engineering", "3550"),
    ("000112", "Sample Lagoon Hospitality", "4870"),
]
PRODUCTS = {"P-0010": 12.50, "P-0025": 48.00, "P-0042": 7.95, "P-0108": 129.00,
            "P-0150": 22.40, "P-0207": 315.00, "P-0333": 64.99, "P-0480": 5.25}
MONTHS = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06", "2026-07"]


def make_month(rng: random.Random, month: str, rows: int = 40) -> pd.DataFrame:
    year, mon = map(int, month.split("-"))
    records = []
    for n in range(1, rows + 1):
        cust_id, name, postcode = rng.choice(CUSTOMERS)
        product = rng.choice(list(PRODUCTS))
        records.append({
            "invoice_id": f"INV-{year}{mon:02d}-{n:04d}",
            "invoice_date": datetime(year, mon, rng.randint(1, 28)),
            "customer_id": cust_id, "customer_name": name, "postcode": postcode,
            "product_code": product, "quantity": rng.randint(1, 25), "unit_price": PRODUCTS[product],
        })
    return pd.DataFrame(records, columns=COLUMNS)


def build_months(seed: int = SEED) -> dict[str, pd.DataFrame]:
    rng = random.Random(seed)
    data = {m: make_month(rng, m) for m in MONTHS}

    # Feb: two January rows re-exported (exact duplicates).
    data["2026-02"] = pd.concat([data["2026-02"], data["2026-01"].iloc[[3, 7]]], ignore_index=True)

    # Mar: IDs saved as numbers, so Excel dropped the leading zeros.
    data["2026-03"]["customer_id"] = data["2026-03"]["customer_id"].astype(int)

    # Apr: invalid values a person might type by hand.
    apr = data["2026-04"].astype(object)
    apr.loc[1, "invoice_date"] = "31/02/2026"
    apr.loc[2, "quantity"] = "two"
    apr.loc[3, "customer_name"] = None
    apr.loc[4, "unit_price"] = -5
    apr.loc[5, "invoice_id"] = "INV-2026-04"
    data["2026-04"] = apr

    # May: inconsistent headers, text dates (dd/mm/yyyy), and a conflicting duplicate of an April invoice.
    may = data["2026-05"].astype(object)
    may["invoice_date"] = [d.strftime("%d/%m/%Y") for d in may["invoice_date"]]
    conflict = data["2026-04"].iloc[[10]].copy()
    conflict["quantity"] = int(conflict["quantity"].iloc[0]) + 5
    conflict["invoice_date"] = conflict["invoice_date"].iloc[0].strftime("%d/%m/%Y")
    may = pd.concat([may, conflict], ignore_index=True)
    may.columns = ["Invoice ID", "Invoice Date", "Customer ID", " Customer Name ",
                   "Postcode", "Product Code", "Quantity", "UNIT PRICE"]
    data["2026-05"] = may

    # Jun: required column missing, so the whole file should be rejected.
    data["2026-06"] = data["2026-06"].drop(columns=["quantity"])
    return data


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Generate synthetic monthly sales workbooks.")
    parser.add_argument("--output", type=Path, default=Path("sample_data/input"))
    args = parser.parse_args(argv)
    args.output.mkdir(parents=True, exist_ok=True)
    for month, df in build_months().items():
        path = args.output / f"sales_{month}.xlsx"
        df.to_excel(path, index=False, sheet_name="Sales")
        print(f"Wrote {path} ({len(df)} rows)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
