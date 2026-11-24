#!/usr/bin/env python3
"""
Weekly Monday ShipHero I_01 Inventory Sync Script.
Fetches live On Hand / Available inventory from ShipHero GraphQL API,
writes shiphero_inventory.json, and updates fallback seeds in psi.html and index.html.
"""
import datetime
import json
import os
import re
import sys
import urllib.request

SKU_MAP = {
    "CDO-999-WT-US": "CDO-999-WT01-US",
    "CCO-999-WT-WW": "CCO-999-WT01-WW",
    "CCO-999-YL-WW": "CCO-999-YL01-WW",
    "CCO-999-OR-WW": "CCO-999-OR01-WW",
    "CCO-999-SV-WW": "CCO-999-SV01-WW",
}

MODEL_ORDER = [
    "CDO-999-WT01-US",
    "CCO-999-WT01-WW",
    "CCO-999-YL01-WW",
    "CCO-999-OR01-WW",
    "CCO-999-SV01-WW",
]


def get_kst_now():
    kst = datetime.timezone(datetime.timedelta(hours=9))
    return datetime.datetime.now(kst)


def get_current_monday_kst(now_kst):
    days_since_monday = now_kst.weekday()  # Monday is 0
    monday = now_kst.date() - datetime.timedelta(days=days_since_monday)
    return monday


def fetch_shiphero_inventory(email, password):
    auth_req = urllib.request.Request(
        "https://public-api.shiphero.com/auth/token",
        data=json.dumps({"username": email, "password": password}).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (ShipHero-PSI-Sync/1.0)",
        },
    )
    with urllib.request.urlopen(auth_req, timeout=20) as resp:
        auth_data = json.loads(resp.read().decode("utf-8"))

    token = auth_data.get("access_token")
    if not token:
        raise RuntimeError("Failed to obtain access_token from ShipHero")

    query = """
    query {
      products {
        data(first: 100) {
          edges {
            node {
              sku
              name
              warehouse_products {
                on_hand
                available
                allocated
                backorder
              }
            }
          }
        }
      }
    }
    """
    gql_req = urllib.request.Request(
        "https://public-api.shiphero.com/graphql",
        data=json.dumps({"query": query}).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
            "User-Agent": "Mozilla/5.0 (ShipHero-PSI-Sync/1.0)",
        },
    )
    with urllib.request.urlopen(gql_req, timeout=20) as resp:
        gql_data = json.loads(resp.read().decode("utf-8"))

    edges = (
        gql_data.get("data", {})
        .get("products", {})
        .get("data", {})
        .get("edges", [])
    )

    items = {}
    for edge in edges:
        node = edge.get("node") or {}
        sh_sku = node.get("sku")
        if sh_sku in SKU_MAP:
            dash_model = SKU_MAP[sh_sku]
            wps = node.get("warehouse_products") or []
            on_hand = sum(int(w.get("on_hand") or 0) for w in wps)
            available = sum(int(w.get("available") or 0) for w in wps)
            allocated = sum(int(w.get("allocated") or 0) for w in wps)
            items[dash_model] = {
                "model": dash_model,
                "shipheroSku": sh_sku,
                "name": node.get("name") or "",
                "inv01OnHand": on_hand,
                "inv01Available": available,
                "allocated": allocated,
            }

    missing = [m for m in MODEL_ORDER if m not in items]
    if missing:
        raise RuntimeError(f"Missing target SKUs from ShipHero response: {missing}")

    return items


def update_html_seeds(html_path, items, monday_str):
    if not os.path.exists(html_path):
        return
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    for model_key, data in items.items():
        oh = data["inv01OnHand"]
        av = data["inv01Available"]
        pattern = (
            r"('"
            + re.escape(model_key)
            + r"'\s*:\s*\{[^}]*?inv01OnHand:\s*)\d+(\s*,\s*inv01Available:\s*)\d+(\s*,\s*inv01Current:\s*)\d+"
        )
        content = re.sub(
            pattern,
            rf"\g<1>{oh}\g<2>{av}\g<3>{av}",
            content,
            flags=re.DOTALL,
        )

    content = re.sub(
        r"const SHIPHERO_SEED_MONDAY\s*=\s*'[^']*';",
        f"const SHIPHERO_SEED_MONDAY = '{monday_str}';",
        content,
    )

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(content)


def main():
    email = os.environ.get("SHIPHERO_EMAIL", "").strip()
    password = os.environ.get("SHIPHERO_PASSWORD", "").strip()
    if not email or not password:
        print("Error: SHIPHERO_EMAIL and SHIPHERO_PASSWORD environment variables are required.", file=sys.stderr)
        sys.exit(1)

    now_kst = get_kst_now()
    monday_date = get_current_monday_kst(now_kst)
    monday_str = monday_date.isoformat()
    iso_year, iso_week, _ = monday_date.isocalendar()
    week_code = f"{iso_week}W"

    items = fetch_shiphero_inventory(email, password)

    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    payload = {
        "mondayDate": monday_str,
        "isoYear": iso_year,
        "weekNum": iso_week,
        "weekCode": week_code,
        "syncedAt": now_kst.isoformat(timespec="seconds"),
        "items": {m: items[m] for m in MODEL_ORDER},
    }

    json_path = os.path.join(repo_root, "shiphero_inventory.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.write("\n")

    update_html_seeds(os.path.join(repo_root, "psi.html"), items, monday_str)
    update_html_seeds(os.path.join(repo_root, "index.html"), items, monday_str)

    summary = ", ".join(
        f"{m}: OH={items[m]['inv01OnHand']}/AV={items[m]['inv01Available']}"
        for m in MODEL_ORDER
    )
    print(f"[ShipHero Monday Sync] {monday_str} ({week_code}) -> {summary}")


if __name__ == "__main__":
    main()
