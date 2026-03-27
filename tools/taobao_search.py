"""
Tool: 1688/Taobao Supplier Search via Apify
=============================================
Searches 1688.com (Alibaba's wholesale platform) for products and
suppliers by keyword using Apify actors.

Input:  APIFY_API_TOKEN in .env, keywords via CLI args
Output: .tmp/taobao_results.json

Usage:
    python3 tools/taobao_search.py --keywords "便携榨汁机,LED灯"
    python3 tools/taobao_search.py --keywords "portable blender" --max-pages 3
    python3 tools/taobao_search.py --keywords "面膜" --price-min 5 --price-max 50
"""

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from apify_client import ApifyClient
from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

TMP_DIR = PROJECT_ROOT / ".tmp"
TMP_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = TMP_DIR / "taobao_results.json"
PROGRESS_FILE = TMP_DIR / "taobao_progress.json"

# Apify actors for 1688 search (in priority order)
ACTORS = [
    "songd/1688-search-scraper",
    "ecomscrape/1688-product-search-scraper",
]

# Default CNY to USD rate (override with CNY_USD_RATE in .env)
DEFAULT_CNY_USD_RATE = 0.14


def get_apify_client():
    token = os.getenv("APIFY_API_TOKEN")
    if not token:
        print("ERROR: APIFY_API_TOKEN not set in .env")
        print("See .env.template for setup instructions.")
        sys.exit(1)
    return ApifyClient(token)


def get_cny_usd_rate():
    rate_str = os.getenv("CNY_USD_RATE")
    if rate_str:
        try:
            return float(rate_str)
        except ValueError:
            pass
    return DEFAULT_CNY_USD_RATE


def parse_count(val):
    """Parse human-readable counts including Chinese suffixes."""
    if isinstance(val, (int, float)):
        return int(val)
    if not isinstance(val, str):
        return 0
    cleaned = val.replace(",", "").replace("+", "").replace("件", "").replace("笔", "").strip()
    if not cleaned or cleaned == "-":
        return 0
    cn_multipliers = {"万": 10_000, "亿": 100_000_000}
    for suffix, mult in cn_multipliers.items():
        if cleaned.endswith(suffix):
            try:
                return int(float(cleaned[:-1]) * mult)
            except ValueError:
                return 0
    multipliers = {"k": 1_000, "m": 1_000_000}
    if cleaned and cleaned[-1].lower() in multipliers:
        try:
            return int(float(cleaned[:-1]) * multipliers[cleaned[-1].lower()])
        except ValueError:
            return 0
    try:
        return int(float(cleaned))
    except ValueError:
        return 0


def parse_price_cny(val):
    """Extract numeric price from CNY string like '¥12.50' or '12.5-25.0'."""
    if isinstance(val, (int, float)):
        return round(float(val), 2), round(float(val), 2)
    if not isinstance(val, str):
        return 0.0, 0.0
    cleaned = val.replace("¥", "").replace("￥", "").replace(",", "").strip()
    if not cleaned or cleaned == "-":
        return 0.0, 0.0
    if "-" in cleaned:
        parts = cleaned.split("-")
        try:
            return round(float(parts[0].strip()), 2), round(float(parts[1].strip()), 2)
        except (ValueError, IndexError):
            return 0.0, 0.0
    try:
        p = round(float(cleaned), 2)
        return p, p
    except ValueError:
        return 0.0, 0.0


def normalize_taobao_product(raw, keyword, cny_rate):
    """Normalize 1688 actor output to unified schema."""
    # Parse price range
    price_val = raw.get("price") or raw.get("priceRange") or raw.get("unitPrice") or 0
    price_min, price_max = parse_price_cny(price_val)
    price_avg = round((price_min + price_max) / 2, 2) if (price_min + price_max) > 0 else 0.0

    # Parse MOQ
    moq_val = raw.get("moq") or raw.get("minOrder") or raw.get("min_order") or raw.get("beginAmount") or 0
    if isinstance(moq_val, str):
        moq_val = moq_val.replace("件起批", "").replace("件", "").strip()
    moq = parse_count(moq_val)

    # Parse supplier info
    supplier = raw.get("supplier") or raw.get("company") or {}
    if isinstance(supplier, str):
        supplier_name = supplier
        supplier_location = ""
        supplier_years = 0
        supplier_verified = False
    else:
        supplier_name = (
            supplier.get("name") or
            supplier.get("companyName") or
            raw.get("shopName") or
            raw.get("seller_name") or
            ""
        )
        supplier_location = (
            supplier.get("location") or
            supplier.get("city") or
            raw.get("location") or
            raw.get("sellerProvince") or
            ""
        )
        supplier_years = parse_count(supplier.get("years") or supplier.get("experienceYear") or 0)
        supplier_verified = bool(
            supplier.get("verified") or
            supplier.get("isGoldSupplier") or
            raw.get("isTradeAssurance") or
            False
        )

    return {
        "product_id": str(raw.get("id") or raw.get("offerId") or raw.get("product_id") or ""),
        "product_name": raw.get("title") or raw.get("subject") or raw.get("product_name") or "",
        "product_name_cn": raw.get("title") or raw.get("subject") or "",
        "product_url": raw.get("url") or raw.get("detailUrl") or raw.get("product_url") or "",
        "image_url": raw.get("image") or raw.get("imageUrl") or raw.get("imgUrl") or "",
        "source_platform": "1688",
        "price": price_avg,
        "price_currency": "CNY",
        "price_usd": round(price_avg * cny_rate, 2),
        "price_range_min": price_min,
        "price_range_max": price_max,
        "total_sold": parse_count(raw.get("sold") or raw.get("monthSold") or raw.get("tradeCount") or 0),
        "order_count": parse_count(raw.get("orderCount") or raw.get("quantitySumMonth") or 0),
        "review_count": parse_count(raw.get("reviews") or raw.get("ratingCount") or 0),
        "rating": float(raw.get("rating") or raw.get("starLevel") or 0),
        "moq": moq,
        "seller_name": supplier_name,
        "seller_url": raw.get("shopUrl") or raw.get("companyUrl") or "",
        "seller_location": supplier_location,
        "supplier_years": supplier_years,
        "supplier_verified": supplier_verified,
        "search_keyword": keyword,
        "discovered_at": datetime.now(timezone.utc).isoformat(),
    }


def search_keyword(client, keyword, max_pages, price_min, price_max, sort_type, actor_id):
    """Run Apify actor for a single keyword search on 1688."""
    print(f"  Searching: '{keyword}' (pages={max_pages}, actor={actor_id})")

    run_input = {
        "keyword": keyword,
        "searchKeyword": keyword,
        "maxPages": max_pages,
        "pages": max_pages,
    }

    if price_min is not None:
        run_input["priceStart"] = price_min
        run_input["minPrice"] = price_min
    if price_max is not None:
        run_input["priceEnd"] = price_max
        run_input["maxPrice"] = price_max
    if sort_type and sort_type != "default":
        run_input["sortType"] = sort_type
        run_input["sort"] = sort_type

    try:
        run = client.actor(actor_id).call(
            run_input=run_input,
            timeout_secs=600,
        )
        items = list(client.dataset(run["defaultDatasetId"]).iterate_items())
        print(f"  Found {len(items)} products for '{keyword}'")
        return items
    except Exception as e:
        print(f"  ERROR with actor {actor_id} for '{keyword}': {e}")
        return None


def search_with_fallback(client, keyword, max_pages, price_min, price_max, sort_type):
    """Try actors in order until one works."""
    for actor_id in ACTORS:
        items = search_keyword(client, keyword, max_pages, price_min, price_max, sort_type, actor_id)
        if items is not None:
            return items
        print(f"  Trying fallback actor...")
    print(f"  WARNING: All actors failed for '{keyword}'")
    return []


def deduplicate(products):
    """Remove duplicate products, keeping first occurrence."""
    seen = set()
    unique = []
    for p in products:
        pid = p["product_id"]
        if pid and pid not in seen:
            seen.add(pid)
            unique.append(p)
        elif not pid:
            key = f"{p['product_name']}|{p['seller_name']}"
            if key not in seen:
                seen.add(key)
                unique.append(p)
    return unique


def load_progress():
    if PROGRESS_FILE.exists():
        with open(PROGRESS_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"completed_keywords": [], "products": []}


def save_progress(progress):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(progress, f, indent=2, ensure_ascii=False)


def main():
    parser = argparse.ArgumentParser(description="Search 1688/Taobao for products via Apify")
    parser.add_argument("--keywords", required=True, help="Comma-separated search keywords")
    parser.add_argument("--max-pages", type=int, default=3, help="Max pages per keyword (default: 3)")
    parser.add_argument("--price-min", type=float, default=None, help="Min price filter (CNY)")
    parser.add_argument("--price-max", type=float, default=None, help="Max price filter (CNY)")
    parser.add_argument("--sort", default="default",
                        choices=["default", "price_asc", "price_desc", "sales_desc"],
                        help="Sort order (default: default)")
    parser.add_argument("--resume", action="store_true", help="Resume from previous interrupted run")
    args = parser.parse_args()

    keywords = [k.strip() for k in args.keywords.split(",") if k.strip()]
    if not keywords:
        print("ERROR: No keywords provided.")
        sys.exit(1)

    cny_rate = get_cny_usd_rate()
    client = get_apify_client()
    progress = load_progress() if args.resume else {"completed_keywords": [], "products": []}
    all_products = progress["products"]
    failed_keywords = []

    print(f"1688/Taobao Search")
    print(f"  Keywords: {keywords}")
    print(f"  Max pages: {args.max_pages}")
    print(f"  Price range: {args.price_min or 'any'} - {args.price_max or 'any'} CNY")
    print(f"  CNY/USD rate: {cny_rate}")
    print()

    for keyword in keywords:
        if keyword in progress["completed_keywords"]:
            print(f"  Skipping '{keyword}' (already completed)")
            continue

        raw_items = search_with_fallback(
            client, keyword, args.max_pages,
            args.price_min, args.price_max, args.sort
        )
        if not raw_items:
            failed_keywords.append(keyword)
            continue

        normalized = [normalize_taobao_product(item, keyword, cny_rate) for item in raw_items]
        all_products.extend(normalized)

        progress["completed_keywords"].append(keyword)
        progress["products"] = all_products
        save_progress(progress)

    unique_products = deduplicate(all_products)
    print(f"\nTotal: {len(unique_products)} unique products from {len(keywords)} keywords")
    if failed_keywords:
        print(f"Failed keywords: {failed_keywords}")

    output = {
        "discovered_at": datetime.now(timezone.utc).isoformat(),
        "keywords": keywords,
        "platform": "1688",
        "cny_usd_rate": cny_rate,
        "total_products": len(unique_products),
        "failed_keywords": failed_keywords,
        "products": unique_products,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"Saved to {OUTPUT_FILE}")

    if PROGRESS_FILE.exists():
        PROGRESS_FILE.unlink()

    print("Done.")


if __name__ == "__main__":
    main()
