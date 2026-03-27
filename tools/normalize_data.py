"""
Tool: Normalize & Merge Multi-Platform Product Data
=====================================================
Reads results from PiPiAds, Douyin, and 1688 searches,
merges them into a unified schema, and flags cross-platform matches.

Input:  .tmp/pipiads_results.json, .tmp/douyin_results.json, .tmp/taobao_results.json
        (any or all — gracefully handles missing files)
Output: .tmp/normalized_products.json

Usage:
    python3 tools/normalize_data.py
    python3 tools/normalize_data.py --sources pipiads,douyin,1688
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

TMP_DIR = PROJECT_ROOT / ".tmp"
OUTPUT_FILE = TMP_DIR / "normalized_products.json"

SOURCE_FILES = {
    "pipiads": TMP_DIR / "pipiads_results.json",
    "douyin": TMP_DIR / "douyin_results.json",
    "1688": TMP_DIR / "taobao_results.json",
}

DEFAULT_CNY_USD_RATE = 0.14

# Unified schema with empty defaults
EMPTY_PRODUCT = {
    # Identity
    "product_id": "",
    "product_name": "",
    "product_name_cn": "",
    "product_url": "",
    "image_url": "",
    "source_platform": "",
    # Pricing
    "price": 0.0,
    "price_currency": "",
    "price_usd": 0.0,
    "price_range_min": 0.0,
    "price_range_max": 0.0,
    # Sales
    "total_sold": 0,
    "sales_7d": 0,
    "sales_30d": 0,
    "review_count": 0,
    "rating": 0.0,
    "order_count": 0,
    # Seller
    "seller_name": "",
    "seller_url": "",
    "seller_location": "",
    # PiPiAds ad intel
    "ad_impressions": 0,
    "ad_likes": 0,
    "ad_comments": 0,
    "ad_shares": 0,
    "ad_creative_url": "",
    "ad_first_seen": "",
    "ad_last_seen": "",
    "ad_duration_days": 0,
    "advertiser_name": "",
    # Douyin social
    "video_likes": 0,
    "video_comments": 0,
    "video_shares": 0,
    "video_views": 0,
    # 1688 sourcing
    "moq": 0,
    "supplier_years": 0,
    "supplier_verified": False,
    # Meta
    "search_keyword": "",
    "discovered_at": "",
    "cross_platform_match": False,
    "matched_platforms": [],
}


def get_cny_usd_rate():
    rate_str = os.getenv("CNY_USD_RATE")
    if rate_str:
        try:
            return float(rate_str)
        except ValueError:
            pass
    return DEFAULT_CNY_USD_RATE


def load_source(source_name, filepath):
    """Load a source file, return products list or empty."""
    if not filepath.exists():
        print(f"  {source_name}: not found ({filepath.name}) — skipping")
        return []
    with open(filepath, encoding="utf-8") as f:
        data = json.load(f)
    products = data.get("products", [])
    print(f"  {source_name}: loaded {len(products)} products")
    return products


def fill_defaults(product):
    """Ensure all unified schema fields exist with defaults."""
    filled = dict(EMPTY_PRODUCT)
    filled.update({k: v for k, v in product.items() if v is not None})
    return filled


def ensure_usd_price(product, cny_rate):
    """Calculate USD price if not set and currency is CNY."""
    if product["price_usd"] == 0 and product["price"] > 0:
        if product["price_currency"] == "CNY":
            product["price_usd"] = round(product["price"] * cny_rate, 2)
        elif product["price_currency"] in ("USD", ""):
            product["price_usd"] = product["price"]
    return product


def extract_chinese_tokens(text):
    """Extract Chinese characters and common tokens for matching."""
    if not text:
        return set()
    # Keep Chinese characters
    chinese = re.findall(r'[\u4e00-\u9fff]+', text)
    # Keep meaningful latin tokens (3+ chars)
    latin = re.findall(r'[a-zA-Z]{3,}', text.lower())
    # Combine all tokens
    tokens = set()
    for segment in chinese:
        # Split Chinese into 2-char bigrams for matching
        for i in range(len(segment) - 1):
            tokens.add(segment[i:i+2])
        if len(segment) >= 3:
            tokens.add(segment)
    tokens.update(latin)
    return tokens


def find_cross_platform_matches(products):
    """Find products that appear on multiple platforms using token overlap."""
    # Group by platform
    by_platform = {}
    for i, p in enumerate(products):
        platform = p["source_platform"]
        if platform not in by_platform:
            by_platform[platform] = []
        by_platform[platform].append(i)

    if len(by_platform) < 2:
        return products

    # Build token index
    token_index = []
    for i, p in enumerate(products):
        name = p.get("product_name_cn") or p.get("product_name") or ""
        tokens = extract_chinese_tokens(name)
        token_index.append(tokens)

    # Compare across platforms (not within same platform)
    platforms = list(by_platform.keys())
    for pi in range(len(platforms)):
        for pj in range(pi + 1, len(platforms)):
            indices_a = by_platform[platforms[pi]]
            indices_b = by_platform[platforms[pj]]

            for ia in indices_a:
                tokens_a = token_index[ia]
                if len(tokens_a) < 2:
                    continue
                for ib in indices_b:
                    tokens_b = token_index[ib]
                    if len(tokens_b) < 2:
                        continue
                    # Calculate Jaccard-like overlap
                    overlap = len(tokens_a & tokens_b)
                    min_size = min(len(tokens_a), len(tokens_b))
                    if min_size > 0 and overlap / min_size >= 0.5:
                        # Match found
                        for idx in (ia, ib):
                            products[idx]["cross_platform_match"] = True
                            matched = set(products[idx].get("matched_platforms") or [])
                            matched.add(products[ia]["source_platform"])
                            matched.add(products[ib]["source_platform"])
                            products[idx]["matched_platforms"] = sorted(matched)

    return products


def main():
    parser = argparse.ArgumentParser(description="Normalize and merge multi-platform product data")
    parser.add_argument("--sources", default="pipiads,douyin,1688",
                        help="Comma-separated sources to include (default: all)")
    args = parser.parse_args()

    requested_sources = [s.strip() for s in args.sources.split(",") if s.strip()]
    cny_rate = get_cny_usd_rate()

    print("Normalize & Merge Product Data")
    print(f"  Sources: {requested_sources}")
    print(f"  CNY/USD rate: {cny_rate}")
    print()

    all_products = []
    loaded_sources = []

    for source_name in requested_sources:
        if source_name not in SOURCE_FILES:
            print(f"  WARNING: Unknown source '{source_name}' — skipping")
            continue
        products = load_source(source_name, SOURCE_FILES[source_name])
        if products:
            loaded_sources.append(source_name)
            for p in products:
                filled = fill_defaults(p)
                filled = ensure_usd_price(filled, cny_rate)
                all_products.append(filled)

    if not all_products:
        print("\nERROR: No products loaded from any source.")
        print("Run at least one search tool first:")
        print("  python3 tools/pipiads_search.py --keywords '...'")
        print("  python3 tools/douyin_search.py --keywords '...'")
        print("  python3 tools/taobao_search.py --keywords '...'")
        sys.exit(1)

    print(f"\n  Total products before matching: {len(all_products)}")

    # Find cross-platform matches
    all_products = find_cross_platform_matches(all_products)
    matches = sum(1 for p in all_products if p["cross_platform_match"])
    print(f"  Cross-platform matches found: {matches}")

    # Sort: cross-platform matches first, then by total_sold
    all_products.sort(key=lambda p: (
        -int(p["cross_platform_match"]),
        -p.get("total_sold", 0),
    ))

    output = {
        "normalized_at": datetime.now(timezone.utc).isoformat(),
        "sources_loaded": loaded_sources,
        "cny_usd_rate": cny_rate,
        "total_products": len(all_products),
        "cross_platform_matches": matches,
        "products": all_products,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"\nSaved {len(all_products)} products to {OUTPUT_FILE}")
    print("Done.")


if __name__ == "__main__":
    main()
