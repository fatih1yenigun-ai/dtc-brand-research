"""
Tool: Douyin Product/Video Search via Apify
=============================================
Searches Douyin (Chinese TikTok) for products and videos by keyword
using Apify actors.

Input:  APIFY_API_TOKEN in .env, keywords via CLI args
Output: .tmp/douyin_results.json

Usage:
    python3 tools/douyin_search.py --keywords "便携榨汁机,LED灯"
    python3 tools/douyin_search.py --keywords "portable blender" --max-items 50
    python3 tools/douyin_search.py --keywords "面膜" --sort most_liked
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

OUTPUT_FILE = TMP_DIR / "douyin_results.json"
PROGRESS_FILE = TMP_DIR / "douyin_progress.json"

# Apify actors for Douyin search (in priority order)
ACTORS = [
    "kuaima/douyin-search",
    "cloudcharlestom/douyin-search-scraper",
]


def get_apify_client():
    token = os.getenv("APIFY_API_TOKEN")
    if not token:
        print("ERROR: APIFY_API_TOKEN not set in .env")
        print("See .env.template for setup instructions.")
        sys.exit(1)
    return ApifyClient(token)


def parse_count(val):
    """Parse human-readable counts like '28.12K', '3.34M', '1.2万', '350.5万'."""
    if isinstance(val, (int, float)):
        return int(val)
    if not isinstance(val, str):
        return 0
    cleaned = val.replace(",", "").replace("+", "").strip()
    if not cleaned or cleaned == "-":
        return 0
    # Chinese number suffixes
    cn_multipliers = {"万": 10_000, "亿": 100_000_000}
    for suffix, mult in cn_multipliers.items():
        if cleaned.endswith(suffix):
            try:
                return int(float(cleaned[:-1]) * mult)
            except ValueError:
                return 0
    # Western number suffixes
    multipliers = {"k": 1_000, "m": 1_000_000, "b": 1_000_000_000}
    if cleaned and cleaned[-1].lower() in multipliers:
        try:
            return int(float(cleaned[:-1]) * multipliers[cleaned[-1].lower()])
        except ValueError:
            return 0
    try:
        return int(float(cleaned))
    except ValueError:
        return 0


def normalize_douyin_product(raw, keyword):
    """Normalize Douyin actor output to unified schema."""
    # Douyin results can be videos, products, or user profiles
    # Extract product info from video metadata where possible
    product_info = raw.get("product") or raw.get("commerce") or {}

    return {
        "product_id": str(raw.get("id") or raw.get("aweme_id") or raw.get("video_id") or ""),
        "product_name": (
            product_info.get("title") or
            raw.get("desc") or
            raw.get("title") or
            raw.get("description") or
            ""
        ),
        "product_name_cn": (
            product_info.get("title") or
            raw.get("desc") or
            raw.get("title") or
            ""
        ),
        "product_url": (
            product_info.get("url") or
            raw.get("url") or
            raw.get("share_url") or
            raw.get("video_url") or
            ""
        ),
        "image_url": (
            product_info.get("image") or
            raw.get("cover") or
            raw.get("thumbnail") or
            raw.get("origin_cover") or
            ""
        ),
        "source_platform": "douyin",
        "price": float(product_info.get("price") or raw.get("price") or 0),
        "price_currency": "CNY",
        "price_usd": 0.0,  # Will be calculated in normalize_data.py
        "total_sold": parse_count(product_info.get("sold") or raw.get("sales") or 0),
        "seller_name": (
            raw.get("author", {}).get("nickname") or
            raw.get("author_name") or
            raw.get("nickname") or
            product_info.get("shop_name") or
            ""
        ) if isinstance(raw.get("author"), dict) else str(raw.get("author") or ""),
        "seller_url": (
            raw.get("author", {}).get("url") or
            raw.get("author_url") or
            ""
        ) if isinstance(raw.get("author"), dict) else "",
        "video_likes": parse_count(raw.get("digg_count") or raw.get("likes") or raw.get("like_count") or 0),
        "video_comments": parse_count(raw.get("comment_count") or raw.get("comments") or 0),
        "video_shares": parse_count(raw.get("share_count") or raw.get("shares") or 0),
        "video_views": parse_count(raw.get("play_count") or raw.get("views") or raw.get("view_count") or 0),
        "search_keyword": keyword,
        "discovered_at": datetime.now(timezone.utc).isoformat(),
    }


def search_keyword(client, keyword, max_items, sort_type, actor_id):
    """Run Apify actor for a single keyword search."""
    print(f"  Searching: '{keyword}' (max={max_items}, sort={sort_type}, actor={actor_id})")

    # Build input based on actor expectations
    run_input = {
        "keyword": keyword,
        "searchTermsOrHashtags": keyword,
        "maxItems": max_items,
        "sort": sort_type,
        "type": "video",  # Search for videos (which may contain product links)
    }

    try:
        run = client.actor(actor_id).call(
            run_input=run_input,
            timeout_secs=600,
        )
        items = list(client.dataset(run["defaultDatasetId"]).iterate_items())
        print(f"  Found {len(items)} results for '{keyword}'")
        return items
    except Exception as e:
        print(f"  ERROR with actor {actor_id} for '{keyword}': {e}")
        return None


def search_with_fallback(client, keyword, max_items, sort_type):
    """Try actors in order until one works."""
    for actor_id in ACTORS:
        items = search_keyword(client, keyword, max_items, sort_type, actor_id)
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
    parser = argparse.ArgumentParser(description="Search Douyin for products/videos via Apify")
    parser.add_argument("--keywords", required=True, help="Comma-separated search keywords")
    parser.add_argument("--max-items", type=int, default=100, help="Max results per keyword (default: 100)")
    parser.add_argument("--sort", default="general",
                        choices=["general", "most_liked", "latest"],
                        help="Sort order (default: general)")
    parser.add_argument("--resume", action="store_true", help="Resume from previous interrupted run")
    args = parser.parse_args()

    keywords = [k.strip() for k in args.keywords.split(",") if k.strip()]
    if not keywords:
        print("ERROR: No keywords provided.")
        sys.exit(1)

    client = get_apify_client()
    progress = load_progress() if args.resume else {"completed_keywords": [], "products": []}
    all_products = progress["products"]
    failed_keywords = []

    print(f"Douyin Search")
    print(f"  Keywords: {keywords}")
    print(f"  Max items: {args.max_items}")
    print(f"  Sort: {args.sort}")
    print()

    for keyword in keywords:
        if keyword in progress["completed_keywords"]:
            print(f"  Skipping '{keyword}' (already completed)")
            continue

        raw_items = search_with_fallback(client, keyword, args.max_items, args.sort)
        if not raw_items:
            failed_keywords.append(keyword)
            continue

        normalized = [normalize_douyin_product(item, keyword) for item in raw_items]
        all_products.extend(normalized)

        progress["completed_keywords"].append(keyword)
        progress["products"] = all_products
        save_progress(progress)

    unique_products = deduplicate(all_products)
    print(f"\nTotal: {len(unique_products)} unique results from {len(keywords)} keywords")
    if failed_keywords:
        print(f"Failed keywords: {failed_keywords}")

    output = {
        "discovered_at": datetime.now(timezone.utc).isoformat(),
        "keywords": keywords,
        "platform": "douyin",
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
