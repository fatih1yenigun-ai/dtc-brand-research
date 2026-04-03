"""
Tool: Amazon Product & Brand Search via Apify
===============================================
Searches Amazon.com for products and brands using Apify actors.
Designed to be swappable with JungleScout API later.

Input:  APIFY_API_TOKEN in .env, search query
Output: List of product dicts with BSR, price, reviews, estimated sales

Usage (CLI):
    python3 tools/amazon_search.py --keyword "yoga mat" --max-items 50
    python3 tools/amazon_search.py --keyword "yoga mat" --search-type brand

Usage (import):
    from amazon_search import search_amazon_products, search_amazon_brand
"""

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

TMP_DIR = PROJECT_ROOT / ".tmp"
TMP_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = TMP_DIR / "amazon_results.json"

# Apify actors for Amazon search (priority order - try free/cheap ones first)
ACTORS = [
    "junglee/amazon-crawler",
    "vaclavrut/amazon-crawler",
    "epctex/amazon-scraper",
]

# Fallback: lightweight actors that search Amazon search results page
SEARCH_ACTORS = [
    "apify/amazon-product-scraper",
    "junglee/amazon-crawler",
]


def get_apify_client():
    """Initialize Apify client."""
    try:
        from apify_client import ApifyClient
    except ImportError:
        print("ERROR: apify-client not installed. Run: pip install apify-client")
        sys.exit(1)

    token = os.getenv("APIFY_API_TOKEN")
    if not token:
        print("ERROR: APIFY_API_TOKEN not set in .env")
        sys.exit(1)
    return ApifyClient(token)


def _try_actor(client, actor_id, run_input, timeout=300):
    """Run an Apify actor and return results, or None on failure."""
    try:
        run = client.actor(actor_id).call(
            run_input=run_input,
            timeout_secs=timeout,
        )
        items = list(client.dataset(run["defaultDatasetId"]).iterate_items())
        if items:
            return items
    except Exception as e:
        print(f"  Actor {actor_id} failed: {e}")
    return None


def search_amazon_products(keyword, max_items=50, category=None):
    """
    Search Amazon.com for products by keyword.

    Args:
        keyword: Search query (e.g. "yoga mat", "wireless earbuds")
        max_items: Maximum results to return
        category: Optional Amazon category filter

    Returns:
        List of normalized product dicts
    """
    client = get_apify_client()
    items = None

    # Try actors with different input schemas
    inputs_to_try = [
        # junglee/amazon-crawler format
        {
            "searchTerms": [keyword],
            "maxItems": max_items,
            "country": "US",
            "categoryUrl": "",
        },
        # Alternative format
        {
            "keyword": keyword,
            "maxResults": max_items,
            "domain": "amazon.com",
        },
        # Generic search format
        {
            "search": keyword,
            "maxItems": max_items,
            "amazonDomain": "amazon.com",
        },
    ]

    for i, actor_id in enumerate(ACTORS):
        input_schema = inputs_to_try[i] if i < len(inputs_to_try) else inputs_to_try[0]
        print(f"  Trying actor: {actor_id}")
        items = _try_actor(client, actor_id, input_schema)
        if items:
            print(f"  Success: {len(items)} results from {actor_id}")
            break

    if not items:
        print("  WARNING: All actors failed. Returning empty results.")
        return []

    # Normalize results
    products = []
    for raw in items[:max_items]:
        product = normalize_amazon_product(raw, keyword)
        if product.get("title"):  # Skip empty results
            products.append(product)

    return products


def search_amazon_brand(brand_name, max_items=50):
    """
    Search Amazon.com for all products by a specific brand.

    Args:
        brand_name: Brand name to search for
        max_items: Maximum results

    Returns:
        List of normalized product dicts
    """
    # Search with brand name as keyword - Amazon search handles brand filtering
    return search_amazon_products(f"{brand_name}", max_items=max_items)


def normalize_amazon_product(raw, keyword):
    """
    Normalize Apify Amazon actor output to unified schema.
    Handles multiple actor output formats.
    """
    from amazon_bsr_estimator import estimate_monthly_sales, estimate_revenue, sales_tier

    # Extract fields (different actors use different field names)
    title = (
        raw.get("title") or
        raw.get("name") or
        raw.get("productTitle") or
        ""
    )

    price_val = _extract_price(raw)

    rating = float(
        raw.get("stars") or
        raw.get("rating") or
        raw.get("averageRating") or
        raw.get("star") or
        0
    )

    review_count = _parse_int(
        raw.get("reviewsCount") or
        raw.get("reviews") or
        raw.get("numberOfReviews") or
        raw.get("totalReviews") or
        raw.get("ratingsTotal") or
        0
    )

    bsr = _extract_bsr(raw)

    brand = (
        raw.get("brand") or
        raw.get("brandName") or
        raw.get("manufacturer") or
        ""
    )

    category = (
        raw.get("category") or
        raw.get("categoryName") or
        raw.get("mainCategory") or
        raw.get("breadcrumbs", [{}])[0].get("name", "") if isinstance(raw.get("breadcrumbs"), list) and raw.get("breadcrumbs") else
        ""
    )

    asin = (
        raw.get("asin") or
        raw.get("ASIN") or
        raw.get("productId") or
        ""
    )

    image = (
        raw.get("thumbnailImage") or
        raw.get("image") or
        raw.get("imageUrl") or
        raw.get("mainImage") or
        ""
    )

    url = (
        raw.get("url") or
        raw.get("productUrl") or
        raw.get("link") or
        (f"https://amazon.com/dp/{asin}" if asin else "")
    )

    # Sales estimation
    monthly_sales = estimate_monthly_sales(bsr, category) if bsr else 0
    monthly_revenue = estimate_revenue(bsr, price_val, category) if bsr and price_val else 0
    tier = sales_tier(monthly_sales) if monthly_sales else "-"

    # Seller count
    sellers = _parse_int(
        raw.get("sellersCount") or
        raw.get("sellers") or
        raw.get("otherSellers") or
        0
    )

    # FBA / Prime
    is_prime = bool(
        raw.get("isPrime") or
        raw.get("prime") or
        raw.get("isFBA") or
        False
    )

    return {
        "asin": asin,
        "title": title,
        "brand": brand,
        "price": price_val,
        "rating": round(rating, 1),
        "review_count": review_count,
        "bsr": bsr,
        "category": category,
        "monthly_sales_est": monthly_sales,
        "monthly_revenue_est": monthly_revenue,
        "sales_tier": tier,
        "sellers": sellers,
        "is_prime": is_prime,
        "image_url": image,
        "url": url,
        "search_keyword": keyword,
        "discovered_at": datetime.now(timezone.utc).isoformat(),
    }


def _extract_price(raw):
    """Extract price from various formats."""
    # Try direct numeric fields
    for key in ["price", "currentPrice", "salePrice", "listPrice", "buyBoxPrice"]:
        val = raw.get(key)
        if isinstance(val, (int, float)) and val > 0:
            return round(float(val), 2)
        if isinstance(val, str):
            cleaned = val.replace("$", "").replace(",", "").strip()
            try:
                p = float(cleaned)
                if p > 0:
                    return round(p, 2)
            except (ValueError, TypeError):
                continue

    # Try nested price object
    price_obj = raw.get("price", {})
    if isinstance(price_obj, dict):
        for key in ["value", "current", "sale", "amount"]:
            val = price_obj.get(key)
            if val:
                try:
                    return round(float(val), 2)
                except (ValueError, TypeError):
                    continue

    return 0.0


def _extract_bsr(raw):
    """Extract Best Sellers Rank from various formats."""
    # Direct BSR field
    for key in ["bsr", "bestSellersRank", "salesRank", "rank", "bestSellerRank"]:
        val = raw.get(key)
        if isinstance(val, (int, float)) and val > 0:
            return int(val)
        if isinstance(val, str):
            try:
                return int(val.replace(",", "").replace("#", "").strip())
            except (ValueError, TypeError):
                continue

    # Nested BSR (some actors return [{category, rank}])
    bsr_list = raw.get("bestSellersRanks") or raw.get("salesRanks") or raw.get("bsrList") or []
    if isinstance(bsr_list, list) and bsr_list:
        for entry in bsr_list:
            if isinstance(entry, dict):
                rank = entry.get("rank") or entry.get("position")
                if rank:
                    try:
                        return int(str(rank).replace(",", "").replace("#", ""))
                    except (ValueError, TypeError):
                        continue

    return 0


def _parse_int(val):
    """Parse various int formats."""
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str):
        cleaned = val.replace(",", "").replace("+", "").replace("#", "").strip()
        try:
            return int(float(cleaned))
        except (ValueError, TypeError):
            return 0
    return 0


def deduplicate(products):
    """Remove duplicate products by ASIN."""
    seen = set()
    unique = []
    for p in products:
        key = p.get("asin") or p.get("title", "")
        if key and key not in seen:
            seen.add(key)
            unique.append(p)
    return unique


# ── CLI ─────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Search Amazon via Apify")
    parser.add_argument("--keyword", required=True, help="Search keyword or brand name")
    parser.add_argument("--max-items", type=int, default=50)
    parser.add_argument("--search-type", choices=["product", "brand"], default="product")
    args = parser.parse_args()

    print(f"Amazon Search ({args.search_type})")
    print(f"  Keyword: {args.keyword}")
    print(f"  Max items: {args.max_items}")
    print()

    if args.search_type == "brand":
        products = search_amazon_brand(args.keyword, args.max_items)
    else:
        products = search_amazon_products(args.keyword, args.max_items)

    products = deduplicate(products)
    print(f"\nTotal: {len(products)} unique products")

    output = {
        "discovered_at": datetime.now(timezone.utc).isoformat(),
        "keyword": args.keyword,
        "search_type": args.search_type,
        "platform": "amazon",
        "total_products": len(products),
        "products": products,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"Saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
