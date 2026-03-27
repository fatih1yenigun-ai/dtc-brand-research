"""
Tool: PiPiAds Ad Spy Search
============================
Searches PiPiAds for winning ads and products by keyword.
Includes an API discovery mode to probe endpoints.

Input:  PIPIADS_API_KEY in .env, keywords via CLI args
Output: .tmp/pipiads_results.json

Usage:
    python3 tools/pipiads_search.py --discover
    python3 tools/pipiads_search.py --keywords "LED light,portable blender" --max-items 50
    python3 tools/pipiads_search.py --keywords "face roller" --period 30 --country US
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

TMP_DIR = PROJECT_ROOT / ".tmp"
TMP_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = TMP_DIR / "pipiads_results.json"
PROGRESS_FILE = TMP_DIR / "pipiads_progress.json"
DISCOVERY_FILE = TMP_DIR / "pipiads_api_discovery.json"
CONFIG_FILE = TMP_DIR / "pipiads_api_config.json"

# Known PiPiAds API base URLs to probe
BASE_URLS = [
    "https://api.pipiads.com",
    "https://api.pipiads.com/api",
    "https://www.pipiads.com/api",
    "https://open.pipiads.com",
    "https://open.pipiads.com/api",
]

# Known endpoints to probe
ENDPOINTS = [
    "/v1/search",
    "/v1/products/search",
    "/v1/ads/search",
    "/v1/ads",
    "/v1/products",
    "/v2/search",
    "/v2/products/search",
    "/v2/ads/search",
    "/search",
    "/products/search",
    "/ads/search",
    "/ads",
    "/products",
    "/ad/search",
    "/product/search",
]

REQUEST_DELAY = 1.0  # seconds between requests


def get_api_key():
    key = os.getenv("PIPIADS_API_KEY")
    if not key:
        print("ERROR: PIPIADS_API_KEY not set in .env")
        print("See .env.template for setup instructions.")
        sys.exit(1)
    return key


def make_request(url, api_key, method="GET", params=None, json_body=None, timeout=15):
    """Try a request with multiple auth patterns."""
    auth_patterns = [
        {"headers": {"Authorization": f"Bearer {api_key}"}, "label": "Bearer token"},
        {"headers": {"X-API-Key": api_key}, "label": "X-API-Key header"},
        {"headers": {"Authorization": api_key}, "label": "Raw Authorization"},
        {"params_extra": {"api_key": api_key}, "label": "api_key query param"},
        {"params_extra": {"apikey": api_key}, "label": "apikey query param"},
        {"params_extra": {"token": api_key}, "label": "token query param"},
    ]

    results = []
    for auth in auth_patterns:
        headers = auth.get("headers", {})
        headers["Content-Type"] = "application/json"
        headers["Accept"] = "application/json"

        req_params = dict(params or {})
        if "params_extra" in auth:
            req_params.update(auth["params_extra"])

        try:
            if method == "GET":
                resp = requests.get(url, headers=headers, params=req_params, timeout=timeout)
            else:
                resp = requests.post(url, headers=headers, params=req_params, json=json_body, timeout=timeout)

            results.append({
                "url": url,
                "method": method,
                "auth_pattern": auth["label"],
                "status_code": resp.status_code,
                "content_type": resp.headers.get("Content-Type", ""),
                "body_preview": resp.text[:500] if resp.text else "",
                "success": 200 <= resp.status_code < 300,
            })

            if 200 <= resp.status_code < 300:
                return results[-1], resp

            time.sleep(0.5)

        except requests.exceptions.ConnectionError:
            results.append({
                "url": url,
                "method": method,
                "auth_pattern": auth["label"],
                "status_code": -1,
                "error": "Connection refused",
                "success": False,
            })
        except requests.exceptions.Timeout:
            results.append({
                "url": url,
                "method": method,
                "auth_pattern": auth["label"],
                "status_code": -2,
                "error": "Timeout",
                "success": False,
            })
        except Exception as e:
            results.append({
                "url": url,
                "method": method,
                "auth_pattern": auth["label"],
                "status_code": -3,
                "error": str(e),
                "success": False,
            })

    return results, None


def discover_api():
    """Probe PiPiAds API endpoints to find working ones."""
    api_key = get_api_key()
    print("PiPiAds API Discovery")
    print("=" * 50)
    print(f"API Key: {api_key[:8]}...{api_key[-4:]}")
    print(f"Testing {len(BASE_URLS)} base URLs x {len(ENDPOINTS)} endpoints\n")

    all_results = []
    working_endpoints = []
    test_params = {"keyword": "test", "q": "test", "search": "test"}

    for base_url in BASE_URLS:
        print(f"\nProbing: {base_url}")
        for endpoint in ENDPOINTS:
            url = f"{base_url}{endpoint}"
            sys.stdout.write(f"  {endpoint} ... ")
            sys.stdout.flush()

            # Try GET first
            results, resp = make_request(url, api_key, method="GET", params=test_params)
            if isinstance(results, list):
                all_results.extend(results)
            else:
                all_results.append(results)

            if resp is not None:
                print(f"SUCCESS ({results['status_code']}) via {results['auth_pattern']}")
                working_endpoints.append({
                    "base_url": base_url,
                    "endpoint": endpoint,
                    "full_url": url,
                    "method": "GET",
                    "auth_pattern": results["auth_pattern"],
                    "status_code": results["status_code"],
                    "body_preview": results.get("body_preview", ""),
                })
                continue

            # Try POST
            results, resp = make_request(
                url, api_key, method="POST",
                json_body={"keyword": "test", "page": 1, "page_size": 5}
            )
            if isinstance(results, list):
                all_results.extend(results)
            else:
                all_results.append(results)

            if resp is not None:
                print(f"SUCCESS ({results['status_code']}) via POST + {results['auth_pattern']}")
                working_endpoints.append({
                    "base_url": base_url,
                    "endpoint": endpoint,
                    "full_url": url,
                    "method": "POST",
                    "auth_pattern": results["auth_pattern"],
                    "status_code": results["status_code"],
                    "body_preview": results.get("body_preview", ""),
                })
                continue

            # Check if we got interesting non-200 responses (401/403 means endpoint exists)
            interesting = [r for r in (results if isinstance(results, list) else [results])
                          if r.get("status_code") in (401, 403, 405, 422)]
            if interesting:
                print(f"EXISTS but auth issue ({interesting[0]['status_code']})")
            else:
                print("no response")

            time.sleep(0.3)

    # Save full discovery results
    discovery_output = {
        "discovered_at": datetime.now(timezone.utc).isoformat(),
        "api_key_prefix": f"{api_key[:8]}...",
        "working_endpoints": working_endpoints,
        "total_probes": len(all_results),
        "all_results": all_results,
    }
    with open(DISCOVERY_FILE, "w", encoding="utf-8") as f:
        json.dump(discovery_output, f, indent=2, ensure_ascii=False)

    # Summary
    print("\n" + "=" * 50)
    print(f"Discovery complete. {len(working_endpoints)} working endpoint(s) found.")
    print(f"Full results saved to: {DISCOVERY_FILE}")

    if working_endpoints:
        best = working_endpoints[0]
        print(f"\nBest endpoint:")
        print(f"  URL: {best['full_url']}")
        print(f"  Method: {best['method']}")
        print(f"  Auth: {best['auth_pattern']}")
        print(f"  Response preview: {best.get('body_preview', '')[:200]}")

        # Save config for future use
        config = {
            "base_url": best["base_url"],
            "endpoint": best["endpoint"],
            "full_url": best["full_url"],
            "method": best["method"],
            "auth_pattern": best["auth_pattern"],
            "discovered_at": datetime.now(timezone.utc).isoformat(),
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
        print(f"\nConfig saved to: {CONFIG_FILE}")
        print("You can now run searches with --keywords")
    else:
        print("\nNo working endpoints found.")
        print("Options:")
        print("  1. Check if your API key is correct")
        print("  2. Check .tmp/pipiads_api_discovery.json for 401/403 responses")
        print("     (these indicate the endpoint exists but auth failed)")
        print("  3. Contact PiPiAds support for API documentation")
        print("  4. Export data manually from pipiads.com as CSV")

    return working_endpoints


def load_config():
    """Load cached API config from discovery."""
    if not CONFIG_FILE.exists():
        print("ERROR: No API config found. Run with --discover first:")
        print("  python3 tools/pipiads_search.py --discover")
        sys.exit(1)
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def search_products(api_key, config, keyword, max_items=50, period=30, country=""):
    """Search PiPiAds for products by keyword using discovered config."""
    url = config["full_url"]
    method = config["method"]
    auth_pattern = config["auth_pattern"]

    # Build auth headers
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    params = {}

    if auth_pattern == "Bearer token":
        headers["Authorization"] = f"Bearer {api_key}"
    elif auth_pattern == "X-API-Key header":
        headers["X-API-Key"] = api_key
    elif auth_pattern == "Raw Authorization":
        headers["Authorization"] = api_key
    elif auth_pattern == "api_key query param":
        params["api_key"] = api_key
    elif auth_pattern == "apikey query param":
        params["apikey"] = api_key
    elif auth_pattern == "token query param":
        params["token"] = api_key

    # Build search payload
    search_params = {
        "keyword": keyword,
        "page": 1,
        "page_size": min(max_items, 100),
        "period": period,
    }
    if country:
        search_params["country"] = country

    all_items = []
    page = 1

    while len(all_items) < max_items:
        search_params["page"] = page
        print(f"    Page {page} (have {len(all_items)}/{max_items})...")

        try:
            if method == "POST":
                resp = requests.post(url, headers=headers, params=params, json=search_params, timeout=30)
            else:
                req_params = {**params, **search_params}
                resp = requests.get(url, headers=headers, params=req_params, timeout=30)

            if resp.status_code == 429:
                print("    Rate limited, waiting 10s...")
                time.sleep(10)
                continue

            if resp.status_code != 200:
                print(f"    ERROR: HTTP {resp.status_code}: {resp.text[:200]}")
                break

            data = resp.json()

            # Try common response structures
            items = (
                data.get("data", {}).get("items") or
                data.get("data", {}).get("products") or
                data.get("data", {}).get("ads") or
                data.get("data", {}).get("results") or
                data.get("items") or
                data.get("products") or
                data.get("ads") or
                data.get("results") or
                (data.get("data") if isinstance(data.get("data"), list) else None) or
                []
            )

            if not items:
                print(f"    No more results on page {page}")
                break

            all_items.extend(items)
            print(f"    Got {len(items)} items")

            # Check if there are more pages
            total = (
                data.get("data", {}).get("total") or
                data.get("total") or
                data.get("data", {}).get("count") or
                0
            )
            if total and len(all_items) >= total:
                break

            page += 1
            time.sleep(REQUEST_DELAY)

        except requests.exceptions.RequestException as e:
            print(f"    ERROR: {e}")
            break

    return all_items[:max_items]


def normalize_pipiads_product(raw, keyword):
    """Normalize PiPiAds raw item to unified schema."""
    return {
        "product_id": str(raw.get("id") or raw.get("product_id") or raw.get("ad_id") or ""),
        "product_name": raw.get("title") or raw.get("product_name") or raw.get("ad_title") or "",
        "product_name_cn": raw.get("title_cn") or raw.get("title_zh") or "",
        "product_url": raw.get("product_url") or raw.get("url") or raw.get("landing_page") or "",
        "image_url": raw.get("image_url") or raw.get("cover") or raw.get("thumbnail") or "",
        "source_platform": "pipiads",
        "price": float(raw.get("price") or raw.get("product_price") or 0),
        "price_currency": raw.get("currency") or "USD",
        "price_usd": float(raw.get("price") or raw.get("product_price") or 0),
        "total_sold": int(raw.get("total_sold") or raw.get("sales") or raw.get("order_count") or 0),
        "ad_impressions": int(raw.get("impressions") or raw.get("imp") or 0),
        "ad_likes": int(raw.get("likes") or raw.get("like_count") or raw.get("digg") or 0),
        "ad_comments": int(raw.get("comments") or raw.get("comment_count") or 0),
        "ad_shares": int(raw.get("shares") or raw.get("share_count") or 0),
        "ad_creative_url": raw.get("video_url") or raw.get("creative_url") or "",
        "ad_first_seen": raw.get("first_seen") or raw.get("created_at") or "",
        "ad_last_seen": raw.get("last_seen") or raw.get("updated_at") or "",
        "ad_duration_days": int(raw.get("duration") or raw.get("running_days") or 0),
        "advertiser_name": raw.get("advertiser") or raw.get("advertiser_name") or raw.get("shop_name") or "",
        "search_keyword": keyword,
        "discovered_at": datetime.now(timezone.utc).isoformat(),
        "_raw": raw,  # Keep raw data for debugging
    }


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
            key = f"{p['product_name']}|{p['advertiser_name']}"
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
    parser = argparse.ArgumentParser(description="Search PiPiAds for winning ads/products")
    parser.add_argument("--discover", action="store_true", help="Probe API endpoints (run first)")
    parser.add_argument("--keywords", help="Comma-separated search keywords")
    parser.add_argument("--max-items", type=int, default=50, help="Max products per keyword (default: 50)")
    parser.add_argument("--period", type=int, default=30, help="Time period in days (default: 30)")
    parser.add_argument("--country", default="", help="Country filter (e.g., US, GB)")
    parser.add_argument("--resume", action="store_true", help="Resume interrupted run")
    args = parser.parse_args()

    if args.discover:
        discover_api()
        return

    if not args.keywords:
        print("ERROR: --keywords required (or use --discover first)")
        parser.print_help()
        sys.exit(1)

    api_key = get_api_key()
    config = load_config()
    keywords = [k.strip() for k in args.keywords.split(",") if k.strip()]

    print(f"PiPiAds Search")
    print(f"  Endpoint: {config['full_url']}")
    print(f"  Keywords: {keywords}")
    print(f"  Max items: {args.max_items}")
    print(f"  Period: {args.period}d")
    print()

    progress = load_progress() if args.resume else {"completed_keywords": [], "products": []}
    all_products = progress["products"]

    for keyword in keywords:
        if keyword in progress["completed_keywords"]:
            print(f"  Skipping '{keyword}' (already completed)")
            continue

        print(f"  Searching: '{keyword}'")
        raw_items = search_products(api_key, config, keyword, args.max_items, args.period, args.country)
        normalized = [normalize_pipiads_product(item, keyword) for item in raw_items]
        all_products.extend(normalized)
        print(f"  Found {len(normalized)} products for '{keyword}'")

        progress["completed_keywords"].append(keyword)
        progress["products"] = all_products
        save_progress(progress)

    # Remove _raw from final output to save space
    for p in all_products:
        p.pop("_raw", None)

    unique_products = deduplicate(all_products)
    print(f"\nTotal: {len(unique_products)} unique products from {len(keywords)} keywords")

    output = {
        "discovered_at": datetime.now(timezone.utc).isoformat(),
        "keywords": keywords,
        "platform": "pipiads",
        "total_products": len(unique_products),
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
