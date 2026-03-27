"""
Tool: Taobao Product Scraper
==============================
Fetches product details and images from Taobao using multiple
API approaches to bypass anti-scraping protections.

Input:  Taobao item IDs
Output: .tmp/taobao_product_details.json + .tmp/product_images/

Usage:
    python3 tools/taobao_scraper.py --ids "996630748998,895209681090"
"""

import argparse
import json
import os
import re
import sys
import time
import urllib.request
import urllib.parse
from pathlib import Path

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TMP_DIR = PROJECT_ROOT / ".tmp"
IMG_DIR = TMP_DIR / "product_images"
TMP_DIR.mkdir(exist_ok=True)
IMG_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = TMP_DIR / "taobao_product_details.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Referer": "https://m.taobao.com/",
}

# Multiple API endpoints to try
API_ENDPOINTS = [
    # Taobao intl detail API (used by overseas version)
    {
        "name": "intl_h5_detail",
        "url": "https://h5api.m.taobao.com/h5/mtop.taobao.detail.getdetail/6.0/",
        "params": lambda item_id: {
            "jsv": "2.7.2",
            "appKey": "12574478",
            "t": str(int(time.time() * 1000)),
            "api": "mtop.taobao.detail.getdetail",
            "v": "6.0",
            "isSec": "0",
            "ecode": "0",
            "timeout": "10000",
            "ttid": "2022@taobao_h5_10.7.0",
            "type": "jsonp",
            "dataType": "jsonp",
            "callback": "mtopjsonp1",
            "data": json.dumps({"itemNumId": str(item_id)}),
        },
    },
    # Taobao global detail API
    {
        "name": "global_detail",
        "url": "https://h5api-intl.m.taobao.com/h5/mtop.taobao.detail.getdetail/6.0/",
        "params": lambda item_id: {
            "jsv": "2.7.2",
            "appKey": "12574478",
            "api": "mtop.taobao.detail.getdetail",
            "v": "6.0",
            "type": "jsonp",
            "dataType": "jsonp",
            "callback": "mtopjsonp1",
            "data": json.dumps({"itemNumId": str(item_id)}),
        },
    },
]


def fetch_url(url, headers=None, timeout=15):
    """Fetch URL using urllib (no external deps needed)."""
    req = urllib.request.Request(url, headers=headers or HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.read().decode("utf-8", errors="replace"), resp.status
    except Exception as e:
        return str(e), 0


def fetch_with_requests(url, headers=None, params=None, timeout=15):
    """Fetch URL using requests library."""
    if not HAS_REQUESTS:
        return None, 0
    try:
        resp = requests.get(url, headers=headers or HEADERS, params=params, timeout=timeout)
        return resp.text, resp.status_code
    except Exception as e:
        return str(e), 0


def parse_jsonp(text):
    """Extract JSON from JSONP response."""
    match = re.search(r'mtopjsonp\d*\((.*)\)', text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            pass
    # Try direct JSON parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    return None


def extract_images_from_html(html, item_id):
    """Extract image URLs from raw HTML/JS content."""
    images = set()
    # Pattern 1: alicdn image URLs
    for match in re.finditer(r'(https?://[a-z0-9]+\.alicdn\.com/[^\s"\'<>]+\.(?:jpg|png|jpeg|webp))', html, re.IGNORECASE):
        url = match.group(1)
        # Filter out tiny icons and tracking pixels
        if "32x32" not in url and "16x16" not in url and "favicon" not in url and "logo" not in url:
            images.add(url)
    # Pattern 2: //img.alicdn.com paths
    for match in re.finditer(r'(//img\.alicdn\.com/[^\s"\'<>]+\.(?:jpg|png|jpeg|webp))', html, re.IGNORECASE):
        images.add("https:" + match.group(1))
    # Pattern 3: imgextra URLs (product detail images)
    for match in re.finditer(r'(https?://img\.alicdn\.com/imgextra/[^\s"\'<>]+\.(?:jpg|png|jpeg|webp))', html, re.IGNORECASE):
        images.add(match.group(1))
    return list(images)


def extract_product_info(data):
    """Extract product info from API response."""
    if not data or not isinstance(data, dict):
        return None

    result = data.get("data", {})
    item = result.get("item", {}) or result.get("itemInfoModel", {})

    if not item:
        # Try nested structures
        for key in ["data", "result", "model"]:
            if key in result and isinstance(result[key], dict):
                item = result[key].get("item", {}) or result[key]
                break

    if not item:
        return None

    title = item.get("title") or item.get("itemTitle") or ""
    images = item.get("images") or item.get("itemImages") or item.get("picsPath") or []
    price = item.get("price") or item.get("reservePrice") or ""

    # Get main images
    if isinstance(images, list):
        image_urls = []
        for img in images:
            if isinstance(img, str):
                url = img if img.startswith("http") else "https:" + img
                image_urls.append(url)
            elif isinstance(img, dict):
                url = img.get("url") or img.get("src") or ""
                if url:
                    url = url if url.startswith("http") else "https:" + url
                    image_urls.append(url)
    else:
        image_urls = []

    return {
        "title": title,
        "price": price,
        "images": image_urls,
        "raw_item_keys": list(item.keys()) if isinstance(item, dict) else [],
    }


def try_mobile_page(item_id):
    """Try to scrape the mobile product page directly."""
    urls = [
        f"https://m.intl.taobao.com/detail/detail.html?id={item_id}",
        f"https://item.taobao.com/item.htm?id={item_id}",
        f"https://h5.m.taobao.com/awp/core/detail.htm?id={item_id}",
    ]

    all_images = []
    title = ""

    for url in urls:
        html, status = fetch_url(url)
        if not html or status == 0:
            continue

        # Try to extract title from meta tags or script data
        title_match = re.search(r'"title"\s*:\s*"([^"]+)"', html)
        if title_match:
            title = title_match.group(1)

        # Extract images
        images = extract_images_from_html(html, item_id)
        all_images.extend(images)

        if title or images:
            break

        time.sleep(0.5)

    return title, list(set(all_images))


def try_apis(item_id):
    """Try multiple API endpoints."""
    for api in API_ENDPOINTS:
        print(f"    Trying {api['name']}...")
        params = api["params"](item_id)

        if HAS_REQUESTS:
            text, status = fetch_with_requests(api["url"], params=params)
        else:
            query = urllib.parse.urlencode(params)
            full_url = f"{api['url']}?{query}"
            text, status = fetch_url(full_url)

        if not text or status == 0:
            continue

        data = parse_jsonp(text)
        if data:
            info = extract_product_info(data)
            if info and (info["title"] or info["images"]):
                return info

        time.sleep(0.5)

    return None


def try_sku_api(item_id):
    """Try Taobao SKU/item info API."""
    url = f"https://acs.m.taobao.com/h5/mtop.taobao.detail.getdesc/6.0/"
    params = {
        "jsv": "2.7.2",
        "appKey": "12574478",
        "api": "mtop.taobao.detail.getdesc",
        "v": "6.0",
        "type": "jsonp",
        "dataType": "jsonp",
        "callback": "mtopjsonp2",
        "data": json.dumps({"id": str(item_id), "type": "1"}),
    }

    if HAS_REQUESTS:
        text, status = fetch_with_requests(url, params=params)
    else:
        query = urllib.parse.urlencode(params)
        full_url = f"{url}?{query}"
        text, status = fetch_url(full_url)

    if text:
        images = extract_images_from_html(text, item_id)
        return images
    return []


def try_share_page_images(item_id):
    """Try to get images from the share/social preview page."""
    # Taobao generates social preview pages with product images
    url = f"https://a.m.taobao.com/i{item_id}.htm"
    html, status = fetch_url(url)
    if html:
        images = extract_images_from_html(html, item_id)
        # Also check for og:image meta tag
        og_match = re.search(r'property="og:image"\s+content="([^"]+)"', html)
        if og_match:
            images.insert(0, og_match.group(1))
        return images
    return []


def download_image(url, filepath):
    """Download an image to the specified path."""
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": HEADERS["User-Agent"],
            "Referer": "https://m.taobao.com/",
        })
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = resp.read()
            if len(data) > 1000:  # Skip tiny files (likely errors)
                with open(filepath, "wb") as f:
                    f.write(data)
                return True
    except Exception as e:
        print(f"      Download failed: {e}")
    return False


def scrape_product(item_id, product_name=""):
    """Scrape a single product using all available methods."""
    print(f"  [{item_id}] {product_name}")

    title = ""
    all_images = []

    # Method 1: Try APIs
    info = try_apis(item_id)
    if info:
        title = info.get("title", "")
        all_images.extend(info.get("images", []))
        if title:
            print(f"    API title: {title}")

    # Method 2: Try mobile pages
    if not title or not all_images:
        page_title, page_images = try_mobile_page(item_id)
        if page_title and not title:
            title = page_title
        all_images.extend(page_images)

    # Method 3: Try share page
    share_images = try_share_page_images(item_id)
    all_images.extend(share_images)

    # Method 4: Try desc API for detail images
    desc_images = try_sku_api(item_id)
    all_images.extend(desc_images)

    # Deduplicate images
    seen = set()
    unique_images = []
    for img in all_images:
        # Normalize URL
        img = img.split("?")[0] if "?" in img and "alicdn" in img else img
        if img not in seen and "alicdn.com" in img:
            seen.add(img)
            unique_images.append(img)

    # Download images
    downloaded = []
    product_dir = IMG_DIR / str(item_id)
    product_dir.mkdir(exist_ok=True)

    for i, img_url in enumerate(unique_images[:20]):  # Max 20 images per product
        ext = "jpg"
        if ".png" in img_url:
            ext = "png"
        elif ".webp" in img_url:
            ext = "webp"

        filepath = product_dir / f"image_{i+1:02d}.{ext}"
        if download_image(img_url, filepath):
            downloaded.append(str(filepath))
            print(f"    Downloaded image {i+1}: {filepath.name}")

    print(f"    Total: {len(downloaded)} images downloaded")

    return {
        "item_id": str(item_id),
        "title": title,
        "product_name": product_name,
        "image_urls": unique_images,
        "downloaded_files": downloaded,
        "images_found": len(unique_images),
        "images_downloaded": len(downloaded),
    }


def main():
    parser = argparse.ArgumentParser(description="Scrape Taobao product details and images")
    parser.add_argument("--ids", required=True, help="Comma-separated Taobao item IDs")
    parser.add_argument("--names", default="", help="Comma-separated product names (optional, same order as IDs)")
    args = parser.parse_args()

    item_ids = [x.strip() for x in args.ids.split(",") if x.strip()]
    names = [x.strip() for x in args.names.split("|")] if args.names else [""] * len(item_ids)

    # Pad names if shorter than IDs
    while len(names) < len(item_ids):
        names.append("")

    print(f"Taobao Product Scraper")
    print(f"  Items: {len(item_ids)}")
    print(f"  Output: {IMG_DIR}")
    print()

    results = []
    for item_id, name in zip(item_ids, names):
        result = scrape_product(item_id, name)
        results.append(result)
        time.sleep(1)

    # Save results
    output = {
        "scraped_at": __import__("datetime").datetime.now().isoformat(),
        "total_items": len(results),
        "total_images": sum(r["images_downloaded"] for r in results),
        "products": results,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"\nResults saved to {OUTPUT_FILE}")
    print(f"Images saved to {IMG_DIR}/")

    # Summary
    print(f"\nSummary:")
    for r in results:
        status = "OK" if r["images_downloaded"] > 0 else "NO IMAGES"
        print(f"  [{r['item_id']}] {r['product_name'][:40]} - {r['images_downloaded']} images [{status}]")


if __name__ == "__main__":
    main()
