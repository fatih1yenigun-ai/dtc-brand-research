"""
Tool: Taobao Product Scraper (Playwright)
==========================================
Uses headless Chromium to render Taobao product pages and
extract product details + download images.

Input:  Taobao item IDs
Output: .tmp/taobao_product_details.json + .tmp/product_images/

Usage:
    python3 tools/taobao_playwright_scraper.py --ids "996630748998,895209681090"
"""

import argparse
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TMP_DIR = PROJECT_ROOT / ".tmp"
IMG_DIR = TMP_DIR / "product_images"
TMP_DIR.mkdir(exist_ok=True)
IMG_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = TMP_DIR / "taobao_product_details.json"


def download_image(url, filepath):
    """Download an image file."""
    try:
        if not url.startswith("http"):
            url = "https:" + url
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
            "Referer": "https://www.taobao.com/",
        })
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = resp.read()
            if len(data) > 5000:
                with open(filepath, "wb") as f:
                    f.write(data)
                return len(data)
    except Exception as e:
        pass
    return 0


def scrape_product(page, item_id, product_name=""):
    """Scrape a single Taobao product page."""
    print(f"\n  [{item_id}] {product_name}")
    url = f"https://item.taobao.com/item.htm?id={item_id}"

    result = {
        "item_id": str(item_id),
        "product_name_given": product_name,
        "title": "",
        "price": "",
        "shop_name": "",
        "image_urls": [],
        "downloaded_files": [],
        "url": url,
    }

    try:
        # Navigate to the product page
        page.goto(url, wait_until="domcontentloaded", timeout=30000)

        # Wait for content to render
        time.sleep(3)

        # Try to wait for product images to load
        try:
            page.wait_for_selector("img", timeout=10000)
        except:
            pass

        # Scroll down to trigger lazy loading
        page.evaluate("window.scrollTo(0, document.body.scrollHeight / 3)")
        time.sleep(2)
        page.evaluate("window.scrollTo(0, document.body.scrollHeight / 2)")
        time.sleep(1)

        # Extract title - try multiple selectors
        title = ""
        for selector in [
            'h1', '[class*="title"]', '[class*="Title"]',
            'meta[property="og:title"]', 'title',
            '[data-spm*="title"]', '.tb-main-title',
        ]:
            try:
                if selector.startswith("meta"):
                    el = page.query_selector(selector)
                    if el:
                        title = el.get_attribute("content") or ""
                elif selector == "title":
                    title = page.title() or ""
                else:
                    el = page.query_selector(selector)
                    if el:
                        title = el.inner_text().strip()
                if title and len(title) > 5 and "淘宝" not in title and "taobao" not in title.lower():
                    break
                else:
                    title = ""
            except:
                continue

        result["title"] = title
        if title:
            print(f"    Title: {title[:60]}")

        # Extract all images from the page
        all_images = set()

        # Method 1: Get all img src attributes
        img_elements = page.query_selector_all("img")
        for img in img_elements:
            src = img.get_attribute("src") or img.get_attribute("data-src") or ""
            if src and "alicdn.com" in src:
                # Clean up URL - remove size suffixes to get full resolution
                src = re.sub(r'_\d+x\d+\.\w+$', '', src)
                src = re.sub(r'\?.*$', '', src) if "imgextra" in src else src
                if not src.startswith("http"):
                    src = "https:" + src
                # Filter out tiny icons
                if any(x in src for x in ["60-60", "72-72", "76-76", "114-114", "120-120", "144-144", "152-152", "180-180", "favicon", "logo", "icon"]):
                    continue
                all_images.add(src)

        # Method 2: Extract from page source/scripts
        content = page.content()
        # Look for image arrays in JavaScript data
        for match in re.finditer(r'(https?://[a-z0-9]+\.alicdn\.com/(?:imgextra|i[1-4]|bao/uploaded)/[^\s"\'<>]+\.(?:jpg|png|jpeg|webp))', content, re.IGNORECASE):
            img_url = match.group(1)
            img_url = re.sub(r'_\d+x\d+\.\w+$', '', img_url)
            if "60-60" not in img_url and "favicon" not in img_url:
                all_images.add(img_url)

        # Method 3: Look for background images in style attributes
        for match in re.finditer(r'background-image:\s*url\(["\']?(https?://[^"\')\s]+)["\']?\)', content):
            img_url = match.group(1)
            if "alicdn.com" in img_url:
                all_images.add(img_url)

        image_list = sorted(all_images)
        result["image_urls"] = image_list
        print(f"    Found {len(image_list)} unique images")

        # Download images
        product_dir = IMG_DIR / str(item_id)
        product_dir.mkdir(exist_ok=True)

        downloaded = []
        for i, img_url in enumerate(image_list[:20]):
            ext = "jpg"
            if ".png" in img_url:
                ext = "png"
            elif ".webp" in img_url:
                ext = "webp"

            filepath = product_dir / f"product_{i + 1:02d}.{ext}"
            size = download_image(img_url, filepath)
            if size > 0:
                downloaded.append({"file": str(filepath), "url": img_url, "size": size})
                print(f"    Downloaded: {filepath.name} ({size:,} bytes)")

        result["downloaded_files"] = downloaded
        result["images_found"] = len(image_list)
        result["images_downloaded"] = len(downloaded)

        # Also take a screenshot of the product page
        screenshot_path = product_dir / "page_screenshot.png"
        try:
            page.screenshot(path=str(screenshot_path), full_page=False)
            print(f"    Screenshot saved: page_screenshot.png")
            result["screenshot"] = str(screenshot_path)
        except:
            pass

    except Exception as e:
        print(f"    ERROR: {e}")
        result["error"] = str(e)

    return result


def main():
    parser = argparse.ArgumentParser(description="Scrape Taobao products with Playwright")
    parser.add_argument("--ids", required=True, help="Comma-separated item IDs")
    parser.add_argument("--names", default="", help="Pipe-separated product names")
    args = parser.parse_args()

    item_ids = [x.strip() for x in args.ids.split(",") if x.strip()]
    names = [x.strip() for x in args.names.split("|")] if args.names else [""] * len(item_ids)
    while len(names) < len(item_ids):
        names.append("")

    print(f"Taobao Playwright Scraper")
    print(f"  Items: {len(item_ids)}")
    print(f"  Output: {IMG_DIR}")

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"],
        )

        context = browser.new_context(
            user_agent="Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
            viewport={"width": 390, "height": 844},
            locale="zh-CN",
            timezone_id="Asia/Shanghai",
        )

        # Set cookies to appear as a returning visitor
        context.add_cookies([
            {"name": "thw", "value": "cn", "domain": ".taobao.com", "path": "/"},
            {"name": "t", "value": "ab66c5e616ee3a0d3", "domain": ".taobao.com", "path": "/"},
        ])

        page = context.new_page()

        # First visit taobao.com to establish session
        print("\n  Establishing session...")
        try:
            page.goto("https://m.taobao.com/", wait_until="domcontentloaded", timeout=15000)
            time.sleep(2)
        except:
            pass

        for item_id, name in zip(item_ids, names):
            result = scrape_product(page, item_id, name)
            results.append(result)
            time.sleep(2)

        browser.close()

    # Save results
    output = {
        "scraped_at": datetime.now().isoformat(),
        "total_items": len(results),
        "total_images_downloaded": sum(r.get("images_downloaded", 0) for r in results),
        "products": results,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"\n{'='*50}")
    print(f"Results saved to {OUTPUT_FILE}")
    print(f"Images saved to {IMG_DIR}/")
    print(f"\nSummary:")
    for r in results:
        dl = r.get("images_downloaded", 0)
        status = "OK" if dl > 0 else "NO IMAGES"
        title = r.get("title", "")[:40] or r.get("product_name_given", "")[:40]
        print(f"  [{r['item_id']}] {title} — {dl} images [{status}]")


if __name__ == "__main__":
    main()
