# Product Research Workflow

## Objective
Research products in depth across PiPiAds (ad spy), Douyin (Chinese TikTok), and 1688/Taobao (suppliers). Output a structured Google Sheet report with product data, sellers, pricing, and ad intelligence.

## Prerequisites
- `.env` configured with all API keys (see `.env.template`)
- Python dependencies installed: `pip install -r requirements.txt`
- Google Sheets credentials set up (service account JSON)
- PiPiAds API discovered: `python3 tools/pipiads_search.py --discover`

## Quick Start

```bash
# 1. Search all platforms (run in order)
python3 tools/pipiads_search.py --keywords "portable blender" --max-items 50
python3 tools/douyin_search.py --keywords "便携榨汁机" --max-items 100
python3 tools/taobao_search.py --keywords "便携榨汁机" --max-pages 3

# 2. Merge results
python3 tools/normalize_data.py

# 3. Export to Google Sheets
python3 tools/export_to_sheets.py
```

## Detailed Steps

### Step 1: Prepare Keywords
- Start with your product keyword in English
- Translate to Chinese for Douyin and 1688 searches
- Use both English (PiPiAds) and Chinese (Douyin, 1688) versions

### Step 2: PiPiAds — Ad Intelligence
```bash
python3 tools/pipiads_search.py --keywords "LED light,portable blender" --max-items 50 --period 30
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `--keywords` | required | Comma-separated keywords |
| `--max-items` | 50 | Max products per keyword |
| `--period` | 30 | Time window in days |
| `--country` | (all) | Country filter (US, GB, etc.) |
| `--discover` | — | Probe API endpoints (run once) |
| `--resume` | — | Resume interrupted search |

**Output:** `.tmp/pipiads_results.json`
**What you get:** Winning ads, impressions, engagement, advertiser info, ad duration

### Step 3: Douyin — Social Commerce
```bash
python3 tools/douyin_search.py --keywords "便携榨汁机" --max-items 100
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `--keywords` | required | Comma-separated keywords (Chinese recommended) |
| `--max-items` | 100 | Max results per keyword |
| `--sort` | general | Sort: general, most_liked, latest |
| `--resume` | — | Resume interrupted search |

**Output:** `.tmp/douyin_results.json`
**What you get:** Product videos, views, likes, sellers, pricing
**Cost estimate:** ~$1-5 per run on Apify

### Step 4: 1688/Taobao — Suppliers
```bash
python3 tools/taobao_search.py --keywords "便携榨汁机" --max-pages 3
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `--keywords` | required | Comma-separated keywords (Chinese recommended) |
| `--max-pages` | 3 | Pages per keyword (~40 products/page) |
| `--price-min` | (none) | Min price in CNY |
| `--price-max` | (none) | Max price in CNY |
| `--sort` | default | Sort: default, price_asc, price_desc, sales_desc |
| `--resume` | — | Resume interrupted search |

**Output:** `.tmp/taobao_results.json`
**What you get:** Supplier pricing, MOQ, location, verification status, order counts
**Cost estimate:** ~$1-3 per run on Apify

### Step 5: Normalize & Merge
```bash
python3 tools/normalize_data.py
```

- Merges all platform data into unified schema
- Calculates CNY→USD prices (rate from .env or default 0.14)
- Finds cross-platform matches using Chinese text token overlap
- Sorts: cross-platform matches first, then by total sold

**Output:** `.tmp/normalized_products.json`

### Step 6: Export to Google Sheets
```bash
python3 tools/export_to_sheets.py
```

Creates 5 tabs:
1. **All Products** — master view across all platforms
2. **PiPiAds Results** — ad-specific data
3. **Douyin Results** — social commerce data
4. **1688 Suppliers** — sourcing data
5. **Summary** — per-keyword aggregation

**Fallback:** `python3 tools/export_to_sheets.py --excel-only` saves to `.tmp/product_research.xlsx`

## Input Types

### Keyword Search
Provide keywords directly via `--keywords`. Use English for PiPiAds, Chinese for Douyin/1688.

### Product URL/Image Search
Currently supported via PiPiAds web interface (image search). For programmatic image search, the PiPiAds API discovery may reveal an image search endpoint.

## Edge Cases & Learned Constraints
<!-- Updated as issues are discovered -->
- Apify actors may change input schemas — check actor page if searches return empty
- 1688 prices are in CNY; conversion rate should be updated periodically
- Douyin results are primarily videos, not pure product listings — product info is extracted from video metadata
- Chinese text matching uses character bigrams; very short product names may produce false matches
- Google Sheets API allows 100 requests per 100 seconds — batch writes handle this

## Troubleshooting

**PiPiAds "No config found" error:**
Run `python3 tools/pipiads_search.py --discover` first to probe API endpoints.

**Apify actor fails:**
The tools try fallback actors automatically. If all fail, check your `APIFY_API_TOKEN` and the actor pages on apify.com for changes.

**Google Sheets export fails:**
Test connection: `python3 tools/export_to_sheets.py --test`
Fallback: `python3 tools/export_to_sheets.py --excel-only`

**Interrupted search:**
All search tools support `--resume` to continue from where they left off.
