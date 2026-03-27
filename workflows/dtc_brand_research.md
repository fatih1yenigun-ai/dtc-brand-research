# DTC Brand Research Workflow

## Objective
Research and compile DTC ecommerce-native brands for a given niche, with revenue estimation, marketing angle analysis, and exportable Excel output.

## Required Inputs
- **Niche/Category**: e.g., "home & bedding", "cosmetics", "fitness"
- **Filters**: Founded year, traffic threshold, ecommerce-only, etc.
- **Language**: Turkish or English for descriptions
- **Target count**: Number of brands to find

## Tools Used
1. `tools/tqs_calculator.py` — TQS (Traffic Quality Score) calculation
2. `tools/brand_metrics.py` — Brand engagement metrics database
3. `tools/marketing_angle_analyzer.py` — Marketing angle clustering
4. `tools/yuvacim_v3_generator.py` — Excel generation (template for new niches)

## Process

### Step 1: Brand Collection
- Create batch data files with 200-300 brands each
- Format: 5-tuple `(name, website, sub_niche, insight, history)`
- Sources: Web research, DTC directories, TikTok/Meta trends
- Filter for ecommerce-native brands only

### Step 2: Metrics Population
- Add brand metrics to `brand_metrics.py`:
  - `founded`, `traffic`, `bounce_pct`, `pages_visit`, `session_sec`, `aov`, `product_count`, `good_upsell`
- Use SimilarWeb free lookups for top brands
- Set category AOV defaults for brands without individual data

### Step 3: TQS Calculation
```
TQS = (0.4 x bounce_score) + (0.35 x pages_score) + (0.25 x duration_score)
```

Scoring tables:
- **Bounce Rate**: 85%+=1, 75-84%=2, 65-74%=3, 55-64%=4, 45-54%=5, 38-44%=6, 30-37%=7, 22-29%=8, 15-21%=9, <15%=10
- **Pages/Visit**: 1.0=1, 1.1-1.4=2, 1.5-1.9=3, 2.0-2.4=4, 2.5-3.0=5, 3.1-3.8=6, 3.9-4.8=7, 4.9-6.0=8, 6.1-8.0=9, 8+=10
- **Session Duration**: 0-10s=1, 11-20s=2, 21-40s=3, 41-60s=4, 1:01-1:45=5, 1:46-2:30=6, 2:31-3:30=7, 3:31-5:00=8, 5:01-8:00=9, 8+=10

TQS → Base Conversion Rate: 1=0.2%, 2=0.3%, 3=0.8%, 4=1.4%, 5=2.3%, 6=2.8%, 7=3.7%, 8=4.5%, 9=5.5%, 10=7.5%

### Niche Multipliers (applied to base rate):
| Niche | Multiplier | Example TQS=5 |
|-------|-----------|----------------|
| Food & Bev | 2.0x | 4.6% |
| Beauty | 1.5x | 3.45% |
| Fashion | 1.0x | 2.3% (base) |
| Electronics | 0.7x | 1.61% |
| Luxury | 0.5x | 1.15% |

If a category doesn't map to any niche, use the base rate (1.0x).
Category-to-niche mapping is auto-detected from Turkish category names in `tqs_calculator.py`.

### Step 4: Revenue Estimation
```
Estimated Revenue = Traffic x AOV x Conversion Rate / 100
```

AOV Rules:
- Use highest-reviewed product price as base
- If 30+ products AND good upsell/bundling → add 10%+

### Step 5: Excel Generation
14 columns: #, Marka Adı, Web Sitesi, Kategori, Alt Niş, Kuruluş Yılı, Aylık Trafik, Tahmini AOV, TQS, Dönüşüm %, Tahmini Gelir, Özellik, Hikaye, Meta Ads

Features:
- Auto-filter on all columns (sortable by revenue, traffic, founded, AOV)
- Color-coded TQS (green=high, yellow=mid, red=low)
- Clickable website and Meta Ads Library links
- Per-category sheets + combined sheet
- Frozen headers

### Step 6: Marketing Angle Analysis (Optional)
- Cluster brands by marketing angle keywords
- Traffic-weight each angle to find real demand
- Generate demand ranking dashboard

## Output
- Excel file in `research_outputs/` with descriptive name + date
- All deliverables go to cloud services or visible folders

## Lessons Learned
- SimilarWeb blocks direct scraping; use free page lookups or paid API
- Build brand data in batches of 200-300 to avoid token limits
- Always use 5-tuple format for future-proofing
- Category AOV defaults fill gaps when individual data isn't available
