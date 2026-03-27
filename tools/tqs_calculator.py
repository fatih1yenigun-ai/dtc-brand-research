#!/usr/bin/env python3
"""
TQS (Traffic Quality Score) Calculator
Estimates conversion rate and revenue from website engagement metrics.

Formula:
  TQS = (0.4 × bounce_score) + (0.35 × pages_score) + (0.25 × duration_score)
  Estimated Revenue = Traffic × AOV × Conversion Rate / 100
"""

# ── Bounce Rate → Score ──────────────────────────────────────────────────────
def bounce_score(bounce_rate_pct: float) -> int:
    if bounce_rate_pct >= 85: return 1
    if bounce_rate_pct >= 75: return 2
    if bounce_rate_pct >= 65: return 3
    if bounce_rate_pct >= 55: return 4
    if bounce_rate_pct >= 45: return 5
    if bounce_rate_pct >= 38: return 6
    if bounce_rate_pct >= 30: return 7
    if bounce_rate_pct >= 22: return 8
    if bounce_rate_pct >= 15: return 9
    return 10

# ── Avg Session Duration (seconds) → Score ───────────────────────────────────
def duration_score(seconds: float) -> int:
    if seconds <= 10: return 1
    if seconds <= 20: return 2
    if seconds <= 40: return 3
    if seconds <= 60: return 4
    if seconds <= 105: return 5   # 1:45
    if seconds <= 150: return 6   # 2:30
    if seconds <= 210: return 7   # 3:30
    if seconds <= 300: return 8   # 5:00
    if seconds <= 480: return 9   # 8:00
    return 10

# ── Pages per Visit → Score ──────────────────────────────────────────────────
def pages_score(pages: float) -> int:
    if pages <= 1.0: return 1
    if pages <= 1.4: return 2
    if pages <= 1.9: return 3
    if pages <= 2.4: return 4
    if pages <= 3.0: return 5
    if pages <= 3.8: return 6
    if pages <= 4.8: return 7
    if pages <= 6.0: return 8
    if pages <= 8.0: return 9
    return 10

# ── TQS → Base Conversion Rate ───────────────────────────────────────────────
TQS_TO_BASE_CONVERSION = {
    1: 0.2, 2: 0.3, 3: 0.8, 4: 1.4, 5: 2.3,
    6: 2.8, 7: 3.7, 8: 4.5, 9: 5.5, 10: 7.5,
}

# ── Niche Multipliers ────────────────────────────────────────────────────────
# Each niche has a multiplier applied to the base conversion rate.
# If a niche is not listed here, use base rate (1.0x).
NICHE_MULTIPLIERS = {
    "food_bev":     2.0,   # Food & Beverage
    "beauty":       1.5,   # Beauty & Cosmetics
    "fashion":      1.0,   # Fashion (same as base)
    "electronics":  0.7,   # Electronics & Tech
    "luxury":       0.5,   # Luxury / High-ticket
}

# Map category names (Turkish) to niche keys for auto-detection
CATEGORY_TO_NICHE = {
    # Food & Bev (2.0x)
    "Yiyecek & Atıştırmalık": "food_bev",
    "İçecek & Kahve": "food_bev",
    "Yiyecek & İçecek": "food_bev",
    # Beauty (1.5x)
    "Güzellik & Cilt Bakımı": "beauty",
    "Cilt Bakımı & Güzellik Araçları": "beauty",
    "Saç Bakımı & Saç Sağlığı": "beauty",
    "Vücut Bakımı & Kişisel Hijyen": "beauty",
    "Erkek Bakım & Tıraş": "beauty",
    "Makyaj & Kozmetik Organizeri": "beauty",
    "Parfüm & Ev Kokusu": "beauty",
    "Diş & Ağız Bakımı": "beauty",
    "Kadın Sağlığı & Regl Bakımı": "beauty",
    # Fashion (1.0x) — same as base
    "Moda & Giyim": "fashion",
    "Moda & Giyim (Kadın)": "fashion",
    "Moda & Giyim (Erkek)": "fashion",
    "İç Giyim & Çorap": "fashion",
    "Ayakkabı & Terlik": "fashion",
    "Aksesuar & Takı": "fashion",
    "Ev Giyim & Loungewear": "fashion",
    "Terlik & Ev Ayakkabısı": "fashion",
    "Giyilebilir Battaniye": "fashion",
    # Electronics (0.7x)
    "Teknoloji Aksesuarları": "electronics",
    "Teknoloji & Elektronik": "electronics",
    "Uyku Teknolojisi": "electronics",
    # Luxury (0.5x)
    # (brands/categories tagged as luxury will use this)
}

# Full niche-specific conversion tables (pre-calculated for quick lookup)
NICHE_CONVERSION_TABLES = {
    "food_bev": {
        1: 0.4, 2: 0.6, 3: 1.6, 4: 2.8, 5: 4.6,
        6: 5.6, 7: 7.4, 8: 9.0, 9: 11.0, 10: 15.0,
    },
    "beauty": {
        1: 0.3, 2: 0.45, 3: 1.2, 4: 2.1, 5: 3.45,
        6: 4.2, 7: 5.55, 8: 6.75, 9: 8.25, 10: 11.25,
    },
    "fashion": {
        1: 0.2, 2: 0.3, 3: 0.8, 4: 1.4, 5: 2.3,
        6: 2.8, 7: 3.7, 8: 4.5, 9: 5.5, 10: 7.5,
    },
    "electronics": {
        1: 0.14, 2: 0.21, 3: 0.56, 4: 0.98, 5: 1.61,
        6: 1.96, 7: 2.59, 8: 3.15, 9: 3.85, 10: 5.25,
    },
    "luxury": {
        1: 0.1, 2: 0.15, 3: 0.4, 4: 0.7, 5: 1.15,
        6: 1.4, 7: 1.85, 8: 2.25, 9: 2.75, 10: 3.75,
    },
}


def calc_tqs(bounce_pct: float, pages_per_visit: float, session_seconds: float) -> float:
    bs = bounce_score(bounce_pct)
    ps = pages_score(pages_per_visit)
    ds = duration_score(session_seconds)
    return round((0.4 * bs) + (0.35 * ps) + (0.25 * ds), 1)


def get_niche_from_category(category: str) -> str:
    """Auto-detect niche from Turkish category name. Returns niche key or None for base."""
    return CATEGORY_TO_NICHE.get(category)


def tqs_to_conversion(tqs: float, niche: str = None) -> float:
    """
    Convert TQS to conversion rate, adjusted by niche.
    niche: "food_bev", "beauty", "fashion", "electronics", "luxury", or None for base rate.
    """
    tqs_rounded = max(1, min(10, round(tqs)))
    if niche and niche in NICHE_CONVERSION_TABLES:
        return NICHE_CONVERSION_TABLES[niche][tqs_rounded]
    return TQS_TO_BASE_CONVERSION[tqs_rounded]


def estimate_revenue(traffic: int, aov: float, conversion_pct: float) -> float:
    return traffic * aov * conversion_pct / 100


def full_estimate(traffic: int, aov: float, bounce_pct: float,
                  pages_per_visit: float, session_seconds: float,
                  niche: str = None, category: str = None) -> dict:
    """
    Full revenue estimate with niche-adjusted conversion.
    Pass niche directly ("food_bev", "beauty", etc.) or pass category for auto-detection.
    If neither is provided or category isn't mapped, uses base rate.
    """
    if not niche and category:
        niche = get_niche_from_category(category)

    tqs = calc_tqs(bounce_pct, pages_per_visit, session_seconds)
    conv = tqs_to_conversion(tqs, niche)
    rev = estimate_revenue(traffic, aov, conv)
    multiplier = NICHE_MULTIPLIERS.get(niche, 1.0) if niche else 1.0
    return {
        "traffic": traffic,
        "aov": aov,
        "bounce_score": bounce_score(bounce_pct),
        "pages_score": pages_score(pages_per_visit),
        "duration_score": duration_score(session_seconds),
        "tqs": tqs,
        "niche": niche or "base",
        "niche_multiplier": multiplier,
        "conversion_pct": conv,
        "est_monthly_revenue": round(rev),
        "est_yearly_revenue": round(rev * 12),
    }


# ── AOV estimation helper ───────────────────────────────────────────────────
def estimate_aov(bestseller_price: float, product_count: int = 0,
                 good_upsell: bool = False) -> float:
    """
    Estimate AOV from bestseller price.
    If 30+ products and good upsell/bundling, add 10%+ to bestseller price.
    """
    aov = bestseller_price
    if product_count >= 30 and good_upsell:
        aov *= 1.10
    return round(aov, 2)


# ── Test / Demo ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 60)
    print("NICHE CONVERSION TABLE")
    print("=" * 60)
    print(f"{'TQS':>4} {'Base':>8} {'Food&Bev':>10} {'Beauty':>10} {'Fashion':>10} {'Electro':>10} {'Luxury':>10}")
    for tqs in range(1, 11):
        base = TQS_TO_BASE_CONVERSION[tqs]
        fb = NICHE_CONVERSION_TABLES["food_bev"][tqs]
        be = NICHE_CONVERSION_TABLES["beauty"][tqs]
        fa = NICHE_CONVERSION_TABLES["fashion"][tqs]
        el = NICHE_CONVERSION_TABLES["electronics"][tqs]
        lu = NICHE_CONVERSION_TABLES["luxury"][tqs]
        print(f"{tqs:>4} {base:>7.2f}% {fb:>9.2f}% {be:>9.2f}% {fa:>9.2f}% {el:>9.2f}% {lu:>9.2f}%")

    print("\n" + "=" * 60)
    print("EXAMPLE: Home brand (base rate)")
    print("=" * 60)
    r1 = full_estimate(traffic=2_400_000, aov=120, bounce_pct=42,
                       pages_per_visit=3.5, session_seconds=195)
    for k, v in r1.items(): print(f"  {k}: {v}")

    print("\n" + "=" * 60)
    print("EXAMPLE: Beauty brand (1.5x)")
    print("=" * 60)
    r2 = full_estimate(traffic=500_000, aov=45, bounce_pct=48,
                       pages_per_visit=4.0, session_seconds=180, niche="beauty")
    for k, v in r2.items(): print(f"  {k}: {v}")

    print("\n" + "=" * 60)
    print("EXAMPLE: Food brand (2.0x)")
    print("=" * 60)
    r3 = full_estimate(traffic=300_000, aov=35, bounce_pct=50,
                       pages_per_visit=3.0, session_seconds=90, niche="food_bev")
    for k, v in r3.items(): print(f"  {k}: {v}")

    print("\n" + "=" * 60)
    print("EXAMPLE: Luxury brand (0.5x)")
    print("=" * 60)
    r4 = full_estimate(traffic=100_000, aov=350, bounce_pct=35,
                       pages_per_visit=5.0, session_seconds=300, niche="luxury")
    for k, v in r4.items(): print(f"  {k}: {v}")

    print("\n" + "=" * 60)
    print("EXAMPLE: Auto-detect from Turkish category")
    print("=" * 60)
    r5 = full_estimate(traffic=200_000, aov=60, bounce_pct=45,
                       pages_per_visit=3.5, session_seconds=150,
                       category="Güzellik & Cilt Bakımı")
    for k, v in r5.items(): print(f"  {k}: {v}")
