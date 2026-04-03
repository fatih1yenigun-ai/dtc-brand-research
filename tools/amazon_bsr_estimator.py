"""
Tool: Amazon BSR-to-Sales Estimator
=====================================
Estimates monthly sales from Best Sellers Rank (BSR) using
logarithmic decay formulas per Amazon category.

Same methodology used by JungleScout, Helium10, etc.
When we switch to JungleScout API, this becomes redundant
as they provide sales estimates directly.

Usage:
    from amazon_bsr_estimator import estimate_monthly_sales, estimate_revenue
"""

import math

# ── BSR-to-Sales coefficients per top-level Amazon.com category ─────────────
# Formula: monthly_sales = A * (BSR ^ -B)
# A = scale factor, B = decay exponent
# Calibrated from publicly available BSR-to-sales research and cross-referenced
# with JungleScout/Helium10 published estimates.
#
# Sources: JungleScout Million Dollar Case Study, WebRetailer BSR analysis
CATEGORY_COEFFICIENTS = {
    # Category: (A, B, category_label)
    "All Departments":          (120000, 0.80, "All Departments"),
    "Arts, Crafts & Sewing":    (30000,  0.75, "Arts, Crafts & Sewing"),
    "Automotive":               (50000,  0.78, "Automotive"),
    "Baby":                     (40000,  0.77, "Baby"),
    "Beauty & Personal Care":   (80000,  0.80, "Beauty & Personal Care"),
    "Books":                    (150000, 0.82, "Books"),
    "CDs & Vinyl":              (10000,  0.70, "CDs & Vinyl"),
    "Cell Phones & Accessories": (60000, 0.78, "Cell Phones & Accessories"),
    "Clothing, Shoes & Jewelry": (100000, 0.82, "Clothing, Shoes & Jewelry"),
    "Electronics":              (80000,  0.80, "Electronics"),
    "Garden & Outdoor":         (50000,  0.78, "Garden & Outdoor"),
    "Grocery & Gourmet Food":   (60000,  0.78, "Grocery & Gourmet Food"),
    "Health & Household":       (80000,  0.80, "Health & Household"),
    "Home & Kitchen":           (100000, 0.82, "Home & Kitchen"),
    "Industrial & Scientific":  (30000,  0.75, "Industrial & Scientific"),
    "Kitchen & Dining":         (70000,  0.80, "Kitchen & Dining"),
    "Musical Instruments":      (20000,  0.72, "Musical Instruments"),
    "Office Products":          (50000,  0.78, "Office Products"),
    "Patio, Lawn & Garden":     (50000,  0.78, "Patio, Lawn & Garden"),
    "Pet Supplies":             (60000,  0.78, "Pet Supplies"),
    "Software":                 (10000,  0.70, "Software"),
    "Sports & Outdoors":        (70000,  0.80, "Sports & Outdoors"),
    "Tools & Home Improvement": (60000,  0.78, "Tools & Home Improvement"),
    "Toys & Games":             (80000,  0.80, "Toys & Games"),
    "Video Games":              (30000,  0.75, "Video Games"),
}

# Default if category not found
DEFAULT_COEFFICIENTS = (80000, 0.80)


def estimate_monthly_sales(bsr, category="All Departments"):
    """
    Estimate monthly unit sales from BSR.

    Args:
        bsr: Best Sellers Rank (int, must be >= 1)
        category: Amazon top-level category name

    Returns:
        Estimated monthly sales (int), minimum 1
    """
    if not bsr or bsr < 1:
        return 0

    coeffs = CATEGORY_COEFFICIENTS.get(category, None)
    if coeffs:
        a, b, _ = coeffs
    else:
        # Try partial match
        matched = False
        for cat_name, (a, b, _) in CATEGORY_COEFFICIENTS.items():
            if cat_name.lower() in category.lower() or category.lower() in cat_name.lower():
                matched = True
                break
        if not matched:
            a, b = DEFAULT_COEFFICIENTS

    sales = a * (bsr ** -b)
    return max(1, round(sales))


def estimate_revenue(bsr, price, category="All Departments"):
    """
    Estimate monthly revenue from BSR and price.

    Args:
        bsr: Best Sellers Rank
        price: Product price in USD
        category: Amazon category

    Returns:
        Estimated monthly revenue in USD
    """
    if not price or price <= 0:
        return 0
    sales = estimate_monthly_sales(bsr, category)
    return round(sales * price, 2)


def sales_tier(monthly_sales):
    """Categorize sales volume into tiers for quick assessment."""
    if monthly_sales >= 3000:
        return "Cok Yuksek"
    elif monthly_sales >= 1000:
        return "Yuksek"
    elif monthly_sales >= 300:
        return "Orta"
    elif monthly_sales >= 50:
        return "Dusuk"
    else:
        return "Cok Dusuk"


def get_categories():
    """Return list of available Amazon categories."""
    return sorted(CATEGORY_COEFFICIENTS.keys())


# Quick self-test
if __name__ == "__main__":
    test_cases = [
        (1, "Home & Kitchen"),
        (100, "Home & Kitchen"),
        (1000, "Home & Kitchen"),
        (5000, "Home & Kitchen"),
        (10000, "Home & Kitchen"),
        (50000, "Home & Kitchen"),
        (100000, "Home & Kitchen"),
    ]
    print("BSR-to-Sales Estimates (Home & Kitchen)")
    print(f"{'BSR':>10} | {'Sales/mo':>10} | {'Rev @ $25':>12} | {'Tier':>12}")
    print("-" * 52)
    for bsr, cat in test_cases:
        sales = estimate_monthly_sales(bsr, cat)
        rev = estimate_revenue(bsr, 25.0, cat)
        tier = sales_tier(sales)
        print(f"{bsr:>10,} | {sales:>10,} | ${rev:>10,.0f} | {tier:>12}")
