"""
Chat tool definitions and execution for AI Danışman & Hintli Danışman.
Both Haiku and Sonnet can call these tools from the chat.
"""
import json
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from tqs_calculator import (
    calc_tqs, tqs_to_conversion, estimate_revenue,
    bounce_score, pages_score, duration_score,
    NICHE_CONVERSION_TABLES, TQS_TO_BASE_CONVERSION, NICHE_MULTIPLIERS,
)

# ── Tool definitions for Claude API ──────────────────────────────────────────
TOOL_DEFINITIONS = [
    {
        "name": "tqs_hesapla",
        "description": "Bir web sitesinin TQS (Traffic Quality Score) skorunu ve tahmini gelirini hesaplar. Bounce rate, sayfa/ziyaret, oturum süresi, trafik ve AOV bilgisi gerekir.",
        "input_schema": {
            "type": "object",
            "properties": {
                "traffic": {"type": "integer", "description": "Aylık trafik (ziyaret sayısı)"},
                "bounce_pct": {"type": "number", "description": "Bounce rate yüzdesi (örn: 45.0)"},
                "pages_visit": {"type": "number", "description": "Sayfa/ziyaret (örn: 3.5)"},
                "session_minutes": {"type": "number", "description": "Ortalama oturum süresi dakika (örn: 2.5)"},
                "aov": {"type": "number", "description": "Ortalama sipariş değeri USD (örn: 80)"},
                "niche": {"type": "string", "description": "Niş tipi: food_bev, beauty, fashion, electronics, luxury veya base", "default": "base"},
            },
            "required": ["traffic", "bounce_pct", "pages_visit", "session_minutes", "aov"],
        },
    },
    {
        "name": "marka_arastir",
        "description": "Verilen anahtar kelime/niş için DTC markaları araştırır. Post-2018, e-ticaret doğumlu markalar bulur.",
        "input_schema": {
            "type": "object",
            "properties": {
                "keyword": {"type": "string", "description": "Araştırılacak niş/anahtar kelime (örn: 'bamboo bedding', 'pet toys')"},
                "count": {"type": "integer", "description": "Bulunacak marka sayısı (5-50)", "default": 15},
            },
            "required": ["keyword"],
        },
    },
    {
        "name": "aov_tahmin",
        "description": "Ürün fiyatları ve yorum sayılarına göre yorum-ağırlıklı AOV tahmini yapar.",
        "input_schema": {
            "type": "object",
            "properties": {
                "products": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "price": {"type": "number"},
                            "reviews": {"type": "integer"},
                        },
                    },
                    "description": "Ürün listesi: [{price: 50, reviews: 200}, ...]",
                },
                "total_product_count": {"type": "integer", "description": "Sitedeki toplam ürün sayısı"},
                "good_upsell": {"type": "boolean", "description": "İyi bundling/upsell var mı?"},
            },
            "required": ["products"],
        },
    },
    {
        "name": "nis_donusum_tablosu",
        "description": "Tüm nişler için TQS → dönüşüm oranı tablosunu gösterir.",
        "input_schema": {
            "type": "object",
            "properties": {},
        },
    },
]


# ── Tool execution ───────────────────────────────────────────────────────────
def execute_tool(tool_name: str, tool_input: dict) -> str:
    """Execute a tool and return result as string."""

    if tool_name == "tqs_hesapla":
        session_sec = tool_input["session_minutes"] * 60
        niche = tool_input.get("niche", "base")
        if niche == "base":
            niche = None

        bs = bounce_score(tool_input["bounce_pct"])
        ps = pages_score(tool_input["pages_visit"])
        ds = duration_score(session_sec)
        tqs = calc_tqs(tool_input["bounce_pct"], tool_input["pages_visit"], session_sec)
        conv = tqs_to_conversion(tqs, niche)
        rev = round(estimate_revenue(tool_input["traffic"], tool_input["aov"], conv))
        multiplier = NICHE_MULTIPLIERS.get(niche, 1.0) if niche else 1.0

        return json.dumps({
            "bounce_score": f"{bs}/10",
            "pages_score": f"{ps}/10",
            "duration_score": f"{ds}/10",
            "tqs": f"{tqs}/10",
            "niche": niche or "base (genel)",
            "niche_multiplier": f"{multiplier}x",
            "conversion_rate": f"{conv}%",
            "estimated_monthly_revenue": f"${rev:,}",
            "estimated_yearly_revenue": f"${rev * 12:,}",
        }, ensure_ascii=False)

    elif tool_name == "marka_arastir":
        try:
            from live_researcher import research_brands
            results = research_brands(tool_input["keyword"], tool_input.get("count", 15))
            summary = []
            for r in results:
                summary.append({
                    "brand": r.get("brand", ""),
                    "website": r.get("website", ""),
                    "category": r.get("category", ""),
                    "insight": r.get("insight", ""),
                    "aov": r.get("est_aov", ""),
                    "founded": r.get("founded", ""),
                })
            return json.dumps(summary, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)}, ensure_ascii=False)

    elif tool_name == "aov_tahmin":
        products = tool_input["products"]
        total_reviews = sum(p.get("reviews", 0) for p in products)
        if total_reviews > 0:
            weighted = sum(p["price"] * p.get("reviews", 1) for p in products) / total_reviews
        else:
            weighted = sum(p["price"] for p in products) / len(products)

        total_count = tool_input.get("total_product_count", 0)
        good_upsell = tool_input.get("good_upsell", False)
        final = weighted * 1.10 if total_count >= 30 and good_upsell else weighted

        return json.dumps({
            "review_weighted_aov": f"${weighted:.2f}",
            "upsell_adjustment": "+10%" if total_count >= 30 and good_upsell else "Yok",
            "final_estimated_aov": f"${final:.2f}",
        }, ensure_ascii=False)

    elif tool_name == "nis_donusum_tablosu":
        table = "TQS | Base | Food&Bev(2x) | Beauty(1.5x) | Fashion(1x) | Electronics(0.7x) | Luxury(0.5x)\n"
        table += "--- | --- | --- | --- | --- | --- | ---\n"
        for tqs in range(1, 11):
            base = TQS_TO_BASE_CONVERSION[tqs]
            fb = NICHE_CONVERSION_TABLES["food_bev"][tqs]
            be = NICHE_CONVERSION_TABLES["beauty"][tqs]
            fa = NICHE_CONVERSION_TABLES["fashion"][tqs]
            el = NICHE_CONVERSION_TABLES["electronics"][tqs]
            lu = NICHE_CONVERSION_TABLES["luxury"][tqs]
            table += f"{tqs} | {base}% | {fb}% | {be}% | {fa}% | {el}% | {lu}%\n"
        return table

    return json.dumps({"error": f"Bilinmeyen tool: {tool_name}"})
