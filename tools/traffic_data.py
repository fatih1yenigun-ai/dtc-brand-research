"""
Traffic data collected from SimilarWeb/SEMrush/Grips for key DTC brands.
Format: {domain: (monthly_visits, est_monthly_revenue_usd, source)}
Data as of late 2025 / early 2026.
"""

TRAFFIC_DATA = {
    # ═══ MEGA BRANDS (2M+ visits/month) ═══
    "brooklinen.com": (2_400_000, 27_000_000, "Grips/SimilarWeb Nov 2025"),
    "casper.com": (3_500_000, 35_000_000, "SimilarWeb estimate"),
    "purple.com": (4_200_000, 45_000_000, "SimilarWeb estimate"),
    "saatva.com": (1_963_000, 27_165_000, "Grips Jan 2026"),

    # ═══ LARGE BRANDS (500K-2M visits/month) ═══
    "cozyearth.com": (1_390_000, 12_000_000, "SimilarWeb Jun 2025"),
    "luluandgeorgia.com": (1_600_000, 8_000_000, "SimilarWeb Aug 2024"),
    "bollandbranch.com": (751_000, 6_000_000, "SimilarWeb Aug 2025"),
    "parachutehome.com": (553_000, 5_500_000, "SimilarWeb Sep 2024"),
    "onequince.com": (800_000, 7_000_000, "SimilarWeb estimate"),
    "wearpact.com": (450_000, 2_500_000, "estimate based on ranking"),
    "bearmattress.com": (600_000, 5_000_000, "estimate"),
    "avocadogreenmattress.com": (700_000, 8_000_000, "estimate from B-Corp data"),
    "nestbedding.com": (350_000, 3_000_000, "estimate"),
    "bedsurehome.com": (500_000, 4_000_000, "SimilarWeb estimate"),

    # ═══ MID-SIZE BRANDS (100K-500K visits/month) ═══
    "coyuchi.com": (313_000, 2_535_000, "Grips Dec 2024"),
    "the-citizenry.com": (305_000, 1_800_000, "Grips Mar 2025"),
    "bearaby.com": (211_000, 1_900_000, "estimate"),
    "kytebaby.com": (400_000, 3_500_000, "$14M annual / 12"),
    "adenandanais.com": (350_000, 2_000_000, "estimate"),
    "mellanni.com": (200_000, 1_500_000, "Amazon-primary, DTC estimate"),
    "kassatex.com": (150_000, 800_000, "estimate"),
    "coophomegoods.com": (300_000, 2_000_000, "Amazon-primary, DTC est"),
    "pillowcube.com": (250_000, 2_500_000, "TikTok viral brand est"),
    "ettitude.com": (180_000, 1_200_000, "estimate"),
    "sundaycitizen.co": (120_000, 800_000, "estimate"),
    "sijohome.com": (121_000, 519_000, "Grips Nov 2025"),
    "rileyhome.com": (130_000, 700_000, "estimate"),
    "magiclinen.com": (200_000, 1_200_000, "estimate"),
    "cariloha.com": (180_000, 1_000_000, "estimate"),
    "dreamcloudsleep.com": (400_000, 3_000_000, "estimate"),
    "nolahmattress.com": (300_000, 2_500_000, "estimate"),
    "ostrichpillow.com": (150_000, 600_000, "estimate"),
    "thecompanystore.com": (400_000, 3_500_000, "estimate"),
    "teklafabrics.com": (180_000, 900_000, "estimate"),
    "pigletinbed.com": (120_000, 800_000, "estimate"),
    "dockandbay.com": (100_000, 500_000, "low estimate"),
    "studiomcgee.com": (350_000, 2_000_000, "Netflix brand est"),
    "garnethill.com": (250_000, 1_500_000, "heritage brand est"),
    "peachskinsheets.com": (150_000, 700_000, "estimate"),
    "dreamlandbabyco.com": (130_000, 800_000, "Shark Tank brand est"),
    "nestedbean.com": (100_000, 600_000, "estimate"),
    "roughlinen.com": (80_000, 400_000, "artisan brand est"),
    "holylamborganics.com": (40_000, 200_000, "niche organic est"),
    "redlandcotton.com": (60_000, 300_000, "small brand est"),
    "sheetsgiggles.com": (80_000, 500_000, "Kickstarter brand est"),
    "balooliving.com": (70_000, 400_000, "estimate"),

    # ═══ SMALLER BRANDS (under 100K) ═══
    "primarygoods.com": (50_000, 250_000, "small brand est"),
    "linoto.com": (30_000, 150_000, "artisan est"),
    "loomahome.com": (25_000, 100_000, "small brand"),
    "solorganix.com": (20_000, 80_000, "small brand"),
    "oliveandcrate.com": (15_000, 60_000, "small brand"),
    "whitelotushome.com": (25_000, 120_000, "niche est"),
    "bedthreads.com.au": (200_000, 1_000_000, "AU brand est"),
    "inbed.com.au": (80_000, 400_000, "AU brand est"),
    "sageandclare.com": (60_000, 300_000, "AU brand est"),
    "snowehome.com": (50_000, 250_000, "small brand"),
    "luxome.com": (40_000, 200_000, "small brand"),
    "flaxhome.com": (15_000, 60_000, "small brand"),
    "cultiver.com.au": (50_000, 250_000, "AU artisan"),
    "pizuna.com": (60_000, 200_000, "estimate"),
    "lunaweightedblanket.com": (80_000, 400_000, "Amazon-primary"),
    "greenlandhomefashions.com": (30_000, 150_000, "small brand"),
    "harborhousehome.com": (20_000, 80_000, "small brand"),
    "underthecanopy.com": (15_000, 60_000, "small brand"),
    "riseandfall.co": (30_000, 150_000, "UK brand est"),
    "luxorlinens.com": (25_000, 100_000, "small brand"),
    "cgkunlimited.com": (20_000, 80_000, "small brand"),
    "downandfeatherco.com": (15_000, 70_000, "niche"),
    "natehome.com": (40_000, 200_000, "celebrity small brand"),
}

# Aggregate by marketing angle
def get_traffic_by_angle(results, traffic_data):
    """
    For each marketing angle, sum up the known traffic of its brands.
    Returns dict of angle -> {total_traffic, total_revenue, known_brands, avg_traffic}
    """
    angle_traffic = {}
    for angle_id, data in results.items():
        total_t = 0
        total_r = 0
        known = 0
        for b in data["brands"]:
            domain = b["website"].replace("frfr", "").replace("frfrfr", "")
            if domain in traffic_data:
                visits, rev, _ = traffic_data[domain]
                total_t += visits
                total_r += rev
                known += 1
        avg_t = total_t / known if known > 0 else 0
        angle_traffic[angle_id] = {
            "total_traffic": total_t,
            "total_revenue": total_r,
            "known_brands": known,
            "total_brands": data["count"],
            "avg_traffic": avg_t,
            "avg_revenue": total_r / known if known > 0 else 0,
        }
    return angle_traffic
