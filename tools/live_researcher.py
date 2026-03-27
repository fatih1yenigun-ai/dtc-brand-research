#!/usr/bin/env python3
"""
Live DTC Brand Researcher
Uses Claude API to research brands in any niche on-demand.
"""
import os
import json
import anthropic
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))

# Support both .env and Streamlit secrets
def _get_api_key():
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        try:
            import streamlit as st
            key = st.secrets.get("ANTHROPIC_API_KEY", "")
        except Exception:
            pass
    return key

client = anthropic.Anthropic(api_key=_get_api_key())

RESEARCH_PROMPT = """Sen bir DTC (Direct-to-Consumer) marka araştırmacısısın.

Aşağıdaki niş/anahtar kelime için E-TİCARET DOĞUMLU, post-2018 DTC markaları bul.

NİŞ: {keyword}
MARKA SAYISI: {count}

KURALLAR:
- SADECE 2018 sonrası kurulmuş markalar
- SADECE e-ticaret doğumlu (online büyümüş, TikTok/Instagram/Meta Ads/Shopify/Amazon)
- Geleneksel/perakende markaları DAHİL ETME
- Her marka için gerçek website ver
- Küçük, inovatif markalar da dahil et

Her marka için şu JSON formatında yanıt ver (başka hiçbir şey yazma, sadece JSON array):

```json
[
  {{
    "brand": "Marka Adı",
    "website": "website.com",
    "category": "Kategori (Türkçe)",
    "sub_niche": "Alt niş (Türkçe)",
    "insight": "Öne çıkan özellik, pazarlama açısı, ne ile farklılaşıyor (Türkçe, 1-2 cümle)",
    "history": "Kuruluş yılı, nasıl büyüdü, bootstrap/VC, gelir bilgisi varsa (Türkçe, 1-2 cümle)",
    "est_aov": 85,
    "founded": 2020
  }}
]
```

SADECE JSON array döndür. Başka açıklama yapma. Markdown code block kullanma."""


def research_brands(keyword: str, count: int = 30) -> list[dict]:
    """
    Research DTC brands for a given keyword/niche.
    Returns list of brand dicts.
    """
    prompt = RESEARCH_PROMPT.format(keyword=keyword, count=count)

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )

    text = message.content[0].text.strip()

    # Clean up response - remove markdown code blocks if present
    if text.startswith("```"):
        text = text.split("\n", 1)[1]  # Remove first line
        if text.endswith("```"):
            text = text[:-3]
        elif "```" in text:
            text = text[:text.rfind("```")]

    text = text.strip()

    try:
        brands = json.loads(text)
        return brands
    except json.JSONDecodeError:
        # Try to find JSON array in the response
        start = text.find("[")
        end = text.rfind("]") + 1
        if start >= 0 and end > start:
            try:
                brands = json.loads(text[start:end])
                return brands
            except json.JSONDecodeError:
                pass
        return []


def research_brands_large(keyword: str, total: int = 100, batch_size: int = 30,
                          progress_callback=None) -> list[dict]:
    """
    Research many brands by making multiple API calls.
    progress_callback(current_count, total) is called after each batch.
    """
    all_brands = []
    seen = set()
    batches_needed = (total + batch_size - 1) // batch_size

    for i in range(batches_needed):
        remaining = total - len(all_brands)
        if remaining <= 0:
            break

        count = min(batch_size, remaining)

        # Adjust prompt to avoid duplicates
        if all_brands:
            existing_names = ", ".join(b["brand"] for b in all_brands[-20:])
            keyword_adjusted = f"{keyword} (bu markaları TEKRAR ETME: {existing_names})"
        else:
            keyword_adjusted = keyword

        try:
            batch = research_brands(keyword_adjusted, count)
            for b in batch:
                if b.get("brand", "").lower() not in seen:
                    all_brands.append(b)
                    seen.add(b["brand"].lower())
        except Exception as e:
            print(f"Batch {i+1} error: {e}")

        if progress_callback:
            progress_callback(len(all_brands), total)

    return all_brands


if __name__ == "__main__":
    print("Testing live researcher...")
    results = research_brands("weighted blanket anxiety", 5)
    for r in results:
        print(f"  {r['brand']} | {r['website']} | {r.get('est_aov', '?')} | {r['insight'][:60]}...")
    print(f"\nFound {len(results)} brands")
