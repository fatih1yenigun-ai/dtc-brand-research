import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import sys
import os
import json

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_APP_DIR, "tools"))

from tqs_calculator import (
    calc_tqs, tqs_to_conversion, estimate_revenue,
    bounce_score, pages_score, duration_score,
    get_niche_from_category, NICHE_CONVERSION_TABLES, TQS_TO_BASE_CONVERSION,
    NICHE_MULTIPLIERS,
)
from brand_metrics import BRAND_METRICS, CATEGORY_AOV_DEFAULTS
from yuvacim_v3_generator import BRANDS, merge_brands
from chat_tools import TOOL_DEFINITIONS, execute_tool
from amazon_search import search_amazon_products, search_amazon_brand, deduplicate as amazon_deduplicate
from amazon_bsr_estimator import estimate_monthly_sales, estimate_revenue as bsr_estimate_revenue, sales_tier, get_categories as get_amazon_categories

# ── Load all batch data ──────────────────────────────────────────────────────
for i in range(1, 20):
    try:
        mod = __import__(f"yuvacim_v3_batch{i}")
        merge_brands(mod.BATCH_DATA)
    except ImportError:
        break


# ── Supabase Persistence ─────────────────────────────────────────────────────
from db import (
    load_folders as db_load_folders,
    create_folder as db_create_folder,
    delete_folder as db_delete_folder,
    load_brands as db_load_brands,
    save_brands_bulk as db_save_brands_bulk,
    remove_brands_by_name as db_remove_brands_by_name,
    move_brands as db_move_brands,
    get_all_saved_count as db_get_all_saved_count,
)


# ── API Key Helper ───────────────────────────────────────────────────────────
def get_api_key():
    """Try st.secrets first, fall back to env var."""
    try:
        key = st.secrets["ANTHROPIC_API_KEY"]
        if key:
            return key
    except (KeyError, FileNotFoundError):
        pass
    from dotenv import load_dotenv
    load_dotenv()
    key = os.getenv("ANTHROPIC_API_KEY")
    if key:
        return key
    return None


# ── Build DataFrame ──────────────────────────────────────────────────────────
@st.cache_data
def load_brands():
    rows = []
    for cat, brands in BRANDS.items():
        niche = get_niche_from_category(cat)
        for b in brands:
            name = b[0]
            website = b[1]
            sub_niche = b[2]
            insight = b[3]
            history = b[4] if len(b) > 4 else "-"

            m = BRAND_METRICS.get(name, {})
            founded = m.get("founded", None)
            traffic = m.get("traffic", None)
            bounce = m.get("bounce_pct", None)
            pages = m.get("pages_visit", None)
            sess = m.get("session_sec", None)
            aov = m.get("aov", CATEGORY_AOV_DEFAULTS.get(cat, 80))

            tqs = None
            conv = None
            est_rev = None
            if bounce and pages and sess:
                tqs = calc_tqs(bounce, pages, sess)
                conv = tqs_to_conversion(tqs, niche)
                if traffic:
                    est_rev = round(estimate_revenue(traffic, aov, conv))

            rows.append({
                "Marka": name,
                "Web Sitesi": website,
                "Kategori": cat,
                "Alt Niş": sub_niche,
                "Kuruluş Yılı": founded,
                "Aylık Trafik": traffic,
                "AOV ($)": aov,
                "TQS": tqs,
                "Dönüşüm %": conv,
                "Tahmini Aylık Gelir ($)": est_rev,
                "Öne Çıkan Özellik": insight,
                "Marka Hikayesi": history,
                "Niş Çarpanı": NICHE_MULTIPLIERS.get(niche, 1.0) if niche else 1.0,
            })
    return pd.DataFrame(rows)


# ── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DTC Marka Araştırma Paneli",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0D1B2A;
        margin-bottom: 0;
    }
    .sub-header {
        font-size: 1rem;
        color: #6c757d;
        margin-top: -10px;
        margin-bottom: 20px;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 12px;
        color: white;
        text-align: center;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
    }
    .metric-label {
        font-size: 0.85rem;
        opacity: 0.85;
    }
    .stDataFrame {
        border-radius: 8px;
    }
    div[data-testid="stSidebar"] {
        background-color: #f8f9fa;
    }
</style>
""", unsafe_allow_html=True)

df = load_brands()

# ── Session state init ───────────────────────────────────────────────────────
if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = []
if "haiku_messages" not in st.session_state:
    st.session_state.haiku_messages = []

# ── Sidebar Navigation ──────────────────────────────────────────────────────
st.sidebar.title("DTC Araştırma Paneli")
page = st.sidebar.radio(
    "Sayfa Seç",
    ["Canlı Araştırma", "Amazon Araştırma", "Kaydedilenler", "Hintli Danışman", "Marka Tarayıcı",
     "Kategori Analizi", "TQS Hesaplayıcı", "AOV Tahminleyici", "Pazarlama Açıları"],
    index=0,
)

# Sidebar folder count
st.sidebar.divider()
folder_names = db_load_folders()
total_saved = db_get_all_saved_count()
st.sidebar.metric("Kayıtlı Marka", total_saved)
st.sidebar.metric("Klasör", len(folder_names))

# ── Sidebar AI Danışman (Haiku) ─────────────────────────────────────────────
st.sidebar.divider()
with st.sidebar.expander("💬 AI Danışman", expanded=False):
    # Display haiku chat history
    for msg in st.session_state.haiku_messages:
        role_label = "Sen" if msg["role"] == "user" else "AI"
        st.markdown(f"**{role_label}:** {msg['content']}")

    haiku_input = st.text_input("Soru sor...", key="haiku_input", placeholder="Kısa bir soru sor...")
    if st.button("Gönder", key="haiku_send", use_container_width=True):
        if haiku_input:
            st.session_state.haiku_messages.append({"role": "user", "content": haiku_input})

            api_key = get_api_key()
            if not api_key:
                st.error("API anahtarı bulunamadı. Streamlit Secrets veya .env dosyasına ANTHROPIC_API_KEY ekleyin.")
            else:
                try:
                    import anthropic
                    client = anthropic.Anthropic(api_key=api_key)

                    system_prompt = """Sen Yuvacım markası için çalışan bir DTC (Direct-to-Consumer) e-ticaret danışmanısın.
Türkçe yanıt ver. Kısa ve öz ol.
Türkiye pazarını iyi bilirsin. Pazarlama açıları, ürün fikirleri, reklam stratejileri konusunda fikir verirsin.
DTC markalar, AOV optimizasyonu, conversion rate, TQS skoru, niş seçimi konularında uzmansın.
Yuvacım bir Türk DTC markasıdır. Kullanıcı Fatih, bir DTC girişimcisidir."""

                    messages = [{"role": m["role"], "content": m["content"]}
                                for m in st.session_state.haiku_messages[-20:]]

                    response = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=1000,
                        system=system_prompt,
                        messages=messages,
                    )
                    reply = response.content[0].text
                    st.session_state.haiku_messages.append({"role": "assistant", "content": reply})
                    st.rerun()
                except Exception as e:
                    st.error(f"Hata: {e}")

    if st.session_state.haiku_messages:
        if st.button("Temizle", key="haiku_clear", use_container_width=True):
            st.session_state.haiku_messages = []
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 0: LIVE RESEARCH
# ══════════════════════════════════════════════════════════════════════════════
if page == "Canlı Araştırma":
    st.markdown('<p class="main-header">Canlı Marka Araştırma</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Bir niş/anahtar kelime gir, Claude sana DTC markaları bulsun</p>', unsafe_allow_html=True)

    # Initialize session state for research results
    if "research_results" not in st.session_state:
        st.session_state.research_results = None
    if "research_history" not in st.session_state:
        st.session_state.research_history = []

    col_r1, col_r2, col_r3 = st.columns([3, 1, 1])
    with col_r1:
        keyword = st.text_input(
            "Niş / Anahtar Kelime",
            placeholder="örn: bamboo bedding, pet accessories, skincare serums, weighted blankets...",
        )
    with col_r2:
        brand_count = st.selectbox("Marka Sayısı", [10, 20, 30, 50, 75, 100], index=2)
    with col_r3:
        niche_type = st.selectbox("Niş Tipi", [
            "base (Genel)", "food_bev (Yiyecek 2.0x)", "beauty (Güzellik 1.5x)",
            "fashion (Moda 1.0x)", "electronics (Elektronik 0.7x)", "luxury (Lüks 0.5x)",
        ])

    search_clicked = st.button("Araştır", type="primary", use_container_width=True)

    if search_clicked and keyword:
        niche_key = niche_type.split(" ")[0] if not niche_type.startswith("base") else None

        try:
            from live_researcher import research_brands, research_brands_large

            with st.status(f"'{keyword}' için araştırma yapılıyor...", expanded=True) as status:
                st.write(f"Claude'a {brand_count} marka sorgulanıyor...")

                if brand_count <= 30:
                    results = research_brands(keyword, brand_count)
                else:
                    progress_bar = st.progress(0)
                    def update_progress(current, total):
                        progress_bar.progress(min(current / total, 1.0))
                        st.write(f"{current}/{total} marka bulundu...")

                    results = research_brands_large(keyword, brand_count, 30, update_progress)

                status.update(label=f"{len(results)} marka bulundu!", state="complete")

            st.session_state.research_results = results
            st.session_state.research_history.append({
                "keyword": keyword,
                "count": len(results),
                "niche": niche_key,
            })

        except Exception as e:
            error_msg = str(e)
            if "credit balance" in error_msg.lower() or "billing" in error_msg.lower():
                st.error("API kredisi yetersiz. console.anthropic.com → Plans & Billing'den kredi yükleyin.")
            else:
                st.error(f"Hata: {error_msg}")

    # Display results
    if st.session_state.research_results:
        results = st.session_state.research_results
        st.divider()
        st.subheader(f"{len(results)} Marka Bulundu")

        # Convert to DataFrame
        rows = []
        niche_key = niche_type.split(" ")[0] if not niche_type.startswith("base") else None
        for r in results:
            aov = r.get("est_aov", 80)
            founded = r.get("founded", None)
            meta_url = f"https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=ALL&q={r.get('brand', '')}&search_type=keyword_unordered"
            website = r.get("website", "")
            url = f"https://{website}" if website and not website.startswith("http") else website

            rows.append({
                "Marka": r.get("brand", ""),
                "Web Sitesi": url,
                "Kategori": r.get("category", ""),
                "Alt Niş": r.get("sub_niche", ""),
                "Kuruluş Yılı": founded,
                "AOV ($)": aov,
                "Öne Çıkan Özellik": r.get("insight", ""),
                "Marka Hikayesi": r.get("history", ""),
                "Meta Ads": meta_url,
            })

        res_df = pd.DataFrame(rows)

        # Top stats
        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("Bulunan Marka", len(res_df))
        rc2.metric("Ort. AOV", f"${res_df['AOV ($)'].mean():.0f}")
        rc3.metric("Kategori", res_df["Kategori"].nunique())

        # Display results table
        st.dataframe(
            res_df,
            use_container_width=True,
            height=400,
            column_config={
                "Web Sitesi": st.column_config.LinkColumn("Web Sitesi", display_text="Ziyaret Et"),
                "AOV ($)": st.column_config.NumberColumn("AOV ($)", format="$%d"),
                "Meta Ads": st.column_config.LinkColumn("Meta Ads", display_text="Reklamları Gör"),
                "Öne Çıkan Özellik": st.column_config.TextColumn("Öne Çıkan Özellik", width="large"),
                "Marka Hikayesi": st.column_config.TextColumn("Marka Hikayesi", width="large"),
            },
        )

        # Save to folder section
        st.divider()
        st.markdown("##### Klasöre Kaydet")

        # Select brands
        brand_list = res_df["Marka"].tolist()
        selected_brands = st.multiselect("Kaydetmek istediğin markaları seç", brand_list, key="save_brand_select")

        # Folder selector + new folder
        current_folders = db_load_folders()
        col_fs1, col_fs2, col_fs3 = st.columns([2, 2, 1])
        with col_fs1:
            target_folder = st.selectbox("Klasör seç", current_folders, key="save_target_folder")
        with col_fs2:
            new_folder_name = st.text_input("veya yeni klasör oluştur", placeholder="Yeni klasör adı...", key="new_folder_inline")
        with col_fs3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Oluştur", key="create_folder_inline", use_container_width=True):
                if new_folder_name:
                    if db_create_folder(new_folder_name):
                        st.success(f"'{new_folder_name}' oluşturuldu!")
                        st.rerun()

        col_sv1, col_sv2 = st.columns(2)
        with col_sv1:
            if st.button("Seçilenleri Kaydet", type="primary", use_container_width=True, key="save_checked_btn"):
                if selected_brands:
                    save_folder = new_folder_name if new_folder_name else target_folder
                    if new_folder_name:
                        db_create_folder(new_folder_name)
                    brands_to_save = [res_df[res_df["Marka"] == name].iloc[0].to_dict() for name in selected_brands]
                    added = db_save_brands_bulk(save_folder, brands_to_save)
                    st.success(f"{added} marka '{save_folder}' klasörüne kaydedildi!")
                else:
                    st.warning("Marka seçin")
        with col_sv2:
            if st.button("Tümünü Kaydet", use_container_width=True, key="save_all_btn"):
                save_folder = new_folder_name if new_folder_name else target_folder
                if new_folder_name:
                    db_create_folder(new_folder_name)
                all_brands = [row.to_dict() for _, row in res_df.iterrows()]
                added = db_save_brands_bulk(save_folder, all_brands)
                st.success(f"{added} marka '{save_folder}' klasörüne kaydedildi!")

        # Category breakdown chart
        if len(res_df) > 5:
            cat_counts = res_df["Kategori"].value_counts().reset_index()
            cat_counts.columns = ["Kategori", "Sayı"]
            fig = px.pie(cat_counts, values="Sayı", names="Kategori",
                         title="Kategori Dağılımı", hole=0.4)
            st.plotly_chart(fig, use_container_width=True)

        # Export
        st.divider()
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            csv = res_df.to_csv(index=False).encode("utf-8")
            st.download_button("CSV İndir", csv,
                f"DTC_Arastirma_{keyword.replace(' ', '_')}.csv", "text/csv",
                use_container_width=True)
        with col_e2:
            from io import BytesIO
            buffer = BytesIO()
            res_df.to_excel(buffer, index=False, engine="openpyxl")
            st.download_button("Excel İndir", buffer.getvalue(),
                f"DTC_Arastirma_{keyword.replace(' ', '_')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: AMAZON ARAŞTIRMA (AMAZON RESEARCH)
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Amazon Araştırma":
    st.markdown('<p class="main-header">Amazon Urun Arastirma</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Amazon.com urun ve marka arastirmasi - BSR bazli satis tahmini</p>', unsafe_allow_html=True)

    # Session state
    if "amazon_results" not in st.session_state:
        st.session_state.amazon_results = None
    if "amazon_search_type" not in st.session_state:
        st.session_state.amazon_search_type = "product"

    # Tabs
    tab_product, tab_brand, tab_bsr_calc = st.tabs(["Urun Arama", "Marka Arama", "BSR Hesaplayici"])

    # ── TAB 1: Product Search ────────────────────────────────────────────────
    with tab_product:
        col_a1, col_a2 = st.columns([3, 1])
        with col_a1:
            amz_keyword = st.text_input(
                "Anahtar Kelime",
                placeholder="orn: yoga mat, wireless earbuds, kitchen organizer...",
                key="amz_keyword",
            )
        with col_a2:
            amz_max = st.selectbox("Sonuc Sayisi", [20, 30, 50, 75, 100], index=1, key="amz_max")

        amz_search = st.button("Amazon'da Ara", type="primary", use_container_width=True, key="amz_search_btn")

        if amz_search and amz_keyword:
            try:
                with st.status(f"'{amz_keyword}' icin Amazon'da arama yapiliyor...", expanded=True) as status:
                    st.write("Apify actor calistiriliyor... (30-60 sn surebilir)")
                    results = search_amazon_products(amz_keyword, max_items=amz_max)
                    results = amazon_deduplicate(results)
                    status.update(label=f"{len(results)} urun bulundu!", state="complete")

                st.session_state.amazon_results = results
                st.session_state.amazon_search_type = "product"
            except Exception as e:
                error_msg = str(e)
                if "APIFY_API_TOKEN" in error_msg:
                    st.error("APIFY_API_TOKEN .env dosyasinda bulunamadi. Lutfen ekleyin.")
                elif "apify_client" in error_msg:
                    st.error("apify-client paketi yuklu degil. Calistirin: pip install apify-client")
                else:
                    st.error(f"Hata: {error_msg}")

    # ── TAB 2: Brand Search ──────────────────────────────────────────────────
    with tab_brand:
        col_b1, col_b2 = st.columns([3, 1])
        with col_b1:
            amz_brand = st.text_input(
                "Marka Adi",
                placeholder="orn: Anker, COSRX, Stanley, Hydro Flask...",
                key="amz_brand",
            )
        with col_b2:
            amz_brand_max = st.selectbox("Sonuc Sayisi", [20, 30, 50, 75, 100], index=2, key="amz_brand_max")

        amz_brand_search = st.button("Markayi Ara", type="primary", use_container_width=True, key="amz_brand_btn")

        if amz_brand_search and amz_brand:
            try:
                with st.status(f"'{amz_brand}' markasi Amazon'da araniyor...", expanded=True) as status:
                    st.write("Apify actor calistiriliyor...")
                    results = search_amazon_brand(amz_brand, max_items=amz_brand_max)
                    results = amazon_deduplicate(results)
                    status.update(label=f"{len(results)} urun bulundu!", state="complete")

                st.session_state.amazon_results = results
                st.session_state.amazon_search_type = "brand"
            except Exception as e:
                st.error(f"Hata: {str(e)}")

    # ── TAB 3: BSR Calculator ────────────────────────────────────────────────
    with tab_bsr_calc:
        st.markdown("##### BSR'den Satis Tahmini")
        st.caption("Bir urunun BSR (Best Sellers Rank) degerini girerek tahmini aylik satisi hesaplayin.")

        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            bsr_input = st.number_input("BSR (Best Sellers Rank)", min_value=1, max_value=5000000, value=1000, key="bsr_input")
        with col_c2:
            bsr_price = st.number_input("Urun Fiyati ($)", min_value=0.01, max_value=10000.0, value=25.0, step=0.5, key="bsr_price")
        with col_c3:
            bsr_category = st.selectbox("Kategori", get_amazon_categories(), index=0, key="bsr_cat")

        if bsr_input:
            est_sales = estimate_monthly_sales(bsr_input, bsr_category)
            est_rev = bsr_estimate_revenue(bsr_input, bsr_price, bsr_category)
            tier = sales_tier(est_sales)

            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("Tahmini Aylik Satis", f"{est_sales:,}")
            mc2.metric("Tahmini Aylik Gelir", f"${est_rev:,.0f}")
            mc3.metric("Tahmini Yillik Gelir", f"${est_rev * 12:,.0f}")
            mc4.metric("Satis Seviyesi", tier)

            # BSR reference table
            st.divider()
            st.markdown("##### BSR Referans Tablosu")
            st.caption(f"Kategori: {bsr_category}")
            ref_bsrs = [1, 50, 100, 500, 1000, 2500, 5000, 10000, 25000, 50000, 100000]
            ref_data = []
            for b in ref_bsrs:
                s = estimate_monthly_sales(b, bsr_category)
                r = bsr_estimate_revenue(b, bsr_price, bsr_category)
                ref_data.append({
                    "BSR": f"#{b:,}",
                    "Aylik Satis": f"{s:,}",
                    f"Aylik Gelir (${bsr_price:.0f})": f"${r:,.0f}",
                    "Seviye": sales_tier(s),
                })
            st.dataframe(pd.DataFrame(ref_data), use_container_width=True, hide_index=True)

    # ── RESULTS DISPLAY (shared between product & brand tabs) ────────────────
    if st.session_state.amazon_results:
        results = st.session_state.amazon_results
        st.divider()

        search_label = "Urun" if st.session_state.amazon_search_type == "product" else "Marka"
        st.subheader(f"{len(results)} {search_label} Sonucu")

        # Convert to DataFrame
        amz_rows = []
        for r in results:
            amz_rows.append({
                "ASIN": r.get("asin", ""),
                "Urun Adi": r.get("title", "")[:80],
                "Marka": r.get("brand", ""),
                "Fiyat ($)": r.get("price", 0),
                "Puan": r.get("rating", 0),
                "Yorum Sayisi": r.get("review_count", 0),
                "BSR": r.get("bsr", 0),
                "Kategori": r.get("category", ""),
                "Aylik Satis (Tah.)": r.get("monthly_sales_est", 0),
                "Aylik Gelir (Tah.)": r.get("monthly_revenue_est", 0),
                "Seviye": r.get("sales_tier", "-"),
                "Prime": "Evet" if r.get("is_prime") else "-",
                "URL": r.get("url", ""),
            })

        amz_df = pd.DataFrame(amz_rows)

        # Filter out zero-price or empty results
        if not amz_df.empty:
            # Top metrics
            valid_prices = amz_df[amz_df["Fiyat ($)"] > 0]
            valid_bsr = amz_df[amz_df["BSR"] > 0]
            valid_rev = amz_df[amz_df["Aylik Gelir (Tah.)"] > 0]

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Toplam Urun", len(amz_df))
            m2.metric("Ort. Fiyat", f"${valid_prices['Fiyat ($)'].mean():.2f}" if not valid_prices.empty else "-")
            m3.metric("Ort. Puan", f"{amz_df['Puan'].mean():.1f}" if amz_df['Puan'].sum() > 0 else "-")
            m4.metric("Ort. Yorum", f"{amz_df['Yorum Sayisi'].mean():,.0f}" if amz_df['Yorum Sayisi'].sum() > 0 else "-")
            m5.metric("Ort. Aylik Gelir", f"${valid_rev['Aylik Gelir (Tah.)'].mean():,.0f}" if not valid_rev.empty else "-")

            # Sort options
            sort_col = st.selectbox("Sirala", [
                "Aylik Gelir (Tah.)", "Aylik Satis (Tah.)", "BSR", "Fiyat ($)",
                "Yorum Sayisi", "Puan",
            ], key="amz_sort")
            ascending = sort_col == "BSR"  # BSR: lower is better
            amz_df_sorted = amz_df.sort_values(sort_col, ascending=ascending)

            # Display table
            st.dataframe(
                amz_df_sorted,
                use_container_width=True,
                height=500,
                column_config={
                    "Fiyat ($)": st.column_config.NumberColumn("Fiyat ($)", format="$%.2f"),
                    "Puan": st.column_config.NumberColumn("Puan", format="%.1f"),
                    "Yorum Sayisi": st.column_config.NumberColumn("Yorum", format="%d"),
                    "BSR": st.column_config.NumberColumn("BSR", format="%d"),
                    "Aylik Satis (Tah.)": st.column_config.NumberColumn("Aylik Satis", format="%d"),
                    "Aylik Gelir (Tah.)": st.column_config.NumberColumn("Aylik Gelir", format="$%d"),
                    "URL": st.column_config.LinkColumn("URL", display_text="Amazon"),
                },
                hide_index=True,
            )

            # Charts
            if not valid_rev.empty and len(valid_rev) > 3:
                col_ch1, col_ch2 = st.columns(2)

                with col_ch1:
                    # Price distribution
                    fig_price = px.histogram(
                        valid_prices, x="Fiyat ($)", nbins=15,
                        title="Fiyat Dagilimi",
                        color_discrete_sequence=["#667eea"],
                    )
                    fig_price.update_layout(height=350)
                    st.plotly_chart(fig_price, use_container_width=True)

                with col_ch2:
                    # Revenue by brand (top 10)
                    if amz_df_sorted["Marka"].nunique() > 1:
                        brand_rev = amz_df_sorted.groupby("Marka")["Aylik Gelir (Tah.)"].sum().sort_values(ascending=False).head(10)
                        fig_brand = px.bar(
                            x=brand_rev.values, y=brand_rev.index,
                            orientation="h",
                            title="Marka Bazinda Tahmini Gelir (Top 10)",
                            labels={"x": "Aylik Gelir ($)", "y": "Marka"},
                            color_discrete_sequence=["#764ba2"],
                        )
                        fig_brand.update_layout(height=350, yaxis=dict(autorange="reversed"))
                        st.plotly_chart(fig_brand, use_container_width=True)
                    else:
                        # Sales tier distribution
                        tier_counts = amz_df_sorted["Seviye"].value_counts().reset_index()
                        tier_counts.columns = ["Seviye", "Sayi"]
                        fig_tier = px.pie(tier_counts, values="Sayi", names="Seviye",
                                         title="Satis Seviyesi Dagilimi", hole=0.4)
                        fig_tier.update_layout(height=350)
                        st.plotly_chart(fig_tier, use_container_width=True)

                # Scatter: Price vs Reviews (bubble = revenue)
                scatter_df = amz_df_sorted[(amz_df_sorted["Fiyat ($)"] > 0) & (amz_df_sorted["Yorum Sayisi"] > 0)]
                if len(scatter_df) > 3:
                    fig_scatter = px.scatter(
                        scatter_df,
                        x="Fiyat ($)", y="Yorum Sayisi",
                        size="Aylik Gelir (Tah.)",
                        color="Seviye",
                        hover_name="Urun Adi",
                        title="Fiyat vs Yorum (Balon = Tahmini Gelir)",
                        color_discrete_map={
                            "Cok Yuksek": "#00b894",
                            "Yuksek": "#00cec9",
                            "Orta": "#fdcb6e",
                            "Dusuk": "#e17055",
                            "Cok Dusuk": "#d63031",
                        },
                    )
                    fig_scatter.update_layout(height=450)
                    st.plotly_chart(fig_scatter, use_container_width=True)

            # Export
            st.divider()
            col_e1, col_e2, col_e3 = st.columns(3)
            with col_e1:
                csv = amz_df_sorted.to_csv(index=False).encode("utf-8")
                st.download_button("CSV Indir", csv,
                    f"Amazon_{amz_keyword if 'amz_keyword' in dir() else 'search'}.csv", "text/csv",
                    use_container_width=True)
            with col_e2:
                from io import BytesIO
                buffer = BytesIO()
                amz_df_sorted.to_excel(buffer, index=False, engine="openpyxl")
                st.download_button("Excel Indir", buffer.getvalue(),
                    f"Amazon_{amz_keyword if 'amz_keyword' in dir() else 'search'}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with col_e3:
                json_str = json.dumps(results, indent=2, ensure_ascii=False).encode("utf-8")
                st.download_button("JSON Indir", json_str,
                    f"Amazon_{amz_keyword if 'amz_keyword' in dir() else 'search'}.json", "application/json",
                    use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: KAYDEDILENLER (SAVED / FAVORITES)
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Kaydedilenler":
    st.markdown('<p class="main-header">Kaydedilenler</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Favori markalarını klasörlerle düzenle</p>', unsafe_allow_html=True)

    # Folder management
    col_fm1, col_fm2, col_fm3 = st.columns([2, 1, 1])
    with col_fm1:
        new_folder = st.text_input("Yeni klasör adı", placeholder="örn: Kozmetik, Yatak, Atıştırmalık...")
    with col_fm2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Klasör Oluştur", type="primary", use_container_width=True):
            if new_folder:
                if db_create_folder(new_folder):
                    st.success(f"'{new_folder}' klasörü oluşturuldu!")
                    st.rerun()
                else:
                    st.warning("Bu klasör zaten var")
    with col_fm3:
        st.markdown("<br>", unsafe_allow_html=True)
        kf = db_load_folders()
        deletable = [f for f in kf if f != "Genel"]
        if deletable:
            delete_folder_name = st.selectbox("Sil", deletable, key="del_folder", label_visibility="collapsed")
            if st.button("Klasörü Sil", use_container_width=True):
                db_delete_folder(delete_folder_name)
                st.success(f"'{delete_folder_name}' silindi")
                st.rerun()

    st.divider()

    # Folder tabs
    kf = db_load_folders()
    if kf:
        tabs = st.tabs(kf)
        for tab, folder_name in zip(tabs, kf):
            with tab:
                brands = db_load_brands(folder_name)
                if brands:
                    st.markdown(f"**{len(brands)}** marka kayıtlı")
                    folder_df = pd.DataFrame(brands)

                    # Display saved brands
                    display_cols = [c for c in ["Marka", "Web Sitesi", "Kategori", "Alt Niş",
                                                "AOV ($)", "Öne Çıkan Özellik", "Meta Ads"] if c in folder_df.columns]
                    col_config = {}
                    if "Web Sitesi" in display_cols:
                        col_config["Web Sitesi"] = st.column_config.LinkColumn("Web Sitesi", display_text="Ziyaret Et")
                    if "Meta Ads" in display_cols:
                        col_config["Meta Ads"] = st.column_config.LinkColumn("Meta Ads", display_text="Reklamları Gör")
                    if "AOV ($)" in display_cols:
                        col_config["AOV ($)"] = st.column_config.NumberColumn("AOV ($)", format="$%d")

                    st.dataframe(folder_df[display_cols], use_container_width=True, height=400,
                                 column_config=col_config)

                    # Remove brands
                    brand_names_list = [b.get("Marka", b.get("brand", "")) for b in brands]
                    remove_brands_sel = st.multiselect(
                        "Kaldırmak istediğin markaları seç", brand_names_list,
                        key=f"remove_{folder_name}",
                    )
                    col_r1, col_r2 = st.columns(2)
                    with col_r1:
                        if st.button("Seçilenleri Kaldır", key=f"rem_btn_{folder_name}"):
                            db_remove_brands_by_name(folder_name, remove_brands_sel)
                            st.success(f"{len(remove_brands_sel)} marka kaldırıldı")
                            st.rerun()
                    with col_r2:
                        other_folders = [f for f in kf if f != folder_name]
                        if other_folders and remove_brands_sel:
                            move_to = st.selectbox("Taşı →", other_folders, key=f"move_{folder_name}")
                            if st.button("Seçilenleri Taşı", key=f"move_btn_{folder_name}"):
                                db_move_brands(folder_name, move_to, remove_brands_sel)
                                st.success(f"Markalar '{move_to}' klasörüne taşındı")
                                st.rerun()

                    # Export folder
                    st.divider()
                    col_x1, col_x2 = st.columns(2)
                    with col_x1:
                        csv = folder_df.to_csv(index=False).encode("utf-8")
                        st.download_button(f"CSV İndir", csv,
                            f"Kaydedilenler_{folder_name}.csv", "text/csv",
                            use_container_width=True, key=f"csv_{folder_name}")
                    with col_x2:
                        from io import BytesIO
                        buf = BytesIO()
                        folder_df.to_excel(buf, index=False, engine="openpyxl")
                        st.download_button(f"Excel İndir", buf.getvalue(),
                            f"Kaydedilenler_{folder_name}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True, key=f"xlsx_{folder_name}")
                else:
                    st.info(f"'{folder_name}' klasörü boş. Canlı Araştırma'dan marka ekleyin.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: HİNTLİ DANIŞMAN (CHAT WITH CLAUDE SONNET + TOOL CALLING)
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Hintli Danışman":
    st.markdown('<p class="main-header">Hintli Danışman</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Sonnet ile güçlendirilmiş DTC danışmanın — araçları kullanabilir</p>', unsafe_allow_html=True)

    # Display chat history
    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Chat input
    if prompt := st.chat_input("Bir şey sor... (örn: 'Bu sitenin TQS skorunu hesapla', 'Bambu niş araştırması yap')"):
        # Add user message
        st.session_state.chat_messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # Generate response with tool calling
        with st.chat_message("assistant"):
            api_key = get_api_key()
            if not api_key:
                st.error("API anahtarı bulunamadı. Streamlit Secrets veya .env dosyasına ANTHROPIC_API_KEY ekleyin.")
            else:
                try:
                    import anthropic
                    client = anthropic.Anthropic(api_key=api_key)

                    # Build context about saved brands
                    saved_context = ""
                    for fname in db_load_folders():
                        fbrands = db_load_brands(fname)
                        if fbrands:
                            brand_names = ", ".join(b.get("Marka", "") for b in fbrands[:20])
                            saved_context += f"\nKlasör '{fname}': {brand_names}"

                    system_prompt = f"""Sen Yuvacım markası için çalışan kıdemli bir DTC (Direct-to-Consumer) e-ticaret danışmanısın.
Türkçe yanıt ver. Adın "Hintli Danışman".

Yuvacım bir Türk DTC markasıdır. Kullanıcı Fatih, Türkiye'de DTC markaları kuran bir girişimcidir.

Kullanıcının kayıtlı markaları:{saved_context if saved_context else ' Henüz kayıtlı marka yok.'}

Görevin:
- DTC marka araştırması konusunda yardım et
- Pazarlama açıları, ürün fikirleri, niş analizi yap
- Türkiye pazarı için öneriler ver
- Marka karşılaştırmaları yap
- AOV, TQS, dönüşüm oranı hakkında bilgi ver
- Gerektiğinde araçları (tools) kullan: TQS hesaplama, marka araştırma, AOV tahmini, niş dönüşüm tablosu

Araçları proaktif kullan. Kullanıcı bir site hakkında sorduğunda TQS hesapla, niş sorduğunda araştırma yap."""

                    # Build message history (last 20 messages)
                    messages = [{"role": m["role"], "content": m["content"]}
                                for m in st.session_state.chat_messages[-20:]]

                    with st.spinner("Düşünüyor..."):
                        response = client.messages.create(
                            model="claude-sonnet-4-20250514",
                            max_tokens=4000,
                            system=system_prompt,
                            tools=TOOL_DEFINITIONS,
                            messages=messages,
                        )

                        # Tool use loop
                        tool_call_count = 0
                        while response.stop_reason == "tool_use" and tool_call_count < 5:
                            tool_call_count += 1
                            tool_results = []
                            for block in response.content:
                                if block.type == "tool_use":
                                    st.info(f"🔧 Araç kullanılıyor: **{block.name}**")
                                    result = execute_tool(block.name, block.input)
                                    tool_results.append({
                                        "type": "tool_result",
                                        "tool_use_id": block.id,
                                        "content": result,
                                    })

                            messages.append({"role": "assistant", "content": response.content})
                            messages.append({"role": "user", "content": tool_results})
                            response = client.messages.create(
                                model="claude-sonnet-4-20250514",
                                max_tokens=4000,
                                system=system_prompt,
                                tools=TOOL_DEFINITIONS,
                                messages=messages,
                            )

                        # Extract final text
                        final_text = "".join(
                            b.text for b in response.content if hasattr(b, "text")
                        )

                    st.markdown(final_text)
                    st.session_state.chat_messages.append({"role": "assistant", "content": final_text})

                except Exception as e:
                    error_msg = str(e)
                    if "credit" in error_msg.lower():
                        st.error("API kredisi yetersiz. console.anthropic.com'dan kredi yükleyin.")
                    else:
                        st.error(f"Hata: {error_msg}")

    # Clear chat button
    if st.session_state.chat_messages:
        if st.button("Sohbeti Temizle"):
            st.session_state.chat_messages = []
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1: BRAND BROWSER
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Marka Tarayıcı":
    st.markdown('<p class="main-header">Marka Tarayıcı</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{len(df)} post-2018 e-ticaret doğumlu DTC marka</p>', unsafe_allow_html=True)

    # Top metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Toplam Marka", f"{len(df):,}")
    with col2:
        st.metric("Kategori", df["Kategori"].nunique())
    with col3:
        has_traffic = df["Aylık Trafik"].notna().sum()
        st.metric("Trafik Verisi Olan", has_traffic)
    with col4:
        has_rev = df["Tahmini Aylık Gelir ($)"].notna().sum()
        st.metric("Gelir Tahmini Olan", has_rev)

    st.divider()

    # Filters
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)

    with col_f1:
        cats = ["Tümü"] + sorted(df["Kategori"].unique().tolist())
        selected_cat = st.selectbox("Kategori", cats)

    with col_f2:
        search = st.text_input("Marka Ara", placeholder="Marka adı...")

    with col_f3:
        traffic_filter = st.selectbox("Trafik Filtresi", [
            "Tümü", "100K+", "50K+", "10K+", "Sadece Verililer"
        ])

    with col_f4:
        sort_by = st.selectbox("Sırala", [
            "Marka (A-Z)", "Tahmini Aylık Gelir ($) ↓", "Aylık Trafik ↓",
            "AOV ($) ↓", "TQS ↓", "Kuruluş Yılı ↓"
        ])

    # Apply filters
    filtered = df.copy()
    if selected_cat != "Tümü":
        filtered = filtered[filtered["Kategori"] == selected_cat]
    if search:
        filtered = filtered[filtered["Marka"].str.contains(search, case=False, na=False)]
    if traffic_filter == "100K+":
        filtered = filtered[filtered["Aylık Trafik"] >= 100_000]
    elif traffic_filter == "50K+":
        filtered = filtered[filtered["Aylık Trafik"] >= 50_000]
    elif traffic_filter == "10K+":
        filtered = filtered[filtered["Aylık Trafik"] >= 10_000]
    elif traffic_filter == "Sadece Verililer":
        filtered = filtered[filtered["Aylık Trafik"].notna()]

    # Sort
    if "Gelir" in sort_by:
        filtered = filtered.sort_values("Tahmini Aylık Gelir ($)", ascending=False, na_position="last")
    elif "Trafik" in sort_by:
        filtered = filtered.sort_values("Aylık Trafik", ascending=False, na_position="last")
    elif "AOV" in sort_by:
        filtered = filtered.sort_values("AOV ($)", ascending=False, na_position="last")
    elif "TQS" in sort_by:
        filtered = filtered.sort_values("TQS", ascending=False, na_position="last")
    elif "Kuruluş" in sort_by:
        filtered = filtered.sort_values("Kuruluş Yılı", ascending=False, na_position="last")
    else:
        filtered = filtered.sort_values("Marka")

    st.markdown(f"**{len(filtered):,}** marka gösteriliyor")

    # Display table
    display_cols = [
        "Marka", "Web Sitesi", "Kategori", "Alt Niş", "Kuruluş Yılı",
        "Aylık Trafik", "AOV ($)", "TQS", "Dönüşüm %",
        "Tahmini Aylık Gelir ($)", "Öne Çıkan Özellik", "Marka Hikayesi"
    ]

    st.dataframe(
        filtered[display_cols],
        use_container_width=True,
        height=600,
        column_config={
            "Web Sitesi": st.column_config.LinkColumn("Web Sitesi", display_text="Ziyaret Et"),
            "Aylık Trafik": st.column_config.NumberColumn("Aylık Trafik", format="%d"),
            "AOV ($)": st.column_config.NumberColumn("AOV ($)", format="$%d"),
            "TQS": st.column_config.ProgressColumn("TQS", min_value=0, max_value=10, format="%.1f"),
            "Dönüşüm %": st.column_config.NumberColumn("Dönüşüm %", format="%.2f%%"),
            "Tahmini Aylık Gelir ($)": st.column_config.NumberColumn("Tahmini Gelir", format="$%d"),
            "Öne Çıkan Özellik": st.column_config.TextColumn("Öne Çıkan Özellik", width="large"),
            "Marka Hikayesi": st.column_config.TextColumn("Marka Hikayesi", width="large"),
        },
    )

    # Expandable brand detail
    st.divider()
    st.subheader("Marka Detayı")
    brand_names = filtered["Marka"].tolist()
    if brand_names:
        selected_brand = st.selectbox("Marka seç", brand_names)
        brand_row = filtered[filtered["Marka"] == selected_brand].iloc[0]

        col_d1, col_d2 = st.columns([1, 2])
        with col_d1:
            st.markdown(f"### {brand_row['Marka']}")
            website = brand_row["Web Sitesi"]
            url = f"https://{website}" if not website.startswith("http") else website
            st.markdown(f"[{website}]({url})")
            meta_url = f"https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=ALL&q={brand_row['Marka']}&search_type=keyword_unordered"
            st.markdown(f"[Meta Reklam Kütüphanesi]({meta_url})")
            st.markdown(f"**Kategori:** {brand_row['Kategori']}")
            st.markdown(f"**Alt Niş:** {brand_row['Alt Niş']}")
            if brand_row["Kuruluş Yılı"]:
                st.markdown(f"**Kuruluş:** {int(brand_row['Kuruluş Yılı'])}")

        with col_d2:
            if brand_row["Aylık Trafik"]:
                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.metric("Aylık Trafik", f"{int(brand_row['Aylık Trafik']):,}")
                mc2.metric("AOV", f"${int(brand_row['AOV ($)'])}")
                mc3.metric("TQS", f"{brand_row['TQS']}")
                mc4.metric("Tahmini Gelir", f"${int(brand_row['Tahmini Aylık Gelir ($)']):,}" if brand_row["Tahmini Aylık Gelir ($)"] else "-")

            st.markdown(f"**Öne Çıkan Özellik:** {brand_row['Öne Çıkan Özellik']}")
            st.markdown(f"**Marka Hikayesi:** {brand_row['Marka Hikayesi']}")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2: CATEGORY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Kategori Analizi":
    st.markdown('<p class="main-header">Kategori Analizi</p>', unsafe_allow_html=True)

    # Category breakdown
    cat_stats = df.groupby("Kategori").agg(
        Marka_Sayısı=("Marka", "count"),
        Ort_AOV=("AOV ($)", "mean"),
        Trafik_Olan=("Aylık Trafik", "count"),
        Toplam_Trafik=("Aylık Trafik", "sum"),
        Toplam_Gelir=("Tahmini Aylık Gelir ($)", "sum"),
    ).reset_index()
    cat_stats = cat_stats.sort_values("Marka_Sayısı", ascending=False)

    # Chart: brands per category
    fig1 = px.bar(
        cat_stats, x="Marka_Sayısı", y="Kategori",
        orientation="h", title="Kategori Bazında Marka Sayısı",
        color="Marka_Sayısı", color_continuous_scale="Viridis",
    )
    fig1.update_layout(height=700, yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig1, use_container_width=True)

    # Chart: revenue by category (for those with data)
    cat_rev = cat_stats[cat_stats["Toplam_Gelir"] > 0].sort_values("Toplam_Gelir", ascending=False)
    if not cat_rev.empty:
        fig2 = px.bar(
            cat_rev, x="Toplam_Gelir", y="Kategori",
            orientation="h", title="Kategori Bazında Tahmini Toplam Aylık Gelir ($)",
            color="Toplam_Gelir", color_continuous_scale="Greens",
        )
        fig2.update_layout(height=500, yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig2, use_container_width=True)

    # AOV comparison
    fig3 = px.bar(
        cat_stats.sort_values("Ort_AOV", ascending=False),
        x="Ort_AOV", y="Kategori",
        orientation="h", title="Kategori Bazında Ortalama AOV ($)",
        color="Ort_AOV", color_continuous_scale="Oranges",
    )
    fig3.update_layout(height=700, yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig3, use_container_width=True)

    # Table
    st.subheader("Detay Tablo")
    st.dataframe(cat_stats, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3: TQS CALCULATOR
# ══════════════════════════════════════════════════════════════════════════════
elif page == "TQS Hesaplayıcı":
    st.markdown('<p class="main-header">TQS Hesaplayıcı</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Herhangi bir web sitesinin tahmini gelirini hesapla</p>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Metrikleri Gir")
        traffic = st.number_input("Aylık Trafik", min_value=0, value=100000, step=10000)
        bounce_pct = st.slider("Bounce Rate (%)", 0.0, 100.0, 45.0, 0.5)
        pages_visit = st.slider("Sayfa / Ziyaret", 1.0, 15.0, 3.5, 0.1)
        session_min = st.slider("Ort. Oturum Süresi (dakika)", 0.0, 15.0, 2.5, 0.25)
        session_sec = session_min * 60
        aov = st.number_input("AOV ($)", min_value=1, value=80, step=5)
        niche = st.selectbox("Niş", [
            "base (Genel)", "food_bev (Yiyecek & İçecek - 2.0x)",
            "beauty (Güzellik - 1.5x)", "fashion (Moda - 1.0x)",
            "electronics (Elektronik - 0.7x)", "luxury (Lüks - 0.5x)",
        ])
        niche_key = niche.split(" ")[0] if niche != "base (Genel)" else None

    with col2:
        st.subheader("Sonuçlar")

        bs = bounce_score(bounce_pct)
        ps = pages_score(pages_visit)
        ds = duration_score(session_sec)
        tqs = calc_tqs(bounce_pct, pages_visit, session_sec)
        conv = tqs_to_conversion(tqs, niche_key)
        rev = round(estimate_revenue(traffic, aov, conv))

        # Score breakdown
        sc1, sc2, sc3 = st.columns(3)
        sc1.metric("Bounce Score", f"{bs}/10")
        sc2.metric("Pages Score", f"{ps}/10")
        sc3.metric("Duration Score", f"{ds}/10")

        st.divider()

        mc1, mc2 = st.columns(2)
        mc1.metric("TQS", f"{tqs}/10")
        multiplier = NICHE_MULTIPLIERS.get(niche_key, 1.0) if niche_key else 1.0
        mc2.metric("Niş Çarpanı", f"{multiplier}x")

        st.divider()

        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("Dönüşüm Oranı", f"{conv}%")
        rc2.metric("Aylık Gelir", f"${rev:,}")
        rc3.metric("Yıllık Gelir", f"${rev * 12:,}")

        # TQS visual gauge
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=tqs,
            title={"text": "Traffic Quality Score"},
            gauge={
                "axis": {"range": [0, 10]},
                "bar": {"color": "#2ecc71" if tqs >= 7 else "#f39c12" if tqs >= 4 else "#e74c3c"},
                "steps": [
                    {"range": [0, 3], "color": "#fadbd8"},
                    {"range": [3, 6], "color": "#fef9e7"},
                    {"range": [6, 10], "color": "#d5f5e3"},
                ],
            },
        ))
        fig.update_layout(height=250)
        st.plotly_chart(fig, use_container_width=True)

    # Niche conversion table
    st.divider()
    st.subheader("Niş Bazlı Dönüşüm Tablosu")
    table_data = []
    for tqs_val in range(1, 11):
        row = {"TQS": tqs_val, "Base": f"{TQS_TO_BASE_CONVERSION[tqs_val]}%"}
        for niche_name, table in NICHE_CONVERSION_TABLES.items():
            label = {"food_bev": "Food & Bev (2.0x)", "beauty": "Beauty (1.5x)",
                     "fashion": "Fashion (1.0x)", "electronics": "Electronics (0.7x)",
                     "luxury": "Luxury (0.5x)"}[niche_name]
            row[label] = f"{table[tqs_val]}%"
        table_data.append(row)
    st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4: AOV ESTIMATOR
# ══════════════════════════════════════════════════════════════════════════════
elif page == "AOV Tahminleyici":
    st.markdown('<p class="main-header">AOV Tahminleyici</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">En çok satanların fiyatına göre AOV hesapla</p>', unsafe_allow_html=True)

    st.subheader("En Çok Satan Ürünleri Gir")

    num_products = st.slider("Kaç ürün gireceksin?", 1, 10, 5)

    products = []
    cols = st.columns(min(num_products, 5))
    for i in range(num_products):
        with cols[i % 5]:
            price = st.number_input(f"Ürün {i+1} Fiyatı ($)", min_value=0.0, value=50.0, step=5.0, key=f"p{i}")
            reviews = st.number_input(f"Yorum Sayısı", min_value=0, value=100, step=10, key=f"r{i}")
            products.append({"price": price, "reviews": reviews})

    st.divider()

    col_a1, col_a2 = st.columns(2)
    with col_a1:
        total_products = st.number_input("Toplam Ürün Sayısı", min_value=1, value=50, step=5)
        good_upsell = st.checkbox("İyi Upsell / Bundling Var mı?", value=True)

    with col_a2:
        # Calculate review-weighted AOV
        total_reviews = sum(p["reviews"] for p in products)
        if total_reviews > 0:
            weighted_aov = sum(p["price"] * p["reviews"] for p in products) / total_reviews
        else:
            weighted_aov = sum(p["price"] for p in products) / len(products)

        # Apply upsell bump
        if total_products >= 30 and good_upsell:
            final_aov = weighted_aov * 1.10
            bump = "+10% upsell"
        else:
            final_aov = weighted_aov
            bump = "Upsell yok"

        st.metric("Yorum Ağırlıklı AOV", f"${weighted_aov:.2f}")
        st.metric("Upsell Düzeltmesi", bump)
        st.metric("Tahmini AOV", f"${final_aov:.2f}")

    # Product breakdown chart
    if products:
        prod_df = pd.DataFrame(products)
        prod_df["Ürün"] = [f"Ürün {i+1}" for i in range(len(products))]
        prod_df["Ağırlık (%)"] = (prod_df["reviews"] / max(total_reviews, 1) * 100).round(1)

        fig = px.bar(
            prod_df, x="Ürün", y="reviews",
            color="price", title="Ürün Bazında Yorum Dağılımı & Fiyat",
            labels={"reviews": "Yorum Sayısı", "price": "Fiyat ($)"},
            color_continuous_scale="Viridis",
        )
        st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 5: MARKETING ANGLES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Pazarlama Açıları":
    st.markdown('<p class="main-header">Pazarlama Açısı Analizi</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Hangi pazarlama açılarında en fazla talep var?</p>', unsafe_allow_html=True)

    # Keyword-based angle detection
    ANGLE_KEYWORDS = {
        "Soğutma / Termal": ["soğut", "cool", "serin", "termal", "sıcaklık"],
        "Ağırlıklı / Anksiyete": ["ağırlıklı", "weighted", "anksiyete", "anxiety", "cam boncuk"],
        "Organik / Sertifikalı": ["organik", "organic", "GOTS", "OEKO-TEX", "sertifika"],
        "Bambu Tekstil": ["bambu", "bamboo", "bambou"],
        "Sürdürülebilir / Eko": ["sürdürülebilir", "geri dönüşüm", "eco", "eko", "çevreci", "sıfır atık"],
        "Uyku Optimizasyonu": ["uyku", "sleep", "horlama", "insomnia"],
        "Memory Foam": ["memory foam", "foam", "viskoelastik"],
        "Keten / Linen": ["keten", "linen", "flax"],
        "TikTok / Viral": ["tiktok", "viral", "sosyal medya"],
        "Instagram Estetik": ["instagram", "pinterest", "estetik", "aesthetic", "minimal"],
        "Otel & Spa Evde": ["otel", "hotel", "spa", "5 yıldız"],
        "Bebek & Anne": ["bebek", "baby", "anne", "hamile", "emzirme"],
        "El Yapımı / Artisan": ["el yapımı", "handmade", "zanaatkar", "artisan"],
        "Kişiselleştirme": ["kişisel", "custom", "monogram", "personal"],
        "Abonelik Modeli": ["abonelik", "subscription", "refill"],
        "Ortopedik / Sağlık": ["ortopedik", "ergonomik", "boyun", "sırt", "omurga"],
        "Lüks / Premium": ["lüks", "premium", "luxury", "ultra"],
        "Modüler / Çok Fonksiyonlu": ["modüler", "modular", "çok kullanım", "multi"],
        "Shark Tank / Crowdfund": ["shark tank", "kickstarter", "crowdfund"],
        "Celebrity / Influencer": ["celebrity", "ünlü", "influencer", "kollab"],
    }

    angle_data = []
    for angle, keywords in ANGLE_KEYWORDS.items():
        matched = df[df["Öne Çıkan Özellik"].str.lower().apply(
            lambda x: any(kw.lower() in str(x) for kw in keywords)
        )]
        matched_with_traffic = matched[matched["Aylık Trafik"].notna()]
        angle_data.append({
            "Pazarlama Açısı": angle,
            "Marka Sayısı": len(matched),
            "Trafik Verisi Olan": len(matched_with_traffic),
            "Toplam Trafik": matched_with_traffic["Aylık Trafik"].sum(),
            "Toplam Gelir ($)": matched_with_traffic["Tahmini Aylık Gelir ($)"].sum(),
            "Ort. AOV ($)": matched["AOV ($)"].mean(),
        })

    angle_df = pd.DataFrame(angle_data).sort_values("Marka Sayısı", ascending=False)

    # Chart
    fig = px.bar(
        angle_df, x="Marka Sayısı", y="Pazarlama Açısı",
        orientation="h", title="Pazarlama Açısı Bazında Marka Sayısı",
        color="Toplam Gelir ($)", color_continuous_scale="RdYlGn",
    )
    fig.update_layout(height=600, yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig, use_container_width=True)

    # Revenue chart
    angle_rev = angle_df[angle_df["Toplam Gelir ($)"] > 0].sort_values("Toplam Gelir ($)", ascending=False)
    if not angle_rev.empty:
        fig2 = px.bar(
            angle_rev, x="Toplam Gelir ($)", y="Pazarlama Açısı",
            orientation="h", title="Trafik Ağırlıklı Talep (Gelire Göre)",
            color="Toplam Gelir ($)", color_continuous_scale="Greens",
        )
        fig2.update_layout(height=500, yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig2, use_container_width=True)

    # Table
    st.subheader("Detay Tablo")
    st.dataframe(
        angle_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Toplam Trafik": st.column_config.NumberColumn(format="%d"),
            "Toplam Gelir ($)": st.column_config.NumberColumn(format="$%d"),
            "Ort. AOV ($)": st.column_config.NumberColumn(format="$%.0f"),
        },
    )

    # Drill-down
    st.divider()
    selected_angle = st.selectbox("Açı Detayı", angle_df["Pazarlama Açısı"].tolist())
    keywords = ANGLE_KEYWORDS[selected_angle]
    drill = df[df["Öne Çıkan Özellik"].str.lower().apply(
        lambda x: any(kw.lower() in str(x) for kw in keywords)
    )]
    st.dataframe(
        drill[["Marka", "Web Sitesi", "Kategori", "Aylık Trafik", "AOV ($)", "TQS", "Tahmini Aylık Gelir ($)", "Öne Çıkan Özellik"]],
        use_container_width=True,
        height=400,
    )
