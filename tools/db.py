"""
Supabase persistence layer for saved folders and brands.
Shared across all users — any changes are visible to everyone.
"""
import os
import math
import streamlit as st
from supabase import create_client


def _get_client():
    """Get Supabase client, trying Streamlit secrets first, then env vars."""
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except Exception:
        from dotenv import load_dotenv
        load_dotenv()
        url = os.getenv("SUPABASE_URL", "")
        key = os.getenv("SUPABASE_KEY", "")
    if not url or not key:
        return None
    return create_client(url, key)


@st.cache_resource
def get_db():
    """Cached Supabase client (one connection per app instance)."""
    return _get_client()


def _clean_value(v):
    """Clean a value for JSON serialization — handle NaN, float('inf'), etc."""
    if v is None:
        return None
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        if v == int(v):
            return int(v)
    return v


def _clean_brand_data(brand_data):
    """Clean brand data dict for Supabase JSONB storage."""
    clean = {}
    for k, v in brand_data.items():
        if k.startswith("_"):
            continue
        clean[k] = _clean_value(v)
    return clean


def load_folders():
    """Load all folder names from Supabase."""
    db = get_db()
    if not db:
        return ["Genel"]
    try:
        result = db.table("folders").select("name").order("created_at").execute()
        names = [r["name"] for r in result.data]
        return names if names else ["Genel"]
    except Exception as e:
        st.error(f"Klasör yükleme hatası: {e}")
        return ["Genel"]


def create_folder(name):
    """Create a new folder. Returns True if created, False if exists."""
    db = get_db()
    if not db:
        st.error("Veritabanı bağlantısı yok")
        return False
    try:
        db.table("folders").insert({"name": name}).execute()
        # Clear the cache so folder list refreshes
        load_folders.clear() if hasattr(load_folders, 'clear') else None
        return True
    except Exception as e:
        if "duplicate" in str(e).lower() or "unique" in str(e).lower() or "23505" in str(e):
            st.warning("Bu klasör zaten var")
        else:
            st.error(f"Klasör oluşturma hatası: {e}")
        return False


def delete_folder(name):
    """Delete a folder and all its saved brands."""
    db = get_db()
    if not db:
        return
    try:
        db.table("saved_brands").delete().eq("folder_name", name).execute()
        db.table("folders").delete().eq("name", name).execute()
    except Exception as e:
        st.error(f"Klasör silme hatası: {e}")


def load_brands(folder_name):
    """Load all saved brands from a folder."""
    db = get_db()
    if not db:
        return []
    try:
        result = (db.table("saved_brands")
                  .select("id, brand_data")
                  .eq("folder_name", folder_name)
                  .order("created_at")
                  .execute())
        return [{"_db_id": r["id"], **r["brand_data"]} for r in result.data]
    except Exception as e:
        st.error(f"Marka yükleme hatası: {e}")
        return []


def save_brands_bulk(folder_name, brands_list):
    """Save multiple brands to a folder, skipping duplicates."""
    db = get_db()
    if not db:
        st.error("Veritabanı bağlantısı yok")
        return 0

    # Ensure folder exists
    try:
        existing_folders = [r["name"] for r in db.table("folders").select("name").execute().data]
        if folder_name not in existing_folders:
            db.table("folders").insert({"name": folder_name}).execute()
    except Exception:
        pass

    # Get existing brand names in folder
    existing = load_brands(folder_name)
    existing_names = set()
    for b in existing:
        name = b.get("Marka", b.get("brand", ""))
        if name:
            existing_names.add(name.lower())

    added = 0
    rows = []
    for brand in brands_list:
        name = brand.get("Marka", brand.get("brand", ""))
        if not name:
            continue
        if name.lower() not in existing_names:
            clean = _clean_brand_data(brand)
            rows.append({"folder_name": folder_name, "brand_data": clean})
            existing_names.add(name.lower())
            added += 1

    if rows:
        try:
            # Insert in batches of 50 to avoid payload limits
            for i in range(0, len(rows), 50):
                batch = rows[i:i+50]
                db.table("saved_brands").insert(batch).execute()
        except Exception as e:
            st.error(f"Kaydetme hatası: {e}")
            return 0
    return added


def remove_brand(db_id):
    """Remove a saved brand by its database ID."""
    db = get_db()
    if not db:
        return
    try:
        db.table("saved_brands").delete().eq("id", db_id).execute()
    except Exception as e:
        st.error(f"Silme hatası: {e}")


def remove_brands_by_name(folder_name, brand_names):
    """Remove brands from a folder by their names."""
    brands = load_brands(folder_name)
    for b in brands:
        bname = b.get("Marka", b.get("brand", ""))
        if bname in brand_names:
            remove_brand(b["_db_id"])


def move_brands(from_folder, to_folder, brand_names):
    """Move brands from one folder to another."""
    brands = load_brands(from_folder)
    to_move = [b for b in brands if b.get("Marka", b.get("brand", "")) in brand_names]

    save_brands_bulk(to_folder, to_move)
    remove_brands_by_name(from_folder, brand_names)


def get_all_saved_count():
    """Get total count of saved brands across all folders."""
    db = get_db()
    if not db:
        return 0
    try:
        result = db.table("saved_brands").select("id", count="exact").execute()
        return result.count or 0
    except Exception:
        return 0
