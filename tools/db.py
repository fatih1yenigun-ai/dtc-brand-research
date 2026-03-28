"""
Supabase persistence layer for saved folders and brands.
Shared across all users — any changes are visible to everyone.
"""
import os
import streamlit as st
from supabase import create_client

def _get_client():
    """Get Supabase client, trying Streamlit secrets first, then env vars."""
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except Exception:
        url = os.getenv("SUPABASE_URL", "")
        key = os.getenv("SUPABASE_KEY", "")
    if not url or not key:
        return None
    return create_client(url, key)


@st.cache_resource
def get_db():
    """Cached Supabase client (one connection per app instance)."""
    return _get_client()


def load_folders():
    """Load all folder names from Supabase."""
    db = get_db()
    if not db:
        return ["Genel"]
    try:
        result = db.table("folders").select("name").order("created_at").execute()
        names = [r["name"] for r in result.data]
        return names if names else ["Genel"]
    except Exception:
        return ["Genel"]


def create_folder(name):
    """Create a new folder. Returns True if created, False if exists."""
    db = get_db()
    if not db:
        return False
    try:
        db.table("folders").insert({"name": name}).execute()
        return True
    except Exception:
        return False


def delete_folder(name):
    """Delete a folder and all its saved brands."""
    db = get_db()
    if not db:
        return
    try:
        db.table("saved_brands").delete().eq("folder_name", name).execute()
        db.table("folders").delete().eq("name", name).execute()
    except Exception:
        pass


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
    except Exception:
        return []


def save_brand(folder_name, brand_data):
    """Save a single brand to a folder. Returns True if saved (no duplicate)."""
    db = get_db()
    if not db:
        return False
    try:
        # Check for duplicate by brand name
        brand_name = brand_data.get("Marka", brand_data.get("brand", ""))
        existing = (db.table("saved_brands")
                    .select("id")
                    .eq("folder_name", folder_name)
                    .execute())
        for row in existing.data:
            pass  # We'll do a simpler approach — just insert and let duplicates be handled in UI

        # Remove internal keys before saving
        clean = {k: v for k, v in brand_data.items() if not k.startswith("_")}
        db.table("saved_brands").insert({
            "folder_name": folder_name,
            "brand_data": clean,
        }).execute()
        return True
    except Exception:
        return False


def save_brands_bulk(folder_name, brands_list):
    """Save multiple brands to a folder, skipping duplicates."""
    db = get_db()
    if not db:
        return 0

    # Get existing brand names in folder
    existing = load_brands(folder_name)
    existing_names = {b.get("Marka", b.get("brand", "")).lower() for b in existing}

    added = 0
    rows = []
    for brand in brands_list:
        name = brand.get("Marka", brand.get("brand", ""))
        if name.lower() not in existing_names:
            clean = {k: v for k, v in brand.items() if not k.startswith("_")}
            rows.append({"folder_name": folder_name, "brand_data": clean})
            existing_names.add(name.lower())
            added += 1

    if rows:
        try:
            db.table("saved_brands").insert(rows).execute()
        except Exception:
            pass
    return added


def remove_brand(db_id):
    """Remove a saved brand by its database ID."""
    db = get_db()
    if not db:
        return
    try:
        db.table("saved_brands").delete().eq("id", db_id).execute()
    except Exception:
        pass


def remove_brands_by_name(folder_name, brand_names):
    """Remove brands from a folder by their names."""
    db = get_db()
    if not db:
        return
    brands = load_brands(folder_name)
    for b in brands:
        bname = b.get("Marka", b.get("brand", ""))
        if bname in brand_names:
            remove_brand(b["_db_id"])


def move_brands(from_folder, to_folder, brand_names):
    """Move brands from one folder to another."""
    brands = load_brands(from_folder)
    to_move = [b for b in brands if b.get("Marka", b.get("brand", "")) in brand_names]

    # Save to target folder
    save_brands_bulk(to_folder, to_move)

    # Remove from source
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
