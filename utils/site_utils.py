from .supabase_client import supabase


def get_all_sites():
    resp = supabase.table("sites").select("*").execute()
    return resp.data or []


def get_site_map():
    """
    返回 dict:
    {code: {"warehouse":..., "name":...}}
    """
    rows = get_all_sites()
    return {row["code"]: row for row in rows}
