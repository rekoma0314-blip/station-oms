from .supabase_client import supabase

def get_all_sites():
    response = supabase.table("sites").select("*").execute()
    return response.data

def get_site_by_code(site_code):
    response = (
        supabase
        .table("sites")
        .select("*")
        .eq("site_code", site_code)
        .execute()
    )
    if response.data:
        return response.data[0]
    return None
