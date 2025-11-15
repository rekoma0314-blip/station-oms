from utils.supabase_client import supabase

# 获取全部站点
def get_all_sites():
    result = supabase.table("sites").select("*").execute()
    return result.data or []

# 按新编码或旧编码查站点
def get_site_by_code(code):
    result = (
        supabase.table("sites")
        .select("*")
        .or_(f"site_code.eq.{code},old_code.eq.{code}")
        .execute()
    )
    return result.data[0] if result.data else None
