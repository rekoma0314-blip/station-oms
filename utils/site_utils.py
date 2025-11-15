from utils.supabase_client import supabase

def get_site_by_code(code: str):
    """
    根据混合编码（code）查询站点信息
    Supabase 表字段：code, name, warehouse, company
    """
    resp = supabase.table("sites") \
                   .select("*") \
                   .eq("code", code) \
                   .execute()
    if resp.data:
        return resp.data[0]
    return None
