import streamlit as st
from supabase import create_client, Client
import pandas as pd
import zipfile, io, re


def read_excel_safe(file) -> pd.DataFrame:
    raw = file.read() if hasattr(file, 'read') else file
    try:
        return pd.read_excel(io.BytesIO(raw), header=None)
    except Exception:
        pass
    try:
        output = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(raw)) as zin:
            with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
                for name in zin.namelist():
                    data = zin.read(name)
                    if name == 'xl/styles.xml':
                        data = re.sub(rb' borderID="\d+"', b'', data)
                        data = re.sub(rb' pivotButton="\d+"', b'', data)
                        data = re.sub(rb' quotePrefix="\d+"', b'', data)
                    zout.writestr(name, data)
        output.seek(0)
        return pd.read_excel(output, header=None)
    except Exception as e:
        raise Exception(f"לא ניתן לקרוא את הקובץ: {e}")


# ===== LITAY DB (DMWA, Brandlight) =====
@st.cache_resource
def get_litay_db() -> Client:
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)

def load_clients_litay():
    try:
        res = get_litay_db().table("client_config").select("client_id,client_name,bank_coa,vendor_index").execute()
        return {r["client_id"]: r for r in res.data}
    except:
        return {}

def load_coa_litay(client_id, lookup_type):
    try:
        res = get_litay_db().table("coa_lookup") \
            .select("lookup_key,coa_credit,coa_debit") \
            .eq("client_id", client_id).eq("lookup_type", lookup_type).execute()
        return {r["lookup_key"]: [r["coa_credit"], r["coa_debit"]] for r in res.data}
    except:
        return {}

def check_duplicates_litay(client_id, file_type, keys):
    try:
        res = get_litay_db().table("processed_keys") \
            .select("txn_key").eq("client_id", client_id) \
            .eq("file_type", file_type).in_("txn_key", list(keys)).execute()
        return {r["txn_key"] for r in res.data}
    except:
        return set()

def save_keys_litay(client_id, file_type, keys, file_name):
    rows = [{"client_id": client_id, "file_type": file_type,
             "txn_key": k, "file_name": file_name} for k in keys]
    try:
        get_litay_db().table("processed_keys").upsert(rows).execute()
    except Exception as e:
        st.warning(f"שגיאה בשמירת היסטוריה: {e}")


# ===== RAHEL MOR DB =====
@st.cache_resource
def get_rahel_db() -> Client:
    url = st.secrets["supabase_rahel"]["SUPABASE_URL"]
    key = st.secrets["supabase_rahel"]["SUPABASE_KEY"]
    return create_client(url, key)

def get_all_clients() -> dict:
    sb = get_rahel_db()
    result = sb.table("clients").select("name, account_number").order("account_number").execute()
    return {row["name"]: row["account_number"] for row in result.data}

def get_clients_full() -> list:
    sb = get_rahel_db()
    result = sb.table("clients").select("account_number, name, id_number").order("account_number").execute()
    return result.data

def add_client(name: str, account_number: int, id_number: str = ""):
    sb = get_rahel_db()
    sb.table("clients").delete().eq("account_number", account_number).execute()
    sb.table("clients").insert({"name": name, "account_number": account_number, "id_number": id_number}).execute()

def delete_client(name: str):
    get_rahel_db().table("clients").delete().eq("name", name).execute()

def get_next_account_number() -> int:
    sb = get_rahel_db()
    result = sb.table("clients").select("account_number").order("account_number", desc=True).limit(1).execute()
    return result.data[0]["account_number"] + 1 if result.data else 3001

def import_clients_from_df(df: pd.DataFrame) -> tuple:
    sb = get_rahel_db()
    header_row = None
    for i in range(min(10, len(df))):
        row_vals = [str(v) for v in df.iloc[i].tolist()]
        if any(k in v for v in row_vals for k in ['חשבון', 'מפתח']):
            header_row = i; break

    if header_row is not None:
        headers = [str(v) for v in df.iloc[header_row].tolist()]
        def find_col(keywords):
            for kw in keywords:
                for j, h in enumerate(headers):
                    if kw in h: return j
            return None
        col_account = find_col(['מפתח חשבון', 'חשבון'])
        col_name    = find_col(['שם החשבון', 'שם חשבון', 'שם'])
        col_id      = find_col(['מס.ע.מ', 'ע.מ', 'ת.ז'])
        data_start  = header_row + 1
    else:
        col_account, col_name, col_id = 5, 6, 10
        data_start = 3

    clients = []
    for i in range(data_start, len(df)):
        row = df.iloc[i]
        try:
            account = row.iloc[col_account] if col_account is not None else None
            name    = row.iloc[col_name] if col_name is not None else None
            id_num  = row.iloc[col_id] if col_id is not None else None
        except Exception:
            continue
        if pd.isna(account) or pd.isna(name): continue
        account_str = str(account).replace(".0", "").strip()
        if not (account_str.isdigit() and len(account_str) == 4
                and account_str.startswith("3") and account_str != "3999"): continue
        name_str = str(name).strip()
        if not name_str or name_str.isdigit() or len(name_str) < 2: continue
        id_str = ""
        if id_num is not None and not pd.isna(id_num):
            id_str = str(id_num).replace(".0", "").strip()
            if id_str in ("0", "nan", "None", ""): id_str = ""
        try:
            clients.append({"name": name_str, "account_number": int(float(account)), "id_number": id_str})
        except Exception:
            pass

    if not clients:
        return 0, "לא נמצאו לקוחות — בדקי שהקובץ הוא אינדקס חשבונות מחשבשבת"
    try:
        sb.table("clients").delete().neq("account_number", 0).execute()
        sb.table("clients").upsert(clients, on_conflict="name").execute()
        return len(clients), None
    except Exception as e:
        return 0, f"שגיאה: {str(e)}"

def get_account_columns() -> dict:
    sb = get_rahel_db()
    result = sb.table("account_columns").select("*").execute()
    return {row["column_name"]: {"account": row["account_number"], "vat_exempt": row["is_vat_exempt"]} for row in result.data}

def add_account_column(column_name: str, account_number: int, is_vat_exempt: bool):
    sb = get_rahel_db()
    sb.table("account_columns").delete().eq("column_name", column_name).execute()
    sb.table("account_columns").insert({"column_name": column_name, "account_number": account_number, "is_vat_exempt": is_vat_exempt}).execute()


# ===== SUPABASE STORAGE =====

def upload_file(client_id: str, file_type: str, filename: str, file_bytes: bytes) -> str:
    """מעלה קובץ ל-Supabase Storage. מחזיר path."""
    import re
    safe_name = re.sub(r'[^a-zA-Z0-9._-]', '_', filename)
    path = f"{client_id}/{file_type}/{safe_name}"
    try:
        sb = get_litay_db()
        # מחק קיים אם יש
        sb.storage.from_("uploads").remove([path])
    except:
        pass
    try:
        sb = get_litay_db()
        sb.storage.from_("uploads").upload(path, file_bytes)
        return path
    except Exception as e:
        return ""


def list_recent_files(client_id: str, file_type: str) -> list:
    """מחזיר רשימת קבצים אחרונים עבור לקוח וסוג קובץ."""
    prefix = f"{client_id}/{file_type}/"
    try:
        sb = get_litay_db()
        files = sb.storage.from_("uploads").list(prefix.rstrip("/"))
        return sorted(
            [f for f in files if f.get("name")],
            key=lambda x: x.get("updated_at", ""),
            reverse=True
        )[:10]
    except:
        return []


def download_file(client_id: str, file_type: str, filename: str) -> bytes:
    """מוריד קובץ מ-Supabase Storage."""
    import re
    safe_name = re.sub(r'[^a-zA-Z0-9._-]', '_', filename)
    path = f"{client_id}/{file_type}/{safe_name}"
    try:
        sb = get_litay_db()
        return sb.storage.from_("uploads").download(path)
    except:
        return b""
