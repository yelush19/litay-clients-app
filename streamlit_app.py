import streamlit as st
import csv, io, hashlib
import pandas as pd
from datetime import datetime

st.set_page_config(
    page_title="ממיר חשבשבת | Litay",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

LOGO_URL = "https://xyukgdxpmbvobbnrpzwt.supabase.co/storage/v1/object/public/assets/logo-litay.png"

# ===== COA DEFAULTS =====
PAYEM_COA_DEFAULT = {"LEVERAGED OUTBOUND|Brandlight Inc.":["502010","911000"],"Upwork|Brandlight Inc.":["502010","962006"],"Calendly|Brandlight Inc.":["502010","901009"],"Linkedin|Brandlight Inc.":["911004","502010"],"OpenAI|Brandlight Inc.":["502010","901009"],"Google|Brandlight Inc.":["502010","901009"],"Notion|Brandlight Inc.":["901009","502010"],"Dropbox|Brandlight Inc.":["901009","502010"],"Adobe|Brandlight Inc.":["502010","901009"]}
VALLEY_CAT_COA = {"Revenue":["100000","700000"],"Employee Salaries & Benefits":["540001","100000"],"Office Rent":["964000","100000"],"Marketing & Advertising":["911003","100000"],"Bank Fees & interest":["990001","100000"],"Professional Services":["962006","100000"],"Credit/Debit banks/Accounts":["502010","100000"],"Technology":["901005","100000"],"Authorities":["990002","100000"],"Vendors":["500001","100000"],"Investment":["100000","600015"],"Recruitment":["962003","100000"]}
LTD_COA_DEFAULT = {}

# לקוחות + מודולים
CLIENT_MODULES = {
    "dmwa":       ["masav", "payit"],
    "rahel_mor":  ["income"],
    "brandlight": ["valley", "payem"],
}

# ===== CSS =====
def inject_css():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;600;700&display=swap');
    html, body, [class*="css"] {{ font-family:'Heebo',sans-serif; direction:rtl; }}
    .app-header {{
        display:flex; align-items:center; justify-content:space-between;
        background:#fff; border-bottom:3px solid #4CAF50;
        padding:12px 24px; margin-bottom:20px;
        border-radius:0 0 8px 8px; box-shadow:0 2px 10px rgba(0,0,0,0.06);
    }}
    .app-header img {{ height:44px; object-fit:contain; }}
    .app-header .title {{ font-size:1.3rem; font-weight:700; color:#1b5e20; margin:0; }}
    .app-header .subtitle {{ font-size:0.8rem; color:#757575; margin:2px 0 0; }}
    .badge {{ display:inline-block; padding:5px 16px; border-radius:20px; font-weight:600; font-size:0.85rem; }}
    .badge-valley {{ background:#e8f5e9; color:#2e7d32; border:2px solid #4caf50; }}
    .badge-payem  {{ background:#e3f2fd; color:#1565c0; border:2px solid #42a5f5; }}
    .badge-masav  {{ background:#fff3e0; color:#e65100; border:2px solid #ff9800; }}
    .badge-payit  {{ background:#f3e5f5; color:#6a1b9a; border:2px solid #ab47bc; }}
    .badge-income {{ background:#e8f5e9; color:#1b5e20; border:2px solid #4caf50; }}
    [data-testid="stSidebar"] {{ background:#f1f8e9; border-left:3px solid #4CAF50; }}
    .stButton>button[kind="primary"] {{
        background:#4CAF50; border:none; border-radius:8px; font-weight:600; transition:background 0.2s;
    }}
    .stButton>button[kind="primary"]:hover {{ background:#388E3C; }}
    .app-footer {{
        text-align:center; font-size:0.75rem; color:#bdbdbd;
        padding:20px 0 8px; border-top:1px solid #f0f0f0; margin-top:32px;
    }}
    </style>
    <div class="app-header">
        <div>
            <p class="title">ממיר תנועות → חשבשבת</p>
            <p class="subtitle">Valley Bank · PayEm · MASAV · Pay-it · הכנסות</p>
        </div>
        <img src="{LOGO_URL}" alt="Litay" />
    </div>
    """, unsafe_allow_html=True)

# ===== AUTH =====
def check_auth():
    if st.session_state.get("authenticated"):
        return True
    try:
        _ = st.secrets["auth"]["users"]
    except (KeyError, FileNotFoundError):
        st.session_state["authenticated"] = True
        return True
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown("### 🔒 כניסה למערכת")
        with st.form("login"):
            u = st.text_input("שם משתמש")
            p = st.text_input("סיסמה", type="password")
            if st.form_submit_button("התחבר", type="primary", use_container_width=True):
                users = st.secrets["auth"]["users"]
                h = hashlib.sha256(p.encode()).hexdigest()
                if u in users and users[u]["password_hash"] == h:
                    st.session_state["authenticated"] = True
                    st.session_state["auth_user"] = u
                    st.rerun()
                else:
                    st.error("שם משתמש או סיסמה שגויים")
    return False

def init_session():
    if "clients" not in st.session_state:
        from utils.db import load_clients_litay
        st.session_state["clients"] = load_clients_litay()
    if "selected_client" not in st.session_state:
        st.session_state["selected_client"] = None
    if "rahel_page" not in st.session_state:
        st.session_state["rahel_page"] = "עיבוד"

def workbook_to_bytes(wb):
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


def file_selector(client_id: str, file_type: str, label: str, 
                  file_ext: list, key: str) -> tuple:
    """
    מציג אפשרות העלאה חדשה או בחירה מקבצים שמורים.
    מחזיר (file_bytes, filename) או (None, None)
    """
    from utils.db import upload_file, list_recent_files, download_file

    recent = list_recent_files(client_id, file_type)

    tab_new, tab_recent = st.tabs(["📤 העלה חדש", f"📂 שמורים ({len(recent)})"])

    with tab_new:
        uploaded = st.file_uploader(label, type=file_ext, key=key)
        if uploaded:
            file_bytes = uploaded.read()
            # שמור ל-Supabase Storage
            path = upload_file(client_id, file_type, uploaded.name, file_bytes)
            if path:
                st.caption(f"✅ נשמר: {uploaded.name}")
            return file_bytes, uploaded.name

    with tab_recent:
        if not recent:
            st.info("אין קבצים שמורים עדיין")
            return None, None
        options = {f["name"]: f for f in recent}
        selected = st.selectbox("בחרי קובץ", list(options.keys()), 
                                 key=f"{key}_recent")
        if st.button("📂 טעני קובץ זה", key=f"{key}_load", type="primary"):
            file_bytes = download_file(client_id, file_type, selected)
            if file_bytes:
                st.success(f"✅ נטען: {selected}")
                return file_bytes, selected
            else:
                st.error("שגיאה בטעינה")

    return None, None

# ===== TABS BY CLIENT =====

def render_masav_tab(client):
    from utils.masav import parse_masav, build_masav_excel, fuzzy_match_vendor
    from utils.db import check_duplicates_litay, save_keys_litay, get_litay_db

    # תמיד טעון מ-session_state כדי לקבל נתונים מעודכנים
    client_id = client["client_id"]
    fresh_client = st.session_state["clients"].get(client_id, client)
    vendor_lookup = fresh_client.get("vendor_index") or {}
    bank_coa      = fresh_client.get("bank_coa", "") or client.get("bank_coa", "")

    if not vendor_lookup:
        st.warning("⚠️ אינדקס ספקים חסר"); return
    if not bank_coa:
        st.warning('⚠️ ח"ן בנק לא מוגדר'); return

    st.success(f"✅ {client['client_name']} | ח\"ן בנק: {bank_coa} | {len(vendor_lookup):,} ספקים")
    st.markdown('<div style="text-align:center"><span class="badge badge-masav">🏦 MASAV</span></div>', unsafe_allow_html=True)

    file_bytes, fname = file_selector(client_id, "masav", "העלי קובץ MASAV (XLSX)", ["xlsx"], "masav_up")
    if not file_bytes: return

    with st.spinner("⏳ מעבד MASAV..."):
        rows_data, batches, errors, unmatched = parse_masav(file_bytes, vendor_lookup, bank_coa)

    if errors:
        with st.expander(f"⚠️ {len(errors)} שגיאות"):
            for e in errors: st.warning(e)
    if not rows_data:
        st.error("לא נמצאו תנועות תקינות"); return

    # ===== טיפול בספקים חסרים =====
    if unmatched:
        st.error(f"⚠️ {len(unmatched)} ספקים ללא ח\"ן — יש למלא לפני הורדה")

        # בנה vendor index עם שמות לצורך fuzzy match
        vendor_with_names = {
            hp: {"account": acc, "name": ""}
            for hp, acc in vendor_lookup.items()
        }

        new_mappings = {}

        for item in unmatched:
            vendor_name = item["vendor"]
            hp          = item["hp"]
            amount      = item["amount"]

            with st.container():
                st.markdown(f"**{vendor_name}** | ח.פ: `{hp}` | סכום: ₪{amount:,.0f}")

                # fuzzy match מול האינדקס
                sugg_name, sugg_acc, sugg_hp, ratio = fuzzy_match_vendor(
                    vendor_name, vendor_with_names
                )

                col1, col2, col3 = st.columns([3, 2, 1])

                if sugg_acc and ratio >= 0.7:
                    col1.info(f"💡 אולי: **{sugg_name}** → ח\"ן {sugg_acc} ({ratio:.0%})")
                    if col2.button(f"✅ כן, זה אותו ספק", key=f"match_{hp}"):
                        new_mappings[hp] = sugg_acc

                manual_acc = col3.text_input("ח\"ן ידני", key=f"manual_{hp}", placeholder="לדוגמה: 1298")
                if manual_acc.strip():
                    new_mappings[hp] = manual_acc.strip()

        if new_mappings:
            if st.button("💾 שמור מיפויים חדשים ועדכן", type="primary"):
                # עדכן vendor_index ב-Supabase
                updated_index = {**vendor_lookup, **new_mappings}
                try:
                    get_litay_db().table("client_config") \
                        .update({"vendor_index": updated_index}) \
                        .eq("client_id", client_id).execute()
                    st.success(f"✅ נשמרו {len(new_mappings)} מיפויים חדשים!")
                    st.rerun()
                except Exception as e:
                    st.error(f"שגיאה: {e}")
            return  # לא מאפשר הורדה עד שהכל מלא

    # ===== תצוגה מקדימה =====
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("תנועות", len(rows_data))
    c2.metric("סה״כ", f"₪{sum(r['amount'] for r in rows_data):,.0f}")
    c3.metric("Batches", len(batches))
    c4.metric("✅ ח\"ן", f"{len(rows_data)}/{len(rows_data)}")

    st.subheader("תצוגה מקדימה")
    prev = [{'תאריך':r['date_fmt'],'תיאור':r['vendor'],'סכום':r['amount'],
             'ח"ן חובה':r['vendor_coa'] or '⚠️','ח"ן זכות':r['bank_coa'],'Batch':r['batch']}
            for r in rows_data[:50]]
    st.dataframe(pd.DataFrame(prev), use_container_width=True, hide_index=True)

    st.divider()
    with st.spinner("⏳ בונה Excel..."):
        wb, checks = build_masav_excel(rows_data, batches)
    today = datetime.now()
    fname = f"MASAV_{client_id}_{today.day}_{today.month}_{today.year}.xlsx"

    st.subheader("✅ בדיקות תקינות")
    cols = st.columns(3)
    cols[0].metric("תנועות", checks['total_rows'])
    cols[1].metric("סה״כ", f"₪{checks['total_amount']:,.2f}")
    cols[2].metric("✅ ח\"ן", f"{checks['coa_covered']}/{checks['total_rows']}")

    st.download_button(f"📥 הורד MASAV Excel ({len(rows_data)} ספקים)",
        workbook_to_bytes(wb), fname,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

def render_payit_tab():
    from utils.payit import extract_data_from_payit_pdf, create_mizrahi_excel

    st.markdown('<div style="text-align:center"><span class="badge badge-payit">💜 Pay-it → מזרחי</span></div>', unsafe_allow_html=True)
    file_bytes, fname = file_selector("dmwa", "payit", "📂 העלי קובץ PDF של Pay-it", ["pdf"], "payit_up")
    if not file_bytes: return

    with st.spinner("⏳ מעבד PDF..."):
        import io as _io
        data = extract_data_from_payit_pdf(_io.BytesIO(file_bytes))

    if not data:
        st.error("❌ לא נמצאו נתונים בקובץ"); return

    df = pd.DataFrame([{
        'שם המוטב': d['fund_name'], 'בנק': d['account'].split('-')[0] if '-' in d['account'] else '',
        'מספר חשבון': d['account'], 'סכום (₪)': float(d['amount'])
    } for d in data])

    c1,c2 = st.columns(2)
    c1.metric("העברות", len(df))
    c2.metric("סה״כ", f"₪{df['סכום (₪)'].sum():,.2f}")

    st.subheader("תצוגה מקדימה")
    st.dataframe(df, use_container_width=True, hide_index=True)

    excel_data = create_mizrahi_excel(data)
    fname = f"מזרחי_גמל_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button("📥 הורד Excel למזרחי", excel_data, fname,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")


def render_valley_payem_tab(client):
    from utils.valley_payem import (parse_valley, build_valley_excel,
                                     parse_payem, build_payem_excel,
                                     detect_file_type, clean_description,
                                     parse_amount, parse_net_amount,
                                     parse_valley_date, parse_payem_date, format_date)
    from utils.db import load_coa_litay, check_duplicates_litay, save_keys_litay

    client_id = client["client_id"]
    f = st.file_uploader("העלי קובץ CSV", type=["csv"], key="csv_up")
    if not f: return

    raw = f.read()
    rows = []
    for enc in ("utf-8-sig", "utf-8", "windows-1255", "latin-1"):
        try:
            decoded = raw.decode(enc)
            candidate = list(csv.reader(io.StringIO(decoded)))
            ft = detect_file_type(candidate)
            if ft != "unknown":
                rows = candidate
                break
            if not rows:
                rows = candidate  # שמור כגיבוי
        except:
            pass
    ftype = detect_file_type(rows)

    if ftype == "valley":
        st.markdown('<div style="text-align:center"><span class="badge badge-valley">🏦 Valley Bank</span></div>', unsafe_allow_html=True)
    elif ftype == "payem":
        st.markdown('<div style="text-align:center"><span class="badge badge-payem">💳 PayEm</span></div>', unsafe_allow_html=True)
    else:
        st.error("לא ניתן לזהות את סוג הקובץ"); return

    coa = load_coa_litay(client_id, ftype)
    if ftype == "valley":
        all_parsed = parse_valley(rows, {**coa}, VALLEY_CAT_COA)
    else:
        ltd_coa = load_coa_litay(client_id, "ltd")
        all_parsed = parse_payem(rows, {**coa}, {**ltd_coa})

    known = check_duplicates_litay(client_id, ftype, {t["key"] for t in all_parsed})
    new_txns = [t for t in all_parsed if t["key"] not in known]
    dupes = [t for t in all_parsed if t["key"] in known]

    c1,c2,c3 = st.columns(3)
    c1.metric("סה״כ", len(all_parsed)); c2.metric("חדשות", len(new_txns)); c3.metric("כבר עובדו", len(dupes))

    if not new_txns:
        st.warning("כל התנועות כבר עובדו")
        if not st.checkbox("הורד מחדש"): return
        data = all_parsed
    elif dupes:
        data = new_txns if st.checkbox("רק תנועות חדשות", value=True) else all_parsed
    else:
        data = all_parsed

    st.subheader("תצוגה מקדימה")
    if ftype == "valley":
        prev = [{"תאריך":t["date_formatted"],"תיאור":t["description"],"סכום":t["amount"],
                 "קטגוריה":t["category"],'ח"ן חובה':t["coa_debit"],'ח"ן זכות':t["coa_credit"]} for t in data[:50]]
    else:
        prev = [{"תיאור":t["description"],"סכום":t["net_amount"],'ח"ן זכות':t["coa_credit"],
                 'ח"ן חובה':t["coa_debit"],"תאריך":t["date_formatted"],"שיוך":t["subsidiary"]} for t in data[:50]]
    st.dataframe(pd.DataFrame(prev), use_container_width=True, hide_index=True)

    st.divider()
    wb, checks = build_valley_excel(data) if ftype == "valley" else build_payem_excel(data)
    today = datetime.now()
    fname = f"{'Valley' if ftype=='valley' else 'PayEm'}_{client_id}_{today.day}_{today.month}_{today.year}.xlsx"

    st.subheader("✅ בדיקות תקינות")
    if ftype == "valley":
        cols = st.columns(3)
        cols[0].metric(f"{'✅' if checks['amount_ok'] else '❌'} סכומים", f"${checks['total_amount']:,.0f}")
        cols[1].metric("יתרה", f"${checks['expected_balance']:,.0f}")
        cols[2].metric(f"{'✅' if checks['no_coa']==0 else '⚠️'} ח\"ן", f"{checks['coa_covered']}/{checks['total_count']}")
    else:
        cols = st.columns(3)
        cols[0].metric(f"{'✅' if checks['count_ok'] else '❌'} שורות", f"INC {checks['count_inc']} + LTD {checks['count_ltd']}")
        cols[1].metric(f"{'✅' if checks['debit_ok'] else '❌'} חובה", f"${checks['data_debit']:,.0f}")
        cols[2].metric(f"{'✅' if checks['credit_ok'] else '❌'} זכות", f"${checks['data_credit']:,.0f}")

    st.download_button(f"📥 הורד Excel ({len(data)} תנועות)", workbook_to_bytes(wb), fname,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    st.info("⚠️ לאחר שווידאת שהקובץ הורד — לחצי לשמור בהיסטוריה")
    if st.button("💾 שמור בהיסטוריה", type="secondary"):
        save_keys_litay(client_id, ftype, [t["key"] for t in data], f.name)
        st.success(f"✅ {len(data)} תנועות נשמרו!"); st.rerun()


def render_income_tab():
    from utils.db import (get_all_clients, get_clients_full, add_client,
                           get_next_account_number, import_clients_from_df,
                           get_account_columns, add_account_column, read_excel_safe)
    from utils.converter import convert_income_file, create_excel_output, create_allocation_report, detect_month_label

    st.markdown('<div style="text-align:center"><span class="badge badge-income">📊 ממיר הכנסות</span></div>', unsafe_allow_html=True)

    page = st.session_state.get("rahel_page", "עיבוד")
    col1, col2 = st.columns([1,5])
    with col1:
        if st.button("📊 עיבוד", type="primary" if page=="עיבוד" else "secondary"):
            st.session_state["rahel_page"] = "עיבוד"; st.rerun()
        if st.button("⚙️ הגדרות", type="primary" if page=="הגדרות" else "secondary"):
            st.session_state["rahel_page"] = "הגדרות"; st.rerun()

    st.divider()

    if page == "עיבוד":
        uploaded_files = st.file_uploader("העלי קובץ הכנסות חודשי (.xlsx)",
                                           type=["xlsx"], accept_multiple_files=True, key="income_up")
        if not uploaded_files:
            st.info("⬆️ העלי קובץ אחד או יותר כדי להתחיל"); return

        for uploaded in uploaded_files:
            df = pd.read_excel(uploaded, header=None)
            month_label = detect_month_label(uploaded.name, df)
            st.success(f"✅ **{uploaded.name}** | חודש: **{month_label}**")

        st.divider()
        if "income_processing" not in st.session_state:
            st.session_state["income_processing"] = False

        if st.button("🚀 עבד והכן קבצי ייבוא", type="primary", use_container_width=True):
            st.session_state["income_processing"] = True

        if st.session_state["income_processing"]:
            clients_dict = get_all_clients()
            extra_cols = get_account_columns()
            all_unmatched, all_unknown_cols = [], []

            dfs = []
            for uploaded in uploaded_files:
                uploaded.seek(0)
                df = read_excel_safe(uploaded)
                dfs.append((uploaded.name, df))

            for name, df in dfs:
                _, _, _, unmatched, unknown_cols = convert_income_file(df, clients_dict, extra_cols)
                for c in unmatched:
                    if c not in all_unmatched: all_unmatched.append(c)
                for c in unknown_cols:
                    if c not in all_unknown_cols: all_unknown_cols.append(c)

            if all_unknown_cols:
                st.error("⚠️ עמודות לא מוכרות — הגדירי חשבון")
                for col_name in all_unknown_cols:
                    with st.form(f"col_{col_name}"):
                        st.markdown(f"**עמודה: {col_name}**")
                        c1, c2 = st.columns([2, 1])
                        acct = c1.number_input("חשבון", min_value=1000, max_value=9999, step=1)
                        exempt = c2.checkbox('פטור ממע"מ')
                        if st.form_submit_button("💾 שמור"):
                            add_account_column(col_name, int(acct), exempt); st.rerun()

            if all_unmatched:
                st.error(f"⚠️ {len(all_unmatched)} לקוחות לא נמצאו")
                next_acct = get_next_account_number()
                for i, client_name in enumerate(all_unmatched):
                    with st.form(f"cl_{client_name}"):
                        st.markdown(f"**{client_name}**")
                        acct = st.number_input("מספר חשבון", min_value=3000, max_value=3999,
                                               step=1, value=next_acct + i)
                        if st.form_submit_button("💾 שמור"):
                            add_client(client_name, int(acct)); st.rerun()

            if not all_unmatched and not all_unknown_cols:
                clients_dict = get_all_clients()
                extra_cols = get_account_columns()
                clients_full = get_clients_full()
                clients_id = {r["name"]: r.get("id_number", "") for r in clients_full}
                all_allocation = []

                for fname, df in dfs:
                    month_label = detect_month_label(fname, df)
                    invoice_rows, receipt_rows, allocation_rows, _, _ = convert_income_file(
                        df, clients_dict, extra_cols, clients_id)
                    all_allocation.extend(allocation_rows)

                    st.divider()
                    c1,c2,c3 = st.columns(3)
                    c1.metric("חשבוניות", len(invoice_rows))
                    c2.metric("קבלות", len(receipt_rows))
                    c3.metric("חודש", month_label)

                    excel_data = create_excel_output(invoice_rows, receipt_rows, month_label)
                    st.download_button(
                        label=f"⬇️ הורד ייבוא — {month_label}",
                        data=excel_data,
                        file_name=f"Hashavshevet_{month_label.replace('.', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, type="primary", key=f"dl_{month_label}")

                if all_allocation:
                    st.divider()
                    st.warning(f"⚠️ נמצאו **{len(all_allocation)}** חשבוניות הדורשות מספר הקצאה")
                    alloc_data = create_allocation_report(all_allocation)
                    st.download_button("📋 הורד דוח הקצאה", data=alloc_data,
                        file_name="דוח_הקצאה.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="dl_allocation")

    elif page == "הגדרות":
        st.subheader("⚙️ הגדרות")
        tab1, tab2 = st.tabs(["👥 לקוחות", "🔢 חשבונות"])
        with tab1:
            st.subheader("ייבוא מאינדקס")
            idx_file = st.file_uploader("CLients_Index_Rachel.xlsx", type=["xlsx"], key="idx")
            if idx_file:
                if st.button("📥 ייבא לקוחות"):
                    try:
                        df_idx = read_excel_safe(idx_file)
                        count, err = import_clients_from_df(df_idx)
                        if err and count == 0: st.error(f"❌ {err}")
                        elif err: st.warning(f"⚠️ {err}"); st.rerun()
                        else: st.success(f"✅ יובאו {count} לקוחות"); st.rerun()
                    except Exception as e:
                        st.error(f"❌ שגיאה: {str(e)}")
            st.divider()
            st.subheader("הוסף לקוח")
            with st.form("add_client"):
                c1, c2, c3 = st.columns([3, 1, 2])
                name = c1.text_input("שם לקוח")
                acct = c2.number_input("חשבון", min_value=3000, max_value=3999, step=1, value=get_next_account_number())
                id_num = c3.text_input("ת.ז / מס.ע.מ")
                if st.form_submit_button("➕ הוסף"):
                    if name.strip(): add_client(name.strip(), int(acct), id_num.strip()); st.rerun()
            st.divider()
            rows = get_clients_full()
            if rows:
                df_show = pd.DataFrame(rows)[["account_number","name","id_number"]]
                df_show.columns = ["חשבון","שם לקוח","ת.ז / מס.ע.מ"]
                st.dataframe(df_show, use_container_width=True, hide_index=True)
        with tab2:
            st.subheader("חשבונות מובנים")
            st.table(pd.DataFrame([
                {"עמודה":"שכט","חשבון":5000,'מע"מ':"חייב 18%"},
                {"עמודה":"נוטריון","חשבון":5004,'מע"מ':"חייב 18%"},
                {"עמודה":"הוצאות","חשבון":5002,'מע"מ':"פטור"},
            ]))
            st.divider()
            st.subheader("עמודות מותאמות")
            extra = get_account_columns()
            if extra:
                for col, cfg in extra.items():
                    st.write(f"**{col}** → {cfg['account']} | {'פטור' if cfg['vat_exempt'] else 'חייב'} מע\"מ")
            else:
                st.info("אין עדיין")


# ===== MAIN =====
def main():
    if not check_auth(): return
    init_session()
    inject_css()

    with st.sidebar:
        try:
            _ = st.secrets["auth"]["users"]
            u = st.session_state.get("auth_user","")
            if u: st.markdown(f"👤 **{u}**")
            if st.button("🚪 התנתק", use_container_width=True):
                st.session_state["authenticated"] = False; st.rerun()
            st.divider()
        except: pass

        st.subheader("👤 בחירת לקוח")
        clients = st.session_state["clients"]

        CLIENT_NAMES = {
            "dmwa":       "ד.מ. פוסט בע\"מ",
            "rahel_mor":  "רחל מור",
            "brandlight": "Brandlight",
        }
        # הוסף לקוחות מ-Supabase + לקוחות קבועים
        all_names = {}
        for cid, cname in CLIENT_NAMES.items():
            if cid in clients:
                all_names[clients[cid]["client_name"]] = cid
            else:
                all_names[cname] = cid

        selected_name = st.selectbox("לקוח", list(all_names.keys()), label_visibility="collapsed")
        selected_id = all_names[selected_name]
        st.session_state["selected_client"] = selected_id

        if selected_id in clients:
            client = clients[selected_id]
            vendor_count = len(client.get("vendor_index") or {})
            if vendor_count > 0:
                st.success(f"✅ {vendor_count:,} ספקים")
            if client.get("bank_coa"):
                st.info(f'ח"ן בנק: **{client["bank_coa"]}**')

        if st.button("🔄 רענן", use_container_width=True):
            from utils.db import load_clients_litay
            st.session_state["clients"] = load_clients_litay(); st.rerun()

        # ===== סרגל צד דינמי לפי לקוח =====
        st.divider()

        if selected_id == "dmwa":
            # ── אינדקס ספקים ──
            st.caption("📋 אינדקס ספקים (מחשבשבת)")
            from utils.db import upload_file, list_recent_files, download_file
            from utils.masav import read_vendor_index_xlsx

            recent_idx = list_recent_files(selected_id, "vendor_index")
            tab_new, tab_saved = st.tabs(["📤 העלה חדש", f"📂 שמורים ({len(recent_idx)})"])

            idx_bytes = None
            idx_name  = None

            with tab_new:
                idx_file = st.file_uploader("XLSX מחשבשבת", type=["xlsx"], key="idx_vendors")
                if idx_file:
                    idx_bytes = idx_file.read()
                    idx_name  = idx_file.name
                    upload_file(selected_id, "vendor_index", idx_name, idx_bytes)

            with tab_saved:
                if not recent_idx:
                    st.info("אין אינדקסים שמורים")
                elif idx_bytes is None:
                    opts = {f["name"]: f for f in recent_idx}
                    sel  = st.selectbox("בחרי אינדקס", list(opts.keys()), key="idx_sel")
                    if st.button("📂 טעני", key="idx_load"):
                        idx_bytes = download_file(selected_id, "vendor_index", sel)
                        idx_name  = sel

            if idx_bytes:
                with st.spinner("קורא אינדקס..."):
                    lookup, errs = read_vendor_index_xlsx(idx_bytes)
                if errs:
                    for e in errs: st.error(e)
                elif lookup:
                    existing = clients.get(selected_id, {}).get("vendor_index") or {}
                    merged   = {**existing, **lookup}
                    try:
                        from utils.db import get_litay_db
                        get_litay_db().table("client_config") \
                            .update({"vendor_index": merged}) \
                            .eq("client_id", selected_id).execute()
                        st.success(f"✅ {len(lookup):,} ספקים עודכנו!")
                        from utils.db import load_clients_litay
                        st.session_state["clients"] = load_clients_litay()
                        st.session_state.pop("masav_up", None)
                        st.rerun()
                    except Exception as e:
                        st.error(f"שגיאה: {e}")

        elif selected_id == "rahel_mor":
            # ── אינדקס לקוחות ──
            st.caption("👥 אינדקס לקוחות (מחשבשבת)")
            idx_file = st.file_uploader("העלי XLSX לקוחות", type=["xlsx"], key="idx_clients")
            if idx_file:
                from utils.db import read_excel_safe, import_clients_from_df
                try:
                    df = read_excel_safe(idx_file)
                    count, err = import_clients_from_df(df)
                    if err and count == 0: st.error(f"❌ {err}")
                    elif err: st.warning(f"⚠️ {err}")
                    else: st.success(f"✅ {count} לקוחות יובאו!")
                except Exception as e:
                    st.error(f"שגיאה: {e}")

        elif selected_id == "brandlight":
            # ── ח"ן JSON ──
            st.caption("📊 ייבוא ח\"ן (JSON)")
            coa_file = st.file_uploader("העלי coa_lookup.json", type=["json"], key="coa_json")
            if coa_file:
                try:
                    import json
                    coa_data = json.loads(coa_file.read().decode("utf-8"))
                    st.session_state["coa"] = coa_data
                    st.success("✅ ח\"ן עודכן!")
                except Exception as e:
                    st.error(f"שגיאה: {e}")

    # טאבים לפי לקוח
    modules = CLIENT_MODULES.get(selected_id, [])

    if selected_id == "dmwa":
        tab1, tab2 = st.tabs(["🏦 MASAV", "💜 Pay-it → מזרחי"])
        client = clients.get(selected_id, {"client_id": selected_id, "client_name": "DMWA", "bank_coa": "", "vendor_index": {}})
        with tab1: render_masav_tab(client)
        with tab2: render_payit_tab()

    elif selected_id == "rahel_mor":
        render_income_tab()

    elif selected_id == "brandlight":
        tab1, = st.tabs(["📄 Valley Bank / PayEm"])
        client = clients.get(selected_id, {"client_id": selected_id, "client_name": "Brandlight", "bank_coa": "", "vendor_index": {}})
        with tab1: render_valley_payem_tab(client)

    else:
        st.info(f"לקוח {selected_name} — אין מודולים מוגדרים עדיין")

    st.markdown(f'<div class="app-footer">Litay Ltd · {datetime.now().year}</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
