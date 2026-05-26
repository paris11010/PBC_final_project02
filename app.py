"""
精品採購管理系統 — 全模組整合版（完整歷史紀錄）
=================================================
pip install streamlit pandas openpyxl plotly numpy
"""

import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import hashlib
import re
import io
import random
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, datetime, timedelta

st.set_page_config(page_title="精品採購管理系統", page_icon="💎", layout="wide")
DB_PATH = "procurement.db"

# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
    .stApp { background-color: #f0f4f8; }
    .banner { background: linear-gradient(135deg, #0a1628, #1a365d, #2b4c7e); padding: 2rem; border-radius: 12px; margin-bottom: 1.5rem; }
    .banner h1 { color: #e2e8f0; text-align: center; font-size: 2rem; margin: 0; letter-spacing: 2px; }
    .banner p { color: #90a4c4; text-align: center; margin-top: 0.5rem; }
    .card { background: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); border-left: 4px solid #2b4c7e; margin-bottom: 1rem; }
    .card h3 { color: #1a365d; margin-top: 0; }
    .step-badge { background: #2b4c7e; color: white; padding: 0.3rem 0.8rem; border-radius: 20px; font-size: 0.8rem; display: inline-block; margin-bottom: 0.5rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="banner"><h1>💎 LUXURY PROCUREMENT SYSTEM</h1><p>精品採購管理系統</p></div>', unsafe_allow_html=True)

# ============================================================
# 類別對應
# ============================================================
big_cat_map = {'RTW':'服飾 (RTW)','HB':'包款 (HB)','SLG':'小皮件 (SLG)','SH':'鞋履 (SH)','ACC':'配飾 (ACC)','JWL':'珠寶 (JWL)'}
sub_cat_map = {
    'TP':'上衣','BT':'下著','DR':'洋裝','OW':'外套',
    'SB':'肩背包','TT':'托特包','CB':'斜背包','CL':'手拿包','BP':'後背包',
    'LW':'長夾','SW':'短夾','CH':'卡夾','PC':'護照套',
    'SN':'運動鞋','LF':'樂福鞋','HL':'高跟鞋','BO':'靴子','SD':'涼鞋',
    'SC':'絲巾','BL':'皮帶','HT':'帽子','NL':'項鍊','BC':'手鍊','RG':'戒指','ER':'耳環'
}

# ============================================================
# 範例資料生成
# ============================================================
random.seed(42)
STORES = ["信義A8","忠孝SOGO","南西店","板橋店","台中新光"]
COLORS = ["黑色","白色","棕色","藍色","紅色","裸色"]
STYLE_NAMES = ["Classic Flap","Tote Bag","Mini Crossbody","Ankle Boot","Silk Scarf",
               "Logo Belt","Chain Necklace","Loafer","Sneaker","Card Holder"]

def gen_sales_sample():
    random.seed(42); records = []
    for month in range(1, 13):
        d = date(2025, month, random.randint(1,28))
        for store in STORES:
            sw = {"信義A8":1.5,"忠孝SOGO":1.3,"南西店":1.0,"板橋店":0.7,"台中新光":0.8}[store]
            for big, subs in [('HB',['SB','TT','CB','CL']),('SH',['SN','LF','BO','SD']),('ACC',['SC','BL','HT']),('RTW',['TP','DR','OW','BT']),('SLG',['LW','SW','CH']),('JWL',['NL','BC','RG','ER'])]:
                for sub in subs:
                    qty = max(1, int(10*sw + random.randint(-3,5)))
                    price = random.choice([8000,12000,18000,25000,38000,55000,78000])
                    sn = f"{big}-{sub}-{random.randint(100,999)}"
                    records.append({'門市':store,'大類別':big,'小類別':sub,'日期':str(d),'貨號':sn,'品名':sub_cat_map.get(sub,sub),
                        '顏色':random.choice(COLORS),'尺寸':random.choice(['S','M','L','0 (ONE SIZE)']),
                        '類型':random.choice(['經典','時尚']),'實售價格':price,'數量':qty,
                        '銷售人員':f'員工{random.randint(1,20)}',
                        '會員編號':f'VIP{random.randint(1000,9999)}' if random.random()>0.3 else '',
                        '會員名稱':f'會員{random.randint(1,200)}' if random.random()>0.3 else None,
                        '會員生日':f'{random.randint(1965,2005)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}' if random.random()>0.3 else None})
    df = pd.DataFrame(records)
    df['銷售總額']=df['實售價格']*df['數量']; df['大類別名稱']=df['大類別'].map(big_cat_map).fillna(df['大類別'])
    df['小類別名稱']=df['小類別'].map(sub_cat_map).fillna(df['小類別']); df['日期_dt']=pd.to_datetime(df['日期'],errors='coerce')
    df['月份']=df['日期_dt'].dt.strftime('%m月'); df['類型']=df['類型'].fillna('未分類')
    if '會員生日' in df.columns:
        df['會員生日_dt']=pd.to_datetime(df['會員生日'],errors='coerce'); df['年齡']=2025-df['會員生日_dt'].dt.year
        df['年齡層']=pd.cut(df['年齡'],bins=[0,20,30,40,50,60,120],labels=['20歲以下','21-30歲','31-40歲','41-50歲','51-60歲','61歲以上'])
    return df

def gen_assortment_sample():
    random.seed(42); records = []
    for big, subs in [('HB',['SB','TT','CB']),('SH',['SN','LF','BO']),('ACC',['SC','BL']),('RTW',['TP','DR','OW']),('SLG',['LW','SW']),('JWL',['NL','RG'])]:
        for sub in subs:
            for i in range(3):
                style = f"SS27{big}{sub}{random.randint(100,999)}"
                records.append({'Style Number':style,'Style Name':f"{sub_cat_map.get(sub,sub)} {chr(65+i)}",'Item Number':f"SS27{big}{sub}{style[-3:]}",
                    'Color Code':f'{random.choice("ABCDEFG")}{random.randint(100,999)}','Color Name':random.choice(COLORS),
                    'Category':big,'Subcategory':sub,'Season':'SS27','PLC':random.choice(['Carryover','Seasonal','Seasonal']),
                    'Wholesale Price(€)':random.randint(150,800),'Suggested Retail Price(€)':random.randint(400,2500),'Total Order Quantity':0,'Style Image':f'IMG-{style}'})
    return pd.DataFrame(records)

A_keywords = ["包裝損傷","外盒破損","保護膜刮痕","輕微壓痕"]
B_keywords = ["刮傷","磨損","掉色","脫線","五金刮傷","車線歪斜","拉鍊不順","展示痕跡","小污漬","輕微掉色"]
C_keywords = ["破裂","斷裂","發霉","腐蝕","破損","嚴重染色","缺件","變形"]

def gen_defect_sample():
    random.seed(42); all_kw = A_keywords+B_keywords+C_keywords; records = []
    for i in range(20):
        style = f"SS27-{random.choice(['HB','SH','ACC','RTW'])}-{random.randint(100,999)}"
        records.append({'Style Image':f'IMG-{style}','Style Number':style,'Item Number':f"ITM{style.replace('-','')}",
            'Style Name':random.choice(STYLE_NAMES),'Color Code':f'{random.choice("ABCDEFG")}{random.randint(100,999)}','Color Name':random.choice(COLORS),
            'Defect Description':random.choice(all_kw)+"，需要檢查",'Repair Cost (NTD)':random.choice([0,500,1000,2000,5000]),
            'Wholesale Price(NTD)':random.randint(5000,40000),'Retail Price(NTD)':random.randint(15000,120000),'Supplier Fault':random.choice(['Yes','No'])})
    return pd.DataFrame(records)

def gen_return_sample():
    random.seed(42); records = []
    for i in range(15):
        style = f"SS27-{random.choice(['HB','SH','ACC','RTW'])}-{random.randint(100,999)}"
        purchase = date(2026, random.randint(3,5), random.randint(1,28)); req = purchase+timedelta(days=random.randint(1,45))
        records.append({'Style Image':f'IMG-{style}','Style Number':style,'Item Number':f"ITM{style.replace('-','')}",
            'Style Name':random.choice(STYLE_NAMES),'Color Code':f'{random.choice("ABCDEFG")}{random.randint(100,999)}','Color Name':random.choice(COLORS),
            'Product Condition':random.choice(['New','Lightly Used','Heavily Used']),'Limited Edition':random.choice(['Yes','No','No','No']),
            'Retail Price(NTD)':random.randint(15000,500000),'Purchase Date':purchase,'Request Date':req})
    return pd.DataFrame(records)

# ============================================================
# 資料庫
# ============================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL, brand TEXT NOT NULL, role TEXT NOT NULL, created TEXT NOT NULL)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, season TEXT NOT NULL,
        date TEXT NOT NULL, summary TEXT)""")
    for col in ['order_data','alloc_data','inventory_data','slow_data','defect_data','return_data']:
        try: conn.execute(f"ALTER TABLE sessions ADD COLUMN {col} TEXT")
        except: pass
    conn.commit(); conn.close()

init_db()

def hash_pw(pw): return hashlib.sha256(pw.encode()).hexdigest()

def register(username, password, brand, role):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("INSERT INTO users (username,password,brand,role,created) VALUES (?,?,?,?,?)",
                     (username, hash_pw(password), brand, role, str(date.today())))
        conn.commit(); conn.close(); return True, "註冊成功！請切換到登入頁面"
    except sqlite3.IntegrityError:
        conn.close(); return False, "此帳號已被使用"

def do_login(username, password):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("SELECT id,username,brand,role FROM users WHERE username=? AND password=?",
                       (username, hash_pw(password)))
    row = cur.fetchone(); conn.close()
    return {"id":row[0],"username":row[1],"brand":row[2],"role":row[3]} if row else None

def save_session_record(uid, season, summary, order_df=None, alloc_df=None, inv_df=None, slow_df=None, defect_df=None, return_df=None):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO sessions (user_id,season,date,summary,order_data,alloc_data,inventory_data,slow_data,defect_data,return_data) VALUES (?,?,?,?,?,?,?,?,?,?)",
        (uid, season, str(date.today()), summary,
         order_df.to_json() if order_df is not None and len(order_df)>0 else None,
         alloc_df.to_json() if alloc_df is not None and len(alloc_df)>0 else None,
         inv_df.to_json() if inv_df is not None and len(inv_df)>0 else None,
         slow_df.to_json() if slow_df is not None and len(slow_df)>0 else None,
         defect_df.to_json() if defect_df is not None and len(defect_df)>0 else None,
         return_df.to_json() if return_df is not None and len(return_df)>0 else None))
    conn.commit(); conn.close()

def get_history(uid):
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT id,season,date,summary FROM sessions WHERE user_id=? ORDER BY id DESC", conn, params=(uid,))
    conn.close(); return df

def get_session_detail(sid):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("SELECT * FROM sessions WHERE id=?", (sid,))
    row = cur.fetchone(); cols = [d[0] for d in cur.description]; conn.close()
    return dict(zip(cols, row)) if row else None

# ============================================================
# xlsx 工具
# ============================================================
def df_to_xlsx(df, sheet_title="Sheet1", title_text=None):
    wb = Workbook(); ws = wb.active; ws.title = sheet_title; sr = 1
    hf = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    hfill = PatternFill("solid", fgColor="1a365d")
    bdr = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    if title_text:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(df.columns),1))
        ws.cell(row=1, column=1, value=title_text).font = Font(bold=True, size=14, name="Arial")
        ws.cell(row=2, column=1, value=f"產出日期：{date.today()}")
        if "user" in st.session_state and st.session_state["user"]:
            u = st.session_state["user"]
            ws.cell(row=3, column=1, value=f"品牌：{u['brand']}　操作者：{u['username']}")
        sr = 5
    for ci, cn in enumerate(df.columns, 1):
        c = ws.cell(row=sr, column=ci, value=cn); c.font=hf; c.fill=hfill; c.alignment=Alignment(horizontal="center"); c.border=bdr
    for ri, row in enumerate(df.itertuples(index=False), sr+1):
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=ri, column=ci, value=val); c.border=bdr; c.font=Font(name="Arial", size=10)
            if isinstance(val, (int, float)): c.alignment=Alignment(horizontal="right"); c.number_format="#,##0"
    for ci in range(1, len(df.columns)+1):
        ml = max(len(str(ws.cell(row=sr, column=ci).value or "")),
                 *[len(str(ws.cell(row=r, column=ci).value or "")) for r in range(sr+1, sr+len(df)+1)])
        ws.column_dimensions[get_column_letter(ci)].width = min(ml+4, 30)
    out = io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

# ============================================================
# 瑕疵品分級邏輯
# ============================================================
def authenticate_item(style_image, style_number, item_number, style_name, color_code, color_name):
    errors = []
    if not str(style_image).strip(): errors.append("Style Image 為空")
    if not re.match(r"^[A-Za-z0-9\-_]{6,20}$", str(style_number).strip()): errors.append(f"Style Number 格式錯誤")
    if not re.match(r"^[A-Za-z0-9\-_]{6,20}$", str(item_number).strip()): errors.append(f"Item Number 格式錯誤")
    if not str(style_name).strip(): errors.append("Style Name 為空")
    if not re.match(r"^[A-Za-z0-9]{2,6}$", str(color_code).strip()): errors.append(f"Color Code 格式錯誤")
    if not str(color_name).strip(): errors.append("Color Name 為空")
    return (True, "OK") if not errors else (False, "; ".join(errors))

def classify_defect(row):
    try:
        is_auth, reason = authenticate_item(row.get("Style Image",""), row.get("Style Number",""),
            row.get("Item Number",""), row.get("Style Name",""), row.get("Color Code",""), row.get("Color Name",""))
        if not is_auth: return ("No","Counterfeit","Destroy",f"驗證失敗: {reason}")
        desc = str(row.get("Defect Description",""))
        supplier = str(row.get("Supplier Fault","")).strip()
        retail = float(row.get("Retail Price(NTD)",0) or 0)
        wholesale = float(row.get("Wholesale Price(NTD)",0) or 0)
        repair = float(row.get("Repair Cost (NTD)",0) or 0)
        for kw in C_keywords:
            if kw in desc:
                if supplier=="Yes": return ("Yes","Grade C","退回供應商",f"偵測到「{kw}」— 嚴重瑕疵，供應商責任")
                else: return ("Yes","Grade C","銷毀",f"偵測到「{kw}」— 嚴重瑕疵")
        for kw in B_keywords:
            if kw in desc:
                if retail==0: return ("Yes","Grade B","待評估","零售價為 0")
                margin = (retail-wholesale-repair)/retail
                if margin>=0.20: return ("Yes","Grade B","維修後銷售",f"偵測到「{kw}」— 利潤率 {margin:.1%} ≥ 20%")
                else: return ("Yes","Grade B","待評估",f"偵測到「{kw}」— 利潤率 {margin:.1%} < 20%")
        for kw in A_keywords:
            if kw in desc: return ("Yes","Grade A","正常銷售",f"偵測到「{kw}」— 輕微包裝瑕疵")
        return ("Yes","Unknown","人工審查","未匹配到瑕疵關鍵字")
    except Exception as e: return ("Error","系統錯誤","檢查資料",str(e))

# ============================================================
# 退換貨邏輯
# ============================================================
def evaluate_return(row):
    try:
        is_auth, reason = authenticate_item(row.get("Style Image",""), row.get("Style Number",""),
            row.get("Item Number",""), row.get("Style Name",""), row.get("Color Code",""), row.get("Color Name",""))
        if not is_auth: return ("拒絕", f"驗證失敗: {reason}")
        purchase = pd.to_datetime(row.get("Purchase Date")); request = pd.to_datetime(row.get("Request Date"))
        days = (request-purchase).days
        if days>30: return ("拒絕", f"超過退貨期限 ({days} 天)")
        if str(row.get("Limited Edition","")).strip()=="Yes": return ("拒絕", "限量款不可退貨")
        condition = row.get("Product Condition","")
        if condition=="Heavily Used": return ("拒絕", "商品有明顯使用痕跡")
        elif condition=="Lightly Used": return ("換貨或購物金", "商品有輕微使用痕跡")
        price = float(row.get("Retail Price(NTD)",0) or 0)
        if price>=300000: return ("人工審查", "高單價精品需主管確認")
        return ("核准退貨", "符合退貨條件")
    except Exception as e: return ("錯誤", str(e))

# ============================================================
# 滯銷品邏輯
# ============================================================
def rule_engine(row, current_date):
    if row.get("is_carryover", False): return "維持正價銷售"
    try: launch_dt = datetime.strptime(str(row.get("launch_date","2026-04-01")), "%Y-%m-%d")
    except: launch_dt = datetime(2026,4,1)
    months = (current_date.year-launch_dt.year)*12+(current_date.month-launch_dt.month)
    if months>12:
        if row.get("brand_tier")=="Ultra_Luxury": return "銷毀"
        return "送 Outlet"
    elif months>6: return "員工特賣"
    else: return "跨店調撥"

# ============================================================
# 登入 / 註冊
# ============================================================
if "user" not in st.session_state: st.session_state["user"] = None

if st.session_state["user"] is None:
    tab_l, tab_r = st.tabs(["🔑 登入", "📝 註冊新帳號"])
    with tab_l:
        st.subheader("登入")
        lu = st.text_input("帳號", key="lu", placeholder="輸入帳號")
        lp = st.text_input("密碼", type="password", key="lp", placeholder="輸入密碼")
        if st.button("登入", type="primary", key="bl"):
            if lu and lp:
                r = do_login(lu, lp)
                if r: st.session_state["user"]=r; st.rerun()
                else: st.error("帳號或密碼錯誤")
    with tab_r:
        st.subheader("註冊新帳號")
        ru = st.text_input("帳號", key="ru"); rp = st.text_input("密碼", type="password", key="rp")
        rp2 = st.text_input("確認密碼", type="password", key="rp2")
        rb = st.text_input("品牌名稱", key="rb", placeholder="例如：Celine")
        rr = st.selectbox("角色", ["採購買手 Buyer","採購主管 Manager","營運經理 Operations"], key="rr")
        if st.button("註冊", type="primary", key="br"):
            if not ru or not rp or not rb: st.warning("請填寫所有欄位")
            elif rp!=rp2: st.error("密碼不一致")
            elif len(rp)<4: st.error("密碼至少4個字")
            else:
                ok, msg = register(ru, rp, rb, rr)
                if ok: st.success(msg)
                else: st.error(msg)
    st.stop()

# ============================================================
# 已登入
# ============================================================
user = st.session_state["user"]
if user is None: st.session_state.clear(); st.rerun()

st.sidebar.markdown(f"### 👤 {user['username']}")
st.sidebar.markdown(f"**品牌：** {user['brand']}　**角色：** {user['role']}")
st.sidebar.divider()
nav = st.sidebar.radio("導航", ["🏠 首頁","📦 新增採購作業","📜 歷史紀錄"])
if st.sidebar.button("🚪 登出"): st.session_state.clear(); st.rerun()

# ============================================================
# 首頁
# ============================================================
if nav == "🏠 首頁":
    st.markdown(f'<div class="card"><h3>歡迎回來，{user["username"]}！</h3><p>品牌：{user["brand"]}　角色：{user["role"]}</p></div>', unsafe_allow_html=True)
    history = get_history(user["id"])
    st.metric("歷史作業次數", f"{len(history)} 次")
    if len(history)>0:
        st.subheader("最近作業")
        for _, r in history.head(5).iterrows():
            st.write(f"📅 **{r['date']}** — {r['season']}　{r['summary'] or ''}")
    else: st.info("還沒有作業紀錄，點左邊「新增採購作業」開始！")

# ============================================================
# 歷史紀錄（完整展開版）
# ============================================================
elif nav == "📜 歷史紀錄":
    st.title("📜 歷史紀錄")
    history = get_history(user["id"])
    if len(history)==0: st.info("無紀錄")
    else:
        for _, r in history.iterrows():
            with st.expander(f"📅 {r['date']} — {r['season']}　{r['summary'] or ''}"):
                d = get_session_detail(r["id"])
                if not d: continue

                if d.get("order_data"):
                    st.subheader("一、叫貨清單")
                    df_o = pd.read_json(io.StringIO(d["order_data"]))
                    st.dataframe(df_o, use_container_width=True, hide_index=True)
                    ox = df_to_xlsx(df_o, "叫貨", "叫貨清單")
                    st.download_button("📥 叫貨清單.xlsx", ox, f"叫貨清單_{r['date']}.xlsx",
                                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_order_{r['id']}")

                if d.get("alloc_data"):
                    st.subheader("二、配貨結果")
                    df_al = pd.read_json(io.StringIO(d["alloc_data"]))
                    if "門市" in df_al.columns and "配貨" in df_al.columns:
                        st.plotly_chart(px.bar(df_al.groupby("門市")["配貨"].sum().reset_index(), x="門市", y="配貨", color="門市"), use_container_width=True)
                    st.dataframe(df_al, use_container_width=True, hide_index=True)

                if d.get("inventory_data"):
                    st.subheader("三、銷售與庫存")
                    df_iv = pd.read_json(io.StringIO(d["inventory_data"]))
                    if "store" in df_iv.columns and "quantity" in df_iv.columns and "sold" in df_iv.columns:
                        df_iv["remaining"] = df_iv["quantity"]-df_iv["sold"]
                        sm = df_iv.groupby("store").agg(配貨=("quantity","sum"),已售=("sold","sum"),剩餘=("remaining","sum")).reset_index()
                        sm["銷售率%"] = (sm["已售"]/sm["配貨"].replace(0,1)*100).round(1)
                        sm.columns = ["門市","配貨","已售","剩餘","銷售率%"]
                        st.dataframe(sm, use_container_width=True, hide_index=True)
                        st.plotly_chart(px.bar(sm, x="門市", y=["配貨","已售","剩餘"], barmode="group", title="各門市銷售狀況"), use_container_width=True)

                if d.get("slow_data"):
                    st.subheader("四、滯銷品")
                    df_sl = pd.read_json(io.StringIO(d["slow_data"]))
                    st.dataframe(df_sl, use_container_width=True, hide_index=True)

                if d.get("defect_data"):
                    st.subheader("五、瑕疵品")
                    df_de = pd.read_json(io.StringIO(d["defect_data"]))
                    st.dataframe(df_de, use_container_width=True, hide_index=True)

                if d.get("return_data"):
                    st.subheader("六、退換貨")
                    df_re = pd.read_json(io.StringIO(d["return_data"]))
                    st.dataframe(df_re, use_container_width=True, hide_index=True)

# ============================================================
# 新增採購作業
# ============================================================
elif nav == "📦 新增採購作業":
    STEPS = ["季度設定","銷售分析 Dashboard","生成 Assortment 叫貨單",
             "配貨給門市","門市實際銷售","庫存調整與調撥",
             "滯銷品處置","瑕疵品處理","退換貨處理","最終報告"]

    if "ps" not in st.session_state: st.session_state["ps"] = 0
    ps = st.session_state["ps"]
    def nxt(): st.session_state["ps"] += 1
    def prv(): st.session_state["ps"] -= 1

    if ps>0:
        st.progress(ps/(len(STEPS)-1), text=f"步驟 {ps}/{len(STEPS)-1}：{STEPS[ps]}")
        st.sidebar.divider(); st.sidebar.button("⬅️ 回上一步", on_click=prv)

    # === Step 0 ===
    if ps == 0:
        st.title("📦 新增採購作業")
        season = st.selectbox("操作季度", ["2026 SS 春夏","2026 AW 秋冬","2025 AW 秋冬"])
        if st.button("開始 →", type="primary"):
            st.session_state["p_season"]=season; nxt(); st.rerun()

    # === Step 1：銷售分析 ===
    elif ps == 1:
        st.markdown('<span class="step-badge">STEP 1</span>', unsafe_allow_html=True)
        st.title("銷售分析 Dashboard")
        method = st.radio("選擇資料來源", ["上傳 Excel","使用範例資料"], key="s1m")

        if method == "上傳 Excel":
            uploaded = st.file_uploader("上傳銷售 Excel（SALES 工作表）", type=["xlsx"], key="s1u")
            if uploaded:
                @st.cache_data
                def load_sales(file):
                    df = pd.read_excel(file, sheet_name='SALES')
                    df['銷售總額']=df['實售價格']*df['數量']; df['大類別名稱']=df['大類別'].map(big_cat_map).fillna(df['大類別'])
                    df['小類別名稱']=df['小類別'].map(sub_cat_map).fillna(df['小類別'])
                    if '日期' in df.columns: df['日期_dt']=pd.to_datetime(df['日期'],errors='coerce'); df['月份']=df['日期_dt'].dt.strftime('%m月')
                    if '類型' in df.columns: df['類型']=df['類型'].fillna('未分類')
                    if '會員生日' in df.columns:
                        df['會員生日_dt']=pd.to_datetime(df['會員生日'],errors='coerce'); df['年齡']=2025-df['會員生日_dt'].dt.year
                        df['年齡層']=pd.cut(df['年齡'],bins=[0,20,30,40,50,60,120],labels=['20歲以下','21-30歲','31-40歲','41-50歲','51-60歲','61歲以上'])
                    return df
                df = load_sales(uploaded); st.success(f"已上傳 {len(df)} 筆")
            else: st.info("等待上傳..."); st.stop()
        else:
            df = gen_sales_sample(); st.success(f"已載入範例資料 {len(df)} 筆")

        stores = st.multiselect("篩選門市", options=df['門市'].unique(), default=df['門市'].unique())
        fdf = df[df['門市'].isin(stores)]

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("總銷售額",f"${fdf['銷售總額'].sum():,.0f}"); c2.metric("售出數量",f"{fdf['數量'].sum():,.0f} 件")
        c3.metric("平均客單價",f"${fdf['實售價格'].mean():,.0f}"); c4.metric("活躍會員",f"{fdf[fdf['會員名稱'].notna()]['會員名稱'].nunique()} 人")
        st.divider()

        cA,cB = st.columns(2)
        with cA: st.plotly_chart(px.pie(fdf.groupby('大類別名稱')['銷售總額'].sum().reset_index(), values='銷售總額', names='大類別名稱', title='大類別佔比'), use_container_width=True)
        with cB: st.plotly_chart(px.bar(fdf.groupby('門市')['銷售總額'].sum().reset_index().sort_values('銷售總額',ascending=True), x='銷售總額', y='門市', orientation='h', title='門市排行', text_auto='.2s'), use_container_width=True)
        st.divider()

        if '銷售人員' in fdf.columns:
            st.subheader("金牌銷售員"); rep=fdf.groupby(['門市','銷售人員']).agg(銷售總額=('銷售總額','sum')).reset_index().sort_values('銷售總額',ascending=False)
            t10=rep.head(10).sort_values('銷售總額',ascending=True); t10['顯示']=t10['門市']+" - "+t10['銷售人員']
            st.plotly_chart(px.bar(t10,x='銷售總額',y='顯示',orientation='h',title='Top 10',text_auto='.2s'), use_container_width=True); st.divider()

        if '月份' in fdf.columns:
            st.subheader("月趨勢"); st.plotly_chart(px.line(fdf.groupby('月份')['銷售總額'].sum().reset_index().sort_values('月份'),x='月份',y='銷售總額',title='月銷售額',markers=True), use_container_width=True); st.divider()

        st.subheader("小類別排行"); st.plotly_chart(px.bar(fdf.groupby(['大類別名稱','小類別名稱'])['銷售總額'].sum().reset_index().sort_values('銷售總額',ascending=False),x='小類別名稱',y='銷售總額',color='大類別名稱',text_auto='.2s'), use_container_width=True); st.divider()

        st.subheader("門市偏好"); store_opts=["全部門市"]+list(df['門市'].unique()); target=st.selectbox("選擇門市",store_opts)
        sdf=df if target=="全部門市" else df[df['門市']==target]
        p1,p2,p3=st.columns(3)
        with p1: st.write("**最愛品項**"); st.dataframe(sdf.groupby('小類別名稱')['數量'].sum().sort_values(ascending=False).head(5))
        with p2: st.write("**最愛顏色**"); st.dataframe(sdf.groupby('顏色')['數量'].sum().sort_values(ascending=False).head(5))
        with p3: st.write("**最愛尺寸**"); st.dataframe(sdf[sdf['尺寸']!='0 (ONE SIZE)'].groupby('尺寸')['數量'].sum().sort_values(ascending=False).head(5))
        st.divider()

        st.subheader("VIP"); vip=fdf[fdf['會員編號'].notna()&(fdf['會員編號']!='')]
        if len(vip)>0:
            vs=vip.groupby(['會員名稱','會員編號']).agg(總消費額=('銷售總額','sum'),購買件數=('數量','sum')).reset_index().sort_values('總消費額',ascending=False).head(20)
            st.dataframe(vs.style.format({'總消費額':'${:,.0f}','購買件數':'{:,.0f}'}), use_container_width=True)
        st.divider()

        st.subheader("需求預測"); growth=st.slider("成長率 (%)",-20,50,10,1); mult=1+growth/100
        fc=fdf.groupby(['大類別名稱','小類別名稱']).agg(本年銷量=('數量','sum')).reset_index()
        fc['預測採購量']=(fc['本年銷量']*mult).astype(int); fc['增長']=fc['預測採購量']-fc['本年銷量']; fc=fc.sort_values('預測採購量',ascending=False)
        fig=go.Figure(); fig.add_trace(go.Bar(x=fc['小類別名稱'],y=fc['本年銷量'],name='本年',marker_color='lightgray'))
        fig.add_trace(go.Bar(x=fc['小類別名稱'],y=fc['預測採購量'],name='預測',marker_color='#2b4c7e')); fig.update_layout(barmode='group',title="本年 vs 預測")
        st.plotly_chart(fig, use_container_width=True); st.dataframe(fc, use_container_width=True, hide_index=True)
        st.download_button("📥 需求預測.xlsx", df_to_xlsx(fc,"預測","需求預測"), "需求預測.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if st.button("下一步 →", type="primary"):
            st.session_state["p_sales"]=df; st.session_state["p_fdf"]=fdf; st.session_state["p_forecast"]=fc; nxt(); st.rerun()

    # === Step 2：Assortment ===
    elif ps == 2:
        st.markdown('<span class="step-badge">STEP 2</span>', unsafe_allow_html=True)
        st.title("生成 Assortment 叫貨單")
        method2 = st.radio("選擇資料來源", ["上傳 Excel","使用範例資料"], key="s2m")
        col1,col2=st.columns(2)
        with col1: budget=st.number_input("採購預算 (NTD)",min_value=0,value=50000000,step=1000000)
        with col2: exchange_rate=st.number_input("歐元兌台幣",min_value=0.0,value=35.5,step=0.1)

        if method2=="上傳 Excel":
            nl=st.file_uploader("新季商品清單",type=["xlsx"],key="s2u")
            if nl:
                try: df_new=pd.read_excel(nl,header=1); df_new=df_new.iloc[1:].copy(); st.success(f"已讀取 {len(df_new)} 筆")
                except Exception as e: st.error(f"讀取失敗：{e}"); st.stop()
            else: st.info("等待上傳..."); st.stop()
        else:
            df_new=gen_assortment_sample(); st.success(f"已載入範例資料 {len(df_new)} 筆")

        if 'Wholesale Price(€)' in df_new.columns: df_new['Wholesale Price(NTD)']=(pd.to_numeric(df_new.get('Wholesale Price(€)',0),errors='coerce')*exchange_rate).round()
        if 'Suggested Retail Price(€)' in df_new.columns: df_new['Retail Price(NTD)']=(pd.to_numeric(df_new.get('Suggested Retail Price(€)',0),errors='coerce')*exchange_rate).round()

        ds=st.session_state.get("p_sales",pd.DataFrame())
        if len(ds)>0 and '實售價格' in ds.columns:
            tr=(ds['實售價格']*ds['數量']).sum(); tq=ds['數量'].sum(); aw=(tr/tq*0.4) if tq>0 else 10000; tt=int(budget/aw)
        else: aw=10000; tt=int(budget/aw)
        st.info(f"預估進貨單價：NT${aw:,.0f}　→　可採購約 {tt:,} 件")

        if 'PLC' in df_new.columns:
            ct=int(tt*0.6); se=tt-ct; df_new['建議採購量']=0
            co=df_new['PLC'].astype(str).str.strip().str.title()=='Carryover'; sm2=df_new['PLC'].astype(str).str.strip().str.title()=='Seasonal'
            if co.sum()>0: df_new.loc[co,'建議採購量']=ct//co.sum()
            if sm2.sum()>0: df_new.loc[sm2,'建議採購量']=se//sm2.sum()
        else: df_new['建議採購量']=tt//len(df_new) if len(df_new)>0 else 0

        dc=[c for c in ['Style Number','Style Name','Category','PLC','Retail Price(NTD)','建議採購量'] if c in df_new.columns]
        edited=st.data_editor(df_new[dc],use_container_width=True,hide_index=True,column_config={"建議採購量":st.column_config.NumberColumn(min_value=0,step=1)})
        st.metric("總叫貨",f"{edited['建議採購量'].sum():,} 件")
        df_new.loc[edited.index,'建議採購量']=edited['建議採購量'].values
        st.download_button("📥 叫貨清單.xlsx",df_to_xlsx(edited,"叫貨","叫貨清單"),"叫貨清單.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if st.button("✅ Approve →", type="primary"):
            st.session_state["p_order"]=edited; st.session_state["p_new_list"]=df_new; nxt(); st.rerun()

    # === Step 3：配貨 ===
    elif ps == 3:
        st.markdown('<span class="step-badge">STEP 3</span>', unsafe_allow_html=True)
        st.title("配貨給門市")
        order=st.session_state.get("p_order",pd.DataFrame())
        if len(order)==0: st.warning("請先完成叫貨單"); st.stop()
        sdf=st.session_state.get("p_sales",pd.DataFrame())
        sl=list(sdf['門市'].unique()) if '門市' in sdf.columns and len(sdf)>0 else ["Store1","Store2","Store3"]
        st.subheader("門市配貨權重"); wc=st.columns(min(len(sl),5)); wt={}
        for i,s in enumerate(sl):
            with wc[i%5]: wt[s]=st.number_input(s,0.1,3.0,1.0,0.1,key=f"w_{i}")
        tw=sum(wt.values()); alloc,inv=[],[]
        sc2='Style Number' if 'Style Number' in order.columns else order.columns[0]
        qc='建議採購量' if '建議採購量' in order.columns else '叫貨數量'
        for _,r in order.iterrows():
            tq=int(r.get(qc,0))
            for s in sl:
                q=max(0,round(tq*wt[s]/tw)); alloc.append({"門市":s,"款式":str(r.get(sc2,"")),"配貨":q})
                inv.append({"store":s,"style":str(r.get(sc2,"")),"quantity":q,"sold":0})
        df_a=pd.DataFrame(alloc)
        st.plotly_chart(px.bar(df_a.groupby("門市")["配貨"].sum().reset_index(),x="門市",y="配貨",color="門市",title="各門市配貨量"), use_container_width=True)
        st.download_button("📥 配貨表.xlsx",df_to_xlsx(df_a,"配貨","配貨表"),"配貨表.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if st.button("下一步 →", type="primary"):
            st.session_state["p_alloc"]=df_a; st.session_state["p_inv"]=pd.DataFrame(inv); nxt(); st.rerun()

    # === Step 4：銷售 ===
    elif ps == 4:
        st.markdown('<span class="step-badge">STEP 4</span>', unsafe_allow_html=True)
        st.title("門市實際銷售")
        inv=st.session_state["p_inv"].copy()
        method4=st.radio("銷售資料",["自動模擬","上傳 Excel"],key="s4m")
        if method4=="上傳 Excel":
            cf=st.file_uploader("上傳當期銷售",type=["xlsx"],key="s4u")
            if cf:
                dc=pd.read_excel(cf); st.success(f"已讀取 {len(dc)} 筆")
                if 'quantity' in dc.columns and 'store' in dc.columns and 'item number' in dc.columns:
                    cg=dc.groupby(['item number','store'])['quantity'].sum()
                    for idx in inv.index: inv.loc[idx,"sold"]=int(cg.get((inv.loc[idx,"style"],inv.loc[idx,"store"]),0))
            else: st.info("等待上傳..."); st.stop()
        else:
            for idx in inv.index: inv.loc[idx,"sold"]=min(int(inv.loc[idx,"quantity"]*random.uniform(0.15,0.95)),inv.loc[idx,"quantity"])
            st.success("已模擬！")
        inv["remaining"]=inv["quantity"]-inv["sold"]; inv["sell_through"]=(inv["sold"]/inv["quantity"].replace(0,1)*100).round(1)
        sm=inv.groupby("store").agg(配貨=("quantity","sum"),已售=("sold","sum"),剩餘=("remaining","sum")).reset_index()
        sm["銷售率%"]=(sm["已售"]/sm["配貨"].replace(0,1)*100).round(1); sm.columns=["門市","配貨","已售","剩餘","銷售率%"]
        st.dataframe(sm,use_container_width=True,hide_index=True)
        st.plotly_chart(px.bar(sm,x="門市",y=["配貨","已售","剩餘"],barmode="group",title="銷售狀況"), use_container_width=True)
        st.download_button("📥 銷售狀況.xlsx",df_to_xlsx(sm,"銷售","銷售狀況"),"銷售狀況.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if st.button("下一步 →", type="primary"): st.session_state["p_inv"]=inv; nxt(); st.rerun()

    # === Step 5：庫存 ===
    elif ps == 5:
        st.markdown('<span class="step-badge">STEP 5</span>', unsafe_allow_html=True)
        st.title("庫存調整與調撥")
        inv=st.session_state["p_inv"].copy(); inv["remaining"]=inv["quantity"]-inv["sold"]
        low=inv[inv["remaining"]<=2]
        if len(low[low["remaining"]==0])>0: st.error(f"🔴 斷貨 {len(low[low['remaining']==0])} 項")
        if len(low[low["remaining"]>0])>0: st.warning(f"🟡 低庫存 {len(low[low['remaining']>0])} 項")
        if len(low)==0: st.success("✅ 全部正常")
        SAFE=2; transfers=[]
        for style in inv["style"].unique():
            bs=inv[inv["style"]==style].groupby("store")["remaining"].sum()
            if len(bs)>=2:
                mx,mn=bs.idxmax(),bs.idxmin()
                if bs[mx]>5 and bs[mn]<SAFE:
                    t=min(SAFE-int(bs[mn]),int(bs[mx])-SAFE)
                    if t>0: transfers.append({"款式":style,"從":mx,"調出":t,"到":mn})
        if transfers:
            st.subheader("🔄 調撥建議"); df_t=pd.DataFrame(transfers); st.dataframe(df_t,use_container_width=True,hide_index=True)
        st.plotly_chart(px.bar(inv.groupby("store")["remaining"].sum().reset_index().rename(columns={"store":"門市","remaining":"剩餘"}),x="門市",y="剩餘",color="門市",title="剩餘庫存"), use_container_width=True)
        if st.button("下一步 →", type="primary"): st.session_state["p_inv"]=inv; nxt(); st.rerun()

    # === Step 6：滯銷品 ===
    elif ps == 6:
        st.markdown('<span class="step-badge">STEP 6</span>', unsafe_allow_html=True)
        st.title("滯銷品處置")
        inv=st.session_state["p_inv"].copy(); inv["remaining"]=inv["quantity"]-inv["sold"]
        inv["sell_through"]=(inv["sold"]/inv["quantity"].replace(0,1)*100).round(1)
        nl=st.session_state.get("p_new_list",pd.DataFrame()); cd=datetime.now(); results=[]
        for _,row in inv.iterrows():
            style=row["style"]; plc="Seasonal"
            if len(nl)>0 and 'Style Number' in nl.columns and 'PLC' in nl.columns:
                m=nl[nl['Style Number'].astype(str).str.strip()==str(style).strip()]
                if len(m)>0: plc=str(m.iloc[0].get('PLC','Seasonal')).strip()
            ic=plc.lower()=="carryover"
            action=rule_engine({"is_carryover":ic,"launch_date":"2026-04-01","brand_tier":"Luxury","current_location":"正價專門店"},cd)
            if row["sell_through"]>50: action="維持正價銷售"
            results.append({"門市":row["store"],"款式":style,"配貨":row["quantity"],"已售":row["sold"],"剩餘":row["remaining"],"銷售率%":row["sell_through"],"處置建議":action})
        df_d=pd.DataFrame(results); slow=df_d[df_d["處置建議"]!="維持正價銷售"]
        if len(slow)>0:
            c1,c2=st.columns(2); c1.metric("需處置",f"{len(slow)} 項"); c2.metric("處置庫存",f"{slow['剩餘'].sum()} 件")
            st.dataframe(slow,use_container_width=True,hide_index=True)
            st.plotly_chart(px.bar(slow.groupby("處置建議").agg(總件數=("剩餘","sum")).reset_index(),x="處置建議",y="總件數",color="處置建議",title="處置分佈"), use_container_width=True)
            st.download_button("📥 滯銷品.xlsx",df_to_xlsx(slow,"滯銷","滯銷品清單"),"滯銷品.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else: st.success("🎉 沒有滯銷品！")
        if st.button("下一步 →", type="primary"):
            st.session_state["p_slow"]=slow if len(slow)>0 else pd.DataFrame(); nxt(); st.rerun()

    # === Step 7：瑕疵品 ===
    elif ps == 7:
        st.markdown('<span class="step-badge">STEP 7</span>', unsafe_allow_html=True)
        st.title("瑕疵品處理")
        method7=st.radio("選擇資料來源",["上傳 Excel","使用範例資料"],key="s7m")
        if method7=="上傳 Excel":
            df7=st.file_uploader("上傳瑕疵品 Excel",type=["xlsx"],key="s7u")
            if df7:
                try: df_def=pd.read_excel(df7,header=1); st.success(f"已讀取 {len(df_def)} 筆")
                except: st.error("讀取失敗"); st.stop()
            else: st.info("等待上傳，或切換範例資料"); st.stop()
        else: df_def=gen_defect_sample(); st.success(f"已載入範例資料 {len(df_def)} 筆")
        try:
            results=df_def.apply(lambda row: classify_defect(row), axis=1)
            rdf=pd.DataFrame(results.tolist(),columns=["驗證","瑕疵等級","建議處理","原因"])
            dfc=pd.concat([df_def[['Style Number','Style Name','Color Code','Defect Description']].reset_index(drop=True),rdf.reset_index(drop=True)],axis=1)
            c1,c2,c3=st.columns(3); c1.metric("總筆數",f"{len(rdf)}"); c2.metric("仿冒品",f"{(rdf['驗證']=='No').sum()}"); c3.metric("Grade C",f"{(rdf['瑕疵等級']=='Grade C').sum()}")
            st.dataframe(dfc,use_container_width=True,hide_index=True)
            gc=rdf['瑕疵等級'].value_counts().reset_index(); gc.columns=["等級","數量"]
            st.plotly_chart(px.bar(gc, x="等級", y="數量", color="等級", title="瑕疵等級分佈"), use_container_width=True)
            st.download_button("📥 瑕疵品.xlsx",df_to_xlsx(dfc,"瑕疵","瑕疵品報告"),"瑕疵品.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.session_state["p_defect"]=dfc
        except Exception as e: st.error(f"處理失敗：{e}")
        if st.button("下一步 →", type="primary"): nxt(); st.rerun()

    # === Step 8：退換貨 ===
    elif ps == 8:
        st.markdown('<span class="step-badge">STEP 8</span>', unsafe_allow_html=True)
        st.title("退換貨處理")
        method8=st.radio("選擇資料來源",["上傳 Excel","使用範例資料"],key="s8m")
        if method8=="上傳 Excel":
            df8=st.file_uploader("上傳退換貨 Excel",type=["xlsx"],key="s8u")
            if df8:
                try: df_ret=pd.read_excel(df8,header=1); st.success(f"已讀取 {len(df_ret)} 筆")
                except: st.error("讀取失敗"); st.stop()
            else: st.info("等待上傳，或切換範例資料"); st.stop()
        else: df_ret=gen_return_sample(); st.success(f"已載入範例資料 {len(df_ret)} 筆")
        try:
            results=df_ret.apply(lambda row: evaluate_return(row), axis=1)
            rdf=pd.DataFrame(results.tolist(),columns=["處理結果","原因"])
            dfc=pd.concat([df_ret[['Style Number','Style Name','Product Condition','Retail Price(NTD)']].reset_index(drop=True),rdf.reset_index(drop=True)],axis=1)
            c1,c2,c3=st.columns(3); c1.metric("總筆數",f"{len(rdf)}"); c2.metric("核准",f"{(rdf['處理結果']=='核准退貨').sum()}"); c3.metric("拒絕",f"{(rdf['處理結果']=='拒絕').sum()}")
            st.dataframe(dfc,use_container_width=True,hide_index=True)
            rc=rdf['處理結果'].value_counts().reset_index(); rc.columns=["結果","數量"]
            st.plotly_chart(px.pie(rc, values="數量", names="結果", title="退換貨結果分佈"), use_container_width=True)
            st.download_button("📥 退換貨.xlsx",df_to_xlsx(dfc,"退換貨","退換貨報告"),"退換貨.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.session_state["p_return"]=dfc
        except Exception as e: st.error(f"處理失敗：{e}")
        if st.button("完成 → 最終報告 📊", type="primary"): nxt(); st.rerun()

    # === Step 9：最終報告 ===
    elif ps == 9:
        st.markdown('<span class="step-badge">FINAL</span>', unsafe_allow_html=True)
        st.title("最終報告"); st.balloons()
        season=st.session_state.get("p_season","")
        st.markdown(f'<div class="card"><h3>{user["brand"]}　{season}</h3><p>操作者：{user["username"]}　日期：{date.today()}</p></div>', unsafe_allow_html=True)

        order=st.session_state.get("p_order",pd.DataFrame()); alloc=st.session_state.get("p_alloc",pd.DataFrame())
        inv=st.session_state.get("p_inv",pd.DataFrame()); slow=st.session_state.get("p_slow",pd.DataFrame())
        defect=st.session_state.get("p_defect",pd.DataFrame()); ret=st.session_state.get("p_return",pd.DataFrame())

        qty_col='建議採購量' if len(order)>0 and '建議採購量' in order.columns else '叫貨數量'

        if len(order)>0:
            st.header("一、叫貨清單"); st.metric("總叫貨",f"{order[qty_col].sum():,} 件")
            st.dataframe(order,use_container_width=True,hide_index=True); st.divider()
        if len(alloc)>0:
            st.header("二、配貨"); st.plotly_chart(px.bar(alloc.groupby("門市")["配貨"].sum().reset_index(),x="門市",y="配貨",color="門市"), use_container_width=True); st.divider()
        if len(inv)>0:
            st.header("三、銷售與庫存"); ic=inv.copy(); ic["remaining"]=ic["quantity"]-ic["sold"]
            sm=ic.groupby("store").agg(配貨=("quantity","sum"),已售=("sold","sum"),剩餘=("remaining","sum")).reset_index()
            sm["銷售率%"]=(sm["已售"]/sm["配貨"].replace(0,1)*100).round(1); sm.columns=["門市","配貨","已售","剩餘","銷售率%"]
            st.dataframe(sm,use_container_width=True,hide_index=True); st.divider()
        if len(slow)>0: st.header("四、滯銷品"); st.dataframe(slow,use_container_width=True,hide_index=True); st.divider()
        if len(defect)>0: st.header("五、瑕疵品"); st.dataframe(defect,use_container_width=True,hide_index=True); st.divider()
        if len(ret)>0: st.header("六、退換貨"); st.dataframe(ret,use_container_width=True,hide_index=True); st.divider()

        st.header("📥 下載報表")
        c1,c2,c3=st.columns(3)
        with c1:
            if len(order)>0: st.download_button("📥 叫貨清單.xlsx",df_to_xlsx(order,"叫貨","叫貨清單"),"叫貨清單.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if len(alloc)>0: st.download_button("📥 配貨表.xlsx",df_to_xlsx(alloc,"配貨","配貨表"),"配貨表.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c2:
            if len(slow)>0: st.download_button("📥 滯銷品.xlsx",df_to_xlsx(slow,"滯銷","滯銷品清單"),"滯銷品.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if len(defect)>0: st.download_button("📥 瑕疵品.xlsx",df_to_xlsx(defect,"瑕疵","瑕疵品報告"),"瑕疵品.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c3:
            if len(ret)>0: st.download_button("📥 退換貨.xlsx",df_to_xlsx(ret,"退換貨","退換貨報告"),"退換貨.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.divider()
        summary=f"叫貨 {order[qty_col].sum():,} 件" if len(order)>0 and qty_col in order.columns else "已完成"

        if st.button("💾 儲存此次作業", type="primary", use_container_width=True):
            save_session_record(
                user["id"], season, summary,
                order_df=order, alloc_df=alloc, inv_df=inv,
                slow_df=slow, defect_df=defect, return_df=ret
            )
            st.success("已儲存！可在「歷史紀錄」查看完整內容")

        if st.button("🔄 開始新作業"):
            for k in [k for k in st.session_state.keys() if k.startswith("p_")]: del st.session_state[k]
            st.session_state["ps"]=0; st.rerun()
