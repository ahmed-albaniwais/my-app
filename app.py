import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter  # ✅ مهم للتلوين + تحويل أرقام الأعمدة لحروف

# =========================================================
# إعداد الصفحة
# =========================================================
st.set_page_config(layout="wide")
st.title("📈 تقرير المنتجات والباقات والمخزون (نهائي - بدون تقرير التوزيع)")

# =========================================================
# أدوات قراءة ذكية (حل مشاكل الترميز/الفواصل)
# =========================================================
def read_csv_smart(uploaded_file):
    encodings = ["utf-8-sig", "utf-8", "cp1256", "latin1"]
    last_err = None
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc)
        except Exception as e:
            last_err = e
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc, engine="python")
        except Exception as e:
            last_err = e
    raise last_err

def pick_first_existing(cols, candidates):
    for c in candidates:
        if c in cols:
            return c
    return None

def normalize_str_col(df, col):
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip()


# =========================================================
# ✅ Mapping مباشر لفئات الباقات
# =========================================================
packages_30 = [
    "RED 30 For 3 months",
    "RED 30,000 Monthly Package",
    "Yooz 30 Extra Data",
    "Yooz 30 Extra Data-3 Months",
    "Yooz 30 Extra Voice",
    "Yooz 30 Extra Voice-3 Months",
    "Yooz 30 Mix",
    "Yooz 30 Mix-3 Months",
    "Yooz 40 Extra Data",
    "Yooz 40 Extra Voice"
]

packages_25 = [
    "RED 25 For 3 months",
    "RED 25,000 Monthly Package",
    "Yooz 25 Extra Data",
    "Yooz 25 Extra Data-3 Months",
    "Yooz 25 Extra Voice",
    "Yooz 25 Mix"
]

packages_15 = [
    "RED 15 For 3 Months",
    "RED 15,000 Monthly Package",
    "Yooz 15 Extra Data",
    "Yooz 15 Extra Data-3 Months",
    "Yooz 15 Extra Voice",
    "Yooz 15 Extra Voice-3 Months",
    "Yooz 15 Mix"
]

packages_20 = [
    "RED 20 For 3 Months",
    "RED 20,000 Monthly Package",
    "Yooz 20 Extra Data",
    "Yooz 20 Extra Voice",
    "Yooz 20 Mix"
]

packages_10 = [
    "Data Line Package Bundle 2",
    "Gift Weekly 4G UL",
    "Monthly10GB",
    "Monthly20GB",
    "Monthly5GB",
    "RED 10,000 for 20 days",
    "Tesla New Data Bundle Monthly",
    "Unlimited 4G Bundle",
    "Unlimited Premium Package FUP",
    "Unlimited Weekly 4G",
    "UnlimitedMonthlyBTL",
    "Yooz 10 Extra Data",
    "Yooz 10 Extra Voice",
    "Yooz 10 Mix"
]

packages_5 = [
    "RED 5,000 Weekly Package",
    "Tesla New Data Bundle Daily",
    "Tesla New Data Bundle Weekly",
    "Unlimited Daily 4G",
    "Weekly3GB",
    "Yooz 5 Extra Data",
    "Yooz 5 Extra Voice",
    "Yooz 5 Mix",
    "Yooz FAF Addon"
]

package_map = {}

for p in packages_30:
    package_map[p] = "باقات 30"

for p in packages_25:
    package_map[p] = "باقات 25"

for p in packages_20:
    package_map[p] = "باقات 20"

for p in packages_15:
    package_map[p] = "باقات 15"

for p in packages_10:
    package_map[p] = "باقات 10"

for p in packages_5:
    package_map[p] = "باقات 5"

# =========================================================
# رفع الملفات
# =========================================================
activation_file = st.sidebar.file_uploader("📥 ملف التفعيل (CSV)", type=["csv"])
customer_file   = st.sidebar.file_uploader("📥 ملف الوكلاء (Excel)", type=["xlsx"])
bundle_file     = st.sidebar.file_uploader("📥 تقرير الباقات (CSV)", type=["csv"])
stock_file      = st.sidebar.file_uploader("📥 تقرير المخزون (CSV)", type=["csv"])

if not (activation_file and customer_file and bundle_file and stock_file):
    st.info("يرجى رفع جميع الملفات المطلوبة لإصدار التقارير.")
    st.stop()

# =========================================================
# قراءة الملفات
# =========================================================
df_activation = read_csv_smart(activation_file)
df_customers  = pd.read_excel(customer_file)
df_bundle     = read_csv_smart(bundle_file)
df_stock      = read_csv_smart(stock_file)

# =========================================================
# توحيد/تنظيف الأعمدة الأساسية
# =========================================================
for c in ["SUB_MSISDN", "DEALER_MSISDN"]:
    normalize_str_col(df_activation, c)

for c in ["SUB_MSISDN", "DEALER_MSISDN", "BUNDLE_NAME", "TYPE_OF_PROD"]:
    if c in df_bundle.columns:
        if c in ["SUB_MSISDN", "DEALER_MSISDN", "TYPE_OF_PROD"]:
            df_bundle[c] = df_bundle[c].astype(str).str.strip()

for c in ["POS_MSISDN", "ITEM_NAME"]:
    normalize_str_col(df_stock, c)

# =========================================================
# تجهيز النقاط من ملف التفعيل نفسه
# =========================================================
if "BUNDLE_POINT" not in df_activation.columns:
    st.error("⚠️ ملف التفعيل لا يحتوي على عمود BUNDLE_POINT.")
    st.stop()

df_activation["BUNDLE_POINT"] = pd.to_numeric(
    df_activation["BUNDLE_POINT"], errors="coerce"
).fillna(0)

# تجميع نقاط كل رقم من ملف التفعيل
points_per_sub = (
    df_activation
    .groupby("SUB_MSISDN", as_index=False)["BUNDLE_POINT"]
    .sum()
)

# =========================================================
# تجهيز تاريخ التفعيل في ملف التفعيل
# =========================================================
act_date_col = pick_first_existing(df_activation.columns, ["ACTIVATION_DATE", "ACTIVATION_DT", "ACT_DATE", "DATE"])
if not act_date_col:
    st.error("⚠️ لم يتم العثور على عمود تاريخ التفعيل في ملف التفعيل (مثل ACTIVATION_DATE/ACTIVATION_DT).")
    st.stop()

df_activation["ACTIVATION_DATE"] = pd.to_datetime(df_activation[act_date_col], errors="coerce").dt.date

# =========================================================
# تجهيز تاريخ التفعيل في ملف الباقات
# =========================================================
bundle_date_col = pick_first_existing(df_bundle.columns, ["ACTIVATION_DATE", "ACTIVATION_DT", "ACT_DATE", "DATE"])
if not bundle_date_col:
    st.error("⚠️ لم يتم العثور على عمود تاريخ التفعيل في تقرير الباقات (مثل ACTIVATION_DATE/ACTIVATION_DT).")
    st.stop()

df_bundle["ACTIVATION_DATE"] = pd.to_datetime(df_bundle[bundle_date_col], errors="coerce").dt.date

# =========================================================
# تجهيز رقم الوكيل normalized (للمخزون)
# =========================================================
if "DEALER_MSISDN" not in df_activation.columns:
    st.error("⚠️ ملف التفعيل لا يحتوي على DEALER_MSISDN.")
    st.stop()

df_activation["DEALER_MSISDN"] = df_activation["DEALER_MSISDN"].astype(str).str.strip()
df_activation["DEALER_MSISDN_normalized"] = df_activation["DEALER_MSISDN"].str[-10:]

# =========================================================
# تجهيز ملف الوكلاء + إعادة تسمية الأعمدة
# =========================================================
if "DEALER_MSISDN" not in df_customers.columns:
    alt = pick_first_existing(df_customers.columns, ["DEALER_MSISDN", "DEALER", "MSISDN", "رقم الوكيل", "رقم"])
    if not alt:
        st.error("⚠️ ملف الوكلاء لا يحتوي على عمود DEALER_MSISDN (أو بديل له).")
        st.stop()
    df_customers = df_customers.rename(columns={alt: "DEALER_MSISDN"})

df_customers["DEALER_MSISDN"] = df_customers["DEALER_MSISDN"].astype(str).str.strip()

df_customers = df_customers.rename(columns={
    "الكود": "CODE",
    "اسم المحل": "SHOP_NAME",
    "اسم المندوب": "REP_NAME",
    "USERLASTNAME": "OWNER_LAST",
    "USERFIRSTNAME": "OWNER_FIRST",
})

if "CODE" not in df_customers.columns:
    alt_code = pick_first_existing(df_customers.columns, ["كود", "كود الوكيل", "DEALER_CODE", "AGENT_CODE", "Code", "code"])
    if alt_code:
        df_customers = df_customers.rename(columns={alt_code: "CODE"})

# =========================================================
# تحديد الفترة
# =========================================================
st.sidebar.subheader("📅 تحديد الفترة")
min_date = min(df_activation["ACTIVATION_DATE"].min(), df_bundle["ACTIVATION_DATE"].min())
max_date = max(df_activation["ACTIVATION_DATE"].max(), df_bundle["ACTIVATION_DATE"].max())

start_date = st.sidebar.date_input("من تاريخ", min_value=min_date, max_value=max_date, value=min_date)
end_date   = st.sidebar.date_input("إلى تاريخ", min_value=min_date, max_value=max_date, value=max_date)

df_activation_f = df_activation[
    (df_activation["ACTIVATION_DATE"] >= start_date) &
    (df_activation["ACTIVATION_DATE"] <= end_date)
].copy()

df_bundle_f = df_bundle[
    (df_bundle["ACTIVATION_DATE"] >= start_date) &
    (df_bundle["ACTIVATION_DATE"] <= end_date)
].copy()

# =========================================================
# دمج التفعيل مع الوكلاء
# =========================================================
df = df_activation_f.merge(df_customers, on="DEALER_MSISDN", how="left")

# =========================================================
# تحديد نوع المنتج من تقرير الباقات (TYPE_OF_PROD)
# =========================================================
if "SUB_MSISDN" not in df_bundle.columns:
    st.error("⚠️ تقرير الباقات لا يحتوي على SUB_MSISDN.")
    st.stop()
if "TYPE_OF_PROD" not in df_bundle.columns:
    st.error("⚠️ تقرير الباقات لا يحتوي على TYPE_OF_PROD.")
    st.stop()

def map_product_type(type_of_prod: str) -> str:
    t = str(type_of_prod).strip().upper()

    if t.startswith("USIM-PRP"):
        return "خط ريد"

    if t == "USIM-YOUT":
        return "خط يوز"

    mapping = {
        "UBSIM-3IN1": "خط مميز",
        "UBSIM-YZS1MC": "يوز مليوني",
        "USIM-40KV": "خط نت 40",
        "USIM-50KV": "نت 50",
        "USIM-GPRS-MIFI4": "ماي فاي",
        "USIM-GPRS-MIFI5": "ماي فاي جديد",
        "USIM-GPRS-PKG": "نت عادي",
        "USIM-GPRS-RTR1": "راوتر",
        "USIM-GPRS-RTR2": "راوتر",
    }
    return mapping.get(t, "غير معروف")

df_bundle_all = df_bundle.copy()
df_bundle_all["SUB_MSISDN"] = df_bundle_all["SUB_MSISDN"].astype(str).str.strip()
df_bundle_all["TYPE_OF_PROD"] = df_bundle_all["TYPE_OF_PROD"].astype(str).str.strip()
df_bundle_all["PRODUCT_TYPE_FROM_BUNDLE"] = df_bundle_all["TYPE_OF_PROD"].apply(map_product_type)

bundle_type_per_sub = (
    df_bundle_all[df_bundle_all["PRODUCT_TYPE_FROM_BUNDLE"] != "غير معروف"]
    .groupby("SUB_MSISDN")["PRODUCT_TYPE_FROM_BUNDLE"]
    .agg(lambda s: s.value_counts().index[0])
    .reset_index()
)

if "SUB_MSISDN" not in df.columns:
    st.error("⚠️ ملف التفعيل لا يحتوي على SUB_MSISDN.")
    st.stop()

df["SUB_MSISDN"] = df["SUB_MSISDN"].astype(str).str.strip()
df = df.merge(
    bundle_type_per_sub.rename(columns={"PRODUCT_TYPE_FROM_BUNDLE": "PRODUCT_TYPE"}),
    on="SUB_MSISDN",
    how="left"
)
df["PRODUCT_TYPE"] = df["PRODUCT_TYPE"].fillna("غير معروف")

# =========================================================
# تصنيف الشحن (إن وجد)
# =========================================================
def classify_amount(value):
    try:
        value = float(value)
    except Exception:
        return "غير مشحونة"
    if pd.isna(value) or value == 0:
        return "غير مشحونة"
    elif value in [5000, 7000, 8000, 9000]:
        return "فئة 5،7،8،9"
    elif value == 6000:
        return "فئة 6"
    elif value >= 10000:
        return "فئة 10 فما فوق"
    else:
        return "أخرى"

if "RECHARGE_AMOUNT" in df.columns:
    df["RECHARGE_CATEGORY"] = df["RECHARGE_AMOUNT"].apply(classify_amount)
else:
    df["RECHARGE_AMOUNT"] = None
    df["RECHARGE_CATEGORY"] = "غير مشحونة"

# =========================================================
# تقرير المخزون
# =========================================================
if "POS_MSISDN" not in df_stock.columns:
    st.error("⚠️ ملف المخزون لا يحتوي على POS_MSISDN.")
    st.stop()

df_stock["POS_MSISDN"] = df_stock["POS_MSISDN"].astype(str).str.strip()
df_stock["ITEM_NAME"] = df_stock["ITEM_NAME"].astype(str).str.strip().str.lower() if "ITEM_NAME" in df_stock.columns else ""

def map_product_from_item(item):
    item = str(item).lower().strip()
    if "pre-paid sim card free" in item or "pre-paid sim card - type2" in item or item == "pre-paid sim card":
        return "خط ريد"
    elif "type2" in item or "red" in item:
        return "خط ريد"
    elif "tozed cpe cat6" in item:
        return "راوتر جديد"
    elif "mifi" in item:
        return "ماي فاي"
    elif "with package" in item:
        return "نت عادي"
    elif "million coin" in item:
        return "يوز مليوني"
    elif "usim" in item and "40000" in item:
        return "خط نت 40"
    elif "youth" in item:
        return "خط يوز"
    elif "visitor sim card with 3gb" in item:
        return "خط زيارة 3 كيكا"
    elif "visitor sim card" in item:
        return "خط زيارة"
    else:
        return "غير معروف"

df_stock["PRODUCT_TYPE"] = df_stock["ITEM_NAME"].apply(map_product_from_item)

if "BALANCE" in df_stock.columns:
    df_stock["BALANCE"] = pd.to_numeric(df_stock["BALANCE"], errors="coerce").fillna(0)
else:
    df_stock["BALANCE"] = 0

df_stock_summary = df_stock.groupby(["POS_MSISDN", "PRODUCT_TYPE"])["BALANCE"].sum().reset_index()
df_stock_summary = df_stock_summary.rename(columns={"POS_MSISDN": "DEALER_MSISDN_normalized", "BALANCE": "STOCK_BALANCE"})
df_stock_summary["DEALER_MSISDN_normalized"] = df_stock_summary["DEALER_MSISDN_normalized"].astype(str).str[-10:]
df["DEALER_MSISDN_normalized"] = df["DEALER_MSISDN_normalized"].astype(str).str[-10:]

# =========================================================
# تقرير المنتجات (ضمن الفترة)
# =========================================================
group_cols = [
    "PRODUCT_TYPE", "SHOP_NAME", "OWNER_FIRST", "OWNER_LAST",
    "DEALER_MSISDN", "REP_NAME", "DEALER_MSISDN_normalized"
]
if "CODE" in df.columns:
    group_cols.insert(4, "CODE")

report = df.groupby(group_cols).agg(
    total_activations=("SUB_MSISDN", "count"),
    total_recharged=("RECHARGE_AMOUNT", lambda x: x.notna().sum()),
    count_5798=("RECHARGE_CATEGORY", lambda x: (x == "فئة 5،7،8،9").sum()),
    count_6=("RECHARGE_CATEGORY", lambda x: (x == "فئة 6").sum()),
    count_10plus=("RECHARGE_CATEGORY", lambda x: (x == "فئة 10 فما فوق").sum()),
    quality_lines=("SUB_MSISDN", lambda x: (df.loc[x.index, "QUALITY"].astype(str).str.upper() == "Y").sum()) if "QUALITY" in df.columns else ("SUB_MSISDN", lambda x: 0)
).reset_index()

report = report.merge(df_stock_summary, on=["DEALER_MSISDN_normalized", "PRODUCT_TYPE"], how="left")
report["STOCK_BALANCE"] = report["STOCK_BALANCE"].fillna(0).astype(int)

# =========================================================
# مخالفات يوز (ضمن الفترة) + اليومي (ضمن الفترة)
# =========================================================
EXCLUDED_BUNDLES = {"Daily300MB", "PUBG"}
EXCEPTION_BUNDLES = {"Daily300MB", "PUBG"}

def is_yooz_bundle(bundle_name: str) -> bool:
    b = str(bundle_name).lower()
    return ("yooz" in b) or ("youth" in b)

sub_to_type = bundle_type_per_sub.rename(columns={"PRODUCT_TYPE_FROM_BUNDLE": "PRODUCT_TYPE"}).copy()

df_yooz_check = df_bundle_f.copy()
df_yooz_check["SUB_MSISDN"] = df_yooz_check["SUB_MSISDN"].astype(str).str.strip()

if "BUNDLE_NAME" in df_yooz_check.columns:
    df_yooz_check["BUNDLE_NAME_CLEAN"] = df_yooz_check["BUNDLE_NAME"].astype("string").str.strip()
else:
    df_yooz_check["BUNDLE_NAME_CLEAN"] = pd.NA

df_yooz_check = df_yooz_check.merge(sub_to_type[["SUB_MSISDN", "PRODUCT_TYPE"]], on="SUB_MSISDN", how="left")
df_yooz_check["IS_YOOZ_BUNDLE"] = df_yooz_check["BUNDLE_NAME_CLEAN"].fillna("").apply(is_yooz_bundle)
df_yooz_check["VIOLATION"] = (
    (df_yooz_check["PRODUCT_TYPE"] == "خط يوز") &
    (~df_yooz_check["IS_YOOZ_BUNDLE"]) &
    (df_yooz_check["BUNDLE_NAME_CLEAN"].notna())
)

yooz_violations_report = df_yooz_check[df_yooz_check["VIOLATION"]].copy()
desired_cols = ["DEALER_MSISDN", "SUB_MSISDN", "PRODUCT_TYPE", "BUNDLE_NAME", "ACTIVATION_DATE"]
existing_cols = [c for c in desired_cols if c in yooz_violations_report.columns]
yooz_violations_report = yooz_violations_report[existing_cols].drop_duplicates()

# يومي ريد/يوز (ضمن الفترة) باستثناء Daily300MB و PUBG
df_daily = df_bundle_f.copy()
df_daily["SUB_MSISDN"] = df_daily["SUB_MSISDN"].astype(str).str.strip()
df_daily = df_daily.merge(sub_to_type[["SUB_MSISDN", "PRODUCT_TYPE"]], on="SUB_MSISDN", how="left")
df_daily["BUNDLE_NAME_CLEAN"] = df_daily["BUNDLE_NAME"].astype("string").str.strip()

df_daily = df_daily[
    (df_daily["PRODUCT_TYPE"].isin(["خط ريد", "خط يوز"])) &
    (df_daily["BUNDLE_NAME_CLEAN"].notna()) &
    (~df_daily["BUNDLE_NAME_CLEAN"].isin(EXCLUDED_BUNDLES))
].copy()

daily_red_yooz_bundle_summary = df_daily.groupby(["ACTIVATION_DATE", "PRODUCT_TYPE"]).agg(
    total_bundles=("BUNDLE_NAME_CLEAN", "count"),
    unique_lines=("SUB_MSISDN", "nunique")
).reset_index()

daily_red_yooz_bundle_pivot = daily_red_yooz_bundle_summary.pivot_table(
    index="ACTIVATION_DATE",
    columns="PRODUCT_TYPE",
    values="total_bundles",
    fill_value=0
).reset_index()

# =========================================================
# ✅ ملخص ريد+يوز — كل التفعيلات + يظهر حتى لو الوكيل غير موجود
# =========================================================
base_lines = df_activation.copy()
base_lines["DEALER_MSISDN"] = base_lines["DEALER_MSISDN"].astype(str).str.strip()
base_lines["SUB_MSISDN"] = base_lines["SUB_MSISDN"].astype(str).str.strip()

base_lines = base_lines.merge(df_customers, on="DEALER_MSISDN", how="left")

base_lines = base_lines.merge(
    bundle_type_per_sub.rename(columns={"PRODUCT_TYPE_FROM_BUNDLE": "PRODUCT_TYPE"}),
    on="SUB_MSISDN",
    how="left"
)
base_lines["PRODUCT_TYPE"] = base_lines["PRODUCT_TYPE"].fillna("غير معروف")
base_lines = base_lines[base_lines["PRODUCT_TYPE"].isin(["خط ريد", "خط يوز"])].copy()

base_lines["SHOP_NAME"] = base_lines["SHOP_NAME"].fillna("وكيل غير موجود بملف الوكلاء")
base_lines["REP_NAME"]  = base_lines["REP_NAME"].fillna("غير معروف")
if "CODE" in base_lines.columns:
    base_lines["CODE"] = base_lines["CODE"].fillna("—")

base_lines["SHOP_NAME_SHOW"] = base_lines["SHOP_NAME"].astype(str) + " | " + base_lines["DEALER_MSISDN"].astype(str)
base_lines["REP_NAME_SHOW"]  = base_lines["REP_NAME"].astype(str)

keep_cols = ["DEALER_MSISDN", "SHOP_NAME_SHOW", "REP_NAME_SHOW", "SUB_MSISDN", "PRODUCT_TYPE"]
if "CODE" in base_lines.columns:
    keep_cols.insert(1, "CODE")
if "QUALITY" in base_lines.columns:
    keep_cols.append("QUALITY")

base_lines = base_lines[keep_cols].drop_duplicates(subset=["DEALER_MSISDN", "SUB_MSISDN"])

df_flags_src = df_bundle.copy()
df_flags_src["SUB_MSISDN"] = df_flags_src["SUB_MSISDN"].astype(str).str.strip()
df_flags_src["BUNDLE_NAME_CLEAN"] = df_flags_src["BUNDLE_NAME"].astype("string").str.strip()

def has_any_bundle(s: pd.Series) -> bool:
    s2 = s.dropna().astype(str).str.strip()
    return (s2 != "").any()

bundle_flags = df_flags_src.groupby("SUB_MSISDN").agg(
    has_any_bundle=("BUNDLE_NAME_CLEAN", has_any_bundle),
    has_yooz_any=("BUNDLE_NAME_CLEAN", lambda s: s.dropna().astype(str).str.strip().pipe(
        lambda x: (
            x.ne("")
            & x.apply(is_yooz_bundle).astype(bool)
        ).any()
    )),
    has_non_yooz_violation=("BUNDLE_NAME_CLEAN", lambda s: s.dropna().astype(str).str.strip().pipe(
        lambda x: (
            x.ne("")
            & ~x.isin(EXCEPTION_BUNDLES)
            & ~x.apply(is_yooz_bundle).astype(bool)
        ).any()
    )),
    has_daily300=("BUNDLE_NAME_CLEAN", lambda s: (s.dropna().astype(str).str.strip() == "Daily300MB").any()),
    has_pubg=("BUNDLE_NAME_CLEAN", lambda s: (s.dropna().astype(str).str.strip() == "PUBG").any()),
).reset_index()

base_lines = base_lines.merge(bundle_flags, on="SUB_MSISDN", how="left")
fill_cols = ["has_any_bundle", "has_yooz_any", "has_non_yooz_violation", "has_daily300", "has_pubg"]
base_lines[fill_cols] = base_lines[fill_cols].fillna(False)

base_lines["HAS_VALID_BUNDLE"] = base_lines["has_any_bundle"]

if "QUALITY" in base_lines.columns:
    base_lines["IS_QUALITY"] = (base_lines["QUALITY"].astype(str).str.upper() == "Y")
else:
    base_lines["IS_QUALITY"] = False

group_cols_sum = ["DEALER_MSISDN", "SHOP_NAME_SHOW", "REP_NAME_SHOW"]
if "CODE" in base_lines.columns:
    group_cols_sum.insert(1, "CODE")

red_yooz_combined_summary = base_lines.groupby(group_cols_sum).agg(
    total_activated_lines=("SUB_MSISDN", "nunique"),
    lines_with_valid_bundle=("SUB_MSISDN", lambda s: s[base_lines.loc[s.index, "HAS_VALID_BUNDLE"]].nunique()),
    lines_without_bundle=("SUB_MSISDN", lambda s: s[~base_lines.loc[s.index, "has_any_bundle"]].nunique()),
    quality_lines=("SUB_MSISDN", lambda s: s[base_lines.loc[s.index, "IS_QUALITY"]].nunique()),
    daily300mb_lines=("SUB_MSISDN", lambda s: s[base_lines.loc[s.index, "has_daily300"]].nunique()),
    pubg_lines=("SUB_MSISDN", lambda s: s[base_lines.loc[s.index, "has_pubg"]].nunique()),
    yooz_lines_with_non_yooz_bundle=("SUB_MSISDN", lambda s: s[
        (base_lines.loc[s.index, "PRODUCT_TYPE"] == "خط يوز") &
        (base_lines.loc[s.index, "has_non_yooz_violation"])
    ].nunique()),
).reset_index()

ordered_cols = [
    "DEALER_MSISDN", "CODE", "SHOP_NAME_SHOW", "REP_NAME_SHOW",
    "total_activated_lines",
    "lines_with_valid_bundle",
    "lines_without_bundle",
    "yooz_lines_with_non_yooz_bundle",
    "quality_lines",
    "daily300mb_lines", "pubg_lines"
]
ordered_cols = [c for c in ordered_cols if c in red_yooz_combined_summary.columns]
red_yooz_combined_summary = red_yooz_combined_summary[ordered_cols]

# =========================================================
# ✅ ملخص جميع المنتجات - أفقي (كل منتج بسطر مستقل)
# =========================================================
base_lines_all = df_activation.copy()
base_lines_all["DEALER_MSISDN"] = base_lines_all["DEALER_MSISDN"].astype(str).str.strip()
base_lines_all["SUB_MSISDN"] = base_lines_all["SUB_MSISDN"].astype(str).str.strip()

base_lines_all = base_lines_all.merge(df_customers, on="DEALER_MSISDN", how="left")

base_lines_all = base_lines_all.merge(
    bundle_type_per_sub.rename(columns={"PRODUCT_TYPE_FROM_BUNDLE": "PRODUCT_TYPE"}),
    on="SUB_MSISDN",
    how="left"
)
base_lines_all["PRODUCT_TYPE"] = base_lines_all["PRODUCT_TYPE"].fillna("غير معروف")

base_lines_all["SHOP_NAME"] = base_lines_all["SHOP_NAME"].fillna("وكيل غير موجود بملف الوكلاء")
base_lines_all["REP_NAME"]  = base_lines_all["REP_NAME"].fillna("غير معروف")
if "CODE" in base_lines_all.columns:
    base_lines_all["CODE"] = base_lines_all["CODE"].fillna("—")

base_lines_all["SHOP_NAME_SHOW"] = base_lines_all["SHOP_NAME"].astype(str) + " | " + base_lines_all["DEALER_MSISDN"].astype(str)
base_lines_all["REP_NAME_SHOW"]  = base_lines_all["REP_NAME"].astype(str)
base_lines_all["DEALER_MSISDN_normalized"] = base_lines_all["DEALER_MSISDN"].astype(str).str[-10:]

keep_cols_all = [
    "DEALER_MSISDN",
    "DEALER_MSISDN_normalized",
    "SHOP_NAME_SHOW",
    "REP_NAME_SHOW",
    "SUB_MSISDN",
    "PRODUCT_TYPE"
]
if "CODE" in base_lines_all.columns:
    keep_cols_all.insert(1, "CODE")
if "QUALITY" in base_lines_all.columns:
    keep_cols_all.append("QUALITY")

base_lines_all = base_lines_all[keep_cols_all].drop_duplicates(subset=["DEALER_MSISDN", "SUB_MSISDN", "PRODUCT_TYPE"])

base_lines_all = base_lines_all.merge(bundle_flags, on="SUB_MSISDN", how="left")
fill_cols_all = ["has_any_bundle", "has_yooz_any", "has_non_yooz_violation", "has_daily300", "has_pubg"]
base_lines_all[fill_cols_all] = base_lines_all[fill_cols_all].fillna(False)

base_lines_all["HAS_VALID_BUNDLE"] = base_lines_all["has_any_bundle"]

if "QUALITY" in base_lines_all.columns:
    base_lines_all["IS_QUALITY"] = (base_lines_all["QUALITY"].astype(str).str.upper() == "Y")
else:
    base_lines_all["IS_QUALITY"] = False

stock_by_product = df_stock_summary.copy()
stock_by_product["DEALER_MSISDN_normalized"] = stock_by_product["DEALER_MSISDN_normalized"].astype(str).str[-10:]
stock_by_product["PRODUCT_TYPE"] = stock_by_product["PRODUCT_TYPE"].astype(str).str.strip()

group_cols_product = ["PRODUCT_TYPE", "DEALER_MSISDN", "SHOP_NAME_SHOW", "REP_NAME_SHOW", "DEALER_MSISDN_normalized"]
if "CODE" in base_lines_all.columns:
    group_cols_product.insert(2, "CODE")

all_products_horizontal_summary = base_lines_all.groupby(group_cols_product).agg(
    total_activated_lines=("SUB_MSISDN", "nunique"),
    total_lines_with_valid_bundle=("SUB_MSISDN", lambda s: s[base_lines_all.loc[s.index, "HAS_VALID_BUNDLE"]].nunique()),
    total_lines_without_bundle=("SUB_MSISDN", lambda s: s[~base_lines_all.loc[s.index, "has_any_bundle"]].nunique()),
    yooz_lines_with_non_yooz_bundle=("SUB_MSISDN", lambda s: s[
        (base_lines_all.loc[s.index, "PRODUCT_TYPE"] == "خط يوز") &
        (base_lines_all.loc[s.index, "has_non_yooz_violation"])
    ].nunique()),
    total_quality_lines=("SUB_MSISDN", lambda s: s[base_lines_all.loc[s.index, "IS_QUALITY"]].nunique()),
).reset_index()

all_products_horizontal_summary = all_products_horizontal_summary.merge(
    stock_by_product,
    on=["DEALER_MSISDN_normalized", "PRODUCT_TYPE"],
    how="left"
)

all_products_horizontal_summary["STOCK_BALANCE"] = (
    all_products_horizontal_summary["STOCK_BALANCE"].fillna(0).astype(int)
)

all_products_horizontal_summary = all_products_horizontal_summary.rename(columns={
    "PRODUCT_TYPE": "المنتج",
    "STOCK_BALANCE": "total_stock_balance"
})


# =========================================================
# ✅ إضافة فئات الباقات إلى ملخص جميع المنتجات
# ✅ اعتماداً على 777 نفسه وبدون تغيير منطق مع/بدون باقة
# =========================================================
bundle_pkg_src = df_bundle.copy()
bundle_pkg_src["SUB_MSISDN"] = bundle_pkg_src["SUB_MSISDN"].astype(str).str.strip()
bundle_pkg_src["BUNDLE_NAME_CLEAN"] = bundle_pkg_src["BUNDLE_NAME"].astype(str).str.strip()
bundle_pkg_src["PACKAGE_GROUP"] = bundle_pkg_src["BUNDLE_NAME_CLEAN"].map(package_map)

bundle_pkg_src = bundle_pkg_src[
    bundle_pkg_src["PACKAGE_GROUP"].notna()
].copy()

# نربط الباقات فقط مع نفس خطوط التفعيل الموجودة في base_lines_all
pkg_lines = base_lines_all[
    ["DEALER_MSISDN", "SUB_MSISDN", "PRODUCT_TYPE"] + (["CODE"] if "CODE" in base_lines_all.columns else [])
].merge(
    bundle_pkg_src[["SUB_MSISDN", "PACKAGE_GROUP"]],
    on="SUB_MSISDN",
    how="inner"
)

# نفس الخط يُحسب مرة واحدة فقط داخل نفس الفئة
pkg_lines = pkg_lines.drop_duplicates(
    subset=["DEALER_MSISDN", "PRODUCT_TYPE", "SUB_MSISDN", "PACKAGE_GROUP"] + (["CODE"] if "CODE" in pkg_lines.columns else [])
)

pkg_group_cols = ["DEALER_MSISDN", "PRODUCT_TYPE"]
if "CODE" in pkg_lines.columns:
    pkg_group_cols.append("CODE")

package_group_summary = (
    pkg_lines.groupby(pkg_group_cols + ["PACKAGE_GROUP"])["SUB_MSISDN"]
    .nunique()
    .unstack(fill_value=0)
    .reset_index()
)

package_group_summary.columns.name = None
package_group_summary = package_group_summary.rename(columns={"PRODUCT_TYPE": "المنتج"})

for c in ["باقات 5", "باقات 10", "باقات 15", "باقات 20", "باقات 25", "باقات 30"]:
    if c not in package_group_summary.columns:
        package_group_summary[c] = 0

merge_keys = ["DEALER_MSISDN", "المنتج"]
if "CODE" in all_products_horizontal_summary.columns and "CODE" in package_group_summary.columns:
    merge_keys.append("CODE")

# إصلاح نوع أعمدة الدمج قبل merge حتى لا يظهر خطأ int64/object
for col in merge_keys:
    if col in all_products_horizontal_summary.columns:
        all_products_horizontal_summary[col] = (
            all_products_horizontal_summary[col]
            .astype(str)
            .str.strip()
            .str.replace(".0", "", regex=False)
        )

    if col in package_group_summary.columns:
        package_group_summary[col] = (
            package_group_summary[col]
            .astype(str)
            .str.strip()
            .str.replace(".0", "", regex=False)
        )

all_products_horizontal_summary = all_products_horizontal_summary.merge(
    package_group_summary,
    on=merge_keys,
    how="left"
)

for c in ["باقات 5", "باقات 10", "باقات 15", "باقات 20", "باقات 25", "باقات 30"]:
    if c in all_products_horizontal_summary.columns:
        all_products_horizontal_summary[c] = all_products_horizontal_summary[c].fillna(0).astype(int)

horizontal_cols = [
    "المنتج",
    "DEALER_MSISDN",
    "CODE",
    "SHOP_NAME_SHOW",
    "REP_NAME_SHOW",
    "total_activated_lines",
    "total_lines_with_valid_bundle",
    "total_lines_without_bundle",
    "yooz_lines_with_non_yooz_bundle",
    "total_quality_lines",
    "باقات 5",
    "باقات 10",
    "باقات 15",
    "باقات 20",
    "باقات 25",
    "باقات 30",
    "total_stock_balance",
]
horizontal_cols = [c for c in horizontal_cols if c in all_products_horizontal_summary.columns]

all_products_horizontal_summary = all_products_horizontal_summary[horizontal_cols]

all_products_horizontal_summary = all_products_horizontal_summary.sort_values(
    by=["DEALER_MSISDN", "المنتج"]
).reset_index(drop=True)

# =========================================================
# ✅ تقرير النقاط — نفس ملخص جميع المنتجات مع نقاط تقرير الجودة
# =========================================================
points_lines = base_lines_all.merge(
    points_per_sub,
    on="SUB_MSISDN",
    how="left"
)

points_lines["BUNDLE_POINT"] = points_lines["BUNDLE_POINT"].fillna(0)

points_group_cols = ["PRODUCT_TYPE", "DEALER_MSISDN"]
if "CODE" in points_lines.columns:
    points_group_cols.append("CODE")

points_summary = (
    points_lines
    .groupby(points_group_cols, as_index=False)["BUNDLE_POINT"]
    .sum()
    .rename(columns={"PRODUCT_TYPE": "المنتج"})
)

points_merge_keys = ["المنتج", "DEALER_MSISDN"]
if "CODE" in all_products_horizontal_summary.columns and "CODE" in points_summary.columns:
    points_merge_keys.append("CODE")

for col in points_merge_keys:
    all_products_horizontal_summary[col] = (
        all_products_horizontal_summary[col]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )
    points_summary[col] = (
        points_summary[col]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )

points_report = all_products_horizontal_summary.merge(
    points_summary,
    on=points_merge_keys,
    how="left"
)

points_report["BUNDLE_POINT"] = points_report["BUNDLE_POINT"].fillna(0)
if (points_report["BUNDLE_POINT"] % 1 == 0).all():
    points_report["BUNDLE_POINT"] = points_report["BUNDLE_POINT"].astype(int)

points_report = points_report.sort_values(
    by=["DEALER_MSISDN", "المنتج"]
).reset_index(drop=True)

# أسماء عربية واضحة داخل ورقة تقرير النقاط
points_report = points_report.rename(columns={
    "total_activated_lines": "التفعيل الكلي",
    "total_lines_with_valid_bundle": "تفعيل بالباقات",
    "total_lines_without_bundle": "تفعيل بدون باقة",
    "yooz_lines_with_non_yooz_bundle": "خط يوز باقة ريد",
    "BUNDLE_POINT": "النقاط"
})

# =========================================================
# تقرير الباقات (كامل) — كل الباقات داخل وخارج الفترة
# =========================================================
if "DEALER_MSISDN" not in df_bundle.columns:
    st.error("⚠️ تقرير الباقات لا يحتوي على DEALER_MSISDN.")
    st.stop()

ALLOWED_TYPE_OF_PROD = {"USIM-PRP", "USIM-YOUT", "USIM-PRP2", "USIM-PRP-FREE"}

if "TYPE_OF_PROD" not in df_bundle.columns:
    st.error("⚠️ تقرير الباقات لا يحتوي على TYPE_OF_PROD.")
    st.stop()

df_bundle["TYPE_OF_PROD_CLEAN"] = df_bundle["TYPE_OF_PROD"].astype(str).str.strip().str.upper()

df_bundle_full_for_report = df_bundle[df_bundle["TYPE_OF_PROD_CLEAN"].isin(ALLOWED_TYPE_OF_PROD)].copy()

df_bundle_full_for_report["DEALER_MSISDN"] = df_bundle_full_for_report["DEALER_MSISDN"].astype(str).str.strip()
df_bundle_full_for_report["BUNDLE_NAME_CLEAN"] = df_bundle_full_for_report["BUNDLE_NAME"].astype("string").str.strip()

all_bundles = sorted(df_bundle_full_for_report["BUNDLE_NAME_CLEAN"].dropna().unique().tolist())

bundle_pivot = df_bundle_full_for_report.pivot_table(
    index="DEALER_MSISDN",
    columns="BUNDLE_NAME_CLEAN",
    values="SUB_MSISDN",
    aggfunc="count",
    fill_value=0
)
bundle_pivot = bundle_pivot.reindex(columns=all_bundles, fill_value=0).reset_index()

# دمج بيانات الوكلاء
base_cols = ["DEALER_MSISDN", "OWNER_FIRST", "OWNER_LAST", "REP_NAME", "SHOP_NAME"]
if "CODE" in df_customers.columns:
    base_cols.insert(1, "CODE")

customer_info = df_customers[base_cols].drop_duplicates()
customer_info["DEALER_MSISDN"] = customer_info["DEALER_MSISDN"].astype(str).str.strip()

bundle_final = bundle_pivot.merge(customer_info, on="DEALER_MSISDN", how="left")
cols_order = [c for c in ["DEALER_MSISDN", "CODE", "OWNER_FIRST", "OWNER_LAST", "REP_NAME", "SHOP_NAME"] if c in bundle_final.columns] + \
             [c for c in bundle_final.columns if c not in ["DEALER_MSISDN", "CODE", "OWNER_FIRST", "OWNER_LAST", "REP_NAME", "SHOP_NAME"]]
bundle_final = bundle_final[cols_order]

# =========================================================
# ✅ كل الباقات المخالفة
# =========================================================
df_b = df_bundle_full_for_report.copy()
df_b["DEALER_MSISDN"] = df_b["DEALER_MSISDN"].astype(str).str.strip()
df_b["SUB_MSISDN"] = df_b["SUB_MSISDN"].astype(str).str.strip()
df_b["BUNDLE_NAME_CLEAN"] = df_b["BUNDLE_NAME"].astype("string").str.strip()

df_b = df_b.merge(sub_to_type[["SUB_MSISDN", "PRODUCT_TYPE"]], on="SUB_MSISDN", how="left")

df_b["IS_YOOZ_BUNDLE"] = df_b["BUNDLE_NAME_CLEAN"].fillna("").apply(is_yooz_bundle)
df_b["IS_VIOLATION"] = (
    (df_b["PRODUCT_TYPE"] == "خط يوز") &
    (df_b["BUNDLE_NAME_CLEAN"].notna()) &
    (df_b["BUNDLE_NAME_CLEAN"].astype(str).str.strip() != "") &
    (~df_b["BUNDLE_NAME_CLEAN"].isin(EXCEPTION_BUNDLES)) &
    (~df_b["IS_YOOZ_BUNDLE"])
)

viol = df_b[df_b["IS_VIOLATION"]].copy()

if viol.empty:
    viol_counts = pd.DataFrame(columns=["DEALER_MSISDN", "BUNDLE_NAME_CLEAN", "VIOLATION_COUNT"])
    viol_total  = pd.DataFrame(columns=["DEALER_MSISDN", "YOOZ_VIOLATIONS_TOTAL"])
    viol_all    = pd.DataFrame(columns=["DEALER_MSISDN", "TOP_VIOLATION_BUNDLES"])
else:
    viol_counts = (
        viol.groupby(["DEALER_MSISDN", "BUNDLE_NAME_CLEAN"])
            .size()
            .reset_index(name="VIOLATION_COUNT")
    )

    viol_total = (
        viol_counts.groupby("DEALER_MSISDN")["VIOLATION_COUNT"]
                   .sum()
                   .reset_index(name="YOOZ_VIOLATIONS_TOTAL")
    )

    def all_violations_text(g):
        g = g.sort_values("VIOLATION_COUNT", ascending=False)
        return "، ".join([f"{r['BUNDLE_NAME_CLEAN']}({int(r['VIOLATION_COUNT'])})" for _, r in g.iterrows()])

    viol_all = (
        viol_counts.groupby("DEALER_MSISDN")
                   .apply(all_violations_text)
                   .reset_index(name="TOP_VIOLATION_BUNDLES")
    )

bundle_final = bundle_final.merge(viol_total, on="DEALER_MSISDN", how="left")
bundle_final = bundle_final.merge(viol_all,   on="DEALER_MSISDN", how="left")

bundle_final["YOOZ_VIOLATIONS_TOTAL"] = bundle_final["YOOZ_VIOLATIONS_TOTAL"].fillna(0).astype(int)
bundle_final["TOP_VIOLATION_BUNDLES"] = bundle_final["TOP_VIOLATION_BUNDLES"].fillna("—")

base_info = [c for c in ["DEALER_MSISDN", "CODE", "OWNER_FIRST", "OWNER_LAST", "REP_NAME", "SHOP_NAME"] if c in bundle_final.columns]
vi_cols = ["YOOZ_VIOLATIONS_TOTAL", "TOP_VIOLATION_BUNDLES"]
bundle_cols = [c for c in bundle_final.columns if c not in (base_info + vi_cols)]
bundle_final = bundle_final[base_info + vi_cols + bundle_cols]

# =========================================================
# تصدير Excel + تلوين المخالفات بالأصفر داخل "تقرير الباقات (كامل)"
# =========================================================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    report.to_excel(writer, sheet_name="تقرير المنتجات", index=False)
    bundle_final.to_excel(writer, sheet_name="تقرير الباقات (كامل)", index=False)
    daily_red_yooz_bundle_pivot.to_excel(writer, sheet_name="يومي ريد ويوز (ملخص)", index=False)
    all_products_horizontal_summary.to_excel(writer, sheet_name="ملخص جميع المنتجات", index=False)
    points_report.to_excel(writer, sheet_name="تقرير النقاط", index=False)

    # =========================================================
    # شيت مساعد للتلوين: Pivot للمخالفات فقط (مخفي)
    # =========================================================
    if viol.empty:
        viol_pivot = pd.DataFrame({"DEALER_MSISDN": bundle_final["DEALER_MSISDN"].astype(str).str.strip()})
        for b in all_bundles:
            viol_pivot[b] = 0
    else:
        viol_pivot = viol.pivot_table(
            index="DEALER_MSISDN",
            columns="BUNDLE_NAME_CLEAN",
            values="SUB_MSISDN",
            aggfunc="count",
            fill_value=0
        ).reindex(columns=all_bundles, fill_value=0).reset_index()

    viol_helper = bundle_final[["DEALER_MSISDN"]].merge(viol_pivot, on="DEALER_MSISDN", how="left").fillna(0)

    for col in bundle_final.columns:
        if col not in viol_helper.columns:
            viol_helper[col] = 0

    viol_helper = viol_helper[bundle_final.columns]
    viol_helper.to_excel(writer, sheet_name="مخالفات_تلوين", index=False)

    workbook = writer.book

    # =========================================================
    # =========================================================
    # =========================================================
    # إضافة فلاتر مباشرة إلى ورقة تقرير النقاط
    # =========================================================
    ws_points = writer.sheets["تقرير النقاط"]
    ws_points.right_to_left()
    ws_points.freeze_panes(1, 0)
    ws_points.hide_gridlines(2)

    if len(points_report.columns) > 0:
        points_table_columns = [
            {"header": str(col)} for col in points_report.columns
        ]

        # تحويل بيانات تقرير النقاط إلى جدول Excel بفلاتر متعددة الاختيار
        ws_points.add_table(
            0,
            0,
            max(len(points_report), 1),
            len(points_report.columns) - 1,
            {
                "name": "PointsReportTable",
                "columns": points_table_columns,
                "style": "Table Style Medium 2",
                "autofilter": True,
            }
        )

        for col_num, col_name in enumerate(points_report.columns):
            width = 15
            if col_name in ["SHOP_NAME_SHOW", "REP_NAME_SHOW"]:
                width = 28
            elif col_name == "DEALER_MSISDN":
                width = 18
            ws_points.set_column(col_num, col_num, width)

    # =========================================================
    # تقرير المندوب اليومي حسب تاريخ التفعيل
    # بدون معادلات، ويعمل على Excel القديم والجوال
    # يمكن تحديد مندوب، تاريخ، ومنتج واحد أو عدة منتجات من الفلاتر
    # =========================================================
    daily_rep_source = df_activation.copy()

    daily_rep_source["DEALER_MSISDN"] = (
        daily_rep_source["DEALER_MSISDN"]
        .astype(str)
        .str.strip()
    )
    daily_rep_source["SUB_MSISDN"] = (
        daily_rep_source["SUB_MSISDN"]
        .astype(str)
        .str.strip()
    )

    # إضافة بيانات المكتب والمندوب
    daily_customer_cols = [
        "DEALER_MSISDN", "SHOP_NAME", "REP_NAME"
    ]
    if "CODE" in df_customers.columns:
        daily_customer_cols.append("CODE")

    daily_rep_source = daily_rep_source.merge(
        df_customers[daily_customer_cols].drop_duplicates(),
        on="DEALER_MSISDN",
        how="left"
    )

    # إضافة نوع الخط من تقرير الباقات
    daily_rep_source = daily_rep_source.merge(
        bundle_type_per_sub.rename(
            columns={"PRODUCT_TYPE_FROM_BUNDLE": "PRODUCT_TYPE"}
        ),
        on="SUB_MSISDN",
        how="left"
    )

    daily_rep_source["PRODUCT_TYPE"] = (
        daily_rep_source["PRODUCT_TYPE"]
        .fillna("غير معروف")
    )
    daily_rep_source["REP_NAME"] = (
        daily_rep_source["REP_NAME"]
        .fillna("غير معروف")
    )
    daily_rep_source["SHOP_NAME"] = (
        daily_rep_source["SHOP_NAME"]
        .fillna("مكتب غير معروف")
    )

    if "CODE" in daily_rep_source.columns:
        daily_rep_source["CODE"] = (
            daily_rep_source["CODE"]
            .fillna("—")
            .astype(str)
            .str.replace(".0", "", regex=False)
        )

    daily_rep_source["BUNDLE_POINT"] = pd.to_numeric(
        daily_rep_source["BUNDLE_POINT"],
        errors="coerce"
    ).fillna(0)

    # التجميع اليومي: كل مكتب + مندوب + نوع خط + تاريخ
    daily_group_cols = [
        "ACTIVATION_DATE",
        "REP_NAME",
        "SHOP_NAME",
        "DEALER_MSISDN",
        "PRODUCT_TYPE",
    ]
    if "CODE" in daily_rep_source.columns:
        daily_group_cols.insert(4, "CODE")

    daily_rep_report = (
        daily_rep_source
        .groupby(daily_group_cols, dropna=False)
        .agg(
            عدد_التفعيلات=("SUB_MSISDN", "nunique"),
            مجموع_النقاط=("BUNDLE_POINT", "sum"),
        )
        .reset_index()
    )

    daily_rep_report = daily_rep_report.rename(columns={
        "ACTIVATION_DATE": "تاريخ التفعيل",
        "REP_NAME": "المندوب",
        "SHOP_NAME": "اسم المكتب",
        "DEALER_MSISDN": "رقم المكتب",
        "CODE": "الكود",
        "PRODUCT_TYPE": "نوع الخط",
        "عدد_التفعيلات": "عدد التفعيلات",
        "مجموع_النقاط": "مجموع النقاط",
    })

    daily_order = [
        "تاريخ التفعيل",
        "المندوب",
        "اسم المكتب",
        "رقم المكتب",
        "الكود",
        "نوع الخط",
        "عدد التفعيلات",
        "مجموع النقاط",
    ]
    daily_order = [
        c for c in daily_order if c in daily_rep_report.columns
    ]
    daily_rep_report = daily_rep_report[daily_order]

    daily_rep_report = daily_rep_report.sort_values(
        by=[
            c for c in [
                "تاريخ التفعيل",
                "المندوب",
                "اسم المكتب",
                "نوع الخط",
            ]
            if c in daily_rep_report.columns
        ]
    ).reset_index(drop=True)

    # كتابة الورقة الجديدة
    daily_sheet_name = "تقرير المندوب اليومي"
    daily_rep_report.to_excel(
        writer,
        sheet_name=daily_sheet_name,
        index=False,
        startrow=3
    )

    ws_daily = writer.sheets[daily_sheet_name]
    ws_daily.right_to_left()
    ws_daily.hide_gridlines(2)

    daily_title_fmt = workbook.add_format({
        "bold": True,
        "font_size": 16,
        "font_color": "#FFFFFF",
        "bg_color": "#1F4E78",
        "align": "center",
        "valign": "vcenter",
        "border": 1,
    })

    daily_note_fmt = workbook.add_format({
        "italic": True,
        "font_color": "#666666",
        "align": "right",
        "valign": "vcenter",
    })

    daily_date_fmt = workbook.add_format({
        "num_format": "dd/mm/yyyy"
    })

    daily_last_col = max(len(daily_rep_report.columns) - 1, 0)

    ws_daily.merge_range(
        0, 0, 0, daily_last_col,
        "التقرير اليومي للمندوب حسب تاريخ التفعيل",
        daily_title_fmt
    )

    ws_daily.merge_range(
        1, 0, 1, daily_last_col,
        "من سهم الفلتر اختر المندوب والتاريخ، ثم اختر نوع خط واحد أو عدة أنواع. سيظهر عدد التفعيلات ومجموع النقاط لكل مكتب.",
        daily_note_fmt
    )

    if len(daily_rep_report.columns) > 0:
        daily_table_columns = [
            {"header": str(col)} for col in daily_rep_report.columns
        ]

        ws_daily.add_table(
            3,
            0,
            3 + max(len(daily_rep_report), 1),
            daily_last_col,
            {
                "name": "DailyRepresentativeTable",
                "columns": daily_table_columns,
                "style": "Table Style Medium 2",
                "autofilter": True,
            }
        )

    ws_daily.freeze_panes(4, 0)

    for col_num, col_name in enumerate(daily_rep_report.columns):
        width = 16
        if col_name in ["اسم المكتب", "المندوب"]:
            width = 28
        elif col_name == "رقم المكتب":
            width = 18
        elif col_name == "تاريخ التفعيل":
            width = 15
            ws_daily.set_column(
                col_num, col_num, width, daily_date_fmt
            )
            continue
        ws_daily.set_column(col_num, col_num, width)

    ws_report = writer.sheets["تقرير الباقات (كامل)"]
    ws_viol   = writer.sheets["مخالفات_تلوين"]
    ws_viol.hide()

    yellow = workbook.add_format({'bg_color': '#FFF59D'})
    wrap_format = workbook.add_format({'text_wrap': True})

    if "TOP_VIOLATION_BUNDLES" in bundle_final.columns:
        col_idx_txt = bundle_final.columns.get_loc("TOP_VIOLATION_BUNDLES")
        ws_report.set_column(col_idx_txt, col_idx_txt, 60, wrap_format)

    info_cols_for_skip = [c for c in [
        "DEALER_MSISDN", "CODE", "OWNER_FIRST", "OWNER_LAST", "REP_NAME", "SHOP_NAME",
        "YOOZ_VIOLATIONS_TOTAL", "TOP_VIOLATION_BUNDLES"
    ] if c in bundle_final.columns]

    first_bundle_col = len(info_cols_for_skip)
    last_row = len(bundle_final)
    last_col = len(bundle_final.columns) - 1

    for col in range(first_bundle_col, last_col + 1):
        col_letter = xlsxwriter.utility.xl_col_to_name(col)
        ws_report.conditional_format(
            1, col,
            last_row, col,
            {
                'type': 'formula',
                'criteria': f"='مخالفات_تلوين'!{col_letter}2>0",
                'format': yellow
            }
        )

# =========================================================
# تجهيز ملف Excel للتنزيل بطريقة متوافقة مع السيرفر
# =========================================================
excel_bytes = output.getvalue()

if not excel_bytes:
    st.error("تعذر إنشاء ملف Excel. يرجى إعادة رفع الملفات والمحاولة مرة أخرى.")
    st.stop()

st.success("تم إعداد التقرير بنجاح")

st.download_button(
    label="تحميل التقرير الكامل (Excel)",
    data=excel_bytes,
    file_name="report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# =========================================================
# بوابة تقارير الأداء الفردية للمندوبين
# =========================================================
import zipfile
import re
from datetime import timedelta

st.divider()
st.header("📊 تقارير الأداء الفردية للمندوبين")
st.caption(
    "اختر المندوب لعرض تقرير متجاوب مناسب للجوال، مع إمكانية تحميل ملف Excel عند الحاجة."
)

def _safe_file_name(value):
    value = re.sub(r'[\\/:*?"<>|]+', "_", str(value)).strip()
    return value or "representative"

def _clean_number(value):
    try:
        number = float(value)
        if number.is_integer():
            return int(number)
        return round(number, 2)
    except Exception:
        return 0

def build_representative_workbook(rep_name):
    """إنشاء ملف Excel مستقل لمندوب واحد."""
    rep_raw = daily_rep_source[
        daily_rep_source["REP_NAME"].astype(str).str.strip() == str(rep_name).strip()
    ].copy()

    # بيانات التقرير حسب نطاق التاريخ المحدد في الشريط الجانبي
    report_raw = rep_raw[
        (rep_raw["ACTIVATION_DATE"] >= start_date) &
        (rep_raw["ACTIVATION_DATE"] <= end_date)
    ].copy()

    # إضافة مؤشرات الباقات للتنبيهات
    flags_for_rep = bundle_flags.copy()
    flags_for_rep["SUB_MSISDN"] = flags_for_rep["SUB_MSISDN"].astype(str).str.strip()

    report_flags = report_raw.merge(
        flags_for_rep,
        on="SUB_MSISDN",
        how="left"
    )
    for flag_col in [
        "has_any_bundle", "has_yooz_any", "has_non_yooz_violation",
        "has_daily300", "has_pubg"
    ]:
        if flag_col not in report_flags.columns:
            report_flags[flag_col] = False
        report_flags[flag_col] = report_flags[flag_col].fillna(False)

    # مؤشرات عامة بدون مقارنة زمنية
    total_activations = int(report_raw["SUB_MSISDN"].nunique())
    total_points = _clean_number(report_raw["BUNDLE_POINT"].sum())
    office_count = int(report_raw["DEALER_MSISDN"].nunique())
    selected_days = max((end_date - start_date).days + 1, 1)
    average_daily = round(total_activations / selected_days, 2)

    # الأداء اليومي
    daily_performance = (
        report_raw.groupby(
            ["ACTIVATION_DATE", "SHOP_NAME", "DEALER_MSISDN", "PRODUCT_TYPE"],
            dropna=False
        )
        .agg(
            **{
                "عدد التفعيلات": ("SUB_MSISDN", "nunique"),
                "مجموع النقاط": ("BUNDLE_POINT", "sum"),
            }
        )
        .reset_index()
        .rename(columns={
            "ACTIVATION_DATE": "تاريخ التفعيل",
            "SHOP_NAME": "اسم المكتب",
            "DEALER_MSISDN": "رقم المكتب",
            "PRODUCT_TYPE": "نوع الخط",
        })
    )

    if not daily_performance.empty:
        daily_performance["مجموع النقاط"] = daily_performance["مجموع النقاط"].apply(_clean_number)
        daily_performance = daily_performance.sort_values(
            ["تاريخ التفعيل", "اسم المكتب", "نوع الخط"]
        ).reset_index(drop=True)

    # ملخص الأيام: ريد + يوز + منتجات النت معاً
    def product_family(product_type):
        product_type = str(product_type).strip()

        if product_type == "خط ريد":
            return "ريد"

        if product_type == "خط يوز" or "يوز" in product_type:
            return "يوز"

        internet_products = {
            "خط نت 40",
            "نت 50",
            "ماي فاي",
            "ماي فاي جديد",
            "نت عادي",
            "راوتر",
            "راوتر جديد",
        }
        if product_type in internet_products:
            return "منتجات النت"

        return "أخرى"

    report_raw["مجموعة المنتج"] = report_raw["PRODUCT_TYPE"].apply(product_family)

    daily_family = (
        report_raw.groupby(
            ["ACTIVATION_DATE", "مجموعة المنتج"],
            dropna=False
        )["SUB_MSISDN"]
        .nunique()
        .reset_index(name="عدد التفعيلات")
    )

    daily_family_pivot = daily_family.pivot_table(
        index="ACTIVATION_DATE",
        columns="مجموعة المنتج",
        values="عدد التفعيلات",
        fill_value=0
    ).reset_index()

    for family_col in ["ريد", "يوز", "منتجات النت", "أخرى"]:
        if family_col not in daily_family_pivot.columns:
            daily_family_pivot[family_col] = 0

    daily_points = (
        report_raw.groupby("ACTIVATION_DATE", dropna=False)["BUNDLE_POINT"]
        .sum()
        .reset_index(name="النقاط")
    )

    daily_chart_data = daily_family_pivot.merge(
        daily_points,
        on="ACTIVATION_DATE",
        how="left"
    )

    daily_chart_data["إجمالي التفعيلات"] = (
        daily_chart_data["ريد"]
        + daily_chart_data["يوز"]
        + daily_chart_data["منتجات النت"]
        + daily_chart_data["أخرى"]
    )

    daily_chart_data = daily_chart_data.rename(
        columns={"ACTIVATION_DATE": "التاريخ"}
    )

    daily_chart_data = daily_chart_data[
        [
            "التاريخ",
            "ريد",
            "يوز",
            "منتجات النت",
            "أخرى",
            "إجمالي التفعيلات",
            "النقاط",
        ]
    ].sort_values("التاريخ").reset_index(drop=True)

    if not daily_chart_data.empty:
        daily_chart_data["النقاط"] = daily_chart_data["النقاط"].apply(_clean_number)
        for family_col in ["ريد", "يوز", "منتجات النت", "أخرى", "إجمالي التفعيلات"]:
            daily_chart_data[family_col] = daily_chart_data[family_col].astype(int)

    # أداء المنتجات بدون مقارنة زمنية
    product_summary = (
        report_raw.groupby("مجموعة المنتج", dropna=False)
        .agg(
            **{
                "عدد التفعيلات": ("SUB_MSISDN", "nunique"),
                "مجموع النقاط": ("BUNDLE_POINT", "sum"),
            }
        )
        .reset_index()
        .rename(columns={"مجموعة المنتج": "المنتج"})
    )

    preferred_products = ["ريد", "يوز", "منتجات النت", "أخرى"]
    product_summary["ترتيب"] = product_summary["المنتج"].apply(
        lambda value: preferred_products.index(value)
        if value in preferred_products else len(preferred_products)
    )
    product_summary = (
        product_summary.sort_values("ترتيب")
        .drop(columns=["ترتيب"])
        .reset_index(drop=True)
    )

    product_summary["عدد التفعيلات"] = product_summary["عدد التفعيلات"].apply(_clean_number)
    product_summary["مجموع النقاط"] = product_summary["مجموع النقاط"].apply(_clean_number)
    product_summary["متوسط النقاط لكل تفعيل"] = product_summary.apply(
        lambda row: round(row["مجموع النقاط"] / row["عدد التفعيلات"], 2)
        if row["عدد التفعيلات"] else 0,
        axis=1
    )

    total_product_activations = product_summary["عدد التفعيلات"].sum()
    total_product_points = product_summary["مجموع النقاط"].sum()

    product_summary["نسبة المساهمة من التفعيلات %"] = product_summary[
        "عدد التفعيلات"
    ].apply(
        lambda value: round(value / total_product_activations * 100, 2)
        if total_product_activations else 0
    )

    product_summary["نسبة المساهمة من النقاط %"] = product_summary[
        "مجموع النقاط"
    ].apply(
        lambda value: round(value / total_product_points * 100, 2)
        if total_product_points else 0
    )

    # أداء المكاتب بدون مقارنة زمنية
    offices_performance = (
        report_raw.groupby(["DEALER_MSISDN", "SHOP_NAME"], dropna=False)
        .agg(
            **{
                "عدد التفعيلات": ("SUB_MSISDN", "nunique"),
                "مجموع النقاط": ("BUNDLE_POINT", "sum"),
            }
        )
        .reset_index()
        .rename(columns={
            "DEALER_MSISDN": "رقم المكتب",
            "SHOP_NAME": "اسم المكتب",
        })
    )

    if not offices_performance.empty:
        offices_performance["عدد التفعيلات"] = offices_performance["عدد التفعيلات"].apply(_clean_number)
        offices_performance["مجموع النقاط"] = offices_performance["مجموع النقاط"].apply(_clean_number)
        offices_performance["متوسط التفعيل اليومي"] = (
            offices_performance["عدد التفعيلات"] / selected_days
        ).round(2)
        offices_performance["متوسط النقاط لكل تفعيل"] = offices_performance.apply(
            lambda row: round(row["مجموع النقاط"] / row["عدد التفعيلات"], 2)
            if row["عدد التفعيلات"] else 0,
            axis=1
        )
        offices_performance["الترتيب"] = (
            offices_performance["عدد التفعيلات"]
            .rank(method="dense", ascending=False)
            .astype(int)
        )
        offices_performance = offices_performance.sort_values(
            ["الترتيب", "اسم المكتب"]
        ).reset_index(drop=True)

    # التنبيهات
    alerts = []

    without_bundle_count = int(
        report_flags.loc[
            ~report_flags["has_any_bundle"],
            "SUB_MSISDN"
        ].nunique()
    )
    if without_bundle_count > 0:
        alerts.append({
            "نوع التنبيه": "خطوط بدون باقة",
            "المكتب": "جميع المكاتب",
            "التفاصيل": f"يوجد {without_bundle_count} خطاً مفعلاً بدون باقة.",
            "الأولوية": "عالية",
        })

    yooz_violation_count = int(
        report_flags.loc[
            (report_flags["PRODUCT_TYPE"] == "خط يوز") &
            (report_flags["has_non_yooz_violation"]),
            "SUB_MSISDN"
        ].nunique()
    )
    if yooz_violation_count > 0:
        alerts.append({
            "نوع التنبيه": "مخالفات يوز",
            "المكتب": "جميع المكاتب",
            "التفاصيل": f"يوجد {yooz_violation_count} خط يوز عليه باقة غير مناسبة.",
            "الأولوية": "عالية",
        })

    # تنبيهات المخزون من تقرير النقاط
    rep_points = points_report[
        points_report["REP_NAME_SHOW"].astype(str).str.strip() == str(rep_name).strip()
    ].copy()

    stock_col = None
    for candidate in ["total_stock_balance", "المخزون"]:
        if candidate in rep_points.columns:
            stock_col = candidate
            break

    if stock_col:
        zero_stock = rep_points[
            pd.to_numeric(rep_points[stock_col], errors="coerce").fillna(0) <= 0
        ]
        for _, row in zero_stock.drop_duplicates(
            subset=[c for c in ["DEALER_MSISDN", "SHOP_NAME_SHOW", "المنتج"] if c in zero_stock.columns]
        ).iterrows():
            alerts.append({
                "نوع التنبيه": "مخزون صفر",
                "المكتب": row.get("SHOP_NAME_SHOW", row.get("DEALER_MSISDN", "غير معروف")),
                "التفاصيل": f"لا يوجد مخزون للمنتج: {row.get('المنتج', 'غير معروف')}.",
                "الأولوية": "عالية",
            })

    if alerts:
        alerts_df = pd.DataFrame(alerts)
    else:
        alerts_df = pd.DataFrame([{
            "نوع التنبيه": "لا توجد تنبيهات",
            "المكتب": "—",
            "التفاصيل": "لا توجد ملاحظات حرجة في البيانات المحددة.",
            "الأولوية": "—",
        }])

    # البيانات التفصيلية: عرض أسماء الباقات بدلاً من رقم الخط
    bundle_names_per_sub = (
        df_flags_src[
            df_flags_src["BUNDLE_NAME_CLEAN"].notna()
            & df_flags_src["BUNDLE_NAME_CLEAN"].astype(str).str.strip().ne("")
        ]
        .groupby("SUB_MSISDN")["BUNDLE_NAME_CLEAN"]
        .agg(lambda values: "، ".join(dict.fromkeys(
            str(value).strip() for value in values if str(value).strip()
        )))
        .reset_index(name="أسماء الباقات")
    )

    raw_details_source = report_raw.merge(
        bundle_names_per_sub,
        on="SUB_MSISDN",
        how="left"
    )
    raw_details_source["أسماء الباقات"] = raw_details_source[
        "أسماء الباقات"
    ].fillna("بدون باقة")

    # إضافة المخزون / CAP حسب المكتب ونوع المنتج
    raw_details_source["DEALER_MSISDN_normalized"] = (
        raw_details_source["DEALER_MSISDN"].astype(str).str[-10:]
    )

    raw_details_source = raw_details_source.merge(
        df_stock_summary[
            ["DEALER_MSISDN_normalized", "PRODUCT_TYPE", "STOCK_BALANCE"]
        ],
        on=["DEALER_MSISDN_normalized", "PRODUCT_TYPE"],
        how="left"
    )
    raw_details_source["STOCK_BALANCE"] = pd.to_numeric(
        raw_details_source["STOCK_BALANCE"], errors="coerce"
    ).fillna(0)

    raw_columns = [
        "ACTIVATION_DATE", "REP_NAME", "SHOP_NAME", "DEALER_MSISDN",
        "PRODUCT_TYPE", "مجموعة المنتج", "أسماء الباقات",
        "STOCK_BALANCE", "BUNDLE_POINT"
    ]
    if "CODE" in raw_details_source.columns:
        raw_columns.insert(4, "CODE")
    raw_columns = [c for c in raw_columns if c in raw_details_source.columns]

    raw_details = raw_details_source[raw_columns].copy().rename(columns={
        "ACTIVATION_DATE": "تاريخ التفعيل",
        "REP_NAME": "المندوب",
        "SHOP_NAME": "اسم المكتب",
        "DEALER_MSISDN": "رقم المكتب",
        "CODE": "الكود",
        "PRODUCT_TYPE": "نوع الخط",
        "مجموعة المنتج": "فئة المنتج",
        "STOCK_BALANCE": "المخزون (CAP)",
        "BUNDLE_POINT": "النقاط",
    })

    # حسابات المندوب التي لم تسجل أي تفعيل في البيانات المحددة
    rep_accounts = df_customers[
        df_customers["REP_NAME"].astype(str).str.strip() == str(rep_name).strip()
    ].copy()

    active_dealers = set(
        report_raw["DEALER_MSISDN"].astype(str).str.strip().dropna().unique()
    )

    no_activation_accounts = rep_accounts[
        ~rep_accounts["DEALER_MSISDN"].astype(str).str.strip().isin(active_dealers)
    ].copy()

    no_activation_accounts["DEALER_MSISDN_normalized"] = (
        no_activation_accounts["DEALER_MSISDN"].astype(str).str[-10:]
    )

    stock_total_by_dealer = (
        df_stock_summary.groupby("DEALER_MSISDN_normalized", as_index=False)[
            "STOCK_BALANCE"
        ].sum()
        .rename(columns={"STOCK_BALANCE": "المخزون الكلي (CAP)"})
    )

    no_activation_accounts = no_activation_accounts.merge(
        stock_total_by_dealer,
        on="DEALER_MSISDN_normalized",
        how="left"
    )
    no_activation_accounts["المخزون الكلي (CAP)"] = pd.to_numeric(
        no_activation_accounts["المخزون الكلي (CAP)"], errors="coerce"
    ).fillna(0)

    no_activation_cols = [
        "DEALER_MSISDN", "SHOP_NAME", "REP_NAME",
        "OWNER_FIRST", "OWNER_LAST", "المخزون الكلي (CAP)"
    ]
    if "CODE" in no_activation_accounts.columns:
        no_activation_cols.insert(1, "CODE")
    no_activation_cols = [
        c for c in no_activation_cols if c in no_activation_accounts.columns
    ]

    no_activation_accounts = no_activation_accounts[
        no_activation_cols
    ].drop_duplicates().rename(columns={
        "DEALER_MSISDN": "رقم الحساب",
        "CODE": "الكود",
        "SHOP_NAME": "اسم المكتب",
        "REP_NAME": "المندوب",
        "OWNER_FIRST": "اسم المالك",
        "OWNER_LAST": "اللقب",
    })

    # إنشاء ملف Excel
    rep_output = BytesIO()

    with pd.ExcelWriter(rep_output, engine="xlsxwriter") as rep_writer:
        workbook = rep_writer.book

        # تنسيقات عامة
        title_fmt = workbook.add_format({
            "bold": True, "font_size": 18, "font_color": "#FFFFFF",
            "bg_color": "#1F4E78", "align": "center", "valign": "vcenter",
        })
        subtitle_fmt = workbook.add_format({
            "font_size": 11, "font_color": "#666666",
            "align": "center", "valign": "vcenter",
        })
        kpi_label_fmt = workbook.add_format({
            "bold": True, "font_color": "#FFFFFF", "bg_color": "#4472C4",
            "align": "center", "border": 1,
        })
        kpi_value_fmt = workbook.add_format({
            "bold": True, "font_size": 16, "bg_color": "#D9EAF7",
            "align": "center", "border": 1,
        })
        percent_fmt = workbook.add_format({"num_format": "0.00%"})
        date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy"})
        green_fmt = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        yellow_fmt = workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500"})
        red_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        # ورقة الملخص
        ws_dashboard = workbook.add_worksheet("ملخص الأداء")
        ws_dashboard.right_to_left()
        ws_dashboard.hide_gridlines(2)
        ws_dashboard.set_column("A:A", 3)
        ws_dashboard.set_column("B:H", 18)

        ws_dashboard.merge_range("B2:H3", f"تقرير أداء المندوب: {rep_name}", title_fmt)
        no_activation_count = len(no_activation_accounts)

        # ملخص الأداء بدون مؤشرات مقارنة بين الفترات
        kpis = [
            ("إجمالي التفعيلات", total_activations),
            ("مجموع النقاط", total_points),
            ("المكاتب النشطة", office_count),
            ("المتوسط اليومي", average_daily),
            ("خطوط بدون باقة", without_bundle_count),
            ("حسابات بدون تفعيل", no_activation_count),
        ]

        for idx, (label, value) in enumerate(kpis):
            row = 4 + (idx // 3) * 3
            col = 1 + (idx % 3) * 2
            ws_dashboard.merge_range(row, col, row, col + 1, label, kpi_label_fmt)
            ws_dashboard.merge_range(row + 1, col, row + 2, col + 1, value, kpi_value_fmt)

        # بيانات الرسوم في أسفل لوحة الملخص
        chart_daily_start = 15
        daily_headers = [
            "التاريخ",
            "ريد",
            "يوز",
            "منتجات النت",
            "أخرى",
            "إجمالي التفعيلات",
            "النقاط",
        ]
        ws_dashboard.write_row(chart_daily_start, 1, daily_headers)

        for row_idx, row in daily_chart_data.iterrows():
            excel_row = chart_daily_start + 1 + row_idx
            ws_dashboard.write_datetime(
                excel_row,
                1,
                pd.Timestamp(row["التاريخ"]).to_pydatetime(),
                date_fmt
            )
            ws_dashboard.write(excel_row, 2, row["ريد"])
            ws_dashboard.write(excel_row, 3, row["يوز"])
            ws_dashboard.write(excel_row, 4, row["منتجات النت"])
            ws_dashboard.write(excel_row, 5, row["أخرى"])
            ws_dashboard.write(excel_row, 6, row["إجمالي التفعيلات"])
            ws_dashboard.write(excel_row, 7, row["النقاط"])

        if len(daily_chart_data) > 0:
            daily_chart = workbook.add_chart({"type": "column", "subtype": "stacked"})

            series_columns = [
                ("ريد", 2),
                ("يوز", 3),
                ("منتجات النت", 4),
            ]

            for series_name, series_col in series_columns:
                daily_chart.add_series({
                    "name": series_name,
                    "categories": [
                        "ملخص الأداء",
                        chart_daily_start + 1,
                        1,
                        chart_daily_start + len(daily_chart_data),
                        1,
                    ],
                    "values": [
                        "ملخص الأداء",
                        chart_daily_start + 1,
                        series_col,
                        chart_daily_start + len(daily_chart_data),
                        series_col,
                    ],
                })

            daily_chart.set_title({"name": "التفعيل اليومي: ريد + يوز + منتجات النت"})
            daily_chart.set_x_axis({"name": "التاريخ", "date_axis": True})
            daily_chart.set_y_axis({
                "name": "عدد التفعيلات",
                "major_gridlines": {"visible": False}
            })
            daily_chart.set_legend({"position": "bottom"})
            ws_dashboard.insert_chart(
                "J15",
                daily_chart,
                {"x_scale": 1.35, "y_scale": 1.1}
            )

        product_chart_start = chart_daily_start + max(len(daily_chart_data), 1) + 4
        ws_dashboard.write_row(product_chart_start, 1, ["المنتج", "عدد التفعيلات"])
        for row_idx, row in product_summary.iterrows():
            ws_dashboard.write(product_chart_start + 1 + row_idx, 1, row["المنتج"])
            ws_dashboard.write(product_chart_start + 1 + row_idx, 2, row["عدد التفعيلات"])

        if len(product_summary) > 0:
            product_chart = workbook.add_chart({"type": "column"})
            product_chart.add_series({
                "name": "التفعيلات",
                "categories": ["ملخص الأداء", product_chart_start + 1, 1,
                               product_chart_start + len(product_summary), 1],
                "values": ["ملخص الأداء", product_chart_start + 1, 2,
                           product_chart_start + len(product_summary), 2],
                "data_labels": {"value": True},
            })
            product_chart.set_title({"name": "التفعيل حسب نوع الخط"})
            product_chart.set_y_axis({"name": "عدد التفعيلات", "major_gridlines": {"visible": False}})
            product_chart.set_legend({"none": True})
            ws_dashboard.insert_chart(
                product_chart_start, 5, product_chart,
                {"x_scale": 1.25, "y_scale": 1.05}
            )

        # كتابة بقية الأوراق
        sheets_data = [
            ("الأداء اليومي", daily_performance),
            ("أداء المكاتب", offices_performance),
            ("أداء المنتجات", product_summary),
            ("التنبيهات", alerts_df),
            ("حسابات بدون تفعيل", no_activation_accounts),
            ("البيانات التفصيلية", raw_details),
        ]

        for sheet_name, dataframe in sheets_data:
            dataframe.to_excel(rep_writer, sheet_name=sheet_name, index=False)
            ws = rep_writer.sheets[sheet_name]
            ws.right_to_left()
            ws.freeze_panes(1, 0)
            ws.hide_gridlines(2)

            if len(dataframe.columns) > 0:
                table_name = re.sub(r"\W+", "", sheet_name) + "Table"
                if not table_name or not table_name[0].isalpha():
                    table_name = "T" + table_name

                ws.add_table(
                    0, 0,
                    max(len(dataframe), 1),
                    len(dataframe.columns) - 1,
                    {
                        "name": table_name,
                        "columns": [{"header": str(c)} for c in dataframe.columns],
                        "style": "Table Style Medium 2",
                    }
                )

            for col_idx, col_name in enumerate(dataframe.columns):
                width = 16
                if col_name in ["اسم المكتب", "التفاصيل"]:
                    width = 34
                elif col_name in ["المندوب", "نوع التنبيه"]:
                    width = 24
                elif col_name in ["تاريخ التفعيل", "التاريخ"]:
                    width = 15
                    ws.set_column(col_idx, col_idx, width, date_fmt)
                    continue
                ws.set_column(col_idx, col_idx, width)

        # تلوين التقييم
        if "التقييم" in offices_performance.columns and len(offices_performance) > 0:
            ws_offices = rep_writer.sheets["أداء المكاتب"]
            eval_col = offices_performance.columns.get_loc("التقييم")
            ws_offices.conditional_format(
                1, eval_col, len(offices_performance), eval_col,
                {"type": "text", "criteria": "containing", "value": "ممتاز", "format": green_fmt}
            )
            ws_offices.conditional_format(
                1, eval_col, len(offices_performance), eval_col,
                {"type": "text", "criteria": "containing", "value": "يحتاج متابعة", "format": yellow_fmt}
            )
            ws_offices.conditional_format(
                1, eval_col, len(offices_performance), eval_col,
                {"type": "text", "criteria": "containing", "value": "ضعيف", "format": red_fmt}
            )
            ws_offices.conditional_format(
                1, eval_col, len(offices_performance), eval_col,
                {"type": "text", "criteria": "containing", "value": "متوقف", "format": red_fmt}
            )

        # تلوين أولوية التنبيهات
        ws_alerts = rep_writer.sheets["التنبيهات"]
        if "الأولوية" in alerts_df.columns and len(alerts_df) > 0:
            priority_col = alerts_df.columns.get_loc("الأولوية")
            ws_alerts.conditional_format(
                1, priority_col, len(alerts_df), priority_col,
                {"type": "text", "criteria": "containing", "value": "عالية", "format": red_fmt}
            )
            ws_alerts.conditional_format(
                1, priority_col, len(alerts_df), priority_col,
                {"type": "text", "criteria": "containing", "value": "متوسطة", "format": yellow_fmt}
            )

    rep_output.seek(0)
    return rep_output.getvalue()



# =========================================================
# لوحة عرض متجاوبة للجوال
# =========================================================
def build_mobile_dashboard_data(rep_name):
    """تجهيز بيانات عرض مبسطة ومتجاوبة للمندوب داخل Streamlit."""
    rep_raw = daily_rep_source[
        daily_rep_source["REP_NAME"].astype(str).str.strip() == str(rep_name).strip()
    ].copy()

    report_raw = rep_raw[
        (rep_raw["ACTIVATION_DATE"] >= start_date) &
        (rep_raw["ACTIVATION_DATE"] <= end_date)
    ].copy()

    selected_days = max((end_date - start_date).days + 1, 1)

    flags_for_rep = bundle_flags.copy()
    flags_for_rep["SUB_MSISDN"] = flags_for_rep["SUB_MSISDN"].astype(str).str.strip()

    report_flags = report_raw.merge(flags_for_rep, on="SUB_MSISDN", how="left")
    for flag_col in [
        "has_any_bundle", "has_yooz_any", "has_non_yooz_violation",
        "has_daily300", "has_pubg"
    ]:
        if flag_col not in report_flags.columns:
            report_flags[flag_col] = False
        report_flags[flag_col] = report_flags[flag_col].fillna(False)

    def product_family_mobile(product_type):
        product_type = str(product_type).strip()
        if product_type == "خط ريد":
            return "ريد"
        if product_type == "خط يوز" or "يوز" in product_type:
            return "يوز"
        internet_products = {
            "خط نت 40", "نت 50", "ماي فاي", "ماي فاي جديد",
            "نت عادي", "راوتر", "راوتر جديد",
        }
        if product_type in internet_products:
            return "منتجات النت"
        return "أخرى"

    report_raw["مجموعة المنتج"] = report_raw["PRODUCT_TYPE"].apply(product_family_mobile)

    total_activations = int(report_raw["SUB_MSISDN"].nunique())
    total_points = _clean_number(report_raw["BUNDLE_POINT"].sum())
    office_count = int(report_raw["DEALER_MSISDN"].nunique())
    average_daily = round(total_activations / selected_days, 2)
    without_bundle_count = int(
        report_flags.loc[~report_flags["has_any_bundle"], "SUB_MSISDN"].nunique()
    )

    daily_summary = (
        report_raw.groupby(["ACTIVATION_DATE", "مجموعة المنتج"], dropna=False)["SUB_MSISDN"]
        .nunique()
        .reset_index(name="عدد التفعيلات")
        .pivot_table(
            index="ACTIVATION_DATE",
            columns="مجموعة المنتج",
            values="عدد التفعيلات",
            fill_value=0
        )
        .reset_index()
    )

    for family_col in ["ريد", "يوز", "منتجات النت", "أخرى"]:
        if family_col not in daily_summary.columns:
            daily_summary[family_col] = 0

    daily_points = (
        report_raw.groupby("ACTIVATION_DATE", dropna=False)["BUNDLE_POINT"]
        .sum()
        .reset_index(name="النقاط")
    )

    daily_summary = daily_summary.merge(
        daily_points, on="ACTIVATION_DATE", how="left"
    ).rename(columns={"ACTIVATION_DATE": "التاريخ"})

    daily_summary["إجمالي التفعيلات"] = (
        daily_summary["ريد"] + daily_summary["يوز"]
        + daily_summary["منتجات النت"] + daily_summary["أخرى"]
    )

    daily_summary = daily_summary[
        ["التاريخ", "ريد", "يوز", "منتجات النت", "أخرى", "إجمالي التفعيلات", "النقاط"]
    ].sort_values("التاريخ").reset_index(drop=True)

    offices = (
        report_raw.groupby(["DEALER_MSISDN", "SHOP_NAME"], dropna=False)
        .agg(
            **{
                "عدد التفعيلات": ("SUB_MSISDN", "nunique"),
                "مجموع النقاط": ("BUNDLE_POINT", "sum"),
            }
        )
        .reset_index()
        .rename(columns={
            "DEALER_MSISDN": "رقم المكتب",
            "SHOP_NAME": "اسم المكتب",
        })
    )

    if not offices.empty:
        offices["متوسط يومي"] = (offices["عدد التفعيلات"] / selected_days).round(2)
        offices["مجموع النقاط"] = offices["مجموع النقاط"].apply(_clean_number)
        offices = offices.sort_values(
            ["عدد التفعيلات", "مجموع النقاط"],
            ascending=[False, False]
        ).reset_index(drop=True)
        offices.insert(0, "الترتيب", range(1, len(offices) + 1))

    products = (
        report_raw.groupby("مجموعة المنتج", dropna=False)
        .agg(
            **{
                "عدد التفعيلات": ("SUB_MSISDN", "nunique"),
                "مجموع النقاط": ("BUNDLE_POINT", "sum"),
            }
        )
        .reset_index()
        .rename(columns={"مجموعة المنتج": "المنتج"})
    )
    if not products.empty:
        products["مجموع النقاط"] = products["مجموع النقاط"].apply(_clean_number)
        products = products.sort_values("عدد التفعيلات", ascending=False).reset_index(drop=True)

    rep_accounts = df_customers[
        df_customers["REP_NAME"].astype(str).str.strip() == str(rep_name).strip()
    ].copy()

    active_dealers = set(
        report_raw["DEALER_MSISDN"].astype(str).str.strip().dropna().unique()
    )

    no_activation = rep_accounts[
        ~rep_accounts["DEALER_MSISDN"].astype(str).str.strip().isin(active_dealers)
    ].copy()

    no_activation["DEALER_MSISDN_normalized"] = (
        no_activation["DEALER_MSISDN"].astype(str).str[-10:]
    )

    stock_total_by_dealer = (
        df_stock_summary.groupby("DEALER_MSISDN_normalized", as_index=False)["STOCK_BALANCE"]
        .sum()
        .rename(columns={"STOCK_BALANCE": "المخزون الكلي (CAP)"})
    )

    no_activation = no_activation.merge(
        stock_total_by_dealer,
        on="DEALER_MSISDN_normalized",
        how="left"
    )
    no_activation["المخزون الكلي (CAP)"] = pd.to_numeric(
        no_activation["المخزون الكلي (CAP)"], errors="coerce"
    ).fillna(0)

    no_activation_cols = [
        "DEALER_MSISDN", "CODE", "SHOP_NAME",
        "OWNER_FIRST", "OWNER_LAST", "المخزون الكلي (CAP)"
    ]
    no_activation_cols = [c for c in no_activation_cols if c in no_activation.columns]
    no_activation = no_activation[no_activation_cols].rename(columns={
        "DEALER_MSISDN": "رقم المكتب",
        "CODE": "الكود",
        "SHOP_NAME": "اسم المكتب",
        "OWNER_FIRST": "اسم المالك",
        "OWNER_LAST": "لقب المالك",
    }).reset_index(drop=True)

    metrics = {
        "إجمالي التفعيلات": total_activations,
        "مجموع النقاط": total_points,
        "المكاتب النشطة": office_count,
        "المتوسط اليومي": average_daily,
        "خطوط بدون باقة": without_bundle_count,
        "حسابات بدون تفعيل": len(no_activation),
    }

    return {
        "metrics": metrics,
        "daily": daily_summary,
        "offices": offices,
        "products": products,
        "no_activation": no_activation,
    }


def render_mobile_dashboard(rep_name):
    dashboard = build_mobile_dashboard_data(rep_name)

    st.markdown(
        """
        <style>
        .mobile-kpi-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 12px;
            margin: 10px 0 18px 0;
        }
        .mobile-kpi-card {
            background: white;
            border: 1px solid #e6e8eb;
            border-radius: 14px;
            padding: 14px 10px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        .mobile-kpi-title {
            color: #5f6368;
            font-size: 0.88rem;
            margin-bottom: 6px;
        }
        .mobile-kpi-value {
            color: #0f5132;
            font-size: 1.55rem;
            font-weight: 700;
            line-height: 1.2;
        }
        @media (max-width: 700px) {
            .block-container {
                padding-top: 1rem;
                padding-left: 0.65rem;
                padding-right: 0.65rem;
            }
            .mobile-kpi-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 8px;
            }
            .mobile-kpi-card {
                padding: 12px 6px;
            }
            .mobile-kpi-title {
                font-size: 0.78rem;
            }
            .mobile-kpi-value {
                font-size: 1.3rem;
            }
            div[data-testid="stDataFrame"] {
                font-size: 0.78rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    metric_cards = "".join(
        f"""
        <div class="mobile-kpi-card">
            <div class="mobile-kpi-title">{label}</div>
            <div class="mobile-kpi-value">{value}</div>
        </div>
        """
        for label, value in dashboard["metrics"].items()
    )
    st.markdown(
        f'<div class="mobile-kpi-grid">{metric_cards}</div>',
        unsafe_allow_html=True,
    )

    tab_summary, tab_daily, tab_offices, tab_products, tab_inactive = st.tabs([
        "📊 الملخص",
        "📅 اليومي",
        "🏢 المكاتب",
        "📦 المنتجات",
        "⛔ بدون تفعيل",
    ])

    with tab_summary:
        st.subheader(f"أداء {rep_name}")
        if dashboard["daily"].empty:
            st.info("لا توجد بيانات ضمن التاريخ المحدد.")
        else:
            chart_data = dashboard["daily"].set_index("التاريخ")[
                ["ريد", "يوز", "منتجات النت", "أخرى"]
            ]
            st.bar_chart(chart_data, use_container_width=True)

            points_data = dashboard["daily"].set_index("التاريخ")[["النقاط"]]
            st.line_chart(points_data, use_container_width=True)

    with tab_daily:
        st.dataframe(
            dashboard["daily"],
            use_container_width=True,
            hide_index=True,
            height=430,
        )

    with tab_offices:
        search_office = st.text_input(
            "بحث باسم أو رقم المكتب",
            key=f"mobile_office_search_{rep_name}",
            placeholder="اكتب جزءاً من اسم المكتب أو رقمه",
        )
        offices_view = dashboard["offices"].copy()
        if search_office and not offices_view.empty:
            mask = (
                offices_view["اسم المكتب"].astype(str).str.contains(
                    search_office, case=False, na=False
                )
                | offices_view["رقم المكتب"].astype(str).str.contains(
                    search_office, case=False, na=False
                )
            )
            offices_view = offices_view[mask]

        st.dataframe(
            offices_view,
            use_container_width=True,
            hide_index=True,
            height=480,
        )

    with tab_products:
        st.dataframe(
            dashboard["products"],
            use_container_width=True,
            hide_index=True,
        )
        if not dashboard["products"].empty:
            st.bar_chart(
                dashboard["products"].set_index("المنتج")[["عدد التفعيلات"]],
                use_container_width=True,
            )

    with tab_inactive:
        st.caption(
            f"عدد الحسابات بدون تفعيل: {len(dashboard['no_activation'])}"
        )
        st.dataframe(
            dashboard["no_activation"],
            use_container_width=True,
            hide_index=True,
            height=480,
        )


available_reps = sorted(
    daily_rep_source["REP_NAME"]
    .dropna()
    .astype(str)
    .str.strip()
    .loc[lambda series: series.ne("") & series.ne("غير معروف")]
    .unique()
    .tolist()
)

if available_reps:
    selected_rep = st.selectbox(
        "اختر المندوب",
        available_reps,
        key="selected_representative_performance"
    )

    st.subheader("📱 العرض المبسط للجوال")
    render_mobile_dashboard(selected_rep)

    st.divider()
    st.subheader("📥 تحميل التقرير")
    selected_rep_excel = build_representative_workbook(selected_rep)

    st.download_button(
        label=f"⬇️ تحميل تقرير أداء {selected_rep}",
        data=selected_rep_excel,
        file_name=f"تقرير_أداء_{_safe_file_name(selected_rep)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_selected_representative"
    )

    with st.expander("تحميل تقارير جميع المندوبين في ملف ZIP"):
        st.caption("قد يستغرق الإنشاء وقتاً أطول عند وجود عدد كبير من المندوبين.")

        if st.button("إنشاء ملف ZIP لجميع المندوبين", key="build_all_reps_zip"):
            zip_output = BytesIO()
            progress = st.progress(0)

            with zipfile.ZipFile(
                zip_output,
                mode="w",
                compression=zipfile.ZIP_DEFLATED
            ) as zip_file:
                for index, rep in enumerate(available_reps, start=1):
                    rep_bytes = build_representative_workbook(rep)
                    zip_file.writestr(
                        f"تقرير_أداء_{_safe_file_name(rep)}.xlsx",
                        rep_bytes
                    )
                    progress.progress(index / len(available_reps))

            zip_output.seek(0)
            st.session_state["all_representatives_zip"] = zip_output.getvalue()
            st.success("تم تجهيز تقارير جميع المندوبين.")

        if "all_representatives_zip" in st.session_state:
            st.download_button(
                label="⬇️ تحميل جميع تقارير المندوبين (ZIP)",
                data=st.session_state["all_representatives_zip"],
                file_name="تقارير_أداء_جميع_المندوبين.zip",
                mime="application/zip",
                key="download_all_reps_zip"
            )
else:
    st.warning("لا توجد أسماء مندوبين متاحة لإنشاء التقارير الفردية.")
