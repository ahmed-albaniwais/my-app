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
        lambda x: (x.ne("") & x.apply(is_yooz_bundle)).any()
    )),
    has_non_yooz_violation=("BUNDLE_NAME_CLEAN", lambda s: s.dropna().astype(str).str.strip().pipe(
        lambda x: (x.ne("") & (~x.isin(EXCEPTION_BUNDLES)) & (~x.apply(is_yooz_bundle))).any()
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
    df_stock_summary.to_excel(writer, sheet_name="تقرير المخزون", index=False)
    yooz_violations_report.to_excel(writer, sheet_name="مخالفات يوز", index=False)
    daily_red_yooz_bundle_summary.to_excel(writer, sheet_name="يومي ريد ويوز (تفصيلي)", index=False)
    daily_red_yooz_bundle_pivot.to_excel(writer, sheet_name="يومي ريد ويوز (ملخص)", index=False)
    red_yooz_combined_summary.to_excel(writer, sheet_name="ملخص ريد+يوز", index=False)
    all_products_horizontal_summary.to_excel(writer, sheet_name="ملخص جميع المنتجات", index=False)

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

st.success("تم إعداد التقرير بنجاح")

st.download_button(
    label="تحميل التقرير الكامل (Excel)",
    data=output.getvalue(),
    file_name="report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
