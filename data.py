import pandas as pd

# ===================== Helper loaders =====================
def _load_raw(filename):
    d = pd.read_csv(filename, encoding="utf-8-sig", low_memory=False)
    d.columns = d.columns.str.replace('\ufeff', '', regex=False).str.strip()
    for col in ["Service_Eng", "Office_Eng", "District_Eng", "Received", "Disposed_Out"]:
        if col not in d.columns:
            d[col] = 0
    d["Received"] = pd.to_numeric(d["Received"], errors="coerce").fillna(0).astype(int)
    d["Disposed_Out"] = pd.to_numeric(d["Disposed_Out"], errors="coerce").fillna(0).astype(int)
    return d

def _load_tt(filename):
    d = pd.read_excel(filename)
    d.columns = d.columns.str.strip()
    for col in ["District_Eng", "Office_Eng", "Service_Eng", "Disposed_Out", "Disposed", "Received", "Total", "Pending"]:
        if col not in d.columns:
            d[col] = 0
    for col in ["Disposed_Out", "Disposed", "Received", "Total", "Pending"]:
        d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0)
    d["Late_Disposed_%"] = (d["Disposed_Out"] / d["Disposed"].replace(0, pd.NA) * 100).fillna(0)
    return d

def _load_mt(filename):
    d = pd.read_csv(filename, encoding="utf-8-sig", low_memory=False)
    d.columns = d.columns.str.replace('\ufeff', '', regex=False).str.strip()
    d = d.rename(columns={
        'Yr': 'Year', 'Mn': 'Month',
        'Service_Eng': 'Service_name', 'District_Eng': 'District_name', 'Office_Eng': 'Office_name',
        'Received': 'application_Received', 'Disposed': 'application_Disposed',
        'Disposed_Out': 'application_Disposed_Out_of_time',
        'Disposed_with_in': 'application_Disposed_with_in_time',
    })
    d['Year'] = d['Year'].astype(str).str.replace('\ufeff', '', regex=False).str.strip()
    d['Year'] = pd.to_numeric(d['Year'], errors='coerce')
    d.dropna(subset=['Year'], inplace=True)
    d['Year'] = d['Year'].astype(int)
    d['Month'] = pd.to_numeric(d['Month'], errors='coerce').fillna(0).astype(int)
    d['Date'] = pd.to_datetime(
        d['Year'].astype(str) + '-' + d['Month'].astype(str).str.zfill(2) + '-01',
        format='%Y-%m-%d'
    )
    d['Month_Year'] = d['Date'].dt.strftime('%b-%Y')
    for col in ['application_Received', 'application_Disposed', 'application_Disposed_Out_of_time', 'application_Disposed_with_in_time']:
        d[col] = pd.to_numeric(d[col], errors='coerce').fillna(0).astype(int)
    d['Efficiency_Percentage'] = d.apply(
        lambda row: (row['application_Disposed_with_in_time'] / row['application_Disposed'] * 100)
        if row['application_Disposed'] > 0 else 0, axis=1
    )
    d['Pending_Applications'] = d['application_Received'] - d['application_Disposed']
    for col in ['District_name', 'Service_name', 'Office_name']:
        d[col] = d[col].astype(str).str.strip()
    return d

def _make_month_options(df_mt):
    months = sorted(df_mt['Month_Year'].unique(), key=lambda x: pd.to_datetime(x, format='%b-%Y'))
    opts = [{'label': m, 'value': m} for m in months]
    opts.insert(0, {'label': 'All Months', 'value': 'ALL_MONTHS'})
    return months, opts

def _district_order(df_tt):
    return df_tt.groupby("District_Eng")["Disposed_Out"].sum().sort_values(ascending=False).index.tolist()

# ===================== Load both FY datasets =====================
df_2425 = _load_raw("digital-data.csv")
df_2526 = _load_raw("digital-data-25-26.csv")

df_tt_2425 = _load_tt("selected_columns.xlsx")
df_tt_2526 = _load_tt("selected_columns-25-26.xlsx")

df_mt_2425 = _load_mt("digital-data.csv")
df_mt_2526 = _load_mt("digital-data-25-26.csv")

df_mt_light_2425 = df_mt_2425[['Month_Year', 'District_name', 'Service_name', 'Office_name', 'application_Disposed_Out_of_time']].copy()
df_mt_light_2526 = df_mt_2526[['Month_Year', 'District_name', 'Service_name', 'Office_name', 'application_Disposed_Out_of_time']].copy()

all_months_2425, month_options_2425 = _make_month_options(df_mt_2425)
all_months_2526, month_options_2526 = _make_month_options(df_mt_2526)

district_order_2425 = _district_order(df_tt_2425)
district_order_2526 = _district_order(df_tt_2526)

# ===================== FY lookup dict =====================
FY_DATA = {
    '2425': {
        'df': df_2425, 'df_tt': df_tt_2425,
        'df_mt': df_mt_2425, 'df_mt_light': df_mt_light_2425,
        'all_months': all_months_2425, 'month_options': month_options_2425,
        'district_order': district_order_2425,
        'label': 'FY 2024-25',
    },
    '2526': {
        'df': df_2526, 'df_tt': df_tt_2526,
        'df_mt': df_mt_2526, 'df_mt_light': df_mt_light_2526,
        'all_months': all_months_2526, 'month_options': month_options_2526,
        'district_order': district_order_2526,
        'label': 'FY 2025-26',
    },
}

# ===================== Backward-compat aliases (used by advanced_analytics, oot_drilldown) =====================
df = df_2425
df_tt = df_tt_2425
df_mt = df_mt_2425
df_mt_light = df_mt_light_2425
all_months = all_months_2425
month_options = month_options_2425
district_order = district_order_2425

hierarchies = {
    "District → Office → Service": ["District_Eng", "Office_Eng", "Service_Eng"],
    "District → Service → Office": ["District_Eng", "Service_Eng", "Office_Eng"],
    "Office → District → Service": ["Office_Eng", "District_Eng", "Service_Eng"],
    "Office → Service → District": ["Office_Eng", "Service_Eng", "District_Eng"],
    "Service → District → Office": ["Service_Eng", "District_Eng", "Office_Eng"],
    "Service → Office → District": ["Service_Eng", "Office_Eng", "District_Eng"],
}

def categorize(x):
    if x < 10: return "A"
    if x < 40: return "B"
    if x < 70: return "C"
    return "D"

def get_district_summary(filtered_df):
    summary = filtered_df.groupby("District_Eng", as_index=False).agg({
        "Disposed_Out": "sum", "Disposed": "sum", "Received": "sum"
    })
    summary["Late_Disposed_%"] = (summary["Disposed_Out"] / summary["Disposed"].replace(0, pd.NA) * 100).fillna(0)
    summary["Category"] = summary["Late_Disposed_%"].apply(categorize)
    summary["Office_Eng"] = ""
    summary["Service_Eng"] = ""
    return summary

initial_summary_tt = get_district_summary(df_tt_2425)

COLOR_PALETTE = {
    'primary': '#1f77b4', 'secondary': '#ff7f0e', 'success': '#2ca02c',
    'danger': '#d62728', 'light': '#f8f9fa', 'dark': '#343a40'
}

# ===================== Advanced Analytics Data =====================
try:
    df_adv = pd.read_csv("Merge-digital-data-22-25.csv", encoding="utf-8-sig", low_memory=False)
    df_adv.columns = df_adv.columns.str.strip()
    for c in ['Year', 'Month', 'application_Disposed_Out_of_time', 'application_Disposed']:
        if c in df_adv.columns:
            df_adv[c] = pd.to_numeric(df_adv[c], errors='coerce').fillna(0).astype(int)
except Exception as e:
    print(f"Warning: Failed to load Advanced Analytics data: {e}")
    df_adv = pd.DataFrame(columns=['Year', 'Month', 'Service_name', 'District_name', 'Office_name', 'application_Disposed_Out_of_time', 'application_Disposed'])