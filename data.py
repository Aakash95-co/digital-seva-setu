# import pandas as pd
#
# # ===================== Metadata Data =====================
# df = pd.read_csv("digital-data.csv", encoding="utf-8-sig", low_memory=False)
# df.columns = df.columns.str.replace('\ufeff', '', regex=False).str.strip()
# for col in["Service_Eng", "Office_Eng", "District_Eng", "Received", "Disposed_Out"]:
#     if col not in df.columns:
#         df[col] = 0
# df["Received"] = pd.to_numeric(df["Received"], errors="coerce").fillna(0).astype(int)
# df["Disposed_Out"] = pd.to_numeric(df["Disposed_Out"], errors="coerce").fillna(0).astype(int)
#
# # ===================== Treemap + Tabular Data =====================
# df_tt = pd.read_excel("selected_columns.xlsx")
# df_tt.columns = df_tt.columns.str.strip()
#
# for col in["District_Eng", "Office_Eng", "Service_Eng", "Disposed_Out", "Disposed", "Received", "Total", "Pending"]:
#     if col not in df_tt.columns:
#         df_tt[col] = 0
# for col in ["Disposed_Out", "Disposed", "Received", "Total", "Pending"]:
#     df_tt[col] = pd.to_numeric(df_tt[col], errors="coerce").fillna(0)
#
# df_tt["Late_Disposed_%"] = (df_tt["Disposed_Out"] / df_tt["Disposed"].replace(0, pd.NA) * 100).fillna(0)
#
# district_order = (
#     df_tt.groupby("District_Eng")["Disposed_Out"]
#     .sum()
#     .sort_values(ascending=False)
#     .index.tolist()
# )
#
# hierarchies = {
#     "District → Office → Service":["District_Eng", "Office_Eng", "Service_Eng"],
#     "District → Service → Office":["District_Eng", "Service_Eng", "Office_Eng"],
#     "Office → District → Service":["Office_Eng", "District_Eng", "Service_Eng"],
#     "Office → Service → District":["Office_Eng", "Service_Eng", "District_Eng"],
#     "Service → District → Office":["Service_Eng", "District_Eng", "Office_Eng"],
#     "Service → Office → District":["Service_Eng", "Office_Eng", "District_Eng"],
# }
#
# def categorize(x):
#     if x < 10: return "A"
#     if x < 40: return "B"
#     if x < 70: return "C"
#     return "D"
#
# def get_district_summary(filtered_df):
#     summary = filtered_df.groupby("District_Eng", as_index=False).agg({
#         "Disposed_Out": "sum",
#         "Disposed": "sum",
#         "Received": "sum"
#     })
#     summary["Late_Disposed_%"] = (summary["Disposed_Out"] / summary["Disposed"].replace(0, pd.NA) * 100).fillna(0)
#     summary["Category"] = summary["Late_Disposed_%"].apply(categorize)
#     summary["Office_Eng"] = ""
#     summary["Service_Eng"] = ""
#     return summary
#
# initial_summary_tt = get_district_summary(df_tt)
#
# # ===================== Monthly Trends Data =====================
# def load_mt_data():
#     df_mt = pd.read_csv("data2.csv", encoding="utf-8-sig", low_memory=False)
#     df_mt.columns = df_mt.columns.str.strip()
#     df_mt['Year'] = pd.to_numeric(df_mt['Year'], errors='coerce')
#     df_mt.dropna(subset=['Year'], inplace=True)
#     df_mt['Year'] = df_mt['Year'].astype(int)
#     df_mt['Month'] = df_mt['Month'].astype(int)
#
#     df_mt['Date'] = pd.to_datetime(
#         df_mt['Year'].astype(str) + '-' + df_mt['Month'].astype(str).str.zfill(2) + '-01',
#         format='%Y-%m-%d'
#     )
#     df_mt['Month_Year'] = df_mt['Date'].dt.strftime('%b-%Y')
#
#     numeric_cols =[
#         'application_Received', 'application_Disposed',
#         'application_Disposed_Out_of_time', 'application_Disposed_with_in_time'
#     ]
#     for col in numeric_cols:
#         df_mt[col] = pd.to_numeric(df_mt[col], errors='coerce').fillna(0).astype(int)
#
#     df_mt['Efficiency_Percentage'] = df_mt.apply(
#         lambda row: (row['application_Disposed_with_in_time'] / row['application_Disposed'] * 100)
#         if row['application_Disposed'] > 0 else 0, axis=1
#     )
#     df_mt['Pending_Applications'] = df_mt['application_Received'] - df_mt['application_Disposed']
#
#     for col in['District_name', 'Service_name', 'Office_name']:
#         df_mt[col] = df_mt[col].astype(str).str.strip()
#
#     return df_mt
#
# df_mt = load_mt_data()
# df_mt_light = df_mt[['Month_Year', 'District_name', 'Service_name', 'Office_name',
#                      'application_Disposed_Out_of_time']].copy()
#
# COLOR_PALETTE = {
#     'primary': '#1f77b4', 'secondary': '#ff7f0e', 'success': '#2ca02c',
#     'danger': '#d62728', 'light': '#f8f9fa', 'dark': '#343a40'
# }
#
# all_months = sorted(df_mt['Month_Year'].unique(), key=lambda x: pd.to_datetime(x, format='%b-%Y'))
# month_options =[{'label': m, 'value': m} for m in all_months]
# month_options.insert(0, {'label': 'All Months', 'value': 'ALL_MONTHS'})


import pandas as pd

# ===================== Metadata Data =====================
df = pd.read_csv("digital-data.csv", encoding="utf-8-sig", low_memory=False)
df.columns = df.columns.str.replace('\ufeff', '', regex=False).str.strip()
for col in["Service_Eng", "Office_Eng", "District_Eng", "Received", "Disposed_Out"]:
    if col not in df.columns:
        df[col] = 0
df["Received"] = pd.to_numeric(df["Received"], errors="coerce").fillna(0).astype(int)
df["Disposed_Out"] = pd.to_numeric(df["Disposed_Out"], errors="coerce").fillna(0).astype(int)

# ===================== Treemap + Tabular Data =====================
df_tt = pd.read_excel("selected_columns.xlsx")
df_tt.columns = df_tt.columns.str.strip()

for col in["District_Eng", "Office_Eng", "Service_Eng", "Disposed_Out", "Disposed", "Received", "Total", "Pending"]:
    if col not in df_tt.columns:
        df_tt[col] = 0
for col in["Disposed_Out", "Disposed", "Received", "Total", "Pending"]:
    df_tt[col] = pd.to_numeric(df_tt[col], errors="coerce").fillna(0)

df_tt["Late_Disposed_%"] = (df_tt["Disposed_Out"] / df_tt["Disposed"].replace(0, pd.NA) * 100).fillna(0)

district_order = (
    df_tt.groupby("District_Eng")["Disposed_Out"]
    .sum()
    .sort_values(ascending=False)
    .index.tolist()
)

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
        "Disposed_Out": "sum",
        "Disposed": "sum",
        "Received": "sum"
    })
    summary["Late_Disposed_%"] = (summary["Disposed_Out"] / summary["Disposed"].replace(0, pd.NA) * 100).fillna(0)
    summary["Category"] = summary["Late_Disposed_%"].apply(categorize)
    summary["Office_Eng"] = ""
    summary["Service_Eng"] = ""
    return summary

initial_summary_tt = get_district_summary(df_tt)

# ===================== Monthly Trends Data =====================
def load_mt_data():
    df_mt = pd.read_csv("data2.csv", encoding="utf-8-sig", low_memory=False)
    df_mt.columns = df_mt.columns.str.strip()
    df_mt['Year'] = pd.to_numeric(df_mt['Year'], errors='coerce')
    df_mt.dropna(subset=['Year'], inplace=True)
    df_mt['Year'] = df_mt['Year'].astype(int)
    df_mt['Month'] = df_mt['Month'].astype(int)

    df_mt['Date'] = pd.to_datetime(
        df_mt['Year'].astype(str) + '-' + df_mt['Month'].astype(str).str.zfill(2) + '-01',
        format='%Y-%m-%d'
    )
    df_mt['Month_Year'] = df_mt['Date'].dt.strftime('%b-%Y')

    numeric_cols =[
        'application_Received', 'application_Disposed',
        'application_Disposed_Out_of_time', 'application_Disposed_with_in_time'
    ]
    for col in numeric_cols:
        df_mt[col] = pd.to_numeric(df_mt[col], errors='coerce').fillna(0).astype(int)

    df_mt['Efficiency_Percentage'] = df_mt.apply(
        lambda row: (row['application_Disposed_with_in_time'] / row['application_Disposed'] * 100)
        if row['application_Disposed'] > 0 else 0, axis=1
    )
    df_mt['Pending_Applications'] = df_mt['application_Received'] - df_mt['application_Disposed']

    for col in ['District_name', 'Service_name', 'Office_name']:
        df_mt[col] = df_mt[col].astype(str).str.strip()

    return df_mt

df_mt = load_mt_data()
df_mt_light = df_mt[['Month_Year', 'District_name', 'Service_name', 'Office_name',
                     'application_Disposed_Out_of_time']].copy()

all_months = sorted(df_mt['Month_Year'].unique(), key=lambda x: pd.to_datetime(x, format='%b-%Y'))
month_options = [{'label': m, 'value': m} for m in all_months]
month_options.insert(0, {'label': 'All Months', 'value': 'ALL_MONTHS'})

COLOR_PALETTE = {
    'primary': '#1f77b4', 'secondary': '#ff7f0e', 'success': '#2ca02c',
    'danger': '#d62728', 'light': '#f8f9fa', 'dark': '#343a40'
}

# ===================== Advanced Analytics Data =====================
try:
    df_adv = pd.read_csv("Merge-digital-data-22-25.csv", encoding="utf-8-sig", low_memory=False)
    # Clean up column names just in case
    df_adv.columns = df_adv.columns.str.strip()
    # Force numerics
    num_cols =['Year', 'Month', 'application_Disposed_Out_of_time', 'application_Disposed']
    for c in num_cols:
        if c in df_adv.columns:
            df_adv[c] = pd.to_numeric(df_adv[c], errors='coerce').fillna(0).astype(int)
except Exception as e:
    print(f"Warning: Failed to load Advanced Analytics data: {e}")
    df_adv = pd.DataFrame(columns=['Year', 'Month', 'Service_name', 'District_name', 'Office_name', 'application_Disposed_Out_of_time', 'application_Disposed'])