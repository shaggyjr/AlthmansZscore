import streamlit as st
import pandas as pd

def load_excel_data1():
    file_path = st.file_uploader("Load Balance Sheet Data", type=['xlsx'])
    if file_path:
        try:
            df1 = pd.read_excel(file_path, sheet_name='Sheet1')
            st.success("Balance Sheet Data loaded successfully.")
            return df1
        except Exception as e:
            st.error("Error loading Balance Sheet Data: " + str(e))
    else:
        st.warning("No file selected.")
        return None

def load_excel_data2():
    file_path = st.file_uploader("Load PL Statement Data", type=['xlsx'])
    if file_path:
        try:
            df2 = pd.read_excel(file_path, sheet_name='Sheet1')
            st.success("PL Statement Data loaded successfully.")
            return df2
        except Exception as e:
            st.error("Error loading PL Statement Data: " + str(e))
    else:
        st.warning("No file selected.")
        return None

def calculate_z_score(df1, df2):
    if df1 is None or df2 is None:
        st.warning("Data not loaded.")
        return

    total_assets = strsearch("Total Assets", df1)
    current_assets = strsearch("Total Current Assets", df1)
    total_liabilities = strsearch("Total Capital And Liabilities", df1)
    current_liabilities = strsearch("Total Current Liabilities", df1)
    tot_shareholders_fund = strsearch("Total Shareholders Funds", df1)
    equity_share_capital = strsearch("Total Share Capital", df1)
    reserve_and_surplus = strsearch("Reserves and Surplus", df1)
    minority_interest = strsearch("Minority Interest", df1)
    ebit = strsearch("Profit/Loss Before Tax", df2)
    sales = strsearch("Total Revenue", df2)

    Working_Capital = current_assets - current_liabilities
    Retained_Earnings = tot_shareholders_fund - equity_share_capital - reserve_and_surplus - minority_interest
    Market_Value_of_Equity = equity_share_capital + reserve_and_surplus + minority_interest
    A = Working_Capital / total_assets
    B = Retained_Earnings / total_assets
    C = ebit / total_assets
    D = Market_Value_of_Equity / total_liabilities
    E = sales / total_assets

    Z = 1.2 * A + 1.4 * B + 3.3 * C + 0.6 * D + 1 * E

    return Z

def strsearch(find, df):
    search = df.eq(find).any(axis=1)
    if not search.any():
        return None
    row_index = search[search].index[0]
    value = int(round(float(df.iloc[row_index, 1])))
    return value

# Main Page
st.set_page_config(page_title="Altman's Z-Score Analyzer", layout="wide")
st.title("Altman's Z-Score Analyzer")
st.header("Excel File Loader")

df1 = load_excel_data1()
df2 = load_excel_data2()

if df1 is not None and df2 is not None:
    Z = calculate_z_score(df1, df2)
    st.subheader("Z-Score Calculation")
    st.markdown(f"Z-Score: **{Z:.2f}**")

    if Z < 1.81:
        st.markdown("Zone: **Distress Zone**", unsafe_allow_html=True)
    elif Z < 2.99:
        st.markdown("Zone: **Grey Zone**", unsafe_allow_html=True)
    else:
        st.markdown("Zone: **Safe Zone**", unsafe_allow_html=True)
