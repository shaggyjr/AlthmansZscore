import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk

df1 = None
df2 = None


def load_excel_data1():
    global df1
    file_path = filedialog.askopenfilename(
        filetypes=[('Excel Files', '*.xlsx')])

    if file_path:
        try:
            df1 = pd.read_excel(file_path, sheet_name='Sheet1')
            print("df1 loaded successfully.")
        except Exception as e:
            print("Error loading Excel file:", str(e))
    else:
        print("No file selected.")


def load_excel_data2():
    global df2
    file_path = filedialog.askopenfilename(
        filetypes=[('Excel Files', '*.xlsx')])

    if file_path:
        try:
            df2 = pd.read_excel(file_path, sheet_name='Sheet1')
            print("df2 loaded successfully.")
        except Exception as e:
            print("Error loading Excel file:", str(e))
    else:
        print("No file selected.")



def strsearch(find, df):
    search = df.eq(find).any(axis=1)
    if not search.any():
        print(f"No rows found for '{find}'.")
        return None
    row_index = search[search].index[0]
    value = int(round(float(df.iloc[row_index, 1])))
    return value

def calc_z():
    total_assets =  strsearch("Total Assets", df1)
    current_assets =  strsearch("Total Current Assets", df1)
    total_liabilities = strsearch("Total Capital And Liabilities", df1)
    current_liabilities = strsearch("Total Current Liabilities", df1)
    tot_shareholders_fund = strsearch("Total Shareholders Funds", df1)
    equity_share_capital = strsearch("Total Share Capital", df1)
    reserve_and_surplus = strsearch("Reserves and Surplus", df1)
    minority_interest = strsearch("Minority Interest", df1)
    ebit = strsearch("Profit/Loss Before Tax",df2)
    sales = strsearch("Total Revenue",df2)
    print(total_assets)
    print(current_assets)
    print(total_liabilities)
    print(current_liabilities)
    print(tot_shareholders_fund)
    print(equity_share_capital)
    print(reserve_and_surplus)
    print(minority_interest)
    print(ebit)
    print(sales)

    Working_Capital = current_assets-current_liabilities
    Retained_Earnings = tot_shareholders_fund-equity_share_capital-reserve_and_surplus-minority_interest
    Market_Value_of_Equity = equity_share_capital+reserve_and_surplus+minority_interest
    print(Working_Capital)
    print(Retained_Earnings)
    print(Market_Value_of_Equity)
    A = Working_Capital/total_assets
    B = Retained_Earnings/total_assets
    C = ebit/total_assets
    D = Market_Value_of_Equity/total_liabilities
    E = sales/total_assets

    Z = 1.2* A +1.4* B + 3.3* C +0.6* D +1* E
    print("Z=",Z)

    if Z < 1.81:
        print("Distress Zone")
    elif Z < 2.99:
        print("Grey Zone")
    else:
        print("Safe Zone")



window = tk.Tk()
window.title("Excel File Loader")
window.geometry("400x200")  # Set window dimensions (width x height)
window.resizable(False, False)  # Disable window resizing

# Create a style for the buttons
button_style = ttk.Style()
button_style.configure("TButton", font=("Arial", 12), padding=10)

load_button1 = ttk.Button(window, text="Load Balance Sheet Data",
                          command=load_excel_data1, style="TButton")
load_button1.pack(pady=20)

load_button2 = ttk.Button(
    window, text="Load PL Statement Data", command=load_excel_data2, style="TButton")
load_button2.pack(pady=10)

calculate_button = ttk.Button(
    window, text="Calculate Z-Score", command=calc_z, style="TButton")
calculate_button.pack(pady=10)

window.mainloop()
