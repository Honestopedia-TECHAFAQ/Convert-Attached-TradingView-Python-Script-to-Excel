import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

def calculate_vwap(vwap_series, length):
    return vwap_series.rolling(window=length).mean()

def calculate_rsi(close_series, length):
    delta = close_series.diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    avg_gain = gain.rolling(window=length).mean()
    avg_loss = loss.rolling(window=length).mean()
    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))
    return rsi.fillna(0)

def fetch_columns(file_path):
    try:
        data = pd.read_csv(file_path)
        columns = data.columns.tolist()
        return columns
    except Exception as e:
        messagebox.showerror("Error", f"Error reading CSV file: {str(e)}")
        return []

def upload_csv():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        columns = fetch_columns(file_path)
        
        if not columns:
            messagebox.showerror("Error", "No columns found or error reading the CSV file.")
            return
        
        vwap_col = simpledialog.askstring("Input", f"Available columns: {columns}\nEnter the column name for VWAP:")
        close_col = simpledialog.askstring("Input", f"Available columns: {columns}\nEnter the column name for Close:")
        
        if vwap_col not in columns or close_col not in columns:
            messagebox.showerror("Error", "Selected columns are not in the CSV file.")
            return
        
        try:
            data = pd.read_csv(file_path)
            vwap = data[vwap_col]
            close = data[close_col]
            vwap_values = [calculate_vwap(vwap, length) for length in len_values]
            rsi_values = calculate_rsi(close, length_rsi)
            output_df = pd.DataFrame({
                'VWAP_len5': vwap_values[0],
                'VWAP_len10': vwap_values[1],
                'VWAP_len15': vwap_values[2],
                'VWAP_len20': vwap_values[3],
                'VWAP_len25': vwap_values[4],
                'RSI_len21': rsi_values
            })
            excel_file = 'output.xlsx'
            with pd.ExcelWriter(excel_file) as writer:
                data.to_excel(writer, sheet_name='Sheet1', index=False) 
                output_df.to_excel(writer, sheet_name='Sheet1', startcol=6, index=False) 

            result_label.config(text="Calculations complete. Check output.xlsx.")
            messagebox.showinfo("Success", f'Excel file "{excel_file}" created successfully.')

        except Exception as e:
            messagebox.showerror("Error", f"Error occurred: {str(e)}")

root = tk.Tk()
root.title("VWAP-RSI Calculator")

len_values = [5, 10, 15, 20, 25]  
length_rsi = 21 

upload_button = tk.Button(root, text="Upload CSV File", command=upload_csv)
upload_button.pack(pady=20)

result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()
