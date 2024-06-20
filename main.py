import json
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def load_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

def filter_group(data, group_key):
    return [item for item in data if item.get('type', '').startswith(group_key)]

def expand_env_column(df):
    env_dicts = []
    
    for index, row in df.iterrows():
        env_entries = row['env']
        env_dict = {entry['name']: entry['value'] for entry in env_entries}
        env_dicts.append(env_dict)
    
    env_data = pd.DataFrame(env_dicts)
    df = pd.concat([df.drop(columns=['env']), env_data], axis=1)
    return df

def json_to_excel(json_file, group_key, excel_file):
    data = load_json(json_file)
    filtered_data = filter_group(data, group_key)
    df = pd.DataFrame(filtered_data)
    
    columns_to_drop = ['info', 'category', 'in', 'out', 'color', 'icon', 'meta', 'status', 'z', 'x', 'y', 'wires', 'd']
    df = df.drop(columns=columns_to_drop, errors='ignore')
    
    df = expand_env_column(df)
    
    df.to_excel(excel_file, index=False, engine='openpyxl')

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
    if file_path:
        process_file(file_path)

def process_file(file_path):
    try:
        group_key = 'subflow'
        excel_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if excel_file:
            json_to_excel(file_path, group_key, excel_file)
            messagebox.showinfo("Success", "Excel file created successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("JSON to Excel Converter")

frame = tk.Frame(root, padx=40, pady=20)
frame.pack(padx=10, pady=10)

btn_select = tk.Button(frame, text="Select JSON File", command=select_file)
btn_select.pack()

root.mainloop()

