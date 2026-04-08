#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
import os
import re

# 👉 Folder path
folder_path = r"C:\Users\Jaya-Suta\OneDrive\Desktop\agritech"

all_data = []

# 👉 Function to clean illegal Excel characters
def clean_illegal_chars(df):
    illegal_char_pattern = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')
    return df.applymap(
        lambda x: illegal_char_pattern.sub('', x) if isinstance(x, str) else x
    )

# 👉 Loop through files
for file in os.listdir(folder_path):

    # ❌ Skip temp/system files
    if file.startswith("~$"):
        continue

    file_path = os.path.join(folder_path, file)

    try:
        # =====================
        # 👉 Excel Files
        # =====================
        if file.lower().endswith((".xlsx", ".xls")):

            excel_file = pd.ExcelFile(file_path)

            for sheet in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet, dtype=str)

                    df['File_Name'] = file
                    df['Sheet_Name'] = sheet

                    all_data.append(df)

                except Exception as e:
                    print(f"❌ Error in sheet: {sheet} | File: {file}")
                    print(e)

        # =====================
        # 👉 CSV Files
        # =====================
        elif file.lower().endswith(".csv"):

            try:
                df = pd.read_csv(
                    file_path,
                    low_memory=False,
                    encoding='utf-8',
                    on_bad_lines='skip'   # 🔥 skip corrupted rows
                )
            except:
                df = pd.read_csv(
                    file_path,
                    low_memory=False,
                    encoding='latin1',
                    on_bad_lines='skip'
                )

            df['File_Name'] = file
            df['Sheet_Name'] = 'CSV'

            all_data.append(df)

    except Exception as e:
        print(f"❌ Error in file: {file}")
        print(e)

# 👉 Merge all data
if all_data:

    final_df = pd.concat(all_data, ignore_index=True)

    # 👉 Clean column names
    final_df.columns = final_df.columns.str.strip().str.upper()

    # 👉 Clean illegal characters (FIXES your Excel error)
    final_df = clean_illegal_chars(final_df)

    # 👉 Output path
    output_path = os.path.join(folder_path, "Final_Merged_File.xlsx")

    # 👉 Save Excel
    final_df.to_excel(output_path, index=False)

    print("✅ Files merged successfully!")
    print("📁 Saved at:", output_path)

else:
    print("⚠️ No data found to merge.")


# In[ ]:




