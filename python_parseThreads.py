import pandas as pd
import re

# --- CONFIG ---
input_file = "ParentList.xlsx"     # your Excel file
input_sheet = "Sheet1"             # sheet containing full Thread IDs in column A
output_file = "Thread_Permutations.xlsx"

# --- LOAD DATA ---
df = pd.read_excel(input_file, sheet_name=input_sheet, header=None)
df.columns = ["FullThreadID"]

all_permutations = []

# --- GENERATE PERMUTATIONS ---
for full_id in df["FullThreadID"]:
    if pd.isna(full_id):
        continue
    
    # remove trailing whitespace/newlines
    full_id = full_id.strip()
    
    # find all positions of '+' and '-'
    cut_points = [m.start() for m in re.finditer(r"[+-]", full_id)]
    
    # go backwards, slice up to each cut point (inclusive)
    for i in range(len(cut_points) - 1, -1, -1):
        truncated = full_id[:cut_points[i] + 1]
        all_permutations.append({
            "OriginalThreadID": full_id,
            "Permutation": truncated
        })

# --- SAVE OUTPUT ---
out_df = pd.DataFrame(all_permutations)
out_df.to_excel(output_file, index=False)

print(f"✅ Generated {len(out_df)} permutations across {len(df)} base ThreadIDs.")
print(f"Saved to: {output_file}")
