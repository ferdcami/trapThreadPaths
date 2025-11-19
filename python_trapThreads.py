import pandas as pd

# --- CONFIG ---
permutations_file = "Thread_Permutations.xlsx"   # Output file from your first script
parentlist_file = "ParentList.xlsx"              # Your master workbook
sheet_name = "All Docs"                          # Sheet with Email Threading IDs
output_file = "ParentList_Flagged.xlsx"          # New file to be created

# --- LOAD DATA ---
print("Loading data...")

# Read permutations (only need the Permutation column)
perms = pd.read_excel(permutations_file, usecols=["Permutation"])
perms_set = set(perms["Permutation"].dropna().astype(str))

# Read All Docs sheet
docs = pd.read_excel(parentlist_file, sheet_name=sheet_name)

# --- FLAG MATCHES ---
print("Flagging matches...")
docs["Flag"] = docs["Email Threading ID"].astype(str).apply(lambda x: "Yes" if x in perms_set else "")

# --- SAVE OUTPUT ---
docs.to_excel(output_file, index=False)
print(f"✅ Done! Matches flagged and saved to {output_file}")
