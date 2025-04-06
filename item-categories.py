import pandas as pd

# ===========================
# This script is useful when you are provided an item spreadsheet with three columns of item categories. It will generate three sheets that you import one-at-a-time into Odoo to have the correet parent-child relationships between the categories.
# ===========================
# ===========================
# CONFIGURATION
# ===========================
INPUT_FILE = "your_input_file.xlsx"  # Replace with your actual input filename
OUTPUT_FILE = "odoo_categories_output.xlsx"  # Output Excel file
TOP_LEVEL_PARENT = "TPI / Armstrong"  # Change this as needed

# These must match your column names
TIER1_COL = 'Category Tier 1'
TIER2_COL = 'Category Tier 2'
TIER3_COL = 'Category Tier 3'


# ===========================
# CATEGORY BUILDING
# ===========================
df_variants = pd.read_excel(INPUT_FILE)
rows = []

for _, row in df_variants.iterrows():
    t1 = row[TIER1_COL]
    t2 = row[TIER2_COL]
    t3 = row[TIER3_COL]

    if pd.notna(t1):
        rows.append({'Category': t1, 'Parent': TOP_LEVEL_PARENT})
    if pd.notna(t2):
        rows.append({'Category': t2, 'Parent': f'{TOP_LEVEL_PARENT} / {t1}'})
    if pd.notna(t3):
        rows.append({'Category': t3, 'Parent': f'{TOP_LEVEL_PARENT} / {t1} / {t2}'})

# Add top-level manually
rows.append({'Category': TOP_LEVEL_PARENT, 'Parent': ''})

# Deduplicate and build dataframe
df_flat = pd.DataFrame(rows).drop_duplicates()

# Tier detection
tier_0 = df_flat[df_flat['Category'] == TOP_LEVEL_PARENT]
tier_1 = df_flat[df_flat['Parent'] == TOP_LEVEL_PARENT]
tier_2 = df_flat[df_flat['Parent'].str.startswith(f'{TOP_LEVEL_PARENT} /') & ~df_flat['Parent'].str.contains(' / ', len(TOP_LEVEL_PARENT) + 3)]
tier_3 = df_flat[~df_flat.isin(tier_0).all(axis=1) & ~df_flat.isin(tier_1).all(axis=1) & ~df_flat.isin(tier_2).all(axis=1)]

# ===========================
# SAVE TO EXCEL
# ===========================
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    tier_0.to_excel(writer, sheet_name="Tier 0", index=False)
    tier_1.to_excel(writer, sheet_name="Tier 1", index=False)
    tier_2.to_excel(writer, sheet_name="Tier 2", index=False)
    tier_3.to_excel(writer, sheet_name="Tier 3", index=False)

print(f"✅ Categories written to: {OUTPUT_FILE}")
