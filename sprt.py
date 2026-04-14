import pandas as pd

# Load the Common names sheet
df_common = pd.read_excel('compounds.xlsx', sheet_name='Common names', header=0)

# Load data.csv - single column, one name per row, possibly ending with comma
with open('data.csv', 'r') as f:
    content = f.read().splitlines()
    # Clean each line: strip whitespace and remove trailing commas
    gene_names = {line.strip().rstrip(',').strip() for line in content if line.strip()}

# Prepare separate lists for present and not present
present_output = []
not_present_output = []

# Iterate over each compound (column)
for compound in df_common.columns:
    # Get non-null values in the column
    values = df_common[compound].dropna().tolist()
    for val in values:
        # Split if multiple names separated by space
        names = str(val).split()
        for name in names:
            present = name in gene_names
            entry = {'Compound': compound, 'Common Name': name}
            
            if present:
                present_output.append(entry)
            else:
                not_present_output.append(entry)

# Create DataFrames
present_df = pd.DataFrame(present_output)
not_present_df = pd.DataFrame(not_present_output)

# Create Excel writer object
with pd.ExcelWriter('compound_matches.xlsx', engine='openpyxl') as writer:
    # Save present matches to Sheet 1
    if not present_df.empty:
        present_df.to_excel(writer, sheet_name='PRESENT Matches', index=False)
    else:
        pd.DataFrame(columns=['Compound', 'Common Name']).to_excel(writer, sheet_name='PRESENT Matches', index=False)
    
    # Save not present matches to Sheet 2
    if not not_present_df.empty:
        not_present_df.to_excel(writer, sheet_name='NOT PRESENT', index=False)
    else:
        pd.DataFrame(columns=['Compound', 'Common Name']).to_excel(writer, sheet_name='NOT PRESENT', index=False)
    
    # Create Summary sheet
    summary_data = {
        'Metric': ['Total Compounds Processed', 'Total Common Names Checked', 
                  'Names PRESENT in data.csv', 'Names NOT PRESENT in data.csv'],
        'Count': [len(df_common.columns), len(present_output + not_present_output),
                 len(present_output), len(not_present_df)]
    }
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

print("=" * 60)
print("COMPOUND-GENE MATCHING ANALYSIS COMPLETED!")
print("=" * 60)
print(f"📊 Total Compounds: {len(df_common.columns)}")
print(f"🔍 Total Names Checked: {len(present_output + not_present_output)}")
print(f"✅ PRESENT Matches: {len(present_output)}")
print(f"❌ NOT PRESENT: {len(not_present_df)}")
print(f"📈 Match Rate: {len(present_output)/(len(present_output + not_present_output)*100):.1f}%" if (present_output + not_present_output) else "📈 Match Rate: 0%")
print("=" * 60)
print("📁 Results saved to 'compound_matches.xlsx' with 3 sheets:")
print("   1. PRESENT Matches - All genes found in data.csv")
print("   2. NOT PRESENT - All genes NOT found in data.csv") 
print("   3. Summary - Overall statistics")
print("=" * 60)

# Show sample data if available
if not present_df.empty:
    print("\n📋 SAMPLE - PRESENT MATCHES:")
    print(present_df.head(3).to_string(index=False))
    print("-" * 40)

if not not_present_df.empty:
    print("\n📋 SAMPLE - NOT PRESENT:")
    print(not_present_df.head(3).to_string(index=False))
    print("-" * 40)
