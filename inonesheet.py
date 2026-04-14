import pandas as pd

# Load the Common names sheet
df_common = pd.read_excel('compounds.xlsx', sheet_name='Common names', header=0)

# Load data.csv - single column, one name per row, possibly ending with comma
with open('data.csv', 'r') as f:
    content = f.read().splitlines()
    # Clean each line: strip whitespace and remove trailing commas
    gene_names = {line.strip().rstrip(',').strip() for line in content if line.strip()}

# Prepare output list
output = []

# Iterate over each compound (column)
for compound in df_common.columns:
    # Get non-null values in the column
    values = df_common[compound].dropna().tolist()
    for val in values:
        # Split if multiple names separated by space
        names = str(val).split()
        for name in names:
            present = 'Yes' if name in gene_names else 'No'
            output.append({'Compound': compound, 'Common Name': name, 'Present': present})

# Create DataFrame
output_df = pd.DataFrame(output)

# Save as Excel file (.xlsx)
output_df.to_excel('matches.xlsx', index=False, sheet_name='Gene Matches')

print(f"Matching completed! Found {len(output)} entries.")
print("Results saved to 'matches.xlsx' - Ready to open in Excel!")
print("\nSummary:")
print(f"Total compounds processed: {len(df_common.columns)}")
print(f"Total common names checked: {len(output)}")
print(f"Present in data.csv: {len([x for x in output if x['Present'] == 'Yes'])}")
print(f"Not present in data.csv: {len([x for x in output if x['Present'] == 'No'])}")
print("\nSample of results:")
print(output_df.head(10))
