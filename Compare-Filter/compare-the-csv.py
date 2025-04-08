import pandas as pd

# Load the CSV files (adjust encoding if needed)
vm_df = pd.read_csv('vm.csv', encoding='utf-8')
inventory_df = pd.read_csv('all_inventory_list.csv', encoding='ISO-8859-1')  # Try 'cp1252' if this fails

# Print column names for reference (optional)
# print("VM Columns:", vm_df.columns)
# print("Inventory Columns:", inventory_df.columns)

# Replace 'VM_Name' with the actual column name that contains machine names
vm_column = 'VM_Name'
inventory_column = 'VM_Name'

# Clean up whitespace and ensure strings
vm_machines = vm_df[vm_column].astype(str).str.strip()
inventory_machines = inventory_df[inventory_column].astype(str).str.strip()

# Filter rows from vm_df that are not in inventory
missing_machines = vm_df[~vm_machines.isin(inventory_machines)]

# Save the result to a CSV file
missing_machines.to_csv('not_in_inventory.csv', index=False)

print("✅ Comparison complete! Results saved to not_in_inventory.csv")
