import pandas as pd

# Specify the file path
file_path = r"C:\Users\kumsam\Desktop\Excel-Automation\file\Large_Excel_Automation_Practice.xlsx"

# Step 1: Load the data
try:
    sales_data = pd.read_excel(file_path, sheet_name="Sales Data")
    print("Data loaded successfully!")
except Exception as e:
    print(f"Error loading data: {e}")

# Step 2: Check the data
print("Shape of data:", sales_data.shape)
print("First few rows of data:\n", sales_data.head())

# Step 3: Check column names
print("Columns in the dataset:", sales_data.columns)

# Step 4: Check for missing values
print("Missing values:\n", sales_data.isnull().sum())

# Step 5: Drop rows with missing 'Quantity' or 'Price'
sales_data = sales_data.dropna(subset=['Quantity', 'Price'])

# Step 6: Convert 'Quantity' and 'Price' to numeric
sales_data['Quantity'] = pd.to_numeric(sales_data['Quantity'], errors='coerce')
sales_data['Price'] = pd.to_numeric(sales_data['Price'], errors='coerce')

# Step 7: Check for missing values after conversion
print("Missing values after conversion:\n", sales_data.isnull().sum())

# Step 8: Add 'Total Revenue' column
sales_data['Total Revenue'] = sales_data['Quantity'] * sales_data['Price']
print("First few rows after adding 'Total Revenue':\n", sales_data.head())

# Step 9: Summarize revenue by product
revenue_by_product = sales_data.groupby('Product')['Total Revenue'].sum()

# Step 10: Display the result
print("Revenue by Product:\n", revenue_by_product)
