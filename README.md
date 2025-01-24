# Excel Automation with Python

This repository contains examples of Excel automation using Python. The examples demonstrate various operations such as data analysis, filtering, sorting, visualizations, conditional formatting, and creating pivot tables. The data used for these examples is in the `Large_Excel_Automation_Practice.xlsx` file.

## Prerequisites

Before running the examples, ensure you have the following installed:

- Python 3.7 or later
- Required Python libraries:
  - `pandas`
  - `openpyxl`
  - `matplotlib`
  - `numpy`

Install the dependencies using pip:

```bash
pip install pandas openpyxl matplotlib numpy
```

## File Overview

### `Large_Excel_Automation_Practice.xlsx`
This Excel file contains three sheets with large datasets:

- **Sales Data**: Includes product sales with details like date, product, quantity, and price.
- **Employee Data**: Contains employee information, including ID, name, department, and salary.
- **Inventory**: Lists inventory items with stock levels and locations.

## Examples

### 1. Sales Data: Data Analysis
**Task**: Calculate total revenue and summarize revenue by product.

```python
# Add a Total Revenue column
sales_data['Total Revenue'] = sales_data['Quantity'] * sales_data['Price']

# Summarize revenue by product
revenue_by_product = sales_data.groupby('Product')['Total Revenue'].sum()
print(revenue_by_product)
```

### 2. Employee Data: Filtering and Sorting
**Task**: Sort employees by salary and filter employees in the IT department.

```python
# Sort by Salary
sorted_employees = employee_data.sort_values(by="Salary", ascending=False)

# Filter employees in the IT department
it_employees = employee_data[employee_data['Department'] == 'IT']
print(it_employees.head())
```

### 3. Inventory Data: Stock Management
**Task**: Identify items with stock below a threshold.

```python
low_stock_items = inventory_data[inventory_data['Stock'] < 50]
print(low_stock_items)
```

### 4. Automating Reports
**Task**: Generate a summary report combining insights from all sheets.

```python
# Summary report example
summary = pd.DataFrame({
    "Metric": ["Total Sales", "Average Salary", "Total Stock"],
    "Value": [total_sales, average_salary, total_stock]
})
summary.to_excel(writer, sheet_name="Summary", index=False)
```

### 5. Visualizations
**Task**: Create a bar chart for product sales.

```python
# Plot a bar chart
sales_by_product.plot(kind='bar', title="Product Sales", color='skyblue')
plt.xlabel("Product")
plt.ylabel("Quantity Sold")
plt.show()
```

### 6. Conditional Formatting
**Task**: Highlight rows in the inventory with stock below 50.

```python
low_stock_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    if row[1].value < 50:
        for cell in row:
            cell.fill = low_stock_fill
```

### 7. Pivot Tables
**Task**: Analyze sales data by product and month.

```python
sales_data['Month'] = pd.to_datetime(sales_data['Date']).dt.to_period('M')
pivot_table = sales_data.pivot_table(
    index='Product', columns='Month', values='Total Revenue', aggfunc='sum', fill_value=0
)
print(pivot_table)
```

## How to Use

1. Clone the repository:

```bash
git clone git@github.com:kumsa-Mergia/Excel-Automation.git
cd Excel-Automation
```

2. Place the `Large_Excel_Automation_Practice.xlsx` file in the root directory.

3. Run any example script to see the output.

4. Modify the examples or add your own scripts to experiment further.


