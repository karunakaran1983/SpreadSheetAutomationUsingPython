import openpyxl

inv_file = openpyxl.load_workbook("inventory.xlsx")
product_list = inv_file["Sheet1"]

# Task 1: List each company with respective product count
# Expected Exercise Result: { 'AAA Company': 43, 'BBB Company': 17, 'CCC Company': 14 }

# print(product_list.max_row) # 75

products_per_supplier = {}

for product_row in range(2, product_list.max_row + 1):
    supplier_name = product_list.cell(product_row, 4).value

    if supplier_name in products_per_supplier:
        current_num_products = products_per_supplier[supplier_name]
        products_per_supplier[supplier_name] = current_num_products + 1
    else:
        print("Adding new supplier", supplier_name)
        products_per_supplier[supplier_name] = 1

print(products_per_supplier)

# Task 2: List products with inventory less than 10

# Task 3: List each company with respective total inventory value

# Task 4: Write to Spreadsheet: Calculate and write inventory value for each product into spreadsheet