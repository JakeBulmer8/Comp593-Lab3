""" 
Description: 
  Divides sales data CSV file into individual order data Excel files.

Usage:
  python process_sales_data.py sales_csv_path

Parameters:
  sales_csv_path = Full path of the sales data CSV file
"""
import sys                 # The "sys" package contains parameters and functions that interact with the system.
import os                  # The "os" module contains functions to interact with the operating system.
from datetime import date  # The "datetime" module allows the work with date as date objects in the script.
import pandas as pd        # The "pandas" library is used to work with data sets.
import re                  # The "re" module is imported for the use of regular expressions.

def main():
    sales_csv_path = get_sales_csv_path()
    orders_dir_path = create_orders_dir(sales_csv_path)
    process_sales_data(sales_csv_path, orders_dir_path)

def get_sales_csv_path():
    """Gets the path of sales data CSV file from the command line

    Returns:
        str: Path of sales data CSV file
    """
    # TODO: Check whether command line parameter provided
    num_params = len(sys.argv) - 1                                  # Puts the length of the first argument for the sript in a variable. 
    if num_params < 1:                                              # If the number of parameters is less than 1
        print('Error: Missing path to sales data CV file')          # Print an error message to the screen
        sys.exit(1)                                                 # Exits the script

    # TODO: Check whether provide parameter is valid path of file
    sales_csv_path = sys.argv[1]                                    # Puts the command argument into a variable.
    if not os.path.isfile(sales_csv_path):                          # If the argument is not the path of a file
        print('Error: Invalid path to sales data CSV file')         # Print an error message to the screen
        sys.exit(1)                                                 # Exits the script

    # TODO: Return path of sales data CSV file

    return sales_csv_path

def create_orders_dir(sales_csv_path):
    """Creates the directory to hold the individual order Excel sheets

    Args:
        sales_csv_path (str): Path of sales data CSV file

    Returns:
        str: Path of orders directory
    """
    # TODO: Get directory in which sales data CSV file resides
    sales_dir_path = os.path.dirname(os.path.abspath(sales_csv_path))

    # TODO: Determine the path of the directory to hold the order data files
    todays_date = date.today().isoformat()
    orders_dir_path = os.path.join(sales_dir_path, f'Orders_{todays_date}')

    # TODO: Create the orders directory if it does not already exist
    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)

    # TODO: Return path of orders directory
    return orders_dir_path

def process_sales_data(sales_csv_path, orders_dir_path):
    """Splits the sales data into individual orders and save to Excel sheets

    Args:
        sales_csv_path (str): Path of sales data CSV file
        orders_dir_path (str): Path of orders directory
    """
    # TODO: Import the sales data from the CSV file into a DataFrame
    sales_df =pd.read_csv(sales_csv_path)

    # TODO: Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])

    # TODO: Remove columns from the DataFrame that are not needed
    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)
    # TODO: Groups orders by ID and iterate
    for order_id, order_df in sales_df.groupby('ORDER ID'):

        # TODO: Remove the 'ORDER ID' column
        order_df.drop(columns=['ORDER ID'], inplace=True)

        # TODO: Sort the items by item number
        order_df.sort_values(by='ITEM NUMBER', inplace=True)

        # TODO: Append a "GRAND TOTAL" row
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL:'], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_df])

        # TODO: Determine the file name and full path of the Excel sheet
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W', '', customer_name) 
        order_file = f'Order{order_id}_{customer_name}.xlsx'
        order_path = os.path.join(orders_dir_path, order_file)

        # TODO: Export the data to an Excel sheet
        sheet_name = f'Order #{order_id}'
        order_df.to_excel(order_path, index=False, sheet_name=sheet_name)

        # TODO: Format the Excel sheet
        writer = pd.ExcelWriter(order_path, engine='xlsxwriter')
        order_df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name] 

        # TODO: Define format for the money columns
        money_fmt = workbook.add_format({'num_format': '$#,##.00'})

        # TODO: Format each colunm
        worksheet.set_column('A:A', 11)
        worksheet.set_column('B:B', 13)
        worksheet.set_column('C:E', 15)
        worksheet.set_column('F:G', 13, money_fmt)
        worksheet.set_column('H:H', 10)
        worksheet.set_column('I:I', 30)

        # TODO: Close the Excelwriter 
        writer.close()
    return

if __name__ == '__main__':
    main()