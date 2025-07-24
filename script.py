import pandas as pd
import os
from sklearn.cluster import KMeans
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# Define paths
DATA_DIR = os.path.join(os.path.dirname(__file__), '..', 'data')
REPORT_DIR = os.path.join(os.path.dirname(__file__), '..', 'reports')

def load_and_prepare_data():
    # Load CSVs
    customers = pd.read_csv(os.path.join(DATA_DIR, 'customers.csv'))
    orders = pd.read_csv(os.path.join(DATA_DIR, 'orders.csv'))
    products = pd.read_csv(os.path.join(DATA_DIR, 'products.csv'))

    # Merge data
    merged = orders.merge(customers, on='customer_id').merge(products, on='product_id')

    # Add total price
    merged['total_price'] = merged['quantity'] * merged['unit_price']
    return merged

def apply_ai_analysis(data):
    # Example AI logic: cluster customers by total_price
    kmeans = KMeans(n_clusters=3, random_state=42)
    data['cluster'] = kmeans.fit_predict(data[['total_price']])
    return data

def create_excel_report(data):
    if not os.path.exists(REPORT_DIR):
        os.makedirs(REPORT_DIR)

    filepath = os.path.join(REPORT_DIR, 'sales_report.xlsx')

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Report"

    # Add header
    for r in dataframe_to_rows(data, index=False, header=True):
        ws.append(r)

    # Format header
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    wb.save(filepath)
    return filepath
