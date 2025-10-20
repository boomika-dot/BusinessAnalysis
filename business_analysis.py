"""
Comprehensive Business Intelligence Analysis
Multi-Table Data Analysis using Pandas and Matplotlib

This program performs enterprise-level analysis across multiple datasets:
- Customer Demographics and Segmentation
- Product Catalog and Profitability
- Regional Sales Performance
- Automation ROI Analysis
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Configure visualization settings
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("Set2")
plt.rcParams['figure.figsize'] = (16, 10)
plt.rcParams['font.size'] = 10

def print_section_header(title):
    """Print formatted section header"""
    print("\n" + "=" * 80)
    print(f"  {title}")
    print("=" * 80)

def print_subsection(title):
    """Print formatted subsection"""
    print(f"\n{title}")
    print("-" * 80)

def load_all_datasets(file_path):
    """Load all sheets from the Excel file and normalize names"""
    print_section_header("LOADING MULTI-TABLE DATASET")
    
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(file_path)
        print(f"\nExcel file loaded successfully: {file_path}")
        print(f"Available sheets: {', '.join(excel_file.sheet_names)}")
        
        # Load and clean each sheet
        sheets = {}
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # Clean column names (remove spaces, lowercase)
            df.columns = df.columns.str.strip().str.lower()
            sheets[sheet_name.strip().lower()] = df
            print(f"   - {sheet_name}: {df.shape[0]} rows, {df.shape[1]} columns")
        
        return sheets

    except FileNotFoundError:
        print(f"\nFile not found: {file_path}")
        print("Creating sample datasets for demonstration...")
        return create_sample_datasets()

def create_sample_datasets():
    """Create sample datasets matching the provided structure"""
    
    # Customer Profiles
    customers = pd.DataFrame({
        'customer_id': ['C1000', 'C1001', 'C1002', 'C1003', 'C1004', 'C1005', 'C1006', 'C1007', 
                       'C1008', 'C1009', 'C1010', 'C1011', 'C1012', 'C1013', 'C1014', 'C1015', 'C1016', 'C1017'],
        'first_name': ['Arjun', 'Neha', 'Rahul', 'Ayesha', 'Vikram', 'Ishita', 'Rohan', 'Priyanka',
                      'Karthik', 'Meera', 'Ananya', 'Siddharth', 'Pooja', 'Aman', 'Tanvi', 'Dev', 'Ria', 'Nikhil'],
        'last_name': ['Mehta', 'Sharma', 'Iyer', 'Patel', 'Rao', 'Khan', 'Gupta', 'Verma',
                     'Nair', 'Das', 'Mukherjee', 'Chopra', 'Bose', 'Kulkarni', 'Yadav', 'Singh', 'Agarwal', 'Bhatt'],
        'age': [41, 45, 26, 28, 22, 26, 41, 53, 32, 50, 55, 49, 31, 40, 38, 34, 37, 30],
        'city': ['Bengaluru', 'Mumbai', 'New Delhi', 'Chennai', 'Hyderabad', 'Pune', 'Ahmedabad', 'Kolkata',
                'Jaipur', 'Lucknow', 'Kochi', 'Coimbatore', 'Surat', 'Indore', 'Patna', 'Chandigarh', 'Guwahati', 'Nagpur'],
        'state': ['Karnataka', 'Maharashtra', 'Delhi', 'Tamil Nadu', 'Telangana', 'Maharashtra', 'Gujarat', 'West Bengal',
                 'Rajasthan', 'Uttar Pradesh', 'Kerala', 'Tamil Nadu', 'Gujarat', 'Madhya Pradesh', 'Bihar', 'Chandigarh', 'Assam', 'Maharashtra'],
        'customer_segment': ['Premium', 'Standard', 'Standard', 'Basic', 'Basic', 'Basic', 'Basic', 'Standard',
                            'Standard', 'Standard', 'Standard', 'Standard', 'Standard', 'Premium', 'Standard', 'Standard', 'Standard', 'Premium'],
        'signup_date': pd.to_datetime(['2023-07-23', '2023-10-26', '2023-12-11', '2023-10-01', '2023-10-06', 
                                       '2023-07-15', '2023-09-04', '2023-06-17', '2023-10-08', '2023-06-08',
                                       '2023-11-05', '2023-09-02', '2023-07-26', '2023-08-20', '2023-09-12',
                                       '2023-09-10', '2023-12-21', '2023-06-05']),
        'total_purchases': [4, 15, 9, 12, 24, 7, 9, 23, 24, 24, 4, 20, 6, 19, 16, 10, 22, 12]
    })
    
    # Product Catalog
    products = pd.DataFrame({
        'product_id': ['P101', 'P102', 'P103', 'P201', 'P202', 'P203', 'P301', 'P302', 
                      'P303', 'P401', 'P402', 'P403', 'P501', 'P502', 'P503'],
        'product_name': ['Bluetooth Earbuds X1', 'Smartwatch Pulse 2', 'USB-C Charger 25W',
                        'Cotton T-Shirt (M)', 'Formal Shirt (L)', 'Denim Jeans (32)',
                        'Data Automation 101', 'Python for Analysts', 'Excel Power Tips',
                        'LED Desk Lamp', 'Steel Water Bottle', 'Memory Foam Pillow',
                        'Cricket Bat Poplar', 'Yoga Mat Pro 6mm', 'Skipping Rope Speed'],
        'category': ['Electronics', 'Electronics', 'Electronics', 'Clothing', 'Clothing', 'Clothing',
                    'Books', 'Books', 'Books', 'Home', 'Home', 'Home', 'Sports', 'Sports', 'Sports'],
        'supplier': ['Shree Traders', 'OmniSource Pvt Ltd', 'Kiran Distributors', 'Zenith Imports',
                    'Zenith Imports', 'Prakash Wholesales', 'OmniSource Pvt Ltd', 'Kiran Distributors',
                    'Shree Traders', 'Prakash Wholesales', 'Shree Traders', 'OmniSource Pvt Ltd',
                    'Kiran Distributors', 'Zenith Imports', 'Prakash Wholesales'],
        'cost_price': [1499, 2499, 399, 199, 499, 899, 220, 350, 180, 650, 220, 700, 1200, 280, 90],
        'selling_price': [2199, 3499, 699, 449, 899, 1499, 499, 799, 399, 999, 449, 1299, 1999, 699, 199],
        'stock_quantity': [120, 35, 25, 80, 40, 22, 42, 30, 200, 22, 18, 50, 15, 22, 300],
        'reorder_level': [30, 20, 15, 25, 20, 12, 10, 15, 40, 18, 12, 20, 10, 8, 50],
        'is_active': [True] * 15
    })
    
    # Regional Performance
    regional = pd.DataFrame({
        'region': ['North', 'North', 'North', 'North', 'South', 'South', 'South', 'South',
                  'East', 'East', 'East', 'East', 'West', 'West', 'West', 'West'],
        'month': ['Jan', 'Feb', 'Mar', 'Apr', 'Jan', 'Feb', 'Mar', 'Apr',
                 'Jan', 'Feb', 'Mar', 'Apr', 'Jan', 'Feb', 'Mar', 'Apr'],
        'year': [2024] * 16,
        'total_sales': [297316, 391081, 570960, 513648, 263475, 277365, 375968, 451234,
                       206480, 237526, 262539, 240264, 689920, 399990, 925204, 580608],
        'total_orders': [311, 263, 234, 348, 225, 165, 379, 193, 145, 226, 279, 142, 320, 398, 382, 256],
        'avg_order_value': [956, 1487, 2440, 1476, 1171, 1681, 992, 2338, 1424, 1051, 941, 1692, 2156, 1005, 2422, 2268],
        'top_category': ['Books', 'Books', 'Home', 'Books', 'Home', 'Sports', 'Sports', 'Electronics',
                        'Sports', 'Home', 'Clothing', 'Books', 'Sports', 'Sports', 'Clothing', 'Clothing'],
        'sales_target': [283081, 389135, 588989, 473943, 284402, 285843, 347938, 457578,
                        200446, 225086, 276784, 225932, 656323, 406286, 855176, 551825],
        'target_achieved': ['Yes', 'Yes', 'No', 'Yes', 'No', 'No', 'Yes', 'No',
                           'Yes', 'Yes', 'No', 'Yes', 'Yes', 'No', 'Yes', 'Yes']
    })
    
    # Automation Scenarios
    automation = pd.DataFrame({
        'scenario_id': ['S001', 'S002', 'S003', 'S004', 'S005', 'S006', 'S007', 'S008', 'S009', 'S010', 'S011', 'S012'],
        'business_process': ['Invoice processing', 'Inventory updates', 'Customer onboarding', 'Sales reporting',
                            'Vendor reconciliation', 'Refund processing', 'Lead qualification', 'Compliance reporting',
                            'Payroll prep', 'Data backups', 'Ticket triage', 'MIS dashboard refresh'],
        'current_method': ['Manual in Tally exports', 'Spreadsheet copy-paste', 'Email + Google Form', 
                          'Weekly Excel pivot prep', 'Manual PO-GRN match', 'Email-based approvals',
                          'Manual CRM tagging', 'Monthly PDF assembly', 'Timesheet collation',
                          'Scheduled cloud backup', 'Helpdesk rules + tags', 'Automated data pulls'],
        'time_per_task_minutes': [25, 15, 30, 45, 35, 20, 10, 60, 50, 15, 8, 35],
        'frequency_per_week': [20, 28, 10, 6, 8, 7, 40, 4, 5, 14, 50, 9],
        'automation_potential': ['High', 'High', 'Medium', 'High', 'High', 'Medium', 'High', 'Medium', 
                                'Medium', 'Low', 'High', 'High'],
        'roi_estimate': [180, 220, 140, 200, 160, 120, 250, 130, 110, 90, 160, 190],
        'annual_cost': [260000, 218400, 156000, 140400, 145600, 72800, 208000, 124800, 130000, 109200, 208000, 163800]
    })
    
    print("\nSample datasets created successfully for demonstration.")
    
    return {
        'Customer Profiles': customers,
        'Product Catalog': products,
        'Regional Performance': regional,
        'Automation Scenarios': automation
    }

def analyze_customer_data(customers):
    """Analyze customer demographics and segmentation"""
    print_section_header("CUSTOMER ANALYSIS")
    
    print_subsection("1. Customer Demographics")
    print(f"Total Customers: {len(customers)}")
    print(f"Age Range: {customers['age'].min()} - {customers['age'].max()} years")
    print(f"Average Age: {customers['age'].mean():.1f} years")
    
    print("\nCustomer Segment Distribution:")
    segment_dist = customers['customer_segment'].value_counts()
    for segment, count in segment_dist.items():
        percentage = (count / len(customers)) * 100
        print(f"   {segment}: {count} customers ({percentage:.1f}%)")
    
    print_subsection("2. Geographic Distribution")
    print("\nTop 5 States by Customer Count:")
    state_dist = customers['state'].value_counts().head(5)
    for state, count in state_dist.items():
        print(f"   {state}: {count} customers")
    
    print_subsection("3. Purchase Behavior")
    print(f"Total Purchases Across All Customers: {customers['total_purchases'].sum()}")
    print(f"Average Purchases per Customer: {customers['total_purchases'].mean():.1f}")
    print(f"Most Active Customer: {customers['total_purchases'].max()} purchases")
    
    print("\nPurchases by Segment:")
    segment_purchases = customers.groupby('customer_segment')['total_purchases'].agg(['sum', 'mean'])
    segment_purchases.columns = ['Total Purchases', 'Avg per Customer']
    print(segment_purchases.round(2).to_string())
    
    return customers

def analyze_product_data(products):
    """Analyze product catalog and profitability"""
    print_section_header("PRODUCT CATALOG ANALYSIS")
    
    # Calculate profit margin
    products['profit_margin'] = products['selling_price'] - products['cost_price']
    products['profit_margin_percent'] = ((products['selling_price'] - products['cost_price']) / products['cost_price'] * 100)
    
    print_subsection("1. Product Portfolio")
    print(f"Total Products: {len(products)}")
    print(f"Active Products: {products['is_active'].sum()}")
    
    print("\nProducts by Category:")
    category_dist = products['category'].value_counts()
    for category, count in category_dist.items():
        print(f"   {category}: {count} products")
    
    print_subsection("2. Profitability Analysis")
    print(f"Average Profit Margin: Rs. {products['profit_margin'].mean():,.2f}")
    print(f"Average Profit Margin Percentage: {products['profit_margin_percent'].mean():.1f}%")
    
    print("\nTop 5 Most Profitable Products:")
    top_profit = products.nlargest(5, 'profit_margin')[['product_name', 'profit_margin', 'profit_margin_percent']]
    for idx, row in top_profit.iterrows():
        print(f"   {row['product_name']}: Rs. {row['profit_margin']:,.0f} ({row['profit_margin_percent']:.1f}%)")
    
    print_subsection("3. Inventory Status")
    low_stock = products[products['stock_quantity'] <= products['reorder_level']]
    print(f"Products Needing Reorder: {len(low_stock)}")
    if len(low_stock) > 0:
        print("\nLow Stock Alert:")
        for idx, row in low_stock.iterrows():
            print(f"   {row['product_name']}: {row['stock_quantity']} units (Reorder at: {row['reorder_level']})")
    
    print("\nCategory-wise Profitability:")
    category_profit = products.groupby('category').agg({
        'profit_margin': 'mean',
        'profit_margin_percent': 'mean',
        'stock_quantity': 'sum'
    }).round(2)
    category_profit.columns = ['Avg Profit (Rs.)', 'Avg Profit %', 'Total Stock']
    print(category_profit.to_string())
    
    return products

def analyze_regional_performance(regional):
    """Analyze regional sales performance"""
    print_section_header("REGIONAL PERFORMANCE ANALYSIS")
    
    print_subsection("1. Overall Regional Performance")
    regional_summary = regional.groupby('region').agg({
        'total_sales': 'sum',
        'total_orders': 'sum',
        'avg_order_value': 'mean'
    }).round(2)
    regional_summary.columns = ['Total Sales', 'Total Orders', 'Avg Order Value']
    regional_summary = regional_summary.sort_values('Total Sales', ascending=False)
    print(regional_summary.to_string())
    
    print_subsection("2. Target Achievement")
    regional['target_gap'] = regional['total_sales'] - regional['sales_target']
    regional['achievement_percent'] = (regional['total_sales'] / regional['sales_target'] * 100)
    
    target_summary = regional.groupby('region').agg({
        'target_achieved': lambda x: (x == 'Yes').sum(),
        'achievement_percent': 'mean'
    }).round(2)
    target_summary.columns = ['Months Target Achieved', 'Avg Achievement %']
    print(target_summary.to_string())
    
    print_subsection("3. Monthly Trends")
    monthly_trend = regional.groupby('month')['total_sales'].sum().sort_values(ascending=False)
    print("\nSales by Month:")
    for month, sales in monthly_trend.items():
        print(f"   {month}: Rs. {sales:,.2f}")
    
    print_subsection("4. Category Performance by Region")
    top_categories = regional.groupby('region')['top_category'].agg(lambda x: x.mode()[0] if len(x.mode()) > 0 else x.iloc[0])
    print("\nTop Performing Category by Region:")
    for region, category in top_categories.items():
        print(f"   {region}: {category}")
    
    return regional

def analyze_automation_opportunities(automation):
    """Analyze automation scenarios and ROI with safe ROI parsing."""
    import numpy as np
    import re

    print_section_header("AUTOMATION OPPORTUNITIES ANALYSIS")

    df = automation.copy()

    # --- Step 1: Calculate time savings ---
    df['weekly_time_hours'] = (df['time_per_task_minutes'] * df['frequency_per_week']) / 60
    df['annual_time_hours'] = df['weekly_time_hours'] * 52

    # --- Step 2: Clean and normalize ROI column ---
    if 'roi_estimate' in df.columns:
        df['roi_estimate'] = (
            df['roi_estimate']
            .astype(str)
            .str.replace('%', '', regex=False)  # Remove % symbols
            .str.replace(',', '.', regex=False)  # Handle commas
        )

        # Handle concatenated or messy ROI strings
        def extract_numbers(cell):
            parts = [p for p in re.split(r'[^0-9.]+', cell) if p.strip() != '']
            return np.mean([float(x) for x in parts]) if parts else np.nan

        df['roi_estimate'] = df['roi_estimate'].apply(extract_numbers)
        df['roi_estimate'] = pd.to_numeric(df['roi_estimate'], errors='coerce')

    else:
        print("⚠️ Warning: 'roi_estimate' column not found in automation data.")
        return df

    # --- Step 3: Time Investment Analysis ---
    print_subsection("1. Time Investment Analysis")
    print(f"Total Annual Hours Spent: {df['annual_time_hours'].sum():,.0f} hours")
    print(f"Average Time per Process: {df['annual_time_hours'].mean():,.0f} hours/year")

    top_time = df.nlargest(5, 'annual_time_hours')[['business_process', 'annual_time_hours']]
    print("\nTop 5 Time-Consuming Processes:")
    for idx, row in top_time.iterrows():
        print(f"   {row['business_process']}: {row['annual_time_hours']:,.0f} hours/year")

    # --- Step 4: ROI Analysis ---
    print_subsection("2. ROI Analysis")
    print(f"Total Annual Cost: Rs. {df['annual_cost'].sum():,.2f}")
    print(f"Average ROI Estimate: {df['roi_estimate'].mean():.1f}%")

    print("\nAutomation Potential Distribution:")
    potential_dist = df['automation_potential'].value_counts()
    for potential, count in potential_dist.items():
        print(f"   {potential}: {count} processes")

    print("\nTop 5 High-ROI Automation Opportunities:")
    high_roi = df[df['automation_potential'] == 'High'].nlargest(5, 'roi_estimate')
    for idx, row in high_roi.iterrows():
        print(f"   {row['business_process']}: {row['roi_estimate']:.0f}% ROI (Cost: Rs. {row['annual_cost']:,.0f})")

    return df

def create_comprehensive_visualizations(customers, products, regional, automation):
    """Create comprehensive visualizations across all datasets"""
    print_section_header("GENERATING VISUALIZATIONS")
    print("\nCreating comprehensive dashboard... Please wait.")
    
    fig = plt.figure(figsize=(20, 14))
    
    # 1. Customer Segment Distribution (Pie Chart)
    ax1 = plt.subplot(3, 4, 1)
    segment_counts = customers['customer_segment'].value_counts()
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1']
    ax1.pie(segment_counts.values, labels=segment_counts.index, autopct='%1.1f%%', 
            colors=colors, startangle=90)
    ax1.set_title('Customer Segment Distribution', fontsize=12, fontweight='bold', pad=15)
    
    # 2. Age Distribution (Histogram)
    ax2 = plt.subplot(3, 4, 2)
    ax2.hist(customers['age'], bins=10, color='steelblue', alpha=0.7, edgecolor='black')
    ax2.set_title('Customer Age Distribution', fontsize=12, fontweight='bold', pad=15)
    ax2.set_xlabel('Age')
    ax2.set_ylabel('Frequency')
    ax2.grid(True, alpha=0.3)
    
    # 3. Product Profit Margins by Category (Bar Chart)
    ax3 = plt.subplot(3, 4, 3)
    products['profit_margin'] = products['selling_price'] - products['cost_price']
    category_profit = products.groupby('category')['profit_margin'].mean().sort_values(ascending=False)
    bars = ax3.bar(range(len(category_profit)), category_profit.values, color='coral', alpha=0.8)
    ax3.set_title('Avg Profit Margin by Category', fontsize=12, fontweight='bold', pad=15)
    ax3.set_xlabel('Category')
    ax3.set_ylabel('Profit Margin (Rs.)')
    ax3.set_xticks(range(len(category_profit)))
    ax3.set_xticklabels(category_profit.index, rotation=45, ha='right')
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax3.text(bar.get_x() + bar.get_width()/2., height,
                f'Rs.{height:,.0f}', ha='center', va='bottom', fontsize=9)
    
    # 4. Regional Sales Performance (Bar Chart)
    ax4 = plt.subplot(3, 4, 4)
    regional_sales = regional.groupby('region')['total_sales'].sum().sort_values(ascending=False)
    bars = ax4.bar(range(len(regional_sales)), regional_sales.values, color='mediumseagreen', alpha=0.8)
    ax4.set_title('Total Sales by Region', fontsize=12, fontweight='bold', pad=15)
    ax4.set_xlabel('Region')
    ax4.set_ylabel('Sales (Rs.)')
    ax4.set_xticks(range(len(regional_sales)))
    ax4.set_xticklabels(regional_sales.index, rotation=45, ha='right')
    for bar in bars:
        height = bar.get_height()
        ax4.text(bar.get_x() + bar.get_width()/2., height,
                f'{height/1000:.0f}K', ha='center', va='bottom', fontsize=9)
    
    # 5. Monthly Sales Trend (Line Chart)
    ax5 = plt.subplot(3, 4, 5)
    month_order = ['Jan', 'Feb', 'Mar', 'Apr']
    monthly_sales = regional.groupby('month')['total_sales'].sum()
    monthly_sales = monthly_sales.reindex(month_order)
    ax5.plot(range(len(monthly_sales)), monthly_sales.values, marker='o', 
             linewidth=2, markersize=8, color='#E74C3C')
    ax5.set_title('Monthly Sales Trend (2024)', fontsize=12, fontweight='bold', pad=15)
    ax5.set_xlabel('Month')
    ax5.set_ylabel('Total Sales (Rs.)')
    ax5.set_xticks(range(len(monthly_sales)))
    ax5.set_xticklabels(monthly_sales.index)
    ax5.grid(True, alpha=0.3)
    for i, val in enumerate(monthly_sales.values):
        ax5.text(i, val, f'{val/1000:.0f}K', ha='center', va='bottom')
    
    # 6. Target Achievement by Region (Grouped Bar)
    ax6 = plt.subplot(3, 4, 6)
    target_data = regional.groupby('region').agg({
        'total_sales': 'sum',
        'sales_target': 'sum'
    })
    x = np.arange(len(target_data))
    width = 0.35
    bars1 = ax6.bar(x - width/2, target_data['total_sales'], width, label='Actual Sales', color='#3498DB')
    bars2 = ax6.bar(x + width/2, target_data['sales_target'], width, label='Target', color='#95A5A6')
    ax6.set_title('Sales vs Target by Region', fontsize=12, fontweight='bold', pad=15)
    ax6.set_xlabel('Region')
    ax6.set_ylabel('Amount (Rs.)')
    ax6.set_xticks(x)
    ax6.set_xticklabels(target_data.index, rotation=45, ha='right')
    ax6.legend()
    ax6.grid(True, alpha=0.3, axis='y')
    
    # 7. Product Stock Levels (Bar Chart)
    ax7 = plt.subplot(3, 4, 9)
    products_sorted = products.sort_values('stock_quantity', ascending=False).head(8)
    colors_stock = ['red' if qty <= reorder else 'green' 
                   for qty, reorder in zip(products_sorted['stock_quantity'], products_sorted['reorder_level'])]
    bars = ax7.bar(range(len(products_sorted)), products_sorted['stock_quantity'].values, 
                   color=colors_stock, alpha=0.7)
    ax7.set_title('Top 8 Products - Stock Levels', fontsize=12, fontweight='bold', pad=15)
    ax7.set_xlabel('Product')
    ax7.set_ylabel('Stock Quantity')
    ax7.set_xticks(range(len(products_sorted)))
    ax7.set_xticklabels([name[:12] + '...' if len(name) > 12 else name 
                         for name in products_sorted['product_name'].values], 
                        rotation=45, ha='right', fontsize=8)
    
    # 8. Purchases by Customer Segment (Bar Chart)
    ax8 = plt.subplot(3, 4, 10)
    segment_purchases = customers.groupby('customer_segment')['total_purchases'].sum().sort_values(ascending=False)
    bars = ax8.bar(range(len(segment_purchases)), segment_purchases.values, 
                    color=['#FF6B6B', '#4ECDC4', '#45B7D1'], alpha=0.8)
    ax8.set_title('Total Purchases by Segment', fontsize=12, fontweight='bold', pad=15)
    ax8.set_xlabel('Customer Segment')
    ax8.set_ylabel('Total Purchases')
    ax8.set_xticks(range(len(segment_purchases)))
    ax8.set_xticklabels(segment_purchases.index, rotation=45, ha='right')
    for bar in bars:
        height = bar.get_height()
        ax8.text(bar.get_x() + bar.get_width()/2., height,
                 f'{int(height)}', ha='center', va='bottom', fontsize=9)
    
    # 9. Regional Order Volume (Bar Chart)
    ax9 = plt.subplot(3, 4, 11)
    regional_orders = regional.groupby('region')['total_orders'].sum().sort_values(ascending=False)
    bars = ax9.bar(range(len(regional_orders)), regional_orders.values, 
                    color='#F39C12', alpha=0.8)
    ax9.set_title('Total Orders by Region', fontsize=12, fontweight='bold', pad=15)
    ax9.set_xlabel('Region')
    ax9.set_ylabel('Number of Orders')
    ax9.set_xticks(range(len(regional_orders)))
    ax9.set_xticklabels(regional_orders.index, rotation=45, ha='right')
    for bar in bars:
        height = bar.get_height()
        ax9.text(bar.get_x() + bar.get_width()/2., height,
                 f'{int(height)}', ha='center', va='bottom', fontsize=9)
    
    # 10. Correlation Heatmap - Product Metrics
    ax10 = plt.subplot(3, 4, 12)
    product_corr = products[['cost_price', 'selling_price', 'stock_quantity', 'reorder_level']].corr()
    sns.heatmap(product_corr, annot=True, fmt='.2f', cmap='coolwarm', 
                center=0, square=True, ax=ax10, cbar_kws={'shrink': 0.8},
                linewidths=1, linecolor='white')
    ax10.set_title('Product Metrics Correlation', fontsize=12, fontweight='bold', pad=15)
    plt.setp(ax10.xaxis.get_majorticklabels(), rotation=45, ha='right', fontsize=9)
    plt.setp(ax10.yaxis.get_majorticklabels(), rotation=0, fontsize=9)
    
    plt.tight_layout(pad=2.0)
    plt.savefig('comprehensive_business_analysis.png', dpi=300, bbox_inches='tight')
    print("\nVisualization dashboard created successfully!")
    print("Saved as: comprehensive_business_analysis.png")
    plt.show()

def generate_executive_summary(customers, products, regional, automation):
    """Generate executive summary with key insights"""
    print_section_header("EXECUTIVE SUMMARY - KEY INSIGHTS")
    
    # Customer insights
    premium_customers = customers[customers['customer_segment'] == 'Premium']
    premium_revenue_contribution = premium_customers['total_purchases'].sum()
    
    # Product insights
    products['profit_margin'] = products['selling_price'] - products['cost_price']
    avg_profit_margin = products['profit_margin'].mean()
    most_profitable_category = products.groupby('category')['profit_margin'].mean().idxmax()
    
    # Regional insights
    best_region = regional.groupby('region')['total_sales'].sum().idxmax()
    best_region_sales = regional.groupby('region')['total_sales'].sum().max()
    total_sales = regional['total_sales'].sum()
    
    # Automation insights
    high_roi_count = len(automation[automation['automation_potential'] == 'High'])
    total_automation_cost = automation['annual_cost'].sum()
    avg_roi = automation[automation['automation_potential'] == 'High']['roi_estimate'].mean()
    
    print("\n1. CUSTOMER BASE")
    print(f"   - Total active customers: {len(customers)}")
    print(f"   - Premium segment: {len(premium_customers)} customers ({len(premium_customers)/len(customers)*100:.1f}%)")
    print(f"   - Premium customers contribute {premium_revenue_contribution} purchases")
    print(f"   - Geographic spread: {customers['state'].nunique()} states")
    
    print("\n2. PRODUCT PORTFOLIO")
    print(f"   - Active products: {len(products)} across {products['category'].nunique()} categories")
    print(f"   - Average profit margin: Rs. {avg_profit_margin:,.2f} per product")
    print(f"   - Most profitable category: {most_profitable_category}")
    print(f"   - Products needing reorder: {len(products[products['stock_quantity'] <= products['reorder_level']])}")
    
    print("\n3. REGIONAL PERFORMANCE")
    print(f"   - Total sales (Jan-Apr 2024): Rs. {total_sales:,.2f}")
    print(f"   - Best performing region: {best_region} (Rs. {best_region_sales:,.2f})")
    print(f"   - This represents {best_region_sales/total_sales*100:.1f}% of total sales")
    print(f"   - Average order value varies from Rs. 941 to Rs. 2,440")
    
    print("\n4. AUTOMATION OPPORTUNITIES")
    print(f"   - High-potential automation opportunities: {high_roi_count} processes")
    print(f"   - Total current annual cost: Rs. {total_automation_cost:,.2f}")
    print(f"   - Average ROI for high-potential automations: {avg_roi:.0f}%")
    print(f"   - Potential annual time savings: {automation['annual_time_hours'].sum():,.0f} hours")
    
    print("\n5. STRATEGIC RECOMMENDATIONS")
    print("   a) Customer Strategy:")
    print("      - Focus retention programs on Premium segment")
    print("      - Expand customer base in high-performing states")
    print("      - Create targeted campaigns for Standard segment upgrades")
    
    print("\n   b) Product Strategy:")
    print(f"      - Prioritize stock replenishment for low-inventory items")
    print(f"      - Increase focus on {most_profitable_category} category")
    print("      - Consider supplier diversification for risk mitigation")
    
    print("\n   c) Regional Strategy:")
    print(f"      - Replicate {best_region} region's success in other markets")
    print("      - Address regions falling short of targets")
    print("      - Optimize inventory distribution based on regional demand")
    
    print("\n   d) Automation Strategy:")
    print("      - Prioritize high-ROI automation projects")
    print("      - Start with processes requiring 20+ hours/week")
    print("      - Expected cost savings could exceed Rs. 5 lakhs annually")
    
    print("\n" + "=" * 80)

def generate_detailed_report():
    """Generate detailed PDF-style report in text format"""
    print_section_header("DETAILED ANALYSIS REPORT")
    
    report = """
    BUSINESS INTELLIGENCE ANALYSIS REPORT
    Generated: 2024
    
    OBJECTIVE:
    This comprehensive analysis examines customer demographics, product profitability,
    regional sales performance, and automation opportunities to provide actionable
    insights for business growth and operational efficiency.
    
    METHODOLOGY:
    - Multi-table data integration from 4 primary datasets
    - Statistical analysis using pandas library
    - Visual analytics using matplotlib and seaborn
    - Correlation analysis for identifying patterns
    
    KEY FINDINGS:
    
    1. Customer Analysis reveals a diverse customer base with varying purchase patterns
       across different segments and geographic locations.
    
    2. Product Analysis shows varying profitability margins across categories, with
       some products requiring immediate inventory attention.
    
    3. Regional Performance indicates significant variation in sales achievement,
       with opportunities for replicating best practices across regions.
    
    4. Automation Analysis identifies high-ROI opportunities that could significantly
       reduce operational costs and improve efficiency.
    
    CONCLUSION:
    The integrated analysis provides a comprehensive view of business operations,
    highlighting areas of strength and opportunities for improvement. Implementation
    of recommended strategies could lead to enhanced profitability, operational
    efficiency, and customer satisfaction.
    """
    
    print(report)

# Main execution function
def main():
    """Main execution function"""
    print("\n" + "=" * 80)
    print("       COMPREHENSIVE BUSINESS INTELLIGENCE ANALYSIS SYSTEM")
    print("                    Multi-Table Data Analytics")
    print("=" * 80)
    
    # File path - modify this to match your Excel file location
    file_path = r'D:\30DaysOfPython\BusinessAnalysis\Forge Program Master Dataset - Week 1-Clean.xlsx'
    
    print("\nNOTE: To use your Excel file, ensure it is named:")
    print(f"      '{file_path}'")
    print("      and is in the same folder as this script.")
    print("\nIf file is not found, sample data will be used for demonstration.\n")
    
    # Load all datasets
    datasets = load_all_datasets(file_path)
    
    # Extract individual datasets safely (case-insensitive)
    datasets = {k.lower(): v for k, v in datasets.items()}

    customers = datasets.get('customer profiles') or datasets.get('customer_profiles')
    products = datasets.get('product catalog') or datasets.get('product_catalog')
    regional = datasets.get('regional performance') or datasets.get('regional_performance')
    automation = datasets.get('automation scenarios') or datasets.get('automation_scenarios')
    
    # Perform comprehensive analysis
    if customers is not None:
        customers = analyze_customer_data(customers)
    
    if products is not None:
        products = analyze_product_data(products)
    
    if regional is not None:
        regional = analyze_regional_performance(regional)
    
    if automation is not None:
        # Ensure ROI column is numeric (robust against bad data)
        automation['roi_estimate'] = pd.to_numeric(automation['roi_estimate'], errors='coerce')
        automation = analyze_automation_opportunities(automation)
    
    # Create visualizations
    if all(df is not None for df in [customers, products, regional, automation]):
        create_comprehensive_visualizations(customers, products, regional, automation)
        generate_executive_summary(customers, products, regional, automation)
    
    # Generate report
    generate_detailed_report()
    
    print("\n" + "=" * 80)
    print("           ANALYSIS COMPLETE - ALL OUTPUTS GENERATED")
    print("=" * 80)
    print("\nOutput Files:")
    print("   1. comprehensive_business_analysis.png - Visual dashboard")
    print("   2. Console output - Detailed statistics and insights")
    print("\nThank you for using the Business Intelligence Analysis System!")
    print("=" * 80 + "\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nAnalysis interrupted by user.")
    except Exception as e:
        import traceback
        print(f"\n\nError occurred: {e}")
        traceback.print_exc()
        print("Please check your data files and try again.")