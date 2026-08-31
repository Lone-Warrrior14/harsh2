import io
import json
import datetime
import pandas as pd
import numpy as np
from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)

def normalize_cols(df):
    """
    Standardize Excel column headers across MB52 and COHV exports.
    """
    df.columns = [str(col).strip() for col in df.columns]
    col_map = {}
    for col in df.columns:
        c_lower = col.lower().replace(' ', '').replace('_', '').replace('-', '').replace('/', '')
        if c_lower in ['article', 'material', 'materialnumber', 'item', 'matnr']:
            col_map[col] = 'Article'
        elif 'valuetransit' in c_lower or ('value' in c_lower and ('transit' in c_lower or 'transfer' in c_lower)):
            col_map[col] = 'Value Transit'
        elif 'value' in c_lower and 'unrestricted' in c_lower:
            col_map[col] = 'Value Unrestricted'
        elif ('transit' in c_lower or 'transfer' in c_lower) and 'value' not in c_lower:
            col_map[col] = 'Transit'
        elif c_lower in ['unrestricted', 'unrestrictedstock', 'unrestricteduse', 'labst', 'available', 'availablestock']:
            col_map[col] = 'Unrestricted'
        elif c_lower in ['salesdocument', 'salesdoc', 'so', 'salesorder', 'vbeln']:
            col_map[col] = 'Sales Document'
        elif c_lower in ['requirementquantity(einheit)', 'requirementquantity', 'reqqty', 'requirementqty', 'quantity', 'bdmng', 'qty']:
            col_map[col] = 'Requirement quantity (EINHEIT)'
        elif c_lower in ['materialdescription', 'description', 'articledescription', 'maktx']:
            col_map[col] = 'Material Description'
        elif c_lower in ['requirementdate', 'deliverydate', 'reqdate', 'basicfinishdate', 'finishdate', 'scheduledate', 'date']:
            col_map[col] = 'Requirement date'
        elif c_lower in ['order', 'aufnr', 'productionorder', 'plannedorder']:
            col_map[col] = 'Order'
        elif c_lower in ['peggedrequirement', 'peggedreq']:
            col_map[col] = 'Pegged requirement'
        elif c_lower in ['phantomitem', 'phantom', 'phantomflag', 'phantomit']:
            col_map[col] = 'Phantom item'
        elif c_lower in ['salesoffice', 'office', 'salesoff']:
            col_map[col] = 'Sales Office'
        elif c_lower in ['site', 'plant', 'werks']:
            col_map[col] = 'Site'
    
    # Avoid duplicate column names after rename
    new_cols = []
    seen = set()
    for col in df.columns:
        target = col_map.get(col, col)
        if target in seen:
            # Append suffix to avoid duplicate DataFrame columns
            count = 1
            alt_target = f"{target}_{count}"
            while alt_target in seen:
                count += 1
                alt_target = f"{target}_{count}"
            seen.add(alt_target)
            col_map[col] = alt_target
        else:
            seen.add(target)
            col_map[col] = target

    df.rename(columns=col_map, inplace=True)
    return df


def calculate_availability_predictions(file_mb52, cohv_files):
    """
    Core calculation engine for Product Availability Predictor.
    Calculates exact Stock Depletion Dates, Runway Days, and Order Fulfillment.
    Combines Unrestricted + Transit/Transfer Stock for total available inventory.
    Excludes rows where Phantom item == 'X'.
    """
    today = pd.Timestamp.now().normalize()

    # 1. Read MB52 stock file
    df_mb52 = pd.read_excel(file_mb52)
    df_mb52 = normalize_cols(df_mb52)

    if 'Article' not in df_mb52.columns:
        raise ValueError(f"MB52 file missing required 'Article' column. Found: {list(df_mb52.columns)}")

    # Remove Phantom items from MB52 if column exists
    if 'Phantom item' in df_mb52.columns:
        df_mb52 = df_mb52[df_mb52['Phantom item'].astype(str).str.strip().str.upper() != 'X'].copy()

    # Filter out rows where Material Description contains 'GLS' (case-insensitive)
    if 'Material Description' in df_mb52.columns:
        df_mb52 = df_mb52[~df_mb52['Material Description'].astype(str).str.contains('gls', case=False, na=False)].copy()

    df_mb52['Article'] = df_mb52['Article'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0')
    
    # Parse Unrestricted Stock (Qty & Value)
    unrest_series = df_mb52['Unrestricted'] if 'Unrestricted' in df_mb52.columns else pd.Series(0, index=df_mb52.index)
    val_unrest_series = df_mb52['Value Unrestricted'] if 'Value Unrestricted' in df_mb52.columns else pd.Series(0, index=df_mb52.index)
    df_mb52['Unrestricted'] = pd.to_numeric(unrest_series, errors='coerce').fillna(0)
    df_mb52['Value Unrestricted'] = pd.to_numeric(val_unrest_series, errors='coerce').fillna(0)

    # Parse Transit / Transfer Stock (Qty & Value)
    transit_series = df_mb52['Transit'] if 'Transit' in df_mb52.columns else pd.Series(0, index=df_mb52.index)
    val_transit_series = df_mb52['Value Transit'] if 'Value Transit' in df_mb52.columns else pd.Series(0, index=df_mb52.index)
    df_mb52['Transit'] = pd.to_numeric(transit_series, errors='coerce').fillna(0)
    df_mb52['Value Transit'] = pd.to_numeric(val_transit_series, errors='coerce').fillna(0)

    # Combined Total Available Stock = Unrestricted + Transit/Transfer
    df_mb52['Total_Available_Stock'] = df_mb52['Unrestricted'] + df_mb52['Transit']
    df_mb52['Total_Stock_Value'] = df_mb52['Value Unrestricted'] + df_mb52['Value Transit']

    # 2. Read and combine COHV requirement files
    df_cohv_list = []
    for f in cohv_files:
        if f:
            fname = getattr(f, 'filename', 'file.xlsx')
            if fname:
                df_temp = pd.read_excel(f)
                df_temp = normalize_cols(df_temp)
                df_cohv_list.append(df_temp)

    if not df_cohv_list:
        raise ValueError("No valid COHV requirement files uploaded.")

    df_cohv = pd.concat(df_cohv_list, ignore_index=True)
    
    req_cols = ['Sales Document', 'Article', 'Requirement quantity (EINHEIT)', 'Requirement date']
    missing = [c for c in req_cols if c not in df_cohv.columns]
    if missing:
        raise ValueError(f"COHV file(s) missing required columns: {missing}. Found: {list(df_cohv.columns)}")

    # CRITICAL: Exclude all rows where Phantom item == 'X'
    if 'Phantom item' in df_cohv.columns:
        df_cohv = df_cohv[df_cohv['Phantom item'].astype(str).str.strip().str.upper() != 'X'].copy()

    if 'Material Description' not in df_cohv.columns:
        df_cohv['Material Description'] = df_cohv['Article']

    # Filter out rows where Material Description contains 'GLS' (case-insensitive)
    if 'Material Description' in df_cohv.columns:
        df_cohv = df_cohv[~df_cohv['Material Description'].astype(str).str.contains('gls', case=False, na=False)].copy()

    df_cohv = df_cohv.dropna(subset=['Article']).copy()
    df_cohv['Article'] = df_cohv['Article'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0')
    df_cohv['Sales Document'] = df_cohv['Sales Document'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    df_cohv['Requirement quantity (EINHEIT)'] = pd.to_numeric(df_cohv['Requirement quantity (EINHEIT)'], errors='coerce').fillna(0)
    df_cohv['Requirement date'] = pd.to_datetime(df_cohv['Requirement date'], errors='coerce')
    
    df_cohv['Requirement_Date_Clean'] = df_cohv['Requirement date'].fillna(pd.Timestamp('2099-12-31'))
    df_cohv = df_cohv.sort_values(by=['Article', 'Requirement_Date_Clean', 'Sales Document']).reset_index(drop=True)

    # MB52 stock aggregation by Article with monetary valuation
    stock_summary = df_mb52.groupby('Article', as_index=False).agg(
        Unrestricted_Stock=('Unrestricted', 'sum'),
        Transit_Stock=('Transit', 'sum'),
        Initial_Stock_MB52=('Total_Available_Stock', 'sum'),
        Initial_Stock_Value=('Total_Stock_Value', 'sum')
    )
    stock_summary['Unit_Price'] = np.where(
        stock_summary['Initial_Stock_MB52'] > 0,
        stock_summary['Initial_Stock_Value'] / stock_summary['Initial_Stock_MB52'],
        0.0
    )

    # 3. Merge MB52 Stock onto requirement lines
    df_lines = pd.merge(df_cohv, stock_summary, on='Article', how='left')
    df_lines['Initial_Stock_MB52'] = df_lines['Initial_Stock_MB52'].fillna(0)

    # Cumulative demand and stock balance calculation
    df_lines['Cumulative_Demand'] = df_lines.groupby('Article')['Requirement quantity (EINHEIT)'].cumsum()
    df_lines['Projected_Stock_Balance'] = df_lines['Initial_Stock_MB52'] - df_lines['Cumulative_Demand']
    df_lines['Opening_Stock_Balance'] = df_lines['Projected_Stock_Balance'] + df_lines['Requirement quantity (EINHEIT)']

    df_lines['Fulfilled_Qty'] = np.minimum(
        df_lines['Requirement quantity (EINHEIT)'],
        np.maximum(0, df_lines['Opening_Stock_Balance'])
    )
    df_lines['Shortage_Qty'] = df_lines['Requirement quantity (EINHEIT)'] - df_lines['Fulfilled_Qty']

    df_lines['Line_Status'] = np.where(
        df_lines['Fulfilled_Qty'] == df_lines['Requirement quantity (EINHEIT)'],
        'FULLY COVERED ON TIME',
        np.where(df_lines['Fulfilled_Qty'] > 0, 'PARTIALLY COVERED', 'STOCKOUT / UNFULFILLABLE')
    )

    # Find exact stock depletion date (first requirement line where shortage occurs)
    stockout_mask = (df_lines['Projected_Stock_Balance'] < 0) | (df_lines['Shortage_Qty'] > 0)
    stockout_rows = df_lines[stockout_mask]
    earliest_stockout = stockout_rows.groupby('Article')['Requirement_Date_Clean'].min().reset_index()
    earliest_stockout.rename(columns={'Requirement_Date_Clean': 'Depletion_Date_Dt'}, inplace=True)

    # 4. Article Level Predictions
    meta_cohv = df_cohv[['Article', 'Material Description']].dropna(subset=['Article', 'Material Description']) if 'Material Description' in df_cohv.columns else pd.DataFrame()
    meta_mb52 = df_mb52[['Article', 'Material Description']].dropna(subset=['Article', 'Material Description']) if 'Material Description' in df_mb52.columns else pd.DataFrame()

    if not meta_cohv.empty:
        meta_cohv['Article'] = meta_cohv['Article'].astype(str).str.strip()
        meta_cohv['Material Description'] = meta_cohv['Material Description'].astype(str).str.strip()
    if not meta_mb52.empty:
        meta_mb52['Article'] = meta_mb52['Article'].astype(str).str.strip()
        meta_mb52['Material Description'] = meta_mb52['Material Description'].astype(str).str.strip()

    all_meta = pd.concat([meta_cohv, meta_mb52], ignore_index=True)
    article_meta = all_meta.drop_duplicates(subset=['Article'], keep='first')
    
    article_demand = df_lines.groupby('Article', as_index=False).agg(
        Total_Future_Demand=('Requirement quantity (EINHEIT)', 'sum'),
        Total_Fulfilled_Qty=('Fulfilled_Qty', 'sum'),
        Total_Shortage_Qty=('Shortage_Qty', 'sum'),
        Earliest_Demand_Date=('Requirement date', 'min'),
        Latest_Demand_Date=('Requirement date', 'max'),
        Total_Orders_Count=('Sales Document', 'nunique')
    )

    all_article_ids = sorted(list(set(stock_summary['Article']).union(set(df_cohv['Article']))))
    all_articles = pd.DataFrame({'Article': all_article_ids})
    
    article_summary = pd.merge(all_articles, stock_summary, on='Article', how='left')
    article_summary['Unrestricted_Stock'] = article_summary['Unrestricted_Stock'].fillna(0)
    article_summary['Transit_Stock'] = article_summary['Transit_Stock'].fillna(0)
    article_summary['Initial_Stock_MB52'] = article_summary['Initial_Stock_MB52'].fillna(0)
    
    article_summary = pd.merge(article_summary, article_meta, on='Article', how='left')
    article_summary['Material Description'] = article_summary['Material Description'].fillna(article_summary['Article'])
    
    article_summary = pd.merge(article_summary, article_demand, on='Article', how='left')
    article_summary['Total_Future_Demand'] = article_summary['Total_Future_Demand'].fillna(0)
    article_summary['Total_Fulfilled_Qty'] = article_summary['Total_Fulfilled_Qty'].fillna(0)
    article_summary['Total_Shortage_Qty'] = article_summary['Total_Shortage_Qty'].fillna(0)
    article_summary['Total_Orders_Count'] = article_summary['Total_Orders_Count'].fillna(0).astype(int)

    article_summary['Initial_Stock_Value'] = article_summary['Initial_Stock_Value'].fillna(0.0)
    article_summary['Unit_Price'] = article_summary['Unit_Price'].fillna(0.0)
    article_summary['Total_Demand_Value'] = (article_summary['Total_Future_Demand'] * article_summary['Unit_Price']).round(2)
    article_summary['Total_Shortage_Value'] = (article_summary['Total_Shortage_Qty'] * article_summary['Unit_Price']).round(2)

    unfulfilled_art = df_lines[df_lines['Shortage_Qty'] > 0.0001].groupby('Article')['Sales Document'].nunique().reset_index()
    unfulfilled_art.rename(columns={'Sales Document': 'Unfulfilled_Orders_Count'}, inplace=True)
    article_summary = pd.merge(article_summary, unfulfilled_art, on='Article', how='left')
    article_summary['Unfulfilled_Orders_Count'] = article_summary['Unfulfilled_Orders_Count'].fillna(0).astype(int)

    # Exclude inactive items (articles with 0 MB52 stock and 0 COHV demand)
    article_summary = article_summary[~((article_summary['Initial_Stock_MB52'] == 0) & (article_summary['Total_Future_Demand'] == 0))].copy()

    article_summary = pd.merge(article_summary, earliest_stockout, on='Article', how='left')
    article_summary['Net_Projected_Balance'] = article_summary['Initial_Stock_MB52'] - article_summary['Total_Future_Demand']
    article_summary['Stock_Coverage_Pct'] = np.where(
        article_summary['Total_Future_Demand'] > 0,
        np.minimum(100.0, (article_summary['Total_Fulfilled_Qty'] / article_summary['Total_Future_Demand']) * 100.0),
        100.0
    ).round(1)

    def get_depletion_info(row):
        depletion_dt = row['Depletion_Date_Dt']
        init_stock = row['Initial_Stock_MB52']
        net_bal = row['Net_Projected_Balance']
        demand = row['Total_Future_Demand']
        
        if pd.isna(depletion_dt) and net_bal >= 0:
            return 'No Depletion Expected', 999
            
        if pd.notna(depletion_dt):
            date_str = depletion_dt.strftime('%Y-%m-%d')
            days = (depletion_dt - today).days
            if init_stock <= 0:
                return f"{date_str} (0 Start Stock)", days
            else:
                return date_str, days
        else:
            if demand > 0:
                return 'Immediate Depletion', 0
            return 'No Depletion Expected', 999

    depletion_results = article_summary.apply(get_depletion_info, axis=1)
    article_summary['Stock_Depletion_Date'] = [r[0] for r in depletion_results]
    article_summary['Days_Until_Depletion'] = [r[1] for r in depletion_results]

    def get_status(row):
        if row['Initial_Stock_MB52'] <= 0 and row['Total_Future_Demand'] > 0:
            return 'CRITICAL STOCKOUT'
        elif row['Net_Projected_Balance'] < 0:
            return 'STOCKOUT RISK'
        else:
            return 'SUFFICIENT STOCK'

    article_summary['Status'] = article_summary.apply(get_status, axis=1)

    status_order = {'CRITICAL STOCKOUT': 0, 'STOCKOUT RISK': 1, 'SUFFICIENT STOCK': 2}
    article_summary['status_rank'] = article_summary['Status'].map(status_order)
    article_summary = article_summary.sort_values(
        by=['status_rank', 'Total_Shortage_Qty', 'Initial_Stock_MB52'],
        ascending=[True, False, True]
    ).drop(columns=['status_rank', 'Depletion_Date_Dt'])

    # 5. MATERIAL DESCRIPTION LEVEL PREDICTIONS
    mat_summary = article_summary.groupby('Material Description', as_index=False).agg(
        Associated_Articles=('Article', lambda x: ', '.join(sorted([str(i) for i in x.unique() if pd.notna(i)]))),
        Unrestricted_Stock=('Unrestricted_Stock', 'sum'),
        Transit_Stock=('Transit_Stock', 'sum'),
        Initial_Stock_MB52=('Initial_Stock_MB52', 'sum'),
        Total_Future_Demand=('Total_Future_Demand', 'sum'),
        Total_Fulfilled_Qty=('Total_Fulfilled_Qty', 'sum'),
        Total_Shortage_Qty=('Total_Shortage_Qty', 'sum'),
        Earliest_Demand_Date=('Earliest_Demand_Date', 'min'),
        Latest_Demand_Date=('Latest_Demand_Date', 'max')
    )

    if not stockout_rows.empty:
        mat_stockout = stockout_rows.groupby('Material Description')['Requirement_Date_Clean'].min().reset_index()
        mat_stockout.rename(columns={'Requirement_Date_Clean': 'Depletion_Date_Dt'}, inplace=True)
        mat_summary = pd.merge(mat_summary, mat_stockout, on='Material Description', how='left')
    else:
        mat_summary['Depletion_Date_Dt'] = pd.NaT

    mat_summary['Net_Projected_Balance'] = mat_summary['Initial_Stock_MB52'] - mat_summary['Total_Future_Demand']
    mat_summary['Stock_Coverage_Pct'] = np.where(
        mat_summary['Total_Future_Demand'] > 0,
        np.minimum(100.0, (mat_summary['Total_Fulfilled_Qty'] / mat_summary['Total_Future_Demand']) * 100.0),
        100.0
    ).round(1)

    mat_orders = df_lines.groupby('Material Description', as_index=False).agg(
        Total_Orders_Count=('Sales Document', 'nunique')
    )
    mat_unfulfilled = df_lines[df_lines['Shortage_Qty'] > 0.0001].groupby('Material Description', as_index=False).agg(
        Unfulfilled_Orders_Count=('Sales Document', 'nunique')
    )
    mat_summary = pd.merge(mat_summary, mat_orders, on='Material Description', how='left')
    mat_summary = pd.merge(mat_summary, mat_unfulfilled, on='Material Description', how='left')
    mat_summary['Total_Orders_Count'] = mat_summary['Total_Orders_Count'].fillna(0).astype(int)
    mat_summary['Unfulfilled_Orders_Count'] = mat_summary['Unfulfilled_Orders_Count'].fillna(0).astype(int)

    mat_depletion_results = mat_summary.apply(get_depletion_info, axis=1)
    mat_summary['Stock_Depletion_Date'] = [r[0] for r in mat_depletion_results]
    mat_summary['Days_Until_Depletion'] = [r[1] for r in mat_depletion_results]
    mat_summary['Status'] = mat_summary.apply(get_status, axis=1)
    
    mat_summary['status_rank'] = mat_summary['Status'].map(status_order)
    mat_summary = mat_summary.sort_values(
        by=['status_rank', 'Total_Shortage_Qty', 'Initial_Stock_MB52'],
        ascending=[True, False, True]
    ).drop(columns=['status_rank', 'Depletion_Date_Dt'])

    # 6. Sales Order Level Prediction
    so_summary = df_lines.groupby('Sales Document', as_index=False).agg(
        Total_Line_Items=('Article', 'count'),
        Distinct_Articles=('Article', 'nunique'),
        Total_Ordered_Qty=('Requirement quantity (EINHEIT)', 'sum'),
        Deliverable_Qty=('Fulfilled_Qty', 'sum'),
        Shortage_Qty=('Shortage_Qty', 'sum'),
        Earliest_Delivery_Date=('Requirement date', 'min'),
        Latest_Delivery_Date=('Requirement date', 'max')
    )

    so_summary['SO_Fulfillment_Pct'] = np.where(
        so_summary['Total_Ordered_Qty'] > 0,
        (so_summary['Deliverable_Qty'] / so_summary['Total_Ordered_Qty'] * 100.0),
        100.0
    ).round(1)

    def get_so_status(row):
        if row['SO_Fulfillment_Pct'] >= 100.0:
            return 'READY FOR FULL DELIVERY'
        elif row['SO_Fulfillment_Pct'] > 0:
            return 'PARTIAL DELIVERY RISK'
        else:
            return 'UNFULFILLABLE (NO STOCK)'

    so_summary['SO_Status'] = so_summary.apply(get_so_status, axis=1)
    # First Come First Serve based on date (Earliest Delivery Date first)
    so_summary = so_summary.sort_values(by=['Earliest_Delivery_Date', 'Sales Document'], ascending=[True, True])

    df_lines_clean = df_lines.copy()
    df_lines_clean['Requirement date_str'] = df_lines_clean['Requirement date'].dt.strftime('%Y-%m-%d').fillna('N/A')

    return article_summary, mat_summary, so_summary, df_lines_clean


def generate_excel_report(article_summary, mat_summary, so_summary, df_lines):
    """
    Generates a multi-tab Excel report containing Stock Depletion Dates and predictions.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#1E293B', 'font_color': '#FFFFFF',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        
        critical_format = workbook.add_format({'bg_color': '#FECDD3', 'font_color': '#9F1239', 'bold': True, 'border': 1})
        risk_format = workbook.add_format({'bg_color': '#FEF08A', 'font_color': '#854D0E', 'bold': True, 'border': 1})
        ok_format = workbook.add_format({'bg_color': '#D1FAE5', 'font_color': '#065F46', 'bold': True, 'border': 1})

        # Sheet 1: Material Description Forecast
        df_mat_export = mat_summary.copy()
        df_mat_export['Earliest_Demand_Date'] = df_mat_export['Earliest_Demand_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
        df_mat_export['Latest_Demand_Date'] = df_mat_export['Latest_Demand_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')

        df_mat_export.to_excel(writer, sheet_name='Material_Description_Forecast', index=False)
        ws_mat = writer.sheets['Material_Description_Forecast']
        ws_mat.freeze_panes(1, 0)
        for col_idx, col_name in enumerate(df_mat_export.columns):
            ws_mat.write(0, col_idx, col_name, header_format)
            ws_mat.set_column(col_idx, col_idx, max(len(col_name) + 3, 16))

        mat_status_col = df_mat_export.columns.get_loc('Status')
        for row_idx, val in enumerate(df_mat_export['Status'], start=1):
            if val == 'CRITICAL STOCKOUT':
                ws_mat.write(row_idx, mat_status_col, val, critical_format)
            elif val == 'STOCKOUT RISK':
                ws_mat.write(row_idx, mat_status_col, val, risk_format)
            else:
                ws_mat.write(row_idx, mat_status_col, val, ok_format)

        # Sheet 2: Article Predictions
        df_art_export = article_summary.copy()
        df_art_export['Earliest_Demand_Date'] = df_art_export['Earliest_Demand_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
        df_art_export['Latest_Demand_Date'] = df_art_export['Latest_Demand_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
        
        df_art_export.to_excel(writer, sheet_name='Article_Availability_Forecast', index=False)
        ws_art = writer.sheets['Article_Availability_Forecast']
        ws_art.freeze_panes(1, 0)
        for col_idx, col_name in enumerate(df_art_export.columns):
            ws_art.write(0, col_idx, col_name, header_format)
            ws_art.set_column(col_idx, col_idx, max(len(col_name) + 3, 14))

        status_col = df_art_export.columns.get_loc('Status')
        for row_idx, val in enumerate(df_art_export['Status'], start=1):
            if val == 'CRITICAL STOCKOUT':
                ws_art.write(row_idx, status_col, val, critical_format)
            elif val == 'STOCKOUT RISK':
                ws_art.write(row_idx, status_col, val, risk_format)
            else:
                ws_art.write(row_idx, status_col, val, ok_format)

        # Sheet 3: Sales Order Fulfillment
        df_so_export = so_summary.copy()
        df_so_export['Earliest_Delivery_Date'] = df_so_export['Earliest_Delivery_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
        df_so_export['Latest_Delivery_Date'] = df_so_export['Latest_Delivery_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')

        df_so_export.to_excel(writer, sheet_name='Sales_Order_Fulfillment', index=False)
        ws_so = writer.sheets['Sales_Order_Fulfillment']
        ws_so.freeze_panes(1, 0)
        for col_idx, col_name in enumerate(df_so_export.columns):
            ws_so.write(0, col_idx, col_name, header_format)
            ws_so.set_column(col_idx, col_idx, max(len(col_name) + 3, 15))

        # Sheet 4: Chronological Details
        cols_to_export = [
            'Requirement date_str', 'Sales Document', 'Order', 'Article', 'Material Description',
            'Requirement quantity (EINHEIT)', 'Initial_Stock_MB52', 'Opening_Stock_Balance',
            'Cumulative_Demand', 'Projected_Stock_Balance', 'Fulfilled_Qty', 'Shortage_Qty', 'Line_Status'
        ]
        df_lines_export = df_lines[cols_to_export].rename(columns={'Requirement date_str': 'Requirement Date'})
        df_lines_export.to_excel(writer, sheet_name='Chronological_Timeline', index=False)
        ws_timeline = writer.sheets['Chronological_Timeline']
        ws_timeline.freeze_panes(1, 0)
        for col_idx, col_name in enumerate(df_lines_export.columns):
            ws_timeline.write(0, col_idx, col_name, header_format)
            ws_timeline.set_column(col_idx, col_idx, max(len(col_name) + 2, 14))

        # Sheet 5: Sales Order Shortage Details
        # This shows only the SO/article lines where shortage occurs.
        df_shortage = df_lines[df_lines['Shortage_Qty'] > 0].copy()
        shortage_cols = [
            'Sales Document', 'Requirement date_str', 'Article', 'Material Description',
            'Requirement quantity (EINHEIT)', 'Initial_Stock_MB52',
            'Opening_Stock_Balance', 'Fulfilled_Qty', 'Shortage_Qty', 'Line_Status'
        ]
        shortage_cols = [c for c in shortage_cols if c in df_shortage.columns]
        df_shortage_export = df_shortage[shortage_cols].rename(
            columns={'Requirement date_str': 'Requirement Date'}
        )
        df_shortage_export = df_shortage_export.sort_values(
            by=['Sales Document', 'Requirement Date', 'Article']
        )
        df_shortage_export.to_excel(writer, sheet_name='SO_Shortage_Details', index=False)
        ws_short = writer.sheets['SO_Shortage_Details']
        ws_short.freeze_panes(1, 0)
        for col_idx, col_name in enumerate(df_shortage_export.columns):
            ws_short.write(0, col_idx, col_name, header_format)
            ws_short.set_column(col_idx, col_idx, max(len(col_name) + 3, 15))

        if 'Shortage_Qty' in df_shortage_export.columns:
            shortage_col = df_shortage_export.columns.get_loc('Shortage_Qty')
            for row_idx, val in enumerate(df_shortage_export['Shortage_Qty'], start=1):
                if float(val) > 0:
                    ws_short.write(row_idx, shortage_col, val, critical_format)

        # Sheet 6: Item-to-Sales Order Shortage Impact Report
        # Groups shortages by Material/Article to detail affected Sales Orders (SICs)
        if not df_shortage.empty:
            # Overall demand per item from all lines (both fulfilled and short)
            total_item_demand = df_lines.groupby(['Material Description', 'Article'], as_index=False).agg(
                Total_Demand_Orders=('Sales Document', 'nunique'),
                Total_Demand_Qty=('Requirement quantity (EINHEIT)', 'sum'),
                Initial_Available_Stock=('Initial_Stock_MB52', 'first')
            )
            
            # Shortage specific metrics per item
            shortage_item_summary = df_shortage.groupby(['Material Description', 'Article'], as_index=False).agg(
                Impacted_SICs_Count=('Sales Document', 'nunique'),
                Impacted_Sales_Orders=('Sales Document', lambda x: ', '.join(sorted([str(i) for i in x.unique() if pd.notna(i)]))),
                Total_Item_Shortage=('Shortage_Qty', 'sum')
            )
            
            impact_df = pd.merge(shortage_item_summary, total_item_demand, on=['Material Description', 'Article'], how='left')
            impact_df = impact_df[[
                'Material Description', 'Article', 'Impacted_SICs_Count', 'Impacted_Sales_Orders',
                'Total_Demand_Orders', 'Total_Demand_Qty', 'Total_Item_Shortage', 'Initial_Available_Stock'
            ]].sort_values(by=['Impacted_SICs_Count', 'Total_Item_Shortage'], ascending=[False, False])
        else:
            impact_df = pd.DataFrame(columns=['Material Description', 'Article', 'Impacted_SICs_Count', 'Impacted_Sales_Orders', 'Total_Demand_Orders', 'Total_Demand_Qty', 'Total_Item_Shortage', 'Initial_Available_Stock'])

        impact_df.to_excel(writer, sheet_name='Item_Shortage_Impact_Report', index=False)
        ws_impact = writer.sheets['Item_Shortage_Impact_Report']
        ws_impact.freeze_panes(1, 0)
        for col_idx, col_name in enumerate(impact_df.columns):
            ws_impact.write(0, col_idx, col_name, header_format)
            ws_impact.set_column(col_idx, col_idx, max(len(col_name) + 3, 20))

    output.seek(0)
    return output


def generate_so_excel_report(so_summary, df_lines):
    """
    Generates a single-tab Excel report for Sales Order Fulfillment.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#1E293B', 'font_color': '#FFFFFF',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        
        critical_format = workbook.add_format({'bg_color': '#FECDD3', 'font_color': '#9F1239', 'bold': True, 'border': 1})
        risk_format = workbook.add_format({'bg_color': '#FEF08A', 'font_color': '#854D0E', 'bold': True, 'border': 1})
        ok_format = workbook.add_format({'bg_color': '#D1FAE5', 'font_color': '#065F46', 'bold': True, 'border': 1})

        df_so_export = so_summary.copy()
        df_so_export['Earliest_Delivery_Date'] = df_so_export['Earliest_Delivery_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
        df_so_export['Latest_Delivery_Date'] = df_so_export['Latest_Delivery_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')

        df_so_export.to_excel(writer, sheet_name='Sales_Order_Fulfillment', index=False)
        ws_so = writer.sheets['Sales_Order_Fulfillment']
        ws_so.freeze_panes(1, 0)
        for col_idx, col_name in enumerate(df_so_export.columns):
            ws_so.write(0, col_idx, col_name, header_format)
            ws_so.set_column(col_idx, col_idx, max(len(col_name) + 3, 15))

        if 'SO_Status' in df_so_export.columns:
            status_col = df_so_export.columns.get_loc('SO_Status')
            for row_idx, val in enumerate(df_so_export['SO_Status'], start=1):
                if 'UNFULFILLABLE' in str(val):
                    ws_so.write(row_idx, status_col, val, critical_format)
                elif 'PARTIAL' in str(val):
                    ws_so.write(row_idx, status_col, val, risk_format)
                else:
                    ws_so.write(row_idx, status_col, val, ok_format)

        # Second sheet: only the lines that have a shortage
        df_shortage = df_lines[df_lines['Shortage_Qty'] > 0].copy()
        shortage_cols = [
            'Sales Document', 'Requirement date_str', 'Article', 'Material Description',
            'Requirement quantity (EINHEIT)', 'Initial_Stock_MB52',
            'Opening_Stock_Balance', 'Fulfilled_Qty', 'Shortage_Qty', 'Line_Status'
        ]
        shortage_cols = [c for c in shortage_cols if c in df_shortage.columns]
        df_shortage_export = df_shortage[shortage_cols].rename(
            columns={'Requirement date_str': 'Requirement Date'}
        )
        df_shortage_export = df_shortage_export.sort_values(
            by=['Sales Document', 'Requirement Date', 'Article']
        )
        df_shortage_export.to_excel(writer, sheet_name='SO_Shortage_Details', index=False)
        ws_short = writer.sheets['SO_Shortage_Details']
        ws_short.freeze_panes(1, 0)
        for col_idx, col_name in enumerate(df_shortage_export.columns):
            ws_short.write(0, col_idx, col_name, header_format)
            ws_short.set_column(col_idx, col_idx, max(len(col_name) + 3, 15))
        if 'Shortage_Qty' in df_shortage_export.columns:
            shortage_col = df_shortage_export.columns.get_loc('Shortage_Qty')
            for row_idx, val in enumerate(df_shortage_export['Shortage_Qty'], start=1):
                if float(val) > 0:
                    ws_short.write(row_idx, shortage_col, val, critical_format)

    output.seek(0)
    return output


@app.route('/', methods=['GET'])
def index():
    return render_template('availability.html')


@app.route('/predict', methods=['POST'])
def predict():
    file_mb52 = request.files.get('mb52_file')
    cohv_files = request.files.getlist('cohv_files')

    if not file_mb52 or not cohv_files or cohv_files[0].filename == '':
        return jsonify({'error': 'Please upload one MB52 stock file and at least one COHV requirement file.'}), 400

    try:
        art_summary, mat_summary, so_summary, df_lines = calculate_availability_predictions(
            file_mb52, cohv_files
        )

        req_format = request.form.get('format', request.args.get('format', 'excel'))

        if req_format == 'excel':
            excel_file = generate_excel_report(mat_summary, art_summary, so_summary, df_lines)
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Product_Availability_Review_{datetime.datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif req_format == 'so_excel':
            excel_file = generate_so_excel_report(so_summary, df_lines)
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Sales_Order_Fulfillment_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif req_format == 'json' or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            tot_stock = float(art_summary['Initial_Stock_MB52'].sum())
            tot_stock_val = float(art_summary['Initial_Stock_Value'].sum())
            tot_demand = float(art_summary['Total_Future_Demand'].sum())
            tot_demand_val = float(art_summary['Total_Demand_Value'].sum())
            tot_shortage = float(art_summary['Total_Shortage_Qty'].sum())
            tot_shortage_val = float(art_summary['Total_Shortage_Value'].sum())
            
            critical_count = int((art_summary['Status'] == 'CRITICAL STOCKOUT').sum())
            risk_count = int((art_summary['Status'] == 'STOCKOUT RISK').sum())
            sufficient_count = int((art_summary['Status'] == 'SUFFICIENT STOCK').sum())

            overall_fulfillment = round(min(100.0, ((tot_demand - tot_shortage) / tot_demand * 100.0)), 1) if tot_demand > 0 else 100.0

            depletion_dates = art_summary[art_summary['Days_Until_Depletion'] < 999]['Stock_Depletion_Date']
            earliest_depletion = depletion_dates.iloc[0] if not depletion_dates.empty else 'No Depletion Expected'

            # Days on Hand (DOH) Formula: Available Stock Value / (Sales Order Value / 30)
            daily_demand_val = tot_demand_val / 30.0 if tot_demand_val > 0 else 0
            days_on_hand = round(tot_stock_val / daily_demand_val, 1) if daily_demand_val > 0 else (999.0 if tot_stock_val > 0 else 0.0)

            # Top 10 Stockout Risk Materials for chart
            top_risk_mat = mat_summary[mat_summary['Total_Shortage_Qty'] > 0].head(10).fillna('')
            top_risk_mat_list = top_risk_mat[['Material Description', 'Initial_Stock_MB52', 'Total_Future_Demand', 'Total_Shortage_Qty', 'Stock_Depletion_Date']].to_dict(orient='records')

            # Top 10 Stockout Risk Articles
            top_risk_df = art_summary[art_summary['Total_Shortage_Qty'] > 0].head(10).fillna('')
            top_risk = top_risk_df[['Article', 'Material Description', 'Initial_Stock_MB52', 'Total_Future_Demand', 'Total_Shortage_Qty', 'Stock_Depletion_Date']].to_dict(orient='records')

            # SO Status Counts
            so_counts = so_summary['SO_Status'].value_counts().to_dict()

            # Material Description Rows
            mat_rows_df = mat_summary.copy()
            mat_rows_df['Earliest_Demand_Date'] = mat_rows_df['Earliest_Demand_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
            mat_rows_df['Latest_Demand_Date'] = mat_rows_df['Latest_Demand_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
            mat_table_rows = mat_rows_df.head(150).fillna('').to_dict(orient='records')

            # Article Rows
            art_rows_df = art_summary.copy()
            art_rows_df['Earliest_Demand_Date'] = art_rows_df['Earliest_Demand_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
            art_rows_df['Latest_Demand_Date'] = art_rows_df['Latest_Demand_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
            art_table_rows = art_rows_df.head(150).fillna('').to_dict(orient='records')

            # Sales Order Rows
            so_table_rows = so_summary.copy()
            so_table_rows['Earliest_Delivery_Date'] = so_table_rows['Earliest_Delivery_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
            so_table_rows['Latest_Delivery_Date'] = so_table_rows['Latest_Delivery_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
            so_table_rows = so_table_rows.fillna('').to_dict(orient='records')

            # Sales Order Line Items Breakdown for Modal Drilldown
            cols_so_items = [
                'Sales Document', 'Article', 'Material Description',
                'Requirement quantity (EINHEIT)', 'Requirement date_str',
                'Initial_Stock_MB52', 'Opening_Stock_Balance', 'Fulfilled_Qty', 'Shortage_Qty'
            ]
            so_line_df = df_lines[cols_so_items].copy()
            so_line_df.rename(columns={
                'Requirement quantity (EINHEIT)': 'Ordered_Qty',
                'Requirement date_str': 'Requirement_Date'
            }, inplace=True)

            # Consolidate repeat line items per Sales Document and Article so each article appears exactly once
            so_line_df = so_line_df.groupby(
                ['Sales Document', 'Article'],
                as_index=False
            ).agg({
                'Material Description': 'first',
                'Requirement_Date': lambda x: f"{x.min()} → {x.max()}" if x.min() != x.max() else x.min(),
                'Ordered_Qty': 'sum',
                'Fulfilled_Qty': 'sum',
                'Shortage_Qty': 'sum',
                'Initial_Stock_MB52': 'first',
                'Opening_Stock_Balance': 'first'
            })

            so_line_df['Line_Status'] = np.where(
                so_line_df['Shortage_Qty'].round(4) <= 0,
                'FULLY COVERED ON TIME',
                np.where(so_line_df['Fulfilled_Qty'].round(4) > 0, 'PARTIALLY COVERED', 'STOCKOUT / UNFULFILLABLE')
            )
            so_line_df = so_line_df.fillna('')

            so_line_dict = {}
            for so_id, group in so_line_df.groupby('Sales Document'):
                clean_key = str(so_id).strip().replace('.0', '')
                so_line_dict[clean_key] = group.to_dict(orient='records')

            # Build Item Shortage Impact Dictionary
            df_shortage_lines = df_lines[df_lines['Shortage_Qty'] > 0.0001].copy()
            item_impact_dict = {}
            if not df_shortage_lines.empty:
                for (mat_desc, art_id), grp in df_shortage_lines.groupby(['Material Description', 'Article']):
                    grp_clean = grp[['Sales Document', 'Requirement date_str', 'Requirement quantity (EINHEIT)', 'Fulfilled_Qty', 'Shortage_Qty']].rename(
                        columns={'Requirement date_str': 'Requirement_Date', 'Requirement quantity (EINHEIT)': 'Ordered_Qty'}
                    ).fillna('')
                    impacted_sos = grp_clean.to_dict(orient='records')
                    
                    # Fetch total demand across all lines (both fulfilled and short)
                    all_item_lines = df_lines[df_lines['Article'] == art_id]
                    if all_item_lines.empty:
                        all_item_lines = df_lines[df_lines['Material Description'] == mat_desc]
                    
                    tot_orders_count = int(all_item_lines['Sales Document'].nunique()) if not all_item_lines.empty else len(grp['Sales Document'].unique())
                    tot_demand_qty_val = float(all_item_lines['Requirement quantity (EINHEIT)'].sum()) if not all_item_lines.empty else float(grp['Requirement quantity (EINHEIT)'].sum())

                    impact_obj = {
                        'Material_Description': str(mat_desc) if pd.notna(mat_desc) else '',
                        'Article': str(art_id) if pd.notna(art_id) else '',
                        'Impacted_SICs_Count': len(grp['Sales Document'].unique()),
                        'Total_Demand_Orders': tot_orders_count,
                        'Total_Demand_Qty': tot_demand_qty_val,
                        'Total_Shortage_Qty': float(grp['Shortage_Qty'].sum()),
                        'Impacted_Sales_Orders': impacted_sos
                    }
                    item_impact_dict[art_id] = impact_obj
                    if mat_desc not in item_impact_dict:
                        item_impact_dict[mat_desc] = impact_obj

            response_dict = {
                'success': True,
                'kpis': {
                    'total_stock': tot_stock,
                    'total_stock_val': tot_stock_val,
                    'total_demand': tot_demand,
                    'total_demand_val': tot_demand_val,
                    'total_shortage': tot_shortage,
                    'total_shortage_val': tot_shortage_val,
                    'fulfillment_rate': overall_fulfillment,
                    'critical_count': critical_count,
                    'risk_count': risk_count,
                    'sufficient_count': sufficient_count,
                    'earliest_depletion': earliest_depletion,
                    'days_on_hand': days_on_hand
                },
                'top_risk_articles': top_risk,
                'top_risk_materials': top_risk_mat_list,
                'so_status_counts': so_counts,
                'mat_table_rows': mat_table_rows,
                'art_table_rows': art_table_rows,
                'so_table_rows': so_table_rows,
                'so_line_dict': so_line_dict,
                'item_impact_dict': item_impact_dict
            }

            def clean_nan(obj):
                if isinstance(obj, float):
                    if np.isnan(obj) or np.isinf(obj):
                        return 0.0
                    return obj
                elif isinstance(obj, dict):
                    return {k: clean_nan(v) for k, v in obj.items()}
                elif isinstance(obj, list):
                    return [clean_nan(v) for v in obj]
                return obj

            sanitized_response = clean_nan(response_dict)
            return app.response_class(
                response=json.dumps(sanitized_response),
                status=200,
                mimetype='application/json'
            )
        
        else:
            excel_file = generate_excel_report(art_summary, mat_summary, so_summary, df_lines)
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Product_Availability_Prediction_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        import traceback
        traceback.print_exc()
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.form.get('format') == 'json':
            return jsonify({'error': str(e)}), 500
        return f"An error occurred during prediction analysis: {str(e)}", 500


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5001))
    print(f"Starting SAP Product Availability Predictor on http://0.0.0.0:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)

